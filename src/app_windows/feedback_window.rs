use crate::accessibility::{handle_accessibility, to_wide};
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetFocus, GetKeyState, SetFocus, VK_ESCAPE, VK_RETURN, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    ES_AUTOHSCROLL, ES_AUTOVSCROLL, ES_MULTILINE, ES_WANTRETURN, GWLP_USERDATA, GetWindowLongPtrW,
    HMENU, IDC_ARROW, LoadCursorW, MB_ICONERROR, MB_OK, MSG, RegisterClassW, SW_SHOW,
    SetForegroundWindow, SetWindowLongPtrW, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE,
    WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFOCUS, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_DLGMODALFRAME, WS_POPUP, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const FEEDBACK_CLASS_NAME: &str = "SonarpadFeedback";
const FEEDBACK_ID_SUBJECT: usize = 7101;
const FEEDBACK_ID_MESSAGE: usize = 7102;
const FEEDBACK_ID_SEND: usize = 7103;
const FEEDBACK_ID_CANCEL: usize = 7104;
const FEEDBACK_TARGET: &str = "ambro86@gmail.com";

struct FeedbackState {
    parent: HWND,
    subject_edit: HWND,
    message_edit: HWND,
    send_button: HWND,
    cancel_button: HWND,
    prev_focus: HWND,
}

struct FeedbackLabels {
    title: String,
    subject: String,
    message: String,
    send: String,
    cancel: String,
    open_mail_error: String,
}

fn labels_for(language: Language) -> FeedbackLabels {
    FeedbackLabels {
        title: i18n::tr(language, "feedback.title"),
        subject: i18n::tr(language, "feedback.label.subject"),
        message: i18n::tr(language, "feedback.label.message"),
        send: i18n::tr(language, "feedback.button.send"),
        cancel: i18n::tr(language, "go_to_time.cancel"),
        open_mail_error: i18n::tr(language, "feedback.error.open_mail"),
    }
}

pub fn open(parent: HWND) {
    unsafe {
        let existing = with_state(parent, |state| state.feedback_window).unwrap_or(HWND(0));
        if existing.0 != 0 {
            SetForegroundWindow(existing);
            return;
        }

        let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
        let labels = labels_for(language);
        let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
        let class_name = to_wide(FEEDBACK_CLASS_NAME);
        let title = to_wide(&labels.title);
        let wc = WNDCLASSW {
            hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
                LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
            ),
            hInstance: hinstance,
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(feedback_wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let state = Box::new(FeedbackState {
            parent,
            subject_edit: HWND(0),
            message_edit: HWND(0),
            send_button: HWND(0),
            cancel_button: HWND(0),
            prev_focus: GetFocus(),
        });
        let state_ptr = Box::into_raw(state);
        let hwnd = CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_POPUP | WS_CAPTION | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            620,
            420,
            parent,
            HMENU(0),
            hinstance,
            Some(state_ptr as *const _),
        );
        if hwnd.0 == 0 {
            let _unused_box = Box::from_raw(state_ptr);
            return;
        }

        crate::enable_window_safe(parent, false);
        with_state(parent, |state| state.feedback_window = hwnd);
    }
}

unsafe extern "system" fn feedback_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "feedback_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || feedback_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn feedback_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let cs = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*cs).lpCreateParams as *mut FeedbackState;
                if init_ptr.is_null() {
                    return LRESULT(0);
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, init_ptr as isize);

                let parent = (*init_ptr).parent;
                let language =
                    with_state(parent, |state| state.settings.language).unwrap_or_default();
                let labels = labels_for(language);
                let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);

                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&labels.subject).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    20,
                    18,
                    560,
                    20,
                    hwnd,
                    HMENU(1),
                    hinstance,
                    None,
                );
                let subject_edit = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    20,
                    42,
                    560,
                    26,
                    hwnd,
                    HMENU(FEEDBACK_ID_SUBJECT as isize),
                    hinstance,
                    None,
                );

                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&labels.message).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    20,
                    84,
                    560,
                    20,
                    hwnd,
                    HMENU(2),
                    hinstance,
                    None,
                );
                let message_edit = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WS_VSCROLL
                        | WINDOW_STYLE(
                            ES_MULTILINE as u32 | ES_AUTOVSCROLL as u32 | ES_WANTRETURN as u32,
                        ),
                    20,
                    108,
                    560,
                    220,
                    hwnd,
                    HMENU(FEEDBACK_ID_MESSAGE as isize),
                    hinstance,
                    None,
                );
                let send_button = CreateWindowExW(
                    Default::default(),
                    w!("BUTTON"),
                    PCWSTR(to_wide(&labels.send).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    368,
                    344,
                    100,
                    30,
                    hwnd,
                    HMENU(FEEDBACK_ID_SEND as isize),
                    hinstance,
                    None,
                );
                let cancel_button = CreateWindowExW(
                    Default::default(),
                    w!("BUTTON"),
                    PCWSTR(to_wide(&labels.cancel).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    480,
                    344,
                    100,
                    30,
                    hwnd,
                    HMENU(FEEDBACK_ID_CANCEL as isize),
                    hinstance,
                    None,
                );

                (*init_ptr).subject_edit = subject_edit;
                (*init_ptr).message_edit = message_edit;
                (*init_ptr).send_button = send_button;
                (*init_ptr).cancel_button = cancel_button;
                SetFocus(subject_edit);
                LRESULT(0)
            }
            WM_SETFOCUS => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut FeedbackState;
                if ptr.is_null() {
                    return LRESULT(0);
                }
                SetFocus((*ptr).subject_edit);
                LRESULT(0)
            }
            WM_COMMAND => {
                if (wparam.0 & 0xffff) == FEEDBACK_ID_SEND {
                    send_feedback(hwnd);
                    return LRESULT(0);
                }
                if (wparam.0 & 0xffff) == FEEDBACK_ID_CANCEL {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => {
                let parent = crate::get_parent_safe(hwnd);
                if parent.0 != 0 {
                    crate::enable_window_safe(parent, true);
                    crate::set_foreground_window_safe(parent);
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut FeedbackState;
                if !ptr.is_null() {
                    let state = Box::from_raw(ptr);
                    with_state(state.parent, |app_state| {
                        app_state.feedback_window = HWND(0)
                    });
                    if state.prev_focus.0 != 0 {
                        SetFocus(state.prev_focus);
                    }
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

pub fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    unsafe {
        if msg.message != WM_KEYDOWN {
            return handle_accessibility(hwnd, msg);
        }

        let key = msg.wParam.0 as u32;
        if key == VK_ESCAPE.0 as u32 {
            crate::log_if_err!(crate::destroy_window_safe(hwnd));
            return true;
        }

        if key == VK_TAB.0 as u32 {
            let shift_down = GetKeyState(VK_SHIFT.0 as i32) < 0;
            if with_feedback_state(hwnd, |state| {
                let focused = GetFocus();
                let target = if shift_down {
                    if focused == state.cancel_button {
                        state.send_button
                    } else if focused == state.send_button {
                        state.message_edit
                    } else if focused == state.message_edit {
                        state.subject_edit
                    } else {
                        state.cancel_button
                    }
                } else if focused == state.subject_edit {
                    state.message_edit
                } else if focused == state.message_edit {
                    state.send_button
                } else if focused == state.send_button {
                    state.cancel_button
                } else {
                    state.subject_edit
                };
                SetFocus(target);
            })
            .is_none()
            {
                crate::log_debug("Failed to access feedback window state");
            }
            return true;
        }

        if key == VK_RETURN.0 as u32 {
            let focus = GetFocus();
            let send_button =
                with_feedback_state(hwnd, |state| state.send_button).unwrap_or(HWND(0));
            if focus == send_button {
                send_feedback(hwnd);
                return true;
            }
        }
    }

    handle_accessibility(hwnd, msg)
}

fn send_feedback(hwnd: HWND) {
    let Some((subject, message, parent)) = with_feedback_state(hwnd, |state| {
        (
            get_window_text(state.subject_edit),
            get_window_text(state.message_edit),
            state.parent,
        )
    }) else {
        crate::log_debug("Failed to read feedback window state");
        return;
    };

    let uri = build_mailto_uri(&subject, &message);
    let uri_wide = to_wide(&uri);
    unsafe {
        let result = ShellExecuteW(
            hwnd,
            w!("open"),
            PCWSTR(uri_wide.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOW,
        );
        if result.0 as isize <= 32 {
            crate::log_debug(&format!("Feedback: ShellExecuteW failed with {}", result.0));
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = labels_for(language);
            let message_wide = to_wide(&labels.open_mail_error);
            let title_wide = to_wide(&labels.title);
            crate::message_box_w_safe(
                hwnd,
                PCWSTR(message_wide.as_ptr()),
                PCWSTR(title_wide.as_ptr()),
                MB_OK | MB_ICONERROR,
            );
            return;
        }
    }

    crate::log_if_err!(crate::destroy_window_safe(hwnd));
}

fn build_mailto_uri(subject: &str, message: &str) -> String {
    let mut uri = format!("mailto:{FEEDBACK_TARGET}");
    let encoded_subject = mailto_encode_component(subject);
    let encoded_message = mailto_encode_component(message);

    if !encoded_subject.is_empty() || !encoded_message.is_empty() {
        uri.push('?');
        if !encoded_subject.is_empty() {
            uri.push_str("subject=");
            uri.push_str(&encoded_subject);
        }
        if !encoded_message.is_empty() {
            if !encoded_subject.is_empty() {
                uri.push('&');
            }
            uri.push_str("body=");
            uri.push_str(&encoded_message);
        }
    }

    uri
}

fn mailto_encode_component(input: &str) -> String {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut out = String::with_capacity(input.len() * 3);
    for byte in input.bytes() {
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'.' | b'_' | b'~' => {
                out.push(byte as char)
            }
            _ => {
                out.push('%');
                out.push(HEX[(byte >> 4) as usize] as char);
                out.push(HEX[(byte & 0x0F) as usize] as char);
            }
        }
    }
    out
}

fn get_window_text(hwnd: HWND) -> String {
    let len = crate::get_window_text_length_w_safe(hwnd);
    if len <= 0 {
        return String::new();
    }

    let mut buffer = vec![0u16; len as usize + 1];
    let written = crate::get_window_text_w_safe(hwnd, &mut buffer);
    String::from_utf16_lossy(&buffer[..written.max(0) as usize])
}

fn with_feedback_state<R>(hwnd: HWND, f: impl FnOnce(&mut FeedbackState) -> R) -> Option<R> {
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut FeedbackState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}
