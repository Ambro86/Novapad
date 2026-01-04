use crate::accessibility::{ES_CENTER, ES_READONLY, handle_accessibility, to_wide};
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{PBM_SETPOS, PBM_SETRANGE, WC_BUTTON};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    GWLP_USERDATA, GetParent, GetWindowLongPtrW, HMENU, IDC_ARROW, IDYES, LoadCursorW,
    MB_ICONWARNING, MB_YESNO, MSG, MessageBoxW, PostMessageW, RegisterClassW, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW, WINDOW_STYLE, WM_APP, WM_CLOSE,
    WM_COMMAND, WM_CREATE, WM_KEYDOWN, WM_NCDESTROY, WM_SETFOCUS, WM_SETFONT, WM_SYSKEYDOWN,
    WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_DLGMODALFRAME, WS_POPUP, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

pub const WM_PODCAST_SAVE_DONE: u32 = WM_APP + 70;
pub const WM_PODCAST_SAVE_CLOSED: u32 = WM_APP + 71;
pub const WM_PODCAST_SAVE_PROGRESS: u32 = WM_APP + 72;
pub const WM_PODCAST_SAVE_CANCEL: u32 = WM_APP + 73;
const SAVE_CLASS_NAME: &str = "NovapadPodcastSave";
const SAVE_ID_CANCEL: usize = 12002;

struct SaveState {
    parent: HWND,
    label: HWND,
    progress: HWND,
    cancel_button: HWND,
    cancel_requested: bool,
    language: Language,
}

fn save_labels(language: Language) -> (String, String, String, String, String, String) {
    (
        i18n::tr(language, "podcast.save.title"),
        i18n::tr(language, "podcast.save.in_progress"),
        i18n::tr(language, "podcast.save.done"),
        i18n::tr(language, "podcast.save.failed"),
        i18n::tr(language, "podcast.save.canceled"),
        i18n::tr(language, "podcast.save.cancel"),
    )
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN {
        let key = msg.wParam.0 as u32;
        if key == VK_ESCAPE.0 as u32 {
            let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(SAVE_ID_CANCEL), LPARAM(0));
            return true;
        }
        if key == VK_RETURN.0 as u32 {
            let focus = GetFocus();
            let cancel = with_save_state(hwnd, |state| state.cancel_button).unwrap_or(HWND(0));
            if cancel.0 != 0 && focus == cancel {
                let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(SAVE_ID_CANCEL), LPARAM(cancel.0));
                return true;
            }
        }
    }
    handle_accessibility(hwnd, msg)
}

pub unsafe fn open(parent: HWND) -> HWND {
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(SAVE_CLASS_NAME);

    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(save_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let window = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(to_wide("Saving").as_ptr()),
        WS_POPUP | WS_CAPTION | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        300,
        150,
        parent,
        HMENU(0),
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if window.0 != 0 {
        EnableWindow(parent, false);
        SetForegroundWindow(window);

        let mut rc_parent = RECT::default();
        let mut rc_dlg = RECT::default();
        let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(parent, &mut rc_parent);
        let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(window, &mut rc_dlg);
        let dlg_w = rc_dlg.right - rc_dlg.left;
        let dlg_h = rc_dlg.bottom - rc_dlg.top;
        let parent_w = rc_parent.right - rc_parent.left;
        let parent_h = rc_parent.bottom - rc_parent.top;
        let x = rc_parent.left + (parent_w - dlg_w) / 2;
        let y = rc_parent.top + (parent_h - dlg_h) / 2;
        let _ =
            windows::Win32::UI::WindowsAndMessaging::MoveWindow(window, x, y, dlg_w, dlg_h, true);
    }
    window
}

unsafe extern "system" fn save_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let main = GetParent(parent);
            let language = with_state(main, |state| state.settings.language).unwrap_or_default();
            let (title, in_progress, _, _, _, cancel_text) = save_labels(language);
            let hfont = with_state(main, |state| state.hfont).unwrap_or(HFONT(0));
            let label_text = format!("{in_progress} 0%");

            let label = CreateWindowExW(
                Default::default(),
                w!("EDIT"),
                PCWSTR(to_wide(&label_text).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE((ES_CENTER | ES_READONLY) as u32),
                20,
                20,
                260,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let progress = CreateWindowExW(
                Default::default(),
                w!("msctls_progress32"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                20,
                50,
                260,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let cancel_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&cancel_text).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                95,
                80,
                90,
                28,
                hwnd,
                HMENU(SAVE_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            let _ = SetWindowTextW(hwnd, PCWSTR(to_wide(&title).as_ptr()));
            let _ = SendMessageW(label, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            let _ = SendMessageW(progress, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            let _ = SendMessageW(
                cancel_button,
                WM_SETFONT,
                WPARAM(hfont.0 as usize),
                LPARAM(1),
            );
            let _ = SendMessageW(progress, PBM_SETRANGE, WPARAM(0), LPARAM((100isize) << 16));
            let _ = SendMessageW(progress, PBM_SETPOS, WPARAM(0), LPARAM(0));

            let state = SaveState {
                parent,
                label,
                progress,
                cancel_button,
                cancel_requested: false,
                language,
            };
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(Box::new(state)) as isize);
            SetFocus(label);
            LRESULT(0)
        }
        WM_SETFOCUS => {
            let _ = with_save_state(hwnd, |state| {
                SetFocus(state.label);
            });
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            if id == SAVE_ID_CANCEL || id == 2 {
                request_cancel(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_KEYDOWN | WM_SYSKEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                request_cancel(hwnd);
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                if let Some(cancel) = with_save_state(hwnd, |state| state.cancel_button) {
                    if cancel.0 != 0 && GetFocus() == cancel {
                        let _ = PostMessageW(
                            hwnd,
                            WM_COMMAND,
                            WPARAM(SAVE_ID_CANCEL),
                            LPARAM(cancel.0),
                        );
                        return LRESULT(0);
                    }
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_PODCAST_SAVE_PROGRESS => {
            let pct = wparam.0.min(100) as usize;
            let _ = with_save_state(hwnd, |state| {
                let _ = SendMessageW(state.progress, PBM_SETPOS, WPARAM(pct), LPARAM(0));
                let (_, in_progress, _, _, _, _) = save_labels(state.language);
                let text = format!("{in_progress} {pct}%");
                let _ = SetWindowTextW(state.label, PCWSTR(to_wide(&text).as_ptr()));
            });
            LRESULT(0)
        }
        WM_PODCAST_SAVE_DONE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_CLOSE => {
            request_cancel(hwnd);
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let parent = with_save_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                EnableWindow(parent, true);
                let _ = PostMessageW(parent, WM_PODCAST_SAVE_CLOSED, WPARAM(0), LPARAM(0));
            }
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
            if ptr != 0 {
                let _ = Box::from_raw(ptr as *mut SaveState);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

fn with_save_state<T>(hwnd: HWND, f: impl FnOnce(&mut SaveState) -> T) -> Option<T> {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut SaveState;
    if ptr.is_null() {
        None
    } else {
        Some(f(unsafe { &mut *ptr }))
    }
}

fn request_cancel(hwnd: HWND) {
    let mut should_post = false;
    let _ = with_save_state(hwnd, |state| {
        if state.cancel_requested {
            return;
        }
        let msg = i18n::tr(state.language, "podcast.cancel_confirm");
        let title = i18n::tr(state.language, "app.confirm_title");
        let msg_w = to_wide(&msg);
        let title_w = to_wide(&title);
        let result = unsafe {
            MessageBoxW(
                hwnd,
                PCWSTR(msg_w.as_ptr()),
                PCWSTR(title_w.as_ptr()),
                MB_YESNO | MB_ICONWARNING,
            )
        };
        if result == IDYES {
            state.cancel_requested = true;
            unsafe {
                EnableWindow(state.cancel_button, false);
            }
            should_post = true;
        }
    });
    if should_post {
        let parent = unsafe { GetParent(hwnd) };
        if parent.0 != 0 {
            unsafe {
                let _ = PostMessageW(parent, WM_PODCAST_SAVE_CANCEL, WPARAM(0), LPARAM(0));
            }
        }
    }
}
