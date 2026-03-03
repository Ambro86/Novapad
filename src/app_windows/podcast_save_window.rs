use crate::accessibility::{ES_CENTER, ES_READONLY, handle_accessibility, to_wide};
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{PBM_SETPOS, PBM_SETRANGE, WC_BUTTON};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, SetFocus, VK_ESCAPE, VK_RETURN, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, GWLP_USERDATA,
    GetParent, HMENU, IDC_ARROW, IDYES, LoadCursorW, MB_ICONWARNING, MB_YESNO, MSG, PostMessageW,
    RegisterClassW, SendMessageW, SetForegroundWindow, SetTimer, SetWindowLongPtrW, SetWindowTextW,
    WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_KEYDOWN, WM_NCDESTROY, WM_SETFOCUS,
    WM_SETFONT, WM_SYSKEYDOWN, WM_TIMER, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_DLGMODALFRAME,
    WS_POPUP, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

pub const WM_PODCAST_SAVE_DONE: u32 = WM_APP + 70;
pub const WM_PODCAST_SAVE_CLOSED: u32 = WM_APP + 71;
pub const WM_PODCAST_SAVE_PROGRESS: u32 = WM_APP + 72;
pub const WM_PODCAST_SAVE_CANCEL: u32 = WM_APP + 73;
const SAVE_CLASS_NAME: &str = "SonarpadPodcastSave";
const SAVE_ID_CANCEL: usize = 12002;
const SAVE_PROGRESS_TIMER_ID: usize = 1;
const SAVE_PROGRESS_TICK_MS: u32 = 250;
const SAVE_PROGRESS_MAX_FAKE: usize = 95;

pub struct SaveDialogLabels {
    pub title: String,
    pub in_progress: String,
    pub cancel: String,
}

struct SaveCreateParams {
    parent: HWND,
    language: Language,
    labels: SaveDialogLabels,
    show_cancel: bool,
}

struct SaveState {
    parent: HWND,
    label: HWND,
    progress: HWND,
    cancel_button: HWND,
    cancel_requested: bool,
    language: Language,
    current_pct: usize,
    labels: SaveDialogLabels,
    show_cancel: bool,
}

fn save_labels(language: Language) -> SaveDialogLabels {
    SaveDialogLabels {
        title: i18n::tr(language, "podcast.save.title"),
        in_progress: i18n::tr(language, "podcast.save.in_progress"),
        cancel: i18n::tr(language, "podcast.save.cancel"),
    }
}

pub fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN {
        let key = msg.wParam.0 as u32;
        if key == VK_ESCAPE.0 as u32 {
            if let Err(_e) =
                crate::post_message_w_safe(hwnd, WM_COMMAND, WPARAM(SAVE_ID_CANCEL), LPARAM(0))
            {
                crate::log_debug(&format!("Error: {:?}", _e));
            }
            return true;
        }
        if key == VK_RETURN.0 as u32 {
            let focus = crate::get_focus_safe();
            let cancel = with_save_state(hwnd, |state| state.cancel_button).unwrap_or(HWND(0));
            if cancel.0 != 0 && focus == cancel {
                if let Err(_e) = crate::post_message_w_safe(
                    hwnd,
                    WM_COMMAND,
                    WPARAM(SAVE_ID_CANCEL),
                    LPARAM(cancel.0),
                ) {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
                return true;
            }
        }
    }
    handle_accessibility(hwnd, msg)
}

pub fn open(parent: HWND) -> HWND {
    unsafe {
        let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
        let class_name = to_wide(SAVE_CLASS_NAME);
        let main = GetParent(parent);
        let language = if main.0 != 0 {
            with_state(main, |state| state.settings.language).unwrap_or_default()
        } else {
            crate::app_windows::podcast_window::language_for_window(parent).unwrap_or_default()
        };
        let labels = save_labels(language);

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

        let params = Box::new(SaveCreateParams {
            parent,
            language,
            labels,
            show_cancel: true,
        });
        let params_ptr = Box::into_raw(params);
        let window = CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR::null(),
            WS_POPUP | WS_CAPTION | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            300,
            150,
            parent,
            HMENU(0),
            hinstance,
            Some(params_ptr as *const std::ffi::c_void),
        );
        if window.0 == 0 {
            let _unused = Box::from_raw(params_ptr);
            return window;
        }

        if window.0 != 0 {
            EnableWindow(parent, false);
            SetForegroundWindow(window);

            let mut rc_parent = RECT::default();
            let mut rc_dlg = RECT::default();
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::GetWindowRect(
                parent,
                &mut rc_parent
            ));
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::GetWindowRect(
                window,
                &mut rc_dlg
            ));
            let dlg_w = rc_dlg.right - rc_dlg.left;
            let dlg_h = rc_dlg.bottom - rc_dlg.top;
            let parent_w = rc_parent.right - rc_parent.left;
            let parent_h = rc_parent.bottom - rc_parent.top;
            let x = rc_parent.left + (parent_w - dlg_w) / 2;
            let y = rc_parent.top + (parent_h - dlg_h) / 2;
            use windows::Win32::UI::WindowsAndMessaging::{HWND_TOP, SWP_SHOWWINDOW, SetWindowPos};
            if let Err(e) = SetWindowPos(window, HWND_TOP, x, y, dlg_w, dlg_h, SWP_SHOWWINDOW) {
                crate::log_debug(&format!("Failed to position save window: {}", e));
            }
        }
        window
    }
}

pub fn open_with_labels(
    parent: HWND,
    language: Language,
    labels: SaveDialogLabels,
    show_cancel: bool,
) -> HWND {
    unsafe {
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

        let params = Box::new(SaveCreateParams {
            parent,
            language,
            labels,
            show_cancel,
        });
        let params_ptr = Box::into_raw(params);
        let window = CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR::null(),
            WS_POPUP | WS_CAPTION | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            300,
            150,
            parent,
            HMENU(0),
            hinstance,
            Some(params_ptr as *const std::ffi::c_void),
        );
        if window.0 == 0 {
            let _unused = Box::from_raw(params_ptr);
            return window;
        }

        EnableWindow(parent, false);
        SetForegroundWindow(window);

        let mut rc_parent = RECT::default();
        let mut rc_dlg = RECT::default();
        crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::GetWindowRect(
            parent,
            &mut rc_parent
        ));
        crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::GetWindowRect(
            window,
            &mut rc_dlg
        ));
        let dlg_w = rc_dlg.right - rc_dlg.left;
        let dlg_h = rc_dlg.bottom - rc_dlg.top;
        let parent_w = rc_parent.right - rc_parent.left;
        let parent_h = rc_parent.bottom - rc_parent.top;
        let x = rc_parent.left + (parent_w - dlg_w) / 2;
        let y = rc_parent.top + (parent_h - dlg_h) / 2;
        use windows::Win32::UI::WindowsAndMessaging::{HWND_TOP, SWP_SHOWWINDOW, SetWindowPos};
        if let Err(e) = SetWindowPos(window, HWND_TOP, x, y, dlg_w, dlg_h, SWP_SHOWWINDOW) {
            crate::log_debug(&format!("Failed to position save window: {}", e));
        }

        window
    }
}

unsafe extern "system" fn save_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "save_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || save_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn save_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let params_ptr = unsafe { (*create_struct).lpCreateParams as *mut SaveCreateParams };
            let params = crate::box_from_raw_safe(params_ptr);
            let parent = params.parent;
            let language = params.language;
            let labels = params.labels;
            let show_cancel = params.show_cancel;
            let main = crate::get_parent_safe(parent);
            let hfont = { with_state(main, |state| state.hfont) }.unwrap_or(HFONT(0));
            let label_text = format!("{} 0%", labels.in_progress);

            let label = unsafe {
                CreateWindowExW(
                    Default::default(),
                    w!("EDIT"),
                    PCWSTR(to_wide(&label_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_CENTER | ES_READONLY),
                    20,
                    20,
                    260,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };

            let progress = unsafe {
                CreateWindowExW(
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
                )
            };

            let cancel_button = if show_cancel {
                unsafe {
                    CreateWindowExW(
                        Default::default(),
                        WC_BUTTON,
                        PCWSTR(to_wide(&labels.cancel).as_ptr()),
                        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                        95,
                        80,
                        90,
                        28,
                        hwnd,
                        HMENU(SAVE_ID_CANCEL as isize),
                        HINSTANCE(0),
                        None,
                    )
                }
            } else {
                HWND(0)
            };

            if let Err(e) =
                crate::set_window_text_w_safe(hwnd, PCWSTR(to_wide(&labels.title).as_ptr()))
            {
                crate::log_debug(&format!("Failed to set title: {}", e));
            }
            unsafe {
                SendMessageW(label, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                SendMessageW(progress, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }
            if cancel_button.0 != 0 {
                crate::send_message_w_safe(
                    cancel_button,
                    WM_SETFONT,
                    WPARAM(hfont.0 as usize),
                    LPARAM(1),
                );
            }
            unsafe {
                SendMessageW(progress, PBM_SETRANGE, WPARAM(0), LPARAM((100isize) << 16));
                SendMessageW(progress, PBM_SETPOS, WPARAM(0), LPARAM(0));
            }

            let state = SaveState {
                parent,
                label,
                progress,
                cancel_button,
                cancel_requested: false,
                language,
                current_pct: 0,
                labels,
                show_cancel,
            };
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(Box::new(state)) as isize);
                SetFocus(label);
            }
            if unsafe { SetTimer(hwnd, SAVE_PROGRESS_TIMER_ID, SAVE_PROGRESS_TICK_MS, None) } == 0 {
                crate::log_debug("Failed to set SAVE_PROGRESS_TIMER");
            }
            LRESULT(0)
        }
        WM_SETFOCUS => {
            if with_save_state(hwnd, |state| {
                crate::set_focus_safe(state.label);
            })
            .is_none()
            {
                crate::log_debug("Failed to access save state");
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = wparam.0 & 0xffff;
            if id == SAVE_ID_CANCEL || id == 2 {
                request_cancel(hwnd);
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_KEYDOWN | WM_SYSKEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                request_cancel(hwnd);
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_TAB.0 as u32
                && let Some((label, cancel, show_cancel)) = with_save_state(hwnd, |state| {
                    (state.label, state.cancel_button, state.show_cancel)
                })
                && show_cancel
                && cancel.0 != 0
            {
                let focus = crate::get_focus_safe();
                let shift_down =
                    (crate::get_key_state_safe(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
                if !shift_down && focus != cancel {
                    crate::set_focus_safe(cancel);
                    return LRESULT(0);
                }
                if shift_down && focus == cancel {
                    crate::set_focus_safe(label);
                    return LRESULT(0);
                }
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32
                && let Some(cancel) = with_save_state(hwnd, |state| state.cancel_button)
                && cancel.0 != 0
                && crate::get_focus_safe() == cancel
            {
                if let Err(_e) = crate::post_message_w_safe(
                    hwnd,
                    WM_COMMAND,
                    WPARAM(SAVE_ID_CANCEL),
                    LPARAM(cancel.0),
                ) {}
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_PODCAST_SAVE_PROGRESS => {
            let pct = wparam.0.min(100);
            if with_save_state(hwnd, |state| {
                crate::send_message_w_safe(state.progress, PBM_SETPOS, WPARAM(pct), LPARAM(0));
                state.current_pct = state.current_pct.max(pct);
                update_progress_label(state);
            })
            .is_none()
            {
                crate::log_debug("Failed to access save state");
            }
            LRESULT(0)
        }
        WM_TIMER => {
            if wparam.0 == SAVE_PROGRESS_TIMER_ID {
                if with_save_state(hwnd, |state| {
                    if state.current_pct < SAVE_PROGRESS_MAX_FAKE {
                        state.current_pct = (state.current_pct + 1).min(SAVE_PROGRESS_MAX_FAKE);
                        crate::send_message_w_safe(
                            state.progress,
                            PBM_SETPOS,
                            WPARAM(state.current_pct),
                            LPARAM(0),
                        );
                        update_progress_label(state);
                    }
                })
                .is_none()
                {
                    crate::log_debug("Failed to access save state");
                }
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_PODCAST_SAVE_DONE => {
            crate::log_if_err!(crate::destroy_window_safe(hwnd));
            LRESULT(0)
        }
        WM_CLOSE => {
            request_cancel(hwnd);
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let parent = with_save_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if let Err(e) = crate::kill_timer_safe(hwnd, SAVE_PROGRESS_TIMER_ID) {
                crate::log_debug(&format!("Failed to kill SAVE_PROGRESS_TIMER: {}", e));
            }
            if parent.0 != 0 {
                crate::enable_window_safe(parent, true);
                // Keep focus within Sonarpad when progress dialogs close (e.g. streaming
                // download -> conversion handoff), avoiding transient desktop focus.
                crate::set_foreground_window_safe(parent);
                if let Err(_e) =
                    crate::post_message_w_safe(parent, WM_PODCAST_SAVE_CLOSED, WPARAM(0), LPARAM(0))
                {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
            }
            let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA);
            if ptr != 0 {
                let _unused_box = crate::box_from_raw_safe(ptr as *mut SaveState);
            }
            LRESULT(0)
        }
        _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
    }
}

fn with_save_state<T>(hwnd: HWND, f: impl FnOnce(&mut SaveState) -> T) -> Option<T> {
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut SaveState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

fn update_progress_label(state: &SaveState) {
    let text = format!("{} {}%", state.labels.in_progress, state.current_pct);
    unsafe {
        if let Err(e) = SetWindowTextW(state.label, PCWSTR(to_wide(&text).as_ptr())) {
            crate::log_debug(&format!("Failed to set label text: {}", e));
        }
    }
}

fn request_cancel(hwnd: HWND) {
    let mut should_post = false;
    if with_save_state(hwnd, |state| {
        if !state.show_cancel || state.cancel_button.0 == 0 {
            return;
        }
        if state.cancel_requested {
            return;
        }
        let msg = i18n::tr(state.language, "podcast.cancel_confirm");
        let title = i18n::tr(state.language, "app.confirm_title");
        let msg_w = to_wide(&msg);
        let title_w = to_wide(&title);
        let result = crate::message_box_w_safe(
            hwnd,
            PCWSTR(msg_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            MB_YESNO | MB_ICONWARNING,
        );
        if result == IDYES {
            state.cancel_requested = true;
            unsafe {
                EnableWindow(state.cancel_button, false);
            }
            should_post = true;
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access save state");
    }
    if should_post {
        let parent = crate::get_parent_safe(hwnd);
        if parent.0 != 0 {
            unsafe {
                if let Err(_e) = PostMessageW(parent, WM_PODCAST_SAVE_CANCEL, WPARAM(0), LPARAM(0))
                {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
            }
        }
    }
}
