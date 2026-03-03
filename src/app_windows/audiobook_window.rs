use crate::accessibility::{ES_CENTER, ES_READONLY, handle_accessibility, to_wide};
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use std::sync::atomic::Ordering;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{PBM_SETPOS, PBM_SETRANGE, WC_BUTTON};
use windows::Win32::UI::Input::KeyboardAndMouse::{SetFocus, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, GWLP_USERDATA,
    HMENU, IDC_ARROW, IDYES, LoadCursorW, MB_ICONWARNING, MB_YESNO, MSG, MoveWindow,
    RegisterClassW, SendMessageW, SetForegroundWindow, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFOCUS, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

const PROGRESS_CLASS_NAME: &str = "SonarpadProgress";
const PROGRESS_ID_CANCEL: usize = 8001;
const WM_UPDATE_PROGRESS: u32 = WM_APP + 6;
pub const WM_SET_PROGRESS_TOTAL: u32 = WM_APP + 8;

struct ProgressDialogState {
    parent: HWND,
    hwnd_pb: HWND,
    hwnd_text: HWND,
    hwnd_cancel: HWND,
    total: usize,
    current: usize,
    language: Language,
}

fn progress_text(language: Language, pct: usize) -> String {
    i18n::tr_f(language, "audiobook.progress", &[("pct", &pct.to_string())])
}

pub fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = crate::get_focus_safe();
        let cancel_btn = with_progress_state(hwnd, |s| s.hwnd_cancel).unwrap_or(HWND(0));
        if focus == cancel_btn {
            request_cancel(hwnd);
            return true;
        }
    }
    handle_accessibility(hwnd, msg)
}

pub fn open(parent: HWND, total: usize) -> HWND {
    unsafe {
        let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
        let class_name = to_wide(PROGRESS_CLASS_NAME);
        let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
        let title_w = to_wide(&i18n::tr(language, "audiobook.title"));

        let wc = WNDCLASSW {
            hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
                LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
            ),
            hInstance: hinstance,
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(progress_wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let hwnd = CreateWindowExW(
            WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            300,
            150,
            None,
            None,
            hinstance,
            Some(parent.0 as *const _),
        );

        if hwnd.0 != 0 {
            if with_progress_state(hwnd, |state| {
                SendMessageW(
                    state.hwnd_pb,
                    PBM_SETRANGE,
                    WPARAM(0),
                    LPARAM((total as isize) << 16),
                );
                state.total = total;
                state.current = 0;
            })
            .is_none()
            {
                crate::log_debug("Failed to access audiobook state");
            }

            // Center window relative to parent
            let mut rc_parent = RECT::default();
            let mut rc_dlg = RECT::default();
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::GetWindowRect(
                parent,
                &mut rc_parent
            ));
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::GetWindowRect(
                hwnd,
                &mut rc_dlg
            ));

            let dlg_w = rc_dlg.right - rc_dlg.left;
            let dlg_h = rc_dlg.bottom - rc_dlg.top;
            let parent_w = rc_parent.right - rc_parent.left;
            let parent_h = rc_parent.bottom - rc_parent.top;

            let x = rc_parent.left + (parent_w - dlg_w) / 2;
            let y = rc_parent.top + (parent_h - dlg_h) / 2;

            crate::log_if_err!(MoveWindow(hwnd, x, y, dlg_w, dlg_h, true));
            if !SetForegroundWindow(hwnd).as_bool() {
                crate::log_debug("Audiobook window: SetForegroundWindow failed");
            }
            if with_progress_state(hwnd, |state| {
                SetFocus(state.hwnd_text);
            })
            .is_none()
            {
                crate::log_debug("Failed to access audiobook state for initial focus");
            }
        }
        hwnd
    }
}

unsafe extern "system" fn progress_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "progress_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || progress_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn progress_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = unsafe { HWND((*create_struct).lpCreateParams as isize) };
            let language =
                { with_state(parent, |state| state.settings.language) }.unwrap_or_default();
            let label_text = progress_text(language, 0);
            let cancel_text = i18n::tr(language, "audiobook.cancel");

            let label = unsafe {
                CreateWindowExW(
                    Default::default(),
                    w!("EDIT"),
                    PCWSTR(to_wide(&label_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_CENTER | ES_READONLY),
                    20,
                    20,
                    240,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };

            let pb = unsafe {
                CreateWindowExW(
                    Default::default(),
                    w!("msctls_progress32"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE,
                    20,
                    50,
                    240,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };

            let hwnd_cancel = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&cancel_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    95,
                    80,
                    90,
                    28,
                    hwnd,
                    HMENU(PROGRESS_ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            let state = Box::new(ProgressDialogState {
                parent,
                hwnd_pb: pb,
                hwnd_text: label,
                hwnd_cancel,
                total: 0,
                current: 0,
                language,
            });
            crate::set_window_long_ptr_w_safe(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            if label.0 != 0 {
                crate::set_focus_safe(label);
            }
            LRESULT(0)
        }
        WM_SETFOCUS => {
            if with_progress_state(hwnd, |state| {
                crate::set_focus_safe(state.hwnd_text);
            })
            .is_none()
            {
                crate::log_debug("Failed to access audiobook state");
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            if cmd_id == PROGRESS_ID_CANCEL || cmd_id == 2 {
                // 2 is IDCANCEL
                request_cancel(hwnd);
                LRESULT(0)
            } else {
                crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
            }
        }
        WM_UPDATE_PROGRESS => {
            let requested = wparam.0;
            if with_progress_state(hwnd, |state| {
                let mut current = requested;
                if state.total > 0 {
                    current = current.min(state.total);
                }
                // Ignore stale/out-of-order progress messages that would move backwards.
                if current < state.current {
                    current = state.current;
                } else {
                    state.current = current;
                }

                crate::send_message_w_safe(state.hwnd_pb, PBM_SETPOS, WPARAM(current), LPARAM(0));
                if state.total > 0 {
                    let pct = ((current * 100) / state.total).min(100);
                    let text = progress_text(state.language, pct);
                    let wide = to_wide(&text);
                    if let Err(e) =
                        crate::set_window_text_w_safe(state.hwnd_text, PCWSTR(wide.as_ptr()))
                    {
                        crate::log_debug(&format!("Failed to set status text: {}", e));
                    }
                }
            })
            .is_none()
            {
                crate::log_debug("Failed to access audiobook state");
            }
            LRESULT(0)
        }
        WM_SET_PROGRESS_TOTAL => {
            let total = wparam.0.max(1);
            if with_progress_state(hwnd, |state| {
                if state.total != total {
                    state.total = total;
                    state.current = 0;
                    unsafe {
                        SendMessageW(
                            state.hwnd_pb,
                            PBM_SETRANGE,
                            WPARAM(0),
                            LPARAM((total as isize) << 16),
                        );
                        SendMessageW(state.hwnd_pb, PBM_SETPOS, WPARAM(0), LPARAM(0));
                    }
                }
            })
            .is_none()
            {
                crate::log_debug("Failed to access audiobook state for SET_TOTAL");
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            request_cancel(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let parent = with_progress_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                crate::set_foreground_window_safe(parent);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr =
                crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut ProgressDialogState;
            if !ptr.is_null() {
                let _unused_box = crate::box_from_raw_safe(ptr);
            }
            LRESULT(0)
        }
        _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
    }
}

pub fn request_cancel(hwnd: HWND) {
    let parent = with_progress_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    let already_cancelled = {
        with_state(parent, |state| {
            state
                .audiobook_cancel
                .as_ref()
                .map(|c| c.load(Ordering::Relaxed))
                .unwrap_or(false)
        })
    }
    .unwrap_or(false);

    if already_cancelled {
        return;
    }

    let language = { with_state(parent, |state| state.settings.language) }.unwrap_or_default();
    let msg = i18n::tr(language, "audiobook.cancel_confirm");
    let title = i18n::tr(language, "app.confirm_title");

    let msg_w = to_wide(&msg);
    let title_w = to_wide(&title);

    if crate::message_box_w_safe(
        hwnd,
        PCWSTR(msg_w.as_ptr()),
        PCWSTR(title_w.as_ptr()),
        MB_YESNO | MB_ICONWARNING,
    ) == IDYES
    {
        if {
            with_state(parent, |state| {
                if let Some(cancel) = &state.audiobook_cancel {
                    cancel.store(true, Ordering::Relaxed);
                }
                state.audiobook_progress = HWND(0);
            })
        }
        .is_none()
        {
            crate::log_debug("Failed to access audiobook state");
        }
        crate::log_if_err!(unsafe { windows::Win32::UI::WindowsAndMessaging::DestroyWindow(hwnd) });
    }
}

fn with_progress_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut ProgressDialogState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut ProgressDialogState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}
