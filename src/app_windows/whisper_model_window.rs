use crate::accessibility::to_wide;
use crate::settings::Language;
use crate::with_state;
use std::sync::{Arc, Mutex};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::UI::Controls::{WC_BUTTON, WC_COMBOBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, GetFocus, VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    DispatchMessageW, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW, IsDialogMessageW,
    LoadCursorW, MSG, PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, TranslateMessage, WINDOW_EX_STYLE, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CB_ADDSTRING, CB_GETCURSEL, CB_SETCURSEL, CBS_DROPDOWNLIST, GWLP_USERDATA,
};
use windows::core::PCWSTR;

const WHISPER_MODEL_DIALOG_CLASS: &str = "SonarpadWhisperModelDialog";
const ID_MODEL_COMBO: usize = 50101;
const ID_OK: usize = 50102;
const ID_CANCEL: usize = 50103;

struct WhisperModelDialogInit {
    parent: HWND,
    language: Language,
    result: Arc<Mutex<Option<String>>>,
}

struct WhisperModelDialogState {
    parent: HWND,
    combo: HWND,
    ok_button: HWND,
    result: Arc<Mutex<Option<String>>>,
}

unsafe extern "system" fn whisper_model_dialog_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "whisper_model_dialog_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || whisper_model_dialog_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn whisper_model_dialog_wndproc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create_struct = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*create_struct).lpCreateParams as *mut WhisperModelDialogInit;
                if init_ptr.is_null() {
                    return LRESULT(0);
                }
                let init = Box::from_raw(init_ptr);
                let hfont: HFONT = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

                let hint = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_STATIC,
                    PCWSTR(
                        to_wide(&crate::i18n::tr(init.language, "whisper.model.dialog_hint"))
                            .as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    16,
                    600,
                    70,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );

                let combo = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    16,
                    94,
                    600,
                    220,
                    hwnd,
                    HMENU(ID_MODEL_COMBO as isize),
                    HINSTANCE(0),
                    None,
                );

                let ok_button = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&crate::i18n::tr(init.language, "options.ok")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    432,
                    132,
                    88,
                    28,
                    hwnd,
                    HMENU(ID_OK as isize),
                    HINSTANCE(0),
                    None,
                );

                let cancel_button = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&crate::i18n::tr(init.language, "options.cancel")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    528,
                    132,
                    88,
                    28,
                    hwnd,
                    HMENU(ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                );

                for control in [hint, combo, ok_button, cancel_button] {
                    if control.0 != 0 && hfont.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                }

                let small = to_wide(&crate::i18n::tr(init.language, "whisper.model.item.fast"));
                let medium = to_wide(&crate::i18n::tr(
                    init.language,
                    "whisper.model.item.balanced",
                ));
                let large = to_wide(&crate::i18n::tr(init.language, "whisper.model.item.hq"));

                SendMessageW(
                    combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(small.as_ptr() as isize),
                );
                SendMessageW(
                    combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(medium.as_ptr() as isize),
                );
                SendMessageW(
                    combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(large.as_ptr() as isize),
                );
                SendMessageW(combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));

                let state = Box::new(WhisperModelDialogState {
                    parent: init.parent,
                    combo,
                    ok_button,
                    result: init.result,
                });
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                SetForegroundWindow(hwnd);
                crate::set_focus_safe(combo);
                LRESULT(0)
            }
            WM_COMMAND => {
                let cmd_id = wparam.0 & 0xffff;
                if cmd_id == ID_OK || cmd_id == 1 {
                    let ptr =
                        GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WhisperModelDialogState;
                    if !ptr.is_null() {
                        let state = &mut *ptr;
                        let sel = SendMessageW(state.combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        let selected = match sel {
                            0 => Some("small_q5_1".to_string()),
                            1 => Some("medium_q5_0".to_string()),
                            2 => Some("large_v3_turbo_q5_0".to_string()),
                            _ => None,
                        };
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) = selected;
                    }
                    if let Err(e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                        crate::log_debug(&format!("whisper model dialog close post failed: {e}"));
                    }
                    return LRESULT(0);
                }
                if cmd_id == ID_CANCEL || cmd_id == 2 {
                    let ptr =
                        GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WhisperModelDialogState;
                    if !ptr.is_null() {
                        let state = &mut *ptr;
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) = None;
                    }
                    if let Err(e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                        crate::log_debug(&format!("whisper model dialog close post failed: {e}"));
                    }
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    if let Err(e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(ID_CANCEL), LPARAM(0)) {
                        crate::log_debug(&format!("whisper model dialog cancel post failed: {e}"));
                    }
                    return LRESULT(0);
                }
                if wparam.0 as u32 == VK_RETURN.0 as u32 {
                    let ptr =
                        GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WhisperModelDialogState;
                    if !ptr.is_null() {
                        let state = &mut *ptr;
                        if GetFocus() == state.ok_button {
                            if let Err(e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(ID_OK), LPARAM(0))
                            {
                                crate::log_debug(&format!(
                                    "whisper model dialog ok post failed: {e}"
                                ));
                            }
                            return LRESULT(0);
                        }
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WhisperModelDialogState;
                if !ptr.is_null() {
                    let state = &mut *ptr;
                    EnableWindow(state.parent, true);
                    SetForegroundWindow(state.parent);
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WhisperModelDialogState;
                if !ptr.is_null() {
                    let _unused = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

pub fn choose_whisper_model(parent: HWND, language: Language) -> Option<String> {
    unsafe {
        let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
        let class_name = to_wide(WHISPER_MODEL_DIALOG_CLASS);
        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hInstance: hinstance,
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(whisper_model_dialog_wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let result = Arc::new(Mutex::new(None));
        let init = Box::new(WhisperModelDialogInit {
            parent,
            language,
            result: Arc::clone(&result),
        });

        let title = to_wide(&crate::i18n::tr(language, "whisper.model.dialog_title"));
        let hwnd = CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            650,
            210,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        );
        if hwnd.0 == 0 {
            return None;
        }

        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);

        let mut msg = MSG::default();
        loop {
            if !crate::is_window_handle_valid(hwnd) {
                break;
            }
            let gm = GetMessageW(&mut msg, HWND(0), 0, 0);
            if gm.0 == 0 || gm.0 == -1 {
                break;
            }
            if crate::app_windows::calendar_window::handle_reminder_alert_message(&msg) {
                continue;
            }
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        EnableWindow(parent, true);
        SetForegroundWindow(parent);

        result.lock().unwrap_or_else(|e| e.into_inner()).clone()
    }
}
