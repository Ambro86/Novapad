use std::sync::{Arc, Mutex};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::WC_BUTTON;
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus, VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW,
    IsDialogMessageW, IsWindow, LB_ADDSTRING, LB_GETCURSEL, LB_GETTEXT, LB_GETTEXTLEN,
    LB_SETCURSEL, LBS_NOTIFY, LoadCursorW, MSG, PostMessageW, RegisterClassW, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::Language;
use crate::with_state;

const INTERPRETER_SELECT_CLASS_NAME: &str = "SonarpadInterpreterSelect";
const ID_LIST: usize = 9201;
const ID_OK: usize = 9202;
const ID_CANCEL: usize = 9203;

struct InterpreterSelectInit {
    parent: HWND,
    items: Vec<String>,
    language: Language,
    result: Arc<Mutex<Option<String>>>,
}

struct InterpreterSelectState {
    list: HWND,
    result: Arc<Mutex<Option<String>>>,
}

pub fn select_interpreter(parent: HWND, items: Vec<String>, language: Language) -> Option<String> {
    if items.is_empty() {
        return None;
    }

    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let class_name = to_wide(INTERPRETER_SELECT_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(interpreter_select_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(InterpreterSelectInit {
        parent,
        items,
        language,
        result: result.clone(),
    });
    let title = to_wide(&i18n::tr(language, "options.interpreter_search.title"));

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            500,
            300,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };

    if hwnd.0 == 0 {
        return None;
    }

    unsafe {
        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);
    }

    let mut msg = MSG::default();
    loop {
        if !unsafe { IsWindow(hwnd).as_bool() } {
            break;
        }
        let res = unsafe { GetMessageW(&mut msg, HWND(0), 0, 0) };
        if res.0 == 0 {
            break;
        }
        unsafe {
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                crate::log_if_err!(PostMessageW(hwnd, WM_COMMAND, WPARAM(ID_OK), LPARAM(0)));
                continue;
            }
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    unsafe {
        EnableWindow(parent, true);
        SetForegroundWindow(parent);
    }

    let guard = result.lock().unwrap_or_else(|e| e.into_inner());
    guard.clone()
}

unsafe extern "system" fn interpreter_select_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "interpreter_select_wndproc",
        || DefWindowProcW(hwnd, msg, wparam, lparam),
        || unsafe { interpreter_select_wndproc_inner(hwnd, msg, wparam, lparam) },
    )
}

unsafe fn interpreter_select_wndproc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*create_struct).lpCreateParams as *mut InterpreterSelectInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = unsafe { Box::from_raw(init_ptr) };
            let parent = init.parent;
            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

            let list = CreateWindowExW(
                Default::default(),
                w!("LISTBOX"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | WINDOW_STYLE(LBS_NOTIFY as u32),
                10,
                10,
                460,
                200,
                hwnd,
                HMENU(ID_LIST as isize),
                HINSTANCE(0),
                None,
            );

            for item in init.items.iter() {
                SendMessageW(
                    list,
                    LB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(item).as_ptr() as isize),
                );
            }
            SendMessageW(list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
            SetFocus(list);

            let ok = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(init.language, "options.ok")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                280,
                220,
                90,
                28,
                hwnd,
                HMENU(ID_OK as isize),
                HINSTANCE(0),
                None,
            );

            let cancel = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(init.language, "options.cancel")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                380,
                220,
                90,
                28,
                hwnd,
                HMENU(ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            if hfont.0 != 0 {
                SendMessageW(list, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                SendMessageW(ok, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                SendMessageW(cancel, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }

            let state = Box::new(InterpreterSelectState {
                list,
                result: init.result.clone(),
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = wparam.0;
            match id {
                ID_OK => {
                    with_interpreter_state(hwnd, |state| {
                        let sel = SendMessageW(state.list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        if sel >= 0 {
                            let len = SendMessageW(
                                state.list,
                                LB_GETTEXTLEN,
                                WPARAM(sel as usize),
                                LPARAM(0),
                            )
                            .0;
                            if len >= 0 {
                                let mut buf = vec![0u16; (len + 1) as usize];
                                SendMessageW(
                                    state.list,
                                    LB_GETTEXT,
                                    WPARAM(sel as usize),
                                    LPARAM(buf.as_mut_ptr() as isize),
                                );
                                let path = String::from_utf16_lossy(&buf[..len as usize]);
                                *state.result.lock().unwrap_or_else(|e| e.into_inner()) =
                                    Some(path);
                            }
                        }
                    });
                    crate::log_if_err!(DestroyWindow(hwnd));
                    LRESULT(0)
                }
                ID_CANCEL => {
                    crate::log_if_err!(DestroyWindow(hwnd));
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            crate::log_if_err!(DestroyWindow(hwnd));
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut InterpreterSelectState;
            if !ptr.is_null() {
                let _unused_box = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_interpreter_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut InterpreterSelectState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut InterpreterSelectState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}
