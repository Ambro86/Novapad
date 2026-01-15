use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{WC_BUTTON, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW,
    IsDialogMessageW, IsWindow, LB_ADDSTRING, LB_GETCOUNT, LB_GETSEL, LB_SETCARETINDEX,
    LB_SETCURSEL, LB_SETSEL, LB_SETTOPINDEX, LBN_SELCHANGE, LBS_MULTIPLESEL, LBS_NOINTEGRALHEIGHT,
    LBS_NOTIFY, LoadCursorW, MSG, PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, SetWindowTextW, TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::Language;
use crate::with_state;

const MARKER_SELECT_CLASS_NAME: &str = "NovapadMarkerSelect";
const MARKER_ID_LIST: usize = 9101;
const MARKER_ID_TOGGLE_ALL: usize = 9102;
const MARKER_ID_OK: usize = 9103;
const MARKER_ID_CANCEL: usize = 9104;

struct MarkerSelectInit {
    parent: HWND,
    items: Vec<String>,
    language: Language,
    result: Arc<Mutex<Option<Vec<usize>>>>,
}

struct MarkerSelectState {
    parent: HWND,
    list: HWND,
    toggle_all: HWND,
    result: Arc<Mutex<Option<Vec<usize>>>>,
}

struct MarkerSelectLabels {
    title: String,
    hint: String,
    toggle_all: String,
    ok: String,
    cancel: String,
}

fn labels(language: Language) -> MarkerSelectLabels {
    MarkerSelectLabels {
        title: i18n::tr(language, "marker_select.title"),
        hint: i18n::tr(language, "marker_select.hint"),
        toggle_all: i18n::tr(language, "marker_select.toggle_all"),
        ok: i18n::tr(language, "marker_select.ok"),
        cancel: i18n::tr(language, "marker_select.cancel"),
    }
}

pub fn select_marker_entries(
    parent: HWND,
    items: &[String],
    language: Language,
) -> Option<Vec<usize>> {
    if items.is_empty() {
        return Some(Vec::new());
    }

    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let class_name = to_wide(MARKER_SELECT_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(marker_select_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(MarkerSelectInit {
        parent,
        items: items.to_vec(),
        language,
        result: result.clone(),
    });
    let labels = labels(language);
    let title = to_wide(&labels.title);

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            460,
            340,
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
                let _ = PostMessageW(
                    hwnd,
                    WM_COMMAND,
                    WPARAM(MARKER_ID_CANCEL as usize),
                    LPARAM(0),
                );
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let list = with_marker_state(hwnd, |state| state.list).unwrap_or(HWND(0));
                if GetFocus() == list {
                    let _ =
                        PostMessageW(hwnd, WM_COMMAND, WPARAM(MARKER_ID_OK as usize), LPARAM(0));
                    continue;
                }
            }
            if IsDialogMessageW(hwnd, &mut msg).as_bool() {
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

    result.lock().unwrap().clone()
}

unsafe extern "system" fn marker_select_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*create_struct).lpCreateParams as *mut MarkerSelectInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = Box::from_raw(init_ptr);
            let labels = labels(init.language);

            let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

            let hint = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.hint).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                14,
                420,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let list = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("LISTBOX"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WINDOW_STYLE((LBS_NOTIFY | LBS_MULTIPLESEL | LBS_NOINTEGRALHEIGHT) as u32),
                16,
                40,
                420,
                220,
                hwnd,
                HMENU(MARKER_ID_LIST as isize),
                HINSTANCE(0),
                None,
            );

            for item in init.items.iter() {
                let _ = SendMessageW(
                    list,
                    LB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(item).as_ptr() as isize),
                );
            }
            let count = SendMessageW(list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0;
            for idx in 0..count {
                let _ = SendMessageW(list, LB_SETSEL, WPARAM(1), LPARAM(idx));
            }
            if count > 0 {
                let _ = SendMessageW(list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
                let _ = SendMessageW(list, LB_SETCARETINDEX, WPARAM(0), LPARAM(0));
                let _ = SendMessageW(list, LB_SETTOPINDEX, WPARAM(0), LPARAM(0));
            }
            let _ = SetFocus(list);

            let toggle_all = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.toggle_all).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                16,
                270,
                140,
                28,
                hwnd,
                HMENU(MARKER_ID_TOGGLE_ALL as isize),
                HINSTANCE(0),
                None,
            );

            let ok = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.ok).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                250,
                270,
                90,
                28,
                hwnd,
                HMENU(MARKER_ID_OK as isize),
                HINSTANCE(0),
                None,
            );

            let cancel = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.cancel).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                346,
                270,
                90,
                28,
                hwnd,
                HMENU(MARKER_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            for control in [hint, list, toggle_all, ok, cancel] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let state = Box::new(MarkerSelectState {
                parent: init.parent,
                list,
                toggle_all,
                result: init.result.clone(),
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            update_toggle_all_label(toggle_all, init.language, list);

            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            let notification = ((wparam.0 >> 16) & 0xffff) as u16;
            if cmd_id == MARKER_ID_LIST && notification as u32 == LBN_SELCHANGE {
                let _ = with_marker_state(hwnd, |state| {
                    let language =
                        with_state(state.parent, |s| s.settings.language).unwrap_or_default();
                    update_toggle_all_label(state.toggle_all, language, state.list);
                });
                LRESULT(0)
            } else if cmd_id == MARKER_ID_TOGGLE_ALL {
                let _ = with_marker_state(hwnd, |state| {
                    let should_select_all = list_has_unselected(state.list);
                    set_all_selected(state.list, should_select_all);
                    let language =
                        with_state(state.parent, |s| s.settings.language).unwrap_or_default();
                    update_toggle_all_label(state.toggle_all, language, state.list);
                });
                LRESULT(0)
            } else if cmd_id == MARKER_ID_OK {
                let _ = with_marker_state(hwnd, |state| {
                    let count = SendMessageW(state.list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0;
                    let mut selected = Vec::new();
                    for idx in 0..count {
                        let is_selected =
                            SendMessageW(state.list, LB_GETSEL, WPARAM(idx as usize), LPARAM(0)).0;
                        if is_selected > 0 {
                            selected.push(idx as usize);
                        }
                    }
                    *state.result.lock().unwrap() = Some(selected);
                });
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                LRESULT(0)
            } else if cmd_id == MARKER_ID_CANCEL {
                let _ = with_marker_state(hwnd, |state| {
                    *state.result.lock().unwrap() = None;
                });
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                LRESULT(0)
            } else {
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(
                    hwnd,
                    WM_COMMAND,
                    WPARAM(MARKER_ID_CANCEL as usize),
                    LPARAM(0),
                );
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let list = with_marker_state(hwnd, |state| state.list).unwrap_or(HWND(0));
                if focus == list {
                    let _ =
                        PostMessageW(hwnd, WM_COMMAND, WPARAM(MARKER_ID_OK as usize), LPARAM(0));
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            let _ = with_marker_state(hwnd, |state| {
                if state.result.lock().unwrap().is_none() {
                    *state.result.lock().unwrap() = None;
                }
            });
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let _ = with_marker_state(hwnd, |state| {
                EnableWindow(state.parent, true);
                SetForegroundWindow(state.parent);
            });
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut MarkerSelectState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_marker_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut MarkerSelectState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut MarkerSelectState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}
fn list_has_unselected(list: HWND) -> bool {
    let count = unsafe { SendMessageW(list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 };
    for idx in 0..count {
        let is_selected =
            unsafe { SendMessageW(list, LB_GETSEL, WPARAM(idx as usize), LPARAM(0)).0 };
        if is_selected == 0 {
            return true;
        }
    }
    false
}

fn set_all_selected(list: HWND, selected: bool) {
    let count = unsafe { SendMessageW(list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 };
    for idx in 0..count {
        let _ = unsafe {
            SendMessageW(
                list,
                LB_SETSEL,
                WPARAM(if selected { 1 } else { 0 }),
                LPARAM(idx),
            )
        };
    }
}

fn update_toggle_all_label(button: HWND, language: Language, list: HWND) {
    let label = if list_has_unselected(list) {
        i18n::tr(language, "marker_select.select_all")
    } else {
        i18n::tr(language, "marker_select.deselect_all")
    };
    let wide = to_wide(&label);
    unsafe {
        let _ = SetWindowTextW(button, PCWSTR(wide.as_ptr()));
    }
}
