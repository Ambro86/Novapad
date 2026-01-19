use crate::accessibility::{EM_SCROLLCARET, handle_accessibility, to_wide};
use crate::audio_player::start_audiobook_at;
use crate::i18n;
use crate::settings::FileFormat;
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::RichEdit::CHARRANGE;
use windows::Win32::UI::Controls::{WC_BUTTON, WC_LISTBOXW};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, GetFocus, SetFocus, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    GWLP_USERDATA, GetWindowLongPtrW, HMENU, IDC_ARROW, IDCANCEL, LB_ADDSTRING, LB_GETCOUNT,
    LB_GETCURSEL, LB_RESETCONTENT, LB_SETCURSEL, LBN_DBLCLK, LBS_HASSTRINGS, LBS_NOTIFY,
    LoadCursorW, MSG, RegisterClassW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW,
    WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY,
    WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const BOOKMARKS_CLASS_NAME: &str = "NovapadBookmarks";
const BOOKMARKS_ID_LIST: usize = 9001;
const BOOKMARKS_ID_DELETE: usize = 9002;
const BOOKMARKS_ID_GOTO: usize = 9003;
const BOOKMARKS_ID_OK: usize = 9004;

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = GetFocus();
        let (list, btn) = with_bookmarks_state(hwnd, |s| (s.hwnd_list, s.hwnd_goto))
            .unwrap_or((HWND(0), HWND(0)));
        if focus == list || focus == btn {
            goto_selected(hwnd);
            return true;
        }
    }
    handle_accessibility(hwnd, msg)
}

struct BookmarksWindowState {
    parent: HWND,
    hwnd_list: HWND,
    hwnd_goto: HWND,
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.bookmarks_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(BOOKMARKS_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(bookmarks_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let title = to_wide(&i18n::tr(language, "bookmarks.title"));

    let window = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        400,
        450,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if window.0 != 0 {
        if with_state(parent, |state| {
            state.bookmarks_window = window;
        })
        .is_none()
        {
            crate::log_debug("Failed to access bookmarks state");
        }
        EnableWindow(parent, false);
        SetForegroundWindow(window);
    }
}

unsafe extern "system" fn bookmarks_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();

            let hwnd_list = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_LISTBOXW,
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_VSCROLL
                    | WS_TABSTOP
                    | WINDOW_STYLE((LBS_NOTIFY | LBS_HASSTRINGS) as u32),
                10,
                10,
                360,
                300,
                hwnd,
                HMENU(BOOKMARKS_ID_LIST as isize),
                HINSTANCE(0),
                None,
            );

            let btn_goto_text = i18n::tr(language, "bookmarks.goto");
            let hwnd_goto = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&btn_goto_text).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                10,
                320,
                110,
                30,
                hwnd,
                HMENU(BOOKMARKS_ID_GOTO as isize),
                HINSTANCE(0),
                None,
            );

            let btn_del_text = i18n::tr(language, "bookmarks.delete");
            let hwnd_delete = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&btn_del_text).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                130,
                320,
                110,
                30,
                hwnd,
                HMENU(BOOKMARKS_ID_DELETE as isize),
                HINSTANCE(0),
                None,
            );

            let hwnd_ok = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                w!("OK"),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                250,
                320,
                110,
                30,
                hwnd,
                HMENU(BOOKMARKS_ID_OK as isize),
                HINSTANCE(0),
                None,
            );

            for ctrl in [hwnd_list, hwnd_goto, hwnd_delete, hwnd_ok] {
                if ctrl.0 != 0 && hfont.0 != 0 {
                    SendMessageW(ctrl, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let state = Box::new(BookmarksWindowState {
                parent,
                hwnd_list,
                hwnd_goto,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            refresh_bookmarks_list(hwnd);

            if SendMessageW(hwnd_list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 > 0 {
                SendMessageW(hwnd_list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
            }
            SetFocus(hwnd_list);

            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            let notify = (wparam.0 >> 16) as u16;
            match cmd_id {
                BOOKMARKS_ID_GOTO => {
                    goto_selected(hwnd);
                    LRESULT(0)
                }
                BOOKMARKS_ID_DELETE => {
                    delete_selected(hwnd);
                    LRESULT(0)
                }
                BOOKMARKS_ID_OK => {
                    crate::log_if_err!(DestroyWindow(hwnd));
                    LRESULT(0)
                }
                BOOKMARKS_ID_LIST if notify == LBN_DBLCLK as u16 => {
                    goto_selected(hwnd);
                    LRESULT(0)
                }
                cmd if cmd == IDCANCEL.0 as usize || cmd == 2 => {
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
        WM_DESTROY => {
            let parent = with_bookmarks_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                EnableWindow(parent, true);
                SetForegroundWindow(parent);
                if with_state(parent, |state| {
                    state.bookmarks_window = HWND(0);
                })
                .is_none()
                {
                    crate::log_debug("Failed to access bookmarks state");
                }
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut BookmarksWindowState;
            if !ptr.is_null() {
                drop(Box::from_raw(ptr));
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_bookmarks_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut BookmarksWindowState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut BookmarksWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

pub unsafe fn refresh_bookmarks_list(hwnd: HWND) {
    let (parent, hwnd_list) = match with_bookmarks_state(hwnd, |s| (s.parent, s.hwnd_list)) {
        Some(v) => v,
        None => return,
    };

    let path = with_state(parent, |state| {
        state.docs.get(state.current).and_then(|d| d.path.clone())
    })
    .flatten();

    let Some(path) = path else {
        return;
    };
    let path_str = path.to_string_lossy().to_string();

    SendMessageW(hwnd_list, LB_RESETCONTENT, WPARAM(0), LPARAM(0));

    if with_state(parent, |state| {
        if let Some(list) = state.bookmarks.files.get(&path_str) {
            for bm in list {
                let text = format!("[{}] {}", bm.timestamp, bm.snippet);
                let wide = to_wide(&text);
                SendMessageW(
                    hwnd_list,
                    LB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(wide.as_ptr() as isize),
                );
            }
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access bookmarks state");
    }

    let count = SendMessageW(hwnd_list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0 as i32;
    if count > 0 && SendMessageW(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32 == -1 {
        SendMessageW(hwnd_list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
    }
}

pub unsafe fn goto_selected(hwnd: HWND) {
    let (parent, hwnd_list) = match with_bookmarks_state(hwnd, |s| (s.parent, s.hwnd_list)) {
        Some(v) => v,
        None => return,
    };

    let sel = SendMessageW(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
    if sel < 0 {
        return;
    }

    let res: Option<(std::path::PathBuf, HWND, FileFormat)> = with_state(parent, |state| {
        state
            .docs
            .get(state.current)
            .and_then(|d| d.path.clone().map(|p| (p, d.hwnd_edit, d.format)))
    })
    .flatten();

    let Some((path, hwnd_edit, format)) = res else {
        return;
    };

    let path_str = path.to_string_lossy().to_string();

    if with_state(parent, |state| {
        if let Some(list) = state.bookmarks.files.get(&path_str)
            && let Some(bm) = list.get(sel as usize)
        {
            if matches!(format, FileFormat::Audiobook) {
                unsafe {
                    start_audiobook_at(parent, &path, bm.position as u64);
                }
            } else {
                let mut cr = CHARRANGE {
                    cpMin: bm.position,
                    cpMax: bm.position,
                };
                unsafe {
                    SendMessageW(
                        hwnd_edit,
                        crate::accessibility::EM_EXSETSEL,
                        WPARAM(0),
                        LPARAM(&mut cr as *mut _ as isize),
                    );
                    SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
                }
            }
            unsafe {
                SetFocus(hwnd_edit);
            }
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access bookmarks state");
    }
    crate::log_if_err!(DestroyWindow(hwnd));
}

pub unsafe fn delete_selected(hwnd: HWND) {
    let (parent, hwnd_list) = match with_bookmarks_state(hwnd, |s| (s.parent, s.hwnd_list)) {
        Some(v) => v,
        None => return,
    };

    let sel = SendMessageW(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
    if sel < 0 {
        return;
    }

    let path = with_state(parent, |state| {
        state.docs.get(state.current).and_then(|d| d.path.clone())
    })
    .flatten();

    let Some(path) = path else {
        return;
    };
    let path_str = path.to_string_lossy().to_string();

    if with_state(parent, |state| {
        if let Some(list) = state.bookmarks.files.get_mut(&path_str)
            && sel < list.len() as i32
        {
            list.remove(sel as usize);
            crate::bookmarks::save_bookmarks(&state.bookmarks);
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access bookmarks state");
    }
    refresh_bookmarks_list(hwnd);
}
