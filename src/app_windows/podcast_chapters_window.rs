use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::WC_BUTTON;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN, VK_SPACE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW,
    IsDialogMessageW, IsWindow, LB_ADDSTRING, LB_GETCURSEL, LB_RESETCONTENT, LB_SETCURSEL,
    LBS_HASSTRINGS, LBS_NOINTEGRALHEIGHT, LBS_NOTIFY, LoadCursorW, MSG, PostMessageW,
    RegisterClassW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, TranslateMessage,
    WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY,
    WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::to_wide;
use crate::i18n;
use crate::podcast::chapters::Chapter;
use crate::settings::Language;
use crate::with_state;

const CHAPTER_LIST_CLASS_NAME: &str = "NovapadChapterList";
const CHAPTER_LIST_ID_LIST: usize = 9201;
const CHAPTER_LIST_ID_OK: usize = 9202;
const CHAPTER_LIST_ID_CANCEL: usize = 9203;

struct ChapterListInit {
    parent: HWND,
    items: Vec<String>,
    language: Language,
    result: Arc<Mutex<Option<usize>>>,
}

struct ChapterListState {
    parent: HWND,
    list: HWND,
    ok: HWND,
    result: Arc<Mutex<Option<usize>>>,
}

pub fn select_chapter(parent: HWND, chapters: &[Chapter], language: Language) -> Option<usize> {
    if chapters.is_empty() {
        return None;
    }
    let items: Vec<String> = chapters
        .iter()
        .map(crate::podcast::chapters::chapter_label)
        .collect();

    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let class_name = to_wide(CHAPTER_LIST_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(chapter_list_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(ChapterListInit {
        parent,
        items,
        language,
        result: result.clone(),
    });
    let title = to_wide(&i18n::tr(language, "podcasts.chapters.title"));

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            440,
            320,
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
        if unsafe { IsWindow(hwnd).as_bool() } == false {
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
                    WPARAM(CHAPTER_LIST_ID_CANCEL as usize),
                    LPARAM(0),
                );
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let (list, ok) = with_chapter_state(hwnd, |state| (state.list, state.ok))
                    .unwrap_or((HWND(0), HWND(0)));
                let focus = GetFocus();
                if focus == list || focus == ok {
                    let _ = PostMessageW(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(CHAPTER_LIST_ID_OK as usize),
                        LPARAM(0),
                    );
                    continue;
                }
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_SPACE.0 as u32 {
                let (list, ok) = with_chapter_state(hwnd, |state| (state.list, state.ok))
                    .unwrap_or((HWND(0), HWND(0)));
                let focus = GetFocus();
                if focus == list || focus == ok {
                    let _ = PostMessageW(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(CHAPTER_LIST_ID_OK as usize),
                        LPARAM(0),
                    );
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

unsafe extern "system" fn chapter_list_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*create_struct).lpCreateParams as *mut ChapterListInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = Box::from_raw(init_ptr);
            let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

            let list = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("LISTBOX"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WINDOW_STYLE((LBS_NOTIFY | LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT) as u32),
                16,
                16,
                392,
                220,
                hwnd,
                HMENU(CHAPTER_LIST_ID_LIST as isize),
                HINSTANCE(0),
                None,
            );

            let _ = SendMessageW(list, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
            for item in init.items.iter() {
                let _ = SendMessageW(
                    list,
                    LB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(item).as_ptr() as isize),
                );
            }
            let _ = SendMessageW(list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
            let _ = SetFocus(list);

            let ok = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(init.language, "marker_select.ok")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                220,
                248,
                90,
                28,
                hwnd,
                HMENU(CHAPTER_LIST_ID_OK as isize),
                HINSTANCE(0),
                None,
            );

            let cancel = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(init.language, "marker_select.cancel")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                318,
                248,
                90,
                28,
                hwnd,
                HMENU(CHAPTER_LIST_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            for control in [list, ok, cancel] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let state = Box::new(ChapterListState {
                parent: init.parent,
                list,
                ok,
                result: init.result,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            match cmd_id {
                CHAPTER_LIST_ID_OK => {
                    let (list, result) =
                        with_chapter_state(hwnd, |state| (state.list, state.result.clone()))
                            .unwrap_or((HWND(0), Arc::new(Mutex::new(None))));
                    if list.0 != 0 {
                        let sel = SendMessageW(list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                        if sel >= 0 {
                            let _ = result.lock().map(|mut r| *r = Some(sel as usize));
                        }
                    }
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                CHAPTER_LIST_ID_CANCEL => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let parent = with_chapter_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                SetForegroundWindow(parent);
            }
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ChapterListState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => DefWindowProcW(hwnd, msg, wparam, lparam),
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_chapter_state<R>(
    hwnd: HWND,
    f: impl FnOnce(&mut ChapterListState) -> R,
) -> Option<R> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ChapterListState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}
