use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::UI::Controls::{
    HTREEITEM, TVGN_CARET, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVINSERTSTRUCTW_0, TVITEMW,
    TVM_ENSUREVISIBLE, TVM_GETITEMW, TVM_GETNEXTITEM, TVM_INSERTITEMW, TVM_SELECTITEM,
    TVS_HASBUTTONS, TVS_HASLINES, TVS_LINESATROOT, TVS_SHOWSELALWAYS, WC_BUTTON, WC_STATIC,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    DispatchMessageW, GWLP_USERDATA, HMENU, IDC_ARROW, IsDialogMessageW, LoadCursorW, MSG,
    PostMessageW, SendMessageW, SetForegroundWindow, TranslateMessage, WINDOW_STYLE, WM_CLOSE,
    WM_COMMAND, WM_CREATE, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::to_wide;
use crate::file_handler::EpubIndexEntry;
use crate::i18n;
use crate::settings::Language;
use crate::with_state;

const EPUB_INDEX_CLASS_NAME: &str = "SonarpadEpubIndex";
const EPUB_INDEX_ID_TREE: usize = 9651;
const EPUB_INDEX_ID_OK: usize = 9652;
const EPUB_INDEX_ID_CANCEL: usize = 9653;

struct EpubIndexInit {
    parent: HWND,
    entries: Vec<EpubIndexEntry>,
    language: Language,
    result: Arc<Mutex<Option<i32>>>,
}

struct EpubIndexState {
    tree: HWND,
    targets: Vec<i32>,
    result: Arc<Mutex<Option<i32>>>,
}

pub fn select_epub_index_entry(
    parent: HWND,
    entries: &[EpubIndexEntry],
    language: Language,
) -> Option<i32> {
    if entries.is_empty() {
        return None;
    }

    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
    let class_name = to_wide(EPUB_INDEX_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(epub_index_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    crate::register_class_w_safe(&wc);

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(EpubIndexInit {
        parent,
        entries: entries.to_vec(),
        language,
        result: result.clone(),
    });
    let title = to_wide(&i18n::tr(language, "epub_index.title"));
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            570,
            440,
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
        if !crate::is_window_handle_valid(hwnd) {
            break;
        }
        let read = crate::get_message_w_safe(&mut msg, HWND(0), 0, 0);
        if read.0 == 0 {
            break;
        }
        unsafe {
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                if let Err(error) =
                    PostMessageW(hwnd, WM_COMMAND, WPARAM(EPUB_INDEX_ID_CANCEL), LPARAM(0))
                {
                    crate::log_debug(&format!("Failed to post EPUB_INDEX_ID_CANCEL: {}", error));
                }
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let tree = with_epub_index_state(hwnd, |state| state.tree).unwrap_or(HWND(0));
                if GetFocus() == tree {
                    if let Err(error) =
                        PostMessageW(hwnd, WM_COMMAND, WPARAM(EPUB_INDEX_ID_OK), LPARAM(0))
                    {
                        crate::log_debug(&format!("Failed to post EPUB_INDEX_ID_OK: {}", error));
                    }
                    continue;
                }
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
    let guard = result.lock().unwrap_or_else(|error| error.into_inner());
    *guard
}

unsafe extern "system" fn epub_index_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "epub_index_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || epub_index_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn epub_index_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create_struct = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*create_struct).lpCreateParams as *mut EpubIndexInit;
                if init_ptr.is_null() {
                    return LRESULT(0);
                }
                let init = Box::from_raw(init_ptr);
                let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

                let hint = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&i18n::tr(init.language, "epub_index.hint")).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    14,
                    520,
                    24,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let tree = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("SysTreeView32"),
                    PCWSTR::null(),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WS_VSCROLL
                        | WINDOW_STYLE(
                            TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS,
                        ),
                    16,
                    44,
                    520,
                    300,
                    hwnd,
                    HMENU(EPUB_INDEX_ID_TREE as isize),
                    HINSTANCE(0),
                    None,
                );
                let ok = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "marker_select.ok")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    346,
                    356,
                    90,
                    28,
                    hwnd,
                    HMENU(EPUB_INDEX_ID_OK as isize),
                    HINSTANCE(0),
                    None,
                );
                let cancel = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "marker_select.cancel")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    446,
                    356,
                    90,
                    28,
                    hwnd,
                    HMENU(EPUB_INDEX_ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                );

                let mut targets = Vec::new();
                let mut first_item = HTREEITEM(0);
                insert_epub_index_entries(
                    tree,
                    HTREEITEM(0),
                    &init.entries,
                    &mut targets,
                    &mut first_item,
                );
                if first_item.0 != 0 {
                    SendMessageW(
                        tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(first_item.0),
                    );
                    SendMessageW(tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(first_item.0));
                }

                for control in [hint, tree, ok, cancel] {
                    if control.0 != 0 && hfont.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                }
                SetFocus(tree);

                let state = Box::new(EpubIndexState {
                    tree,
                    targets,
                    result: init.result.clone(),
                });
                crate::set_window_long_ptr_w_safe(
                    hwnd,
                    GWLP_USERDATA,
                    Box::into_raw(state) as isize,
                );
                LRESULT(0)
            }
            WM_COMMAND => match wparam.0 & 0xFFFF {
                EPUB_INDEX_ID_OK => {
                    if let Some(target) = selected_epub_target(hwnd) {
                        with_epub_index_state(hwnd, |state| {
                            *state
                                .result
                                .lock()
                                .unwrap_or_else(|error| error.into_inner()) = Some(target);
                        });
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    }
                    LRESULT(0)
                }
                EPUB_INDEX_ID_CANCEL => {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            },
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr =
                    crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut EpubIndexState;
                if !ptr.is_null() {
                    let _unused_state = Box::from_raw(ptr);
                    crate::set_window_long_ptr_w_safe(hwnd, GWLP_USERDATA, 0);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn insert_epub_index_entries(
    tree: HWND,
    parent: HTREEITEM,
    entries: &[EpubIndexEntry],
    targets: &mut Vec<i32>,
    first_item: &mut HTREEITEM,
) {
    for entry in entries {
        let target_index = targets.len();
        targets.push(entry.target_utf16);
        let item = insert_tree_item(tree, parent, &entry.title, target_index as isize);
        if first_item.0 == 0 && item.0 != 0 {
            *first_item = item;
        }
        if item.0 != 0 && !entry.children.is_empty() {
            insert_epub_index_entries(tree, item, &entry.children, targets, first_item);
        }
    }
}

fn insert_tree_item(tree: HWND, parent: HTREEITEM, label: &str, value_index: isize) -> HTREEITEM {
    let text = to_wide(label);
    let mut insert = TVINSERTSTRUCTW {
        hParent: parent,
        hInsertAfter: HTREEITEM(0xffff0002u32 as isize),
        Anonymous: TVINSERTSTRUCTW_0 {
            item: TVITEMW {
                mask: TVIF_TEXT | TVIF_PARAM,
                pszText: windows::core::PWSTR(text.as_ptr() as *mut _),
                lParam: LPARAM(value_index),
                ..Default::default()
            },
        },
    };
    HTREEITEM(
        crate::send_message_w_safe(
            tree,
            TVM_INSERTITEMW,
            WPARAM(0),
            LPARAM(&mut insert as *mut _ as isize),
        )
        .0,
    )
}

fn selected_epub_target(hwnd: HWND) -> Option<i32> {
    with_epub_index_state(hwnd, |state| {
        let caret = HTREEITEM(
            crate::send_message_w_safe(
                state.tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
            .0,
        );
        if caret.0 == 0 {
            return None;
        }
        let mut item = TVITEMW {
            mask: TVIF_PARAM,
            hItem: caret,
            ..Default::default()
        };
        if crate::send_message_w_safe(
            state.tree,
            TVM_GETITEMW,
            WPARAM(0),
            LPARAM(&mut item as *mut _ as isize),
        )
        .0 == 0
        {
            return None;
        }
        if item.lParam.0 < 0 {
            return None;
        }
        state.targets.get(item.lParam.0 as usize).copied()
    })
    .flatten()
}

fn with_epub_index_state<F, R>(hwnd: HWND, callback: F) -> Option<R>
where
    F: FnOnce(&mut EpubIndexState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut EpubIndexState;
    crate::with_raw_mut_ptr_safe(ptr, callback)
}
