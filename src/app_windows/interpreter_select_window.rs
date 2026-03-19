use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::UI::Controls::{
    HTREEITEM, TVE_EXPAND, TVGN_CARET, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVINSERTSTRUCTW_0,
    TVITEMW, TVM_ENSUREVISIBLE, TVM_EXPAND, TVM_GETITEMW, TVM_GETNEXTITEM, TVM_INSERTITEMW,
    TVM_SELECTITEM, TVS_HASBUTTONS, TVS_HASLINES, TVS_LINESATROOT, TVS_SHOWSELALWAYS, WC_BUTTON,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, SetFocus, VK_DOWN, VK_END, VK_ESCAPE, VK_HOME, VK_LEFT, VK_NEXT, VK_PRIOR,
    VK_RETURN, VK_RIGHT, VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    DispatchMessageW, GWLP_USERDATA, HMENU, HWND_TOPMOST, IDC_ARROW, IsDialogMessageW,
    LB_ADDSTRING, LB_GETCURSEL, LB_GETTEXT, LB_GETTEXTLEN, LB_SETCURSEL, LBS_NOTIFY, LoadCursorW,
    MSG, PostMessageW, SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW, SendMessageW, SetForegroundWindow,
    SetWindowPos, TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_KEYDOWN,
    WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
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
const ID_SECONDARY: usize = 9204;

#[derive(Clone)]
pub(crate) struct GroupedSelectItem {
    pub(crate) label: String,
    pub(crate) value: String,
}

#[derive(Clone)]
pub(crate) struct GroupedSelectGroup {
    pub(crate) label: String,
    pub(crate) items: Vec<GroupedSelectItem>,
}

#[derive(Clone)]
pub(crate) enum InterpreterSelectionResult {
    Item(String),
    SecondaryAction,
}

enum InterpreterDialogInitMode {
    List(Vec<String>),
    Tree(Vec<GroupedSelectGroup>),
}

struct InterpreterSelectInit {
    parent: HWND,
    mode: InterpreterDialogInitMode,
    language: Language,
    secondary_action_label: Option<String>,
    initial_tree_value: Option<String>,
    result: Arc<Mutex<Option<InterpreterSelectionResult>>>,
}

#[derive(Default)]
struct InterpreterSelectOptions {
    suppress_parent_restore_on_accept: bool,
    pin_topmost: bool,
    secondary_action_label: Option<String>,
    initial_tree_value: Option<String>,
}

enum ControlKind {
    List(HWND),
    Tree(HWND),
}

struct InterpreterSelectState {
    control: ControlKind,
    tree_values: Vec<String>,
    result: Arc<Mutex<Option<InterpreterSelectionResult>>>,
}

pub fn select_interpreter(
    parent: HWND,
    items: Vec<String>,
    language: Language,
    title: String,
) -> Option<String> {
    match select_interpreter_internal(
        parent,
        InterpreterDialogInitMode::List(items),
        language,
        title,
        InterpreterSelectOptions::default(),
    ) {
        Some(InterpreterSelectionResult::Item(value)) => Some(value),
        _ => None,
    }
}

pub fn select_interpreter_without_parent_restore_on_accept(
    parent: HWND,
    items: Vec<String>,
    language: Language,
    title: String,
) -> Option<String> {
    match select_interpreter_internal(
        parent,
        InterpreterDialogInitMode::List(items),
        language,
        title,
        InterpreterSelectOptions {
            suppress_parent_restore_on_accept: true,
            ..Default::default()
        },
    ) {
        Some(InterpreterSelectionResult::Item(value)) => Some(value),
        _ => None,
    }
}

pub fn select_interpreter_with_secondary_action(
    parent: HWND,
    items: Vec<String>,
    language: Language,
    title: String,
    secondary_action_label: String,
) -> Option<InterpreterSelectionResult> {
    select_interpreter_internal(
        parent,
        InterpreterDialogInitMode::List(items),
        language,
        title,
        InterpreterSelectOptions {
            secondary_action_label: Some(secondary_action_label),
            ..Default::default()
        },
    )
}

pub fn select_grouped_interpreter(
    parent: HWND,
    groups: Vec<GroupedSelectGroup>,
    language: Language,
    title: String,
    initial_value: Option<String>,
) -> Option<String> {
    match select_interpreter_internal(
        parent,
        InterpreterDialogInitMode::Tree(groups),
        language,
        title,
        InterpreterSelectOptions {
            initial_tree_value: initial_value,
            ..Default::default()
        },
    ) {
        Some(InterpreterSelectionResult::Item(value)) => Some(value),
        _ => None,
    }
}

fn select_interpreter_internal(
    parent: HWND,
    mode: InterpreterDialogInitMode,
    language: Language,
    title: String,
    options: InterpreterSelectOptions,
) -> Option<InterpreterSelectionResult> {
    match &mode {
        InterpreterDialogInitMode::List(items) if items.is_empty() => return None,
        InterpreterDialogInitMode::Tree(groups) if groups.is_empty() => return None,
        _ => {}
    }

    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
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
    crate::register_class_w_safe(&wc);

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(InterpreterSelectInit {
        parent,
        mode,
        language,
        secondary_action_label: options.secondary_action_label,
        initial_tree_value: options.initial_tree_value,
        result: result.clone(),
    });
    let title = to_wide(&title);

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            620,
            360,
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
        if options.pin_topmost
            && let Err(err) = SetWindowPos(
                hwnd,
                HWND_TOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
            )
        {
            crate::log_debug(&format!(
                "Failed to pin interpreter selection window {:?} as topmost: {}",
                hwnd, err
            ));
        }
    }

    let mut msg = MSG::default();
    loop {
        if !crate::is_window_handle_valid(hwnd) {
            break;
        }
        let res = crate::get_message_w_safe(&mut msg, HWND(0), 0, 0);
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
            let focused = crate::get_focus_safe();
            let control_has_navigation_key =
                with_interpreter_state(hwnd, |state| match state.control {
                    ControlKind::List(list) => {
                        focused == list && is_list_navigation_key(msg.wParam.0 as u32)
                    }
                    ControlKind::Tree(tree) => {
                        focused == tree && is_tree_navigation_key(msg.wParam.0 as u32)
                    }
                })
                .unwrap_or(false);
            if control_has_navigation_key {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
                continue;
            }
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    let result_value = result.lock().unwrap_or_else(|e| e.into_inner()).clone();
    unsafe {
        if !(options.suppress_parent_restore_on_accept && result_value.is_some()) {
            EnableWindow(parent, true);
        }
        if result_value.is_none() {
            SetForegroundWindow(parent);
        }
    }

    result_value
}

unsafe extern "system" fn interpreter_select_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "interpreter_select_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || interpreter_select_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn interpreter_select_wndproc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = unsafe { (*create_struct).lpCreateParams as *mut InterpreterSelectInit };
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = crate::box_from_raw_safe(init_ptr);
            let parent = init.parent;
            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

            let (control, tree_values) = match init.mode {
                InterpreterDialogInitMode::List(items) => {
                    let list = unsafe {
                        CreateWindowExW(
                            Default::default(),
                            w!("LISTBOX"),
                            PCWSTR::null(),
                            WS_CHILD
                                | WS_VISIBLE
                                | WS_TABSTOP
                                | WS_VSCROLL
                                | WINDOW_STYLE(LBS_NOTIFY as u32),
                            10,
                            10,
                            580,
                            250,
                            hwnd,
                            HMENU(ID_LIST as isize),
                            HINSTANCE(0),
                            None,
                        )
                    };

                    for item in items {
                        crate::send_message_w_safe(
                            list,
                            LB_ADDSTRING,
                            WPARAM(0),
                            LPARAM(to_wide(&item).as_ptr() as isize),
                        );
                    }
                    unsafe {
                        SendMessageW(list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
                        SetFocus(list);
                    }
                    (ControlKind::List(list), Vec::new())
                }
                InterpreterDialogInitMode::Tree(groups) => {
                    let tree = unsafe {
                        CreateWindowExW(
                            Default::default(),
                            w!("SysTreeView32"),
                            PCWSTR::null(),
                            WS_CHILD
                                | WS_VISIBLE
                                | WS_TABSTOP
                                | WS_VSCROLL
                                | WINDOW_STYLE(
                                    TVS_HASBUTTONS
                                        | TVS_HASLINES
                                        | TVS_LINESATROOT
                                        | TVS_SHOWSELALWAYS,
                                ),
                            10,
                            10,
                            580,
                            250,
                            hwnd,
                            HMENU(ID_LIST as isize),
                            HINSTANCE(0),
                            None,
                        )
                    };
                    let mut tree_values = Vec::new();
                    let mut first_group = HTREEITEM(0);
                    let mut initial_selection = HTREEITEM(0);
                    for group in groups {
                        let parent_item = insert_tree_item(tree, HTREEITEM(0), &group.label, -1);
                        if parent_item.0 == 0 {
                            continue;
                        }
                        if first_group.0 == 0 {
                            first_group = parent_item;
                        }
                        for item in group.items {
                            let value_index = tree_values.len();
                            let item_value = item.value;
                            tree_values.push(item_value.clone());
                            let child_item = insert_tree_item(
                                tree,
                                parent_item,
                                &item.label,
                                value_index as isize,
                            );
                            if init.initial_tree_value.as_deref() == Some(item_value.as_str()) {
                                initial_selection = child_item;
                                crate::send_message_w_safe(
                                    tree,
                                    TVM_EXPAND,
                                    WPARAM(TVE_EXPAND.0 as usize),
                                    LPARAM(parent_item.0),
                                );
                            }
                        }
                    }
                    let target = if initial_selection.0 != 0 {
                        initial_selection
                    } else {
                        first_group
                    };
                    if target.0 != 0 {
                        crate::send_message_w_safe(
                            tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(target.0),
                        );
                        crate::send_message_w_safe(
                            tree,
                            TVM_ENSUREVISIBLE,
                            WPARAM(0),
                            LPARAM(target.0),
                        );
                    }
                    unsafe {
                        SetFocus(tree);
                    }
                    (ControlKind::Tree(tree), tree_values)
                }
            };

            let secondary_button = init.secondary_action_label.as_ref().map(|label| unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(label).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    10,
                    280,
                    240,
                    28,
                    hwnd,
                    HMENU(ID_SECONDARY as isize),
                    HINSTANCE(0),
                    None,
                )
            });

            let ok = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "options.ok")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    390,
                    280,
                    90,
                    28,
                    hwnd,
                    HMENU(ID_OK as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            let cancel = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "options.cancel")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    490,
                    280,
                    90,
                    28,
                    hwnd,
                    HMENU(ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            if hfont.0 != 0 {
                let control_hwnd = match control {
                    ControlKind::List(list) => list,
                    ControlKind::Tree(tree) => tree,
                };
                unsafe {
                    SendMessageW(
                        control_hwnd,
                        WM_SETFONT,
                        WPARAM(hfont.0 as usize),
                        LPARAM(1),
                    );
                    if let Some(button) = secondary_button {
                        SendMessageW(button, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                    SendMessageW(ok, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    SendMessageW(cancel, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let state = Box::new(InterpreterSelectState {
                control,
                tree_values,
                result: init.result.clone(),
            });
            crate::set_window_long_ptr_w_safe(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = wparam.0;
            match id {
                ID_OK => {
                    with_interpreter_state(hwnd, |state| match &state.control {
                        ControlKind::List(list) => {
                            let sel = crate::send_message_w_safe(
                                *list,
                                LB_GETCURSEL,
                                WPARAM(0),
                                LPARAM(0),
                            )
                            .0;
                            if sel >= 0 {
                                let len = crate::send_message_w_safe(
                                    *list,
                                    LB_GETTEXTLEN,
                                    WPARAM(sel as usize),
                                    LPARAM(0),
                                )
                                .0;
                                if len >= 0 {
                                    let mut buf = vec![0u16; (len + 1) as usize];
                                    crate::send_message_w_safe(
                                        *list,
                                        LB_GETTEXT,
                                        WPARAM(sel as usize),
                                        LPARAM(buf.as_mut_ptr() as isize),
                                    );
                                    let value = String::from_utf16_lossy(&buf[..len as usize]);
                                    *state.result.lock().unwrap_or_else(|e| e.into_inner()) =
                                        Some(InterpreterSelectionResult::Item(value));
                                }
                            }
                        }
                        ControlKind::Tree(tree) => {
                            let caret = HTREEITEM(
                                crate::send_message_w_safe(
                                    *tree,
                                    TVM_GETNEXTITEM,
                                    WPARAM(TVGN_CARET as usize),
                                    LPARAM(0),
                                )
                                .0,
                            );
                            if caret.0 != 0 {
                                let mut item = TVITEMW {
                                    mask: TVIF_PARAM | TVIF_TEXT,
                                    hItem: caret,
                                    ..Default::default()
                                };
                                let mut text = vec![0u16; 512];
                                item.pszText = windows::core::PWSTR(text.as_mut_ptr());
                                item.cchTextMax = text.len() as i32;
                                if crate::send_message_w_safe(
                                    *tree,
                                    TVM_GETITEMW,
                                    WPARAM(0),
                                    LPARAM(&mut item as *mut _ as isize),
                                )
                                .0 != 0
                                {
                                    if item.lParam.0 >= 0 {
                                        let index = item.lParam.0 as usize;
                                        if let Some(value) = state.tree_values.get(index) {
                                            *state
                                                .result
                                                .lock()
                                                .unwrap_or_else(|e| e.into_inner()) = Some(
                                                InterpreterSelectionResult::Item(value.clone()),
                                            );
                                        }
                                    } else {
                                        crate::send_message_w_safe(
                                            *tree,
                                            TVM_EXPAND,
                                            WPARAM(TVE_EXPAND.0 as usize),
                                            LPARAM(caret.0),
                                        );
                                    }
                                }
                            }
                        }
                    });
                    if with_interpreter_state(hwnd, |state| {
                        state
                            .result
                            .lock()
                            .unwrap_or_else(|e| e.into_inner())
                            .is_some()
                    })
                    .unwrap_or(false)
                    {
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    }
                    LRESULT(0)
                }
                ID_SECONDARY => {
                    with_interpreter_state(hwnd, |state| {
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) =
                            Some(InterpreterSelectionResult::SecondaryAction);
                    });
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    LRESULT(0)
                }
                ID_CANCEL => {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    LRESULT(0)
                }
                _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            crate::log_if_err!(crate::destroy_window_safe(hwnd));
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA)
                as *mut InterpreterSelectState;
            if !ptr.is_null() {
                let _unused_box = crate::box_from_raw_safe(ptr);
            }
            LRESULT(0)
        }
        _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
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

fn with_interpreter_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut InterpreterSelectState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut InterpreterSelectState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

fn is_list_navigation_key(key: u32) -> bool {
    key == VK_UP.0 as u32
        || key == VK_DOWN.0 as u32
        || key == VK_PRIOR.0 as u32
        || key == VK_NEXT.0 as u32
        || key == VK_HOME.0 as u32
        || key == VK_END.0 as u32
}

fn is_tree_navigation_key(key: u32) -> bool {
    is_list_navigation_key(key) || key == VK_LEFT.0 as u32 || key == VK_RIGHT.0 as u32
}
