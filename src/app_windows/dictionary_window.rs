use crate::accessibility::{handle_accessibility, to_wide};
use crate::i18n;
use crate::settings::{DictionaryEntry, Language, save_settings};
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{WC_BUTTON, WC_LISTBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus, VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CreateWindowExW, DefWindowProcW, DestroyWindow,
    ES_AUTOHSCROLL, GWLP_USERDATA, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU,
    IDC_ARROW, IDCANCEL, IDOK, LB_ADDSTRING, LB_GETCOUNT, LB_GETCURSEL, LB_GETITEMDATA,
    LB_RESETCONTENT, LB_SETCURSEL, LB_SETITEMDATA, LBN_SELCHANGE, LBS_HASSTRINGS, LBS_NOTIFY,
    LoadCursorW, MSG, PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, SetWindowTextW, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE,
    WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const DICTIONARY_CLASS_NAME: &str = "NovapadDictionary";
const DICTIONARY_ENTRY_CLASS_NAME: &str = "NovapadDictionaryEntry";

const DICT_ID_LIST: usize = 9101;
const DICT_ID_ADD: usize = 9102;
const DICT_ID_EDIT: usize = 9103;
const DICT_ID_REMOVE: usize = 9104;
const DICT_ID_CLOSE: usize = 9105;

const DICT_ENTRY_ID_ORIGINAL: usize = 9201;
const DICT_ENTRY_ID_REPLACEMENT: usize = 9202;
const DICT_ENTRY_ID_OK: usize = 9203;
const DICT_ENTRY_ID_CANCEL: usize = 9204;
const DICT_FOCUS_LIST_MSG: u32 = WM_APP + 9;

struct DictionaryWindowState {
    parent: HWND,
    hwnd_list: HWND,
    hwnd_edit: HWND,
    hwnd_remove: HWND,
}

struct DictionaryEntryState {
    parent: HWND,
    owner: HWND,
    edit_original: HWND,
    edit_replacement: HWND,
    ok_button: HWND,
    index: Option<usize>,
}

struct DictionaryLabels {
    title: String,
    add: String,
    edit: String,
    remove: String,
    close: String,
    entry_title_add: String,
    entry_title_edit: String,
    entry_original: String,
    entry_replacement: String,
    entry_ok: String,
    entry_cancel: String,
}

fn dictionary_labels(language: Language) -> DictionaryLabels {
    DictionaryLabels {
        title: i18n::tr(language, "dictionary.title"),
        add: i18n::tr(language, "dictionary.add"),
        edit: i18n::tr(language, "dictionary.edit"),
        remove: i18n::tr(language, "dictionary.remove"),
        close: i18n::tr(language, "dictionary.close"),
        entry_title_add: i18n::tr(language, "dictionary.entry_title_add"),
        entry_title_edit: i18n::tr(language, "dictionary.entry_title_edit"),
        entry_original: i18n::tr(language, "dictionary.entry_original"),
        entry_replacement: i18n::tr(language, "dictionary.entry_replacement"),
        entry_ok: i18n::tr(language, "dictionary.entry_ok"),
        entry_cancel: i18n::tr(language, "dictionary.entry_cancel"),
    }
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN {
        if msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
            crate::log_if_err!(DestroyWindow(hwnd));
            return true;
        }
        if msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
            return handle_accessibility(hwnd, msg);
        }
    }
    handle_accessibility(hwnd, msg)
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.dictionary_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(DICTIONARY_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(dictionary_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = dictionary_labels(language);
    let title = to_wide(&labels.title);

    let window = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        0,
        0,
        520,
        430,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if window.0 != 0 {
        if with_state(parent, |state| {
            state.dictionary_window = window;
        })
        .is_none()
        {
            crate::log_debug("Failed to access dictionary state");
        }
        EnableWindow(parent, false);
        SetForegroundWindow(window);
    }
}

unsafe extern "system" fn dictionary_wndproc(
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
            let labels = dictionary_labels(language);

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
                480,
                270,
                hwnd,
                HMENU(DICT_ID_LIST as isize),
                HINSTANCE(0),
                None,
            );

            let hwnd_add = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.add).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                10,
                290,
                240,
                30,
                hwnd,
                HMENU(DICT_ID_ADD as isize),
                HINSTANCE(0),
                None,
            );

            let hwnd_edit = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.edit).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                260,
                290,
                230,
                30,
                hwnd,
                HMENU(DICT_ID_EDIT as isize),
                HINSTANCE(0),
                None,
            );

            let hwnd_remove = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.remove).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                10,
                330,
                240,
                30,
                hwnd,
                HMENU(DICT_ID_REMOVE as isize),
                HINSTANCE(0),
                None,
            );

            let hwnd_close = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.close).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                260,
                330,
                230,
                30,
                hwnd,
                HMENU(DICT_ID_CLOSE as isize),
                HINSTANCE(0),
                None,
            );

            for ctrl in [hwnd_list, hwnd_add, hwnd_edit, hwnd_remove, hwnd_close] {
                if ctrl.0 != 0 && hfont.0 != 0 {
                    SendMessageW(ctrl, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let state = Box::new(DictionaryWindowState {
                parent,
                hwnd_list,
                hwnd_edit,
                hwnd_remove,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            refresh_dictionary_list(hwnd);
            update_button_states(hwnd);
            SetFocus(hwnd_list);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            let notify = (wparam.0 >> 16) as u16;
            match cmd_id {
                DICT_ID_ADD => {
                    open_entry_dialog(hwnd, None);
                    LRESULT(0)
                }
                DICT_ID_EDIT => {
                    if let Some(index) = selected_dictionary_index(hwnd) {
                        open_entry_dialog(hwnd, Some(index));
                    }
                    LRESULT(0)
                }
                DICT_ID_REMOVE => {
                    remove_selected_entry(hwnd);
                    LRESULT(0)
                }
                DICT_ID_CLOSE => {
                    crate::log_if_err!(DestroyWindow(hwnd));
                    LRESULT(0)
                }
                DICT_ID_LIST if notify == LBN_SELCHANGE as u16 => {
                    update_button_states(hwnd);
                    LRESULT(0)
                }
                cmd if cmd == IDCANCEL.0 as usize || cmd == 2 => {
                    crate::log_if_err!(DestroyWindow(hwnd));
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        DICT_FOCUS_LIST_MSG => {
            let list = with_dictionary_state(hwnd, |s| s.hwnd_list).unwrap_or(HWND(0));
            if list.0 != 0 {
                SetForegroundWindow(hwnd);
                SetFocus(list);
            }
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                crate::log_if_err!(DestroyWindow(hwnd));
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            crate::log_if_err!(DestroyWindow(hwnd));
            LRESULT(0)
        }
        WM_DESTROY => {
            let parent = with_dictionary_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                EnableWindow(parent, true);
                SetForegroundWindow(parent);
                SetFocus(parent);
                if let Some(edit) = crate::get_active_edit(parent) {
                    SetFocus(edit);
                }
                if with_state(parent, |state| {
                    state.dictionary_window = HWND(0);
                })
                .is_none()
                {
                    crate::log_debug("Failed to access dictionary state");
                }
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DictionaryWindowState;
            if !ptr.is_null() {
                drop(Box::from_raw(ptr));
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_dictionary_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut DictionaryWindowState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DictionaryWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe fn update_button_states(hwnd: HWND) {
    let (hwnd_list, hwnd_edit, hwnd_remove) =
        match with_dictionary_state(hwnd, |s| (s.hwnd_list, s.hwnd_edit, s.hwnd_remove)) {
            Some(values) => values,
            None => return,
        };

    let count = SendMessageW(hwnd_list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0;
    let sel = SendMessageW(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let has_selection = count > 0 && sel >= 0;
    EnableWindow(hwnd_edit, has_selection);
    EnableWindow(hwnd_remove, has_selection);
}

pub unsafe fn refresh_dictionary_list(hwnd: HWND) {
    let (parent, hwnd_list) = match with_dictionary_state(hwnd, |s| (s.parent, s.hwnd_list)) {
        Some(values) => values,
        None => return,
    };

    let selected = SendMessageW(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    SendMessageW(hwnd_list, LB_RESETCONTENT, WPARAM(0), LPARAM(0));

    let entries = with_state(parent, |state| state.settings.dictionary.clone()).unwrap_or_default();
    for (idx, entry) in entries.iter().enumerate() {
        let label = format!("{} -> {}", entry.original, entry.replacement);
        let lb_idx = SendMessageW(
            hwnd_list,
            LB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&label).as_ptr() as isize),
        )
        .0;
        if lb_idx >= 0 {
            SendMessageW(
                hwnd_list,
                LB_SETITEMDATA,
                WPARAM(lb_idx as usize),
                LPARAM(idx as isize),
            );
        }
    }

    let count = SendMessageW(hwnd_list, LB_GETCOUNT, WPARAM(0), LPARAM(0)).0;
    if count > 0 {
        let target = if selected >= 0 && selected < count {
            selected
        } else {
            0
        };
        SendMessageW(hwnd_list, LB_SETCURSEL, WPARAM(target as usize), LPARAM(0));
    }
    update_button_states(hwnd);
}

unsafe fn selected_dictionary_index(hwnd: HWND) -> Option<usize> {
    let hwnd_list = with_dictionary_state(hwnd, |s| s.hwnd_list).unwrap_or(HWND(0));
    if hwnd_list.0 == 0 {
        return None;
    }
    let sel = SendMessageW(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel < 0 {
        return None;
    }
    let idx = SendMessageW(hwnd_list, LB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as isize;
    if idx < 0 {
        return None;
    }
    Some(idx as usize)
}

unsafe fn remove_selected_entry(hwnd: HWND) {
    let (parent, _list) = match with_dictionary_state(hwnd, |s| (s.parent, s.hwnd_list)) {
        Some(values) => values,
        None => return,
    };
    let Some(index) = selected_dictionary_index(hwnd) else {
        return;
    };
    if with_state(parent, |state| {
        if index < state.settings.dictionary.len() {
            state.settings.dictionary.remove(index);
        }
        save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to access dictionary state");
    }
    refresh_dictionary_list(hwnd);
    if let Err(_e) = PostMessageW(hwnd, DICT_FOCUS_LIST_MSG, WPARAM(0), LPARAM(0)) {
        crate::log_debug(&format!("Error: {:?}", _e));
    }
}

unsafe fn open_entry_dialog(owner: HWND, index: Option<usize>) {
    let parent = with_dictionary_state(owner, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(DICTIONARY_ENTRY_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(dictionary_entry_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = dictionary_labels(language);
    let title = if index.is_some() {
        &labels.entry_title_edit
    } else {
        &labels.entry_title_add
    };

    let params = Box::new(DictionaryEntryState {
        parent,
        owner,
        edit_original: HWND(0),
        edit_replacement: HWND(0),
        ok_button: HWND(0),
        index,
    });

    let dialog = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(to_wide(title).as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        0,
        0,
        420,
        220,
        owner,
        None,
        hinstance,
        Some(Box::into_raw(params) as *const std::ffi::c_void),
    );

    if dialog.0 != 0 {
        if with_state(parent, |state| {
            state.dictionary_entry_dialog = dialog;
        })
        .is_none()
        {
            crate::log_debug("Failed to access dictionary state");
        }
        EnableWindow(owner, false);
        SetForegroundWindow(dialog);
    }
}

unsafe extern "system" fn dictionary_entry_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let state_ptr = (*create_struct).lpCreateParams as *mut DictionaryEntryState;
            if state_ptr.is_null() {
                return LRESULT(0);
            }
            let mut state = Box::from_raw(state_ptr);
            let language = with_state(state.parent, |s| s.settings.language).unwrap_or_default();
            let labels = dictionary_labels(language);
            let hfont = with_state(state.parent, |s| s.hfont).unwrap_or(HFONT(0));

            let label_original = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.entry_original).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                10,
                380,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_original = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                10,
                32,
                380,
                24,
                hwnd,
                HMENU(DICT_ENTRY_ID_ORIGINAL as isize),
                HINSTANCE(0),
                None,
            );

            let label_replacement = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.entry_replacement).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                64,
                380,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_replacement = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                10,
                86,
                380,
                24,
                hwnd,
                HMENU(DICT_ENTRY_ID_REPLACEMENT as isize),
                HINSTANCE(0),
                None,
            );

            let ok_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.entry_ok).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                200,
                130,
                90,
                28,
                hwnd,
                HMENU(DICT_ENTRY_ID_OK as isize),
                HINSTANCE(0),
                None,
            );
            let cancel_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.entry_cancel).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                300,
                130,
                90,
                28,
                hwnd,
                HMENU(DICT_ENTRY_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            for ctrl in [
                label_original,
                edit_original,
                label_replacement,
                edit_replacement,
                ok_button,
                cancel_button,
            ] {
                if ctrl.0 != 0 && hfont.0 != 0 {
                    SendMessageW(ctrl, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            if let Some(index) = state.index
                && let Some(entry) =
                    with_state(state.parent, |s| s.settings.dictionary.get(index).cloned())
                        .unwrap_or(None)
            {
                if let Err(_e) =
                    SetWindowTextW(edit_original, PCWSTR(to_wide(&entry.original).as_ptr()))
                {
                    crate::log_debug(&format!("Failed to set edit_original text: {:?}", _e));
                }
                if let Err(_e) = SetWindowTextW(
                    edit_replacement,
                    PCWSTR(to_wide(&entry.replacement).as_ptr()),
                ) {
                    crate::log_debug(&format!("Failed to set edit_replacement text: {:?}", _e));
                }
            }

            state.edit_original = edit_original;
            state.edit_replacement = edit_replacement;
            state.ok_button = ok_button;
            let state_ptr = Box::into_raw(state);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, state_ptr as isize);
            SetFocus(edit_original);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            match cmd_id {
                DICT_ENTRY_ID_OK => {
                    apply_entry_dialog(hwnd);
                    LRESULT(0)
                }
                cmd if cmd == IDOK.0 as usize => {
                    apply_entry_dialog(hwnd);
                    LRESULT(0)
                }
                DICT_ENTRY_ID_CANCEL | 2 => {
                    crate::log_if_err!(DestroyWindow(hwnd));
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                crate::log_if_err!(DestroyWindow(hwnd));
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            crate::log_if_err!(DestroyWindow(hwnd));
            LRESULT(0)
        }
        WM_DESTROY => {
            let (owner, parent) =
                with_entry_state(hwnd, |s| (s.owner, s.parent)).unwrap_or((HWND(0), HWND(0)));
            if owner.0 != 0 {
                EnableWindow(owner, true);
                SetForegroundWindow(owner);
                if let Err(_e) = PostMessageW(owner, DICT_FOCUS_LIST_MSG, WPARAM(0), LPARAM(0)) {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
            }
            if parent.0 != 0
                && with_state(parent, |state| {
                    state.dictionary_entry_dialog = HWND(0);
                })
                .is_none()
            {
                crate::log_debug("Failed to access dictionary state");
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DictionaryEntryState;
            if !ptr.is_null() {
                drop(Box::from_raw(ptr));
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_entry_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut DictionaryEntryState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DictionaryEntryState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe fn apply_entry_dialog(hwnd: HWND) {
    let (parent, owner, edit_original, edit_replacement, index) =
        match with_entry_state(hwnd, |s| {
            (
                s.parent,
                s.owner,
                s.edit_original,
                s.edit_replacement,
                s.index,
            )
        }) {
            Some(values) => values,
            None => return,
        };

    let original = get_window_text(edit_original);
    let replacement = get_window_text(edit_replacement);
    if original.trim().is_empty() {
        return;
    }

    if with_state(parent, |state| {
        match index {
            Some(idx) => {
                if idx < state.settings.dictionary.len() {
                    state.settings.dictionary[idx] = DictionaryEntry {
                        original,
                        replacement,
                    };
                }
            }
            None => {
                state.settings.dictionary.push(DictionaryEntry {
                    original,
                    replacement,
                });
            }
        }
        save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to access dictionary state");
    }

    refresh_dictionary_list(owner);
    if let Err(_e) = PostMessageW(owner, DICT_FOCUS_LIST_MSG, WPARAM(0), LPARAM(0)) {
        crate::log_debug(&format!("Error: {:?}", _e));
    }
    crate::log_if_err!(DestroyWindow(hwnd));
}

unsafe fn get_window_text(hwnd: HWND) -> String {
    let len = GetWindowTextLengthW(hwnd);
    if len <= 0 {
        return String::new();
    }
    let mut buf = vec![0u16; (len + 1) as usize];
    let read = GetWindowTextW(hwnd, &mut buf);
    String::from_utf16_lossy(&buf[..read as usize])
}
