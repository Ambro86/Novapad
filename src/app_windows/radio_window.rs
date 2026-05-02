use crate::accessibility::{PlayerCommand, handle_player_keyboard, to_wide};
use crate::i18n;
use crate::launch_stream_url_in_mpv;
use crate::settings::{Language, RadioFavorite, load_settings, save_settings};
use serde::Deserialize;
use std::collections::HashMap;
use std::time::Duration;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, DEFAULT_GUI_FONT, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreatePopupMenu, CreateWindowExW,
    DefWindowProcW, DestroyMenu, DestroyWindow, DispatchMessageW, ES_AUTOHSCROLL, GWLP_USERDATA,
    GetCursorPos, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW, IsDialogMessageW, IsWindow,
    LoadCursorW, MB_ICONINFORMATION, MB_OK, MF_STRING, MoveWindow, RegisterClassW, SW_SHOW,
    SetWindowLongPtrW, ShowWindow, TPM_NONOTIFY, TPM_RETURNCMD, TrackPopupMenu, TranslateMessage,
    WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU, WM_CREATE,
    WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL, WM_SETFONT, WM_SIZE, WNDCLASSW, WS_BORDER,
    WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_DLGMODALFRAME, WS_OVERLAPPED, WS_POPUP,
    WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

const CLASS_NAME: &str = "SonarpadRadioWindow";
const RADIO_RESULTS_PAGE_SIZE: usize = 25;
const RADIO_BROWSER_LIMIT: &str = "100000";

const ID_EDIT_SEARCH: usize = 1001;
const ID_COMBO_LANGUAGE: usize = 1002;
const ID_LIST_RESULTS: usize = 1003;
const ID_BUTTON_SEARCH: usize = 1005;
const ID_BUTTON_PREV: usize = 1006;
const ID_BUTTON_NEXT: usize = 1007;
const ID_BUTTON_CLOSE: usize = 1011;
const ID_LABEL_PAGE: usize = 1012;
const ID_LABEL_SEARCH: usize = 1013;
const ID_LABEL_LANGUAGE: usize = 1014;

const CB_ADDSTRING: u32 = 0x0143;
const CB_SETCURSEL: u32 = 0x014E;
const CB_GETCURSEL: u32 = 0x0147;
const LB_ADDSTRING: u32 = 0x0180;
const LB_RESETCONTENT: u32 = 0x0184;
const LB_SETCURSEL: u32 = 0x0186;
const LB_GETCURSEL: u32 = 0x0188;
const WM_RADIO_FOCUS_RESULTS: u32 = WM_APP + 77;
const ID_CONTEXT_ADD_FAVORITE: usize = 1;
const ID_CONTEXT_REMOVE_FAVORITE: usize = 2;

#[derive(Clone, Deserialize)]
struct RadioStation {
    name: String,
    stream_url: String,
}

#[derive(Deserialize)]
struct RadioBrowserStation {
    #[serde(default)]
    name: String,
    #[serde(default)]
    url: String,
    #[serde(default)]
    url_resolved: String,
}

struct RadioDialogState {
    parent: HWND,
    language: Language,
    edit_search: HWND,
    combo_language: HWND,
    list_results: HWND,
    label_page: HWND,
    all_results: Vec<RadioFavorite>,
    page: usize,
    languages: Vec<(String, String)>,
    stations_by_language: HashMap<String, Vec<RadioStation>>,
}

pub fn open(parent: HWND) {
    unsafe {
        let language = load_settings().language;
        let hinstance = HINSTANCE(GetModuleHandleW(None).map(|m| m.0).unwrap_or(0));
        let class_name = to_wide(CLASS_NAME);
        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hInstance: hinstance,
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let title = to_wide(&tr(language, "radio.title", "Radio da tutto il mondo"));
        let hwnd = CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            760,
            430,
            parent,
            HMENU(0),
            hinstance,
            Some(parent.0 as *const std::ffi::c_void),
        );
        if hwnd.0 == 0 {
            return;
        }
        ShowWindow(hwnd, SW_SHOW);

        let mut msg = Default::default();
        while IsWindow(hwnd).as_bool() && GetMessageW(&mut msg, HWND(0), 0, 0).as_bool() {
            match route_player_keyboard(hwnd, parent, &msg) {
                RadioLoopAction::NotHandled => {}
                RadioLoopAction::Handled => continue,
            }
            if handle_enter_key(hwnd, &msg) {
                continue;
            }
            if !IsDialogMessageW(hwnd, &msg).as_bool() {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    }
}

enum RadioLoopAction {
    NotHandled,
    Handled,
}

fn handle_enter_key(hwnd: HWND, msg: &windows::Win32::UI::WindowsAndMessaging::MSG) -> bool {
    if msg.message != WM_KEYDOWN || msg.wParam.0 as u32 != VK_RETURN.0 as u32 {
        return false;
    }
    let target = with_radio_state(hwnd, |state| {
        if msg.hwnd == state.list_results {
            Some(RadioEnterTarget::Results)
        } else if msg.hwnd == state.combo_language
            || crate::is_child_safe(state.combo_language, msg.hwnd)
        {
            Some(RadioEnterTarget::Language)
        } else {
            None
        }
    })
    .flatten();
    match target {
        Some(RadioEnterTarget::Results) => open_selected(hwnd),
        Some(RadioEnterTarget::Language) => show_all_and_focus(hwnd),
        None => return false,
    }
    true
}

enum RadioEnterTarget {
    Results,
    Language,
}

fn route_player_keyboard(
    hwnd: HWND,
    parent: HWND,
    msg: &windows::Win32::UI::WindowsAndMessaging::MSG,
) -> RadioLoopAction {
    let is_escape = msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32;
    if is_escape {
        crate::log_debug(&format!(
            "Radio: ESC in loop hwnd={:?} msg_hwnd={:?}",
            hwnd, msg.hwnd
        ));
    }
    let Some((has_player, skip_seconds)) = crate::with_state(parent, |state| {
        (
            state.active_audiobook.is_some() || state.active_mpv_session.is_some(),
            state.settings.audiobook_skip_seconds,
        )
    }) else {
        if is_escape {
            crate::log_debug("Radio: ESC parent AppState unavailable");
        }
        return RadioLoopAction::NotHandled;
    };
    if !has_player {
        if is_escape {
            crate::log_debug("Radio: ESC no active player, focusing results list");
            focus_results_list(hwnd);
            return RadioLoopAction::Handled;
        }
        return RadioLoopAction::NotHandled;
    }
    let command = handle_player_keyboard(msg, skip_seconds);
    if is_escape {
        crate::log_debug(&format!("Radio: ESC player command {:?}", command));
    }
    if matches!(command, PlayerCommand::None) {
        return RadioLoopAction::NotHandled;
    }
    let stop = matches!(command, PlayerCommand::Stop);
    crate::handle_player_command(parent, command);
    if stop {
        crate::log_debug(
            "Radio: ESC stopped player, closing player document and focusing results list",
        );
        crate::editor_manager::close_current_document(parent);
        focus_results_list(hwnd);
        crate::log_if_err!(crate::post_message_w_safe(
            hwnd,
            WM_RADIO_FOCUS_RESULTS,
            WPARAM(0),
            LPARAM(0)
        ));
    }
    RadioLoopAction::Handled
}

fn focus_results_list(hwnd: HWND) {
    crate::set_foreground_window_safe(hwnd);
    if let Some(list_results) = with_radio_state(hwnd, |state| state.list_results) {
        crate::log_debug(&format!(
            "Radio: focus results list hwnd={:?} list={:?} before={:?}",
            hwnd,
            list_results,
            crate::get_focus_safe()
        ));
        crate::set_focus_safe(list_results);
        crate::send_message_w_safe(
            hwnd,
            WM_NEXTDLGCTL,
            WPARAM(list_results.0 as usize),
            LPARAM(1),
        );
        crate::log_debug(&format!(
            "Radio: focus results list after={:?}",
            crate::get_focus_safe()
        ));
    } else {
        crate::log_debug("Radio: focus results list failed, state unavailable");
    }
}

fn show_station_context_menu(hwnd: HWND, wparam: WPARAM, lparam: LPARAM) -> bool {
    let target = HWND(wparam.0 as isize);
    let Some((list_results, language)) =
        with_radio_state(hwnd, |state| (state.list_results, state.language))
    else {
        return false;
    };
    if target.0 != 0 && target != hwnd && target != list_results {
        return false;
    }
    if selected_result(hwnd).is_none() {
        return false;
    }

    let menu = match unsafe { CreatePopupMenu() } {
        Ok(menu) => menu,
        Err(err) => {
            crate::log_debug(&format!("Radio: CreatePopupMenu failed: {}", err));
            return false;
        }
    };

    let add_label = to_wide(&tr(language, "radio.add_favorite", "Aggiungi ai preferiti"));
    if let Err(err) = unsafe {
        AppendMenuW(
            menu,
            MF_STRING,
            ID_CONTEXT_ADD_FAVORITE,
            PCWSTR(add_label.as_ptr()),
        )
    } {
        crate::log_debug(&format!("Radio: AppendMenuW add favorite failed: {}", err));
        crate::log_if_err!(unsafe { DestroyMenu(menu) });
        return false;
    }

    let remove_label = to_wide(&tr(
        language,
        "radio.remove_favorite",
        "Rimuovi dai preferiti",
    ));
    if let Err(err) = unsafe {
        AppendMenuW(
            menu,
            MF_STRING,
            ID_CONTEXT_REMOVE_FAVORITE,
            PCWSTR(remove_label.as_ptr()),
        )
    } {
        crate::log_debug(&format!(
            "Radio: AppendMenuW remove favorite failed: {}",
            err
        ));
        crate::log_if_err!(unsafe { DestroyMenu(menu) });
        return false;
    }

    let point = if lparam.0 == -1 {
        let mut pt = POINT::default();
        if let Err(err) = unsafe { GetCursorPos(&mut pt) } {
            crate::log_debug(&format!("Radio: GetCursorPos failed: {}", err));
            crate::log_if_err!(unsafe { DestroyMenu(menu) });
            return false;
        }
        pt
    } else {
        POINT {
            x: (lparam.0 as u32 & 0xFFFF) as i16 as i32,
            y: ((lparam.0 as u32 >> 16) & 0xFFFF) as i16 as i32,
        }
    };

    let command = unsafe {
        TrackPopupMenu(
            menu,
            TPM_RETURNCMD | TPM_NONOTIFY,
            point.x,
            point.y,
            0,
            hwnd,
            None,
        )
    };
    crate::log_if_err!(unsafe { DestroyMenu(menu) });

    match command.0 as usize {
        ID_CONTEXT_ADD_FAVORITE => add_selected_favorite(hwnd),
        ID_CONTEXT_REMOVE_FAVORITE => remove_selected_favorite(hwnd),
        0 => {}
        other => crate::log_debug(&format!("Radio: unknown context menu command {}", other)),
    }
    true
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_RADIO_FOCUS_RESULTS => {
                crate::log_debug("Radio: deferred focus results message");
                focus_results_list(hwnd);
                LRESULT(0)
            }
            WM_CREATE => {
                let create = lparam.0 as *const CREATESTRUCTW;
                let parent = HWND((*create).lpCreateParams as isize);
                create_controls(hwnd, parent);
                LRESULT(0)
            }
            WM_SIZE => {
                layout(hwnd);
                LRESULT(0)
            }
            WM_CONTEXTMENU => {
                if show_station_context_menu(hwnd, wparam, lparam) {
                    LRESULT(0)
                } else {
                    DefWindowProcW(hwnd, msg, wparam, lparam)
                }
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                match id {
                    ID_BUTTON_SEARCH => {
                        search(hwnd);
                        LRESULT(0)
                    }
                    ID_BUTTON_PREV => {
                        change_page(hwnd, -1);
                        LRESULT(0)
                    }
                    ID_BUTTON_NEXT => {
                        change_page(hwnd, 1);
                        LRESULT(0)
                    }
                    ID_BUTTON_CLOSE => {
                        crate::log_if_err!(DestroyWindow(hwnd));
                        LRESULT(0)
                    }
                    _ => DefWindowProcW(hwnd, msg, wparam, lparam),
                }
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    crate::log_debug(&format!("Radio: ESC in wndproc hwnd={:?}", hwnd));
                    let parent = with_radio_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
                    crate::log_if_err!(DestroyWindow(hwnd));
                    if parent.0 != 0 {
                        crate::restore_editor_focus(parent);
                    }
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(DestroyWindow(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => LRESULT(0),
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut RadioDialogState;
                if !ptr.is_null() {
                    let _box = Box::from_raw(ptr);
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn create_controls(hwnd: HWND, parent: HWND) {
    let language = load_settings().language;
    let font = HFONT(crate::get_stock_object_safe(DEFAULT_GUI_FONT).0);
    let languages = radio_menu_languages(language);
    let stations_by_language = embedded_radio_stations();

    let label_search = create_static(
        hwnd,
        &tr(language, "radio.search_text", "Cerca radio:"),
        ID_LABEL_SEARCH,
    );
    let edit_search = create_edit(hwnd, ID_EDIT_SEARCH);
    let label_lang = create_static(
        hwnd,
        &tr(language, "radio.language", "Lingua:"),
        ID_LABEL_LANGUAGE,
    );
    let combo_language = create_combo(hwnd, ID_COMBO_LANGUAGE);
    let list_results = create_listbox(hwnd, ID_LIST_RESULTS);
    let label_page = create_static(hwnd, "", ID_LABEL_PAGE);
    let btn_search = create_button(
        hwnd,
        &tr(language, "radio.search", "Ricerca"),
        ID_BUTTON_SEARCH,
        true,
    );
    let btn_prev = create_button(
        hwnd,
        &tr(language, "radio.previous", "Precedenti"),
        ID_BUTTON_PREV,
        false,
    );
    let btn_next = create_button(
        hwnd,
        &tr(language, "radio.next", "Successivi"),
        ID_BUTTON_NEXT,
        false,
    );
    let btn_close = create_button(
        hwnd,
        &tr(language, "radio.close", "Chiudi"),
        ID_BUTTON_CLOSE,
        false,
    );

    for hwnd_ctrl in [
        label_search,
        edit_search,
        label_lang,
        combo_language,
        list_results,
        label_page,
        btn_search,
        btn_prev,
        btn_next,
        btn_close,
    ] {
        crate::send_message_w_safe(hwnd_ctrl, WM_SETFONT, WPARAM(font.0 as usize), LPARAM(1));
    }
    for (_, label) in &languages {
        let w = to_wide(label);
        crate::send_message_w_safe(
            combo_language,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(w.as_ptr() as isize),
        );
    }
    crate::send_message_w_safe(combo_language, CB_SETCURSEL, WPARAM(0), LPARAM(0));

    let mut state = Box::new(RadioDialogState {
        parent,
        language,
        edit_search,
        combo_language,
        list_results,
        label_page,
        all_results: Vec::new(),
        page: 0,
        languages,
        stations_by_language,
    });
    state.all_results = initial_results(&state);
    crate::set_window_long_ptr_w_safe(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
    populate_results(hwnd);
    layout(hwnd);
    crate::set_focus_safe(edit_search);
}

fn create_static(parent: HWND, text: &str, id: usize) -> HWND {
    let t = to_wide(text);
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            w!("STATIC"),
            PCWSTR(t.as_ptr()),
            WS_CHILD | WS_VISIBLE,
            0,
            0,
            0,
            0,
            parent,
            HMENU(id as isize),
            HINSTANCE(0),
            None,
        )
    }
}
fn create_edit(parent: HWND, id: usize) -> HWND {
    unsafe {
        CreateWindowExW(
            WS_EX_CLIENTEDGE,
            w!("EDIT"),
            PCWSTR::null(),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
            0,
            0,
            0,
            0,
            parent,
            HMENU(id as isize),
            HINSTANCE(0),
            None,
        )
    }
}
fn create_combo(parent: HWND, id: usize) -> HWND {
    unsafe {
        CreateWindowExW(
            WS_EX_CLIENTEDGE,
            w!("COMBOBOX"),
            PCWSTR::null(),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(0x0003),
            0,
            0,
            0,
            200,
            parent,
            HMENU(id as isize),
            HINSTANCE(0),
            None,
        )
    }
}
fn create_listbox(parent: HWND, id: usize) -> HWND {
    unsafe {
        CreateWindowExW(
            WS_EX_CLIENTEDGE,
            w!("LISTBOX"),
            PCWSTR::null(),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(0x0001 | 0x00200000 | 0x00100000),
            0,
            0,
            0,
            0,
            parent,
            HMENU(id as isize),
            HINSTANCE(0),
            None,
        )
    }
}
fn create_button(parent: HWND, text: &str, id: usize, default: bool) -> HWND {
    let t = to_wide(text);
    let mut style = WS_CHILD | WS_VISIBLE | WS_TABSTOP;
    if default {
        style |= WINDOW_STYLE(BS_DEFPUSHBUTTON as u32);
    }
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            w!("BUTTON"),
            PCWSTR(t.as_ptr()),
            style,
            0,
            0,
            0,
            0,
            parent,
            HMENU(id as isize),
            HINSTANCE(0),
            None,
        )
    }
}

fn layout(hwnd: HWND) {
    let mut rc = Default::default();
    crate::log_if_err!(crate::get_client_rect_safe(hwnd, &mut rc));
    let w = rc.right - rc.left;
    let margin = 10;
    let label_w = 80;
    let row_h = 25;
    let mut y = margin;
    move_id(hwnd, ID_LABEL_SEARCH, margin, y + 4, label_w, row_h);
    move_id(
        hwnd,
        ID_EDIT_SEARCH,
        margin + label_w,
        y,
        w - margin * 2 - label_w,
        row_h,
    );
    y += 35;
    move_id(hwnd, ID_LABEL_LANGUAGE, margin, y + 4, label_w, row_h);
    move_id(
        hwnd,
        ID_COMBO_LANGUAGE,
        margin + label_w,
        y,
        w - margin * 2 - label_w,
        140,
    );
    y += 38;
    move_id(hwnd, ID_BUTTON_SEARCH, w - margin - 100, y, 100, 30);
    y += 42;
    let list_h = (rc.bottom - y - 92).max(80);
    move_id(hwnd, ID_LIST_RESULTS, margin, y, w - margin * 2, list_h);
    y += list_h + 8;
    move_id(hwnd, ID_BUTTON_PREV, margin, y, 100, 28);
    move_id(
        hwnd,
        ID_LABEL_PAGE,
        margin + 110,
        y + 5,
        w - margin * 2 - 220,
        28,
    );
    move_id(hwnd, ID_BUTTON_NEXT, w - margin - 100, y, 100, 28);
    y += 36;
    move_id(hwnd, ID_BUTTON_CLOSE, w - margin - 90, y, 90, 30);
}
fn move_id(hwnd: HWND, id: usize, x: i32, y: i32, w: i32, h: i32) {
    let child = crate::get_dlg_item_safe(hwnd, id as i32);
    if child.0 != 0 {
        crate::log_if_err!(unsafe { MoveWindow(child, x, y, w, h, true) });
    }
}

fn with_radio_state<R>(hwnd: HWND, f: impl FnOnce(&mut RadioDialogState) -> R) -> Option<R> {
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut RadioDialogState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

fn tr(language: Language, key: &str, fallback: &str) -> String {
    let value = i18n::tr(language, key);
    if value == key {
        fallback.to_string()
    } else {
        value
    }
}

fn get_edit_text(hwnd: HWND) -> String {
    let len = crate::get_window_text_length_w_safe(hwnd);
    let mut buf = vec![0u16; len as usize + 1];
    crate::get_window_text_w_safe(hwnd, &mut buf);
    String::from_utf16_lossy(&buf[..len as usize])
}

fn radio_menu_languages(language: Language) -> Vec<(String, String)> {
    vec![
        ("it".into(), tr(language, "radio.lang.it", "Italiano")),
        ("en".into(), tr(language, "radio.lang.en", "Inglese")),
        (
            "country:de".into(),
            tr(language, "radio.lang.de", "Germania"),
        ),
        (
            "country:ch".into(),
            tr(language, "radio.lang.ch", "Svizzera"),
        ),
        ("es".into(), tr(language, "radio.lang.es", "Spagnolo")),
        ("pt".into(), tr(language, "radio.lang.pt", "Portoghese")),
        ("sv".into(), tr(language, "radio.lang.sv", "Svedese")),
        ("vi".into(), tr(language, "radio.lang.vi", "Vietnamita")),
        ("cs".into(), tr(language, "radio.lang.cs", "Ceco")),
        ("pl".into(), tr(language, "radio.lang.pl", "Polacco")),
        ("fr".into(), tr(language, "radio.lang.fr", "Francese")),
        ("sr".into(), tr(language, "radio.lang.sr", "Serbo")),
        ("uk".into(), tr(language, "radio.lang.uk", "Ucraino")),
        ("lt".into(), tr(language, "radio.lang.lt", "Lituano")),
        ("ru".into(), tr(language, "radio.lang.ru", "Russo")),
        ("zh".into(), tr(language, "radio.lang.zh", "Cinese")),
    ]
}

fn embedded_radio_stations() -> HashMap<String, Vec<RadioStation>> {
    serde_json::from_str::<HashMap<String, Vec<RadioStation>>>(include_str!(
        "../../i18n/radio.json"
    ))
    .unwrap_or_default()
    .into_iter()
    .map(|(k, v)| (k, normalize_radio_stations(v)))
    .collect()
}

fn initial_results(state: &RadioDialogState) -> Vec<RadioFavorite> {
    let mut results = load_settings().radio_favorites;
    if let Some(stations) = state.stations_by_language.get("it") {
        for s in stations.iter().take(50) {
            results.push(favorite_from_station("it", s));
        }
    }
    normalize_favorites(results)
}

fn selected_language_code(state: &RadioDialogState) -> String {
    let idx = crate::send_message_w_safe(state.combo_language, CB_GETCURSEL, WPARAM(0), LPARAM(0))
        .0
        .max(0) as usize;
    state
        .languages
        .get(idx)
        .map(|(c, _)| c.clone())
        .unwrap_or_else(|| "it".into())
}

fn show_all(hwnd: HWND) {
    let should_populate = with_radio_state(hwnd, |state| {
        let code = selected_language_code(state);
        let mut results = state
            .stations_by_language
            .get(&code)
            .cloned()
            .unwrap_or_default();
        if results.is_empty() {
            match fetch_radio_browser_stations(&code) {
                Ok(v) => results = v,
                Err(e) => {
                    message(hwnd, &tr(state.language, "radio.error", "Errore"), &e);
                    return false;
                }
            }
        }
        state.all_results = normalize_favorites(
            results
                .iter()
                .map(|s| favorite_from_station(&code, s))
                .collect(),
        );
        state.page = 0;
        true
    })
    .unwrap_or(false);
    if should_populate {
        populate_results(hwnd);
    }
}

fn show_all_and_focus(hwnd: HWND) {
    show_all(hwnd);
    focus_results_list(hwnd);
}

fn search(hwnd: HWND) {
    let should_populate = with_radio_state(hwnd, |state| {
        let keyword = normalized_radio_name(&get_edit_text(state.edit_search));
        if keyword.is_empty() {
            message(
                hwnd,
                "Radio",
                &tr(
                    state.language,
                    "radio.enter_search",
                    "Inserisci un testo da cercare.",
                ),
            );
            return false;
        }
        let code = selected_language_code(state);
        let mut stations = state
            .stations_by_language
            .get(&code)
            .cloned()
            .unwrap_or_default();
        if stations.is_empty() {
            stations = fetch_radio_browser_stations(&code).unwrap_or_default();
        }
        let mut results: Vec<_> = stations
            .iter()
            .filter(|s| radio_name_matches_keyword(&s.name, &keyword))
            .map(|s| favorite_from_station(&code, s))
            .collect();
        results.sort_by(|a, b| {
            radio_search_rank(&a.name, &keyword)
                .cmp(&radio_search_rank(&b.name, &keyword))
                .then_with(|| a.name.cmp(&b.name))
        });
        state.all_results = normalize_favorites(results);
        state.page = 0;
        if state.all_results.is_empty() {
            message(
                hwnd,
                "Radio",
                &tr(state.language, "radio.none_found", "Nessuna radio trovata."),
            );
        }
        true
    })
    .unwrap_or(false);
    if should_populate {
        populate_results(hwnd);
    }
}

fn populate_results(hwnd: HWND) {
    with_radio_state(hwnd, |state| {
        crate::send_message_w_safe(state.list_results, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let total_pages = state
            .all_results
            .len()
            .div_ceil(RADIO_RESULTS_PAGE_SIZE)
            .max(1);
        state.page = state.page.min(total_pages.saturating_sub(1));
        let start = state.page * RADIO_RESULTS_PAGE_SIZE;
        let end = (start + RADIO_RESULTS_PAGE_SIZE).min(state.all_results.len());
        for item in &state.all_results[start..end] {
            let w = to_wide(&radio_label(item, state.language));
            crate::send_message_w_safe(
                state.list_results,
                LB_ADDSTRING,
                WPARAM(0),
                LPARAM(w.as_ptr() as isize),
            );
        }
        if end > start {
            crate::send_message_w_safe(state.list_results, LB_SETCURSEL, WPARAM(0), LPARAM(0));
        }
        let page = i18n::tr_f(
            state.language,
            "radio.page",
            &[
                ("current", &(state.page + 1).to_string()),
                ("total", &total_pages.to_string()),
            ],
        );
        let text = if page == "radio.page" {
            format!("Pagina {} di {}", state.page + 1, total_pages)
        } else {
            page
        };
        let w = to_wide(&text);
        crate::log_if_err!(crate::set_window_text_w_safe(
            state.label_page,
            PCWSTR(w.as_ptr())
        ));
    });
}

fn change_page(hwnd: HWND, delta: isize) {
    let should_populate = with_radio_state(hwnd, |state| {
        let total_pages = state
            .all_results
            .len()
            .div_ceil(RADIO_RESULTS_PAGE_SIZE)
            .max(1);
        state.page =
            (state.page as isize + delta).clamp(0, total_pages.saturating_sub(1) as isize) as usize;
        true
    })
    .unwrap_or(false);
    if should_populate {
        populate_results(hwnd);
    }
}

fn selected_result(hwnd: HWND) -> Option<RadioFavorite> {
    with_radio_state(hwnd, selected_result_from_state).flatten()
}

fn selected_result_from_state(state: &mut RadioDialogState) -> Option<RadioFavorite> {
    let sel = crate::send_message_w_safe(state.list_results, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel < 0 {
        return None;
    }
    let index = state.page * RADIO_RESULTS_PAGE_SIZE + sel as usize;
    state.all_results.get(index).cloned()
}

fn open_selected(hwnd: HWND) {
    if let Some((parent, item)) = with_radio_state(hwnd, |state| {
        selected_result_from_state(state).map(|item| (state.parent, item))
    })
    .flatten()
        && let Err(err) =
            launch_stream_url_in_mpv(parent, &item.stream_url, Some(&item.name), None, None, None)
    {
        message(hwnd, "Radio", &err);
    }
}
fn add_selected_favorite(hwnd: HWND) {
    let Some(item) = selected_result(hwnd) else {
        return;
    };
    let mut settings = load_settings();
    if settings
        .radio_favorites
        .iter()
        .any(|f| f.stream_url == item.stream_url)
    {
        message(
            hwnd,
            "Radio",
            &tr(
                settings.language,
                "radio.already_favorite",
                "La radio selezionata è già nei preferiti.",
            ),
        );
        return;
    }
    settings.radio_favorites.push(item.clone());
    settings.radio_favorites = normalize_favorites(settings.radio_favorites.clone());
    let lang = settings.language;
    save_settings(settings);
    message(
        hwnd,
        "Radio",
        &i18n::tr_f(lang, "radio.favorite_added", &[("name", &item.name)]),
    );
}
fn remove_selected_favorite(hwnd: HWND) {
    let Some(item) = selected_result(hwnd) else {
        return;
    };
    let mut settings = load_settings();
    let before = settings.radio_favorites.len();
    settings
        .radio_favorites
        .retain(|f| f.stream_url != item.stream_url);
    if settings.radio_favorites.len() != before {
        let lang = settings.language;
        save_settings(settings);
        message(
            hwnd,
            "Radio",
            &i18n::tr_f(lang, "radio.favorite_removed", &[("name", &item.name)]),
        );
    }
}

fn message(hwnd: HWND, title: &str, body: &str) {
    let t = to_wide(title);
    let b = to_wide(body);
    crate::message_box_w_safe(
        hwnd,
        PCWSTR(b.as_ptr()),
        PCWSTR(t.as_ptr()),
        MB_OK | MB_ICONINFORMATION,
    );
}

fn favorite_from_station(language_code: &str, station: &RadioStation) -> RadioFavorite {
    RadioFavorite {
        language_code: language_code.to_string(),
        name: station.name.clone(),
        stream_url: station.stream_url.clone(),
    }
}
fn normalize_favorites(mut items: Vec<RadioFavorite>) -> Vec<RadioFavorite> {
    items.retain(|x| !x.name.trim().is_empty() && !x.stream_url.trim().is_empty());
    items.sort_by(|a, b| {
        canonical_radio_name(&a.name)
            .cmp(&canonical_radio_name(&b.name))
            .then_with(|| a.stream_url.cmp(&b.stream_url))
    });
    items.dedup_by(|a, b| {
        canonical_radio_name(&a.name) == canonical_radio_name(&b.name)
            || a.stream_url == b.stream_url
    });
    items
}
fn radio_label(f: &RadioFavorite, language: Language) -> String {
    if f.language_code == "custom" || f.language_code == "it" {
        f.name.clone()
    } else {
        format!(
            "{} ({})",
            f.name,
            language_label(language, &f.language_code)
        )
    }
}
fn language_label(language: Language, code: &str) -> String {
    radio_menu_languages(language)
        .into_iter()
        .find(|(c, _)| c == code)
        .map(|(_, l)| l)
        .unwrap_or_else(|| code.to_string())
}
fn normalized_radio_name(value: &str) -> String {
    value
        .to_lowercase()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}
fn canonical_radio_name(value: &str) -> String {
    let mut n = normalized_radio_name(value);
    if let Some(rest) = n.strip_prefix("radio rai ") {
        n = format!("rai radio {rest}");
    }
    n.replace("rai radiouno", "rai radio 1")
        .replace("rai radiodue", "rai radio 2")
        .replace("rai radiotre", "rai radio 3")
        .replace("rai radio1", "rai radio 1")
        .replace("rai radio2", "rai radio 2")
        .replace("rai radio3", "rai radio 3")
        .replace("rai radio uno", "rai radio 1")
        .replace("rai radio due", "rai radio 2")
        .replace("rai radio tre", "rai radio 3")
}
fn normalize_radio_stations(mut stations: Vec<RadioStation>) -> Vec<RadioStation> {
    for s in &mut stations {
        s.name = s
            .name
            .replace('&', "")
            .split_whitespace()
            .collect::<Vec<_>>()
            .join(" ");
        if canonical_radio_name(&s.name) == "rai radio tutta italiana" {
            s.name = "Rai Radio Tutta Italiana".into();
        }
    }
    stations.retain(|s| !s.name.trim().is_empty() && !s.stream_url.trim().is_empty());
    stations.sort_by(|a, b| {
        radio_name_priority(&a.name)
            .cmp(&radio_name_priority(&b.name))
            .then_with(|| a.name.len().cmp(&b.name.len()))
            .then_with(|| a.name.cmp(&b.name))
    });
    stations.dedup_by(|a, b| {
        canonical_radio_name(&a.name) == canonical_radio_name(&b.name)
            || a.stream_url == b.stream_url
    });
    stations
}
fn radio_name_priority(value: &str) -> (usize, usize, String) {
    let n = normalized_radio_name(value);
    let c = canonical_radio_name(value);
    (
        usize::from(!n.starts_with("rai radio ")),
        usize::from(!n.starts_with("rai ")),
        c,
    )
}
fn radio_search_rank(name: &str, keyword: &str) -> (usize, usize, usize, String) {
    let n = normalized_radio_name(name);
    let c = canonical_radio_name(name);
    let exact = n == keyword;
    let starts = n.starts_with(keyword);
    let word = n.contains(&format!(" {keyword}"));
    let pos = n.find(keyword).unwrap_or(usize::MAX);
    let tier = if exact {
        0
    } else if starts {
        1
    } else if word {
        2
    } else {
        3
    };
    (
        tier,
        if keyword == "rai" && c.starts_with("rai radio ") {
            0
        } else {
            1
        },
        pos,
        c,
    )
}
fn radio_name_matches_keyword(name: &str, keyword: &str) -> bool {
    let c = canonical_radio_name(name);
    c == keyword
        || c.starts_with(&format!("{keyword} "))
        || c.contains(&format!(" {keyword} "))
        || (keyword.len() >= 4 && c.split_whitespace().any(|w| w.starts_with(keyword)))
}
fn radio_browser_language_name(code: &str) -> &str {
    match code {
        "cs" => "czech",
        "en" => "english",
        "es" => "spanish",
        "fr" => "french",
        "it" => "italian",
        "lt" => "lithuanian",
        "pl" => "polish",
        "pt" => "portuguese",
        "ru" => "russian",
        "sr" => "serbian",
        "sv" => "swedish",
        "uk" => "ukrainian",
        "vi" => "vietnamese",
        "zh" => "chinese",
        _ => code,
    }
}
fn fetch_radio_browser_stations(language_code: &str) -> Result<Vec<RadioStation>, String> {
    let client = reqwest::blocking::Client::builder()
        .timeout(Duration::from_secs(10))
        .user_agent("Sonarpad Radio/1.0")
        .build()
        .map_err(|e| e.to_string())?;
    let mut last = None;
    for mirror in [
        "https://de1.api.radio-browser.info",
        "https://fi1.api.radio-browser.info",
        "https://at1.api.radio-browser.info",
    ] {
        let mut req = client
            .get(format!("{mirror}/json/stations/search"))
            .query(&[
                ("hidebroken", "true"),
                ("order", "clickcount"),
                ("reverse", "true"),
                ("limit", RADIO_BROWSER_LIMIT),
            ]);
        req = if let Some(country) = language_code.strip_prefix("country:") {
            req.query(&[("countrycode", country)])
        } else {
            req.query(&[
                ("language", radio_browser_language_name(language_code)),
                ("languageExact", "true"),
            ])
        };
        match req
            .send()
            .and_then(|r| r.error_for_status())
            .and_then(|r| r.json::<Vec<RadioBrowserStation>>())
        {
            Ok(v) => {
                return Ok(normalize_radio_stations(
                    v.into_iter()
                        .filter_map(|s| {
                            let url = if s.url_resolved.trim().is_empty() {
                                s.url.trim().to_string()
                            } else {
                                s.url_resolved.trim().to_string()
                            };
                            if url.is_empty() {
                                None
                            } else {
                                Some(RadioStation {
                                    name: if s.name.trim().is_empty() {
                                        url.clone()
                                    } else {
                                        s.name.replace('&', "")
                                    },
                                    stream_url: url,
                                })
                            }
                        })
                        .collect(),
                ));
            }
            Err(e) => last = Some(e.to_string()),
        }
    }
    Err(last.unwrap_or_else(|| "radio browser request failed".to_string()))
}
