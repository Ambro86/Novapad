use crate::accessibility::{handle_accessibility, to_wide};
use crate::editor_manager::{apply_word_wrap_to_all_edits, update_window_title};
use crate::settings::{
    Language, ModifiedMarkerPosition, OpenBehavior, TRUSTED_CLIENT_TOKEN, TtsEngine,
    VOICE_LIST_URL, VoiceInfo, save_settings_with_default_copy, sync_context_menu,
};
use crate::{i18n, rebuild_menus, refresh_voice_panel, tts_engine, with_state};
use std::sync::Arc;
use std::sync::atomic::AtomicBool;
use std::thread;
use tokio::sync::mpsc;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    BST_CHECKED, NMHDR, TCIF_TEXT, TCITEMW, TCM_GETCURSEL, TCM_INSERTITEMW, TCM_SETCURSEL,
    TCN_SELCHANGE, WC_BUTTON, WC_COMBOBOXW, WC_STATIC, WC_TABCONTROLW,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, GetKeyState, SetFocus, VK_CONTROL, VK_ESCAPE, VK_RETURN, VK_SHIFT,
    VK_TAB,
};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCOUNT,
    CB_GETCURSEL, CB_GETDROPPEDSTATE, CB_GETITEMDATA, CB_RESETCONTENT, CB_SETCURSEL,
    CB_SETITEMDATA, CBN_SELCHANGE, CBS_DROPDOWNLIST, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW,
    DefWindowProcW, DestroyWindow, ES_AUTOHSCROLL, ES_PASSWORD, GWLP_USERDATA, GetParent,
    GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW, LoadCursorW, MSG,
    PostMessageW, RegisterClassW, SW_HIDE, SW_SHOW, SW_SHOWNORMAL, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW, ShowWindow, WINDOW_STYLE, WM_APP,
    WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL,
    WM_NOTIFY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, PWSTR, w};

const OPTIONS_CLASS_NAME: &str = "NovapadOptions";
const OPTIONS_ID_LANG: usize = 6001;
const OPTIONS_ID_MODIFIED_MARKER_POSITION: usize = 6023;
const OPTIONS_ID_OPEN: usize = 6002;
const OPTIONS_ID_TTS_ENGINE: usize = 6012;
const OPTIONS_ID_VOICE: usize = 6003;
const OPTIONS_ID_MULTILINGUAL: usize = 6004;
const OPTIONS_ID_SPLIT_ON_NEWLINE: usize = 6007;
const OPTIONS_ID_WORD_WRAP: usize = 6008;
const OPTIONS_ID_SMART_QUOTES: usize = 6025;
const OPTIONS_ID_CONTEXT_MENU: usize = 6026;
const OPTIONS_ID_SPELLCHECK_ENABLED: usize = 6027;
const OPTIONS_ID_SPELLCHECK_LANGUAGE: usize = 6028;
const OPTIONS_ID_MOVE_CURSOR: usize = 6009;
const OPTIONS_ID_TTS_SPEED: usize = 6014;
const OPTIONS_ID_TTS_PITCH: usize = 6020;
const OPTIONS_ID_TTS_VOLUME: usize = 6021;
const OPTIONS_ID_TTS_PREVIEW: usize = 6022;
const OPTIONS_ID_TTS_MANUAL_TUNING: usize = 6031;
const OPTIONS_ID_TTS_SPEED_EDIT: usize = 6032;
const OPTIONS_ID_TTS_PITCH_EDIT: usize = 6033;
const OPTIONS_ID_TTS_VOLUME_EDIT: usize = 6034;
const OPTIONS_ID_AUDIO_SKIP: usize = 6010;
const OPTIONS_ID_AUDIO_SPLIT: usize = 6011;
const OPTIONS_ID_AUDIO_SPLIT_TEXT: usize = 6013;
const OPTIONS_ID_AUDIO_SPLIT_REQUIRE_NEWLINE: usize = 6016;
const OPTIONS_ID_PODCAST_CACHE_LIMIT: usize = 6030;
const OPTIONS_ID_PODCASTINDEX_KEY: usize = 6035;
const OPTIONS_ID_PODCASTINDEX_SECRET: usize = 6036;
const OPTIONS_ID_PODCASTINDEX_SIGNUP: usize = 6037;
const OPTIONS_ID_WRAP_WIDTH: usize = 6017;
const OPTIONS_ID_QUOTE_PREFIX: usize = 6018;
const OPTIONS_ID_CHECK_UPDATES: usize = 6015;
const OPTIONS_ID_PROMPT_PROGRAM: usize = 6019;
const OPTIONS_ID_TABS: usize = 6024;

const OPTIONS_ID_OK: usize = 6005;
const OPTIONS_ID_CANCEL: usize = 6006;

const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
const WM_TTS_SAPI_VOICES_LOADED: u32 = WM_APP + 8;
const AUDIOBOOK_SPLIT_BY_TEXT: u32 = u32::MAX;

const OPTIONS_TAB_GENERAL: i32 = 0;
const OPTIONS_TAB_VOICE: i32 = 1;
const OPTIONS_TAB_EDITOR: i32 = 2;
const OPTIONS_TAB_AUDIO: i32 = 3;
const OPTIONS_TAB_COUNT: i32 = 4;

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
        let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
        if ctrl_down {
            let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
            if let Some(tabs) = with_options_state(hwnd, |state| state.hwnd_tabs) {
                if tabs.0 != 0 {
                    let current = SendMessageW(tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                    let mut next = if shift_down { current - 1 } else { current + 1 };
                    if next < 0 {
                        next = OPTIONS_TAB_COUNT - 1;
                    } else if next >= OPTIONS_TAB_COUNT {
                        next = 0;
                    }
                    let _ = SendMessageW(tabs, TCM_SETCURSEL, WPARAM(next as usize), LPARAM(0));
                    set_active_tab(hwnd, next);
                    SetFocus(tabs);
                    let _ = PostMessageW(hwnd, WM_NEXTDLGCTL, WPARAM(tabs.0 as usize), LPARAM(1));
                    return true;
                }
            }
        }
    }
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = GetFocus();
        if GetParent(focus) == hwnd {
            let dropped = SendMessageW(focus, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 != 0;
            if !dropped {
                let _ = with_options_state(hwnd, |state| {
                    let _ = SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(OPTIONS_ID_OK),
                        LPARAM(state.ok_button.0),
                    );
                });
                return true;
            }
        }
    }
    handle_accessibility(hwnd, msg)
}

struct OptionsDialogState {
    parent: HWND,
    hwnd_tabs: HWND,
    focus_initialized: bool,
    label_language: HWND,
    label_modified_marker_position: HWND,
    label_open: HWND,
    label_tts_engine: HWND,
    label_voice: HWND,
    label_tts_speed: HWND,
    label_tts_pitch: HWND,
    label_tts_volume: HWND,
    button_tts_preview: HWND,
    combo_lang: HWND,
    combo_modified_marker_position: HWND,
    combo_open: HWND,
    combo_tts_engine: HWND,
    combo_voice: HWND,
    combo_tts_speed: HWND,
    combo_tts_pitch: HWND,
    combo_tts_volume: HWND,
    edit_tts_speed: HWND,
    edit_tts_pitch: HWND,
    edit_tts_volume: HWND,
    checkbox_tts_manual: HWND,
    label_audio_skip: HWND,
    combo_audio_skip: HWND,
    label_audio_split: HWND,
    combo_audio_split: HWND,
    label_audio_split_text: HWND,
    edit_audio_split_text: HWND,
    checkbox_audio_split_requires_newline: HWND,
    label_podcast_cache_limit: HWND,
    edit_podcast_cache_limit: HWND,
    label_podcastindex_key: HWND,
    edit_podcastindex_key: HWND,
    label_podcastindex_secret: HWND,
    edit_podcastindex_secret: HWND,
    button_podcastindex_signup: HWND,
    checkbox_multilingual: HWND,
    checkbox_split_on_newline: HWND,
    checkbox_word_wrap: HWND,
    checkbox_smart_quotes: HWND,
    checkbox_spellcheck: HWND,
    label_spellcheck_language: HWND,
    combo_spellcheck_language: HWND,
    label_wrap_width: HWND,
    edit_wrap_width: HWND,
    label_quote_prefix: HWND,
    edit_quote_prefix: HWND,
    checkbox_move_cursor: HWND,
    checkbox_check_updates: HWND,
    checkbox_context_menu: HWND,
    label_prompt_program: HWND,
    combo_prompt_program: HWND,
    ok_button: HWND,
}

struct OptionsLabels {
    title: String,
    tab_general: String,
    tab_voice: String,
    tab_editor: String,
    tab_audio: String,
    label_language: String,
    label_modified_marker_position: String,
    label_open: String,
    label_tts_engine: String,
    label_voice: String,
    label_multilingual: String,
    label_tts_speed: String,
    label_tts_pitch: String,
    label_tts_volume: String,
    label_tts_preview: String,
    label_tts_manual_tuning: String,
    label_split_on_newline: String,
    label_word_wrap: String,
    label_smart_quotes: String,
    label_spellcheck: String,
    label_spellcheck_language: String,
    label_wrap_width: String,
    label_quote_prefix: String,
    label_move_cursor: String,
    label_check_updates: String,
    label_context_menu: String,
    label_prompt_program: String,
    label_audio_skip: String,
    label_audio_split: String,
    label_audio_split_text: String,
    label_audio_split_requires_newline: String,
    label_podcast_cache_limit: String,
    label_podcastindex_key: String,
    label_podcastindex_secret: String,
    label_podcastindex_signup: String,
    lang_it: String,
    lang_en: String,
    lang_es: String,
    lang_pt: String,
    lang_vi: String,
    marker_position_end: String,
    marker_position_beginning: String,
    open_new_tab: String,
    open_new_window: String,
    engine_edge: String,
    engine_sapi5: String,

    split_none: String,
    split_by_text: String,
    split_parts: String,
    spellcheck_lang_follow: String,
    spellcheck_lang_en_us: String,
    spellcheck_lang_en_gb: String,
    spellcheck_lang_it: String,
    spellcheck_lang_es: String,
    spellcheck_lang_pt_br: String,
    spellcheck_lang_fr: String,
    spellcheck_lang_de: String,
    prompt_cmd: String,
    prompt_powershell: String,
    prompt_codex: String,
    ok: String,
    cancel: String,
    voices_empty: String,
}

fn options_labels(language: Language) -> OptionsLabels {
    OptionsLabels {
        title: i18n::tr(language, "options.title"),
        tab_general: i18n::tr(language, "options.tab.general"),
        tab_voice: i18n::tr(language, "options.tab.voice"),
        tab_editor: i18n::tr(language, "options.tab.editor"),
        tab_audio: i18n::tr(language, "options.tab.audio"),
        label_language: i18n::tr(language, "options.label.language"),
        label_modified_marker_position: i18n::tr(
            language,
            "options.label.modified_marker_position",
        ),
        label_open: i18n::tr(language, "options.label.open"),
        label_tts_engine: i18n::tr(language, "options.label.tts_engine"),
        label_voice: i18n::tr(language, "options.label.voice"),
        label_multilingual: i18n::tr(language, "options.label.multilingual"),
        label_tts_speed: i18n::tr(language, "tts_tuning.label_speed"),
        label_tts_pitch: i18n::tr(language, "tts_tuning.label_pitch"),
        label_tts_volume: i18n::tr(language, "tts_tuning.label_volume"),
        label_tts_preview: i18n::tr(language, "options.label.voice_preview"),
        label_tts_manual_tuning: i18n::tr(language, "options.label.tts_manual_tuning"),
        label_split_on_newline: i18n::tr(language, "options.label.split_on_newline"),
        label_word_wrap: i18n::tr(language, "options.label.word_wrap"),
        label_smart_quotes: i18n::tr(language, "options.label.smart_quotes"),
        label_spellcheck: i18n::tr(language, "options.label.spellcheck"),
        label_spellcheck_language: i18n::tr(language, "options.label.spellcheck_language"),
        label_wrap_width: i18n::tr(language, "options.label.wrap_width"),
        label_quote_prefix: i18n::tr(language, "options.label.quote_prefix"),
        label_move_cursor: i18n::tr(language, "options.label.move_cursor"),
        label_check_updates: i18n::tr(language, "options.label.check_updates"),
        label_context_menu: i18n::tr(language, "options.label.context_menu"),
        label_prompt_program: i18n::tr(language, "options.label.prompt_program"),
        label_audio_skip: i18n::tr(language, "options.label.audio_skip"),
        label_audio_split: i18n::tr(language, "options.label.audio_split"),
        label_audio_split_text: i18n::tr(language, "options.label.audio_split_text"),
        label_audio_split_requires_newline: i18n::tr(
            language,
            "options.label.audio_split_requires_newline",
        ),
        label_podcast_cache_limit: i18n::tr(language, "options.label.podcast_cache_limit"),
        label_podcastindex_key: i18n::tr(language, "options.label.podcastindex_key"),
        label_podcastindex_secret: i18n::tr(language, "options.label.podcastindex_secret"),
        label_podcastindex_signup: i18n::tr(language, "options.button.podcastindex_signup"),
        lang_it: i18n::tr(language, "options.lang.it"),
        lang_en: i18n::tr(language, "options.lang.en"),
        lang_es: i18n::tr(language, "options.lang.es"),
        lang_pt: i18n::tr(language, "options.lang.pt"),
        lang_vi: i18n::tr(language, "options.lang.vi"),
        marker_position_end: i18n::tr(language, "options.modified_marker_position.end"),
        marker_position_beginning: i18n::tr(language, "options.modified_marker_position.beginning"),
        open_new_tab: i18n::tr(language, "options.open.new_tab"),
        open_new_window: i18n::tr(language, "options.open.new_window"),
        engine_edge: i18n::tr(language, "options.engine.edge"),
        engine_sapi5: i18n::tr(language, "options.engine.sapi5"),

        split_none: i18n::tr(language, "options.split.none"),
        split_by_text: i18n::tr(language, "options.split.by_text"),
        split_parts: i18n::tr(language, "options.split.parts"),
        spellcheck_lang_follow: i18n::tr(language, "options.spellcheck.lang.follow"),
        spellcheck_lang_en_us: i18n::tr(language, "options.spellcheck.lang.en_us"),
        spellcheck_lang_en_gb: i18n::tr(language, "options.spellcheck.lang.en_gb"),
        spellcheck_lang_it: i18n::tr(language, "options.spellcheck.lang.it"),
        spellcheck_lang_es: i18n::tr(language, "options.spellcheck.lang.es"),
        spellcheck_lang_pt_br: i18n::tr(language, "options.spellcheck.lang.pt_br"),
        spellcheck_lang_fr: i18n::tr(language, "options.spellcheck.lang.fr"),
        spellcheck_lang_de: i18n::tr(language, "options.spellcheck.lang.de"),
        prompt_cmd: i18n::tr(language, "options.prompt.cmd"),
        prompt_powershell: i18n::tr(language, "options.prompt.powershell"),
        prompt_codex: i18n::tr(language, "options.prompt.codex"),
        ok: i18n::tr(language, "options.ok"),
        cancel: i18n::tr(language, "options.cancel"),
        voices_empty: i18n::tr(language, "options.voices.empty"),
    }
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.options_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(OPTIONS_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(options_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = options_labels(language);
    let title = to_wide(&labels.title);

    let dialog = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        520,
        710,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if dialog.0 != 0 {
        let _ = with_state(parent, |state| {
            state.options_dialog = dialog;
        });
        EnableWindow(parent, true);
        SetForegroundWindow(dialog);
        ensure_voice_lists_loaded(parent, language);
    }
}

pub unsafe fn refresh_voices(hwnd: HWND) {
    let (parent, combo_voice, combo_engine, checkbox) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_voice,
            state.combo_tts_engine,
            state.checkbox_multilingual,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();

    // Determine current engine from combo if possible, otherwise settings
    let engine_sel = SendMessageW(combo_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let engine = if engine_sel >= 0 {
        match engine_sel {
            1 => TtsEngine::Sapi5,

            _ => TtsEngine::Edge,
        }
    } else {
        settings.tts_engine
    };

    let voices = with_state(parent, |state| match engine {
        TtsEngine::Edge => state.edge_voices.clone(),
        TtsEngine::Sapi5 => state.sapi_voices.clone(),
    })
    .unwrap_or_default();

    let labels = options_labels(settings.language);
    let only_multilingual =
        SendMessageW(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;

    // Multilingual checkbox only relevant for Edge voices?
    // SAPI voices usually don't have "Multilingual" in name in the same way, but let's keep logic if applicable.
    // Generally assume SAPI voices are local and we list all.
    let filter_multilingual = if engine == TtsEngine::Edge {
        only_multilingual
    } else {
        false
    };

    // Disable multilingual checkbox for SAPI
    EnableWindow(checkbox, engine == TtsEngine::Edge);

    // If switching engine, we might not have the correct "selected" voice in settings yet if we haven't saved.
    // But we pass settings.tts_voice. If it's an ID from other engine, it won't match, so it selects default/first.
    populate_voice_combo(
        combo_voice,
        &voices,
        &settings.tts_voice,
        filter_multilingual,
        &labels,
    );
}

unsafe extern "system" fn options_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = options_labels(language);

            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

            let hwnd_tabs = CreateWindowExW(
                Default::default(),
                WC_TABCONTROLW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                20,
                10,
                460,
                28,
                hwnd,
                HMENU(OPTIONS_ID_TABS as isize),
                HINSTANCE(0),
                None,
            );
            let tab_labels = [
                labels.tab_general.clone(),
                labels.tab_voice.clone(),
                labels.tab_editor.clone(),
                labels.tab_audio.clone(),
            ];
            for (index, label) in tab_labels.iter().enumerate() {
                let mut text = to_wide(label);
                let mut item = TCITEMW {
                    mask: TCIF_TEXT,
                    pszText: PWSTR(text.as_mut_ptr()),
                    ..Default::default()
                };
                let _ = SendMessageW(
                    hwnd_tabs,
                    TCM_INSERTITEMW,
                    WPARAM(index),
                    LPARAM(&mut item as *mut _ as isize),
                );
            }
            let _ = SendMessageW(hwnd_tabs, TCM_SETCURSEL, WPARAM(0), LPARAM(0));

            let mut y = 50;
            let label_lang = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_language).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_lang = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                120,
                hwnd,
                HMENU(OPTIONS_ID_LANG as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_modified_marker_position = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_modified_marker_position).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_modified_marker_position = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                120,
                hwnd,
                HMENU(OPTIONS_ID_MODIFIED_MARKER_POSITION as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_open = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_open).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_open = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                120,
                hwnd,
                HMENU(OPTIONS_ID_OPEN as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_tts_engine = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_tts_engine).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_tts_engine = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                120,
                hwnd,
                HMENU(OPTIONS_ID_TTS_ENGINE as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_voice = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_voice).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_voice = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                140,
                hwnd,
                HMENU(OPTIONS_ID_VOICE as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let checkbox_multilingual = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_multilingual).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_MULTILINGUAL as isize),
                HINSTANCE(0),
                None,
            );
            y += 28;

            let label_tts_speed = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_tts_speed).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_tts_speed = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                140,
                hwnd,
                HMENU(OPTIONS_ID_TTS_SPEED as isize),
                HINSTANCE(0),
                None,
            );
            let edit_tts_speed = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                300,
                22,
                hwnd,
                HMENU(OPTIONS_ID_TTS_SPEED_EDIT as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_tts_pitch = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_tts_pitch).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_tts_pitch = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                140,
                hwnd,
                HMENU(OPTIONS_ID_TTS_PITCH as isize),
                HINSTANCE(0),
                None,
            );
            let edit_tts_pitch = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                300,
                22,
                hwnd,
                HMENU(OPTIONS_ID_TTS_PITCH_EDIT as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_tts_volume = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_tts_volume).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_tts_volume = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                140,
                hwnd,
                HMENU(OPTIONS_ID_TTS_VOLUME as isize),
                HINSTANCE(0),
                None,
            );
            let edit_tts_volume = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                300,
                22,
                hwnd,
                HMENU(OPTIONS_ID_TTS_VOLUME_EDIT as isize),
                HINSTANCE(0),
                None,
            );
            y += 36;

            let button_tts_preview = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_tts_preview).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                170,
                y,
                300,
                26,
                hwnd,
                HMENU(OPTIONS_ID_TTS_PREVIEW as isize),
                HINSTANCE(0),
                None,
            );
            y += 36;

            let label_audio_skip = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_audio_skip).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_audio_skip = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                140,
                hwnd,
                HMENU(OPTIONS_ID_AUDIO_SKIP as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let label_audio_split = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_audio_split).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_audio_split = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                140,
                hwnd,
                HMENU(OPTIONS_ID_AUDIO_SPLIT as isize),
                HINSTANCE(0),
                None,
            );
            y += 34;

            let label_audio_split_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_audio_split_text).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_audio_split_text = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                300,
                22,
                hwnd,
                HMENU(OPTIONS_ID_AUDIO_SPLIT_TEXT as isize),
                HINSTANCE(0),
                None,
            );
            y += 34;

            let checkbox_audio_split_requires_newline = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_audio_split_requires_newline).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_AUDIO_SPLIT_REQUIRE_NEWLINE as isize),
                HINSTANCE(0),
                None,
            );
            y += 24;

            let label_podcast_cache_limit = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_podcast_cache_limit).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_podcast_cache_limit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                80,
                22,
                hwnd,
                HMENU(OPTIONS_ID_PODCAST_CACHE_LIMIT as isize),
                HINSTANCE(0),
                None,
            );
            y += 30;

            let label_podcastindex_key = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_podcastindex_key).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_podcastindex_key = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                300,
                22,
                hwnd,
                HMENU(OPTIONS_ID_PODCASTINDEX_KEY as isize),
                HINSTANCE(0),
                None,
            );
            y += 30;

            let label_podcastindex_secret = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_podcastindex_secret).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_podcastindex_secret = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WINDOW_STYLE((ES_AUTOHSCROLL | ES_PASSWORD) as u32),
                170,
                y - 2,
                300,
                22,
                hwnd,
                HMENU(OPTIONS_ID_PODCASTINDEX_SECRET as isize),
                HINSTANCE(0),
                None,
            );
            y += 30;

            let button_podcastindex_signup = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_podcastindex_signup).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                170,
                y,
                300,
                26,
                hwnd,
                HMENU(OPTIONS_ID_PODCASTINDEX_SIGNUP as isize),
                HINSTANCE(0),
                None,
            );
            y += 34;

            let checkbox_tts_manual = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_tts_manual_tuning).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_TTS_MANUAL_TUNING as isize),
                HINSTANCE(0),
                None,
            );
            y += 24;

            let checkbox_split_on_newline = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_split_on_newline).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_SPLIT_ON_NEWLINE as isize),
                HINSTANCE(0),
                None,
            );
            y += 24;

            let checkbox_word_wrap = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_word_wrap).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_WORD_WRAP as isize),
                HINSTANCE(0),
                None,
            );
            y += 24;

            let checkbox_smart_quotes = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_smart_quotes).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_SMART_QUOTES as isize),
                HINSTANCE(0),
                None,
            );
            y += 26;

            let checkbox_spellcheck = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_spellcheck).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_SPELLCHECK_ENABLED as isize),
                HINSTANCE(0),
                None,
            );
            y += 26;

            let label_spellcheck_language = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_spellcheck_language).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_spellcheck_language = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                200,
                hwnd,
                HMENU(OPTIONS_ID_SPELLCHECK_LANGUAGE as isize),
                HINSTANCE(0),
                None,
            );
            y += 30;

            let label_wrap_width = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_wrap_width).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_wrap_width = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                80,
                22,
                hwnd,
                HMENU(OPTIONS_ID_WRAP_WIDTH as isize),
                HINSTANCE(0),
                None,
            );
            y += 30;

            let label_quote_prefix = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_quote_prefix).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let edit_quote_prefix = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                170,
                y - 2,
                120,
                22,
                hwnd,
                HMENU(OPTIONS_ID_QUOTE_PREFIX as isize),
                HINSTANCE(0),
                None,
            );
            y += 30;

            let checkbox_move_cursor = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_move_cursor).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_MOVE_CURSOR as isize),
                HINSTANCE(0),
                None,
            );
            y += 24;

            let checkbox_check_updates = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_check_updates).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_CHECK_UPDATES as isize),
                HINSTANCE(0),
                None,
            );
            y += 24;

            let checkbox_context_menu = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.label_context_menu).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                170,
                y,
                300,
                20,
                hwnd,
                HMENU(OPTIONS_ID_CONTEXT_MENU as isize),
                HINSTANCE(0),
                None,
            );
            y += 28;

            let label_prompt_program = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.label_prompt_program).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                y,
                140,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_prompt_program = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170,
                y - 2,
                300,
                120,
                hwnd,
                HMENU(OPTIONS_ID_PROMPT_PROGRAM as isize),
                HINSTANCE(0),
                None,
            );
            y += 40;

            let ok_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.ok).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                280,
                y,
                90,
                28,
                hwnd,
                HMENU(OPTIONS_ID_OK as isize),
                HINSTANCE(0),
                None,
            );
            let cancel_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.cancel).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                380,
                y,
                90,
                28,
                hwnd,
                HMENU(OPTIONS_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            for control in [
                hwnd_tabs,
                label_lang,
                combo_lang,
                label_modified_marker_position,
                combo_modified_marker_position,
                label_open,
                combo_open,
                label_tts_engine,
                combo_tts_engine,
                label_voice,
                combo_voice,
                label_tts_speed,
                combo_tts_speed,
                label_tts_pitch,
                combo_tts_pitch,
                label_tts_volume,
                combo_tts_volume,
                edit_tts_speed,
                edit_tts_pitch,
                edit_tts_volume,
                button_tts_preview,
                label_audio_skip,
                combo_audio_skip,
                label_audio_split,
                combo_audio_split,
                label_audio_split_text,
                edit_audio_split_text,
                checkbox_audio_split_requires_newline,
                label_podcast_cache_limit,
                edit_podcast_cache_limit,
                label_podcastindex_key,
                edit_podcastindex_key,
                label_podcastindex_secret,
                edit_podcastindex_secret,
                button_podcastindex_signup,
                checkbox_tts_manual,
                checkbox_multilingual,
                checkbox_split_on_newline,
                checkbox_word_wrap,
                checkbox_smart_quotes,
                checkbox_spellcheck,
                label_spellcheck_language,
                combo_spellcheck_language,
                label_wrap_width,
                edit_wrap_width,
                label_quote_prefix,
                edit_quote_prefix,
                checkbox_move_cursor,
                checkbox_check_updates,
                checkbox_context_menu,
                label_prompt_program,
                combo_prompt_program,
                ok_button,
                cancel_button,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let dialog_state = Box::new(OptionsDialogState {
                parent,
                hwnd_tabs,
                focus_initialized: false,
                label_language: label_lang,
                label_modified_marker_position,
                label_open,
                label_tts_engine,
                label_voice,
                label_tts_speed,
                label_tts_pitch,
                label_tts_volume,
                button_tts_preview,
                combo_lang,
                combo_modified_marker_position,
                combo_open,
                combo_tts_engine,
                combo_voice,
                combo_tts_speed,
                combo_tts_pitch,
                combo_tts_volume,
                edit_tts_speed,
                edit_tts_pitch,
                edit_tts_volume,
                checkbox_tts_manual,
                label_audio_skip,
                combo_audio_skip,
                label_audio_split,
                combo_audio_split,
                label_audio_split_text,
                edit_audio_split_text,
                checkbox_audio_split_requires_newline,
                label_podcast_cache_limit,
                edit_podcast_cache_limit,
                label_podcastindex_key,
                edit_podcastindex_key,
                label_podcastindex_secret,
                edit_podcastindex_secret,
                button_podcastindex_signup,
                checkbox_multilingual,
                checkbox_split_on_newline,
                checkbox_word_wrap,
                checkbox_smart_quotes,
                checkbox_spellcheck,
                label_spellcheck_language,
                combo_spellcheck_language,
                label_wrap_width,
                edit_wrap_width,
                label_quote_prefix,
                edit_quote_prefix,
                checkbox_move_cursor,
                checkbox_check_updates,
                checkbox_context_menu,
                label_prompt_program,
                combo_prompt_program,
                ok_button,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(dialog_state) as isize);
            initialize_options_dialog(hwnd);
            set_active_tab(hwnd, OPTIONS_TAB_GENERAL);
            LRESULT(0)
        }
        WM_NOTIFY => {
            let hdr = &*(lparam.0 as *const NMHDR);
            if hdr.idFrom == OPTIONS_ID_TABS as usize && hdr.code == TCN_SELCHANGE as u32 {
                let tabs = with_options_state(hwnd, |state| state.hwnd_tabs).unwrap_or(HWND(0));
                if tabs.0 != 0 {
                    let index = SendMessageW(tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                    set_active_tab(hwnd, index);
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            let code = (wparam.0 >> 16) as u32;
            match cmd_id {
                OPTIONS_ID_OK => {
                    apply_options_dialog(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_CANCEL | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_MULTILINGUAL => {
                    refresh_voices(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_TTS_PREVIEW => {
                    preview_voice(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_TTS_ENGINE => {
                    if code == CBN_SELCHANGE {
                        // When engine changes, verify if we need to load SAPI voices
                        let combo =
                            with_options_state(hwnd, |s| s.combo_tts_engine).unwrap_or(HWND(0));
                        let sel = SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        if sel == 1 {
                            // SAPI5
                            let parent = with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                            let has_sapi =
                                with_state(parent, |s| !s.sapi_voices.is_empty()).unwrap_or(false);
                            if !has_sapi {
                                let lang =
                                    with_state(parent, |s| s.settings.language).unwrap_or_default();
                                ensure_sapi_voices_loaded(parent, lang);
                            }
                        }

                        refresh_voices(hwnd);
                    }
                    LRESULT(0)
                }
                OPTIONS_ID_AUDIO_SPLIT => {
                    if code == CBN_SELCHANGE {
                        update_audio_split_text_visibility(hwnd);
                    }
                    LRESULT(0)
                }
                OPTIONS_ID_SPELLCHECK_ENABLED => {
                    update_spellcheck_language_visibility(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_TTS_MANUAL_TUNING => {
                    update_tts_manual_visibility(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_PODCASTINDEX_SIGNUP => {
                    open_podcastindex_signup();
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_TAB.0 as u32 {
                let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
                if ctrl_down {
                    let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
                    let tabs = with_options_state(hwnd, |state| state.hwnd_tabs).unwrap_or(HWND(0));
                    if tabs.0 != 0 {
                        let current =
                            SendMessageW(tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                        let mut next = if shift_down { current - 1 } else { current + 1 };
                        if next < 0 {
                            next = OPTIONS_TAB_COUNT - 1;
                        } else if next >= OPTIONS_TAB_COUNT {
                            next = 0;
                        }
                        let _ = SendMessageW(tabs, TCM_SETCURSEL, WPARAM(next as usize), LPARAM(0));
                        set_active_tab(hwnd, next);
                        return LRESULT(0);
                    }
                }
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let is_tts_combo = with_options_state(hwnd, |state| {
                    focus == state.combo_voice
                        || focus == state.combo_tts_speed
                        || focus == state.combo_tts_pitch
                        || focus == state.combo_tts_volume
                        || focus == state.edit_tts_speed
                        || focus == state.edit_tts_pitch
                        || focus == state.edit_tts_volume
                })
                .unwrap_or(false);
                if is_tts_combo {
                    apply_options_dialog(hwnd);
                    return LRESULT(0);
                }
            } else if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_DESTROY => {
            let parent = with_options_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                EnableWindow(parent, true);
                SetForegroundWindow(parent);
                SetFocus(parent);
                if let Some(edit) = crate::get_active_edit(parent) {
                    SetFocus(edit);
                }
                let _ = PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
                let _ = with_state(parent, |state| {
                    state.options_dialog = HWND(0);
                });
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_options_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut OptionsDialogState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe fn initialize_options_dialog(hwnd: HWND) {
    let (
        parent,
        combo_lang,
        combo_modified_marker_position,
        combo_open,
        combo_tts_engine,
        _combo_voice,
        combo_tts_speed,
        combo_tts_pitch,
        combo_tts_volume,
        edit_tts_speed,
        edit_tts_pitch,
        edit_tts_volume,
        combo_audio_skip,
        combo_audio_split,
        _label_audio_split_text,
        edit_audio_split_text,
        checkbox_audio_split_requires_newline,
        _label_podcast_cache_limit,
        edit_podcast_cache_limit,
        _label_podcastindex_key,
        edit_podcastindex_key,
        _label_podcastindex_secret,
        edit_podcastindex_secret,
        _button_podcastindex_signup,
        checkbox_tts_manual,
        checkbox_multilingual,
        checkbox_split_on_newline,
        checkbox_word_wrap,
        checkbox_smart_quotes,
        checkbox_spellcheck,
        combo_spellcheck_language,
        _label_wrap_width,
        edit_wrap_width,
        _label_quote_prefix,
        edit_quote_prefix,
        checkbox_move_cursor,
        checkbox_check_updates,
        checkbox_context_menu,
        _label_prompt_program,
        combo_prompt_program,
    ) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_lang,
            state.combo_modified_marker_position,
            state.combo_open,
            state.combo_tts_engine,
            state.combo_voice,
            state.combo_tts_speed,
            state.combo_tts_pitch,
            state.combo_tts_volume,
            state.edit_tts_speed,
            state.edit_tts_pitch,
            state.edit_tts_volume,
            state.combo_audio_skip,
            state.combo_audio_split,
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
            state.label_podcast_cache_limit,
            state.edit_podcast_cache_limit,
            state.label_podcastindex_key,
            state.edit_podcastindex_key,
            state.label_podcastindex_secret,
            state.edit_podcastindex_secret,
            state.button_podcastindex_signup,
            state.checkbox_tts_manual,
            state.checkbox_multilingual,
            state.checkbox_split_on_newline,
            state.checkbox_word_wrap,
            state.checkbox_smart_quotes,
            state.checkbox_spellcheck,
            state.combo_spellcheck_language,
            state.label_wrap_width,
            state.edit_wrap_width,
            state.label_quote_prefix,
            state.edit_quote_prefix,
            state.checkbox_move_cursor,
            state.checkbox_check_updates,
            state.checkbox_context_menu,
            state.label_prompt_program,
            state.combo_prompt_program,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let labels = options_labels(settings.language);

    let _ = SendMessageW(combo_lang, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(
        combo_lang,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.lang_it).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_lang,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.lang_en).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_lang,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.lang_es).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_lang,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.lang_pt).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_lang,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.lang_vi).as_ptr() as isize),
    );
    let lang_index = match settings.language {
        Language::Italian => 0,
        Language::English => 1,
        Language::Spanish => 2,
        Language::Portuguese => 3,
        Language::Vietnamese => 4,
    };
    let _ = SendMessageW(combo_lang, CB_SETCURSEL, WPARAM(lang_index), LPARAM(0));

    let _ = SendMessageW(combo_lang, CB_SETCURSEL, WPARAM(lang_index), LPARAM(0));

    let _ = SendMessageW(
        combo_modified_marker_position,
        CB_RESETCONTENT,
        WPARAM(0),
        LPARAM(0),
    );
    let _ = SendMessageW(
        combo_modified_marker_position,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.marker_position_end).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_modified_marker_position,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.marker_position_beginning).as_ptr() as isize),
    );
    let position_index = match settings.modified_marker_position {
        ModifiedMarkerPosition::Beginning => 1,
        _ => 0,
    };
    let _ = SendMessageW(
        combo_modified_marker_position,
        CB_SETCURSEL,
        WPARAM(position_index),
        LPARAM(0),
    );

    let _ = SendMessageW(combo_open, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(
        combo_open,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.open_new_tab).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_open,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.open_new_window).as_ptr() as isize),
    );
    let open_index = match settings.open_behavior {
        OpenBehavior::NewTab => 0,
        OpenBehavior::NewWindow => 1,
    };
    let _ = SendMessageW(combo_open, CB_SETCURSEL, WPARAM(open_index), LPARAM(0));

    let _ = SendMessageW(combo_tts_engine, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(
        combo_tts_engine,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.engine_edge).as_ptr() as isize),
    );
    let _ = SendMessageW(
        combo_tts_engine,
        CB_ADDSTRING,
        WPARAM(0),
        LPARAM(to_wide(&labels.engine_sapi5).as_ptr() as isize),
    );

    let engine_index = match settings.tts_engine {
        TtsEngine::Edge => 0,
        TtsEngine::Sapi5 => 1,
    };
    let _ = SendMessageW(
        combo_tts_engine,
        CB_SETCURSEL,
        WPARAM(engine_index),
        LPARAM(0),
    );

    let speed_items = [
        (
            i18n::tr(settings.language, "tts_tuning.speed.extremely_slow"),
            -100,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.speed.very_slow"),
            -60,
        ),
        (i18n::tr(settings.language, "tts_tuning.speed.slow"), -35),
        (
            i18n::tr(settings.language, "tts_tuning.speed.a_bit_slow"),
            -20,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.speed.slightly_slow"),
            -10,
        ),
        (i18n::tr(settings.language, "tts_tuning.speed.normal"), 0),
        (
            i18n::tr(settings.language, "tts_tuning.speed.slightly_fast"),
            10,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.speed.a_bit_fast"),
            20,
        ),
        (i18n::tr(settings.language, "tts_tuning.speed.fast"), 35),
        (
            i18n::tr(settings.language, "tts_tuning.speed.very_fast"),
            50,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.speed.super_fast"),
            100,
        ),
    ];
    let pitch_items = [
        (
            i18n::tr(settings.language, "tts_tuning.pitch.very_low"),
            -12,
        ),
        (i18n::tr(settings.language, "tts_tuning.pitch.low"), -10),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.a_bit_low"),
            -7,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.slightly_low"),
            -5,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.a_little_lower"),
            -2,
        ),
        (i18n::tr(settings.language, "tts_tuning.pitch.normal"), 0),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.a_little_higher"),
            2,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.slightly_high"),
            5,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.a_bit_high"),
            7,
        ),
        (i18n::tr(settings.language, "tts_tuning.pitch.high"), 9),
        (
            i18n::tr(settings.language, "tts_tuning.pitch.very_high"),
            12,
        ),
    ];
    let volume_items = [
        (
            i18n::tr(settings.language, "tts_tuning.volume.very_low"),
            25,
        ),
        (i18n::tr(settings.language, "tts_tuning.volume.low"), 40),
        (
            i18n::tr(settings.language, "tts_tuning.volume.a_bit_low"),
            55,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.volume.medium_low"),
            70,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.volume.slightly_low"),
            85,
        ),
        (i18n::tr(settings.language, "tts_tuning.volume.normal"), 100),
        (
            i18n::tr(settings.language, "tts_tuning.volume.slightly_high"),
            115,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.volume.medium_high"),
            130,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.volume.a_bit_high"),
            145,
        ),
        (i18n::tr(settings.language, "tts_tuning.volume.high"), 160),
        (
            i18n::tr(settings.language, "tts_tuning.volume.very_high"),
            180,
        ),
        (
            i18n::tr(settings.language, "tts_tuning.volume.maximum"),
            200,
        ),
    ];
    init_tts_combo(combo_tts_speed, &speed_items);
    init_tts_combo(combo_tts_pitch, &pitch_items);
    init_tts_combo(combo_tts_volume, &volume_items);
    select_combo_value(combo_tts_speed, settings.tts_rate);
    select_combo_value(combo_tts_pitch, settings.tts_pitch);
    select_combo_value(combo_tts_volume, settings.tts_volume);
    let _ = SetWindowTextW(
        edit_tts_speed,
        PCWSTR(to_wide(&settings.tts_rate.to_string()).as_ptr()),
    );
    let _ = SetWindowTextW(
        edit_tts_pitch,
        PCWSTR(to_wide(&settings.tts_pitch.to_string()).as_ptr()),
    );
    let _ = SetWindowTextW(
        edit_tts_volume,
        PCWSTR(to_wide(&settings.tts_volume.to_string()).as_ptr()),
    );
    let _ = SendMessageW(
        checkbox_tts_manual,
        BM_SETCHECK,
        WPARAM(if settings.tts_manual_tuning {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    update_tts_manual_visibility(hwnd);

    let _ = SendMessageW(
        checkbox_multilingual,
        BM_SETCHECK,
        WPARAM(if settings.tts_only_multilingual {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_audio_split_requires_newline,
        BM_SETCHECK,
        WPARAM(if settings.audiobook_split_text_requires_newline {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_split_on_newline,
        BM_SETCHECK,
        WPARAM(if settings.split_on_newline {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_word_wrap,
        BM_SETCHECK,
        WPARAM(if settings.word_wrap {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_smart_quotes,
        BM_SETCHECK,
        WPARAM(if settings.smart_quotes {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_spellcheck,
        BM_SETCHECK,
        WPARAM(if settings.spellcheck_enabled {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        combo_spellcheck_language,
        CB_RESETCONTENT,
        WPARAM(0),
        LPARAM(0),
    );
    let spellcheck_options = [
        (labels.spellcheck_lang_follow.clone(), "follow"),
        (labels.spellcheck_lang_en_us.clone(), "en-US"),
        (labels.spellcheck_lang_en_gb.clone(), "en-GB"),
        (labels.spellcheck_lang_it.clone(), "it-IT"),
        (labels.spellcheck_lang_es.clone(), "es-ES"),
        (labels.spellcheck_lang_pt_br.clone(), "pt-BR"),
        (labels.spellcheck_lang_fr.clone(), "fr-FR"),
        (labels.spellcheck_lang_de.clone(), "de-DE"),
    ];
    let mut selected_idx = 0;
    let current_val = if settings.spellcheck_language_mode
        == crate::settings::SpellcheckLanguageMode::FollowEditorLanguage
    {
        "follow"
    } else {
        &settings.spellcheck_fixed_language
    };

    for (i, (label, val)) in spellcheck_options.iter().enumerate() {
        let _ = SendMessageW(
            combo_spellcheck_language,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(label).as_ptr() as isize),
        );
        if *val == current_val {
            selected_idx = i;
        }
    }
    let _ = SendMessageW(
        combo_spellcheck_language,
        CB_SETCURSEL,
        WPARAM(selected_idx),
        LPARAM(0),
    );
    update_spellcheck_language_visibility(hwnd);
    let wrap_text = settings.wrap_width.to_string();
    let _ = SetWindowTextW(edit_wrap_width, PCWSTR(to_wide(&wrap_text).as_ptr()));
    let _ = SetWindowTextW(
        edit_quote_prefix,
        PCWSTR(to_wide(&settings.quote_prefix).as_ptr()),
    );
    let _ = SendMessageW(
        checkbox_move_cursor,
        BM_SETCHECK,
        WPARAM(if settings.move_cursor_during_reading {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_check_updates,
        BM_SETCHECK,
        WPARAM(if settings.check_updates_on_startup {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );
    let _ = SendMessageW(
        checkbox_context_menu,
        BM_SETCHECK,
        WPARAM(if settings.context_menu_open_with {
            BST_CHECKED.0 as usize
        } else {
            0
        }),
        LPARAM(0),
    );

    let _ = SendMessageW(combo_prompt_program, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let prompt_options = [
        labels.prompt_cmd.clone(),
        labels.prompt_powershell.clone(),
        labels.prompt_codex.clone(),
    ];
    for label in prompt_options.iter() {
        let _ = SendMessageW(
            combo_prompt_program,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(label).as_ptr() as isize),
        );
    }
    let program = settings.prompt_program.to_ascii_lowercase();
    let program_idx = if program.contains("powershell") {
        1
    } else if program.contains("codex") {
        2
    } else {
        0
    };
    let _ = SendMessageW(
        combo_prompt_program,
        CB_SETCURSEL,
        WPARAM(program_idx),
        LPARAM(0),
    );

    let _ = SendMessageW(combo_audio_skip, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let skip_options = [
        (10, "10 s"),
        (30, "30 s"),
        (60, "1 m"),
        (120, "2 m"),
        (300, "5 m"),
    ];
    let mut selected_idx = 2;
    for (secs, label) in skip_options.iter() {
        let idx = SendMessageW(
            combo_audio_skip,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(label).as_ptr() as isize),
        )
        .0 as usize;
        let _ = SendMessageW(
            combo_audio_skip,
            CB_SETITEMDATA,
            WPARAM(idx),
            LPARAM(*secs as isize),
        );
        if *secs == settings.audiobook_skip_seconds {
            selected_idx = idx;
        }
    }
    let _ = SendMessageW(
        combo_audio_skip,
        CB_SETCURSEL,
        WPARAM(selected_idx),
        LPARAM(0),
    );

    let _ = SendMessageW(combo_audio_split, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let split_options = [
        (0, labels.split_none.clone()),
        (AUDIOBOOK_SPLIT_BY_TEXT, labels.split_by_text.clone()),
        (2, format!("2 {}", labels.split_parts)),
        (4, format!("4 {}", labels.split_parts)),
        (6, format!("6 {}", labels.split_parts)),
        (8, format!("8 {}", labels.split_parts)),
    ];
    let mut selected_split_idx = 0;
    for (parts, label) in split_options.iter() {
        let idx = SendMessageW(
            combo_audio_split,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(label).as_ptr() as isize),
        )
        .0 as usize;
        let _ = SendMessageW(
            combo_audio_split,
            CB_SETITEMDATA,
            WPARAM(idx),
            LPARAM(*parts as isize),
        );
        if (settings.audiobook_split_by_text && *parts == AUDIOBOOK_SPLIT_BY_TEXT)
            || (!settings.audiobook_split_by_text && *parts == settings.audiobook_split)
        {
            selected_split_idx = idx;
        }
    }
    let _ = SendMessageW(
        combo_audio_split,
        CB_SETCURSEL,
        WPARAM(selected_split_idx),
        LPARAM(0),
    );

    let split_text_wide = to_wide(&settings.audiobook_split_text);
    let _ = SetWindowTextW(edit_audio_split_text, PCWSTR(split_text_wide.as_ptr()));
    update_audio_split_text_visibility(hwnd);

    let cache_limit_text = settings.podcast_cache_limit_mb.to_string();
    let _ = SetWindowTextW(
        edit_podcast_cache_limit,
        PCWSTR(to_wide(&cache_limit_text).as_ptr()),
    );
    let _ = SetWindowTextW(
        edit_podcastindex_key,
        PCWSTR(to_wide(&settings.podcast_index_api_key).as_ptr()),
    );
    let secret = crate::settings::decrypt_podcast_index_secret(&settings.podcast_index_api_secret)
        .unwrap_or_default();
    let _ = SetWindowTextW(edit_podcastindex_secret, PCWSTR(to_wide(&secret).as_ptr()));

    refresh_voices(hwnd);
}

unsafe fn populate_voice_combo(
    combo_voice: HWND,
    voices: &[VoiceInfo],
    selected: &str,
    only_multilingual: bool,
    labels: &OptionsLabels,
) {
    let _ = SendMessageW(combo_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    if voices.is_empty() {
        let label = &labels.voices_empty;
        // We could also check if it's loading, but SAPI loads fast.
        // For Edge, it might be loading.
        // We can check if "loading" logic is needed, but "voices_empty" is safe default.
        let _ = SendMessageW(
            combo_voice,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(label).as_ptr() as isize),
        );
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        return;
    }
    let mut selected_index: Option<usize> = None;
    let mut combo_index = 0usize;

    for (voice_index, voice) in voices.iter().enumerate() {
        if only_multilingual && !voice.is_multilingual {
            continue;
        }
        let label = format!("{} ({})", voice.short_name, voice.locale);
        let wide = to_wide(&label);
        let idx = SendMessageW(
            combo_voice,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(wide.as_ptr() as isize),
        )
        .0;
        if idx >= 0 {
            let _ = SendMessageW(
                combo_voice,
                CB_SETITEMDATA,
                WPARAM(idx as usize),
                LPARAM(voice_index as isize),
            );
            if voice.short_name == selected {
                selected_index = Some(combo_index);
            }
            combo_index += 1;
        }
    }

    if let Some(idx) = selected_index {
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
    } else if combo_index > 0 {
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
    }
}

fn init_tts_combo(hwnd: HWND, items: &[(String, i32)]) {
    unsafe {
        let _ = SendMessageW(hwnd, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for (label, value) in items {
            let idx = SendMessageW(
                hwnd,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            )
            .0 as usize;
            let _ = SendMessageW(hwnd, CB_SETITEMDATA, WPARAM(idx), LPARAM(*value as isize));
        }
    }
}

fn select_combo_value(hwnd: HWND, value: i32) {
    unsafe {
        let count = SendMessageW(
            hwnd,
            windows::Win32::UI::WindowsAndMessaging::CB_GETCOUNT,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        for i in 0..count {
            let data = SendMessageW(hwnd, CB_GETITEMDATA, WPARAM(i as usize), LPARAM(0)).0 as i32;
            if data == value {
                let _ = SendMessageW(hwnd, CB_SETCURSEL, WPARAM(i as usize), LPARAM(0));
                break;
            }
        }
    }
}

fn combo_value(hwnd: HWND) -> i32 {
    unsafe {
        let sel = SendMessageW(hwnd, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if sel < 0 {
            return 0;
        }
        SendMessageW(hwnd, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as i32
    }
}

const TTS_RATE_MIN: i32 = -100;
const TTS_RATE_MAX: i32 = 100;
const TTS_PITCH_MIN: i32 = -12;
const TTS_PITCH_MAX: i32 = 12;
const TTS_VOLUME_MIN: i32 = 25;
const TTS_VOLUME_MAX: i32 = 200;

fn read_tts_edit_value(edit: HWND, fallback: i32, min: i32, max: i32) -> i32 {
    unsafe {
        let len = GetWindowTextLengthW(edit);
        if len <= 0 {
            return fallback;
        }
        let mut buf = vec![0u16; (len + 1) as usize];
        let read = GetWindowTextW(edit, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        if let Ok(parsed) = text.trim().parse::<i32>() {
            parsed.clamp(min, max)
        } else {
            fallback
        }
    }
}

fn select_combo_nearest_value(hwnd: HWND, value: i32) {
    unsafe {
        let count = SendMessageW(hwnd, CB_GETCOUNT, WPARAM(0), LPARAM(0)).0;
        if count <= 0 {
            return;
        }
        let mut best_idx = 0;
        let mut best_diff = i32::MAX;
        for i in 0..count {
            let data = SendMessageW(hwnd, CB_GETITEMDATA, WPARAM(i as usize), LPARAM(0)).0 as i32;
            let diff = (data - value).abs();
            if diff < best_diff {
                best_diff = diff;
                best_idx = i;
            }
        }
        let _ = SendMessageW(hwnd, CB_SETCURSEL, WPARAM(best_idx as usize), LPARAM(0));
    }
}

unsafe fn update_tts_manual_visibility(hwnd: HWND) {
    let (checkbox, combo_speed, combo_pitch, combo_volume, edit_speed, edit_pitch, edit_volume) =
        match with_options_state(hwnd, |state| {
            (
                state.checkbox_tts_manual,
                state.combo_tts_speed,
                state.combo_tts_pitch,
                state.combo_tts_volume,
                state.edit_tts_speed,
                state.edit_tts_pitch,
                state.edit_tts_volume,
            )
        }) {
            Some(values) => values,
            None => return,
        };
    let manual =
        SendMessageW(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    if manual {
        let rate = combo_value(combo_speed);
        let pitch = combo_value(combo_pitch);
        let volume = combo_value(combo_volume);
        let _ = SetWindowTextW(edit_speed, PCWSTR(to_wide(&rate.to_string()).as_ptr()));
        let _ = SetWindowTextW(edit_pitch, PCWSTR(to_wide(&pitch.to_string()).as_ptr()));
        let _ = SetWindowTextW(edit_volume, PCWSTR(to_wide(&volume.to_string()).as_ptr()));
    } else {
        let rate = read_tts_edit_value(edit_speed, 0, TTS_RATE_MIN, TTS_RATE_MAX);
        let pitch = read_tts_edit_value(edit_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX);
        let volume = read_tts_edit_value(edit_volume, 100, TTS_VOLUME_MIN, TTS_VOLUME_MAX);
        select_combo_nearest_value(combo_speed, rate);
        select_combo_nearest_value(combo_pitch, pitch);
        select_combo_nearest_value(combo_volume, volume);
    }
    ShowWindow(combo_speed, if manual { SW_HIDE } else { SW_SHOW });
    ShowWindow(combo_pitch, if manual { SW_HIDE } else { SW_SHOW });
    ShowWindow(combo_volume, if manual { SW_HIDE } else { SW_SHOW });
    ShowWindow(edit_speed, if manual { SW_SHOW } else { SW_HIDE });
    ShowWindow(edit_pitch, if manual { SW_SHOW } else { SW_HIDE });
    ShowWindow(edit_volume, if manual { SW_SHOW } else { SW_HIDE });
    EnableWindow(combo_speed, !manual);
    EnableWindow(combo_pitch, !manual);
    EnableWindow(combo_volume, !manual);
    EnableWindow(edit_speed, manual);
    EnableWindow(edit_pitch, manual);
    EnableWindow(edit_volume, manual);
}

fn open_podcastindex_signup() {
    unsafe {
        let url = to_wide("https://api.podcastindex.org/signup");
        let _ = ShellExecuteW(
            HWND(0),
            w!("open"),
            PCWSTR(url.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        );
    }
}

unsafe fn preview_voice(hwnd: HWND) {
    let (
        parent,
        combo_tts_engine,
        combo_voice,
        combo_tts_speed,
        combo_tts_pitch,
        combo_tts_volume,
        edit_tts_speed,
        edit_tts_pitch,
        edit_tts_volume,
        checkbox_tts_manual,
    ) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_tts_engine,
            state.combo_voice,
            state.combo_tts_speed,
            state.combo_tts_pitch,
            state.combo_tts_volume,
            state.edit_tts_speed,
            state.edit_tts_pitch,
            state.edit_tts_volume,
            state.checkbox_tts_manual,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let (language, split_on_newline, dictionary) = with_state(parent, |state| {
        (
            state.settings.language,
            state.settings.split_on_newline,
            state.settings.dictionary.clone(),
        )
    })
    .unwrap_or((Language::Italian, true, Vec::new()));

    let text = i18n::tr(language, "tts.preview_text");
    if text.trim().is_empty() {
        return;
    }

    let engine_sel = SendMessageW(combo_tts_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let engine = if engine_sel == 1 {
        TtsEngine::Sapi5
    } else {
        TtsEngine::Edge
    };
    let voices = with_state(parent, |state| match engine {
        TtsEngine::Edge => state.edge_voices.clone(),
        TtsEngine::Sapi5 => state.sapi_voices.clone(),
    })
    .unwrap_or_default();

    let voice_sel = SendMessageW(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel < 0 {
        return;
    }
    let voice_index = SendMessageW(
        combo_voice,
        CB_GETITEMDATA,
        WPARAM(voice_sel as usize),
        LPARAM(0),
    )
    .0 as usize;
    if voice_index >= voices.len() {
        return;
    }
    let voice = voices[voice_index].short_name.clone();

    let manual = SendMessageW(checkbox_tts_manual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
        == BST_CHECKED.0;
    let rate = if manual {
        read_tts_edit_value(edit_tts_speed, 0, TTS_RATE_MIN, TTS_RATE_MAX)
    } else {
        combo_value(combo_tts_speed)
    };
    let pitch = if manual {
        read_tts_edit_value(edit_tts_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX)
    } else {
        combo_value(combo_tts_pitch)
    };
    let volume = if manual {
        read_tts_edit_value(edit_tts_volume, 100, TTS_VOLUME_MIN, TTS_VOLUME_MAX)
    } else {
        combo_value(combo_tts_volume)
    };
    let chunks = tts_engine::split_into_tts_chunks(&text, split_on_newline, &dictionary);

    match engine {
        TtsEngine::Edge => {
            tts_engine::start_tts_playback_with_chunks(
                parent, text, voice, chunks, 0, rate, pitch, volume,
            );
        }
        TtsEngine::Sapi5 => {
            tts_engine::stop_tts_playback(parent);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            let _ = with_state(parent, |state| {
                state.tts_session = Some(tts_engine::TtsSession {
                    id: state.tts_next_session_id,
                    command_tx,
                    cancel: cancel.clone(),
                    paused: false,
                    initial_caret_pos: 0,
                });
                state.tts_next_session_id += 1;
            });
            let chunk_strings: Vec<String> = chunks.into_iter().map(|c| c.text_to_read).collect();
            let _ = crate::sapi5_engine::play_sapi(
                chunk_strings,
                voice,
                rate,
                pitch,
                volume,
                cancel,
                command_rx,
            );
        }
    }
}

unsafe fn apply_options_dialog(hwnd: HWND) {
    let (
        parent,
        combo_lang,
        combo_modified_marker_position,
        combo_open,
        combo_tts_engine,
        combo_voice,
        combo_tts_speed,
        combo_tts_pitch,
        combo_tts_volume,
        edit_tts_speed,
        edit_tts_pitch,
        edit_tts_volume,
        combo_audio_skip,
        combo_audio_split,
        edit_audio_split_text,
        checkbox_audio_split_requires_newline,
        edit_podcast_cache_limit,
        edit_podcastindex_key,
        edit_podcastindex_secret,
        checkbox_tts_manual,
        checkbox_multilingual,
        checkbox_split_on_newline,
        checkbox_word_wrap,
        checkbox_smart_quotes,
        checkbox_spellcheck,
        combo_spellcheck_language,
        edit_wrap_width,
        edit_quote_prefix,
        checkbox_move_cursor,
        checkbox_check_updates,
        checkbox_context_menu,
        combo_prompt_program,
    ) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_lang,
            state.combo_modified_marker_position,
            state.combo_open,
            state.combo_tts_engine,
            state.combo_voice,
            state.combo_tts_speed,
            state.combo_tts_pitch,
            state.combo_tts_volume,
            state.edit_tts_speed,
            state.edit_tts_pitch,
            state.edit_tts_volume,
            state.combo_audio_skip,
            state.combo_audio_split,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
            state.edit_podcast_cache_limit,
            state.edit_podcastindex_key,
            state.edit_podcastindex_secret,
            state.checkbox_tts_manual,
            state.checkbox_multilingual,
            state.checkbox_split_on_newline,
            state.checkbox_word_wrap,
            state.checkbox_smart_quotes,
            state.checkbox_spellcheck,
            state.combo_spellcheck_language,
            state.edit_wrap_width,
            state.edit_quote_prefix,
            state.checkbox_move_cursor,
            state.checkbox_check_updates,
            state.checkbox_context_menu,
            state.combo_prompt_program,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let mut settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let old_language = settings.language;
    let old_marker_position = settings.modified_marker_position;
    let old_word_wrap = settings.word_wrap;
    let old_context_menu = settings.context_menu_open_with;
    let old_spellcheck_enabled = settings.spellcheck_enabled;
    let old_spellcheck_mode = settings.spellcheck_language_mode;
    let old_spellcheck_fixed_language = settings.spellcheck_fixed_language.clone();
    let (old_engine, old_voice, old_rate, old_pitch, old_volume, was_tts_active) =
        with_state(parent, |state| {
            (
                state.settings.tts_engine,
                state.settings.tts_voice.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
                state.tts_session.is_some(),
            )
        })
        .unwrap_or((
            settings.tts_engine,
            settings.tts_voice.clone(),
            settings.tts_rate,
            settings.tts_pitch,
            settings.tts_volume,
            false,
        ));

    let lang_sel = SendMessageW(combo_lang, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.language = match lang_sel {
        1 => Language::English,
        2 => Language::Spanish,
        3 => Language::Portuguese,
        4 => Language::Vietnamese,
        _ => Language::Italian,
    };

    let marker_sel = SendMessageW(
        combo_modified_marker_position,
        CB_GETCURSEL,
        WPARAM(0),
        LPARAM(0),
    )
    .0;
    settings.modified_marker_position = if marker_sel == 1 {
        ModifiedMarkerPosition::Beginning
    } else {
        ModifiedMarkerPosition::End
    };

    let open_sel = SendMessageW(combo_open, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.open_behavior = if open_sel == 1 {
        OpenBehavior::NewWindow
    } else {
        OpenBehavior::NewTab
    };

    let engine_sel = SendMessageW(combo_tts_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.tts_engine = match engine_sel {
        1 => TtsEngine::Sapi5,

        _ => TtsEngine::Edge,
    };

    settings.tts_manual_tuning =
        SendMessageW(checkbox_tts_manual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    if settings.tts_manual_tuning {
        settings.tts_rate = read_tts_edit_value(
            edit_tts_speed,
            settings.tts_rate,
            TTS_RATE_MIN,
            TTS_RATE_MAX,
        );
        settings.tts_pitch = read_tts_edit_value(
            edit_tts_pitch,
            settings.tts_pitch,
            TTS_PITCH_MIN,
            TTS_PITCH_MAX,
        );
        settings.tts_volume = read_tts_edit_value(
            edit_tts_volume,
            settings.tts_volume,
            TTS_VOLUME_MIN,
            TTS_VOLUME_MAX,
        );
    } else {
        settings.tts_rate = combo_value(combo_tts_speed);
        settings.tts_pitch = combo_value(combo_tts_pitch);
        settings.tts_volume = combo_value(combo_tts_volume);
    }

    settings.tts_only_multilingual =
        SendMessageW(checkbox_multilingual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    settings.audiobook_split_text_requires_newline = SendMessageW(
        checkbox_audio_split_requires_newline,
        BM_GETCHECK,
        WPARAM(0),
        LPARAM(0),
    )
    .0 as u32
        == BST_CHECKED.0;
    settings.split_on_newline =
        SendMessageW(checkbox_split_on_newline, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    settings.word_wrap = SendMessageW(checkbox_word_wrap, BM_GETCHECK, WPARAM(0), LPARAM(0)).0
        as u32
        == BST_CHECKED.0;
    settings.smart_quotes = SendMessageW(checkbox_smart_quotes, BM_GETCHECK, WPARAM(0), LPARAM(0)).0
        as u32
        == BST_CHECKED.0;
    settings.spellcheck_enabled =
        SendMessageW(checkbox_spellcheck, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    let spellcheck_sel = SendMessageW(
        combo_spellcheck_language,
        CB_GETCURSEL,
        WPARAM(0),
        LPARAM(0),
    )
    .0;
    if spellcheck_sel == 0 {
        settings.spellcheck_language_mode =
            crate::settings::SpellcheckLanguageMode::FollowEditorLanguage;
    } else {
        settings.spellcheck_language_mode = crate::settings::SpellcheckLanguageMode::FixedLanguage;
        let val = match spellcheck_sel {
            1 => "en-US",
            2 => "en-GB",
            3 => "it-IT",
            4 => "es-ES",
            5 => "pt-BR",
            6 => "fr-FR",
            7 => "de-DE",
            _ => "en-US",
        };
        settings.spellcheck_fixed_language = val.to_string();
    }

    let width_len = GetWindowTextLengthW(edit_wrap_width);
    if width_len >= 0 {
        let mut buf = vec![0u16; (width_len + 1) as usize];
        let read = GetWindowTextW(edit_wrap_width, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        if let Ok(parsed) = text.trim().parse::<u32>() {
            if parsed > 0 {
                settings.wrap_width = parsed;
            }
        }
    }
    let prefix_len = GetWindowTextLengthW(edit_quote_prefix);
    if prefix_len >= 0 {
        let mut buf = vec![0u16; (prefix_len + 1) as usize];
        let read = GetWindowTextW(edit_quote_prefix, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        settings.quote_prefix = text;
    }
    settings.move_cursor_during_reading =
        SendMessageW(checkbox_move_cursor, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    settings.check_updates_on_startup =
        SendMessageW(checkbox_check_updates, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    settings.context_menu_open_with =
        SendMessageW(checkbox_context_menu, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;

    let prompt_sel = SendMessageW(combo_prompt_program, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.prompt_program = match prompt_sel {
        1 => "powershell.exe".to_string(),
        2 => "codex".to_string(),
        _ => "cmd.exe".to_string(),
    };

    let voices = with_state(parent, |state| match settings.tts_engine {
        TtsEngine::Edge => state.edge_voices.clone(),
        TtsEngine::Sapi5 => state.sapi_voices.clone(),
    })
    .unwrap_or_default();

    let voice_sel = SendMessageW(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel >= 0 {
        let voice_index = SendMessageW(
            combo_voice,
            CB_GETITEMDATA,
            WPARAM(voice_sel as usize),
            LPARAM(0),
        )
        .0 as usize;
        if voice_index < voices.len() {
            settings.tts_voice = voices[voice_index].short_name.clone();
        }
    }

    let skip_sel = SendMessageW(combo_audio_skip, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if skip_sel >= 0 {
        let skip_secs = SendMessageW(
            combo_audio_skip,
            CB_GETITEMDATA,
            WPARAM(skip_sel as usize),
            LPARAM(0),
        )
        .0;
        settings.audiobook_skip_seconds = skip_secs as u32;
    }

    let split_sel = SendMessageW(combo_audio_split, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if split_sel >= 0 {
        let split_parts = SendMessageW(
            combo_audio_split,
            CB_GETITEMDATA,
            WPARAM(split_sel as usize),
            LPARAM(0),
        )
        .0;
        let split_parts = split_parts as u32;
        if split_parts == AUDIOBOOK_SPLIT_BY_TEXT {
            settings.audiobook_split_by_text = true;
            settings.audiobook_split = 0;
        } else {
            settings.audiobook_split_by_text = false;
            settings.audiobook_split = split_parts;
        }
    }

    let text_len = GetWindowTextLengthW(edit_audio_split_text);
    if text_len >= 0 {
        let mut buf = vec![0u16; (text_len + 1) as usize];
        let read = GetWindowTextW(edit_audio_split_text, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        settings.audiobook_split_text = text;
    }

    let cache_len = GetWindowTextLengthW(edit_podcast_cache_limit);
    if cache_len >= 0 {
        let mut buf = vec![0u16; (cache_len + 1) as usize];
        let read = GetWindowTextW(edit_podcast_cache_limit, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        if let Ok(parsed) = text.trim().parse::<u32>() {
            settings.podcast_cache_limit_mb = parsed.clamp(100, 2048);
        }
    }
    let key_len = GetWindowTextLengthW(edit_podcastindex_key);
    if key_len >= 0 {
        let mut buf = vec![0u16; (key_len + 1) as usize];
        let read = GetWindowTextW(edit_podcastindex_key, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        settings.podcast_index_api_key = text.trim().to_string();
    }
    let secret_len = GetWindowTextLengthW(edit_podcastindex_secret);
    if secret_len >= 0 {
        let mut buf = vec![0u16; (secret_len + 1) as usize];
        let read = GetWindowTextW(edit_podcastindex_secret, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        let trimmed = text.trim();
        if trimmed.is_empty() {
            settings.podcast_index_api_secret.clear();
        } else {
            settings.podcast_index_api_secret =
                crate::settings::encrypt_podcast_index_secret(trimmed);
        }
    }

    let _ = with_state(parent, |state| {
        state.settings = settings.clone();
    });
    let new_language = settings.language;
    let keep_default_copy = false;

    save_settings_with_default_copy(settings.clone(), keep_default_copy);
    if settings.spellcheck_enabled != old_spellcheck_enabled
        || settings.spellcheck_language_mode != old_spellcheck_mode
        || settings.spellcheck_fixed_language != old_spellcheck_fixed_language
    {
        crate::reset_spellcheck_state(parent);
    }
    if settings.context_menu_open_with != old_context_menu
        || (settings.context_menu_open_with && old_language != new_language)
    {
        sync_context_menu(&settings);
    }

    if old_language != new_language {
        rebuild_menus(parent);
    }
    if old_marker_position != settings.modified_marker_position {
        update_window_title(parent);
    }
    if old_word_wrap != settings.word_wrap {
        apply_word_wrap_to_all_edits(parent, settings.word_wrap);
    }
    refresh_voice_panel(parent);
    if was_tts_active
        && (old_engine != settings.tts_engine
            || old_voice != settings.tts_voice
            || old_rate != settings.tts_rate
            || old_pitch != settings.tts_pitch
            || old_volume != settings.tts_volume)
    {
        crate::restart_tts_from_current_offset(parent);
    }
    if parent.0 != 0 {
        let _ = PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
    }
    let _ = DestroyWindow(hwnd);
}

unsafe fn update_audio_split_text_visibility(hwnd: HWND) {
    let (
        combo_audio_split,
        label_audio_split_text,
        edit_audio_split_text,
        checkbox_audio_split_requires_newline,
    ) = match with_options_state(hwnd, |state| {
        (
            state.combo_audio_split,
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let split_sel = SendMessageW(combo_audio_split, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let selected = if split_sel >= 0 {
        let split_parts = SendMessageW(
            combo_audio_split,
            CB_GETITEMDATA,
            WPARAM(split_sel as usize),
            LPARAM(0),
        )
        .0 as u32;
        split_parts == AUDIOBOOK_SPLIT_BY_TEXT
    } else {
        false
    };

    let show = if selected { SW_SHOW } else { SW_HIDE };
    ShowWindow(label_audio_split_text, show);
    ShowWindow(edit_audio_split_text, show);
    ShowWindow(checkbox_audio_split_requires_newline, show);
    EnableWindow(edit_audio_split_text, selected);
    EnableWindow(checkbox_audio_split_requires_newline, selected);
}

unsafe fn update_spellcheck_language_visibility(hwnd: HWND) {
    let (checkbox_spellcheck, label_spellcheck_language, combo_spellcheck_language) =
        match with_options_state(hwnd, |state| {
            (
                state.checkbox_spellcheck,
                state.label_spellcheck_language,
                state.combo_spellcheck_language,
            )
        }) {
            Some(values) => values,
            None => return,
        };
    let spellcheck_enabled = SendMessageW(checkbox_spellcheck, BM_GETCHECK, WPARAM(0), LPARAM(0)).0
        as u32
        == BST_CHECKED.0;
    EnableWindow(label_spellcheck_language, spellcheck_enabled);
    EnableWindow(combo_spellcheck_language, spellcheck_enabled);
}

unsafe fn set_active_tab(hwnd: HWND, index: i32) {
    let focus_first = with_options_state(hwnd, |state| {
        if state.focus_initialized {
            false
        } else {
            state.focus_initialized = true;
            true
        }
    })
    .unwrap_or(false);
    let _ = with_options_state(hwnd, |state| {
        let show_general = index == OPTIONS_TAB_GENERAL;
        let show_voice = index == OPTIONS_TAB_VOICE;
        let show_editor = index == OPTIONS_TAB_EDITOR;
        let show_audio = index == OPTIONS_TAB_AUDIO;

        for control in [
            state.label_language,
            state.combo_lang,
            state.label_modified_marker_position,
            state.combo_modified_marker_position,
            state.label_open,
            state.combo_open,
            state.label_prompt_program,
            state.combo_prompt_program,
            state.checkbox_check_updates,
            state.checkbox_context_menu,
        ] {
            ShowWindow(control, if show_general { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_tts_engine,
            state.combo_tts_engine,
            state.label_voice,
            state.combo_voice,
            state.label_tts_speed,
            state.combo_tts_speed,
            state.label_tts_pitch,
            state.combo_tts_pitch,
            state.label_tts_volume,
            state.combo_tts_volume,
            state.edit_tts_speed,
            state.edit_tts_pitch,
            state.edit_tts_volume,
            state.button_tts_preview,
            state.checkbox_multilingual,
            state.checkbox_tts_manual,
            state.checkbox_split_on_newline,
        ] {
            ShowWindow(control, if show_voice { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.checkbox_word_wrap,
            state.checkbox_smart_quotes,
            state.checkbox_spellcheck,
            state.label_spellcheck_language,
            state.combo_spellcheck_language,
            state.label_wrap_width,
            state.edit_wrap_width,
            state.label_quote_prefix,
            state.edit_quote_prefix,
            state.checkbox_move_cursor,
        ] {
            ShowWindow(control, if show_editor { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_audio_skip,
            state.combo_audio_skip,
            state.label_audio_split,
            state.combo_audio_split,
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
            state.label_podcast_cache_limit,
            state.edit_podcast_cache_limit,
            state.label_podcastindex_key,
            state.edit_podcastindex_key,
            state.label_podcastindex_secret,
            state.edit_podcastindex_secret,
            state.button_podcastindex_signup,
        ] {
            ShowWindow(control, if show_audio { SW_SHOW } else { SW_HIDE });
        }
    });

    if index == OPTIONS_TAB_AUDIO {
        update_audio_split_text_visibility(hwnd);
    } else if index == OPTIONS_TAB_VOICE {
        update_tts_manual_visibility(hwnd);
    } else if let Some((label, edit, checkbox)) = with_options_state(hwnd, |state| {
        (
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
        )
    }) {
        ShowWindow(label, SW_HIDE);
        ShowWindow(edit, SW_HIDE);
        ShowWindow(checkbox, SW_HIDE);
        EnableWindow(edit, false);
        EnableWindow(checkbox, false);
    }

    if focus_first {
        focus_tab_first(hwnd, index);
    }
}

unsafe fn focus_tab_first(hwnd: HWND, index: i32) {
    let target = with_options_state(hwnd, |state| match index {
        OPTIONS_TAB_GENERAL => state.combo_lang,
        OPTIONS_TAB_VOICE => state.combo_tts_engine,
        OPTIONS_TAB_EDITOR => state.checkbox_word_wrap,
        OPTIONS_TAB_AUDIO => state.combo_audio_skip,
        _ => HWND(0),
    })
    .unwrap_or(HWND(0));

    if target.0 != 0 {
        SetFocus(target);
        let _ = PostMessageW(hwnd, WM_NEXTDLGCTL, WPARAM(target.0 as usize), LPARAM(1));
    }
}

pub(crate) fn ensure_voice_lists_loaded(hwnd: HWND, language: Language) {
    let (has_edge, has_sapi) = unsafe {
        with_state(hwnd, |state| {
            (!state.edge_voices.is_empty(), !state.sapi_voices.is_empty())
        })
    }
    .unwrap_or((false, false));

    if !has_edge {
        thread::spawn(move || {
            match fetch_voice_list() {
                Ok(list) => {
                    let payload = Box::new(list);
                    let _ = unsafe {
                        windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                            hwnd,
                            WM_TTS_VOICES_LOADED,
                            WPARAM(0),
                            LPARAM(Box::into_raw(payload) as isize),
                        )
                    };
                }
                Err(err) => {
                    // Log error but don't show message box for background load unless critical
                    // For now keeping it to avoid spamming user if offline
                    crate::log_debug(&format!("Failed to load Edge voices: {}", err));
                }
            }
        });
    }

    if !has_sapi {
        ensure_sapi_voices_loaded(hwnd, language);
    }
}

fn ensure_sapi_voices_loaded(hwnd: HWND, _language: Language) {
    thread::spawn(move || match crate::sapi5_engine::list_sapi_voices() {
        Ok(list) => {
            let payload = Box::new(list);
            let _ = unsafe {
                windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                    hwnd,
                    WM_TTS_SAPI_VOICES_LOADED,
                    WPARAM(0),
                    LPARAM(Box::into_raw(payload) as isize),
                )
            };
        }
        Err(err) => {
            crate::log_debug(&format!("Failed to load SAPI voices: {}", err));
        }
    });
}

fn fetch_voice_list() -> Result<Vec<VoiceInfo>, String> {
    let url = format!(
        "{}?trustedclienttoken={}",
        VOICE_LIST_URL, TRUSTED_CLIENT_TOKEN
    );
    let resp = reqwest::blocking::get(url).map_err(|err| err.to_string())?;
    let value: serde_json::Value = resp.json().map_err(|err| err.to_string())?;
    let Some(voices) = value.as_array() else {
        return Err("Risposta non valida".to_string());
    };

    let mut results = Vec::new();
    for voice in voices {
        let short_name = voice["ShortName"].as_str().unwrap_or("").to_string();
        if short_name.is_empty() {
            continue;
        }
        let locale = voice["Locale"].as_str().unwrap_or("").to_string();
        let is_multilingual = short_name.contains("Multilingual");
        results.push(VoiceInfo {
            short_name,
            locale,
            is_multilingual,
        });
    }
    results.sort_by(|a, b| a.short_name.cmp(&b.short_name));
    Ok(results)
}
