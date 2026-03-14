use crate::accessibility::{handle_accessibility, to_wide};
use crate::app_windows::interpreter_select_window;
use crate::editor_manager::{
    apply_indent_settings_to_all_edits, apply_word_wrap_to_all_edits, insert_voice_tag_at_caret,
    update_window_title,
};
use crate::settings::{
    AudiobookPartNamingMode, DEFAULT_VOICE_PROFILE_NAME, Language, ListDateDisplayMode,
    ListTimeDisplayMode, ModifiedMarkerPosition, OpenBehavior, PodcastDeleteConfirmMode,
    RssDeleteConfirmMode, RssPodcastUnreadLabelPosition, ShortcutBinding, ShortcutSettings,
    SubtitleReadMode, TRUSTED_CLIENT_TOKEN, TtsEngine, VOICE_LIST_URL, VoiceInfo, VoiceProfile,
    format_shortcut, save_settings_with_default_copy, sync_context_menu, sync_start_menu_shortcuts,
    voice_profile_from_settings_fields,
};
use crate::{i18n, rebuild_menus, refresh_voice_panel, tts_engine, with_state};
use reqwest::blocking::Client;
use reqwest::header::USER_AGENT;
use std::process::Command;
use std::sync::Arc;
use std::sync::atomic::AtomicBool;
use std::thread;
use std::time::Duration;
use tokio::sync::mpsc;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{
    BST_CHECKED, NMHDR, SetScrollInfo, ShowScrollBar, TCIF_TEXT, TCITEMW, TCM_GETCURSEL,
    TCM_INSERTITEMW, TCM_SETCURSEL, TCN_SELCHANGE, WC_BUTTON, WC_COMBOBOXW, WC_EDIT, WC_STATIC,
    WC_TABCONTROLW,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetAsyncKeyState, GetFocus, GetKeyState, SetFocus, VK_CONTROL, VK_ESCAPE,
    VK_LCONTROL, VK_LMENU, VK_LSHIFT, VK_MENU, VK_RCONTROL, VK_RETURN, VK_RMENU, VK_RSHIFT,
    VK_SHIFT, VK_SPACE, VK_TAB,
};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCOUNT,
    CB_GETCURSEL, CB_GETDROPPEDSTATE, CB_GETITEMDATA, CB_RESETCONTENT, CB_SETCURSEL,
    CB_SETITEMDATA, CBN_SELCHANGE, CBS_DROPDOWNLIST, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW,
    DefWindowProcW, ES_AUTOHSCROLL, ES_PASSWORD, ES_READONLY, GWLP_USERDATA, GetClassNameW,
    GetClientRect, GetParent, GetScrollInfo, GetWindowLongPtrW, GetWindowTextLengthW,
    GetWindowTextW, HMENU, IDC_ARROW, LoadCursorW, MB_ICONWARNING, MB_OK, MSG, MessageBoxW,
    MoveWindow, PostMessageW, RegisterClassW, SB_BOTTOM, SB_LINEDOWN, SB_LINEUP, SB_PAGEDOWN,
    SB_PAGEUP, SB_THUMBPOSITION, SB_THUMBTRACK, SB_TOP, SB_VERT, SCROLLINFO, SIF_PAGE, SIF_POS,
    SIF_RANGE, SIF_TRACKPOS, SW_HIDE, SW_SHOW, SW_SHOWNORMAL, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, SetWindowTextW, ShowWindow, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_MOUSEWHEEL, WM_NCDESTROY, WM_NEXTDLGCTL, WM_NOTIFY,
    WM_SETFONT, WM_VSCROLL, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, PWSTR, w};

const OPTIONS_CLASS_NAME: &str = "SonarpadOptions";
const OPTIONS_ID_LANG: usize = 6001;
const OPTIONS_ID_MODIFIED_MARKER_POSITION: usize = 6023;
const OPTIONS_ID_OPEN: usize = 6002;
const OPTIONS_ID_TTS_ENGINE: usize = 6012;
const OPTIONS_ID_TTS_VOICE_LANGUAGE: usize = 6075;
const OPTIONS_ID_VOICE: usize = 6003;
const OPTIONS_ID_MULTILINGUAL: usize = 6004;
const OPTIONS_ID_SPLIT_ON_NEWLINE: usize = 6007;
const OPTIONS_ID_WORD_WRAP: usize = 6008;
const OPTIONS_ID_SMART_QUOTES: usize = 6025;
const OPTIONS_ID_STRIP_MARKDOWN_KEEP_BULLETS: usize = 6039;
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
const OPTIONS_ID_DIALOGUE_VOICE_RATE_EDIT: usize = 6098;
const OPTIONS_ID_DIALOGUE_VOICE_PITCH_EDIT: usize = 6099;
const OPTIONS_ID_DIALOGUE_VOICE_VOLUME_EDIT: usize = 6100;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_RATE_EDIT: usize = 6101;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PITCH_EDIT: usize = 6102;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_VOLUME_EDIT: usize = 6103;
const OPTIONS_ID_SHOW_MEDIA_SAVE_CONFIRMATION: usize = 6104;
const OPTIONS_ID_AUDIO_SKIP: usize = 6010;
const OPTIONS_ID_AUDIO_SPLIT: usize = 6011;
const OPTIONS_ID_AUDIOBOOK_SAVE_FOLDER: usize = 6064;
const OPTIONS_ID_AUDIOBOOK_SAVE_FOLDER_BROWSE: usize = 6065;
const OPTIONS_ID_DEFAULT_SAVE_FOLDER_KIND: usize = 6110;
const OPTIONS_ID_AUDIO_SPLIT_MINUTES: usize = 6058;
const OPTIONS_ID_AUDIO_SPLIT_START_NUMBER: usize = 6059;
const OPTIONS_ID_AUDIOBOOK_PART_NAMING: usize = 6120;
const OPTIONS_ID_AUDIO_SPLIT_TEXT: usize = 6013;
const OPTIONS_ID_AUDIO_SPLIT_PARTS_COUNT: usize = 6087;
const OPTIONS_ID_AUDIO_SPLIT_REQUIRE_NEWLINE: usize = 6016;
const OPTIONS_ID_AUDIO_SPLIT_EPUB_CHAPTERS: usize = 6063;
const OPTIONS_ID_SUBTITLE_MODE: usize = 6046;
const OPTIONS_ID_SUBTITLE_DUCKING: usize = 6045;
const OPTIONS_ID_SUBTITLE_OFFSET: usize = 6047;
const OPTIONS_ID_PODCAST_CACHE_LIMIT: usize = 6030;
const OPTIONS_ID_PODCASTINDEX_KEY: usize = 6035;
const OPTIONS_ID_PODCASTINDEX_SECRET: usize = 6036;
const OPTIONS_ID_PODCASTINDEX_SIGNUP: usize = 6037;
const OPTIONS_ID_WHISPER_MODEL: usize = 6111;
const OPTIONS_ID_WHISPER_CUDA: usize = 6112;
const OPTIONS_ID_WHISPER_KEEP_ORIGINAL_LANGUAGE: usize = 6113;
const OPTIONS_ID_WHISPER_INCLUDE_TIMESTAMPS: usize = 6114;
const OPTIONS_ID_DICTATION_MICROPHONE: usize = 6121;
const OPTIONS_ID_DICTIONARY_TRANSLATION: usize = 6038;
const OPTIONS_ID_WIKIPEDIA_LANGUAGE: usize = 6040;
const OPTIONS_ID_INTERPRETER_PATH: usize = 6041;
const OPTIONS_ID_INTERPRETER_BROWSE: usize = 6042;
const OPTIONS_ID_INTERPRETER_SEARCH: usize = 6043;
const OPTIONS_ID_WRAP_WIDTH: usize = 6017;
const OPTIONS_ID_QUOTE_PREFIX: usize = 6018;
const OPTIONS_ID_INDENT_MODE: usize = 6060;
const OPTIONS_ID_TAB_WIDTH: usize = 6061;
const OPTIONS_ID_SPACE_WIDTH: usize = 6062;
const OPTIONS_ID_CHECK_UPDATES: usize = 6015;
const OPTIONS_ID_CHECK_BETA_UPDATES: usize = 6105;
const OPTIONS_ID_SEND_CRASH_REPORTS: usize = 6048;
const OPTIONS_ID_USE_LEGACY_NAME: usize = 6057;
const OPTIONS_ID_CONFIRM_DELETE_RSS_MODE: usize = 6066;
const OPTIONS_ID_ANNOUNCE_UNREAD_RSS_PODCAST: usize = 6067;
const OPTIONS_ID_UNREAD_LABEL_POSITION: usize = 6078;
const OPTIONS_ID_CONFIRM_DELETE_PODCAST_MODE: usize = 6068;
const OPTIONS_ID_RSS_QUICK_COPY_MODE: usize = 6069;
const OPTIONS_ID_RSS_SHOW_ARTICLE_PREVIEW: usize = 6119;
const OPTIONS_ID_RSS_DATE_DISPLAY: usize = 6083;
const OPTIONS_ID_RSS_TIME_DISPLAY: usize = 6084;
const OPTIONS_ID_PODCAST_DATE_DISPLAY: usize = 6085;
const OPTIONS_ID_PODCAST_TIME_DISPLAY: usize = 6086;
const OPTIONS_ID_MANAGE_ASSOCIATIONS: usize = 6044;
const OPTIONS_ID_PROMPT_PROGRAM: usize = 6019;
const OPTIONS_ID_NETWORK_PROXY: usize = 6075;
const OPTIONS_ID_NETWORK_PROXY_USERNAME: usize = 6076;
const OPTIONS_ID_NETWORK_PROXY_PASSWORD: usize = 6077;
const OPTIONS_ID_TABS: usize = 6024;
const OPTIONS_ID_USE_DIALOGUE_VOICE: usize = 6049;
const OPTIONS_ID_DIALOGUE_VOICE: usize = 6050;
const OPTIONS_ID_DIALOGUE_VOICE_PREVIEW: usize = 6051;
const OPTIONS_ID_DIALOGUE_VOICE_RATE: usize = 6052;
const OPTIONS_ID_DIALOGUE_VOICE_PITCH: usize = 6053;
const OPTIONS_ID_DIALOGUE_VOICE_VOLUME: usize = 6054;
const OPTIONS_ID_DIALOGUE_TTS_ENGINE: usize = 6055;
const OPTIONS_ID_DIALOGUE_VOICE_LANGUAGE: usize = 6082;
const OPTIONS_ID_TTS_INSERT_TAG: usize = 6056;
const OPTIONS_ID_DIALOGUE_OPEN_QUOTE: usize = 6079;
const OPTIONS_ID_DIALOGUE_CLOSE_QUOTE: usize = 6080;
const OPTIONS_ID_DIALOGUE_ALLOW_MULTILINE: usize = 6081;
const OPTIONS_ID_DIALOGUE_USE_SECONDARY_VOICE: usize = 6088;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE: usize = 6089;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_RATE: usize = 6090;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PITCH: usize = 6091;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_VOLUME: usize = 6092;
const OPTIONS_ID_DIALOGUE_SECONDARY_TTS_ENGINE: usize = 6093;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_LANGUAGE: usize = 6094;
const OPTIONS_ID_DIALOGUE_MULTILINGUAL: usize = 6095;
const OPTIONS_ID_DIALOGUE_SECONDARY_MULTILINGUAL: usize = 6096;
const OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PREVIEW: usize = 6097;
const OPTIONS_ID_SHORTCUT_ACTION: usize = 6070;
const OPTIONS_ID_SHORTCUT_VALUE: usize = 6071;
const OPTIONS_ID_SHORTCUT_CHANGE: usize = 6072;
const OPTIONS_ID_SHORTCUT_RESET: usize = 6073;
const OPTIONS_ID_SHORTCUT_RESET_ALL: usize = 6074;
const OPTIONS_ID_VOICE_PROFILE: usize = 6115;
const OPTIONS_ID_RENAME_VOICE_PROFILE: usize = 6116;
const OPTIONS_ID_ADD_VOICE_PROFILE: usize = 6117;
const OPTIONS_ID_DELETE_VOICE_PROFILE: usize = 6118;

const OPTIONS_ID_OK: usize = 6005;
const OPTIONS_ID_CANCEL: usize = 6006;

const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
const WM_TTS_SAPI_VOICES_LOADED: u32 = WM_APP + 8;
const AUDIOBOOK_SPLIT_BY_TEXT: u32 = u32::MAX;
const AUDIOBOOK_SPLIT_BY_TIME: u32 = u32::MAX - 1;
const AUDIOBOOK_SPLIT_BY_PARTS: u32 = u32::MAX - 2;

const OPTIONS_TAB_GENERAL: i32 = 0;
const OPTIONS_TAB_VOICE: i32 = 1;
const OPTIONS_TAB_EDITOR: i32 = 2;
const OPTIONS_TAB_AUDIO: i32 = 3;
const OPTIONS_TAB_RSS_PODCAST: i32 = 4;
const OPTIONS_TAB_AI_TRANSCRIPTION: i32 = 5;
const OPTIONS_TAB_SHORTCUTS: i32 = 6;
const OPTIONS_TAB_COUNT: i32 = 7;
const OPTIONS_CONTENT_TOP: i32 = 50;
const OPTIONS_MARGIN_X: i32 = 20;
const OPTIONS_LABEL_WIDTH: i32 = 245;
const OPTIONS_CONTROL_X: i32 = OPTIONS_MARGIN_X + OPTIONS_LABEL_WIDTH + 16;
const OPTIONS_CONTROL_WIDTH: i32 = 350;
const OPTIONS_ROW_HEIGHT: i32 = 32;
const OPTIONS_ROW_HEIGHT_COMPACT: i32 = 26;
const OPTIONS_CHECKBOX_HEIGHT: i32 = 22;
const OPTIONS_BUTTON_HEIGHT: i32 = 28;
const OPTIONS_COMBO_HEIGHT: i32 = 28;
const OPTIONS_COMBO_DROPDOWN_HEIGHT: i32 = 180;
const OPTIONS_EDIT_HEIGHT: i32 = 24;
const OPTIONS_SECTION_GAP: i32 = 12;
const OPTIONS_SCROLL_LINE: i32 = 32;
const OPTIONS_CONTENT_BOTTOM_GAP: i32 = 40;
const OPTIONS_WHEEL_DELTA: i32 = 120;
const DEFAULT_SAVE_FOLDER_AUDIOBOOK: u32 = 0;
const DEFAULT_SAVE_FOLDER_MEDIA: u32 = 1;
const DEFAULT_SAVE_FOLDER_DOCUMENTS: u32 = 2;

fn proxy_is_valid(proxy_url: &str, username: &str, password: &str) -> Result<(), String> {
    let mut proxy = reqwest::Proxy::all(proxy_url).map_err(|e| e.to_string())?;
    if !username.trim().is_empty() {
        proxy = proxy.basic_auth(username.trim(), password.trim());
    }
    let client = Client::builder()
        .proxy(proxy)
        .connect_timeout(Duration::from_secs(8))
        .timeout(Duration::from_secs(15))
        .build()
        .map_err(|e| e.to_string())?;
    let response = client
        .get("https://it.wikipedia.org/w/api.php?action=query&meta=siteinfo&format=json")
        .header(
            USER_AGENT,
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Sonarpad",
        )
        .send()
        .map_err(|e| e.to_string())?;
    if response.status().is_success() {
        Ok(())
    } else {
        Err(format!("HTTP {}", response.status()))
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ShortcutAction {
    ReadPauseResume,
    ReadStart,
    ReadStop,
    ExecuteFile,
    Audiobook,
    BatchAudiobooks,
    RecordPodcast,
    Dictation,
    ConvertAudio,
    OpenRss,
    OpenPodcasts,
    OpenDictionary,
    OpenOptions,
    OpenTerminal,
    ImportWikipedia,
    ImportYoutube,
    Find,
    QuoteLines,
    UnquoteLines,
    MediaPrev,
    MediaNext,
    ChapterPrev,
    ChapterNext,
}

impl ShortcutAction {
    const ALL: [ShortcutAction; 23] = [
        ShortcutAction::ReadPauseResume,
        ShortcutAction::ReadStart,
        ShortcutAction::ReadStop,
        ShortcutAction::ExecuteFile,
        ShortcutAction::Audiobook,
        ShortcutAction::BatchAudiobooks,
        ShortcutAction::RecordPodcast,
        ShortcutAction::Dictation,
        ShortcutAction::ConvertAudio,
        ShortcutAction::OpenRss,
        ShortcutAction::OpenPodcasts,
        ShortcutAction::OpenDictionary,
        ShortcutAction::OpenOptions,
        ShortcutAction::OpenTerminal,
        ShortcutAction::ImportWikipedia,
        ShortcutAction::ImportYoutube,
        ShortcutAction::Find,
        ShortcutAction::QuoteLines,
        ShortcutAction::UnquoteLines,
        ShortcutAction::MediaPrev,
        ShortcutAction::MediaNext,
        ShortcutAction::ChapterPrev,
        ShortcutAction::ChapterNext,
    ];
}

fn shortcut_action_label(language: Language, action: ShortcutAction) -> String {
    match action {
        ShortcutAction::ChapterPrev => i18n::tr(language, "playback.chapter_prev"),
        ShortcutAction::ChapterNext => i18n::tr(language, "playback.chapter_next"),
        _ => i18n::tr(language, shortcut_action_i18n_key(action)),
    }
}

fn shortcut_action_i18n_key(action: ShortcutAction) -> &'static str {
    match action {
        ShortcutAction::ReadPauseResume => "options.shortcuts.action.read_pause_resume",
        ShortcutAction::ReadStart => "options.shortcuts.action.read_start",
        ShortcutAction::ReadStop => "options.shortcuts.action.read_stop",
        ShortcutAction::ExecuteFile => "options.shortcuts.action.execute_file",
        ShortcutAction::Audiobook => "options.shortcuts.action.audiobook",
        ShortcutAction::BatchAudiobooks => "options.shortcuts.action.batch_audiobooks",
        ShortcutAction::RecordPodcast => "options.shortcuts.action.record_podcast",
        ShortcutAction::Dictation => "options.shortcuts.action.dictation",
        ShortcutAction::ConvertAudio => "options.shortcuts.action.convert_audio",
        ShortcutAction::OpenRss => "options.shortcuts.action.open_rss",
        ShortcutAction::OpenPodcasts => "options.shortcuts.action.open_podcasts",
        ShortcutAction::OpenDictionary => "options.shortcuts.action.open_dictionary",
        ShortcutAction::OpenOptions => "options.shortcuts.action.open_options",
        ShortcutAction::OpenTerminal => "options.shortcuts.action.open_terminal",
        ShortcutAction::ImportWikipedia => "options.shortcuts.action.import_wikipedia",
        ShortcutAction::ImportYoutube => "options.shortcuts.action.import_youtube",
        ShortcutAction::Find => "options.shortcuts.action.find",
        ShortcutAction::QuoteLines => "options.shortcuts.action.quote_lines",
        ShortcutAction::UnquoteLines => "options.shortcuts.action.unquote_lines",
        ShortcutAction::MediaPrev => "options.shortcuts.action.media_prev",
        ShortcutAction::MediaNext => "options.shortcuts.action.media_next",
        ShortcutAction::ChapterPrev => "playback.chapter_prev",
        ShortcutAction::ChapterNext => "playback.chapter_next",
    }
}

fn shortcut_tab_title(language: Language) -> String {
    i18n::tr(language, "options.shortcuts.tab")
}

fn shortcuts_label_action(language: Language) -> String {
    i18n::tr(language, "options.shortcuts.label_action")
}

fn shortcuts_label_value(language: Language) -> String {
    i18n::tr(language, "options.shortcuts.label_combination")
}

fn shortcuts_change_label(language: Language) -> String {
    i18n::tr(language, "options.shortcuts.button_change")
}

fn shortcuts_reset_label(language: Language) -> String {
    i18n::tr(language, "options.shortcuts.button_default")
}

fn shortcuts_reset_all_label(language: Language) -> String {
    i18n::tr(language, "options.shortcuts.button_reset_all")
}

fn shortcut_binding_for_action(
    settings: &ShortcutSettings,
    action: ShortcutAction,
) -> ShortcutBinding {
    match action {
        ShortcutAction::ReadPauseResume => settings.read_pause_resume,
        ShortcutAction::ReadStart => settings.read_start,
        ShortcutAction::ReadStop => settings.read_stop,
        ShortcutAction::ExecuteFile => settings.execute_file,
        ShortcutAction::Audiobook => settings.audiobook,
        ShortcutAction::BatchAudiobooks => settings.batch_audiobooks,
        ShortcutAction::RecordPodcast => settings.record_podcast,
        ShortcutAction::Dictation => settings.dictation,
        ShortcutAction::ConvertAudio => settings.convert_audio,
        ShortcutAction::OpenRss => settings.open_rss,
        ShortcutAction::OpenPodcasts => settings.open_podcasts,
        ShortcutAction::OpenDictionary => settings.open_dictionary,
        ShortcutAction::OpenOptions => settings.open_options,
        ShortcutAction::OpenTerminal => settings.open_terminal,
        ShortcutAction::ImportWikipedia => settings.import_wikipedia,
        ShortcutAction::ImportYoutube => settings.import_youtube,
        ShortcutAction::Find => settings.find,
        ShortcutAction::QuoteLines => settings.quote_lines,
        ShortcutAction::UnquoteLines => settings.unquote_lines,
        ShortcutAction::MediaPrev => settings.media_prev,
        ShortcutAction::MediaNext => settings.media_next,
        ShortcutAction::ChapterPrev => settings.chapter_prev,
        ShortcutAction::ChapterNext => settings.chapter_next,
    }
}

fn set_shortcut_binding_for_action(
    settings: &mut ShortcutSettings,
    action: ShortcutAction,
    binding: ShortcutBinding,
) {
    match action {
        ShortcutAction::ReadPauseResume => settings.read_pause_resume = binding,
        ShortcutAction::ReadStart => settings.read_start = binding,
        ShortcutAction::ReadStop => settings.read_stop = binding,
        ShortcutAction::ExecuteFile => settings.execute_file = binding,
        ShortcutAction::Audiobook => settings.audiobook = binding,
        ShortcutAction::BatchAudiobooks => settings.batch_audiobooks = binding,
        ShortcutAction::RecordPodcast => settings.record_podcast = binding,
        ShortcutAction::Dictation => settings.dictation = binding,
        ShortcutAction::ConvertAudio => settings.convert_audio = binding,
        ShortcutAction::OpenRss => settings.open_rss = binding,
        ShortcutAction::OpenPodcasts => settings.open_podcasts = binding,
        ShortcutAction::OpenDictionary => settings.open_dictionary = binding,
        ShortcutAction::OpenOptions => settings.open_options = binding,
        ShortcutAction::OpenTerminal => settings.open_terminal = binding,
        ShortcutAction::ImportWikipedia => settings.import_wikipedia = binding,
        ShortcutAction::ImportYoutube => settings.import_youtube = binding,
        ShortcutAction::Find => settings.find = binding,
        ShortcutAction::QuoteLines => settings.quote_lines = binding,
        ShortcutAction::UnquoteLines => settings.unquote_lines = binding,
        ShortcutAction::MediaPrev => settings.media_prev = binding,
        ShortcutAction::MediaNext => settings.media_next = binding,
        ShortcutAction::ChapterPrev => settings.chapter_prev = binding,
        ShortcutAction::ChapterNext => settings.chapter_next = binding,
    }
}

fn find_shortcut_conflict(
    settings: &ShortcutSettings,
    current_action: ShortcutAction,
    candidate: ShortcutBinding,
) -> Option<ShortcutAction> {
    ShortcutAction::ALL.iter().copied().find(|action| {
        *action != current_action && shortcut_binding_for_action(settings, *action) == candidate
    })
}

fn is_modifier_vk(key: u16) -> bool {
    matches!(key, 0x10 | 0x11 | 0x12 | 0xA0..=0xA5)
}

fn modifier_down(vk: i32) -> bool {
    ((crate::get_key_state_safe(vk) & (0x8000u16 as i16)) != 0)
        || ((unsafe { GetAsyncKeyState(vk) } & (0x8000u16 as i16)) != 0)
}

fn modifier_down_any(vks: &[i32]) -> bool {
    vks.iter().copied().any(modifier_down)
}

fn move_options_focus_tab(hwnd: HWND, backwards: bool) {
    let current = crate::get_focus_safe();
    let next = crate::get_next_dlg_tab_item_safe(hwnd, current, backwards);
    if next.0 != 0 {
        crate::set_focus_safe(next);
    }
}

pub fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    unsafe {
        if (msg.message == WM_KEYDOWN
            || msg.message == windows::Win32::UI::WindowsAndMessaging::WM_SYSKEYDOWN)
            && with_options_state(hwnd, |state| state.shortcut_capture_pending).unwrap_or(false)
        {
            let edit_shortcut_value =
                with_options_state(hwnd, |state| state.edit_shortcut_value).unwrap_or(HWND(0));
            if edit_shortcut_value.0 == 0 {
                return false;
            }

            let key = msg.wParam.0 as u16;
            if key == VK_ESCAPE.0 {
                if with_options_state(hwnd, |state| {
                    state.shortcut_capture_pending = false;
                })
                .is_none()
                {
                    crate::log_debug("Failed to access state in options_window");
                }
                update_shortcut_binding_text(hwnd);
                return true;
            }
            if key == VK_TAB.0 {
                if with_options_state(hwnd, |state| {
                    state.shortcut_capture_pending = false;
                })
                .is_none()
                {
                    crate::log_debug("Failed to access state in options_window");
                }
                update_shortcut_binding_text(hwnd);
                let shift_down =
                    modifier_down_any(&[VK_SHIFT.0 as i32, VK_LSHIFT.0 as i32, VK_RSHIFT.0 as i32]);
                move_options_focus_tab(hwnd, shift_down);
                return true;
            }
            if key == VK_RETURN.0 || key == VK_SPACE.0 {
                return true;
            }
            if is_modifier_vk(key) {
                return true;
            }

            let normalized_key = if (b'a' as u16..=b'z' as u16).contains(&key) {
                key - 32
            } else {
                key
            };
            let action = selected_shortcut_action(hwnd);
            let ctrl = modifier_down_any(&[
                VK_CONTROL.0 as i32,
                VK_LCONTROL.0 as i32,
                VK_RCONTROL.0 as i32,
            ]);
            let shift =
                modifier_down_any(&[VK_SHIFT.0 as i32, VK_LSHIFT.0 as i32, VK_RSHIFT.0 as i32]);
            let alt = modifier_down_any(&[VK_MENU.0 as i32, VK_LMENU.0 as i32, VK_RMENU.0 as i32]);
            let candidate = ShortcutBinding::new(ctrl, shift, alt, normalized_key);

            let conflict = with_options_state(hwnd, |state| {
                find_shortcut_conflict(&state.shortcut_draft, action, candidate).map(
                    |conflict_action| {
                        let language = with_state(state.parent, |app| app.settings.language)
                            .unwrap_or_default();
                        let shortcut = format_shortcut(candidate);
                        let conflict_label = shortcut_action_label(language, conflict_action);
                        let message = i18n::tr_f(
                            language,
                            "options.shortcuts.duplicate_error",
                            &[("shortcut", &shortcut), ("action", &conflict_label)],
                        );
                        (language, message)
                    },
                )
            })
            .flatten();
            if let Some((language, message)) = conflict {
                crate::show_error(hwnd, language, &message);
                update_shortcut_binding_text(hwnd);
                return true;
            }

            if with_options_state(hwnd, |state| {
                set_shortcut_binding_for_action(&mut state.shortcut_draft, action, candidate);
                state.shortcut_capture_pending = false;
            })
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            update_shortcut_binding_text(hwnd);
            return true;
        }

        if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
            let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
            if ctrl_down {
                let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;

                if let Some(tabs) = with_options_state(hwnd, |state| state.hwnd_tabs)
                    && tabs.0 != 0
                {
                    let current = SendMessageW(tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                    let mut next = if shift_down { current - 1 } else { current + 1 };
                    if next < 0 {
                        next = OPTIONS_TAB_COUNT - 1;
                    } else if next >= OPTIONS_TAB_COUNT {
                        next = 0;
                    }
                    SendMessageW(tabs, TCM_SETCURSEL, WPARAM(next as usize), LPARAM(0));
                    set_active_tab(hwnd, next);
                    SetFocus(tabs);

                    // Force update focus for screen readers
                    if let Err(_e) =
                        PostMessageW(hwnd, WM_NEXTDLGCTL, WPARAM(tabs.0 as usize), LPARAM(1))
                    {
                        crate::log_debug(&format!("Error posting WM_NEXTDLGCTL: {:?}", _e));
                    }
                    return true;
                }
            }
        }
        if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
            let focus = GetFocus();
            if GetParent(focus) == hwnd {
                let dropped = SendMessageW(focus, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 != 0;
                if !dropped {
                    // If focus is on the insert tag button, insert the tag before closing
                    let is_insert_tag_button =
                        with_options_state(hwnd, |state| focus == state.button_tts_insert_tag)
                            .unwrap_or(false);
                    if is_insert_tag_button {
                        insert_voice_tag_from_options(hwnd);
                    }
                    if with_options_state(hwnd, |state| {
                        SendMessageW(
                            hwnd,
                            WM_COMMAND,
                            WPARAM(OPTIONS_ID_OK),
                            LPARAM(state.ok_button.0),
                        );
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access state in options_window");
                    }
                    return true;
                }
            }
        }
        handle_accessibility(hwnd, msg)
    }
}

struct OptionsDialogState {
    parent: HWND,
    hwnd_tabs: HWND,
    focus_initialized: bool,
    label_language: HWND,
    label_modified_marker_position: HWND,
    label_open: HWND,
    label_tts_engine: HWND,
    label_tts_voice_language: HWND,
    label_voice_profile: HWND,
    label_voice: HWND,
    label_tts_speed: HWND,
    label_tts_pitch: HWND,
    label_tts_volume: HWND,
    button_tts_preview: HWND,
    button_tts_insert_tag: HWND,
    button_rename_voice_profile: HWND,
    button_add_voice_profile: HWND,
    button_delete_voice_profile: HWND,
    combo_lang: HWND,
    combo_modified_marker_position: HWND,
    combo_open: HWND,
    combo_tts_engine: HWND,
    combo_tts_voice_language: HWND,
    combo_voice_profile: HWND,
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
    label_default_save_folder_kind: HWND,
    combo_default_save_folder_kind: HWND,
    label_audiobook_save_folder: HWND,
    edit_audiobook_save_folder: HWND,
    button_audiobook_save_folder_browse: HWND,
    checkbox_show_media_save_confirmation: HWND,
    label_audio_split: HWND,
    combo_audio_split: HWND,
    label_audio_split_minutes: HWND,
    combo_audio_split_minutes: HWND,
    label_audio_split_parts_count: HWND,
    edit_audio_split_parts_count: HWND,
    label_audio_split_start_number: HWND,
    combo_audio_split_start_number: HWND,
    label_audiobook_part_naming: HWND,
    combo_audiobook_part_naming: HWND,
    label_audio_split_text: HWND,
    edit_audio_split_text: HWND,
    checkbox_audio_split_requires_newline: HWND,
    checkbox_audio_split_epub_chapters: HWND,
    checkbox_subtitle_ducking: HWND,
    label_subtitle_offset: HWND,
    edit_subtitle_offset: HWND,
    label_podcast_cache_limit: HWND,
    edit_podcast_cache_limit: HWND,
    checkbox_rss_show_article_preview: HWND,
    checkbox_announce_unread_rss_podcast: HWND,
    label_unread_label_position: HWND,
    combo_unread_label_position: HWND,
    label_rss_date_display: HWND,
    combo_rss_date_display: HWND,
    label_rss_time_display: HWND,
    combo_rss_time_display: HWND,
    label_podcast_date_display: HWND,
    combo_podcast_date_display: HWND,
    label_podcast_time_display: HWND,
    combo_podcast_time_display: HWND,
    label_podcastindex_key: HWND,
    edit_podcastindex_key: HWND,
    label_podcastindex_secret: HWND,
    edit_podcastindex_secret: HWND,
    label_whisper_model: HWND,
    combo_whisper_model: HWND,
    checkbox_whisper_cuda: HWND,
    checkbox_whisper_keep_original_language: HWND,
    checkbox_whisper_include_timestamps: HWND,
    label_dictation_microphone: HWND,
    combo_dictation_microphone: HWND,
    button_podcastindex_signup: HWND,
    checkbox_multilingual: HWND,
    checkbox_use_dialogue_voice: HWND,
    label_dialogue_voice: HWND,
    combo_dialogue_voice: HWND,
    button_dialogue_voice_preview: HWND,
    button_dialogue_secondary_voice_preview: HWND,
    label_dialogue_engine: HWND,
    combo_dialogue_engine: HWND,
    label_dialogue_voice_language: HWND,
    combo_dialogue_voice_language: HWND,
    label_dialogue_voice_rate: HWND,
    combo_dialogue_voice_rate: HWND,
    edit_dialogue_voice_rate: HWND,
    label_dialogue_voice_pitch: HWND,
    combo_dialogue_voice_pitch: HWND,
    edit_dialogue_voice_pitch: HWND,
    label_dialogue_voice_volume: HWND,
    combo_dialogue_voice_volume: HWND,
    edit_dialogue_voice_volume: HWND,
    checkbox_dialogue_multilingual: HWND,
    checkbox_dialogue_use_secondary_voice: HWND,
    label_dialogue_secondary_engine: HWND,
    combo_dialogue_secondary_engine: HWND,
    label_dialogue_secondary_voice_language: HWND,
    combo_dialogue_secondary_voice_language: HWND,
    label_dialogue_secondary_voice: HWND,
    combo_dialogue_secondary_voice: HWND,
    label_dialogue_secondary_voice_rate: HWND,
    combo_dialogue_secondary_voice_rate: HWND,
    edit_dialogue_secondary_voice_rate: HWND,
    label_dialogue_secondary_voice_pitch: HWND,
    combo_dialogue_secondary_voice_pitch: HWND,
    edit_dialogue_secondary_voice_pitch: HWND,
    label_dialogue_secondary_voice_volume: HWND,
    combo_dialogue_secondary_voice_volume: HWND,
    edit_dialogue_secondary_voice_volume: HWND,
    checkbox_dialogue_secondary_multilingual: HWND,
    label_dialogue_open_quote: HWND,
    edit_dialogue_open_quote: HWND,
    label_dialogue_close_quote: HWND,
    edit_dialogue_close_quote: HWND,
    checkbox_dialogue_allow_multiline: HWND,
    checkbox_split_on_newline: HWND,
    checkbox_word_wrap: HWND,
    checkbox_smart_quotes: HWND,
    checkbox_strip_markdown_keep_bullets: HWND,
    checkbox_spellcheck: HWND,
    label_spellcheck_language: HWND,
    combo_spellcheck_language: HWND,
    label_dictionary_translation: HWND,
    combo_dictionary_translation: HWND,
    label_wikipedia_language: HWND,
    combo_wikipedia_language: HWND,
    label_wrap_width: HWND,
    edit_wrap_width: HWND,
    label_indentation: HWND,
    combo_indentation: HWND,
    label_tab_width: HWND,
    combo_tab_width: HWND,
    label_space_width: HWND,
    combo_space_width: HWND,
    label_quote_prefix: HWND,
    edit_quote_prefix: HWND,
    label_interpreter_path: HWND,
    edit_interpreter_path: HWND,
    button_interpreter_browse: HWND,
    button_interpreter_search: HWND,
    label_subtitle_mode: HWND,
    combo_subtitle_mode: HWND,
    checkbox_move_cursor: HWND,
    checkbox_check_updates: HWND,
    checkbox_check_beta_updates: HWND,
    checkbox_send_crash_reports: HWND,
    checkbox_use_legacy_name: HWND,
    checkbox_context_menu: HWND,
    label_confirm_delete_rss_mode: HWND,
    combo_confirm_delete_rss_mode: HWND,
    label_confirm_delete_podcast_mode: HWND,
    combo_confirm_delete_podcast_mode: HWND,
    label_rss_quick_copy_mode: HWND,
    combo_rss_quick_copy_mode: HWND,
    label_file_associations: HWND,
    button_manage_associations: HWND,
    label_prompt_program: HWND,
    combo_prompt_program: HWND,
    label_network_proxy: HWND,
    edit_network_proxy: HWND,
    label_network_proxy_username: HWND,
    edit_network_proxy_username: HWND,
    label_network_proxy_password: HWND,
    edit_network_proxy_password: HWND,
    label_shortcut_action: HWND,
    combo_shortcut_action: HWND,
    label_shortcut_value: HWND,
    edit_shortcut_value: HWND,
    button_shortcut_change: HWND,
    button_shortcut_reset: HWND,
    button_shortcut_reset_all: HWND,
    shortcut_draft: ShortcutSettings,
    shortcut_capture_pending: bool,
    tts_voice_language_codes: Vec<String>,
    dialogue_voice_language_codes: Vec<String>,
    secondary_dialogue_voice_language_codes: Vec<String>,
    dictation_microphone_device_ids: Vec<String>,
    voice_profiles: Vec<VoiceProfile>,
    active_voice_profile_name: String,
    active_tab: i32,
    scroll_offsets: [i32; OPTIONS_TAB_COUNT as usize],
    content_heights: [i32; OPTIONS_TAB_COUNT as usize],
    default_save_folder_selection: u32,
    default_save_folder_audiobook: String,
    default_save_folder_media: String,
    default_save_folder_documents: String,
    ok_button: HWND,
    cancel_button: HWND,
}

struct OptionsLabels {
    title: String,
    tab_general: String,
    tab_voice: String,
    tab_editor: String,
    tab_audio: String,
    tab_rss_podcast: String,
    tab_ai_transcription: String,
    tab_shortcuts: String,
    label_language: String,
    label_modified_marker_position: String,
    label_open: String,
    label_tts_engine: String,
    label_tts_voice_language: String,
    label_voice_profile: String,
    label_voice: String,
    button_rename_voice_profile: String,
    button_add_voice_profile: String,
    button_delete_voice_profile: String,
    label_multilingual: String,
    label_use_dialogue_voice: String,
    label_dialogue_voice: String,
    label_dialogue_voice_preview: String,
    label_dialogue_engine: String,
    label_dialogue_voice_language: String,
    label_dialogue_voice_rate: String,
    label_dialogue_voice_pitch: String,
    label_dialogue_voice_volume: String,
    label_dialogue_use_secondary_voice: String,
    label_dialogue_secondary_engine: String,
    label_dialogue_secondary_voice_language: String,
    label_dialogue_secondary_voice: String,
    label_dialogue_open_quote: String,
    label_dialogue_close_quote: String,
    label_dialogue_allow_multiline: String,
    label_tts_speed: String,
    label_tts_pitch: String,
    label_tts_volume: String,
    label_tts_preview: String,
    label_tts_insert_tag: String,
    label_tts_manual_tuning: String,
    label_split_on_newline: String,
    label_word_wrap: String,
    label_smart_quotes: String,
    label_strip_markdown_keep_bullets: String,
    label_spellcheck: String,
    label_spellcheck_language: String,
    label_dictionary_translation: String,
    label_wikipedia_language: String,
    label_wrap_width: String,
    label_indentation: String,
    indent_default: String,
    indent_tabs: String,
    indent_spaces: String,
    label_tab_width: String,
    label_space_width: String,
    label_quote_prefix: String,
    label_interpreter_path: String,
    label_interpreter_browse: String,
    label_interpreter_search: String,
    label_subtitle_mode: String,
    label_move_cursor: String,
    label_check_updates: String,
    label_check_beta_updates: String,
    label_send_crash_reports: String,
    label_use_legacy_name: String,
    label_context_menu: String,
    label_confirm_delete_rss_mode: String,
    label_confirm_delete_podcast_mode: String,
    label_rss_quick_copy_mode: String,
    label_file_associations: String,
    label_manage_associations: String,
    label_prompt_program: String,
    label_network_proxy: String,
    label_network_proxy_username: String,
    label_network_proxy_password: String,
    label_shortcut_action: String,
    label_shortcut_value: String,
    label_shortcut_change: String,
    label_shortcut_reset: String,
    label_shortcut_reset_all: String,
    label_audio_skip: String,
    label_default_save_folder_kind: String,
    option_default_save_folder_audiobooks: String,
    option_default_save_folder_media: String,
    option_default_save_folder_documents: String,
    label_audiobook_save_folder: String,
    label_audiobook_save_folder_browse: String,
    label_show_media_save_confirmation: String,
    label_audio_split: String,
    label_audio_split_minutes: String,
    label_audio_split_parts_count: String,
    label_audio_split_start_number: String,
    label_audiobook_part_naming: String,
    label_audio_split_text: String,
    label_audio_split_requires_newline: String,
    label_audio_split_epub_chapters: String,
    option_audiobook_part_naming_title_number: String,
    option_audiobook_part_naming_number_only: String,
    option_audiobook_part_naming_number_title: String,
    label_subtitle_ducking: String,
    label_subtitle_offset: String,
    label_podcast_cache_limit: String,
    label_rss_show_article_preview: String,
    label_announce_unread_rss_podcast: String,
    label_unread_label_position: String,
    label_rss_date_display: String,
    label_rss_time_display: String,
    label_podcast_date_display: String,
    label_podcast_time_display: String,
    label_podcastindex_key: String,
    label_podcastindex_secret: String,
    label_whisper_model: String,
    label_whisper_cuda: String,
    label_whisper_keep_original_language: String,
    label_whisper_include_timestamps: String,
    label_dictation_microphone: String,
    option_podcast_device_default: String,
    whisper_model_small: String,
    whisper_model_medium: String,
    whisper_model_large: String,
    label_podcastindex_signup: String,
    lang_it: String,
    lang_en: String,
    lang_es: String,
    lang_pt: String,
    lang_sv: String,
    lang_vi: String,
    lang_cs: String,
    lang_pl: String,
    lang_fr: String,
    lang_sr: String,
    lang_uk: String,
    lang_lt: String,
    lang_ru: String,
    lang_zh: String,
    marker_position_end: String,
    marker_position_beginning: String,
    open_new_tab: String,
    open_new_window: String,
    engine_edge: String,
    engine_sapi5: String,
    engine_sapi4: String,
    subtitle_mode_off: String,
    subtitle_mode_nvda: String,
    subtitle_mode_user: String,
    subtitle_mode_record: String,

    split_none: String,
    split_by_time: String,
    split_by_text: String,
    split_by_parts: String,
    spellcheck_lang_follow: String,
    spellcheck_lang_en_us: String,
    spellcheck_lang_en_gb: String,
    spellcheck_lang_it: String,
    spellcheck_lang_es: String,
    spellcheck_lang_pt_br: String,
    spellcheck_lang_fr: String,
    spellcheck_lang_de: String,
    spellcheck_lang_ru: String,
    dictionary_translation_auto: String,
    dictionary_translation_none: String,
    wikipedia_language_auto: String,
    prompt_cmd: String,
    prompt_powershell: String,
    prompt_codex: String,
    confirm_delete_feed: String,
    confirm_delete_article: String,
    confirm_delete_podcast: String,
    confirm_delete_episode: String,
    confirm_delete_both: String,
    confirm_delete_none: String,
    rss_quick_copy_title: String,
    rss_quick_copy_url: String,
    rss_quick_copy_content: String,
    rss_quick_copy_all: String,
    unread_label_position_before: String,
    unread_label_position_after: String,
    list_date_always: String,
    list_date_never: String,
    list_time_always: String,
    list_time_never: String,
    list_time_only_if_multiple_same_day: String,
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
        tab_rss_podcast: i18n::tr(language, "options.tab.rss_podcast"),
        tab_ai_transcription: i18n::tr(language, "options.tab.ai_transcription"),
        tab_shortcuts: shortcut_tab_title(language),
        label_language: i18n::tr(language, "options.label.language"),
        label_modified_marker_position: i18n::tr(
            language,
            "options.label.modified_marker_position",
        ),
        label_open: i18n::tr(language, "options.label.open"),
        label_tts_engine: i18n::tr(language, "options.label.tts_engine"),
        label_tts_voice_language: i18n::tr(language, "options.label.voice_language"),
        label_voice_profile: i18n::tr(language, "options.label.voice_profile"),
        label_voice: i18n::tr(language, "options.label.voice"),
        button_rename_voice_profile: i18n::tr(language, "options.button.apply_selected_profile"),
        button_add_voice_profile: i18n::tr(language, "options.button.add_profile"),
        button_delete_voice_profile: i18n::tr(language, "options.button.delete_profile"),
        label_multilingual: i18n::tr(language, "options.label.multilingual"),
        label_use_dialogue_voice: i18n::tr(language, "options.label.use_dialogue_voice"),
        label_dialogue_voice: i18n::tr(language, "options.label.dialogue_voice"),
        label_dialogue_voice_preview: i18n::tr(language, "options.label.dialogue_voice_preview"),
        label_dialogue_engine: i18n::tr(language, "options.label.tts_engine"),
        label_dialogue_voice_language: i18n::tr(language, "options.label.voice_language"),
        label_dialogue_voice_rate: i18n::tr(language, "tts_tuning.label_speed"),
        label_dialogue_voice_pitch: i18n::tr(language, "tts_tuning.label_pitch"),
        label_dialogue_voice_volume: i18n::tr(language, "tts_tuning.label_volume"),
        label_dialogue_use_secondary_voice: i18n::tr(
            language,
            "options.label.dialogue_use_secondary_voice",
        ),
        label_dialogue_secondary_engine: i18n::tr(
            language,
            "options.label.dialogue_secondary_engine",
        ),
        label_dialogue_secondary_voice_language: i18n::tr(
            language,
            "options.label.dialogue_secondary_voice_language",
        ),
        label_dialogue_secondary_voice: i18n::tr(
            language,
            "options.label.dialogue_secondary_voice",
        ),
        label_dialogue_open_quote: i18n::tr(language, "dialogue_voice.apply.open_quote_title"),
        label_dialogue_close_quote: i18n::tr(language, "dialogue_voice.apply.close_quote_title"),
        label_dialogue_allow_multiline: i18n::tr(language, "dialogue_voice.apply.multiline_body"),
        label_tts_speed: i18n::tr(language, "tts_tuning.label_speed"),
        label_tts_pitch: i18n::tr(language, "tts_tuning.label_pitch"),
        label_tts_volume: i18n::tr(language, "tts_tuning.label_volume"),
        label_tts_preview: i18n::tr(language, "options.label.voice_preview"),
        label_tts_insert_tag: i18n::tr(language, "options.label.insert_voice_tag"),
        label_tts_manual_tuning: i18n::tr(language, "options.label.tts_manual_tuning"),
        label_split_on_newline: i18n::tr(language, "options.label.split_on_newline"),
        label_word_wrap: i18n::tr(language, "options.label.word_wrap"),
        label_smart_quotes: i18n::tr(language, "options.label.smart_quotes"),
        label_strip_markdown_keep_bullets: i18n::tr(
            language,
            "options.label.strip_markdown_keep_bullets",
        ),
        label_spellcheck: i18n::tr(language, "options.label.spellcheck"),
        label_spellcheck_language: i18n::tr(language, "options.label.spellcheck_language"),
        label_dictionary_translation: i18n::tr(language, "options.label.dictionary_translation"),
        label_wikipedia_language: i18n::tr(language, "options.label.wikipedia_language"),
        label_wrap_width: i18n::tr(language, "options.label.wrap_width"),
        label_indentation: i18n::tr(language, "options.label.indentation"),
        indent_default: i18n::tr(language, "options.indent.default"),
        indent_tabs: i18n::tr(language, "options.indent.tabs"),
        indent_spaces: i18n::tr(language, "options.indent.spaces"),
        label_tab_width: i18n::tr(language, "options.label.tab_width"),
        label_space_width: i18n::tr(language, "options.label.space_width"),
        label_quote_prefix: i18n::tr(language, "options.label.quote_prefix"),
        label_interpreter_path: i18n::tr(language, "options.label.interpreter_path"),
        label_interpreter_browse: i18n::tr(language, "options.button.browse"),
        label_interpreter_search: i18n::tr(language, "options.button.search_computer"),
        label_subtitle_mode: i18n::tr(language, "options.label.subtitle_mode"),
        label_move_cursor: i18n::tr(language, "options.label.move_cursor"),
        label_check_updates: i18n::tr(language, "options.label.check_updates"),
        label_check_beta_updates: {
            let value = i18n::tr(language, "options.label.check_beta_updates");
            if value == "options.label.check_beta_updates" {
                "Check beta updates".to_string()
            } else {
                value
            }
        },
        label_send_crash_reports: i18n::tr(language, "options.label.send_crash_reports"),
        label_use_legacy_name: i18n::tr(language, "options.label.legacy_name"),
        label_context_menu: i18n::tr(language, "options.label.context_menu"),
        label_confirm_delete_rss_mode: i18n::tr(language, "options.label.confirm_delete_rss_mode"),
        label_confirm_delete_podcast_mode: i18n::tr(
            language,
            "options.label.confirm_delete_podcast_mode",
        ),
        label_rss_quick_copy_mode: i18n::tr(language, "options.label.rss_quick_copy_mode"),
        label_file_associations: i18n::tr(language, "options.label.file_associations"),
        label_manage_associations: i18n::tr(language, "options.button.manage_associations"),
        label_prompt_program: i18n::tr(language, "options.label.prompt_program"),
        label_network_proxy: i18n::tr(language, "options.label.network_proxy"),
        label_network_proxy_username: i18n::tr(language, "options.label.network_proxy_username"),
        label_network_proxy_password: i18n::tr(language, "options.label.network_proxy_password"),
        label_shortcut_action: shortcuts_label_action(language),
        label_shortcut_value: shortcuts_label_value(language),
        label_shortcut_change: shortcuts_change_label(language),
        label_shortcut_reset: shortcuts_reset_label(language),
        label_shortcut_reset_all: shortcuts_reset_all_label(language),
        label_audio_skip: i18n::tr(language, "options.label.audio_skip"),
        label_default_save_folder_kind: {
            let value = i18n::tr(language, "options.label.default_save_folder_kind");
            if value == "options.label.default_save_folder_kind" {
                "Default folder type:".to_string()
            } else {
                value
            }
        },
        option_default_save_folder_audiobooks: {
            let value = i18n::tr(language, "options.choice.default_save_folder.audiobooks");
            if value == "options.choice.default_save_folder.audiobooks" {
                "Audiobooks".to_string()
            } else {
                value
            }
        },
        option_default_save_folder_media: {
            let value = i18n::tr(language, "options.choice.default_save_folder.media");
            if value == "options.choice.default_save_folder.media" {
                "Media".to_string()
            } else {
                value
            }
        },
        option_default_save_folder_documents: {
            let value = i18n::tr(language, "options.choice.default_save_folder.documents");
            if value == "options.choice.default_save_folder.documents" {
                "Documents".to_string()
            } else {
                value
            }
        },
        label_audiobook_save_folder: {
            let value = i18n::tr(language, "options.label.default_save_folder_path");
            if value == "options.label.default_save_folder_path" {
                "Default folder path:".to_string()
            } else {
                value
            }
        },
        label_audiobook_save_folder_browse: i18n::tr(language, "options.button.browse"),
        label_show_media_save_confirmation: i18n::tr(
            language,
            "options.label.show_media_save_confirmation",
        ),
        label_audio_split: i18n::tr(language, "options.label.audio_split"),
        label_audio_split_minutes: i18n::tr(language, "options.label.audio_split_minutes"),
        label_audio_split_parts_count: i18n::tr(language, "options.label.audio_split_parts_count"),
        label_audio_split_start_number: i18n::tr(
            language,
            "options.label.audio_split_start_number",
        ),
        label_audiobook_part_naming: i18n::tr(language, "options.label.audiobook_part_naming"),
        label_audio_split_text: i18n::tr(language, "options.label.audio_split_text"),
        label_audio_split_requires_newline: i18n::tr(
            language,
            "options.label.audio_split_requires_newline",
        ),
        label_audio_split_epub_chapters: i18n::tr(
            language,
            "options.label.audio_split_epub_chapters",
        ),
        option_audiobook_part_naming_title_number: i18n::tr(
            language,
            "options.choice.audiobook_part_naming.title_number",
        ),
        option_audiobook_part_naming_number_only: i18n::tr(
            language,
            "options.choice.audiobook_part_naming.number_only",
        ),
        option_audiobook_part_naming_number_title: i18n::tr(
            language,
            "options.choice.audiobook_part_naming.number_title",
        ),
        label_subtitle_ducking: i18n::tr(language, "options.label.subtitle_ducking"),
        label_subtitle_offset: i18n::tr(language, "options.label.subtitle_offset"),
        label_podcast_cache_limit: i18n::tr(language, "options.label.podcast_cache_limit"),
        label_rss_show_article_preview: i18n::tr(
            language,
            "options.label.rss_show_article_preview",
        ),
        label_announce_unread_rss_podcast: i18n::tr(
            language,
            "options.label.announce_unread_rss_podcast",
        ),
        label_unread_label_position: i18n::tr(language, "options.label.unread_label_position"),
        label_rss_date_display: i18n::tr(language, "options.label.rss_date_display"),
        label_rss_time_display: i18n::tr(language, "options.label.rss_time_display"),
        label_podcast_date_display: i18n::tr(language, "options.label.podcast_date_display"),
        label_podcast_time_display: i18n::tr(language, "options.label.podcast_time_display"),
        label_podcastindex_key: i18n::tr(language, "options.label.podcastindex_key"),
        label_podcastindex_secret: i18n::tr(language, "options.label.podcastindex_secret"),
        label_whisper_model: i18n::tr(language, "options.label.whisper_model"),
        label_whisper_cuda: i18n::tr(language, "options.label.whisper_cuda"),
        label_whisper_keep_original_language: i18n::tr(
            language,
            "options.label.whisper_keep_original_language",
        ),
        label_whisper_include_timestamps: i18n::tr(
            language,
            "options.label.whisper_include_timestamps",
        ),
        label_dictation_microphone: i18n::tr(language, "podcast.mic_device"),
        option_podcast_device_default: i18n::tr(language, "podcast.device.default"),
        whisper_model_small: i18n::tr(language, "options.whisper_model.small"),
        whisper_model_medium: i18n::tr(language, "options.whisper_model.medium"),
        whisper_model_large: i18n::tr(language, "options.whisper_model.large"),
        label_podcastindex_signup: i18n::tr(language, "options.button.podcastindex_signup"),
        lang_it: i18n::tr(language, "options.lang.it"),
        lang_en: i18n::tr(language, "options.lang.en"),
        lang_es: i18n::tr(language, "options.lang.es"),
        lang_pt: i18n::tr(language, "options.lang.pt"),
        lang_sv: i18n::tr(language, "options.lang.sv"),
        lang_vi: i18n::tr(language, "options.lang.vi"),
        lang_cs: i18n::tr(language, "options.lang.cs"),
        lang_pl: i18n::tr(language, "options.lang.pl"),
        lang_fr: i18n::tr(language, "options.lang.fr"),
        lang_sr: {
            let value = i18n::tr(language, "options.lang.sr");
            if value == "options.lang.sr" {
                "Serbian".to_string()
            } else {
                value
            }
        },
        lang_uk: {
            let value = i18n::tr(language, "options.lang.uk");
            if value == "options.lang.uk" {
                "Ukrainian".to_string()
            } else {
                value
            }
        },
        lang_lt: {
            let value = i18n::tr(language, "options.lang.lt");
            if value == "options.lang.lt" {
                "Lithuanian".to_string()
            } else {
                value
            }
        },
        lang_ru: {
            let value = i18n::tr(language, "options.lang.ru");
            if value == "options.lang.ru" {
                "Russian".to_string()
            } else {
                value
            }
        },
        lang_zh: {
            let value = i18n::tr(language, "options.lang.zh");
            if value == "options.lang.zh" {
                "Chinese".to_string()
            } else {
                value
            }
        },
        marker_position_end: i18n::tr(language, "options.modified_marker_position.end"),
        marker_position_beginning: i18n::tr(language, "options.modified_marker_position.beginning"),
        open_new_tab: i18n::tr(language, "options.open.new_tab"),
        open_new_window: i18n::tr(language, "options.open.new_window"),
        engine_edge: i18n::tr(language, "options.engine.edge"),
        engine_sapi5: i18n::tr(language, "options.engine.sapi5"),
        engine_sapi4: "SAPI 4".to_string(),
        subtitle_mode_off: i18n::tr(language, "options.subtitle_mode.off"),
        subtitle_mode_nvda: i18n::tr(language, "options.subtitle_mode.nvda"),
        subtitle_mode_user: i18n::tr(language, "options.subtitle_mode.user"),
        subtitle_mode_record: i18n::tr(language, "options.subtitle_mode.record"),

        split_none: i18n::tr(language, "options.split.none"),
        split_by_time: i18n::tr(language, "options.split.by_time"),
        split_by_text: i18n::tr(language, "options.split.by_text"),
        split_by_parts: i18n::tr(language, "options.split.by_parts"),
        spellcheck_lang_follow: i18n::tr(language, "options.spellcheck.lang.follow"),
        spellcheck_lang_en_us: i18n::tr(language, "options.spellcheck.lang.en_us"),
        spellcheck_lang_en_gb: i18n::tr(language, "options.spellcheck.lang.en_gb"),
        spellcheck_lang_it: i18n::tr(language, "options.spellcheck.lang.it"),
        spellcheck_lang_es: i18n::tr(language, "options.spellcheck.lang.es"),
        spellcheck_lang_pt_br: i18n::tr(language, "options.spellcheck.lang.pt_br"),
        spellcheck_lang_fr: i18n::tr(language, "options.spellcheck.lang.fr"),
        spellcheck_lang_de: i18n::tr(language, "options.spellcheck.lang.de"),
        spellcheck_lang_ru: i18n::tr(language, "options.spellcheck.lang.ru"),
        dictionary_translation_auto: i18n::tr(language, "options.dictionary_translation.auto"),
        dictionary_translation_none: i18n::tr(language, "options.dictionary_translation.none"),
        wikipedia_language_auto: i18n::tr(language, "options.wikipedia_language.auto"),
        prompt_cmd: i18n::tr(language, "options.prompt.cmd"),
        prompt_powershell: i18n::tr(language, "options.prompt.powershell"),
        prompt_codex: i18n::tr(language, "options.prompt.codex"),
        confirm_delete_feed: i18n::tr(language, "options.confirm_delete.feed"),
        confirm_delete_article: i18n::tr(language, "options.confirm_delete.article"),
        confirm_delete_podcast: i18n::tr(language, "options.confirm_delete.podcast"),
        confirm_delete_episode: i18n::tr(language, "options.confirm_delete.episode"),
        confirm_delete_both: i18n::tr(language, "options.confirm_delete.both"),
        confirm_delete_none: i18n::tr(language, "options.confirm_delete.none"),
        rss_quick_copy_title: i18n::tr(language, "options.rss_quick_copy.title"),
        rss_quick_copy_url: i18n::tr(language, "options.rss_quick_copy.url"),
        rss_quick_copy_content: i18n::tr(language, "options.rss_quick_copy.content"),
        rss_quick_copy_all: i18n::tr(language, "options.rss_quick_copy.all"),
        unread_label_position_before: i18n::tr(language, "options.unread_label_position.before"),
        unread_label_position_after: i18n::tr(language, "options.unread_label_position.after"),
        list_date_always: i18n::tr(language, "options.list_date.always"),
        list_date_never: i18n::tr(language, "options.list_date.never"),
        list_time_always: i18n::tr(language, "options.list_time.always"),
        list_time_never: i18n::tr(language, "options.list_time.never"),
        list_time_only_if_multiple_same_day: i18n::tr(
            language,
            "options.list_time.only_if_multiple_same_day",
        ),
        ok: i18n::tr(language, "options.ok"),
        cancel: i18n::tr(language, "options.cancel"),
        voices_empty: i18n::tr(language, "options.voices.empty"),
    }
}

fn voice_locale_language_code(locale: &str) -> Option<String> {
    let base = locale.split(['-', '_']).next()?.trim();
    if base.is_empty() {
        return None;
    }
    Some(base.to_ascii_lowercase())
}

fn localized_voice_language_name(language: Language, labels: &OptionsLabels, code: &str) -> String {
    let key = format!("voice.lang.{}", code);
    let localized = i18n::tr(language, &key);
    if localized != key {
        return localized;
    }
    match code {
        "it" => labels.lang_it.clone(),
        "en" => labels.lang_en.clone(),
        "es" => labels.lang_es.clone(),
        "pt" => labels.lang_pt.clone(),
        "sv" => labels.lang_sv.clone(),
        "vi" => labels.lang_vi.clone(),
        "cs" => labels.lang_cs.clone(),
        "pl" => labels.lang_pl.clone(),
        "fr" => labels.lang_fr.clone(),
        "sr" => labels.lang_sr.clone(),
        "uk" => labels.lang_uk.clone(),
        "lt" => labels.lang_lt.clone(),
        "ru" => labels.lang_ru.clone(),
        "zh" => labels.lang_zh.clone(),
        "de" => match language {
            Language::Italian => "Tedesco".to_string(),
            Language::Spanish => "Aleman".to_string(),
            Language::Portuguese => "Alemao".to_string(),
            Language::French => "Allemand".to_string(),
            _ => "German".to_string(),
        },
        _ => code.to_ascii_uppercase(),
    }
}

fn collect_voice_language_codes(voices: &[VoiceInfo]) -> Vec<String> {
    let mut codes: Vec<String> = voices
        .iter()
        .filter_map(|v| voice_locale_language_code(&v.locale))
        .collect();
    codes.sort();
    codes.dedup();
    codes
}

fn next_voice_profile_name(profiles: &[VoiceProfile]) -> String {
    let mut n = 1usize;
    loop {
        let candidate = format!("Profilo {n}");
        if !profiles
            .iter()
            .any(|p| p.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
        n += 1;
    }
}

fn refresh_voice_profile_combo(hwnd: HWND) {
    let Some((combo, profiles, active_name)) = with_options_state(hwnd, |state| {
        (
            state.combo_voice_profile,
            state.voice_profiles.clone(),
            state.active_voice_profile_name.clone(),
        )
    }) else {
        return;
    };

    unsafe {
        SendMessageW(combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let mut selected_index: usize = 0;
        for (idx, profile) in profiles.iter().enumerate() {
            let _added = SendMessageW(
                combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&profile.name).as_ptr() as isize),
            );
            if profile.name.eq_ignore_ascii_case(&active_name) {
                selected_index = idx;
            }
        }
        if !profiles.is_empty() {
            SendMessageW(combo, CB_SETCURSEL, WPARAM(selected_index), LPARAM(0));
        }
    }

    update_voice_profile_delete_button_visibility(hwnd);
}

fn update_voice_profile_delete_button_visibility(hwnd: HWND) {
    let Some((rename_button, delete_button, selected_name)) = with_options_state(hwnd, |state| {
        let sel = unsafe {
            SendMessageW(
                state.combo_voice_profile,
                CB_GETCURSEL,
                WPARAM(0),
                LPARAM(0),
            )
            .0
        };
        let selected = if sel >= 0 {
            state
                .voice_profiles
                .get(sel as usize)
                .map(|p| p.name.clone())
                .unwrap_or_else(|| DEFAULT_VOICE_PROFILE_NAME.to_string())
        } else {
            DEFAULT_VOICE_PROFILE_NAME.to_string()
        };
        (
            state.button_rename_voice_profile,
            state.button_delete_voice_profile,
            selected,
        )
    }) else {
        return;
    };

    let hide_actions = selected_name.eq_ignore_ascii_case(DEFAULT_VOICE_PROFILE_NAME);
    crate::show_window_safe(rename_button, if hide_actions { SW_HIDE } else { SW_SHOW });
    crate::show_window_safe(delete_button, if hide_actions { SW_HIDE } else { SW_SHOW });
}

fn voices_for_engine(parent: HWND, engine: TtsEngine) -> Vec<VoiceInfo> {
    with_state(parent, |state| match engine {
        TtsEngine::Edge => state.edge_voices.clone(),
        TtsEngine::Sapi5 => state.sapi_voices.clone(),
        TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
    })
    .unwrap_or_default()
}

fn select_voice_combo_by_short_name(combo: HWND, voices: &[VoiceInfo], short_name: &str) {
    if short_name.trim().is_empty() {
        return;
    }
    let count = unsafe { SendMessageW(combo, CB_GETCOUNT, WPARAM(0), LPARAM(0)).0 };
    if count <= 0 {
        return;
    }
    for idx in 0..count {
        let voice_index = unsafe {
            SendMessageW(combo, CB_GETITEMDATA, WPARAM(idx as usize), LPARAM(0)).0 as usize
        };
        if let Some(voice) = voices.get(voice_index)
            && voice.short_name == short_name
        {
            unsafe {
                SendMessageW(combo, CB_SETCURSEL, WPARAM(idx as usize), LPARAM(0));
            }
            break;
        }
    }
}

fn apply_selected_voice_profile(hwnd: HWND) {
    let Some((
        parent,
        profile,
        combo_tts_engine,
        combo_dialogue_engine,
        combo_dialogue_secondary_engine,
        checkbox_multilingual,
        checkbox_tts_manual,
        checkbox_use_dialogue_voice,
        checkbox_dialogue_use_secondary_voice,
        combo_tts_speed,
        combo_tts_pitch,
        combo_tts_volume,
        edit_tts_speed,
        edit_tts_pitch,
        edit_tts_volume,
        combo_dialogue_voice_rate,
        combo_dialogue_voice_pitch,
        combo_dialogue_voice_volume,
        edit_dialogue_voice_rate,
        edit_dialogue_voice_pitch,
        edit_dialogue_voice_volume,
        combo_dialogue_secondary_voice_rate,
        combo_dialogue_secondary_voice_pitch,
        combo_dialogue_secondary_voice_volume,
        edit_dialogue_secondary_voice_rate,
        edit_dialogue_secondary_voice_pitch,
        edit_dialogue_secondary_voice_volume,
        combo_voice,
        combo_dialogue_voice,
        combo_dialogue_secondary_voice,
    )) = with_options_state(hwnd, |state| {
        let sel = unsafe {
            SendMessageW(
                state.combo_voice_profile,
                CB_GETCURSEL,
                WPARAM(0),
                LPARAM(0),
            )
            .0
        };
        let selected = if sel >= 0 {
            state.voice_profiles.get(sel as usize).cloned()
        } else {
            None
        };
        (
            state.parent,
            selected,
            state.combo_tts_engine,
            state.combo_dialogue_engine,
            state.combo_dialogue_secondary_engine,
            state.checkbox_multilingual,
            state.checkbox_tts_manual,
            state.checkbox_use_dialogue_voice,
            state.checkbox_dialogue_use_secondary_voice,
            state.combo_tts_speed,
            state.combo_tts_pitch,
            state.combo_tts_volume,
            state.edit_tts_speed,
            state.edit_tts_pitch,
            state.edit_tts_volume,
            state.combo_dialogue_voice_rate,
            state.combo_dialogue_voice_pitch,
            state.combo_dialogue_voice_volume,
            state.edit_dialogue_voice_rate,
            state.edit_dialogue_voice_pitch,
            state.edit_dialogue_voice_volume,
            state.combo_dialogue_secondary_voice_rate,
            state.combo_dialogue_secondary_voice_pitch,
            state.combo_dialogue_secondary_voice_volume,
            state.edit_dialogue_secondary_voice_rate,
            state.edit_dialogue_secondary_voice_pitch,
            state.edit_dialogue_secondary_voice_volume,
            state.combo_voice,
            state.combo_dialogue_voice,
            state.combo_dialogue_secondary_voice,
        )
    })
    else {
        return;
    };

    let Some(profile) = profile else {
        return;
    };

    unsafe {
        let tts_engine_index = match profile.tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(
            combo_tts_engine,
            CB_SETCURSEL,
            WPARAM(tts_engine_index),
            LPARAM(0),
        );

        let dialogue_engine_index = match profile.dialogue_tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(
            combo_dialogue_engine,
            CB_SETCURSEL,
            WPARAM(dialogue_engine_index),
            LPARAM(0),
        );

        let dialogue_secondary_engine_index = match profile.dialogue_secondary_tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(
            combo_dialogue_secondary_engine,
            CB_SETCURSEL,
            WPARAM(dialogue_secondary_engine_index),
            LPARAM(0),
        );

        SendMessageW(
            checkbox_multilingual,
            BM_SETCHECK,
            WPARAM(if profile.tts_only_multilingual {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_tts_manual,
            BM_SETCHECK,
            WPARAM(if profile.tts_manual_tuning {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_use_dialogue_voice,
            BM_SETCHECK,
            WPARAM(if profile.use_dialogue_voice {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_dialogue_use_secondary_voice,
            BM_SETCHECK,
            WPARAM(if profile.dialogue_use_secondary_voice {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
    }

    select_combo_value(combo_tts_speed, profile.tts_rate);
    select_combo_value(combo_tts_pitch, profile.tts_pitch);
    select_combo_value(combo_tts_volume, profile.tts_volume);
    select_combo_value(combo_dialogue_voice_rate, profile.dialogue_voice_rate);
    select_combo_value(combo_dialogue_voice_pitch, profile.dialogue_voice_pitch);
    select_combo_value(combo_dialogue_voice_volume, profile.dialogue_voice_volume);
    select_combo_value(
        combo_dialogue_secondary_voice_rate,
        profile.dialogue_secondary_voice_rate,
    );
    select_combo_value(
        combo_dialogue_secondary_voice_pitch,
        profile.dialogue_secondary_voice_pitch,
    );
    select_combo_value(
        combo_dialogue_secondary_voice_volume,
        profile.dialogue_secondary_voice_volume,
    );

    if let Err(e) = crate::set_window_text_w_safe(
        edit_tts_speed,
        PCWSTR(to_wide(&tts_ui_value_from_internal(profile.tts_rate).to_string()).as_ptr()),
    ) {
        crate::log_debug(&format!("Failed to set tts speed edit from profile: {e}"));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_tts_pitch,
        PCWSTR(to_wide(&tts_ui_value_from_internal(profile.tts_pitch).to_string()).as_ptr()),
    ) {
        crate::log_debug(&format!("Failed to set tts pitch edit from profile: {e}"));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_tts_volume,
        PCWSTR(to_wide(&profile.tts_volume.to_string()).as_ptr()),
    ) {
        crate::log_debug(&format!("Failed to set tts volume edit from profile: {e}"));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_dialogue_voice_rate,
        PCWSTR(
            to_wide(&tts_ui_value_from_internal(profile.dialogue_voice_rate).to_string()).as_ptr(),
        ),
    ) {
        crate::log_debug(&format!(
            "Failed to set dialogue rate edit from profile: {e}"
        ));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_dialogue_voice_pitch,
        PCWSTR(
            to_wide(&tts_ui_value_from_internal(profile.dialogue_voice_pitch).to_string()).as_ptr(),
        ),
    ) {
        crate::log_debug(&format!(
            "Failed to set dialogue pitch edit from profile: {e}"
        ));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_dialogue_voice_volume,
        PCWSTR(to_wide(&profile.dialogue_voice_volume.to_string()).as_ptr()),
    ) {
        crate::log_debug(&format!(
            "Failed to set dialogue volume edit from profile: {e}"
        ));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_dialogue_secondary_voice_rate,
        PCWSTR(
            to_wide(&tts_ui_value_from_internal(profile.dialogue_secondary_voice_rate).to_string())
                .as_ptr(),
        ),
    ) {
        crate::log_debug(&format!(
            "Failed to set secondary dialogue rate edit from profile: {e}"
        ));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_dialogue_secondary_voice_pitch,
        PCWSTR(
            to_wide(
                &tts_ui_value_from_internal(profile.dialogue_secondary_voice_pitch).to_string(),
            )
            .as_ptr(),
        ),
    ) {
        crate::log_debug(&format!(
            "Failed to set secondary dialogue pitch edit from profile: {e}"
        ));
    }
    if let Err(e) = crate::set_window_text_w_safe(
        edit_dialogue_secondary_voice_volume,
        PCWSTR(to_wide(&profile.dialogue_secondary_voice_volume.to_string()).as_ptr()),
    ) {
        crate::log_debug(&format!(
            "Failed to set secondary dialogue volume edit from profile: {e}"
        ));
    }

    if with_options_state(hwnd, |state| {
        state.active_voice_profile_name = profile.name.clone();
    })
    .is_none()
    {
        crate::log_debug("Failed to access options state when applying profile");
    }

    update_tts_manual_visibility(hwnd);
    refresh_voices(hwnd);
    update_dialogue_voice_visibility(hwnd);
    let tts_voices = voices_for_engine(parent, profile.tts_engine);
    let dialogue_voices = voices_for_engine(parent, profile.dialogue_tts_engine);
    let dialogue_secondary_voices =
        voices_for_engine(parent, profile.dialogue_secondary_tts_engine);
    select_voice_combo_by_short_name(combo_voice, &tts_voices, &profile.tts_voice);
    select_voice_combo_by_short_name(
        combo_dialogue_voice,
        &dialogue_voices,
        &profile.dialogue_voice,
    );
    select_voice_combo_by_short_name(
        combo_dialogue_secondary_voice,
        &dialogue_secondary_voices,
        &profile.dialogue_secondary_voice,
    );
    relayout_active_tab_content(hwnd);
    update_voice_profile_delete_button_visibility(hwnd);
}

fn unique_voice_profile_name_for_rename(
    base_name: &str,
    profiles: &[VoiceProfile],
    skip_index: usize,
) -> String {
    let requested = base_name.trim();
    if requested.is_empty() {
        return String::new();
    }
    let exists = |name: &str| {
        profiles
            .iter()
            .enumerate()
            .any(|(idx, profile)| idx != skip_index && profile.name.eq_ignore_ascii_case(name))
    };
    if !exists(requested) {
        return requested.to_string();
    }
    let mut n = 2usize;
    loop {
        let candidate = format!("{requested} {n}");
        if !exists(&candidate) {
            return candidate;
        }
        n += 1;
    }
}

fn rename_selected_voice_profile(hwnd: HWND) {
    let Some((parent, selected_index, selected_name)) = with_options_state(hwnd, |state| {
        let sel = unsafe {
            SendMessageW(
                state.combo_voice_profile,
                CB_GETCURSEL,
                WPARAM(0),
                LPARAM(0),
            )
            .0
        };
        if sel < 0 {
            return None;
        }
        let idx = sel as usize;
        state
            .voice_profiles
            .get(idx)
            .map(|profile| (state.parent, idx, profile.name.clone()))
    })
    .flatten() else {
        return;
    };
    if selected_name.eq_ignore_ascii_case(DEFAULT_VOICE_PROFILE_NAME) {
        return;
    }

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let title = i18n::tr(language, "options.button.apply_selected_profile");
    let body = i18n::tr(language, "options.label.voice_profile");
    let Some(new_name_input) = crate::app_windows::prompt_window::prompt_user(
        hwnd,
        &title,
        &body,
        &selected_name,
        language,
    ) else {
        return;
    };
    let requested_name = new_name_input.trim().to_string();
    if requested_name.is_empty() {
        return;
    }

    let updated = with_options_state(hwnd, |state| {
        if selected_index >= state.voice_profiles.len() {
            return;
        }
        if state.voice_profiles[selected_index]
            .name
            .eq_ignore_ascii_case(DEFAULT_VOICE_PROFILE_NAME)
        {
            return;
        }
        let old_name = state.voice_profiles[selected_index].name.clone();
        let renamed = unique_voice_profile_name_for_rename(
            &requested_name,
            &state.voice_profiles,
            selected_index,
        );
        if renamed.is_empty() {
            return;
        }
        state.voice_profiles[selected_index].name = renamed.clone();
        if state
            .active_voice_profile_name
            .eq_ignore_ascii_case(&old_name)
        {
            state.active_voice_profile_name = renamed;
        }
    });
    if updated.is_none() {
        crate::log_debug("Failed to access options state when renaming profile");
        return;
    }
    refresh_voice_profile_combo(hwnd);
    update_voice_profile_delete_button_visibility(hwnd);
    if let Some(combo_voice_profile) = with_options_state(hwnd, |state| state.combo_voice_profile) {
        crate::set_focus_safe(combo_voice_profile);
    }
}

fn add_voice_profile(hwnd: HWND) {
    let updated = with_options_state(hwnd, |state| {
        let profile = VoiceProfile {
            name: next_voice_profile_name(&state.voice_profiles),
            ..Default::default()
        };
        state.voice_profiles.push(profile);
    });
    if updated.is_none() {
        crate::log_debug("Failed to access options state when adding profile");
        return;
    }
    refresh_voice_profile_combo(hwnd);
    update_voice_profile_delete_button_visibility(hwnd);
}

fn delete_selected_voice_profile(hwnd: HWND) {
    let updated = with_options_state(hwnd, |state| {
        let sel = unsafe {
            SendMessageW(
                state.combo_voice_profile,
                CB_GETCURSEL,
                WPARAM(0),
                LPARAM(0),
            )
            .0
        };
        if sel < 0 {
            return;
        }
        let idx = sel as usize;
        if idx >= state.voice_profiles.len() {
            return;
        }
        if state.voice_profiles[idx]
            .name
            .eq_ignore_ascii_case(DEFAULT_VOICE_PROFILE_NAME)
        {
            return;
        }
        state.voice_profiles.remove(idx);
        state.active_voice_profile_name = DEFAULT_VOICE_PROFILE_NAME.to_string();
    });
    if updated.is_none() {
        crate::log_debug("Failed to access options state when deleting profile");
        return;
    }
    refresh_voice_profile_combo(hwnd);
    update_voice_profile_delete_button_visibility(hwnd);
}
pub fn open(parent: HWND) {
    unsafe {
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
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_VSCROLL,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            700,
            760,
            parent,
            None,
            hinstance,
            Some(parent.0 as *const std::ffi::c_void),
        );

        if dialog.0 != 0 {
            if with_state(parent, |state| {
                state.options_dialog = dialog;
            })
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            EnableWindow(parent, true);
            SetForegroundWindow(dialog);
            ensure_voice_lists_loaded(parent, language);
        }
    }
}

pub fn refresh_voices(hwnd: HWND) {
    unsafe {
        let (
            parent,
            hwnd_tabs,
            combo_voice,
            combo_dialogue_voice,
            combo_dialogue_secondary_voice,
            combo_engine,
            combo_dialogue_engine,
            combo_dialogue_secondary_engine,
            checkbox,
            checkbox_dialogue_multilingual,
            checkbox_dialogue_secondary_multilingual,
            label_tts_voice_language,
            combo_tts_voice_language,
            label_dialogue_voice_language,
            combo_dialogue_voice_language,
            label_dialogue_secondary_voice_language,
            combo_dialogue_secondary_voice_language,
        ) = match with_options_state(hwnd, |state| {
            (
                state.parent,
                state.hwnd_tabs,
                state.combo_voice,
                state.combo_dialogue_voice,
                state.combo_dialogue_secondary_voice,
                state.combo_tts_engine,
                state.combo_dialogue_engine,
                state.combo_dialogue_secondary_engine,
                state.checkbox_multilingual,
                state.checkbox_dialogue_multilingual,
                state.checkbox_dialogue_secondary_multilingual,
                state.label_tts_voice_language,
                state.combo_tts_voice_language,
                state.label_dialogue_voice_language,
                state.combo_dialogue_voice_language,
                state.label_dialogue_secondary_voice_language,
                state.combo_dialogue_secondary_voice_language,
            )
        }) {
            Some(values) => values,
            None => return,
        };
        let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
        let voice_tab_active = hwnd_tabs.0 != 0
            && SendMessageW(hwnd_tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32
                == OPTIONS_TAB_VOICE;

        // Determine current engine from combo if possible, otherwise settings
        let engine_sel = SendMessageW(combo_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        let engine = if engine_sel >= 0 {
            match engine_sel {
                1 => TtsEngine::Sapi5,
                2 => TtsEngine::Sapi4,

                _ => TtsEngine::Edge,
            }
        } else {
            settings.tts_engine
        };

        let voices = with_state(parent, |state| match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
        .unwrap_or_default();
        let dialogue_engine_sel =
            SendMessageW(combo_dialogue_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        let dialogue_engine = if dialogue_engine_sel >= 0 {
            match dialogue_engine_sel {
                1 => TtsEngine::Sapi5,
                2 => TtsEngine::Sapi4,
                _ => TtsEngine::Edge,
            }
        } else {
            settings.dialogue_tts_engine
        };
        let dialogue_voices = with_state(parent, |state| match dialogue_engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
        .unwrap_or_default();
        let dialogue_secondary_engine_sel = SendMessageW(
            combo_dialogue_secondary_engine,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        let dialogue_secondary_engine = if dialogue_secondary_engine_sel >= 0 {
            match dialogue_secondary_engine_sel {
                1 => TtsEngine::Sapi5,
                2 => TtsEngine::Sapi4,
                _ => TtsEngine::Edge,
            }
        } else {
            settings.dialogue_secondary_tts_engine
        };
        let dialogue_secondary_voices =
            with_state(parent, |state| match dialogue_secondary_engine {
                TtsEngine::Edge => state.edge_voices.clone(),
                TtsEngine::Sapi5 => state.sapi_voices.clone(),
                TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
            })
            .unwrap_or_default();

        let labels = options_labels(settings.language);
        let only_multilingual =
            SendMessageW(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
        let dialogue_only_multilingual = SendMessageW(
            checkbox_dialogue_multilingual,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        let dialogue_secondary_only_multilingual = SendMessageW(
            checkbox_dialogue_secondary_multilingual,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;

        let filter_multilingual = engine == TtsEngine::Edge && only_multilingual;
        let dialogue_filter_multilingual =
            dialogue_engine == TtsEngine::Edge && dialogue_only_multilingual;
        let dialogue_secondary_filter_multilingual =
            dialogue_secondary_engine == TtsEngine::Edge && dialogue_secondary_only_multilingual;

        EnableWindow(checkbox, engine == TtsEngine::Edge);
        EnableWindow(
            checkbox_dialogue_multilingual,
            dialogue_engine == TtsEngine::Edge,
        );
        EnableWindow(
            checkbox_dialogue_secondary_multilingual,
            dialogue_secondary_engine == TtsEngine::Edge,
        );

        let show_language_combo =
            voice_tab_active && engine == TtsEngine::Edge && !only_multilingual;
        ShowWindow(
            label_tts_voice_language,
            if show_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        ShowWindow(
            combo_tts_voice_language,
            if show_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        EnableWindow(combo_tts_voice_language, show_language_combo);

        let show_dialogue_language_combo =
            voice_tab_active && dialogue_engine == TtsEngine::Edge && !dialogue_only_multilingual;
        ShowWindow(
            label_dialogue_voice_language,
            if show_dialogue_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        ShowWindow(
            combo_dialogue_voice_language,
            if show_dialogue_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        EnableWindow(combo_dialogue_voice_language, show_dialogue_language_combo);

        let show_dialogue_secondary_language_combo = voice_tab_active
            && dialogue_secondary_engine == TtsEngine::Edge
            && !dialogue_secondary_only_multilingual;
        ShowWindow(
            label_dialogue_secondary_voice_language,
            if show_dialogue_secondary_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        ShowWindow(
            combo_dialogue_secondary_voice_language,
            if show_dialogue_secondary_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        EnableWindow(
            combo_dialogue_secondary_voice_language,
            show_dialogue_secondary_language_combo,
        );

        let mut language_filter: Option<String> = None;
        if show_language_combo {
            let previous_selection = with_options_state(hwnd, |state| {
                let sel = SendMessageW(
                    state.combo_tts_voice_language,
                    CB_GETCURSEL,
                    WPARAM(0),
                    LPARAM(0),
                )
                .0;
                if sel >= 0 {
                    state.tts_voice_language_codes.get(sel as usize).cloned()
                } else {
                    None
                }
            })
            .flatten();

            let mut codes = collect_voice_language_codes(&voices);
            if !codes.is_empty() {
                let selected_from_voice = voices
                    .iter()
                    .find(|v| v.short_name == settings.tts_voice)
                    .and_then(|v| voice_locale_language_code(&v.locale));
                let selected_code = previous_selection
                    .filter(|code| codes.contains(code))
                    .or(selected_from_voice.filter(|code| codes.contains(code)))
                    .unwrap_or_else(|| codes[0].clone());
                SendMessageW(
                    combo_tts_voice_language,
                    CB_RESETCONTENT,
                    WPARAM(0),
                    LPARAM(0),
                );
                let mut selected_index: Option<usize> = None;
                for (idx, code) in codes.iter().enumerate() {
                    let label = localized_voice_language_name(settings.language, &labels, code);
                    let added = SendMessageW(
                        combo_tts_voice_language,
                        CB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(to_wide(&label).as_ptr() as isize),
                    )
                    .0;
                    if added >= 0 && *code == selected_code {
                        selected_index = Some(idx);
                    }
                }
                SendMessageW(
                    combo_tts_voice_language,
                    CB_SETCURSEL,
                    WPARAM(selected_index.unwrap_or(0)),
                    LPARAM(0),
                );
                language_filter = Some(selected_code);
            }
            if with_options_state(hwnd, |state| {
                state.tts_voice_language_codes = std::mem::take(&mut codes);
            })
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
        } else if with_options_state(hwnd, |state| {
            state.tts_voice_language_codes.clear();
            SendMessageW(
                state.combo_tts_voice_language,
                CB_RESETCONTENT,
                WPARAM(0),
                LPARAM(0),
            );
        })
        .is_none()
        {
            crate::log_debug("Failed to access state in options_window");
        }

        let mut dialogue_language_filter: Option<String> = None;
        if show_dialogue_language_combo {
            let previous_selection = with_options_state(hwnd, |state| {
                let sel = SendMessageW(
                    state.combo_dialogue_voice_language,
                    CB_GETCURSEL,
                    WPARAM(0),
                    LPARAM(0),
                )
                .0;
                if sel >= 0 {
                    state
                        .dialogue_voice_language_codes
                        .get(sel as usize)
                        .cloned()
                } else {
                    None
                }
            })
            .flatten();

            let mut codes = collect_voice_language_codes(&dialogue_voices);
            if !codes.is_empty() {
                let selected_from_voice = dialogue_voices
                    .iter()
                    .find(|v| v.short_name == settings.dialogue_voice)
                    .and_then(|v| voice_locale_language_code(&v.locale));
                let selected_code = previous_selection
                    .filter(|code| codes.contains(code))
                    .or(selected_from_voice.filter(|code| codes.contains(code)))
                    .unwrap_or_else(|| codes[0].clone());
                SendMessageW(
                    combo_dialogue_voice_language,
                    CB_RESETCONTENT,
                    WPARAM(0),
                    LPARAM(0),
                );
                let mut selected_index: Option<usize> = None;
                for (idx, code) in codes.iter().enumerate() {
                    let label = localized_voice_language_name(settings.language, &labels, code);
                    let added = SendMessageW(
                        combo_dialogue_voice_language,
                        CB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(to_wide(&label).as_ptr() as isize),
                    )
                    .0;
                    if added >= 0 && *code == selected_code {
                        selected_index = Some(idx);
                    }
                }
                SendMessageW(
                    combo_dialogue_voice_language,
                    CB_SETCURSEL,
                    WPARAM(selected_index.unwrap_or(0)),
                    LPARAM(0),
                );
                dialogue_language_filter = Some(selected_code);
            }
            if with_options_state(hwnd, |state| {
                state.dialogue_voice_language_codes = std::mem::take(&mut codes);
            })
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
        } else if with_options_state(hwnd, |state| {
            state.dialogue_voice_language_codes.clear();
            SendMessageW(
                state.combo_dialogue_voice_language,
                CB_RESETCONTENT,
                WPARAM(0),
                LPARAM(0),
            );
        })
        .is_none()
        {
            crate::log_debug("Failed to access state in options_window");
        }

        let mut dialogue_secondary_language_filter: Option<String> = None;
        if show_dialogue_secondary_language_combo {
            let previous_selection = with_options_state(hwnd, |state| {
                let sel = SendMessageW(
                    state.combo_dialogue_secondary_voice_language,
                    CB_GETCURSEL,
                    WPARAM(0),
                    LPARAM(0),
                )
                .0;
                if sel >= 0 {
                    state
                        .secondary_dialogue_voice_language_codes
                        .get(sel as usize)
                        .cloned()
                } else {
                    None
                }
            })
            .flatten();

            let mut codes = collect_voice_language_codes(&dialogue_secondary_voices);
            if !codes.is_empty() {
                let selected_from_voice = dialogue_secondary_voices
                    .iter()
                    .find(|v| v.short_name == settings.dialogue_secondary_voice)
                    .and_then(|v| voice_locale_language_code(&v.locale));
                let selected_code = previous_selection
                    .filter(|code| codes.contains(code))
                    .or(selected_from_voice.filter(|code| codes.contains(code)))
                    .unwrap_or_else(|| codes[0].clone());
                SendMessageW(
                    combo_dialogue_secondary_voice_language,
                    CB_RESETCONTENT,
                    WPARAM(0),
                    LPARAM(0),
                );
                let mut selected_index: Option<usize> = None;
                for (idx, code) in codes.iter().enumerate() {
                    let label = localized_voice_language_name(settings.language, &labels, code);
                    let added = SendMessageW(
                        combo_dialogue_secondary_voice_language,
                        CB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(to_wide(&label).as_ptr() as isize),
                    )
                    .0;
                    if added >= 0 && *code == selected_code {
                        selected_index = Some(idx);
                    }
                }
                SendMessageW(
                    combo_dialogue_secondary_voice_language,
                    CB_SETCURSEL,
                    WPARAM(selected_index.unwrap_or(0)),
                    LPARAM(0),
                );
                dialogue_secondary_language_filter = Some(selected_code);
            }
            if with_options_state(hwnd, |state| {
                state.secondary_dialogue_voice_language_codes = std::mem::take(&mut codes);
            })
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
        } else if with_options_state(hwnd, |state| {
            state.secondary_dialogue_voice_language_codes.clear();
            SendMessageW(
                state.combo_dialogue_secondary_voice_language,
                CB_RESETCONTENT,
                WPARAM(0),
                LPARAM(0),
            );
        })
        .is_none()
        {
            crate::log_debug("Failed to access state in options_window");
        }

        // If switching engine, we might not have the correct "selected" voice in settings yet if we haven't saved.
        // But we pass settings.tts_voice. If it's an ID from other engine, it won't match, so it selects default/first.
        populate_voice_combo(
            combo_voice,
            &voices,
            &settings.tts_voice,
            filter_multilingual,
            language_filter.as_deref(),
            &labels,
        );

        // Also populate dialogue voice combo
        populate_voice_combo(
            combo_dialogue_voice,
            &dialogue_voices,
            &settings.dialogue_voice,
            dialogue_filter_multilingual,
            dialogue_language_filter.as_deref(),
            &labels,
        );
        populate_voice_combo(
            combo_dialogue_secondary_voice,
            &dialogue_secondary_voices,
            &settings.dialogue_secondary_voice,
            dialogue_secondary_filter_multilingual,
            dialogue_secondary_language_filter.as_deref(),
            &labels,
        );
    }
}

unsafe extern "system" fn options_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "options_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || options_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn options_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create_struct = lparam.0 as *const CREATESTRUCTW;
                let parent = HWND((*create_struct).lpCreateParams as isize);
                let language =
                    with_state(parent, |state| state.settings.language).unwrap_or_default();
                let labels = options_labels(language);

                let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

                let hwnd_tabs = CreateWindowExW(
                    Default::default(),
                    WC_TABCONTROLW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    20,
                    10,
                    640,
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
                    labels.tab_rss_podcast.clone(),
                    labels.tab_ai_transcription.clone(),
                    labels.tab_shortcuts.clone(),
                ];
                for (index, label) in tab_labels.iter().enumerate() {
                    let mut text = to_wide(label);
                    let mut item = TCITEMW {
                        mask: TCIF_TEXT,
                        pszText: PWSTR(text.as_mut_ptr()),
                        ..Default::default()
                    };
                    SendMessageW(
                        hwnd_tabs,
                        TCM_INSERTITEMW,
                        WPARAM(index),
                        LPARAM(&mut item as *mut _ as isize),
                    );
                }
                SendMessageW(hwnd_tabs, TCM_SETCURSEL, WPARAM(0), LPARAM(0));

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

                let label_voice_profile = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_voice_profile).as_ptr()),
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
                let combo_voice_profile = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_VOICE_PROFILE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let button_rename_voice_profile = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.button_rename_voice_profile).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y,
                    300,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_RENAME_VOICE_PROFILE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let button_add_voice_profile = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.button_add_voice_profile).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y,
                    145,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_ADD_VOICE_PROFILE as isize),
                    HINSTANCE(0),
                    None,
                );
                let button_delete_voice_profile = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.button_delete_voice_profile).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    325,
                    y,
                    145,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_DELETE_VOICE_PROFILE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;

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

                let label_tts_voice_language = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_tts_voice_language).as_ptr()),
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
                let combo_tts_voice_language = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_TTS_VOICE_LANGUAGE as isize),
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
                y += 30;

                let button_tts_insert_tag = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_tts_insert_tag).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y,
                    300,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_TTS_INSERT_TAG as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let checkbox_use_dialogue_voice = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_use_dialogue_voice).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_USE_DIALOGUE_VOICE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 26;

                let label_dialogue_engine = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_engine).as_ptr()),
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
                let combo_dialogue_engine = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    120,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_TTS_ENGINE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_dialogue_voice_language = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_language).as_ptr()),
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
                let combo_dialogue_voice_language = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    120,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_LANGUAGE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_dialogue_voice = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice).as_ptr()),
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
                let combo_dialogue_voice = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let checkbox_dialogue_multilingual = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_multilingual).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_MULTILINGUAL as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 26;

                let button_dialogue_voice_preview = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_preview).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y,
                    300,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_PREVIEW as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 32;

                let label_dialogue_voice_rate = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_rate).as_ptr()),
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
                let combo_dialogue_voice_rate = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_RATE as isize),
                    HINSTANCE(0),
                    None,
                );
                let edit_dialogue_voice_rate = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_RATE_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_dialogue_voice_pitch = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_pitch).as_ptr()),
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
                let combo_dialogue_voice_pitch = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_PITCH as isize),
                    HINSTANCE(0),
                    None,
                );
                let edit_dialogue_voice_pitch = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_PITCH_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_dialogue_voice_volume = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_volume).as_ptr()),
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
                let combo_dialogue_voice_volume = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_VOLUME as isize),
                    HINSTANCE(0),
                    None,
                );
                let edit_dialogue_voice_volume = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_VOICE_VOLUME_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 36;

                let checkbox_dialogue_use_secondary_voice = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_dialogue_use_secondary_voice).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_USE_SECONDARY_VOICE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 26;

                let label_dialogue_secondary_engine = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_secondary_engine).as_ptr()),
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
                let combo_dialogue_secondary_engine = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    120,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_TTS_ENGINE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_dialogue_secondary_voice_language = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_secondary_voice_language).as_ptr()),
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
                let combo_dialogue_secondary_voice_language = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    120,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_LANGUAGE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_dialogue_secondary_voice = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_secondary_voice).as_ptr()),
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
                let combo_dialogue_secondary_voice = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let checkbox_dialogue_secondary_multilingual = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_multilingual).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_MULTILINGUAL as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 26;

                let button_dialogue_secondary_voice_preview = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_preview).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y,
                    300,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PREVIEW as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 32;

                let label_dialogue_secondary_voice_rate = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_rate).as_ptr()),
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
                let combo_dialogue_secondary_voice_rate = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_RATE as isize),
                    HINSTANCE(0),
                    None,
                );
                let edit_dialogue_secondary_voice_rate = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_RATE_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_dialogue_secondary_voice_pitch = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_pitch).as_ptr()),
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
                let combo_dialogue_secondary_voice_pitch = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PITCH as isize),
                    HINSTANCE(0),
                    None,
                );
                let edit_dialogue_secondary_voice_pitch = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PITCH_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_dialogue_secondary_voice_volume = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_voice_volume).as_ptr()),
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
                let combo_dialogue_secondary_voice_volume = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_VOLUME as isize),
                    HINSTANCE(0),
                    None,
                );
                let edit_dialogue_secondary_voice_volume = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_VOLUME_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_dialogue_open_quote = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_open_quote).as_ptr()),
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
                let edit_dialogue_open_quote = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_EDIT,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_OPEN_QUOTE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_dialogue_close_quote = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dialogue_close_quote).as_ptr()),
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
                let edit_dialogue_close_quote = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_EDIT,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_CLOSE_QUOTE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let checkbox_dialogue_allow_multiline = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_dialogue_allow_multiline).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    38,
                    hwnd,
                    HMENU(OPTIONS_ID_DIALOGUE_ALLOW_MULTILINE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 44;

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

                let label_default_save_folder_kind = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_default_save_folder_kind).as_ptr()),
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
                let combo_default_save_folder_kind = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_DEFAULT_SAVE_FOLDER_KIND as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_audiobook_save_folder = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_audiobook_save_folder).as_ptr()),
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
                let edit_audiobook_save_folder = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    220,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIOBOOK_SAVE_FOLDER as isize),
                    HINSTANCE(0),
                    None,
                );
                let button_audiobook_save_folder_browse = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_audiobook_save_folder_browse).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    395,
                    y - 2,
                    75,
                    24,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIOBOOK_SAVE_FOLDER_BROWSE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 36;

                let checkbox_show_media_save_confirmation = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_show_media_save_confirmation).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    360,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_SHOW_MEDIA_SAVE_CONFIRMATION as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 28;

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

                let label_audio_split_minutes = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_audio_split_minutes).as_ptr()),
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
                let combo_audio_split_minutes = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIO_SPLIT_MINUTES as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;

                let label_audio_split_parts_count = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_audio_split_parts_count).as_ptr()),
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
                let edit_audio_split_parts_count = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIO_SPLIT_PARTS_COUNT as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;

                let label_audio_split_start_number = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_audio_split_start_number).as_ptr()),
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
                let combo_audio_split_start_number = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIO_SPLIT_START_NUMBER as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;

                let label_audiobook_part_naming = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_audiobook_part_naming).as_ptr()),
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
                let combo_audiobook_part_naming = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIOBOOK_PART_NAMING as isize),
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

                let checkbox_audio_split_epub_chapters = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_audio_split_epub_chapters).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_AUDIO_SPLIT_EPUB_CHAPTERS as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 24;

                let checkbox_subtitle_ducking = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_subtitle_ducking).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_SUBTITLE_DUCKING as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 26;

                let label_subtitle_mode = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_subtitle_mode).as_ptr()),
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
                let combo_subtitle_mode = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    140,
                    hwnd,
                    HMENU(OPTIONS_ID_SUBTITLE_MODE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_subtitle_offset = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_subtitle_offset).as_ptr()),
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
                let edit_subtitle_offset = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_SUBTITLE_OFFSET as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_confirm_delete_rss_mode = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_confirm_delete_rss_mode).as_ptr()),
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
                let combo_confirm_delete_rss_mode = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_CONFIRM_DELETE_RSS_MODE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_confirm_delete_podcast_mode = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_confirm_delete_podcast_mode).as_ptr()),
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
                let combo_confirm_delete_podcast_mode = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_CONFIRM_DELETE_PODCAST_MODE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_rss_quick_copy_mode = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_rss_quick_copy_mode).as_ptr()),
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
                let combo_rss_quick_copy_mode = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_RSS_QUICK_COPY_MODE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

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

                let checkbox_rss_show_article_preview = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_rss_show_article_preview).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    360,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_RSS_SHOW_ARTICLE_PREVIEW as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 28;

                let checkbox_announce_unread_rss_podcast = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_announce_unread_rss_podcast).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    360,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_ANNOUNCE_UNREAD_RSS_PODCAST as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 28;

                let label_unread_label_position = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_unread_label_position).as_ptr()),
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
                let combo_unread_label_position = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_UNREAD_LABEL_POSITION as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_rss_date_display = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_rss_date_display).as_ptr()),
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
                let combo_rss_date_display = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_RSS_DATE_DISPLAY as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_rss_time_display = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_rss_time_display).as_ptr()),
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
                let combo_rss_time_display = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_RSS_TIME_DISPLAY as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_podcast_date_display = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_podcast_date_display).as_ptr()),
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
                let combo_podcast_date_display = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_PODCAST_DATE_DISPLAY as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_podcast_time_display = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_podcast_time_display).as_ptr()),
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
                let combo_podcast_time_display = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    180,
                    hwnd,
                    HMENU(OPTIONS_ID_PODCAST_TIME_DISPLAY as isize),
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

                let label_whisper_model = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_whisper_model).as_ptr()),
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
                let combo_whisper_model = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_WHISPER_MODEL as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                for model_label in [
                    labels.whisper_model_small.as_str(),
                    labels.whisper_model_medium.as_str(),
                    labels.whisper_model_large.as_str(),
                ] {
                    SendMessageW(
                        combo_whisper_model,
                        CB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(to_wide(model_label).as_ptr() as isize),
                    );
                }
                SendMessageW(combo_whisper_model, CB_SETCURSEL, WPARAM(0), LPARAM(0));
                let checkbox_whisper_cuda = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_whisper_cuda).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_WHISPER_CUDA as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;
                let checkbox_whisper_keep_original_language = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_whisper_keep_original_language).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    420,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_WHISPER_KEEP_ORIGINAL_LANGUAGE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;
                let checkbox_whisper_include_timestamps = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_whisper_include_timestamps).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    420,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_WHISPER_INCLUDE_TIMESTAMPS as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_dictation_microphone = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dictation_microphone).as_ptr()),
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
                let combo_dictation_microphone = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_DICTATION_MICROPHONE as isize),
                    HINSTANCE(0),
                    None,
                );
                let default_index = SendMessageW(
                    combo_dictation_microphone,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&labels.option_podcast_device_default).as_ptr() as isize),
                )
                .0 as usize;
                SendMessageW(
                    combo_dictation_microphone,
                    CB_SETITEMDATA,
                    WPARAM(default_index),
                    LPARAM(0),
                );
                let mut dictation_microphone_device_ids =
                    vec![crate::settings::PODCAST_DEVICE_DEFAULT.to_string()];
                if let Ok(devices) = crate::podcast_recorder::list_input_devices() {
                    for device in devices {
                        let idx = SendMessageW(
                            combo_dictation_microphone,
                            CB_ADDSTRING,
                            WPARAM(0),
                            LPARAM(to_wide(&device.name).as_ptr() as isize),
                        )
                        .0 as usize;
                        SendMessageW(
                            combo_dictation_microphone,
                            CB_SETITEMDATA,
                            WPARAM(idx),
                            LPARAM(dictation_microphone_device_ids.len() as isize),
                        );
                        dictation_microphone_device_ids.push(device.id);
                    }
                }
                SendMessageW(
                    combo_dictation_microphone,
                    CB_SETCURSEL,
                    WPARAM(0),
                    LPARAM(0),
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

                let checkbox_strip_markdown_keep_bullets = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_strip_markdown_keep_bullets).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_STRIP_MARKDOWN_KEEP_BULLETS as isize),
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

                let label_dictionary_translation = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_dictionary_translation).as_ptr()),
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
                let combo_dictionary_translation = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_DICTIONARY_TRANSLATION as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_wikipedia_language = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_wikipedia_language).as_ptr()),
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
                let combo_wikipedia_language = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_WIKIPEDIA_LANGUAGE as isize),
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

                let label_indentation = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_indentation).as_ptr()),
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
                let combo_indentation = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_INDENT_MODE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_tab_width = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_tab_width).as_ptr()),
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
                let combo_tab_width = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    80,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_TAB_WIDTH as isize),
                    HINSTANCE(0),
                    None,
                );
                let label_space_width = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_space_width).as_ptr()),
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
                let combo_space_width = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    80,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_SPACE_WIDTH as isize),
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

                let label_interpreter_path = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_interpreter_path).as_ptr()),
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
                let edit_interpreter_path = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    140,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_INTERPRETER_PATH as isize),
                    HINSTANCE(0),
                    None,
                );
                let button_interpreter_browse = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_interpreter_browse).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    315,
                    y - 2,
                    70,
                    24,
                    hwnd,
                    HMENU(OPTIONS_ID_INTERPRETER_BROWSE as isize),
                    HINSTANCE(0),
                    None,
                );
                let button_interpreter_search = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_interpreter_search).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    390,
                    y - 2,
                    100,
                    24,
                    hwnd,
                    HMENU(OPTIONS_ID_INTERPRETER_SEARCH as isize),
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

                let checkbox_check_beta_updates = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_check_beta_updates).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_CHECK_BETA_UPDATES as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 24;

                let checkbox_send_crash_reports = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_send_crash_reports).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_SEND_CRASH_REPORTS as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 24;

                let checkbox_use_legacy_name = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_use_legacy_name).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    170,
                    y,
                    300,
                    20,
                    hwnd,
                    HMENU(OPTIONS_ID_USE_LEGACY_NAME as isize),
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

                let label_file_associations = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_file_associations).as_ptr()),
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
                let button_manage_associations = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_manage_associations).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y - 2,
                    300,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_MANAGE_ASSOCIATIONS as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;

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
                y += 30;

                let label_network_proxy = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_network_proxy).as_ptr()),
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
                let edit_network_proxy = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_NETWORK_PROXY as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_network_proxy_username = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_network_proxy_username).as_ptr()),
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
                let edit_network_proxy_username = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    170,
                    y - 2,
                    300,
                    22,
                    hwnd,
                    HMENU(OPTIONS_ID_NETWORK_PROXY_USERNAME as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 30;

                let label_network_proxy_password = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_network_proxy_password).as_ptr()),
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
                let edit_network_proxy_password = CreateWindowExW(
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
                    HMENU(OPTIONS_ID_NETWORK_PROXY_PASSWORD as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 40;

                let label_shortcut_action = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_shortcut_action).as_ptr()),
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
                let combo_shortcut_action = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    170,
                    y - 2,
                    300,
                    200,
                    hwnd,
                    HMENU(OPTIONS_ID_SHORTCUT_ACTION as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;
                let label_shortcut_value = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.label_shortcut_value).as_ptr()),
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
                let edit_shortcut_value = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_EDIT,
                    PCWSTR::null(),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WINDOW_STYLE(ES_AUTOHSCROLL as u32)
                        | WINDOW_STYLE(ES_READONLY as u32),
                    170,
                    y - 2,
                    300,
                    24,
                    hwnd,
                    HMENU(OPTIONS_ID_SHORTCUT_VALUE as isize),
                    HINSTANCE(0),
                    None,
                );
                y += 34;
                let button_shortcut_change = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_shortcut_change).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y - 2,
                    146,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_SHORTCUT_CHANGE as isize),
                    HINSTANCE(0),
                    None,
                );
                let button_shortcut_reset = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_shortcut_reset).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    324,
                    y - 2,
                    146,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_SHORTCUT_RESET as isize),
                    HINSTANCE(0),
                    None,
                );
                let button_shortcut_reset_all = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.label_shortcut_reset_all).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    170,
                    y + 30,
                    300,
                    26,
                    hwnd,
                    HMENU(OPTIONS_ID_SHORTCUT_RESET_ALL as isize),
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
                    label_tts_voice_language,
                    combo_tts_voice_language,
                    label_voice_profile,
                    combo_voice_profile,
                    button_rename_voice_profile,
                    button_add_voice_profile,
                    button_delete_voice_profile,
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
                    button_tts_insert_tag,
                    label_audio_skip,
                    combo_audio_skip,
                    label_default_save_folder_kind,
                    combo_default_save_folder_kind,
                    label_audiobook_save_folder,
                    edit_audiobook_save_folder,
                    button_audiobook_save_folder_browse,
                    checkbox_show_media_save_confirmation,
                    label_audio_split,
                    combo_audio_split,
                    label_audio_split_minutes,
                    combo_audio_split_minutes,
                    label_audio_split_parts_count,
                    edit_audio_split_parts_count,
                    label_audio_split_start_number,
                    combo_audio_split_start_number,
                    label_audiobook_part_naming,
                    combo_audiobook_part_naming,
                    label_audio_split_text,
                    edit_audio_split_text,
                    checkbox_audio_split_requires_newline,
                    checkbox_audio_split_epub_chapters,
                    checkbox_subtitle_ducking,
                    label_subtitle_offset,
                    edit_subtitle_offset,
                    label_podcast_cache_limit,
                    edit_podcast_cache_limit,
                    checkbox_rss_show_article_preview,
                    checkbox_announce_unread_rss_podcast,
                    label_unread_label_position,
                    combo_unread_label_position,
                    label_rss_date_display,
                    combo_rss_date_display,
                    label_rss_time_display,
                    combo_rss_time_display,
                    label_podcast_date_display,
                    combo_podcast_date_display,
                    label_podcast_time_display,
                    combo_podcast_time_display,
                    label_podcastindex_key,
                    edit_podcastindex_key,
                    label_podcastindex_secret,
                    edit_podcastindex_secret,
                    label_whisper_model,
                    combo_whisper_model,
                    checkbox_whisper_cuda,
                    checkbox_whisper_keep_original_language,
                    checkbox_whisper_include_timestamps,
                    label_dictation_microphone,
                    combo_dictation_microphone,
                    button_podcastindex_signup,
                    checkbox_tts_manual,
                    checkbox_multilingual,
                    checkbox_use_dialogue_voice,
                    label_dialogue_engine,
                    combo_dialogue_engine,
                    label_dialogue_voice_language,
                    combo_dialogue_voice_language,
                    label_dialogue_voice,
                    combo_dialogue_voice,
                    checkbox_dialogue_multilingual,
                    label_dialogue_voice_rate,
                    combo_dialogue_voice_rate,
                    edit_dialogue_voice_rate,
                    label_dialogue_voice_pitch,
                    combo_dialogue_voice_pitch,
                    edit_dialogue_voice_pitch,
                    label_dialogue_voice_volume,
                    combo_dialogue_voice_volume,
                    edit_dialogue_voice_volume,
                    checkbox_dialogue_use_secondary_voice,
                    label_dialogue_secondary_engine,
                    combo_dialogue_secondary_engine,
                    label_dialogue_secondary_voice_language,
                    combo_dialogue_secondary_voice_language,
                    label_dialogue_secondary_voice,
                    combo_dialogue_secondary_voice,
                    checkbox_dialogue_secondary_multilingual,
                    label_dialogue_secondary_voice_rate,
                    combo_dialogue_secondary_voice_rate,
                    edit_dialogue_secondary_voice_rate,
                    label_dialogue_secondary_voice_pitch,
                    combo_dialogue_secondary_voice_pitch,
                    edit_dialogue_secondary_voice_pitch,
                    label_dialogue_secondary_voice_volume,
                    combo_dialogue_secondary_voice_volume,
                    edit_dialogue_secondary_voice_volume,
                    label_dialogue_open_quote,
                    edit_dialogue_open_quote,
                    label_dialogue_close_quote,
                    edit_dialogue_close_quote,
                    checkbox_dialogue_allow_multiline,
                    button_dialogue_voice_preview,
                    button_dialogue_secondary_voice_preview,
                    checkbox_split_on_newline,
                    checkbox_word_wrap,
                    checkbox_smart_quotes,
                    checkbox_strip_markdown_keep_bullets,
                    checkbox_spellcheck,
                    label_spellcheck_language,
                    combo_spellcheck_language,
                    label_dictionary_translation,
                    combo_dictionary_translation,
                    label_wikipedia_language,
                    combo_wikipedia_language,
                    label_wrap_width,
                    edit_wrap_width,
                    label_indentation,
                    combo_indentation,
                    label_tab_width,
                    combo_tab_width,
                    label_space_width,
                    combo_space_width,
                    label_quote_prefix,
                    edit_quote_prefix,
                    label_interpreter_path,
                    edit_interpreter_path,
                    button_interpreter_browse,
                    label_subtitle_mode,
                    combo_subtitle_mode,
                    label_subtitle_offset,
                    edit_subtitle_offset,
                    checkbox_move_cursor,
                    checkbox_check_updates,
                    checkbox_check_beta_updates,
                    checkbox_send_crash_reports,
                    checkbox_use_legacy_name,
                    checkbox_context_menu,
                    label_confirm_delete_rss_mode,
                    combo_confirm_delete_rss_mode,
                    label_confirm_delete_podcast_mode,
                    combo_confirm_delete_podcast_mode,
                    label_rss_quick_copy_mode,
                    combo_rss_quick_copy_mode,
                    label_file_associations,
                    button_manage_associations,
                    label_prompt_program,
                    combo_prompt_program,
                    label_network_proxy,
                    edit_network_proxy,
                    label_network_proxy_username,
                    edit_network_proxy_username,
                    label_network_proxy_password,
                    edit_network_proxy_password,
                    label_shortcut_action,
                    combo_shortcut_action,
                    label_shortcut_value,
                    edit_shortcut_value,
                    button_shortcut_change,
                    button_shortcut_reset,
                    button_shortcut_reset_all,
                    ok_button,
                    cancel_button,
                ] {
                    if control.0 != 0 && hfont.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
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
                    label_tts_voice_language,
                    label_voice_profile,
                    label_voice,
                    label_tts_speed,
                    label_tts_pitch,
                    label_tts_volume,
                    button_tts_preview,
                    button_tts_insert_tag,
                    button_rename_voice_profile,
                    button_add_voice_profile,
                    button_delete_voice_profile,
                    combo_lang,
                    combo_modified_marker_position,
                    combo_open,
                    combo_tts_engine,
                    combo_tts_voice_language,
                    combo_voice_profile,
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
                    label_default_save_folder_kind,
                    combo_default_save_folder_kind,
                    label_audiobook_save_folder,
                    edit_audiobook_save_folder,
                    button_audiobook_save_folder_browse,
                    checkbox_show_media_save_confirmation,
                    label_audio_split,
                    combo_audio_split,
                    label_audio_split_minutes,
                    combo_audio_split_minutes,
                    label_audio_split_parts_count,
                    edit_audio_split_parts_count,
                    label_audio_split_start_number,
                    combo_audio_split_start_number,
                    label_audiobook_part_naming,
                    combo_audiobook_part_naming,
                    label_audio_split_text,
                    edit_audio_split_text,
                    checkbox_audio_split_requires_newline,
                    checkbox_audio_split_epub_chapters,
                    checkbox_subtitle_ducking,
                    label_subtitle_offset,
                    edit_subtitle_offset,
                    label_podcast_cache_limit,
                    edit_podcast_cache_limit,
                    checkbox_rss_show_article_preview,
                    checkbox_announce_unread_rss_podcast,
                    label_unread_label_position,
                    combo_unread_label_position,
                    label_rss_date_display,
                    combo_rss_date_display,
                    label_rss_time_display,
                    combo_rss_time_display,
                    label_podcast_date_display,
                    combo_podcast_date_display,
                    label_podcast_time_display,
                    combo_podcast_time_display,
                    label_podcastindex_key,
                    edit_podcastindex_key,
                    label_podcastindex_secret,
                    edit_podcastindex_secret,
                    label_whisper_model,
                    combo_whisper_model,
                    checkbox_whisper_cuda,
                    checkbox_whisper_keep_original_language,
                    checkbox_whisper_include_timestamps,
                    label_dictation_microphone,
                    combo_dictation_microphone,
                    button_podcastindex_signup,
                    checkbox_multilingual,
                    checkbox_use_dialogue_voice,
                    label_dialogue_engine,
                    combo_dialogue_engine,
                    label_dialogue_voice_language,
                    combo_dialogue_voice_language,
                    label_dialogue_voice,
                    combo_dialogue_voice,
                    checkbox_dialogue_multilingual,
                    label_dialogue_voice_rate,
                    combo_dialogue_voice_rate,
                    edit_dialogue_voice_rate,
                    label_dialogue_voice_pitch,
                    combo_dialogue_voice_pitch,
                    edit_dialogue_voice_pitch,
                    label_dialogue_voice_volume,
                    combo_dialogue_voice_volume,
                    edit_dialogue_voice_volume,
                    checkbox_dialogue_use_secondary_voice,
                    label_dialogue_secondary_engine,
                    combo_dialogue_secondary_engine,
                    label_dialogue_secondary_voice_language,
                    combo_dialogue_secondary_voice_language,
                    label_dialogue_secondary_voice,
                    combo_dialogue_secondary_voice,
                    checkbox_dialogue_secondary_multilingual,
                    label_dialogue_secondary_voice_rate,
                    combo_dialogue_secondary_voice_rate,
                    edit_dialogue_secondary_voice_rate,
                    label_dialogue_secondary_voice_pitch,
                    combo_dialogue_secondary_voice_pitch,
                    edit_dialogue_secondary_voice_pitch,
                    label_dialogue_secondary_voice_volume,
                    combo_dialogue_secondary_voice_volume,
                    edit_dialogue_secondary_voice_volume,
                    label_dialogue_open_quote,
                    edit_dialogue_open_quote,
                    label_dialogue_close_quote,
                    edit_dialogue_close_quote,
                    checkbox_dialogue_allow_multiline,
                    button_dialogue_voice_preview,
                    button_dialogue_secondary_voice_preview,
                    checkbox_split_on_newline,
                    checkbox_word_wrap,
                    checkbox_smart_quotes,
                    checkbox_strip_markdown_keep_bullets,
                    checkbox_spellcheck,
                    label_spellcheck_language,
                    combo_spellcheck_language,
                    label_dictionary_translation,
                    combo_dictionary_translation,
                    label_wikipedia_language,
                    combo_wikipedia_language,
                    label_wrap_width,
                    edit_wrap_width,
                    label_indentation,
                    combo_indentation,
                    label_tab_width,
                    combo_tab_width,
                    label_space_width,
                    combo_space_width,
                    label_quote_prefix,
                    edit_quote_prefix,
                    label_interpreter_path,
                    edit_interpreter_path,
                    button_interpreter_browse,
                    button_interpreter_search,
                    label_subtitle_mode,
                    combo_subtitle_mode,
                    checkbox_move_cursor,
                    checkbox_check_updates,
                    checkbox_check_beta_updates,
                    checkbox_send_crash_reports,
                    checkbox_use_legacy_name,
                    checkbox_context_menu,
                    label_confirm_delete_rss_mode,
                    combo_confirm_delete_rss_mode,
                    label_confirm_delete_podcast_mode,
                    combo_confirm_delete_podcast_mode,
                    label_rss_quick_copy_mode,
                    combo_rss_quick_copy_mode,
                    label_file_associations,
                    button_manage_associations,
                    label_prompt_program,
                    combo_prompt_program,
                    label_network_proxy,
                    edit_network_proxy,
                    label_network_proxy_username,
                    edit_network_proxy_username,
                    label_network_proxy_password,
                    edit_network_proxy_password,
                    label_shortcut_action,
                    combo_shortcut_action,
                    label_shortcut_value,
                    edit_shortcut_value,
                    button_shortcut_change,
                    button_shortcut_reset,
                    button_shortcut_reset_all,
                    shortcut_draft: ShortcutSettings::default(),
                    shortcut_capture_pending: false,
                    tts_voice_language_codes: Vec::new(),
                    dialogue_voice_language_codes: Vec::new(),
                    secondary_dialogue_voice_language_codes: Vec::new(),
                    dictation_microphone_device_ids,
                    voice_profiles: Vec::new(),
                    active_voice_profile_name: DEFAULT_VOICE_PROFILE_NAME.to_string(),
                    active_tab: OPTIONS_TAB_GENERAL,
                    scroll_offsets: [0; OPTIONS_TAB_COUNT as usize],
                    content_heights: [0; OPTIONS_TAB_COUNT as usize],
                    default_save_folder_selection: 0,
                    default_save_folder_audiobook: String::new(),
                    default_save_folder_media: String::new(),
                    default_save_folder_documents: String::new(),
                    ok_button,
                    cancel_button,
                });
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(dialog_state) as isize);
                initialize_options_dialog(hwnd);
                set_active_tab(hwnd, OPTIONS_TAB_GENERAL);
                LRESULT(0)
            }
            WM_NOTIFY => {
                let hdr = &*(lparam.0 as *const NMHDR);
                if hdr.idFrom == OPTIONS_ID_TABS && hdr.code == TCN_SELCHANGE {
                    let tabs = with_options_state(hwnd, |state| state.hwnd_tabs).unwrap_or(HWND(0));
                    if tabs.0 != 0 {
                        let index =
                            SendMessageW(tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                        set_active_tab(hwnd, index);
                        return LRESULT(0);
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_COMMAND => {
                let cmd_id = wparam.0 & 0xffff;
                let code = (wparam.0 >> 16) as u32;
                match cmd_id {
                    OPTIONS_ID_OK => {
                        apply_options_dialog(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_CANCEL | 2 => {
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
                        LRESULT(0)
                    }
                    OPTIONS_ID_MULTILINGUAL => {
                        refresh_voices(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_MULTILINGUAL
                    | OPTIONS_ID_DIALOGUE_SECONDARY_MULTILINGUAL => {
                        refresh_voices(hwnd);
                        update_dialogue_voice_visibility(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_TTS_VOICE_LANGUAGE => {
                        if code == CBN_SELCHANGE {
                            refresh_voices(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_VOICE_LANGUAGE => {
                        if code == CBN_SELCHANGE {
                            refresh_voices(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_LANGUAGE => {
                        if code == CBN_SELCHANGE {
                            refresh_voices(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_TTS_PREVIEW => {
                        preview_voice(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_TTS_INSERT_TAG => {
                        insert_voice_tag_from_options(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_VOICE_PROFILE => {
                        if code == CBN_SELCHANGE {
                            // Selecting a profile immediately loads its voice settings in the form.
                            apply_selected_voice_profile(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_RENAME_VOICE_PROFILE => {
                        rename_selected_voice_profile(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_ADD_VOICE_PROFILE => {
                        add_voice_profile(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_DELETE_VOICE_PROFILE => {
                        delete_selected_voice_profile(hwnd);
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
                                let parent =
                                    with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                                let has_sapi = with_state(parent, |s| !s.sapi_voices.is_empty())
                                    .unwrap_or(false);
                                if !has_sapi {
                                    let lang = with_state(parent, |s| s.settings.language)
                                        .unwrap_or_default();
                                    ensure_sapi_voices_loaded(parent, lang);
                                }
                            }

                            refresh_voices(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_TTS_ENGINE => {
                        if code == CBN_SELCHANGE {
                            let combo = with_options_state(hwnd, |s| s.combo_dialogue_engine)
                                .unwrap_or(HWND(0));
                            let sel = SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                            if sel == 1 {
                                let parent =
                                    with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                                let has_sapi = with_state(parent, |s| !s.sapi_voices.is_empty())
                                    .unwrap_or(false);
                                if !has_sapi {
                                    let lang = with_state(parent, |s| s.settings.language)
                                        .unwrap_or_default();
                                    ensure_sapi_voices_loaded(parent, lang);
                                }
                            }
                            refresh_voices(hwnd);
                            update_dialogue_voice_visibility(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_SECONDARY_TTS_ENGINE => {
                        if code == CBN_SELCHANGE {
                            let combo =
                                with_options_state(hwnd, |s| s.combo_dialogue_secondary_engine)
                                    .unwrap_or(HWND(0));
                            let sel = SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                            if sel == 1 {
                                let parent =
                                    with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                                let has_sapi = with_state(parent, |s| !s.sapi_voices.is_empty())
                                    .unwrap_or(false);
                                if !has_sapi {
                                    let lang = with_state(parent, |s| s.settings.language)
                                        .unwrap_or_default();
                                    ensure_sapi_voices_loaded(parent, lang);
                                }
                            }
                            refresh_voices(hwnd);
                            update_dialogue_voice_visibility(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_AUDIO_SPLIT => {
                        if code == CBN_SELCHANGE {
                            update_audio_split_visibility(hwnd);
                            relayout_active_tab_content(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_SUBTITLE_MODE => {
                        if code == CBN_SELCHANGE {
                            update_subtitle_ducking_visibility(hwnd);
                            relayout_active_tab_content(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_INDENT_MODE => {
                        if code == CBN_SELCHANGE {
                            update_indentation_visibility(hwnd);
                            relayout_active_tab_content(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_SPELLCHECK_ENABLED => {
                        update_spellcheck_language_visibility(hwnd);
                        relayout_active_tab_content(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_TTS_MANUAL_TUNING => {
                        update_tts_manual_visibility(hwnd);
                        relayout_active_tab_content(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_USE_DIALOGUE_VOICE => {
                        update_dialogue_voice_visibility(hwnd);
                        relayout_active_tab_content(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_USE_SECONDARY_VOICE => {
                        update_dialogue_voice_visibility(hwnd);
                        relayout_active_tab_content(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_VOICE_PREVIEW => {
                        preview_dialogue_voice(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_DIALOGUE_SECONDARY_VOICE_PREVIEW => {
                        preview_dialogue_secondary_voice(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_DEFAULT_SAVE_FOLDER_KIND => {
                        if code == CBN_SELCHANGE {
                            on_default_save_folder_kind_changed(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_PODCASTINDEX_SIGNUP => {
                        open_podcastindex_signup();
                        LRESULT(0)
                    }
                    OPTIONS_ID_INTERPRETER_BROWSE => {
                        browse_for_interpreter(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_AUDIOBOOK_SAVE_FOLDER_BROWSE => {
                        browse_for_audiobook_folder(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_INTERPRETER_SEARCH => {
                        search_for_interpreter(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_MANAGE_ASSOCIATIONS => {
                        crate::settings::register_application_capabilities();
                        let url = to_wide("ms-settings:defaultapps?registeredApp=Sonarpad");
                        ShellExecuteW(
                            HWND(0),
                            w!("open"),
                            PCWSTR(url.as_ptr()),
                            PCWSTR::null(),
                            PCWSTR::null(),
                            SW_SHOWNORMAL,
                        );
                        LRESULT(0)
                    }
                    OPTIONS_ID_SHORTCUT_ACTION => {
                        if code == CBN_SELCHANGE {
                            if with_options_state(hwnd, |state| {
                                state.shortcut_capture_pending = false;
                            })
                            .is_none()
                            {
                                crate::log_debug("Failed to access state in options_window");
                            }
                            update_shortcut_binding_text(hwnd);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_SHORTCUT_CHANGE => {
                        if with_options_state(hwnd, |state| {
                            state.shortcut_capture_pending = true;
                        })
                        .is_none()
                        {
                            crate::log_debug("Failed to access state in options_window");
                        }
                        update_shortcut_binding_text(hwnd);
                        if let Some(control) =
                            with_options_state(hwnd, |state| state.edit_shortcut_value)
                        {
                            SetFocus(control);
                        }
                        LRESULT(0)
                    }
                    OPTIONS_ID_SHORTCUT_RESET => {
                        let action = selected_shortcut_action(hwnd);
                        let defaults = ShortcutSettings::default();
                        let default_binding = shortcut_binding_for_action(&defaults, action);
                        let conflict = with_options_state(hwnd, |state| {
                            find_shortcut_conflict(&state.shortcut_draft, action, default_binding)
                                .map(|conflict_action| {
                                    let language =
                                        with_state(state.parent, |app| app.settings.language)
                                            .unwrap_or_default();
                                    let shortcut = format_shortcut(default_binding);
                                    let conflict_label =
                                        shortcut_action_label(language, conflict_action);
                                    let message = i18n::tr_f(
                                        language,
                                        "options.shortcuts.duplicate_error",
                                        &[("shortcut", &shortcut), ("action", &conflict_label)],
                                    );
                                    (language, message)
                                })
                        })
                        .flatten();
                        if let Some((language, message)) = conflict {
                            crate::show_error(hwnd, language, &message);
                            update_shortcut_binding_text(hwnd);
                            return LRESULT(0);
                        }
                        if with_options_state(hwnd, |state| {
                            set_shortcut_binding_for_action(
                                &mut state.shortcut_draft,
                                action,
                                default_binding,
                            );
                            state.shortcut_capture_pending = false;
                        })
                        .is_none()
                        {
                            crate::log_debug("Failed to access state in options_window");
                        }
                        update_shortcut_binding_text(hwnd);
                        LRESULT(0)
                    }
                    OPTIONS_ID_SHORTCUT_RESET_ALL => {
                        if with_options_state(hwnd, |state| {
                            state.shortcut_draft = ShortcutSettings::default();
                            state.shortcut_capture_pending = false;
                        })
                        .is_none()
                        {
                            crate::log_debug("Failed to access state in options_window");
                        }
                        update_shortcut_binding_text(hwnd);
                        LRESULT(0)
                    }
                    _ => DefWindowProcW(hwnd, msg, wparam, lparam),
                }
            }
            WM_VSCROLL => {
                if handle_options_vscroll(hwnd, wparam) {
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_MOUSEWHEEL => {
                if handle_options_mouse_wheel(hwnd, wparam) {
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_TAB.0 as u32 {
                    let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
                    if ctrl_down {
                        let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
                        let tabs =
                            with_options_state(hwnd, |state| state.hwnd_tabs).unwrap_or(HWND(0));
                        if tabs.0 != 0 {
                            let current =
                                SendMessageW(tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
                            let mut next = if shift_down { current - 1 } else { current + 1 };
                            if next < 0 {
                                next = OPTIONS_TAB_COUNT - 1;
                            } else if next >= OPTIONS_TAB_COUNT {
                                next = 0;
                            }
                            SendMessageW(tabs, TCM_SETCURSEL, WPARAM(next as usize), LPARAM(0));
                            set_active_tab(hwnd, next);
                            SetFocus(tabs);
                            return LRESULT(0);
                        }
                    }
                }
                if wparam.0 as u32 == VK_RETURN.0 as u32 {
                    let focus = GetFocus();
                    let is_tts_combo = with_options_state(hwnd, |state| {
                        focus == state.combo_voice
                            || focus == state.combo_voice_profile
                            || focus == state.combo_tts_voice_language
                            || focus == state.combo_dialogue_voice
                            || focus == state.combo_dialogue_voice_language
                            || focus == state.edit_dialogue_voice_rate
                            || focus == state.edit_dialogue_voice_pitch
                            || focus == state.edit_dialogue_voice_volume
                            || focus == state.combo_dialogue_secondary_engine
                            || focus == state.combo_dialogue_secondary_voice_language
                            || focus == state.combo_dialogue_secondary_voice
                            || focus == state.combo_dialogue_secondary_voice_rate
                            || focus == state.combo_dialogue_secondary_voice_pitch
                            || focus == state.combo_dialogue_secondary_voice_volume
                            || focus == state.edit_dialogue_secondary_voice_rate
                            || focus == state.edit_dialogue_secondary_voice_pitch
                            || focus == state.edit_dialogue_secondary_voice_volume
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
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_DESTROY => {
                let parent = with_options_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
                if parent.0 != 0 {
                    EnableWindow(parent, true);
                    SetForegroundWindow(parent);
                    // Only focus editor if not in player mode (audiobook)
                    if !crate::editor_manager::is_current_audiobook(parent) {
                        SetFocus(parent);
                        if let Some(edit) = crate::get_active_edit(parent) {
                            SetFocus(edit);
                        }
                        if let Err(_e) =
                            PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0))
                        {
                            crate::log_debug(&format!("Error: {:?}", _e));
                        }
                    }
                    if with_state(parent, |state| {
                        state.options_dialog = HWND(0);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access state in options_window");
                    }
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
                if !ptr.is_null() {
                    let _unused_box = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn with_options_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut OptionsDialogState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

fn initialize_options_dialog(hwnd: HWND) {
    unsafe {
        let (
            parent,
            combo_lang,
            combo_modified_marker_position,
            combo_open,
            combo_tts_engine,
            _combo_tts_voice_language,
            _combo_voice,
            combo_tts_speed,
            combo_tts_pitch,
            combo_tts_volume,
            edit_tts_speed,
            edit_tts_pitch,
            edit_tts_volume,
            combo_audio_skip,
            combo_default_save_folder_kind,
            _label_audiobook_save_folder,
            edit_audiobook_save_folder,
            _button_audiobook_save_folder_browse,
            checkbox_show_media_save_confirmation,
            combo_audio_split,
            combo_audio_split_minutes,
            edit_audio_split_parts_count,
            combo_audio_split_start_number,
            combo_audiobook_part_naming,
            _label_audio_split_text,
            edit_audio_split_text,
            checkbox_audio_split_requires_newline,
            checkbox_audio_split_epub_chapters,
            checkbox_subtitle_ducking,
            _label_podcast_cache_limit,
            edit_podcast_cache_limit,
            checkbox_rss_show_article_preview,
            checkbox_announce_unread_rss_podcast,
            _label_unread_label_position,
            combo_unread_label_position,
            _label_rss_date_display,
            combo_rss_date_display,
            _label_rss_time_display,
            combo_rss_time_display,
            _label_podcast_date_display,
            combo_podcast_date_display,
            _label_podcast_time_display,
            combo_podcast_time_display,
            _label_podcastindex_key,
            edit_podcastindex_key,
            _label_podcastindex_secret,
            edit_podcastindex_secret,
            _label_whisper_model,
            combo_whisper_model,
            _label_dictation_microphone,
            _combo_dictation_microphone,
            _button_podcastindex_signup,
            checkbox_tts_manual,
            checkbox_multilingual,
            checkbox_use_dialogue_voice,
            combo_dialogue_engine,
            checkbox_dialogue_multilingual,
            _label_dialogue_voice_language,
            _combo_dialogue_voice_language,
            combo_dialogue_voice,
            combo_dialogue_voice_rate,
            edit_dialogue_voice_rate,
            combo_dialogue_voice_pitch,
            edit_dialogue_voice_pitch,
            combo_dialogue_voice_volume,
            edit_dialogue_voice_volume,
            checkbox_dialogue_use_secondary_voice,
            _label_dialogue_secondary_engine,
            combo_dialogue_secondary_engine,
            checkbox_dialogue_secondary_multilingual,
            _label_dialogue_secondary_voice_language,
            _combo_dialogue_secondary_voice_language,
            _label_dialogue_secondary_voice,
            combo_dialogue_secondary_voice,
            combo_dialogue_secondary_voice_rate,
            edit_dialogue_secondary_voice_rate,
            combo_dialogue_secondary_voice_pitch,
            edit_dialogue_secondary_voice_pitch,
            combo_dialogue_secondary_voice_volume,
            edit_dialogue_secondary_voice_volume,
            _label_dialogue_open_quote,
            edit_dialogue_open_quote,
            _label_dialogue_close_quote,
            edit_dialogue_close_quote,
            checkbox_dialogue_allow_multiline,
            checkbox_split_on_newline,
            checkbox_word_wrap,
            checkbox_smart_quotes,
            checkbox_strip_markdown_keep_bullets,
            checkbox_spellcheck,
            combo_spellcheck_language,
            _label_dictionary_translation,
            combo_dictionary_translation,
            _label_wikipedia_language,
            combo_wikipedia_language,
            _label_wrap_width,
            edit_wrap_width,
            _label_indentation,
            combo_indentation,
            _label_tab_width,
            combo_tab_width,
            _label_space_width,
            combo_space_width,
            _label_quote_prefix,
            edit_quote_prefix,
            _label_interpreter_path,
            edit_interpreter_path,
            _button_interpreter_browse,
            _label_subtitle_mode,
            combo_subtitle_mode,
            _label_subtitle_offset,
            edit_subtitle_offset,
            checkbox_move_cursor,
            checkbox_check_updates,
            checkbox_check_beta_updates,
            checkbox_send_crash_reports,
            checkbox_use_legacy_name,
            checkbox_context_menu,
            combo_confirm_delete_rss_mode,
            combo_confirm_delete_podcast_mode,
            combo_rss_quick_copy_mode,
            _label_prompt_program,
            combo_prompt_program,
            _label_network_proxy,
            edit_network_proxy,
            _label_network_proxy_username,
            edit_network_proxy_username,
            _label_network_proxy_password,
            edit_network_proxy_password,
            combo_shortcut_action,
            _edit_shortcut_value,
        ) = match with_options_state(hwnd, |state| {
            (
                state.parent,
                state.combo_lang,
                state.combo_modified_marker_position,
                state.combo_open,
                state.combo_tts_engine,
                state.combo_tts_voice_language,
                state.combo_voice,
                state.combo_tts_speed,
                state.combo_tts_pitch,
                state.combo_tts_volume,
                state.edit_tts_speed,
                state.edit_tts_pitch,
                state.edit_tts_volume,
                state.combo_audio_skip,
                state.combo_default_save_folder_kind,
                state.label_audiobook_save_folder,
                state.edit_audiobook_save_folder,
                state.button_audiobook_save_folder_browse,
                state.checkbox_show_media_save_confirmation,
                state.combo_audio_split,
                state.combo_audio_split_minutes,
                state.edit_audio_split_parts_count,
                state.combo_audio_split_start_number,
                state.combo_audiobook_part_naming,
                state.label_audio_split_text,
                state.edit_audio_split_text,
                state.checkbox_audio_split_requires_newline,
                state.checkbox_audio_split_epub_chapters,
                state.checkbox_subtitle_ducking,
                state.label_podcast_cache_limit,
                state.edit_podcast_cache_limit,
                state.checkbox_rss_show_article_preview,
                state.checkbox_announce_unread_rss_podcast,
                state.label_unread_label_position,
                state.combo_unread_label_position,
                state.label_rss_date_display,
                state.combo_rss_date_display,
                state.label_rss_time_display,
                state.combo_rss_time_display,
                state.label_podcast_date_display,
                state.combo_podcast_date_display,
                state.label_podcast_time_display,
                state.combo_podcast_time_display,
                state.label_podcastindex_key,
                state.edit_podcastindex_key,
                state.label_podcastindex_secret,
                state.edit_podcastindex_secret,
                state.label_whisper_model,
                state.combo_whisper_model,
                state.label_dictation_microphone,
                state.combo_dictation_microphone,
                state.button_podcastindex_signup,
                state.checkbox_tts_manual,
                state.checkbox_multilingual,
                state.checkbox_use_dialogue_voice,
                state.combo_dialogue_engine,
                state.checkbox_dialogue_multilingual,
                state.label_dialogue_voice_language,
                state.combo_dialogue_voice_language,
                state.combo_dialogue_voice,
                state.combo_dialogue_voice_rate,
                state.edit_dialogue_voice_rate,
                state.combo_dialogue_voice_pitch,
                state.edit_dialogue_voice_pitch,
                state.combo_dialogue_voice_volume,
                state.edit_dialogue_voice_volume,
                state.checkbox_dialogue_use_secondary_voice,
                state.label_dialogue_secondary_engine,
                state.combo_dialogue_secondary_engine,
                state.checkbox_dialogue_secondary_multilingual,
                state.label_dialogue_secondary_voice_language,
                state.combo_dialogue_secondary_voice_language,
                state.label_dialogue_secondary_voice,
                state.combo_dialogue_secondary_voice,
                state.combo_dialogue_secondary_voice_rate,
                state.edit_dialogue_secondary_voice_rate,
                state.combo_dialogue_secondary_voice_pitch,
                state.edit_dialogue_secondary_voice_pitch,
                state.combo_dialogue_secondary_voice_volume,
                state.edit_dialogue_secondary_voice_volume,
                state.label_dialogue_open_quote,
                state.edit_dialogue_open_quote,
                state.label_dialogue_close_quote,
                state.edit_dialogue_close_quote,
                state.checkbox_dialogue_allow_multiline,
                state.checkbox_split_on_newline,
                state.checkbox_word_wrap,
                state.checkbox_smart_quotes,
                state.checkbox_strip_markdown_keep_bullets,
                state.checkbox_spellcheck,
                state.combo_spellcheck_language,
                state.label_dictionary_translation,
                state.combo_dictionary_translation,
                state.label_wikipedia_language,
                state.combo_wikipedia_language,
                state.label_wrap_width,
                state.edit_wrap_width,
                state.label_indentation,
                state.combo_indentation,
                state.label_tab_width,
                state.combo_tab_width,
                state.label_space_width,
                state.combo_space_width,
                state.label_quote_prefix,
                state.edit_quote_prefix,
                state.label_interpreter_path,
                state.edit_interpreter_path,
                state.button_interpreter_browse,
                state.label_subtitle_mode,
                state.combo_subtitle_mode,
                state.label_subtitle_offset,
                state.edit_subtitle_offset,
                state.checkbox_move_cursor,
                state.checkbox_check_updates,
                state.checkbox_check_beta_updates,
                state.checkbox_send_crash_reports,
                state.checkbox_use_legacy_name,
                state.checkbox_context_menu,
                state.combo_confirm_delete_rss_mode,
                state.combo_confirm_delete_podcast_mode,
                state.combo_rss_quick_copy_mode,
                state.label_prompt_program,
                state.combo_prompt_program,
                state.label_network_proxy,
                state.edit_network_proxy,
                state.label_network_proxy_username,
                state.edit_network_proxy_username,
                state.label_network_proxy_password,
                state.edit_network_proxy_password,
                state.combo_shortcut_action,
                state.edit_shortcut_value,
            )
        }) {
            Some(values) => values,
            None => return,
        };

        let mut settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
        let missing_dialogue_settings = settings.dialogue_voice.trim().is_empty()
            || settings.dialogue_opening_quote.trim().is_empty()
            || settings.dialogue_closing_quote.trim().is_empty();
        if missing_dialogue_settings
            && let Some(cfg) = crate::dialogue_voice::load_dialogue_voice_config()
        {
            settings.dialogue_tts_engine = cfg.engine;
            settings.dialogue_voice = cfg.voice;
            settings.dialogue_use_secondary_voice = cfg.use_secondary_voice;
            settings.dialogue_secondary_voice = cfg.secondary_voice;
            settings.dialogue_secondary_tts_engine = cfg.secondary_engine;
            settings.dialogue_secondary_voice_rate = cfg.secondary_rate;
            settings.dialogue_secondary_voice_pitch = cfg.secondary_pitch;
            settings.dialogue_secondary_voice_volume = cfg.secondary_volume;
            settings.dialogue_voice_rate = cfg.rate;
            settings.dialogue_voice_pitch = cfg.pitch;
            settings.dialogue_voice_volume = cfg.volume;
            settings.dialogue_opening_quote = cfg.opening_quote;
            settings.dialogue_closing_quote = cfg.closing_quote;
            settings.dialogue_allow_multiline = cfg.allow_multiline;
        }
        let labels = options_labels(settings.language);

        let mut voice_profiles = settings.voice_profiles.clone();
        if !voice_profiles
            .iter()
            .any(|p| p.name.eq_ignore_ascii_case(DEFAULT_VOICE_PROFILE_NAME))
        {
            voice_profiles.push(voice_profile_from_settings_fields(
                DEFAULT_VOICE_PROFILE_NAME.to_string(),
                &settings,
            ));
        }
        let mut active_voice_profile_name = settings.active_voice_profile.trim().to_string();
        if active_voice_profile_name.is_empty()
            || !voice_profiles
                .iter()
                .any(|p| p.name.eq_ignore_ascii_case(&active_voice_profile_name))
        {
            active_voice_profile_name = DEFAULT_VOICE_PROFILE_NAME.to_string();
        }
        if with_options_state(hwnd, |state| {
            state.voice_profiles = voice_profiles;
            state.active_voice_profile_name = active_voice_profile_name;
        })
        .is_none()
        {
            crate::log_debug("Failed to initialize voice profile state");
        }
        refresh_voice_profile_combo(hwnd);

        SendMessageW(combo_lang, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_it).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_en).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_es).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_pt).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_sv).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_vi).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_cs).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_pl).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_fr).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_sr).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_uk).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_lt).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_ru).as_ptr() as isize),
        );
        SendMessageW(
            combo_lang,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.lang_zh).as_ptr() as isize),
        );

        let lang_index = match settings.language {
            Language::Italian => 0,
            Language::English => 1,
            Language::Spanish => 2,
            Language::Portuguese => 3,
            Language::Swedish => 4,
            Language::Vietnamese => 5,
            Language::Czech => 6,
            Language::Polish => 7,
            Language::French => 8,
            Language::Serbian => 9,
            Language::Ukrainian => 10,
            Language::Lithuanian => 11,
            Language::Russian => 12,
            Language::Chinese => 13,
        };
        SendMessageW(combo_lang, CB_SETCURSEL, WPARAM(lang_index), LPARAM(0));

        SendMessageW(combo_lang, CB_SETCURSEL, WPARAM(lang_index), LPARAM(0));

        SendMessageW(
            combo_modified_marker_position,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        SendMessageW(
            combo_modified_marker_position,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.marker_position_end).as_ptr() as isize),
        );
        SendMessageW(
            combo_modified_marker_position,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.marker_position_beginning).as_ptr() as isize),
        );
        let position_index = match settings.modified_marker_position {
            ModifiedMarkerPosition::Beginning => 1,
            _ => 0,
        };
        SendMessageW(
            combo_modified_marker_position,
            CB_SETCURSEL,
            WPARAM(position_index),
            LPARAM(0),
        );

        SendMessageW(combo_open, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(
            combo_open,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.open_new_tab).as_ptr() as isize),
        );
        SendMessageW(
            combo_open,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.open_new_window).as_ptr() as isize),
        );
        let open_index = match settings.open_behavior {
            OpenBehavior::NewTab => 0,
            OpenBehavior::NewWindow => 1,
        };
        SendMessageW(combo_open, CB_SETCURSEL, WPARAM(open_index), LPARAM(0));

        SendMessageW(combo_tts_engine, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(
            combo_tts_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_edge).as_ptr() as isize),
        );
        SendMessageW(
            combo_tts_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi5).as_ptr() as isize),
        );
        SendMessageW(
            combo_tts_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi4).as_ptr() as isize),
        );

        let engine_index = match settings.tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(
            combo_tts_engine,
            CB_SETCURSEL,
            WPARAM(engine_index),
            LPARAM(0),
        );

        SendMessageW(combo_dialogue_engine, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(
            combo_dialogue_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_edge).as_ptr() as isize),
        );
        SendMessageW(
            combo_dialogue_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi5).as_ptr() as isize),
        );
        SendMessageW(
            combo_dialogue_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi4).as_ptr() as isize),
        );
        let dialogue_engine_index = match settings.dialogue_tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(
            combo_dialogue_engine,
            CB_SETCURSEL,
            WPARAM(dialogue_engine_index),
            LPARAM(0),
        );
        SendMessageW(
            combo_dialogue_secondary_engine,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        SendMessageW(
            combo_dialogue_secondary_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_edge).as_ptr() as isize),
        );
        SendMessageW(
            combo_dialogue_secondary_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi5).as_ptr() as isize),
        );
        SendMessageW(
            combo_dialogue_secondary_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi4).as_ptr() as isize),
        );
        let dialogue_secondary_engine_index = match settings.dialogue_secondary_tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(
            combo_dialogue_secondary_engine,
            CB_SETCURSEL,
            WPARAM(dialogue_secondary_engine_index),
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
        init_tts_combo(combo_dialogue_voice_rate, &speed_items);
        init_tts_combo(combo_dialogue_voice_pitch, &pitch_items);
        init_tts_combo(combo_dialogue_voice_volume, &volume_items);
        init_tts_combo(combo_dialogue_secondary_voice_rate, &speed_items);
        init_tts_combo(combo_dialogue_secondary_voice_pitch, &pitch_items);
        init_tts_combo(combo_dialogue_secondary_voice_volume, &volume_items);
        select_combo_value(combo_tts_speed, settings.tts_rate);
        select_combo_value(combo_tts_pitch, settings.tts_pitch);
        select_combo_value(combo_tts_volume, settings.tts_volume);
        select_combo_value(combo_dialogue_voice_rate, settings.dialogue_voice_rate);
        select_combo_value(combo_dialogue_voice_pitch, settings.dialogue_voice_pitch);
        select_combo_value(combo_dialogue_voice_volume, settings.dialogue_voice_volume);
        select_combo_value(
            combo_dialogue_secondary_voice_rate,
            settings.dialogue_secondary_voice_rate,
        );
        select_combo_value(
            combo_dialogue_secondary_voice_pitch,
            settings.dialogue_secondary_voice_pitch,
        );
        select_combo_value(
            combo_dialogue_secondary_voice_volume,
            settings.dialogue_secondary_voice_volume,
        );
        if let Err(_e) = SetWindowTextW(
            edit_tts_speed,
            PCWSTR(to_wide(&tts_ui_value_from_internal(settings.tts_rate).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set speed text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_tts_pitch,
            PCWSTR(to_wide(&tts_ui_value_from_internal(settings.tts_pitch).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set pitch text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_tts_volume,
            PCWSTR(to_wide(&settings.tts_volume.to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set volume text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_voice_rate,
            PCWSTR(
                to_wide(&tts_ui_value_from_internal(settings.dialogue_voice_rate).to_string())
                    .as_ptr(),
            ),
        ) {
            crate::log_debug(&format!("Failed to set dialogue speed text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_voice_pitch,
            PCWSTR(
                to_wide(&tts_ui_value_from_internal(settings.dialogue_voice_pitch).to_string())
                    .as_ptr(),
            ),
        ) {
            crate::log_debug(&format!("Failed to set dialogue pitch text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_voice_volume,
            PCWSTR(to_wide(&settings.dialogue_voice_volume.to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set dialogue volume text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_secondary_voice_rate,
            PCWSTR(
                to_wide(
                    &tts_ui_value_from_internal(settings.dialogue_secondary_voice_rate).to_string(),
                )
                .as_ptr(),
            ),
        ) {
            crate::log_debug(&format!(
                "Failed to set secondary dialogue speed text: {:?}",
                _e
            ));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_secondary_voice_pitch,
            PCWSTR(
                to_wide(
                    &tts_ui_value_from_internal(settings.dialogue_secondary_voice_pitch)
                        .to_string(),
                )
                .as_ptr(),
            ),
        ) {
            crate::log_debug(&format!(
                "Failed to set secondary dialogue pitch text: {:?}",
                _e
            ));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_secondary_voice_volume,
            PCWSTR(to_wide(&settings.dialogue_secondary_voice_volume.to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!(
                "Failed to set secondary dialogue volume text: {:?}",
                _e
            ));
        }
        SendMessageW(
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

        SendMessageW(
            checkbox_multilingual,
            BM_SETCHECK,
            WPARAM(if settings.tts_only_multilingual {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_use_dialogue_voice,
            BM_SETCHECK,
            WPARAM(if settings.use_dialogue_voice {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_dialogue_multilingual,
            BM_SETCHECK,
            WPARAM(0),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_dialogue_use_secondary_voice,
            BM_SETCHECK,
            WPARAM(if settings.dialogue_use_secondary_voice {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_dialogue_secondary_multilingual,
            BM_SETCHECK,
            WPARAM(0),
            LPARAM(0),
        );
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_open_quote,
            PCWSTR(to_wide(&settings.dialogue_opening_quote).as_ptr()),
        ) {
            crate::log_debug(&format!(
                "Failed to set dialogue opening quote text: {:?}",
                _e
            ));
        }
        if let Err(_e) = SetWindowTextW(
            edit_dialogue_close_quote,
            PCWSTR(to_wide(&settings.dialogue_closing_quote).as_ptr()),
        ) {
            crate::log_debug(&format!(
                "Failed to set dialogue closing quote text: {:?}",
                _e
            ));
        }
        SendMessageW(
            checkbox_dialogue_allow_multiline,
            BM_SETCHECK,
            WPARAM(if settings.dialogue_allow_multiline {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        // Populate dialogue voice combo - will be filled in refresh_voices
        SendMessageW(combo_dialogue_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(
            combo_dialogue_secondary_voice,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        update_dialogue_voice_visibility(hwnd);

        SendMessageW(
            checkbox_audio_split_requires_newline,
            BM_SETCHECK,
            WPARAM(if settings.audiobook_split_text_requires_newline {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_audio_split_epub_chapters,
            BM_SETCHECK,
            WPARAM(if settings.audiobook_split_by_epub_chapter {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_subtitle_ducking,
            BM_SETCHECK,
            WPARAM(if settings.subtitle_mix_ducking {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_split_on_newline,
            BM_SETCHECK,
            WPARAM(if settings.split_on_newline {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_word_wrap,
            BM_SETCHECK,
            WPARAM(if settings.word_wrap {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_smart_quotes,
            BM_SETCHECK,
            WPARAM(if settings.smart_quotes {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_strip_markdown_keep_bullets,
            BM_SETCHECK,
            WPARAM(if settings.strip_markdown_keep_bullets {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_spellcheck,
            BM_SETCHECK,
            WPARAM(if settings.spellcheck_enabled {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
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
            ("Polski (Polska)".to_string(), "pl-PL"),
            (labels.spellcheck_lang_ru.clone(), "ru-RU"),
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
            SendMessageW(
                combo_spellcheck_language,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            );
            if *val == current_val {
                selected_idx = i;
            }
        }
        SendMessageW(
            combo_spellcheck_language,
            CB_SETCURSEL,
            WPARAM(selected_idx),
            LPARAM(0),
        );
        update_spellcheck_language_visibility(hwnd);

        SendMessageW(
            combo_dictionary_translation,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        let dictionary_translation_options = [
            (labels.dictionary_translation_auto.clone(), "auto"),
            (labels.dictionary_translation_none.clone(), "none"),
            (labels.lang_it.clone(), "it"),
            (labels.lang_en.clone(), "en"),
            (labels.lang_es.clone(), "es"),
            (labels.lang_pt.clone(), "pt"),
            (labels.lang_sv.clone(), "sv"),
            (labels.lang_vi.clone(), "vi"),
            (labels.lang_cs.clone(), "cs"),
            (labels.lang_pl.clone(), "pl"),
            (labels.lang_fr.clone(), "fr"),
            (labels.lang_uk.clone(), "uk"),
            (labels.lang_lt.clone(), "lt"),
            (labels.lang_ru.clone(), "ru"),
            (labels.lang_zh.clone(), "zh"),
        ];
        let current_dict_lang = settings
            .dictionary_translation_language
            .trim()
            .to_ascii_lowercase();
        let mut dict_selected_idx = 0;
        for (i, (label, val)) in dictionary_translation_options.iter().enumerate() {
            SendMessageW(
                combo_dictionary_translation,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            );
            if *val == current_dict_lang {
                dict_selected_idx = i;
            }
        }
        SendMessageW(
            combo_dictionary_translation,
            CB_SETCURSEL,
            WPARAM(dict_selected_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_wikipedia_language,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        let wikipedia_language_options = [
            (labels.wikipedia_language_auto.clone(), "auto"),
            (labels.lang_it.clone(), "it"),
            (labels.lang_en.clone(), "en"),
            (labels.lang_es.clone(), "es"),
            (labels.lang_pt.clone(), "pt"),
            (labels.lang_sv.clone(), "sv"),
            (labels.lang_vi.clone(), "vi"),
            (labels.lang_cs.clone(), "cs"),
            (labels.lang_pl.clone(), "pl"),
            (labels.lang_fr.clone(), "fr"),
            (labels.lang_uk.clone(), "uk"),
            (labels.lang_lt.clone(), "lt"),
            (labels.lang_ru.clone(), "ru"),
            (labels.lang_zh.clone(), "zh"),
        ];
        let current_wikipedia_lang = settings.wikipedia_language.trim().to_ascii_lowercase();
        let mut wiki_selected_idx = 0;
        for (i, (label, val)) in wikipedia_language_options.iter().enumerate() {
            SendMessageW(
                combo_wikipedia_language,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            );
            if *val == current_wikipedia_lang {
                wiki_selected_idx = i;
            }
        }
        SendMessageW(
            combo_wikipedia_language,
            CB_SETCURSEL,
            WPARAM(wiki_selected_idx),
            LPARAM(0),
        );

        let wrap_text = settings.wrap_width.to_string();
        if let Err(_e) = SetWindowTextW(edit_wrap_width, PCWSTR(to_wide(&wrap_text).as_ptr())) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }

        SendMessageW(combo_indentation, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let indent_options = [
            labels.indent_default.clone(),
            labels.indent_tabs.clone(),
            labels.indent_spaces.clone(),
        ];
        for label in indent_options.iter() {
            SendMessageW(
                combo_indentation,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            );
        }
        let indent_index = match settings.indentation_mode {
            crate::settings::IndentationMode::Default => 0,
            crate::settings::IndentationMode::Tabs => 1,
            crate::settings::IndentationMode::Spaces => 2,
        };
        SendMessageW(
            combo_indentation,
            CB_SETCURSEL,
            WPARAM(indent_index),
            LPARAM(0),
        );

        SendMessageW(combo_tab_width, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(combo_space_width, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let mut tab_sel = 0usize;
        let mut space_sel = 0usize;
        for (idx, width) in [2u32, 4, 6, 8].iter().enumerate() {
            let label = width.to_string();
            let tab_idx = SendMessageW(
                combo_tab_width,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_tab_width,
                CB_SETITEMDATA,
                WPARAM(tab_idx),
                LPARAM(*width as isize),
            );
            if *width == settings.indent_tab_width {
                tab_sel = idx;
            }

            let space_idx = SendMessageW(
                combo_space_width,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_space_width,
                CB_SETITEMDATA,
                WPARAM(space_idx),
                LPARAM(*width as isize),
            );
            if *width == settings.indent_space_width {
                space_sel = idx;
            }
        }
        SendMessageW(combo_tab_width, CB_SETCURSEL, WPARAM(tab_sel), LPARAM(0));
        SendMessageW(
            combo_space_width,
            CB_SETCURSEL,
            WPARAM(space_sel),
            LPARAM(0),
        );
        update_indentation_visibility(hwnd);

        if let Err(_e) = SetWindowTextW(
            edit_quote_prefix,
            PCWSTR(to_wide(&settings.quote_prefix).as_ptr()),
        ) {}
        if let Err(_e) = SetWindowTextW(
            edit_interpreter_path,
            PCWSTR(to_wide(&settings.interpreter_path).as_ptr()),
        ) {}
        if let Err(_e) = SetWindowTextW(
            edit_subtitle_offset,
            PCWSTR(to_wide(&settings.subtitle_offset_ms.to_string()).as_ptr()),
        ) {}
        SendMessageW(combo_subtitle_mode, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let subtitle_mode_options = [
            labels.subtitle_mode_off.clone(),
            labels.subtitle_mode_nvda.clone(),
            labels.subtitle_mode_user.clone(),
            labels.subtitle_mode_record.clone(),
        ];
        for label in subtitle_mode_options.iter() {
            SendMessageW(
                combo_subtitle_mode,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            );
        }
        let subtitle_mode_index = match settings.subtitle_read_mode {
            SubtitleReadMode::Off => 0,
            SubtitleReadMode::Nvda => 1,
            SubtitleReadMode::User => 2,
            SubtitleReadMode::Record => 3,
            _ => 2,
        };
        SendMessageW(
            combo_subtitle_mode,
            CB_SETCURSEL,
            WPARAM(subtitle_mode_index),
            LPARAM(0),
        );
        update_subtitle_ducking_visibility(hwnd);
        SendMessageW(
            checkbox_move_cursor,
            BM_SETCHECK,
            WPARAM(if settings.move_cursor_during_reading {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_check_updates,
            BM_SETCHECK,
            WPARAM(if settings.check_updates_on_startup {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_check_beta_updates,
            BM_SETCHECK,
            WPARAM(if settings.check_beta_updates_on_startup {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_send_crash_reports,
            BM_SETCHECK,
            WPARAM(if settings.send_crash_reports {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_use_legacy_name,
            BM_SETCHECK,
            WPARAM(if settings.use_legacy_name {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_context_menu,
            BM_SETCHECK,
            WPARAM(if settings.context_menu_open_with {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            combo_confirm_delete_rss_mode,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.confirm_delete_feed.clone(),
            labels.confirm_delete_article.clone(),
            labels.confirm_delete_both.clone(),
            labels.confirm_delete_none.clone(),
        ] {
            SendMessageW(
                combo_confirm_delete_rss_mode,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let rss_confirm_idx = match settings.rss_delete_confirm_mode {
            RssDeleteConfirmMode::Feed => 0,
            RssDeleteConfirmMode::Article => 1,
            RssDeleteConfirmMode::Both => 2,
            RssDeleteConfirmMode::None => 3,
        };
        SendMessageW(
            combo_confirm_delete_rss_mode,
            CB_SETCURSEL,
            WPARAM(rss_confirm_idx),
            LPARAM(0),
        );
        SendMessageW(
            combo_confirm_delete_podcast_mode,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.confirm_delete_podcast.clone(),
            labels.confirm_delete_episode.clone(),
            labels.confirm_delete_both.clone(),
            labels.confirm_delete_none.clone(),
        ] {
            SendMessageW(
                combo_confirm_delete_podcast_mode,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let podcast_confirm_idx = match settings.podcast_delete_confirm_mode {
            PodcastDeleteConfirmMode::Podcast => 0,
            PodcastDeleteConfirmMode::Episode => 1,
            PodcastDeleteConfirmMode::Both => 2,
            PodcastDeleteConfirmMode::None => 3,
        };
        SendMessageW(
            combo_confirm_delete_podcast_mode,
            CB_SETCURSEL,
            WPARAM(podcast_confirm_idx),
            LPARAM(0),
        );
        SendMessageW(
            combo_rss_quick_copy_mode,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.rss_quick_copy_title.clone(),
            labels.rss_quick_copy_url.clone(),
            labels.rss_quick_copy_content.clone(),
            labels.rss_quick_copy_all.clone(),
        ] {
            SendMessageW(
                combo_rss_quick_copy_mode,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let rss_quick_copy_idx = match settings.rss_quick_copy_mode {
            crate::settings::RssQuickCopyMode::Title => 0,
            crate::settings::RssQuickCopyMode::Url => 1,
            crate::settings::RssQuickCopyMode::Content => 2,
            crate::settings::RssQuickCopyMode::All => 3,
        };
        SendMessageW(
            combo_rss_quick_copy_mode,
            CB_SETCURSEL,
            WPARAM(rss_quick_copy_idx),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_rss_show_article_preview,
            BM_SETCHECK,
            WPARAM(if settings.rss_show_article_preview {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            checkbox_announce_unread_rss_podcast,
            BM_SETCHECK,
            WPARAM(if settings.announce_unread_rss_podcast_items {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );
        SendMessageW(
            combo_unread_label_position,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.unread_label_position_before.clone(),
            labels.unread_label_position_after.clone(),
        ] {
            SendMessageW(
                combo_unread_label_position,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let unread_label_idx = match settings.rss_podcast_unread_label_position {
            RssPodcastUnreadLabelPosition::Before => 0,
            RssPodcastUnreadLabelPosition::After => 1,
        };
        SendMessageW(
            combo_unread_label_position,
            CB_SETCURSEL,
            WPARAM(unread_label_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_rss_date_display,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.list_date_always.clone(),
            labels.list_date_never.clone(),
        ] {
            SendMessageW(
                combo_rss_date_display,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let rss_date_idx = match settings.rss_articles_date_display {
            ListDateDisplayMode::Always => 0,
            ListDateDisplayMode::Never => 1,
        };
        SendMessageW(
            combo_rss_date_display,
            CB_SETCURSEL,
            WPARAM(rss_date_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_rss_time_display,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.list_time_always.clone(),
            labels.list_time_never.clone(),
            labels.list_time_only_if_multiple_same_day.clone(),
        ] {
            SendMessageW(
                combo_rss_time_display,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let rss_time_idx = match settings.rss_articles_time_display {
            ListTimeDisplayMode::Always => 0,
            ListTimeDisplayMode::Never => 1,
            ListTimeDisplayMode::OnlyIfMultipleSameDay => 2,
        };
        SendMessageW(
            combo_rss_time_display,
            CB_SETCURSEL,
            WPARAM(rss_time_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_podcast_date_display,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.list_date_always.clone(),
            labels.list_date_never.clone(),
        ] {
            SendMessageW(
                combo_podcast_date_display,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let podcast_date_idx = match settings.podcast_episodes_date_display {
            ListDateDisplayMode::Always => 0,
            ListDateDisplayMode::Never => 1,
        };
        SendMessageW(
            combo_podcast_date_display,
            CB_SETCURSEL,
            WPARAM(podcast_date_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_podcast_time_display,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for option in [
            labels.list_time_always.clone(),
            labels.list_time_never.clone(),
            labels.list_time_only_if_multiple_same_day.clone(),
        ] {
            SendMessageW(
                combo_podcast_time_display,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&option).as_ptr() as isize),
            );
        }
        let podcast_time_idx = match settings.podcast_episodes_time_display {
            ListTimeDisplayMode::Always => 0,
            ListTimeDisplayMode::Never => 1,
            ListTimeDisplayMode::OnlyIfMultipleSameDay => 2,
        };
        SendMessageW(
            combo_podcast_time_display,
            CB_SETCURSEL,
            WPARAM(podcast_time_idx),
            LPARAM(0),
        );

        SendMessageW(combo_prompt_program, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let prompt_options = [
            labels.prompt_cmd.clone(),
            labels.prompt_powershell.clone(),
            labels.prompt_codex.clone(),
        ];
        for label in prompt_options.iter() {
            SendMessageW(
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
        SendMessageW(
            combo_prompt_program,
            CB_SETCURSEL,
            WPARAM(program_idx),
            LPARAM(0),
        );
        if let Err(_e) = SetWindowTextW(
            edit_network_proxy,
            PCWSTR(to_wide(&settings.network_proxy_url).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set network proxy text: {:?}", _e));
        }
        if let Err(_e) = SetWindowTextW(
            edit_network_proxy_username,
            PCWSTR(to_wide(&settings.network_proxy_username).as_ptr()),
        ) {
            crate::log_debug(&format!(
                "Failed to set network proxy username text: {:?}",
                _e
            ));
        }
        if let Err(_e) = SetWindowTextW(
            edit_network_proxy_password,
            PCWSTR(to_wide(&settings.network_proxy_password).as_ptr()),
        ) {
            crate::log_debug(&format!(
                "Failed to set network proxy password text: {:?}",
                _e
            ));
        }
        if with_options_state(hwnd, |state| {
            state.shortcut_draft = settings.shortcuts.clone();
            state.shortcut_capture_pending = false;
        })
        .is_none()
        {
            crate::log_debug("Failed to access state in options_window");
        }
        SendMessageW(combo_shortcut_action, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for action in ShortcutAction::ALL {
            SendMessageW(
                combo_shortcut_action,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(
                    to_wide(&shortcut_action_label(settings.language, action)).as_ptr() as isize,
                ),
            );
        }
        SendMessageW(combo_shortcut_action, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        update_shortcut_binding_text(hwnd);

        SendMessageW(combo_audio_skip, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let skip_options = [
            (1, "1 s"),
            (2, "2 s"),
            (3, "3 s"),
            (4, "4 s"),
            (5, "5 s"),
            (6, "6 s"),
            (7, "7 s"),
            (8, "8 s"),
            (9, "9 s"),
            (10, "10 s"),
            (15, "15 s"),
            (20, "20 s"),
            (30, "30 s"),
            (45, "45 s"),
            (60, "1 m"),
            (90, "1.5 m"),
            (120, "2 m"),
            (180, "3 m"),
            (300, "5 m"),
            (600, "10 m"),
            (900, "15 m"),
            (1800, "30 m"),
            (3600, "1 h"),
            (7200, "2 h"),
        ];
        let mut selected_idx = 0;
        for (secs, label) in skip_options.iter() {
            let idx = SendMessageW(
                combo_audio_skip,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_audio_skip,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(*secs as isize),
            );
            if *secs == 60 || *secs == settings.audiobook_skip_seconds {
                selected_idx = idx;
            }
        }
        SendMessageW(
            combo_audio_skip,
            CB_SETCURSEL,
            WPARAM(selected_idx),
            LPARAM(0),
        );
        if with_options_state(hwnd, |state| {
            state.default_save_folder_selection = DEFAULT_SAVE_FOLDER_AUDIOBOOK;
            state.default_save_folder_audiobook = settings.audiobook_save_folder.clone();
            state.default_save_folder_media = settings.media_save_folder.clone();
            state.default_save_folder_documents = settings.documents_save_folder.clone();
        })
        .is_none()
        {
            crate::log_debug("Failed to access state in options_window");
        }
        SendMessageW(
            combo_default_save_folder_kind,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        for (kind, label) in [
            (
                DEFAULT_SAVE_FOLDER_AUDIOBOOK,
                labels.option_default_save_folder_audiobooks.clone(),
            ),
            (
                DEFAULT_SAVE_FOLDER_MEDIA,
                labels.option_default_save_folder_media.clone(),
            ),
            (
                DEFAULT_SAVE_FOLDER_DOCUMENTS,
                labels.option_default_save_folder_documents.clone(),
            ),
        ] {
            let idx = SendMessageW(
                combo_default_save_folder_kind,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_default_save_folder_kind,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(kind as isize),
            );
        }
        SendMessageW(
            combo_default_save_folder_kind,
            CB_SETCURSEL,
            WPARAM(0),
            LPARAM(0),
        );
        if let Err(_e) = SetWindowTextW(
            edit_audiobook_save_folder,
            PCWSTR(to_wide(&settings.audiobook_save_folder).as_ptr()),
        ) {}
        SendMessageW(
            checkbox_show_media_save_confirmation,
            BM_SETCHECK,
            WPARAM(if settings.show_media_save_confirmation {
                BST_CHECKED.0 as usize
            } else {
                0
            }),
            LPARAM(0),
        );

        SendMessageW(combo_audio_split, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let split_options = [
            (0, labels.split_none.clone()),
            (AUDIOBOOK_SPLIT_BY_TIME, labels.split_by_time.clone()),
            (AUDIOBOOK_SPLIT_BY_TEXT, labels.split_by_text.clone()),
            (AUDIOBOOK_SPLIT_BY_PARTS, labels.split_by_parts.clone()),
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
            SendMessageW(
                combo_audio_split,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(*parts as isize),
            );
            if (settings.audiobook_split_by_time && *parts == AUDIOBOOK_SPLIT_BY_TIME)
                || (settings.audiobook_split_by_text && *parts == AUDIOBOOK_SPLIT_BY_TEXT)
                || (!settings.audiobook_split_by_time
                    && !settings.audiobook_split_by_text
                    && settings.audiobook_split > 0
                    && *parts == AUDIOBOOK_SPLIT_BY_PARTS)
            {
                selected_split_idx = idx;
            }
        }
        SendMessageW(
            combo_audio_split,
            CB_SETCURSEL,
            WPARAM(selected_split_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_audio_split_minutes,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        let mut selected_minutes_idx = 0;
        for minutes in 1..=60u32 {
            let label = format!("{minutes}");
            let idx = SendMessageW(
                combo_audio_split_minutes,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_audio_split_minutes,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(minutes as isize),
            );
            if minutes == settings.audiobook_split_minutes {
                selected_minutes_idx = idx;
            }
        }
        SendMessageW(
            combo_audio_split_minutes,
            CB_SETCURSEL,
            WPARAM(selected_minutes_idx),
            LPARAM(0),
        );

        let split_parts_value = if settings.audiobook_split == 0 {
            2
        } else {
            settings.audiobook_split.clamp(1, 100)
        };
        let split_parts_wide = to_wide(&split_parts_value.to_string());
        if let Err(_e) = SetWindowTextW(
            edit_audio_split_parts_count,
            PCWSTR(split_parts_wide.as_ptr()),
        ) {}

        SendMessageW(
            combo_audio_split_start_number,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        let mut selected_start_idx = 0;
        for number in 1..=99u32 {
            let label = format!("{number:02}");
            let idx = SendMessageW(
                combo_audio_split_start_number,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_audio_split_start_number,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(number as isize),
            );
            if number == settings.audiobook_split_start_number {
                selected_start_idx = idx;
            }
        }
        SendMessageW(
            combo_audio_split_start_number,
            CB_SETCURSEL,
            WPARAM(selected_start_idx),
            LPARAM(0),
        );

        SendMessageW(
            combo_audiobook_part_naming,
            CB_RESETCONTENT,
            WPARAM(0),
            LPARAM(0),
        );
        let naming_options = [
            (
                AudiobookPartNamingMode::TitleNumber,
                labels.option_audiobook_part_naming_title_number.clone(),
            ),
            (
                AudiobookPartNamingMode::NumberOnly,
                labels.option_audiobook_part_naming_number_only.clone(),
            ),
            (
                AudiobookPartNamingMode::NumberTitle,
                labels.option_audiobook_part_naming_number_title.clone(),
            ),
        ];
        let mut selected_naming_idx = 0usize;
        for (mode, label) in naming_options {
            let idx = SendMessageW(
                combo_audiobook_part_naming,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                combo_audiobook_part_naming,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(mode as isize),
            );
            if mode == settings.audiobook_part_naming_mode {
                selected_naming_idx = idx;
            }
        }
        SendMessageW(
            combo_audiobook_part_naming,
            CB_SETCURSEL,
            WPARAM(selected_naming_idx),
            LPARAM(0),
        );

        let split_text_wide = to_wide(&settings.audiobook_split_text);
        if let Err(_e) = SetWindowTextW(edit_audio_split_text, PCWSTR(split_text_wide.as_ptr())) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        update_audio_split_visibility(hwnd);

        let cache_limit_text = settings.podcast_cache_limit_mb.to_string();
        if let Err(_e) = SetWindowTextW(
            edit_podcast_cache_limit,
            PCWSTR(to_wide(&cache_limit_text).as_ptr()),
        ) {}
        if let Err(_e) = SetWindowTextW(
            edit_podcastindex_key,
            PCWSTR(to_wide(&settings.podcast_index_api_key).as_ptr()),
        ) {}
        let secret =
            crate::settings::decrypt_podcast_index_secret(&settings.podcast_index_api_secret)
                .unwrap_or_default();
        if let Err(_e) = SetWindowTextW(edit_podcastindex_secret, PCWSTR(to_wide(&secret).as_ptr()))
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        let whisper_index = match settings.whisper_model_profile.as_str() {
            "medium_q5_0" => 1usize,
            "large_v3_turbo_q5_0" => 2usize,
            _ => 0usize,
        };
        SendMessageW(
            combo_whisper_model,
            CB_SETCURSEL,
            WPARAM(whisper_index),
            LPARAM(0),
        );
        if let Some(checkbox_whisper_cuda) =
            with_options_state(hwnd, |state| state.checkbox_whisper_cuda)
        {
            SendMessageW(
                checkbox_whisper_cuda,
                BM_SETCHECK,
                WPARAM(if settings.whisper_cuda_enabled {
                    BST_CHECKED.0 as usize
                } else {
                    0
                }),
                LPARAM(0),
            );
        }
        if let Some(checkbox_whisper_keep_original_language) =
            with_options_state(hwnd, |state| state.checkbox_whisper_keep_original_language)
        {
            SendMessageW(
                checkbox_whisper_keep_original_language,
                BM_SETCHECK,
                WPARAM(if settings.whisper_keep_original_language {
                    BST_CHECKED.0 as usize
                } else {
                    0
                }),
                LPARAM(0),
            );
        }
        if let Some(checkbox_whisper_include_timestamps) =
            with_options_state(hwnd, |state| state.checkbox_whisper_include_timestamps)
        {
            SendMessageW(
                checkbox_whisper_include_timestamps,
                BM_SETCHECK,
                WPARAM(if settings.whisper_include_timestamps {
                    BST_CHECKED.0 as usize
                } else {
                    0
                }),
                LPARAM(0),
            );
        }
        if let Some((combo_dictation_microphone, device_ids)) = with_options_state(hwnd, |state| {
            (
                state.combo_dictation_microphone,
                state.dictation_microphone_device_ids.clone(),
            )
        }) {
            let selected = device_ids
                .iter()
                .position(|id| id == &settings.dictation_microphone_device_id)
                .unwrap_or(0);
            SendMessageW(
                combo_dictation_microphone,
                CB_SETCURSEL,
                WPARAM(selected),
                LPARAM(0),
            );
        }

        refresh_voices(hwnd);
    }
}

fn populate_voice_combo(
    combo_voice: HWND,
    voices: &[VoiceInfo],
    selected: &str,
    only_multilingual: bool,
    language_filter: Option<&str>,
    labels: &OptionsLabels,
) {
    unsafe {
        SendMessageW(combo_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        if voices.is_empty() {
            let label = &labels.voices_empty;
            // We could also check if it's loading, but SAPI loads fast.
            // For Edge, it might be loading.
            // We can check if "loading" logic is needed, but "voices_empty" is safe default.
            SendMessageW(
                combo_voice,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            );
            SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
            return;
        }
        let mut selected_index: Option<usize> = None;
        let mut combo_index = 0usize;

        for (voice_index, voice) in voices.iter().enumerate() {
            if only_multilingual && !voice.is_multilingual {
                continue;
            }
            if let Some(filter) = language_filter {
                let Some(code) = voice_locale_language_code(&voice.locale) else {
                    continue;
                };
                if code != filter {
                    continue;
                }
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
                SendMessageW(
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
            SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
        } else if combo_index > 0 {
            SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        }
    }
}

fn init_tts_combo(hwnd: HWND, items: &[(String, i32)]) {
    unsafe {
        SendMessageW(hwnd, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for (label, value) in items {
            let idx = SendMessageW(
                hwnd,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(label).as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(hwnd, CB_SETITEMDATA, WPARAM(idx), LPARAM(*value as isize));
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
                SendMessageW(hwnd, CB_SETCURSEL, WPARAM(i as usize), LPARAM(0));
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

fn selected_voice_short_name_from_combo_text(combo: HWND) -> Option<String> {
    let len = crate::get_window_text_length_w_safe(combo);
    if len <= 0 {
        return None;
    }

    let mut buf = vec![0u16; (len + 1) as usize];
    let read = crate::get_window_text_w_safe(combo, &mut buf);
    if read <= 0 {
        return None;
    }

    let label = String::from_utf16_lossy(&buf[..read as usize]);
    let short_name = label.split(" (").next().unwrap_or("").trim();
    if short_name.is_empty() {
        None
    } else {
        Some(short_name.to_string())
    }
}

fn selected_shortcut_action(hwnd: HWND) -> ShortcutAction {
    let sel = unsafe {
        with_options_state(hwnd, |state| {
            SendMessageW(
                state.combo_shortcut_action,
                CB_GETCURSEL,
                WPARAM(0),
                LPARAM(0),
            )
            .0
        })
        .unwrap_or(0)
    };
    let idx = if sel < 0 { 0usize } else { sel as usize };
    ShortcutAction::ALL
        .get(idx)
        .copied()
        .unwrap_or(ShortcutAction::ReadPauseResume)
}

fn update_shortcut_binding_text(hwnd: HWND) {
    let action = selected_shortcut_action(hwnd);
    let (edit, binding, waiting, language) = with_options_state(hwnd, |state| {
        let language =
            { with_state(state.parent, |app| app.settings.language) }.unwrap_or_default();
        (
            state.edit_shortcut_value,
            shortcut_binding_for_action(&state.shortcut_draft, action),
            state.shortcut_capture_pending,
            language,
        )
    })
    .unwrap_or((
        HWND(0),
        ShortcutBinding::new(false, false, false, 0),
        false,
        Language::English,
    ));
    if edit.0 == 0 {
        return;
    }
    let text = if waiting {
        i18n::tr(language, "options.shortcuts.capture_hint")
    } else {
        format_shortcut(binding)
    };
    crate::log_if_err!(crate::set_window_text_w_safe(
        edit,
        PCWSTR(to_wide(&text).as_ptr())
    ));
}

const TTS_RATE_MIN: i32 = -100;
const TTS_RATE_MAX: i32 = 100;
const TTS_PITCH_MIN: i32 = -12;
const TTS_PITCH_MAX: i32 = 12;
const TTS_VOLUME_MIN: i32 = 25;
const TTS_VOLUME_MAX: i32 = 200;
const TTS_UI_OFFSET: i32 = 100;

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

fn tts_ui_value_from_internal(value: i32) -> i32 {
    value + TTS_UI_OFFSET
}

fn read_tts_tuning_edit_value(edit: HWND, fallback_internal: i32, min: i32, max: i32) -> i32 {
    let ui_min = min + TTS_UI_OFFSET;
    let ui_max = max + TTS_UI_OFFSET;
    let ui_fallback = tts_ui_value_from_internal(fallback_internal).clamp(ui_min, ui_max);
    let ui_value = read_tts_edit_value(edit, ui_fallback, ui_min, ui_max);
    (ui_value - TTS_UI_OFFSET).clamp(min, max)
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
        SendMessageW(hwnd, CB_SETCURSEL, WPARAM(best_idx as usize), LPARAM(0));
    }
}

fn update_tts_manual_visibility(hwnd: HWND) {
    let (
        checkbox,
        combo_speed,
        combo_pitch,
        combo_volume,
        edit_speed,
        edit_pitch,
        edit_volume,
        combo_dialogue_voice_rate,
        edit_dialogue_voice_rate,
        combo_dialogue_voice_pitch,
        edit_dialogue_voice_pitch,
        combo_dialogue_voice_volume,
        edit_dialogue_voice_volume,
        combo_dialogue_secondary_voice_rate,
        edit_dialogue_secondary_voice_rate,
        combo_dialogue_secondary_voice_pitch,
        edit_dialogue_secondary_voice_pitch,
        combo_dialogue_secondary_voice_volume,
        edit_dialogue_secondary_voice_volume,
    ) = match with_options_state(hwnd, |state| {
        (
            state.checkbox_tts_manual,
            state.combo_tts_speed,
            state.combo_tts_pitch,
            state.combo_tts_volume,
            state.edit_tts_speed,
            state.edit_tts_pitch,
            state.edit_tts_volume,
            state.combo_dialogue_voice_rate,
            state.edit_dialogue_voice_rate,
            state.combo_dialogue_voice_pitch,
            state.edit_dialogue_voice_pitch,
            state.combo_dialogue_voice_volume,
            state.edit_dialogue_voice_volume,
            state.combo_dialogue_secondary_voice_rate,
            state.edit_dialogue_secondary_voice_rate,
            state.combo_dialogue_secondary_voice_pitch,
            state.edit_dialogue_secondary_voice_pitch,
            state.combo_dialogue_secondary_voice_volume,
            state.edit_dialogue_secondary_voice_volume,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    let manual = crate::send_message_w_safe(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
        == BST_CHECKED.0;
    if manual {
        let rate = combo_value(combo_speed);
        let pitch = combo_value(combo_pitch);
        let volume = combo_value(combo_volume);
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_speed,
            PCWSTR(to_wide(&tts_ui_value_from_internal(rate).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_pitch,
            PCWSTR(to_wide(&tts_ui_value_from_internal(pitch).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_volume,
            PCWSTR(to_wide(&volume.to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        let d_rate = combo_value(combo_dialogue_voice_rate);
        let d_pitch = combo_value(combo_dialogue_voice_pitch);
        let d_volume = combo_value(combo_dialogue_voice_volume);
        let sd_rate = combo_value(combo_dialogue_secondary_voice_rate);
        let sd_pitch = combo_value(combo_dialogue_secondary_voice_pitch);
        let sd_volume = combo_value(combo_dialogue_secondary_voice_volume);
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_dialogue_voice_rate,
            PCWSTR(to_wide(&tts_ui_value_from_internal(d_rate).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_dialogue_voice_pitch,
            PCWSTR(to_wide(&tts_ui_value_from_internal(d_pitch).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_dialogue_voice_volume,
            PCWSTR(to_wide(&d_volume.to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_dialogue_secondary_voice_rate,
            PCWSTR(to_wide(&tts_ui_value_from_internal(sd_rate).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_dialogue_secondary_voice_pitch,
            PCWSTR(to_wide(&tts_ui_value_from_internal(sd_pitch).to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = crate::set_window_text_w_safe(
            edit_dialogue_secondary_voice_volume,
            PCWSTR(to_wide(&sd_volume.to_string()).as_ptr()),
        ) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
    } else {
        let rate = read_tts_tuning_edit_value(edit_speed, 0, TTS_RATE_MIN, TTS_RATE_MAX);
        let pitch = read_tts_tuning_edit_value(edit_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX);
        let volume = read_tts_edit_value(edit_volume, 100, TTS_VOLUME_MIN, TTS_VOLUME_MAX);
        let d_rate =
            read_tts_tuning_edit_value(edit_dialogue_voice_rate, 0, TTS_RATE_MIN, TTS_RATE_MAX);
        let d_pitch =
            read_tts_tuning_edit_value(edit_dialogue_voice_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX);
        let d_volume = read_tts_edit_value(
            edit_dialogue_voice_volume,
            100,
            TTS_VOLUME_MIN,
            TTS_VOLUME_MAX,
        );
        let sd_rate = read_tts_tuning_edit_value(
            edit_dialogue_secondary_voice_rate,
            0,
            TTS_RATE_MIN,
            TTS_RATE_MAX,
        );
        let sd_pitch = read_tts_tuning_edit_value(
            edit_dialogue_secondary_voice_pitch,
            0,
            TTS_PITCH_MIN,
            TTS_PITCH_MAX,
        );
        let sd_volume = read_tts_edit_value(
            edit_dialogue_secondary_voice_volume,
            100,
            TTS_VOLUME_MIN,
            TTS_VOLUME_MAX,
        );
        select_combo_nearest_value(combo_speed, rate);
        select_combo_nearest_value(combo_pitch, pitch);
        select_combo_nearest_value(combo_volume, volume);
        select_combo_nearest_value(combo_dialogue_voice_rate, d_rate);
        select_combo_nearest_value(combo_dialogue_voice_pitch, d_pitch);
        select_combo_nearest_value(combo_dialogue_voice_volume, d_volume);
        select_combo_nearest_value(combo_dialogue_secondary_voice_rate, sd_rate);
        select_combo_nearest_value(combo_dialogue_secondary_voice_pitch, sd_pitch);
        select_combo_nearest_value(combo_dialogue_secondary_voice_volume, sd_volume);
    }
    unsafe {
        ShowWindow(combo_speed, if manual { SW_HIDE } else { SW_SHOW });
        ShowWindow(combo_pitch, if manual { SW_HIDE } else { SW_SHOW });
        ShowWindow(combo_volume, if manual { SW_HIDE } else { SW_SHOW });
        ShowWindow(edit_speed, if manual { SW_SHOW } else { SW_HIDE });
        ShowWindow(edit_pitch, if manual { SW_SHOW } else { SW_HIDE });
        ShowWindow(edit_volume, if manual { SW_SHOW } else { SW_HIDE });
        ShowWindow(
            combo_dialogue_voice_rate,
            if manual { SW_HIDE } else { SW_SHOW },
        );
        ShowWindow(
            combo_dialogue_voice_pitch,
            if manual { SW_HIDE } else { SW_SHOW },
        );
        ShowWindow(
            combo_dialogue_voice_volume,
            if manual { SW_HIDE } else { SW_SHOW },
        );
        ShowWindow(
            edit_dialogue_voice_rate,
            if manual { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            edit_dialogue_voice_pitch,
            if manual { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            edit_dialogue_voice_volume,
            if manual { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            combo_dialogue_secondary_voice_rate,
            if manual { SW_HIDE } else { SW_SHOW },
        );
        ShowWindow(
            combo_dialogue_secondary_voice_pitch,
            if manual { SW_HIDE } else { SW_SHOW },
        );
        ShowWindow(
            combo_dialogue_secondary_voice_volume,
            if manual { SW_HIDE } else { SW_SHOW },
        );
        ShowWindow(
            edit_dialogue_secondary_voice_rate,
            if manual { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            edit_dialogue_secondary_voice_pitch,
            if manual { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            edit_dialogue_secondary_voice_volume,
            if manual { SW_SHOW } else { SW_HIDE },
        );
        EnableWindow(combo_speed, !manual);
        EnableWindow(combo_pitch, !manual);
        EnableWindow(combo_volume, !manual);
        EnableWindow(edit_speed, manual);
        EnableWindow(edit_pitch, manual);
        EnableWindow(edit_volume, manual);
        EnableWindow(combo_dialogue_voice_rate, !manual);
        EnableWindow(combo_dialogue_voice_pitch, !manual);
        EnableWindow(combo_dialogue_voice_volume, !manual);
        EnableWindow(edit_dialogue_voice_rate, manual);
        EnableWindow(edit_dialogue_voice_pitch, manual);
        EnableWindow(edit_dialogue_voice_volume, manual);
        EnableWindow(combo_dialogue_secondary_voice_rate, !manual);
        EnableWindow(combo_dialogue_secondary_voice_pitch, !manual);
        EnableWindow(combo_dialogue_secondary_voice_volume, !manual);
        EnableWindow(edit_dialogue_secondary_voice_rate, manual);
        EnableWindow(edit_dialogue_secondary_voice_pitch, manual);
        EnableWindow(edit_dialogue_secondary_voice_volume, manual);
    }
}

fn update_dialogue_voice_visibility(hwnd: HWND) {
    let (
        checkbox,
        label,
        combo,
        button,
        secondary_button,
        label_engine,
        combo_engine,
        label_voice_language,
        combo_voice_language,
        label_rate,
        combo_rate,
        edit_rate,
        label_pitch,
        combo_pitch,
        edit_pitch,
        label_volume,
        combo_volume,
        edit_volume,
        checkbox_dialogue_multilingual,
        checkbox_use_secondary_voice,
        label_secondary_engine,
        combo_secondary_engine,
        label_secondary_voice_language,
        combo_secondary_voice_language,
        label_secondary_voice,
        combo_secondary_voice,
        checkbox_dialogue_secondary_multilingual,
        label_secondary_rate,
        combo_secondary_rate,
        edit_secondary_rate,
        label_secondary_pitch,
        combo_secondary_pitch,
        edit_secondary_pitch,
        label_secondary_volume,
        combo_secondary_volume,
        edit_secondary_volume,
        label_open_quote,
        edit_open_quote,
        label_close_quote,
        edit_close_quote,
        checkbox_multiline,
    ) = match with_options_state(hwnd, |state| {
        (
            state.checkbox_use_dialogue_voice,
            state.label_dialogue_voice,
            state.combo_dialogue_voice,
            state.button_dialogue_voice_preview,
            state.button_dialogue_secondary_voice_preview,
            state.label_dialogue_engine,
            state.combo_dialogue_engine,
            state.label_dialogue_voice_language,
            state.combo_dialogue_voice_language,
            state.label_dialogue_voice_rate,
            state.combo_dialogue_voice_rate,
            state.edit_dialogue_voice_rate,
            state.label_dialogue_voice_pitch,
            state.combo_dialogue_voice_pitch,
            state.edit_dialogue_voice_pitch,
            state.label_dialogue_voice_volume,
            state.combo_dialogue_voice_volume,
            state.edit_dialogue_voice_volume,
            state.checkbox_dialogue_multilingual,
            state.checkbox_dialogue_use_secondary_voice,
            state.label_dialogue_secondary_engine,
            state.combo_dialogue_secondary_engine,
            state.label_dialogue_secondary_voice_language,
            state.combo_dialogue_secondary_voice_language,
            state.label_dialogue_secondary_voice,
            state.combo_dialogue_secondary_voice,
            state.checkbox_dialogue_secondary_multilingual,
            state.label_dialogue_secondary_voice_rate,
            state.combo_dialogue_secondary_voice_rate,
            state.edit_dialogue_secondary_voice_rate,
            state.label_dialogue_secondary_voice_pitch,
            state.combo_dialogue_secondary_voice_pitch,
            state.edit_dialogue_secondary_voice_pitch,
            state.label_dialogue_secondary_voice_volume,
            state.combo_dialogue_secondary_voice_volume,
            state.edit_dialogue_secondary_voice_volume,
            state.label_dialogue_open_quote,
            state.edit_dialogue_open_quote,
            state.label_dialogue_close_quote,
            state.edit_dialogue_close_quote,
            state.checkbox_dialogue_allow_multiline,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    let controls = [
        checkbox,
        label,
        combo,
        button,
        label_engine,
        combo_engine,
        label_voice_language,
        combo_voice_language,
        label_rate,
        combo_rate,
        edit_rate,
        label_pitch,
        combo_pitch,
        edit_pitch,
        label_volume,
        combo_volume,
        edit_volume,
        checkbox_dialogue_multilingual,
        checkbox_use_secondary_voice,
        label_secondary_engine,
        combo_secondary_engine,
        label_secondary_voice_language,
        combo_secondary_voice_language,
        label_secondary_voice,
        combo_secondary_voice,
        checkbox_dialogue_secondary_multilingual,
        secondary_button,
        label_secondary_rate,
        combo_secondary_rate,
        edit_secondary_rate,
        label_secondary_pitch,
        combo_secondary_pitch,
        edit_secondary_pitch,
        label_secondary_volume,
        combo_secondary_volume,
        edit_secondary_volume,
        label_open_quote,
        edit_open_quote,
        label_close_quote,
        edit_close_quote,
        checkbox_multiline,
    ];
    let voice_tab_active = with_options_state(hwnd, |state| {
        state.hwnd_tabs.0 != 0
            && crate::send_message_w_safe(state.hwnd_tabs, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0
                as i32
                == OPTIONS_TAB_VOICE
    })
    .unwrap_or(false);
    if !voice_tab_active {
        for control in controls {
            unsafe {
                ShowWindow(control, SW_HIDE);
                EnableWindow(control, false);
            }
        }
        return;
    }
    let enabled = crate::send_message_w_safe(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
        == BST_CHECKED.0;
    let secondary_enabled = enabled
        && (crate::send_message_w_safe(
            checkbox_use_secondary_voice,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0);
    let secondary_engine_sel =
        crate::send_message_w_safe(combo_secondary_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let dialogue_engine_sel =
        crate::send_message_w_safe(combo_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let dialogue_engine_is_edge = dialogue_engine_sel <= 0;
    let secondary_engine_is_edge = secondary_engine_sel <= 0;
    let dialogue_only_multilingual = crate::send_message_w_safe(
        checkbox_dialogue_multilingual,
        BM_GETCHECK,
        WPARAM(0),
        LPARAM(0),
    )
    .0 as u32
        == BST_CHECKED.0;
    let dialogue_secondary_only_multilingual = crate::send_message_w_safe(
        checkbox_dialogue_secondary_multilingual,
        BM_GETCHECK,
        WPARAM(0),
        LPARAM(0),
    )
    .0 as u32
        == BST_CHECKED.0;
    let show_dialogue_lang_combo =
        enabled && dialogue_engine_is_edge && !dialogue_only_multilingual;
    let show_secondary_lang_combo =
        secondary_enabled && secondary_engine_is_edge && !dialogue_secondary_only_multilingual;
    let manual_tuning =
        with_options_state(hwnd, |state| state.checkbox_tts_manual).is_some_and(|h| {
            (crate::send_message_w_safe(h, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32)
                == BST_CHECKED.0
        });
    for control in controls {
        let is_toggle = control == checkbox;
        let is_secondary_toggle = control == checkbox_use_secondary_voice;
        let is_primary_combo_tuning =
            control == combo_rate || control == combo_pitch || control == combo_volume;
        let is_primary_edit_tuning =
            control == edit_rate || control == edit_pitch || control == edit_volume;
        let is_secondary_combo_tuning = control == combo_secondary_rate
            || control == combo_secondary_pitch
            || control == combo_secondary_volume;
        let is_secondary_edit_tuning = control == edit_secondary_rate
            || control == edit_secondary_pitch
            || control == edit_secondary_volume;
        let is_secondary_control = control == label_secondary_voice
            || control == label_secondary_engine
            || control == combo_secondary_engine
            || control == label_secondary_voice_language
            || control == combo_secondary_voice_language
            || control == combo_secondary_voice
            || control == checkbox_dialogue_secondary_multilingual
            || control == secondary_button
            || control == label_secondary_rate
            || control == combo_secondary_rate
            || control == edit_secondary_rate
            || control == label_secondary_pitch
            || control == combo_secondary_pitch
            || control == edit_secondary_pitch
            || control == label_secondary_volume
            || control == combo_secondary_volume
            || control == edit_secondary_volume;
        let visible = if is_toggle || enabled {
            SW_SHOW
        } else {
            SW_HIDE
        };
        let is_secondary_lang_control =
            control == label_secondary_voice_language || control == combo_secondary_voice_language;
        let is_primary_lang_control =
            control == label_voice_language || control == combo_voice_language;
        let is_primary_multilingual_toggle = control == checkbox_dialogue_multilingual;
        let is_secondary_multilingual_toggle = control == checkbox_dialogue_secondary_multilingual;
        let visible = if is_secondary_control {
            if secondary_enabled { SW_SHOW } else { SW_HIDE }
        } else {
            visible
        };
        let visible = if is_primary_multilingual_toggle {
            if enabled && dialogue_engine_is_edge {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_secondary_multilingual_toggle {
            if secondary_enabled && secondary_engine_is_edge {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_primary_lang_control {
            if show_dialogue_lang_combo {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_secondary_lang_control {
            if show_secondary_lang_combo {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_primary_combo_tuning {
            if enabled && !manual_tuning {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_primary_edit_tuning {
            if enabled && manual_tuning {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_secondary_combo_tuning {
            if secondary_enabled && !manual_tuning {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        let visible = if is_secondary_edit_tuning {
            if secondary_enabled && manual_tuning {
                SW_SHOW
            } else {
                SW_HIDE
            }
        } else {
            visible
        };
        crate::show_window_safe(control, visible);
        if is_primary_lang_control {
            crate::enable_window_safe(control, show_dialogue_lang_combo);
        } else if is_secondary_lang_control {
            crate::enable_window_safe(control, show_secondary_lang_combo);
        } else if is_primary_combo_tuning {
            crate::enable_window_safe(control, enabled && !manual_tuning);
        } else if is_primary_edit_tuning {
            crate::enable_window_safe(control, enabled && manual_tuning);
        } else if is_secondary_combo_tuning {
            crate::enable_window_safe(control, secondary_enabled && !manual_tuning);
        } else if is_secondary_edit_tuning {
            crate::enable_window_safe(control, secondary_enabled && manual_tuning);
        } else if is_primary_multilingual_toggle {
            crate::enable_window_safe(control, enabled && dialogue_engine_is_edge);
        } else if is_secondary_multilingual_toggle {
            crate::enable_window_safe(control, secondary_enabled && secondary_engine_is_edge);
        } else if is_secondary_control {
            crate::enable_window_safe(control, secondary_enabled);
        } else {
            crate::enable_window_safe(control, is_toggle || is_secondary_toggle || enabled);
        }
    }
}

fn open_podcastindex_signup() {
    unsafe {
        let url = to_wide("https://api.podcastindex.org/signup");
        ShellExecuteW(
            HWND(0),
            w!("open"),
            PCWSTR(url.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        );
    }
}

fn preview_voice(hwnd: HWND) {
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

    let (language, split_on_newline, dictionary) = {
        with_state(parent, |state| {
            (
                state.settings.language,
                state.settings.split_on_newline,
                state.settings.dictionary.clone(),
            )
        })
    }
    .unwrap_or((Language::Italian, true, Vec::new()));

    let text = i18n::tr(language, "tts.preview_text");
    if text.trim().is_empty() {
        return;
    }

    let engine_sel =
        crate::send_message_w_safe(combo_tts_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let engine = match engine_sel {
        1 => TtsEngine::Sapi5,
        2 => TtsEngine::Sapi4,
        _ => TtsEngine::Edge,
    };
    let voices = {
        with_state(parent, |state| match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
    }
    .unwrap_or_default();

    let voice_sel = crate::send_message_w_safe(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel < 0 {
        return;
    }
    let voice_index = crate::send_message_w_safe(
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

    let manual = crate::send_message_w_safe(checkbox_tts_manual, BM_GETCHECK, WPARAM(0), LPARAM(0))
        .0 as u32
        == BST_CHECKED.0;
    let rate = if manual {
        read_tts_tuning_edit_value(edit_tts_speed, 0, TTS_RATE_MIN, TTS_RATE_MAX)
    } else {
        combo_value(combo_tts_speed)
    };
    let pitch = if manual {
        read_tts_tuning_edit_value(edit_tts_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX)
    } else {
        combo_value(combo_tts_pitch)
    };
    let volume = if manual {
        read_tts_edit_value(edit_tts_volume, 100, TTS_VOLUME_MIN, TTS_VOLUME_MAX)
    } else {
        combo_value(combo_tts_volume)
    };
    let chunks = tts_engine::split_into_tts_chunks(&text, split_on_newline, &dictionary, engine);

    match engine {
        TtsEngine::Edge => {
            let options = tts_engine::TtsPlaybackOptions {
                hwnd: parent,
                engine,
                cleaned: text,
                voice,
                chunks,
                initial_caret_pos: 0,
                rate,
                pitch,
                volume,
            };
            tts_engine::start_tts_playback_with_chunks(options);
        }
        TtsEngine::Sapi4 => {
            tts_engine::stop_tts_playback(parent);
            let voice_idx = if let Some(hash_pos) = voice.find('#') {
                let rest = &voice[hash_pos + 1..];
                if let Some(pipe_pos) = rest.find('|') {
                    rest[..pipe_pos].parse::<i32>().unwrap_or(1)
                } else {
                    rest.parse::<i32>().unwrap_or(1)
                }
            } else {
                1
            };
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(parent, |state| {
                    state.tts_session = Some(tts_engine::TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos: 0,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            crate::sapi4_engine::play_sapi4(
                voice_idx, text, rate, pitch, volume, cancel, command_rx,
            );
        }
        TtsEngine::Sapi5 => {
            tts_engine::stop_tts_playback(parent);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(parent, |state| {
                    state.tts_session = Some(tts_engine::TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos: 0,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            if let Err(e) = crate::sapi5_engine::play_sapi(
                vec![text],
                voice,
                rate,
                pitch,
                volume,
                cancel,
                command_rx,
            ) {
                crate::log_debug(&format!("SAPI5 test playback failed: {}", e));
            }
        }
    }
}

fn insert_voice_tag_from_options(hwnd: HWND) {
    let (
        parent,
        combo_tts_engine,
        combo_voice,
        checkbox_manual,
        combo_speed,
        combo_pitch,
        combo_volume,
        edit_speed,
        edit_pitch,
        edit_volume,
    ) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_tts_engine,
            state.combo_voice,
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

    let engine_sel =
        crate::send_message_w_safe(combo_tts_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let engine = match engine_sel {
        1 => TtsEngine::Sapi5,
        2 => TtsEngine::Sapi4,
        _ => TtsEngine::Edge,
    };
    let voices = {
        with_state(parent, |state| match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
    }
    .unwrap_or_default();

    let voice_sel = crate::send_message_w_safe(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel < 0 {
        return;
    }
    let voice_index = crate::send_message_w_safe(
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
    if voice.trim().is_empty() {
        return;
    }

    let manual = crate::send_message_w_safe(checkbox_manual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0
        as u32
        == BST_CHECKED.0;
    let rate = if manual {
        read_tts_tuning_edit_value(edit_speed, 0, TTS_RATE_MIN, TTS_RATE_MAX)
    } else {
        combo_value(combo_speed)
    };
    let pitch = if manual {
        read_tts_tuning_edit_value(edit_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX)
    } else {
        combo_value(combo_pitch)
    };
    let volume = if manual {
        read_tts_edit_value(edit_volume, 100, TTS_VOLUME_MIN, TTS_VOLUME_MAX)
    } else {
        combo_value(combo_volume)
    };

    insert_voice_tag_at_caret(parent, engine, &voice, rate, pitch, volume);
}

fn preview_dialogue_voice(hwnd: HWND) {
    let (
        parent,
        combo_dialogue_engine,
        combo_dialogue_voice,
        combo_dialogue_voice_rate,
        combo_dialogue_voice_pitch,
        combo_dialogue_voice_volume,
        checkbox_tts_manual,
        edit_dialogue_voice_rate,
        edit_dialogue_voice_pitch,
        edit_dialogue_voice_volume,
    ) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_dialogue_engine,
            state.combo_dialogue_voice,
            state.combo_dialogue_voice_rate,
            state.combo_dialogue_voice_pitch,
            state.combo_dialogue_voice_volume,
            state.checkbox_tts_manual,
            state.edit_dialogue_voice_rate,
            state.edit_dialogue_voice_pitch,
            state.edit_dialogue_voice_volume,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let language =
        { with_state(parent, |state| state.settings.language) }.unwrap_or(Language::Italian);

    let text = i18n::tr(language, "tts.preview_text");
    if text.trim().is_empty() {
        return;
    }

    let engine_sel =
        crate::send_message_w_safe(combo_dialogue_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let engine = match engine_sel {
        1 => TtsEngine::Sapi5,
        2 => TtsEngine::Sapi4,
        _ => TtsEngine::Edge,
    };
    let voices = {
        with_state(parent, |state| match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
    }
    .unwrap_or_default();

    let voice_sel =
        crate::send_message_w_safe(combo_dialogue_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel < 0 {
        return;
    }
    let voice_index = crate::send_message_w_safe(
        combo_dialogue_voice,
        CB_GETITEMDATA,
        WPARAM(voice_sel as usize),
        LPARAM(0),
    )
    .0 as usize;
    if voice_index >= voices.len() {
        return;
    }
    let voice = voices[voice_index].short_name.clone();

    let manual = crate::send_message_w_safe(checkbox_tts_manual, BM_GETCHECK, WPARAM(0), LPARAM(0))
        .0 as u32
        == BST_CHECKED.0;
    let rate = if manual {
        read_tts_tuning_edit_value(edit_dialogue_voice_rate, 0, TTS_RATE_MIN, TTS_RATE_MAX)
    } else {
        combo_value(combo_dialogue_voice_rate)
    };
    let pitch = if manual {
        read_tts_tuning_edit_value(edit_dialogue_voice_pitch, 0, TTS_PITCH_MIN, TTS_PITCH_MAX)
    } else {
        combo_value(combo_dialogue_voice_pitch)
    };
    let volume = if manual {
        read_tts_edit_value(
            edit_dialogue_voice_volume,
            100,
            TTS_VOLUME_MIN,
            TTS_VOLUME_MAX,
        )
    } else {
        combo_value(combo_dialogue_voice_volume)
    };

    match engine {
        TtsEngine::Edge => {
            let chunks = tts_engine::split_into_tts_chunks(&text, false, &[], engine);
            let options = tts_engine::TtsPlaybackOptions {
                hwnd: parent,
                engine,
                cleaned: text,
                voice,
                chunks,
                initial_caret_pos: 0,
                rate,
                pitch,
                volume,
            };
            tts_engine::start_tts_playback_with_chunks(options);
        }
        TtsEngine::Sapi4 => {
            tts_engine::stop_tts_playback(parent);
            let voice_idx = if let Some(hash_pos) = voice.find('#') {
                let rest = &voice[hash_pos + 1..];
                if let Some(pipe_pos) = rest.find('|') {
                    rest[..pipe_pos].parse::<i32>().unwrap_or(1)
                } else {
                    rest.parse::<i32>().unwrap_or(1)
                }
            } else {
                1
            };
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(parent, |state| {
                    state.tts_session = Some(tts_engine::TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos: 0,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            crate::sapi4_engine::play_sapi4(
                voice_idx, text, rate, pitch, volume, cancel, command_rx,
            );
        }
        TtsEngine::Sapi5 => {
            tts_engine::stop_tts_playback(parent);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(parent, |state| {
                    state.tts_session = Some(tts_engine::TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos: 0,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            if let Err(e) = crate::sapi5_engine::play_sapi(
                vec![text],
                voice,
                rate,
                pitch,
                volume,
                cancel,
                command_rx,
            ) {
                crate::log_debug(&format!("SAPI5 test playback failed: {}", e));
            }
        }
    }
}

fn preview_dialogue_secondary_voice(hwnd: HWND) {
    let (
        parent,
        combo_dialogue_secondary_engine,
        combo_dialogue_secondary_voice,
        combo_dialogue_secondary_voice_rate,
        combo_dialogue_secondary_voice_pitch,
        combo_dialogue_secondary_voice_volume,
        checkbox_tts_manual,
        edit_dialogue_secondary_voice_rate,
        edit_dialogue_secondary_voice_pitch,
        edit_dialogue_secondary_voice_volume,
    ) = match with_options_state(hwnd, |state| {
        (
            state.parent,
            state.combo_dialogue_secondary_engine,
            state.combo_dialogue_secondary_voice,
            state.combo_dialogue_secondary_voice_rate,
            state.combo_dialogue_secondary_voice_pitch,
            state.combo_dialogue_secondary_voice_volume,
            state.checkbox_tts_manual,
            state.edit_dialogue_secondary_voice_rate,
            state.edit_dialogue_secondary_voice_pitch,
            state.edit_dialogue_secondary_voice_volume,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let language =
        { with_state(parent, |state| state.settings.language) }.unwrap_or(Language::Italian);

    let text = i18n::tr(language, "tts.preview_text");
    if text.trim().is_empty() {
        return;
    }

    let engine_sel = crate::send_message_w_safe(
        combo_dialogue_secondary_engine,
        CB_GETCURSEL,
        WPARAM(0),
        LPARAM(0),
    )
    .0;
    let engine = match engine_sel {
        1 => TtsEngine::Sapi5,
        2 => TtsEngine::Sapi4,
        _ => TtsEngine::Edge,
    };
    let voices = {
        with_state(parent, |state| match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
    }
    .unwrap_or_default();

    let voice_sel = crate::send_message_w_safe(
        combo_dialogue_secondary_voice,
        CB_GETCURSEL,
        WPARAM(0),
        LPARAM(0),
    )
    .0;
    if voice_sel < 0 {
        return;
    }
    let voice_index = crate::send_message_w_safe(
        combo_dialogue_secondary_voice,
        CB_GETITEMDATA,
        WPARAM(voice_sel as usize),
        LPARAM(0),
    )
    .0 as usize;
    if voice_index >= voices.len() {
        return;
    }
    let voice = voices[voice_index].short_name.clone();

    let manual = crate::send_message_w_safe(checkbox_tts_manual, BM_GETCHECK, WPARAM(0), LPARAM(0))
        .0 as u32
        == BST_CHECKED.0;
    let rate = if manual {
        read_tts_tuning_edit_value(
            edit_dialogue_secondary_voice_rate,
            0,
            TTS_RATE_MIN,
            TTS_RATE_MAX,
        )
    } else {
        combo_value(combo_dialogue_secondary_voice_rate)
    };
    let pitch = if manual {
        read_tts_tuning_edit_value(
            edit_dialogue_secondary_voice_pitch,
            0,
            TTS_PITCH_MIN,
            TTS_PITCH_MAX,
        )
    } else {
        combo_value(combo_dialogue_secondary_voice_pitch)
    };
    let volume = if manual {
        read_tts_edit_value(
            edit_dialogue_secondary_voice_volume,
            100,
            TTS_VOLUME_MIN,
            TTS_VOLUME_MAX,
        )
    } else {
        combo_value(combo_dialogue_secondary_voice_volume)
    };

    match engine {
        TtsEngine::Edge => {
            let chunks = tts_engine::split_into_tts_chunks(&text, false, &[], engine);
            let options = tts_engine::TtsPlaybackOptions {
                hwnd: parent,
                engine,
                cleaned: text,
                voice,
                chunks,
                initial_caret_pos: 0,
                rate,
                pitch,
                volume,
            };
            tts_engine::start_tts_playback_with_chunks(options);
        }
        TtsEngine::Sapi4 => {
            tts_engine::stop_tts_playback(parent);
            let voice_idx = if let Some(hash_pos) = voice.find('#') {
                let rest = &voice[hash_pos + 1..];
                if let Some(pipe_pos) = rest.find('|') {
                    rest[..pipe_pos].parse::<i32>().unwrap_or(1)
                } else {
                    rest.parse::<i32>().unwrap_or(1)
                }
            } else {
                1
            };
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(parent, |state| {
                    state.tts_session = Some(tts_engine::TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos: 0,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            crate::sapi4_engine::play_sapi4(
                voice_idx, text, rate, pitch, volume, cancel, command_rx,
            );
        }
        TtsEngine::Sapi5 => {
            tts_engine::stop_tts_playback(parent);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(parent, |state| {
                    state.tts_session = Some(tts_engine::TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos: 0,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to access state in options_window");
            }
            if let Err(e) = crate::sapi5_engine::play_sapi(
                vec![text],
                voice,
                rate,
                pitch,
                volume,
                cancel,
                command_rx,
            ) {
                crate::log_debug(&format!("SAPI5 test playback failed: {}", e));
            }
        }
    }
}

fn apply_options_dialog(hwnd: HWND) {
    unsafe {
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
            edit_audiobook_save_folder,
            checkbox_show_media_save_confirmation,
            combo_audio_split,
            combo_audio_split_minutes,
            edit_audio_split_parts_count,
            combo_audio_split_start_number,
            combo_audiobook_part_naming,
            edit_audio_split_text,
            checkbox_audio_split_requires_newline,
            checkbox_audio_split_epub_chapters,
            checkbox_subtitle_ducking,
            edit_podcast_cache_limit,
            checkbox_rss_show_article_preview,
            checkbox_announce_unread_rss_podcast,
            combo_unread_label_position,
            combo_rss_date_display,
            combo_rss_time_display,
            combo_podcast_date_display,
            combo_podcast_time_display,
            edit_podcastindex_key,
            edit_podcastindex_secret,
            combo_whisper_model,
            checkbox_tts_manual,
            checkbox_multilingual,
            checkbox_use_dialogue_voice,
            combo_dialogue_engine,
            combo_dialogue_voice,
            checkbox_dialogue_use_secondary_voice,
            combo_dialogue_secondary_engine,
            combo_dialogue_secondary_voice,
            _combo_dialogue_secondary_voice_language,
            combo_dialogue_secondary_voice_rate,
            edit_dialogue_secondary_voice_rate,
            combo_dialogue_secondary_voice_pitch,
            edit_dialogue_secondary_voice_pitch,
            combo_dialogue_secondary_voice_volume,
            edit_dialogue_secondary_voice_volume,
            combo_dialogue_voice_rate,
            edit_dialogue_voice_rate,
            combo_dialogue_voice_pitch,
            edit_dialogue_voice_pitch,
            combo_dialogue_voice_volume,
            edit_dialogue_voice_volume,
            edit_dialogue_open_quote,
            edit_dialogue_close_quote,
            checkbox_dialogue_allow_multiline,
            checkbox_split_on_newline,
            checkbox_word_wrap,
            checkbox_smart_quotes,
            checkbox_strip_markdown_keep_bullets,
            checkbox_spellcheck,
            combo_spellcheck_language,
            combo_dictionary_translation,
            combo_wikipedia_language,
            edit_wrap_width,
            combo_indentation,
            combo_tab_width,
            combo_space_width,
            edit_quote_prefix,
            edit_interpreter_path,
            _button_interpreter_browse,
            _button_interpreter_search,
            combo_subtitle_mode,
            edit_subtitle_offset,
            checkbox_move_cursor,
            checkbox_check_updates,
            checkbox_check_beta_updates,
            checkbox_send_crash_reports,
            checkbox_use_legacy_name,
            checkbox_context_menu,
            combo_confirm_delete_rss_mode,
            combo_confirm_delete_podcast_mode,
            combo_rss_quick_copy_mode,
            combo_prompt_program,
            edit_network_proxy,
            edit_network_proxy_username,
            edit_network_proxy_password,
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
                state.edit_audiobook_save_folder,
                state.checkbox_show_media_save_confirmation,
                state.combo_audio_split,
                state.combo_audio_split_minutes,
                state.edit_audio_split_parts_count,
                state.combo_audio_split_start_number,
                state.combo_audiobook_part_naming,
                state.edit_audio_split_text,
                state.checkbox_audio_split_requires_newline,
                state.checkbox_audio_split_epub_chapters,
                state.checkbox_subtitle_ducking,
                state.edit_podcast_cache_limit,
                state.checkbox_rss_show_article_preview,
                state.checkbox_announce_unread_rss_podcast,
                state.combo_unread_label_position,
                state.combo_rss_date_display,
                state.combo_rss_time_display,
                state.combo_podcast_date_display,
                state.combo_podcast_time_display,
                state.edit_podcastindex_key,
                state.edit_podcastindex_secret,
                state.combo_whisper_model,
                state.checkbox_tts_manual,
                state.checkbox_multilingual,
                state.checkbox_use_dialogue_voice,
                state.combo_dialogue_engine,
                state.combo_dialogue_voice,
                state.checkbox_dialogue_use_secondary_voice,
                state.combo_dialogue_secondary_engine,
                state.combo_dialogue_secondary_voice,
                state.combo_dialogue_secondary_voice_language,
                state.combo_dialogue_secondary_voice_rate,
                state.edit_dialogue_secondary_voice_rate,
                state.combo_dialogue_secondary_voice_pitch,
                state.edit_dialogue_secondary_voice_pitch,
                state.combo_dialogue_secondary_voice_volume,
                state.edit_dialogue_secondary_voice_volume,
                state.combo_dialogue_voice_rate,
                state.edit_dialogue_voice_rate,
                state.combo_dialogue_voice_pitch,
                state.edit_dialogue_voice_pitch,
                state.combo_dialogue_voice_volume,
                state.edit_dialogue_voice_volume,
                state.edit_dialogue_open_quote,
                state.edit_dialogue_close_quote,
                state.checkbox_dialogue_allow_multiline,
                state.checkbox_split_on_newline,
                state.checkbox_word_wrap,
                state.checkbox_smart_quotes,
                state.checkbox_strip_markdown_keep_bullets,
                state.checkbox_spellcheck,
                state.combo_spellcheck_language,
                state.combo_dictionary_translation,
                state.combo_wikipedia_language,
                state.edit_wrap_width,
                state.combo_indentation,
                state.combo_tab_width,
                state.combo_space_width,
                state.edit_quote_prefix,
                state.edit_interpreter_path,
                state.button_interpreter_browse,
                state.button_interpreter_search,
                state.combo_subtitle_mode,
                state.edit_subtitle_offset,
                state.checkbox_move_cursor,
                state.checkbox_check_updates,
                state.checkbox_check_beta_updates,
                state.checkbox_send_crash_reports,
                state.checkbox_use_legacy_name,
                state.checkbox_context_menu,
                state.combo_confirm_delete_rss_mode,
                state.combo_confirm_delete_podcast_mode,
                state.combo_rss_quick_copy_mode,
                state.combo_prompt_program,
                state.edit_network_proxy,
                state.edit_network_proxy_username,
                state.edit_network_proxy_password,
            )
        }) {
            Some(values) => values,
            None => return,
        };

        let mut settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
        let old_language = settings.language;
        let old_marker_position = settings.modified_marker_position;
        let old_word_wrap = settings.word_wrap;
        let old_indent_mode = settings.indentation_mode;
        let old_tab_width = settings.indent_tab_width;
        let old_space_width = settings.indent_space_width;
        let old_context_menu = settings.context_menu_open_with;
        let old_use_legacy_name = settings.use_legacy_name;
        let old_spellcheck_enabled = settings.spellcheck_enabled;
        let old_spellcheck_mode = settings.spellcheck_language_mode;
        let old_spellcheck_fixed_language = settings.spellcheck_fixed_language.clone();
        let old_shortcuts = settings.shortcuts.clone();
        settings.shortcuts = with_options_state(hwnd, |state| state.shortcut_draft.clone())
            .unwrap_or_else(ShortcutSettings::default);
        let res = with_state(parent, |state| {
            (
                state.settings.tts_engine,
                state.settings.tts_voice.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
                state.settings.use_dialogue_voice,
                state.settings.dialogue_voice.clone(),
                state.settings.dialogue_use_secondary_voice,
                state.settings.dialogue_secondary_voice.clone(),
                state.settings.dialogue_secondary_voice_rate,
                state.settings.dialogue_secondary_voice_pitch,
                state.settings.dialogue_secondary_voice_volume,
                state.settings.dialogue_secondary_tts_engine,
                state.settings.dialogue_voice_rate,
                state.settings.dialogue_voice_pitch,
                state.settings.dialogue_voice_volume,
                state.settings.dialogue_tts_engine,
                state.settings.dialogue_opening_quote.clone(),
                state.settings.dialogue_closing_quote.clone(),
                state.settings.dialogue_allow_multiline,
                state.tts_session.is_some(),
            )
        });
        let (
            old_engine,
            old_voice,
            old_rate,
            old_pitch,
            old_volume,
            old_use_dialogue_voice,
            old_dialogue_voice,
            old_dialogue_use_secondary_voice,
            old_dialogue_secondary_voice,
            old_dialogue_secondary_rate,
            old_dialogue_secondary_pitch,
            old_dialogue_secondary_volume,
            old_dialogue_secondary_engine,
            old_dialogue_rate,
            old_dialogue_pitch,
            old_dialogue_volume,
            old_dialogue_engine,
            old_dialogue_opening_quote,
            old_dialogue_closing_quote,
            old_dialogue_allow_multiline,
            was_tts_active,
        ) = res.unwrap_or((
            settings.tts_engine,
            settings.tts_voice.clone(),
            settings.tts_rate,
            settings.tts_pitch,
            settings.tts_volume,
            settings.use_dialogue_voice,
            settings.dialogue_voice.clone(),
            settings.dialogue_use_secondary_voice,
            settings.dialogue_secondary_voice.clone(),
            settings.dialogue_secondary_voice_rate,
            settings.dialogue_secondary_voice_pitch,
            settings.dialogue_secondary_voice_volume,
            settings.dialogue_secondary_tts_engine,
            settings.dialogue_voice_rate,
            settings.dialogue_voice_pitch,
            settings.dialogue_voice_volume,
            settings.dialogue_tts_engine,
            settings.dialogue_opening_quote.clone(),
            settings.dialogue_closing_quote.clone(),
            settings.dialogue_allow_multiline,
            false,
        ));

        let lang_sel = SendMessageW(combo_lang, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.language = match lang_sel {
            1 => Language::English,
            2 => Language::Spanish,
            3 => Language::Portuguese,
            4 => Language::Swedish,
            5 => Language::Vietnamese,
            6 => Language::Czech,
            7 => Language::Polish,
            8 => Language::French,
            9 => Language::Serbian,
            10 => Language::Ukrainian,
            11 => Language::Lithuanian,
            12 => Language::Russian,
            13 => Language::Chinese,
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
            2 => TtsEngine::Sapi4,

            _ => TtsEngine::Edge,
        };

        let dialogue_engine_sel =
            SendMessageW(combo_dialogue_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.dialogue_tts_engine = match dialogue_engine_sel {
            1 => TtsEngine::Sapi5,
            2 => TtsEngine::Sapi4,
            _ => TtsEngine::Edge,
        };
        let dialogue_secondary_engine_sel = SendMessageW(
            combo_dialogue_secondary_engine,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.dialogue_secondary_tts_engine = match dialogue_secondary_engine_sel {
            1 => TtsEngine::Sapi5,
            2 => TtsEngine::Sapi4,
            _ => TtsEngine::Edge,
        };

        settings.tts_manual_tuning =
            SendMessageW(checkbox_tts_manual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        if settings.tts_manual_tuning {
            settings.tts_rate = read_tts_tuning_edit_value(
                edit_tts_speed,
                settings.tts_rate,
                TTS_RATE_MIN,
                TTS_RATE_MAX,
            );
            settings.tts_pitch = read_tts_tuning_edit_value(
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
        settings.audiobook_split_by_epub_chapter = SendMessageW(
            checkbox_audio_split_epub_chapters,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        settings.subtitle_mix_ducking =
            SendMessageW(checkbox_subtitle_ducking, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        settings.split_on_newline =
            SendMessageW(checkbox_split_on_newline, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        settings.word_wrap = SendMessageW(checkbox_word_wrap, BM_GETCHECK, WPARAM(0), LPARAM(0)).0
            as u32
            == BST_CHECKED.0;
        settings.smart_quotes =
            SendMessageW(checkbox_smart_quotes, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        settings.strip_markdown_keep_bullets = SendMessageW(
            checkbox_strip_markdown_keep_bullets,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
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
            settings.spellcheck_language_mode =
                crate::settings::SpellcheckLanguageMode::FixedLanguage;
            let val = match spellcheck_sel {
                1 => "en-US",
                2 => "en-GB",
                3 => "it-IT",
                4 => "es-ES",
                5 => "pt-BR",
                6 => "fr-FR",
                7 => "de-DE",
                8 => "pl-PL",
                9 => "ru-RU",
                _ => "en-US",
            };
            settings.spellcheck_fixed_language = val.to_string();
        }

        let dict_sel = SendMessageW(
            combo_dictionary_translation,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        let dict_values = [
            "auto", "none", "it", "en", "es", "pt", "sv", "vi", "cs", "pl", "fr", "uk", "lt", "ru",
            "zh",
        ];
        settings.dictionary_translation_language = if dict_sel >= 0 {
            dict_values
                .get(dict_sel as usize)
                .unwrap_or(&"auto")
                .to_string()
        } else {
            "auto".to_string()
        };

        let wiki_sel = SendMessageW(combo_wikipedia_language, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        let wiki_values = [
            "auto", "it", "en", "es", "pt", "sv", "vi", "cs", "pl", "fr", "uk", "lt", "ru", "zh",
        ];
        settings.wikipedia_language = if wiki_sel >= 0 {
            wiki_values
                .get(wiki_sel as usize)
                .unwrap_or(&"auto")
                .to_string()
        } else {
            "auto".to_string()
        };

        let width_len = GetWindowTextLengthW(edit_wrap_width);
        if width_len >= 0 {
            let mut buf = vec![0u16; (width_len + 1) as usize];
            let read = GetWindowTextW(edit_wrap_width, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            if let Ok(parsed) = text.trim().parse::<u32>()
                && parsed > 0
            {
                settings.wrap_width = parsed;
            }
        }

        let indent_sel = SendMessageW(combo_indentation, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.indentation_mode = match indent_sel {
            1 => crate::settings::IndentationMode::Tabs,
            2 => crate::settings::IndentationMode::Spaces,
            _ => crate::settings::IndentationMode::Default,
        };
        let tab_sel = SendMessageW(combo_tab_width, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if tab_sel >= 0 {
            let tab_width = SendMessageW(
                combo_tab_width,
                CB_GETITEMDATA,
                WPARAM(tab_sel as usize),
                LPARAM(0),
            )
            .0 as u32;
            if matches!(tab_width, 2 | 4 | 6 | 8) {
                settings.indent_tab_width = tab_width;
            }
        }
        let space_sel = SendMessageW(combo_space_width, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if space_sel >= 0 {
            let space_width = SendMessageW(
                combo_space_width,
                CB_GETITEMDATA,
                WPARAM(space_sel as usize),
                LPARAM(0),
            )
            .0 as u32;
            if matches!(space_width, 2 | 4 | 6 | 8) {
                settings.indent_space_width = space_width;
            }
        }
        let prefix_len = GetWindowTextLengthW(edit_quote_prefix);
        if prefix_len >= 0 {
            let mut buf = vec![0u16; (prefix_len + 1) as usize];
            let read = GetWindowTextW(edit_quote_prefix, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            settings.quote_prefix = text;
        }
        let interpreter_len = GetWindowTextLengthW(edit_interpreter_path);
        if interpreter_len >= 0 {
            let mut buf = vec![0u16; (interpreter_len + 1) as usize];
            let read = GetWindowTextW(edit_interpreter_path, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            settings.interpreter_path = text;
        }
        let subtitle_sel = SendMessageW(combo_subtitle_mode, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.subtitle_read_mode = match subtitle_sel {
            1 => SubtitleReadMode::Nvda,
            2 => SubtitleReadMode::User,
            3 => SubtitleReadMode::Record,
            _ => SubtitleReadMode::Off,
        };
        let offset_len = GetWindowTextLengthW(edit_subtitle_offset);
        if offset_len >= 0 {
            let mut buf = vec![0u16; (offset_len + 1) as usize];
            let read = GetWindowTextW(edit_subtitle_offset, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            if let Ok(parsed) = text.trim().parse::<i32>() {
                settings.subtitle_offset_ms = parsed;
            }
        }
        settings.move_cursor_during_reading =
            SendMessageW(checkbox_move_cursor, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        settings.check_updates_on_startup =
            SendMessageW(checkbox_check_updates, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        settings.check_beta_updates_on_startup = SendMessageW(
            checkbox_check_beta_updates,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        settings.send_crash_reports = SendMessageW(
            checkbox_send_crash_reports,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        settings.use_legacy_name =
            SendMessageW(checkbox_use_legacy_name, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        settings.context_menu_open_with =
            SendMessageW(checkbox_context_menu, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                == BST_CHECKED.0;
        let rss_confirm_sel = SendMessageW(
            combo_confirm_delete_rss_mode,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.rss_delete_confirm_mode = match rss_confirm_sel {
            0 => RssDeleteConfirmMode::Feed,
            1 => RssDeleteConfirmMode::Article,
            3 => RssDeleteConfirmMode::None,
            _ => RssDeleteConfirmMode::Both,
        };
        let podcast_confirm_sel = SendMessageW(
            combo_confirm_delete_podcast_mode,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.podcast_delete_confirm_mode = match podcast_confirm_sel {
            0 => PodcastDeleteConfirmMode::Podcast,
            1 => PodcastDeleteConfirmMode::Episode,
            3 => PodcastDeleteConfirmMode::None,
            _ => PodcastDeleteConfirmMode::Both,
        };
        let rss_quick_copy_sel = SendMessageW(
            combo_rss_quick_copy_mode,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.rss_quick_copy_mode = match rss_quick_copy_sel {
            1 => crate::settings::RssQuickCopyMode::Url,
            2 => crate::settings::RssQuickCopyMode::Content,
            3 => crate::settings::RssQuickCopyMode::All,
            _ => crate::settings::RssQuickCopyMode::Title,
        };
        settings.rss_show_article_preview = SendMessageW(
            checkbox_rss_show_article_preview,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        settings.confirm_delete_rss_podcast =
            !matches!(settings.rss_delete_confirm_mode, RssDeleteConfirmMode::None)
                || !matches!(
                    settings.podcast_delete_confirm_mode,
                    PodcastDeleteConfirmMode::None
                );
        settings.announce_unread_rss_podcast_items = SendMessageW(
            checkbox_announce_unread_rss_podcast,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        let unread_label_position_sel = SendMessageW(
            combo_unread_label_position,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.rss_podcast_unread_label_position = if unread_label_position_sel == 1 {
            RssPodcastUnreadLabelPosition::After
        } else {
            RssPodcastUnreadLabelPosition::Before
        };
        let rss_date_sel =
            SendMessageW(combo_rss_date_display, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.rss_articles_date_display = if rss_date_sel == 1 {
            ListDateDisplayMode::Never
        } else {
            ListDateDisplayMode::Always
        };
        let rss_time_sel =
            SendMessageW(combo_rss_time_display, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.rss_articles_time_display = match rss_time_sel {
            1 => ListTimeDisplayMode::Never,
            2 => ListTimeDisplayMode::OnlyIfMultipleSameDay,
            _ => ListTimeDisplayMode::Always,
        };
        let podcast_date_sel = SendMessageW(
            combo_podcast_date_display,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.podcast_episodes_date_display = if podcast_date_sel == 1 {
            ListDateDisplayMode::Never
        } else {
            ListDateDisplayMode::Always
        };
        let podcast_time_sel = SendMessageW(
            combo_podcast_time_display,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        settings.podcast_episodes_time_display = match podcast_time_sel {
            1 => ListTimeDisplayMode::Never,
            2 => ListTimeDisplayMode::OnlyIfMultipleSameDay,
            _ => ListTimeDisplayMode::Always,
        };

        let prompt_sel = SendMessageW(combo_prompt_program, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.prompt_program = match prompt_sel {
            1 => "powershell.exe".to_string(),
            2 => "codex".to_string(),
            _ => "cmd.exe".to_string(),
        };
        let proxy_len = GetWindowTextLengthW(edit_network_proxy);
        if proxy_len >= 0 {
            let mut buf = vec![0u16; (proxy_len + 1) as usize];
            let read = GetWindowTextW(edit_network_proxy, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            settings.network_proxy_url = text.trim().to_string();
        }
        let proxy_user_len = GetWindowTextLengthW(edit_network_proxy_username);
        if proxy_user_len >= 0 {
            let mut buf = vec![0u16; (proxy_user_len + 1) as usize];
            let read = GetWindowTextW(edit_network_proxy_username, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            settings.network_proxy_username = text.trim().to_string();
        }
        let proxy_password_len = GetWindowTextLengthW(edit_network_proxy_password);
        if proxy_password_len >= 0 {
            let mut buf = vec![0u16; (proxy_password_len + 1) as usize];
            let read = GetWindowTextW(edit_network_proxy_password, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            settings.network_proxy_password = text.trim().to_string();
        }
        if !settings.network_proxy_url.is_empty()
            && let Err(err) = proxy_is_valid(
                &settings.network_proxy_url,
                &settings.network_proxy_username,
                &settings.network_proxy_password,
            )
        {
            crate::log_debug(&format!("Invalid proxy removed: {}", err));
            let warning = i18n::tr(settings.language, "options.proxy.invalid");
            let title = i18n::tr(settings.language, "options.title");
            MessageBoxW(
                hwnd,
                PCWSTR(to_wide(&warning).as_ptr()),
                PCWSTR(to_wide(&title).as_ptr()),
                MB_OK | MB_ICONWARNING,
            );
            settings.network_proxy_url.clear();
            settings.network_proxy_username.clear();
            settings.network_proxy_password.clear();
        }

        let voices = with_state(parent, |state| match settings.tts_engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
        .unwrap_or_default();
        let dialogue_voices = with_state(parent, |state| match settings.dialogue_tts_engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
            TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
        })
        .unwrap_or_default();
        let dialogue_secondary_voices = with_state(parent, |state| {
            match settings.dialogue_secondary_tts_engine {
                TtsEngine::Edge => state.edge_voices.clone(),
                TtsEngine::Sapi5 => state.sapi_voices.clone(),
                TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
            }
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

        // Dialogue voice settings
        settings.use_dialogue_voice = SendMessageW(
            checkbox_use_dialogue_voice,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;
        settings.dialogue_use_secondary_voice = SendMessageW(
            checkbox_dialogue_use_secondary_voice,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;

        let dialogue_voice_sel =
            SendMessageW(combo_dialogue_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if dialogue_voice_sel >= 0 {
            let voice_index = SendMessageW(
                combo_dialogue_voice,
                CB_GETITEMDATA,
                WPARAM(dialogue_voice_sel as usize),
                LPARAM(0),
            )
            .0 as usize;
            if voice_index < dialogue_voices.len() {
                settings.dialogue_voice = dialogue_voices[voice_index].short_name.clone();
            } else if let Some(short_name) =
                selected_voice_short_name_from_combo_text(combo_dialogue_voice)
            {
                settings.dialogue_voice = short_name;
            }
        }
        let dialogue_secondary_voice_sel = SendMessageW(
            combo_dialogue_secondary_voice,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        if dialogue_secondary_voice_sel >= 0 {
            let voice_index = SendMessageW(
                combo_dialogue_secondary_voice,
                CB_GETITEMDATA,
                WPARAM(dialogue_secondary_voice_sel as usize),
                LPARAM(0),
            )
            .0 as usize;
            if voice_index < dialogue_secondary_voices.len() {
                settings.dialogue_secondary_voice =
                    dialogue_secondary_voices[voice_index].short_name.clone();
            } else if let Some(short_name) =
                selected_voice_short_name_from_combo_text(combo_dialogue_secondary_voice)
            {
                settings.dialogue_secondary_voice = short_name;
            }
        }
        if settings.tts_manual_tuning {
            settings.dialogue_voice_rate = read_tts_tuning_edit_value(
                edit_dialogue_voice_rate,
                settings.dialogue_voice_rate,
                TTS_RATE_MIN,
                TTS_RATE_MAX,
            );
            settings.dialogue_voice_pitch = read_tts_tuning_edit_value(
                edit_dialogue_voice_pitch,
                settings.dialogue_voice_pitch,
                TTS_PITCH_MIN,
                TTS_PITCH_MAX,
            );
            settings.dialogue_voice_volume = read_tts_edit_value(
                edit_dialogue_voice_volume,
                settings.dialogue_voice_volume,
                TTS_VOLUME_MIN,
                TTS_VOLUME_MAX,
            );
            settings.dialogue_secondary_voice_rate = read_tts_tuning_edit_value(
                edit_dialogue_secondary_voice_rate,
                settings.dialogue_secondary_voice_rate,
                TTS_RATE_MIN,
                TTS_RATE_MAX,
            );
            settings.dialogue_secondary_voice_pitch = read_tts_tuning_edit_value(
                edit_dialogue_secondary_voice_pitch,
                settings.dialogue_secondary_voice_pitch,
                TTS_PITCH_MIN,
                TTS_PITCH_MAX,
            );
            settings.dialogue_secondary_voice_volume = read_tts_edit_value(
                edit_dialogue_secondary_voice_volume,
                settings.dialogue_secondary_voice_volume,
                TTS_VOLUME_MIN,
                TTS_VOLUME_MAX,
            );
        } else {
            settings.dialogue_voice_rate = combo_value(combo_dialogue_voice_rate);
            settings.dialogue_voice_pitch = combo_value(combo_dialogue_voice_pitch);
            settings.dialogue_voice_volume = combo_value(combo_dialogue_voice_volume);
            settings.dialogue_secondary_voice_rate =
                combo_value(combo_dialogue_secondary_voice_rate);
            settings.dialogue_secondary_voice_pitch =
                combo_value(combo_dialogue_secondary_voice_pitch);
            settings.dialogue_secondary_voice_volume =
                combo_value(combo_dialogue_secondary_voice_volume);
        }
        let dialogue_open_quote_len = GetWindowTextLengthW(edit_dialogue_open_quote);
        if dialogue_open_quote_len >= 0 {
            let mut buf = vec![0u16; (dialogue_open_quote_len + 1) as usize];
            let read = GetWindowTextW(edit_dialogue_open_quote, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            let trimmed = text.trim();
            settings.dialogue_opening_quote = if trimmed.is_empty() {
                "\"|\u{201C}|\u{00AB}|\u{201E}".to_string()
            } else {
                trimmed.to_string()
            };
        }
        let dialogue_close_quote_len = GetWindowTextLengthW(edit_dialogue_close_quote);
        if dialogue_close_quote_len >= 0 {
            let mut buf = vec![0u16; (dialogue_close_quote_len + 1) as usize];
            let read = GetWindowTextW(edit_dialogue_close_quote, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            let trimmed = text.trim();
            settings.dialogue_closing_quote = if trimmed.is_empty() {
                "\"|\u{201D}|\u{00BB}".to_string()
            } else {
                trimmed.to_string()
            };
        }
        settings.dialogue_allow_multiline = SendMessageW(
            checkbox_dialogue_allow_multiline,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;

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
            let split_mode = SendMessageW(
                combo_audio_split,
                CB_GETITEMDATA,
                WPARAM(split_sel as usize),
                LPARAM(0),
            )
            .0;
            let split_mode = split_mode as u32;
            if split_mode == AUDIOBOOK_SPLIT_BY_TIME {
                settings.audiobook_split_by_time = true;
                settings.audiobook_split_by_text = false;
                settings.audiobook_split = 0;
            } else if split_mode == AUDIOBOOK_SPLIT_BY_TEXT {
                settings.audiobook_split_by_time = false;
                settings.audiobook_split_by_text = true;
                settings.audiobook_split = 0;
            } else if split_mode == AUDIOBOOK_SPLIT_BY_PARTS {
                settings.audiobook_split_by_time = false;
                settings.audiobook_split_by_text = false;
                let parts_len = GetWindowTextLengthW(edit_audio_split_parts_count);
                if parts_len >= 0 {
                    let mut buf = vec![0u16; (parts_len + 1) as usize];
                    let read = GetWindowTextW(edit_audio_split_parts_count, &mut buf);
                    let text = String::from_utf16_lossy(&buf[..read as usize]);
                    let trimmed = text.trim();
                    let parsed = trimmed.parse::<u32>().ok();
                    if !matches!(parsed, Some(1..=100)) {
                        let error_title = i18n::tr(settings.language, "app.error_title");
                        let error_message =
                            i18n::tr(settings.language, "options.error.audio_split_parts_invalid");
                        MessageBoxW(
                            hwnd,
                            PCWSTR(to_wide(&error_message).as_ptr()),
                            PCWSTR(to_wide(&error_title).as_ptr()),
                            MB_OK | MB_ICONWARNING,
                        );
                        SetFocus(edit_audio_split_parts_count);
                        return;
                    }
                    settings.audiobook_split = parsed.unwrap_or(2);
                } else {
                    settings.audiobook_split = settings.audiobook_split.clamp(1, 100);
                }
            } else {
                settings.audiobook_split_by_time = false;
                settings.audiobook_split_by_text = false;
                settings.audiobook_split = 0;
            }
        }

        let minutes_sel = SendMessageW(
            combo_audio_split_minutes,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        if minutes_sel >= 0 {
            let minutes = SendMessageW(
                combo_audio_split_minutes,
                CB_GETITEMDATA,
                WPARAM(minutes_sel as usize),
                LPARAM(0),
            )
            .0 as u32;
            settings.audiobook_split_minutes = minutes.clamp(1, 60);
        }

        let start_sel = SendMessageW(
            combo_audio_split_start_number,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        if start_sel >= 0 {
            let start_number = SendMessageW(
                combo_audio_split_start_number,
                CB_GETITEMDATA,
                WPARAM(start_sel as usize),
                LPARAM(0),
            )
            .0 as u32;
            settings.audiobook_split_start_number = start_number.clamp(1, 99);
        }

        let naming_sel = SendMessageW(
            combo_audiobook_part_naming,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        if naming_sel >= 0 {
            let naming_mode = SendMessageW(
                combo_audiobook_part_naming,
                CB_GETITEMDATA,
                WPARAM(naming_sel as usize),
                LPARAM(0),
            )
            .0 as u32;
            settings.audiobook_part_naming_mode = match naming_mode {
                1 => AudiobookPartNamingMode::NumberOnly,
                2 => AudiobookPartNamingMode::NumberTitle,
                _ => AudiobookPartNamingMode::TitleNumber,
            };
        }

        let text_len = GetWindowTextLengthW(edit_audio_split_text);
        if text_len >= 0 {
            let mut buf = vec![0u16; (text_len + 1) as usize];
            let read = GetWindowTextW(edit_audio_split_text, &mut buf);
            let text = String::from_utf16_lossy(&buf[..read as usize]);
            settings.audiobook_split_text = text;
        }
        persist_default_save_folder_edit(hwnd);
        if let Some((audiobook, media, documents)) = with_options_state(hwnd, |state| {
            (
                state.default_save_folder_audiobook.clone(),
                state.default_save_folder_media.clone(),
                state.default_save_folder_documents.clone(),
            )
        }) {
            settings.audiobook_save_folder = audiobook.trim().to_string();
            settings.media_save_folder = media.trim().to_string();
            settings.documents_save_folder = documents.trim().to_string();
        } else {
            let audiobook_folder_len = GetWindowTextLengthW(edit_audiobook_save_folder);
            if audiobook_folder_len >= 0 {
                let mut buf = vec![0u16; (audiobook_folder_len + 1) as usize];
                let read = GetWindowTextW(edit_audiobook_save_folder, &mut buf);
                let text = String::from_utf16_lossy(&buf[..read as usize]);
                settings.audiobook_save_folder = text.trim().to_string();
            }
        }
        settings.show_media_save_confirmation = SendMessageW(
            checkbox_show_media_save_confirmation,
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 as u32
            == BST_CHECKED.0;

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
        let whisper_sel = SendMessageW(combo_whisper_model, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        settings.whisper_model_profile = match whisper_sel {
            1 => "medium_q5_0".to_string(),
            2 => "large_v3_turbo_q5_0".to_string(),
            _ => "small_q5_1".to_string(),
        };
        if let Some(checkbox_whisper_cuda) =
            with_options_state(hwnd, |state| state.checkbox_whisper_cuda)
        {
            settings.whisper_cuda_enabled =
                SendMessageW(checkbox_whisper_cuda, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
                    == BST_CHECKED.0;
        }
        if let Some(checkbox_whisper_keep_original_language) =
            with_options_state(hwnd, |state| state.checkbox_whisper_keep_original_language)
        {
            settings.whisper_keep_original_language = SendMessageW(
                checkbox_whisper_keep_original_language,
                BM_GETCHECK,
                WPARAM(0),
                LPARAM(0),
            )
            .0 as u32
                == BST_CHECKED.0;
        }
        if let Some(checkbox_whisper_include_timestamps) =
            with_options_state(hwnd, |state| state.checkbox_whisper_include_timestamps)
        {
            settings.whisper_include_timestamps = SendMessageW(
                checkbox_whisper_include_timestamps,
                BM_GETCHECK,
                WPARAM(0),
                LPARAM(0),
            )
            .0 as u32
                == BST_CHECKED.0;
        }
        if let Some((combo_dictation_microphone, device_ids)) = with_options_state(hwnd, |state| {
            (
                state.combo_dictation_microphone,
                state.dictation_microphone_device_ids.clone(),
            )
        }) {
            let sel = SendMessageW(
                combo_dictation_microphone,
                CB_GETCURSEL,
                WPARAM(0),
                LPARAM(0),
            )
            .0;
            settings.dictation_microphone_device_id = device_ids
                .get(sel.max(0) as usize)
                .cloned()
                .unwrap_or_else(|| crate::settings::PODCAST_DEVICE_DEFAULT.to_string());
        }

        let dialogue_cfg = crate::dialogue_voice::DialogueVoiceConfig {
            engine: settings.dialogue_tts_engine,
            voice: settings.dialogue_voice.clone(),
            use_secondary_voice: settings.dialogue_use_secondary_voice,
            secondary_voice: settings.dialogue_secondary_voice.clone(),
            secondary_engine: settings.dialogue_secondary_tts_engine,
            secondary_rate: settings.dialogue_secondary_voice_rate,
            secondary_pitch: settings.dialogue_secondary_voice_pitch,
            secondary_volume: settings.dialogue_secondary_voice_volume,
            rate: settings.dialogue_voice_rate,
            pitch: settings.dialogue_voice_pitch,
            volume: settings.dialogue_voice_volume,
            opening_quote: settings.dialogue_opening_quote.clone(),
            closing_quote: settings.dialogue_closing_quote.clone(),
            allow_multiline: settings.dialogue_allow_multiline,
        };
        if let Err(err) = crate::dialogue_voice::save_dialogue_voice_config(&dialogue_cfg) {
            crate::log_debug(&format!("Failed to save dialogue config: {}", err));
        }

        let (mut dialog_profiles, mut active_profile_name) = with_options_state(hwnd, |state| {
            (
                state.voice_profiles.clone(),
                state.active_voice_profile_name.clone(),
            )
        })
        .unwrap_or((Vec::new(), DEFAULT_VOICE_PROFILE_NAME.to_string()));

        if active_profile_name.trim().is_empty() {
            active_profile_name = DEFAULT_VOICE_PROFILE_NAME.to_string();
        }
        if !dialog_profiles
            .iter()
            .any(|p| p.name.eq_ignore_ascii_case(DEFAULT_VOICE_PROFILE_NAME))
        {
            dialog_profiles.push(voice_profile_from_settings_fields(
                DEFAULT_VOICE_PROFILE_NAME.to_string(),
                &settings,
            ));
        }

        let current_profile =
            voice_profile_from_settings_fields(active_profile_name.clone(), &settings);
        if let Some(existing) = dialog_profiles
            .iter_mut()
            .find(|p| p.name.eq_ignore_ascii_case(&active_profile_name))
        {
            *existing = current_profile;
        } else {
            dialog_profiles.push(current_profile);
        }

        settings.active_voice_profile = active_profile_name;
        settings.voice_profiles = dialog_profiles;

        if with_state(parent, |state| {
            state.settings = settings.clone();
        })
        .is_none()
        {
            crate::log_debug("Failed to access state in options_window");
        }
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
        if settings.use_legacy_name != old_use_legacy_name {
            sync_start_menu_shortcuts(&settings);
            update_window_title(parent);
        }

        if old_language != new_language || old_shortcuts != settings.shortcuts {
            rebuild_menus(parent);
        }
        if old_marker_position != settings.modified_marker_position {
            update_window_title(parent);
        }
        if old_word_wrap != settings.word_wrap {
            apply_word_wrap_to_all_edits(parent, settings.word_wrap);
        }
        if old_indent_mode != settings.indentation_mode
            || old_tab_width != settings.indent_tab_width
            || old_space_width != settings.indent_space_width
        {
            apply_indent_settings_to_all_edits(parent, &settings);
        }
        refresh_voice_panel(parent);
        if was_tts_active
            && (old_engine != settings.tts_engine
                || old_voice != settings.tts_voice
                || old_rate != settings.tts_rate
                || old_pitch != settings.tts_pitch
                || old_volume != settings.tts_volume
                || old_use_dialogue_voice != settings.use_dialogue_voice
                || old_dialogue_voice != settings.dialogue_voice
                || old_dialogue_use_secondary_voice != settings.dialogue_use_secondary_voice
                || old_dialogue_secondary_voice != settings.dialogue_secondary_voice
                || old_dialogue_secondary_rate != settings.dialogue_secondary_voice_rate
                || old_dialogue_secondary_pitch != settings.dialogue_secondary_voice_pitch
                || old_dialogue_secondary_volume != settings.dialogue_secondary_voice_volume
                || old_dialogue_secondary_engine != settings.dialogue_secondary_tts_engine
                || old_dialogue_rate != settings.dialogue_voice_rate
                || old_dialogue_pitch != settings.dialogue_voice_pitch
                || old_dialogue_volume != settings.dialogue_voice_volume
                || old_dialogue_engine != settings.dialogue_tts_engine
                || old_dialogue_opening_quote != settings.dialogue_opening_quote
                || old_dialogue_closing_quote != settings.dialogue_closing_quote
                || old_dialogue_allow_multiline != settings.dialogue_allow_multiline)
        {
            crate::restart_tts_from_current_offset(parent);
        }
        if parent.0 != 0
            && let Err(_e) = PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0))
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        crate::log_if_err!(crate::destroy_window_safe(hwnd));
    }
}

fn update_audio_split_visibility(hwnd: HWND) {
    let (
        combo_audio_split,
        label_audio_split_text,
        edit_audio_split_text,
        checkbox_audio_split_requires_newline,
        label_audio_split_minutes,
        combo_audio_split_minutes,
        label_audio_split_parts_count,
        edit_audio_split_parts_count,
        label_audio_split_start_number,
        combo_audio_split_start_number,
    ) = match with_options_state(hwnd, |state| {
        (
            state.combo_audio_split,
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
            state.label_audio_split_minutes,
            state.combo_audio_split_minutes,
            state.label_audio_split_parts_count,
            state.edit_audio_split_parts_count,
            state.label_audio_split_start_number,
            state.combo_audio_split_start_number,
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let split_sel =
        crate::send_message_w_safe(combo_audio_split, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let (selected_text, selected_time, selected_parts) = if split_sel >= 0 {
        let split_parts = crate::send_message_w_safe(
            combo_audio_split,
            CB_GETITEMDATA,
            WPARAM(split_sel as usize),
            LPARAM(0),
        )
        .0 as u32;
        (
            split_parts == AUDIOBOOK_SPLIT_BY_TEXT,
            split_parts == AUDIOBOOK_SPLIT_BY_TIME,
            split_parts == AUDIOBOOK_SPLIT_BY_PARTS,
        )
    } else {
        (false, false, false)
    };

    let show_text = if selected_text { SW_SHOW } else { SW_HIDE };
    unsafe {
        ShowWindow(label_audio_split_text, show_text);
        ShowWindow(edit_audio_split_text, show_text);
        ShowWindow(checkbox_audio_split_requires_newline, show_text);
        EnableWindow(edit_audio_split_text, selected_text);
        EnableWindow(checkbox_audio_split_requires_newline, selected_text);
    }

    let show_time = if selected_time { SW_SHOW } else { SW_HIDE };
    unsafe {
        ShowWindow(label_audio_split_minutes, show_time);
        ShowWindow(combo_audio_split_minutes, show_time);
        EnableWindow(combo_audio_split_minutes, selected_time);
    }
    let show_parts = if selected_parts { SW_SHOW } else { SW_HIDE };
    unsafe {
        ShowWindow(label_audio_split_parts_count, show_parts);
        ShowWindow(edit_audio_split_parts_count, show_parts);
        EnableWindow(edit_audio_split_parts_count, selected_parts);
        ShowWindow(label_audio_split_start_number, show_time);
        ShowWindow(combo_audio_split_start_number, show_time);
        EnableWindow(combo_audio_split_start_number, selected_time);
    }
}

fn update_subtitle_ducking_visibility(hwnd: HWND) {
    let (combo_subtitle_mode, checkbox_subtitle_ducking) = match with_options_state(hwnd, |state| {
        (state.combo_subtitle_mode, state.checkbox_subtitle_ducking)
    }) {
        Some(values) => values,
        None => return,
    };
    let sel = crate::send_message_w_safe(combo_subtitle_mode, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let show = sel == 3;
    unsafe {
        ShowWindow(
            checkbox_subtitle_ducking,
            if show { SW_SHOW } else { SW_HIDE },
        );
        EnableWindow(checkbox_subtitle_ducking, show);
    }
}

fn update_spellcheck_language_visibility(hwnd: HWND) {
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
    let spellcheck_enabled =
        crate::send_message_w_safe(checkbox_spellcheck, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
            == BST_CHECKED.0;
    unsafe {
        EnableWindow(label_spellcheck_language, spellcheck_enabled);
        EnableWindow(combo_spellcheck_language, spellcheck_enabled);
    }
}

fn update_indentation_visibility(hwnd: HWND) {
    let (combo_indentation, label_tab_width, combo_tab_width, label_space_width, combo_space_width) =
        match with_options_state(hwnd, |state| {
            (
                state.combo_indentation,
                state.label_tab_width,
                state.combo_tab_width,
                state.label_space_width,
                state.combo_space_width,
            )
        }) {
            Some(values) => values,
            None => return,
        };

    let sel = crate::send_message_w_safe(combo_indentation, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let show_tab = sel == 1;
    let show_space = sel == 2;

    unsafe {
        ShowWindow(label_tab_width, if show_tab { SW_SHOW } else { SW_HIDE });
        ShowWindow(combo_tab_width, if show_tab { SW_SHOW } else { SW_HIDE });
        EnableWindow(combo_tab_width, show_tab);
    }

    unsafe {
        ShowWindow(
            label_space_width,
            if show_space { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            combo_space_width,
            if show_space { SW_SHOW } else { SW_HIDE },
        );
        EnableWindow(combo_space_width, show_space);
    }
}

fn move_control_best_effort(name: &str, hwnd: HWND, x: i32, y: i32, w: i32, h: i32) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        let actual_height = adjusted_control_height(hwnd, h);
        if MoveWindow(hwnd, x, y, w, actual_height, true).is_err() {
            crate::log_debug(&format!("MoveWindow failed for {}", name));
        }
    }
}

fn adjusted_control_height(hwnd: HWND, height: i32) -> i32 {
    if height != OPTIONS_COMBO_HEIGHT || !is_combo_box(hwnd) {
        return height;
    }
    OPTIONS_COMBO_DROPDOWN_HEIGHT
}

fn is_combo_box(hwnd: HWND) -> bool {
    unsafe {
        let mut class_name = [0u16; 32];
        let len = GetClassNameW(hwnd, &mut class_name);
        if len == 0 {
            return false;
        }
        let class_name = String::from_utf16_lossy(&class_name[..len as usize]);
        class_name.eq_ignore_ascii_case("ComboBox")
    }
}

fn layout_dialog_buttons(hwnd: HWND, ok_button: HWND, cancel_button: HWND) {
    let mut rect = RECT::default();
    unsafe {
        if GetClientRect(hwnd, &mut rect).is_err() {
            crate::log_debug("GetClientRect failed for options dialog");
            return;
        }
    }
    let button_y = (rect.bottom - 40).max(OPTIONS_CONTENT_TOP);
    move_control_best_effort("options_ok", ok_button, rect.right - 200, button_y, 90, 28);
    move_control_best_effort(
        "options_cancel",
        cancel_button,
        rect.right - 100,
        button_y,
        90,
        28,
    );
}

fn layout_label_control(
    label_name: &str,
    label: HWND,
    control_name: &str,
    control: HWND,
    y: i32,
    control_height: i32,
) -> i32 {
    move_control_best_effort(
        label_name,
        label,
        OPTIONS_MARGIN_X,
        y,
        OPTIONS_LABEL_WIDTH,
        20,
    );
    move_control_best_effort(
        control_name,
        control,
        OPTIONS_CONTROL_X,
        y - 2,
        OPTIONS_CONTROL_WIDTH,
        control_height,
    );
    y + OPTIONS_ROW_HEIGHT
}

fn layout_checkbox(name: &str, checkbox: HWND, y: i32) -> i32 {
    move_control_best_effort(
        name,
        checkbox,
        OPTIONS_CONTROL_X,
        y,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_CHECKBOX_HEIGHT,
    );
    y + OPTIONS_ROW_HEIGHT
}

fn layout_button(name: &str, button: HWND, y: i32) -> i32 {
    move_control_best_effort(
        name,
        button,
        OPTIONS_CONTROL_X,
        y,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_BUTTON_HEIGHT,
    );
    y + OPTIONS_ROW_HEIGHT
}

fn layout_button_compact(name: &str, button: HWND, y: i32) -> i32 {
    move_control_best_effort(
        name,
        button,
        OPTIONS_CONTROL_X,
        y,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_BUTTON_HEIGHT,
    );
    y + OPTIONS_ROW_HEIGHT_COMPACT
}

fn layout_button_compact_height(name: &str, button: HWND, y: i32, height: i32) -> i32 {
    move_control_best_effort(
        name,
        button,
        OPTIONS_CONTROL_X,
        y,
        OPTIONS_CONTROL_WIDTH,
        height,
    );
    y + OPTIONS_ROW_HEIGHT_COMPACT
}

fn layout_general_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_label_control(
        "label_language",
        state.label_language,
        "combo_lang",
        state.combo_lang,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_modified_marker_position",
        state.label_modified_marker_position,
        "combo_modified_marker_position",
        state.combo_modified_marker_position,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_open",
        state.label_open,
        "combo_open",
        state.combo_open,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_prompt_program",
        state.label_prompt_program,
        "combo_prompt_program",
        state.combo_prompt_program,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_network_proxy",
        state.label_network_proxy,
        "edit_network_proxy",
        state.edit_network_proxy,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_network_proxy_username",
        state.label_network_proxy_username,
        "edit_network_proxy_username",
        state.edit_network_proxy_username,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_network_proxy_password",
        state.label_network_proxy_password,
        "edit_network_proxy_password",
        state.edit_network_proxy_password,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y += OPTIONS_SECTION_GAP;
    y = layout_checkbox("checkbox_check_updates", state.checkbox_check_updates, y);
    y = layout_checkbox(
        "checkbox_check_beta_updates",
        state.checkbox_check_beta_updates,
        y,
    );
    y = layout_checkbox(
        "checkbox_send_crash_reports",
        state.checkbox_send_crash_reports,
        y,
    );
    y = layout_checkbox(
        "checkbox_use_legacy_name",
        state.checkbox_use_legacy_name,
        y,
    );
    y = layout_checkbox("checkbox_context_menu", state.checkbox_context_menu, y);
    y += OPTIONS_SECTION_GAP;
    y = layout_label_control(
        "label_file_associations",
        state.label_file_associations,
        "button_manage_associations",
        state.button_manage_associations,
        y,
        OPTIONS_BUTTON_HEIGHT,
    );
    y + scroll_offset
}

fn layout_voice_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_label_control(
        "label_voice_profile",
        state.label_voice_profile,
        "combo_voice_profile",
        state.combo_voice_profile,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_button_compact(
        "button_rename_voice_profile",
        state.button_rename_voice_profile,
        y,
    );
    move_control_best_effort(
        "button_add_voice_profile",
        state.button_add_voice_profile,
        OPTIONS_CONTROL_X,
        y,
        (OPTIONS_CONTROL_WIDTH - 10) / 2,
        OPTIONS_BUTTON_HEIGHT,
    );
    move_control_best_effort(
        "button_delete_voice_profile",
        state.button_delete_voice_profile,
        OPTIONS_CONTROL_X + (OPTIONS_CONTROL_WIDTH - 10) / 2 + 10,
        y,
        (OPTIONS_CONTROL_WIDTH - 10) / 2,
        OPTIONS_BUTTON_HEIGHT,
    );
    y += OPTIONS_ROW_HEIGHT;
    y = layout_label_control(
        "label_tts_engine",
        state.label_tts_engine,
        "combo_tts_engine",
        state.combo_tts_engine,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_tts_voice_language",
        state.label_tts_voice_language,
        "combo_tts_voice_language",
        state.combo_tts_voice_language,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_voice",
        state.label_voice,
        "combo_voice",
        state.combo_voice,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_checkbox("checkbox_multilingual", state.checkbox_multilingual, y);
    y = layout_label_control(
        "label_tts_speed",
        state.label_tts_speed,
        "combo_tts_speed",
        state.combo_tts_speed,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_tts_speed",
        state.edit_tts_speed,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_tts_pitch",
        state.label_tts_pitch,
        "combo_tts_pitch",
        state.combo_tts_pitch,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_tts_pitch",
        state.edit_tts_pitch,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_tts_volume",
        state.label_tts_volume,
        "combo_tts_volume",
        state.combo_tts_volume,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_tts_volume",
        state.edit_tts_volume,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_checkbox("checkbox_tts_manual", state.checkbox_tts_manual, y);
    y = layout_button_compact("button_tts_preview", state.button_tts_preview, y);
    y = layout_button_compact("button_tts_insert_tag", state.button_tts_insert_tag, y);
    y += OPTIONS_SECTION_GAP;
    y = layout_checkbox(
        "checkbox_split_on_newline",
        state.checkbox_split_on_newline,
        y,
    );
    y = layout_checkbox(
        "checkbox_use_dialogue_voice",
        state.checkbox_use_dialogue_voice,
        y,
    );
    y += OPTIONS_SECTION_GAP;
    y = layout_label_control(
        "label_dialogue_engine",
        state.label_dialogue_engine,
        "combo_dialogue_engine",
        state.combo_dialogue_engine,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_voice",
        state.label_dialogue_voice,
        "combo_dialogue_voice",
        state.combo_dialogue_voice,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_checkbox(
        "checkbox_dialogue_multilingual",
        state.checkbox_dialogue_multilingual,
        y,
    );
    y = layout_button_compact(
        "button_dialogue_voice_preview",
        state.button_dialogue_voice_preview,
        y,
    );
    y = layout_label_control(
        "label_dialogue_voice_rate",
        state.label_dialogue_voice_rate,
        "combo_dialogue_voice_rate",
        state.combo_dialogue_voice_rate,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_dialogue_voice_rate",
        state.edit_dialogue_voice_rate,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_voice_pitch",
        state.label_dialogue_voice_pitch,
        "combo_dialogue_voice_pitch",
        state.combo_dialogue_voice_pitch,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_dialogue_voice_pitch",
        state.edit_dialogue_voice_pitch,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_voice_volume",
        state.label_dialogue_voice_volume,
        "combo_dialogue_voice_volume",
        state.combo_dialogue_voice_volume,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_dialogue_voice_volume",
        state.edit_dialogue_voice_volume,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_checkbox(
        "checkbox_dialogue_use_secondary_voice",
        state.checkbox_dialogue_use_secondary_voice,
        y,
    );
    y = layout_label_control(
        "label_dialogue_secondary_engine",
        state.label_dialogue_secondary_engine,
        "combo_dialogue_secondary_engine",
        state.combo_dialogue_secondary_engine,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_secondary_voice_language",
        state.label_dialogue_secondary_voice_language,
        "combo_dialogue_secondary_voice_language",
        state.combo_dialogue_secondary_voice_language,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_secondary_voice",
        state.label_dialogue_secondary_voice,
        "combo_dialogue_secondary_voice",
        state.combo_dialogue_secondary_voice,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_checkbox(
        "checkbox_dialogue_secondary_multilingual",
        state.checkbox_dialogue_secondary_multilingual,
        y,
    );
    y = layout_button_compact(
        "button_dialogue_secondary_voice_preview",
        state.button_dialogue_secondary_voice_preview,
        y,
    );
    y = layout_label_control(
        "label_dialogue_secondary_voice_rate",
        state.label_dialogue_secondary_voice_rate,
        "combo_dialogue_secondary_voice_rate",
        state.combo_dialogue_secondary_voice_rate,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_dialogue_secondary_voice_rate",
        state.edit_dialogue_secondary_voice_rate,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_secondary_voice_pitch",
        state.label_dialogue_secondary_voice_pitch,
        "combo_dialogue_secondary_voice_pitch",
        state.combo_dialogue_secondary_voice_pitch,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_dialogue_secondary_voice_pitch",
        state.edit_dialogue_secondary_voice_pitch,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_secondary_voice_volume",
        state.label_dialogue_secondary_voice_volume,
        "combo_dialogue_secondary_voice_volume",
        state.combo_dialogue_secondary_voice_volume,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    move_control_best_effort(
        "edit_dialogue_secondary_voice_volume",
        state.edit_dialogue_secondary_voice_volume,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        OPTIONS_CONTROL_WIDTH,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_open_quote",
        state.label_dialogue_open_quote,
        "edit_dialogue_open_quote",
        state.edit_dialogue_open_quote,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_dialogue_close_quote",
        state.label_dialogue_close_quote,
        "edit_dialogue_close_quote",
        state.edit_dialogue_close_quote,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_checkbox(
        "checkbox_dialogue_allow_multiline",
        state.checkbox_dialogue_allow_multiline,
        y,
    );
    y + scroll_offset
}

fn layout_editor_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_checkbox("checkbox_word_wrap", state.checkbox_word_wrap, y);
    y = layout_checkbox("checkbox_smart_quotes", state.checkbox_smart_quotes, y);
    y = layout_checkbox(
        "checkbox_strip_markdown_keep_bullets",
        state.checkbox_strip_markdown_keep_bullets,
        y,
    );
    y = layout_checkbox("checkbox_spellcheck", state.checkbox_spellcheck, y);
    y = layout_label_control(
        "label_spellcheck_language",
        state.label_spellcheck_language,
        "combo_spellcheck_language",
        state.combo_spellcheck_language,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_dictionary_translation",
        state.label_dictionary_translation,
        "combo_dictionary_translation",
        state.combo_dictionary_translation,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_wikipedia_language",
        state.label_wikipedia_language,
        "combo_wikipedia_language",
        state.combo_wikipedia_language,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_wrap_width",
        state.label_wrap_width,
        "edit_wrap_width",
        state.edit_wrap_width,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_indentation",
        state.label_indentation,
        "combo_indentation",
        state.combo_indentation,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_tab_width",
        state.label_tab_width,
        "combo_tab_width",
        state.combo_tab_width,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_space_width",
        state.label_space_width,
        "combo_space_width",
        state.combo_space_width,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_quote_prefix",
        state.label_quote_prefix,
        "edit_quote_prefix",
        state.edit_quote_prefix,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_interpreter_path",
        state.label_interpreter_path,
        "edit_interpreter_path",
        state.edit_interpreter_path,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_button_compact_height(
        "button_interpreter_browse",
        state.button_interpreter_browse,
        y,
        OPTIONS_BUTTON_HEIGHT,
    );
    y = layout_button_compact_height(
        "button_interpreter_search",
        state.button_interpreter_search,
        y,
        OPTIONS_BUTTON_HEIGHT,
    );
    y += OPTIONS_SECTION_GAP;
    y = layout_checkbox("checkbox_move_cursor", state.checkbox_move_cursor, y);
    y + scroll_offset
}

fn layout_audio_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let (show_text_split, show_time_split, show_parts_split) = unsafe {
        let split_sel = SendMessageW(state.combo_audio_split, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if split_sel >= 0 {
            let split_parts = SendMessageW(
                state.combo_audio_split,
                CB_GETITEMDATA,
                WPARAM(split_sel as usize),
                LPARAM(0),
            )
            .0 as u32;
            (
                split_parts == AUDIOBOOK_SPLIT_BY_TEXT,
                split_parts == AUDIOBOOK_SPLIT_BY_TIME,
                split_parts == AUDIOBOOK_SPLIT_BY_PARTS,
            )
        } else {
            (false, false, false)
        }
    };
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_label_control(
        "label_audio_skip",
        state.label_audio_skip,
        "combo_audio_skip",
        state.combo_audio_skip,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_default_save_folder_kind",
        state.label_default_save_folder_kind,
        "combo_default_save_folder_kind",
        state.combo_default_save_folder_kind,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_audiobook_save_folder",
        state.label_audiobook_save_folder,
        "edit_audiobook_save_folder",
        state.edit_audiobook_save_folder,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    let browse_width = 72;
    let spacing = 6;
    let edit_width = (OPTIONS_CONTROL_WIDTH - browse_width - spacing).max(80);
    move_control_best_effort(
        "edit_audiobook_save_folder",
        state.edit_audiobook_save_folder,
        OPTIONS_CONTROL_X,
        y - OPTIONS_ROW_HEIGHT - 2,
        edit_width,
        OPTIONS_EDIT_HEIGHT,
    );
    move_control_best_effort(
        "button_audiobook_save_folder_browse",
        state.button_audiobook_save_folder_browse,
        OPTIONS_CONTROL_X + edit_width + spacing,
        y - OPTIONS_ROW_HEIGHT - 2,
        browse_width,
        OPTIONS_BUTTON_HEIGHT,
    );
    y = layout_checkbox(
        "checkbox_show_media_save_confirmation",
        state.checkbox_show_media_save_confirmation,
        y,
    );
    y = layout_label_control(
        "label_audio_split",
        state.label_audio_split,
        "combo_audio_split",
        state.combo_audio_split,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    if show_time_split {
        y = layout_label_control(
            "label_audio_split_minutes",
            state.label_audio_split_minutes,
            "combo_audio_split_minutes",
            state.combo_audio_split_minutes,
            y,
            OPTIONS_COMBO_HEIGHT,
        );
        y = layout_label_control(
            "label_audio_split_start_number",
            state.label_audio_split_start_number,
            "combo_audio_split_start_number",
            state.combo_audio_split_start_number,
            y,
            OPTIONS_COMBO_HEIGHT,
        );
    }
    y = layout_label_control(
        "label_audiobook_part_naming",
        state.label_audiobook_part_naming,
        "combo_audiobook_part_naming",
        state.combo_audiobook_part_naming,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    if show_parts_split {
        y = layout_label_control(
            "label_audio_split_parts_count",
            state.label_audio_split_parts_count,
            "edit_audio_split_parts_count",
            state.edit_audio_split_parts_count,
            y,
            OPTIONS_EDIT_HEIGHT,
        );
    }
    if show_text_split {
        y = layout_label_control(
            "label_audio_split_text",
            state.label_audio_split_text,
            "edit_audio_split_text",
            state.edit_audio_split_text,
            y,
            OPTIONS_EDIT_HEIGHT,
        );
        y = layout_checkbox(
            "checkbox_audio_split_requires_newline",
            state.checkbox_audio_split_requires_newline,
            y,
        );
    }
    y = layout_checkbox(
        "checkbox_audio_split_epub_chapters",
        state.checkbox_audio_split_epub_chapters,
        y,
    );
    y += OPTIONS_SECTION_GAP;
    y = layout_checkbox(
        "checkbox_subtitle_ducking",
        state.checkbox_subtitle_ducking,
        y,
    );
    y = layout_label_control(
        "label_subtitle_mode",
        state.label_subtitle_mode,
        "combo_subtitle_mode",
        state.combo_subtitle_mode,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_subtitle_offset",
        state.label_subtitle_offset,
        "edit_subtitle_offset",
        state.edit_subtitle_offset,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y + scroll_offset
}

fn layout_rss_podcast_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_label_control(
        "label_confirm_delete_rss_mode",
        state.label_confirm_delete_rss_mode,
        "combo_confirm_delete_rss_mode",
        state.combo_confirm_delete_rss_mode,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_confirm_delete_podcast_mode",
        state.label_confirm_delete_podcast_mode,
        "combo_confirm_delete_podcast_mode",
        state.combo_confirm_delete_podcast_mode,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_rss_quick_copy_mode",
        state.label_rss_quick_copy_mode,
        "combo_rss_quick_copy_mode",
        state.combo_rss_quick_copy_mode,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_checkbox(
        "checkbox_rss_show_article_preview",
        state.checkbox_rss_show_article_preview,
        y,
    );
    y = layout_checkbox(
        "checkbox_announce_unread_rss_podcast",
        state.checkbox_announce_unread_rss_podcast,
        y,
    );
    y = layout_label_control(
        "label_unread_label_position",
        state.label_unread_label_position,
        "combo_unread_label_position",
        state.combo_unread_label_position,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_rss_date_display",
        state.label_rss_date_display,
        "combo_rss_date_display",
        state.combo_rss_date_display,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_rss_time_display",
        state.label_rss_time_display,
        "combo_rss_time_display",
        state.combo_rss_time_display,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_podcast_date_display",
        state.label_podcast_date_display,
        "combo_podcast_date_display",
        state.combo_podcast_date_display,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_podcast_time_display",
        state.label_podcast_time_display,
        "combo_podcast_time_display",
        state.combo_podcast_time_display,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y += OPTIONS_SECTION_GAP;
    y = layout_label_control(
        "label_podcast_cache_limit",
        state.label_podcast_cache_limit,
        "edit_podcast_cache_limit",
        state.edit_podcast_cache_limit,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_podcastindex_key",
        state.label_podcastindex_key,
        "edit_podcastindex_key",
        state.edit_podcastindex_key,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_label_control(
        "label_podcastindex_secret",
        state.label_podcastindex_secret,
        "edit_podcastindex_secret",
        state.edit_podcastindex_secret,
        y,
        OPTIONS_EDIT_HEIGHT,
    );
    y = layout_button(
        "button_podcastindex_signup",
        state.button_podcastindex_signup,
        y,
    );
    y + scroll_offset
}

fn layout_ai_transcription_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_label_control(
        "label_whisper_model",
        state.label_whisper_model,
        "combo_whisper_model",
        state.combo_whisper_model,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_checkbox("checkbox_whisper_cuda", state.checkbox_whisper_cuda, y);
    y = layout_checkbox(
        "checkbox_whisper_keep_original_language",
        state.checkbox_whisper_keep_original_language,
        y,
    );
    y = layout_checkbox(
        "checkbox_whisper_include_timestamps",
        state.checkbox_whisper_include_timestamps,
        y,
    );
    y + scroll_offset
}

fn layout_shortcuts_tab(state: &OptionsDialogState, scroll_offset: i32) -> i32 {
    let mut y = OPTIONS_CONTENT_TOP - scroll_offset;
    y = layout_label_control(
        "label_shortcut_action",
        state.label_shortcut_action,
        "combo_shortcut_action",
        state.combo_shortcut_action,
        y,
        OPTIONS_COMBO_HEIGHT,
    );
    y = layout_label_control(
        "label_shortcut_value",
        state.label_shortcut_value,
        "edit_shortcut_value",
        state.edit_shortcut_value,
        y,
        OPTIONS_EDIT_HEIGHT,
    );

    layout_button("button_shortcut_change", state.button_shortcut_change, y);
    layout_button("button_shortcut_reset", state.button_shortcut_reset, y);
    unsafe {
        crate::log_if_err!(MoveWindow(
            state.button_shortcut_change,
            OPTIONS_CONTROL_X,
            y,
            (OPTIONS_CONTROL_WIDTH / 2) - 4,
            OPTIONS_BUTTON_HEIGHT,
            true,
        ));
        crate::log_if_err!(MoveWindow(
            state.button_shortcut_reset,
            OPTIONS_CONTROL_X + (OPTIONS_CONTROL_WIDTH / 2) + 4,
            y,
            (OPTIONS_CONTROL_WIDTH / 2) - 4,
            OPTIONS_BUTTON_HEIGHT,
            true,
        ));
    }
    y = layout_button(
        "button_shortcut_reset_all",
        state.button_shortcut_reset_all,
        y + OPTIONS_ROW_HEIGHT,
    );
    y + scroll_offset
}

fn tab_array_index(index: i32) -> usize {
    index.clamp(0, OPTIONS_TAB_COUNT - 1) as usize
}

fn options_viewport_height(hwnd: HWND) -> i32 {
    let mut rect = RECT::default();
    unsafe {
        if GetClientRect(hwnd, &mut rect).is_err() {
            crate::log_debug("GetClientRect failed for options viewport");
            return 0;
        }
    }
    let button_y = (rect.bottom - 40).max(OPTIONS_CONTENT_TOP);
    (button_y - OPTIONS_CONTENT_TOP - OPTIONS_CONTENT_BOTTOM_GAP).max(0)
}

fn clamp_scroll_offset(offset: i32, content_height: i32, viewport_height: i32) -> i32 {
    let max_offset = (content_height - viewport_height).max(0);
    offset.clamp(0, max_offset)
}

fn layout_tab_content(state: &OptionsDialogState, index: i32, scroll_offset: i32) -> i32 {
    match index {
        OPTIONS_TAB_GENERAL => layout_general_tab(state, scroll_offset),
        OPTIONS_TAB_VOICE => layout_voice_tab(state, scroll_offset),
        OPTIONS_TAB_EDITOR => layout_editor_tab(state, scroll_offset),
        OPTIONS_TAB_AUDIO => layout_audio_tab(state, scroll_offset),
        OPTIONS_TAB_RSS_PODCAST => layout_rss_podcast_tab(state, scroll_offset),
        OPTIONS_TAB_AI_TRANSCRIPTION => layout_ai_transcription_tab(state, scroll_offset),
        OPTIONS_TAB_SHORTCUTS => layout_shortcuts_tab(state, scroll_offset),
        _ => OPTIONS_CONTENT_TOP,
    }
}

fn update_options_scrollbar(hwnd: HWND, state: &OptionsDialogState) {
    let idx = tab_array_index(state.active_tab);
    let viewport_height = options_viewport_height(hwnd);
    let content_height = state.content_heights[idx].max(0);
    let max_offset = (content_height - viewport_height).max(0);
    let position = clamp_scroll_offset(state.scroll_offsets[idx], content_height, viewport_height);

    let info = SCROLLINFO {
        cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
        fMask: SIF_RANGE | SIF_PAGE | SIF_POS,
        nMin: 0,
        nMax: max_offset,
        nPage: viewport_height.max(1) as u32,
        nPos: position,
        ..Default::default()
    };
    if unsafe { SetScrollInfo(hwnd, SB_VERT, &info, true) } == 0 {
        crate::log_debug("SetScrollInfo failed for options dialog");
    }
    if unsafe { ShowScrollBar(hwnd, SB_VERT, max_offset > 0) }.is_err() {
        crate::log_debug("ShowScrollBar failed for options dialog");
    }
}

fn relayout_active_tab_content(hwnd: HWND) {
    if with_options_state(hwnd, |state| {
        let idx = tab_array_index(state.active_tab);
        let viewport_height = options_viewport_height(hwnd);
        let content_height = state.content_heights[idx].max(0);
        state.scroll_offsets[idx] =
            clamp_scroll_offset(state.scroll_offsets[idx], content_height, viewport_height);
        let scroll_offset = state.scroll_offsets[idx];
        state.content_heights[idx] = layout_tab_content(state, state.active_tab, scroll_offset);
        layout_dialog_buttons(hwnd, state.ok_button, state.cancel_button);
        update_options_scrollbar(hwnd, state);
    })
    .is_none()
    {
        crate::log_debug("Failed to access state in options_window");
    }
}

fn wheel_delta_from_wparam(wparam: WPARAM) -> i32 {
    (((wparam.0 >> 16) & 0xffff) as i16) as i32
}

fn handle_options_mouse_wheel(hwnd: HWND, wparam: WPARAM) -> bool {
    let delta = wheel_delta_from_wparam(wparam);
    if delta == 0 {
        return false;
    }
    let steps = (delta.abs() / OPTIONS_WHEEL_DELTA).max(1);
    let delta_pixels = steps * OPTIONS_SCROLL_LINE;

    let changed = with_options_state(hwnd, |state| {
        let idx = tab_array_index(state.active_tab);
        let viewport_height = options_viewport_height(hwnd);
        let content_height = state.content_heights[idx].max(0);
        let old = state.scroll_offsets[idx];
        let mut new = if delta > 0 {
            old - delta_pixels
        } else {
            old + delta_pixels
        };
        new = clamp_scroll_offset(new, content_height, viewport_height);
        if new != old {
            state.scroll_offsets[idx] = new;
            true
        } else {
            false
        }
    })
    .unwrap_or(false);

    if changed {
        relayout_active_tab_content(hwnd);
    }
    changed
}

fn handle_options_vscroll(hwnd: HWND, wparam: WPARAM) -> bool {
    let command = (wparam.0 & 0xffff) as u32;
    let thumb_pos = ((wparam.0 >> 16) & 0xffff) as i32;

    let changed = with_options_state(hwnd, |state| {
        let idx = tab_array_index(state.active_tab);
        let viewport_height = options_viewport_height(hwnd);
        let content_height = state.content_heights[idx].max(0);
        let max_offset = (content_height - viewport_height).max(0);
        if max_offset <= 0 {
            if state.scroll_offsets[idx] != 0 {
                state.scroll_offsets[idx] = 0;
                return true;
            }
            return false;
        }

        let old = state.scroll_offsets[idx];
        let page_step = (viewport_height - OPTIONS_ROW_HEIGHT).max(OPTIONS_SCROLL_LINE);
        let mut new = old;
        match command {
            c if c == SB_LINEUP.0 as u32 => new -= OPTIONS_SCROLL_LINE,
            c if c == SB_LINEDOWN.0 as u32 => new += OPTIONS_SCROLL_LINE,
            c if c == SB_PAGEUP.0 as u32 => new -= page_step,
            c if c == SB_PAGEDOWN.0 as u32 => new += page_step,
            c if c == SB_TOP.0 as u32 => new = 0,
            c if c == SB_BOTTOM.0 as u32 => new = max_offset,
            c if c == SB_THUMBPOSITION.0 as u32 || c == SB_THUMBTRACK.0 as u32 => {
                let mut info = SCROLLINFO {
                    cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
                    fMask: SIF_TRACKPOS,
                    ..Default::default()
                };
                if unsafe { GetScrollInfo(hwnd, SB_VERT, &mut info) }.is_ok() {
                    new = info.nTrackPos;
                } else {
                    new = thumb_pos;
                }
            }
            _ => {}
        }
        new = new.clamp(0, max_offset);
        if new != old {
            state.scroll_offsets[idx] = new;
            true
        } else {
            false
        }
    })
    .unwrap_or(false);

    if changed {
        relayout_active_tab_content(hwnd);
    }
    changed
}

fn set_active_tab(hwnd: HWND, index: i32) {
    let index = index.clamp(0, OPTIONS_TAB_COUNT - 1);
    let focus_first = with_options_state(hwnd, |state| {
        if state.focus_initialized {
            false
        } else {
            state.focus_initialized = true;
            true
        }
    })
    .unwrap_or(false);
    if with_options_state(hwnd, |state| {
        state.active_tab = index;
        let idx = tab_array_index(index);
        let viewport_height = options_viewport_height(hwnd);
        state.scroll_offsets[idx] = clamp_scroll_offset(
            state.scroll_offsets[idx],
            state.content_heights[idx],
            viewport_height,
        );
        let scroll_offset = state.scroll_offsets[idx];
        state.content_heights[idx] = layout_tab_content(state, index, scroll_offset);
        layout_dialog_buttons(hwnd, state.ok_button, state.cancel_button);
        update_options_scrollbar(hwnd, state);

        let show_general = index == OPTIONS_TAB_GENERAL;
        let show_voice = index == OPTIONS_TAB_VOICE;
        let show_editor = index == OPTIONS_TAB_EDITOR;
        let show_audio = index == OPTIONS_TAB_AUDIO;
        let show_rss_podcast = index == OPTIONS_TAB_RSS_PODCAST;
        let show_ai_transcription = index == OPTIONS_TAB_AI_TRANSCRIPTION;
        let show_shortcuts = index == OPTIONS_TAB_SHORTCUTS;

        for control in [
            state.label_language,
            state.combo_lang,
            state.label_modified_marker_position,
            state.combo_modified_marker_position,
            state.label_open,
            state.combo_open,
            state.label_prompt_program,
            state.combo_prompt_program,
            state.label_network_proxy,
            state.edit_network_proxy,
            state.label_network_proxy_username,
            state.edit_network_proxy_username,
            state.label_network_proxy_password,
            state.edit_network_proxy_password,
            state.checkbox_check_updates,
            state.checkbox_check_beta_updates,
            state.checkbox_send_crash_reports,
            state.checkbox_use_legacy_name,
            state.checkbox_context_menu,
            state.label_file_associations,
            state.button_manage_associations,
        ] {
            crate::show_window_safe(control, if show_general { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_tts_engine,
            state.combo_tts_engine,
            state.label_tts_voice_language,
            state.combo_tts_voice_language,
            state.label_voice_profile,
            state.combo_voice_profile,
            state.button_rename_voice_profile,
            state.button_add_voice_profile,
            state.button_delete_voice_profile,
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
            state.button_tts_insert_tag,
            state.checkbox_multilingual,
            state.checkbox_tts_manual,
            state.checkbox_split_on_newline,
            state.checkbox_use_dialogue_voice,
            state.label_dialogue_engine,
            state.combo_dialogue_engine,
            state.label_dialogue_voice_language,
            state.combo_dialogue_voice_language,
            state.label_dialogue_voice,
            state.combo_dialogue_voice,
            state.checkbox_dialogue_multilingual,
            state.label_dialogue_voice_rate,
            state.combo_dialogue_voice_rate,
            state.edit_dialogue_voice_rate,
            state.label_dialogue_voice_pitch,
            state.combo_dialogue_voice_pitch,
            state.edit_dialogue_voice_pitch,
            state.label_dialogue_voice_volume,
            state.combo_dialogue_voice_volume,
            state.edit_dialogue_voice_volume,
            state.checkbox_dialogue_use_secondary_voice,
            state.label_dialogue_secondary_engine,
            state.combo_dialogue_secondary_engine,
            state.label_dialogue_secondary_voice_language,
            state.combo_dialogue_secondary_voice_language,
            state.label_dialogue_secondary_voice,
            state.combo_dialogue_secondary_voice,
            state.checkbox_dialogue_secondary_multilingual,
            state.label_dialogue_secondary_voice_rate,
            state.combo_dialogue_secondary_voice_rate,
            state.edit_dialogue_secondary_voice_rate,
            state.label_dialogue_secondary_voice_pitch,
            state.combo_dialogue_secondary_voice_pitch,
            state.edit_dialogue_secondary_voice_pitch,
            state.label_dialogue_secondary_voice_volume,
            state.combo_dialogue_secondary_voice_volume,
            state.edit_dialogue_secondary_voice_volume,
            state.label_dialogue_open_quote,
            state.edit_dialogue_open_quote,
            state.label_dialogue_close_quote,
            state.edit_dialogue_close_quote,
            state.checkbox_dialogue_allow_multiline,
            state.button_dialogue_voice_preview,
            state.button_dialogue_secondary_voice_preview,
        ] {
            crate::show_window_safe(control, if show_voice { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.checkbox_word_wrap,
            state.checkbox_smart_quotes,
            state.checkbox_strip_markdown_keep_bullets,
            state.checkbox_spellcheck,
            state.label_spellcheck_language,
            state.combo_spellcheck_language,
            state.label_dictionary_translation,
            state.combo_dictionary_translation,
            state.label_wikipedia_language,
            state.combo_wikipedia_language,
            state.label_wrap_width,
            state.edit_wrap_width,
            state.label_indentation,
            state.combo_indentation,
            state.label_tab_width,
            state.combo_tab_width,
            state.label_space_width,
            state.combo_space_width,
            state.label_quote_prefix,
            state.edit_quote_prefix,
            state.label_interpreter_path,
            state.edit_interpreter_path,
            state.button_interpreter_browse,
            state.button_interpreter_search,
            state.checkbox_move_cursor,
        ] {
            crate::show_window_safe(control, if show_editor { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_audio_skip,
            state.combo_audio_skip,
            state.label_default_save_folder_kind,
            state.combo_default_save_folder_kind,
            state.label_audiobook_save_folder,
            state.edit_audiobook_save_folder,
            state.button_audiobook_save_folder_browse,
            state.checkbox_show_media_save_confirmation,
            state.label_audio_split,
            state.combo_audio_split,
            state.label_audio_split_minutes,
            state.combo_audio_split_minutes,
            state.label_audio_split_parts_count,
            state.edit_audio_split_parts_count,
            state.label_audio_split_start_number,
            state.combo_audio_split_start_number,
            state.label_audiobook_part_naming,
            state.combo_audiobook_part_naming,
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
            state.checkbox_audio_split_epub_chapters,
            state.checkbox_subtitle_ducking,
            state.label_subtitle_mode,
            state.combo_subtitle_mode,
            state.label_subtitle_offset,
            state.edit_subtitle_offset,
        ] {
            crate::show_window_safe(control, if show_audio { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_confirm_delete_rss_mode,
            state.combo_confirm_delete_rss_mode,
            state.label_confirm_delete_podcast_mode,
            state.combo_confirm_delete_podcast_mode,
            state.label_rss_quick_copy_mode,
            state.combo_rss_quick_copy_mode,
            state.label_podcast_cache_limit,
            state.edit_podcast_cache_limit,
            state.checkbox_rss_show_article_preview,
            state.checkbox_announce_unread_rss_podcast,
            state.label_unread_label_position,
            state.combo_unread_label_position,
            state.label_rss_date_display,
            state.combo_rss_date_display,
            state.label_rss_time_display,
            state.combo_rss_time_display,
            state.label_podcast_date_display,
            state.combo_podcast_date_display,
            state.label_podcast_time_display,
            state.combo_podcast_time_display,
            state.label_podcastindex_key,
            state.edit_podcastindex_key,
            state.label_podcastindex_secret,
            state.edit_podcastindex_secret,
            state.button_podcastindex_signup,
        ] {
            crate::show_window_safe(control, if show_rss_podcast { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_shortcut_action,
            state.combo_shortcut_action,
            state.label_shortcut_value,
            state.edit_shortcut_value,
            state.button_shortcut_change,
            state.button_shortcut_reset,
            state.button_shortcut_reset_all,
        ] {
            crate::show_window_safe(control, if show_shortcuts { SW_SHOW } else { SW_HIDE });
        }

        for control in [
            state.label_whisper_model,
            state.combo_whisper_model,
            state.checkbox_whisper_cuda,
            state.checkbox_whisper_keep_original_language,
            state.checkbox_whisper_include_timestamps,
            state.label_dictation_microphone,
            state.combo_dictation_microphone,
        ] {
            crate::show_window_safe(
                control,
                if show_ai_transcription {
                    SW_SHOW
                } else {
                    SW_HIDE
                },
            );
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access state in options_window");
    }

    if index == OPTIONS_TAB_AUDIO {
        update_audio_split_visibility(hwnd);
        update_subtitle_ducking_visibility(hwnd);
    } else if index == OPTIONS_TAB_VOICE {
        refresh_voices(hwnd);
        update_tts_manual_visibility(hwnd);
        update_dialogue_voice_visibility(hwnd);
        update_voice_profile_delete_button_visibility(hwnd);
    } else if index == OPTIONS_TAB_EDITOR {
        update_indentation_visibility(hwnd);
    } else if index == OPTIONS_TAB_SHORTCUTS {
        update_shortcut_binding_text(hwnd);
    } else if let Some((
        label_text,
        edit_text,
        checkbox,
        label_minutes,
        combo_minutes,
        label_parts,
        edit_parts,
        label_start,
        combo_start,
    )) = with_options_state(hwnd, |state| {
        (
            state.label_audio_split_text,
            state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline,
            state.label_audio_split_minutes,
            state.combo_audio_split_minutes,
            state.label_audio_split_parts_count,
            state.edit_audio_split_parts_count,
            state.label_audio_split_start_number,
            state.combo_audio_split_start_number,
        )
    }) {
        unsafe {
            ShowWindow(label_text, SW_HIDE);
            ShowWindow(edit_text, SW_HIDE);
            ShowWindow(checkbox, SW_HIDE);
            ShowWindow(label_minutes, SW_HIDE);
            ShowWindow(combo_minutes, SW_HIDE);
            ShowWindow(label_parts, SW_HIDE);
            ShowWindow(edit_parts, SW_HIDE);
            ShowWindow(label_start, SW_HIDE);
            ShowWindow(combo_start, SW_HIDE);
            EnableWindow(edit_text, false);
            EnableWindow(checkbox, false);
            EnableWindow(edit_parts, false);
            EnableWindow(combo_minutes, false);
            EnableWindow(combo_start, false);
        }
    }

    if focus_first {
        focus_tab_first(hwnd, index);
    }
}

fn focus_tab_first(hwnd: HWND, index: i32) {
    let target = with_options_state(hwnd, |state| match index {
        OPTIONS_TAB_GENERAL => state.combo_lang,
        OPTIONS_TAB_VOICE => state.combo_tts_engine,
        OPTIONS_TAB_EDITOR => state.checkbox_word_wrap,
        OPTIONS_TAB_AUDIO => state.combo_audio_skip,
        OPTIONS_TAB_RSS_PODCAST => state.combo_confirm_delete_rss_mode,
        OPTIONS_TAB_AI_TRANSCRIPTION => state.combo_whisper_model,
        OPTIONS_TAB_SHORTCUTS => state.combo_shortcut_action,
        _ => HWND(0),
    })
    .unwrap_or(HWND(0));

    if target.0 != 0 {
        crate::set_focus_safe(target);
        if let Err(_e) =
            crate::post_message_w_safe(hwnd, WM_NEXTDLGCTL, WPARAM(target.0 as usize), LPARAM(1))
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
    }
}

pub(crate) fn ensure_voice_lists_loaded(hwnd: HWND, language: Language) {
    let (has_edge, has_sapi) = {
        with_state(hwnd, |state| {
            (!state.edge_voices.is_empty(), !state.sapi_voices.is_empty())
        })
        .unwrap_or((false, false))
    };

    if !has_edge {
        thread::spawn(move || {
            match fetch_voice_list() {
                Ok(list) => {
                    let payload = Box::new(list);
                    unsafe {
                        if let Err(e) = windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                            hwnd,
                            WM_TTS_VOICES_LOADED,
                            WPARAM(0),
                            LPARAM(Box::into_raw(payload) as isize),
                        ) {
                            crate::log_debug(&format!(
                                "Failed to post WM_TTS_VOICES_LOADED: {}",
                                e
                            ));
                        }
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
            unsafe {
                if let Err(e) = windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                    hwnd,
                    WM_TTS_SAPI_VOICES_LOADED,
                    WPARAM(0),
                    LPARAM(Box::into_raw(payload) as isize),
                ) {
                    crate::log_debug(&format!("Failed to post WM_TTS_SAPI_VOICES_LOADED: {}", e));
                }
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

fn browse_for_interpreter(hwnd: HWND) {
    let mut buffer = [0u16; 1024];
    let language = { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    let executables_label = i18n::tr(language, "dialog.executables");
    let all_files_label = i18n::tr(language, "dialog.all_files");
    let filter = to_wide(&format!(
        "{} (*.exe)\0*.exe\0{} (*.*)\0*.*\0\0",
        executables_label, all_files_label
    ));
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        ..Default::default()
    };

    if crate::get_open_file_name_w_safe(&mut ofn).as_bool() {
        let len = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
        let path = String::from_utf16_lossy(&buffer[..len]);
        if let Some(edit) = with_options_state(hwnd, |state| state.edit_interpreter_path) {
            crate::log_if_err!(crate::set_window_text_w_safe(
                edit,
                PCWSTR(to_wide(&path).as_ptr())
            ));
        }
    }
}

fn selected_default_save_folder_kind(hwnd: HWND) -> u32 {
    with_options_state(hwnd, |state| {
        let sel = crate::send_message_w_safe(
            state.combo_default_save_folder_kind,
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0;
        if sel < 0 {
            return DEFAULT_SAVE_FOLDER_AUDIOBOOK;
        }
        let data = crate::send_message_w_safe(
            state.combo_default_save_folder_kind,
            CB_GETITEMDATA,
            WPARAM(sel as usize),
            LPARAM(0),
        )
        .0;
        if data < 0 {
            DEFAULT_SAVE_FOLDER_AUDIOBOOK
        } else {
            data as u32
        }
    })
    .unwrap_or(DEFAULT_SAVE_FOLDER_AUDIOBOOK)
}

fn persist_default_save_folder_edit(hwnd: HWND) {
    let text =
        if let Some(edit) = with_options_state(hwnd, |state| state.edit_audiobook_save_folder) {
            let len = crate::get_window_text_length_w_safe(edit);
            if len > 0 {
                let mut buf = vec![0u16; (len + 1) as usize];
                let read = crate::get_window_text_w_safe(edit, &mut buf);
                String::from_utf16_lossy(&buf[..read as usize])
                    .trim()
                    .to_string()
            } else {
                String::new()
            }
        } else {
            String::new()
        };
    let _updated = with_options_state(hwnd, |state| match state.default_save_folder_selection {
        DEFAULT_SAVE_FOLDER_MEDIA => state.default_save_folder_media = text,
        DEFAULT_SAVE_FOLDER_DOCUMENTS => state.default_save_folder_documents = text,
        _ => state.default_save_folder_audiobook = text,
    });
}

fn refresh_default_save_folder_edit(hwnd: HWND) {
    let data = with_options_state(hwnd, |state| {
        let text = match state.default_save_folder_selection {
            DEFAULT_SAVE_FOLDER_MEDIA => state.default_save_folder_media.clone(),
            DEFAULT_SAVE_FOLDER_DOCUMENTS => state.default_save_folder_documents.clone(),
            _ => state.default_save_folder_audiobook.clone(),
        };
        (state.edit_audiobook_save_folder, text)
    });
    if let Some((edit, text)) = data {
        crate::log_if_err!(crate::set_window_text_w_safe(
            edit,
            PCWSTR(to_wide(&text).as_ptr())
        ));
    }
}

fn on_default_save_folder_kind_changed(hwnd: HWND) {
    persist_default_save_folder_edit(hwnd);
    let new_kind = selected_default_save_folder_kind(hwnd);
    let _updated = with_options_state(hwnd, |state| {
        state.default_save_folder_selection = new_kind;
    });
    refresh_default_save_folder_edit(hwnd);
}

fn browse_for_audiobook_folder(hwnd: HWND) {
    browse_for_folder_into_edit(hwnd, |state| state.edit_audiobook_save_folder);
    persist_default_save_folder_edit(hwnd);
}

fn browse_for_folder_into_edit<F>(hwnd: HWND, edit_selector: F)
where
    F: Fn(&mut OptionsDialogState) -> HWND,
{
    let language = with_options_state(hwnd, |state| {
        { with_state(state.parent, |app| app.settings.language) }.unwrap_or_default()
    })
    .unwrap_or_default();
    if let Some(folder) =
        crate::app_windows::find_in_files_window::browse_for_folder(hwnd, language)
    {
        let path = folder.to_string_lossy().to_string();
        if let Some(edit) = with_options_state(hwnd, |state| edit_selector(state))
            && let Err(_e) = crate::set_window_text_w_safe(edit, PCWSTR(to_wide(&path).as_ptr()))
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
    }
}

fn search_for_interpreter(hwnd: HWND) {
    let query = if let Some(edit) = with_options_state(hwnd, |state| state.edit_interpreter_path) {
        let len = crate::get_window_text_length_w_safe(edit);
        if len > 0 {
            let mut buf = vec![0u16; (len + 1) as usize];
            let read = crate::get_window_text_w_safe(edit, &mut buf);
            String::from_utf16_lossy(&buf[..read as usize])
        } else {
            String::new()
        }
    } else {
        String::new()
    };

    let parent = with_options_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
    let language = { with_state(parent, |state| state.settings.language) }.unwrap_or_default();

    if query.trim().is_empty() {
        let msg = i18n::tr(language, "options.interpreter_search.empty_query");
        crate::show_info(hwnd, language, &msg);
        return;
    }

    // Try both with and without .exe extension
    let exec_name = if query.to_lowercase().ends_with(".exe") {
        query.clone()
    } else {
        format!("{}.exe", query)
    };

    let output = Command::new("where").arg(&exec_name).output();
    let mut paths = match output {
        Ok(out) if out.status.success() => {
            let s = String::from_utf8_lossy(&out.stdout);
            s.lines()
                .map(|l| l.trim().to_string())
                .filter(|l| !l.is_empty())
                .collect::<Vec<_>>()
        }
        _ => Vec::new(),
    };

    // If still no paths, try exact query
    if paths.is_empty() {
        let output = Command::new("where").arg(&query).output();
        if let Ok(out) = output
            && out.status.success()
        {
            let s = String::from_utf8_lossy(&out.stdout);
            paths = s
                .lines()
                .map(|l| l.trim().to_string())
                .filter(|l| !l.is_empty())
                .collect::<Vec<_>>();
        }
    }

    if paths.is_empty() {
        let msg = i18n::tr_f(
            language,
            "options.interpreter_search.no_results",
            &[("query", &query)],
        );
        crate::show_info(hwnd, language, &msg);
        return;
    }

    if let Some(selected) = interpreter_select_window::select_interpreter(hwnd, paths, language)
        && let Some(edit) = with_options_state(hwnd, |state| state.edit_interpreter_path)
    {
        crate::log_if_err!(crate::set_window_text_w_safe(
            edit,
            PCWSTR(to_wide(&selected).as_ptr())
        ));
    }
}
