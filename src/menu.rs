use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::{Language, ShortcutBinding, ShortcutSettings, format_shortcut};
use crate::with_state;
use std::path::Path;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreateMenu, DeleteMenu, DestroyMenu, DrawMenuBar, GetMenu, GetMenuItemCount,
    GetSubMenu, HMENU, InsertMenuW, IsMenu, MENU_ITEM_FLAGS, MF_BYPOSITION, MF_CHECKED, MF_GRAYED,
    MF_POPUP, MF_SEPARATOR, MF_STRING, MF_UNCHECKED, SetMenu,
};
use windows::core::PCWSTR;

pub const IDM_FILE_NEW: usize = 1001;
pub const IDM_FILE_OPEN: usize = 1002;
pub const IDM_FILE_SAVE: usize = 1003;
pub const IDM_FILE_SAVE_AS: usize = 1004;
pub const IDM_FILE_SAVE_ALL: usize = 1005;
pub const IDM_FILE_CLOSE: usize = 1006;
pub const IDM_FILE_EXIT: usize = 1007;
pub const IDM_FILE_READ_START: usize = 1008;
pub const IDM_FILE_READ_PAUSE: usize = 1009;
pub const IDM_FILE_READ_STOP: usize = 1010;
pub const IDM_FILE_READ_PREVIOUS_SENTENCE: usize = 1017;
pub const IDM_FILE_READ_NEXT_SENTENCE: usize = 1018;
pub const IDM_FILE_EXECUTE: usize = 1015;
pub const IDM_FILE_AUDIOBOOK: usize = 1011;
pub const IDM_FILE_PODCAST: usize = 1012;
pub const IDM_FILE_BATCH_AUDIOBOOK: usize = 1013;
pub const IDM_FILE_CLOSE_OTHERS: usize = 1014;
pub const IDM_FILE_CONVERT_AUDIO: usize = 1016;
pub const IDM_EDIT_UNDO: usize = 2001;
pub const IDM_EDIT_CUT: usize = 2002;
pub const IDM_EDIT_COPY: usize = 2003;
pub const IDM_EDIT_PASTE: usize = 2004;
pub const IDM_EDIT_SELECT_ALL: usize = 2005;
pub const IDM_EDIT_FIND: usize = 2006;
pub const IDM_EDIT_FIND_NEXT: usize = 2007;
pub const IDM_EDIT_REPLACE: usize = 2008;
pub const IDM_EDIT_FIND_IN_FILES: usize = 2009;
pub const IDM_EDIT_STRIP_MARKDOWN: usize = 2010;
pub const IDM_EDIT_NORMALIZE_WHITESPACE: usize = 2011;
pub const IDM_EDIT_HARD_LINE_BREAK: usize = 2012;
pub const IDM_EDIT_ORDER_ITEMS: usize = 2013;
pub const IDM_EDIT_KEEP_UNIQUE_ITEMS: usize = 2014;
pub const IDM_EDIT_REVERSE_ITEMS: usize = 2015;
pub const IDM_EDIT_QUOTE_LINES: usize = 2016;
pub const IDM_EDIT_UNQUOTE_LINES: usize = 2017;
pub const IDM_EDIT_TEXT_STATS: usize = 2018;
pub const IDM_EDIT_JOIN_LINES: usize = 2019;
pub const IDM_EDIT_CLEAN_EOL_HYPHENS: usize = 2020;
pub const IDM_EDIT_REMOVE_DUPLICATE_LINES: usize = 2021;
pub const IDM_EDIT_REMOVE_DUPLICATE_CONSECUTIVE_LINES: usize = 2022;
pub const IDM_EDIT_PREV_SPELLING_ERROR: usize = 2023;
pub const IDM_EDIT_NEXT_SPELLING_ERROR: usize = 2024;
pub const IDM_EDIT_GO_TO_LINE: usize = 2025;
pub const IDM_EDIT_AUDIOBOOK_SELECTION: usize = 2026;
pub const IDM_EDIT_AUTO_FORMAT_TTS: usize = 2027;
pub const IDM_EDIT_INDENT: usize = 2028;
pub const IDM_EDIT_OUTDENT: usize = 2029;
pub const IDM_EDIT_FIND_PREVIOUS: usize = 2030;
pub const IDM_EDIT_INSERT_ELLIPSIS: usize = 2031;
pub const IDM_SPELLCHECK_SUGGESTION_BASE: usize = 12000;
pub const IDM_SPELLCHECK_SUGGESTION_MAX: usize = 10;
pub const IDM_SPELLCHECK_ADD_TO_DICTIONARY: usize = 12100;
pub const IDM_SPELLCHECK_IGNORE_ONCE: usize = 12101;
pub const IDM_PLAYBACK_PLAY_PAUSE: usize = 8001;
pub const IDM_PLAYBACK_STOP: usize = 8002;
pub const IDM_PLAYBACK_SEEK_FORWARD: usize = 8003;
pub const IDM_PLAYBACK_SEEK_BACKWARD: usize = 8004;
pub const IDM_PLAYBACK_GO_TO_TIME: usize = 8005;
pub const IDM_PLAYBACK_ANNOUNCE_TIME: usize = 8006;
pub const IDM_PLAYBACK_VOLUME_UP: usize = 8007;
pub const IDM_PLAYBACK_VOLUME_DOWN: usize = 8008;
pub const IDM_PLAYBACK_VOLUME_RESET: usize = 8024;
pub const IDM_PLAYBACK_MUTE_TOGGLE: usize = 8009;
pub const IDM_PLAYBACK_SPEED_UP: usize = 8010;
pub const IDM_PLAYBACK_SPEED_DOWN: usize = 8011;
pub const IDM_PLAYBACK_SPEED_RESET: usize = 8016;
pub const IDM_PLAYBACK_PITCH_UP: usize = 8017;
pub const IDM_PLAYBACK_PITCH_DOWN: usize = 8018;
pub const IDM_PLAYBACK_PITCH_RESET: usize = 8021;
pub const IDM_PLAYBACK_TRACK_PREV: usize = 8022;
pub const IDM_PLAYBACK_TRACK_NEXT: usize = 8023;
pub const IDM_PLAYBACK_SEEK_TO_START: usize = 8025;
pub const IDM_PLAYBACK_SEEK_TO_END: usize = 8026;
pub const IDM_PLAYBACK_TRANSCRIBE_CURRENT: usize = 8027;
pub const IDM_PLAYBACK_TRANSCRIBE_CURRENT_FOLDER: usize = 8028;
pub const IDM_PLAYBACK_TRANSCRIBE_CANCEL: usize = 8030;
pub const IDM_PLAYBACK_ADD_SUBTITLES: usize = 8019;
pub const IDM_PLAYBACK_REMOVE_SUBTITLES: usize = 8020;
pub const IDM_PLAYBACK_CHAPTER_PREV: usize = 8012;
pub const IDM_PLAYBACK_CHAPTER_NEXT: usize = 8013;
pub const IDM_PLAYBACK_CHAPTER_LIST: usize = 8014;
pub const IDM_PLAYBACK_DOWNLOAD_EPISODE: usize = 8015;
pub const IDM_PLAYBACK_AUDIO_TRACK_BASE: usize = 8100;
pub const IDM_PLAYBACK_AUDIO_TRACK_MAX: usize = 20;
pub const IDM_INSERT_BOOKMARK: usize = 2101;
pub const IDM_MANAGE_BOOKMARKS: usize = 2102;
pub const IDM_INSERT_CLEAR_BOOKMARKS: usize = 2103;
pub const IDM_GOTO_NEXT_BOOKMARK: usize = 2104;
pub const IDM_GOTO_PREV_BOOKMARK: usize = 2105;
pub const IDM_AUTOMATIC_BOOKMARK: usize = 2106;
pub const IDM_NEXT_TAB: usize = 3001;
pub const IDM_WINDOW_OPEN_DOCUMENTS: usize = 3002;
pub const IDM_WINDOW_CLOSE_ALL: usize = 3003;
pub const IDM_VIEW_SHOW_VOICES: usize = 6101;
pub const IDM_VIEW_SHOW_FAVORITES: usize = 6102;
pub const IDM_VIEW_READ_ONLY: usize = 6103;
pub const IDM_VIEW_WORD_WRAP: usize = 6104;
pub const IDM_VIEW_SHOW_VIDEO_DURING_PLAYBACK: usize = 6105;
pub const IDM_VIEW_FONT_ARIAL: usize = 6401;
pub const IDM_VIEW_FONT_CALIBRI: usize = 6402;
pub const IDM_VIEW_FONT_CONSOLAS: usize = 6403;
pub const IDM_VIEW_FONT_SEGOE_UI: usize = 6404;
pub const IDM_VIEW_FONT_TAHOMA: usize = 6405;
pub const IDM_VIEW_FONT_VERDANA: usize = 6406;
pub const IDM_VIEW_FONT_TIMES_NEW_ROMAN: usize = 6407;
pub const IDM_VIEW_FONT_GEORGIA: usize = 6408;
pub const IDM_VIEW_TEXT_COLOR_BLACK: usize = 6201;
pub const IDM_VIEW_TEXT_COLOR_DARK_BLUE: usize = 6202;
pub const IDM_VIEW_TEXT_COLOR_DARK_GREEN: usize = 6203;
pub const IDM_VIEW_TEXT_COLOR_DARK_BROWN: usize = 6204;
pub const IDM_VIEW_TEXT_COLOR_DARK_GRAY: usize = 6205;
pub const IDM_VIEW_TEXT_COLOR_LIGHT_BLUE: usize = 6206;
pub const IDM_VIEW_TEXT_COLOR_LIGHT_GREEN: usize = 6207;
pub const IDM_VIEW_TEXT_COLOR_LIGHT_BROWN: usize = 6208;
pub const IDM_VIEW_TEXT_COLOR_LIGHT_GRAY: usize = 6209;
pub const IDM_VIEW_TEXT_SIZE_SMALL: usize = 6301;
pub const IDM_VIEW_TEXT_SIZE_NORMAL: usize = 6302;
pub const IDM_VIEW_TEXT_SIZE_LARGE: usize = 6303;
pub const IDM_VIEW_TEXT_SIZE_XLARGE: usize = 6304;
pub const IDM_VIEW_TEXT_SIZE_XXLARGE: usize = 6305;
pub const IDM_FILE_RECENT_BASE: usize = 4000;
pub const IDM_FILE_RECENT_CLEAR: usize = 4010;
pub const IDM_TOOLS_OPTIONS: usize = 5001;
pub const IDM_TOOLS_DICTIONARY: usize = 5002;
pub const IDM_TOOLS_IMPORT_YOUTUBE: usize = 5003;
pub const IDM_TOOLS_PROMPT: usize = 5004;
pub const IDM_TOOLS_RSS: usize = 5005;
pub const IDM_TOOLS_PODCASTS: usize = 5006;
pub const IDM_TOOLS_DICTIONARY_LOOKUP: usize = 5007;
pub const IDM_TOOLS_WIKIPEDIA_IMPORT: usize = 5008;
pub const IDM_TOOLS_STREAM_AUDIO: usize = 5009;
pub const IDM_TOOLS_BDCIECHI: usize = 5010;
pub const IDM_TOOLS_TOGGLE_DICTATION: usize = 5011;
pub const IDM_TOOLS_RAI_AUDIODESCRIZIONI: usize = 5012;
pub const IDM_TOOLS_RAIPLAYSOUND: usize = 5013;
pub const IDM_TOOLS_RAIPLAY: usize = 5014;
pub const IDM_TOOLS_ITALIAONLINE: usize = 5015;
pub const IDM_TOOLS_RADIO: usize = 5016;
pub const IDM_HELP_GUIDE: usize = 7001;
pub const IDM_HELP_ABOUT: usize = 7002;
pub const IDM_HELP_CHECK_UPDATES: usize = 7003;
pub const IDM_HELP_CHANGELOG: usize = 7004;
pub const IDM_HELP_FEEDBACK: usize = 7005;
pub const IDM_HELP_DONATIONS: usize = 7006;
pub const IDM_HELP_EXPORT_DIAGNOSTICS: usize = 7007;
pub const MAX_RECENT: usize = 5;

pub struct MenuLabels {
    pub menu_file: String,
    pub menu_edit: String,
    pub menu_view: String,
    pub menu_insert: String,
    pub menu_window: String,
    pub menu_voice_audio: String,
    pub menu_tools: String,
    pub menu_help: String,
    pub menu_options: String,
    pub menu_dictionary: String,
    pub menu_dictionary_lookup: String,
    pub menu_wikipedia_import: String,
    pub menu_import_youtube: String,
    pub menu_stream_audio: String,
    pub menu_bdciechi: String,
    pub menu_rai_audiodescrizioni: String,
    pub menu_raiplay: String,
    pub menu_raiplaysound: String,
    pub menu_italiaonline: String,
    pub menu_radio: String,
    pub menu_prompt: String,
    pub menu_rss: String,
    pub menu_podcasts: String,
    pub view_text_color: String,
    pub view_text_size: String,
    pub view_text_color_black: String,
    pub view_text_color_dark_blue: String,
    pub view_text_color_dark_green: String,
    pub view_text_color_dark_brown: String,
    pub view_text_color_dark_gray: String,
    pub view_text_color_light_blue: String,
    pub view_text_color_light_green: String,
    pub view_text_color_light_brown: String,
    pub view_text_color_light_gray: String,
    pub view_text_size_small: String,
    pub view_text_size_normal: String,
    pub view_text_size_large: String,
    pub view_text_size_xlarge: String,
    pub view_text_size_xxlarge: String,
    pub view_show_voices: String,
    pub view_show_favorites: String,
    pub view_read_only: String,
    pub view_word_wrap: String,
    pub view_show_video_during_playback: String,
    pub view_font: String,
    pub file_new: String,
    pub file_open: String,
    pub file_save: String,
    pub file_save_as: String,
    pub file_save_all: String,
    pub file_close: String,
    pub file_recent: String,
    file_read_start: String,
    file_read_previous_sentence: String,
    file_read_next_sentence: String,
    file_execute: String,
    file_read_pause: String,
    pub file_read_stop: String,
    pub file_audiobook: String,
    pub file_podcast: String,
    pub file_batch_audiobooks: String,
    pub file_convert_audio: String,
    pub file_exit: String,
    pub edit_undo: String,
    pub edit_cut: String,
    pub edit_copy: String,
    pub edit_paste: String,
    pub edit_select_all: String,
    pub edit_find: String,
    pub edit_find_next: String,
    pub edit_find_previous: String,
    pub edit_replace: String,
    pub edit_find_in_files: String,
    pub edit_prev_spelling_error: String,
    pub edit_next_spelling_error: String,
    pub edit_goto_line: String,
    pub edit_text_menu: String,
    pub edit_auto_format_tts: String,
    pub edit_strip_markdown: String,
    pub edit_normalize_whitespace: String,
    pub edit_hard_line_break: String,
    pub edit_order_items: String,
    pub edit_keep_unique_items: String,
    pub edit_reverse_items: String,
    pub edit_quote_lines: String,
    pub edit_unquote_lines: String,
    pub edit_indent: String,
    pub edit_outdent: String,
    pub edit_text_stats: String,
    pub edit_join_lines: String,
    pub edit_clean_eol_hyphens: String,
    pub edit_remove_duplicate_lines: String,
    pub edit_remove_duplicate_consecutive_lines: String,
    pub insert_bookmark: String,
    pub automatic_bookmark: String,
    pub insert_goto_next_bookmark: String,
    pub insert_goto_prev_bookmark: String,
    pub insert_clear_bookmarks: String,
    pub manage_bookmarks: String,
    pub window_open_documents: String,
    pub window_close_others: String,
    pub window_close_all: String,
    pub help_guide: String,
    pub help_changelog: String,
    pub help_feedback: String,
    pub help_export_diagnostics: String,
    pub help_donations: String,
    pub help_check_updates: String,
    pub help_about: String,
    pub recent_empty: String,
    pub recent_clear: String,
}

pub fn menu_labels(language: Language) -> MenuLabels {
    MenuLabels {
        menu_file: i18n::tr(language, "menu.file"),
        menu_edit: i18n::tr(language, "menu.edit"),
        menu_view: i18n::tr(language, "menu.view"),
        menu_insert: i18n::tr(language, "menu.insert"),
        menu_window: i18n::tr(language, "menu.window"),
        menu_voice_audio: i18n::tr(language, "menu.voice_audio"),
        menu_tools: i18n::tr(language, "menu.tools"),
        menu_help: i18n::tr(language, "menu.help"),
        menu_options: i18n::tr(language, "menu.options"),
        menu_dictionary: i18n::tr(language, "menu.dictionary"),
        menu_dictionary_lookup: i18n::tr(language, "menu.dictionary_lookup"),
        menu_wikipedia_import: i18n::tr(language, "menu.wikipedia_import"),
        menu_import_youtube: i18n::tr(language, "menu.import_youtube"),
        menu_stream_audio: i18n::tr(language, "menu.stream_audio"),
        menu_bdciechi: if language == Language::Italian {
            i18n::tr(language, "excluded_from_testing.bdciechi.menu_label")
        } else {
            String::new()
        },
        menu_rai_audiodescrizioni: if language == Language::Italian {
            "Rai audiodescrizioni...\tAlt+Shift+A".to_string()
        } else {
            String::new()
        },
        menu_raiplay: if language == Language::Italian {
            "RaiPlay...\tAlt+Shift+P".to_string()
        } else {
            String::new()
        },
        menu_raiplaysound: if language == Language::Italian {
            "RaiPlay Sound...\tCtrl+Shift+S".to_string()
        } else {
            String::new()
        },
        menu_italiaonline: if language == Language::Italian {
            "Pagine Bianche e Gialle...\tAlt+Shift+G".to_string()
        } else {
            String::new()
        },
        menu_radio: if language == Language::Italian {
            i18n::tr(language, "menu.radio")
        } else {
            String::new()
        },
        menu_prompt: i18n::tr(language, "menu.prompt"),
        menu_rss: i18n::tr(language, "menu.rss"),
        menu_podcasts: i18n::tr(language, "menu.podcasts"),
        view_text_color: i18n::tr(language, "view.text_color"),
        view_text_size: i18n::tr(language, "view.text_size"),
        view_text_color_black: i18n::tr(language, "view.text_color.black"),
        view_text_color_dark_blue: i18n::tr(language, "view.text_color.dark_blue"),
        view_text_color_dark_green: i18n::tr(language, "view.text_color.dark_green"),
        view_text_color_dark_brown: i18n::tr(language, "view.text_color.dark_brown"),
        view_text_color_dark_gray: i18n::tr(language, "view.text_color.dark_gray"),
        view_text_color_light_blue: i18n::tr(language, "view.text_color.light_blue"),
        view_text_color_light_green: i18n::tr(language, "view.text_color.light_green"),
        view_text_color_light_brown: i18n::tr(language, "view.text_color.light_brown"),
        view_text_color_light_gray: i18n::tr(language, "view.text_color.light_gray"),
        view_text_size_small: i18n::tr(language, "view.text_size.small"),
        view_text_size_normal: i18n::tr(language, "view.text_size.normal"),
        view_text_size_large: i18n::tr(language, "view.text_size.large"),
        view_text_size_xlarge: i18n::tr(language, "view.text_size.xlarge"),
        view_text_size_xxlarge: i18n::tr(language, "view.text_size.xxlarge"),
        view_show_voices: i18n::tr(language, "view.show_voices"),
        view_show_favorites: i18n::tr(language, "view.show_favorites"),
        view_read_only: i18n::tr(language, "view.read_only"),
        view_word_wrap: i18n::tr(language, "view.word_wrap"),
        view_show_video_during_playback: i18n::tr(language, "view.show_video_during_playback"),
        view_font: i18n::tr(language, "view.font"),
        file_new: i18n::tr(language, "file.new"),
        file_open: i18n::tr(language, "file.open"),
        file_save: i18n::tr(language, "file.save"),
        file_save_as: i18n::tr(language, "file.save_as"),
        file_save_all: i18n::tr(language, "file.save_all"),
        file_close: i18n::tr(language, "file.close"),
        file_recent: i18n::tr(language, "file.recent"),
        file_read_start: i18n::tr(language, "file.read_start"),
        file_read_previous_sentence: i18n::tr(language, "file.read_previous_sentence"),
        file_read_next_sentence: i18n::tr(language, "file.read_next_sentence"),
        file_execute: i18n::tr(language, "file.execute"),
        file_read_pause: i18n::tr(language, "file.read_pause"),
        file_read_stop: i18n::tr(language, "file.read_stop"),
        file_audiobook: i18n::tr(language, "file.audiobook"),
        file_podcast: i18n::tr(language, "file.podcast"),
        file_batch_audiobooks: i18n::tr(language, "file.batch_audiobooks"),
        file_convert_audio: i18n::tr(language, "file.convert_audio"),
        file_exit: i18n::tr(language, "file.exit"),
        edit_undo: i18n::tr(language, "edit.undo"),
        edit_cut: i18n::tr(language, "edit.cut"),
        edit_copy: i18n::tr(language, "edit.copy"),
        edit_paste: i18n::tr(language, "edit.paste"),
        edit_select_all: i18n::tr(language, "edit.select_all"),
        edit_find: i18n::tr(language, "edit.find"),
        edit_find_next: i18n::tr(language, "edit.find_next"),
        edit_find_previous: i18n::tr(language, "edit.find_previous"),
        edit_replace: i18n::tr(language, "edit.replace"),
        edit_find_in_files: i18n::tr(language, "edit.find_in_files"),
        edit_prev_spelling_error: i18n::tr(language, "edit.prev_spelling_error"),
        edit_next_spelling_error: i18n::tr(language, "edit.next_spelling_error"),
        edit_goto_line: i18n::tr(language, "menu.goto_line"),
        edit_text_menu: i18n::tr(language, "edit.text_menu"),
        edit_auto_format_tts: i18n::tr(language, "edit.auto_format_tts"),
        edit_strip_markdown: i18n::tr(language, "edit.strip_markdown"),
        edit_normalize_whitespace: i18n::tr(language, "edit.normalize_whitespace"),
        edit_hard_line_break: i18n::tr(language, "edit.hard_line_break"),
        edit_order_items: i18n::tr(language, "edit.order_items"),
        edit_keep_unique_items: i18n::tr(language, "edit.keep_unique_items"),
        edit_reverse_items: i18n::tr(language, "edit.reverse_items"),
        edit_quote_lines: i18n::tr(language, "edit.quote_lines"),
        edit_unquote_lines: i18n::tr(language, "edit.unquote_lines"),
        edit_indent: i18n::tr(language, "edit.indent"),
        edit_outdent: i18n::tr(language, "edit.outdent"),
        edit_text_stats: i18n::tr(language, "edit.text_stats"),
        edit_join_lines: i18n::tr(language, "edit.join_lines"),
        edit_clean_eol_hyphens: i18n::tr(language, "edit.clean_eol_hyphens"),
        edit_remove_duplicate_lines: i18n::tr(language, "edit.remove_duplicate_lines"),
        edit_remove_duplicate_consecutive_lines: i18n::tr(
            language,
            "edit.remove_duplicate_consecutive_lines",
        ),
        insert_bookmark: i18n::tr(language, "insert.bookmark"),
        automatic_bookmark: i18n::tr(language, "bookmarks.automatic_bookmark"),
        insert_goto_next_bookmark: i18n::tr(language, "insert.goto_next_bookmark"),
        insert_goto_prev_bookmark: i18n::tr(language, "insert.goto_prev_bookmark"),
        insert_clear_bookmarks: i18n::tr(language, "insert.clear_bookmarks"),
        manage_bookmarks: i18n::tr(language, "insert.manage_bookmarks"),
        window_open_documents: i18n::tr(language, "window.open_documents"),
        window_close_others: i18n::tr(language, "window.close_others"),
        window_close_all: i18n::tr(language, "window.close_all"),
        help_guide: i18n::tr(language, "help.guide"),
        help_changelog: i18n::tr(language, "help.changelog"),
        help_feedback: i18n::tr(language, "help.feedback"),
        help_export_diagnostics: i18n::tr(language, "help.export_diagnostics"),
        help_donations: i18n::tr(language, "help.donations"),
        help_check_updates: i18n::tr(language, "help.check_updates"),
        help_about: i18n::tr(language, "help.about"),
        recent_empty: i18n::tr(language, "recent.empty"),
        recent_clear: i18n::tr(language, "recent.clear"),
    }
}

fn strip_menu_shortcut(label: &str) -> &str {
    match label.split_once('\t') {
        Some((left, _)) => left,
        None => label,
    }
}

fn label_with_shortcut(label: &str, binding: ShortcutBinding) -> String {
    format!(
        "{}\t{}",
        strip_menu_shortcut(label),
        format_shortcut(binding)
    )
}

pub fn update_playback_menu(hwnd: HWND, show: bool) {
    unsafe {
        let hmenu = {
            let attached = GetMenu(hwnd);
            if attached.0 != 0 {
                attached
            } else if with_state(hwnd, |state| state.local_mpv_video_mode_active).unwrap_or(false) {
                with_state(hwnd, |state| state.local_mpv_hidden_menu).unwrap_or(HMENU(0))
            } else {
                HMENU(0)
            }
        };
        if hmenu.0 == 0 {
            return;
        }
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        let existing = with_state(hwnd, |state| state.playback_menu).unwrap_or(HMENU(0));
        let show_download = with_state(hwnd, |state| {
            let is_raiplay_live = state.active_podcast_episode_from_rai
                == crate::RaiAudioOrigin::RaiPlay
                && !state.raiplay_live_audio_variants.is_empty();
            let from_rss_doc = state
                .docs
                .get(state.current)
                .map(|doc| doc.from_rss)
                .unwrap_or(false);
            let cached_media_doc = state
                .active_audiobook
                .as_ref()
                .map(|player| crate::is_local_cached_media_path(&player.path))
                .or_else(|| {
                    state.docs.get(state.current).and_then(|doc| {
                        if matches!(doc.format, crate::settings::FileFormat::Audiobook) {
                            doc.path
                                .as_ref()
                                .map(|path| crate::is_local_cached_media_path(path))
                        } else {
                            None
                        }
                    })
                })
                .unwrap_or(false);
            !is_raiplay_live
                && (from_rss_doc || state.active_podcast_episode_url.is_some() || cached_media_doc)
        })
        .unwrap_or(false);
        let stream_direct_no_download = with_state(hwnd, |state| {
            state
                .active_podcast_episode_cache
                .as_ref()
                .map(|p| p.to_string_lossy())
                .map(|s| s.starts_with("http://") || s.starts_with("https://"))
                .unwrap_or(false)
        })
        .unwrap_or(false);
        let direct_stream_playback = with_state(hwnd, |state| {
            if let Some(player) = state.active_audiobook.as_ref() {
                let s = player.path.to_string_lossy();
                if s.starts_with("http://") || s.starts_with("https://") {
                    return true;
                }
            }
            state
                .docs
                .get(state.current)
                .and_then(|doc| {
                    if matches!(doc.format, crate::settings::FileFormat::Audiobook) {
                        doc.path.as_ref()
                    } else {
                        None
                    }
                })
                .map(|p| p.to_string_lossy())
                .map(|s| s.starts_with("http://") || s.starts_with("https://"))
                .unwrap_or(false)
        })
        .unwrap_or(false);
        let mpv_playback_active =
            with_state(hwnd, |state| state.active_mpv_session.is_some()).unwrap_or(false);
        let hls_m3u8_playback = with_state(hwnd, |state| {
            let active_player_url = state
                .active_audiobook
                .as_ref()
                .map(|player| player.path.to_string_lossy().to_ascii_lowercase());
            let active_doc_url = state
                .docs
                .get(state.current)
                .and_then(|doc| {
                    if matches!(doc.format, crate::settings::FileFormat::Audiobook) {
                        doc.path.as_ref()
                    } else {
                        None
                    }
                })
                .map(|path| path.to_string_lossy().to_ascii_lowercase());
            let active_episode_url = state
                .active_podcast_episode_url
                .as_ref()
                .map(|url| url.to_ascii_lowercase());
            let active_cache = state
                .active_podcast_episode_cache
                .as_ref()
                .map(|path| path.to_string_lossy().to_ascii_lowercase());

            [
                active_player_url,
                active_doc_url,
                active_episode_url,
                active_cache,
            ]
            .into_iter()
            .flatten()
            .any(|value| value.contains(".m3u8"))
        })
        .unwrap_or(false);
        let raiplay_live_stream_playback = with_state(hwnd, |state| {
            state.active_podcast_episode_from_rai == crate::RaiAudioOrigin::RaiPlay
                && !state.raiplay_live_audio_variants.is_empty()
        })
        .unwrap_or(false);
        if show {
            if existing.0 != 0 {
                remove_playback_menu(hmenu, existing);
                with_state(hwnd, |state| state.playback_menu = HMENU(0));
            }
            let playback_menu = CreateMenu().unwrap_or(HMENU(0));
            if playback_menu.0 == 0 {
                return;
            }
            let title = i18n::tr(language, "menu.playback");
            let play_pause = i18n::tr(language, "playback.play_pause");
            let stop = i18n::tr(language, "playback.stop");
            let seek_forward = i18n::tr(language, "playback.seek_forward");
            let seek_backward = i18n::tr(language, "playback.seek_backward");
            let seek_to_start = i18n::tr(language, "playback.seek_to_start");
            let seek_to_end = i18n::tr(language, "playback.seek_to_end");
            let shortcuts = with_state(hwnd, |state| state.settings.shortcuts.clone())
                .unwrap_or_else(ShortcutSettings::default);
            let track_prev = label_with_shortcut(
                &i18n::tr(language, "playback.track_prev"),
                shortcuts.media_prev,
            );
            let track_next = label_with_shortcut(
                &i18n::tr(language, "playback.track_next"),
                shortcuts.media_next,
            );
            let go_to_time = i18n::tr(language, "playback.go_to_time");
            let announce_time = i18n::tr(language, "playback.announce_time");
            let volume_up = i18n::tr(language, "playback.volume_up");
            let volume_down = i18n::tr(language, "playback.volume_down");
            let volume_reset = i18n::tr(language, "playback.volume_reset");
            let speed_up = i18n::tr(language, "playback.speed_up");
            let speed_down = i18n::tr(language, "playback.speed_down");
            let speed_reset = i18n::tr(language, "playback.speed_reset");
            let pitch_up = i18n::tr(language, "playback.pitch_up");
            let pitch_down = i18n::tr(language, "playback.pitch_down");
            let pitch_reset = i18n::tr(language, "playback.pitch_reset");
            let reset_menu_label = i18n::tr(language, "playback.reset_menu");
            let mute_toggle = i18n::tr(language, "playback.mute_toggle");
            let add_subtitles = i18n::tr(language, "playback.add_subtitles");
            let remove_subtitles = i18n::tr(language, "playback.remove_subtitles");
            let chapter_prev = label_with_shortcut(
                &i18n::tr(language, "playback.chapter_prev"),
                shortcuts.chapter_prev,
            );
            let chapter_next = label_with_shortcut(
                &i18n::tr(language, "playback.chapter_next"),
                shortcuts.chapter_next,
            );
            let chapter_list = i18n::tr(language, "playback.chapter_list");
            let download_episode = format!(
                "{}\tAlt+Shift+E",
                i18n::tr(language, "playback.download_episode")
            );
            let transcribe_current = i18n::tr(language, "playback.transcribe_current");
            let transcribe_current_folder = format!(
                "{}\tAlt+Shift+C",
                i18n::tr(language, "playback.transcribe_current_folder")
            );
            let has_chapters = with_state(hwnd, |state| !state.active_podcast_chapters.is_empty())
                .unwrap_or(false);
            let (audio_tracks, selected_track) = with_state(hwnd, |state| {
                (
                    state.available_audio_tracks.clone(),
                    state.selected_audio_track,
                )
            })
            .unwrap_or_default();

            append_menu_string(
                playback_menu,
                MF_STRING,
                IDM_PLAYBACK_PLAY_PAUSE,
                &play_pause,
            );
            append_menu_string(playback_menu, MF_STRING, IDM_PLAYBACK_STOP, &stop);
            let seek_flags = if direct_stream_playback {
                MF_STRING | MF_GRAYED
            } else {
                MF_STRING
            };
            append_menu_string(
                playback_menu,
                seek_flags,
                IDM_PLAYBACK_SEEK_FORWARD,
                &seek_forward,
            );
            append_menu_string(
                playback_menu,
                seek_flags,
                IDM_PLAYBACK_SEEK_BACKWARD,
                &seek_backward,
            );
            append_menu_string(
                playback_menu,
                seek_flags,
                IDM_PLAYBACK_SEEK_TO_START,
                &seek_to_start,
            );
            append_menu_string(
                playback_menu,
                seek_flags,
                IDM_PLAYBACK_SEEK_TO_END,
                &seek_to_end,
            );
            append_menu_string(
                playback_menu,
                MF_STRING,
                IDM_PLAYBACK_TRACK_PREV,
                &track_prev,
            );
            append_menu_string(
                playback_menu,
                MF_STRING,
                IDM_PLAYBACK_TRACK_NEXT,
                &track_next,
            );
            if has_chapters {
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_CHAPTER_PREV,
                    &chapter_prev,
                );
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_CHAPTER_NEXT,
                    &chapter_next,
                );
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_CHAPTER_LIST,
                    &chapter_list,
                );
            }
            if show_download && !stream_direct_no_download && !raiplay_live_stream_playback {
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_DOWNLOAD_EPISODE,
                    &download_episode,
                );
            }
            if !hls_m3u8_playback {
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_TRANSCRIBE_CURRENT,
                    &transcribe_current,
                );
            }
            if !hls_m3u8_playback && !direct_stream_playback {
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_TRANSCRIBE_CURRENT_FOLDER,
                    &transcribe_current_folder,
                );
            }
            append_menu_string(
                playback_menu,
                if direct_stream_playback {
                    MF_STRING | MF_GRAYED
                } else {
                    MF_STRING
                },
                IDM_PLAYBACK_GO_TO_TIME,
                &go_to_time,
            );
            append_menu_string(
                playback_menu,
                MF_STRING,
                IDM_PLAYBACK_ANNOUNCE_TIME,
                &announce_time,
            );
            if !raiplay_live_stream_playback {
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_ADD_SUBTITLES,
                    &add_subtitles,
                );
                append_menu_string(
                    playback_menu,
                    MF_STRING,
                    IDM_PLAYBACK_REMOVE_SUBTITLES,
                    &remove_subtitles,
                );
            }
            // Audio tracks submenu
            if !audio_tracks.is_empty() {
                let audio_track_label = i18n::tr(language, "playback.audio_track");
                let audio_tracks_menu = CreateMenu().unwrap_or(HMENU(0));
                if audio_tracks_menu.0 != 0 {
                    for (i, track) in audio_tracks.iter().enumerate() {
                        if i >= IDM_PLAYBACK_AUDIO_TRACK_MAX {
                            break;
                        }
                        let mut label = format!("{}", track.index);
                        if let Some(ref lang) = track.language {
                            label.push_str(&format!(" [{}]", lang));
                        }
                        if let Some(ref title) = track.title {
                            label.push_str(&format!(" - {}", title));
                        }
                        if track.channels > 0 && track.sample_rate > 0 && !track.codec.is_empty() {
                            label.push_str(&format!(
                                " ({}, {}ch, {}Hz)",
                                track.codec, track.channels, track.sample_rate
                            ));
                        }
                        let is_selected = selected_track
                            .map(|s| s == track.index)
                            .unwrap_or(track.is_default);
                        let flags = if is_selected {
                            MF_STRING | MF_CHECKED
                        } else {
                            MF_STRING | MF_UNCHECKED
                        };
                        append_menu_string(
                            audio_tracks_menu,
                            flags,
                            IDM_PLAYBACK_AUDIO_TRACK_BASE + i,
                            &label,
                        );
                    }
                    append_menu_string(
                        playback_menu,
                        MF_POPUP,
                        audio_tracks_menu.0 as usize,
                        &audio_track_label,
                    );
                }
            }
            append_menu_string(playback_menu, MF_STRING, IDM_PLAYBACK_VOLUME_UP, &volume_up);
            append_menu_string(
                playback_menu,
                MF_STRING,
                IDM_PLAYBACK_VOLUME_DOWN,
                &volume_down,
            );
            let rate_flags = if direct_stream_playback && !mpv_playback_active {
                MF_STRING | MF_GRAYED
            } else {
                MF_STRING
            };
            append_menu_string(playback_menu, rate_flags, IDM_PLAYBACK_SPEED_UP, &speed_up);
            append_menu_string(
                playback_menu,
                rate_flags,
                IDM_PLAYBACK_SPEED_DOWN,
                &speed_down,
            );
            append_menu_string(playback_menu, rate_flags, IDM_PLAYBACK_PITCH_UP, &pitch_up);
            append_menu_string(
                playback_menu,
                rate_flags,
                IDM_PLAYBACK_PITCH_DOWN,
                &pitch_down,
            );
            let reset_menu = CreateMenu().unwrap_or(HMENU(0));
            if reset_menu.0 != 0 {
                append_menu_string(
                    reset_menu,
                    MF_STRING,
                    IDM_PLAYBACK_VOLUME_RESET,
                    &volume_reset,
                );
                append_menu_string(
                    reset_menu,
                    if direct_stream_playback && !mpv_playback_active {
                        MF_STRING | MF_GRAYED
                    } else {
                        MF_STRING
                    },
                    IDM_PLAYBACK_SPEED_RESET,
                    &speed_reset,
                );
                append_menu_string(
                    reset_menu,
                    if direct_stream_playback && !mpv_playback_active {
                        MF_STRING | MF_GRAYED
                    } else {
                        MF_STRING
                    },
                    IDM_PLAYBACK_PITCH_RESET,
                    &pitch_reset,
                );
                append_menu_string(
                    playback_menu,
                    MF_POPUP,
                    reset_menu.0 as usize,
                    &reset_menu_label,
                );
            }
            append_menu_string(
                playback_menu,
                MF_STRING,
                IDM_PLAYBACK_MUTE_TOGGLE,
                &mute_toggle,
            );
            let wide = to_wide(&title);
            crate::log_if_err!(InsertMenuW(
                hmenu,
                0,
                MF_BYPOSITION | MF_POPUP,
                playback_menu.0 as usize,
                PCWSTR(wide.as_ptr()),
            ));
            with_state(hwnd, |state| state.playback_menu = playback_menu);
            crate::log_if_err!(DrawMenuBar(hwnd));
        } else if existing.0 != 0 {
            remove_playback_menu(hmenu, existing);
            with_state(hwnd, |state| state.playback_menu = HMENU(0));
            crate::log_if_err!(DrawMenuBar(hwnd));
        }
    }
}

fn remove_playback_menu(hmenu: HMENU, playback_menu: HMENU) {
    unsafe {
        if hmenu.0 == 0 || playback_menu.0 == 0 {
            return;
        }
        let mut detached = false;
        let count = GetMenuItemCount(hmenu);
        if count > 0 {
            for index in 0..count {
                if GetSubMenu(hmenu, index) == playback_menu {
                    crate::log_if_err!(DeleteMenu(hmenu, index as u32, MF_BYPOSITION));
                    detached = true;
                    break;
                }
            }
        }
        if detached {
            if IsMenu(playback_menu).as_bool() {
                crate::log_if_err!(DestroyMenu(playback_menu));
            } else {
                crate::log_debug("remove_playback_menu: playback menu handle already invalid");
            }
        }
    }
}

pub fn create_menus(hwnd: HWND, language: Language) -> (HMENU, HMENU) {
    unsafe {
        let hmenu = CreateMenu().unwrap_or(HMENU(0));
        let file_menu = CreateMenu().unwrap_or(HMENU(0));
        let recent_menu = CreateMenu().unwrap_or(HMENU(0));
        let edit_menu = CreateMenu().unwrap_or(HMENU(0));
        let view_menu = CreateMenu().unwrap_or(HMENU(0));
        let view_font_menu = CreateMenu().unwrap_or(HMENU(0));
        let view_color_menu = CreateMenu().unwrap_or(HMENU(0));
        let view_size_menu = CreateMenu().unwrap_or(HMENU(0));
        let insert_menu = CreateMenu().unwrap_or(HMENU(0));
        let window_menu = CreateMenu().unwrap_or(HMENU(0));
        let voice_audio_menu = CreateMenu().unwrap_or(HMENU(0));
        let tools_menu = CreateMenu().unwrap_or(HMENU(0));
        let help_menu = CreateMenu().unwrap_or(HMENU(0));

        let mut labels = menu_labels(language);
        let shortcuts = with_state(hwnd, |state| state.settings.shortcuts.clone())
            .unwrap_or_else(ShortcutSettings::default);
        labels.file_read_start = label_with_shortcut(&labels.file_read_start, shortcuts.read_start);
        labels.file_read_previous_sentence = label_with_shortcut(
            &labels.file_read_previous_sentence,
            shortcuts.read_previous_sentence,
        );
        labels.file_read_next_sentence = label_with_shortcut(
            &labels.file_read_next_sentence,
            shortcuts.read_next_sentence,
        );
        labels.file_read_pause =
            label_with_shortcut(&labels.file_read_pause, shortcuts.read_pause_resume);
        labels.file_read_stop = label_with_shortcut(&labels.file_read_stop, shortcuts.read_stop);
        labels.file_execute = label_with_shortcut(&labels.file_execute, shortcuts.execute_file);
        labels.file_audiobook = label_with_shortcut(&labels.file_audiobook, shortcuts.audiobook);
        labels.file_batch_audiobooks =
            label_with_shortcut(&labels.file_batch_audiobooks, shortcuts.batch_audiobooks);
        labels.file_podcast = label_with_shortcut(&labels.file_podcast, shortcuts.record_podcast);
        labels.file_convert_audio =
            label_with_shortcut(&labels.file_convert_audio, shortcuts.convert_audio);
        labels.menu_rss = label_with_shortcut(&labels.menu_rss, shortcuts.open_rss);
        labels.menu_podcasts = label_with_shortcut(&labels.menu_podcasts, shortcuts.open_podcasts);
        labels.menu_dictionary =
            label_with_shortcut(&labels.menu_dictionary, shortcuts.open_dictionary);
        labels.menu_options = label_with_shortcut(&labels.menu_options, shortcuts.open_options);
        labels.menu_prompt = label_with_shortcut(&labels.menu_prompt, shortcuts.open_terminal);
        labels.menu_wikipedia_import =
            label_with_shortcut(&labels.menu_wikipedia_import, shortcuts.import_wikipedia);
        labels.menu_import_youtube =
            label_with_shortcut(&labels.menu_import_youtube, shortcuts.import_youtube);
        labels.edit_find = label_with_shortcut(&labels.edit_find, shortcuts.find);
        labels.edit_quote_lines =
            label_with_shortcut(&labels.edit_quote_lines, shortcuts.quote_lines);
        labels.edit_unquote_lines =
            label_with_shortcut(&labels.edit_unquote_lines, shortcuts.unquote_lines);

        append_menu_string(file_menu, MF_STRING, IDM_FILE_NEW, &labels.file_new);
        append_menu_string(file_menu, MF_STRING, IDM_FILE_OPEN, &labels.file_open);
        append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE, &labels.file_save);
        append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE_AS, &labels.file_save_as);
        append_menu_string(
            file_menu,
            MF_STRING,
            IDM_FILE_SAVE_ALL,
            &labels.file_save_all,
        );
        crate::log_if_err!(AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(
            file_menu,
            MF_POPUP,
            recent_menu.0 as usize,
            &labels.file_recent,
        );
        crate::log_if_err!(AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(file_menu, MF_STRING, IDM_FILE_EXECUTE, &labels.file_execute);
        crate::log_if_err!(AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(file_menu, MF_STRING, IDM_FILE_EXIT, &labels.file_exit);
        append_menu_string(hmenu, MF_POPUP, file_menu.0 as usize, &labels.menu_file);

        append_menu_string(edit_menu, MF_STRING, IDM_EDIT_UNDO, &labels.edit_undo);
        crate::log_if_err!(AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(edit_menu, MF_STRING, IDM_EDIT_CUT, &labels.edit_cut);
        append_menu_string(edit_menu, MF_STRING, IDM_EDIT_COPY, &labels.edit_copy);
        append_menu_string(edit_menu, MF_STRING, IDM_EDIT_PASTE, &labels.edit_paste);
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_SELECT_ALL,
            &labels.edit_select_all,
        );
        crate::log_if_err!(AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(edit_menu, MF_STRING, IDM_EDIT_FIND, &labels.edit_find);
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_FIND_IN_FILES,
            &labels.edit_find_in_files,
        );
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_FIND_NEXT,
            &labels.edit_find_next,
        );
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_FIND_PREVIOUS,
            &labels.edit_find_previous,
        );
        append_menu_string(edit_menu, MF_STRING, IDM_EDIT_REPLACE, &labels.edit_replace);
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_GO_TO_LINE,
            &labels.edit_goto_line,
        );
        crate::log_if_err!(AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_PREV_SPELLING_ERROR,
            &labels.edit_prev_spelling_error,
        );
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_NEXT_SPELLING_ERROR,
            &labels.edit_next_spelling_error,
        );
        crate::log_if_err!(AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        let text_menu = CreateMenu().unwrap_or(HMENU(0));
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_AUTO_FORMAT_TTS,
            &labels.edit_auto_format_tts,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_STRIP_MARKDOWN,
            &labels.edit_strip_markdown,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_NORMALIZE_WHITESPACE,
            &labels.edit_normalize_whitespace,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_HARD_LINE_BREAK,
            &labels.edit_hard_line_break,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_JOIN_LINES,
            &labels.edit_join_lines,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_CLEAN_EOL_HYPHENS,
            &labels.edit_clean_eol_hyphens,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_ORDER_ITEMS,
            &labels.edit_order_items,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_KEEP_UNIQUE_ITEMS,
            &labels.edit_keep_unique_items,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_REVERSE_ITEMS,
            &labels.edit_reverse_items,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_QUOTE_LINES,
            &labels.edit_quote_lines,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_UNQUOTE_LINES,
            &labels.edit_unquote_lines,
        );
        append_menu_string(text_menu, MF_STRING, IDM_EDIT_INDENT, &labels.edit_indent);
        append_menu_string(text_menu, MF_STRING, IDM_EDIT_OUTDENT, &labels.edit_outdent);
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_REMOVE_DUPLICATE_LINES,
            &labels.edit_remove_duplicate_lines,
        );
        append_menu_string(
            text_menu,
            MF_STRING,
            IDM_EDIT_REMOVE_DUPLICATE_CONSECUTIVE_LINES,
            &labels.edit_remove_duplicate_consecutive_lines,
        );
        append_menu_string(
            edit_menu,
            MF_POPUP,
            text_menu.0 as usize,
            &labels.edit_text_menu,
        );
        crate::log_if_err!(AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(
            edit_menu,
            MF_STRING,
            IDM_EDIT_TEXT_STATS,
            &labels.edit_text_stats,
        );
        append_menu_string(hmenu, MF_POPUP, edit_menu.0 as usize, &labels.menu_edit);

        append_menu_string(
            view_menu,
            MF_STRING,
            IDM_VIEW_SHOW_VOICES,
            &labels.view_show_voices,
        );
        append_menu_string(
            view_menu,
            MF_STRING,
            IDM_VIEW_SHOW_FAVORITES,
            &labels.view_show_favorites,
        );
        append_menu_string(
            view_menu,
            MF_STRING,
            IDM_VIEW_READ_ONLY,
            &labels.view_read_only,
        );
        append_menu_string(
            view_menu,
            MF_STRING,
            IDM_VIEW_WORD_WRAP,
            &labels.view_word_wrap,
        );
        append_menu_string(
            view_menu,
            MF_STRING,
            IDM_VIEW_SHOW_VIDEO_DURING_PLAYBACK,
            &labels.view_show_video_during_playback,
        );
        crate::log_if_err!(AppendMenuW(view_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(view_font_menu, MF_STRING, IDM_VIEW_FONT_ARIAL, "Arial");
        append_menu_string(view_font_menu, MF_STRING, IDM_VIEW_FONT_CALIBRI, "Calibri");
        append_menu_string(
            view_font_menu,
            MF_STRING,
            IDM_VIEW_FONT_CONSOLAS,
            "Consolas",
        );
        append_menu_string(
            view_font_menu,
            MF_STRING,
            IDM_VIEW_FONT_SEGOE_UI,
            "Segoe UI",
        );
        append_menu_string(view_font_menu, MF_STRING, IDM_VIEW_FONT_TAHOMA, "Tahoma");
        append_menu_string(view_font_menu, MF_STRING, IDM_VIEW_FONT_VERDANA, "Verdana");
        append_menu_string(
            view_font_menu,
            MF_STRING,
            IDM_VIEW_FONT_TIMES_NEW_ROMAN,
            "Times New Roman",
        );
        append_menu_string(view_font_menu, MF_STRING, IDM_VIEW_FONT_GEORGIA, "Georgia");
        append_menu_string(
            view_menu,
            MF_POPUP,
            view_font_menu.0 as usize,
            &labels.view_font,
        );
        crate::log_if_err!(AppendMenuW(view_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_BLACK,
            &labels.view_text_color_black,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_DARK_BLUE,
            &labels.view_text_color_dark_blue,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_DARK_GREEN,
            &labels.view_text_color_dark_green,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_DARK_BROWN,
            &labels.view_text_color_dark_brown,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_DARK_GRAY,
            &labels.view_text_color_dark_gray,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_LIGHT_BLUE,
            &labels.view_text_color_light_blue,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_LIGHT_GREEN,
            &labels.view_text_color_light_green,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_LIGHT_BROWN,
            &labels.view_text_color_light_brown,
        );
        append_menu_string(
            view_color_menu,
            MF_STRING,
            IDM_VIEW_TEXT_COLOR_LIGHT_GRAY,
            &labels.view_text_color_light_gray,
        );
        append_menu_string(
            view_size_menu,
            MF_STRING,
            IDM_VIEW_TEXT_SIZE_SMALL,
            &labels.view_text_size_small,
        );
        append_menu_string(
            view_size_menu,
            MF_STRING,
            IDM_VIEW_TEXT_SIZE_NORMAL,
            &labels.view_text_size_normal,
        );
        append_menu_string(
            view_size_menu,
            MF_STRING,
            IDM_VIEW_TEXT_SIZE_LARGE,
            &labels.view_text_size_large,
        );
        append_menu_string(
            view_size_menu,
            MF_STRING,
            IDM_VIEW_TEXT_SIZE_XLARGE,
            &labels.view_text_size_xlarge,
        );
        append_menu_string(
            view_size_menu,
            MF_STRING,
            IDM_VIEW_TEXT_SIZE_XXLARGE,
            &labels.view_text_size_xxlarge,
        );
        append_menu_string(
            view_menu,
            MF_POPUP,
            view_color_menu.0 as usize,
            &labels.view_text_color,
        );
        append_menu_string(
            view_menu,
            MF_POPUP,
            view_size_menu.0 as usize,
            &labels.view_text_size,
        );
        append_menu_string(hmenu, MF_POPUP, view_menu.0 as usize, &labels.menu_view);

        append_menu_string(
            insert_menu,
            MF_STRING,
            IDM_INSERT_BOOKMARK,
            &labels.insert_bookmark,
        );
        append_menu_string(
            insert_menu,
            MF_STRING,
            IDM_AUTOMATIC_BOOKMARK,
            &labels.automatic_bookmark,
        );
        append_menu_string(
            insert_menu,
            MF_STRING,
            IDM_GOTO_NEXT_BOOKMARK,
            &labels.insert_goto_next_bookmark,
        );
        append_menu_string(
            insert_menu,
            MF_STRING,
            IDM_GOTO_PREV_BOOKMARK,
            &labels.insert_goto_prev_bookmark,
        );
        append_menu_string(
            insert_menu,
            MF_STRING,
            IDM_INSERT_CLEAR_BOOKMARKS,
            &labels.insert_clear_bookmarks,
        );
        append_menu_string(
            insert_menu,
            MF_STRING,
            IDM_MANAGE_BOOKMARKS,
            &labels.manage_bookmarks,
        );
        append_menu_string(hmenu, MF_POPUP, insert_menu.0 as usize, &labels.menu_insert);

        append_menu_string(window_menu, MF_STRING, IDM_FILE_CLOSE, &labels.file_close);
        append_menu_string(
            window_menu,
            MF_STRING,
            IDM_WINDOW_OPEN_DOCUMENTS,
            &labels.window_open_documents,
        );
        crate::log_if_err!(AppendMenuW(window_menu, MF_SEPARATOR, 0, PCWSTR::null()));
        append_menu_string(
            window_menu,
            MF_STRING,
            IDM_FILE_CLOSE_OTHERS,
            &format!("{}\tCtrl+Shift+W", labels.window_close_others),
        );
        append_menu_string(
            window_menu,
            MF_STRING,
            IDM_WINDOW_CLOSE_ALL,
            &labels.window_close_all,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_READ_START,
            &labels.file_read_start,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_READ_PREVIOUS_SENTENCE,
            &labels.file_read_previous_sentence,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_READ_NEXT_SENTENCE,
            &labels.file_read_next_sentence,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_READ_PAUSE,
            &labels.file_read_pause,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_READ_STOP,
            &labels.file_read_stop,
        );
        crate::log_if_err!(AppendMenuW(
            voice_audio_menu,
            MF_SEPARATOR,
            0,
            PCWSTR::null()
        ));
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_AUDIOBOOK,
            &labels.file_audiobook,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_BATCH_AUDIOBOOK,
            &labels.file_batch_audiobooks,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_PODCAST,
            &labels.file_podcast,
        );
        append_menu_string(
            voice_audio_menu,
            MF_STRING,
            IDM_FILE_CONVERT_AUDIO,
            &labels.file_convert_audio,
        );
        append_menu_string(
            hmenu,
            MF_POPUP,
            voice_audio_menu.0 as usize,
            &labels.menu_voice_audio,
        );

        append_menu_string(tools_menu, MF_STRING, IDM_TOOLS_PROMPT, &labels.menu_prompt);
        append_menu_string(tools_menu, MF_STRING, IDM_TOOLS_RSS, &labels.menu_rss);
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_PODCASTS,
            &labels.menu_podcasts,
        );
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_DICTIONARY,
            &labels.menu_dictionary,
        );
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_DICTIONARY_LOOKUP,
            &labels.menu_dictionary_lookup,
        );
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_WIKIPEDIA_IMPORT,
            &labels.menu_wikipedia_import,
        );
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_IMPORT_YOUTUBE,
            &labels.menu_import_youtube,
        );
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_STREAM_AUDIO,
            &labels.menu_stream_audio,
        );
        if language == Language::Italian {
            append_menu_string(
                tools_menu,
                MF_STRING,
                IDM_TOOLS_RAI_AUDIODESCRIZIONI,
                &labels.menu_rai_audiodescrizioni,
            );
            append_menu_string(
                tools_menu,
                MF_STRING,
                IDM_TOOLS_RAIPLAY,
                &labels.menu_raiplay,
            );
            append_menu_string(
                tools_menu,
                MF_STRING,
                IDM_TOOLS_RAIPLAYSOUND,
                &labels.menu_raiplaysound,
            );
            append_menu_string(
                tools_menu,
                MF_STRING,
                IDM_TOOLS_ITALIAONLINE,
                &labels.menu_italiaonline,
            );
            append_menu_string(tools_menu, MF_STRING, IDM_TOOLS_RADIO, &labels.menu_radio);
            append_menu_string(
                tools_menu,
                MF_STRING,
                IDM_TOOLS_BDCIECHI,
                &labels.menu_bdciechi,
            );
        }
        append_menu_string(
            tools_menu,
            MF_STRING,
            IDM_TOOLS_OPTIONS,
            &labels.menu_options,
        );
        append_menu_string(hmenu, MF_POPUP, tools_menu.0 as usize, &labels.menu_tools);
        append_menu_string(hmenu, MF_POPUP, window_menu.0 as usize, &labels.menu_window);

        append_menu_string(help_menu, MF_STRING, IDM_HELP_GUIDE, &labels.help_guide);
        append_menu_string(
            help_menu,
            MF_STRING,
            IDM_HELP_CHANGELOG,
            &labels.help_changelog,
        );
        append_menu_string(
            help_menu,
            MF_STRING,
            IDM_HELP_FEEDBACK,
            &labels.help_feedback,
        );
        append_menu_string(
            help_menu,
            MF_STRING,
            IDM_HELP_DONATIONS,
            &labels.help_donations,
        );
        append_menu_string(
            help_menu,
            MF_STRING,
            IDM_HELP_CHECK_UPDATES,
            &labels.help_check_updates,
        );
        append_menu_string(
            help_menu,
            MF_STRING,
            IDM_HELP_EXPORT_DIAGNOSTICS,
            &labels.help_export_diagnostics,
        );
        append_menu_string(help_menu, MF_STRING, IDM_HELP_ABOUT, &labels.help_about);
        append_menu_string(hmenu, MF_POPUP, help_menu.0 as usize, &labels.menu_help);

        crate::log_if_err!(SetMenu(hwnd, hmenu));
        (hmenu, recent_menu)
    }
}

pub fn update_recent_menu(hwnd: HWND, hmenu_recent: HMENU) {
    let count = unsafe { GetMenuItemCount(hmenu_recent) };
    if count > 0 {
        for _ in 0..count {
            crate::log_if_err!(unsafe { DeleteMenu(hmenu_recent, 0, MF_BYPOSITION) });
        }
    }

    let (files, language): (Vec<std::path::PathBuf>, Language) = {
        with_state(hwnd, |state| {
            (state.recent_files.clone(), state.settings.language)
        })
    }
    .unwrap_or_default();
    if files.is_empty() {
        let labels = menu_labels(language);
        append_menu_string(hmenu_recent, MF_STRING | MF_GRAYED, 0, &labels.recent_empty);
    } else {
        let labels = menu_labels(language);
        for (i, path) in files.iter().enumerate() {
            let label = format!("&{} {}", i + 1, abbreviate_recent_label(path));
            let wide = to_wide(&label);
            crate::log_if_err!(unsafe {
                AppendMenuW(
                    hmenu_recent,
                    MF_STRING,
                    IDM_FILE_RECENT_BASE + i,
                    PCWSTR(wide.as_ptr()),
                )
            });
        }
        crate::log_if_err!(unsafe { AppendMenuW(hmenu_recent, MF_SEPARATOR, 0, PCWSTR::null()) });
        append_menu_string(
            hmenu_recent,
            MF_STRING,
            IDM_FILE_RECENT_CLEAR,
            &labels.recent_clear,
        );
    }
    crate::log_if_err!(unsafe { DrawMenuBar(hwnd) });
}

pub fn abbreviate_recent_label(path: &Path) -> String {
    let filename = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
    let parent = path.parent().and_then(|p| p.to_str()).unwrap_or("");
    if parent.is_empty() {
        return filename.to_string();
    }
    let mut suffix = parent.to_string();
    if suffix.len() > 24 {
        let start = suffix.len().saturating_sub(24);
        // Advance to the next char boundary so we never slice inside a multi-byte character.
        let start = suffix.ceil_char_boundary(start);
        suffix = format!("...{}", &suffix[start..]);
    }
    format!("{filename} - {suffix}")
}

pub fn append_menu_string(menu: HMENU, flags: MENU_ITEM_FLAGS, id: usize, text: &str) {
    let wide = to_wide(text);
    crate::log_if_err!(unsafe { AppendMenuW(menu, flags, id, PCWSTR(wide.as_ptr())) });
}
