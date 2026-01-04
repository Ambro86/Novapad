#![allow(clippy::let_unit_value)]
use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use std::path::Path;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreateMenu, DeleteMenu, DrawMenuBar, GetMenuItemCount, HMENU, MENU_ITEM_FLAGS,
    MF_BYPOSITION, MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING, SetMenu,
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
pub const IDM_FILE_AUDIOBOOK: usize = 1011;
pub const IDM_FILE_PODCAST: usize = 1012;
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
pub const IDM_INSERT_BOOKMARK: usize = 2101;
pub const IDM_MANAGE_BOOKMARKS: usize = 2102;
pub const IDM_NEXT_TAB: usize = 3001;
pub const IDM_VIEW_SHOW_VOICES: usize = 6101;
pub const IDM_VIEW_SHOW_FAVORITES: usize = 6102;
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
pub const IDM_TOOLS_OPTIONS: usize = 5001;
pub const IDM_TOOLS_DICTIONARY: usize = 5002;
pub const IDM_TOOLS_IMPORT_YOUTUBE: usize = 5003;
pub const IDM_TOOLS_PROMPT: usize = 5004;
pub const IDM_HELP_GUIDE: usize = 7001;
pub const IDM_HELP_ABOUT: usize = 7002;
pub const IDM_HELP_CHECK_UPDATES: usize = 7003;
pub const IDM_HELP_CHANGELOG: usize = 7004;
pub const MAX_RECENT: usize = 5;

pub struct MenuLabels {
    pub menu_file: String,
    pub menu_edit: String,
    pub menu_view: String,
    pub menu_insert: String,
    pub menu_tools: String,
    pub menu_help: String,
    pub menu_options: String,
    pub menu_dictionary: String,
    pub menu_import_youtube: String,
    pub menu_prompt: String,
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
    pub file_new: String,
    pub file_open: String,
    pub file_save: String,
    pub file_save_as: String,
    pub file_save_all: String,
    pub file_close: String,
    pub file_recent: String,
    pub file_read_start: String,
    pub file_read_pause: String,
    pub file_read_stop: String,
    pub file_audiobook: String,
    pub file_podcast: String,
    pub file_exit: String,
    pub edit_undo: String,
    pub edit_cut: String,
    pub edit_copy: String,
    pub edit_paste: String,
    pub edit_select_all: String,
    pub edit_find: String,
    pub edit_find_next: String,
    pub edit_replace: String,
    pub edit_find_in_files: String,
    pub edit_strip_markdown: String,
    pub edit_normalize_whitespace: String,
    pub edit_hard_line_break: String,
    pub edit_order_items: String,
    pub edit_keep_unique_items: String,
    pub edit_reverse_items: String,
    pub edit_quote_lines: String,
    pub edit_unquote_lines: String,
    pub edit_text_stats: String,
    pub edit_join_lines: String,
    pub insert_bookmark: String,
    pub manage_bookmarks: String,
    pub help_guide: String,
    pub help_changelog: String,
    pub help_check_updates: String,
    pub help_about: String,
    pub recent_empty: String,
}

pub fn menu_labels(language: Language) -> MenuLabels {
    MenuLabels {
        menu_file: i18n::tr(language, "menu.file"),
        menu_edit: i18n::tr(language, "menu.edit"),
        menu_view: i18n::tr(language, "menu.view"),
        menu_insert: i18n::tr(language, "menu.insert"),
        menu_tools: i18n::tr(language, "menu.tools"),
        menu_help: i18n::tr(language, "menu.help"),
        menu_options: i18n::tr(language, "menu.options"),
        menu_dictionary: i18n::tr(language, "menu.dictionary"),
        menu_import_youtube: i18n::tr(language, "menu.import_youtube"),
        menu_prompt: i18n::tr(language, "menu.prompt"),
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
        file_new: i18n::tr(language, "file.new"),
        file_open: i18n::tr(language, "file.open"),
        file_save: i18n::tr(language, "file.save"),
        file_save_as: i18n::tr(language, "file.save_as"),
        file_save_all: i18n::tr(language, "file.save_all"),
        file_close: i18n::tr(language, "file.close"),
        file_recent: i18n::tr(language, "file.recent"),
        file_read_start: i18n::tr(language, "file.read_start"),
        file_read_pause: i18n::tr(language, "file.read_pause"),
        file_read_stop: i18n::tr(language, "file.read_stop"),
        file_audiobook: i18n::tr(language, "file.audiobook"),
        file_podcast: i18n::tr(language, "file.podcast"),
        file_exit: i18n::tr(language, "file.exit"),
        edit_undo: i18n::tr(language, "edit.undo"),
        edit_cut: i18n::tr(language, "edit.cut"),
        edit_copy: i18n::tr(language, "edit.copy"),
        edit_paste: i18n::tr(language, "edit.paste"),
        edit_select_all: i18n::tr(language, "edit.select_all"),
        edit_find: i18n::tr(language, "edit.find"),
        edit_find_next: i18n::tr(language, "edit.find_next"),
        edit_replace: i18n::tr(language, "edit.replace"),
        edit_find_in_files: i18n::tr(language, "edit.find_in_files"),
        edit_strip_markdown: i18n::tr(language, "edit.strip_markdown"),
        edit_normalize_whitespace: i18n::tr(language, "edit.normalize_whitespace"),
        edit_hard_line_break: i18n::tr(language, "edit.hard_line_break"),
        edit_order_items: i18n::tr(language, "edit.order_items"),
        edit_keep_unique_items: i18n::tr(language, "edit.keep_unique_items"),
        edit_reverse_items: i18n::tr(language, "edit.reverse_items"),
        edit_quote_lines: i18n::tr(language, "edit.quote_lines"),
        edit_unquote_lines: i18n::tr(language, "edit.unquote_lines"),
        edit_text_stats: i18n::tr(language, "edit.text_stats"),
        edit_join_lines: i18n::tr(language, "edit.join_lines"),
        insert_bookmark: i18n::tr(language, "insert.bookmark"),
        manage_bookmarks: i18n::tr(language, "insert.manage_bookmarks"),
        help_guide: i18n::tr(language, "help.guide"),
        help_changelog: i18n::tr(language, "help.changelog"),
        help_check_updates: i18n::tr(language, "help.check_updates"),
        help_about: i18n::tr(language, "help.about"),
        recent_empty: i18n::tr(language, "recent.empty"),
    }
}

pub unsafe fn create_menus(hwnd: HWND, language: Language) -> (HMENU, HMENU) {
    let hmenu = CreateMenu().unwrap_or(HMENU(0));
    let file_menu = CreateMenu().unwrap_or(HMENU(0));
    let recent_menu = CreateMenu().unwrap_or(HMENU(0));
    let edit_menu = CreateMenu().unwrap_or(HMENU(0));
    let view_menu = CreateMenu().unwrap_or(HMENU(0));
    let view_color_menu = CreateMenu().unwrap_or(HMENU(0));
    let view_size_menu = CreateMenu().unwrap_or(HMENU(0));
    let insert_menu = CreateMenu().unwrap_or(HMENU(0));
    let tools_menu = CreateMenu().unwrap_or(HMENU(0));
    let help_menu = CreateMenu().unwrap_or(HMENU(0));

    let labels = menu_labels(language);

    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_NEW, &labels.file_new);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_OPEN, &labels.file_open);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE, &labels.file_save);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE_AS, &labels.file_save_as);
    let _ = append_menu_string(
        file_menu,
        MF_STRING,
        IDM_FILE_SAVE_ALL,
        &labels.file_save_all,
    );
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_CLOSE, &labels.file_close);
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(
        file_menu,
        MF_POPUP,
        recent_menu.0 as usize,
        &labels.file_recent,
    );
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(
        file_menu,
        MF_STRING,
        IDM_FILE_READ_START,
        &labels.file_read_start,
    );
    let _ = append_menu_string(
        file_menu,
        MF_STRING,
        IDM_FILE_READ_PAUSE,
        &labels.file_read_pause,
    );
    let _ = append_menu_string(
        file_menu,
        MF_STRING,
        IDM_FILE_READ_STOP,
        &labels.file_read_stop,
    );
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(
        file_menu,
        MF_STRING,
        IDM_FILE_AUDIOBOOK,
        &labels.file_audiobook,
    );
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_PODCAST, &labels.file_podcast);
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_EXIT, &labels.file_exit);
    let _ = append_menu_string(hmenu, MF_POPUP, file_menu.0 as usize, &labels.menu_file);

    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_UNDO, &labels.edit_undo);
    let _ = AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_CUT, &labels.edit_cut);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_COPY, &labels.edit_copy);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_PASTE, &labels.edit_paste);
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_SELECT_ALL,
        &labels.edit_select_all,
    );
    let _ = AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_FIND, &labels.edit_find);
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_FIND_IN_FILES,
        &labels.edit_find_in_files,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_FIND_NEXT,
        &labels.edit_find_next,
    );
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_REPLACE, &labels.edit_replace);
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_STRIP_MARKDOWN,
        &labels.edit_strip_markdown,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_NORMALIZE_WHITESPACE,
        &labels.edit_normalize_whitespace,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_HARD_LINE_BREAK,
        &labels.edit_hard_line_break,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_ORDER_ITEMS,
        &labels.edit_order_items,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_KEEP_UNIQUE_ITEMS,
        &labels.edit_keep_unique_items,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_REVERSE_ITEMS,
        &labels.edit_reverse_items,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_QUOTE_LINES,
        &labels.edit_quote_lines,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_UNQUOTE_LINES,
        &labels.edit_unquote_lines,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_JOIN_LINES,
        &labels.edit_join_lines,
    );
    let _ = append_menu_string(
        edit_menu,
        MF_STRING,
        IDM_EDIT_TEXT_STATS,
        &labels.edit_text_stats,
    );
    let _ = append_menu_string(hmenu, MF_POPUP, edit_menu.0 as usize, &labels.menu_edit);

    let _ = append_menu_string(
        view_menu,
        MF_STRING,
        IDM_VIEW_SHOW_VOICES,
        &labels.view_show_voices,
    );
    let _ = append_menu_string(
        view_menu,
        MF_STRING,
        IDM_VIEW_SHOW_FAVORITES,
        &labels.view_show_favorites,
    );
    let _ = AppendMenuW(view_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_BLACK,
        &labels.view_text_color_black,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_DARK_BLUE,
        &labels.view_text_color_dark_blue,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_DARK_GREEN,
        &labels.view_text_color_dark_green,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_DARK_BROWN,
        &labels.view_text_color_dark_brown,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_DARK_GRAY,
        &labels.view_text_color_dark_gray,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_LIGHT_BLUE,
        &labels.view_text_color_light_blue,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_LIGHT_GREEN,
        &labels.view_text_color_light_green,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_LIGHT_BROWN,
        &labels.view_text_color_light_brown,
    );
    let _ = append_menu_string(
        view_color_menu,
        MF_STRING,
        IDM_VIEW_TEXT_COLOR_LIGHT_GRAY,
        &labels.view_text_color_light_gray,
    );
    let _ = append_menu_string(
        view_size_menu,
        MF_STRING,
        IDM_VIEW_TEXT_SIZE_SMALL,
        &labels.view_text_size_small,
    );
    let _ = append_menu_string(
        view_size_menu,
        MF_STRING,
        IDM_VIEW_TEXT_SIZE_NORMAL,
        &labels.view_text_size_normal,
    );
    let _ = append_menu_string(
        view_size_menu,
        MF_STRING,
        IDM_VIEW_TEXT_SIZE_LARGE,
        &labels.view_text_size_large,
    );
    let _ = append_menu_string(
        view_size_menu,
        MF_STRING,
        IDM_VIEW_TEXT_SIZE_XLARGE,
        &labels.view_text_size_xlarge,
    );
    let _ = append_menu_string(
        view_size_menu,
        MF_STRING,
        IDM_VIEW_TEXT_SIZE_XXLARGE,
        &labels.view_text_size_xxlarge,
    );
    let _ = append_menu_string(
        view_menu,
        MF_POPUP,
        view_color_menu.0 as usize,
        &labels.view_text_color,
    );
    let _ = append_menu_string(
        view_menu,
        MF_POPUP,
        view_size_menu.0 as usize,
        &labels.view_text_size,
    );
    let _ = append_menu_string(hmenu, MF_POPUP, view_menu.0 as usize, &labels.menu_view);

    let _ = append_menu_string(
        insert_menu,
        MF_STRING,
        IDM_INSERT_BOOKMARK,
        &labels.insert_bookmark,
    );
    let _ = append_menu_string(
        insert_menu,
        MF_STRING,
        IDM_MANAGE_BOOKMARKS,
        &labels.manage_bookmarks,
    );
    let _ = append_menu_string(hmenu, MF_POPUP, insert_menu.0 as usize, &labels.menu_insert);

    let _ = append_menu_string(tools_menu, MF_STRING, IDM_TOOLS_PROMPT, &labels.menu_prompt);
    let _ = append_menu_string(
        tools_menu,
        MF_STRING,
        IDM_TOOLS_OPTIONS,
        &labels.menu_options,
    );
    let _ = append_menu_string(
        tools_menu,
        MF_STRING,
        IDM_TOOLS_DICTIONARY,
        &labels.menu_dictionary,
    );
    let _ = append_menu_string(
        tools_menu,
        MF_STRING,
        IDM_TOOLS_IMPORT_YOUTUBE,
        &labels.menu_import_youtube,
    );
    let _ = append_menu_string(hmenu, MF_POPUP, tools_menu.0 as usize, &labels.menu_tools);

    let _ = append_menu_string(help_menu, MF_STRING, IDM_HELP_GUIDE, &labels.help_guide);
    let _ = append_menu_string(
        help_menu,
        MF_STRING,
        IDM_HELP_CHANGELOG,
        &labels.help_changelog,
    );
    let _ = append_menu_string(
        help_menu,
        MF_STRING,
        IDM_HELP_CHECK_UPDATES,
        &labels.help_check_updates,
    );
    let _ = append_menu_string(help_menu, MF_STRING, IDM_HELP_ABOUT, &labels.help_about);
    let _ = append_menu_string(hmenu, MF_POPUP, help_menu.0 as usize, &labels.menu_help);

    let _ = SetMenu(hwnd, hmenu);
    (hmenu, recent_menu)
}

pub unsafe fn update_recent_menu(hwnd: HWND, hmenu_recent: HMENU) {
    let count = GetMenuItemCount(hmenu_recent);
    if count > 0 {
        for _ in 0..count {
            let _ = DeleteMenu(hmenu_recent, 0, MF_BYPOSITION);
        }
    }

    let (files, language): (Vec<std::path::PathBuf>, Language) = with_state(hwnd, |state| {
        (state.recent_files.clone(), state.settings.language)
    })
    .unwrap_or_default();
    if files.is_empty() {
        let labels = menu_labels(language);
        let _ = append_menu_string(hmenu_recent, MF_STRING | MF_GRAYED, 0, &labels.recent_empty);
    } else {
        for (i, path) in files.iter().enumerate() {
            let label = format!("&{} {}", i + 1, abbreviate_recent_label(path));
            let wide = to_wide(&label);
            let _ = AppendMenuW(
                hmenu_recent,
                MF_STRING,
                IDM_FILE_RECENT_BASE + i,
                PCWSTR(wide.as_ptr()),
            );
        }
    }
    let _ = DrawMenuBar(hwnd);
}

pub fn abbreviate_recent_label(path: &Path) -> String {
    let filename = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
    let parent = path.parent().and_then(|p| p.to_str()).unwrap_or("");
    if parent.is_empty() {
        return filename.to_string();
    }
    let mut suffix = parent.to_string();
    if suffix.len() > 24 {
        suffix = format!("...ருங்கள்{}", &suffix[suffix.len().saturating_sub(24)..]);
    }
    format!("{filename} - {suffix}")
}

pub unsafe fn append_menu_string(menu: HMENU, flags: MENU_ITEM_FLAGS, id: usize, text: &str) {
    let wide = to_wide(text);
    let _ = AppendMenuW(menu, flags, id, PCWSTR(wide.as_ptr()));
}
