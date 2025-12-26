use windows::core::{PCWSTR};
use windows::Win32::Foundation::{HWND};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreateMenu, DeleteMenu, DrawMenuBar, GetMenuItemCount,
    SetMenu, HMENU, MENU_ITEM_FLAGS, MF_BYPOSITION,
    MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING
};
use crate::settings::Language;
use crate::accessibility::to_wide;
use crate::with_state;
use std::path::Path;

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
pub const IDM_EDIT_UNDO: usize = 2001;
pub const IDM_EDIT_CUT: usize = 2002;
pub const IDM_EDIT_COPY: usize = 2003;
pub const IDM_EDIT_PASTE: usize = 2004;
pub const IDM_EDIT_SELECT_ALL: usize = 2005;
pub const IDM_EDIT_FIND: usize = 2006;
pub const IDM_EDIT_FIND_NEXT: usize = 2007;
pub const IDM_EDIT_REPLACE: usize = 2008;
pub const IDM_INSERT_BOOKMARK: usize = 2101;
pub const IDM_MANAGE_BOOKMARKS: usize = 2102;
pub const IDM_NEXT_TAB: usize = 3001;
pub const IDM_FILE_RECENT_BASE: usize = 4000;
pub const IDM_TOOLS_OPTIONS: usize = 5001;
pub const IDM_HELP_GUIDE: usize = 7001;
pub const IDM_HELP_ABOUT: usize = 7002;
pub const MAX_RECENT: usize = 5;

pub struct MenuLabels {
    pub menu_file: &'static str,
    pub menu_edit: &'static str,
    pub menu_insert: &'static str,
    pub menu_tools: &'static str,
    pub menu_help: &'static str,
    pub menu_options: &'static str,
    pub file_new: &'static str,
    pub file_open: &'static str,
    pub file_save: &'static str,
    pub file_save_as: &'static str,
    pub file_save_all: &'static str,
    pub file_close: &'static str,
    pub file_recent: &'static str,
    pub file_read_start: &'static str,
    pub file_read_pause: &'static str,
    pub file_read_stop: &'static str,
    pub file_audiobook: &'static str,
    pub file_exit: &'static str,
    pub edit_undo: &'static str,
    pub edit_cut: &'static str,
    pub edit_copy: &'static str,
    pub edit_paste: &'static str,
    pub edit_select_all: &'static str,
    pub edit_find: &'static str,
    pub edit_find_next: &'static str,
    pub edit_replace: &'static str,
    pub insert_bookmark: &'static str,
    pub manage_bookmarks: &'static str,
    pub help_guide: &'static str,
    pub help_about: &'static str,
    pub recent_empty: &'static str,
}

pub fn menu_labels(language: Language) -> MenuLabels {
    match language {
        Language::Italian => MenuLabels {
            menu_file: "&File",
            menu_edit: "&Modifica",
            menu_insert: "&Inserisci",
            menu_tools: "S&trumenti",
            menu_help: "&Aiuto",
            menu_options: "&Opzioni...",
            file_new: "&Nuovo\tCtrl+N",
            file_open: "&Apri...\tCtrl+O",
            file_save: "&Salva\tCtrl+S",
            file_save_as: "Salva &come...",
            file_save_all: "Salva &tutto\tCtrl+Shift+S",
            file_close: "&Chiudi tab\tCtrl+W",
            file_recent: "File &recenti",
            file_read_start: "Avvia lettura\tF5",
            file_read_pause: "Pausa lettura\tF4",
            file_read_stop: "Stop lettura\tF6",
            file_audiobook: "Registra audiolibro...\tCtrl+R",
            file_exit: "&Esci",
            edit_undo: "&Annulla\tCtrl+Z",
            edit_cut: "&Taglia\tCtrl+X",
            edit_copy: "&Copia\tCtrl+C",
            edit_paste: "&Incolla\tCtrl+V",
            edit_select_all: "Seleziona &tutto\tCtrl+A",
            edit_find: "&Trova...\tCtrl+F",
            edit_find_next: "Trova &successivo\tF3",
            edit_replace: "&Sostituisci...\tCtrl+H",
            insert_bookmark: "Inserisci &segnalibro\tCtrl+B",
            manage_bookmarks: "&Gestisci segnalibri...",
            help_guide: "&Guida",
            help_about: "Informazioni &sul programma",
            recent_empty: "Nessun file recente",
        },
        Language::English => MenuLabels {
            menu_file: "&File",
            menu_edit: "&Edit",
            menu_insert: "&Insert",
            menu_tools: "&Tools",
            menu_help: "&Help",
            menu_options: "&Options...",
            file_new: "&New\tCtrl+N",
            file_open: "&Open...\tCtrl+O",
            file_save: "&Save\tCtrl+S",
            file_save_as: "Save &As...",
            file_save_all: "Save &All\tCtrl+Shift+S",
            file_close: "&Close tab\tCtrl+W",
            file_recent: "Recent &Files",
            file_read_start: "Start reading\tF5",
            file_read_pause: "Pause reading\tF4",
            file_read_stop: "Stop reading\tF6",
            file_audiobook: "Record audiobook...\tCtrl+R",
            file_exit: "E&xit",
            edit_undo: "&Undo\tCtrl+Z",
            edit_cut: "Cu&t\tCtrl+X",
            edit_copy: "&Copy\tCtrl+C",
            edit_paste: "&Paste\tCtrl+V",
            edit_select_all: "Select &All\tCtrl+A",
            edit_find: "&Find...\tCtrl+F",
            edit_find_next: "Find &Next\tF3",
            edit_replace: "&Replace...\tCtrl+H",
            insert_bookmark: "Insert &Bookmark\tCtrl+B",
            manage_bookmarks: "&Manage Bookmarks...",
            help_guide: "&Guide",
            help_about: "&About the program",
            recent_empty: "No recent files",
        },
    }
}

pub unsafe fn create_menus(hwnd: HWND, language: Language) -> (HMENU, HMENU) {
    let hmenu = CreateMenu().unwrap_or(HMENU(0));
    let file_menu = CreateMenu().unwrap_or(HMENU(0));
    let recent_menu = CreateMenu().unwrap_or(HMENU(0));
    let edit_menu = CreateMenu().unwrap_or(HMENU(0));
    let insert_menu = CreateMenu().unwrap_or(HMENU(0));
    let tools_menu = CreateMenu().unwrap_or(HMENU(0));
    let help_menu = CreateMenu().unwrap_or(HMENU(0));

    let labels = menu_labels(language);

    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_NEW, labels.file_new);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_OPEN, labels.file_open);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE, labels.file_save);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE_AS, labels.file_save_as);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_SAVE_ALL, labels.file_save_all);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_CLOSE, labels.file_close);
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(file_menu, MF_POPUP, recent_menu.0 as usize, labels.file_recent);
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_READ_START, labels.file_read_start);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_READ_PAUSE, labels.file_read_pause);
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_READ_STOP, labels.file_read_stop);
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_AUDIOBOOK, labels.file_audiobook);
    let _ = AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(file_menu, MF_STRING, IDM_FILE_EXIT, labels.file_exit);
    let _ = append_menu_string(hmenu, MF_POPUP, file_menu.0 as usize, labels.menu_file);

    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_UNDO, labels.edit_undo);
    let _ = AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_CUT, labels.edit_cut);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_COPY, labels.edit_copy);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_PASTE, labels.edit_paste);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_SELECT_ALL, labels.edit_select_all);
    let _ = AppendMenuW(edit_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_FIND, labels.edit_find);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_FIND_NEXT, labels.edit_find_next);
    let _ = append_menu_string(edit_menu, MF_STRING, IDM_EDIT_REPLACE, labels.edit_replace);
    let _ = append_menu_string(hmenu, MF_POPUP, edit_menu.0 as usize, labels.menu_edit);

    let _ = append_menu_string(insert_menu, MF_STRING, IDM_INSERT_BOOKMARK, labels.insert_bookmark);
    let _ = append_menu_string(insert_menu, MF_STRING, IDM_MANAGE_BOOKMARKS, labels.manage_bookmarks);
    let _ = append_menu_string(hmenu, MF_POPUP, insert_menu.0 as usize, labels.menu_insert);

    let _ = append_menu_string(tools_menu, MF_STRING, IDM_TOOLS_OPTIONS, labels.menu_options);
    let _ = append_menu_string(hmenu, MF_POPUP, tools_menu.0 as usize, labels.menu_tools);

    let _ = append_menu_string(help_menu, MF_STRING, IDM_HELP_GUIDE, labels.help_guide);
    let _ = append_menu_string(help_menu, MF_STRING, IDM_HELP_ABOUT, labels.help_about);
    let _ = append_menu_string(hmenu, MF_POPUP, help_menu.0 as usize, labels.menu_help);

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
        let _ = append_menu_string(hmenu_recent, MF_STRING | MF_GRAYED, 0, labels.recent_empty);
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
