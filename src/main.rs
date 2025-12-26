#![allow(unsafe_op_in_unsafe_fn)]
#![windows_subsystem = "windows"]

mod accessibility;
use accessibility::*;
mod settings;
use editor_manager::{Document};
use settings::*;
mod bookmarks;
use bookmarks::*;
mod tts_engine;
use tts_engine::*;
mod file_handler;
use file_handler::*;
mod menu;
use menu::*;
mod search;
use search::*;
mod audio_player;
use audio_player::*;
mod editor_manager;
use editor_manager::*;
mod app_windows;

use std::io::Write;

use std::mem::size_of;

use std::path::{Path, PathBuf};

use std::sync::atomic::{AtomicBool};

use std::sync::Arc;

use chrono::Local;

use serde::{Deserialize, Serialize};

use windows::core::{w, PCWSTR, PWSTR};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};

use windows::Win32::Graphics::Gdi::{GetStockObject, HBRUSH, HFONT, COLOR_WINDOW, DEFAULT_GUI_FONT};

use windows::Win32::System::DataExchange::COPYDATASTRUCT;

use windows::Win32::System::LibraryLoader::{GetModuleHandleW, LoadLibraryW};

use windows::Win32::UI::Controls::RichEdit::{

    CHARRANGE, EM_EXSETSEL, EM_EXGETSEL, TEXTRANGEW, EM_GETTEXTRANGE

};

use windows::Win32::UI::Controls::{

    InitCommonControlsEx, ICC_TAB_CLASSES, INITCOMMONCONTROLSEX, NMHDR, TCM_GETCURSEL,

    TCN_SELCHANGE, WC_TABCONTROLW,

};



use windows::Win32::UI::Controls::Dialogs::{

    FINDREPLACEW, FINDREPLACE_FLAGS, GetOpenFileNameW, GetSaveFileNameW,

    OPENFILENAMEW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_OVERWRITEPROMPT,

    OFN_PATHMUSTEXIST,

};

use windows::Win32::UI::Input::KeyboardAndMouse::{

    GetKeyState, SetFocus, VK_CONTROL, VK_F3, VK_F4, VK_F5,

    VK_F6, VK_TAB

};

use windows::Win32::UI::Shell::{DragAcceptFiles, DragFinish, DragQueryFileW, HDROP};

use windows::Win32::UI::WindowsAndMessaging::{
    CreateAcceleratorTableW, CreateWindowExW,
    DefWindowProcW, DestroyWindow, DispatchMessageW, FindWindowW,
    GetMessageW, GetWindowLongPtrW, LoadCursorW, LoadIconW,
    MessageBoxW, PostQuitMessage, RegisterClassW, SendMessageW,
    RegisterWindowMessageW, SetForegroundWindow, SetWindowLongPtrW,
    PostMessageW, WM_APP,
    TranslateAcceleratorW, TranslateMessage, CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT,
    GWLP_USERDATA, CREATESTRUCTW, EN_CHANGE,
    HMENU, HCURSOR, HICON, IDC_ARROW, IDI_APPLICATION,
    MB_ICONERROR, MB_ICONINFORMATION, MB_OK,
    MSG, WM_CLOSE,
    WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_DROPFILES, WM_KEYDOWN, WM_NOTIFY, WM_SIZE, WM_TIMER, WNDCLASSW, WS_CHILD,
    WS_CLIPCHILDREN, WS_OVERLAPPEDWINDOW, WS_VISIBLE, ACCEL, FVIRTKEY,
    FCONTROL, FSHIFT, WM_SETFOCUS, WM_NCDESTROY, HACCEL, WM_UNDO, WM_CUT, WM_COPY,
    WM_PASTE, WM_COPYDATA, KillTimer, SetTimer,
};





const WM_PDF_LOADED: u32 = WM_APP + 1;
const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
const WM_TTS_AUDIOBOOK_DONE: u32 = WM_APP + 4;
const WM_UPDATE_PROGRESS: u32 = WM_APP + 6;
const COPYDATA_OPEN_FILE: usize = 1;

struct PdfLoadResult {
    hwnd_edit: HWND,
    path: PathBuf,
    result: Result<String, String>,
}

struct PdfLoadingState {
    hwnd_edit: HWND,
    timer_id: usize,
    frame: usize,
}

fn log_path() -> Option<PathBuf> {
    let base = std::env::var_os("APPDATA")?;
    let mut path = PathBuf::from(base);
    path.push("Novapad");
    path.push("Novapad.log");
    Some(path)
}

fn log_debug(message: &str) {
    let Some(path) = log_path() else {
        return;
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(mut log) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let timestamp = Local::now().format("%Y-%m-%d %H:%M:%S");
        let _ = writeln!(log, "[{timestamp}] {message}");
    }
}






#[derive(Default)]
pub(crate) struct AppState {
    hwnd_tab: HWND,
    docs: Vec<Document>,
    current: usize,
    untitled_count: usize,
    hfont: HFONT,
    hmenu_recent: HMENU,
    recent_files: Vec<PathBuf>,
    settings: AppSettings,
    bookmarks: BookmarkStore,
    find_dialog: HWND,
    replace_dialog: HWND,
    options_dialog: HWND,
    help_window: HWND,
    bookmarks_window: HWND,
    find_msg: u32,
    find_text: Vec<u16>,
    replace_text: Vec<u16>,
    find_replace: Option<FINDREPLACEW>,
    replace_replace: Option<FINDREPLACEW>,
    last_find_flags: FINDREPLACE_FLAGS,
    pdf_loading: Vec<PdfLoadingState>,
    next_timer_id: usize,
    tts_session: Option<TtsSession>,
    tts_next_session_id: u64,
    voice_list: Vec<VoiceInfo>,
    audiobook_progress: HWND,
    audiobook_cancel: Option<Arc<AtomicBool>>,
    active_audiobook: Option<AudiobookPlayer>,
}

#[derive(Default, Serialize, Deserialize)]
struct RecentFileStore {
    files: Vec<String>,
}

fn main() -> windows::core::Result<()> {
    log_debug("Application started.");

    unsafe {
        let _ = LoadLibraryW(w!("Msftedit.dll"));
        let hinstance = HINSTANCE(GetModuleHandleW(None)?.0);
        let class_name = w!("NovapadWin32");

        let wc = WNDCLASSW {
            hCursor: HCURSOR(LoadCursorW(None, IDC_ARROW)?.0),
            hIcon: HICON(LoadIconW(None, IDI_APPLICATION)?.0),
            hInstance: hinstance,
            lpszClassName: class_name,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let args: Vec<String> = std::env::args().collect();
        let extra_paths: Vec<String> = if args.len() > 1 {
            args[1..].to_vec()
        } else {
            Vec::new()
        };
        let settings = load_settings();
        let file_to_open = extra_paths.first().cloned();
        if !extra_paths.is_empty() {
            if settings.open_behavior == OpenBehavior::NewTab {
                let existing = FindWindowW(class_name, PCWSTR::null());
                if existing.0 != 0 {
                    for path in &extra_paths {
                        if !send_open_file(existing, path) {
                            break;
                        }
                    }
                    SetForegroundWindow(existing);
                    return Ok(());
                }
            }
        }
        let lp_param = &file_to_open as *const Option<String> as *const std::ffi::c_void;

        let hwnd = CreateWindowExW(
            Default::default(),
            class_name,
            w!("Novapad"),
            WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            900,
            700,
            None,
            None,
            hinstance,
            Some(lp_param),
        );

        if hwnd.0 == 0 {
            return Ok(());
        }

        let accel = create_accelerators();
        let mut msg = MSG::default();
        while GetMessageW(&mut msg, HWND(0), 0, 0).into() {
            // Priority 1: Global navigation keys (Ctrl+Tab)
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
                if (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0 {
                    next_tab_with_prompt(hwnd);
                    continue;
                }
            }

            let mut handled = false;
            let _ = with_state(hwnd, |state| {
                // Audiobook keyboard controls (ONLY if no secondary window is open)
                if msg.message == WM_KEYDOWN {
                    let is_audiobook = state.docs.get(state.current).map(|d| matches!(d.format, FileFormat::Audiobook)).unwrap_or(false);
                    let secondary_open = state.bookmarks_window.0 != 0 || state.options_dialog.0 != 0 || state.help_window.0 != 0;
                    
                    if is_audiobook && !secondary_open {
                        match handle_player_keyboard(&msg, state.settings.audiobook_skip_seconds) {
                            PlayerAction::TogglePause => {
                                toggle_audiobook_pause(hwnd);
                                handled = true;
                                return;
                            }
                            PlayerAction::Seek(amount) => {
                                seek_audiobook(hwnd, amount);
                                handled = true;
                                return;
                            }
                            PlayerAction::Volume(delta) => {
                                change_audiobook_volume(hwnd, delta);
                                handled = true;
                                return;
                            }
                            PlayerAction::BlockNavigation => {
                                handled = true;
                                return;
                            }
                            PlayerAction::None => {}
                        }
                    }
                }

                if state.find_dialog.0 != 0 && handle_accessibility(state.find_dialog, &msg) {
                    handled = true;
                    return;
                }
                if state.replace_dialog.0 != 0
                    && handle_accessibility(state.replace_dialog, &msg)
                {
                    handled = true;
                    return;
                }

                if state.help_window.0 != 0 {
                    // Manual TAB handling for Help window
                    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
                        app_windows::help_window::handle_tab(state.help_window);
                        handled = true;
                        return;
                    }

                    if handle_accessibility(state.help_window, &msg) {
                        handled = true;
                        return;
                    }
                }

                if state.options_dialog.0 != 0 {
                    if app_windows::options_window::handle_navigation(state.options_dialog, &msg) {
                        handled = true;
                        return;
                    }
                }

                if state.audiobook_progress.0 != 0 {
                    if app_windows::audiobook_window::handle_navigation(state.audiobook_progress, &msg) {
                        handled = true;
                        return;
                    }
                }

                if state.bookmarks_window.0 != 0 {
                    if app_windows::bookmarks_window::handle_navigation(state.bookmarks_window, &msg) {
                        handled = true;
                        return;
                    }
                }
            });
            if handled {
                continue;
            }
            if accel.0 != 0 && TranslateAcceleratorW(hwnd, accel, &msg) != 0 {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    Ok(())
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    if let Some(find_msg) = with_state(hwnd, |state| state.find_msg) {
        if msg == find_msg {
            handle_find_message(hwnd, lparam);
            return LRESULT(0);
        }
    }

    match msg {
        WM_CREATE => {
            let mut icc = INITCOMMONCONTROLSEX {
                dwSize: size_of::<INITCOMMONCONTROLSEX>() as u32,
                dwICC: ICC_TAB_CLASSES,
            };
            InitCommonControlsEx(&mut icc);

            let hwnd_tab = CreateWindowExW(
                Default::default(),
                WC_TABCONTROLW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let hfont = HFONT(GetStockObject(DEFAULT_GUI_FONT).0);
            let find_msg = RegisterWindowMessageW(w!("commdlg_FindReplace"));
            let settings = load_settings();
            let bookmarks = load_bookmarks();
            let (_, recent_menu) = create_menus(hwnd, settings.language);
            let recent_files = load_recent_files();
            let state = Box::new(AppState {
                hwnd_tab,
                docs: Vec::new(),
                current: 0,
                untitled_count: 0,
                hfont,
                hmenu_recent: recent_menu,
                recent_files,
                settings,
                bookmarks,
                find_dialog: HWND(0),
                replace_dialog: HWND(0),
                options_dialog: HWND(0),
                help_window: HWND(0),
                bookmarks_window: HWND(0),
                find_msg,
                find_text: vec![0u16; 256],
                replace_text: vec![0u16; 256],
                find_replace: None,
                replace_replace: None,
                last_find_flags: FINDREPLACE_FLAGS(0),
                pdf_loading: Vec::new(),
                next_timer_id: 1,
                tts_session: None,
                tts_next_session_id: 1,
                voice_list: Vec::new(),
                audiobook_progress: HWND(0),
                audiobook_cancel: None,
                active_audiobook: None,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            update_recent_menu(hwnd, recent_menu);
            
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let lp_create_params = (*create_struct).lpCreateParams as *const Option<String>;
            let file_to_open = if !lp_create_params.is_null() {
                (*lp_create_params).as_ref()
            } else {
                None
            };

            if let Some(path_str) = file_to_open {
                editor_manager::open_document(hwnd, Path::new(path_str));
            } else {
                editor_manager::new_document(hwnd);
            }
            
            editor_manager::layout_children(hwnd);
            DragAcceptFiles(hwnd, true);
            LRESULT(0)
        }
        WM_SIZE => {
            editor_manager::layout_children(hwnd);
            LRESULT(0)
        }
        WM_SETFOCUS => {
            let _ = with_state(hwnd, |state| {
                if let Some(doc) = state.docs.get(state.current) {
                    if matches!(doc.format, FileFormat::Audiobook) {
                        unsafe { SetFocus(state.hwnd_tab); }
                    } else {
                        unsafe { SetFocus(doc.hwnd_edit); }
                    }
                }
            });
            LRESULT(0)
        }
        WM_NOTIFY => {
            let hdr = &*(lparam.0 as *const NMHDR);
            if hdr.code == TCN_SELCHANGE && hdr.hwndFrom == editor_manager::get_tab(hwnd) {
                attempt_switch_to_selected_tab(hwnd);
                return LRESULT(0);
            }
            if hdr.code == EN_CHANGE as u32 {
                editor_manager::mark_dirty_from_edit(hwnd, hdr.hwndFrom);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_TIMER => {
            handle_pdf_loading_timer(hwnd, wparam.0 as usize);
            LRESULT(0)
        }
        WM_PDF_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = Box::from_raw(lparam.0 as *mut PdfLoadResult);
            handle_pdf_loaded(hwnd, *payload);
            LRESULT(0)
        }
        WM_TTS_VOICES_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = Box::from_raw(lparam.0 as *mut Vec<VoiceInfo>);
            let voices: Vec<VoiceInfo> = *payload;
            let _ = with_state(hwnd, |state| {
                state.voice_list = voices.clone();
            });
            if let Some(dialog) = with_state(hwnd, |state| state.options_dialog) {
                if dialog.0 != 0 {
                    app_windows::options_window::refresh_voices(dialog);
                }
            }
            LRESULT(0)
        }
        WM_TTS_PLAYBACK_DONE => {
            let session_id = wparam.0 as u64;
            let _ = with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session {
                    if current.id == session_id {
                        state.tts_session = None;
                        prevent_sleep(false);
                    }
                }
            });
            LRESULT(0)
        }
        WM_TTS_CHUNK_START => {
            let session_id = wparam.0 as u64;
            let offset = lparam.0 as i32;
            let _ = with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session {
                    if current.id == session_id && state.settings.move_cursor_during_reading {
                        if let Some(doc) = state.docs.get(state.current) {
                            let new_pos = current.initial_caret_pos + offset;
                            let mut cr = CHARRANGE { cpMin: new_pos, cpMax: new_pos };
                            unsafe {
                                SendMessageW(doc.hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
                                SendMessageW(doc.hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
                            }
                        }
                    }
                }
            });
            LRESULT(0)
        }
        WM_TTS_PLAYBACK_ERROR => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = Box::from_raw(lparam.0 as *mut String);
            let message: String = *payload;
            let session_id = wparam.0 as u64;
            let mut should_show = false;
            let _ = with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session {
                    if current.id == session_id {
                        state.tts_session = None;
                        prevent_sleep(false);
                        should_show = true;
                    }
                }
            });
            if should_show {
                let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                show_error(hwnd, language, &message);
            } else {
                log_debug(&format!(
                    "TTS error ignored for session {session_id}: {message}"
                ));
            }
            LRESULT(0)
        }
        WM_TTS_AUDIOBOOK_DONE => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            
            let _ = with_state(hwnd, |state| {
                if state.audiobook_progress.0 != 0 {
                     let _ = DestroyWindow(state.audiobook_progress);
                     state.audiobook_progress = HWND(0);
                     state.audiobook_cancel = None;
                }
                if let Some(doc) = state.docs.get(state.current) {
                    SetFocus(doc.hwnd_edit);
                }
            });

            let payload = Box::from_raw(lparam.0 as *mut AudiobookResult);
            let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
            let title = to_wide(if payload.success {
                audiobook_done_title(language)
            } else {
                error_title(language)
            });
            let message = to_wide(&payload.message);
            let flags = if payload.success { MB_OK | MB_ICONINFORMATION } else { MB_OK | MB_ICONERROR };
            MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), flags);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == u32::from(VK_TAB.0)
                && (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0
            {
                next_tab_with_prompt(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            let notification = (wparam.0 >> 16) as u16;
            if u32::from(notification) == EN_CHANGE {
                mark_dirty_from_edit(hwnd, HWND(lparam.0));
                return LRESULT(0);
            }

            if cmd_id >= IDM_FILE_RECENT_BASE && cmd_id < IDM_FILE_RECENT_BASE + MAX_RECENT {
                open_recent_by_index(hwnd, cmd_id - IDM_FILE_RECENT_BASE);
                return LRESULT(0);
            }

            match cmd_id {
                IDM_FILE_NEW => {
                    log_debug("Menu: New document");
                    editor_manager::new_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_OPEN => {
                    log_debug("Menu: Open document");
                    if let Some(path) = open_file_dialog(hwnd) {
                        editor_manager::open_document(hwnd, &path);
                    }
                    LRESULT(0)
                }
                IDM_FILE_SAVE => {
                    log_debug("Menu: Save document");
                    let _ = editor_manager::save_current_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_SAVE_AS => {
                    log_debug("Menu: Save document as");
                    let _ = editor_manager::save_current_document_as(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_SAVE_ALL => {
                    log_debug("Menu: Save all documents");
                    let _ = editor_manager::save_all_documents(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_CLOSE => {
                    log_debug("Menu: Close document");
                    editor_manager::close_current_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_EXIT => {
                    log_debug("Menu: Exit");
                    let _ = editor_manager::try_close_app(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_READ_START => {
                    log_debug("Menu: Start reading");
                    tts_engine::start_tts_from_caret(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_READ_PAUSE => {
                    log_debug("Menu: Pause/resume reading");
                    tts_engine::toggle_tts_pause(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_READ_STOP => {
                    log_debug("Menu: Stop reading");
                    tts_engine::stop_tts_playback(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_AUDIOBOOK => {
                    log_debug("Menu: Record audiobook");
                    tts_engine::start_audiobook(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_UNDO => {
                    log_debug("Menu: Undo");
                    editor_manager::send_to_active_edit(hwnd, WM_UNDO);
                    LRESULT(0)
                }
                IDM_EDIT_CUT => {
                    log_debug("Menu: Cut");
                    editor_manager::send_to_active_edit(hwnd, WM_CUT);
                    LRESULT(0)
                }
                IDM_EDIT_COPY => {
                    log_debug("Menu: Copy");
                    editor_manager::send_to_active_edit(hwnd, WM_COPY);
                    LRESULT(0)
                }
                IDM_EDIT_PASTE => {
                    log_debug("Menu: Paste");
                    editor_manager::send_to_active_edit(hwnd, WM_PASTE);
                    LRESULT(0)
                }
                IDM_EDIT_SELECT_ALL => {
                    log_debug("Menu: Select All");
                    editor_manager::select_all_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_FIND => {
                    log_debug("Menu: Find");
                    search::open_find_dialog(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_FIND_NEXT => {
                    log_debug("Menu: Find next");
                    search::find_next_from_state(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_REPLACE => {
                    log_debug("Menu: Replace");
                    search::open_replace_dialog(hwnd);
                    LRESULT(0)
                }
                IDM_INSERT_BOOKMARK => {
                    log_debug("Menu: Insert Bookmark");
                    insert_bookmark(hwnd);
                    LRESULT(0)
                }
                IDM_MANAGE_BOOKMARKS => {
                    log_debug("Menu: Manage Bookmarks");
                    app_windows::bookmarks_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_NEXT_TAB => {
                    next_tab_with_prompt(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_OPTIONS => {
                    log_debug("Menu: Options");
                    app_windows::options_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_GUIDE => {
                    log_debug("Menu: Guide");
                    app_windows::help_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_ABOUT => {
                    log_debug("Menu: About");
                    app_windows::about_window::show(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            let _ = try_close_app(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            PostQuitMessage(0);
            LRESULT(0)
        }
        WM_DROPFILES => {
            handle_drop_files(hwnd, HDROP(wparam.0 as isize));
            LRESULT(0)
        }
        WM_COPYDATA => {
            let cds = &*(lparam.0 as *const COPYDATASTRUCT);
            if cds.dwData == COPYDATA_OPEN_FILE && !cds.lpData.is_null() {
                let path = from_wide(cds.lpData as *const u16);
                if !path.is_empty() {
                    open_document(hwnd, Path::new(&path));
                    SetForegroundWindow(hwnd);
                    if let Some(hwnd_edit) = get_active_edit(hwnd) {
                        SetFocus(hwnd_edit);
                    }
                }
                return LRESULT(1);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
            if !ptr.is_null() {
                drop(Box::from_raw(ptr));
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

























unsafe fn create_accelerators() -> HACCEL {
    let virt = FCONTROL | FVIRTKEY;
    let virt_shift = FCONTROL | FSHIFT | FVIRTKEY;
    let mut accels = [
        ACCEL { fVirt: virt, key: 'N' as u16, cmd: IDM_FILE_NEW as u16 },
        ACCEL { fVirt: virt, key: 'O' as u16, cmd: IDM_FILE_OPEN as u16 },
        ACCEL { fVirt: virt, key: 'S' as u16, cmd: IDM_FILE_SAVE as u16 },
        ACCEL { fVirt: virt_shift, key: 'S' as u16, cmd: IDM_FILE_SAVE_ALL as u16 },
        ACCEL { fVirt: virt, key: 'W' as u16, cmd: IDM_FILE_CLOSE as u16 },
        ACCEL { fVirt: virt, key: 'F' as u16, cmd: IDM_EDIT_FIND as u16 },
        ACCEL { fVirt: FVIRTKEY, key: VK_F3.0 as u16, cmd: IDM_EDIT_FIND_NEXT as u16 },
        ACCEL { fVirt: virt, key: 'H' as u16, cmd: IDM_EDIT_REPLACE as u16 },
        ACCEL { fVirt: virt, key: 'A' as u16, cmd: IDM_EDIT_SELECT_ALL as u16 },
        ACCEL { fVirt: virt, key: VK_TAB.0 as u16, cmd: IDM_NEXT_TAB as u16 },
        ACCEL { fVirt: FVIRTKEY, key: VK_F4.0 as u16, cmd: IDM_FILE_READ_PAUSE as u16 },
        ACCEL { fVirt: FVIRTKEY, key: VK_F5.0 as u16, cmd: IDM_FILE_READ_START as u16 },
        ACCEL { fVirt: FVIRTKEY, key: VK_F6.0 as u16, cmd: IDM_FILE_READ_STOP as u16 },
        ACCEL { fVirt: virt, key: 'R' as u16, cmd: IDM_FILE_AUDIOBOOK as u16 },
        ACCEL { fVirt: virt, key: 'B' as u16, cmd: IDM_INSERT_BOOKMARK as u16 },
    ];
    CreateAcceleratorTableW(&mut accels).unwrap_or(HACCEL(0))
}


































pub(crate) unsafe fn get_active_edit(hwnd: HWND) -> Option<HWND> {
    with_state(hwnd, |state| state.docs.get(state.current).map(|doc| doc.hwnd_edit)).flatten()
}







unsafe fn insert_bookmark(hwnd: HWND) {
    let (hwnd_edit, path, format): (HWND, std::path::PathBuf, FileFormat) = match with_state(hwnd, |state| {
        state.docs.get(state.current).and_then(|doc| {
            doc.path.clone().map(|p| (doc.hwnd_edit, p, doc.format))
        })
    }) {
        Some(Some(values)) => values,
        _ => return,
    };

    if matches!(format, FileFormat::Audiobook) {
        let (pos, snippet) = with_state(hwnd, |state| {
            if let Some(player) = &mut state.active_audiobook {
                let current_total = if player.is_paused {
                    player.accumulated_seconds
                } else {
                    player.accumulated_seconds + player.start_instant.elapsed().as_secs()
                };
                let mins = current_total / 60;
                let secs = current_total % 60;
                (current_total as i32, format!("Posizione audio: {:02}:{:02}", mins, secs))
            } else {
                (0, "Audio non in riproduzione".to_string())
            }
        }).unwrap_or((0, String::new()));

        let bookmark = Bookmark {
            position: pos,
            snippet,
            timestamp: Local::now().format("%Y-%m-%d %H:%M:%S").to_string(),
        };

        let path_str = path.to_string_lossy().to_string();
        let bookmarks_window = with_state(hwnd, |state| {
            let list = state.bookmarks.files.entry(path_str).or_default();
            list.push(bookmark);
            save_bookmarks(&state.bookmarks);
            state.bookmarks_window
        }).unwrap_or(HWND(0));

        if bookmarks_window.0 != 0 {
            unsafe { app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window); }
        }
        return;
    }

    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    unsafe { SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize)); }
    
    let pos = cr.cpMax;
    
    // 1. Try to get up to 60 characters AFTER the cursor
    let mut buffer = vec![0u16; 62];
    let mut tr = TEXTRANGEW {
        chrg: CHARRANGE { cpMin: pos, cpMax: pos + 60 },
        lpstrText: PWSTR(buffer.as_mut_ptr()),
    };
    let copied = unsafe { SendMessageW(hwnd_edit, EM_GETTEXTRANGE, WPARAM(0), LPARAM(&mut tr as *mut _ as isize)).0 as usize };
    let mut snippet = String::from_utf16_lossy(&buffer[..copied]);
    
    // Stop at the first newline
    if let Some(idx) = snippet.find(|c| c == '\r' || c == '\n') {
        snippet.truncate(idx);
    }
    
    // 2. If the resulting snippet is empty (e.g. cursor at end of line), take text BEFORE the cursor
    if snippet.trim().is_empty() && pos > 0 {
        let start_pre = (pos - 60).max(0);
        let mut buffer_pre = vec![0u16; 62];
        let mut tr_pre = TEXTRANGEW {
            chrg: CHARRANGE { cpMin: start_pre, cpMax: pos },
            lpstrText: PWSTR(buffer_pre.as_mut_ptr()),
        };
        let copied_pre = unsafe { SendMessageW(hwnd_edit, EM_GETTEXTRANGE, WPARAM(0), LPARAM(&mut tr_pre as *mut _ as isize)).0 as usize };
        let mut snippet_pre = String::from_utf16_lossy(&buffer_pre[..copied_pre]);
        
        // Take text after the last newline in this prefix
        if let Some(idx) = snippet_pre.rfind(|c| c == '\r' || c == '\n') {
            snippet_pre = snippet_pre[idx+1..].to_string();
        }
        snippet = snippet_pre;
    }

    let bookmark = Bookmark {
        position: pos,
        snippet: snippet.trim().to_string(),
        timestamp: Local::now().format("%Y-%m-%d %H:%M:%S").to_string(),
    };

    let path_str = path.to_string_lossy().to_string();
    let bookmarks_window = with_state(hwnd, |state| {
        let list = state.bookmarks.files.entry(path_str).or_default();
        list.push(bookmark);
        save_bookmarks(&state.bookmarks);
        state.bookmarks_window
    }).unwrap_or(HWND(0));

    if bookmarks_window.0 != 0 {
        unsafe { app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window); }
    }
}

pub(crate) unsafe fn goto_first_bookmark(hwnd_edit: HWND, path: &Path, bookmarks: &BookmarkStore, format: FileFormat) {
    let path_str = path.to_string_lossy().to_string();
    if let Some(list) = bookmarks.files.get(&path_str) {
        if let Some(bm) = list.first() {
            if matches!(format, FileFormat::Audiobook) {
                // Audiobook position is handled by playback start
            } else {
                let mut cr = CHARRANGE { cpMin: bm.position, cpMax: bm.position };
                SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
                SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
            }
        }
    }
}
















pub(crate) unsafe fn rebuild_menus(hwnd: HWND) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let (_, recent_menu) = create_menus(hwnd, language);
    let _ = with_state(hwnd, |state| {
        state.hmenu_recent = recent_menu;
    });
    update_recent_menu(hwnd, recent_menu);
}

pub(crate) unsafe fn push_recent_file(hwnd: HWND, path: &Path) {
    let (hmenu_recent, files) = match with_state(hwnd, |state| {
        state.recent_files.retain(|p| p != path);
        state.recent_files.insert(0, path.to_path_buf());
        if state.recent_files.len() > MAX_RECENT {
            state.recent_files.truncate(MAX_RECENT);
        }
        (state.hmenu_recent, state.recent_files.clone())
    }) {
        Some(values) => values,
        None => return,
    };
    update_recent_menu(hwnd, hmenu_recent);
    save_recent_files(&files);
}

unsafe fn open_recent_by_index(hwnd: HWND, index: usize) {
    let path = with_state(hwnd, |state| state.recent_files.get(index).cloned()).unwrap_or(None);
    let Some(path) = path else { return; };
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if !path.exists() {
        show_error(hwnd, language, recent_missing_message(language));
        let files = with_state(hwnd, |state| {
            state.recent_files.retain(|p| p != &path);
            update_recent_menu(hwnd, state.hmenu_recent);
            state.recent_files.clone()
        })
        .unwrap_or_default();
        save_recent_files(&files);
        return;
    }
    open_document(hwnd, &path);
}







pub(crate) unsafe fn with_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut AppState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}





pub(crate) unsafe fn open_pdf_document_async(hwnd: HWND, path: &Path) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let path_buf = path.to_path_buf();
    let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File").to_string();
    let (hwnd_edit, new_index) = with_state(hwnd, |state| {
        let hwnd_edit = create_edit(hwnd, state.hfont, state.settings.word_wrap);
        editor_manager::set_edit_text(hwnd_edit, &pdf_loading_placeholder(0));
        let doc = Document {
            title: title.clone(),
            path: Some(path_buf.clone()),
            hwnd_edit,
            dirty: false,
            format: FileFormat::Pdf,
        };
        state.docs.push(doc);
        insert_tab(state.hwnd_tab, &title, (state.docs.len() - 1) as i32);
        (hwnd_edit, state.docs.len() - 1)
    })
    .unwrap_or((HWND(0), 0));

    if hwnd_edit.0 == 0 {
        return;
    }
    select_tab(hwnd, new_index);

    start_pdf_loading_animation(hwnd, hwnd_edit);

    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let result = read_pdf_text(&path_buf, language);
        let payload = Box::new(PdfLoadResult {
            hwnd_edit,
            path: path_buf,
            result,
        });
        unsafe {
            let payload_ptr = Box::into_raw(payload);
            if PostMessageW(
                hwnd_main,
                WM_PDF_LOADED,
                WPARAM(0),
                LPARAM(payload_ptr as isize),
            )
            .is_err()
            {
                let _ = Box::from_raw(payload_ptr);
            }
        }
    });
}

unsafe fn handle_pdf_loaded(hwnd: HWND, payload: PdfLoadResult) {
    let PdfLoadResult {
        hwnd_edit,
        path,
        result,
    } = payload;
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

    stop_pdf_loading_animation(hwnd, hwnd_edit);

    let doc_index = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .enumerate()
            .find_map(|(i, doc)| (doc.hwnd_edit == hwnd_edit).then_some(i))
    })
    .flatten();

    let Some(index) = doc_index else {
        return;
    };

    match result {
        Ok(text) => {
            editor_manager::set_edit_text(hwnd_edit, &text);
            let _ = with_state(hwnd, |state| {
                goto_first_bookmark(hwnd_edit, &path, &state.bookmarks, FileFormat::Pdf);
            });
            show_info(hwnd, language, pdf_loaded_message(language));
            let mut update_title = false;
            let _ = with_state(hwnd, |state| {
                if let Some(doc) = state.docs.get_mut(index) {
                    doc.dirty = false;
                    update_tab_title(state.hwnd_tab, index, &doc.title, false);
                    update_title = state.current == index;
                }
            });
            if update_title {
                update_window_title(hwnd);
            }
            push_recent_file(hwnd, &path);
        }
        Err(message) => {
            show_error(hwnd, language, &message);
            let _ = close_document_at(hwnd, index);
        }
    }
}

unsafe fn start_pdf_loading_animation(hwnd: HWND, hwnd_edit: HWND) {
    let timer_id = with_state(hwnd, |state| {
        let timer_id = state.next_timer_id;
        state.next_timer_id = state.next_timer_id.saturating_add(1);
        state.pdf_loading.push(PdfLoadingState {
            hwnd_edit,
            timer_id,
            frame: 0,
        });
        timer_id
    })
    .unwrap_or(0);

    if timer_id == 0 {
        return;
    }

    if SetTimer(hwnd, timer_id, 120, None) == 0 {
        stop_pdf_loading_animation(hwnd, hwnd_edit);
    }
}

unsafe fn stop_pdf_loading_animation(hwnd: HWND, hwnd_edit: HWND) {
    let mut timer_id = None;
    let _ = with_state(hwnd, |state| {
        if let Some(pos) = state
            .pdf_loading
            .iter()
            .position(|entry| entry.hwnd_edit == hwnd_edit)
        {
            timer_id = Some(state.pdf_loading[pos].timer_id);
            state.pdf_loading.swap_remove(pos);
        }
    });
    if let Some(timer_id) = timer_id {
        let _ = KillTimer(hwnd, timer_id);
    }
}

unsafe fn handle_pdf_loading_timer(hwnd: HWND, timer_id: usize) {
    let mut target = None;
    let _ = with_state(hwnd, |state| {
        if let Some(entry) = state
            .pdf_loading
            .iter_mut()
            .find(|entry| entry.timer_id == timer_id)
        {
            entry.frame = entry.frame.wrapping_add(1);
            target = Some((entry.hwnd_edit, entry.frame));
        }
    });

    if let Some((hwnd_edit, frame)) = target {
        editor_manager::set_edit_text(hwnd_edit, &pdf_loading_placeholder(frame));
    }
}

pub(crate) fn pdf_loading_placeholder(frame: usize) -> String {
    let spinner = ['|', '/', '-', '\\'][frame % 4];
    let bar_width = 24;
    let filled = frame % (bar_width + 1);
    let bar = format!(
        "{}{}",
        "#".repeat(filled),
        "-".repeat(bar_width.saturating_sub(filled))
    );
    format!(
        "Caricamento PDF...\r\n\r\n[{bar}]\r\nAnalisi in corso {spinner}"
    )
}

unsafe fn handle_drop_files(hwnd: HWND, hdrop: HDROP) {
    let count = DragQueryFileW(hdrop, 0xFFFFFFFF, None);
    for index in 0..count {
        let mut buffer = [0u16; 260];
        let len = DragQueryFileW(hdrop, index, Some(&mut buffer));
        if len == 0 {
            continue;
        }
        let path = PathBuf::from(from_wide(buffer.as_ptr()));
        if path.as_os_str().is_empty() {
            continue;
        }
        open_document(hwnd, &path);
    }
    DragFinish(hdrop);
}



unsafe fn next_tab_with_prompt(hwnd: HWND) {
    let (current, count) = match with_state(hwnd, |state| {
        if state.docs.is_empty() {
            return None;
        }
        let current = state.current;
        Some((current, state.docs.len()))
    }) {
        Some(Some(values)) => values,
        _ => return,
    };
    if count <= 1 {
        return;
    }
    let next = (current + 1) % count;
    select_tab(hwnd, next);
}

unsafe fn attempt_switch_to_selected_tab(hwnd: HWND) {
    let (current, hwnd_tab, count) = match with_state(hwnd, |state| {
        if state.docs.is_empty() {
            return None;
        }
        let current = state.current;
        Some((
            current,
            state.hwnd_tab,
            state.docs.len(),
        ))
    }) {
        Some(Some(values)) => values,
        _ => return,
    };
    let sel = SendMessageW(hwnd_tab, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
    if sel < 0 {
        return;
    }
    let sel = sel as usize;
    if sel >= count || sel == current {
        return;
    }
    select_tab(hwnd, sel);
}




fn suggested_filename_from_text(text: &str) -> Option<String> {
    let first_line = text.lines().next().unwrap_or("").trim();
    if first_line.is_empty() {
        return None;
    }
    let sanitized = sanitize_filename(first_line);
    if sanitized.is_empty() {
        None
    } else {
        Some(sanitized)
    }
}

fn sanitize_filename(input: &str) -> String {
    let mut out = String::new();
    for ch in input.chars() {
        if ch.is_ascii_control() {
            continue;
        }
        match ch {
            '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|' => out.push(' '),
            _ => out.push(ch),
        }
    }
    let mut cleaned = out
        .trim()
        .trim_end_matches(|c| c == '.' || c == ' ')
        .to_string();
    if cleaned.is_empty() {
        return cleaned;
    }
    if cleaned.len() > 120 {
        cleaned.truncate(120);
    }
    if is_reserved_filename(&cleaned) {
        cleaned.push('_');
    }
    cleaned
}

fn is_reserved_filename(name: &str) -> bool {
    let upper = name
        .trim_end_matches(|c| c == '.' || c == ' ')
        .to_ascii_uppercase();
    matches!(
        upper.as_str(),
        "CON"
            | "PRN"
            | "AUX"
            | "NUL"
            | "COM1"
            | "COM2"
            | "COM3"
            | "COM4"
            | "COM5"
            | "COM6"
            | "COM7"
            | "COM8"
            | "COM9"
            | "LPT1"
            | "LPT2"
            | "LPT3"
            | "LPT4"
            | "LPT5"
            | "LPT6"
            | "LPT7"
            | "LPT8"
            | "LPT9"
    )
}

unsafe fn open_file_dialog(hwnd: HWND) -> Option<PathBuf> {
    let filter = to_wide(
        "TXT (*.txt)\0*.txt\0PDF (*.pdf)\0*.pdf\0EPUB (*.epub)\0*.epub\0MP3 (*.mp3)\0*.mp3\0Word (*.doc;*.docx)\0*.doc;*.docx\0Excel (*.xls;*.xlsx)\0*.xls;*.xlsx\0RTF (*.rtf)\0*.rtf\0Tutti i file (*.*)\0*.*\0",
    );
    let mut file_buf = [0u16; 260];
    let mut ofn = OPENFILENAMEW {
        lStructSize: size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(file_buf.as_mut_ptr()),
        nMaxFile: file_buf.len() as u32,
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        ..Default::default()
    };
    let ok = GetOpenFileNameW(&mut ofn);
    if ok.as_bool() {
        Some(PathBuf::from(from_wide(file_buf.as_ptr())))
    } else {
        None
    }
}

unsafe fn save_file_dialog(hwnd: HWND, suggested_name: Option<&str>) -> Option<PathBuf> {
    let filter = to_wide(
        "TXT (*.txt)\0*.txt\0PDF (*.pdf)\0*.pdf\0EPUB (*.epub)\0*.epub\0Word (*.doc;*.docx)\0*.doc;*.docx\0Excel (*.xls;*.xlsx)\0*.xls;*.xlsx\0RTF (*.rtf)\0*.rtf\0Tutti i file (*.*)\0*.*\0",
    );
    let mut file_buf = [0u16; 260];
    if let Some(name) = suggested_name {
        let mut idx = 0usize;
        for &ch in to_wide(name).iter() {
            if ch == 0 || idx >= file_buf.len() - 1 {
                break;
            }
            file_buf[idx] = ch;
            idx += 1;
        }
        file_buf[idx] = 0;
    }
    let mut ofn = OPENFILENAMEW {
        lStructSize: size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(file_buf.as_mut_ptr()),
        nMaxFile: file_buf.len() as u32,
        Flags: OFN_EXPLORER | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST,
        ..Default::default()
    };
    let ok = GetSaveFileNameW(&mut ofn);
    if ok.as_bool() {
        let mut path = PathBuf::from(from_wide(file_buf.as_ptr()));
        if path.extension().is_none() {
            match ofn.nFilterIndex {
                1 => { path.set_extension("txt"); },
                2 => { path.set_extension("pdf"); },
                3 => { path.set_extension("docx"); },
                4 => { path.set_extension("xlsx"); },
                5 => { path.set_extension("rtf"); },
                _ => {},
            }
        }
        Some(path)
    } else {
        None
    }
}

pub(crate) unsafe fn save_audio_dialog(hwnd: HWND, suggested_name: Option<&str>) -> Option<PathBuf> {
    let mut file_buf = vec![0u16; 4096];
    if let Some(name) = suggested_name {
        let mut name_wide = to_wide(name);
        if let Some(0) = name_wide.last() { name_wide.pop(); }
        let copy_len = name_wide.len().min(file_buf.len() - 1);
        file_buf[..copy_len].copy_from_slice(&name_wide[..copy_len]);
    }
    let filter = to_wide("MP3 Files (*.mp3)\0*.mp3\0All Files (*.*)\0*.*\0\0");
    let title = to_wide("Audiobook");
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFile: PWSTR(file_buf.as_mut_ptr()),
        nMaxFile: file_buf.len() as u32,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrTitle: PCWSTR(title.as_ptr()),
        Flags: OFN_EXPLORER | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST,
        ..Default::default()
    };
    if GetSaveFileNameW(&mut ofn).as_bool() {
        let path = PathBuf::from(from_wide(file_buf.as_ptr()));
        let mut path = path;
        if path.extension().is_none() { path.set_extension("mp3"); }
        Some(path)
    } else { None }
}

pub(crate) unsafe fn show_error(hwnd: HWND, language: Language, message: &str) {
    log_debug(&format!("Error shown: {message}"));
    let wide = to_wide(message);
    let title = to_wide(error_title(language));
    MessageBoxW(hwnd, PCWSTR(wide.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONERROR);
}

pub(crate) unsafe fn show_info(hwnd: HWND, language: Language, message: &str) {
    log_debug(&format!("Info shown: {message}"));
    let wide = to_wide(message);
    let title = to_wide(info_title(language));
    MessageBoxW(hwnd, PCWSTR(wide.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONINFORMATION);
}

pub(crate) unsafe fn send_open_file(hwnd: HWND, path: &str) -> bool {
    let wide = to_wide(path);
    let data = COPYDATASTRUCT {
        dwData: COPYDATA_OPEN_FILE,
        cbData: (wide.len() * std::mem::size_of::<u16>()) as u32,
        lpData: wide.as_ptr() as *mut std::ffi::c_void,
    };
    SendMessageW(hwnd, WM_COPYDATA, WPARAM(0), LPARAM(&data as *const _ as isize));
    true
}

pub(crate) fn recent_store_path() -> Option<PathBuf> {
    let base = std::env::var_os("APPDATA")?;
    let mut path = PathBuf::from(base);
    path.push("Novapad");
    path.push("recent.json");
    Some(path)
}

fn load_recent_files() -> Vec<PathBuf> {
    let Some(path) = recent_store_path() else {
        return Vec::new();
    };
    let data = std::fs::read_to_string(path).ok();
    let Some(data) = data else {
        return Vec::new();
    };
    let store: RecentFileStore = serde_json::from_str(&data).unwrap_or_default();
    store
        .files
        .into_iter()
        .map(PathBuf::from)
        .filter(|p| !p.as_os_str().is_empty())
        .collect()
}

fn save_recent_files(files: &[PathBuf]) {
    let Some(path) = recent_store_path() else {
        return;
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let store = RecentFileStore {
        files: files.iter().map(|p| p.to_string_lossy().to_string()).collect(),
    };
    if let Ok(json) = serde_json::to_string_pretty(&store) {
        let _ = std::fs::write(path, json);
    }
}
