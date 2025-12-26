#![allow(unsafe_op_in_unsafe_fn)]
#![windows_subsystem = "windows"]

mod accessibility;
use accessibility::*;
mod settings;
use settings::*;
mod bookmarks;
use bookmarks::*;
mod tts_engine;
use tts_engine::*;
mod file_handler;
use file_handler::*;
mod app_windows;

use std::fmt::Display;

use std::io::{Write};

use std::mem::size_of;

use std::path::{Path, PathBuf};

use std::sync::atomic::{AtomicBool};

use std::sync::Arc;

use std::time::Duration;

use chrono::Local;

use serde::{Deserialize, Serialize};

use rodio::{Decoder, OutputStream, Sink, Source};

use windows::core::{w, PCWSTR, PWSTR};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM, BOOL};

use windows::Win32::Graphics::Gdi::{GetStockObject, HBRUSH, HFONT, COLOR_WINDOW, DEFAULT_GUI_FONT, InvalidateRect, UpdateWindow};

use windows::Win32::System::DataExchange::COPYDATASTRUCT;

use windows::Win32::System::LibraryLoader::{GetModuleHandleW, LoadLibraryW};

use windows::Win32::UI::Controls::RichEdit::{

    MSFTEDIT_CLASS, EM_SETEVENTMASK, ENM_CHANGE, FINDTEXTEXW, CHARRANGE, EM_FINDTEXTEXW, EM_EXSETSEL, EM_EXGETSEL,

    TEXTRANGEW, EM_GETTEXTRANGE

};

use windows::Win32::UI::Controls::{

    InitCommonControlsEx, ICC_TAB_CLASSES, INITCOMMONCONTROLSEX, NMHDR, TCITEMW, TCIF_TEXT,

    TCM_ADJUSTRECT, TCM_DELETEITEM, TCM_GETCURSEL, TCM_INSERTITEMW, TCM_SETCURSEL, TCM_SETITEMW,

    TCN_SELCHANGE, WC_TABCONTROLW, EM_GETMODIFY, EM_SETMODIFY, EM_SETREADONLY,

};



use windows::Win32::UI::Controls::Dialogs::{

    FindTextW, ReplaceTextW, FINDREPLACEW, FINDREPLACE_FLAGS, FR_DIALOGTERM, FR_DOWN, FR_FINDNEXT,

    FR_MATCHCASE, FR_REPLACE, FR_REPLACEALL, FR_WHOLEWORD, GetOpenFileNameW, GetSaveFileNameW,

    OPENFILENAMEW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_OVERWRITEPROMPT,

    OFN_PATHMUSTEXIST,

};

use windows::Win32::UI::Input::KeyboardAndMouse::{

    GetKeyState, SetFocus, VK_CONTROL, VK_F3, VK_F4, VK_F5,

    VK_F6, VK_TAB

};

use windows::Win32::UI::Shell::{DragAcceptFiles, DragFinish, DragQueryFileW, HDROP};

use windows::Win32::UI::WindowsAndMessaging::{

    AppendMenuW, CreateAcceleratorTableW, CreateMenu, CreateWindowExW,

    DefWindowProcW, DeleteMenu, DestroyWindow, DispatchMessageW, DrawMenuBar, FindWindowW,

    GetClientRect, GetMenuItemCount, GetMessageW, GetWindowLongPtrW, LoadCursorW, LoadIconW,

    MessageBoxW, MoveWindow, PostQuitMessage, RegisterClassW, SendMessageW,

    SetMenu, RegisterWindowMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW,

    ShowWindow, PostMessageW, WM_APP,

    TranslateAcceleratorW, TranslateMessage, CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT,

    EN_CHANGE, GWLP_USERDATA, CREATESTRUCTW,

    HMENU, HCURSOR, HICON, IDC_ARROW, IDI_APPLICATION, IDYES, IDNO, MENU_ITEM_FLAGS,

    MB_ICONERROR, MB_ICONINFORMATION, MB_ICONWARNING, MB_OK, MB_YESNOCANCEL, MF_BYPOSITION,

    MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING, MSG, SW_HIDE, SW_SHOW, WM_CLOSE,

    WM_COMMAND,

    WM_CREATE, WM_DESTROY, WM_DROPFILES, WM_KEYDOWN, WM_NOTIFY, WM_SIZE, WM_TIMER, WNDCLASSW, WS_CHILD,

    WS_CLIPCHILDREN, WS_EX_CLIENTEDGE, WS_OVERLAPPEDWINDOW, WS_VISIBLE, ES_AUTOVSCROLL,

    ES_AUTOHSCROLL, ES_MULTILINE, ES_WANTRETURN, WS_HSCROLL, WS_VSCROLL, ACCEL, FVIRTKEY,

    FCONTROL, FSHIFT, WM_SETFOCUS, WM_NCDESTROY, HACCEL, WM_UNDO, WM_CUT, WM_COPY, WINDOW_STYLE,

    WM_PASTE, WM_GETTEXT, WM_GETTEXTLENGTH, WM_COPYDATA, KillTimer, SetTimer,

    WM_SETREDRAW,

};





const IDM_FILE_NEW: usize = 1001;
const IDM_FILE_OPEN: usize = 1002;
const IDM_FILE_SAVE: usize = 1003;
const IDM_FILE_SAVE_AS: usize = 1004;
const IDM_FILE_SAVE_ALL: usize = 1005;
const IDM_FILE_CLOSE: usize = 1006;
const IDM_FILE_EXIT: usize = 1007;
const IDM_FILE_READ_START: usize = 1008;
const IDM_FILE_READ_PAUSE: usize = 1009;
const IDM_FILE_READ_STOP: usize = 1010;
const IDM_FILE_AUDIOBOOK: usize = 1011;
const IDM_EDIT_UNDO: usize = 2001;
const IDM_EDIT_CUT: usize = 2002;
const IDM_EDIT_COPY: usize = 2003;
const IDM_EDIT_PASTE: usize = 2004;
const IDM_EDIT_SELECT_ALL: usize = 2005;
const IDM_EDIT_FIND: usize = 2006;
const IDM_EDIT_FIND_NEXT: usize = 2007;
const IDM_EDIT_REPLACE: usize = 2008;
const IDM_INSERT_BOOKMARK: usize = 2101;
const IDM_MANAGE_BOOKMARKS: usize = 2102;
const IDM_NEXT_TAB: usize = 3001;
const IDM_FILE_RECENT_BASE: usize = 4000;
const IDM_TOOLS_OPTIONS: usize = 5001;
const IDM_HELP_GUIDE: usize = 7001;
const IDM_HELP_ABOUT: usize = 7002;
const MAX_RECENT: usize = 5;
const WM_PDF_LOADED: u32 = WM_APP + 1;
const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
const WM_TTS_AUDIOBOOK_DONE: u32 = WM_APP + 4;
const WM_UPDATE_PROGRESS: u32 = WM_APP + 6;
const FIND_DIALOG_ID: isize = 1;
const REPLACE_DIALOG_ID: isize = 2;
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

#[derive(Default)]
struct Document {
    title: String,
    path: Option<PathBuf>,
    hwnd_edit: HWND,
    dirty: bool,
    format: FileFormat,
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






struct AudiobookPlayer {
    path: PathBuf,
    sink: Arc<Sink>,
    _stream: OutputStream, // Must be kept alive
    is_paused: bool,
    start_instant: std::time::Instant,
    accumulated_seconds: u64,
    volume: f32,
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
                open_document(hwnd, Path::new(path_str));
            } else {
                new_document(hwnd);
            }
            
            layout_children(hwnd);
            DragAcceptFiles(hwnd, true);
            LRESULT(0)
        }
        WM_SIZE => {
            layout_children(hwnd);
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
            if hdr.code == TCN_SELCHANGE && hdr.hwndFrom == get_tab(hwnd) {
                attempt_switch_to_selected_tab(hwnd);
                return LRESULT(0);
            }
            if hdr.code == EN_CHANGE as u32 {
                mark_dirty_from_edit(hwnd, hdr.hwndFrom);
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
                    new_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_OPEN => {
                    log_debug("Menu: Open document");
                    if let Some(path) = open_file_dialog(hwnd) {
                        open_document(hwnd, &path);
                    }
                    LRESULT(0)
                }
                IDM_FILE_SAVE => {
                    log_debug("Menu: Save document");
                    let _ = save_current_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_SAVE_AS => {
                    log_debug("Menu: Save document as");
                    let _ = save_current_document_as(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_SAVE_ALL => {
                    log_debug("Menu: Save all documents");
                    let _ = save_all_documents(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_CLOSE => {
                    log_debug("Menu: Close document");
                    close_current_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_EXIT => {
                    log_debug("Menu: Exit");
                    let _ = try_close_app(hwnd);
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
                    send_to_active_edit(hwnd, WM_UNDO);
                    LRESULT(0)
                }
                IDM_EDIT_CUT => {
                    send_to_active_edit(hwnd, WM_CUT);
                    LRESULT(0)
                }
                IDM_EDIT_COPY => {
                    send_to_active_edit(hwnd, WM_COPY);
                    LRESULT(0)
                }
                IDM_EDIT_PASTE => {
                    send_to_active_edit(hwnd, WM_PASTE);
                    LRESULT(0)
                }
                IDM_EDIT_SELECT_ALL => {
                    select_all_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_FIND => {
                    log_debug("Menu: Find");
                    open_find_dialog(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_FIND_NEXT => {
                    log_debug("Menu: Find next");
                    find_next_from_state(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_REPLACE => {
                    log_debug("Menu: Replace");
                    open_replace_dialog(hwnd);
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

unsafe fn create_menus(hwnd: HWND, language: Language) -> (HMENU, HMENU) {
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

struct MenuLabels {
    menu_file: &'static str,
    menu_edit: &'static str,
    menu_insert: &'static str,
    menu_tools: &'static str,
    menu_help: &'static str,
    menu_options: &'static str,
    file_new: &'static str,
    file_open: &'static str,
    file_save: &'static str,
    file_save_as: &'static str,
    file_save_all: &'static str,
    file_close: &'static str,
    file_recent: &'static str,
    file_read_start: &'static str,
    file_read_pause: &'static str,
    file_read_stop: &'static str,
    file_audiobook: &'static str,
    file_exit: &'static str,
    edit_undo: &'static str,
    edit_cut: &'static str,
    edit_copy: &'static str,
    edit_paste: &'static str,
    edit_select_all: &'static str,
    edit_find: &'static str,
    edit_find_next: &'static str,
    edit_replace: &'static str,
    insert_bookmark: &'static str,
    manage_bookmarks: &'static str,
    help_guide: &'static str,
    help_about: &'static str,
    recent_empty: &'static str,
}

fn menu_labels(language: Language) -> MenuLabels {
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



fn untitled_base(language: Language) -> &'static str {
    match language {
        Language::Italian => "Senza titolo",
        Language::English => "Untitled",
    }
}

fn untitled_title(language: Language, count: usize) -> String {
    format!("{} {}", untitled_base(language), count)
}

fn recent_missing_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Il file recente non esiste piu'.",
        Language::English => "The recent file no longer exists.",
    }
}

fn confirm_save_message(language: Language, title: &str) -> String {

    match language {

        Language::Italian => format!("Il documento \"{}\" e' modificato. Salvare?", title),

        Language::English => format!("The document \"{}\" has been modified. Save?", title),

    }

}



fn confirm_title(language: Language) -> &'static str {

    match language {

        Language::Italian => "Conferma",

        Language::English => "Confirm",

    }

}



fn error_title(language: Language) -> &'static str {

    match language {

        Language::Italian => "Errore",

        Language::English => "Error",

    }

}



pub(crate) fn tts_no_text_message(language: Language) -> &'static str {

    match language {

        Language::Italian => "Non c'e' testo da leggere.",

        Language::English => "There is no text to read.",

    }

}



fn audiobook_done_title(language: Language) -> &'static str {

    match language {

        Language::Italian => "Audiolibro",

        Language::English => "Audiobook",

    }

}



fn info_title(language: Language) -> &'static str {

    match language {

        Language::Italian => "Info",

        Language::English => "Info",

    }

}



fn pdf_loaded_message(language: Language) -> &'static str {

    match language {

        Language::Italian => "PDF caricato.",

        Language::English => "PDF loaded.",

    }

}



fn text_not_found_message(language: Language) -> &'static str {

    match language {

        Language::Italian => "Testo non trovato.",

        Language::English => "Text not found.",

    }

}



fn find_title(language: Language) -> &'static str {

    match language {

        Language::Italian => "Trova",

        Language::English => "Find",

    }

}



fn error_open_file_message(language: Language, err: impl Display) -> String {

    match language {

        Language::Italian => format!("Errore apertura file: {err}"),

        Language::English => format!("Error opening file: {err}"),

    }

}



fn error_save_file_message(language: Language, err: impl Display) -> String {

    match language {

        Language::Italian => format!("Errore salvataggio file: {err}"),

        Language::English => format!("Error saving file: {err}"),

    }

}



unsafe fn append_menu_string(menu: HMENU, flags: MENU_ITEM_FLAGS, id: usize, text: &str) {
    let wide = to_wide(text);
    let _ = AppendMenuW(menu, flags, id, PCWSTR(wide.as_ptr()));
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

unsafe fn open_find_dialog(hwnd: HWND) {
    let has_dialog = with_state(hwnd, |state| state.find_dialog.0 != 0).unwrap_or(false);
    if has_dialog {
        let _ = with_state(hwnd, |state| {
            SetFocus(state.find_dialog);
        });
        return;
    }

    let _ = with_state(hwnd, |state| {
        let fr = FINDREPLACEW {
            lStructSize: size_of::<FINDREPLACEW>() as u32,
            hwndOwner: hwnd,
            Flags: FR_DOWN,
            lpstrFindWhat: PWSTR(state.find_text.as_mut_ptr()),
            wFindWhatLen: state.find_text.len() as u16,
            lCustData: LPARAM(FIND_DIALOG_ID),
            ..Default::default()
        };
        state.find_replace = Some(fr);
        if let Some(ref mut fr) = state.find_replace {
            let dialog = FindTextW(fr);
            state.find_dialog = dialog;
        }
    });
}

unsafe fn open_replace_dialog(hwnd: HWND) {
    let has_dialog = with_state(hwnd, |state| state.replace_dialog.0 != 0).unwrap_or(false);
    if has_dialog {
        let _ = with_state(hwnd, |state| {
            SetFocus(state.replace_dialog);
        });
        return;
    }

    let _ = with_state(hwnd, |state| {
        let fr = FINDREPLACEW {
            lStructSize: size_of::<FINDREPLACEW>() as u32,
            hwndOwner: hwnd,
            Flags: FR_DOWN,
            lpstrFindWhat: PWSTR(state.find_text.as_mut_ptr()),
            wFindWhatLen: state.find_text.len() as u16,
            lpstrReplaceWith: PWSTR(state.replace_text.as_mut_ptr()),
            wReplaceWithLen: state.replace_text.len() as u16,
            lCustData: LPARAM(REPLACE_DIALOG_ID),
            ..Default::default()
        };
        state.replace_replace = Some(fr);
        if let Some(ref mut fr) = state.replace_replace {
            let dialog = ReplaceTextW(fr);
            state.replace_dialog = dialog;
        }
    });
}




























unsafe fn handle_find_message(hwnd: HWND, lparam: LPARAM) {
    let fr = &*(lparam.0 as *const FINDREPLACEW);
    if (fr.Flags & FR_DIALOGTERM) != FINDREPLACE_FLAGS(0) {
        let _ = with_state(hwnd, |state| {
            if fr.lCustData.0 == FIND_DIALOG_ID {
                state.find_dialog = HWND(0);
                state.find_replace = None;
            } else if fr.lCustData.0 == REPLACE_DIALOG_ID {
                state.replace_dialog = HWND(0);
                state.replace_replace = None;
            }
        });
        return;
    }

    if (fr.Flags & (FR_FINDNEXT | FR_REPLACE | FR_REPLACEALL)) == FINDREPLACE_FLAGS(0) {
        return;
    }

    let search = from_wide(fr.lpstrFindWhat.0);
    if search.is_empty() {
        return;
    }

    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

    let find_flags = extract_find_flags(fr.Flags);
    let _ = with_state(hwnd, |state| {
        state.last_find_flags = find_flags;
    });

    if (fr.Flags & FR_REPLACEALL) != FINDREPLACE_FLAGS(0) {
        replace_all(hwnd, hwnd_edit, &search, &from_wide(fr.lpstrReplaceWith.0), find_flags);
        return;
    }

    if (fr.Flags & FR_REPLACE) != FINDREPLACE_FLAGS(0) {
        let replace = from_wide(fr.lpstrReplaceWith.0);
        let replaced = replace_selection_if_match(hwnd_edit, &search, &replace, find_flags);
        let found = find_next(hwnd_edit, &search, find_flags, true);
        if !replaced && !found {
            let message = to_wide(text_not_found_message(language));
            let title = to_wide(find_title(language));
            MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
        }
        return;
    }

    if find_next(hwnd_edit, &search, find_flags, true) {
        return;
    }
    let message = to_wide(text_not_found_message(language));
    let title = to_wide(find_title(language));
    MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
}

unsafe fn find_next_from_state(hwnd: HWND) {
    let (search, flags, language) = with_state(hwnd, |state| {
        let search = from_wide(state.find_text.as_ptr());
        (search, state.last_find_flags, state.settings.language)
    })
    .unwrap_or((String::new(), FINDREPLACE_FLAGS(0), Language::default()));
    if search.is_empty() {
        open_find_dialog(hwnd);
        return;
    }
    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    if !find_next(hwnd_edit, &search, flags, true) {
        let message = to_wide(text_not_found_message(language));
        let title = to_wide(find_title(language));
        MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
    }
}

pub(crate) unsafe fn get_active_edit(hwnd: HWND) -> Option<HWND> {
    with_state(hwnd, |state| state.docs.get(state.current).map(|doc| doc.hwnd_edit)).flatten()
}

fn extract_find_flags(flags: FINDREPLACE_FLAGS) -> FINDREPLACE_FLAGS {
    let mut out = FINDREPLACE_FLAGS(0);
    if (flags & FR_MATCHCASE) != FINDREPLACE_FLAGS(0) {
        out |= FR_MATCHCASE;
    }
    if (flags & FR_WHOLEWORD) != FINDREPLACE_FLAGS(0) {
        out |= FR_WHOLEWORD;
    }
    if (flags & FR_DOWN) != FINDREPLACE_FLAGS(0) {
        out |= FR_DOWN;
    }
    out
}

unsafe fn find_next(
    hwnd_edit: HWND,
    search: &str,
    flags: FINDREPLACE_FLAGS,
    wrap: bool,
) -> bool {
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
    
    let down = (flags & FR_DOWN) != FINDREPLACE_FLAGS(0);
    
    let mut ft = FINDTEXTEXW {
        chrg: CHARRANGE {
            cpMin: if down { cr.cpMax } else { cr.cpMin },
            cpMax: if down { -1 } else { 0 },
        },
        lpstrText: PCWSTR(to_wide(search).as_ptr()),
        chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
    };

    let result = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
    
    if result.0 != -1 {
        let mut sel = ft.chrgText;
        // Swap to put caret at the beginning
        std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
        SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut sel as *mut _ as isize));
        SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        SetFocus(hwnd_edit);
        return true;
    }

    if wrap {
        ft.chrg.cpMin = if down { 0 } else { -1 };
        ft.chrg.cpMax = if down { -1 } else { 0 };
        let result = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
        if result.0 != -1 {
            let mut sel = ft.chrgText;
            std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
            SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut sel as *mut _ as isize));
            SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
            SetFocus(hwnd_edit);
            return true;
        }
    }
    false
}



unsafe fn replace_selection_if_match(
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
) -> bool {
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
    
    if cr.cpMin == cr.cpMax {
        return false;
    }

    let wide_search = to_wide(search);
    let mut ft = FINDTEXTEXW {
        chrg: cr,
        lpstrText: PCWSTR(wide_search.as_ptr()),
        chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
    };
    
    let res = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
    
    if res.0 == cr.cpMin as isize && ft.chrgText.cpMax == cr.cpMax {
        let replace_wide = to_wide(replace);
        SendMessageW(
            hwnd_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(replace_wide.as_ptr() as isize),
        );
        true
    } else {
        false
    }
}

unsafe fn replace_all(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
) {
    if search.is_empty() {
        return;
    }
    let mut start = 0i32;
    let mut replaced_any = false;
    let replace_wide = to_wide(replace);
    
    loop {
        let mut ft = FINDTEXTEXW {
            chrg: CHARRANGE {
                cpMin: start,
                cpMax: -1,
            },
            lpstrText: PCWSTR(to_wide(search).as_ptr()),
            chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
        };

        let res = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
        
        if res.0 != -1 {
            SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut ft.chrgText as *mut _ as isize));
            SendMessageW(
                hwnd_edit,
                EM_REPLACESEL,
                WPARAM(1),
                LPARAM(replace_wide.as_ptr() as isize),
            );
            replaced_any = true;
            
            let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
            SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
            start = cr.cpMax;
        } else {
            break;
        }
    }
    
    if !replaced_any {
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        let message = to_wide(text_not_found_message(language));
        let title = to_wide(find_title(language));
        MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
    }
}

unsafe fn update_recent_menu(hwnd: HWND, hmenu_recent: HMENU) {
    let count = GetMenuItemCount(hmenu_recent);
    if count > 0 {
        for _ in 0..count {
            let _ = DeleteMenu(hmenu_recent, 0, MF_BYPOSITION);
        }
    }

    let (files, language) = with_state(hwnd, |state| {
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

unsafe fn goto_first_bookmark(hwnd_edit: HWND, path: &Path, bookmarks: &BookmarkStore, format: FileFormat) {
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
















unsafe fn start_audiobook_playback(hwnd: HWND, path: &Path) {
    let path_buf = path.to_path_buf();
    
    let bookmark_pos = with_state(hwnd, |state| {
        state.bookmarks.files.get(&path_buf.to_string_lossy().to_string())
            .and_then(|list| list.last()) // Use LAST bookmark for audio
            .map(|bm| bm.position)
            .unwrap_or(0)
    }).unwrap_or(0);

    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path_buf) {
            Ok(f) => f,
            Err(_) => return,
        };
        
        let source: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        // Skip to bookmark position if any
        if bookmark_pos > 0 {
            let skipped = source.skip_duration(std::time::Duration::from_secs(bookmark_pos as u64));
            sink.append(skipped);
        } else {
            sink.append(source);
        }

        let player = AudiobookPlayer {
            path: path_buf.clone(),
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: bookmark_pos as u64,
            volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}


unsafe fn toggle_audiobook_pause(hwnd: HWND) {
    with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.is_paused {
                player.sink.play();
                player.is_paused = false;
                player.start_instant = std::time::Instant::now();
            } else {
                player.sink.pause();
                player.is_paused = true;
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
            }
        }
    });
}

unsafe fn seek_audiobook(hwnd: HWND, seconds: i64) {
    let (path, current_pos) = match with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if !player.is_paused {
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
                player.start_instant = std::time::Instant::now();
            }
            let new_pos = (player.accumulated_seconds as i64 + seconds).max(0);
            player.accumulated_seconds = new_pos as u64;
            Some((player.path.clone(), new_pos))
        } else {
            None
        }
    }) {
        Some(Some(v)) => v,
        _ => return,
    };

    stop_audiobook_playback(hwnd);
    
    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = OutputStream::try_default().unwrap();
        let sink: Arc<Sink> = Arc::new(Sink::try_new(&handle).unwrap());
        let file = std::fs::File::open(&path).unwrap();
        let source: Decoder<_> = Decoder::new(std::io::BufReader::new(file)).unwrap();
        
        use rodio::Source;
        let skipped = source.skip_duration(Duration::from_secs(current_pos as u64));
        sink.append(skipped);

        let player = AudiobookPlayer {
            path,
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: current_pos as u64,
            volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

unsafe fn stop_audiobook_playback(hwnd: HWND) {
    with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            player.sink.stop();
        }
    });
}

pub(crate) unsafe fn start_audiobook_at(hwnd: HWND, path: &Path, seconds: u64) {
    stop_audiobook_playback(hwnd);
    let path_buf = path.to_path_buf();
    let hwnd_main = hwnd;
    
    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path_buf) {
            Ok(f) => f,
            Err(_) => return,
        };
        
        let source: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        if seconds > 0 {
            let skipped = source.skip_duration(std::time::Duration::from_secs(seconds));
            sink.append(skipped);
        } else {
            sink.append(source);
        }

        let player = AudiobookPlayer {
            path: path_buf.clone(),
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

unsafe fn change_audiobook_volume(hwnd: HWND, delta: f32) {
    with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            player.volume = (player.volume + delta).clamp(0.0, 1.0);
            player.sink.set_volume(player.volume);
        }
    });
}


pub(crate) unsafe fn rebuild_menus(hwnd: HWND) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let (_, recent_menu) = create_menus(hwnd, language);
    let _ = with_state(hwnd, |state| {
        state.hmenu_recent = recent_menu;
    });
    update_recent_menu(hwnd, recent_menu);
}

unsafe fn push_recent_file(hwnd: HWND, path: &Path) {
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

unsafe fn sync_dirty_from_edit(hwnd: HWND, index: usize) -> bool {
    let mut hwnd_edit = HWND(0);
    let mut is_dirty = false;
    let mut is_current = false;
    let _ = with_state(hwnd, |state| {
        if let Some(doc) = state.docs.get(index) {
            hwnd_edit = doc.hwnd_edit;
            is_dirty = doc.dirty;
            is_current = state.current == index;
        }
    });

    if hwnd_edit.0 == 0 {
        return is_dirty;
    }

    let modified = SendMessageW(hwnd_edit, EM_GETMODIFY, WPARAM(0), LPARAM(0)).0 != 0;
    if modified && !is_dirty {
        let _ = with_state(hwnd, |state| {
            if let Some(doc) = state.docs.get_mut(index) {
                doc.dirty = true;
                update_tab_title(state.hwnd_tab, index, &doc.title, true);
            }
        });
        if is_current {
            update_window_title(hwnd);
        }
    }
    is_dirty || modified
}

unsafe fn confirm_save_if_dirty_entry(hwnd: HWND, index: usize, title: &str) -> bool {
    if !sync_dirty_from_edit(hwnd, index) {
        return true;
    }

    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let message = confirm_save_message(language, title);
    let wide = to_wide(&message);
    let confirm = to_wide(confirm_title(language));
    let result = MessageBoxW(
        hwnd,
        PCWSTR(wide.as_ptr()),
        PCWSTR(confirm.as_ptr()),
        MB_YESNOCANCEL | MB_ICONWARNING,
    );
    match result {
        IDYES => save_document_at(hwnd, index, false),
        IDNO => true,
        _ => false,
    }
}

unsafe fn close_current_document(hwnd: HWND) {
    let index = match with_state(hwnd, |state| state.current) {
        Some(index) => index,
        None => return,
    };
    let _ = close_document_at(hwnd, index);
}

unsafe fn close_document_at(hwnd: HWND, index: usize) -> bool {
    let (current, hwnd_tab, count, title) = match with_state(hwnd, |state| {
        if index >= state.docs.len() {
            return None;
        }
        Some((
            state.current,
            state.hwnd_tab,
            state.docs.len(),
            state.docs[index].title.clone(),
        ))
    }) {
        Some(Some(values)) => values,
        _ => return false,
    };
    if index >= count {
        return false;
    }
    if !confirm_save_if_dirty_entry(hwnd, index, &title) {
        return false;
    }

    let mut was_empty = false;
    let mut new_hwnd_edit = None;
    let mut update_title = false;
    let mut closing_hwnd_edit = HWND(0);
    let _ = with_state(hwnd, |state| {
        let was_current = index == current;
        let doc = state.docs.remove(index);
        closing_hwnd_edit = doc.hwnd_edit;
        SendMessageW(hwnd_tab, TCM_DELETEITEM, WPARAM(index), LPARAM(0));

        if state.docs.is_empty() {
            was_empty = true;
            return;
        }

        if was_current {
            let idx = if index >= state.docs.len() {
                state.docs.len() - 1
            } else {
                index
            };
            state.current = idx;
            SendMessageW(hwnd_tab, TCM_SETCURSEL, WPARAM(idx), LPARAM(0));
            new_hwnd_edit = state.docs.get(idx).map(|doc| doc.hwnd_edit);
            update_title = true;
        } else if index < state.current {
            state.current -= 1;
            SendMessageW(hwnd_tab, TCM_SETCURSEL, WPARAM(state.current), LPARAM(0));
        }
    });

    if closing_hwnd_edit.0 != 0 {
        stop_pdf_loading_animation(hwnd, closing_hwnd_edit);
        let _ = DestroyWindow(closing_hwnd_edit);
    }

    if was_empty {
        new_document(hwnd);
        return true;
    }

    if let Some(hwnd_edit) = new_hwnd_edit {
        let is_audiobook = with_state(hwnd, |state| {
            state.docs.get(state.current).map(|d| matches!(d.format, FileFormat::Audiobook)).unwrap_or(false)
        }).unwrap_or(false);

        if is_audiobook {
            ShowWindow(hwnd_edit, SW_HIDE);
            let hwnd_tab = with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0));
            if hwnd_tab.0 != 0 { SetFocus(hwnd_tab); }
        } else {
            ShowWindow(hwnd_edit, SW_SHOW);
            SetFocus(hwnd_edit);
        }
        update_window_title(hwnd);
        layout_children(hwnd);
    } else if update_title {
        update_window_title(hwnd);
    }

    true
}

unsafe fn try_close_app(hwnd: HWND) -> bool {
    let entries = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .enumerate()
            .map(|(i, doc)| (i, doc.title.clone()))
            .collect::<Vec<_>>()
    })
    .unwrap_or_default();
    for (index, title) in entries {
        if !confirm_save_if_dirty_entry(hwnd, index, &title) {
            return false;
        }
    }
    let _ = DestroyWindow(hwnd);
    true
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

unsafe fn get_current_index(hwnd: HWND) -> usize {
    with_state(hwnd, |state| state.current).unwrap_or(0)
}

unsafe fn get_tab(hwnd: HWND) -> HWND {
    with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0))
}

unsafe fn new_document(hwnd: HWND) {
    let new_index = with_state(hwnd, |state| {
        state.untitled_count += 1;
        let title = untitled_title(state.settings.language, state.untitled_count);
        let hwnd_edit = create_edit(hwnd, state.hfont, state.settings.word_wrap);
        let doc = Document {
            title: title.clone(),
            path: None,
            hwnd_edit,
            dirty: false,
            format: FileFormat::Text(TextEncoding::Utf8),
        };
        state.docs.push(doc);
        insert_tab(state.hwnd_tab, &title, (state.docs.len() - 1) as i32);
        state.docs.len() - 1
    })
    .unwrap_or(0);
    select_tab(hwnd, new_index);
}

unsafe fn open_document(hwnd: HWND, path: &Path) {
    log_debug(&format!("Open document: {}", path.display()));

    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if is_pdf_path(path) {
        open_pdf_document_async(hwnd, path);
        return;
    }
    let (content, format) = if is_docx_path(path) {
        match read_docx_text(path, language) {
            Ok(text) => (text, FileFormat::Docx),
            Err(message) => {
                show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_epub_path(path) {
        match read_epub_text(path, language) {
            Ok(text) => (text, FileFormat::Epub),
            Err(message) => {
                show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_mp3_path(path) {
        (String::new(), FileFormat::Audiobook)
    } else if is_doc_path(path) {
        match read_doc_text(path, language) {
            Ok(text) => (text, FileFormat::Doc),
            Err(message) => {
                show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_spreadsheet_path(path) {
        match read_spreadsheet_text(path, language) {
            Ok(text) => (text, FileFormat::Spreadsheet),
            Err(message) => {
                show_error(hwnd, language, &message);
                return;
            }
        }
    } else {
        match std::fs::read(path) {
            Ok(bytes) => match decode_text(&bytes, language) {
                Ok((text, encoding)) => (text, FileFormat::Text(encoding)),
                Err(message) => {
                    show_error(hwnd, language, &message);
                    return;
                }
            },
            Err(err) => {
                show_error(hwnd, language, &error_open_file_message(language, err));
                return;
            }
        }
    };

    let new_index = with_state(hwnd, |state| {
        let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
        let hwnd_edit = create_edit(hwnd, state.hfont, state.settings.word_wrap);
        set_edit_text(hwnd_edit, &content);

        let doc = Document {
            title: title.to_string(),
            path: Some(path.to_path_buf()),
            hwnd_edit,
            dirty: false,
            format,
        };
        if matches!(format, FileFormat::Audiobook) {
            unsafe {
                SendMessageW(hwnd_edit, EM_SETREADONLY, WPARAM(1), LPARAM(0));
                ShowWindow(hwnd_edit, SW_HIDE);
            }
        }
        state.docs.push(doc);
        insert_tab(state.hwnd_tab, title, (state.docs.len() - 1) as i32);
        goto_first_bookmark(hwnd_edit, path, &state.bookmarks, format);
        state.docs.len() - 1
    })
    .unwrap_or(0);
    select_tab(hwnd, new_index);
    if matches!(format, FileFormat::Audiobook) {
        unsafe { start_audiobook_playback(hwnd, path); }
    }
    push_recent_file(hwnd, path);
}

unsafe fn open_pdf_document_async(hwnd: HWND, path: &Path) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let path_buf = path.to_path_buf();
    let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File").to_string();
    let (hwnd_edit, new_index) = with_state(hwnd, |state| {
        let hwnd_edit = create_edit(hwnd, state.hfont, state.settings.word_wrap);
        set_edit_text(hwnd_edit, &pdf_loading_placeholder(0));
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
            set_edit_text(hwnd_edit, &text);
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
        set_edit_text(hwnd_edit, &pdf_loading_placeholder(frame));
    }
}

fn pdf_loading_placeholder(frame: usize) -> String {
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

unsafe fn save_current_document(hwnd: HWND) -> bool {
    save_document_at(hwnd, get_current_index(hwnd), false)
}

unsafe fn save_current_document_as(hwnd: HWND) -> bool {
    save_document_at(hwnd, get_current_index(hwnd), true)
}

unsafe fn save_all_documents(hwnd: HWND) -> bool {
    let dirty_indices = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .enumerate()
            .filter_map(|(i, doc)| if doc.dirty { Some(i) } else { None })
            .collect::<Vec<_>>()
    })
    .unwrap_or_default();
    for index in dirty_indices {
        if !save_document_at(hwnd, index, false) {
            return false;
        }
    }
    true
}

unsafe fn save_document_at(hwnd: HWND, index: usize, force_dialog: bool) -> bool {
    let path = match with_state(hwnd, |state| {
        if state.docs.is_empty() || index >= state.docs.len() {
            return None;
        }
        let language = state.settings.language;
        let text = get_edit_text(state.docs[index].hwnd_edit);
        let suggested_name = suggested_filename_from_text(&text)
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| state.docs[index].title.clone());

        let path = if !force_dialog {
            state.docs[index].path.clone()
        } else {
            None
        };
        let path = match path {
            Some(path) => path,
            None => match save_file_dialog(hwnd, Some(&suggested_name)) {
                Some(path) => path,
                None => return None,
            },
        };

        let is_docx = is_docx_path(&path);
        let is_pdf = is_pdf_path(&path);
        if is_docx {
            if let Err(message) = write_docx_text(&path, &text, language) {
                show_error(hwnd, language, &message);
                return None;
            }
            state.docs[index].format = FileFormat::Docx;
        } else if is_pdf {
            let pdf_title = path.file_stem().and_then(|s| s.to_str()).unwrap_or("Documento");
            if let Err(message) = write_pdf_text(&path, pdf_title, &text, language) {
                show_error(hwnd, language, &message);
                return None;
            }
            state.docs[index].format = FileFormat::Pdf;
        } else {
            let encoding = match state.docs[index].format {
                FileFormat::Text(enc) => enc,
                FileFormat::Docx | FileFormat::Doc | FileFormat::Pdf | FileFormat::Spreadsheet | FileFormat::Epub | FileFormat::Audiobook => TextEncoding::Utf8,
            };
            let bytes = encode_text(&text, encoding);
            if let Err(err) = std::fs::write(&path, bytes) {
                show_error(hwnd, language, &error_save_file_message(language, err));
                return None;
            }
            state.docs[index].format = FileFormat::Text(encoding);
        }

        let hwnd_edit = state.docs[index].hwnd_edit;
        state.docs[index].path = Some(path.clone());
        state.docs[index].dirty = false;
        SendMessageW(hwnd_edit, EM_SETMODIFY, WPARAM(0), LPARAM(0));
        let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
        state.docs[index].title = title.to_string();
        update_tab_title(state.hwnd_tab, index, &state.docs[index].title, false);
        if index == state.current {
            update_window_title(hwnd);
        }
        Some(path)
    }) {
        Some(Some(path)) => path,
        _ => return false,
    };
    push_recent_file(hwnd, &path);
    true
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

unsafe fn select_tab(hwnd: HWND, index: usize) {
    let result = with_state(hwnd, |state| {
        if index >= state.docs.len() {
            return None;
        }
        let prev = state.current;
        let prev_edit = state.docs.get(prev).map(|doc| doc.hwnd_edit);
        let new_doc = state.docs.get(index);
        let new_edit = new_doc.map(|doc| doc.hwnd_edit);
        let is_audiobook = new_doc.map(|doc| matches!(doc.format, FileFormat::Audiobook)).unwrap_or(false);
        state.current = index;
        Some((state.hwnd_tab, prev_edit, new_edit, is_audiobook))
    })
    .flatten();

    let Some((hwnd_tab, prev_edit, new_edit, is_audiobook)) = result else {
        return;
    };

    if let Some(hwnd_edit) = prev_edit {
        ShowWindow(hwnd_edit, SW_HIDE);
    }
    SendMessageW(hwnd_tab, TCM_SETCURSEL, WPARAM(index), LPARAM(0));
    if let Some(hwnd_edit) = new_edit {
        if is_audiobook {
            ShowWindow(hwnd_edit, SW_HIDE);
            SetFocus(hwnd_tab);
        } else {
            ShowWindow(hwnd_edit, SW_SHOW);
            SetFocus(hwnd_edit);
        }
    }
    update_window_title(hwnd);
    layout_children(hwnd);
}

unsafe fn insert_tab(hwnd_tab: HWND, title: &str, index: i32) {
    let mut text = to_wide(title);
    let mut item = TCITEMW {
        mask: TCIF_TEXT,
        pszText: PWSTR(text.as_mut_ptr()),
        ..Default::default()
    };
    SendMessageW(hwnd_tab, TCM_INSERTITEMW, WPARAM(index as usize), LPARAM(&mut item as *mut _ as isize));
}

unsafe fn update_tab_title(hwnd_tab: HWND, index: usize, title: &str, dirty: bool) {
    let label = if dirty {
        format!("{title}*")
    } else {
        title.to_string()
    };
    let mut text = to_wide(&label);
    let mut item = TCITEMW {
        mask: TCIF_TEXT,
        pszText: PWSTR(text.as_mut_ptr()),
        ..Default::default()
    };
    SendMessageW(hwnd_tab, TCM_SETITEMW, WPARAM(index), LPARAM(&mut item as *mut _ as isize));
}

unsafe fn mark_dirty_from_edit(hwnd: HWND, hwnd_edit: HWND) {
    let _ = with_state(hwnd, |state| {
        for (i, doc) in state.docs.iter_mut().enumerate() {
            if doc.hwnd_edit == hwnd_edit && !doc.dirty {
                doc.dirty = true;
                update_tab_title(state.hwnd_tab, i, &doc.title, true);
                update_window_title(hwnd);
                break;
            }
        }
    });
}

unsafe fn update_window_title(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(doc) = state.docs.get(state.current) {
            let suffix = if doc.dirty { "*" } else { "" };
            let untitled = untitled_base(state.settings.language);
            let display_title = if doc.title.starts_with(untitled) {
                untitled
            } else {
                doc.title.as_str()
            };
            let title = format!("{display_title}{suffix} - Novapad");
            let _ = SetWindowTextW(hwnd, PCWSTR(to_wide(&title).as_ptr()));
        }
    });
}

unsafe fn layout_children(hwnd: HWND) {
    let state_data = with_state(hwnd, |state| {
        (state.hwnd_tab, state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>())
    });
    let Some((hwnd_tab, edit_handles)) = state_data else {
        return;
    };
    let mut rc = RECT::default();
    if GetClientRect(hwnd, &mut rc).is_err() {
        return;
    }
    let width = rc.right - rc.left;
    let height = rc.bottom - rc.top;
    let _ = MoveWindow(hwnd_tab, 0, 0, width, height, true);

    let mut display = rc;
    SendMessageW(
        hwnd_tab,
        TCM_ADJUSTRECT,
        WPARAM(0),
        LPARAM(&mut display as *mut _ as isize),
    );
    let edit_width = display.right - display.left;
    let edit_height = display.bottom - display.top;
    for hwnd_edit in edit_handles {
        let _ = MoveWindow(
            hwnd_edit,
            display.left,
            display.top,
            edit_width,
            edit_height,
            true,
        );
    }
}

unsafe fn create_edit(parent: HWND, hfont: HFONT, word_wrap: bool) -> HWND {
    let mut base_style = WS_CHILD.0 | WS_VSCROLL.0 | ES_MULTILINE as u32 | ES_AUTOVSCROLL as u32 | ES_WANTRETURN as u32;

    if !word_wrap {
        base_style |= WS_HSCROLL.0 | ES_AUTOHSCROLL as u32;
    }

    let style = WINDOW_STYLE(base_style);

    let hwnd_edit = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        MSFTEDIT_CLASS,
        PCWSTR::null(),
        style,
        0,
        0,
        0,
        0,
        parent,
        HMENU(0),
        HINSTANCE(0),
        None,
    );
    SendMessageW(hwnd_edit, EM_LIMITTEXT, WPARAM(0x7FFFFFFE), LPARAM(0));
    SendMessageW(hwnd_edit, EM_SETEVENTMASK, WPARAM(0), LPARAM(ENM_CHANGE as isize));
    SendMessageW(hwnd_edit, EM_SETREADONLY, WPARAM(0), LPARAM(0));
    SendMessageW(hwnd_edit, EM_SETMODIFY, WPARAM(0), LPARAM(0));
    SendMessageW(
        hwnd_edit,
        windows::Win32::UI::WindowsAndMessaging::WM_SETFONT,
        WPARAM(hfont.0 as usize),
        LPARAM(1),
    );
    ShowWindow(hwnd_edit, SW_HIDE);
    hwnd_edit
}

pub(crate) unsafe fn apply_word_wrap_to_all_edits(hwnd: HWND, word_wrap: bool) {
    let edits = with_state(hwnd, |state| state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>())
        .unwrap_or_default();

    for edit in edits {
        apply_word_wrap_to_edit(edit, word_wrap);
    }
}

unsafe fn apply_word_wrap_to_edit(hwnd_edit: HWND, word_wrap: bool) {
    use windows::Win32::UI::WindowsAndMessaging::{
        GetWindowLongPtrW, SetWindowLongPtrW, SetWindowPos, GWL_STYLE, SWP_FRAMECHANGED, SWP_NOMOVE,
        SWP_NOSIZE, SWP_NOZORDER,
    };

    let _ = SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(0), LPARAM(0));

    let mut style = GetWindowLongPtrW(hwnd_edit, GWL_STYLE) as u32;

    if word_wrap {
        style &= !(WS_HSCROLL.0);
        style &= !(ES_AUTOHSCROLL as u32);
    } else {
        style |= WS_HSCROLL.0;
        style |= ES_AUTOHSCROLL as u32;
    }

    let _ = SetWindowLongPtrW(hwnd_edit, GWL_STYLE, style as isize);
    let _ = SetWindowPos(
        hwnd_edit,
        HWND(0),
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED,
    );

    let _ = SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(1), LPARAM(0));
    let _ = InvalidateRect(hwnd_edit, None, BOOL(1));
    let _ = UpdateWindow(hwnd_edit);
}


unsafe fn send_to_active_edit(hwnd: HWND, msg: u32) {
    let _ = with_state(hwnd, |state| {
        if let Some(doc) = state.docs.get(state.current) {
            SendMessageW(doc.hwnd_edit, msg, WPARAM(0), LPARAM(0));
        }
    });
}

unsafe fn select_all_active_edit(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(doc) = state.docs.get(state.current) {
            SendMessageW(doc.hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(-1));
        }
    });
}

unsafe fn set_edit_text(hwnd_edit: HWND, text: &str) {
    let wide = to_wide_normalized(text);
    
    SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(0), LPARAM(0));
    let _ = SetWindowTextW(hwnd_edit, PCWSTR(wide.as_ptr()));
    SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(1), LPARAM(0));
    
    let _ = InvalidateRect(hwnd_edit, None, BOOL(1));
    let _ = UpdateWindow(hwnd_edit);
    SendMessageW(hwnd_edit, EM_SETMODIFY, WPARAM(0), LPARAM(0));
}

pub(crate) unsafe fn get_edit_text(hwnd_edit: HWND) -> String {
    let len = SendMessageW(hwnd_edit, WM_GETTEXTLENGTH, WPARAM(0), LPARAM(0)).0 as usize;
    if len == 0 {
        return String::new();
    }
    let mut buffer = vec![0u16; len + 1];
    SendMessageW(
        hwnd_edit,
        WM_GETTEXT,
        WPARAM((len + 1) as usize),
        LPARAM(buffer.as_mut_ptr() as isize),
    );
    String::from_utf16_lossy(&buffer[..len])
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

fn abbreviate_recent_label(path: &Path) -> String {
    let filename = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
    let parent = path.parent().and_then(|p| p.to_str()).unwrap_or("");
    if parent.is_empty() {
        return filename.to_string();
    }
    let mut suffix = parent.to_string();
    if suffix.len() > 24 {
        suffix = format!("...{}", &suffix[suffix.len().saturating_sub(24)..]);
    }
    format!("{filename} - {suffix}")
}