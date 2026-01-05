#![allow(unsafe_op_in_unsafe_fn)]
#![allow(
    clippy::collapsible_if,
    clippy::manual_pattern_char_comparison,
    clippy::manual_range_contains,
    clippy::needless_return,
    clippy::unnecessary_cast,
    clippy::unnecessary_mut_passed
)]
#![windows_subsystem = "windows"]

mod accessibility;
use accessibility::*;
mod conpty;
mod settings;
use editor_manager::Document;
use settings::*;
mod bookmarks;
use bookmarks::*;
mod tts_engine;
use tts_engine::*;
mod file_handler;
mod mf_encoder;
mod sapi5_engine;
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
mod i18n;
mod podcast_recorder;
mod updater;

use std::io::Write;
use std::mem::size_of;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::Once;
use std::sync::atomic::AtomicBool;
use std::time::{Duration, Instant};

use chrono::Local;
use serde::{Deserialize, Serialize};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    COLOR_WINDOW, DEFAULT_GUI_FONT, GetStockObject, HBRUSH, HFONT,
};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::Win32::System::DataExchange::COPYDATASTRUCT;
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, LoadLibraryW};
use windows::Win32::UI::Controls::Dialogs::{
    FINDREPLACE_FLAGS, FINDREPLACEW, GetSaveFileNameW, OFN_EXPLORER, OFN_OVERWRITEPROMPT,
    OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::RichEdit::{
    CHARRANGE, EM_EXGETSEL, EM_EXSETSEL, EM_GETTEXTRANGE, TEXTRANGEW,
};
use windows::Win32::UI::Controls::{
    BST_CHECKED, ICC_TAB_CLASSES, INITCOMMONCONTROLSEX, InitCommonControlsEx, NMHDR, TCM_GETCURSEL,
    TCN_SELCHANGE, WC_BUTTON, WC_COMBOBOXW, WC_STATIC, WC_TABCONTROLW,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, GetKeyState, SetActiveWindow, SetFocus, VK_CONTROL, VK_ESCAPE, VK_F1,
    VK_F2, VK_F3, VK_F4, VK_F5, VK_F6, VK_RETURN, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows::Win32::UI::Shell::{
    DragAcceptFiles, DragFinish, DragQueryFileW, FileSaveDialog, HDROP, IFileDialog,
    IFileDialogControlEvents, IFileDialogControlEvents_Impl, IFileDialogCustomize,
    IFileDialogEvents, IFileDialogEvents_Impl, IFileSaveDialog, IShellItem,
};
use windows::Win32::UI::WindowsAndMessaging::{
    ACCEL, AppendMenuW, BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, CB_ADDSTRING, CB_GETCURSEL,
    CB_GETDROPPEDSTATE, CB_GETITEMDATA, CB_RESETCONTENT, CB_SETCURSEL, CB_SETITEMDATA,
    CBN_SELCHANGE, CBS_DROPDOWNLIST, CREATESTRUCTW, CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT,
    CheckMenuItem, CreateAcceleratorTableW, CreatePopupMenu, CreateWindowExW, DefWindowProcW,
    DestroyWindow, DispatchMessageW, EN_CHANGE, FALT, FCONTROL, FSHIFT, FVIRTKEY, FindWindowW,
    GWLP_USERDATA, GetCursorPos, GetMenu, GetMessageW, GetParent, GetWindowLongPtrW, HACCEL,
    HCURSOR, HICON, HMENU, IDC_ARROW, IDI_APPLICATION, KillTimer, LoadCursorW, LoadIconW,
    MB_ICONERROR, MB_ICONINFORMATION, MB_OK, MF_BYCOMMAND, MF_CHECKED, MF_STRING, MF_UNCHECKED,
    MSG, MessageBoxW, PostMessageW, PostQuitMessage, RegisterClassW, RegisterWindowMessageW,
    SW_HIDE, SW_SHOW, SendMessageW, SetForegroundWindow, SetTimer, SetWindowLongPtrW,
    SetWindowTextW, ShowWindow, TPM_RIGHTBUTTON, TrackPopupMenu, TranslateAcceleratorW,
    TranslateMessage, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU, WM_COPY,
    WM_COPYDATA, WM_CREATE, WM_CUT, WM_DESTROY, WM_DROPFILES, WM_KEYDOWN, WM_NCDESTROY,
    WM_NEXTDLGCTL, WM_NOTIFY, WM_NULL, WM_PASTE, WM_SETFOCUS, WM_SETFONT, WM_SIZE, WM_SYSKEYDOWN,
    WM_TIMER, WM_UNDO, WNDCLASSW, WS_CHILD, WS_CLIPCHILDREN, WS_EX_CLIENTEDGE, WS_OVERLAPPEDWINDOW,
    WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{Interface, PCWSTR, PWSTR, implement, w};

const EM_SCROLLCARET: u32 = 0x00B7;

use crate::app_windows::find_in_files_window::FindInFilesCache;

const WM_PDF_LOADED: u32 = WM_APP + 1;
const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
const WM_TTS_AUDIOBOOK_DONE: u32 = WM_APP + 4;
const WM_TTS_PLAYBACK_ERROR: u32 = WM_APP + 5;
const WM_UPDATE_PROGRESS: u32 = WM_APP + 6;
const WM_TTS_CHUNK_START: u32 = WM_APP + 7;
const WM_TTS_SAPI_VOICES_LOADED: u32 = WM_APP + 8;
pub const WM_FOCUS_EDITOR: u32 = WM_APP + 30;
const FOCUS_EDITOR_TIMER_ID: usize = 1;
const FOCUS_EDITOR_TIMER_ID2: usize = 2;
const COPYDATA_OPEN_FILE: usize = 1;
const VOICE_PANEL_ID_ENGINE: usize = 8001;
const VOICE_PANEL_ID_VOICE: usize = 8002;
const VOICE_PANEL_ID_MULTILINGUAL: usize = 8003;
const VOICE_PANEL_ID_FAVORITES: usize = 8004;
const VOICE_MENU_ID_ADD_FAVORITE: u32 = 9001;
const VOICE_MENU_ID_REMOVE_FAVORITE: u32 = 9002;

pub(crate) fn focus_editor(hwnd: HWND) {
    unsafe {
        SetForegroundWindow(hwnd);
        SetActiveWindow(hwnd);
        if let Some(hwnd_edit) = with_state(hwnd, |state| {
            state.docs.get(state.current).map(|doc| doc.hwnd_edit)
        })
        .flatten()
        {
            SetFocus(hwnd_edit);
            let _ = SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
            let _ = PostMessageW(hwnd, WM_NEXTDLGCTL, WPARAM(hwnd_edit.0 as usize), LPARAM(1));
        }
    }
}

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

const MAX_LOG_SIZE: u64 = 150 * 1024;

fn log_lock_path(log_path: &Path) -> Option<PathBuf> {
    let parent = log_path.parent()?;
    Some(parent.join("Novapad.log.lock"))
}

fn truncate_log_if_needed(path: &Path) {
    static LOG_INIT: Once = Once::new();
    LOG_INIT.call_once(|| {
        let Some(lock_path) = log_lock_path(path) else {
            return;
        };
        let start = Instant::now();
        let mut lock_acquired = false;
        while start.elapsed() < Duration::from_millis(200) {
            match std::fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&lock_path)
            {
                Ok(mut file) => {
                    let _ = writeln!(file, "{}", std::process::id());
                    lock_acquired = true;
                    break;
                }
                Err(err) if err.kind() == std::io::ErrorKind::AlreadyExists => {
                    std::thread::sleep(Duration::from_millis(50));
                }
                Err(_) => {
                    break;
                }
            }
        }
        if lock_acquired {
            let needs_truncate = path.metadata().ok().map(|m| m.len() > MAX_LOG_SIZE) == Some(true);
            if needs_truncate {
                let mut truncated = false;
                if let Ok(mut file) = std::fs::OpenOptions::new()
                    .create(true)
                    .write(true)
                    .truncate(true)
                    .open(path)
                {
                    let _ = writeln!(file, "[INFO] log truncated (exceeded 150 KB)");
                    truncated = true;
                }
                if !truncated {
                    let _ = std::fs::remove_file(path);
                    if let Ok(mut file) = std::fs::OpenOptions::new()
                        .create(true)
                        .write(true)
                        .truncate(true)
                        .open(path)
                    {
                        let _ = writeln!(file, "[INFO] log truncated (exceeded 150 KB)");
                    }
                }
            }
            let _ = std::fs::remove_file(&lock_path);
        }
    });
}

pub(crate) fn log_debug(message: &str) {
    let Some(path) = log_path() else {
        return;
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    truncate_log_if_needed(&path);
    if let Ok(mut log) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
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
    changelog_window: HWND,
    bookmarks_window: HWND,
    dictionary_window: HWND,
    dictionary_entry_dialog: HWND,
    tts_tuning_dialog: HWND,
    prompt_window: HWND,
    podcast_window: HWND,
    podcast_save_window: HWND,
    batch_audiobooks_window: HWND,
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
    tts_last_offset: i32,
    edge_voices: Vec<VoiceInfo>,
    sapi_voices: Vec<VoiceInfo>,
    audiobook_progress: HWND,
    audiobook_cancel: Option<Arc<AtomicBool>>,
    active_audiobook: Option<AudiobookPlayer>,
    voice_panel_visible: bool,
    voice_label_engine: HWND,
    voice_combo_engine: HWND,
    voice_label_voice: HWND,
    voice_combo_voice: HWND,
    voice_checkbox_multilingual: HWND,
    voice_favorites_visible: bool,
    voice_label_favorites: HWND,
    voice_combo_favorites: HWND,
    voice_context_voice: Option<FavoriteVoice>,
    find_in_files_cache: Option<FindInFilesCache>,
    normalize_undo: Option<NormalizeUndo>,
    normalize_skip_change: bool,
}

#[derive(Default, Serialize, Deserialize)]
struct RecentFileStore {
    files: Vec<String>,
}

fn main() -> windows::core::Result<()> {
    log_debug("Application started.");

    let args: Vec<String> = std::env::args().collect();
    if args.iter().any(|arg| arg == "--self-update") {
        match updater::run_self_update(&args) {
            Ok(code) => std::process::exit(code),
            Err(err) => {
                log_debug(&format!("Self-update failed: {err}"));
                std::process::exit(2);
            }
        }
    }
    updater::cleanup_backup_on_start();
    updater::cleanup_update_lock_on_start();
    updater::cleanup_update_temp_on_start();

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
        updater::check_pending_update(hwnd, false);

        let current_version = env!("CARGO_PKG_VERSION");
        let mut show_changelog = false;
        let _ = with_state(hwnd, |state| {
            let last_seen = state.settings.last_seen_changelog_version.clone();
            if last_seen.is_empty() {
                state.settings.last_seen_changelog_version = current_version.to_string();
                save_settings(state.settings.clone());
                return;
            }
            if last_seen != current_version {
                state.settings.last_seen_changelog_version = current_version.to_string();
                save_settings(state.settings.clone());
                show_changelog = true;
            }
        });
        if show_changelog {
            app_windows::help_window::open_changelog(hwnd);
        }

        let check_updates =
            with_state(hwnd, |state| state.settings.check_updates_on_startup).unwrap_or(true);
        if check_updates {
            updater::check_for_update(hwnd, false);
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
            if msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN {
                if msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                    let save_hwnd =
                        with_state(hwnd, |state| state.podcast_save_window).unwrap_or(HWND(0));
                    if save_hwnd.0 != 0 {
                        let _ = PostMessageW(save_hwnd, WM_COMMAND, WPARAM(2), LPARAM(0));
                        continue;
                    }
                }
            }
            if msg.message == WM_SYSKEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F4.0) {
                let (prompt_hwnd, prompt_open) = with_state(hwnd, |state| {
                    (state.prompt_window, state.prompt_window.0 != 0)
                })
                .unwrap_or((HWND(0), false));
                if prompt_open {
                    let target = msg.hwnd;
                    let target_parent = GetParent(target);
                    let prompt_target = target == prompt_hwnd || target_parent == prompt_hwnd;
                    let main_target = target == hwnd || target_parent == hwnd;
                    if main_target && !prompt_target {
                        editor_manager::close_current_document(hwnd);
                        continue;
                    }
                }
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == 'Z' as u32 {
                let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
                let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
                if ctrl_down && !shift_down {
                    if let Some(hwnd_edit) = get_active_edit(hwnd) {
                        if GetFocus() == hwnd_edit {
                            if !editor_manager::try_normalize_undo(hwnd) {
                                SendMessageW(hwnd_edit, WM_UNDO, WPARAM(0), LPARAM(0));
                            }
                            continue;
                        }
                    }
                }
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F1.0) {
                if is_tts_active(hwnd) {
                    cycle_favorite_voice(hwnd, -1);
                    continue;
                }
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F2.0) {
                if is_tts_active(hwnd) {
                    cycle_favorite_voice(hwnd, 1);
                    continue;
                }
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
                if (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) == 0 {
                    if handle_voice_panel_tab(hwnd) {
                        continue;
                    }
                }
            }

            let mut handled = false;
            let _ = with_state(hwnd, |state| {
                // Audiobook keyboard controls (ONLY if no secondary window is open)
                if msg.message == WM_KEYDOWN {
                    let is_audiobook = state
                        .docs
                        .get(state.current)
                        .map(|d| matches!(d.format, FileFormat::Audiobook))
                        .unwrap_or(false);
                    let secondary_open = state.bookmarks_window.0 != 0
                        || state.options_dialog.0 != 0
                        || state.help_window.0 != 0
                        || state.changelog_window.0 != 0
                        || state.dictionary_window.0 != 0
                        || state.podcast_window.0 != 0;
                    let secondary_open = secondary_open
                        || state.dictionary_entry_dialog.0 != 0
                        || state.tts_tuning_dialog.0 != 0;

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
                if state.replace_dialog.0 != 0 && handle_accessibility(state.replace_dialog, &msg) {
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
                if state.changelog_window.0 != 0 {
                    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
                        app_windows::help_window::handle_tab(state.changelog_window);
                        handled = true;
                        return;
                    }

                    if handle_accessibility(state.changelog_window, &msg) {
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

                if state.podcast_window.0 != 0 {
                    if app_windows::podcast_window::handle_navigation(state.podcast_window, &msg) {
                        handled = true;
                        return;
                    }
                }

                if state.podcast_save_window.0 != 0 {
                    if app_windows::podcast_save_window::handle_navigation(
                        state.podcast_save_window,
                        &msg,
                    ) {
                        handled = true;
                        return;
                    }
                }

                if state.audiobook_progress.0 != 0 {
                    if app_windows::audiobook_window::handle_navigation(
                        state.audiobook_progress,
                        &msg,
                    ) {
                        handled = true;
                        return;
                    }
                }

                if state.bookmarks_window.0 != 0 {
                    if app_windows::bookmarks_window::handle_navigation(
                        state.bookmarks_window,
                        &msg,
                    ) {
                        handled = true;
                        return;
                    }
                }

                if state.dictionary_window.0 != 0 {
                    if app_windows::dictionary_window::handle_navigation(
                        state.dictionary_window,
                        &msg,
                    ) {
                        handled = true;
                        return;
                    }
                }

                if state.dictionary_entry_dialog.0 != 0 {
                    if handle_accessibility(state.dictionary_entry_dialog, &msg) {
                        handled = true;
                        return;
                    }
                }
                if state.batch_audiobooks_window.0 != 0 {
                    if app_windows::batch_audiobooks_window::handle_navigation(
                        state.batch_audiobooks_window,
                        &msg,
                    ) {
                        handled = true;
                        return;
                    }
                    if handle_accessibility(state.batch_audiobooks_window, &msg) {
                        handled = true;
                        return;
                    }
                }

                if state.tts_tuning_dialog.0 != 0 {
                    if app_windows::tts_tuning_window::handle_navigation(
                        state.tts_tuning_dialog,
                        &msg,
                    ) {
                        handled = true;
                        return;
                    }
                }

                if state.prompt_window.0 != 0 {
                    if app_windows::prompt_window::handle_navigation(state.prompt_window, &msg) {
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
            let panel_labels = voice_panel_labels(settings.language);
            let _panel_labels = panel_labels;
            let empty_label = to_wide("");
            let label_engine = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(empty_label.as_ptr()),
                WS_CHILD,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_engine = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                140,
                hwnd,
                HMENU(VOICE_PANEL_ID_ENGINE as isize),
                HINSTANCE(0),
                None,
            );
            let label_voice = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(empty_label.as_ptr()),
                WS_CHILD,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_voice = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                160,
                hwnd,
                HMENU(VOICE_PANEL_ID_VOICE as isize),
                HINSTANCE(0),
                None,
            );
            let checkbox_multilingual = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(empty_label.as_ptr()),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(VOICE_PANEL_ID_MULTILINGUAL as isize),
                HINSTANCE(0),
                None,
            );
            let label_favorites = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(empty_label.as_ptr()),
                WS_CHILD,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let combo_favorites = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                160,
                hwnd,
                HMENU(VOICE_PANEL_ID_FAVORITES as isize),
                HINSTANCE(0),
                None,
            );
            for control in [
                label_engine,
                combo_engine,
                label_voice,
                combo_voice,
                checkbox_multilingual,
                label_favorites,
                combo_favorites,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
                ShowWindow(control, SW_HIDE);
            }
            let state = Box::new(AppState {
                hwnd_tab,
                docs: Vec::new(),
                current: 0,
                untitled_count: 0,
                hfont,
                hmenu_recent: recent_menu,
                recent_files,
                settings: settings.clone(),
                bookmarks,
                find_dialog: HWND(0),
                replace_dialog: HWND(0),
                options_dialog: HWND(0),
                help_window: HWND(0),
                changelog_window: HWND(0),
                bookmarks_window: HWND(0),
                dictionary_window: HWND(0),
                dictionary_entry_dialog: HWND(0),
                tts_tuning_dialog: HWND(0),
                prompt_window: HWND(0),
                podcast_window: HWND(0),
                podcast_save_window: HWND(0),
                batch_audiobooks_window: HWND(0),
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
                tts_last_offset: 0,
                edge_voices: Vec::new(),
                sapi_voices: Vec::new(),
                audiobook_progress: HWND(0),
                audiobook_cancel: None,
                active_audiobook: None,
                voice_panel_visible: false,
                voice_label_engine: label_engine,
                voice_combo_engine: combo_engine,
                voice_label_voice: label_voice,
                voice_combo_voice: combo_voice,
                voice_checkbox_multilingual: checkbox_multilingual,
                voice_favorites_visible: false,
                voice_label_favorites: label_favorites,
                voice_combo_favorites: combo_favorites,
                voice_context_voice: None,
                find_in_files_cache: None,
                normalize_undo: None,
                normalize_skip_change: false,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            update_recent_menu(hwnd, recent_menu);
            if settings.show_voice_panel {
                set_voice_panel_visible_internal(hwnd, true, false);
            }
            if settings.show_favorite_panel {
                set_favorites_panel_visible_internal(hwnd, true, false);
            }

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
            editor_manager::apply_text_limit_to_all_edits(hwnd);
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
                        unsafe {
                            SetFocus(state.hwnd_tab);
                        }
                    } else {
                        unsafe {
                            SetFocus(doc.hwnd_edit);
                        }
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
            if wparam.0 == FOCUS_EDITOR_TIMER_ID || wparam.0 == FOCUS_EDITOR_TIMER_ID2 {
                let _ = KillTimer(hwnd, wparam.0);
                focus_editor(hwnd);
                return LRESULT(0);
            }
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
                state.edge_voices = voices.clone();
            });
            if let Some(dialog) = with_state(hwnd, |state| state.options_dialog) {
                if dialog.0 != 0 {
                    app_windows::options_window::refresh_voices(dialog);
                }
            }
            refresh_voice_panel(hwnd);
            LRESULT(0)
        }
        WM_TTS_SAPI_VOICES_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = Box::from_raw(lparam.0 as *mut Vec<VoiceInfo>);
            let voices: Vec<VoiceInfo> = *payload;
            let _ = with_state(hwnd, |state| {
                state.sapi_voices = voices.clone();
            });
            if let Some(dialog) = with_state(hwnd, |state| state.options_dialog) {
                if dialog.0 != 0 {
                    app_windows::options_window::refresh_voices(dialog);
                }
            }
            refresh_voice_panel(hwnd);
            LRESULT(0)
        }
        WM_TTS_PLAYBACK_DONE => {
            let session_id = wparam.0 as u64;
            let _ = with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session {
                    if current.id == session_id {
                        state.tts_session = None;
                        state.tts_last_offset = 0;
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
                    if current.id == session_id {
                        state.tts_last_offset = offset;
                    }
                    if current.id == session_id && state.settings.move_cursor_during_reading {
                        if let Some(doc) = state.docs.get(state.current) {
                            let new_pos = current.initial_caret_pos + offset;
                            let mut cr = CHARRANGE {
                                cpMin: new_pos,
                                cpMax: new_pos,
                            };
                            unsafe {
                                SendMessageW(
                                    doc.hwnd_edit,
                                    EM_EXSETSEL,
                                    WPARAM(0),
                                    LPARAM(&mut cr as *mut _ as isize),
                                );
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
                        state.tts_last_offset = 0;
                        prevent_sleep(false);
                        should_show = true;
                    }
                }
            });
            if should_show {
                let language =
                    with_state(hwnd, |state| state.settings.language).unwrap_or_default();
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
            let title = if payload.success {
                audiobook_done_title(language)
            } else {
                error_title(language)
            };
            let title = to_wide(&title);
            let message = to_wide(&payload.message);
            let flags = if payload.success {
                MB_OK | MB_ICONINFORMATION
            } else {
                MB_OK | MB_ICONERROR
            };
            MessageBoxW(
                hwnd,
                PCWSTR(message.as_ptr()),
                PCWSTR(title.as_ptr()),
                flags,
            );
            LRESULT(0)
        }
        WM_FOCUS_EDITOR => {
            focus_editor(hwnd);
            let _ = SetTimer(hwnd, FOCUS_EDITOR_TIMER_ID, 80, None);
            let _ = SetTimer(hwnd, FOCUS_EDITOR_TIMER_ID2, 200, None);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == u32::from(VK_F1.0) {
                cycle_favorite_voice(hwnd, -1);
                return LRESULT(0);
            }
            if wparam.0 as u32 == u32::from(VK_F2.0) {
                cycle_favorite_voice(hwnd, 1);
                return LRESULT(0);
            }
            if wparam.0 as u32 == u32::from(VK_TAB.0)
                && (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0
            {
                next_tab_with_prompt(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CONTEXTMENU => {
            let target = HWND(wparam.0 as isize);
            let (combo_voice, combo_favorites) = with_state(hwnd, |state| {
                (state.voice_combo_voice, state.voice_combo_favorites)
            })
            .unwrap_or((HWND(0), HWND(0)));
            if (target == combo_voice && combo_voice.0 != 0)
                || (target == combo_favorites && combo_favorites.0 != 0)
            {
                show_voice_context_menu(hwnd, target, lparam);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            let notification = (wparam.0 >> 16) as u16;
            if u32::from(notification) == EN_CHANGE {
                editor_manager::handle_normalize_edit_change(hwnd, HWND(lparam.0));
                mark_dirty_from_edit(hwnd, HWND(lparam.0));
                return LRESULT(0);
            }
            if cmd_id == VOICE_PANEL_ID_ENGINE && u32::from(notification) == CBN_SELCHANGE {
                handle_voice_panel_engine_change(hwnd);
                return LRESULT(0);
            }
            if cmd_id == VOICE_PANEL_ID_VOICE && u32::from(notification) == CBN_SELCHANGE {
                handle_voice_panel_voice_change(hwnd);
                return LRESULT(0);
            }
            if cmd_id == VOICE_PANEL_ID_FAVORITES && u32::from(notification) == CBN_SELCHANGE {
                handle_voice_panel_favorite_change(hwnd);
                return LRESULT(0);
            }
            if cmd_id == VOICE_PANEL_ID_MULTILINGUAL {
                handle_voice_panel_multilingual_toggle(hwnd);
                return LRESULT(0);
            }
            if cmd_id == VOICE_MENU_ID_ADD_FAVORITE as usize {
                handle_voice_context_favorite(hwnd, true);
                return LRESULT(0);
            }
            if cmd_id == VOICE_MENU_ID_REMOVE_FAVORITE as usize {
                handle_voice_context_favorite(hwnd, false);
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
                    if let Some((path, encoding)) = open_file_dialog_with_encoding(hwnd) {
                        open_document_with_encoding(hwnd, &path, encoding);
                        if with_state(hwnd, |state| state.prompt_window.0 != 0).unwrap_or(false) {
                            focus_editor(hwnd);
                        }
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
                IDM_FILE_BATCH_AUDIOBOOK => {
                    log_debug("Menu: Batch audiobooks");
                    app_windows::batch_audiobooks_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_PODCAST => {
                    log_debug("Menu: Record podcast");
                    app_windows::podcast_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_UNDO => {
                    log_debug("Menu: Undo");
                    if !editor_manager::try_normalize_undo(hwnd) {
                        editor_manager::send_to_active_edit(hwnd, WM_UNDO);
                    }
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
                IDM_EDIT_FIND_IN_FILES => {
                    log_debug("Menu: Find in files");
                    app_windows::find_in_files_window::open_find_in_files_dialog(hwnd);
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
                IDM_EDIT_STRIP_MARKDOWN => {
                    log_debug("Menu: Strip Markdown");
                    editor_manager::strip_markdown_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_NORMALIZE_WHITESPACE => {
                    log_debug("Menu: Normalize whitespace");
                    editor_manager::normalize_whitespace_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_HARD_LINE_BREAK => {
                    log_debug("Menu: Hard line break");
                    editor_manager::hard_line_break_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_ORDER_ITEMS => {
                    log_debug("Menu: Order items");
                    editor_manager::order_items_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_KEEP_UNIQUE_ITEMS => {
                    log_debug("Menu: Keep unique items");
                    editor_manager::keep_unique_items_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_REVERSE_ITEMS => {
                    log_debug("Menu: Reverse items");
                    editor_manager::reverse_items_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_QUOTE_LINES => {
                    log_debug("Menu: Quote lines");
                    editor_manager::quote_lines_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_UNQUOTE_LINES => {
                    log_debug("Menu: Unquote lines");
                    editor_manager::unquote_lines_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_TEXT_STATS => {
                    log_debug("Menu: Text stats");
                    editor_manager::text_stats_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_JOIN_LINES => {
                    log_debug("Menu: Join lines");
                    editor_manager::join_lines_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_CLEAN_EOL_HYPHENS => {
                    log_debug("Menu: Clean EOL hyphens");
                    editor_manager::clean_end_of_line_hyphens_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_VIEW_SHOW_VOICES => {
                    log_debug("Menu: Toggle voice panel");
                    toggle_voice_panel(hwnd);
                    LRESULT(0)
                }
                IDM_VIEW_SHOW_FAVORITES => {
                    log_debug("Menu: Toggle favorite voices panel");
                    toggle_favorites_panel(hwnd);
                    LRESULT(0)
                }
                cmd_id if text_color_from_menu_id(cmd_id).is_some() => {
                    let color = text_color_from_menu_id(cmd_id);
                    update_text_preferences(hwnd, color, None);
                    LRESULT(0)
                }
                cmd_id if text_size_from_menu_id(cmd_id).is_some() => {
                    let size = text_size_from_menu_id(cmd_id);
                    update_text_preferences(hwnd, None, size);
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
                IDM_TOOLS_DICTIONARY => {
                    log_debug("Menu: Dictionary");
                    app_windows::dictionary_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_IMPORT_YOUTUBE => {
                    log_debug("Menu: Import YouTube transcript");
                    app_windows::youtube_transcript_window::import_youtube_transcript(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_PROMPT => {
                    log_debug("Menu: Prompt");
                    app_windows::prompt_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_GUIDE => {
                    log_debug("Menu: Guide");
                    app_windows::help_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_CHANGELOG => {
                    log_debug("Menu: Changelog");
                    app_windows::help_window::open_changelog(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_CHECK_UPDATES => {
                    log_debug("Menu: Check updates");
                    updater::check_for_update(hwnd, true);
                    LRESULT(0)
                }
                IDM_HELP_PENDING_UPDATE => {
                    log_debug("Menu: Pending update");
                    updater::check_pending_update(hwnd, true);
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
                    let _ = PostMessageW(hwnd, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
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

unsafe fn cycle_favorite_voice(hwnd: HWND, direction: i32) {
    let (favorites, current_engine, current_voice) = with_state(hwnd, |state| {
        (
            state.settings.favorite_voices.clone(),
            state.settings.tts_engine,
            state.settings.tts_voice.clone(),
        )
    })
    .unwrap_or((Vec::new(), TtsEngine::Edge, String::new()));
    if favorites.is_empty() {
        return;
    }
    let mut current_idx = favorites
        .iter()
        .position(|fav| fav.engine == current_engine && fav.short_name == current_voice);
    if current_idx.is_none() {
        current_idx = Some(if direction >= 0 {
            0
        } else {
            favorites.len().saturating_sub(1)
        });
    }
    let idx = current_idx.unwrap_or(0);
    let len = favorites.len() as i32;
    let mut next_idx = idx as i32 + direction;
    if next_idx < 0 {
        next_idx = len - 1;
    } else if next_idx >= len {
        next_idx = 0;
    }
    let Some(next_fav) = favorites.get(next_idx as usize).cloned() else {
        return;
    };
    if next_fav.engine == current_engine && next_fav.short_name == current_voice {
        return;
    }
    let _ = with_state(hwnd, |state| {
        state.settings.tts_engine = next_fav.engine;
        state.settings.tts_voice = next_fav.short_name.clone();
    });
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
    refresh_voice_panel(hwnd);
    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
        save_settings(settings);
    }
    restart_tts_from_current_offset(hwnd);
}

unsafe fn is_tts_active(hwnd: HWND) -> bool {
    with_state(hwnd, |state| state.tts_session.is_some()).unwrap_or(false)
}

struct VoicePanelLabels {
    label_engine: String,
    label_voice: String,
    label_favorites: String,
    label_multilingual: String,
    engine_edge: String,
    engine_sapi: String,
    voices_empty: String,
    favorites_empty: String,
    add_favorite: String,
    remove_favorite: String,
}

fn voice_panel_labels(language: Language) -> VoicePanelLabels {
    VoicePanelLabels {
        label_engine: i18n::tr(language, "voice_panel.label_engine"),
        label_voice: i18n::tr(language, "voice_panel.label_voice"),
        label_favorites: i18n::tr(language, "voice_panel.label_favorites"),
        label_multilingual: i18n::tr(language, "voice_panel.label_multilingual"),
        engine_edge: i18n::tr(language, "voice_panel.engine_edge"),
        engine_sapi: i18n::tr(language, "voice_panel.engine_sapi"),
        voices_empty: i18n::tr(language, "voice_panel.voices_empty"),
        favorites_empty: i18n::tr(language, "voice_panel.favorites_empty"),
        add_favorite: i18n::tr(language, "voice_panel.add_favorite"),
        remove_favorite: i18n::tr(language, "voice_panel.remove_favorite"),
    }
}

fn text_color_menu_id(text_color: u32) -> usize {
    match text_color {
        0x000000 => IDM_VIEW_TEXT_COLOR_BLACK,
        0x800000 => IDM_VIEW_TEXT_COLOR_DARK_BLUE,
        0x006400 => IDM_VIEW_TEXT_COLOR_DARK_GREEN,
        0x002850 => IDM_VIEW_TEXT_COLOR_DARK_BROWN,
        0x404040 => IDM_VIEW_TEXT_COLOR_DARK_GRAY,
        0xFFCC99 => IDM_VIEW_TEXT_COLOR_LIGHT_BLUE,
        0x99CC99 => IDM_VIEW_TEXT_COLOR_LIGHT_GREEN,
        0x99B2CC => IDM_VIEW_TEXT_COLOR_LIGHT_BROWN,
        0xC0C0C0 => IDM_VIEW_TEXT_COLOR_LIGHT_GRAY,
        _ => IDM_VIEW_TEXT_COLOR_BLACK,
    }
}

fn text_color_from_menu_id(cmd_id: usize) -> Option<u32> {
    match cmd_id {
        IDM_VIEW_TEXT_COLOR_BLACK => Some(0x000000),
        IDM_VIEW_TEXT_COLOR_DARK_BLUE => Some(0x800000),
        IDM_VIEW_TEXT_COLOR_DARK_GREEN => Some(0x006400),
        IDM_VIEW_TEXT_COLOR_DARK_BROWN => Some(0x002850),
        IDM_VIEW_TEXT_COLOR_DARK_GRAY => Some(0x404040),
        IDM_VIEW_TEXT_COLOR_LIGHT_BLUE => Some(0xFFCC99),
        IDM_VIEW_TEXT_COLOR_LIGHT_GREEN => Some(0x99CC99),
        IDM_VIEW_TEXT_COLOR_LIGHT_BROWN => Some(0x99B2CC),
        IDM_VIEW_TEXT_COLOR_LIGHT_GRAY => Some(0xC0C0C0),
        _ => None,
    }
}

fn text_size_menu_id(text_size: i32) -> usize {
    match text_size {
        10 => IDM_VIEW_TEXT_SIZE_SMALL,
        12 => IDM_VIEW_TEXT_SIZE_NORMAL,
        16 => IDM_VIEW_TEXT_SIZE_LARGE,
        20 => IDM_VIEW_TEXT_SIZE_XLARGE,
        24 => IDM_VIEW_TEXT_SIZE_XXLARGE,
        _ => IDM_VIEW_TEXT_SIZE_NORMAL,
    }
}

fn text_size_from_menu_id(cmd_id: usize) -> Option<i32> {
    match cmd_id {
        IDM_VIEW_TEXT_SIZE_SMALL => Some(10),
        IDM_VIEW_TEXT_SIZE_NORMAL => Some(12),
        IDM_VIEW_TEXT_SIZE_LARGE => Some(16),
        IDM_VIEW_TEXT_SIZE_XLARGE => Some(20),
        IDM_VIEW_TEXT_SIZE_XXLARGE => Some(24),
        _ => None,
    }
}

unsafe fn update_text_preferences(hwnd: HWND, text_color: Option<u32>, text_size: Option<i32>) {
    let mut changed = false;
    let mut next_color = None;
    let mut next_size = None;
    let _ = with_state(hwnd, |state| {
        if let Some(color) = text_color {
            if state.settings.text_color != color {
                state.settings.text_color = color;
                changed = true;
            }
            next_color = Some(state.settings.text_color);
        } else {
            next_color = Some(state.settings.text_color);
        }
        if let Some(size) = text_size {
            if state.settings.text_size != size {
                state.settings.text_size = size;
                changed = true;
            }
            next_size = Some(state.settings.text_size);
        } else {
            next_size = Some(state.settings.text_size);
        }
    });

    let (color, size) = match (next_color, next_size) {
        (Some(c), Some(s)) => (c, s),
        _ => return,
    };
    if changed {
        if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
            save_settings(settings);
        }
    }
    editor_manager::apply_text_appearance_to_all_edits(hwnd, color, size);
    update_voice_panel_menu_check(hwnd);
}

unsafe fn update_voice_panel_menu_check(hwnd: HWND) {
    let (visible, favorites_visible, text_color, text_size) = with_state(hwnd, |state| {
        (
            state.voice_panel_visible,
            state.voice_favorites_visible,
            state.settings.text_color,
            state.settings.text_size,
        )
    })
    .unwrap_or((false, false, 0x000000, 12));
    let hmenu = GetMenu(hwnd);
    if hmenu.0 == 0 {
        return;
    }
    let flags = if visible { MF_CHECKED } else { MF_UNCHECKED };
    let _ = CheckMenuItem(hmenu, IDM_VIEW_SHOW_VOICES as u32, (MF_BYCOMMAND | flags).0);
    let fav_flags = if favorites_visible {
        MF_CHECKED
    } else {
        MF_UNCHECKED
    };
    let _ = CheckMenuItem(
        hmenu,
        IDM_VIEW_SHOW_FAVORITES as u32,
        (MF_BYCOMMAND | fav_flags).0,
    );

    let color_items = [
        IDM_VIEW_TEXT_COLOR_BLACK,
        IDM_VIEW_TEXT_COLOR_DARK_BLUE,
        IDM_VIEW_TEXT_COLOR_DARK_GREEN,
        IDM_VIEW_TEXT_COLOR_DARK_BROWN,
        IDM_VIEW_TEXT_COLOR_DARK_GRAY,
        IDM_VIEW_TEXT_COLOR_LIGHT_BLUE,
        IDM_VIEW_TEXT_COLOR_LIGHT_GREEN,
        IDM_VIEW_TEXT_COLOR_LIGHT_BROWN,
        IDM_VIEW_TEXT_COLOR_LIGHT_GRAY,
    ];
    let selected_color = text_color_menu_id(text_color);
    for item in color_items {
        let item_flags = if item == selected_color {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
        let _ = CheckMenuItem(hmenu, item as u32, (MF_BYCOMMAND | item_flags).0);
    }

    let size_items = [
        IDM_VIEW_TEXT_SIZE_SMALL,
        IDM_VIEW_TEXT_SIZE_NORMAL,
        IDM_VIEW_TEXT_SIZE_LARGE,
        IDM_VIEW_TEXT_SIZE_XLARGE,
        IDM_VIEW_TEXT_SIZE_XXLARGE,
    ];
    let selected_size = text_size_menu_id(text_size);
    for item in size_items {
        let item_flags = if item == selected_size {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
        let _ = CheckMenuItem(hmenu, item as u32, (MF_BYCOMMAND | item_flags).0);
    }
}

unsafe fn toggle_voice_panel(hwnd: HWND) {
    let visible = with_state(hwnd, |state| state.voice_panel_visible).unwrap_or(false);
    set_voice_panel_visible(hwnd, !visible);
}

unsafe fn set_voice_panel_visible(hwnd: HWND, visible: bool) {
    set_voice_panel_visible_internal(hwnd, visible, true);
}

unsafe fn set_voice_panel_visible_internal(hwnd: HWND, visible: bool, persist: bool) {
    let (label_engine, combo_engine, label_voice, combo_voice, checkbox_multilingual, changed) =
        match with_state(hwnd, |state| {
            let changed = state.settings.show_voice_panel != visible;
            state.voice_panel_visible = visible;
            state.settings.show_voice_panel = visible;
            (
                state.voice_label_engine,
                state.voice_combo_engine,
                state.voice_label_voice,
                state.voice_combo_voice,
                state.voice_checkbox_multilingual,
                changed,
            )
        }) {
            Some(values) => values,
            None => return,
        };

    let show = if visible { SW_SHOW } else { SW_HIDE };
    for control in [
        label_engine,
        combo_engine,
        label_voice,
        combo_voice,
        checkbox_multilingual,
    ] {
        if control.0 != 0 {
            ShowWindow(control, show);
        }
    }
    update_voice_panel_menu_check(hwnd);
    if visible {
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
        refresh_voice_panel(hwnd);
    }
    if persist && changed {
        if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
            save_settings(settings);
        }
    }
    clear_voice_labels_if_hidden(hwnd);
    editor_manager::layout_children(hwnd);
}

unsafe fn toggle_favorites_panel(hwnd: HWND) {
    let visible = with_state(hwnd, |state| state.voice_favorites_visible).unwrap_or(false);
    set_favorites_panel_visible(hwnd, !visible);
}

unsafe fn set_favorites_panel_visible(hwnd: HWND, visible: bool) {
    set_favorites_panel_visible_internal(hwnd, visible, true);
}

unsafe fn set_favorites_panel_visible_internal(hwnd: HWND, visible: bool, persist: bool) {
    let (label_favorites, combo_favorites, changed) = match with_state(hwnd, |state| {
        let changed = state.settings.show_favorite_panel != visible;
        state.voice_favorites_visible = visible;
        state.settings.show_favorite_panel = visible;
        (
            state.voice_label_favorites,
            state.voice_combo_favorites,
            changed,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    let show = if visible { SW_SHOW } else { SW_HIDE };
    for control in [label_favorites, combo_favorites] {
        if control.0 != 0 {
            ShowWindow(control, show);
        }
    }
    update_voice_panel_menu_check(hwnd);
    if visible {
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
        refresh_voice_panel(hwnd);
    }
    if persist && changed {
        if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
            save_settings(settings);
        }
    }
    clear_voice_labels_if_hidden(hwnd);
    editor_manager::layout_children(hwnd);
}

pub(crate) unsafe fn refresh_voice_panel(hwnd: HWND) {
    let (
        voice_visible,
        label_engine,
        combo_engine,
        label_voice,
        combo_voice,
        checkbox_multilingual,
        favorites_visible,
        label_favorites,
        combo_favorites,
    ) = match with_state(hwnd, |state| {
        (
            state.voice_panel_visible,
            state.voice_label_engine,
            state.voice_combo_engine,
            state.voice_label_voice,
            state.voice_combo_voice,
            state.voice_checkbox_multilingual,
            state.voice_favorites_visible,
            state.voice_label_favorites,
            state.voice_combo_favorites,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    if !voice_visible && !favorites_visible {
        return;
    }

    let settings = with_state(hwnd, |state| state.settings.clone()).unwrap_or_default();
    let labels = voice_panel_labels(settings.language);
    if voice_visible {
        let label_engine_wide = to_wide(&labels.label_engine);
        let label_voice_wide = to_wide(&labels.label_voice);
        let _ = SetWindowTextW(label_engine, PCWSTR(label_engine_wide.as_ptr()));
        let _ = SetWindowTextW(label_voice, PCWSTR(label_voice_wide.as_ptr()));
        let label_multi_wide = to_wide(&labels.label_multilingual);
        let _ = SetWindowTextW(checkbox_multilingual, PCWSTR(label_multi_wide.as_ptr()));
    }
    if favorites_visible && label_favorites.0 != 0 {
        let label_fav_wide = to_wide(&labels.label_favorites);
        let _ = SetWindowTextW(label_favorites, PCWSTR(label_fav_wide.as_ptr()));
    }

    if voice_visible && combo_engine.0 != 0 && combo_voice.0 != 0 {
        let _ = SendMessageW(combo_engine, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(
            combo_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_edge).as_ptr() as isize),
        );
        let _ = SendMessageW(
            combo_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi).as_ptr() as isize),
        );
        let engine_index = match settings.tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
        };
        let _ = SendMessageW(combo_engine, CB_SETCURSEL, WPARAM(engine_index), LPARAM(0));
        let is_edge = matches!(settings.tts_engine, TtsEngine::Edge);
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
        EnableWindow(checkbox_multilingual, is_edge);
        let multi_show = if is_edge { SW_SHOW } else { SW_HIDE };
        ShowWindow(checkbox_multilingual, multi_show);
    }

    if voice_visible {
        let voices = with_state(hwnd, |state| match settings.tts_engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
        })
        .unwrap_or_default();
        populate_voice_panel_combo(
            combo_voice,
            &voices,
            &settings.tts_voice,
            settings.tts_only_multilingual,
            &labels.voices_empty,
        );
    }
    if favorites_visible {
        populate_favorites_combo(
            combo_favorites,
            &settings.favorite_voices,
            settings.tts_engine,
            &settings.tts_voice,
            &labels,
        );
    }
}

unsafe fn refresh_voice_panel_voice_list(hwnd: HWND) {
    let (voice_visible, combo_voice, checkbox_multilingual) = match with_state(hwnd, |state| {
        (
            state.voice_panel_visible,
            state.voice_combo_voice,
            state.voice_checkbox_multilingual,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    if !voice_visible || combo_voice.0 == 0 {
        return;
    }

    let settings = with_state(hwnd, |state| state.settings.clone()).unwrap_or_default();
    let labels = voice_panel_labels(settings.language);
    let is_edge = matches!(settings.tts_engine, TtsEngine::Edge);
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
    EnableWindow(checkbox_multilingual, is_edge);
    let multi_show = if is_edge { SW_SHOW } else { SW_HIDE };
    ShowWindow(checkbox_multilingual, multi_show);

    let voices = with_state(hwnd, |state| match settings.tts_engine {
        TtsEngine::Edge => state.edge_voices.clone(),
        TtsEngine::Sapi5 => state.sapi_voices.clone(),
    })
    .unwrap_or_default();
    populate_voice_panel_combo(
        combo_voice,
        &voices,
        &settings.tts_voice,
        settings.tts_only_multilingual,
        &labels.voices_empty,
    );
}

unsafe fn clear_voice_labels_if_hidden(hwnd: HWND) {
    let (
        voice_visible,
        favorites_visible,
        label_engine,
        label_voice,
        checkbox_multilingual,
        label_favorites,
    ) = match with_state(hwnd, |state| {
        (
            state.voice_panel_visible,
            state.voice_favorites_visible,
            state.voice_label_engine,
            state.voice_label_voice,
            state.voice_checkbox_multilingual,
            state.voice_label_favorites,
        )
    }) {
        Some(values) => values,
        None => return,
    };
    if voice_visible || favorites_visible {
        return;
    }
    let empty = to_wide("");
    let _ = SetWindowTextW(label_engine, PCWSTR(empty.as_ptr()));
    let _ = SetWindowTextW(label_voice, PCWSTR(empty.as_ptr()));
    let _ = SetWindowTextW(checkbox_multilingual, PCWSTR(empty.as_ptr()));
    let _ = SetWindowTextW(label_favorites, PCWSTR(empty.as_ptr()));
}

unsafe fn populate_voice_panel_combo(
    combo_voice: HWND,
    voices: &[VoiceInfo],
    selected: &str,
    only_multilingual: bool,
    empty_label: &str,
) {
    let _ = SendMessageW(combo_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    if voices.is_empty() {
        let _ = SendMessageW(
            combo_voice,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(empty_label).as_ptr() as isize),
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
        let idx = SendMessageW(
            combo_voice,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&label).as_ptr() as isize),
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

unsafe fn populate_favorites_combo(
    combo_favorites: HWND,
    favorites: &[FavoriteVoice],
    selected_engine: TtsEngine,
    selected_voice: &str,
    labels: &VoicePanelLabels,
) {
    let _ = SendMessageW(combo_favorites, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    if favorites.is_empty() {
        let _ = SendMessageW(
            combo_favorites,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.favorites_empty).as_ptr() as isize),
        );
        let _ = SendMessageW(combo_favorites, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        return;
    }
    let mut selected_index: Option<usize> = None;
    for (idx, fav) in favorites.iter().enumerate() {
        let engine_label = match fav.engine {
            TtsEngine::Edge => &labels.engine_edge,
            TtsEngine::Sapi5 => &labels.engine_sapi,
        };
        let label = format!("{} ({})", fav.short_name, engine_label);
        let cb_idx = SendMessageW(
            combo_favorites,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&label).as_ptr() as isize),
        )
        .0;
        if cb_idx >= 0 {
            let _ = SendMessageW(
                combo_favorites,
                CB_SETITEMDATA,
                WPARAM(cb_idx as usize),
                LPARAM(idx as isize),
            );
            if fav.short_name == selected_voice && fav.engine == selected_engine {
                selected_index = Some(cb_idx as usize);
            }
        }
    }
    if let Some(idx) = selected_index {
        let _ = SendMessageW(combo_favorites, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
    } else {
        let _ = SendMessageW(combo_favorites, CB_SETCURSEL, WPARAM(0), LPARAM(0));
    }
}

unsafe fn handle_voice_panel_engine_change(hwnd: HWND) {
    let (combo_engine, language) = match with_state(hwnd, |state| {
        (state.voice_combo_engine, state.settings.language)
    }) {
        Some(values) => values,
        None => return,
    };
    let sel = SendMessageW(combo_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let new_engine = if sel == 1 {
        TtsEngine::Sapi5
    } else {
        TtsEngine::Edge
    };
    let (old_engine, old_voice) = with_state(hwnd, |state| {
        (state.settings.tts_engine, state.settings.tts_voice.clone())
    })
    .unwrap_or((TtsEngine::Edge, String::new()));
    let _ = with_state(hwnd, |state| {
        state.settings.tts_engine = new_engine;
    });
    app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
    refresh_voice_panel(hwnd);
    let mut new_voice = old_voice.clone();
    if let Some(voice_name) = current_voice_selection(hwnd, new_engine) {
        let _ = with_state(hwnd, |state| {
            state.settings.tts_voice = voice_name.clone();
        });
        new_voice = voice_name;
    }
    let changed = new_engine != old_engine || new_voice != old_voice;
    if changed {
        if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
            save_settings(settings);
        }
        restart_tts_from_current_offset(hwnd);
    }
}

unsafe fn handle_voice_panel_voice_change(hwnd: HWND) {
    let engine = with_state(hwnd, |state| state.settings.tts_engine).unwrap_or_default();
    if let Some(voice_name) = current_voice_selection(hwnd, engine) {
        let old_voice =
            with_state(hwnd, |state| state.settings.tts_voice.clone()).unwrap_or_default();
        if voice_name != old_voice {
            let _ = with_state(hwnd, |state| {
                state.settings.tts_voice = voice_name;
            });
            if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
                save_settings(settings);
            }
            restart_tts_from_current_offset(hwnd);
        }
    }
}

unsafe fn handle_voice_panel_multilingual_toggle(hwnd: HWND) {
    let (checkbox, is_edge) = with_state(hwnd, |state| {
        (
            state.voice_checkbox_multilingual,
            matches!(state.settings.tts_engine, TtsEngine::Edge),
        )
    })
    .unwrap_or((HWND(0), false));
    if checkbox.0 == 0 {
        return;
    }
    if !is_edge {
        return;
    }
    let checked =
        SendMessageW(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    let _ = with_state(hwnd, |state| {
        state.settings.tts_only_multilingual = checked;
    });
    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
        save_settings(settings);
    }
    refresh_voice_panel_voice_list(hwnd);
}

unsafe fn handle_voice_panel_favorite_change(hwnd: HWND) {
    let (combo_favorites, favorites) = with_state(hwnd, |state| {
        (
            state.voice_combo_favorites,
            state.settings.favorite_voices.clone(),
        )
    })
    .unwrap_or((HWND(0), Vec::new()));
    if combo_favorites.0 == 0 || favorites.is_empty() {
        return;
    }
    let sel = SendMessageW(combo_favorites, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel < 0 {
        return;
    }
    let fav_idx = SendMessageW(
        combo_favorites,
        CB_GETITEMDATA,
        WPARAM(sel as usize),
        LPARAM(0),
    )
    .0 as usize;
    let Some(fav) = favorites.get(fav_idx).cloned() else {
        return;
    };
    let (old_engine, old_voice) = with_state(hwnd, |state| {
        (state.settings.tts_engine, state.settings.tts_voice.clone())
    })
    .unwrap_or((TtsEngine::Edge, String::new()));
    if fav.engine == old_engine && fav.short_name == old_voice {
        return;
    }
    let _ = with_state(hwnd, |state| {
        state.settings.tts_engine = fav.engine;
        state.settings.tts_voice = fav.short_name.clone();
    });
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
    refresh_voice_panel(hwnd);
    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
        save_settings(settings);
    }
    restart_tts_from_current_offset(hwnd);
}

unsafe fn current_voice_selection(hwnd: HWND, engine: TtsEngine) -> Option<String> {
    let (combo_voice, voices) = with_state(hwnd, |state| {
        let list = match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
        };
        (state.voice_combo_voice, list)
    })?;
    if voices.is_empty() || combo_voice.0 == 0 {
        return None;
    }
    let sel = SendMessageW(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel < 0 {
        return None;
    }
    let voice_index =
        SendMessageW(combo_voice, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as usize;
    voices.get(voice_index).map(|v| v.short_name.clone())
}

unsafe fn current_favorite_selection(hwnd: HWND) -> Option<FavoriteVoice> {
    let (combo_favorites, favorites) = with_state(hwnd, |state| {
        (
            state.voice_combo_favorites,
            state.settings.favorite_voices.clone(),
        )
    })?;
    if combo_favorites.0 == 0 || favorites.is_empty() {
        return None;
    }
    let sel = SendMessageW(combo_favorites, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel < 0 {
        return None;
    }
    let fav_idx = SendMessageW(
        combo_favorites,
        CB_GETITEMDATA,
        WPARAM(sel as usize),
        LPARAM(0),
    )
    .0 as usize;
    favorites.get(fav_idx).cloned()
}

pub(crate) unsafe fn restart_tts_from_current_offset(hwnd: HWND) {
    let mut restart = None;
    let _ = with_state(hwnd, |state| {
        if let Some(session) = &state.tts_session {
            if let Some(doc) = state.docs.get(state.current) {
                if matches!(doc.format, FileFormat::Audiobook) {
                    return;
                }
                let pos = (session.initial_caret_pos + state.tts_last_offset).max(0);
                restart = Some((doc.hwnd_edit, pos));
            }
        }
    });
    let Some((hwnd_edit, pos)) = restart else {
        return;
    };
    tts_engine::stop_tts_playback(hwnd);
    let pos = adjust_tts_restart_pos(hwnd_edit, pos);
    let mut cr = CHARRANGE {
        cpMin: pos,
        cpMax: pos,
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );
    tts_engine::start_tts_from_caret(hwnd);
}

unsafe fn adjust_tts_restart_pos(hwnd_edit: HWND, pos: i32) -> i32 {
    if pos <= 0 {
        return 0;
    }
    let text = editor_manager::get_edit_text(hwnd_edit);
    if text.is_empty() {
        return pos;
    }
    let normalized = text.replace("\r\n", "\n");
    let mut items: Vec<(usize, usize, bool)> = Vec::new();
    let mut offset = 0usize;
    for ch in normalized.chars() {
        let start = offset;
        let len = ch.len_utf16();
        let end = start + len;
        let is_word = ch.is_alphanumeric() || ch == '_';
        items.push((start, end, is_word));
        offset = end;
    }
    if offset == 0 {
        return pos;
    }
    let mut pos_usize = pos as usize;
    if pos_usize > offset {
        pos_usize = offset;
    }

    let mut prev: Option<usize> = None;
    let mut next: Option<usize> = None;
    for (idx, (start, end, _)) in items.iter().enumerate() {
        if *end <= pos_usize {
            prev = Some(idx);
            continue;
        }
        if *start >= pos_usize {
            next = Some(idx);
            break;
        }
        next = Some(idx);
        break;
    }

    let prev_is_word = prev
        .and_then(|idx| items.get(idx))
        .map(|v| v.2)
        .unwrap_or(false);
    let next_is_word = next
        .and_then(|idx| items.get(idx))
        .map(|v| v.2)
        .unwrap_or(false);
    if prev_is_word && next_is_word {
        let mut idx = prev.unwrap();
        while idx > 0 && items[idx - 1].2 {
            idx -= 1;
        }
        return items[idx].0 as i32;
    }
    pos
}

unsafe fn show_voice_context_menu(hwnd: HWND, target: HWND, lparam: LPARAM) {
    let (combo_voice, combo_favorites, engine, language) = with_state(hwnd, |state| {
        (
            state.voice_combo_voice,
            state.voice_combo_favorites,
            state.settings.tts_engine,
            state.settings.language,
        )
    })
    .unwrap_or((HWND(0), HWND(0), TtsEngine::Edge, Language::Italian));
    let labels = voice_panel_labels(language);

    let mut action_id = VOICE_MENU_ID_ADD_FAVORITE;
    let mut action_label = labels.add_favorite;
    let mut ctx_voice: Option<FavoriteVoice> = None;

    if target == combo_favorites {
        if let Some(fav) = current_favorite_selection(hwnd) {
            action_id = VOICE_MENU_ID_REMOVE_FAVORITE;
            action_label = labels.remove_favorite;
            ctx_voice = Some(fav);
        }
    } else if target == combo_voice {
        let Some(voice_name) = current_voice_selection(hwnd, engine) else {
            return;
        };
        let is_favorite = with_state(hwnd, |state| {
            state
                .settings
                .favorite_voices
                .iter()
                .any(|fav| fav.engine == engine && fav.short_name == voice_name)
        })
        .unwrap_or(false);
        if is_favorite {
            action_id = VOICE_MENU_ID_REMOVE_FAVORITE;
            action_label = labels.remove_favorite;
        }
        ctx_voice = Some(FavoriteVoice {
            engine,
            short_name: voice_name,
        });
    } else {
        return;
    }

    let Some(ctx) = ctx_voice else {
        return;
    };
    let menu = CreatePopupMenu().unwrap_or(HMENU(0));
    if menu.0 == 0 {
        return;
    }
    let _ = AppendMenuW(
        menu,
        MF_STRING,
        action_id as usize,
        PCWSTR(to_wide(&action_label).as_ptr()),
    );
    let _ = with_state(hwnd, |state| {
        state.voice_context_voice = Some(ctx);
    });

    let mut x = (lparam.0 & 0xffff) as i32;
    let mut y = ((lparam.0 >> 16) & 0xffff) as i32;
    if x == -1 && y == -1 {
        let mut pt = windows::Win32::Foundation::POINT::default();
        let _ = GetCursorPos(&mut pt);
        x = pt.x;
        y = pt.y;
    }

    SetForegroundWindow(hwnd);
    let _ = TrackPopupMenu(menu, TPM_RIGHTBUTTON, x, y, 0, hwnd, None);
    let _ = PostMessageW(hwnd, WM_NULL, WPARAM(0), LPARAM(0));
}

unsafe fn handle_voice_context_favorite(hwnd: HWND, add: bool) {
    let ctx = with_state(hwnd, |state| state.voice_context_voice.clone()).unwrap_or(None);
    let Some(fav) = ctx else {
        return;
    };
    if add {
        add_favorite_voice(hwnd, fav.engine, &fav.short_name);
    } else {
        remove_favorite_voice(hwnd, fav.engine, &fav.short_name);
    }
    let _ = with_state(hwnd, |state| {
        state.voice_context_voice = None;
    });
}

unsafe fn add_favorite_voice(hwnd: HWND, engine: TtsEngine, voice_name: &str) {
    let _ = with_state(hwnd, |state| {
        if state
            .settings
            .favorite_voices
            .iter()
            .any(|fav| fav.engine == engine && fav.short_name == voice_name)
        {
            return;
        }
        state.settings.favorite_voices.push(FavoriteVoice {
            engine,
            short_name: voice_name.to_string(),
        });
    });
    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
        save_settings(settings);
    }
    refresh_voice_panel(hwnd);
}

unsafe fn remove_favorite_voice(hwnd: HWND, engine: TtsEngine, voice_name: &str) {
    let _ = with_state(hwnd, |state| {
        state
            .settings
            .favorite_voices
            .retain(|fav| !(fav.engine == engine && fav.short_name == voice_name));
    });
    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
        save_settings(settings);
    }
    refresh_voice_panel(hwnd);
}

unsafe fn handle_voice_panel_tab(hwnd: HWND) -> bool {
    let (
        visible,
        combo_engine,
        combo_voice,
        checkbox_multilingual,
        combo_favorites,
        favorites_visible,
        is_edge,
        hwnd_tab,
    ) = match with_state(hwnd, |state| {
        (
            state.voice_panel_visible,
            state.voice_combo_engine,
            state.voice_combo_voice,
            state.voice_checkbox_multilingual,
            state.voice_combo_favorites,
            state.voice_favorites_visible,
            matches!(state.settings.tts_engine, TtsEngine::Edge),
            state.hwnd_tab,
        )
    }) {
        Some(values) => values,
        None => return false,
    };
    if !visible && !favorites_visible {
        return false;
    }
    let focus = GetFocus();
    if focus.0 == 0 {
        return false;
    }
    let is_combo_focus = focus == combo_engine
        || focus == combo_voice
        || (is_edge && focus == checkbox_multilingual)
        || (favorites_visible && focus == combo_favorites);
    if is_combo_focus {
        let dropped = SendMessageW(focus, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 != 0;
        if dropped {
            return false;
        }
    }
    let (mut hwnd_edit, is_audiobook) = with_state(hwnd, |state| {
        let doc = state.docs.get(state.current);
        let hwnd_edit = doc.map(|d| d.hwnd_edit).unwrap_or(HWND(0));
        let is_audiobook = doc
            .map(|d| matches!(d.format, FileFormat::Audiobook))
            .unwrap_or(false);
        (hwnd_edit, is_audiobook)
    })
    .unwrap_or((HWND(0), false));
    if is_audiobook {
        hwnd_edit = hwnd_tab;
    }
    if focus != hwnd_edit
        && focus != combo_engine
        && focus != combo_voice
        && focus != hwnd_tab
        && !(is_edge && focus == checkbox_multilingual)
        && !(favorites_visible && focus == combo_favorites)
    {
        return false;
    }
    let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
    if focus == hwnd_edit || focus == hwnd_tab {
        if visible {
            SetFocus(combo_engine);
        } else if favorites_visible {
            SetFocus(combo_favorites);
        }
        return true;
    }
    let fallback_edit = if hwnd_edit.0 != 0 {
        hwnd_edit
    } else {
        hwnd_tab
    };
    if focus == combo_engine {
        let target = if shift_down {
            fallback_edit
        } else {
            combo_voice
        };
        if target.0 != 0 {
            SetFocus(target);
            return true;
        }
    }
    if focus == combo_voice {
        let target = if shift_down {
            combo_engine
        } else if is_edge {
            checkbox_multilingual
        } else if favorites_visible {
            combo_favorites
        } else {
            fallback_edit
        };
        if target.0 != 0 {
            SetFocus(target);
            return true;
        }
    }
    if is_edge && focus == checkbox_multilingual {
        let target = if shift_down {
            combo_voice
        } else if favorites_visible {
            combo_favorites
        } else {
            fallback_edit
        };
        if target.0 != 0 {
            SetFocus(target);
            return true;
        }
    }
    if favorites_visible && focus == combo_favorites {
        let target = if shift_down {
            if visible {
                if is_edge {
                    checkbox_multilingual
                } else {
                    combo_voice
                }
            } else {
                fallback_edit
            }
        } else {
            fallback_edit
        };
        if target.0 != 0 {
            SetFocus(target);
            return true;
        }
    }
    false
}

unsafe fn create_accelerators() -> HACCEL {
    let virt = FCONTROL | FVIRTKEY;
    let virt_shift = FCONTROL | FSHIFT | FVIRTKEY;
    let virt_alt = FALT | FVIRTKEY;
    let virt_alt_shift = FALT | FSHIFT | FVIRTKEY;
    let mut accels = [
        ACCEL {
            fVirt: virt,
            key: 'N' as u16,
            cmd: IDM_FILE_NEW as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'O' as u16,
            cmd: IDM_FILE_OPEN as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'S' as u16,
            cmd: IDM_FILE_SAVE as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'S' as u16,
            cmd: IDM_FILE_SAVE_ALL as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'W' as u16,
            cmd: IDM_FILE_CLOSE as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'F' as u16,
            cmd: IDM_EDIT_FIND as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'F' as u16,
            cmd: IDM_EDIT_FIND_IN_FILES as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'M' as u16,
            cmd: IDM_EDIT_STRIP_MARKDOWN as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'H' as u16,
            cmd: IDM_EDIT_HARD_LINE_BREAK as u16,
        },
        ACCEL {
            fVirt: virt_alt_shift,
            key: 'O' as u16,
            cmd: IDM_EDIT_ORDER_ITEMS as u16,
        },
        ACCEL {
            fVirt: virt_alt_shift,
            key: 'K' as u16,
            cmd: IDM_EDIT_KEEP_UNIQUE_ITEMS as u16,
        },
        ACCEL {
            fVirt: virt_alt_shift,
            key: 'Z' as u16,
            cmd: IDM_EDIT_REVERSE_ITEMS as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: VK_RETURN.0 as u16,
            cmd: IDM_EDIT_NORMALIZE_WHITESPACE as u16,
        },
        ACCEL {
            fVirt: FVIRTKEY,
            key: VK_F3.0 as u16,
            cmd: IDM_EDIT_FIND_NEXT as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'H' as u16,
            cmd: IDM_EDIT_REPLACE as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'A' as u16,
            cmd: IDM_EDIT_SELECT_ALL as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'Q' as u16,
            cmd: IDM_EDIT_QUOTE_LINES as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'Q' as u16,
            cmd: IDM_EDIT_UNQUOTE_LINES as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'J' as u16,
            cmd: IDM_EDIT_JOIN_LINES as u16,
        },
        ACCEL {
            fVirt: virt_alt,
            key: 'Y' as u16,
            cmd: IDM_EDIT_TEXT_STATS as u16,
        },
        ACCEL {
            fVirt: virt,
            key: VK_TAB.0 as u16,
            cmd: IDM_NEXT_TAB as u16,
        },
        ACCEL {
            fVirt: FVIRTKEY,
            key: VK_F4.0 as u16,
            cmd: IDM_FILE_READ_PAUSE as u16,
        },
        ACCEL {
            fVirt: FVIRTKEY,
            key: VK_F5.0 as u16,
            cmd: IDM_FILE_READ_START as u16,
        },
        ACCEL {
            fVirt: FVIRTKEY,
            key: VK_F6.0 as u16,
            cmd: IDM_FILE_READ_STOP as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'R' as u16,
            cmd: IDM_FILE_AUDIOBOOK as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'B' as u16,
            cmd: IDM_FILE_BATCH_AUDIOBOOK as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'R' as u16,
            cmd: IDM_FILE_PODCAST as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'Y' as u16,
            cmd: IDM_TOOLS_IMPORT_YOUTUBE as u16,
        },
        ACCEL {
            fVirt: virt_shift,
            key: 'P' as u16,
            cmd: IDM_TOOLS_PROMPT as u16,
        },
        ACCEL {
            fVirt: virt,
            key: 'B' as u16,
            cmd: IDM_INSERT_BOOKMARK as u16,
        },
    ];
    CreateAcceleratorTableW(&mut accels).unwrap_or(HACCEL(0))
}

pub(crate) unsafe fn get_active_edit(hwnd: HWND) -> Option<HWND> {
    with_state(hwnd, |state| {
        state.docs.get(state.current).map(|doc| doc.hwnd_edit)
    })
    .flatten()
}

unsafe fn insert_bookmark(hwnd: HWND) {
    let (hwnd_edit, path, format): (HWND, std::path::PathBuf, FileFormat) =
        match with_state(hwnd, |state| {
            state
                .docs
                .get(state.current)
                .and_then(|doc| doc.path.clone().map(|p| (doc.hwnd_edit, p, doc.format)))
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
                (
                    current_total as i32,
                    format!("Posizione audio: {:02}:{:02}", mins, secs),
                )
            } else {
                (0, "Audio non in riproduzione".to_string())
            }
        })
        .unwrap_or((0, String::new()));

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
        })
        .unwrap_or(HWND(0));

        if bookmarks_window.0 != 0 {
            unsafe {
                app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window);
            }
        }
        return;
    }

    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut cr as *mut _ as isize),
        );
    }

    let pos = cr.cpMax;

    // 1. Try to get up to 60 characters AFTER the cursor
    let mut buffer = vec![0u16; 62];
    let mut tr = TEXTRANGEW {
        chrg: CHARRANGE {
            cpMin: pos,
            cpMax: pos + 60,
        },
        lpstrText: PWSTR(buffer.as_mut_ptr()),
    };
    let copied = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_GETTEXTRANGE,
            WPARAM(0),
            LPARAM(&mut tr as *mut _ as isize),
        )
        .0 as usize
    };
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
            chrg: CHARRANGE {
                cpMin: start_pre,
                cpMax: pos,
            },
            lpstrText: PWSTR(buffer_pre.as_mut_ptr()),
        };
        let copied_pre = unsafe {
            SendMessageW(
                hwnd_edit,
                EM_GETTEXTRANGE,
                WPARAM(0),
                LPARAM(&mut tr_pre as *mut _ as isize),
            )
            .0 as usize
        };
        let mut snippet_pre = String::from_utf16_lossy(&buffer_pre[..copied_pre]);

        // Take text after the last newline in this prefix
        if let Some(idx) = snippet_pre.rfind(|c| c == '\r' || c == '\n') {
            snippet_pre = snippet_pre[idx + 1..].to_string();
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
    })
    .unwrap_or(HWND(0));

    if bookmarks_window.0 != 0 {
        unsafe {
            app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window);
        }
    }
}

pub(crate) unsafe fn goto_first_bookmark(
    hwnd_edit: HWND,
    path: &Path,
    bookmarks: &BookmarkStore,
    format: FileFormat,
) {
    let path_str = path.to_string_lossy().to_string();
    if let Some(list) = bookmarks.files.get(&path_str) {
        if let Some(bm) = list.first() {
            if matches!(format, FileFormat::Audiobook) {
                // Audiobook position is handled by playback start
            } else {
                let mut cr = CHARRANGE {
                    cpMin: bm.position,
                    cpMax: bm.position,
                };
                SendMessageW(
                    hwnd_edit,
                    EM_EXSETSEL,
                    WPARAM(0),
                    LPARAM(&mut cr as *mut _ as isize),
                );
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
    update_voice_panel_menu_check(hwnd);
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

fn spawn_new_window_with_path(path: &Path) -> bool {
    let Ok(exe) = std::env::current_exe() else {
        return false;
    };
    std::process::Command::new(exe).arg(path).spawn().is_ok()
}

unsafe fn open_document_with_encoding(hwnd: HWND, path: &Path, encoding: Option<TextEncoding>) {
    let behavior =
        with_state(hwnd, |state| state.settings.open_behavior).unwrap_or(OpenBehavior::NewTab);
    if behavior == OpenBehavior::NewWindow {
        if spawn_new_window_with_path(path) {
            return;
        }
    }
    editor_manager::open_document_with_encoding(hwnd, path, encoding);
}

unsafe fn open_path_with_behavior(hwnd: HWND, path: &Path) {
    open_document_with_encoding(hwnd, path, None);
}

unsafe fn open_recent_by_index(hwnd: HWND, index: usize) {
    let path = with_state(hwnd, |state| state.recent_files.get(index).cloned()).unwrap_or(None);
    let Some(path) = path else {
        return;
    };
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if !path.exists() {
        show_error(hwnd, language, &recent_missing_message(language));
        let files = with_state(hwnd, |state| {
            state.recent_files.retain(|p| p != &path);
            update_recent_menu(hwnd, state.hmenu_recent);
            state.recent_files.clone()
        })
        .unwrap_or_default();
        save_recent_files(&files);
        return;
    }
    open_path_with_behavior(hwnd, &path);
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
    let title = path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or("File")
        .to_string();
    let (hwnd_edit, new_index) = with_state(hwnd, |state| {
        let hwnd_edit = create_edit(
            hwnd,
            state.hfont,
            state.settings.word_wrap,
            state.settings.text_color,
            state.settings.text_size,
        );
        editor_manager::set_edit_text(hwnd_edit, &pdf_loading_placeholder(0));
        let doc = Document {
            title: title.clone(),
            path: Some(path_buf.clone()),
            hwnd_edit,
            dirty: false,
            format: FileFormat::Pdf,
            opened_text_encoding: None,
            current_save_text_encoding: None,
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
            show_info(hwnd, language, &pdf_loaded_message(language));
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
    format!("Caricamento PDF...\r\n\r\n[{bar}]\r\nAnalisi in corso {spinner}")
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
        open_path_with_behavior(hwnd, &path);
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
        Some((current, state.hwnd_tab, state.docs.len()))
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

pub(crate) fn sanitize_filename(input: &str) -> String {
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

pub(crate) unsafe fn save_audio_dialog(
    hwnd: HWND,
    suggested_name: Option<&str>,
) -> Option<PathBuf> {
    let mut file_buf = vec![0u16; 4096];
    if let Some(name) = suggested_name {
        let mut name_wide = to_wide(name);
        if let Some(0) = name_wide.last() {
            name_wide.pop();
        }
        let copy_len = name_wide.len().min(file_buf.len() - 1);
        file_buf[..copy_len].copy_from_slice(&name_wide[..copy_len]);
    }
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let filter_raw = i18n::tr(language, "dialog.save_audio_filter");
    let filter = to_wide(&filter_raw.replace("\\0", "\0"));
    let title = to_wide(&i18n::tr(language, "dialog.save_audio_title"));
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
        if path.extension().is_none() {
            path.set_extension("mp3");
        }
        Some(path)
    } else {
        None
    }
}

pub(crate) unsafe fn show_error(hwnd: HWND, language: Language, message: &str) {
    log_debug(&format!("Error shown: {message}"));
    let wide = to_wide(message);
    let title = to_wide(&error_title(language));
    MessageBoxW(
        hwnd,
        PCWSTR(wide.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONERROR,
    );
}

pub(crate) unsafe fn show_info(hwnd: HWND, language: Language, message: &str) {
    log_debug(&format!("Info shown: {message}"));
    let wide = to_wide(message);
    let title = to_wide(&info_title(language));
    MessageBoxW(
        hwnd,
        PCWSTR(wide.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONINFORMATION,
    );
}

pub(crate) unsafe fn send_open_file(hwnd: HWND, path: &str) -> bool {
    let wide = to_wide(path);
    let data = COPYDATASTRUCT {
        dwData: COPYDATA_OPEN_FILE,
        cbData: (wide.len() * std::mem::size_of::<u16>()) as u32,
        lpData: wide.as_ptr() as *mut std::ffi::c_void,
    };
    SendMessageW(
        hwnd,
        WM_COPYDATA,
        WPARAM(0),
        LPARAM(&data as *const _ as isize),
    );
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
        files: files
            .iter()
            .map(|p| p.to_string_lossy().to_string())
            .collect(),
    };
    if let Ok(json) = serde_json::to_string_pretty(&store) {
        let _ = std::fs::write(path, json);
    }
}

#[implement(IFileDialogEvents, IFileDialogControlEvents)]
struct CustomFileDialogEventHandler {
    _encoding_label: String,
    _encodings: Vec<String>,
    _initial_encoding: TextEncoding,
    _is_save_dialog: bool,
}

impl IFileDialogEvents_Impl for CustomFileDialogEventHandler {
    fn OnFileOk(&self, _pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnFolderChange(&self, _pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnFolderChanging(
        &self,
        _pfd: Option<&IFileDialog>,
        _psi: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnSelectionChange(&self, _pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnShareViolation(
        &self,
        _pfd: Option<&IFileDialog>,
        _psi: Option<&IShellItem>,
    ) -> windows::core::Result<windows::Win32::UI::Shell::FDE_SHAREVIOLATION_RESPONSE> {
        Ok(windows::Win32::UI::Shell::FDESVR_DEFAULT)
    }
    fn OnTypeChange(&self, pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        unsafe {
            let pfd = pfd.unwrap();
            let filter_index = pfd.GetFileTypeIndex()?;
            crate::log_debug(&format!("OnTypeChange: filter_index = {}", filter_index));
            let pfdc: IFileDialogCustomize = pfd.cast()?;
            // Show encoding only for TXT:
            // - Open dialog: TXT is index 2
            // - Save dialog: TXT is index 1
            let is_txt = if self._is_save_dialog {
                filter_index == 1
            } else {
                filter_index == 2
            };
            if is_txt {
                crate::log_debug("OnTypeChange: showing encoding combobox");
                // Show the ComboBox (101)
                pfdc.SetControlState(
                    101,
                    windows::Win32::UI::Shell::CDCS_VISIBLE
                        | windows::Win32::UI::Shell::CDCS_ENABLED,
                )?;
            } else {
                crate::log_debug("OnTypeChange: hiding encoding combobox");
                // Hide the ComboBox (101)
                pfdc.SetControlState(101, windows::Win32::UI::Shell::CDCS_INACTIVE)?;
            }
        }
        Ok(())
    }
    fn OnOverwrite(
        &self,
        _pfd: Option<&IFileDialog>,
        _psi: Option<&IShellItem>,
    ) -> windows::core::Result<windows::Win32::UI::Shell::FDE_OVERWRITE_RESPONSE> {
        Ok(windows::Win32::UI::Shell::FDEOR_DEFAULT)
    }
}

impl IFileDialogControlEvents_Impl for CustomFileDialogEventHandler {
    fn OnItemSelected(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
        _dwiditem: u32,
    ) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnButtonClicked(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
    ) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnCheckButtonToggled(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
        _pbchecked: windows::Win32::Foundation::BOOL,
    ) -> windows::core::Result<()> {
        Ok(())
    }
    fn OnControlActivating(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
    ) -> windows::core::Result<()> {
        Ok(())
    }
}

fn encoding_to_index(enc: TextEncoding) -> u32 {
    match enc {
        TextEncoding::Ansi => 0,
        TextEncoding::Utf8 => 1,
        TextEncoding::Utf8Bom => 2,
        TextEncoding::Utf16Le => 3,
        TextEncoding::Utf16Be => 4,
    }
}

fn index_to_encoding(index: u32) -> TextEncoding {
    match index {
        0 => TextEncoding::Ansi,
        1 => TextEncoding::Utf8,
        2 => TextEncoding::Utf8Bom,
        3 => TextEncoding::Utf16Le,
        4 => TextEncoding::Utf16Be,
        _ => TextEncoding::Utf8,
    }
}

pub(crate) unsafe fn open_file_dialog_with_encoding(
    hwnd: HWND,
) -> Option<(PathBuf, Option<TextEncoding>)> {
    log_debug("open_file_dialog_with_encoding called");
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

    use windows::Win32::UI::Shell::FileOpenDialog;
    use windows::Win32::UI::Shell::IFileOpenDialog;

    let pfd: IFileOpenDialog = match CoCreateInstance(&FileOpenDialog, None, CLSCTX_ALL) {
        Ok(dialog) => {
            log_debug("FileOpenDialog created successfully");
            dialog
        }
        Err(e) => {
            log_debug(&format!("Failed to create FileOpenDialog: {:?}", e));
            return None;
        }
    };

    let filter_raw = i18n::tr(language, "dialog.open_filter");
    let parts: Vec<&str> = filter_raw.split("\\0").collect();
    let mut spec = Vec::new();
    let mut pattern_wides = Vec::new();
    let mut name_wides = Vec::new();
    for i in (0..parts.len().saturating_sub(1)).step_by(2) {
        if parts[i].is_empty() {
            break;
        }
        name_wides.push(to_wide(parts[i]));
        pattern_wides.push(to_wide(parts[i + 1]));
    }
    for i in 0..name_wides.len() {
        spec.push(COMDLG_FILTERSPEC {
            pszName: PCWSTR(name_wides[i].as_ptr()),
            pszSpec: PCWSTR(pattern_wides[i].as_ptr()),
        });
    }
    pfd.SetFileTypes(&spec).ok()?;
    pfd.SetFileTypeIndex(1).ok()?; // Default to "All supported formats"

    let pfdc: IFileDialogCustomize = pfd.cast().ok()?;
    let encoding_label = i18n::tr(language, "dialog.encoding_label");
    let encodings = vec![
        i18n::tr(language, "encoding.ansi"),
        i18n::tr(language, "encoding.utf8"),
        i18n::tr(language, "encoding.utf8bom"),
        i18n::tr(language, "encoding.utf16le"),
        i18n::tr(language, "encoding.utf16be"),
    ];

    log_debug("Adding encoding controls to open dialog");

    // Use ComboBox with "Codifica: " prefix in each item for NVDA
    pfdc.AddComboBox(101).ok()?;

    for (i, enc_name) in encodings.iter().enumerate() {
        let item_text = format!("{} {}", encoding_label, enc_name);
        pfdc.AddControlItem(101, i as u32, PCWSTR(to_wide(&item_text).as_ptr()))
            .ok()?;
    }
    pfdc.SetSelectedControlItem(101, encoding_to_index(TextEncoding::Utf8))
        .ok()?;

    let handler: IFileDialogEvents = CustomFileDialogEventHandler {
        _encoding_label: encoding_label,
        _encodings: encodings,
        _initial_encoding: TextEncoding::Utf8,
        _is_save_dialog: false,
    }
    .into();
    let cookie = pfd.Advise(&handler).ok()?;
    log_debug(&format!(
        "Event handler registered with cookie: {:?}",
        cookie
    ));

    // Trigger OnTypeChange to set initial visibility
    // Default index 1 = "All supported formats", encoding will be hidden
    log_debug("Triggering initial OnTypeChange");
    let _ = pfd.SetFileTypeIndex(1);

    log_debug("Showing open dialog");
    if pfd.Show(hwnd).is_ok() {
        log_debug("Dialog closed with OK");
        let item = pfd.GetResult().ok()?;
        let path_ptr = item
            .GetDisplayName(windows::Win32::UI::Shell::SIGDN_FILESYSPATH)
            .ok()?;
        let path_str = from_wide(path_ptr.0);
        CoTaskMemFree(Some(path_ptr.0 as *const _));

        let selected_encoding_idx = pfdc.GetSelectedControlItem(101).ok()?;
        let filter_index = pfd.GetFileTypeIndex().ok()?;

        let path = PathBuf::from(path_str);

        // Only return encoding for text files (filter index 2 = TXT)
        let encoding = if filter_index == 2 {
            Some(index_to_encoding(selected_encoding_idx))
        } else {
            None
        };

        pfd.Unadvise(cookie).ok()?;
        Some((path, encoding))
    } else {
        pfd.Unadvise(cookie).ok()?;
        None
    }
}

pub(crate) unsafe fn save_file_dialog_with_encoding(
    hwnd: HWND,
    suggested_name: Option<&str>,
    initial_encoding: TextEncoding,
) -> Option<(PathBuf, TextEncoding)> {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

    let pfd: IFileSaveDialog = CoCreateInstance(&FileSaveDialog, None, CLSCTX_ALL).ok()?;

    let filter_raw = i18n::tr(language, "dialog.save_filter");
    let parts: Vec<&str> = filter_raw.split("\\0").collect();
    let mut spec = Vec::new();
    let mut pattern_wides = Vec::new();
    let mut name_wides = Vec::new();
    for i in (0..parts.len().saturating_sub(1)).step_by(2) {
        if parts[i].is_empty() {
            break;
        }
        name_wides.push(to_wide(parts[i]));
        pattern_wides.push(to_wide(parts[i + 1]));
    }
    for i in 0..name_wides.len() {
        spec.push(COMDLG_FILTERSPEC {
            pszName: PCWSTR(name_wides[i].as_ptr()),
            pszSpec: PCWSTR(pattern_wides[i].as_ptr()),
        });
    }
    pfd.SetFileTypes(&spec).ok()?;
    pfd.SetFileTypeIndex(1).ok()?; // Default to TXT
    pfd.SetDefaultExtension(w!("txt")).ok()?;

    if let Some(name) = suggested_name {
        pfd.SetFileName(PCWSTR(to_wide(name).as_ptr())).ok()?;
    }

    let pfdc: IFileDialogCustomize = pfd.cast().ok()?;
    let encoding_label = i18n::tr(language, "dialog.encoding_label");
    let encodings = vec![
        i18n::tr(language, "encoding.ansi"),
        i18n::tr(language, "encoding.utf8"),
        i18n::tr(language, "encoding.utf8bom"),
        i18n::tr(language, "encoding.utf16le"),
        i18n::tr(language, "encoding.utf16be"),
    ];

    // Use ComboBox with "Codifica: " prefix in each item for NVDA
    pfdc.AddComboBox(101).ok()?;

    for (i, enc_name) in encodings.iter().enumerate() {
        let item_text = format!("{} {}", encoding_label, enc_name);
        pfdc.AddControlItem(101, i as u32, PCWSTR(to_wide(&item_text).as_ptr()))
            .ok()?;
    }
    pfdc.SetSelectedControlItem(101, encoding_to_index(initial_encoding))
        .ok()?;

    let handler: IFileDialogEvents = CustomFileDialogEventHandler {
        _encoding_label: encoding_label,
        _encodings: encodings,
        _initial_encoding: initial_encoding,
        _is_save_dialog: true,
    }
    .into();
    let cookie = pfd.Advise(&handler).ok()?;

    // Trigger OnTypeChange to set initial visibility (filter index 1 = TXT for save dialog)
    let _ = pfd.SetFileTypeIndex(1);

    if pfd.Show(hwnd).is_ok() {
        let item = pfd.GetResult().ok()?;
        let path_ptr = item
            .GetDisplayName(windows::Win32::UI::Shell::SIGDN_FILESYSPATH)
            .ok()?;
        let path_str = from_wide(path_ptr.0);
        CoTaskMemFree(Some(path_ptr.0 as *const _));

        let selected_encoding_idx = pfdc.GetSelectedControlItem(101).ok()?;
        let filter_index = pfd.GetFileTypeIndex().ok()?;

        let mut path = PathBuf::from(path_str);
        if path.extension().is_none() {
            match filter_index {
                1 => {
                    path.set_extension("txt");
                }
                2 => {
                    path.set_extension("pdf");
                }
                3 => {
                    path.set_extension("docx");
                }
                4 => {
                    path.set_extension("xlsx");
                }
                5 => {
                    path.set_extension("rtf");
                }
                7 => {
                    path.set_extension("html");
                }
                _ => {}
            }
        }

        pfd.Unadvise(cookie).ok()?;
        Some((path, index_to_encoding(selected_encoding_idx)))
    } else {
        pfd.Unadvise(cookie).ok()?;
        None
    }
}
