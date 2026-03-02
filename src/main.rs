#![deny(warnings)]
#![deny(clippy::unwrap_used)]
#![deny(clippy::expect_used)]
#![deny(let_underscore_drop)]
#![allow(unsafe_op_in_unsafe_fn)]
#![windows_subsystem = "windows"]

mod accessibility;
mod com_guard;
mod curl_client;
mod embedded_deps;
mod macros;
use accessibility::*;
mod conpty;
mod diagnostics;
mod sentry_integration;
mod settings;
mod telemetry;
mod watchdog;
use editor_manager::Document;
use settings::*;
mod bookmarks;
use bookmarks::*;
mod tts_engine;
use tts_engine::*;
mod file_handler;
mod mf_encoder;
mod panic_guard;

mod sapi4_engine;
mod sapi5_engine;

use file_handler::*;
mod menu;
use menu::*;
mod search;
use search::*;
mod audio_player;
use audio_player::*;
mod bass_ffmpeg_stream;
mod bass_output;
mod bass_sys;
mod editor_manager;
mod ffmpeg_dyn;
mod ffmpeg_export;
mod ffmpeg_source;
mod subtitle_wasapi;
mod subtitles;
use editor_manager::*;
mod app_windows;
mod audio_monitor;
mod audio_utils;
mod dialogue_voice;
mod i18n;
mod podcast;
mod podcast_recorder;
mod spellcheck;
mod text_ops;
mod tools;
mod updater;
mod wikipedia;
mod wiktionary;
mod win_ocr;

use std::collections::{HashMap, HashSet};
use std::io::Write;
use std::mem::size_of;
use std::path::{Path, PathBuf};
use std::sync::Once;
use std::sync::atomic::AtomicBool;
use std::sync::{Arc, Mutex, mpsc};
use std::time::{Duration, Instant};

use chrono::Local;
use serde::{Deserialize, Serialize};

use windows::Win32::Foundation::{
    BOOL, ERROR_INVALID_PARAMETER, ERROR_INVALID_WINDOW_HANDLE, GetLastError, HINSTANCE, HWND,
    LPARAM, LRESULT, POINT, SetLastError, WIN32_ERROR, WPARAM,
};
use windows::Win32::Graphics::Gdi::{
    COLOR_WINDOW, DEFAULT_GUI_FONT, DeleteObject, GetObjectW, GetStockObject, HBRUSH, HFONT,
    InvalidateRect, LOGFONTW, ScreenToClient,
};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::Win32::System::DataExchange::{COPYDATASTRUCT, IsClipboardFormatAvailable};
use windows::Win32::System::Diagnostics::Debug::MessageBeep;
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, LoadLibraryW};
use windows::Win32::System::Threading::{AttachThreadInput, GetCurrentThreadId};
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::Dialogs::{
    FINDREPLACE_FLAGS, FINDREPLACEW, GetOpenFileNameW, GetSaveFileNameW, OFN_EXPLORER,
    OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_OVERWRITEPROMPT, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::RichEdit::{
    CFE_AUTOBACKCOLOR, CFM_BACKCOLOR, CHARFORMAT2W, CHARRANGE, EM_EXGETSEL, EM_EXSETSEL,
    EM_GETTEXTRANGE, EM_SETCHARFORMAT, EM_SETEVENTMASK, EN_SELCHANGE, ENM_CHANGE, ENM_SELCHANGE,
    SCF_SELECTION, TEXTRANGEW,
};
use windows::Win32::UI::Controls::{
    BST_CHECKED, EM_GETMODIFY, EM_SETMODIFY, ICC_BAR_CLASSES, ICC_TAB_CLASSES,
    INITCOMMONCONTROLSEX, InitCommonControlsEx, NMHDR, SB_SETTEXTW, STATUSCLASSNAMEW,
    TCM_GETCURSEL, TCN_SELCHANGE, WC_BUTTON, WC_COMBOBOXW, WC_STATIC, WC_TABCONTROLW,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, GetKeyState, SetActiveWindow, SetFocus, VK_APPS, VK_CONTROL, VK_ESCAPE,
    VK_F1, VK_F2, VK_F3, VK_F4, VK_F7, VK_F8, VK_F9, VK_F10, VK_MEDIA_PLAY_PAUSE, VK_MENU, VK_NEXT,
    VK_OEM_COMMA, VK_OEM_PERIOD, VK_PRIOR, VK_RETURN, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows::Win32::UI::Shell::{
    DragAcceptFiles, DragFinish, DragQueryFileW, FileSaveDialog, HDROP, IFileDialog,
    IFileDialogControlEvents, IFileDialogControlEvents_Impl, IFileDialogCustomize,
    IFileDialogEvents, IFileDialogEvents_Impl, IFileSaveDialog, IShellItem,
    SHCreateItemFromParsingName, ShellExecuteW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    ACCEL, AllowSetForegroundWindow, AppendMenuW, BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX,
    CB_ADDSTRING, CB_GETCOUNT, CB_GETCURSEL, CB_GETDROPPEDSTATE, CB_GETITEMDATA, CB_RESETCONTENT,
    CB_SETCURSEL, CB_SETITEMDATA, CBN_SELCHANGE, CBS_DROPDOWNLIST, CHILDID_SELF, CREATESTRUCTW,
    CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT, CallWindowProcW, CheckMenuItem, CreateAcceleratorTableW,
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DeleteMenu, DestroyWindow, DispatchMessageW,
    DrawMenuBar, EN_CHANGE, EN_KILLFOCUS, ES_AUTOHSCROLL, EVENT_OBJECT_FOCUS, EnableMenuItem,
    EnumWindows, FALT, FCONTROL, FSHIFT, FVIRTKEY, FindWindowW, GWLP_USERDATA, GWLP_WNDPROC,
    GetClassNameW, GetCursorPos, GetForegroundWindow, GetMenu, GetMenuItemCount, GetMessageW,
    GetParent, GetSubMenu, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW,
    GetWindowThreadProcessId, HACCEL, HCURSOR, HICON, HMENU, IDC_ARROW, IDI_APPLICATION, IDYES,
    IsChild, IsIconic, IsWindow, KillTimer, LoadCursorW, LoadIconW, MB_ICONASTERISK, MB_ICONERROR,
    MB_ICONINFORMATION, MB_OK, MB_YESNO, MESSAGEBOX_RESULT, MESSAGEBOX_STYLE, MF_BYCOMMAND,
    MF_BYPOSITION, MF_CHECKED, MF_ENABLED, MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING,
    MF_UNCHECKED, MSG, MessageBoxW, ModifyMenuW, OBJID_CLIENT, PostMessageW, PostQuitMessage,
    RegisterClassW, RegisterWindowMessageW, SW_HIDE, SW_RESTORE, SW_SHOW, SW_SHOWMAXIMIZED,
    SW_SHOWNORMAL, SendMessageW, SetForegroundWindow, SetTimer, SetWindowLongPtrW, SetWindowTextW,
    ShowWindow, TPM_RETURNCMD, TPM_RIGHTBUTTON, TrackPopupMenu, TranslateAcceleratorW,
    TranslateMessage, WINDOW_STYLE, WM_ACTIVATE, WM_APP, WM_APPCOMMAND, WM_CLOSE, WM_COMMAND,
    WM_CONTEXTMENU, WM_COPY, WM_COPYDATA, WM_CREATE, WM_CUT, WM_DESTROY, WM_DROPFILES,
    WM_INITMENUPOPUP, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL, WM_NOTIFY, WM_NULL, WM_PASTE,
    WM_SETFOCUS, WM_SETFONT, WM_SETREDRAW, WM_SIZE, WM_SYSKEYDOWN, WM_TIMER, WNDCLASSW, WNDPROC,
    WS_CHILD, WS_CLIPCHILDREN, WS_EX_CLIENTEDGE, WS_OVERLAPPEDWINDOW, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{HSTRING, Interface, PCWSTR, PWSTR, implement, w};

const EM_SCROLLCARET: u32 = 0x00B7;
const EM_CHARFROMPOS: u32 = 0x00D7;
const EM_LINEFROMCHAR: u32 = 0x00C9;
const EM_LINEINDEX: u32 = 0x00BB;
const EM_LINELENGTH: u32 = 0x00C1;
const EM_SETSEL: u32 = 0x00B1;
const EM_CANUNDO: u32 = 0x00C6;
const APPCOMMAND_MEDIA_PLAY_PAUSE: usize = 14;

use crate::app_windows::find_in_files_window::{
    FindInFilesCache, apply_find_in_files_selection, focus_find_in_files_results,
};
use crate::podcast::chapters::Chapter;

const WM_PDF_LOADED: u32 = WM_APP + 1;
const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
pub(crate) const WM_DOCUMENT_LOADED: u32 = WM_APP + 9;
const WM_TTS_AUDIOBOOK_DONE: u32 = WM_APP + 4;
const WM_TTS_PLAYBACK_ERROR: u32 = WM_APP + 5;
const WM_UPDATE_PROGRESS: u32 = WM_APP + 6;
const WM_TTS_CHUNK_START: u32 = WM_APP + 7;
const WM_TTS_SAPI_VOICES_LOADED: u32 = WM_APP + 8;
const WM_TTS_START: u32 = WM_APP + 10;

pub const WM_FOCUS_EDITOR: u32 = WM_APP + 30;
pub const WM_UPDATE_DIALOG: u32 = WM_APP + 80;
pub const WM_UPDATE_PROGRESS_OPEN: u32 = WM_APP + 84;
pub const WM_UPDATE_PROGRESS_SET: u32 = WM_APP + 85;
pub const WM_UPDATE_PROGRESS_CLOSE: u32 = WM_APP + 86;
const WM_AUTO_UPDATE_CHECK: u32 = WM_APP + 81;
const WM_CHECK_PENDING_UPDATE: u32 = WM_APP + 82;
const WM_SHOW_CHANGELOG: u32 = WM_APP + 83;
const WM_PODCAST_CHAPTERS_READY: u32 = WM_APP + 31;
const WM_DICTIONARY_LOADED: u32 = WM_APP + 32;
const WM_PODCAST_EPISODE_SAVE_RESULT: u32 = WM_APP + 33;
const FOCUS_EDITOR_TIMER_ID: usize = 1;
const FOCUS_EDITOR_TIMER_ID2: usize = 2;
const FOCUS_EDITOR_TIMER_ID3: usize = 3;
const FOCUS_EDITOR_TIMER_ID4: usize = 4;
const CHAPTER_ANNOUNCE_TIMER_ID: usize = 5;
const SPELLCHECK_HIGHLIGHT_TIMER_ID: usize = 6;
const AUDIO_PLAYLIST_TIMER_ID: usize = 7;
const SPELLCHECK_HIGHLIGHT_DEBOUNCE_MS: u32 = 100;
const PDF_OCR_PROMPT_TIMEOUT_COPYDATA_SECS: u64 = 30;
const COPYDATA_OPEN_FILE: usize = 1;
const VOICE_PANEL_ID_ENGINE: usize = 21001;
const VOICE_PANEL_ID_LANGUAGE: usize = 21012;
const VOICE_PANEL_ID_VOICE: usize = 21002;
const VOICE_PANEL_ID_MULTILINGUAL: usize = 21003;
const VOICE_PANEL_ID_FAVORITES: usize = 21004;
const VOICE_PANEL_ID_SPEED: usize = 21005;
const VOICE_PANEL_ID_PITCH: usize = 21006;
const VOICE_PANEL_ID_VOLUME: usize = 21007;
const VOICE_PANEL_ID_SPEED_EDIT: usize = 21008;
const VOICE_PANEL_ID_PITCH_EDIT: usize = 21009;
const VOICE_PANEL_ID_VOLUME_EDIT: usize = 21010;
const VOICE_PANEL_ID_INSERT_TAG: usize = 21011;
const MAIN_STATUS_ID: usize = 22001;
const VOICE_MENU_ID_ADD_FAVORITE: u32 = 9001;
const VOICE_MENU_ID_REMOVE_FAVORITE: u32 = 9002;
const WINDOW_MENU_INDEX: i32 = 7;
const WINDOW_DOC_MENU_BASE: usize = 11_000;
const WINDOW_DOC_MENU_MAX: usize = 200;
const WINDOW_DOC_MENU_SEPARATOR_ID: usize = 10_999;

fn bring_window_to_foreground(hwnd: HWND) {
    unsafe {
        let foreground = GetForegroundWindow();
        let current_thread = GetCurrentThreadId();
        let mut attached_thread = None;
        if foreground.0 != 0 {
            let foreground_thread = GetWindowThreadProcessId(foreground, None);
            if foreground_thread != 0 && foreground_thread != current_thread {
                if AttachThreadInput(foreground_thread, current_thread, true).as_bool() {
                    attached_thread = Some(foreground_thread);
                } else {
                    log_debug("AttachThreadInput (attach) failed");
                }
            }
        }

        if IsIconic(hwnd).as_bool() {
            ShowWindow(hwnd, SW_RESTORE);
        } else {
            ShowWindow(hwnd, SW_SHOW);
        }
        if !SetForegroundWindow(hwnd).as_bool() {
            log_debug("SetForegroundWindow failed");
        }
        SetActiveWindow(hwnd);
        notify_active_editor_focus(hwnd, false);

        if let Some(foreground_thread) = attached_thread
            && !AttachThreadInput(foreground_thread, current_thread, false).as_bool()
        {
            log_debug("AttachThreadInput (detach) failed");
        }
    }
}

fn notify_active_editor_focus(hwnd: HWND, notify_when_audiobook: bool) {
    let is_audiobook = unsafe {
        with_state(hwnd, |state| {
            state
                .docs
                .get(state.current)
                .map(|doc| matches!(doc.format, FileFormat::Audiobook))
                .unwrap_or(false)
        })
    }
    .unwrap_or(false);

    if is_audiobook {
        if !notify_when_audiobook {
            return;
        }
        let tab_hwnd = unsafe { with_state(hwnd, |state| state.hwnd_tab) }.unwrap_or(HWND(0));
        if tab_hwnd.0 != 0 {
            // SAFETY: `tab_hwnd` is owned by this process and used only for an accessibility focus event.
            unsafe {
                NotifyWinEvent(
                    EVENT_OBJECT_FOCUS,
                    tab_hwnd,
                    OBJID_CLIENT.0,
                    CHILDID_SELF as i32,
                );
            }
        }
        return;
    }

    if let Some(hwnd_edit) = unsafe { get_active_edit(hwnd) } {
        // SAFETY: `hwnd_edit` is the current editor handle managed by app state.
        unsafe {
            NotifyWinEvent(
                EVENT_OBJECT_FOCUS,
                hwnd_edit,
                OBJID_CLIENT.0,
                CHILDID_SELF as i32,
            );
        }
    }
}

pub(crate) fn focus_editor(hwnd: HWND) {
    if has_secondary_window_open(hwnd) {
        return;
    }
    bring_window_to_foreground(hwnd);
    unsafe {
        let result = with_state(hwnd, |state| {
            state
                .docs
                .get(state.current)
                .map(|doc| (doc.hwnd_edit, matches!(doc.format, FileFormat::Audiobook)))
        })
        .flatten();

        if let Some((hwnd_edit, is_audiobook)) = result {
            if is_audiobook {
                let tab_hwnd = with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0));
                if tab_hwnd.0 != 0 {
                    SetFocus(tab_hwnd);
                }
                return;
            }
            SetFocus(hwnd_edit);
            SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
            SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
            crate::log_if_err!(PostMessageW(
                hwnd,
                WM_NEXTDLGCTL,
                WPARAM(hwnd_edit.0 as usize),
                LPARAM(1)
            ));
            NotifyWinEvent(
                EVENT_OBJECT_FOCUS,
                hwnd_edit,
                OBJID_CLIENT.0,
                CHILDID_SELF as i32,
            );
        }
    }
}

pub(crate) fn reset_spellcheck_state(hwnd: HWND) {
    unsafe {
        if with_state(hwnd, |state| {
            state.spellcheck_manager.clear_cache();
            state.spellcheck_last_announce = None;
            state.spellcheck_context = None;
            state.spellcheck_typing_in_progress = false;
        })
        .is_none()
        {
            crate::log_debug("Failed to reset spellcheck state");
        }
    }
}

struct PdfLoadResult {
    hwnd_edit: HWND,
    path: PathBuf,
    result: Result<PdfTextResult, String>,
    from_copydata: bool,
}

pub struct UpdateDialogRequest {
    pub text: String,
    pub title: String,
    pub flags: MESSAGEBOX_STYLE,
    pub response_tx: mpsc::Sender<i32>,
}

pub struct UpdateProgressOpenRequest {
    pub language: Language,
    pub response_tx: mpsc::Sender<isize>,
}

struct PdfLoadingState {
    hwnd_edit: HWND,
    timer_id: usize,
    frame: usize,
    start_time: Instant,
    ocr_timeout_secs: u64,
}

struct PodcastEpisodeSaveResult {
    language: Language,
    target_path: PathBuf,
    error: Option<String>,
}

struct PodcastChaptersReady {
    key: String,
    chapters: Option<Vec<Chapter>>,
}

#[derive(Clone, PartialEq, Eq)]
struct SpellcheckAnnounceKey {
    doc_id: isize,
    line_index: i32,
    start_utf8: usize,
    end_utf8: usize,
    line_hash: u64,
    language: String,
}

#[derive(Clone)]
struct SpellcheckContextMenuState {
    hwnd_edit: HWND,
    line_start: i32,
    language: String,
    word_range: (usize, usize),
    word: String,
    line_text: String,
    suggestions: Vec<String>,
}

struct SpellcheckWordContext {
    doc_id: isize,
    line_index: i32,
    line_start: i32,
    line_text: String,
    line_hash: u64,
    word_range: (usize, usize),
    word: String,
}

fn log_path() -> Option<PathBuf> {
    let mut path = settings::settings_dir();
    path.push("Sonarpad.log");
    Some(path)
}

const MAX_LOG_SIZE: u64 = 150 * 1024;

fn log_lock_path(log_path: &Path) -> Option<PathBuf> {
    let parent = log_path.parent()?;
    Some(parent.join("Sonarpad.log.lock"))
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
                    if writeln!(file, "{}", std::process::id()).is_err() {
                        return;
                    }
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
                    if writeln!(file, "[INFO] log truncated (exceeded 150 KB)").is_err() {
                        return;
                    } else {
                        truncated = true;
                    }
                }
                if !truncated {
                    if std::fs::remove_file(path).is_err() {
                        return;
                    }
                    if let Ok(mut file) = std::fs::OpenOptions::new()
                        .create(true)
                        .write(true)
                        .truncate(true)
                        .open(path)
                        && writeln!(file, "[INFO] log truncated (exceeded 150 KB)").is_err()
                    {
                        return;
                    }
                }
            }
            if std::fs::remove_file(&lock_path).is_err() {}
        }
    });
}

pub(crate) fn log_debug(message: &str) {
    // Push to telemetry ring buffer for hang diagnostics
    telemetry::push_log_line(message);

    let Some(path) = log_path() else {
        return;
    };
    if let Some(parent) = path.parent()
        && std::fs::create_dir_all(parent).is_err()
    {
        return;
    }
    truncate_log_if_needed(&path);
    if let Ok(mut log) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        let timestamp = Local::now().format("%Y-%m-%d %H:%M:%S");
        if writeln!(log, "[{timestamp}] {message}").is_err() {}
    }
}

fn kill_timer_best_effort(hwnd: HWND, timer_id: usize, context: &str) {
    unsafe {
        SetLastError(WIN32_ERROR(0));
        if let Err(err) = KillTimer(hwnd, timer_id) {
            let last_error = GetLastError();
            if last_error == WIN32_ERROR(0)
                || last_error == ERROR_INVALID_PARAMETER
                || last_error == ERROR_INVALID_WINDOW_HANDLE
            {
                return;
            }
            log_debug(&format!(
                "{} failed: {:?} (win32={})",
                context, err, last_error.0
            ));
        }
    }
}

fn clean_menu_label(label: &str) -> String {
    let main = label.split('\t').next().unwrap_or(label);
    let mut cleaned = String::with_capacity(main.len());
    for ch in main.chars() {
        if ch != '&' {
            cleaned.push(ch);
        }
    }
    // Remove accelerator patterns like "(&X)" or "(X)" at the end
    let trimmed = cleaned.trim();
    if let Some(pos) = trimmed.rfind(" (") {
        let suffix = &trimmed[pos..];
        // Check if it matches pattern " (&X)" or " (X)" where X is a single char
        if suffix.len() <= 5 && suffix.ends_with(')') {
            return trimmed[..pos].trim().to_string();
        }
    }
    trimmed.to_string()
}

fn confirm_menu_action(hwnd: HWND, key: &str) {
    let language = unsafe { with_state(hwnd, |state| state.settings.language).unwrap_or_default() };
    let label = i18n::tr(language, key);
    let cleaned = clean_menu_label(&label);
    if !cleaned.is_empty() {
        if key.starts_with("edit.") {
            unsafe {
                with_state(hwnd, |state| {
                    state.undo_action_label = Some(cleaned.clone())
                });
            }
        }
        let message = i18n::tr_f(language, "app.action_completed", &[("action", &cleaned)]);
        unsafe {
            show_info(hwnd, language, &message);
        }
    }
}

fn announce_menu_action_screen_reader(hwnd: HWND, key: &str) {
    let language = unsafe { with_state(hwnd, |state| state.settings.language).unwrap_or_default() };
    let label = i18n::tr(language, key);
    let cleaned = clean_menu_label(&label);
    if !cleaned.is_empty() {
        let message = i18n::tr_f(language, "app.action_completed", &[("action", &cleaned)]);
        crate::accessibility::screen_reader_speak(&message);
    }
}

fn dictionary_cache_key(language: Language, pref: &str, word: &str) -> String {
    let lang = match language {
        Language::Italian => "it",
        Language::English => "en",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::Swedish => "sv",
        Language::Vietnamese => "vi",
        Language::Czech => "cs",
        Language::Polish => "pl",
        Language::French => "fr",
        Language::Serbian => "sr",
        Language::Ukrainian => "uk",
        Language::Lithuanian => "lt",
        Language::Chinese => "zh",
    };
    format!(
        "{}|{}|{}",
        lang,
        pref.trim().to_ascii_lowercase(),
        word.trim().to_ascii_lowercase()
    )
}

fn dictionary_cache_path() -> std::path::PathBuf {
    settings::settings_dir().join("dictionary_cache.json")
}

fn load_dictionary_cache() -> HashMap<String, Vec<String>> {
    let path = dictionary_cache_path();
    if !path.exists() {
        return HashMap::new();
    }
    match std::fs::read_to_string(&path) {
        Ok(content) => serde_json::from_str(&content).unwrap_or_default(),
        Err(_) => HashMap::new(),
    }
}

fn save_dictionary_cache(cache: &HashMap<String, Vec<String>>) {
    let path = dictionary_cache_path();
    if let Ok(content) = serde_json::to_string(cache) {
        crate::log_if_err!(std::fs::write(path, content));
    }
}

pub(crate) fn is_dictionary_not_found_cache_entry(language: Language, lines: &[String]) -> bool {
    lines.len() == 1 && lines[0] == i18n::tr(language, "dictionary.not_found")
}

pub(crate) fn update_dictionary_cache(hwnd: HWND, key: String, lines: Vec<String>) {
    unsafe {
        if with_state(hwnd, |state| {
            state.dictionary_cache.insert(key, lines);
            save_dictionary_cache(&state.dictionary_cache);
        })
        .is_none()
        {
            crate::log_debug("Failed to update dictionary cache state");
        }
    }
}

pub(crate) fn remove_dictionary_cache(hwnd: HWND, key: &str) {
    unsafe {
        if with_state(hwnd, |state| {
            let removed = state.dictionary_cache.remove(key).is_some();
            if removed {
                save_dictionary_cache(&state.dictionary_cache);
            }
        })
        .is_none()
        {
            crate::log_debug("Failed to remove dictionary cache state");
        }
    }
}

struct DictionaryLookupResult {
    key: String,
    lines: Vec<String>,
    generation: usize,
    cacheable: bool,
}

fn start_dictionary_lookup(
    hwnd_val: isize,
    word: String,
    language: Language,
    pref: String,
    key: String,
    generation: usize,
) {
    std::thread::spawn(move || {
        let (lines, cacheable) = match wiktionary::lookup_for_language(&word, language, &pref) {
            Ok(entry) => (wiktionary::format_menu_lines(language, &entry), true),
            Err(wiktionary::LookupError::NotFound { .. }) => {
                (vec![i18n::tr(language, "dictionary.not_found")], false)
            }
            Err(err) => {
                log_debug(&format!("Dictionary lookup failed: {err}"));
                (vec![i18n::tr(language, "dictionary.not_found")], false)
            }
        };
        let result = Box::new(DictionaryLookupResult {
            key,
            lines,
            generation,
            cacheable,
        });
        let hwnd = HWND(hwnd_val);
        if unsafe { IsWindow(hwnd).as_bool() } {
            unsafe {
                crate::log_if_err!(PostMessageW(
                    hwnd,
                    WM_DICTIONARY_LOADED,
                    WPARAM(0),
                    LPARAM(Box::into_raw(result) as isize),
                ));
            }
        }
    });
}

fn prefetch_dictionary_for_selection(hwnd: HWND, hwnd_edit: HWND) {
    let mut range = CHARRANGE::default();
    // SAFETY: `hwnd_edit` is an edit control and `range` is valid writable memory.
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut range as *mut _ as isize),
        );
    }
    let start = range.cpMin;
    let end = range.cpMax;
    if start >= end || (end - start) > 50 {
        return;
    }
    let len = (end - start) as usize;
    let mut buf = vec![0u16; len + 1];
    let mut tr = TEXTRANGEW {
        chrg: CHARRANGE {
            cpMin: start,
            cpMax: end,
        },
        lpstrText: windows::core::PWSTR(buf.as_mut_ptr()),
    };
    // SAFETY: `tr` points to `buf`, which is allocated for the requested range plus terminator.
    let copied = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_GETTEXTRANGE,
            WPARAM(0),
            LPARAM(&mut tr as *mut _ as isize),
        )
    }
    .0 as usize;
    if copied == 0 {
        return;
    }
    let selected = String::from_utf16_lossy(&buf[..copied]);
    let trimmed = selected.trim();
    if trimmed.is_empty() || trimmed.contains(char::is_whitespace) {
        return;
    }
    let word = trimmed.to_string();

    let prefetch_info = unsafe {
        with_state(hwnd, |state| {
            let language = state.settings.language;
            let pref = state.settings.dictionary_translation_language.clone();
            let key = dictionary_cache_key(language, &pref, &word);
            if let Some(lines) = state.dictionary_cache.get(&key).cloned() {
                if is_dictionary_not_found_cache_entry(language, &lines) {
                    state.dictionary_cache.remove(&key);
                    save_dictionary_cache(&state.dictionary_cache);
                } else {
                    return None;
                }
            }
            if state.dictionary_pending_lookup.as_ref() == Some(&key) {
                return None;
            }
            state.dictionary_pending_lookup = Some(key.clone());
            let generation = state.dictionary_prefetch_generation;
            Some((word.clone(), language, pref, key, generation))
        })
    }
    .flatten();

    if let Some((word, language, pref, key, generation)) = prefetch_info {
        start_dictionary_lookup(hwnd.0, word, language, pref, key, generation);
    }
}

fn format_time_hms(seconds: u64) -> String {
    let hours = seconds / 3600;
    let minutes = (seconds % 3600) / 60;
    let secs = seconds % 60;
    if hours > 0 {
        format!("{:02}:{:02}:{:02}", hours, minutes, secs)
    } else {
        format!("{:02}:{:02}", minutes, secs)
    }
}

fn audiobook_position_ms_from_state(state: &AppState) -> Option<u64> {
    let player = state.active_audiobook.as_ref()?;
    if let Some(pos_secs) = player.position_secs() {
        return Some((pos_secs * 1000.0).max(0.0) as u64);
    }
    let accumulated_ms = player.accumulated_seconds.saturating_mul(1000);
    if player.is_paused {
        return Some(accumulated_ms);
    }
    let elapsed_real_ms = player.start_instant.elapsed().as_millis() as f64;
    let elapsed_audio_ms = (elapsed_real_ms * player.speed as f64) as u64;
    Some(accumulated_ms.saturating_add(elapsed_audio_ms))
}

fn update_chapter_announcement(hwnd: HWND) {
    let (current_pos_ms, chapters, last_idx, language) = unsafe {
        with_state(hwnd, |state| {
            (
                audiobook_position_ms_from_state(state),
                state.active_podcast_chapters.clone(),
                state.last_announced_chapter_index,
                state.settings.language,
            )
        })
    }
    .unwrap_or((None, Vec::new(), None, Language::default()));
    let Some(current_pos_ms) = current_pos_ms else {
        return;
    };
    if chapters.is_empty() {
        if unsafe { with_state(hwnd, |state| state.last_announced_chapter_index = None) }.is_none()
        {
            crate::log_debug("Failed to clear last announced chapter index");
        }
        return;
    }
    let current_idx = crate::podcast::chapters::current_chapter_index(current_pos_ms, &chapters);
    if current_idx == last_idx {
        return;
    }
    if unsafe {
        with_state(hwnd, |state| {
            state.last_announced_chapter_index = current_idx
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to update last announced chapter index");
    }
    let Some(idx) = current_idx else {
        return;
    };
    if let Some(chapter) = chapters.get(idx) {
        let message = i18n::tr_f(
            language,
            "playback.chapter_announce",
            &[("title", &chapter.title)],
        );
        nvda_speak(&message);
    }
}

fn announce_current_chapter_on_start(
    hwnd: HWND,
    chapters: &[Chapter],
    current_pos_ms: Option<u64>,
    language: Language,
) {
    if chapters.is_empty() {
        return;
    }
    let current_idx = current_pos_ms
        .and_then(|pos| crate::podcast::chapters::current_chapter_index(pos, chapters))
        .or(Some(0));
    unsafe {
        if with_state(hwnd, |state| {
            state.last_announced_chapter_index = current_idx
        })
        .is_none()
        {
            crate::log_debug("Failed to update last announced chapter index");
        }
    };
    let Some(idx) = current_idx else {
        return;
    };
    if let Some(chapter) = chapters.get(idx) {
        let message = i18n::tr_f(
            language,
            "playback.chapter_announce",
            &[("title", &chapter.title)],
        );
        nvda_speak(&message);
    }
}

pub(crate) fn clear_active_podcast_chapters(hwnd: HWND) {
    unsafe {
        if with_state(hwnd, |state| {
            state.active_podcast_chapters_key = None;
            state.active_podcast_chapters.clear();
            state.last_announced_chapter_index = None;
            state.active_podcast_episode_url = None;
            state.active_podcast_episode_title = None;
            state.active_podcast_episode_cache = None;
        })
        .is_none()
        {
            crate::log_debug("Failed to clear active podcast chapters");
        }
        kill_timer_best_effort(
            hwnd,
            CHAPTER_ANNOUNCE_TIMER_ID,
            "KillTimer CHAPTER_ANNOUNCE",
        );
    }
}

pub(crate) fn reset_active_podcast_chapters_for_playback(hwnd: HWND) {
    let (has_pending, has_active) = unsafe {
        with_state(hwnd, |state| {
            (
                state.pending_podcast_chapters_key.is_some(),
                state.active_podcast_chapters_key.is_some(),
            )
        })
        .unwrap_or((false, false))
    };
    if has_pending {
        unsafe {
            with_state(hwnd, |state| {
                state.active_podcast_chapters_key = None;
                state.active_podcast_chapters.clear();
                state.last_announced_chapter_index = None;
                state.active_podcast_episode_url = None;
                state.active_podcast_episode_title = None;
                state.active_podcast_episode_cache = None;
            });
            kill_timer_best_effort(
                hwnd,
                CHAPTER_ANNOUNCE_TIMER_ID,
                "KillTimer CHAPTER_ANNOUNCE",
            );
        }
        return;
    }
    if !has_active {
        clear_active_podcast_chapters(hwnd);
    }
}

pub(crate) fn set_pending_podcast_chapters_key(hwnd: HWND, key: Option<String>) {
    unsafe { with_state(hwnd, |state| state.pending_podcast_chapters_key = key) };
}

pub(crate) fn activate_pending_podcast_chapters(hwnd: HWND) {
    let (chapters, language, should_announce_unavailable, current_pos_ms) = unsafe {
        with_state(hwnd, |state| {
            let key = state.pending_podcast_chapters_key.take();
            state.active_podcast_chapters_key = key.clone();
            state.last_announced_chapter_index = None;
            if let Some(key) = key.as_ref()
                && let Some(cached) = state.podcast_chapters_cache.get(key)
            {
                match cached {
                    Some(list) => {
                        state.active_podcast_chapters = list.clone();
                        return (
                            list.clone(),
                            state.settings.language,
                            false,
                            audiobook_position_ms_from_state(state),
                        );
                    }
                    None => {
                        state.active_podcast_chapters.clear();
                        return (
                            Vec::new(),
                            state.settings.language,
                            true,
                            audiobook_position_ms_from_state(state),
                        );
                    }
                }
            }
            state.active_podcast_chapters.clear();
            (
                Vec::new(),
                state.settings.language,
                false,
                audiobook_position_ms_from_state(state),
            )
        })
        .unwrap_or((Vec::new(), Language::default(), false, None))
    };
    unsafe {
        if !chapters.is_empty() {
            if SetTimer(hwnd, CHAPTER_ANNOUNCE_TIMER_ID, 500, None) == 0 {
                crate::log_debug("Failed to set CHAPTER_ANNOUNCE_TIMER");
            }
            announce_current_chapter_on_start(hwnd, &chapters, current_pos_ms, language);
        } else {
            kill_timer_best_effort(
                hwnd,
                CHAPTER_ANNOUNCE_TIMER_ID,
                "KillTimer CHAPTER_ANNOUNCE",
            );
        }
        if should_announce_unavailable {
            let message = i18n::tr(language, "playback.chapters_unavailable");
            nvda_speak(&message);
        }
        crate::menu::update_playback_menu(hwnd, true);
    }
}

pub(crate) fn set_active_podcast_episode_info(
    hwnd: HWND,
    url: Option<String>,
    title: Option<String>,
    cache_path: Option<PathBuf>,
) {
    if let Some(url_value) = url {
        unsafe {
            let has_pending_chapters =
                with_state(hwnd, |state| state.pending_podcast_chapters_key.is_some())
                    .unwrap_or(false);
            if with_state(hwnd, |state| {
                state.active_podcast_episode_url = Some(url_value.clone());
                state.active_podcast_episode_title = title;
                state.active_podcast_episode_cache = cache_path;
            })
            .is_none()
            {
                crate::log_debug("Failed to set active podcast episode info");
            }

            if !has_pending_chapters {
                let chapters_url = extract_embedded_chapters_url(&url_value)
                    .or_else(|| extract_buzzsprout_chapters_url(&url_value));
                if let Some(chapters_url) = chapters_url {
                    let key = format!("url_chapters:{url_value}");
                    set_pending_podcast_chapters_key(hwnd, Some(key.clone()));
                    prefetch_podcast_chapters(hwnd, key, chapters_url);
                    activate_pending_podcast_chapters(hwnd);
                }
            }
        }
    }
}

pub(crate) fn download_active_podcast_episode(hwnd: HWND) {
    let (url, title, cache_path, language) = unsafe {
        with_state(hwnd, |state| {
            (
                state.active_podcast_episode_url.clone(),
                state.active_podcast_episode_title.clone(),
                state.active_podcast_episode_cache.clone(),
                state.settings.language,
            )
        })
        .unwrap_or((None, None, None, Language::default()))
    };
    download_podcast_episode(hwnd, url, title, cache_path, language);
}

fn post_podcast_episode_save_result(hwnd: HWND, payload: PodcastEpisodeSaveResult) {
    let ptr = Box::into_raw(Box::new(payload));
    unsafe {
        if let Err(err) = PostMessageW(
            hwnd,
            WM_PODCAST_EPISODE_SAVE_RESULT,
            WPARAM(0),
            LPARAM(ptr as isize),
        ) {
            log_debug(&format!(
                "Failed to post WM_PODCAST_EPISODE_SAVE_RESULT: {}",
                err
            ));
            let _drop_payload = Box::from_raw(ptr);
        }
    }
}

pub(crate) fn download_podcast_episode(
    hwnd: HWND,
    url: Option<String>,
    title: Option<String>,
    cache_path: Option<PathBuf>,
    language: Language,
) {
    let suggested_name = title
        .as_deref()
        .and_then(suggested_filename_from_text)
        .unwrap_or_else(|| "podcast_episode".to_string());
    let mut ext = cache_path
        .as_ref()
        .and_then(|p| p.extension().and_then(|e| e.to_str()))
        .map(|e| e.to_string());
    if ext.is_none()
        && let Some(url) = url.as_deref()
    {
        ext = Path::new(url)
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_string());
    }
    let ext = ext.unwrap_or_else(|| "mp3".to_string());
    let suggested_full = format!("{}.{}", suggested_name, ext);
    let target = save_podcast_episode_dialog(hwnd, language, &suggested_full);
    let Some(target) = target else {
        return;
    };
    let target = ensure_path_extension(target, &ext);
    let cache_path = cache_path.clone();
    let url = url.clone();
    let hwnd_copy = hwnd;
    std::thread::spawn(move || {
        let Some(cache_path) = cache_path.as_ref() else {
            let err = "no cache path".to_string();
            log_debug(&format!("podcast_episode_save_failed {}", err));
            post_podcast_episode_save_result(
                hwnd_copy,
                PodcastEpisodeSaveResult {
                    language,
                    target_path: target.clone(),
                    error: Some(err),
                },
            );
            return;
        };

        if !cache_path.exists() {
            let Some(url) = url.as_ref() else {
                let err = "no cache and no source URL".to_string();
                log_debug(&format!("podcast_episode_save_failed {}", err));
                post_podcast_episode_save_result(
                    hwnd_copy,
                    PodcastEpisodeSaveResult {
                        language,
                        target_path: target.clone(),
                        error: Some(err),
                    },
                );
                return;
            };
            log_debug(&format!(
                "podcast_episode_save: cache missing, downloading from {}",
                url
            ));
            let rt = tokio::runtime::Builder::new_current_thread()
                .enable_all()
                .build()
                .ok();
            let bytes = if let Some(rt) = rt {
                rt.block_on(crate::tools::rss::fetch_url_bytes(
                    url,
                    crate::tools::rss::RssFetchConfig::default(),
                ))
            } else {
                let err = "failed to create async runtime".to_string();
                log_debug(&format!("podcast_episode_save_failed {}", err));
                post_podcast_episode_save_result(
                    hwnd_copy,
                    PodcastEpisodeSaveResult {
                        language,
                        target_path: target.clone(),
                        error: Some(err),
                    },
                );
                return;
            };
            match bytes {
                Ok(b) => {
                    if let Some(parent) = cache_path.parent() {
                        crate::log_if_err!(std::fs::create_dir_all(parent));
                    }
                    if let Err(e) = std::fs::write(cache_path, b) {
                        let err = format!("failed to write to cache: {}", e);
                        log_debug(&format!("podcast_episode_save: {}", err));
                        post_podcast_episode_save_result(
                            hwnd_copy,
                            PodcastEpisodeSaveResult {
                                language,
                                target_path: target.clone(),
                                error: Some(err),
                            },
                        );
                        return;
                    }
                }
                Err(e) => {
                    let err = format!("download failed: {}", e);
                    log_debug(&format!("podcast_episode_save: {}", err));
                    post_podcast_episode_save_result(
                        hwnd_copy,
                        PodcastEpisodeSaveResult {
                            language,
                            target_path: target.clone(),
                            error: Some(err),
                        },
                    );
                    return;
                }
            }
        }

        if std::fs::copy(cache_path, &target).is_ok() {
            log_debug(&format!(
                "podcast_episode_saved src=cache dst={}",
                target.to_string_lossy()
            ));
            post_podcast_episode_save_result(
                hwnd_copy,
                PodcastEpisodeSaveResult {
                    language,
                    target_path: target,
                    error: None,
                },
            );
        } else {
            let err = format!("copy failed to dst={}", target.to_string_lossy());
            log_debug(&format!(
                "podcast_episode_save_failed copy dst={}",
                target.to_string_lossy()
            ));
            post_podcast_episode_save_result(
                hwnd_copy,
                PodcastEpisodeSaveResult {
                    language,
                    target_path: target,
                    error: Some(err),
                },
            );
        }
    });
}

fn ensure_path_extension(mut path: PathBuf, desired_ext: &str) -> PathBuf {
    let has_nonempty_extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| !e.trim().is_empty())
        .unwrap_or(false);
    if !has_nonempty_extension {
        path.set_extension(desired_ext);
    }
    path
}

fn save_podcast_episode_dialog(
    hwnd: HWND,
    language: Language,
    suggested_name: &str,
) -> Option<PathBuf> {
    let raw_filter = i18n::tr(language, "podcasts.download_filter");
    let filter = to_wide(&raw_filter.replace("\\0", "\0"));
    let mut buffer = vec![0u16; 4096];
    let wide_name = to_wide(suggested_name);
    for (i, ch) in wide_name
        .iter()
        .enumerate()
        .take(buffer.len().saturating_sub(1))
    {
        buffer[i] = *ch;
    }
    let initial_dir = PathBuf::from(settings::default_media_save_folder());
    crate::log_if_err!(std::fs::create_dir_all(&initial_dir));
    let initial_dir_wide = to_wide(&initial_dir.to_string_lossy());
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrInitialDir: PCWSTR(initial_dir_wide.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        Flags: OFN_EXPLORER | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT,
        ..Default::default()
    };
    if !unsafe { GetSaveFileNameW(&mut ofn) }.as_bool() {
        return None;
    }
    let end = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
    if end == 0 {
        return None;
    }
    Some(PathBuf::from(String::from_utf16_lossy(&buffer[..end])))
}
pub(crate) fn prefetch_podcast_chapters(hwnd: HWND, key: String, url: String) {
    let should_fetch = unsafe {
        with_state(hwnd, |state| {
            !state.podcast_chapters_cache.contains_key(&key)
        })
        .unwrap_or(false)
    };
    if !should_fetch {
        return;
    }
    let config = unsafe {
        with_state(hwnd, |state| {
            crate::tools::rss::config_from_settings(&state.settings)
        })
        .unwrap_or_else(crate::tools::rss::RssHttpConfig::default)
    };
    if let Err(err) = crate::tools::rss::init_http(config) {
        log_debug(&format!("rss_http_init_error: {}", err));
    }
    let fetch_config = unsafe {
        with_state(hwnd, |state| {
            crate::tools::rss::fetch_config_from_settings(&state.settings)
        })
        .unwrap_or_else(crate::tools::rss::RssFetchConfig::default)
    };
    let fallback_url = extract_embedded_chapters_url(&url);
    let hwnd_copy = hwnd;
    std::thread::spawn(move || {
        let rt = match tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
        {
            Ok(rt) => rt,
            Err(e) => {
                log_debug(&format!("Failed to build tokio runtime: {}", e));
                return;
            }
        };
        let chapters = fetch_chapters_with_fallback(&rt, &url, &fallback_url, fetch_config);
        let msg = Box::new(PodcastChaptersReady { key, chapters });
        unsafe {
            crate::log_if_err!(PostMessageW(
                hwnd_copy,
                WM_PODCAST_CHAPTERS_READY,
                WPARAM(0),
                LPARAM(Box::into_raw(msg) as isize),
            ));
        }
    });
}

pub(crate) fn cache_podcast_chapters(hwnd: HWND, key: String, chapters: Vec<Chapter>) {
    unsafe {
        with_state(hwnd, |state| {
            state.podcast_chapters_cache.insert(key, Some(chapters));
        })
    };
}

fn fetch_chapters_with_fallback(
    rt: &tokio::runtime::Runtime,
    url: &str,
    fallback_url: &Option<String>,
    fetch_config: crate::tools::rss::RssFetchConfig,
) -> Option<Vec<Chapter>> {
    match rt.block_on(crate::tools::rss::fetch_url_bytes(url, fetch_config)) {
        Ok(bytes) => {
            let parsed = crate::podcast::chapters::parse_chapters_json(&bytes);
            if !parsed.is_empty() {
                log_debug(&format!(
                    "podcast_chapters_ok url={} count={}",
                    url,
                    parsed.len()
                ));
                return Some(parsed);
            }
            log_debug(&format!("podcast_chapters_empty url={}", url));
        }
        Err(err) => {
            log_debug(&format!("podcast_chapters_fetch_error {}", err));
        }
    }
    let fallback_url = fallback_url.as_ref()?;
    match rt.block_on(crate::tools::rss::fetch_url_bytes(
        fallback_url,
        fetch_config,
    )) {
        Ok(bytes) => {
            let parsed = crate::podcast::chapters::parse_chapters_json(&bytes);
            if parsed.is_empty() {
                log_debug(&format!("podcast_chapters_empty url={}", fallback_url));
                None
            } else {
                log_debug(&format!(
                    "podcast_chapters_ok url={} count={}",
                    fallback_url,
                    parsed.len()
                ));
                Some(parsed)
            }
        }
        Err(err) => {
            log_debug(&format!("podcast_chapters_fetch_error {}", err));
            None
        }
    }
}

pub(crate) fn extract_embedded_chapters_url(url: &str) -> Option<String> {
    let marker = "/chapters/";
    let idx = url.rfind(marker)?;
    let tail = &url[idx + marker.len()..];
    if tail.starts_with("http://") || tail.starts_with("https://") {
        Some(tail.to_string())
    } else {
        None
    }
}

pub(crate) fn extract_buzzsprout_chapters_url(url: &str) -> Option<String> {
    // Supports direct Buzzsprout URLs and wrapped URLs (e.g. op3.dev/.../www.buzzsprout.com/...).
    // Expected pattern: ".../<show_id>/episodes/<episode_id>-...".
    let marker = "/episodes/";
    let episode_idx = url.find(marker)?;
    let before = &url[..episode_idx];
    let show_id = before
        .rsplit('/')
        .find(|seg| !seg.is_empty() && seg.chars().all(|c| c.is_ascii_digit()))?;
    let after = &url[episode_idx + marker.len()..];
    let episode_id: String = after.chars().take_while(|c| c.is_ascii_digit()).collect();
    if episode_id.is_empty() {
        return None;
    }
    Some(format!(
        "https://www.buzzsprout.com/{show_id}/{episode_id}/chapters.json"
    ))
}

fn announce_player_time(hwnd: HWND) {
    const END_STOP_TOLERANCE_SECS: u64 = 1;

    let (current_raw, path, fallback_total, fallback_is_stopped_same_path, language) = unsafe {
        with_state(hwnd, |state| {
            let current = state
                .active_audiobook
                .as_ref()
                .map(|player| audiobook_position_secs(player).max(0.0).floor() as u64);
            let active_path = state
                .active_audiobook
                .as_ref()
                .map(|player| player.path.clone());
            let doc_path = if active_path.is_none() {
                state.docs.get(state.current).and_then(|doc| {
                    if matches!(doc.format, FileFormat::Audiobook) {
                        doc.path.clone()
                    } else {
                        None
                    }
                })
            } else {
                None
            };
            let path = active_path.or(doc_path);
            let fallback_total = state.active_audiobook.as_ref().and_then(|player| {
                player
                    .duration_secs()
                    .map(|secs| secs.max(0.0).floor() as u64)
            });
            let fallback_is_stopped_same_path = path
                .as_ref()
                .zip(state.last_stopped_audiobook.as_ref())
                .map(|(p, stopped)| p == stopped)
                .unwrap_or(false);
            (
                current,
                path,
                fallback_total,
                fallback_is_stopped_same_path,
                state.settings.language,
            )
        })
    }
    .unwrap_or((None, None, None, false, Language::default()));
    let Some(path) = path else {
        return;
    };

    let total = audiobook_duration_secs(&path).or(fallback_total);
    let Some(current_raw) = current_raw.or_else(|| {
        if fallback_is_stopped_same_path {
            // After reaching EOF and stopping, treat time as restarted from 0:00.
            Some(0)
        } else {
            total.map(|_| 0)
        }
    }) else {
        return;
    };
    let (current, should_stop) = if let Some(total) = total {
        let overrun = current_raw > total;
        if overrun {
            log_debug(&format!(
                "Audio player: position beyond duration (current={} total={}), clamping",
                current_raw, total
            ));
        }
        (
            current_raw.min(total),
            overrun && current_raw >= total.saturating_add(END_STOP_TOLERANCE_SECS),
        )
    } else {
        (current_raw, false)
    };

    let current_str = format_time_hms(current);
    let message = if let Some(total) = total {
        let total_str = format_time_hms(total);
        i18n::tr_f(
            language,
            "player.time_announce",
            &[("current", &current_str), ("total", &total_str)],
        )
    } else {
        i18n::tr_f(
            language,
            "player.time_announce_no_total",
            &[("current", &current_str)],
        )
    };
    nvda_speak(&message);
    if should_stop {
        stop_audiobook_playback(hwnd);
    }
}

fn announce_player_pitch(language: Language, pitch: f32) {
    let pitch_text = format!("{:+.0}", pitch);
    let message = i18n::tr_f(language, "player.pitch_announce", &[("pitch", &pitch_text)]);
    crate::accessibility::nvda_speak(&message);
}

fn announce_player_volume(hwnd: HWND) {
    let volume = crate::audio_player::audiobook_volume_level(hwnd);
    let language = unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    let Some(volume) = volume else {
        return;
    };
    let percent = (volume * 100.0).round().clamp(0.0, 300.0) as u32;
    let message = i18n::tr_f(
        language,
        "player.volume_announce",
        &[("pct", &percent.to_string())],
    );
    nvda_speak(&message);
}

fn announce_player_speed(language: Language, speed: f32) {
    let scaled = (speed * 10.0).round() / 10.0;
    let speed_text = if (scaled.fract() - 0.0).abs() < f32::EPSILON {
        format!("{:.0}", scaled)
    } else {
        format!("{:.1}", scaled)
    };
    let message = i18n::tr_f(language, "player.speed_announce", &[("speed", &speed_text)]);
    nvda_speak(&message);
}

fn announce_chapters_unavailable(language: Language) {
    let message = i18n::tr(language, "playback.chapters_unavailable");
    nvda_speak(&message);
}

fn seek_to_chapter_index(hwnd: HWND, chapters: &[Chapter], index: usize) {
    let Some(chapter) = chapters.get(index) else {
        return;
    };
    log_debug(&format!(
        "podcast_chapter_seek index={} start_ms={} title={}",
        index, chapter.start_ms, chapter.title
    ));
    let target_secs = chapter.start_ms.saturating_add(999) / 1000;
    unsafe {
        match seek_audiobook_to(hwnd, target_secs) {
            Ok(()) => {
                let language = if let Some(lang) = with_state(hwnd, |state| {
                    state.last_announced_chapter_index = Some(index);
                    state.settings.language
                }) {
                    lang
                } else {
                    crate::log_debug("Failed to update last announced chapter index after seek");
                    Language::default()
                };
                // Announce target chapter immediately; avoid recomputing from position too early
                // because some backends can report the old position for a short time after seek.
                let message = i18n::tr_f(
                    language,
                    "playback.chapter_announce",
                    &[("title", &chapter.title)],
                );
                nvda_speak(&message);
            }
            Err(err) => crate::log_debug(&format!("podcast_chapter_seek_error {}", err)),
        }
    }
}

fn handle_chapter_navigation(hwnd: HWND, direction: i32) {
    let (chapters, language, current_pos_ms, last_idx) = unsafe {
        with_state(hwnd, |state| {
            (
                state.active_podcast_chapters.clone(),
                state.settings.language,
                audiobook_position_ms_from_state(state),
                state.last_announced_chapter_index,
            )
        })
    }
    .unwrap_or((Vec::new(), Language::default(), None, None));
    if chapters.is_empty() {
        announce_chapters_unavailable(language);
        return;
    }
    let current_idx = current_pos_ms
        .and_then(|pos| crate::podcast::chapters::current_chapter_index(pos, &chapters))
        .or(last_idx);
    if direction < 0 && matches!(current_idx, Some(0)) {
        unsafe {
            match seek_audiobook_to(hwnd, 0) {
                Ok(()) => {
                    if with_state(hwnd, |state| {
                        // Intro segment before first chapter: no active chapter index.
                        state.last_announced_chapter_index = None;
                    })
                    .is_none()
                    {
                        crate::log_debug(
                            "Failed to clear last announced chapter index after intro seek",
                        );
                    }
                }
                Err(err) => crate::log_debug(&format!("podcast_chapter_intro_seek_error {}", err)),
            }
        }
        return;
    }

    let target = if direction > 0 {
        match current_idx {
            Some(idx) if idx + 1 < chapters.len() => Some(idx + 1),
            None => Some(0),
            _ => None,
        }
    } else {
        match current_idx {
            Some(idx) if idx > 0 => Some(idx - 1),
            _ => Some(0),
        }
    };
    if let Some(index) = target {
        seek_to_chapter_index(hwnd, &chapters, index);
    }
}

fn handle_chapter_list(hwnd: HWND) {
    let (chapters, language) = unsafe {
        with_state(hwnd, |state| {
            (
                state.active_podcast_chapters.clone(),
                state.settings.language,
            )
        })
    }
    .unwrap_or((Vec::new(), Language::default()));
    if chapters.is_empty() {
        announce_chapters_unavailable(language);
        return;
    }
    if let Some(index) =
        app_windows::podcast_chapters_window::select_chapter(hwnd, &chapters, language)
    {
        seek_to_chapter_index(hwnd, &chapters, index);
    }
}

fn is_direct_stream_url_path(path: &Path) -> bool {
    let s = path.to_string_lossy();
    s.starts_with("http://") || s.starts_with("https://")
}

fn is_stream_cache_media(path: &Path) -> bool {
    let name = path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if name.starts_with("stream_") {
        return true;
    }
    // Streaming downloads can now be renamed to media title, so they may not keep
    // the old "stream_*" prefix. Treat WebM files in podcast cache as stream media.
    if path
        .extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("webm"))
    {
        let full = path.to_string_lossy().to_ascii_lowercase();
        return full.contains("\\podcast cache\\");
    }
    false
}

fn is_direct_stream_playback_active(hwnd: HWND) -> bool {
    unsafe {
        with_state(hwnd, |state| {
            if let Some(player) = state.active_audiobook.as_ref()
                && is_direct_stream_url_path(&player.path)
            {
                return true;
            }
            state
                .docs
                .get(state.current)
                .and_then(|doc| {
                    if matches!(doc.format, FileFormat::Audiobook) {
                        doc.path.as_ref()
                    } else {
                        None
                    }
                })
                .is_some_and(|path| is_direct_stream_url_path(path))
        })
        .unwrap_or(false)
    }
}

fn handle_player_command(hwnd: HWND, command: PlayerCommand) {
    let disable_seek_rate_pitch = matches!(
        command,
        PlayerCommand::Seek(_)
            | PlayerCommand::SeekToStart
            | PlayerCommand::SeekToEnd
            | PlayerCommand::GoToTime
            | PlayerCommand::Speed(_)
            | PlayerCommand::SpeedReset
            | PlayerCommand::Pitch(_)
            | PlayerCommand::PitchReset
    ) && is_direct_stream_playback_active(hwnd);
    if disable_seek_rate_pitch {
        let language =
            unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
        let message = i18n::tr(language, "playback.direct_stream_command_disabled");
        if !message.is_empty() {
            crate::accessibility::screen_reader_speak(&message);
        }
        return;
    }
    match command {
        PlayerCommand::TogglePause => {
            toggle_audiobook_pause(hwnd);
        }
        PlayerCommand::Stop => {
            stop_audiobook_playback(hwnd);
        }
        PlayerCommand::Seek(amount) => {
            seek_audiobook(hwnd, amount);
        }
        PlayerCommand::SeekToStart => unsafe {
            if let Err(err) = seek_audiobook_to(hwnd, 0) {
                if err == "No active audiobook" {
                    let restart_path = with_state(hwnd, |state| {
                        let doc = state.docs.get(state.current)?;
                        if matches!(doc.format, FileFormat::Audiobook) {
                            doc.path.clone()
                        } else {
                            None
                        }
                    })
                    .flatten();
                    if let Some(path) = restart_path {
                        start_audiobook_at(hwnd, &path, 0);
                    }
                } else {
                    log_debug(&format!("Seek to start failed: {}", err));
                }
            }
        },
        PlayerCommand::SeekToEnd => unsafe {
            if let Some(result) = with_state(hwnd, |state| {
                let player = state.active_audiobook.as_ref()?;
                let live_total = player.duration_secs().map(|s| s.max(0.0).floor() as u64);
                let file_total = crate::audio_player::audiobook_duration_secs(&player.path);
                let total = match (file_total, live_total) {
                    // Prefer the larger duration to avoid BASS underestimation on long MP3.
                    (Some(file), Some(live)) => Some(file.max(live)),
                    (Some(file), None) => Some(file),
                    (None, Some(live)) => Some(live),
                    (None, None) => None,
                }?;
                // Seek just before the exact end to avoid immediate stop logic.
                let target = total.saturating_sub(1);
                Some(seek_audiobook_to(hwnd, target))
            })
            .flatten()
                && let Err(err) = result
            {
                log_debug(&format!("Seek to end failed: {}", err));
            }
        },
        PlayerCommand::Volume(delta) => {
            change_audiobook_volume(hwnd, delta);
            announce_player_volume(hwnd);
        }
        PlayerCommand::VolumeReset => {
            reset_audiobook_volume(hwnd);
            announce_player_volume(hwnd);
        }
        PlayerCommand::Speed(delta) => {
            let language =
                unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
            let speed = change_audiobook_speed(hwnd, delta);
            if let Some(speed) = speed {
                announce_player_speed(language, speed);
            }
        }
        PlayerCommand::Pitch(delta) => {
            let language =
                unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
            let pitch = change_audiobook_pitch(hwnd, delta);
            if let Some(pitch) = pitch {
                announce_player_pitch(language, pitch);
            }
        }
        PlayerCommand::SpeedReset => {
            let language =
                unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
            let speed = reset_audiobook_speed(hwnd);
            if let Some(speed) = speed {
                announce_player_speed(language, speed);
            }
        }
        PlayerCommand::PitchReset => {
            let language =
                unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
            let pitch = reset_audiobook_pitch(hwnd);
            if let Some(pitch) = pitch {
                announce_player_pitch(language, pitch);
            }
        }
        PlayerCommand::MuteToggle => {
            toggle_audiobook_mute(hwnd);
        }
        PlayerCommand::GoToTime => {
            app_windows::go_to_time_window::open(hwnd);
        }
        PlayerCommand::AnnounceTime => {
            announce_player_time(hwnd);
        }
        PlayerCommand::ChapterPrev => {
            handle_chapter_navigation(hwnd, -1);
        }
        PlayerCommand::ChapterNext => {
            handle_chapter_navigation(hwnd, 1);
        }
        PlayerCommand::ChapterList => {
            handle_chapter_list(hwnd);
        }
        PlayerCommand::TrackPrev => {
            if !switch_audio_playlist_relative(hwnd, -1) {
                crate::log_debug("Audio player: no previous track available in playlist");
            }
        }
        PlayerCommand::TrackNext => {
            if !switch_audio_playlist_relative(hwnd, 1) {
                crate::log_debug("Audio player: no next track available in playlist");
            }
        }
        PlayerCommand::BlockNavigation | PlayerCommand::None => {}
    }
}

fn has_secondary_window_open(hwnd: HWND) -> bool {
    unsafe {
        with_state(hwnd, |state| {
            state.find_dialog.0 != 0
                || state.replace_dialog.0 != 0
                || state.options_dialog.0 != 0
                || state.help_window.0 != 0
                || state.changelog_window.0 != 0
                || state.donations_window.0 != 0
                || state.bookmarks_window.0 != 0
                || state.dictionary_window.0 != 0
                || state.dictionary_entry_dialog.0 != 0
                || state.wiktionary_window.0 != 0
                || state.wikipedia_window.0 != 0
                || state.prompt_window.0 != 0
                || state.podcast_window.0 != 0
                || state.podcast_save_window.0 != 0
                || state.update_progress_window.0 != 0
                || state.batch_audiobooks_window.0 != 0
                || state.convert_audio_window.0 != 0
                || state.podcasts_window.0 != 0
                || state.podcasts_add_dialog.0 != 0
                || state.podcasts_categories_dialog.0 != 0
                || state.podcasts_description_dialog.0 != 0
                || state.rss_window.0 != 0
                || state.rss_add_dialog.0 != 0
                || state.go_to_time_dialog.0 != 0
        })
        .unwrap_or(false)
    }
}

fn should_force_editor_focus_on_foreground(hwnd: HWND) -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        with_state(hwnd, |state| {
            let is_reader_mode = state
                .docs
                .get(state.current)
                .map(|doc| matches!(doc.format, FileFormat::Audiobook))
                .unwrap_or(false);
            let audiobook_progress_in_foreground =
                state.audiobook_progress.0 != 0 && foreground == state.audiobook_progress;
            state.update_progress_window.0 == 0
                && !audiobook_progress_in_foreground
                && !is_reader_mode
        })
        .unwrap_or(false)
    }
}

fn force_active_editor_focus(hwnd: HWND) {
    if !should_force_editor_focus_on_foreground(hwnd) {
        return;
    }
    unsafe {
        if let Some(hwnd_edit) = get_active_edit(hwnd) {
            SetFocus(hwnd_edit);
            SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
            SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
            crate::log_if_err!(PostMessageW(
                hwnd,
                WM_NEXTDLGCTL,
                WPARAM(hwnd_edit.0 as usize),
                LPARAM(1)
            ));
            NotifyWinEvent(
                EVENT_OBJECT_FOCUS,
                hwnd_edit,
                OBJID_CLIENT.0,
                CHILDID_SELF as i32,
            );
        }
    }
}

fn schedule_editor_focus_retry(hwnd: HWND) {
    unsafe {
        if SetTimer(hwnd, FOCUS_EDITOR_TIMER_ID, 80, None) == 0 {
            crate::log_debug("Failed to set FOCUS_EDITOR_TIMER_ID");
        }
        if SetTimer(hwnd, FOCUS_EDITOR_TIMER_ID2, 200, None) == 0 {
            crate::log_debug("Failed to set FOCUS_EDITOR_TIMER_ID2");
        }
        if SetTimer(hwnd, FOCUS_EDITOR_TIMER_ID3, 350, None) == 0 {
            crate::log_debug("Failed to set FOCUS_EDITOR_TIMER_ID3");
        }
        if SetTimer(hwnd, FOCUS_EDITOR_TIMER_ID4, 600, None) == 0 {
            crate::log_debug("Failed to set FOCUS_EDITOR_TIMER_ID4");
        }
    }
}

#[derive(Default)]
pub(crate) struct AppState {
    hwnd_tab: HWND,
    hwnd_status: HWND,
    docs: Vec<Document>,
    current: usize,
    untitled_count: usize,
    hfont: HFONT,
    hfont_custom: bool,
    hmenu_recent: HMENU,
    recent_files: Vec<PathBuf>,
    settings: AppSettings,
    bookmarks: BookmarkStore,
    find_dialog: HWND,
    replace_dialog: HWND,
    options_dialog: HWND,
    help_window: HWND,
    changelog_window: HWND,
    donations_window: HWND,
    bookmarks_window: HWND,
    dictionary_window: HWND,
    dictionary_entry_dialog: HWND,
    wiktionary_window: HWND,
    wikipedia_window: HWND,
    prompt_window: HWND,
    podcast_window: HWND,
    podcast_save_window: HWND,
    update_progress_window: HWND,
    batch_audiobooks_window: HWND,
    convert_audio_window: HWND,
    podcasts_window: HWND,
    podcasts_add_dialog: HWND,
    podcasts_categories_dialog: HWND,
    podcasts_description_dialog: HWND,
    rss_window: HWND,
    rss_add_dialog: HWND, // Input dialog for RSS
    go_to_time_dialog: HWND,
    playback_menu: HMENU,
    find_msg: u32,
    find_text: Vec<u16>,
    replace_text: Vec<u16>,
    find_replace: Option<FINDREPLACEW>,
    replace_replace: Option<FINDREPLACEW>,
    last_find_flags: FINDREPLACE_FLAGS,
    find_use_regex: bool,
    find_dot_matches_newline: bool,
    find_wrap_around: bool,
    find_match_case: bool,
    find_whole_word: bool,
    find_replace_in_selection: bool,
    find_replace_in_all_docs: bool,
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
    audiobook_session_id: u64,
    last_stopped_audiobook: Option<std::path::PathBuf>,
    active_podcast_episode_url: Option<String>,
    active_podcast_episode_title: Option<String>,
    active_podcast_episode_cache: Option<PathBuf>,
    podcast_chapters_cache: HashMap<String, Option<Vec<Chapter>>>,
    pending_podcast_chapters_key: Option<String>,
    active_podcast_chapters_key: Option<String>,
    active_podcast_chapters: Vec<Chapter>,
    last_announced_chapter_index: Option<usize>,
    available_audio_tracks: Vec<crate::ffmpeg_source::AudioStreamInfo>,
    selected_audio_track: Option<i32>,
    audio_playlist: Vec<PathBuf>,
    audio_playlist_index: Option<usize>,
    audio_ffmpeg_retry_for: Option<PathBuf>,
    voice_panel_visible: bool,
    voice_label_engine: HWND,
    voice_combo_engine: HWND,
    voice_label_language: HWND,
    voice_combo_language: HWND,
    voice_language_codes: Vec<String>,
    voice_label_voice: HWND,
    voice_combo_voice: HWND,
    voice_button_insert_tag: HWND,
    voice_label_speed: HWND,
    voice_combo_speed: HWND,
    voice_edit_speed: HWND,
    voice_label_pitch: HWND,
    voice_combo_pitch: HWND,
    voice_edit_pitch: HWND,
    voice_label_volume: HWND,
    voice_combo_volume: HWND,
    voice_edit_volume: HWND,
    voice_checkbox_multilingual: HWND,
    voice_favorites_visible: bool,
    voice_label_favorites: HWND,
    voice_combo_favorites: HWND,
    voice_combo_voice_proc: WNDPROC,
    voice_combo_favorites_proc: WNDPROC,
    voice_context_voice: Option<FavoriteVoice>,
    find_in_files_cache: Option<FindInFilesCache>,
    pending_find_in_files: Option<PendingFindInFilesSelection>,
    normalize_undo: Option<NormalizeUndo>,
    undo_action_label: Option<String>,
    normalize_skip_change: bool,
    spellcheck_manager: spellcheck::SpellcheckManager,
    spellcheck_last_announce: Option<SpellcheckAnnounceKey>,
    spellcheck_context: Option<SpellcheckContextMenuState>,
    spellcheck_space_trigger: Option<HWND>,
    spellcheck_typing_in_progress: bool,
    spellcheck_highlight_pending: Option<HWND>,
    spellcheck_last_highlighted_line: Option<(isize, i32)>, // (doc_id, line_index)
    dictionary_context_menu: HMENU,
    dictionary_context_word: String,
    dictionary_context_language: Language,
    dictionary_context_pref: String,
    dictionary_context_loaded: bool,
    dictionary_context_expanded: bool,
    dictionary_cache: HashMap<String, Vec<String>>,
    dictionary_pending_lookup: Option<String>,
    dictionary_prefetch_generation: usize,
    large_text_editors: HashSet<isize>,
}

#[derive(Clone)]
pub(crate) struct PendingFindInFilesSelection {
    pub(crate) path: PathBuf,
    pub(crate) snippet: String,
    pub(crate) term: String,
    pub(crate) start_utf16: i32,
    pub(crate) len_utf16: i32,
}

#[derive(Default, Serialize, Deserialize)]
struct RecentFileStore {
    files: Vec<String>,
}

fn main() -> windows::core::Result<()> {
    // Initialize COM for the main UI thread (STA)
    let _com = com_guard::ComGuard::new_sta().ok();

    // Estrai le dipendenze embedded (DLL, certificati, ecc.)
    if let Err(e) = embedded_deps::extract_all() {
        log_debug(&format!("Warning: Failed to extract embedded deps: {}", e));
    }
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
    if args.iter().any(|arg| arg == "-h" || arg == "--help") {
        print_cli_help();
        std::process::exit(0);
    }
    if args.iter().any(|arg| arg == "--version") {
        println!("Sonarpad {}", env!("CARGO_PKG_VERSION"));
        std::process::exit(0);
    }
    updater::cleanup_backup_on_start();
    updater::cleanup_update_lock_on_start();
    updater::cleanup_update_temp_on_start();

    // Inizializza Sentry per crash reporting (opt-in)
    {
        let settings = load_settings();
        sentry_integration::init(settings.send_crash_reports, option_env!("SENTRY_DSN"));
    }

    // Inizializza telemetry per hang diagnostics
    telemetry::init();

    // Installa panic hook (logga + invia a Sentry)
    sentry_integration::install_panic_hook();

    // Error boundary: cattura errori fatali
    if let Err(e) = run_app(&args) {
        sentry_integration::capture_fatal_windows_error("run_app", &e);
        sentry_integration::flush(2);
        return Err(e);
    }

    Ok(())
}

fn print_cli_help() {
    println!("Sonarpad {}", env!("CARGO_PKG_VERSION"));
    println!("Usage:");
    println!("  sonarpad.exe [OPTIONS] [FILES...]");
    println!();
    println!("Options:");
    println!("  -h, --help         Show this help message and exit");
    println!("  --version          Show version and exit");
    println!("  --self-update      Internal updater mode (do not use manually)");
    println!();
    println!("Arguments:");
    println!("  FILES...           One or more files to open");
}

fn is_large_text_editor(hwnd: HWND, hwnd_edit: HWND) -> bool {
    unsafe {
        with_state(hwnd, |state| {
            state.large_text_editors.contains(&hwnd_edit.0)
        })
        .unwrap_or(false)
    }
}

/// Core dell'applicazione - separato per error boundary
fn run_app(args: &[String]) -> windows::core::Result<()> {
    unsafe {
        crate::log_if_err!(LoadLibraryW(w!("Msftedit.dll")));
        let hinstance = HINSTANCE(GetModuleHandleW(None)?.0);
        let class_name = w!("SonarpadWin32");

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
        let current_version = env!("CARGO_PKG_VERSION");
        let mut settings = load_settings();
        if settings.last_seen_changelog_version != current_version
            && app_windows::rss_window::sync_default_sources_for_settings(&mut settings)
        {
            save_settings(settings.clone());
        }
        crate::settings::sync_start_menu_shortcuts(&settings);
        let file_to_open = extra_paths.first().cloned();
        if !extra_paths.is_empty() && settings.open_behavior == OpenBehavior::NewTab {
            let existing = FindWindowW(class_name, PCWSTR::null());
            if existing.0 != 0 {
                // Send paths to existing window via WM_COPYDATA
                let joined = extra_paths.join("|");
                let wide = to_wide(&joined);
                let mut cds = COPYDATASTRUCT {
                    dwData: 1, // 1 = open files
                    cbData: (wide.len() * 2) as u32,
                    lpData: wide.as_ptr() as *mut std::ffi::c_void,
                };
                let mut existing_pid = 0u32;
                let existing_thread = GetWindowThreadProcessId(existing, Some(&mut existing_pid));
                if existing_thread == 0 {
                    log_debug("GetWindowThreadProcessId failed for existing window");
                } else if existing_pid != 0 {
                    crate::log_if_err!(AllowSetForegroundWindow(existing_pid));
                }
                SendMessageW(
                    existing,
                    WM_COPYDATA,
                    WPARAM(0),
                    LPARAM(&mut cds as *mut _ as isize),
                );
                return Ok(());
            }
        }
        let lp_param = &file_to_open as *const Option<String> as *const std::ffi::c_void;

        let title_wide = to_wide(crate::settings::app_display_name(&settings));
        let hwnd = CreateWindowExW(
            Default::default(),
            class_name,
            PCWSTR(title_wide.as_ptr()),
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
        refresh_voice_panel(hwnd);
        crate::log_if_err!(PostMessageW(
            hwnd,
            WM_CHECK_PENDING_UPDATE,
            WPARAM(0),
            LPARAM(0)
        ));

        let mut show_changelog = false;
        let mut cleanup_legacy_context_menu = false;
        with_state(hwnd, |state| {
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
                cleanup_legacy_context_menu = true;
            }
        });
        if cleanup_legacy_context_menu {
            crate::settings::cleanup_legacy_context_menu_entries();
        }
        if show_changelog {
            let hwnd_val = hwnd.0;
            std::thread::spawn(move || {
                std::thread::sleep(std::time::Duration::from_secs(1));
                if let Err(e) =
                    PostMessageW(HWND(hwnd_val), WM_SHOW_CHANGELOG, WPARAM(0), LPARAM(0))
                {
                    crate::log_debug(&format!("Failed to post show changelog message: {}", e));
                }
            });
        }

        let check_updates =
            with_state(hwnd, |state| state.settings.check_updates_on_startup).unwrap_or(true);
        if check_updates {
            let hwnd_val = hwnd.0;
            std::thread::spawn(move || {
                std::thread::sleep(std::time::Duration::from_secs(1));
                if let Err(e) =
                    PostMessageW(HWND(hwnd_val), WM_AUTO_UPDATE_CHECK, WPARAM(0), LPARAM(0))
                {
                    crate::log_debug(&format!("Failed to post startup update check: {}", e));
                }
            });
        }

        let accel = create_accelerators();
        let mut msg = MSG::default();

        // Avvia watchdog per rilevare freeze
        let watchdog = watchdog::start_watchdog(watchdog::WatchdogConfig::default());
        let mut heartbeat_counter = 0u32;

        while GetMessageW(&mut msg, HWND(0), 0, 0).into() {
            // Heartbeat ogni ~100 messaggi per non impattare performance
            heartbeat_counter = heartbeat_counter.wrapping_add(1);
            if heartbeat_counter.is_multiple_of(100) {
                watchdog.heartbeat();
            }
            // Priority 1: Global navigation keys (Ctrl+Tab)
            if msg.message == WM_KEYDOWN
                && msg.wParam.0 as u32 == VK_TAB.0 as u32
                && (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0
            {
                let options_hwnd =
                    with_state(hwnd, |state| state.options_dialog).unwrap_or(HWND(0));
                if options_hwnd.0 != 0 {
                    if app_windows::options_window::handle_navigation(options_hwnd, &msg) {
                        continue;
                    }
                } else {
                    // Switch tabs in main window
                    let tab_hwnd = with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0));
                    if tab_hwnd.0 != 0 {
                        let count = SendMessageW(
                            tab_hwnd,
                            windows::Win32::UI::Controls::TCM_GETITEMCOUNT,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0;
                        if count > 1 {
                            let cur = SendMessageW(
                                tab_hwnd,
                                windows::Win32::UI::Controls::TCM_GETCURSEL,
                                WPARAM(0),
                                LPARAM(0),
                            )
                            .0;
                            let shift_down =
                                (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
                            let next = if shift_down {
                                if cur == 0 { count - 1 } else { cur - 1 }
                            } else if cur == count - 1 {
                                0
                            } else {
                                cur + 1
                            };
                            editor_manager::select_tab(hwnd, next as usize);
                        }
                    }
                    continue;
                }
            }
            if msg.message == WM_CONTEXTMENU && msg.lParam.0 == -1 {
                let rss_hwnd = with_state(hwnd, |state| state.rss_window).unwrap_or(HWND(0));
                if rss_hwnd.0 != 0 {
                    let mut cur = msg.hwnd;
                    let mut rss_target = false;
                    while cur.0 != 0 {
                        if cur == rss_hwnd {
                            app_windows::rss_window::show_context_menu_from_keyboard(rss_hwnd);
                            rss_target = true;
                            break;
                        }
                        cur = GetParent(cur);
                    }
                    if rss_target {
                        continue;
                    }
                }
                let podcasts_hwnd =
                    with_state(hwnd, |state| state.podcasts_window).unwrap_or(HWND(0));
                if podcasts_hwnd.0 != 0 {
                    let mut cur = msg.hwnd;
                    let mut podcasts_target = false;
                    while cur.0 != 0 {
                        if cur == podcasts_hwnd {
                            app_windows::podcasts_window::show_context_menu_from_keyboard(
                                podcasts_hwnd,
                            );
                            podcasts_target = true;
                            break;
                        }
                        cur = GetParent(cur);
                    }
                    if podcasts_target {
                        continue;
                    }
                }
            }
            if msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN {
                let key = msg.wParam.0 as u32;
                let is_context_key = key == u32::from(VK_APPS.0)
                    || (key == u32::from(VK_F10.0) && GetKeyState(VK_SHIFT.0 as i32) < 0);
                if is_context_key {
                    let rss_hwnd = with_state(hwnd, |state| state.rss_window).unwrap_or(HWND(0));
                    if rss_hwnd.0 != 0 {
                        let mut cur = msg.hwnd;
                        let mut rss_target = false;
                        while cur.0 != 0 {
                            if cur == rss_hwnd {
                                app_windows::rss_window::show_context_menu_from_keyboard(rss_hwnd);
                                rss_target = true;
                                break;
                            }
                            cur = GetParent(cur);
                        }
                        if rss_target {
                            continue;
                        }
                    }
                    let podcasts_hwnd =
                        with_state(hwnd, |state| state.podcasts_window).unwrap_or(HWND(0));
                    if podcasts_hwnd.0 != 0 {
                        let mut cur = msg.hwnd;
                        let mut podcasts_target = false;
                        while cur.0 != 0 {
                            if cur == podcasts_hwnd {
                                app_windows::podcasts_window::show_context_menu_from_keyboard(
                                    podcasts_hwnd,
                                );
                                podcasts_target = true;
                                break;
                            }
                            cur = GetParent(cur);
                        }
                        if podcasts_target {
                            continue;
                        }
                    }
                }
            }
            if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
                && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32
            {
                let rss_hwnd = with_state(hwnd, |state| state.rss_window).unwrap_or(HWND(0));
                if rss_hwnd.0 != 0
                    && let Some(hwnd_edit) = get_active_edit(hwnd)
                    && GetFocus() == hwnd_edit
                    && editor_manager::current_document_is_from_rss(hwnd)
                {
                    app_windows::rss_window::focus_library(rss_hwnd);
                    continue;
                }
                let save_hwnd =
                    with_state(hwnd, |state| state.podcast_save_window).unwrap_or(HWND(0));
                if save_hwnd.0 != 0 {
                    crate::log_if_err!(PostMessageW(save_hwnd, WM_COMMAND, WPARAM(2), LPARAM(0)));
                    continue;
                }
                let update_progress_hwnd =
                    with_state(hwnd, |state| state.update_progress_window).unwrap_or(HWND(0));
                if update_progress_hwnd.0 != 0 {
                    crate::log_if_err!(PostMessageW(
                        update_progress_hwnd,
                        WM_COMMAND,
                        WPARAM(2),
                        LPARAM(0)
                    ));
                    continue;
                }
                if let Some(hwnd_edit) = get_active_edit(hwnd)
                    && GetFocus() == hwnd_edit
                    && focus_find_in_files_results()
                {
                    continue;
                }
            }
            if (msg.message == WM_KEYDOWN
                && msg.wParam.0 as u32 == u32::from(VK_MEDIA_PLAY_PAUSE.0))
                || (msg.message == WM_APPCOMMAND
                    && appcommand_from_lparam(msg.lParam) == APPCOMMAND_MEDIA_PLAY_PAUSE)
            {
                let has_player =
                    with_state(hwnd, |state| state.active_audiobook.is_some()).unwrap_or(false);
                if has_player {
                    handle_player_command(hwnd, PlayerCommand::TogglePause);
                    continue;
                }
                if is_tts_active(hwnd) {
                    tts_engine::toggle_tts_pause(hwnd);
                    continue;
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
                let alt_down = (GetKeyState(VK_MENU.0 as i32) & (0x8000u16 as i16)) != 0;
                if ctrl_down
                    && !shift_down
                    && !alt_down
                    && let Some(hwnd_edit) = get_active_edit(hwnd)
                    && GetFocus() == hwnd_edit
                {
                    if !editor_manager::try_normalize_undo(hwnd) {
                        editor_manager::undo_active_edit_skip_navigation(hwnd);
                    }
                    continue;
                }
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F1.0) {
                app_windows::help_window::open(hwnd);
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F2.0) {
                updater::check_for_update(hwnd, true);
                continue;
            }
            if msg.message == WM_KEYDOWN
                && msg.wParam.0 as u32 == u32::from(VK_F9.0)
                && is_tts_active(hwnd)
            {
                cycle_favorite_voice(hwnd, -1);
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F10.0) {
                // F10 is normally used for menu, so only use it for voice cycling during TTS
                if is_tts_active(hwnd) {
                    cycle_favorite_voice(hwnd, 1);
                    continue;
                }
            }
            // F7/F8 for spelling navigation
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F7.0) {
                go_to_spelling_error(hwnd, false);
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == u32::from(VK_F8.0) {
                go_to_spelling_error(hwnd, true);
                continue;
            }
            if msg.message == WM_KEYDOWN
                && msg.wParam.0 as u32 == VK_TAB.0 as u32
                && (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) == 0
                && handle_voice_panel_tab(hwnd)
            {
                continue;
            }

            // Enter on the voice panel "Insert Tag" button: behave like the options dialog
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let is_insert_tag = focus.0 != 0
                    && with_state(hwnd, |state| focus == state.voice_button_insert_tag)
                        .unwrap_or(false);
                if is_insert_tag {
                    insert_voice_tag_from_voice_panel(hwnd);
                    if let Some(hwnd_edit) = get_active_edit(hwnd) {
                        SetFocus(hwnd_edit);
                    }
                    continue;
                }
            }

            let mut handled = false;
            with_state(hwnd, |state| {
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
                        || state.donations_window.0 != 0
                        || state.dictionary_window.0 != 0
                        || state.podcast_window.0 != 0;
                    let secondary_open = secondary_open
                        || state.dictionary_entry_dialog.0 != 0
                        || state.go_to_time_dialog.0 != 0
                        || state.podcasts_add_dialog.0 != 0
                        || state.podcasts_categories_dialog.0 != 0
                        || state.podcasts_description_dialog.0 != 0;

                    // Exclude voice panel controls from player keyboard handling
                    let is_voice_panel_control = is_focus_in_voice_panel(hwnd);

                    let is_main_target = msg.hwnd == hwnd || IsChild(hwnd, msg.hwnd).as_bool();
                    if is_audiobook && !secondary_open && is_main_target && !is_voice_panel_control
                    {
                        let command =
                            handle_player_keyboard(&msg, state.settings.audiobook_skip_seconds);
                        if !matches!(command, PlayerCommand::None) {
                            if matches!(command, PlayerCommand::BlockNavigation) {
                                handled = true;
                                return;
                            }
                            let is_stop = matches!(command, PlayerCommand::Stop);
                            let podcasts_window = state.podcasts_window;
                            if is_stop {
                                // close_current_document() already stops audiobook playback
                                // for audiobook tabs, so avoid duplicate stop work here.
                                editor_manager::close_current_document(hwnd);
                                if podcasts_window.0 != 0 {
                                    SetForegroundWindow(podcasts_window);
                                    app_windows::podcasts_window::focus_library(podcasts_window);
                                }
                            } else {
                                handle_player_command(hwnd, command);
                            }
                            handled = true;
                            return;
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
                if state.go_to_time_dialog.0 != 0
                    && app_windows::go_to_time_window::handle_navigation(
                        state.go_to_time_dialog,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if app_windows::help_window::handle_readonly_navigation(&msg) {
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
                if state.donations_window.0 != 0 {
                    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_TAB.0 as u32 {
                        app_windows::help_window::handle_tab(state.donations_window);
                        handled = true;
                        return;
                    }

                    if handle_accessibility(state.donations_window, &msg) {
                        handled = true;
                        return;
                    }
                }

                if state.options_dialog.0 != 0
                    && app_windows::options_window::handle_navigation(state.options_dialog, &msg)
                {
                    handled = true;
                    return;
                }

                if state.podcast_window.0 != 0
                    && app_windows::podcast_window::handle_navigation(state.podcast_window, &msg)
                {
                    handled = true;
                    return;
                }

                if state.podcast_save_window.0 != 0
                    && app_windows::podcast_save_window::handle_navigation(
                        state.podcast_save_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.update_progress_window.0 != 0
                    && app_windows::podcast_save_window::handle_navigation(
                        state.update_progress_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }

                if state.audiobook_progress.0 != 0
                    && app_windows::audiobook_window::handle_navigation(
                        state.audiobook_progress,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }

                if state.bookmarks_window.0 != 0
                    && app_windows::bookmarks_window::handle_navigation(
                        state.bookmarks_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }

                if state.dictionary_window.0 != 0
                    && app_windows::dictionary_window::handle_navigation(
                        state.dictionary_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }

                if state.wiktionary_window.0 != 0
                    && app_windows::wiktionary_window::handle_navigation(
                        state.wiktionary_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.wikipedia_window.0 != 0
                    && app_windows::wikipedia_window::handle_navigation(
                        state.wikipedia_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }

                if state.dictionary_entry_dialog.0 != 0
                    && handle_accessibility(state.dictionary_entry_dialog, &msg)
                {
                    handled = true;
                    return;
                }

                if state.batch_audiobooks_window.0 != 0
                    && app_windows::batch_audiobooks_window::handle_navigation(
                        state.batch_audiobooks_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.batch_audiobooks_window.0 != 0
                    && handle_accessibility(state.batch_audiobooks_window, &msg)
                {
                    handled = true;
                    return;
                }
                if state.convert_audio_window.0 != 0
                    && app_windows::convert_audio_window::handle_navigation(
                        state.convert_audio_window,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.convert_audio_window.0 != 0
                    && handle_accessibility(state.convert_audio_window, &msg)
                {
                    handled = true;
                    return;
                }

                if state.prompt_window.0 != 0
                    && app_windows::prompt_window::handle_navigation(state.prompt_window, &msg)
                {
                    handled = true;
                    return;
                }

                if state.rss_window.0 != 0 && handle_accessibility(state.rss_window, &msg) {
                    handled = true;
                    return;
                }

                if state.rss_add_dialog.0 != 0 && handle_accessibility(state.rss_add_dialog, &msg) {
                    handled = true;
                    return;
                }
                if state.podcasts_description_dialog.0 != 0
                    && app_windows::podcasts_window::handle_navigation(
                        state.podcasts_description_dialog,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.podcasts_categories_dialog.0 != 0
                    && app_windows::podcasts_window::handle_navigation(
                        state.podcasts_categories_dialog,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.podcasts_add_dialog.0 != 0
                    && app_windows::podcasts_window::handle_navigation(
                        state.podcasts_add_dialog,
                        &msg,
                    )
                {
                    handled = true;
                    return;
                }
                if state.podcasts_window.0 != 0
                    && app_windows::podcasts_window::handle_navigation(state.podcasts_window, &msg)
                {
                    handled = true;
                }
            });
            if handled {
                continue;
            }
            if handle_custom_shortcuts(hwnd, &msg) {
                continue;
            }
            if accel.0 != 0 && TranslateAcceleratorW(hwnd, accel, &msg) != 0 {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        // Ferma watchdog prima di uscire
        watchdog.stop();

        Ok(())
    }
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    crate::panic_guard::guard(
        "wndproc",
        || DefWindowProcW(hwnd, msg, wparam, lparam),
        || unsafe { wndproc_inner(hwnd, msg, wparam, lparam) },
    )
}

unsafe fn wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    if let Some(find_msg) = with_state(hwnd, |state| state.find_msg)
        && msg == find_msg
    {
        handle_find_message(hwnd, lparam);
        return LRESULT(0);
    }

    match msg {
        WM_CREATE => {
            let icc = INITCOMMONCONTROLSEX {
                dwSize: size_of::<INITCOMMONCONTROLSEX>() as u32,
                dwICC: ICC_TAB_CLASSES | ICC_BAR_CLASSES,
            };
            InitCommonControlsEx(&icc);

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
            let hwnd_status = CreateWindowExW(
                Default::default(),
                STATUSCLASSNAMEW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(MAIN_STATUS_ID as isize),
                HINSTANCE(0),
                None,
            );

            let find_msg = RegisterWindowMessageW(w!("commdlg_FindReplace"));
            let settings = load_settings();
            let hfont = create_ui_font(&settings.editor_font_face, None, settings.text_size)
                .unwrap_or_else(|| HFONT(GetStockObject(DEFAULT_GUI_FONT).0));
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
            let label_language = CreateWindowExW(
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
            let combo_language = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                140,
                hwnd,
                HMENU(VOICE_PANEL_ID_LANGUAGE as isize),
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
            let button_insert_tag = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(empty_label.as_ptr()),
                WS_CHILD | WS_TABSTOP,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(VOICE_PANEL_ID_INSERT_TAG as isize),
                HINSTANCE(0),
                None,
            );
            let label_speed = CreateWindowExW(
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
            let combo_speed = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                140,
                hwnd,
                HMENU(VOICE_PANEL_ID_SPEED as isize),
                HINSTANCE(0),
                None,
            );
            let edit_speed = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(VOICE_PANEL_ID_SPEED_EDIT as isize),
                HINSTANCE(0),
                None,
            );
            let label_pitch = CreateWindowExW(
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
            let combo_pitch = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                140,
                hwnd,
                HMENU(VOICE_PANEL_ID_PITCH as isize),
                HINSTANCE(0),
                None,
            );
            let edit_pitch = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(VOICE_PANEL_ID_PITCH_EDIT as isize),
                HINSTANCE(0),
                None,
            );
            let label_volume = CreateWindowExW(
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
            let combo_volume = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                0,
                0,
                0,
                140,
                hwnd,
                HMENU(VOICE_PANEL_ID_VOLUME as isize),
                HINSTANCE(0),
                None,
            );
            let edit_volume = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(VOICE_PANEL_ID_VOLUME_EDIT as isize),
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
            let combo_voice_proc = if combo_voice.0 != 0 {
                let proc_ptr = voice_combo_subclass_proc as *const () as usize;
                let old = SetWindowLongPtrW(combo_voice, GWLP_WNDPROC, proc_ptr as isize);
                std::mem::transmute::<isize, WNDPROC>(old)
            } else {
                None
            };
            let combo_favorites_proc = if combo_favorites.0 != 0 {
                let proc_ptr = voice_combo_subclass_proc as *const () as usize;
                let old = SetWindowLongPtrW(combo_favorites, GWLP_WNDPROC, proc_ptr as isize);
                std::mem::transmute::<isize, WNDPROC>(old)
            } else {
                None
            };
            for control in [
                label_engine,
                combo_engine,
                label_language,
                combo_language,
                label_voice,
                combo_voice,
                button_insert_tag,
                label_speed,
                combo_speed,
                edit_speed,
                label_pitch,
                combo_pitch,
                edit_pitch,
                label_volume,
                combo_volume,
                edit_volume,
                checkbox_multilingual,
                label_favorites,
                combo_favorites,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
                ShowWindow(control, SW_HIDE);
            }
            let state = Box::new(AppState {
                hwnd_tab,
                hwnd_status,
                docs: Vec::new(),
                current: 0,
                untitled_count: 0,
                hfont,
                hfont_custom: !settings.editor_font_face.trim().is_empty() && hfont.0 != 0,
                hmenu_recent: recent_menu,
                recent_files,
                settings: settings.clone(),
                bookmarks,
                find_dialog: HWND(0),
                replace_dialog: HWND(0),
                options_dialog: HWND(0),
                help_window: HWND(0),
                changelog_window: HWND(0),
                donations_window: HWND(0),
                bookmarks_window: HWND(0),
                dictionary_window: HWND(0),
                dictionary_entry_dialog: HWND(0),
                wiktionary_window: HWND(0),
                wikipedia_window: HWND(0),
                prompt_window: HWND(0),
                podcast_window: HWND(0),
                rss_window: HWND(0),
                podcasts_window: HWND(0),
                podcasts_add_dialog: HWND(0),
                podcasts_categories_dialog: HWND(0),
                podcasts_description_dialog: HWND(0),
                rss_add_dialog: HWND(0),
                go_to_time_dialog: HWND(0),
                playback_menu: HMENU(0),
                podcast_save_window: HWND(0),
                update_progress_window: HWND(0),
                batch_audiobooks_window: HWND(0),
                convert_audio_window: HWND(0),

                find_msg,
                find_text: vec![0u16; 256],
                replace_text: vec![0u16; 256],
                find_replace: None,
                replace_replace: None,
                last_find_flags: FINDREPLACE_FLAGS(0),
                find_use_regex: false,
                find_dot_matches_newline: false,
                find_wrap_around: true,
                find_match_case: false,
                find_whole_word: false,
                find_replace_in_selection: false,
                find_replace_in_all_docs: false,
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
                audiobook_session_id: 0,
                last_stopped_audiobook: None,
                active_podcast_episode_url: None,
                active_podcast_episode_title: None,
                active_podcast_episode_cache: None,
                podcast_chapters_cache: HashMap::new(),
                pending_podcast_chapters_key: None,
                active_podcast_chapters_key: None,
                active_podcast_chapters: Vec::new(),
                last_announced_chapter_index: None,
                available_audio_tracks: Vec::new(),
                selected_audio_track: None,
                audio_playlist: Vec::new(),
                audio_playlist_index: None,
                audio_ffmpeg_retry_for: None,
                voice_panel_visible: false,
                voice_label_engine: label_engine,
                voice_combo_engine: combo_engine,
                voice_label_language: label_language,
                voice_combo_language: combo_language,
                voice_language_codes: Vec::new(),
                voice_label_voice: label_voice,
                voice_combo_voice: combo_voice,
                voice_button_insert_tag: button_insert_tag,
                voice_label_speed: label_speed,
                voice_combo_speed: combo_speed,
                voice_edit_speed: edit_speed,
                voice_label_pitch: label_pitch,
                voice_combo_pitch: combo_pitch,
                voice_edit_pitch: edit_pitch,
                voice_label_volume: label_volume,
                voice_combo_volume: combo_volume,
                voice_edit_volume: edit_volume,
                voice_checkbox_multilingual: checkbox_multilingual,
                voice_favorites_visible: false,
                voice_label_favorites: label_favorites,
                voice_combo_favorites: combo_favorites,
                voice_combo_voice_proc: combo_voice_proc,
                voice_combo_favorites_proc: combo_favorites_proc,
                voice_context_voice: None,
                find_in_files_cache: None,
                pending_find_in_files: None,
                normalize_undo: None,
                undo_action_label: None,
                normalize_skip_change: false,
                spellcheck_manager: spellcheck::SpellcheckManager::default(),
                spellcheck_last_announce: None,
                spellcheck_context: None,
                spellcheck_space_trigger: None,
                spellcheck_typing_in_progress: false,
                spellcheck_highlight_pending: None,
                spellcheck_last_highlighted_line: None,
                dictionary_context_menu: HMENU(0),
                dictionary_context_word: String::new(),
                dictionary_context_language: Language::default(),
                dictionary_context_pref: String::new(),
                dictionary_context_loaded: false,
                dictionary_context_expanded: false,
                dictionary_cache: load_dictionary_cache(),
                dictionary_pending_lookup: None,
                dictionary_prefetch_generation: 0,
                large_text_editors: HashSet::new(),
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            if SetTimer(hwnd, AUDIO_PLAYLIST_TIMER_ID, 700, None) == 0 {
                log_debug("Failed to set AUDIO_PLAYLIST_TIMER");
            }

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
                ShowWindow(hwnd, SW_SHOWMAXIMIZED);
                bring_window_to_foreground(hwnd);

                notify_active_editor_focus(hwnd, true);
                crate::log_if_err!(PostMessageW(hwnd, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0)));
            } else {
                editor_manager::new_document(hwnd);
                crate::log_if_err!(PostMessageW(hwnd, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0)));
            }

            editor_manager::layout_children(hwnd);
            editor_manager::apply_text_limit_to_all_edits(hwnd);
            update_main_status_bar(hwnd);
            DragAcceptFiles(hwnd, true);
            LRESULT(0)
        }
        WM_SIZE => {
            editor_manager::layout_children(hwnd);
            LRESULT(0)
        }
        WM_SETFOCUS => {
            force_active_editor_focus(hwnd);
            LRESULT(0)
        }
        WM_ACTIVATE => {
            let is_activating = (wparam.0 & 0xFFFF) != 0;
            if is_activating && should_force_editor_focus_on_foreground(hwnd) {
                force_active_editor_focus(hwnd);
                schedule_editor_focus_retry(hwnd);
            }
            LRESULT(0)
        }
        WM_NOTIFY => {
            let hdr = &*(lparam.0 as *const NMHDR);
            if hdr.code == TCN_SELCHANGE && hdr.hwndFrom == editor_manager::get_tab(hwnd) {
                // Cancel pending spellcheck highlight when switching tabs
                kill_timer_best_effort(
                    hwnd,
                    SPELLCHECK_HIGHLIGHT_TIMER_ID,
                    "KillTimer SPELLCHECK_HIGHLIGHT",
                );
                with_state(hwnd, |state| {
                    state.spellcheck_highlight_pending = None;
                    state.spellcheck_last_highlighted_line = None;
                });
                attempt_switch_to_selected_tab(hwnd);
                return LRESULT(0);
            }
            if hdr.code == EN_CHANGE {
                editor_manager::mark_dirty_from_edit(hwnd, hdr.hwndFrom);
                let active_edit = get_active_edit(hwnd).unwrap_or(HWND(0));
                if active_edit == hdr.hwndFrom {
                    let language =
                        with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                    let label = i18n::tr(language, "undo.action.text");
                    with_state(hwnd, |state| state.undo_action_label = Some(label));
                    update_main_status_bar(hwnd);
                }
                return LRESULT(0);
            }
            if hdr.code == EN_SELCHANGE {
                // Only process if editor has focus to avoid focus issues during file open etc.
                if GetFocus() == hdr.hwndFrom {
                    let is_large_editor = is_large_text_editor(hwnd, hdr.hwndFrom);
                    if !is_large_editor {
                        handle_spellcheck_selection_change(hwnd, hdr.hwndFrom);
                        prefetch_dictionary_for_selection(hwnd, hdr.hwndFrom);
                        trigger_spellcheck_highlight(hwnd, hdr.hwndFrom);
                        update_main_status_bar(hwnd);
                    }
                }
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_TIMER => {
            if wparam.0 == FOCUS_EDITOR_TIMER_ID
                || wparam.0 == FOCUS_EDITOR_TIMER_ID2
                || wparam.0 == FOCUS_EDITOR_TIMER_ID3
                || wparam.0 == FOCUS_EDITOR_TIMER_ID4
            {
                kill_timer_best_effort(hwnd, wparam.0, "KillTimer FOCUS_EDITOR");
                force_active_editor_focus(hwnd);
                return LRESULT(0);
            }
            if wparam.0 == CHAPTER_ANNOUNCE_TIMER_ID {
                update_chapter_announcement(hwnd);
                return LRESULT(0);
            }
            if wparam.0 == SPELLCHECK_HIGHLIGHT_TIMER_ID {
                kill_timer_best_effort(
                    hwnd,
                    SPELLCHECK_HIGHLIGHT_TIMER_ID,
                    "KillTimer SPELLCHECK_HIGHLIGHT",
                );
                handle_spellcheck_highlight_timer(hwnd);
                return LRESULT(0);
            }
            if wparam.0 == AUDIO_PLAYLIST_TIMER_ID {
                handle_audio_playlist_timer(hwnd);
                return LRESULT(0);
            }
            handle_pdf_loading_timer(hwnd, wparam.0);
            LRESULT(0)
        }
        WM_PODCAST_CHAPTERS_READY => {
            let ptr = lparam.0 as *mut PodcastChaptersReady;
            if ptr.is_null() {
                return LRESULT(0);
            }
            let msg = unsafe { Box::from_raw(ptr) };
            let (apply_now, chapters, language, announce_unavailable, current_pos_ms) =
                with_state(hwnd, |state| {
                    let chapters = msg.chapters.clone();
                    state
                        .podcast_chapters_cache
                        .insert(msg.key.clone(), chapters.clone());
                    let apply_now = state
                        .active_podcast_chapters_key
                        .as_deref()
                        .map(|k| k == msg.key.as_str())
                        .unwrap_or(false);
                    if apply_now {
                        state.last_announced_chapter_index = None;
                        if let Some(list) = chapters.clone() {
                            state.active_podcast_chapters = list.clone();
                            return (
                                true,
                                list,
                                state.settings.language,
                                false,
                                audiobook_position_ms_from_state(state),
                            );
                        }
                        state.active_podcast_chapters.clear();
                        return (
                            true,
                            Vec::new(),
                            state.settings.language,
                            true,
                            audiobook_position_ms_from_state(state),
                        );
                    }
                    (
                        false,
                        Vec::new(),
                        state.settings.language,
                        false,
                        audiobook_position_ms_from_state(state),
                    )
                })
                .unwrap_or((false, Vec::new(), Language::default(), false, None));
            if apply_now {
                if !chapters.is_empty() {
                    if SetTimer(hwnd, CHAPTER_ANNOUNCE_TIMER_ID, 500, None) == 0 {
                        crate::log_debug("Failed to set CHAPTER_ANNOUNCE_TIMER");
                    }
                    announce_current_chapter_on_start(hwnd, &chapters, current_pos_ms, language);
                } else {
                    kill_timer_best_effort(
                        hwnd,
                        CHAPTER_ANNOUNCE_TIMER_ID,
                        "KillTimer CHAPTER_ANNOUNCE",
                    );
                }
                if announce_unavailable {
                    let message = i18n::tr(language, "playback.chapters_unavailable");
                    nvda_speak(&message);
                }
                crate::menu::update_playback_menu(hwnd, true);
            }
            LRESULT(0)
        }
        WM_PODCAST_EPISODE_SAVE_RESULT => {
            let ptr = lparam.0 as *mut PodcastEpisodeSaveResult;
            if ptr.is_null() {
                return LRESULT(0);
            }
            let payload = Box::from_raw(ptr);
            if let Some(err) = payload.error {
                let body = i18n::tr_f(
                    payload.language,
                    "podcasts.save_error_body",
                    &[("err", &err)],
                );
                let title = i18n::tr(payload.language, "podcasts.save_error_title");
                let body_w = to_wide(&body);
                let title_w = to_wide(&title);
                message_box_modal(
                    hwnd,
                    PCWSTR(body_w.as_ptr()),
                    PCWSTR(title_w.as_ptr()),
                    MB_OK | MB_ICONERROR,
                );
                return LRESULT(0);
            }

            let show_confirmation =
                with_state(hwnd, |state| state.settings.show_media_save_confirmation)
                    .unwrap_or(true);
            if !show_confirmation {
                return LRESULT(0);
            }

            let path_text = payload.target_path.to_string_lossy().to_string();
            let saved_line = i18n::tr_f(
                payload.language,
                "podcasts.save_confirm_body",
                &[("path", &path_text)],
            );
            let open_label = i18n::tr(payload.language, "podcasts.save_confirm_open_folder");
            let body = format!("{saved_line}\n\n{open_label}?");
            let title = i18n::tr(payload.language, "podcasts.save_confirm_title");
            let body_w = to_wide(&body);
            let title_w = to_wide(&title);
            let response = message_box_modal(
                hwnd,
                PCWSTR(body_w.as_ptr()),
                PCWSTR(title_w.as_ptr()),
                MB_YESNO | MB_ICONINFORMATION,
            );
            if response == IDYES {
                let folder = payload
                    .target_path
                    .parent()
                    .map(Path::to_path_buf)
                    .unwrap_or_else(|| payload.target_path.clone());
                let folder_wide = to_wide(&folder.to_string_lossy());
                let open_res = ShellExecuteW(
                    hwnd,
                    w!("open"),
                    PCWSTR(folder_wide.as_ptr()),
                    PCWSTR::null(),
                    PCWSTR::null(),
                    SW_SHOWNORMAL,
                );
                if open_res.0 as isize <= 32 {
                    log_debug(&format!(
                        "podcast_episode_save_open_folder_failed path={} code={}",
                        folder.to_string_lossy(),
                        open_res.0
                    ));
                }
            }
            LRESULT(0)
        }
        WM_DICTIONARY_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let result = unsafe { Box::from_raw(lparam.0 as *mut DictionaryLookupResult) };
            let updated = with_state(hwnd, |state| {
                let current_gen = state.dictionary_prefetch_generation;
                if result.generation != current_gen {
                    return false;
                }
                if result.cacheable {
                    state
                        .dictionary_cache
                        .insert(result.key.clone(), result.lines.clone());
                } else {
                    state.dictionary_cache.remove(&result.key);
                }
                state.dictionary_pending_lookup = None;
                save_dictionary_cache(&state.dictionary_cache);

                if state.dictionary_context_menu.0 != 0 && !state.dictionary_context_loaded {
                    let key = dictionary_cache_key(
                        state.dictionary_context_language,
                        &state.dictionary_context_pref,
                        &state.dictionary_context_word,
                    );
                    if key == result.key {
                        let hmenu = state.dictionary_context_menu;
                        let count = GetMenuItemCount(hmenu);
                        if count > 0 {
                            for _ in 0..count {
                                crate::log_if_err!(DeleteMenu(hmenu, 0, MF_BYPOSITION));
                            }
                        }
                        for line in &result.lines {
                            let display = format!(" {}", line);
                            crate::log_if_err!(AppendMenuW(
                                hmenu,
                                MF_STRING | MF_GRAYED,
                                0,
                                PCWSTR(to_wide(&display).as_ptr()),
                            ));
                        }
                        state.dictionary_context_loaded = true;
                        return true;
                    }
                }
                false
            })
            .unwrap_or(false);
            if updated {
                let (hmenu, expanded) = with_state(hwnd, |state| {
                    (
                        state.dictionary_context_menu,
                        state.dictionary_context_expanded,
                    )
                })
                .unwrap_or((HMENU(0), false));
                if hmenu.0 != 0 && expanded {
                    crate::log_if_err!(DrawMenuBar(hwnd));
                    use windows::Win32::UI::Input::KeyboardAndMouse::{
                        KEYBD_EVENT_FLAGS, KEYEVENTF_KEYUP, VK_LEFT, VK_RIGHT, keybd_event,
                    };
                    keybd_event(VK_LEFT.0 as u8, 0, KEYBD_EVENT_FLAGS(0), 0);
                    keybd_event(VK_LEFT.0 as u8, 0, KEYEVENTF_KEYUP, 0);
                    keybd_event(VK_RIGHT.0 as u8, 0, KEYBD_EVENT_FLAGS(0), 0);
                    keybd_event(VK_RIGHT.0 as u8, 0, KEYEVENTF_KEYUP, 0);
                }
            }
            LRESULT(0)
        }
        WM_UPDATE_DIALOG => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            crate::log_debug("UI: Received WM_UPDATE_DIALOG");
            let req = unsafe { Box::from_raw(lparam.0 as *mut UpdateDialogRequest) };
            let result = unsafe {
                MessageBoxW(
                    hwnd,
                    &HSTRING::from(&req.text),
                    &HSTRING::from(&req.title),
                    req.flags,
                )
            };
            crate::log_debug(&format!("UI: Update dialog result: {:?}", result));
            if let Err(e) = req.response_tx.send(result.0) {
                crate::log_debug(&format!("UI: Failed to send response to channel: {}", e));
            } else {
                crate::log_debug("UI: Response sent to channel successfully");
            }
            LRESULT(0)
        }
        WM_UPDATE_PROGRESS_OPEN => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let req = unsafe { Box::from_raw(lparam.0 as *mut UpdateProgressOpenRequest) };
            let labels = app_windows::podcast_save_window::SaveDialogLabels {
                title: i18n::tr(req.language, "updater.title"),
                in_progress: i18n::tr(req.language, "podcast.save.in_progress"),
                cancel: i18n::tr(req.language, "podcast.save.cancel"),
            };
            let dialog = app_windows::podcast_save_window::open_with_labels(
                hwnd,
                req.language,
                labels,
                false,
            );
            with_state(hwnd, |state| {
                state.update_progress_window = dialog;
            });
            if let Err(e) = req.response_tx.send(dialog.0) {
                crate::log_debug(&format!(
                    "UI: Failed to send update progress dialog handle: {}",
                    e
                ));
            }
            LRESULT(0)
        }
        WM_UPDATE_PROGRESS_SET => {
            let pct = wparam.0.min(100);
            let dialog = with_state(hwnd, |state| state.update_progress_window).unwrap_or(HWND(0));
            if dialog.0 != 0 {
                crate::log_if_err!(PostMessageW(
                    dialog,
                    app_windows::podcast_save_window::WM_PODCAST_SAVE_PROGRESS,
                    WPARAM(pct),
                    LPARAM(0)
                ));
            }
            LRESULT(0)
        }
        WM_UPDATE_PROGRESS_CLOSE => {
            let dialog = with_state(hwnd, |state| state.update_progress_window).unwrap_or(HWND(0));
            if dialog.0 != 0 {
                crate::log_if_err!(PostMessageW(
                    dialog,
                    app_windows::podcast_save_window::WM_PODCAST_SAVE_DONE,
                    WPARAM(0),
                    LPARAM(0)
                ));
                with_state(hwnd, |state| {
                    state.update_progress_window = HWND(0);
                });
            }
            LRESULT(0)
        }
        app_windows::podcast_save_window::WM_PODCAST_SAVE_CLOSED => {
            with_state(hwnd, |state| {
                state.update_progress_window = HWND(0);
            });
            LRESULT(0)
        }
        WM_AUTO_UPDATE_CHECK => {
            if !has_secondary_window_open(hwnd) {
                updater::check_for_update(hwnd, false);
            }
            LRESULT(0)
        }
        WM_CHECK_PENDING_UPDATE => {
            if !has_secondary_window_open(hwnd) {
                updater::check_pending_update(hwnd, false);
            }
            LRESULT(0)
        }
        WM_SHOW_CHANGELOG => {
            if !has_secondary_window_open(hwnd) {
                app_windows::help_window::open_changelog(hwnd);
            }
            LRESULT(0)
        }
        WM_PDF_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = unsafe { Box::from_raw(lparam.0 as *mut PdfLoadResult) };
            handle_pdf_loaded(hwnd, *payload);
            LRESULT(0)
        }
        WM_DOCUMENT_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload =
                unsafe { Box::from_raw(lparam.0 as *mut editor_manager::DocumentLoadResult) };
            handle_document_loaded(hwnd, *payload);
            LRESULT(0)
        }
        WM_TTS_VOICES_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = unsafe { Box::from_raw(lparam.0 as *mut Vec<VoiceInfo>) };
            let voices: Vec<VoiceInfo> = *payload;
            with_state(hwnd, |state| {
                state.edge_voices = voices.clone();
            });
            if let Some(dialog) = with_state(hwnd, |state| state.options_dialog)
                && dialog.0 != 0
            {
                app_windows::options_window::refresh_voices(dialog);
            }
            refresh_voice_panel(hwnd);
            LRESULT(0)
        }
        WM_TTS_SAPI_VOICES_LOADED => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = unsafe { Box::from_raw(lparam.0 as *mut Vec<VoiceInfo>) };
            let voices: Vec<VoiceInfo> = *payload;
            with_state(hwnd, |state| {
                state.sapi_voices = voices.clone();
            });
            if let Some(dialog) = with_state(hwnd, |state| state.options_dialog)
                && dialog.0 != 0
            {
                app_windows::options_window::refresh_voices(dialog);
            }
            refresh_voice_panel(hwnd);
            LRESULT(0)
        }
        WM_TTS_START => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = unsafe { Box::from_raw(lparam.0 as *mut tts_engine::TtsPlaybackOptions) };
            tts_engine::start_tts_playback_with_chunks(*payload);
            LRESULT(0)
        }

        WM_TTS_PLAYBACK_DONE => {
            let session_id = wparam.0 as u64;
            with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session
                    && current.id == session_id
                {
                    state.tts_session = None;
                    state.tts_last_offset = 0;
                    prevent_sleep(false);
                }
            });
            LRESULT(0)
        }
        WM_TTS_CHUNK_START => {
            let session_id = wparam.0 as u64;
            let offset = lparam.0 as i32;
            with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session
                    && current.id == session_id
                {
                    let safe_offset = clamp_tts_chunk_offset(state.tts_last_offset, offset);
                    if safe_offset != offset {
                        log_debug(&format!(
                            "TTS: normalized non-monotonic offset session={} prev={} new={} safe={}",
                            session_id, state.tts_last_offset, offset, safe_offset
                        ));
                    }
                    state.tts_last_offset = safe_offset;
                    if state.settings.move_cursor_during_reading
                        && let Some(doc) = state.docs.get(state.current)
                    {
                        let new_pos = current.initial_caret_pos + safe_offset;
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
            });
            LRESULT(0)
        }
        WM_TTS_PLAYBACK_ERROR => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = unsafe { Box::from_raw(lparam.0 as *mut String) };
            let message: String = *payload;
            let session_id = wparam.0 as u64;
            let mut should_show = false;
            with_state(hwnd, |state| {
                if let Some(current) = &state.tts_session
                    && current.id == session_id
                {
                    state.tts_session = None;
                    state.tts_last_offset = 0;
                    prevent_sleep(false);
                    should_show = true;
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

            with_state(hwnd, |state| {
                if state.audiobook_progress.0 != 0 {
                    crate::log_if_err!(DestroyWindow(state.audiobook_progress));
                    state.audiobook_progress = HWND(0);
                    state.audiobook_cancel = None;
                }
                if let Some(doc) = state.docs.get(state.current) {
                    SetFocus(doc.hwnd_edit);
                }
            });

            let payload = unsafe { Box::from_raw(lparam.0 as *mut AudiobookResult) };
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
            if should_force_editor_focus_on_foreground(hwnd) {
                force_active_editor_focus(hwnd);
                schedule_editor_focus_retry(hwnd);
            } else if !has_secondary_window_open(hwnd) {
                focus_editor(hwnd);
            }
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == u32::from(VK_F9.0) {
                cycle_favorite_voice(hwnd, -1);
                return LRESULT(0);
            }
            if wparam.0 as u32 == u32::from(VK_F10.0) {
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
        WM_INITMENUPOPUP => {
            let hmenu = HMENU(wparam.0 as isize);
            let main_menu = GetMenu(hwnd);
            if main_menu.0 != 0 {
                let edit_menu = GetSubMenu(main_menu, 1);
                if edit_menu == hmenu {
                    let can_undo = can_undo_now(hwnd);
                    let flags = if can_undo {
                        MF_BYCOMMAND | MF_ENABLED
                    } else {
                        MF_BYCOMMAND | MF_GRAYED
                    };
                    let _enabled = EnableMenuItem(hmenu, IDM_EDIT_UNDO as u32, flags);
                    let language =
                        with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                    let undo_label = build_undo_menu_label(hwnd, language);
                    let menu_flags = if can_undo {
                        MF_BYCOMMAND | MF_STRING | MF_ENABLED
                    } else {
                        MF_BYCOMMAND | MF_STRING | MF_GRAYED
                    };
                    let _modified = ModifyMenuW(
                        hmenu,
                        IDM_EDIT_UNDO as u32,
                        menu_flags,
                        IDM_EDIT_UNDO,
                        PCWSTR(to_wide(&undo_label).as_ptr()),
                    );
                    let can_cut_copy = has_active_text_selection(hwnd);
                    let cut_copy_flags = if can_cut_copy {
                        MF_BYCOMMAND | MF_ENABLED
                    } else {
                        MF_BYCOMMAND | MF_GRAYED
                    };
                    let _enabled = EnableMenuItem(hmenu, IDM_EDIT_CUT as u32, cut_copy_flags);
                    let _enabled = EnableMenuItem(hmenu, IDM_EDIT_COPY as u32, cut_copy_flags);
                    let paste_flags = if can_paste_now(hwnd) {
                        MF_BYCOMMAND | MF_ENABLED
                    } else {
                        MF_BYCOMMAND | MF_GRAYED
                    };
                    let _enabled = EnableMenuItem(hmenu, IDM_EDIT_PASTE as u32, paste_flags);
                }
                let window_menu = GetSubMenu(main_menu, WINDOW_MENU_INDEX);
                if window_menu == hmenu {
                    refresh_window_open_documents_menu(hwnd, hmenu);
                }
            }
            let ctx = with_state(hwnd, |state| {
                if state.dictionary_context_menu != hmenu || state.dictionary_context_loaded {
                    return None;
                }
                state.dictionary_context_expanded = true;
                let key = dictionary_cache_key(
                    state.dictionary_context_language,
                    &state.dictionary_context_pref,
                    &state.dictionary_context_word,
                );
                let not_found = i18n::tr(state.dictionary_context_language, "dictionary.not_found");
                let cached = match state.dictionary_cache.get(&key) {
                    Some(lines) if lines.len() == 1 && lines[0] == not_found => {
                        state.dictionary_cache.remove(&key);
                        save_dictionary_cache(&state.dictionary_cache);
                        None
                    }
                    Some(lines) => Some(lines.clone()),
                    None => None,
                };
                let pending = state.dictionary_pending_lookup.as_ref() == Some(&key);
                Some((
                    state.dictionary_context_word.clone(),
                    state.dictionary_context_language,
                    state.dictionary_context_pref.clone(),
                    key,
                    cached,
                    pending,
                    state.dictionary_prefetch_generation,
                ))
            })
            .flatten();
            let Some((word, language, pref, key, cached, pending, generation)) = ctx else {
                return DefWindowProcW(hwnd, msg, wparam, lparam);
            };

            let count = GetMenuItemCount(hmenu);
            if count > 0 {
                for _ in 0..count {
                    crate::log_if_err!(DeleteMenu(hmenu, 0, MF_BYPOSITION));
                }
            }

            match cached {
                Some(lines) => {
                    for line in lines {
                        let display = line.replace('&', "");
                        crate::log_if_err!(AppendMenuW(
                            hmenu,
                            MF_STRING | MF_GRAYED,
                            0,
                            PCWSTR(to_wide(&display).as_ptr()),
                        ));
                    }
                    with_state(hwnd, |state| {
                        state.dictionary_context_loaded = true;
                    });
                }
                None => {
                    let loading_msg = i18n::tr(language, "dictionary.loading");
                    crate::log_if_err!(AppendMenuW(
                        hmenu,
                        MF_STRING | MF_GRAYED,
                        0,
                        PCWSTR(to_wide(&loading_msg).as_ptr()),
                    ));
                    if !pending {
                        with_state(hwnd, |state| {
                            state.dictionary_pending_lookup = Some(key.clone());
                        });
                        start_dictionary_lookup(hwnd.0, word, language, pref, key, generation);
                    }
                }
            }
            LRESULT(0)
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
            let cmd_id = wparam.0 & 0xffff;
            let notification = (wparam.0 >> 16) as u16;
            if u32::from(notification) == EN_CHANGE {
                if is_voice_panel_tuning_edit(hwnd, HWND(lparam.0)) {
                    return LRESULT(0);
                }
                editor_manager::handle_normalize_edit_change(hwnd, HWND(lparam.0));
                mark_dirty_from_edit(hwnd, HWND(lparam.0));
                let active_edit = get_active_edit(hwnd).unwrap_or(HWND(0));
                if active_edit == HWND(lparam.0) {
                    let language =
                        with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                    let label = i18n::tr(language, "undo.action.text");
                    with_state(hwnd, |state| state.undo_action_label = Some(label));
                    update_main_status_bar(hwnd);
                }
                return LRESULT(0);
            }
            if cmd_id == VOICE_PANEL_ID_ENGINE && u32::from(notification) == CBN_SELCHANGE {
                handle_voice_panel_engine_change(hwnd);
                return LRESULT(0);
            }
            if cmd_id == VOICE_PANEL_ID_LANGUAGE && u32::from(notification) == CBN_SELCHANGE {
                refresh_voice_panel_voice_list(hwnd);
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
            if cmd_id == VOICE_PANEL_ID_INSERT_TAG {
                insert_voice_tag_from_voice_panel(hwnd);
                return LRESULT(0);
            }
            if (cmd_id == VOICE_PANEL_ID_SPEED
                || cmd_id == VOICE_PANEL_ID_PITCH
                || cmd_id == VOICE_PANEL_ID_VOLUME)
                && u32::from(notification) == CBN_SELCHANGE
            {
                handle_voice_panel_tuning_combo_change(hwnd);
                return LRESULT(0);
            }
            if (cmd_id == VOICE_PANEL_ID_SPEED_EDIT
                || cmd_id == VOICE_PANEL_ID_PITCH_EDIT
                || cmd_id == VOICE_PANEL_ID_VOLUME_EDIT)
                && u32::from(notification) == EN_KILLFOCUS
            {
                handle_voice_panel_tuning_edit_change(hwnd);
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
            if (IDM_SPELLCHECK_SUGGESTION_BASE
                ..IDM_SPELLCHECK_SUGGESTION_BASE + IDM_SPELLCHECK_SUGGESTION_MAX)
                .contains(&cmd_id)
            {
                let index = cmd_id - IDM_SPELLCHECK_SUGGESTION_BASE;
                handle_spellcheck_suggestion(hwnd, index);
                return LRESULT(0);
            }
            if cmd_id == IDM_SPELLCHECK_ADD_TO_DICTIONARY {
                handle_spellcheck_add_to_dictionary(hwnd);
                return LRESULT(0);
            }
            if cmd_id == IDM_SPELLCHECK_IGNORE_ONCE {
                handle_spellcheck_ignore_once(hwnd);
                return LRESULT(0);
            }

            if (IDM_FILE_RECENT_BASE..IDM_FILE_RECENT_BASE + MAX_RECENT).contains(&cmd_id) {
                let index = cmd_id - IDM_FILE_RECENT_BASE;
                if let Some(path) =
                    with_state(hwnd, |state| state.recent_files.get(index).cloned()).flatten()
                {
                    editor_manager::open_document(hwnd, &path);
                }
                return LRESULT(0);
            }
            if cmd_id == IDM_FILE_RECENT_CLEAR {
                clear_recent_files(hwnd);
                return LRESULT(0);
            }

            match cmd_id {
                IDM_FILE_NEW => {
                    log_debug("Menu: New document");
                    editor_manager::new_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_OPEN => {
                    if unsafe { with_state(hwnd, |_| {}).is_none() } {
                        log_debug("Menu: Open document ignored (not initialized)");
                        return LRESULT(0);
                    }
                    log_debug("Menu: Open document");
                    // Cancel spellcheck highlight to avoid focus issues
                    kill_timer_best_effort(
                        hwnd,
                        SPELLCHECK_HIGHLIGHT_TIMER_ID,
                        "KillTimer SPELLCHECK_HIGHLIGHT",
                    );
                    with_state(hwnd, |state| {
                        state.spellcheck_highlight_pending = None;
                        state.spellcheck_last_highlighted_line = None;
                    });
                    if let Some(selected) = open_file_dialog_with_encoding(hwnd) {
                        if selected.iter().all(|(path, _)| is_audio_path(path)) {
                            let audio_paths = selected
                                .into_iter()
                                .map(|(path, _)| path)
                                .collect::<Vec<_>>();
                            queue_audio_files_and_play(hwnd, audio_paths);
                        } else {
                            for (path, encoding) in selected {
                                open_document_with_encoding(hwnd, &path, encoding);
                            }
                        }
                        if with_state(hwnd, |state| state.prompt_window.0 != 0).unwrap_or(false) {
                            focus_editor(hwnd);
                        }
                    }
                    LRESULT(0)
                }
                IDM_FILE_SAVE => {
                    log_debug("Menu: Save document");
                    editor_manager::save_current_document(hwnd);
                    editor_manager::refresh_current_editor_visual(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_SAVE_AS => {
                    log_debug("Menu: Save document as");
                    editor_manager::save_current_document_as(hwnd);
                    editor_manager::refresh_current_editor_visual(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_SAVE_ALL => {
                    log_debug("Menu: Save all documents");
                    editor_manager::save_all_documents(hwnd);
                    editor_manager::refresh_current_editor_visual(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_CLOSE => {
                    log_debug("Menu: Close document");
                    editor_manager::close_current_document(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_CLOSE_OTHERS => {
                    log_debug("Menu: Close other files");
                    if editor_manager::close_other_documents(hwnd) {
                        close_other_windows(hwnd);
                    }
                    LRESULT(0)
                }
                IDM_FILE_EXIT => {
                    log_debug("Menu: Exit");
                    editor_manager::try_close_app(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_READ_START => {
                    log_debug("Menu: Start reading");
                    tts_engine::start_tts_from_caret(hwnd);
                    LRESULT(0)
                }
                IDM_FILE_EXECUTE => {
                    log_debug("Menu: Execute file");
                    execute_current_file(hwnd);
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
                IDM_EDIT_AUDIOBOOK_SELECTION => {
                    log_debug("Menu: Record audiobook from selection");
                    tts_engine::start_audiobook_from_selection(hwnd);
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
                IDM_FILE_CONVERT_AUDIO => {
                    log_debug("Menu: Convert audio");
                    app_windows::convert_audio_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_UNDO => {
                    log_debug("Menu: Undo");
                    if !editor_manager::try_normalize_undo(hwnd) {
                        editor_manager::undo_active_edit_skip_navigation(hwnd);
                    }
                    with_state(hwnd, |state| state.undo_action_label = None);
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
                    announce_menu_action_screen_reader(hwnd, "edit.select_all");
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
                IDM_EDIT_FIND_PREVIOUS => {
                    log_debug("Menu: Find previous");
                    search::find_previous_from_state(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_REPLACE => {
                    log_debug("Menu: Replace");
                    search::open_replace_dialog(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_GO_TO_LINE => {
                    log_debug("Menu: Go to Line");
                    if let Some(hwnd_edit) = get_active_edit(hwnd) {
                        let language =
                            with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                        let title = i18n::tr(language, "goto_line.prompt_title");
                        let body = i18n::tr(language, "goto_line.prompt_body");
                        if let Some(res) = app_windows::prompt_window::prompt_user(
                            hwnd, &title, &body, "", language,
                        ) && let Ok(line) = res.trim().parse::<usize>()
                            && line > 0
                        {
                            let line_idx = line - 1;
                            let char_idx =
                                SendMessageW(hwnd_edit, EM_LINEINDEX, WPARAM(line_idx), LPARAM(0))
                                    .0;
                            if char_idx != -1 {
                                SendMessageW(
                                    hwnd_edit,
                                    EM_SETSEL,
                                    WPARAM(char_idx as usize),
                                    LPARAM(char_idx as isize),
                                );
                                SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
                            }
                        }
                    }
                    LRESULT(0)
                }
                IDM_EDIT_PREV_SPELLING_ERROR => {
                    log_debug("Menu: Previous spelling error");
                    go_to_spelling_error(hwnd, false);
                    LRESULT(0)
                }
                IDM_EDIT_NEXT_SPELLING_ERROR => {
                    log_debug("Menu: Next spelling error");
                    go_to_spelling_error(hwnd, true);
                    LRESULT(0)
                }
                IDM_EDIT_STRIP_MARKDOWN => {
                    log_debug("Menu: Strip Markdown");
                    if editor_manager::strip_markdown_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.strip_markdown");
                        if let Some(hwnd_edit) = get_active_edit(hwnd) {
                            SetFocus(hwnd_edit);
                        }
                    }
                    LRESULT(0)
                }
                IDM_EDIT_AUTO_FORMAT_TTS => {
                    log_debug("Menu: Auto format TTS");
                    if editor_manager::auto_format_tts_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.auto_format_tts");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_NORMALIZE_WHITESPACE => {
                    log_debug("Menu: Normalize whitespace");
                    if editor_manager::normalize_whitespace_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.normalize_whitespace");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_HARD_LINE_BREAK => {
                    log_debug("Menu: Hard line break");
                    if editor_manager::hard_line_break_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.hard_line_break");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_ORDER_ITEMS => {
                    log_debug("Menu: Order items");
                    if editor_manager::order_items_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.order_items");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_KEEP_UNIQUE_ITEMS => {
                    log_debug("Menu: Keep unique items");
                    if editor_manager::keep_unique_items_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.keep_unique_items");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_REVERSE_ITEMS => {
                    log_debug("Menu: Reverse items");
                    if editor_manager::reverse_items_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.reverse_items");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_QUOTE_LINES => {
                    log_debug("Menu: Quote lines");
                    if editor_manager::quote_lines_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.quote_lines");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_UNQUOTE_LINES => {
                    log_debug("Menu: Unquote lines");
                    if editor_manager::unquote_lines_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.unquote_lines");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_INDENT => {
                    log_debug("Menu: Indent");
                    editor_manager::indent_active_edit(hwnd, false);
                    LRESULT(0)
                }
                IDM_EDIT_OUTDENT => {
                    log_debug("Menu: Outdent");
                    editor_manager::indent_active_edit(hwnd, true);
                    LRESULT(0)
                }
                IDM_EDIT_INSERT_ELLIPSIS => {
                    log_debug("Menu: Insert ellipsis");
                    if !editor_manager::is_current_audiobook(hwnd)
                        && let Some(hwnd_edit) = get_active_edit(hwnd)
                    {
                        let text = to_wide("…");
                        SendMessageW(
                            hwnd_edit,
                            EM_REPLACESEL,
                            WPARAM(1),
                            LPARAM(text.as_ptr() as isize),
                        );
                    }
                    LRESULT(0)
                }
                IDM_EDIT_TEXT_STATS => {
                    log_debug("Menu: Text stats");
                    editor_manager::text_stats_active_edit(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_JOIN_LINES => {
                    log_debug("Menu: Join lines");
                    if editor_manager::join_lines_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.join_lines");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_CLEAN_EOL_HYPHENS => {
                    log_debug("Menu: Clean EOL hyphens");
                    if editor_manager::clean_end_of_line_hyphens_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.clean_eol_hyphens");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_REMOVE_DUPLICATE_LINES => {
                    log_debug("Menu: Remove duplicate lines");
                    if editor_manager::remove_duplicate_lines_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.remove_duplicate_lines");
                    }
                    LRESULT(0)
                }
                IDM_EDIT_REMOVE_DUPLICATE_CONSECUTIVE_LINES => {
                    log_debug("Menu: Remove duplicate consecutive lines");
                    if editor_manager::remove_duplicate_consecutive_lines_active_edit(hwnd) {
                        confirm_menu_action(hwnd, "edit.remove_duplicate_consecutive_lines");
                    }
                    LRESULT(0)
                }
                IDM_PLAYBACK_PLAY_PAUSE => {
                    handle_player_command(hwnd, PlayerCommand::TogglePause);
                    LRESULT(0)
                }
                IDM_PLAYBACK_STOP => {
                    handle_player_command(hwnd, PlayerCommand::Stop);
                    LRESULT(0)
                }
                IDM_PLAYBACK_SEEK_FORWARD => {
                    let skip_seconds =
                        with_state(hwnd, |state| state.settings.audiobook_skip_seconds)
                            .unwrap_or(0);
                    handle_player_command(hwnd, PlayerCommand::Seek(skip_seconds as i64));
                    LRESULT(0)
                }
                IDM_PLAYBACK_SEEK_BACKWARD => {
                    let skip_seconds =
                        with_state(hwnd, |state| state.settings.audiobook_skip_seconds)
                            .unwrap_or(0);
                    handle_player_command(hwnd, PlayerCommand::Seek(-(skip_seconds as i64)));
                    LRESULT(0)
                }
                IDM_PLAYBACK_SEEK_TO_START => {
                    handle_player_command(hwnd, PlayerCommand::SeekToStart);
                    LRESULT(0)
                }
                IDM_PLAYBACK_SEEK_TO_END => {
                    handle_player_command(hwnd, PlayerCommand::SeekToEnd);
                    LRESULT(0)
                }
                IDM_PLAYBACK_TRACK_PREV => {
                    handle_player_command(hwnd, PlayerCommand::TrackPrev);
                    LRESULT(0)
                }
                IDM_PLAYBACK_TRACK_NEXT => {
                    handle_player_command(hwnd, PlayerCommand::TrackNext);
                    LRESULT(0)
                }
                IDM_PLAYBACK_CHAPTER_PREV => {
                    handle_chapter_navigation(hwnd, -1);
                    LRESULT(0)
                }
                IDM_PLAYBACK_CHAPTER_NEXT => {
                    handle_chapter_navigation(hwnd, 1);
                    LRESULT(0)
                }
                IDM_PLAYBACK_CHAPTER_LIST => {
                    handle_chapter_list(hwnd);
                    LRESULT(0)
                }
                IDM_PLAYBACK_DOWNLOAD_EPISODE => {
                    download_active_podcast_episode(hwnd);
                    LRESULT(0)
                }
                IDM_PLAYBACK_GO_TO_TIME => {
                    handle_player_command(hwnd, PlayerCommand::GoToTime);
                    LRESULT(0)
                }
                IDM_PLAYBACK_ANNOUNCE_TIME => {
                    handle_player_command(hwnd, PlayerCommand::AnnounceTime);
                    LRESULT(0)
                }
                IDM_PLAYBACK_ADD_SUBTITLES => {
                    if let Some(path) = open_subtitle_file_dialog(hwnd) {
                        let language =
                            with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                        match audio_player::set_audiobook_subtitle_override(hwnd, &path) {
                            Ok(()) => {}
                            Err(code) => {
                                let key = match code.as_str() {
                                    "invalid_subtitle" => "playback.add_subtitles_invalid",
                                    "no_media" => "playback.add_subtitles_no_media",
                                    _ => "playback.add_subtitles_state",
                                };
                                let message = i18n::tr(language, key);
                                show_error(hwnd, language, &message);
                            }
                        }
                    }
                    LRESULT(0)
                }
                IDM_PLAYBACK_REMOVE_SUBTITLES => {
                    let language =
                        with_state(hwnd, |state| state.settings.language).unwrap_or_default();
                    match audio_player::clear_audiobook_subtitle_override(hwnd) {
                        Ok(()) => {}
                        Err(code) => {
                            let key = match code.as_str() {
                                "no_media" => "playback.remove_subtitles_no_media",
                                _ => "playback.remove_subtitles_state",
                            };
                            let message = i18n::tr(language, key);
                            show_error(hwnd, language, &message);
                        }
                    }
                    LRESULT(0)
                }
                IDM_PLAYBACK_VOLUME_UP => {
                    handle_player_command(hwnd, PlayerCommand::Volume(0.1));
                    LRESULT(0)
                }
                IDM_PLAYBACK_VOLUME_DOWN => {
                    handle_player_command(hwnd, PlayerCommand::Volume(-0.1));
                    LRESULT(0)
                }
                IDM_PLAYBACK_VOLUME_RESET => {
                    handle_player_command(hwnd, PlayerCommand::VolumeReset);
                    LRESULT(0)
                }
                IDM_PLAYBACK_SPEED_UP => {
                    handle_player_command(hwnd, PlayerCommand::Speed(0.1));
                    LRESULT(0)
                }
                IDM_PLAYBACK_SPEED_DOWN => {
                    handle_player_command(hwnd, PlayerCommand::Speed(-0.1));
                    LRESULT(0)
                }
                IDM_PLAYBACK_SPEED_RESET => {
                    handle_player_command(hwnd, PlayerCommand::SpeedReset);
                    LRESULT(0)
                }
                IDM_PLAYBACK_PITCH_UP => {
                    handle_player_command(hwnd, PlayerCommand::Pitch(1.0));
                    LRESULT(0)
                }
                IDM_PLAYBACK_PITCH_DOWN => {
                    handle_player_command(hwnd, PlayerCommand::Pitch(-1.0));
                    LRESULT(0)
                }
                IDM_PLAYBACK_PITCH_RESET => {
                    handle_player_command(hwnd, PlayerCommand::PitchReset);
                    LRESULT(0)
                }
                IDM_PLAYBACK_MUTE_TOGGLE => {
                    handle_player_command(hwnd, PlayerCommand::MuteToggle);
                    LRESULT(0)
                }
                cmd_id
                    if (IDM_PLAYBACK_AUDIO_TRACK_BASE
                        ..IDM_PLAYBACK_AUDIO_TRACK_BASE + IDM_PLAYBACK_AUDIO_TRACK_MAX)
                        .contains(&cmd_id) =>
                {
                    let track_menu_index = cmd_id - IDM_PLAYBACK_AUDIO_TRACK_BASE;
                    let track_index = with_state(hwnd, |state| {
                        state
                            .available_audio_tracks
                            .get(track_menu_index)
                            .map(|t| t.index)
                    })
                    .flatten();
                    if let Some(idx) = track_index {
                        audio_player::switch_audio_track(hwnd, idx);
                    }
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
                IDM_VIEW_READ_ONLY => {
                    let new_read_only = with_state(hwnd, |state| {
                        state.settings.editor_read_only = !state.settings.editor_read_only;
                        state.settings.editor_read_only
                    })
                    .unwrap_or(false);
                    editor_manager::apply_read_only_to_all_edits(hwnd, new_read_only);
                    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
                        save_settings(settings);
                    }
                    update_voice_panel_menu_check(hwnd);
                    LRESULT(0)
                }
                IDM_VIEW_WORD_WRAP => {
                    let new_word_wrap = with_state(hwnd, |state| {
                        state.settings.word_wrap = !state.settings.word_wrap;
                        state.settings.word_wrap
                    })
                    .unwrap_or(true);
                    editor_manager::apply_word_wrap_to_all_edits(hwnd, new_word_wrap);
                    if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
                        save_settings(settings);
                    }
                    update_voice_panel_menu_check(hwnd);
                    LRESULT(0)
                }
                cmd_id if font_face_from_menu_id(cmd_id).is_some() => {
                    let face = font_face_from_menu_id(cmd_id).unwrap_or("Segoe UI");
                    apply_ui_font(hwnd, face.to_string());
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
                IDM_GOTO_NEXT_BOOKMARK => {
                    log_debug("Menu: Go to next bookmark");
                    goto_relative_bookmark(hwnd, true);
                    LRESULT(0)
                }
                IDM_GOTO_PREV_BOOKMARK => {
                    log_debug("Menu: Go to previous bookmark");
                    goto_relative_bookmark(hwnd, false);
                    LRESULT(0)
                }
                IDM_INSERT_CLEAR_BOOKMARKS => {
                    log_debug("Menu: Clear Current Bookmarks");
                    if clear_current_bookmarks(hwnd) {
                        confirm_menu_action(hwnd, "insert.clear_bookmarks");
                    }
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
                IDM_WINDOW_OPEN_DOCUMENTS => {
                    open_documents_popup(hwnd);
                    LRESULT(0)
                }
                cmd_id if window_doc_menu_index_from_command(cmd_id).is_some() => {
                    if let Some(index) = window_doc_menu_index_from_command(cmd_id) {
                        select_tab(hwnd, index);
                    }
                    LRESULT(0)
                }
                IDM_WINDOW_CLOSE_ALL => {
                    log_debug("Menu: Close all documents");
                    editor_manager::close_all_documents(hwnd);
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
                IDM_TOOLS_DICTIONARY_LOOKUP => {
                    log_debug("Menu: Dictionary lookup");
                    open_dictionary_lookup(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_WIKIPEDIA_IMPORT => {
                    log_debug("Menu: Wikipedia import");
                    app_windows::wikipedia_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_IMPORT_YOUTUBE => {
                    log_debug("Menu: Import YouTube transcript");
                    app_windows::youtube_transcript_window::import_youtube_transcript(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_STREAM_AUDIO => {
                    log_debug("Menu: Stream audio from URL");
                    app_windows::youtube_transcript_window::play_streaming_audio_from_url(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_PROMPT => {
                    log_debug("Menu: Prompt");
                    app_windows::prompt_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_RSS => {
                    log_debug("Menu: RSS");
                    app_windows::rss_window::open(hwnd);
                    LRESULT(0)
                }
                IDM_TOOLS_PODCASTS => {
                    log_debug("Menu: Podcasts");
                    app_windows::podcasts_window::open(hwnd);
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
                IDM_HELP_DONATIONS => {
                    log_debug("Menu: Donations");
                    app_windows::help_window::open_donations(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_CHECK_UPDATES => {
                    log_debug("Menu: Check updates");
                    updater::check_for_update(hwnd, true);
                    LRESULT(0)
                }
                IDM_HELP_ABOUT => {
                    log_debug("Menu: About");
                    app_windows::about_window::show(hwnd);
                    LRESULT(0)
                }
                IDM_HELP_EXPORT_DIAGNOSTICS => {
                    log_debug("Menu: Export diagnostics");
                    export_diagnostics_dialog(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            try_close_app(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            PostQuitMessage(0);
            LRESULT(0)
        }
        WM_DROPFILES => {
            handle_drop_files(hwnd, HDROP(wparam.0 as isize));
            focus_editor(hwnd);
            LRESULT(0)
        }
        WM_COPYDATA => {
            let cds = &*(lparam.0 as *const COPYDATASTRUCT);
            if cds.dwData == COPYDATA_OPEN_FILE && !cds.lpData.is_null() {
                let len_u16 = (cds.cbData as usize) / 2;
                let slice = std::slice::from_raw_parts(cds.lpData as *const u16, len_u16);
                let len = if len_u16 > 0 && slice[len_u16 - 1] == 0 {
                    len_u16 - 1
                } else {
                    len_u16
                };
                let joined_paths = String::from_utf16_lossy(&slice[..len]);
                if !joined_paths.is_empty() {
                    for path_str in joined_paths.split('|') {
                        if !path_str.is_empty() {
                            editor_manager::open_document_from_copydata(hwnd, Path::new(path_str));
                        }
                    }
                    ShowWindow(hwnd, SW_SHOWMAXIMIZED);
                    bring_window_to_foreground(hwnd);
                    focus_editor(hwnd);
                }
                return LRESULT(1);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
            if !ptr.is_null() {
                let state_box = unsafe { Box::from_raw(ptr) };
                if state_box.hfont_custom
                    && state_box.hfont.0 != 0
                    && !DeleteObject(state_box.hfont).as_bool()
                {
                    log_debug("DeleteObject failed for AppState font on destroy");
                }
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

fn cycle_favorite_voice(hwnd: HWND, direction: i32) {
    let (favorites, current_engine, current_voice) = unsafe {
        with_state(hwnd, |state| {
            (
                state.settings.favorite_voices.clone(),
                state.settings.tts_engine,
                state.settings.tts_voice.clone(),
            )
        })
    }
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
    unsafe {
        with_state(hwnd, |state| {
            state.settings.tts_engine = next_fav.engine;
            state.settings.tts_voice = next_fav.short_name.clone();
        });
    }
    let language = unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
    unsafe {
        refresh_voice_panel(hwnd);
    }
    if let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) } {
        save_settings(settings);
    }
    unsafe {
        restart_tts_from_current_offset(hwnd);
    }
}

fn is_tts_active(hwnd: HWND) -> bool {
    unsafe { with_state(hwnd, |state| state.tts_session.is_some()).unwrap_or(false) }
}

struct VoicePanelLabels {
    label_engine: String,
    label_language: String,
    label_voice: String,
    label_speed: String,
    label_pitch: String,
    label_volume: String,
    label_favorites: String,
    label_multilingual: String,
    button_insert_tag: String,
    engine_edge: String,
    engine_sapi: String,
    engine_sapi4: String,
    voices_empty: String,
    favorites_empty: String,
    add_favorite: String,
    remove_favorite: String,
}

fn voice_panel_labels(language: Language) -> VoicePanelLabels {
    VoicePanelLabels {
        label_engine: i18n::tr(language, "voice_panel.label_engine"),
        label_language: i18n::tr(language, "voice_panel.label_language"),
        label_voice: i18n::tr(language, "voice_panel.label_voice"),
        label_speed: i18n::tr(language, "tts_tuning.label_speed"),
        label_pitch: i18n::tr(language, "tts_tuning.label_pitch"),
        label_volume: i18n::tr(language, "tts_tuning.label_volume"),
        label_favorites: i18n::tr(language, "voice_panel.label_favorites"),
        label_multilingual: i18n::tr(language, "voice_panel.label_multilingual"),
        button_insert_tag: i18n::tr(language, "voice_panel.insert_tag"),
        engine_edge: i18n::tr(language, "voice_panel.engine_edge"),
        engine_sapi: i18n::tr(language, "voice_panel.engine_sapi"),
        engine_sapi4: i18n::tr(language, "voice_panel.engine_sapi4"),
        voices_empty: i18n::tr(language, "voice_panel.voices_empty"),
        favorites_empty: i18n::tr(language, "voice_panel.favorites_empty"),
        add_favorite: i18n::tr(language, "voice_panel.add_favorite"),
        remove_favorite: i18n::tr(language, "voice_panel.remove_favorite"),
    }
}

fn voice_locale_language_code(locale: &str) -> Option<String> {
    let base = locale.split(['-', '_']).next()?.trim();
    if base.is_empty() {
        return None;
    }
    Some(base.to_ascii_lowercase())
}

fn localized_voice_language_name(language: Language, code: &str) -> String {
    let key = format!("voice.lang.{}", code);
    let localized = i18n::tr(language, &key);
    if localized != key {
        return localized;
    }
    let from_i18n = |key: &str| i18n::tr(language, key);
    match code {
        "it" => from_i18n("options.lang.it"),
        "en" => from_i18n("options.lang.en"),
        "es" => from_i18n("options.lang.es"),
        "pt" => from_i18n("options.lang.pt"),
        "sv" => from_i18n("options.lang.sv"),
        "vi" => from_i18n("options.lang.vi"),
        "cs" => from_i18n("options.lang.cs"),
        "pl" => from_i18n("options.lang.pl"),
        "fr" => from_i18n("options.lang.fr"),
        "sr" => {
            let value = from_i18n("options.lang.sr");
            if value == "options.lang.sr" {
                "Serbian".to_string()
            } else {
                value
            }
        }
        "uk" => {
            let value = from_i18n("options.lang.uk");
            if value == "options.lang.uk" {
                "Ukrainian".to_string()
            } else {
                value
            }
        }
        "lt" => {
            let value = from_i18n("options.lang.lt");
            if value == "options.lang.lt" {
                "Lithuanian".to_string()
            } else {
                value
            }
        }
        "zh" => {
            let value = from_i18n("options.lang.zh");
            if value == "options.lang.zh" {
                "Chinese".to_string()
            } else {
                value
            }
        }
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

const TTS_RATE_MIN: i32 = -100;
const TTS_RATE_MAX: i32 = 100;
const TTS_PITCH_MIN: i32 = -12;
const TTS_PITCH_MAX: i32 = 12;
const TTS_VOLUME_MIN: i32 = 25;
const TTS_VOLUME_MAX: i32 = 200;
const TTS_UI_OFFSET: i32 = 100;

fn clamp_tts_chunk_offset(previous: i32, incoming: i32) -> i32 {
    previous.max(incoming.max(0))
}

fn init_tts_panel_combo(hwnd: HWND, items: &[(String, i32)]) {
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

fn combo_value(hwnd: HWND) -> i32 {
    unsafe {
        let sel = SendMessageW(hwnd, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if sel < 0 {
            return 0;
        }
        SendMessageW(hwnd, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as i32
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
        SendMessageW(hwnd, CB_SETCURSEL, WPARAM(best_idx as usize), LPARAM(0));
    }
}

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

fn font_face_from_menu_id(cmd_id: usize) -> Option<&'static str> {
    match cmd_id {
        IDM_VIEW_FONT_ARIAL => Some("Arial"),
        IDM_VIEW_FONT_CALIBRI => Some("Calibri"),
        IDM_VIEW_FONT_CONSOLAS => Some("Consolas"),
        IDM_VIEW_FONT_SEGOE_UI => Some("Segoe UI"),
        IDM_VIEW_FONT_TAHOMA => Some("Tahoma"),
        IDM_VIEW_FONT_VERDANA => Some("Verdana"),
        IDM_VIEW_FONT_TIMES_NEW_ROMAN => Some("Times New Roman"),
        IDM_VIEW_FONT_GEORGIA => Some("Georgia"),
        _ => None,
    }
}

fn create_ui_font(
    face_name: &str,
    base_font: Option<HFONT>,
    fallback_text_size: i32,
) -> Option<HFONT> {
    if face_name.trim().is_empty() {
        return None;
    }
    let mut logfont = LOGFONTW::default();
    if let Some(font) = base_font
        && font.0 != 0
    {
        let copied = unsafe {
            GetObjectW(
                font,
                size_of::<LOGFONTW>() as i32,
                Some((&mut logfont as *mut LOGFONTW).cast()),
            )
        };
        if copied == 0 {
            log_debug("GetObjectW failed for base font; using fallback LOGFONT");
        }
    }
    if logfont.lfHeight == 0 {
        let points = fallback_text_size.max(1);
        // LOGFONT height is in pixels; negative means character height.
        logfont.lfHeight = -((points * 96 + 36) / 72);
    }
    logfont.lfFaceName.fill(0);
    let face_wide = to_wide(face_name);
    let mut i = 0usize;
    while i + 1 < face_wide.len() && i < logfont.lfFaceName.len() {
        logfont.lfFaceName[i] = face_wide[i];
        i += 1;
    }
    let hfont = unsafe { windows::Win32::Graphics::Gdi::CreateFontIndirectW(&logfont) };
    if hfont.0 == 0 { None } else { Some(hfont) }
}

fn apply_ui_font(hwnd: HWND, face_name: String) {
    let (base_font, text_size) =
        unsafe { with_state(hwnd, |state| (state.hfont, state.settings.text_size)) }
            .unwrap_or((HFONT(0), 12));
    let custom_font = create_ui_font(
        &face_name,
        if base_font.0 != 0 {
            Some(base_font)
        } else {
            None
        },
        text_size,
    );
    let is_custom = custom_font.is_some() && !face_name.trim().is_empty();
    let new_font_resolved =
        custom_font.unwrap_or_else(|| unsafe { HFONT(GetStockObject(DEFAULT_GUI_FONT).0) });
    let Some((new_font, old_font, old_custom)) = unsafe {
        with_state(hwnd, |state| {
            let old_font = state.hfont;
            let old_custom = state.hfont_custom;
            state.hfont = new_font_resolved;
            state.hfont_custom = is_custom;
            state.settings.editor_font_face = face_name.clone();
            Some((new_font_resolved, old_font, old_custom))
        })
    }
    .flatten() else {
        return;
    };

    if old_custom
        && old_font.0 != 0
        && old_font != new_font
        && !unsafe { DeleteObject(old_font) }.as_bool()
    {
        log_debug("DeleteObject failed for previous UI font");
    }

    let controls = unsafe {
        with_state(hwnd, |state| {
            vec![
                state.hwnd_tab,
                state.voice_label_engine,
                state.voice_combo_engine,
                state.voice_label_language,
                state.voice_combo_language,
                state.voice_label_voice,
                state.voice_combo_voice,
                state.voice_button_insert_tag,
                state.voice_label_speed,
                state.voice_combo_speed,
                state.voice_edit_speed,
                state.voice_label_pitch,
                state.voice_combo_pitch,
                state.voice_edit_pitch,
                state.voice_label_volume,
                state.voice_combo_volume,
                state.voice_edit_volume,
                state.voice_checkbox_multilingual,
                state.voice_label_favorites,
                state.voice_combo_favorites,
            ]
        })
    }
    .unwrap_or_default();
    for control in controls {
        if control.0 != 0 {
            unsafe {
                SendMessageW(control, WM_SETFONT, WPARAM(new_font.0 as usize), LPARAM(1));
            }
        }
    }
    editor_manager::apply_font_to_all_edits(hwnd, new_font);
    if let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) } {
        save_settings(settings);
    }
}

fn update_text_preferences(hwnd: HWND, text_color: Option<u32>, text_size: Option<i32>) {
    let mut changed = false;
    let mut next_color = None;
    let mut next_size = None;
    unsafe {
        with_state(hwnd, |state| {
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
    }

    let (color, size) = match (next_color, next_size) {
        (Some(c), Some(s)) => (c, s),
        _ => return,
    };
    if changed && let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) } {
        save_settings(settings);
    }
    editor_manager::apply_text_appearance_to_all_edits(hwnd, color, size);
    update_voice_panel_menu_check(hwnd);
}

fn update_voice_panel_menu_check(hwnd: HWND) {
    let (visible, favorites_visible, text_color, text_size, read_only, word_wrap) = unsafe {
        with_state(hwnd, |state| {
            (
                state.voice_panel_visible,
                state.voice_favorites_visible,
                state.settings.text_color,
                state.settings.text_size,
                state.settings.editor_read_only,
                state.settings.word_wrap,
            )
        })
    }
    .unwrap_or((false, false, 0x000000, 12, false, true));
    let hmenu = unsafe { GetMenu(hwnd) };
    if hmenu.0 == 0 {
        return;
    }
    let flags = if visible { MF_CHECKED } else { MF_UNCHECKED };
    if unsafe { CheckMenuItem(hmenu, IDM_VIEW_SHOW_VOICES as u32, (MF_BYCOMMAND | flags).0) }
        == 0xFFFFFFFF
    {
        crate::log_debug("CheckMenuItem failed for IDM_VIEW_SHOW_VOICES");
    }
    let fav_flags = if favorites_visible {
        MF_CHECKED
    } else {
        MF_UNCHECKED
    };
    unsafe {
        CheckMenuItem(
            hmenu,
            IDM_VIEW_SHOW_FAVORITES as u32,
            (MF_BYCOMMAND | fav_flags).0,
        );
    }
    let read_only_flags = if read_only { MF_CHECKED } else { MF_UNCHECKED };
    unsafe {
        CheckMenuItem(
            hmenu,
            IDM_VIEW_READ_ONLY as u32,
            (MF_BYCOMMAND | read_only_flags).0,
        );
    }
    let wrap_flags = if word_wrap { MF_CHECKED } else { MF_UNCHECKED };
    unsafe {
        CheckMenuItem(
            hmenu,
            IDM_VIEW_WORD_WRAP as u32,
            (MF_BYCOMMAND | wrap_flags).0,
        );
    }

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
        if unsafe { CheckMenuItem(hmenu, item as u32, (MF_BYCOMMAND | item_flags).0) } == 0xFFFFFFFF
        {
            crate::log_debug("CheckMenuItem failed for view item");
        }
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
        if unsafe { CheckMenuItem(hmenu, item as u32, (MF_BYCOMMAND | item_flags).0) } == 0xFFFFFFFF
        {
            crate::log_debug("CheckMenuItem failed for color item");
        }
    }
}

fn toggle_voice_panel(hwnd: HWND) {
    let visible = unsafe { with_state(hwnd, |state| state.voice_panel_visible) }.unwrap_or(false);
    set_voice_panel_visible(hwnd, !visible);
}

fn set_voice_panel_visible(hwnd: HWND, visible: bool) {
    set_voice_panel_visible_internal(hwnd, visible, true);
}

fn set_voice_panel_visible_internal(hwnd: HWND, visible: bool, persist: bool) {
    let (
        label_engine,
        combo_engine,
        label_language,
        combo_language,
        label_voice,
        combo_voice,
        button_insert_tag,
        label_speed,
        combo_speed,
        edit_speed,
        label_pitch,
        combo_pitch,
        edit_pitch,
        label_volume,
        combo_volume,
        edit_volume,
        checkbox_multilingual,
        changed,
    ) = match unsafe {
        with_state(hwnd, |state| {
            let changed = state.settings.show_voice_panel != visible;
            state.voice_panel_visible = visible;
            state.settings.show_voice_panel = visible;
            (
                state.voice_label_engine,
                state.voice_combo_engine,
                state.voice_label_language,
                state.voice_combo_language,
                state.voice_label_voice,
                state.voice_combo_voice,
                state.voice_button_insert_tag,
                state.voice_label_speed,
                state.voice_combo_speed,
                state.voice_edit_speed,
                state.voice_label_pitch,
                state.voice_combo_pitch,
                state.voice_edit_pitch,
                state.voice_label_volume,
                state.voice_combo_volume,
                state.voice_edit_volume,
                state.voice_checkbox_multilingual,
                changed,
            )
        })
    } {
        Some(values) => values,
        None => return,
    };

    let show = if visible { SW_SHOW } else { SW_HIDE };
    for control in [
        label_engine,
        combo_engine,
        label_language,
        combo_language,
        label_voice,
        combo_voice,
        button_insert_tag,
        label_speed,
        combo_speed,
        edit_speed,
        label_pitch,
        combo_pitch,
        edit_pitch,
        label_volume,
        combo_volume,
        edit_volume,
        checkbox_multilingual,
    ] {
        if control.0 != 0 {
            unsafe {
                ShowWindow(control, show);
            }
        }
    }
    update_voice_panel_menu_check(hwnd);
    if visible {
        let language =
            unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
        app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
        unsafe {
            refresh_voice_panel(hwnd);
        }
    }
    if persist
        && changed
        && let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) }
    {
        save_settings(settings);
    }
    clear_voice_labels_if_hidden(hwnd);
    editor_manager::layout_children(hwnd);
}

fn toggle_favorites_panel(hwnd: HWND) {
    let visible =
        unsafe { with_state(hwnd, |state| state.voice_favorites_visible) }.unwrap_or(false);
    set_favorites_panel_visible(hwnd, !visible);
}

fn set_favorites_panel_visible(hwnd: HWND, visible: bool) {
    set_favorites_panel_visible_internal(hwnd, visible, true);
}

fn set_favorites_panel_visible_internal(hwnd: HWND, visible: bool, persist: bool) {
    let (label_favorites, combo_favorites, changed) = match unsafe {
        with_state(hwnd, |state| {
            let changed = state.settings.show_favorite_panel != visible;
            state.voice_favorites_visible = visible;
            state.settings.show_favorite_panel = visible;
            (
                state.voice_label_favorites,
                state.voice_combo_favorites,
                changed,
            )
        })
    } {
        Some(values) => values,
        None => return,
    };
    let show = if visible { SW_SHOW } else { SW_HIDE };
    for control in [label_favorites, combo_favorites] {
        if control.0 != 0 {
            unsafe {
                ShowWindow(control, show);
            }
        }
    }
    update_voice_panel_menu_check(hwnd);
    if visible {
        let language =
            unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
        app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
        unsafe {
            refresh_voice_panel(hwnd);
        }
    }
    if persist
        && changed
        && let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) }
    {
        save_settings(settings);
    }
    clear_voice_labels_if_hidden(hwnd);
    editor_manager::layout_children(hwnd);
}

pub(crate) unsafe fn refresh_voice_panel(hwnd: HWND) {
    let (
        voice_visible,
        label_engine,
        combo_engine,
        label_language,
        combo_language,
        label_voice,
        combo_voice,
        button_insert_tag,
        label_speed,
        combo_speed,
        edit_speed,
        label_pitch,
        combo_pitch,
        edit_pitch,
        label_volume,
        combo_volume,
        edit_volume,
        checkbox_multilingual,
        favorites_visible,
        label_favorites,
        combo_favorites,
    ) = match with_state(hwnd, |state| {
        (
            state.voice_panel_visible,
            state.voice_label_engine,
            state.voice_combo_engine,
            state.voice_label_language,
            state.voice_combo_language,
            state.voice_label_voice,
            state.voice_combo_voice,
            state.voice_button_insert_tag,
            state.voice_label_speed,
            state.voice_combo_speed,
            state.voice_edit_speed,
            state.voice_label_pitch,
            state.voice_combo_pitch,
            state.voice_edit_pitch,
            state.voice_label_volume,
            state.voice_combo_volume,
            state.voice_edit_volume,
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
        let label_language_wide = to_wide(&labels.label_language);
        let label_voice_wide = to_wide(&labels.label_voice);
        let label_speed_wide = to_wide(&labels.label_speed);
        let label_pitch_wide = to_wide(&labels.label_pitch);
        let label_volume_wide = to_wide(&labels.label_volume);
        crate::log_if_err!(SetWindowTextW(
            label_engine,
            PCWSTR(label_engine_wide.as_ptr())
        ));
        crate::log_if_err!(SetWindowTextW(
            label_language,
            PCWSTR(label_language_wide.as_ptr())
        ));
        crate::log_if_err!(SetWindowTextW(
            label_voice,
            PCWSTR(label_voice_wide.as_ptr())
        ));
        let button_insert_wide = to_wide(&labels.button_insert_tag);
        crate::log_if_err!(SetWindowTextW(
            button_insert_tag,
            PCWSTR(button_insert_wide.as_ptr())
        ));
        crate::log_if_err!(SetWindowTextW(
            label_speed,
            PCWSTR(label_speed_wide.as_ptr())
        ));
        crate::log_if_err!(SetWindowTextW(
            label_pitch,
            PCWSTR(label_pitch_wide.as_ptr())
        ));
        crate::log_if_err!(SetWindowTextW(
            label_volume,
            PCWSTR(label_volume_wide.as_ptr())
        ));
        let label_multi_wide = to_wide(&labels.label_multilingual);
        crate::log_if_err!(SetWindowTextW(
            checkbox_multilingual,
            PCWSTR(label_multi_wide.as_ptr())
        ));
    }
    if favorites_visible && label_favorites.0 != 0 {
        let label_fav_wide = to_wide(&labels.label_favorites);
        crate::log_if_err!(SetWindowTextW(
            label_favorites,
            PCWSTR(label_fav_wide.as_ptr())
        ));
    }

    if voice_visible && combo_engine.0 != 0 && combo_voice.0 != 0 {
        SendMessageW(combo_engine, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        SendMessageW(
            combo_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_edge).as_ptr() as isize),
        );
        SendMessageW(
            combo_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi).as_ptr() as isize),
        );
        SendMessageW(
            combo_engine,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&labels.engine_sapi4).as_ptr() as isize),
        );
        let engine_index = match settings.tts_engine {
            TtsEngine::Edge => 0,
            TtsEngine::Sapi5 => 1,
            TtsEngine::Sapi4 => 2,
        };
        SendMessageW(combo_engine, CB_SETCURSEL, WPARAM(engine_index), LPARAM(0));
        let is_edge = matches!(settings.tts_engine, TtsEngine::Edge);
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
        EnableWindow(checkbox_multilingual, is_edge);
        let multi_show = if is_edge { SW_SHOW } else { SW_HIDE };
        ShowWindow(checkbox_multilingual, multi_show);
        let show_language_combo = is_edge && !settings.tts_only_multilingual;
        ShowWindow(
            label_language,
            if show_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        ShowWindow(
            combo_language,
            if show_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        EnableWindow(combo_language, show_language_combo);
    }

    if voice_visible {
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
        init_tts_panel_combo(combo_speed, &speed_items);
        init_tts_panel_combo(combo_pitch, &pitch_items);
        init_tts_panel_combo(combo_volume, &volume_items);
        select_combo_nearest_value(combo_speed, settings.tts_rate);
        select_combo_nearest_value(combo_pitch, settings.tts_pitch);
        select_combo_nearest_value(combo_volume, settings.tts_volume);
        crate::log_if_err!(SetWindowTextW(
            edit_speed,
            PCWSTR(to_wide(&tts_ui_value_from_internal(settings.tts_rate).to_string()).as_ptr()),
        ));
        crate::log_if_err!(SetWindowTextW(
            edit_pitch,
            PCWSTR(to_wide(&tts_ui_value_from_internal(settings.tts_pitch).to_string()).as_ptr()),
        ));
        crate::log_if_err!(SetWindowTextW(
            edit_volume,
            PCWSTR(to_wide(&settings.tts_volume.to_string()).as_ptr()),
        ));
        let manual = settings.tts_manual_tuning;
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
        let voices: Vec<crate::settings::VoiceInfo> =
            with_state(hwnd, |state| match settings.tts_engine {
                TtsEngine::Edge => state.edge_voices.clone(),
                TtsEngine::Sapi5 => state.sapi_voices.clone(),
                TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
            })
            .unwrap_or_default();
        let mut language_filter: Option<String> = None;
        let show_language_combo =
            matches!(settings.tts_engine, TtsEngine::Edge) && !settings.tts_only_multilingual;
        if show_language_combo {
            let previous_selection = with_state(hwnd, |state| {
                let sel = SendMessageW(combo_language, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                if sel >= 0 {
                    state.voice_language_codes.get(sel as usize).cloned()
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
                SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
                let mut selected_index: Option<usize> = None;
                for (idx, code) in codes.iter().enumerate() {
                    let label = localized_voice_language_name(settings.language, code);
                    let added = SendMessageW(
                        combo_language,
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
                    combo_language,
                    CB_SETCURSEL,
                    WPARAM(selected_index.unwrap_or(0)),
                    LPARAM(0),
                );
                language_filter = Some(selected_code);
            }
            if with_state(hwnd, |state| {
                state.voice_language_codes = std::mem::take(&mut codes);
            })
            .is_none()
            {
                log_debug("Failed to update voice language codes");
            }
        } else {
            SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
            if with_state(hwnd, |state| state.voice_language_codes.clear()).is_none() {
                log_debug("Failed to clear voice language codes");
            }
        }
        populate_voice_panel_combo(
            combo_voice,
            &voices,
            &settings.tts_voice,
            settings.tts_only_multilingual,
            language_filter.as_deref(),
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

fn refresh_voice_panel_voice_list(hwnd: HWND) {
    unsafe {
        let (voice_visible, combo_voice, checkbox_multilingual, label_language, combo_language) =
            match with_state(hwnd, |state| {
                (
                    state.voice_panel_visible,
                    state.voice_combo_voice,
                    state.voice_checkbox_multilingual,
                    state.voice_label_language,
                    state.voice_combo_language,
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
        EnableWindow(checkbox_multilingual, is_edge);
        let multi_show = if is_edge { SW_SHOW } else { SW_HIDE };
        ShowWindow(checkbox_multilingual, multi_show);
        let show_language_combo = is_edge && !settings.tts_only_multilingual;
        ShowWindow(
            label_language,
            if show_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        ShowWindow(
            combo_language,
            if show_language_combo {
                SW_SHOW
            } else {
                SW_HIDE
            },
        );
        EnableWindow(combo_language, show_language_combo);

        let voices: Vec<crate::settings::VoiceInfo> =
            with_state(hwnd, |state| match settings.tts_engine {
                TtsEngine::Edge => state.edge_voices.clone(),
                TtsEngine::Sapi5 => state.sapi_voices.clone(),
                TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
            })
            .unwrap_or_default();
        let mut language_filter: Option<String> = None;
        if show_language_combo {
            let previous_selection = with_state(hwnd, |state| {
                let sel = SendMessageW(combo_language, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                if sel >= 0 {
                    state.voice_language_codes.get(sel as usize).cloned()
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
                SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
                let mut selected_index: Option<usize> = None;
                for (idx, code) in codes.iter().enumerate() {
                    let label = localized_voice_language_name(settings.language, code);
                    let added = SendMessageW(
                        combo_language,
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
                    combo_language,
                    CB_SETCURSEL,
                    WPARAM(selected_index.unwrap_or(0)),
                    LPARAM(0),
                );
                language_filter = Some(selected_code);
            }
            if with_state(hwnd, |state| {
                state.voice_language_codes = std::mem::take(&mut codes);
            })
            .is_none()
            {
                log_debug("Failed to update voice language codes");
            }
        } else {
            SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
            if with_state(hwnd, |state| state.voice_language_codes.clear()).is_none() {
                log_debug("Failed to clear voice language codes");
            }
        }
        populate_voice_panel_combo(
            combo_voice,
            &voices,
            &settings.tts_voice,
            settings.tts_only_multilingual,
            language_filter.as_deref(),
            &labels.voices_empty,
        );
    }
}

fn clear_voice_labels_if_hidden(hwnd: HWND) {
    unsafe {
        let (
            voice_visible,
            favorites_visible,
            label_engine,
            label_language,
            label_voice,
            label_speed,
            label_pitch,
            label_volume,
            checkbox_multilingual,
            label_favorites,
        ) = match with_state(hwnd, |state| {
            (
                state.voice_panel_visible,
                state.voice_favorites_visible,
                state.voice_label_engine,
                state.voice_label_language,
                state.voice_label_voice,
                state.voice_label_speed,
                state.voice_label_pitch,
                state.voice_label_volume,
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
        crate::log_if_err!(SetWindowTextW(label_engine, PCWSTR(empty.as_ptr())));
        crate::log_if_err!(SetWindowTextW(label_language, PCWSTR(empty.as_ptr())));
        crate::log_if_err!(SetWindowTextW(label_voice, PCWSTR(empty.as_ptr())));
        crate::log_if_err!(SetWindowTextW(label_speed, PCWSTR(empty.as_ptr())));
        crate::log_if_err!(SetWindowTextW(label_pitch, PCWSTR(empty.as_ptr())));
        crate::log_if_err!(SetWindowTextW(label_volume, PCWSTR(empty.as_ptr())));
        crate::log_if_err!(SetWindowTextW(
            checkbox_multilingual,
            PCWSTR(empty.as_ptr())
        ));
        crate::log_if_err!(SetWindowTextW(label_favorites, PCWSTR(empty.as_ptr())));
    }
}

fn populate_voice_panel_combo(
    combo_voice: HWND,
    voices: &[VoiceInfo],
    selected: &str,
    only_multilingual: bool,
    language_filter: Option<&str>,
    empty_label: &str,
) {
    unsafe {
        SendMessageW(combo_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        if voices.is_empty() {
            SendMessageW(
                combo_voice,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(empty_label).as_ptr() as isize),
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
            let idx = SendMessageW(
                combo_voice,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&label).as_ptr() as isize),
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

fn populate_favorites_combo(
    combo_favorites: HWND,
    favorites: &[FavoriteVoice],
    selected_engine: TtsEngine,
    selected_voice: &str,
    labels: &VoicePanelLabels,
) {
    unsafe {
        SendMessageW(combo_favorites, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        if favorites.is_empty() {
            SendMessageW(
                combo_favorites,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&labels.favorites_empty).as_ptr() as isize),
            );
            SendMessageW(combo_favorites, CB_SETCURSEL, WPARAM(0), LPARAM(0));
            return;
        }
        let mut selected_index: Option<usize> = None;
        for (idx, fav) in favorites.iter().enumerate() {
            let engine_label = match fav.engine {
                TtsEngine::Edge => &labels.engine_edge,
                TtsEngine::Sapi5 => &labels.engine_sapi,
                TtsEngine::Sapi4 => &labels.engine_sapi,
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
                SendMessageW(
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
            SendMessageW(combo_favorites, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
        } else {
            SendMessageW(combo_favorites, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        }
    }
}

fn handle_voice_panel_engine_change(hwnd: HWND) {
    unsafe {
        let (combo_engine, language) = match with_state(hwnd, |state| {
            (state.voice_combo_engine, state.settings.language)
        }) {
            Some(values) => values,
            None => return,
        };
        let sel = SendMessageW(combo_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        let new_engine = match sel {
            1 => TtsEngine::Sapi5,
            2 => TtsEngine::Sapi4,
            _ => TtsEngine::Edge,
        };
        let (old_engine, old_voice) = with_state(hwnd, |state| {
            (state.settings.tts_engine, state.settings.tts_voice.clone())
        })
        .unwrap_or((TtsEngine::Edge, String::new()));
        with_state(hwnd, |state| {
            state.settings.tts_engine = new_engine;
        });
        app_windows::options_window::ensure_voice_lists_loaded(hwnd, language);
        refresh_voice_panel(hwnd);
        let mut new_voice = old_voice.clone();
        if let Some(voice_name) = current_voice_selection(hwnd, new_engine) {
            with_state(hwnd, |state| {
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
}

fn handle_voice_panel_voice_change(hwnd: HWND) {
    unsafe {
        let engine = with_state(hwnd, |state| state.settings.tts_engine).unwrap_or_default();
        if let Some(voice_name) = current_voice_selection(hwnd, engine) {
            let old_voice =
                with_state(hwnd, |state| state.settings.tts_voice.clone()).unwrap_or_default();
            if voice_name != old_voice {
                with_state(hwnd, |state| {
                    state.settings.tts_voice = voice_name;
                });
                if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
                    save_settings(settings);
                }
                restart_tts_from_current_offset(hwnd);
            }
        }
    }
}

fn handle_voice_panel_multilingual_toggle(hwnd: HWND) {
    unsafe {
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
        with_state(hwnd, |state| {
            state.settings.tts_only_multilingual = checked;
        });
        if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
            save_settings(settings);
        }
        refresh_voice_panel_voice_list(hwnd);
    }
}

fn insert_voice_tag_from_voice_panel(hwnd: HWND) {
    unsafe {
        let (engine, rate, pitch, volume) = with_state(hwnd, |state| {
            (
                state.settings.tts_engine,
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
            )
        })
        .unwrap_or_default();
        let voice_name = current_voice_selection(hwnd, engine)
            .or_else(|| with_state(hwnd, |state| Some(state.settings.tts_voice.clone())).flatten())
            .unwrap_or_default();
        if voice_name.trim().is_empty() {
            return;
        }
        crate::editor_manager::insert_voice_tag_at_caret(
            hwnd,
            engine,
            &voice_name,
            rate,
            pitch,
            volume,
        );
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        let label = clean_menu_label(&i18n::tr(language, "options.label.insert_voice_tag"));
        if !label.is_empty() {
            with_state(hwnd, |state| state.undo_action_label = Some(label));
        }
    }
}

fn is_voice_panel_tuning_edit(hwnd: HWND, target: HWND) -> bool {
    if target.0 == 0 {
        return false;
    }
    unsafe {
        with_state(hwnd, |state| {
            target == state.voice_edit_speed
                || target == state.voice_edit_pitch
                || target == state.voice_edit_volume
        })
    }
    .unwrap_or(false)
}

fn handle_voice_panel_tuning_combo_change(hwnd: HWND) {
    unsafe {
        let (combo_speed, combo_pitch, combo_volume, was_active, old_rate, old_pitch, old_volume) =
            with_state(hwnd, |state| {
                (
                    state.voice_combo_speed,
                    state.voice_combo_pitch,
                    state.voice_combo_volume,
                    state.tts_session.is_some(),
                    state.settings.tts_rate,
                    state.settings.tts_pitch,
                    state.settings.tts_volume,
                )
            })
            .unwrap_or((HWND(0), HWND(0), HWND(0), false, 0, 0, 100));
        if combo_speed.0 == 0 || combo_pitch.0 == 0 || combo_volume.0 == 0 {
            return;
        }
        let rate = combo_value(combo_speed);
        let pitch = combo_value(combo_pitch);
        let volume = combo_value(combo_volume);
        let changed = with_state(hwnd, |state| {
            if state.settings.tts_rate != rate
                || state.settings.tts_pitch != pitch
                || state.settings.tts_volume != volume
            {
                state.settings.tts_rate = rate;
                state.settings.tts_pitch = pitch;
                state.settings.tts_volume = volume;
                true
            } else {
                false
            }
        })
        .unwrap_or(false);
        if changed {
            if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
                save_settings(settings);
            }
            if was_active && (old_rate != rate || old_pitch != pitch || old_volume != volume) {
                restart_tts_from_current_offset(hwnd);
            }
        }
    }
}

fn handle_voice_panel_tuning_edit_change(hwnd: HWND) {
    unsafe {
        let (edit_speed, edit_pitch, edit_volume, was_active, old_rate, old_pitch, old_volume) =
            with_state(hwnd, |state| {
                (
                    state.voice_edit_speed,
                    state.voice_edit_pitch,
                    state.voice_edit_volume,
                    state.tts_session.is_some(),
                    state.settings.tts_rate,
                    state.settings.tts_pitch,
                    state.settings.tts_volume,
                )
            })
            .unwrap_or((HWND(0), HWND(0), HWND(0), false, 0, 0, 100));
        if edit_speed.0 == 0 || edit_pitch.0 == 0 || edit_volume.0 == 0 {
            return;
        }
        let rate = read_tts_tuning_edit_value(edit_speed, old_rate, TTS_RATE_MIN, TTS_RATE_MAX);
        let pitch = read_tts_tuning_edit_value(edit_pitch, old_pitch, TTS_PITCH_MIN, TTS_PITCH_MAX);
        let volume = read_tts_edit_value(edit_volume, old_volume, TTS_VOLUME_MIN, TTS_VOLUME_MAX);
        let changed = with_state(hwnd, |state| {
            if state.settings.tts_rate != rate
                || state.settings.tts_pitch != pitch
                || state.settings.tts_volume != volume
            {
                state.settings.tts_rate = rate;
                state.settings.tts_pitch = pitch;
                state.settings.tts_volume = volume;
                true
            } else {
                false
            }
        })
        .unwrap_or(false);
        if changed {
            if let Some(settings) = with_state(hwnd, |state| state.settings.clone()) {
                save_settings(settings);
            }
            if was_active && (old_rate != rate || old_pitch != pitch || old_volume != volume) {
                restart_tts_from_current_offset(hwnd);
            }
        }
    }
}

fn handle_voice_panel_favorite_change(hwnd: HWND) {
    unsafe {
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
        with_state(hwnd, |state| {
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
}

fn current_voice_selection(hwnd: HWND, engine: TtsEngine) -> Option<String> {
    let (combo_voice, voices) = unsafe {
        with_state(hwnd, |state| {
            let list = match engine {
                TtsEngine::Edge => state.edge_voices.clone(),
                TtsEngine::Sapi5 => state.sapi_voices.clone(),
                TtsEngine::Sapi4 => crate::sapi4_engine::get_voices(),
            };
            (state.voice_combo_voice, list)
        })
    }?;
    if voices.is_empty() || combo_voice.0 == 0 {
        return None;
    }
    let sel = unsafe { SendMessageW(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if sel < 0 {
        return None;
    }
    let voice_index = unsafe {
        SendMessageW(combo_voice, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as usize
    };
    voices.get(voice_index).map(|v| v.short_name.clone())
}

fn current_favorite_selection(hwnd: HWND) -> Option<FavoriteVoice> {
    let (combo_favorites, favorites) = unsafe {
        with_state(hwnd, |state| {
            (
                state.voice_combo_favorites,
                state.settings.favorite_voices.clone(),
            )
        })
    }?;
    if combo_favorites.0 == 0 || favorites.is_empty() {
        return None;
    }
    let sel = unsafe { SendMessageW(combo_favorites, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if sel < 0 {
        return None;
    }
    let fav_idx = unsafe {
        SendMessageW(
            combo_favorites,
            CB_GETITEMDATA,
            WPARAM(sel as usize),
            LPARAM(0),
        )
        .0 as usize
    };
    favorites.get(fav_idx).cloned()
}

unsafe extern "system" fn voice_combo_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "voice_combo_subclass_proc",
        || DefWindowProcW(hwnd, msg, wparam, lparam),
        || {
            if msg == WM_CONTEXTMENU {
                let parent = GetParent(hwnd);
                if parent.0 != 0 {
                    show_voice_context_menu(parent, hwnd, lparam);
                    return LRESULT(0);
                }
            }
            if msg == WM_KEYDOWN
                && wparam.0 as u32 == u32::from(VK_F10.0)
                && GetKeyState(VK_SHIFT.0 as i32) < 0
            {
                let parent = GetParent(hwnd);
                if parent.0 != 0 {
                    show_voice_context_menu(parent, hwnd, LPARAM(-1));
                    return LRESULT(0);
                }
            }

            let parent = GetParent(hwnd);
            let prev_proc = if parent.0 != 0 {
                with_state(parent, |s| {
                    if hwnd == s.voice_combo_voice {
                        s.voice_combo_voice_proc
                    } else if hwnd == s.voice_combo_favorites {
                        s.voice_combo_favorites_proc
                    } else {
                        None
                    }
                })
                .unwrap_or(None)
            } else {
                None
            };
            if let Some(proc) = prev_proc {
                CallWindowProcW(Some(proc), hwnd, msg, wparam, lparam)
            } else {
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        },
    )
}

pub(crate) unsafe fn restart_tts_from_current_offset(hwnd: HWND) {
    let mut restart = None;
    with_state(hwnd, |state| {
        if let Some(session) = &state.tts_session
            && let Some(doc) = state.docs.get(state.current)
        {
            if matches!(doc.format, FileFormat::Audiobook) {
                return;
            }
            let pos = (session.initial_caret_pos + state.tts_last_offset).max(0);
            restart = Some((doc.hwnd_edit, pos));
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

fn adjust_tts_restart_pos(hwnd_edit: HWND, pos: i32) -> i32 {
    if pos <= 0 {
        return 0;
    }
    let text = editor_manager::get_edit_text(hwnd_edit);
    if text.is_empty() {
        return pos;
    }
    let mut items: Vec<(usize, usize, bool)> = Vec::new();
    let mut offset = 0usize;
    for ch in text.chars() {
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
    if prev_is_word
        && next_is_word
        && let Some(mut idx) = prev
    {
        while idx > 0 && items[idx - 1].2 {
            idx -= 1;
        }
        return items[idx].0 as i32;
    }
    pos
}

fn spellcheck_caret_char_index(hwnd_edit: HWND) -> Option<i32> {
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut selection as *mut _ as isize),
        );
    }
    if selection.cpMin < 0 {
        None
    } else {
        Some(selection.cpMin)
    }
}

fn spellcheck_char_index_from_lparam(hwnd_edit: HWND, lparam: LPARAM) -> Option<i32> {
    if lparam.0 == -1 {
        return spellcheck_caret_char_index(hwnd_edit);
    }
    let x = (lparam.0 & 0xffff) as i32;
    let y = ((lparam.0 >> 16) & 0xffff) as i32;
    if x == -1 && y == -1 {
        return spellcheck_caret_char_index(hwnd_edit);
    }
    let mut pt = POINT { x, y };
    if !unsafe { ScreenToClient(hwnd_edit, &mut pt).as_bool() } {
        crate::log_debug("ScreenToClient failed");
    }
    let res = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_CHARFROMPOS,
            WPARAM(0),
            LPARAM(&pt as *const _ as isize),
        )
        .0 as i32
    };
    if res < 0 { None } else { Some(res) }
}

fn spellcheck_line_info(hwnd_edit: HWND, char_index: i32) -> Option<(i32, i32, String)> {
    if char_index < 0 {
        return None;
    }
    let line_index = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_LINEFROMCHAR,
            WPARAM(char_index as usize),
            LPARAM(0),
        )
        .0 as i32
    };
    if line_index < 0 {
        return None;
    }
    let line_start = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_LINEINDEX,
            WPARAM(line_index as usize),
            LPARAM(0),
        )
        .0 as i32
    };
    if line_start < 0 {
        return None;
    }
    let line_len = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_LINELENGTH,
            WPARAM(line_start as usize),
            LPARAM(0),
        )
        .0 as i32
    };
    if line_len <= 0 {
        return Some((line_index, line_start, String::new()));
    }
    let mut buf = vec![0u16; (line_len + 1) as usize];
    let mut range = TEXTRANGEW {
        chrg: CHARRANGE {
            cpMin: line_start,
            cpMax: line_start + line_len,
        },
        lpstrText: PWSTR(buf.as_mut_ptr()),
    };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_GETTEXTRANGE,
            WPARAM(0),
            LPARAM(&mut range as *mut _ as isize),
        );
    }
    let line_len = line_len.max(0) as usize;
    let line_text = String::from_utf16_lossy(&buf[..line_len]);
    Some((line_index, line_start, line_text))
}

fn spellcheck_word_context_from_char_index(
    hwnd_edit: HWND,
    char_index: i32,
) -> Option<SpellcheckWordContext> {
    let (line_index, line_start, line_text) = spellcheck_line_info(hwnd_edit, char_index)?;
    if line_text.is_empty() {
        return None;
    }
    let offset_utf16 = (char_index - line_start).max(0) as u32;
    let caret_byte = spellcheck::utf16_offset_to_utf8_byte_offset(&line_text, offset_utf16);
    let word_range = spellcheck::word_range_at(&line_text, caret_byte)?;
    let word = line_text
        .get(word_range.0..word_range.1)
        .unwrap_or("")
        .to_string();
    if word.is_empty() {
        return None;
    }
    let line_hash = spellcheck::hash_line(&line_text);
    Some(SpellcheckWordContext {
        doc_id: hwnd_edit.0,
        line_index,
        line_start,
        line_text,
        line_hash,
        word_range,
        word,
    })
}

fn spellcheck_word_context_from_lparam(
    hwnd_edit: HWND,
    lparam: LPARAM,
) -> Option<SpellcheckWordContext> {
    let char_index = spellcheck_char_index_from_lparam(hwnd_edit, lparam)?;
    spellcheck_word_context_from_char_index(hwnd_edit, char_index)
}

fn handle_spellcheck_selection_change(hwnd: HWND, hwnd_edit: HWND) {
    let announce_allowed = unsafe {
        with_state(hwnd, |state| {
            if state.spellcheck_space_trigger == Some(hwnd_edit) {
                // Force re-highlight of current line when space/punctuation is pressed.
                state.spellcheck_last_highlighted_line = None;
                state.spellcheck_space_trigger = None;
                state.spellcheck_typing_in_progress = false;
                return true;
            }
            !state.spellcheck_typing_in_progress
        })
    }
    .unwrap_or(false);
    let Some(caret_index) = spellcheck_caret_char_index(hwnd_edit) else {
        unsafe {
            with_state(hwnd, |state| state.spellcheck_last_announce = None);
        }
        return;
    };
    let Some(word_ctx) = spellcheck_word_context_from_char_index(hwnd_edit, caret_index) else {
        unsafe {
            with_state(hwnd, |state| state.spellcheck_last_announce = None);
        }
        return;
    };

    let (announce_msg, fallback_msg) = unsafe {
        with_state(hwnd, |state| {
            let settings = &state.settings;
            let Some(resolution) = state.spellcheck_manager.resolve_language(settings) else {
                state.spellcheck_last_announce = None;
                return (None, None);
            };
            let language_ui = settings.language;
            let fallback_msg = if resolution.announce_fallback {
                Some(i18n::tr_f(
                    language_ui,
                    "spellcheck.language_fallback",
                    &[
                        ("requested", &resolution.requested),
                        ("language", &resolution.effective),
                    ],
                ))
            } else {
                None
            };

            let miss = state.spellcheck_manager.is_word_misspelled(
                word_ctx.doc_id,
                word_ctx.line_index,
                &word_ctx.line_text,
                word_ctx.word_range,
                &resolution.effective,
            );
            if let Some(miss) = miss {
                let key = SpellcheckAnnounceKey {
                    doc_id: word_ctx.doc_id,
                    line_index: word_ctx.line_index,
                    start_utf8: miss.start,
                    end_utf8: miss.end,
                    line_hash: word_ctx.line_hash,
                    language: resolution.effective.clone(),
                };
                if announce_allowed && state.spellcheck_last_announce.as_ref() != Some(&key) {
                    state.spellcheck_last_announce = Some(key);
                    let msg = i18n::tr_f(
                        language_ui,
                        "spellcheck.announce_misspelled",
                        &[("word", &word_ctx.word)],
                    );
                    return (Some(msg), fallback_msg);
                }
                return (None, fallback_msg);
            }
            state.spellcheck_last_announce = None;
            (None, fallback_msg)
        })
    }
    .unwrap_or((None, None));

    if let Some(message) = fallback_msg {
        log_debug(&format!("Spellcheck: {message}"));
        nvda_speak(&message);
    }
    if let Some(message) = announce_msg {
        nvda_speak(&message);
    }
}

/// Triggers the debounced spellcheck highlight timer
fn trigger_spellcheck_highlight(hwnd: HWND, hwnd_edit: HWND) {
    let should_start_timer = unsafe {
        with_state(hwnd, |state| {
            if !state.settings.spellcheck_enabled {
                return false;
            }
            state.spellcheck_highlight_pending = Some(hwnd_edit);
            true
        })
    }
    .unwrap_or(false);

    if should_start_timer {
        // Reset/start the debounce timer only if spellcheck is enabled
        unsafe {
            SetTimer(
                hwnd,
                SPELLCHECK_HIGHLIGHT_TIMER_ID,
                SPELLCHECK_HIGHLIGHT_DEBOUNCE_MS,
                None,
            );
        }
    }
}

/// Called when the debounce timer fires - highlights misspellings on current line
fn handle_spellcheck_highlight_timer(hwnd: HWND) {
    let Some(hwnd_edit) =
        (unsafe { with_state(hwnd, |state| state.spellcheck_highlight_pending.take()).flatten() })
    else {
        return;
    };

    // Don't do anything if editor doesn't have focus
    if unsafe { GetFocus() } != hwnd_edit {
        return;
    }

    let Some(caret_index) = spellcheck_caret_char_index(hwnd_edit) else {
        return;
    };

    let Some((line_index, line_start, line_text)) = spellcheck_line_info(hwnd_edit, caret_index)
    else {
        return;
    };

    let doc_id = hwnd_edit.0;

    // Check if we're on the same line as before - no need to re-highlight
    let should_highlight = unsafe {
        with_state(hwnd, |state| {
            let last = state.spellcheck_last_highlighted_line;
            if last == Some((doc_id, line_index)) {
                return false;
            }
            state.spellcheck_last_highlighted_line = Some((doc_id, line_index));
            true
        })
    }
    .unwrap_or(false);

    if !should_highlight {
        return;
    }

    // Get misspellings for this line
    let misspellings = unsafe {
        with_state(hwnd, |state| {
            let settings = &state.settings;
            let resolution = state.spellcheck_manager.resolve_language(settings)?;
            let misses = state.spellcheck_manager.check_line(
                doc_id,
                line_index,
                &line_text,
                &resolution.effective,
            );
            Some((misses, settings.text_color))
        })
    }
    .flatten();

    let Some((misspellings, text_color)) = misspellings else {
        return;
    };

    // First, reset the entire line to normal formatting
    reset_line_formatting(hwnd_edit, line_start, line_text.len(), text_color);

    // Then highlight each misspelled word with red background
    for miss in misspellings {
        let start_utf16 = spellcheck::utf8_byte_offset_to_utf16_units(&line_text, miss.start);
        let end_utf16 = spellcheck::utf8_byte_offset_to_utf16_units(&line_text, miss.end);
        let abs_start = line_start + start_utf16;
        let abs_end = line_start + end_utf16;
        highlight_misspelled_word(hwnd_edit, abs_start, abs_end);
    }
}

/// Resets line formatting to normal (removes red background)
fn reset_line_formatting(hwnd_edit: HWND, line_start: i32, line_len: usize, _text_color: u32) {
    unsafe {
        if line_len == 0 {
            return;
        }

        // Check if this editor still has focus - don't mess with selection if not
        if GetFocus() != hwnd_edit {
            return;
        }

        // Save modified state - formatting changes should not mark document as dirty
        let was_modified = SendMessageW(hwnd_edit, EM_GETMODIFY, WPARAM(0), LPARAM(0)).0 != 0;

        // Disable change notifications during formatting
        SendMessageW(hwnd_edit, EM_SETEVENTMASK, WPARAM(0), LPARAM(0));

        // Lock the control to prevent redraws/scrolling
        SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(0), LPARAM(0));

        // Save current selection
        let mut old_sel = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut old_sel as *mut _ as isize),
        );

        // Select the line
        let line_end = line_start + line_len as i32;
        let mut sel = CHARRANGE {
            cpMin: line_start,
            cpMax: line_end,
        };
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut sel as *mut _ as isize),
        );

        // Apply normal formatting (remove background color)
        let mut format = CHARFORMAT2W::default();
        format.Base.cbSize = std::mem::size_of::<CHARFORMAT2W>() as u32;
        format.Base.dwMask = CFM_BACKCOLOR;
        format.Base.dwEffects = CFE_AUTOBACKCOLOR; // Use default background
        SendMessageW(
            hwnd_edit,
            EM_SETCHARFORMAT,
            WPARAM(SCF_SELECTION as usize),
            LPARAM(&mut format as *mut _ as isize),
        );

        // Restore selection
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut old_sel as *mut _ as isize),
        );

        // Unlock redraw
        SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(1), LPARAM(0));
        // Force repaint to avoid stale/blank regions after formatting changes.
        if !InvalidateRect(hwnd_edit, None, BOOL(1)).as_bool() {
            crate::log_debug("InvalidateRect failed in reset_line_formatting");
        }

        // Re-enable change notifications
        SendMessageW(
            hwnd_edit,
            EM_SETEVENTMASK,
            WPARAM(0),
            LPARAM((ENM_CHANGE | ENM_SELCHANGE) as isize),
        );

        // Restore modified state
        SendMessageW(
            hwnd_edit,
            EM_SETMODIFY,
            WPARAM(if was_modified { 1 } else { 0 }),
            LPARAM(0),
        );
    }
}

/// Highlights a misspelled word with red background
fn highlight_misspelled_word(hwnd_edit: HWND, start: i32, end: i32) {
    unsafe {
        // Check if this editor still has focus - don't mess with selection if not
        if GetFocus() != hwnd_edit {
            return;
        }

        // Save modified state - formatting changes should not mark document as dirty
        let was_modified = SendMessageW(hwnd_edit, EM_GETMODIFY, WPARAM(0), LPARAM(0)).0 != 0;

        // Disable change notifications during formatting
        SendMessageW(hwnd_edit, EM_SETEVENTMASK, WPARAM(0), LPARAM(0));

        // Lock the control to prevent redraws/scrolling
        SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(0), LPARAM(0));

        // Save current selection
        let mut old_sel = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut old_sel as *mut _ as isize),
        );

        // Select the misspelled word
        let mut sel = CHARRANGE {
            cpMin: start,
            cpMax: end,
        };
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut sel as *mut _ as isize),
        );

        // Apply red background color
        let mut format = CHARFORMAT2W::default();
        format.Base.cbSize = std::mem::size_of::<CHARFORMAT2W>() as u32;
        format.Base.dwMask = CFM_BACKCOLOR;
        format.crBackColor = windows::Win32::Foundation::COLORREF(0x0000FF); // Red in BGR format
        SendMessageW(
            hwnd_edit,
            EM_SETCHARFORMAT,
            WPARAM(SCF_SELECTION as usize),
            LPARAM(&mut format as *mut _ as isize),
        );

        // Restore selection
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut old_sel as *mut _ as isize),
        );

        // Unlock redraw
        SendMessageW(hwnd_edit, WM_SETREDRAW, WPARAM(1), LPARAM(0));
        // Force repaint to avoid stale/blank regions after formatting changes.
        if !InvalidateRect(hwnd_edit, None, BOOL(1)).as_bool() {
            crate::log_debug("InvalidateRect failed in highlight_misspelled_word");
        }

        // Re-enable change notifications
        SendMessageW(
            hwnd_edit,
            EM_SETEVENTMASK,
            WPARAM(0),
            LPARAM((ENM_CHANGE | ENM_SELCHANGE) as isize),
        );

        // Restore modified state
        SendMessageW(
            hwnd_edit,
            EM_SETMODIFY,
            WPARAM(if was_modified { 1 } else { 0 }),
            LPARAM(0),
        );
    }
}

pub(crate) fn show_editor_context_menu(hwnd: HWND, hwnd_edit: HWND, lparam: LPARAM) {
    unsafe {
        let language_ui = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        let labels = menu_labels(language_ui);
        let dictionary_pref = with_state(hwnd, |state| {
            state.settings.dictionary_translation_language.clone()
        })
        .unwrap_or_else(|| "auto".to_string());

        let mut spell_status = None;
        let mut spell_context = None;
        let mut fallback_msg = None;

        if let Some(word_ctx) = spellcheck_word_context_from_lparam(hwnd_edit, lparam) {
            let (status, suggestions, language, fallback) = with_state(hwnd, |state| {
                let settings = &state.settings;
                let Some(resolution) = state.spellcheck_manager.resolve_language(settings) else {
                    return (None, Vec::new(), None, None);
                };
                let fallback_msg = if resolution.announce_fallback {
                    Some(i18n::tr_f(
                        settings.language,
                        "spellcheck.language_fallback",
                        &[
                            ("requested", &resolution.requested),
                            ("language", &resolution.effective),
                        ],
                    ))
                } else {
                    None
                };
                let miss = state.spellcheck_manager.is_word_misspelled(
                    word_ctx.doc_id,
                    word_ctx.line_index,
                    &word_ctx.line_text,
                    word_ctx.word_range,
                    &resolution.effective,
                );
                if miss.is_some() {
                    let suggestions = state
                        .spellcheck_manager
                        .suggestions(&word_ctx.word, &resolution.effective);
                    (
                        Some(true),
                        suggestions,
                        Some(resolution.effective.clone()),
                        fallback_msg,
                    )
                } else {
                    (
                        Some(false),
                        Vec::new(),
                        Some(resolution.effective.clone()),
                        fallback_msg,
                    )
                }
            })
            .unwrap_or((None, Vec::new(), None, None));

            spell_status = status;
            fallback_msg = fallback;
            if status == Some(true) {
                let suggestions = suggestions
                    .into_iter()
                    .take(menu::IDM_SPELLCHECK_SUGGESTION_MAX)
                    .collect::<Vec<_>>();
                if let Some(language) = language {
                    spell_context = Some(SpellcheckContextMenuState {
                        hwnd_edit,
                        line_start: word_ctx.line_start,
                        language,
                        word_range: word_ctx.word_range,
                        word: word_ctx.word,
                        line_text: word_ctx.line_text,
                        suggestions,
                    });
                }
            }
        }

        with_state(hwnd, |state| {
            state.spellcheck_context = spell_context.clone();
        });

        if let Some(message) = fallback_msg {
            log_debug(&format!("Spellcheck: {message}"));
            nvda_speak(&message);
        }

        let menu = CreatePopupMenu().unwrap_or(HMENU(0));
        if menu.0 == 0 {
            return;
        }

        if let Some(word_ctx) = spellcheck_word_context_from_lparam(hwnd_edit, lparam)
            && let Ok(submenu) = CreatePopupMenu()
            && submenu.0 != 0
        {
            let placeholder = format!(" {}", i18n::tr(language_ui, "dictionary.menu_expand"));
            crate::log_if_err!(AppendMenuW(
                submenu,
                MF_STRING | MF_GRAYED,
                0,
                PCWSTR(to_wide(&placeholder).as_ptr()),
            ));
            let label = i18n::tr(language_ui, "context_menu.dictionary");
            crate::log_if_err!(AppendMenuW(
                menu,
                MF_POPUP,
                submenu.0 as usize,
                PCWSTR(to_wide(&label).as_ptr()),
            ));
            crate::log_if_err!(AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()));
            let prefetch_info = with_state(hwnd, |state| {
                state.dictionary_context_menu = submenu;
                state.dictionary_context_word = word_ctx.word.clone();
                state.dictionary_context_language = language_ui;
                state.dictionary_context_pref = dictionary_pref.clone();
                state.dictionary_context_loaded = false;
                state.dictionary_prefetch_generation =
                    state.dictionary_prefetch_generation.wrapping_add(1);
                let generation = state.dictionary_prefetch_generation;

                let key = dictionary_cache_key(language_ui, &dictionary_pref, &word_ctx.word);
                if let Some(lines) = state.dictionary_cache.get(&key).cloned() {
                    if is_dictionary_not_found_cache_entry(language_ui, &lines) {
                        state.dictionary_cache.remove(&key);
                        save_dictionary_cache(&state.dictionary_cache);
                    } else {
                        return None;
                    }
                }
                if state.dictionary_pending_lookup.as_ref() == Some(&key) {
                    return None;
                }
                state.dictionary_pending_lookup = Some(key.clone());
                Some((word_ctx.word.clone(), key, generation))
            })
            .flatten();
            if let Some((word, key, generation)) = prefetch_info {
                start_dictionary_lookup(
                    hwnd.0,
                    word,
                    language_ui,
                    dictionary_pref.clone(),
                    key,
                    generation,
                );
            }
        }

        if let Some(status) = spell_status {
            if status {
                let label = i18n::tr(language_ui, "context_menu.spelling_misspelled");
                crate::log_if_err!(AppendMenuW(
                    menu,
                    MF_STRING | MF_GRAYED,
                    0,
                    PCWSTR(to_wide(&label).as_ptr()),
                ));
                if let Ok(submenu) = CreatePopupMenu()
                    && submenu.0 != 0
                {
                    let suggestions = spell_context
                        .as_ref()
                        .map(|ctx| ctx.suggestions.as_slice())
                        .unwrap_or(&[]);
                    if suggestions.is_empty() {
                        let none_label =
                            i18n::tr(language_ui, "context_menu.spelling_no_suggestions");
                        crate::log_if_err!(AppendMenuW(
                            submenu,
                            MF_STRING | MF_GRAYED,
                            0,
                            PCWSTR(to_wide(&none_label).as_ptr()),
                        ));
                    } else {
                        for (idx, suggestion) in suggestions.iter().enumerate() {
                            let id = menu::IDM_SPELLCHECK_SUGGESTION_BASE + idx;
                            crate::log_if_err!(AppendMenuW(
                                submenu,
                                MF_STRING,
                                id,
                                PCWSTR(to_wide(suggestion).as_ptr()),
                            ));
                        }
                    }
                    crate::log_if_err!(AppendMenuW(submenu, MF_SEPARATOR, 0, PCWSTR::null()));
                    let add_label =
                        i18n::tr(language_ui, "context_menu.spelling_add_to_dictionary");
                    let ignore_label = i18n::tr(language_ui, "context_menu.spelling_ignore_once");
                    crate::log_if_err!(AppendMenuW(
                        submenu,
                        MF_STRING,
                        menu::IDM_SPELLCHECK_ADD_TO_DICTIONARY,
                        PCWSTR(to_wide(&add_label).as_ptr()),
                    ));
                    crate::log_if_err!(AppendMenuW(
                        submenu,
                        MF_STRING,
                        menu::IDM_SPELLCHECK_IGNORE_ONCE,
                        PCWSTR(to_wide(&ignore_label).as_ptr()),
                    ));
                    let suggestions_label =
                        i18n::tr(language_ui, "context_menu.spelling_suggestions");
                    crate::log_if_err!(AppendMenuW(
                        menu,
                        MF_POPUP,
                        submenu.0 as usize,
                        PCWSTR(to_wide(&suggestions_label).as_ptr()),
                    ));
                }
            } else {
                let label = i18n::tr(language_ui, "context_menu.spelling_ok");
                crate::log_if_err!(AppendMenuW(
                    menu,
                    MF_STRING | MF_GRAYED,
                    0,
                    PCWSTR(to_wide(&label).as_ptr()),
                ));
            }
            crate::log_if_err!(AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()));
        }

        let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut selection as *mut _ as isize),
        );
        if selection.cpMin > selection.cpMax {
            std::mem::swap(&mut selection.cpMin, &mut selection.cpMax);
        }
        let has_selection = selection.cpMin != selection.cpMax;
        if has_selection {
            let label = i18n::tr(language_ui, "context_menu.audiobook_selection");
            crate::log_if_err!(AppendMenuW(
                menu,
                MF_STRING,
                menu::IDM_EDIT_AUDIOBOOK_SELECTION,
                PCWSTR(to_wide(&label).as_ptr()),
            ));
            crate::log_if_err!(AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()));
        }

        let undo_flags = if can_undo_now(hwnd) {
            MF_STRING
        } else {
            MF_STRING | MF_GRAYED
        };
        let cut_copy_flags = if has_selection {
            MF_STRING
        } else {
            MF_STRING | MF_GRAYED
        };
        let paste_flags = if can_paste_now(hwnd) {
            MF_STRING
        } else {
            MF_STRING | MF_GRAYED
        };
        let undo_label = build_undo_menu_label(hwnd, language_ui);
        crate::log_if_err!(AppendMenuW(
            menu,
            undo_flags,
            IDM_EDIT_UNDO,
            PCWSTR(to_wide(&undo_label).as_ptr()),
        ));
        crate::log_if_err!(AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()));
        crate::log_if_err!(AppendMenuW(
            menu,
            cut_copy_flags,
            IDM_EDIT_CUT,
            PCWSTR(to_wide(&labels.edit_cut).as_ptr()),
        ));
        crate::log_if_err!(AppendMenuW(
            menu,
            cut_copy_flags,
            IDM_EDIT_COPY,
            PCWSTR(to_wide(&labels.edit_copy).as_ptr()),
        ));
        crate::log_if_err!(AppendMenuW(
            menu,
            paste_flags,
            IDM_EDIT_PASTE,
            PCWSTR(to_wide(&labels.edit_paste).as_ptr()),
        ));
        crate::log_if_err!(AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()));
        crate::log_if_err!(AppendMenuW(
            menu,
            MF_STRING,
            IDM_EDIT_SELECT_ALL,
            PCWSTR(to_wide(&labels.edit_select_all).as_ptr()),
        ));

        let mut x = (lparam.0 & 0xffff) as i32;
        let mut y = ((lparam.0 >> 16) & 0xffff) as i32;
        if x == -1 && y == -1 {
            let mut pt = POINT::default();
            crate::log_if_err!(GetCursorPos(&mut pt));
            x = pt.x;
            y = pt.y;
        }
        SetForegroundWindow(hwnd);
        if !TrackPopupMenu(menu, TPM_RIGHTBUTTON, x, y, 0, hwnd, None).as_bool() {
            crate::log_debug("TrackPopupMenu failed");
        }
        crate::log_if_err!(PostMessageW(hwnd, WM_NULL, WPARAM(0), LPARAM(0)));
        with_state(hwnd, |state| {
            state.dictionary_context_menu = HMENU(0);
            state.dictionary_context_word.clear();
            state.dictionary_context_pref.clear();
            state.dictionary_context_loaded = false;
            state.dictionary_context_expanded = false;
        });
    }
}

fn open_dictionary_lookup(hwnd: HWND) {
    app_windows::wiktionary_window::open(hwnd);
}

fn can_undo_now(hwnd: HWND) -> bool {
    if unsafe { with_state(hwnd, |state| state.normalize_undo.is_some()).unwrap_or(false) } {
        return true;
    }
    let Some(hwnd_edit) = (unsafe { get_active_edit(hwnd) }) else {
        return false;
    };
    // SAFETY: querying EM_CANUNDO on the active edit control is side-effect free.
    unsafe { SendMessageW(hwnd_edit, EM_CANUNDO, WPARAM(0), LPARAM(0)).0 != 0 }
}

pub(crate) fn update_main_status_bar(hwnd: HWND) {
    let (hwnd_status, language) =
        unsafe { with_state(hwnd, |state| (state.hwnd_status, state.settings.language)) }
            .unwrap_or((HWND(0), Language::default()));
    if hwnd_status.0 == 0 {
        return;
    }
    let (chars, words, line, col) = if let Some(hwnd_edit) = unsafe { get_active_edit(hwnd) } {
        let text = editor_manager::get_edit_text(hwnd_edit);
        let chars = text.chars().count();
        let words = text.split_whitespace().count();
        let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
        unsafe {
            SendMessageW(
                hwnd_edit,
                EM_EXGETSEL,
                WPARAM(0),
                LPARAM(&mut selection as *mut _ as isize),
            );
        }
        let caret = selection.cpMax.max(0);
        let line_idx = unsafe {
            SendMessageW(
                hwnd_edit,
                EM_LINEFROMCHAR,
                WPARAM(caret as usize),
                LPARAM(0),
            )
            .0 as i32
        };
        let line_start = unsafe {
            SendMessageW(
                hwnd_edit,
                EM_LINEINDEX,
                WPARAM(line_idx.max(0) as usize),
                LPARAM(0),
            )
            .0 as i32
        };
        let col = (caret - line_start).max(0) + 1;
        (chars, words, line_idx.max(0) + 1, col)
    } else {
        (0, 0, 1, 1)
    };
    let chars_str = chars.to_string();
    let words_str = words.to_string();
    let chars_label = i18n::tr_f(
        language,
        "text_stats.characters_with_spaces",
        &[("count", &chars_str)],
    );
    let words_label = i18n::tr_f(language, "text_stats.words", &[("count", &words_str)]);
    let chars_label = chars_label.trim_end_matches('.');
    let words_label = words_label.trim_end_matches('.');
    let label = format!(
        "{}. | {}. | Ln {}, Col {}",
        chars_label, words_label, line, col
    );
    let label_wide = to_wide(&label);
    unsafe {
        SendMessageW(
            hwnd_status,
            SB_SETTEXTW,
            WPARAM(0),
            LPARAM(label_wide.as_ptr() as isize),
        );
    }
}

fn build_undo_menu_label(hwnd: HWND, language: Language) -> String {
    let base = i18n::tr(language, "edit.undo");
    let action = unsafe { with_state(hwnd, |state| state.undo_action_label.clone()).flatten() };
    let Some(action) = action else {
        return base;
    };
    let action = action.trim();
    if action.is_empty() {
        return base;
    }
    let mut split = base.splitn(2, '\t');
    let caption = split.next().unwrap_or("").trim();
    let accel = split.next();
    if caption.is_empty() {
        return base;
    }
    match accel {
        Some(accel) if !accel.is_empty() => format!("{}: {}\t{}", caption, action, accel),
        _ => format!("{}: {}", caption, action),
    }
}

fn has_active_text_selection(hwnd: HWND) -> bool {
    let Some(hwnd_edit) = (unsafe { get_active_edit(hwnd) }) else {
        return false;
    };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    // SAFETY: `selection` is valid writable memory and `hwnd_edit` is the active edit control.
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut selection as *mut _ as isize),
        );
    }
    selection.cpMin != selection.cpMax
}

fn can_paste_now(hwnd: HWND) -> bool {
    if unsafe { get_active_edit(hwnd) }.is_none() {
        return false;
    }
    // CF_UNICODETEXT = 13
    unsafe { IsClipboardFormatAvailable(13).is_ok() }
}

fn show_voice_context_menu(hwnd: HWND, target: HWND, lparam: LPARAM) {
    let (combo_voice, combo_favorites, engine, language) = unsafe {
        with_state(hwnd, |state| {
            (
                state.voice_combo_voice,
                state.voice_combo_favorites,
                state.settings.tts_engine,
                state.settings.language,
            )
        })
    }
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
        let is_favorite = unsafe {
            with_state(hwnd, |state| {
                state
                    .settings
                    .favorite_voices
                    .iter()
                    .any(|fav| fav.engine == engine && fav.short_name == voice_name)
            })
        }
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
    unsafe {
        let menu = CreatePopupMenu().unwrap_or(HMENU(0));
        if menu.0 == 0 {
            return;
        }
        crate::log_if_err!(AppendMenuW(
            menu,
            MF_STRING,
            action_id as usize,
            PCWSTR(to_wide(&action_label).as_ptr()),
        ));
        with_state(hwnd, |state| {
            state.voice_context_voice = Some(ctx);
        });

        let mut x = (lparam.0 & 0xffff) as i32;
        let mut y = ((lparam.0 >> 16) & 0xffff) as i32;
        if x == -1 && y == -1 {
            let mut pt = windows::Win32::Foundation::POINT::default();
            crate::log_if_err!(GetCursorPos(&mut pt));
            x = pt.x;
            y = pt.y;
        }

        SetForegroundWindow(hwnd);
        if !TrackPopupMenu(menu, TPM_RIGHTBUTTON, x, y, 0, hwnd, None).as_bool() {
            crate::log_debug("TrackPopupMenu failed");
        }
        crate::log_if_err!(PostMessageW(hwnd, WM_NULL, WPARAM(0), LPARAM(0)));
    }
}

fn replace_spellcheck_word(hwnd_edit: HWND, ctx: &SpellcheckContextMenuState, replacement: &str) {
    let start_utf16 = ctx.line_start
        + spellcheck::utf8_byte_offset_to_utf16_units(&ctx.line_text, ctx.word_range.0);
    let end_utf16 = ctx.line_start
        + spellcheck::utf8_byte_offset_to_utf16_units(&ctx.line_text, ctx.word_range.1);
    let mut range = CHARRANGE {
        cpMin: start_utf16,
        cpMax: end_utf16,
    };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut range as *mut _ as isize),
        );
    }
    let wide = to_wide(replacement);
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(wide.as_ptr() as isize),
        );
    }
    let new_end =
        start_utf16 + spellcheck::utf8_byte_offset_to_utf16_units(replacement, replacement.len());
    let mut new_sel = CHARRANGE {
        cpMin: new_end,
        cpMax: new_end,
    };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut new_sel as *mut _ as isize),
        );
    }
}

fn handle_spellcheck_suggestion(hwnd: HWND, index: usize) {
    let ctx = unsafe { with_state(hwnd, |state| state.spellcheck_context.clone()) }.unwrap_or(None);
    let Some(ctx) = ctx else {
        return;
    };
    let Some(replacement) = ctx.suggestions.get(index).cloned() else {
        return;
    };
    if ctx.hwnd_edit.0 != 0 {
        replace_spellcheck_word(ctx.hwnd_edit, &ctx, &replacement);
    }
    unsafe {
        with_state(hwnd, |state| {
            state.spellcheck_manager.clear_cache();
            state.spellcheck_last_announce = None;
            state.spellcheck_context = None;
        });
    }
}

fn handle_spellcheck_add_to_dictionary(hwnd: HWND) {
    let ctx = unsafe { with_state(hwnd, |state| state.spellcheck_context.clone()) }.unwrap_or(None);
    let Some(ctx) = ctx else {
        return;
    };
    unsafe {
        with_state(hwnd, |state| {
            state
                .spellcheck_manager
                .add_to_dictionary(&ctx.word, &ctx.language);
            state.spellcheck_last_announce = None;
            state.spellcheck_context = None;
        });
    }
}

fn handle_spellcheck_ignore_once(hwnd: HWND) {
    let ctx = unsafe { with_state(hwnd, |state| state.spellcheck_context.clone()) }.unwrap_or(None);
    let Some(ctx) = ctx else {
        return;
    };
    unsafe {
        with_state(hwnd, |state| {
            state
                .spellcheck_manager
                .ignore_once(&ctx.word, &ctx.language);
            state.spellcheck_last_announce = None;
            state.spellcheck_context = None;
        });
    }
}

/// Navigate to next (forward=true) or previous (forward=false) spelling error
fn go_to_spelling_error(hwnd: HWND, forward: bool) {
    use windows::Win32::UI::Controls::RichEdit::{CHARRANGE, EM_EXGETSEL, EM_EXSETSEL};

    let Some(hwnd_edit) = (unsafe { get_active_edit(hwnd) }) else {
        return;
    };

    // Get spellcheck language
    let resolution = unsafe {
        with_state(hwnd, |state| {
            state.spellcheck_manager.resolve_language(&state.settings)
        })
    }
    .flatten();
    let Some(resolution) = resolution else {
        // Spellcheck disabled or no language available
        return;
    };

    // Get current cursor position
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut cr as *mut _ as isize),
        );
    }
    let current_pos = if forward { cr.cpMax } else { cr.cpMin };

    // Get document info
    let doc_id = unsafe {
        with_state(hwnd, |state| {
            state
                .docs
                .iter()
                .find(|d| d.hwnd_edit == hwnd_edit)
                .map(|d| d.hwnd_edit.0)
        })
    }
    .flatten()
    .unwrap_or(0);

    let text = editor_manager::get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    // Collect all misspellings from all lines
    let mut all_errors: Vec<(i32, i32)> = Vec::new(); // (start_utf16, end_utf16)

    let mut line_start_utf16 = 0i32;
    for (line_idx, line) in text.lines().enumerate() {
        let misspellings = unsafe {
            with_state(hwnd, |state| {
                state.spellcheck_manager.check_line(
                    doc_id,
                    line_idx as i32,
                    line,
                    &resolution.effective,
                )
            })
        }
        .unwrap_or_default();

        for m in misspellings {
            // Convert byte offsets to UTF-16 offsets
            let start_byte = m.start;
            let end_byte = m.end;
            let prefix = &line[..start_byte.min(line.len())];
            let word_part = &line[start_byte.min(line.len())..end_byte.min(line.len())];
            let start_utf16_in_line: i32 = prefix.encode_utf16().count() as i32;
            let word_utf16_len: i32 = word_part.encode_utf16().count() as i32;

            let abs_start = line_start_utf16 + start_utf16_in_line;
            let abs_end = abs_start + word_utf16_len;
            all_errors.push((abs_start, abs_end));
        }

        // Account for line ending (could be \r\n or \n)
        let line_utf16_len: i32 = line.encode_utf16().count() as i32;
        line_start_utf16 += line_utf16_len + 1; // +1 for \n (simplified)
    }

    if all_errors.is_empty() {
        return;
    }

    // Find the next/previous error relative to current position (no wrap-around)
    let target = if forward {
        // Find first error after current position
        all_errors.iter().find(|(start, _)| *start > current_pos)
    } else {
        // Find last error before current position
        all_errors.iter().rev().find(|(_, end)| *end < current_pos)
    };

    if let Some(&(start, end)) = target {
        // Select the misspelled word
        let mut new_range = CHARRANGE {
            cpMin: start,
            cpMax: end,
        };
        unsafe {
            SendMessageW(
                hwnd_edit,
                EM_EXSETSEL,
                WPARAM(0),
                LPARAM(&mut new_range as *mut _ as isize),
            );
            // Scroll to make visible
            SendMessageW(
                hwnd_edit,
                crate::accessibility::EM_SCROLLCARET,
                WPARAM(0),
                LPARAM(0),
            );
        }
    }
}

fn handle_voice_context_favorite(hwnd: HWND, add: bool) {
    let ctx =
        unsafe { with_state(hwnd, |state| state.voice_context_voice.clone()) }.unwrap_or(None);
    let Some(fav) = ctx else {
        return;
    };
    if add {
        add_favorite_voice(hwnd, fav.engine, &fav.short_name);
    } else {
        remove_favorite_voice(hwnd, fav.engine, &fav.short_name);
    }
    unsafe {
        with_state(hwnd, |state| {
            state.voice_context_voice = None;
        });
    }
}

fn add_favorite_voice(hwnd: HWND, engine: TtsEngine, voice_name: &str) {
    unsafe {
        with_state(hwnd, |state| {
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
    }
    if let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) } {
        save_settings(settings);
    }
    unsafe {
        refresh_voice_panel(hwnd);
    }
}

fn remove_favorite_voice(hwnd: HWND, engine: TtsEngine, voice_name: &str) {
    unsafe {
        with_state(hwnd, |state| {
            state
                .settings
                .favorite_voices
                .retain(|fav| !(fav.engine == engine && fav.short_name == voice_name));
        });
    }
    if let Some(settings) = unsafe { with_state(hwnd, |state| state.settings.clone()) } {
        save_settings(settings);
    }
    unsafe {
        refresh_voice_panel(hwnd);
    }
}

fn is_focus_in_voice_panel(hwnd: HWND) -> bool {
    let focus = unsafe { GetFocus() };
    if focus.0 == 0 {
        return false;
    }

    let mut class_buf = [0u16; 64];
    let len = unsafe { GetClassNameW(focus, &mut class_buf) };
    let class_name = String::from_utf16_lossy(&class_buf[..len as usize]);
    if class_name == "ComboLBox" {
        return true;
    }

    unsafe {
        with_state(hwnd, |state| {
            if !state.voice_panel_visible && !state.voice_favorites_visible {
                return false;
            }
            let is_match =
                |ctrl: HWND| ctrl.0 != 0 && (focus == ctrl || IsChild(ctrl, focus).as_bool());
            is_match(state.voice_combo_engine)
                || is_match(state.voice_combo_language)
                || is_match(state.voice_combo_voice)
                || is_match(state.voice_button_insert_tag)
                || is_match(state.voice_combo_speed)
                || is_match(state.voice_combo_pitch)
                || is_match(state.voice_combo_volume)
                || is_match(state.voice_edit_speed)
                || is_match(state.voice_edit_pitch)
                || is_match(state.voice_edit_volume)
                || is_match(state.voice_checkbox_multilingual)
                || is_match(state.voice_combo_favorites)
        })
    }
    .unwrap_or(false)
}

fn handle_voice_panel_tab(hwnd: HWND) -> bool {
    unsafe {
        let (
            visible,
            combo_engine,
            combo_language,
            combo_voice,
            button_insert_tag,
            combo_speed,
            combo_pitch,
            combo_volume,
            edit_speed,
            edit_pitch,
            edit_volume,
            checkbox_multilingual,
            combo_favorites,
            favorites_visible,
            is_edge,
            only_multilingual,
            manual_tuning,
            hwnd_tab,
        ) = match with_state(hwnd, |state| {
            (
                state.voice_panel_visible,
                state.voice_combo_engine,
                state.voice_combo_language,
                state.voice_combo_voice,
                state.voice_button_insert_tag,
                state.voice_combo_speed,
                state.voice_combo_pitch,
                state.voice_combo_volume,
                state.voice_edit_speed,
                state.voice_edit_pitch,
                state.voice_edit_volume,
                state.voice_checkbox_multilingual,
                state.voice_combo_favorites,
                state.voice_favorites_visible,
                matches!(state.settings.tts_engine, TtsEngine::Edge),
                state.settings.tts_only_multilingual,
                state.settings.tts_manual_tuning,
                state.hwnd_tab,
            )
        }) {
            Some(values) => values,
            None => return false,
        };
        if !visible && !favorites_visible {
            return false;
        }
        let raw_focus = GetFocus();
        if raw_focus.0 == 0 {
            return false;
        }
        let focus = if raw_focus == combo_engine || IsChild(combo_engine, raw_focus).as_bool() {
            combo_engine
        } else if raw_focus == combo_language || IsChild(combo_language, raw_focus).as_bool() {
            combo_language
        } else if raw_focus == combo_voice || IsChild(combo_voice, raw_focus).as_bool() {
            combo_voice
        } else if raw_focus == button_insert_tag || IsChild(button_insert_tag, raw_focus).as_bool()
        {
            button_insert_tag
        } else if raw_focus == combo_speed || IsChild(combo_speed, raw_focus).as_bool() {
            combo_speed
        } else if raw_focus == combo_pitch || IsChild(combo_pitch, raw_focus).as_bool() {
            combo_pitch
        } else if raw_focus == combo_volume || IsChild(combo_volume, raw_focus).as_bool() {
            combo_volume
        } else if raw_focus == combo_favorites || IsChild(combo_favorites, raw_focus).as_bool() {
            combo_favorites
        } else {
            raw_focus
        };
        let is_combo_focus = focus == combo_engine
            || (is_edge && !only_multilingual && focus == combo_language)
            || focus == combo_voice
            || focus == button_insert_tag
            || (!manual_tuning && focus == combo_speed)
            || (!manual_tuning && focus == combo_pitch)
            || (!manual_tuning && focus == combo_volume)
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
        let speed_control = if manual_tuning {
            edit_speed
        } else {
            combo_speed
        };
        let pitch_control = if manual_tuning {
            edit_pitch
        } else {
            combo_pitch
        };
        let volume_control = if manual_tuning {
            edit_volume
        } else {
            combo_volume
        };
        if focus != hwnd_edit
            && focus != combo_engine
            && focus != combo_language
            && focus != combo_voice
            && focus != button_insert_tag
            && focus != speed_control
            && focus != pitch_control
            && focus != volume_control
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
        let mut order = Vec::new();
        if visible {
            order.push(combo_engine);
            if is_edge && !only_multilingual {
                order.push(combo_language);
            }
            order.push(combo_voice);
            order.push(button_insert_tag);
            order.push(speed_control);
            order.push(pitch_control);
            order.push(volume_control);
            if is_edge {
                order.push(checkbox_multilingual);
            }
        }
        if favorites_visible {
            order.push(combo_favorites);
        }
        let Some(idx) = order.iter().position(|item| *item == focus) else {
            return false;
        };
        if shift_down {
            if idx == 0 {
                if fallback_edit.0 != 0 {
                    SetFocus(fallback_edit);
                    return true;
                }
                return false;
            }
            let target = order[idx - 1];
            if target.0 != 0 {
                SetFocus(target);
                return true;
            }
        } else {
            if idx + 1 >= order.len() {
                if fallback_edit.0 != 0 {
                    SetFocus(fallback_edit);
                    return true;
                }
                return false;
            }
            let target = order[idx + 1];
            if target.0 != 0 {
                SetFocus(target);
                return true;
            }
        }
        false
    }
}

fn is_modifier_vk(key: u16) -> bool {
    matches!(key, 0x10 | 0x11 | 0x12 | 0xA0..=0xA5)
}

fn appcommand_from_lparam(lparam: LPARAM) -> usize {
    ((lparam.0 as usize) >> 16) & 0x7ff
}

fn shortcut_matches_message(binding: ShortcutBinding, msg: &MSG) -> bool {
    if msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN {
        return false;
    }
    let key = msg.wParam.0 as u16;
    if is_modifier_vk(key) {
        return false;
    }
    // SAFETY: key-state reads are thread-local Win32 queries with no aliasing requirements.
    let ctrl_down = unsafe { (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0 };
    // SAFETY: key-state reads are thread-local Win32 queries with no aliasing requirements.
    let shift_down = unsafe { (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0 };
    // SAFETY: key-state reads are thread-local Win32 queries with no aliasing requirements.
    let alt_down = unsafe { (GetKeyState(VK_MENU.0 as i32) & (0x8000u16 as i16)) != 0 };
    key == binding.key
        && ctrl_down == binding.ctrl
        && shift_down == binding.shift
        && alt_down == binding.alt
}

fn dispatch_shortcut_command(hwnd: HWND, cmd: usize) {
    unsafe {
        crate::log_if_err!(PostMessageW(hwnd, WM_COMMAND, WPARAM(cmd), LPARAM(0)));
    }
}

fn handle_custom_shortcuts(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN {
        return false;
    }
    let options_hwnd = unsafe { with_state(hwnd, |state| state.options_dialog) }.unwrap_or(HWND(0));
    if options_hwnd.0 != 0
        && (msg.hwnd == options_hwnd || unsafe { IsChild(options_hwnd, msg.hwnd).as_bool() })
    {
        return false;
    }

    let shortcuts = unsafe { with_state(hwnd, |state| state.settings.shortcuts.clone()) }
        .unwrap_or_else(ShortcutSettings::default);

    // Dedicated shortcut requested for streaming audio.
    // NOTE: this intentionally takes precedence over the accelerator table.
    let key = msg.wParam.0 as u16;
    let ctrl_down = unsafe { (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0 };
    let shift_down = unsafe { (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0 };
    let alt_down = unsafe { (GetKeyState(VK_MENU.0 as i32) & (0x8000u16 as i16)) != 0 };
    if key == 'S' as u16 && !ctrl_down && shift_down && alt_down {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_STREAM_AUDIO);
        return true;
    }

    if shortcut_matches_message(shortcuts.read_pause_resume, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_READ_PAUSE);
        return true;
    }
    if shortcut_matches_message(shortcuts.read_start, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_READ_START);
        return true;
    }
    if shortcut_matches_message(shortcuts.read_stop, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_READ_STOP);
        return true;
    }
    if shortcut_matches_message(shortcuts.execute_file, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_EXECUTE);
        return true;
    }
    if shortcut_matches_message(shortcuts.audiobook, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_AUDIOBOOK);
        return true;
    }
    if shortcut_matches_message(shortcuts.batch_audiobooks, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_BATCH_AUDIOBOOK);
        return true;
    }
    if shortcut_matches_message(shortcuts.record_podcast, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_PODCAST);
        return true;
    }
    if shortcut_matches_message(shortcuts.convert_audio, msg) {
        dispatch_shortcut_command(hwnd, IDM_FILE_CONVERT_AUDIO);
        return true;
    }
    if shortcut_matches_message(shortcuts.open_rss, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_RSS);
        return true;
    }
    if shortcut_matches_message(shortcuts.open_podcasts, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_PODCASTS);
        return true;
    }
    if shortcut_matches_message(shortcuts.open_dictionary, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_DICTIONARY);
        return true;
    }
    if shortcut_matches_message(shortcuts.open_options, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_OPTIONS);
        return true;
    }
    if shortcut_matches_message(shortcuts.open_terminal, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_PROMPT);
        return true;
    }
    if shortcut_matches_message(shortcuts.import_wikipedia, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_WIKIPEDIA_IMPORT);
        return true;
    }
    if shortcut_matches_message(shortcuts.import_youtube, msg) {
        dispatch_shortcut_command(hwnd, IDM_TOOLS_IMPORT_YOUTUBE);
        return true;
    }
    if shortcut_matches_message(shortcuts.find, msg) {
        dispatch_shortcut_command(hwnd, IDM_EDIT_FIND);
        return true;
    }
    if shortcut_matches_message(shortcuts.quote_lines, msg) {
        dispatch_shortcut_command(hwnd, IDM_EDIT_QUOTE_LINES);
        return true;
    }
    if shortcut_matches_message(shortcuts.unquote_lines, msg) {
        dispatch_shortcut_command(hwnd, IDM_EDIT_UNQUOTE_LINES);
        return true;
    }
    if shortcut_matches_message(shortcuts.media_prev, msg) {
        dispatch_shortcut_command(hwnd, IDM_PLAYBACK_TRACK_PREV);
        return true;
    }
    if shortcut_matches_message(shortcuts.media_next, msg) {
        dispatch_shortcut_command(hwnd, IDM_PLAYBACK_TRACK_NEXT);
        return true;
    }
    if shortcut_matches_message(shortcuts.chapter_prev, msg) {
        dispatch_shortcut_command(hwnd, IDM_PLAYBACK_CHAPTER_PREV);
        return true;
    }
    if shortcut_matches_message(shortcuts.chapter_next, msg) {
        dispatch_shortcut_command(hwnd, IDM_PLAYBACK_CHAPTER_NEXT);
        return true;
    }
    false
}

fn create_accelerators() -> HACCEL {
    unsafe {
        let virt = FCONTROL | FVIRTKEY;
        let virt_shift = FCONTROL | FSHIFT | FVIRTKEY;
        let virt_shift_only = FSHIFT | FVIRTKEY;
        let virt_alt = FALT | FVIRTKEY;
        let virt_alt_shift = FALT | FSHIFT | FVIRTKEY;
        let accels = [
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
                fVirt: virt_shift,
                key: 'W' as u16,
                cmd: IDM_FILE_CLOSE_OTHERS as u16,
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
                key: VK_RETURN.0,
                cmd: IDM_EDIT_NORMALIZE_WHITESPACE as u16,
            },
            ACCEL {
                fVirt: FVIRTKEY,
                key: VK_F3.0,
                cmd: IDM_EDIT_FIND_NEXT as u16,
            },
            ACCEL {
                fVirt: virt_shift_only,
                key: VK_F3.0,
                cmd: IDM_EDIT_FIND_PREVIOUS as u16,
            },
            ACCEL {
                fVirt: virt,
                key: 'H' as u16,
                cmd: IDM_EDIT_REPLACE as u16,
            },
            ACCEL {
                fVirt: virt,
                key: 'J' as u16,
                cmd: IDM_EDIT_GO_TO_LINE as u16,
            },
            ACCEL {
                fVirt: virt,
                key: 'A' as u16,
                cmd: IDM_EDIT_SELECT_ALL as u16,
            },
            ACCEL {
                fVirt: virt_shift,
                key: VK_OEM_PERIOD.0,
                cmd: IDM_EDIT_INDENT as u16,
            },
            ACCEL {
                fVirt: virt,
                key: VK_OEM_PERIOD.0,
                cmd: IDM_EDIT_INSERT_ELLIPSIS as u16,
            },
            ACCEL {
                fVirt: virt_shift,
                key: VK_OEM_COMMA.0,
                cmd: IDM_EDIT_OUTDENT as u16,
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
                key: 'D' as u16,
                cmd: IDM_EDIT_REMOVE_DUPLICATE_LINES as u16,
            },
            ACCEL {
                fVirt: virt_shift,
                key: 'C' as u16,
                cmd: IDM_EDIT_REMOVE_DUPLICATE_CONSECUTIVE_LINES as u16,
            },
            ACCEL {
                fVirt: virt_alt_shift,
                key: 'H' as u16,
                cmd: IDM_EDIT_CLEAN_EOL_HYPHENS as u16,
            },
            ACCEL {
                fVirt: virt_alt_shift,
                key: 'D' as u16,
                cmd: IDM_TOOLS_DICTIONARY_LOOKUP as u16,
            },
            ACCEL {
                fVirt: virt,
                key: VK_TAB.0,
                cmd: IDM_NEXT_TAB as u16,
            },
            ACCEL {
                fVirt: virt_shift_only,
                key: VK_NEXT.0,
                cmd: IDM_GOTO_NEXT_BOOKMARK as u16,
            },
            ACCEL {
                fVirt: virt_shift_only,
                key: VK_PRIOR.0,
                cmd: IDM_GOTO_PREV_BOOKMARK as u16,
            },
            ACCEL {
                fVirt: virt_shift,
                key: 'G' as u16,
                cmd: IDM_MANAGE_BOOKMARKS as u16,
            },
            ACCEL {
                fVirt: virt_shift,
                key: 'L' as u16,
                cmd: IDM_INSERT_CLEAR_BOOKMARKS as u16,
            },
            ACCEL {
                fVirt: virt,
                key: 'B' as u16,
                cmd: IDM_INSERT_BOOKMARK as u16,
            },
        ];
        CreateAcceleratorTableW(&accels).unwrap_or(HACCEL(0))
    }
}

unsafe extern "system" fn enum_close_other_windows(hwnd: HWND, lparam: LPARAM) -> BOOL {
    crate::panic_guard::guard(
        "enum_close_other_windows",
        || BOOL(1),
        || {
            let current = HWND(lparam.0);
            if hwnd == current {
                return BOOL(1);
            }
            let mut buf = [0u16; 64];
            let len = GetClassNameW(hwnd, &mut buf);
            if len == 0 {
                return BOOL(1);
            }
            let name = String::from_utf16_lossy(&buf[..len as usize]);
            if name == "SonarpadWin32" {
                crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
            }
            BOOL(1)
        },
    )
}

fn close_other_windows(hwnd: HWND) {
    unsafe {
        crate::log_if_err!(EnumWindows(Some(enum_close_other_windows), LPARAM(hwnd.0)));
    }
}

pub(crate) unsafe fn get_active_edit(hwnd: HWND) -> Option<HWND> {
    with_state(hwnd, |state| {
        state.docs.get(state.current).map(|doc| doc.hwnd_edit)
    })
    .flatten()
}

const UNSAVED_BOOKMARK_PREFIX: &str = "__unsaved__:";

fn bookmark_storage_key(path: Option<&Path>, hwnd_edit: HWND) -> (String, bool) {
    if let Some(path) = path {
        (path.to_string_lossy().to_string(), true)
    } else {
        (format!("{UNSAVED_BOOKMARK_PREFIX}{}", hwnd_edit.0), false)
    }
}

fn insert_bookmark(hwnd: HWND) {
    unsafe {
        let (hwnd_edit, path, format): (HWND, Option<std::path::PathBuf>, FileFormat) =
            with_state(hwnd, |state| {
                state
                    .docs
                    .get(state.current)
                    .map(|doc| (doc.hwnd_edit, doc.path.clone(), doc.format))
            })
            .flatten()
            .unwrap_or((HWND(0), None, FileFormat::default()));
        if hwnd_edit.0 == 0 {
            return;
        }
        let (storage_key, persist_to_disk) = bookmark_storage_key(path.as_deref(), hwnd_edit);

        if matches!(format, FileFormat::Audiobook) {
            let (pos, snippet) = with_state(hwnd, |state| {
                if let Some(player) = &mut state.active_audiobook {
                    let position_secs = crate::audio_player::audiobook_position_secs(player);
                    audio_bookmark_position_and_snippet(position_secs)
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

            let bookmarks_window = with_state(hwnd, |state| {
                let list = state
                    .bookmarks
                    .files
                    .entry(storage_key.clone())
                    .or_default();
                list.push(bookmark);
                if persist_to_disk {
                    save_bookmarks(&state.bookmarks);
                }
                state.bookmarks_window
            })
            .unwrap_or(HWND(0));

            if bookmarks_window.0 != 0 {
                app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window);
            }
            confirm_menu_action(hwnd, "insert.bookmark");
            return;
        }

        let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut cr as *mut _ as isize),
        );

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
        let copied = SendMessageW(
            hwnd_edit,
            EM_GETTEXTRANGE,
            WPARAM(0),
            LPARAM(&mut tr as *mut _ as isize),
        )
        .0 as usize;
        let mut snippet = String::from_utf16_lossy(&buffer[..copied]);

        // Stop at the first newline
        if let Some(idx) = snippet.find(['\r', '\n']) {
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
            let copied_pre = SendMessageW(
                hwnd_edit,
                EM_GETTEXTRANGE,
                WPARAM(0),
                LPARAM(&mut tr_pre as *mut _ as isize),
            )
            .0 as usize;
            let mut snippet_pre = String::from_utf16_lossy(&buffer_pre[..copied_pre]);

            // Take text after the last newline in this prefix
            if let Some(idx) = snippet_pre.rfind(['\r', '\n']) {
                snippet_pre = snippet_pre[idx + 1..].to_string();
            }
            snippet = snippet_pre;
        }

        let bookmark = Bookmark {
            position: pos,
            snippet: snippet.trim().to_string(),
            timestamp: Local::now().format("%Y-%m-%d %H:%M:%S").to_string(),
        };

        let bookmarks_window = with_state(hwnd, |state| {
            let list = state
                .bookmarks
                .files
                .entry(storage_key.clone())
                .or_default();
            list.push(bookmark);
            if persist_to_disk {
                save_bookmarks(&state.bookmarks);
            }
            state.bookmarks_window
        })
        .unwrap_or(HWND(0));

        if bookmarks_window.0 != 0 {
            app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window);
        }
        confirm_menu_action(hwnd, "insert.bookmark");
    }
}

fn audio_bookmark_position_and_snippet(position_secs: f64) -> (i32, String) {
    let current_total = position_secs.max(0.0).floor() as u64;
    let mins = current_total / 60;
    let secs = current_total % 60;
    (
        current_total as i32,
        format!("Posizione audio: {:02}:{:02}", mins, secs),
    )
}

fn goto_relative_bookmark(hwnd: HWND, forward: bool) -> bool {
    let (path, hwnd_edit, format): (Option<std::path::PathBuf>, HWND, FileFormat) = unsafe {
        with_state(hwnd, |state| {
            state
                .docs
                .get(state.current)
                .map(|doc| (doc.path.clone(), doc.hwnd_edit, doc.format))
        })
    }
    .flatten()
    .unwrap_or((None, HWND(0), FileFormat::default()));
    if hwnd_edit.0 == 0 {
        return false;
    }

    let (storage_key, _) = bookmark_storage_key(path.as_deref(), hwnd_edit);
    let Some(bookmarks) = unsafe {
        with_state(hwnd, |state| {
            state.bookmarks.files.get(&storage_key).cloned()
        })
    }
    .flatten() else {
        return false;
    };
    if bookmarks.is_empty() {
        return false;
    }

    let current_pos = if matches!(format, FileFormat::Audiobook) {
        unsafe {
            with_state(hwnd, |state| {
                let active = state.active_audiobook.as_ref()?;
                let current_path = path.as_ref()?;
                if active.path != *current_path {
                    return Some(0);
                }
                Some(crate::audio_player::audiobook_position_secs(active).floor() as i32)
            })
        }
        .flatten()
        .unwrap_or(0)
    } else {
        let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
        unsafe {
            SendMessageW(
                hwnd_edit,
                EM_EXGETSEL,
                WPARAM(0),
                LPARAM(&mut cr as *mut _ as isize),
            );
        }
        if forward { cr.cpMax } else { cr.cpMin }
    };

    let target = if forward {
        bookmarks.iter().find(|bm| bm.position > current_pos)
    } else {
        bookmarks.iter().rev().find(|bm| bm.position < current_pos)
    };
    let Some(target) = target else {
        return false;
    };
    let target_position = target.position;
    let target_snippet = target.snippet.trim().to_string();

    if matches!(format, FileFormat::Audiobook) {
        let Some(path) = path.as_ref() else {
            return false;
        };
        crate::audio_player::start_audiobook_at(hwnd, path, target_position.max(0) as u64);
        if !target_snippet.is_empty() {
            crate::accessibility::screen_reader_speak(&target_snippet);
        }
        return true;
    }

    let mut cr = CHARRANGE {
        cpMin: target_position,
        cpMax: target_position,
    };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut cr as *mut _ as isize),
        );
        SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        SetFocus(hwnd_edit);
    }
    announce_bookmark_target_line(hwnd_edit, target_position, &target_snippet);
    true
}

fn announce_bookmark_target_line(hwnd_edit: HWND, position: i32, fallback: &str) {
    let line = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_LINEFROMCHAR,
            WPARAM(position.max(0) as usize),
            LPARAM(0),
        )
        .0 as i32
    };
    if line < 0 {
        if !fallback.is_empty() {
            crate::accessibility::screen_reader_speak(fallback);
        }
        return;
    }
    let line_start =
        unsafe { SendMessageW(hwnd_edit, EM_LINEINDEX, WPARAM(line as usize), LPARAM(0)).0 as i32 };
    if line_start < 0 {
        if !fallback.is_empty() {
            crate::accessibility::screen_reader_speak(fallback);
        }
        return;
    }
    let line_len = unsafe {
        SendMessageW(
            hwnd_edit,
            EM_LINELENGTH,
            WPARAM(line_start.max(0) as usize),
            LPARAM(0),
        )
        .0 as i32
    };
    if line_len <= 0 {
        if !fallback.is_empty() {
            crate::accessibility::screen_reader_speak(fallback);
        }
        return;
    }
    let mut buf = vec![0u16; (line_len + 1) as usize];
    let mut range = TEXTRANGEW {
        chrg: CHARRANGE {
            cpMin: line_start,
            cpMax: line_start + line_len,
        },
        lpstrText: PWSTR(buf.as_mut_ptr()),
    };
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_GETTEXTRANGE,
            WPARAM(0),
            LPARAM(&mut range as *mut _ as isize),
        );
    }
    let text = String::from_utf16_lossy(&buf[..line_len as usize]);
    let trimmed = text.trim();
    if !trimmed.is_empty() {
        crate::accessibility::screen_reader_speak(trimmed);
    } else if !fallback.is_empty() {
        crate::accessibility::screen_reader_speak(fallback);
    }
}

fn clear_current_bookmarks(hwnd: HWND) -> bool {
    let (path, hwnd_edit) = unsafe {
        with_state(hwnd, |state| {
            state
                .docs
                .get(state.current)
                .map(|doc| (doc.path.clone(), doc.hwnd_edit))
        })
    }
    .flatten()
    .unwrap_or((None, HWND(0)));
    if hwnd_edit.0 == 0 {
        return false;
    }
    let (storage_key, persist_to_disk) = bookmark_storage_key(path.as_deref(), hwnd_edit);

    let (removed, bookmarks_window) = unsafe {
        with_state(hwnd, |state| {
            let removed = state.bookmarks.files.remove(&storage_key).is_some();
            if removed && persist_to_disk {
                save_bookmarks(&state.bookmarks);
            }
            (removed, state.bookmarks_window)
        })
    }
    .unwrap_or((false, HWND(0)));

    if bookmarks_window.0 != 0 {
        app_windows::bookmarks_window::refresh_bookmarks_list(bookmarks_window);
    }
    removed
}

pub(crate) fn goto_first_bookmark(
    hwnd_edit: HWND,
    path: &Path,
    bookmarks: &BookmarkStore,
    format: FileFormat,
) {
    let path_str = path.to_string_lossy().to_string();
    if let Some(list) = bookmarks.files.get(&path_str)
        && let Some(bm) = list.first()
    {
        if matches!(format, FileFormat::Audiobook) {
            // Audiobook position is handled by playback start
        } else {
            let mut cr = CHARRANGE {
                cpMin: bm.position,
                cpMax: bm.position,
            };
            unsafe {
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

pub(crate) fn rebuild_menus(hwnd: HWND) {
    let language = unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    let (_, recent_menu) = create_menus(hwnd, language);
    unsafe {
        with_state(hwnd, |state| {
            state.hmenu_recent = recent_menu;
        });
    }
    update_recent_menu(hwnd, recent_menu);
    update_voice_panel_menu_check(hwnd);
}

pub(crate) fn push_recent_file(hwnd: HWND, path: &Path) {
    let (hmenu_recent, files) = match unsafe {
        with_state(hwnd, |state| {
            state.recent_files.retain(|p| p != path);
            state.recent_files.insert(0, path.to_path_buf());
            if state.recent_files.len() > MAX_RECENT {
                state.recent_files.truncate(MAX_RECENT);
            }
            (state.hmenu_recent, state.recent_files.clone())
        })
    } {
        Some(values) => values,
        None => return,
    };
    update_recent_menu(hwnd, hmenu_recent);
    save_recent_files(&files);
}

pub(crate) fn clear_recent_files(hwnd: HWND) {
    let hmenu_recent = unsafe {
        with_state(hwnd, |state| {
            state.recent_files.clear();
            state.hmenu_recent
        })
    }
    .unwrap_or(HMENU(0));
    if hmenu_recent.0 != 0 {
        update_recent_menu(hwnd, hmenu_recent);
    }
    save_recent_files(&[]);
}

fn spawn_new_window_with_path(path: &Path) -> bool {
    let Ok(exe) = std::env::current_exe() else {
        return false;
    };
    std::process::Command::new(exe).arg(path).spawn().is_ok()
}

#[cfg(test)]
mod tests {
    use super::{audio_bookmark_position_and_snippet, clamp_tts_chunk_offset};

    #[test]
    fn audio_bookmark_position_rounds_down_and_formats() {
        let (pos, snippet) = audio_bookmark_position_and_snippet(73.9);
        assert_eq!(pos, 73);
        assert_eq!(snippet, "Posizione audio: 01:13");
    }

    #[test]
    fn audio_bookmark_position_clamps_negative() {
        let (pos, snippet) = audio_bookmark_position_and_snippet(-2.0);
        assert_eq!(pos, 0);
        assert_eq!(snippet, "Posizione audio: 00:00");
    }

    #[test]
    fn audio_bookmark_position_formats_over_an_hour() {
        let (pos, snippet) = audio_bookmark_position_and_snippet(3723.0);
        assert_eq!(pos, 3723);
        assert_eq!(snippet, "Posizione audio: 62:03");
    }

    #[test]
    fn tts_chunk_offset_stays_monotonic() {
        assert_eq!(clamp_tts_chunk_offset(120, 140), 140);
        assert_eq!(clamp_tts_chunk_offset(120, 120), 120);
        assert_eq!(clamp_tts_chunk_offset(120, 90), 120);
    }

    #[test]
    fn tts_chunk_offset_clamps_negative_input() {
        assert_eq!(clamp_tts_chunk_offset(0, -7), 0);
        assert_eq!(clamp_tts_chunk_offset(25, -1), 25);
    }
}

fn open_document_with_encoding(hwnd: HWND, path: &Path, encoding: Option<TextEncoding>) {
    let behavior = unsafe { with_state(hwnd, |state| state.settings.open_behavior) }
        .unwrap_or(OpenBehavior::NewTab);
    if behavior == OpenBehavior::NewWindow && spawn_new_window_with_path(path) {
        return;
    }
    editor_manager::open_document_with_encoding(hwnd, path, encoding);
}

fn play_audio_playlist_item(hwnd: HWND, index: usize) {
    let path = unsafe {
        with_state(hwnd, |state| {
            if index >= state.audio_playlist.len() {
                return None;
            }
            state.audio_playlist_index = Some(index);
            state.audio_ffmpeg_retry_for = None;
            Some(state.audio_playlist[index].clone())
        })
        .flatten()
    };
    let Some(path) = path else {
        return;
    };

    let tab_index = editor_manager::ensure_audio_document_tab(hwnd, &path);
    if let Some(tab_index) = tab_index {
        editor_manager::select_tab(hwnd, tab_index);
    }
    audio_player::stop_audiobook_playback(hwnd);
    audio_player::start_audiobook_playback(hwnd, &path);
}

pub(crate) fn queue_audio_files_and_play(hwnd: HWND, paths: Vec<PathBuf>) {
    unsafe {
        if paths.is_empty() {
            return;
        }
        for path in &paths {
            if editor_manager::ensure_audio_document_tab(hwnd, path).is_none() {
                crate::log_debug(&format!(
                    "Audio player: failed to ensure audio tab for {}",
                    path.display()
                ));
            }
            push_recent_file(hwnd, path);
        }
        let target_index = with_state(hwnd, |state| {
            let play_path = paths[0].clone();
            if state.audio_playlist.is_empty() {
                for path in &paths {
                    if !state.audio_playlist.iter().any(|p| p == path) {
                        state.audio_playlist.push(path.clone());
                    }
                }
                let idx = state
                    .audio_playlist
                    .iter()
                    .position(|p| p == &play_path)
                    .unwrap_or(0);
                state.audio_playlist_index = Some(idx);
                return idx;
            }
            let current = state
                .audio_playlist_index
                .filter(|idx| *idx < state.audio_playlist.len())
                .unwrap_or(0);
            let mut insert_at = current.saturating_add(1);
            for path in &paths {
                if state.audio_playlist.iter().any(|p| p == path) {
                    continue;
                }
                state.audio_playlist.insert(insert_at, path.clone());
                insert_at = insert_at.saturating_add(1);
            }
            let target = state
                .audio_playlist
                .iter()
                .position(|p| p == &play_path)
                .unwrap_or(current);
            state.audio_playlist_index = Some(target);
            target
        })
        .unwrap_or(0);

        play_audio_playlist_item(hwnd, target_index);
    }
}

fn switch_audio_playlist_relative(hwnd: HWND, delta: i32) -> bool {
    let target = unsafe {
        with_state(hwnd, |state| {
            if state.audio_playlist.is_empty() {
                return None;
            }
            let current = state.audio_playlist_index?;
            let next = if delta > 0 {
                current.checked_add(delta as usize)?
            } else {
                current.checked_sub(delta.unsigned_abs() as usize)?
            };
            if next >= state.audio_playlist.len() {
                return None;
            }
            Some(next)
        })
        .flatten()
    };
    let Some(target) = target else {
        return false;
    };
    play_audio_playlist_item(hwnd, target);
    true
}

fn handle_audio_playlist_timer(hwnd: HWND) {
    let (is_paused, should_advance, current_seconds, elapsed_since_start) = unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.is_paused {
                return Some((true, false, 0_u64, std::time::Duration::from_secs(0)));
            }
            let current = audio_player::audiobook_position_secs(player)
                .max(0.0)
                .floor() as u64;
            let near_end_player = player
                .duration_secs()
                .map(|d| current.saturating_add(1) >= d.max(0.0).floor() as u64)
                .unwrap_or(false);
            let near_end_file = audio_player::audiobook_duration_secs(&player.path)
                .map(|duration| current.saturating_add(1) >= duration)
                .unwrap_or(false);
            // Treat as ended if either live player duration or file duration confirms near-end.
            let should_advance = near_end_player || near_end_file;
            Some((
                false,
                should_advance,
                current,
                player.start_instant.elapsed(),
            ))
        })
        .flatten()
        .unwrap_or((false, false, 0_u64, std::time::Duration::from_secs(0)))
    };
    if is_paused {
        return;
    }
    let output_stopped = audio_player::audiobook_output_stopped(hwnd).unwrap_or(false);
    if !should_advance && !output_stopped {
        return;
    }
    // Retry with forced FFmpeg streaming only in the startup window.
    // This avoids collisions with manual seeks/skips during normal playback.
    if !should_advance
        && output_stopped
        && current_seconds == 0
        && elapsed_since_start.as_secs() <= 5
        && audio_player::retry_current_with_ffmpeg_stream(hwnd)
    {
        return;
    }
    // In the first seconds after startup/resume FFmpeg streaming may transiently report
    // "stopped" before stable output; avoid stopping the playlist too early.
    if !should_advance && output_stopped && elapsed_since_start.as_secs() <= 2 {
        return;
    }
    // If output stops mid-file (e.g. transient backend failure), recover by restarting
    // from the current position instead of stopping playback entirely.
    if !should_advance && output_stopped {
        let is_stream_cache = unsafe {
            with_state(hwnd, |state| {
                state
                    .active_audiobook
                    .as_ref()
                    .map(|player| is_stream_cache_media(&player.path))
                    .unwrap_or(false)
            })
            .unwrap_or(false)
        };
        if is_stream_cache {
            log_debug("Audio player: stream cache output stopped, skipping mid-file auto-restart");
        } else {
            let restart = unsafe {
                with_state(hwnd, |state| {
                    let player = state.active_audiobook.as_ref()?;
                    let position_secs = audio_player::audiobook_position_secs(player)
                        .max(0.0)
                        .floor() as u64;
                    Some((player.path.clone(), position_secs))
                })
                .flatten()
            };
            if let Some((path, position_secs)) = restart {
                audio_player::start_audiobook_at(hwnd, &path, position_secs);
                return;
            }
        }
    }
    if !switch_audio_playlist_relative(hwnd, 1) {
        // Keep the player instance alive when no next track is available.
        // This matches non-hard-stop behavior (seek back and resume works).
    }
}

fn open_path_with_behavior(hwnd: HWND, path: &Path) {
    if is_audio_path(path) {
        queue_audio_files_and_play(hwnd, vec![path.to_path_buf()]);
        return;
    }
    open_document_with_encoding(hwnd, path, None);
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

pub(crate) fn open_pdf_document_async(hwnd: HWND, path: &Path, from_copydata: bool) {
    unsafe {
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
            editor_manager::set_edit_text(hwnd_edit, &pdf_loading_placeholder(0, language));
            let doc = Document {
                title: title.clone(),
                path: Some(path_buf.clone()),
                hwnd_edit,
                dirty: false,
                format: FileFormat::Pdf,
                opened_text_encoding: None,
                current_save_text_encoding: None,
                from_rss: false,
                is_temporary: false,
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

        let ocr_timeout_secs = if from_copydata {
            PDF_OCR_PROMPT_TIMEOUT_COPYDATA_SECS
        } else {
            0
        };
        start_pdf_loading_animation(hwnd, hwnd_edit, ocr_timeout_secs);

        let hwnd_main = hwnd;
        std::thread::spawn(move || {
            let result = match std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                read_pdf_text_with_status(&path_buf, language)
            })) {
                Ok(result) => result,
                Err(panic) => {
                    let panic_msg = if let Some(s) = panic.downcast_ref::<&str>() {
                        (*s).to_string()
                    } else if let Some(s) = panic.downcast_ref::<String>() {
                        s.clone()
                    } else {
                        "unknown panic payload".to_string()
                    };
                    crate::log_debug(&format!("PDF load thread panic caught: {}", panic_msg));
                    Err(crate::i18n::tr_f(
                        language,
                        "file_handler.pdf_read_error",
                        &[("err", "PDF extraction crashed unexpectedly")],
                    ))
                }
            };
            let payload = Box::new(PdfLoadResult {
                hwnd_edit,
                path: path_buf,
                result,
                from_copydata,
            });
            let payload_ptr = Box::into_raw(payload);
            if let Err(e) = PostMessageW(
                hwnd_main,
                WM_PDF_LOADED,
                WPARAM(0),
                LPARAM(payload_ptr as isize),
            ) {
                crate::log_debug(&format!("Failed to post WM_PDF_LOADED: {}", e));
                let _unused_box = Box::from_raw(payload_ptr);
            }
        });
    }
}

fn start_ocr_for_pdf(
    hwnd: HWND,
    hwnd_edit: HWND,
    path: PathBuf,
    from_copydata: bool,
    language: Language,
) {
    start_pdf_loading_animation(hwnd, hwnd_edit, 0);
    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let final_result = match std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            win_ocr::recognize_text_from_pdf(&path, language)
        })) {
            Ok(ocr_result) => match ocr_result {
                Ok(text) => Ok(PdfTextResult::Text(text)),
                Err(e) => Err(e),
            },
            Err(panic) => {
                let panic_msg = if let Some(s) = panic.downcast_ref::<&str>() {
                    (*s).to_string()
                } else if let Some(s) = panic.downcast_ref::<String>() {
                    s.clone()
                } else {
                    "unknown panic payload".to_string()
                };
                crate::log_debug(&format!("PDF OCR thread panic caught: {}", panic_msg));
                Err(crate::i18n::tr_f(
                    language,
                    "file_handler.pdf_read_error",
                    &[("err", "PDF OCR crashed unexpectedly")],
                ))
            }
        };
        let payload = Box::new(PdfLoadResult {
            hwnd_edit,
            path,
            result: final_result,
            from_copydata,
        });
        unsafe {
            let payload_ptr = Box::into_raw(payload);
            if let Err(e) = PostMessageW(
                hwnd_main,
                WM_PDF_LOADED,
                WPARAM(0),
                LPARAM(payload_ptr as isize),
            ) {
                crate::log_debug(&format!("Failed to post WM_PDF_LOADED (OCR): {}", e));
                let _unused_box = Box::from_raw(payload_ptr);
            }
        }
    });
}

fn handle_pdf_loaded(hwnd: HWND, payload: PdfLoadResult) {
    unsafe {
        let PdfLoadResult {
            hwnd_edit,
            path,
            result,
            from_copydata,
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
            Ok(PdfTextResult::Text(text)) => {
                editor_manager::set_edit_text(hwnd_edit, &text);
                with_state(hwnd, |state| {
                    goto_first_bookmark(hwnd_edit, &path, &state.bookmarks, FileFormat::Pdf);
                });
                let msg = pdf_loaded_message(language);
                crate::log_debug(&format!("Info (speech): {msg}"));
                crate::accessibility::nvda_speak(&msg);
                crate::log_if_err!(MessageBeep(MB_ICONASTERISK));
                let mut update_title = false;
                with_state(hwnd, |state| {
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
            Ok(PdfTextResult::NoText) => {
                start_ocr_for_pdf(hwnd, hwnd_edit, path, from_copydata, language);
            }
            Err(message) => {
                // Instead of closing the document, show error message as placeholder text
                let error_placeholder = format!(
                    "{}\n\n{}",
                    message,
                    i18n::tr(language, "app.pdf_error_hint")
                );
                editor_manager::set_edit_text(hwnd_edit, &error_placeholder);
                show_error(hwnd, language, &message);
                let mut update_title = false;
                with_state(hwnd, |state| {
                    if let Some(doc) = state.docs.get_mut(index) {
                        doc.dirty = false;
                        update_tab_title(state.hwnd_tab, index, &doc.title, false);
                        update_title = state.current == index;
                    }
                });
                if update_title {
                    update_window_title(hwnd);
                }
            }
        }
    }
}

fn handle_document_loaded(hwnd: HWND, payload: editor_manager::DocumentLoadResult) {
    unsafe {
        let editor_manager::DocumentLoadResult { path, result } = payload;
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        const LARGE_FILE_NO_WRAP_THRESHOLD_BYTES: u64 = 15 * 1024 * 1024;

        let loaded = match result {
            Ok(loaded) => loaded,
            Err(message) => {
                show_error(hwnd, language, &message);
                return;
            }
        };
        let Some(loaded) = loaded else {
            return;
        };
        let large_file_no_wrap = std::fs::metadata(&path)
            .map(|meta| meta.len() >= LARGE_FILE_NO_WRAP_THRESHOLD_BYTES)
            .unwrap_or(false);

        let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
        let (hwnd_edit, new_index) = with_state(hwnd, |state| {
            let use_word_wrap = state.settings.word_wrap && !large_file_no_wrap;
            let hwnd_edit = editor_manager::create_edit(
                hwnd,
                state.hfont,
                use_word_wrap,
                state.settings.text_color,
                state.settings.text_size,
            );
            editor_manager::set_edit_text(hwnd_edit, &loaded.content);

            let doc = Document {
                title: title.to_string(),
                path: Some(path.clone()),
                hwnd_edit,
                dirty: false,
                format: loaded.format,
                opened_text_encoding: loaded.opened_text_encoding,
                current_save_text_encoding: None,
                from_rss: false,
                is_temporary: false,
            };
            state.docs.push(doc);
            if large_file_no_wrap {
                state.large_text_editors.insert(hwnd_edit.0);
            } else {
                state.large_text_editors.remove(&hwnd_edit.0);
            }
            insert_tab(state.hwnd_tab, title, (state.docs.len() - 1) as i32);
            goto_first_bookmark(hwnd_edit, &path, &state.bookmarks, loaded.format);
            (hwnd_edit, state.docs.len() - 1)
        })
        .unwrap_or((HWND(0), 0));

        if hwnd_edit.0 == 0 {
            return;
        }
        if large_file_no_wrap {
            log_debug(&format!(
                "Large file mode: disabled word wrap for '{}' (>= {} bytes)",
                path.display(),
                LARGE_FILE_NO_WRAP_THRESHOLD_BYTES
            ));
        }

        editor_manager::select_tab(hwnd, new_index);
        push_recent_file(hwnd, &path);

        let pending = with_state(hwnd, |state| state.pending_find_in_files.clone()).unwrap_or(None);
        if let Some(pending) = pending
            && pending.path == path
        {
            apply_find_in_files_selection(
                hwnd_edit,
                &pending.snippet,
                &pending.term,
                pending.start_utf16,
                pending.len_utf16,
            );
            with_state(hwnd, |state| {
                state.pending_find_in_files = None;
            });
        }
    }
}

fn start_pdf_loading_animation(hwnd: HWND, hwnd_edit: HWND, ocr_timeout_secs: u64) {
    unsafe {
        let timer_id = with_state(hwnd, |state| {
            let timer_id = state.next_timer_id;
            state.next_timer_id = state.next_timer_id.saturating_add(1);
            state.pdf_loading.push(PdfLoadingState {
                hwnd_edit,
                timer_id,
                frame: 0,
                start_time: Instant::now(),
                ocr_timeout_secs,
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
}

fn stop_pdf_loading_animation(hwnd: HWND, hwnd_edit: HWND) {
    unsafe {
        let mut timer_id = None;
        with_state(hwnd, |state| {
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
            kill_timer_best_effort(hwnd, timer_id, "KillTimer PDF_LOADING");
        }
    }
}

fn handle_pdf_loading_timer(hwnd: HWND, timer_id: usize) {
    unsafe {
        let mut target = None;
        let mut should_timeout = false;
        with_state(hwnd, |state| {
            if let Some(entry) = state
                .pdf_loading
                .iter_mut()
                .find(|entry| entry.timer_id == timer_id)
            {
                entry.frame = entry.frame.wrapping_add(1);
                let timeout_secs = entry.ocr_timeout_secs;
                if timeout_secs == 0 || entry.start_time.elapsed().as_secs() >= timeout_secs {
                    should_timeout = true;
                }
                target = Some((entry.hwnd_edit, entry.frame, timeout_secs));
            }
        });

        if should_timeout && let Some((hwnd_edit, _, timeout_secs)) = target {
            stop_pdf_loading_animation(hwnd, hwnd_edit);
            let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
            let from_copydata = timeout_secs != 0;
            let path = with_state(hwnd, |state| {
                state
                    .docs
                    .iter()
                    .find(|d| d.hwnd_edit == hwnd_edit)
                    .and_then(|d| d.path.clone())
            })
            .flatten();

            if let Some(path) = path {
                start_ocr_for_pdf(hwnd, hwnd_edit, path, from_copydata, language);
            } else {
                let err = i18n::tr(language, "app.pdf_error_hint");
                editor_manager::set_edit_text(hwnd_edit, &err);
            }
            return;
        }

        if let Some((hwnd_edit, frame, _)) = target {
            let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
            editor_manager::set_edit_text(hwnd_edit, &pdf_loading_placeholder(frame, language));
        }
    }
}

pub(crate) fn pdf_loading_placeholder(frame: usize, language: crate::settings::Language) -> String {
    let spinner = ['|', '/', '-', '\\'][frame % 4];
    let bar_width = 24;
    let filled = frame % (bar_width + 1);
    let bar = format!(
        "{}{}",
        "#".repeat(filled),
        "-".repeat(bar_width.saturating_sub(filled))
    );
    let loading = i18n::tr(language, "app.pdf_loading");
    let analyzing = i18n::tr(language, "app.pdf_analyzing");
    format!("{loading}\r\n\r\n[{bar}]\r\n{analyzing} {spinner}")
}

fn handle_drop_files(hwnd: HWND, hdrop: HDROP) {
    let count = unsafe { DragQueryFileW(hdrop, 0xFFFFFFFF, None) };
    let mut dropped_paths = Vec::new();
    for index in 0..count {
        let mut buffer = [0u16; 260];
        let len = unsafe { DragQueryFileW(hdrop, index, Some(&mut buffer)) };
        if len == 0 {
            continue;
        }
        let path = PathBuf::from(String::from_utf16_lossy(&buffer[..len as usize]));
        if path.as_os_str().is_empty() {
            continue;
        }
        dropped_paths.push(path);
    }
    if !dropped_paths.is_empty() && dropped_paths.iter().all(|path| is_audio_path(path)) {
        queue_audio_files_and_play(hwnd, dropped_paths);
    } else {
        for path in dropped_paths {
            open_path_with_behavior(hwnd, &path);
        }
    }
    unsafe {
        DragFinish(hdrop);
    }
}

fn next_tab_with_prompt(hwnd: HWND) {
    let (current, count) = match unsafe {
        with_state(hwnd, |state| {
            if state.docs.is_empty() {
                return None;
            }
            let current = state.current;
            Some((current, state.docs.len()))
        })
    } {
        Some(Some(values)) => values,
        _ => return,
    };
    if count <= 1 {
        return;
    }
    let next = (current + 1) % count;
    select_tab(hwnd, next);
}

fn open_documents_popup(hwnd: HWND) {
    let docs = unsafe {
        with_state(hwnd, |state| {
            state
                .docs
                .iter()
                .enumerate()
                .map(|(idx, doc)| (idx, doc.title.clone()))
                .collect::<Vec<_>>()
        })
    }
    .unwrap_or_default();
    if docs.len() <= 1 {
        return;
    }

    let menu = unsafe { CreatePopupMenu().unwrap_or(HMENU(0)) };
    if menu.0 == 0 {
        return;
    }

    for (idx, title) in docs.iter().take(200) {
        let id = (WINDOW_DOC_MENU_BASE + *idx) as u32;
        let display = if title.trim().is_empty() {
            format!("Documento {}", idx + 1)
        } else {
            title.clone()
        };
        let label = format!("&{} {}", idx + 1, display);
        unsafe {
            crate::log_if_err!(AppendMenuW(
                menu,
                MF_STRING,
                id as usize,
                PCWSTR(to_wide(&label).as_ptr()),
            ));
        }
    }

    let mut pt = POINT::default();
    if unsafe { GetCursorPos(&mut pt).is_err() } {
        return;
    }
    let command = unsafe {
        TrackPopupMenu(
            menu,
            TPM_RIGHTBUTTON | TPM_RETURNCMD,
            pt.x,
            pt.y,
            0,
            hwnd,
            None,
        )
    };
    if let Some(index) = window_doc_menu_index_from_command(command.0 as usize) {
        select_tab(hwnd, index);
    }
}

fn window_doc_menu_index_from_command(cmd_id: usize) -> Option<usize> {
    if (WINDOW_DOC_MENU_BASE..(WINDOW_DOC_MENU_BASE + WINDOW_DOC_MENU_MAX)).contains(&cmd_id) {
        Some(cmd_id - WINDOW_DOC_MENU_BASE)
    } else {
        None
    }
}

fn refresh_window_open_documents_menu(hwnd: HWND, window_menu: HMENU) {
    unsafe {
        crate::log_if_err!(DeleteMenu(
            window_menu,
            WINDOW_DOC_MENU_SEPARATOR_ID as u32,
            MF_BYCOMMAND
        ));
    }
    for idx in 0..WINDOW_DOC_MENU_MAX {
        let cmd_id = (WINDOW_DOC_MENU_BASE + idx) as u32;
        unsafe {
            crate::log_if_err!(DeleteMenu(window_menu, cmd_id, MF_BYCOMMAND));
        }
    }

    let docs = unsafe {
        with_state(hwnd, |state| {
            state
                .docs
                .iter()
                .enumerate()
                .map(|(idx, doc)| (idx, doc.title.clone()))
                .collect::<Vec<_>>()
        })
    }
    .unwrap_or_default();
    if docs.len() <= 1 {
        return;
    }

    unsafe {
        crate::log_if_err!(AppendMenuW(
            window_menu,
            MF_SEPARATOR,
            WINDOW_DOC_MENU_SEPARATOR_ID,
            PCWSTR::null()
        ));
    }
    for (idx, title) in docs.iter().take(WINDOW_DOC_MENU_MAX) {
        let display = if title.trim().is_empty() {
            format!("Documento {}", idx + 1)
        } else {
            title.clone()
        };
        let label = format!("&{} {}", idx + 1, display);
        unsafe {
            crate::log_if_err!(AppendMenuW(
                window_menu,
                MF_STRING,
                WINDOW_DOC_MENU_BASE + idx,
                PCWSTR(to_wide(&label).as_ptr()),
            ));
        }
    }
}

fn now_ms() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis() as u64
}

fn suggest_extension_for_interpreter(interpreter: &str) -> &'static str {
    let lower = interpreter.to_lowercase();
    if lower.contains("python") || lower.ends_with("py.exe") || lower.ends_with("py") {
        "py"
    } else if lower.contains("java") {
        "java"
    } else if lower.contains("node") {
        "js"
    } else if lower.contains("ruby") {
        "rb"
    } else if lower.contains("perl") {
        "pl"
    } else if lower.contains("php") {
        "php"
    } else if lower.contains("lua") {
        "lua"
    } else if lower.contains("htm") || lower.contains("html") {
        "html"
    } else if lower.contains("powershell") || lower.contains("pwsh") {
        "ps1"
    } else if lower.contains("bash") || lower.contains("sh.exe") || lower == "sh" {
        "sh"
    } else {
        "txt"
    }
}

fn execute_current_file(hwnd: HWND) {
    let (path, content, interpreter) = match unsafe {
        with_state(hwnd, |state| {
            let doc = state.docs.get(state.current)?;
            let path = doc.path.clone();
            let hwnd_edit = doc.hwnd_edit;
            let content = editor_manager::get_edit_text(hwnd_edit);
            let interpreter = state.settings.interpreter_path.clone();
            Some((path, content, interpreter))
        })
    } {
        Some(Some(v)) => v,
        _ => return,
    };

    let (exec_path, working_dir) = if let Some(p) = path {
        let dir = p.parent().map(|d| d.to_path_buf());
        (p, dir)
    } else {
        // Not saved: create temp file
        let ext = suggest_extension_for_interpreter(&interpreter);
        let temp_dir = std::env::temp_dir();
        let temp_file = temp_dir.join(format!("sonarpad_exec_{}.{}", now_ms(), ext));
        if let Err(e) = std::fs::write(&temp_file, content) {
            log_debug(&format!("Failed to write temp file for execution: {}", e));
            return;
        }
        (temp_file, Some(temp_dir))
    };

    let extension = exec_path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    if extension == "html" || extension == "htm" {
        let res = unsafe {
            ShellExecuteW(
                hwnd,
                w!("open"),
                PCWSTR(to_wide(&exec_path.to_string_lossy()).as_ptr()),
                PCWSTR::null(),
                PCWSTR::null(),
                SW_SHOWNORMAL,
            )
        };
        if res.0 as isize <= 32 {
            crate::log_debug(&format!("ShellExecuteW failed to open HTML: {}", res.0));
        }
        return;
    }

    let command = format!("\"{}\" \"{}\"", interpreter, exec_path.to_string_lossy());
    app_windows::prompt_window::open_with_command(hwnd, Some(command), working_dir);
}

fn attempt_switch_to_selected_tab(hwnd: HWND) {
    let (current, hwnd_tab, count) = match unsafe {
        with_state(hwnd, |state| {
            if state.docs.is_empty() {
                return None;
            }
            let current = state.current;
            Some((current, state.hwnd_tab, state.docs.len()))
        })
    } {
        Some(Some(values)) => values,
        _ => return,
    };
    let sel = unsafe { SendMessageW(hwnd_tab, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32 };
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
    let mut cleaned = out.trim().trim_end_matches(['.', ' ']).to_string();
    if cleaned.is_empty() {
        return cleaned;
    }
    if cleaned.len() > 120 {
        let mut idx = 120;
        while idx > 0 && !cleaned.is_char_boundary(idx) {
            idx -= 1;
        }
        cleaned.truncate(idx);
    }
    if is_reserved_filename(&cleaned) {
        cleaned.push('_');
    }
    cleaned
}

fn is_reserved_filename(name: &str) -> bool {
    let upper = name.trim_end_matches(['.', ' ']).to_ascii_uppercase();
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

pub(crate) struct SaveAudioDialogResult {
    pub path: PathBuf,
    pub create_parts_folder: bool,
}

pub(crate) fn save_audio_dialog(
    hwnd: HWND,
    suggested_name: Option<&str>,
    show_split_folder_option: bool,
) -> Option<SaveAudioDialogResult> {
    unsafe {
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        let initial_dir_setting =
            with_state(hwnd, |state| state.settings.audiobook_save_folder.clone())
                .unwrap_or_default();
        let initial_dir = initial_dir_setting.trim().to_string();
        if !initial_dir.is_empty() {
            crate::log_if_err!(std::fs::create_dir_all(&initial_dir));
        }
        let pfd: IFileSaveDialog = CoCreateInstance(&FileSaveDialog, None, CLSCTX_ALL).ok()?;

        let filter_raw = i18n::tr(language, "dialog.save_audio_filter");
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
        pfd.SetFileTypeIndex(1).ok()?;
        pfd.SetDefaultExtension(w!("mp3")).ok()?;
        pfd.SetTitle(PCWSTR(
            to_wide(&i18n::tr(language, "dialog.save_audio_title")).as_ptr(),
        ))
        .ok()?;

        if !initial_dir.is_empty() {
            let initial_dir_w = to_wide(&initial_dir);
            if let Ok(shell_folder) = SHCreateItemFromParsingName::<_, _, IShellItem>(
                PCWSTR(initial_dir_w.as_ptr()),
                None,
            ) {
                let _unused = pfd.SetDefaultFolder(&shell_folder);
                let _unused = pfd.SetFolder(&shell_folder);
            }
        }

        if let Some(name) = suggested_name {
            let default_name = Path::new(name)
                .file_name()
                .and_then(|n| n.to_str())
                .filter(|n| !n.trim().is_empty())
                .unwrap_or(name);
            pfd.SetFileName(PCWSTR(to_wide(default_name).as_ptr()))
                .ok()?;
        }

        let current_bitrate =
            with_state(hwnd, |state| state.settings.audiobook_m4b_bitrate).unwrap_or(128);
        let bitrate_options = [64u32, 80, 96, 128, 160, 192, 256];
        let initial_bitrate = if bitrate_options.contains(&current_bitrate) {
            current_bitrate
        } else {
            128
        };
        let selected_bitrate = Arc::new(Mutex::new(initial_bitrate));

        const AUDIO_SAVE_BITRATE_BUTTON_ID: u32 = 201;
        const AUDIO_SAVE_SPLIT_FOLDER_CHECKBOX_ID: u32 = 202;
        let pfdc: IFileDialogCustomize = pfd.cast().ok()?;
        pfdc.AddPushButton(
            AUDIO_SAVE_BITRATE_BUTTON_ID,
            PCWSTR(to_wide(&audiobook_bitrate_button_label(language, initial_bitrate)).as_ptr()),
        )
        .ok()?;
        if show_split_folder_option {
            pfdc.AddCheckButton(
                AUDIO_SAVE_SPLIT_FOLDER_CHECKBOX_ID,
                PCWSTR(to_wide(&i18n::tr(language, "dialog.save_audio_split_folder")).as_ptr()),
                BOOL(1),
            )
            .ok()?;
        }

        let handler: IFileDialogEvents = AudiobookBitrateDialogHandler {
            parent: hwnd,
            language,
            selected_bitrate: selected_bitrate.clone(),
            allowed_bitrates: bitrate_options.to_vec(),
        }
        .into();
        let cookie = pfd.Advise(&handler).ok()?;

        let show_ok = pfd.Show(hwnd).is_ok();
        pfd.Unadvise(cookie).ok()?;

        if show_ok {
            let item = pfd.GetResult().ok()?;
            let path_ptr = item
                .GetDisplayName(windows::Win32::UI::Shell::SIGDN_FILESYSPATH)
                .ok()?;
            let path_str = path_ptr.to_string().unwrap_or_default();
            CoTaskMemFree(Some(path_ptr.0 as *const _));

            let mut path = PathBuf::from(path_str);
            let has_valid_ext = path
                .extension()
                .and_then(|e| e.to_str())
                .map(str::trim)
                .is_some_and(|e| !e.is_empty());
            if !has_valid_ext {
                let filter_index = pfd.GetFileTypeIndex().ok().unwrap_or(1);
                if filter_index == 2 {
                    path.set_extension("m4b");
                } else {
                    path.set_extension("mp3");
                }
            }

            let selected_bitrate = selected_bitrate
                .lock()
                .map(|v| *v)
                .unwrap_or(current_bitrate.clamp(64, 256));
            log_debug(&format!(
                "Save audio dialog: selected bitrate {} kbps (current {} kbps)",
                selected_bitrate, current_bitrate
            ));
            if let Some(settings) = with_state(hwnd, |state| {
                state.settings.audiobook_m4b_bitrate = selected_bitrate;
                state.settings.clone()
            }) {
                save_settings(settings);
            } else {
                log_debug("Failed to access settings for audiobook bitrate");
            }

            let create_parts_folder = if show_split_folder_option {
                pfdc.GetCheckButtonState(AUDIO_SAVE_SPLIT_FOLDER_CHECKBOX_ID)
                    .ok()
                    .is_none_or(|state| state.as_bool())
            } else {
                false
            };

            Some(SaveAudioDialogResult {
                path,
                create_parts_folder,
            })
        } else {
            None
        }
    }
}

fn audiobook_bitrate_button_label(language: Language, bitrate_kbps: u32) -> String {
    let bitrate_label = i18n::tr(language, "podcast.bitrate");
    format!("{bitrate_label} ({bitrate_kbps} kbps)")
}

#[implement(IFileDialogEvents, IFileDialogControlEvents)]
struct AudiobookBitrateDialogHandler {
    parent: HWND,
    language: Language,
    selected_bitrate: Arc<Mutex<u32>>,
    allowed_bitrates: Vec<u32>,
}

impl IFileDialogEvents_Impl for AudiobookBitrateDialogHandler {
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
    fn OnTypeChange(&self, _pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
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

impl IFileDialogControlEvents_Impl for AudiobookBitrateDialogHandler {
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
        pfdc: Option<&IFileDialogCustomize>,
        dwidctl: u32,
    ) -> windows::core::Result<()> {
        const AUDIO_SAVE_BITRATE_BUTTON_ID: u32 = 201;
        if dwidctl != AUDIO_SAVE_BITRATE_BUTTON_ID {
            return Ok(());
        }
        let current = self
            .selected_bitrate
            .lock()
            .map(|v| *v)
            .unwrap_or(128)
            .clamp(64, 256);
        let menu = unsafe { CreatePopupMenu().unwrap_or(HMENU(0)) };
        if menu.0 == 0 {
            return Ok(());
        }
        let mut ids = Vec::new();
        for (i, bitrate) in self.allowed_bitrates.iter().enumerate() {
            let id = 10_000u32 + i as u32;
            ids.push((id, *bitrate));
            let text = if *bitrate == current {
                format!("* {bitrate} kbps")
            } else {
                format!("{bitrate} kbps")
            };
            unsafe {
                crate::log_if_err!(AppendMenuW(
                    menu,
                    MF_STRING,
                    id as usize,
                    PCWSTR(to_wide(&text).as_ptr()),
                ));
            }
        }
        let mut pt = POINT::default();
        if unsafe { GetCursorPos(&mut pt).is_err() } {
            return Ok(());
        }
        let owner = {
            let fg = unsafe { GetForegroundWindow() };
            if fg.0 != 0 { fg } else { self.parent }
        };
        let command = unsafe {
            TrackPopupMenu(
                menu,
                TPM_RIGHTBUTTON | TPM_RETURNCMD,
                pt.x,
                pt.y,
                0,
                owner,
                None,
            )
        };
        let chosen = ids
            .iter()
            .find(|(id, _)| *id == command.0 as u32)
            .map(|(_, bitrate)| *bitrate);
        if let Some(next) = chosen {
            if let Ok(mut guard) = self.selected_bitrate.lock() {
                *guard = next;
            }
            if let Some(pfdc) = pfdc {
                unsafe {
                    let label = audiobook_bitrate_button_label(self.language, next);
                    pfdc.SetControlLabel(dwidctl, PCWSTR(to_wide(&label).as_ptr()))
                        .ok();
                }
            }
        }
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
/// Mostra dialog per salvare file diagnostici zip.
fn export_diagnostics_dialog(hwnd: HWND) {
    unsafe {
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

        // Genera nome file con timestamp
        let timestamp = chrono::Local::now().format("%Y%m%d_%H%M%S");
        let default_name = format!("sonarpad_diagnostics_{}.zip", timestamp);
        let mut default_wide = to_wide(&default_name);
        default_wide.resize(260, 0);

        let zip_archive_label = i18n::tr(language, "dialog.zip_archive");
        let all_files_label = i18n::tr(language, "dialog.all_files");
        let filter = to_wide(&format!(
            "{} (*.zip)\0*.zip\0{} (*.*)\0*.*\0\0",
            zip_archive_label, all_files_label
        ));
        let title = to_wide(&i18n::tr(language, "dialog.export_diagnostics_title"));

        let mut ofn = OPENFILENAMEW {
            lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
            hwndOwner: hwnd,
            lpstrFile: PWSTR(default_wide.as_mut_ptr()),
            nMaxFile: default_wide.len() as u32,
            lpstrFilter: PCWSTR(filter.as_ptr()),
            lpstrTitle: PCWSTR(title.as_ptr()),
            lpstrDefExt: PCWSTR(to_wide("zip").as_ptr()),
            nFilterIndex: 1,
            Flags: OFN_EXPLORER | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST,
            ..Default::default()
        };

        if !GetSaveFileNameW(&mut ofn).as_bool() {
            return;
        }

        let len = default_wide
            .iter()
            .position(|&c| c == 0)
            .unwrap_or(default_wide.len());
        let path = PathBuf::from(String::from_utf16_lossy(&default_wide[..len]));

        match diagnostics::export_diagnostics_zip(&path) {
            Ok(()) => {
                let message = i18n::tr(language, "dialog.export_diagnostics_success");
                show_info(hwnd, language, &message);
            }
            Err(e) => {
                let message = format!(
                    "{}: {}",
                    i18n::tr(language, "dialog.export_diagnostics_error"),
                    e
                );
                show_error(hwnd, language, &message);
            }
        }
    }
}

pub(crate) unsafe fn show_error(hwnd: HWND, language: Language, message: &str) {
    show_error_with_id(hwnd, language, message, None);
}

pub(crate) unsafe fn show_error_with_id(
    hwnd: HWND,
    language: Language,
    message: &str,
    event_id: Option<&sentry_integration::EventId>,
) {
    let full_message = format!(
        "{}{}",
        message,
        sentry_integration::format_event_id(event_id)
    );
    log_debug(&format!("Error shown: {full_message}"));
    let wide = to_wide(&full_message);
    let title = to_wide(&error_title(language));
    watchdog::enter_modal_dialog();
    MessageBoxW(
        hwnd,
        PCWSTR(wide.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONERROR,
    );
    watchdog::exit_modal_dialog();
}

pub(crate) unsafe fn show_info(hwnd: HWND, language: Language, message: &str) {
    log_debug(&format!("Info shown: {message}"));
    let wide = to_wide(message);
    let title = to_wide(&info_title(language));
    watchdog::enter_modal_dialog();
    MessageBoxW(
        hwnd,
        PCWSTR(wide.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONINFORMATION,
    );
    watchdog::exit_modal_dialog();
}

/// Mostra un MessageBox generico sospendendo il watchdog durante la visualizzazione.
/// Usare per i MessageBox diretti che non passano da show_error/show_info.
pub(crate) unsafe fn message_box_modal(
    hwnd: HWND,
    message: PCWSTR,
    title: PCWSTR,
    flags: MESSAGEBOX_STYLE,
) -> MESSAGEBOX_RESULT {
    watchdog::enter_modal_dialog();
    let result = MessageBoxW(hwnd, message, title, flags);
    watchdog::exit_modal_dialog();
    result
}

pub(crate) fn recent_store_path() -> Option<PathBuf> {
    let mut path = settings::settings_dir();
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
        crate::log_if_err!(std::fs::create_dir_all(parent));
    }
    let store = RecentFileStore {
        files: files
            .iter()
            .map(|p| p.to_string_lossy().to_string())
            .collect(),
    };
    if let Ok(json) = serde_json::to_string_pretty(&store) {
        crate::log_if_err!(std::fs::write(path, json));
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
            let Some(pfd) = pfd else {
                return Ok(());
            };
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
) -> Option<Vec<(PathBuf, Option<TextEncoding>)>> {
    log_debug("open_file_dialog_with_encoding called");
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

    use windows::Win32::UI::Shell::FileOpenDialog;
    use windows::Win32::UI::Shell::IFileOpenDialog;
    use windows::Win32::UI::Shell::{FILEOPENDIALOGOPTIONS, FOS_ALLOWMULTISELECT};

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
    let mut options = pfd.GetOptions().ok()?;
    options |= FILEOPENDIALOGOPTIONS(FOS_ALLOWMULTISELECT.0);
    pfd.SetOptions(options).ok()?;

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
    crate::log_if_err!(pfd.SetFileTypeIndex(1));

    log_debug("Showing open dialog");
    if pfd.Show(hwnd).is_ok() {
        log_debug("Dialog closed with OK");
        let selected_encoding_idx = pfdc.GetSelectedControlItem(101).ok()?;
        let filter_index = pfd.GetFileTypeIndex().ok()?;
        let manual_encoding = if filter_index == 2 {
            Some(index_to_encoding(selected_encoding_idx))
        } else {
            None
        };
        let mut out = Vec::new();
        if let Ok(items) = pfd.GetResults()
            && let Ok(count) = items.GetCount()
        {
            for i in 0..count {
                if let Ok(item) = items.GetItemAt(i)
                    && let Ok(path_ptr) =
                        item.GetDisplayName(windows::Win32::UI::Shell::SIGDN_FILESYSPATH)
                {
                    let path_str = path_ptr.to_string().unwrap_or_default();
                    CoTaskMemFree(Some(path_ptr.0 as *const _));
                    if !path_str.is_empty() {
                        out.push((PathBuf::from(path_str), None));
                    }
                }
            }
        }

        if out.is_empty() {
            let item = pfd.GetResult().ok()?;
            let path_ptr = item
                .GetDisplayName(windows::Win32::UI::Shell::SIGDN_FILESYSPATH)
                .ok()?;
            let path_str = path_ptr.to_string().unwrap_or_default();
            CoTaskMemFree(Some(path_ptr.0 as *const _));
            if path_str.is_empty() {
                pfd.Unadvise(cookie).ok()?;
                return None;
            }
            out.push((PathBuf::from(path_str), manual_encoding));
        } else if out.len() == 1 {
            out[0].1 = manual_encoding;
        }

        pfd.Unadvise(cookie).ok()?;
        Some(out)
    } else {
        pfd.Unadvise(cookie).ok()?;
        None
    }
}

pub(crate) fn open_subtitle_file_dialog(hwnd: HWND) -> Option<PathBuf> {
    let language = unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    let title = i18n::tr(language, "dialog.open_subtitles_title");
    let filter_label = i18n::tr(language, "dialog.subtitles_filter");
    let all_files_label = i18n::tr(language, "dialog.all_files");

    let pattern = crate::subtitles::SUBTITLE_EXTENSIONS
        .iter()
        .map(|ext| format!("*.{ext}"))
        .collect::<Vec<_>>()
        .join(";");

    let filter = format!(
        "{} ({})\0{}\0{} (*.*)\0*.*\0",
        filter_label, pattern, pattern, all_files_label
    );
    let mut filter_wide: Vec<u16> = filter.encode_utf16().collect();
    filter_wide.push(0);

    let mut buffer = [0u16; 1024];
    let title_wide = to_wide(&title);
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter_wide.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        lpstrTitle: PCWSTR(title_wide.as_ptr()),
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        ..Default::default()
    };

    if !unsafe { GetOpenFileNameW(&mut ofn) }.as_bool() {
        return None;
    }
    let len = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
    if len == 0 {
        return None;
    }
    Some(PathBuf::from(String::from_utf16_lossy(&buffer[..len])))
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
    let initial_dir = settings::default_documents_save_folder();
    crate::log_if_err!(std::fs::create_dir_all(&initial_dir));
    let initial_dir_w = to_wide(&initial_dir);
    if let Ok(shell_folder) =
        SHCreateItemFromParsingName::<_, _, IShellItem>(PCWSTR(initial_dir_w.as_ptr()), None)
    {
        let _unused = pfd.SetDefaultFolder(&shell_folder);
        let _unused = pfd.SetFolder(&shell_folder);
    }

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
    crate::log_if_err!(pfd.SetFileTypeIndex(1));

    if pfd.Show(hwnd).is_ok() {
        let item = pfd.GetResult().ok()?;
        let path_ptr = item
            .GetDisplayName(windows::Win32::UI::Shell::SIGDN_FILESYSPATH)
            .ok()?;
        let path_str = path_ptr.to_string().unwrap_or_default();
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
