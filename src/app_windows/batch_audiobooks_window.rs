use std::collections::VecDeque;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::sync::{Arc, Mutex};
use std::time::Duration;

use chrono::Local;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_ALLOWMULTISELECT, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY,
    OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{
    BST_CHECKED, LVCOLUMNW, LVIF_TEXT, LVITEMW, LVM_DELETEALLITEMS, LVM_DELETEITEM,
    LVM_GETITEMCOUNT, LVM_GETNEXTITEM, LVM_INSERTCOLUMNW, LVM_INSERTITEMW,
    LVM_SETEXTENDEDLISTVIEWSTYLE, LVM_SETITEMTEXTW, LVNI_SELECTED, LVS_EX_FULLROWSELECT,
    LVS_EX_GRIDLINES, LVS_REPORT, LVS_SHOWSELALWAYS, PBM_SETPOS, PBM_SETRANGE, PROGRESS_CLASSW,
    WC_BUTTON, WC_COMBOBOXW, WC_EDIT, WC_STATIC,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetKeyState, SetFocus, VK_ESCAPE, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL, CB_SETCURSEL, CBS_DROPDOWNLIST,
    CREATESTRUCTW, CreateWindowExW, DefWindowProcW, DestroyWindow, ES_AUTOVSCROLL, ES_MULTILINE,
    GetNextDlgTabItem, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW,
    IsWindow, KillTimer, LoadCursorW, MSG, PostMessageW, RegisterClassW, SendMessageW,
    SetForegroundWindow, SetTimer, SetWindowLongPtrW, SetWindowTextW, WINDOW_STYLE, WM_APP,
    WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WM_TIMER,
    WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME,
    WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, PWSTR, w};

use crate::accessibility::{EM_REPLACESEL, ES_READONLY, to_wide, to_wide_normalized};
use crate::app_windows::find_in_files_window::browse_for_folder;
use crate::file_handler::{
    decode_text, is_audio_path, is_doc_path, is_docx_path, is_epub_path, is_html_path, is_pdf_path,
    is_ppt_path, is_pptx_path, is_spreadsheet_path, read_doc_text, read_docx_text,
    read_epub_chapters, read_epub_text, read_html_text, read_pdf_text, read_ppt_text,
    read_spreadsheet_text,
};
use crate::i18n;
use crate::settings::{DictionaryEntry, Language, TtsEngine};
use crate::tts_engine::{
    MixedAudiobookConfig, TtsChunk, build_audiobook_parts_by_positions,
    build_audiobook_parts_from_sections, build_mixed_audiobook_parts_by_positions,
    build_mixed_audiobook_parts_from_sections, collect_marker_entries, has_voice_tags,
    prepare_tts_text, render_mixed_audiobook_part, run_tts_audiobook_part, split_into_tts_chunks,
    split_text, strip_dashed_lines,
};
use crate::{log_debug, sanitize_filename, show_error, with_state};

const BATCH_CLASS_NAME: &str = "SonarpadBatchAudiobooks";

const BATCH_ID_LIST: usize = 9401;
const BATCH_ID_ADD_FILES: usize = 9402;
const BATCH_ID_ADD_FOLDER: usize = 9403;
const BATCH_ID_REMOVE: usize = 9404;
const BATCH_ID_CLEAR: usize = 9405;
const BATCH_ID_OUTPUT_EDIT: usize = 9406;
const BATCH_ID_OUTPUT_BROWSE: usize = 9407;
const BATCH_ID_FORMAT: usize = 9408;
const BATCH_ID_SUBFOLDER: usize = 9409;
const BATCH_ID_AVOID_OVERWRITE: usize = 9410;
const BATCH_ID_START: usize = 9411;
const BATCH_ID_CANCEL: usize = 9412;
const BATCH_ID_CLOSE: usize = 9413;
const BATCH_ID_LOG: usize = 9414;
const BATCH_ID_PROGRESS: usize = 9415;

const WM_BATCH_EVENT: u32 = WM_APP + 120;
const BATCH_TIMER_ID: usize = 1;
const EM_LIMITTEXT: u32 = 0x00C5;

const EM_SETSEL: u32 = 0x00B1;

#[repr(u32)]
#[derive(Clone, Copy, PartialEq, Eq)]
enum BatchStatusCode {
    Pending = 0,
    Running = 1,
    Done = 2,
    Failed = 3,
    Canceled = 4,
}

pub fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN {
        let vk = msg.wParam.0 as u32;
        if vk == VK_ESCAPE.0 as u32 {
            if let Err(e) = unsafe { PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) } {
                crate::log_debug(&format!("Failed to post WM_CLOSE: {}", e));
            }
            return true;
        }
        if vk == VK_TAB.0 as u32 {
            let log_edit = with_batch_state(hwnd, |state| state.log_edit);
            let Some(log_edit) = log_edit else {
                crate::log_debug("Failed to access batch state in handle_navigation");
                return false;
            };
            if msg.hwnd == log_edit {
                let shift_down =
                    (unsafe { GetKeyState(VK_SHIFT.0 as i32) } & (0x8000u16 as i16)) != 0;
                let next = unsafe { GetNextDlgTabItem(hwnd, log_edit, shift_down) };
                if next.0 != 0 {
                    unsafe { SetFocus(next) };
                    return true;
                }
            }
        }
    }
    false
}

struct BatchItem {
    input: PathBuf,
    status: BatchStatusCode,
    output: String,
}

struct BatchLabels {
    title: String,
    col_input: String,
    col_status: String,
    col_output: String,
    add_files: String,
    add_folder: String,
    remove_selected: String,
    clear: String,
    output_folder: String,
    browse: String,
    format: String,
    format_mp3: String,
    format_wav: String,
    option_subfolder: String,
    option_avoid_overwrite: String,
    start: String,
    cancel: String,
    close: String,
    log_label: String,
    progress_label: String,
    status_pending: String,
    status_running: String,
    status_done: String,
    status_failed: String,
    status_canceled: String,
    empty_queue: String,
    no_output_folder: String,
    invalid_output_folder: String,
    wav_not_supported: String,
    log_start: String,
    log_cancel_requested: String,
}

struct BatchState {
    hwnd: HWND,
    parent: HWND,
    list: HWND,
    add_files: HWND,
    add_folder: HWND,
    remove: HWND,
    clear: HWND,
    output_edit: HWND,
    browse: HWND,
    format_combo: HWND,
    checkbox_subfolder: HWND,
    checkbox_avoid_overwrite: HWND,
    start_button: HWND,
    cancel_button: HWND,
    close_button: HWND,
    log_edit: HWND,
    progress_label: HWND,
    progress_bar: HWND,
    language: Language,
    items: Vec<BatchItem>,
    running: bool,
    cancel_flag: Option<Arc<AtomicBool>>,
    message_queue: Arc<Mutex<VecDeque<BatchMessage>>>,
}

struct StatusUpdate {
    index: usize,
    status: BatchStatusCode,
    output: String,
}

struct ProgressUpdate {
    completed: usize,
    total: usize,
}

struct DoneUpdate {
    report_path: Option<PathBuf>,
}

enum BatchMessage {
    Log(String),
    Status(StatusUpdate),
    Progress(ProgressUpdate),
    Done(DoneUpdate),
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum AudioFormat {
    Mp3,
    Wav,
}

struct BatchSettings {
    output_folder: PathBuf,
    format: AudioFormat,
    create_subfolder: bool,
    avoid_overwrite: bool,
}

struct TtsSettings {
    voice: String,
    split_on_newline: bool,
    audiobook_split: u32,
    audiobook_split_by_text: bool,
    audiobook_split_text: String,
    audiobook_split_text_requires_newline: bool,
    audiobook_split_by_epub_chapter: bool,
    audiobook_split_by_time: bool,
    audiobook_split_minutes: u32,
    audiobook_split_start_number: u32,
    audiobook_m4b_bitrate: u32,
    tts_engine: TtsEngine,
    dictionary: Vec<DictionaryEntry>,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    language: Language,
}

struct BatchResultItem {
    input: PathBuf,
    status: BatchStatusCode,
    outputs: Vec<PathBuf>,
    error: Option<String>,
}

fn labels(language: Language) -> BatchLabels {
    BatchLabels {
        title: i18n::tr(language, "batch_audiobooks.title"),
        col_input: i18n::tr(language, "batch_audiobooks.col.input"),
        col_status: i18n::tr(language, "batch_audiobooks.col.status"),
        col_output: i18n::tr(language, "batch_audiobooks.col.output"),
        add_files: i18n::tr(language, "batch_audiobooks.add_files"),
        add_folder: i18n::tr(language, "batch_audiobooks.add_folder"),
        remove_selected: i18n::tr(language, "batch_audiobooks.remove_selected"),
        clear: i18n::tr(language, "batch_audiobooks.clear"),
        output_folder: i18n::tr(language, "batch_audiobooks.output_folder"),
        browse: i18n::tr(language, "batch_audiobooks.browse"),
        format: i18n::tr(language, "batch_audiobooks.format"),
        format_mp3: i18n::tr(language, "batch_audiobooks.format.mp3"),
        format_wav: i18n::tr(language, "batch_audiobooks.format.wav"),
        option_subfolder: i18n::tr(language, "batch_audiobooks.option_subfolder"),
        option_avoid_overwrite: i18n::tr(language, "batch_audiobooks.option_avoid_overwrite"),
        start: i18n::tr(language, "batch_audiobooks.start"),
        cancel: i18n::tr(language, "batch_audiobooks.cancel"),
        close: i18n::tr(language, "batch_audiobooks.close"),
        log_label: i18n::tr(language, "batch_audiobooks.log_label"),
        progress_label: i18n::tr(language, "batch_audiobooks.progress_label"),
        status_pending: i18n::tr(language, "batch_audiobooks.status.pending"),
        status_running: i18n::tr(language, "batch_audiobooks.status.running"),
        status_done: i18n::tr(language, "batch_audiobooks.status.done"),
        status_failed: i18n::tr(language, "batch_audiobooks.status.failed"),
        status_canceled: i18n::tr(language, "batch_audiobooks.status.canceled"),
        empty_queue: i18n::tr(language, "batch_audiobooks.error.empty_queue"),
        no_output_folder: i18n::tr(language, "batch_audiobooks.error.no_output_folder"),
        invalid_output_folder: i18n::tr(language, "batch_audiobooks.error.invalid_output_folder"),
        wav_not_supported: i18n::tr(language, "batch_audiobooks.error.wav_not_supported"),
        log_start: i18n::tr(language, "batch_audiobooks.log.start"),
        log_cancel_requested: i18n::tr(language, "batch_audiobooks.log.cancel_requested"),
    }
}

fn status_text(labels: &BatchLabels, status: BatchStatusCode) -> &str {
    match status {
        BatchStatusCode::Pending => &labels.status_pending,
        BatchStatusCode::Running => &labels.status_running,
        BatchStatusCode::Done => &labels.status_done,
        BatchStatusCode::Failed => &labels.status_failed,
        BatchStatusCode::Canceled => &labels.status_canceled,
    }
}

pub fn open(parent: HWND) {
    let existing =
        unsafe { with_state(parent, |state| state.batch_audiobooks_window).unwrap_or(HWND(0)) };
    if existing.0 != 0 {
        if unsafe { !IsWindow(existing).as_bool() } {
            let result = unsafe {
                with_state(parent, |state| {
                    state.batch_audiobooks_window = HWND(0);
                })
            };
            if result.is_none() {
                crate::log_debug("Failed to access batch window state at L264");
            }
        } else {
            unsafe {
                SetForegroundWindow(existing);
            }
            return;
        }
    }

    let language =
        unsafe { with_state(parent, |state| state.settings.language) }.unwrap_or_default();
    let class_name = to_wide(BATCH_CLASS_NAME);
    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(batch_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let labels = labels(language);
    let title = to_wide(&labels.title);

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            120,
            120,
            760,
            620,
            parent,
            HMENU(0),
            hinstance,
            None,
        )
    };
    if hwnd.0 == 0 {
        return;
    }

    if unsafe {
        with_state(parent, |state| {
            state.batch_audiobooks_window = hwnd;
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access batch window state at L316");
    }
}

unsafe extern "system" fn batch_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let result = std::panic::catch_unwind(|| match msg {
        WM_CREATE => {
            let cs = &*(lparam.0 as *const CREATESTRUCTW);
            let parent = cs.hwndParent;
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = labels(language);
            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

            let list = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("SysListView32"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(LVS_REPORT | LVS_SHOWSELALWAYS),
                16,
                16,
                720,
                190,
                hwnd,
                HMENU(BATCH_ID_LIST as isize),
                HINSTANCE(0),
                None,
            );
            SendMessageW(
                list,
                LVM_SETEXTENDEDLISTVIEWSTYLE,
                WPARAM(0),
                LPARAM((LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES) as isize),
            );

            insert_column(list, 0, &labels.col_input, 360);
            insert_column(list, 1, &labels.col_status, 120);
            insert_column(list, 2, &labels.col_output, 220);

            let add_files = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.add_files).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                16,
                214,
                120,
                26,
                hwnd,
                HMENU(BATCH_ID_ADD_FILES as isize),
                HINSTANCE(0),
                None,
            );
            let add_folder = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.add_folder).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                144,
                214,
                120,
                26,
                hwnd,
                HMENU(BATCH_ID_ADD_FOLDER as isize),
                HINSTANCE(0),
                None,
            );
            let remove = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.remove_selected).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                272,
                214,
                140,
                26,
                hwnd,
                HMENU(BATCH_ID_REMOVE as isize),
                HINSTANCE(0),
                None,
            );
            let clear = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.clear).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                420,
                214,
                100,
                26,
                hwnd,
                HMENU(BATCH_ID_CLEAR as isize),
                HINSTANCE(0),
                None,
            );

            let output_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.output_folder).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                246,
                200,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let output_edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                16,
                268,
                540,
                24,
                hwnd,
                HMENU(BATCH_ID_OUTPUT_EDIT as isize),
                HINSTANCE(0),
                None,
            );
            let browse = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.browse).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                564,
                268,
                172,
                24,
                hwnd,
                HMENU(BATCH_ID_OUTPUT_BROWSE as isize),
                HINSTANCE(0),
                None,
            );

            let format_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.format).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                300,
                80,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let format_combo = CreateWindowExW(
                Default::default(),
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                96,
                296,
                120,
                200,
                hwnd,
                HMENU(BATCH_ID_FORMAT as isize),
                HINSTANCE(0),
                None,
            );
            SendMessageW(
                format_combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&labels.format_mp3).as_ptr() as isize),
            );
            SendMessageW(
                format_combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(&labels.format_wav).as_ptr() as isize),
            );
            SendMessageW(format_combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));

            let checkbox_subfolder = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.option_subfolder).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                16,
                328,
                300,
                22,
                hwnd,
                HMENU(BATCH_ID_SUBFOLDER as isize),
                HINSTANCE(0),
                None,
            );
            let checkbox_avoid_overwrite = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.option_avoid_overwrite).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                16,
                352,
                320,
                22,
                hwnd,
                HMENU(BATCH_ID_AVOID_OVERWRITE as isize),
                HINSTANCE(0),
                None,
            );
            SendMessageW(
                checkbox_avoid_overwrite,
                windows::Win32::UI::WindowsAndMessaging::BM_SETCHECK,
                WPARAM(BST_CHECKED.0 as usize),
                LPARAM(0),
            );

            let progress_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.progress_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                380,
                260,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let progress_bar = CreateWindowExW(
                Default::default(),
                PROGRESS_CLASSW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                16,
                402,
                720,
                18,
                hwnd,
                HMENU(BATCH_ID_PROGRESS as isize),
                HINSTANCE(0),
                None,
            );
            SendMessageW(
                progress_bar,
                PBM_SETRANGE,
                WPARAM(0),
                LPARAM(((0u16 as u32) | ((100u16 as u32) << 16)) as isize),
            );

            let start_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.start).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                16,
                426,
                120,
                28,
                hwnd,
                HMENU(BATCH_ID_START as isize),
                HINSTANCE(0),
                None,
            );
            let cancel_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.cancel).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                144,
                426,
                120,
                28,
                hwnd,
                HMENU(BATCH_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );
            let close_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.close).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                616,
                426,
                120,
                28,
                hwnd,
                HMENU(BATCH_ID_CLOSE as isize),
                HINSTANCE(0),
                None,
            );
            EnableWindow(cancel_button, false);

            let log_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.log_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                462,
                200,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let log_edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WINDOW_STYLE(ES_READONLY | (ES_MULTILINE as u32) | (ES_AUTOVSCROLL as u32)),
                16,
                484,
                720,
                112,
                hwnd,
                HMENU(BATCH_ID_LOG as isize),
                HINSTANCE(0),
                None,
            );
            SendMessageW(log_edit, EM_LIMITTEXT, WPARAM(0x7FFF_FFFEusize), LPARAM(0));

            for control in [
                list,
                add_files,
                add_folder,
                remove,
                clear,
                output_label,
                output_edit,
                browse,
                format_label,
                format_combo,
                checkbox_subfolder,
                checkbox_avoid_overwrite,
                progress_label,
                progress_bar,
                start_button,
                cancel_button,
                close_button,
                log_label,
                log_edit,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let initial_folder = std::env::current_dir()
                .unwrap_or_else(|_| PathBuf::from("."))
                .to_string_lossy()
                .to_string();
            if let Err(e) = SetWindowTextW(output_edit, PCWSTR(to_wide(&initial_folder).as_ptr())) {
                crate::log_debug(&format!("Failed to set initial output folder: {}", e));
            }
            SetFocus(list);

            let message_queue = Arc::new(Mutex::new(VecDeque::new()));

            let state = Box::new(BatchState {
                hwnd,
                parent,
                list,
                add_files,
                add_folder,
                remove,
                clear,
                output_edit,
                browse,
                format_combo,
                checkbox_subfolder,
                checkbox_avoid_overwrite,
                start_button,
                cancel_button,
                close_button,
                log_edit,
                progress_label,
                progress_bar,
                language,
                items: Vec::new(),
                running: false,
                cancel_flag: None,
                message_queue,
            });
            SetWindowLongPtrW(
                hwnd,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                Box::into_raw(state) as isize,
            );
            if SetTimer(hwnd, BATCH_TIMER_ID, 200, None) == 0 {
                crate::log_debug("Failed to set BATCH_TIMER");
            }

            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            match cmd_id {
                BATCH_ID_ADD_FILES => {
                    if with_batch_state(hwnd, |state| {
                        add_files_to_queue(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_ADD_FOLDER => {
                    if with_batch_state(hwnd, |state| {
                        add_folder_to_queue(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_REMOVE => {
                    if with_batch_state(hwnd, |state| {
                        remove_selected_items(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_CLEAR => {
                    if with_batch_state(hwnd, |state| {
                        clear_items(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_OUTPUT_BROWSE => {
                    if with_batch_state(hwnd, |state| {
                        if let Some(folder) = browse_for_folder(hwnd, state.language) {
                            let wide = to_wide(folder.to_string_lossy().as_ref());
                            if let Err(e) = SetWindowTextW(state.output_edit, PCWSTR(wide.as_ptr()))
                            {
                                crate::log_debug(&format!("Failed to set output folder: {}", e));
                            }
                        }
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_START => {
                    if with_batch_state(hwnd, |state| {
                        start_batch(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_CANCEL => {
                    if with_batch_state(hwnd, |state| {
                        request_cancel(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access batch state");
                    }
                    LRESULT(0)
                }
                BATCH_ID_CLOSE => {
                    if let Err(e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                        crate::log_debug(&format!(
                            "Failed to post WM_CLOSE on cancellation: {}",
                            e
                        ));
                    }
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            let allow_close = with_batch_state(hwnd, |state| {
                if state.running {
                    request_cancel(state);
                    false
                } else {
                    true
                }
            })
            .unwrap_or(true);
            if allow_close {
                unsafe {
                    crate::log_if_err!(DestroyWindow(hwnd));
                }
            }
            LRESULT(0)
        }
        WM_BATCH_EVENT => {
            handle_batch_messages(hwnd);
            LRESULT(0)
        }
        WM_TIMER => {
            if wparam.0 == BATCH_TIMER_ID {
                handle_batch_messages(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_DESTROY => {
            log_debug("Batch WM_DESTROY");
            if let Err(e) = KillTimer(hwnd, BATCH_TIMER_ID) {
                crate::log_debug(&format!("Failed to kill BATCH_TIMER: {}", e));
            }
            if with_batch_state(hwnd, |state| {
                state.running = false;
                state.cancel_flag = None;
            })
            .is_none()
            {
                crate::log_debug("Failed to access batch state");
            }
            unsafe {
                if with_state(
                    windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd),
                    |state| {
                        state.batch_audiobooks_window = HWND(0);
                    },
                )
                .is_none()
                {
                    crate::log_debug("Failed to access parent state in WM_DESTROY");
                }
            }
            let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            crate::focus_editor(parent);
            LRESULT(0)
        }
        WM_NCDESTROY => {
            log_debug("Batch WM_NCDESTROY");
            let ptr = unsafe {
                GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                    as *mut BatchState
            };
            if !ptr.is_null() {
                unsafe {
                    SetWindowLongPtrW(
                        hwnd,
                        windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                        0,
                    );
                    let _unused_box = Box::from_raw(ptr);
                }
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    });
    match result {
        Ok(res) => res,
        Err(_) => {
            log_debug("Batch window panic in wndproc.");
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
    }
}

fn with_batch_state<T>(hwnd: HWND, f: impl FnOnce(&mut BatchState) -> T) -> Option<T> {
    let ptr = unsafe {
        GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
            as *mut BatchState
    };
    if ptr.is_null() {
        None
    } else {
        Some(unsafe { f(&mut *ptr) })
    }
}

fn insert_column(list: HWND, index: i32, text: &str, width: i32) {
    let mut wide = to_wide(text);
    let column = LVCOLUMNW {
        mask: windows::Win32::UI::Controls::LVCF_TEXT
            | windows::Win32::UI::Controls::LVCF_WIDTH
            | windows::Win32::UI::Controls::LVCF_SUBITEM,
        pszText: PWSTR(wide.as_mut_ptr()),
        cchTextMax: (wide.len().saturating_sub(1)) as i32,
        cx: width,
        iSubItem: index,
        ..Default::default()
    };
    unsafe {
        SendMessageW(
            list,
            LVM_INSERTCOLUMNW,
            WPARAM(index as usize),
            LPARAM(&column as *const _ as isize),
        );
    }
}

fn insert_list_item(list: HWND, index: i32, input: &str, status: &str, output: &str) {
    let mut input_wide = to_wide(input);
    let item = LVITEMW {
        mask: LVIF_TEXT,
        iItem: index,
        iSubItem: 0,
        pszText: PWSTR(input_wide.as_mut_ptr()),
        ..Default::default()
    };
    unsafe {
        SendMessageW(
            list,
            LVM_INSERTITEMW,
            WPARAM(0),
            LPARAM(&item as *const _ as isize),
        );
    }
    set_list_subitem(list, index, 1, status);
    set_list_subitem(list, index, 2, output);
}

fn set_list_subitem(list: HWND, index: i32, subitem: i32, text: &str) {
    let mut wide = to_wide(text);
    let item = LVITEMW {
        iSubItem: subitem,
        pszText: PWSTR(wide.as_mut_ptr()),
        ..Default::default()
    };
    unsafe {
        SendMessageW(
            list,
            LVM_SETITEMTEXTW,
            WPARAM(index as usize),
            LPARAM(&item as *const _ as isize),
        );
    }
}

fn add_files_to_queue(state: &mut BatchState) {
    if state.running {
        return;
    }
    let Some(paths) = open_files_dialog(state.hwnd, state.language) else {
        return;
    };
    for path in paths {
        push_queue_item(state, path);
    }
    update_progress(state, 0, state.items.len());
}

fn add_folder_to_queue(state: &mut BatchState) {
    if state.running {
        return;
    }
    let Some(folder) = browse_for_folder(state.hwnd, state.language) else {
        return;
    };
    for path in collect_folder_files(&folder) {
        push_queue_item(state, path);
    }
    update_progress(state, 0, state.items.len());
}

fn push_queue_item(state: &mut BatchState, path: PathBuf) {
    if path.as_os_str().is_empty() {
        return;
    }
    if !is_supported_batch_input_path(&path) {
        crate::log_debug(&format!(
            "Batch: skipping unsupported or temp input: {}",
            path.display()
        ));
        return;
    }
    if state.items.iter().any(|item| item.input == path) {
        return;
    }
    let labels = labels(state.language);
    let idx = state.items.len() as i32;
    let display = path.to_string_lossy().to_string();
    insert_list_item(state.list, idx, &display, &labels.status_pending, "");
    state.items.push(BatchItem {
        input: path,
        status: BatchStatusCode::Pending,
        output: String::new(),
    });
}

fn remove_selected_items(state: &mut BatchState) {
    if state.running {
        return;
    }
    let mut indices = Vec::new();
    let mut current = -1i32;
    loop {
        let next = unsafe {
            SendMessageW(
                state.list,
                LVM_GETNEXTITEM,
                WPARAM(current as isize as usize),
                LPARAM(LVNI_SELECTED as isize),
            )
            .0 as i32
        };
        if next == -1 {
            break;
        }
        indices.push(next as usize);
        current = next;
    }
    if indices.is_empty() {
        return;
    }
    indices.sort_unstable_by(|a, b| b.cmp(a));
    for idx in indices {
        if idx < state.items.len() {
            state.items.remove(idx);
            unsafe {
                SendMessageW(state.list, LVM_DELETEITEM, WPARAM(idx), LPARAM(0));
            }
        }
    }
    update_progress(state, 0, state.items.len());
}

fn clear_items(state: &mut BatchState) {
    if state.running {
        return;
    }
    state.items.clear();
    unsafe {
        SendMessageW(state.list, LVM_DELETEALLITEMS, WPARAM(0), LPARAM(0));
    }
    update_progress(state, 0, 0);
}

fn start_batch(state: &mut BatchState) {
    if state.running {
        return;
    }
    let labels = labels(state.language);
    if state.items.is_empty() {
        unsafe {
            show_error(state.hwnd, state.language, &labels.empty_queue);
        }
        return;
    }
    let output_folder = read_control_text(state.output_edit);
    if output_folder.trim().is_empty() {
        unsafe {
            show_error(state.hwnd, state.language, &labels.no_output_folder);
        }
        return;
    }
    let output_folder = PathBuf::from(output_folder.trim());
    if !output_folder.exists()
        && let Err(err) = std::fs::create_dir_all(&output_folder)
    {
        unsafe {
            show_error(
                state.hwnd,
                state.language,
                &format!("{}: {}", labels.invalid_output_folder, err),
            );
        }
        return;
    }
    let format_sel =
        unsafe { SendMessageW(state.format_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)) }.0 as i32;
    let format = if format_sel == 1 {
        AudioFormat::Wav
    } else {
        AudioFormat::Mp3
    };
    let (tts_engine, tts_voice) = unsafe {
        with_state(state.parent, |app| {
            (app.settings.tts_engine, app.settings.tts_voice.clone())
        })
        .unwrap_or((TtsEngine::Edge, "it-IT-IsabellaNeural".to_string()))
    };
    if tts_engine == TtsEngine::Edge && format == AudioFormat::Wav {
        unsafe {
            show_error(state.hwnd, state.language, &labels.wav_not_supported);
        }
        return;
    }

    let create_subfolder = is_checked(state.checkbox_subfolder);
    let avoid_overwrite = is_checked(state.checkbox_avoid_overwrite);

    let tts_settings = load_tts_settings(state.parent, tts_voice, state.language);
    let batch_settings = BatchSettings {
        output_folder,
        format,
        create_subfolder,
        avoid_overwrite,
    };

    set_running(state, true);
    append_log(state, &labels.log_start);
    unsafe {
        SetFocus(state.log_edit);
    }

    let items = state.items.iter().map(|item| item.input.clone()).collect();
    let hwnd = state.hwnd;
    let message_queue = state.message_queue.clone();
    let cancel_flag = Arc::new(AtomicBool::new(false));
    state.cancel_flag = Some(cancel_flag.clone());

    std::thread::spawn(move || {
        run_batch(
            hwnd,
            items,
            batch_settings,
            tts_settings,
            cancel_flag,
            message_queue,
        );
    });
}

fn request_cancel(state: &mut BatchState) {
    if !state.running {
        return;
    }
    if let Some(cancel) = &state.cancel_flag {
        cancel.store(true, Ordering::SeqCst);
    }
    let labels = labels(state.language);
    append_log(state, &labels.log_cancel_requested);
    unsafe {
        EnableWindow(state.cancel_button, false);
    }
}

fn set_running(state: &mut BatchState, running: bool) {
    state.running = running;
    unsafe {
        EnableWindow(state.add_files, !running);
        EnableWindow(state.add_folder, !running);
        EnableWindow(state.remove, !running);
        EnableWindow(state.clear, !running);
        EnableWindow(state.output_edit, !running);
        EnableWindow(state.browse, !running);
        EnableWindow(state.format_combo, !running);
        EnableWindow(state.checkbox_subfolder, !running);
        EnableWindow(state.checkbox_avoid_overwrite, !running);
        EnableWindow(state.start_button, !running);
        EnableWindow(state.cancel_button, running);
        EnableWindow(state.close_button, !running);
    }
}

fn finish_batch(state: &mut BatchState, _report_path: Option<&PathBuf>) -> (HWND, Language, HWND) {
    set_running(state, false);
    state.cancel_flag = None;
    (state.parent, state.language, state.list)
}

fn update_item_status(state: &mut BatchState, index: usize, status: BatchStatusCode, output: &str) {
    let labels = labels(state.language);
    if index >= state.items.len() {
        return;
    }
    if let Some(item) = state.items.get_mut(index) {
        item.status = status;
        item.output = output.to_string();
    }
    if unsafe { !IsWindow(state.list).as_bool() } {
        return;
    }
    let count =
        unsafe { SendMessageW(state.list, LVM_GETITEMCOUNT, WPARAM(0), LPARAM(0)) }.0 as usize;
    if index >= count {
        return;
    }
    set_list_subitem(state.list, index as i32, 1, status_text(&labels, status));
    set_list_subitem(state.list, index as i32, 2, output);
}

fn handle_batch_messages(hwnd: HWND) {
    let mut messages: Vec<BatchMessage> = Vec::new();
    let mut done_dialog: Option<(HWND, Language, HWND)> = None;
    let mut done_report: Option<PathBuf> = None;
    if with_batch_state(hwnd, |state| {
        if let Ok(mut queue) = state.message_queue.lock() {
            while let Some(msg) = queue.pop_front() {
                messages.push(msg);
            }
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access batch state");
    }
    if messages.is_empty() {
        return;
    }
    for msg in &messages {
        if let BatchMessage::Done(update) = msg {
            done_report = update.report_path.clone();
            break;
        }
    }
    if with_batch_state(hwnd, |state| {
        if done_report.is_some() {
            done_dialog = Some(finish_batch(state, done_report.as_ref()));
            return;
        }
        for msg in messages {
            match msg {
                BatchMessage::Log(line) => append_log(state, &line),
                BatchMessage::Status(update) => {
                    update_item_status(state, update.index, update.status, &update.output);
                }
                BatchMessage::Progress(update) => {
                    update_progress(state, update.completed, update.total);
                }
                BatchMessage::Done(update) => {
                    done_dialog = Some(finish_batch(state, update.report_path.as_ref()));
                }
            }
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access batch state");
    }
    if let Some((parent, language, list)) = done_dialog {
        if unsafe { !IsWindow(parent).as_bool() } {
            log_debug("Batch finished: parent window invalid, skipping dialog.");
            return;
        }
        log_debug("Batch finished. Showing completion dialog.");
        unsafe {
            crate::show_info(
                parent,
                language,
                &i18n::tr(language, "batch_audiobooks.done"),
            );
        }
        if unsafe { IsWindow(list).as_bool() } {
            unsafe {
                SetFocus(list);
            }
        }
        log_debug("Batch completion dialog closed.");
    }
}

fn update_progress(state: &mut BatchState, completed: usize, total: usize) {
    let percent = if total == 0 {
        0
    } else {
        ((completed as f32 / total as f32) * 100.0).round() as u32
    };
    if is_window_valid(state.progress_bar, "progress_bar") {
        unsafe {
            SendMessageW(
                state.progress_bar,
                PBM_SETPOS,
                WPARAM(percent as usize),
                LPARAM(0),
            );
        }
    }
    let label = i18n::tr_f(
        state.language,
        "batch_audiobooks.progress_text",
        &[
            ("done", &completed.to_string()),
            ("total", &total.to_string()),
        ],
    );
    if is_window_valid(state.progress_label, "progress_label") {
        let wide = to_wide(&label);
        unsafe {
            if let Err(e) = SetWindowTextW(state.progress_label, PCWSTR(wide.as_ptr())) {
                crate::log_debug(&format!("Failed to set progress label: {}", e));
            }
        }
    }
}

fn append_log(state: &mut BatchState, line: &str) {
    if !is_window_valid(state.log_edit, "log_edit") {
        return;
    }
    let mut text = line.to_string();
    if !text.ends_with('\n') {
        text.push('\n');
    }
    let wide = to_wide_normalized(&text);
    unsafe {
        let len = GetWindowTextLengthW(state.log_edit);
        let len = if len < 0 { 0 } else { len };
        SendMessageW(
            state.log_edit,
            EM_SETSEL,
            WPARAM(len as usize),
            LPARAM(len as isize),
        );
        SendMessageW(
            state.log_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(wide.as_ptr() as isize),
        );
    }
}

fn is_window_valid(hwnd: HWND, label: &str) -> bool {
    let ok = unsafe { IsWindow(hwnd).as_bool() };
    if !ok {
        log_debug(&format!("Batch: {label} window invalid. hwnd={}", hwnd.0));
    }
    ok
}

fn open_files_dialog(hwnd: HWND, language: Language) -> Option<Vec<PathBuf>> {
    let filter_raw = i18n::tr(language, "dialog.open_filter");
    let filter = to_wide(&filter_raw.replace("\\0", "\0"));
    let mut buffer = vec![0u16; 4096];
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        Flags: OFN_EXPLORER
            | OFN_FILEMUSTEXIST
            | OFN_PATHMUSTEXIST
            | OFN_HIDEREADONLY
            | OFN_ALLOWMULTISELECT,
        ..Default::default()
    };
    if !unsafe { GetOpenFileNameW(&mut ofn).as_bool() } {
        return None;
    }
    Some(parse_multi_select(&buffer))
}

fn parse_multi_select(buffer: &[u16]) -> Vec<PathBuf> {
    let mut parts: Vec<String> = Vec::new();
    let mut current: Vec<u16> = Vec::new();
    for &ch in buffer {
        if ch == 0 {
            if current.is_empty() {
                break;
            }
            parts.push(String::from_utf16_lossy(&current));
            current.clear();
        } else {
            current.push(ch);
        }
    }
    if parts.is_empty() {
        return Vec::new();
    }
    if parts.len() == 1 {
        return vec![PathBuf::from(&parts[0])];
    }
    let dir = PathBuf::from(&parts[0]);
    parts[1..].iter().map(|name| dir.join(name)).collect()
}

fn collect_folder_files(folder: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    let mut stack = vec![folder.to_path_buf()];
    while let Some(dir) = stack.pop() {
        let entries = match std::fs::read_dir(&dir) {
            Ok(entries) => entries,
            Err(e) => {
                crate::log_debug(&format!(
                    "Batch: failed to read dir {}: {}",
                    dir.display(),
                    e
                ));
                continue;
            }
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
                continue;
            }
            if is_supported_batch_input_path(&path) {
                files.push(path);
            }
        }
    }
    files
}

fn is_supported_batch_input_path(path: &Path) -> bool {
    if is_temp_batch_path(path) {
        return false;
    }
    if is_audio_path(path) {
        return false;
    }
    if is_pdf_path(path)
        || is_docx_path(path)
        || is_doc_path(path)
        || is_pptx_path(path)
        || is_ppt_path(path)
        || is_epub_path(path)
        || is_html_path(path)
        || is_spreadsheet_path(path)
    {
        return true;
    }
    matches!(
        path.extension()
            .and_then(|s| s.to_str())
            .map(|s| s.to_ascii_lowercase())
            .as_deref(),
        None | Some("txt" | "md" | "rtf" | "csv" | "log")
    )
}

fn is_temp_batch_path(path: &Path) -> bool {
    let name = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
    let lower = name.to_ascii_lowercase();
    if lower.ends_with(".tmp") || lower.ends_with(".part") {
        return true;
    }
    false
}

fn is_checked(hwnd: HWND) -> bool {
    unsafe {
        SendMessageW(
            hwnd,
            windows::Win32::UI::WindowsAndMessaging::BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
        .0 == BST_CHECKED.0 as isize
    }
}

fn read_control_text(hwnd: HWND) -> String {
    let len = unsafe { GetWindowTextLengthW(hwnd) } as usize;
    if len == 0 {
        return String::new();
    }
    let mut buffer = vec![0u16; len + 1];
    unsafe {
        GetWindowTextW(hwnd, &mut buffer);
    }
    String::from_utf16_lossy(&buffer[..len])
}

fn load_tts_settings(parent: HWND, voice: String, language: Language) -> TtsSettings {
    unsafe {
        with_state(parent, |state| TtsSettings {
            voice,
            split_on_newline: state.settings.split_on_newline,
            audiobook_split: state.settings.audiobook_split,
            audiobook_split_by_text: state.settings.audiobook_split_by_text,
            audiobook_split_text: state.settings.audiobook_split_text.clone(),
            audiobook_split_text_requires_newline: state
                .settings
                .audiobook_split_text_requires_newline,
            audiobook_split_by_epub_chapter: state.settings.audiobook_split_by_epub_chapter,
            audiobook_split_by_time: state.settings.audiobook_split_by_time,
            audiobook_split_minutes: state.settings.audiobook_split_minutes,
            audiobook_split_start_number: state.settings.audiobook_split_start_number,
            audiobook_m4b_bitrate: state.settings.audiobook_m4b_bitrate,
            tts_engine: state.settings.tts_engine,
            dictionary: state.settings.dictionary.clone(),
            tts_rate: state.settings.tts_rate,
            tts_pitch: state.settings.tts_pitch,
            tts_volume: state.settings.tts_volume,
            language,
        })
        .unwrap_or_else(|| TtsSettings {
            voice: "it-IT-IsabellaNeural".to_string(),
            split_on_newline: true,
            audiobook_split: 0,
            audiobook_split_by_text: false,
            audiobook_split_text: String::new(),
            audiobook_split_text_requires_newline: true,
            audiobook_split_by_epub_chapter: false,
            audiobook_split_by_time: false,
            audiobook_split_minutes: 5,
            audiobook_split_start_number: 1,
            audiobook_m4b_bitrate: 128,
            tts_engine: TtsEngine::Edge,
            dictionary: Vec::new(),
            tts_rate: 0,
            tts_pitch: 0,
            tts_volume: 100,
            language,
        })
    }
}

fn run_batch(
    hwnd: HWND,
    items: Vec<PathBuf>,
    batch_settings: BatchSettings,
    tts_settings: TtsSettings,
    cancel: Arc<AtomicBool>,
    message_queue: Arc<Mutex<VecDeque<BatchMessage>>>,
) {
    // Initialize COM once for the entire batch thread so that SAPI and Media
    // Foundation objects stay valid across all items.  Without this, each
    // `speak_sapi_to_file` call creates and destroys its own COM apartment,
    // which can invalidate Media Foundation state that is initialized only
    // once (via `Once`), leading to crashes after several items.
    let _com_guard = match crate::com_guard::ComGuard::new_sta() {
        Ok(g) => Some(g),
        Err(e) => {
            log_debug(&format!("Batch: COM init failed: {e}"));
            None
        }
    };

    let mut sapi_voice: Option<crate::sapi5_engine::SapiVoice> = None;
    if matches!(tts_settings.tts_engine, TtsEngine::Sapi5) {
        match crate::sapi5_engine::SapiVoice::new() {
            Ok(voice) => {
                log_debug("Batch: SAPI5 voice initialized for reuse.");
                sapi_voice = Some(voice);
            }
            Err(e) => {
                log_debug(&format!("Batch: SAPI5 voice init failed: {e}"));
            }
        }
    }

    let total = items.len();
    let mut completed = 0usize;
    let mut results: Vec<BatchResultItem> = Vec::new();
    log_debug(&format!("Batch worker start. items={total}"));
    for (index, input) in items.into_iter().enumerate() {
        log_debug(&format!(
            "Batch item start. index={} path={}",
            index + 1,
            input.display()
        ));
        if cancel.load(Ordering::SeqCst) {
            log_debug(&format!(
                "Batch item canceled before start. index={} path={}",
                index + 1,
                input.display()
            ));
            post_status(
                &message_queue,
                hwnd,
                index,
                BatchStatusCode::Canceled,
                String::new(),
            );
            results.push(BatchResultItem {
                input,
                status: BatchStatusCode::Canceled,
                outputs: Vec::new(),
                error: None,
            });
            continue;
        }
        post_status(
            &message_queue,
            hwnd,
            index,
            BatchStatusCode::Running,
            String::new(),
        );
        let mut attempts = 0;
        loop {
            attempts += 1;
            log_debug(&format!(
                "Batch item attempt. index={} attempt={} path={}",
                index + 1,
                attempts,
                input.display()
            ));
            match export_single_audiobook(
                &input,
                &batch_settings,
                &tts_settings,
                cancel.clone(),
                sapi_voice.as_ref(),
            ) {
                Ok(outputs) => {
                    completed += 1;
                    log_debug(&format!(
                        "Batch item done. index={} outputs={} path={}",
                        index + 1,
                        outputs.len(),
                        input.display()
                    ));
                    let output_label = if outputs.len() > 1 {
                        i18n::tr(tts_settings.language, "batch_audiobooks.output_multiple")
                    } else {
                        outputs
                            .first()
                            .map(|p| p.to_string_lossy().to_string())
                            .unwrap_or_default()
                    };
                    post_status(
                        &message_queue,
                        hwnd,
                        index,
                        BatchStatusCode::Done,
                        output_label,
                    );
                    post_log(
                        &message_queue,
                        hwnd,
                        &i18n::tr_f(
                            tts_settings.language,
                            "batch_audiobooks.log.done",
                            &[("file", &input.to_string_lossy())],
                        ),
                    );
                    results.push(BatchResultItem {
                        input,
                        status: BatchStatusCode::Done,
                        outputs,
                        error: None,
                    });
                    break;
                }
                Err(err) => {
                    log_debug(&format!(
                        "Batch item error. index={} attempt={} path={} err={}",
                        index + 1,
                        attempts,
                        input.display(),
                        err
                    ));
                    let transient = is_transient_error(&err);
                    if transient && attempts <= 2 && !cancel.load(Ordering::SeqCst) {
                        post_log(
                            &message_queue,
                            hwnd,
                            &i18n::tr_f(
                                tts_settings.language,
                                "batch_audiobooks.log.retry",
                                &[("file", &input.to_string_lossy()), ("err", &err)],
                            ),
                        );
                        let wait = if attempts == 1 { 1 } else { 3 };
                        post_log(
                            &message_queue,
                            hwnd,
                            &i18n::tr_f(
                                tts_settings.language,
                                "batch_audiobooks.log.retry_wait",
                                &[("seconds", &wait.to_string())],
                            ),
                        );
                        std::thread::sleep(Duration::from_secs(wait));
                        continue;
                    }
                    let output_label = String::new();
                    post_status(
                        &message_queue,
                        hwnd,
                        index,
                        BatchStatusCode::Failed,
                        output_label,
                    );
                    post_log(
                        &message_queue,
                        hwnd,
                        &i18n::tr_f(
                            tts_settings.language,
                            "batch_audiobooks.log.failed",
                            &[("file", &input.to_string_lossy()), ("err", &err)],
                        ),
                    );
                    if is_antivirus_related(&err) {
                        post_log(
                            &message_queue,
                            hwnd,
                            &i18n::tr(tts_settings.language, "batch_audiobooks.log.antivirus_hint"),
                        );
                    }
                    results.push(BatchResultItem {
                        input,
                        status: BatchStatusCode::Failed,
                        outputs: Vec::new(),
                        error: Some(err),
                    });
                    break;
                }
            }
        }
        if cancel.load(Ordering::SeqCst) {
            completed = completed.min(total);
        }
        post_progress(&message_queue, hwnd, completed, total);
        log_debug(&format!(
            "Batch item end. index={} completed={}/{}",
            index + 1,
            completed,
            total
        ));
    }
    log_debug("Batch worker finished items. Writing report.");
    let report_path = match write_report(&batch_settings, &tts_settings, &results) {
        Ok(path) => Some(path),
        Err(e) => {
            log_debug(&format!("Batch report failed: {}", e));
            None
        }
    };
    log_debug(&format!(
        "Batch report done. path={}",
        report_path
            .as_ref()
            .map(|p| p.display().to_string())
            .unwrap_or_else(|| "(none)".to_string())
    ));
    post_done(&message_queue, hwnd, report_path);
}

fn export_single_audiobook(
    input: &Path,
    batch_settings: &BatchSettings,
    tts: &TtsSettings,
    cancel: Arc<AtomicBool>,
    sapi_voice: Option<&crate::sapi5_engine::SapiVoice>,
) -> Result<Vec<PathBuf>, String> {
    let use_epub_split = tts.audiobook_split_by_epub_chapter && is_epub_path(input);
    let epub_chapters = if use_epub_split {
        read_epub_chapters(input, tts.language)?
    } else {
        Vec::new()
    };
    let text = if use_epub_split {
        String::new()
    } else {
        read_text_for_audiobook(input, tts.language)?
    };
    if use_epub_split {
        if epub_chapters
            .iter()
            .all(|chapter| chapter.trim().is_empty())
        {
            return Err(crate::settings::tts_no_text_message(tts.language));
        }
    } else if text.trim().is_empty() {
        return Err(crate::settings::tts_no_text_message(tts.language));
    }
    let cleaned = if use_epub_split {
        String::new()
    } else {
        strip_dashed_lines(&text)
    };
    let split_by_time = if use_epub_split {
        false
    } else {
        tts.audiobook_split_by_time
    };
    let split_minutes = tts.audiobook_split_minutes.clamp(1, 20);
    let split_start_number = tts.audiobook_split_start_number.clamp(1, 99);
    let split_by_text = tts.audiobook_split_by_text && !split_by_time && !use_epub_split;
    let split_parts = if split_by_time || use_epub_split {
        0
    } else {
        tts.audiobook_split
    };
    let mixed_needed = if use_epub_split {
        epub_chapters.iter().any(|chapter| has_voice_tags(chapter))
    } else {
        has_voice_tags(&cleaned)
    };

    let (parts, mixed_parts) = if use_epub_split {
        if mixed_needed {
            (
                None,
                build_mixed_audiobook_parts_from_sections(
                    &epub_chapters,
                    tts.split_on_newline,
                    &tts.dictionary,
                    tts.tts_engine,
                ),
            )
        } else {
            (
                build_audiobook_parts_from_sections(
                    &epub_chapters,
                    tts.split_on_newline,
                    &tts.dictionary,
                    tts.tts_engine,
                ),
                None,
            )
        }
    } else if split_by_text {
        let (normalized, entries) = collect_marker_entries(
            &cleaned,
            &tts.audiobook_split_text,
            tts.audiobook_split_text_requires_newline,
        );
        if entries.is_empty() {
            (None, None)
        } else {
            let positions: Vec<usize> = entries.iter().map(|e| e.pos).collect();
            if mixed_needed {
                (
                    None,
                    build_mixed_audiobook_parts_by_positions(
                        &normalized,
                        &positions,
                        tts.split_on_newline,
                        &tts.dictionary,
                        tts.tts_engine,
                    ),
                )
            } else {
                (
                    build_audiobook_parts_by_positions(
                        &normalized,
                        &positions,
                        tts.split_on_newline,
                        &tts.dictionary,
                        tts.tts_engine,
                    ),
                    None,
                )
            }
        }
    } else {
        (None, None)
    };
    if use_epub_split && parts.is_none() && mixed_parts.is_none() {
        return Err(crate::settings::tts_no_text_message(tts.language));
    }

    if mixed_needed {
        let parts = match mixed_parts {
            Some(parts) => parts,
            None => {
                let chunks = split_into_tts_chunks(
                    &cleaned,
                    tts.split_on_newline,
                    &tts.dictionary,
                    tts.tts_engine,
                );
                split_tts_chunks_by_count(&chunks, split_parts)
            }
        };
        if parts.is_empty() {
            return Err(crate::settings::tts_no_text_message(tts.language));
        }
        let output_paths = build_output_paths(input, parts.len(), batch_settings, tts.language)?;
        match export_parts_mixed(&parts, &output_paths, tts, cancel.clone()) {
            Ok(()) => {
                if split_by_time {
                    let Some(output) = output_paths.first() else {
                        return Err("Batch: missing output path for time split".to_string());
                    };
                    return segment_batch_output(output, split_minutes, split_start_number);
                }
                Ok(output_paths)
            }
            Err(err) => {
                if !cancel.load(Ordering::SeqCst) {
                    for path in &output_paths {
                        if let Err(e) = std::fs::remove_file(path) {
                            crate::log_debug(&format!("Failed to remove temp batch file: {}", e));
                        }
                    }
                }
                Err(err)
            }
        }
    } else {
        let parts = match parts {
            Some(parts) => parts,
            None => {
                let prepared = prepare_tts_text(&cleaned, tts.split_on_newline, &tts.dictionary);
                let chunks = split_text(&prepared);
                split_chunks_by_count(&chunks, split_parts)
            }
        };
        if parts.is_empty() {
            return Err(crate::settings::tts_no_text_message(tts.language));
        }
        let output_paths = build_output_paths(input, parts.len(), batch_settings, tts.language)?;
        match export_parts(&parts, &output_paths, tts, cancel.clone(), sapi_voice) {
            Ok(()) => {
                if split_by_time {
                    let Some(output) = output_paths.first() else {
                        return Err("Batch: missing output path for time split".to_string());
                    };
                    return segment_batch_output(output, split_minutes, split_start_number);
                }
                Ok(output_paths)
            }
            Err(err) => {
                if !cancel.load(Ordering::SeqCst) {
                    for path in &output_paths {
                        if let Err(e) = std::fs::remove_file(path) {
                            crate::log_debug(&format!("Failed to remove temp batch file: {}", e));
                        }
                    }
                }
                Err(err)
            }
        }
    }
}

fn segment_batch_output(
    output: &Path,
    split_minutes: u32,
    split_start_number: u32,
) -> Result<Vec<PathBuf>, String> {
    let stem = output
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook");
    let ext = output.extension().and_then(|s| s.to_str());
    let pattern = if let Some(ext) = ext {
        output.with_file_name(format!("{stem}_part%02d.{ext}"))
    } else {
        output.with_file_name(format!("{stem}_part%02d"))
    };
    let segment_seconds = split_minutes.saturating_mul(60);
    crate::ffmpeg_export::segment_audio_file(
        output,
        &pattern,
        segment_seconds,
        split_start_number,
    )?;

    let parent = output
        .parent()
        .ok_or_else(|| "Batch: output has no parent directory".to_string())?;
    let prefix = format!("{stem}_part");
    let mut outputs = Vec::new();
    for entry in
        std::fs::read_dir(parent).map_err(|e| format!("Batch: failed to read output dir: {}", e))?
    {
        let entry = entry.map_err(|e| format!("Batch: output dir entry failed: {}", e))?;
        let path = entry.path();
        let Some(file_name) = path.file_name().and_then(|s| s.to_str()) else {
            continue;
        };
        if !file_name.starts_with(&prefix) {
            continue;
        }
        if let Some(ext) = ext
            && path.extension().and_then(|s| s.to_str()).unwrap_or("") != ext
        {
            continue;
        }
        outputs.push(path);
    }
    outputs.sort();
    if outputs.is_empty() {
        return Err("Batch: no segments were generated".to_string());
    }
    if let Err(e) = std::fs::remove_file(output) {
        crate::log_debug(&format!(
            "Failed to remove original audiobook after segmenting: {}",
            e
        ));
    }
    Ok(outputs)
}

fn split_chunks_by_count(chunks: &[String], split_parts: u32) -> Vec<Vec<String>> {
    let parts = if split_parts == 0 {
        1
    } else {
        split_parts as usize
    };
    let total_chunks = chunks.len();
    if total_chunks == 0 {
        return Vec::new();
    }
    let parts = if total_chunks < parts {
        total_chunks
    } else {
        parts
    };
    let chunks_per_part = total_chunks.div_ceil(parts);
    let mut out = Vec::new();
    for part_idx in 0..parts {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }
        out.push(chunks[start_idx..end_idx].to_vec());
    }
    out
}

fn split_tts_chunks_by_count(chunks: &[TtsChunk], split_parts: u32) -> Vec<Vec<TtsChunk>> {
    let parts = if split_parts == 0 {
        1
    } else {
        split_parts as usize
    };
    let total_chunks = chunks.len();
    if total_chunks == 0 {
        return Vec::new();
    }
    let parts = if total_chunks < parts {
        total_chunks
    } else {
        parts
    };
    let chunks_per_part = total_chunks.div_ceil(parts);
    let mut out = Vec::new();
    for part_idx in 0..parts {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }
        out.push(chunks[start_idx..end_idx].to_vec());
    }
    out
}

fn build_output_paths(
    input: &Path,
    parts_len: usize,
    settings: &BatchSettings,
    language: Language,
) -> Result<Vec<PathBuf>, String> {
    let base = input
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook");
    let mut base_name = sanitize_filename(base);
    if base_name.is_empty() {
        base_name = "audiobook".to_string();
    }
    let ext = match settings.format {
        AudioFormat::Mp3 => "mp3",
        AudioFormat::Wav => "wav",
    };

    let mut output_dir = settings.output_folder.clone();
    if settings.create_subfolder {
        output_dir = ensure_unique_folder(&output_dir.join(&base_name), settings.avoid_overwrite)?;
        if let Err(err) = std::fs::create_dir_all(&output_dir) {
            return Err(format!(
                "{}: {}",
                i18n::tr(language, "batch_audiobooks.error.invalid_output_folder"),
                err
            ));
        }
    }

    let width = std::cmp::max(2, parts_len.to_string().len());
    let mut outputs = Vec::new();
    for idx in 0..parts_len {
        let file_name = if parts_len > 1 {
            format!("{base_name} - {:0width$}.{ext}", idx + 1, width = width)
        } else {
            format!("{base_name}.{ext}")
        };
        let path = output_dir.join(file_name);
        let path = if settings.avoid_overwrite {
            unique_path(&path)
        } else {
            if path.exists() {
                return Err(format!("File already exists: {}", path.display()));
            }
            path
        };
        outputs.push(path);
    }
    Ok(outputs)
}

fn ensure_unique_folder(base: &Path, avoid_overwrite: bool) -> Result<PathBuf, String> {
    if !base.exists() {
        return Ok(base.to_path_buf());
    }
    if !avoid_overwrite {
        return Err(format!("Folder already exists: {}", base.display()));
    }
    for idx in 1..1000 {
        let candidate = PathBuf::from(format!("{} ({})", base.display(), idx));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    Err("Unable to find a unique folder name.".to_string())
}

fn unique_path(path: &Path) -> PathBuf {
    if !path.exists() {
        return path.to_path_buf();
    }
    let stem = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook");
    let ext = path.extension().and_then(|s| s.to_str());
    for idx in 1..1000 {
        let candidate = if let Some(ext) = ext {
            path.with_file_name(format!("{stem} ({idx}).{ext}"))
        } else {
            path.with_file_name(format!("{stem} ({idx})"))
        };
        if !candidate.exists() {
            return candidate;
        }
    }
    path.to_path_buf()
}

fn export_parts(
    parts: &[Vec<String>],
    outputs: &[PathBuf],
    tts: &TtsSettings,
    cancel: Arc<AtomicBool>,
    sapi_voice: Option<&crate::sapi5_engine::SapiVoice>,
) -> Result<(), String> {
    if parts.len() != outputs.len() {
        return Err("Output count mismatch.".to_string());
    }
    let mut progress = 0usize;
    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if cancel.load(Ordering::SeqCst) {
            return Err(i18n::tr(tts.language, "tts.cancelled"));
        }
        let output = &outputs[part_idx];
        match tts.tts_engine {
            TtsEngine::Edge => {
                let options = crate::tts_engine::AudiobookCommonOptions {
                    voice: &tts.voice,
                    output,
                    progress_hwnd: HWND(0),
                    cancel: cancel.clone(),
                    language: tts.language,
                    audiobook_bitrate_kbps: tts.audiobook_m4b_bitrate,
                    rate: tts.tts_rate,
                    pitch: tts.tts_pitch,
                    volume: tts.tts_volume,
                    sapi4_threads: None,
                };
                run_tts_audiobook_part(part_chunks, &mut progress, &options)?;
            }
            TtsEngine::Sapi4 => {
                // SAPI4 not implemented for batch yet in this context, or maybe it was empty before too?
                // The original code had empty block for Sapi4.
            }
            TtsEngine::Sapi5 => {
                crate::log_debug(&format!(
                    "Batch SAPI5 part start. part_idx={} chunks={} output={}",
                    part_idx + 1,
                    part_chunks.len(),
                    output.display()
                ));
                let is_mp3 = output
                    .extension()
                    .is_some_and(|e| e.eq_ignore_ascii_case("mp3"));
                if is_mp3 {
                    let wav_path = output.with_extension("wav.tmp");
                    let options = crate::sapi5_engine::SapiExportOptions {
                        chunks: part_chunks,
                        voice_name: &tts.voice,
                        output_path: &wav_path,
                        language: tts.language,
                        rate: tts.tts_rate,
                        pitch: tts.tts_pitch,
                        volume: tts.tts_volume,
                        audiobook_bitrate_kbps: tts.audiobook_m4b_bitrate,
                        cancel: cancel.clone(),
                    };
                    if let Some(voice) = sapi_voice {
                        crate::sapi5_engine::speak_sapi_to_file_with_sapi_voice(
                            voice,
                            options,
                            |_chunk_idx| {
                                progress += 1;
                            },
                        )?;
                    } else {
                        crate::sapi5_engine::speak_sapi_to_file(options, |_chunk_idx| {
                            progress += 1;
                        })?;
                    }

                    crate::log_debug("Batch SAPI5: converting WAV to MP3 via FFmpeg.");
                    let convert_res = convert_wav_to_mp3_ffmpeg_with_timeout(
                        &wav_path,
                        output,
                        tts.audiobook_m4b_bitrate,
                        cancel.clone(),
                    );
                    match convert_res {
                        Ok(()) => {
                            if let Err(e) = std::fs::remove_file(&wav_path) {
                                crate::log_debug(&format!(
                                    "Batch SAPI5: failed to remove temp WAV: {}",
                                    e
                                ));
                            }
                        }
                        Err(e) => {
                            if e == "FFmpeg MP3 conversion timed out" {
                                crate::log_debug("Batch SAPI5: MP3 conversion timed out.");
                            } else {
                                let dest_wav = output.with_extension("wav");
                                if let Err(rename_err) = std::fs::rename(&wav_path, &dest_wav) {
                                    crate::log_debug(&format!(
                                        "Batch SAPI5: failed to preserve WAV: {}",
                                        rename_err
                                    ));
                                } else {
                                    crate::log_debug(&format!(
                                        "Batch SAPI5: preserved WAV at {:?}",
                                        dest_wav
                                    ));
                                }
                            }
                            return Err(format!("FFmpeg MP3 conversion failed: {}", e));
                        }
                    }
                } else {
                    let options = crate::sapi5_engine::SapiExportOptions {
                        chunks: part_chunks,
                        voice_name: &tts.voice,
                        output_path: output,
                        language: tts.language,
                        rate: tts.tts_rate,
                        pitch: tts.tts_pitch,
                        volume: tts.tts_volume,
                        audiobook_bitrate_kbps: tts.audiobook_m4b_bitrate,
                        cancel: cancel.clone(),
                    };
                    if let Some(voice) = sapi_voice {
                        crate::sapi5_engine::speak_sapi_to_file_with_sapi_voice(
                            voice,
                            options,
                            |_chunk_idx| {
                                progress += 1;
                            },
                        )?;
                    } else {
                        crate::sapi5_engine::speak_sapi_to_file(options, |_chunk_idx| {
                            progress += 1;
                        })?;
                    }
                }
                crate::log_debug(&format!(
                    "Batch SAPI5 part done. part_idx={} output={}",
                    part_idx + 1,
                    output.display()
                ));
            }
        }
    }
    Ok(())
}

fn convert_wav_to_mp3_ffmpeg_with_timeout(
    wav_path: &Path,
    mp3_path: &Path,
    bitrate_kbps: u32,
    cancel: Arc<AtomicBool>,
) -> Result<(), String> {
    let data_size = crate::audio_utils::get_wav_data_size(wav_path).unwrap_or(0) as u64;
    let bytes_per_sec = 44_100u64 * 2u64;
    let duration_secs = if bytes_per_sec == 0 {
        0
    } else {
        (data_size / bytes_per_sec) as u64
    };
    let timeout_secs = (duration_secs.saturating_mul(10))
        .saturating_add(10)
        .max(30);
    let timeout = Duration::from_secs(timeout_secs);
    let (tx, rx) = mpsc::channel();
    let wav_path = wav_path.to_path_buf();
    let mp3_path = mp3_path.to_path_buf();
    let cancel = cancel.clone();
    std::thread::spawn(move || {
        let settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(bitrate_kbps),
        };
        let res = crate::ffmpeg_export::convert_audio_file(
            &wav_path,
            &mp3_path,
            &settings,
            Some(cancel),
            None,
        );
        let _unused = tx.send(res);
    });
    match rx.recv_timeout(timeout) {
        Ok(res) => res,
        Err(mpsc::RecvTimeoutError::Timeout) => Err("FFmpeg MP3 conversion timed out".to_string()),
        Err(_) => Err("FFmpeg MP3 conversion worker failed".to_string()),
    }
}

fn export_parts_mixed(
    parts: &[Vec<TtsChunk>],
    outputs: &[PathBuf],
    tts: &TtsSettings,
    cancel: Arc<AtomicBool>,
) -> Result<(), String> {
    if parts.len() != outputs.len() {
        return Err("Output count mismatch.".to_string());
    }
    let mut progress = 0usize;
    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if cancel.load(Ordering::SeqCst) {
            return Err(i18n::tr(tts.language, "tts.cancelled"));
        }
        let output = &outputs[part_idx];
        let options = crate::tts_engine::AudiobookCommonOptions {
            voice: &tts.voice,
            output,
            progress_hwnd: HWND(0),
            cancel: cancel.clone(),
            language: tts.language,
            audiobook_bitrate_kbps: tts.audiobook_m4b_bitrate,
            rate: tts.tts_rate,
            pitch: tts.tts_pitch,
            volume: tts.tts_volume,
            sapi4_threads: None,
        };
        let config = MixedAudiobookConfig {
            main_engine: tts.tts_engine,
        };
        render_mixed_audiobook_part(part_chunks, &mut progress, output, &options, &config)?;
    }
    Ok(())
}

fn read_text_for_audiobook(path: &Path, language: Language) -> Result<String, String> {
    if is_pdf_path(path) {
        return read_pdf_text(path, language);
    }
    if is_docx_path(path) {
        return read_docx_text(path, language);
    }
    if is_pptx_path(path) || is_ppt_path(path) {
        return read_ppt_text(path, language);
    }
    if is_epub_path(path) {
        return read_epub_text(path, language);
    }
    if is_html_path(path) {
        return read_html_text(path, language).map(|(text, _)| text);
    }
    if is_doc_path(path) {
        return read_doc_text(path, language);
    }
    if is_spreadsheet_path(path) {
        return read_spreadsheet_text(path, language);
    }
    let bytes = std::fs::read(path)
        .map_err(|err| crate::settings::error_open_file_message(language, err))?;
    decode_text(&bytes, language).map(|(text, _)| text)
}

fn is_transient_error(err: &str) -> bool {
    let lower = err.to_ascii_lowercase();
    lower.contains("timeout")
        || lower.contains("tempor")
        || lower.contains("websocket")
        || lower.contains("connection")
        || lower.contains("rate limit")
        || lower.contains("429")
        || lower.contains("502")
        || lower.contains("503")
        || lower.contains("service unavailable")
}

fn is_antivirus_related(err: &str) -> bool {
    let lower = err.to_ascii_lowercase();
    lower.contains("access is denied")
        || lower.contains("permission")
        || lower.contains("cannot access the file")
        || lower.contains("blocked")
}

fn post_status(
    queue: &Arc<Mutex<VecDeque<BatchMessage>>>,
    hwnd: HWND,
    index: usize,
    status: BatchStatusCode,
    output: String,
) {
    if unsafe { !IsWindow(hwnd).as_bool() } {
        return;
    }
    if let Ok(mut queue) = queue.lock() {
        queue.push_back(BatchMessage::Status(StatusUpdate {
            index,
            status,
            output,
        }));
    }
}

fn post_log(queue: &Arc<Mutex<VecDeque<BatchMessage>>>, hwnd: HWND, line: &str) {
    if unsafe { !IsWindow(hwnd).as_bool() } {
        return;
    }
    if let Ok(mut queue) = queue.lock() {
        queue.push_back(BatchMessage::Log(line.to_string()));
    }
}

fn post_progress(
    queue: &Arc<Mutex<VecDeque<BatchMessage>>>,
    hwnd: HWND,
    completed: usize,
    total: usize,
) {
    if unsafe { !IsWindow(hwnd).as_bool() } {
        return;
    }
    if let Ok(mut queue) = queue.lock() {
        queue.push_back(BatchMessage::Progress(ProgressUpdate { completed, total }));
    }
}

fn post_done(queue: &Arc<Mutex<VecDeque<BatchMessage>>>, hwnd: HWND, report_path: Option<PathBuf>) {
    if unsafe { !IsWindow(hwnd).as_bool() } {
        return;
    }
    if let Ok(mut queue) = queue.lock() {
        queue.push_back(BatchMessage::Done(DoneUpdate { report_path }));
    }
    log_debug(&format!("Batch post_done queued. hwnd={}", hwnd.0));
}

fn write_report(
    batch: &BatchSettings,
    tts: &TtsSettings,
    results: &[BatchResultItem],
) -> Result<PathBuf, String> {
    let timestamp = Local::now().format("%Y-%m-%d %H:%M:%S");
    let mut lines = Vec::new();
    lines.push(format!("Batch report - {timestamp}"));
    lines.push(format!("Voice: {}", tts.voice));
    lines.push(format!(
        "Format: {}",
        match batch.format {
            AudioFormat::Mp3 => "MP3",
            AudioFormat::Wav => "WAV",
        }
    ));
    let split_desc = if tts.audiobook_split_by_time {
        format!("Split by time: {} minutes", tts.audiobook_split_minutes)
    } else if tts.audiobook_split_by_text {
        format!("Split by text: {}", tts.audiobook_split_text)
    } else if tts.audiobook_split == 0 {
        "Split: disabled".to_string()
    } else {
        format!("Split parts: {}", tts.audiobook_split)
    };
    lines.push(split_desc);
    lines.push(String::new());

    for item in results {
        let status = match item.status {
            BatchStatusCode::Done => "Done",
            BatchStatusCode::Failed => "Failed",
            BatchStatusCode::Canceled => "Canceled",
            BatchStatusCode::Running => "Running",
            BatchStatusCode::Pending => "Pending",
        };
        lines.push(format!("{} - {}", status, item.input.display()));
        if !item.outputs.is_empty() {
            for out in &item.outputs {
                lines.push(format!("  Output: {}", out.display()));
            }
        }
        if let Some(err) = &item.error {
            lines.push(format!("  Error: {}", err));
        }
        lines.push(String::new());
    }

    let report_name = i18n::tr(tts.language, "batch_audiobooks.report_filename");
    let report_path = batch.output_folder.join(report_name);
    std::fs::write(&report_path, lines.join("\r\n")).map_err(|e| e.to_string())?;
    Ok(report_path)
}
