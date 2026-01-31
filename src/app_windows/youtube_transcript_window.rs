use std::cmp::Ordering as CmpOrdering;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::CREATE_NO_WINDOW;
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::RichEdit::{CHARRANGE, EM_EXSETSEL};
use windows::Win32::UI::Controls::{
    BST_CHECKED, EM_SCROLLCARET, EM_SETSEL, WC_BUTTON, WC_COMBOBOXW, WC_STATIC,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL,
    CB_RESETCONTENT, CB_SETCURSEL, CBS_DROPDOWNLIST, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW,
    DefWindowProcW, DestroyWindow, DispatchMessageW, EN_CHANGE, ES_MULTILINE, ES_READONLY,
    GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW, IDYES, IsDialogMessageW,
    IsWindow, LoadCursorW, MB_ICONQUESTION, MB_YESNO, MSG, MessageBoxW, PostMessageW,
    RegisterClassW, SW_HIDE, SW_SHOW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW,
    SetWindowTextW, ShowWindow, TranslateMessage, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::{EM_REPLACESEL, to_wide, to_wide_normalized};
use crate::editor_manager::get_edit_text;
use crate::i18n;
use crate::settings::{Language, confirm_title, save_settings, settings_dir};
use crate::with_state;
use crate::{WM_FOCUS_EDITOR, get_active_edit, show_error};
use url::Url;

const YT_IMPORT_CLASS_NAME: &str = "SonarpadYouTubeTranscript";
const YT_ID_URL: usize = 9301;
const YT_ID_LOAD: usize = 9302;
const YT_ID_LANG: usize = 9303;
const YT_ID_TIMESTAMP: usize = 9304;
const YT_ID_OK: usize = 9305;
const YT_ID_CANCEL: usize = 9306;
const WM_YT_LOAD_COMPLETE: u32 = WM_APP + 40;
const WM_YT_TEXT_COMPLETE: u32 = WM_APP + 41;
const WM_YT_LOAD_CANCEL: u32 = WM_APP + 42;
const WM_YT_TEXT_CANCEL: u32 = WM_APP + 43;
const EVENT_OBJECT_FOCUS: u32 = 0x8005;
const EVENT_OBJECT_VALUECHANGE: u32 = 0x800E;
const OBJID_CLIENT: i32 = -4;
const CHILDID_SELF: i32 = 0;
const YTDLP_EXE_NAME: &str = "yt-dlp.exe";
const YTDLP_DOWNLOAD_URL: &str =
    "https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe";
const YTDLP_LATEST_API_URL: &str = "https://api.github.com/repos/yt-dlp/yt-dlp/releases/latest";
const YTDLP_USER_AGENT: &str = "Sonarpad/yt-dlp";
const YTDLP_SOCKET_TIMEOUT_SECS: &str = "10";
static YTDLP_UPDATE_CHECKED: AtomicBool = AtomicBool::new(false);

// YouTube InnerTube API constants
const INNERTUBE_API_URL: &str = "https://www.youtube.com/youtubei/v1/player";
const INNERTUBE_CLIENT_NAME: &str = "ANDROID";
const INNERTUBE_CLIENT_VERSION: &str = "19.29.37";
const INNERTUBE_USER_AGENT: &str = "com.google.android.youtube/19.29.37";

#[derive(Clone)]
struct ImportResult {
    text: String,
    include_timestamps: bool,
}

struct ImportInit {
    parent: HWND,
    language: Language,
    include_timestamps: bool,
    result: Arc<Mutex<Option<ImportResult>>>, // Corrected from `лық>` to `>`
}

struct ImportState {
    parent: HWND,
    language: Language,
    url_edit: HWND,
    load_button: HWND,
    lang_combo: HWND,
    timestamp_check: HWND,
    ok_button: HWND,
    status_label: HWND,
    loading: bool,
    cancelled: Arc<AtomicBool>,
    transcripts: Vec<YtSubtitleOption>,
    result: Arc<Mutex<Option<ImportResult>>>, // Corrected from `лық>` to `>`
}

struct Labels {
    title: String,
    url: String,
    load: String,
    language: String,
    include_timestamps: String,
    loading_languages: String,
    loading_transcript: String,
    ok: String,
    cancel: String,
    auto: String,
    invalid_url: String,
    no_transcript: String,
    import_error: String,
    no_document: String,
    ytdlp_prompt_download: String,
    ytdlp_download_failed: String,
    ytdlp_update_failed: String,
}

#[derive(Clone)]
struct YtSubtitleOption {
    language: String,
    code: String,
    is_generated: bool,
}

fn labels(language: Language) -> Labels {
    Labels {
        title: i18n::tr(language, "youtube.title"),
        url: i18n::tr(language, "youtube.url"),
        load: i18n::tr(language, "youtube.load"),
        language: i18n::tr(language, "youtube.language"),
        include_timestamps: i18n::tr(language, "youtube.include_timestamps"),
        loading_languages: i18n::tr(language, "youtube.loading_languages"),
        loading_transcript: i18n::tr(language, "youtube.loading_transcript"),
        ok: i18n::tr(language, "youtube.ok"),
        cancel: i18n::tr(language, "youtube.cancel"),
        auto: i18n::tr(language, "youtube.auto"),
        invalid_url: i18n::tr(language, "youtube.invalid_url"),
        no_transcript: i18n::tr(language, "youtube.no_transcript"),
        import_error: i18n::tr(language, "youtube.import_error"),
        no_document: i18n::tr(language, "youtube.no_document"),
        ytdlp_prompt_download: i18n::tr(language, "youtube.ytdlp_prompt_download"),
        ytdlp_download_failed: i18n::tr(language, "youtube.ytdlp_download_failed"),
        ytdlp_update_failed: i18n::tr(language, "youtube.ytdlp_update_failed"),
    }
}

fn ytdlp_command(path: &Path) -> Command {
    let mut cmd = Command::new(path);
    cmd.stdin(Stdio::null());
    cmd.creation_flags(CREATE_NO_WINDOW.0);
    cmd
}

pub fn import_youtube_transcript(parent: HWND) {
    let (language, include_timestamps) = unsafe {
        with_state(parent, |state| {
            (
                state.settings.language,
                state.settings.youtube_include_timestamps,
            )
        })
        .unwrap_or((Language::Italian, true))
    };
    let Some(result) = show_import_dialog(parent, language, include_timestamps) else {
        unsafe {
            if let Err(e) = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0)) {
                crate::log_debug(&format!("Failed to post WM_FOCUS_EDITOR: {}", e));
            }
        }
        return;
    };
    unsafe {
        if with_state(parent, |state| {
            state.settings.youtube_include_timestamps = result.include_timestamps;
            save_settings(state.settings.clone());
        })
        .is_none()
        {
            crate::log_debug("Failed to update YouTube settings state");
        }
    }

    let text = result.text;

    unsafe {
        let Some(hwnd_edit) = get_active_edit(parent) else {
            show_error(parent, language, &labels(language).no_document);
            if let Err(e) = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0)) {
                crate::log_debug(&format!("Failed to post WM_FOCUS_EDITOR: {}", e));
            }
            return;
        };
        SetFocus(hwnd_edit);
        let existing = get_edit_text(hwnd_edit);
        let combined = if existing.is_empty() {
            text
        } else {
            format!("{text}\n\n{existing}")
        };
        let wide = to_wide_normalized(&combined);
        SendMessageW(hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(-1));
        SendMessageW(
            hwnd_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(wide.as_ptr() as isize),
        );
        let end = combined.len() as i32;
        SendMessageW(
            hwnd_edit,
            EM_SETSEL,
            WPARAM(end as usize),
            LPARAM(end as isize),
        );
        let cr = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&cr as *const _ as isize),
        );
        SendMessageW(hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(0));
        SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        NotifyWinEvent(
            EVENT_OBJECT_VALUECHANGE,
            hwnd_edit,
            OBJID_CLIENT,
            CHILDID_SELF,
        );
        NotifyWinEvent(EVENT_OBJECT_FOCUS, hwnd_edit, OBJID_CLIENT, CHILDID_SELF);
        if let Err(e) = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0)) {
            crate::log_debug(&format!("Error: {:?}", e));
        }
    }
}

fn show_import_dialog(
    parent: HWND,
    language: Language,
    include_timestamps: bool,
) -> Option<ImportResult> {
    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let class_name = to_wide(YT_IMPORT_CLASS_NAME);
    let wc = windows::Win32::UI::WindowsAndMessaging::WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(import_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(ImportInit {
        parent,
        language,
        include_timestamps,
        result: result.clone(),
    });
    let labels = labels(language);
    let title = to_wide(&labels.title);

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            520,
            240,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };

    if hwnd.0 == 0 {
        return None;
    }

    unsafe {
        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);
    }

    let mut msg = MSG::default();
    loop {
        if !unsafe { IsWindow(hwnd).as_bool() } {
            break;
        }
        let res = unsafe { GetMessageW(&mut msg, HWND(0), 0, 0) };
        if res.0 == 0 || res.0 == -1 {
            break;
        }
        unsafe {
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                if let Err(_e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_CANCEL), LPARAM(0)) {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let ok = with_import_state(hwnd, |state| state.ok_button).unwrap_or(HWND(0));
                if GetFocus() == ok {
                    if let Err(_e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_OK), LPARAM(0)) {
                        crate::log_debug(&format!("Error: {:?}", _e));
                    }
                    continue;
                }
            }
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    unsafe {
        EnableWindow(parent, true);
        SetForegroundWindow(parent);
    }

    result.lock().unwrap_or_else(|e| e.into_inner()).clone()
}

unsafe extern "system" fn import_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "import_wndproc",
        || DefWindowProcW(hwnd, msg, wparam, lparam),
        || unsafe { import_wndproc_inner(hwnd, msg, wparam, lparam) },
    )
}

unsafe fn import_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*create_struct).lpCreateParams as *mut ImportInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = Box::from_raw(init_ptr);
            let labels = labels(init.language);
            let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

            let label_url = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.url).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                18,
                90,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let url_edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                110,
                16,
                290,
                22,
                hwnd,
                HMENU(YT_ID_URL as isize),
                HINSTANCE(0),
                None,
            );

            let load_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.load).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                410,
                15,
                90,
                26,
                hwnd,
                HMENU(YT_ID_LOAD as isize),
                HINSTANCE(0),
                None,
            );

            let label_lang = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.language).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                60,
                90,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let lang_combo = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                110,
                58,
                290,
                140,
                hwnd,
                HMENU(YT_ID_LANG as isize),
                HINSTANCE(0),
                None,
            );

            let status_label = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WINDOW_STYLE((ES_READONLY | ES_MULTILINE) as u32),
                110,
                86,
                390,
                32,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let timestamp_check = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.include_timestamps).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                110,
                112,
                260,
                22,
                hwnd,
                HMENU(YT_ID_TIMESTAMP as isize),
                HINSTANCE(0),
                None,
            );

            let ok_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.ok).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                310,
                154,
                90,
                28,
                hwnd,
                HMENU(YT_ID_OK as isize),
                HINSTANCE(0),
                None,
            );

            let cancel_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.cancel).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                410,
                154,
                90,
                28,
                hwnd,
                HMENU(YT_ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            for control in [
                label_url,
                url_edit,
                load_button,
                label_lang,
                lang_combo,
                status_label,
                timestamp_check,
                ok_button,
                cancel_button,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }
            ShowWindow(status_label, SW_HIDE);

            let state = Box::new(ImportState {
                parent: init.parent,
                language: init.language,
                url_edit,
                load_button,
                lang_combo,
                timestamp_check,
                ok_button,
                status_label,
                loading: false,
                cancelled: Arc::new(AtomicBool::new(false)),
                transcripts: Vec::new(),
                result: init.result.clone(),
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            let initial_check = if init.include_timestamps {
                BST_CHECKED.0
            } else {
                0
            };
            SendMessageW(
                timestamp_check,
                BM_SETCHECK,
                WPARAM(initial_check as usize),
                LPARAM(0),
            );
            SetFocus(url_edit);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            let notification = ((wparam.0 >> 16) & 0xffff) as u16;
            if cmd_id == YT_ID_LOAD {
                start_load_languages(hwnd);
                LRESULT(0)
            } else if cmd_id == YT_ID_OK {
                if with_import_state(hwnd, |state| {
                    if state.loading {
                        return;
                    }
                    if state.transcripts.is_empty() {
                        if !start_load_languages(hwnd) {
                            crate::log_debug("Failed to start load languages");
                        }
                        return;
                    }
                    let idx = SendMessageW(state.lang_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                    if idx < 0 || idx as usize >= state.transcripts.len() {
                        return;
                    }
                    let include_timestamps =
                        SendMessageW(state.timestamp_check, BM_GETCHECK, WPARAM(0), LPARAM(0)).0
                            == BST_CHECKED.0 as isize;

                    let transcript = state.transcripts[idx as usize].clone();
                    let url = read_edit_text(state.url_edit);
                    // Start loading text instead of closing immediately
                    if !start_load_transcript_text(hwnd, transcript, include_timestamps, url) {
                        crate::log_debug("Failed to start load transcript text");
                    }
                })
                .is_none()
                {
                    crate::log_debug("Failed to access import state");
                }
                LRESULT(0)
            } else if cmd_id == YT_ID_CANCEL {
                if with_import_state(hwnd, |state| {
                    state.cancelled.store(true, Ordering::SeqCst);
                    *state.result.lock().unwrap_or_else(|e| e.into_inner()) = None;
                })
                .is_none()
                {
                    crate::log_debug("Failed to access import state");
                }
                if let Err(_e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
                LRESULT(0)
            } else if cmd_id == YT_ID_URL && notification as u32 == EN_CHANGE {
                if with_import_state(hwnd, |state| {
                    if state.loading {
                        return;
                    }
                    state.transcripts.clear();
                    SendMessageW(state.lang_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
                })
                .is_none()
                {
                    crate::log_debug("Failed to access import state");
                }
                LRESULT(0)
            } else {
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        }
        WM_YT_LOAD_COMPLETE => {
            let result = unsafe { Box::from_raw(lparam.0 as *mut LoadResult) };
            finish_load_languages(hwnd, *result);
            LRESULT(0)
        }
        WM_YT_LOAD_CANCEL => {
            reset_languages_loading_state(hwnd);
            LRESULT(0)
        }
        WM_YT_TEXT_COMPLETE => {
            let result = unsafe { Box::from_raw(lparam.0 as *mut TextLoadResult) };
            finish_load_text(hwnd, *result);
            LRESULT(0)
        }
        WM_YT_TEXT_CANCEL => {
            reset_text_loading_state(hwnd);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                if let Err(_e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_CANCEL), LPARAM(0)) {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let url_edit = with_import_state(hwnd, |state| state.url_edit).unwrap_or(HWND(0));
                if focus == url_edit {
                    if let Err(_e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_LOAD), LPARAM(0)) {
                        crate::log_debug(&format!("Error: {:?}", _e));
                    }
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            crate::log_if_err!(DestroyWindow(hwnd));
            LRESULT(0)
        }
        WM_DESTROY => {
            if with_import_state(hwnd, |state| {
                EnableWindow(state.parent, true);
                SetForegroundWindow(state.parent);
                // Only focus editor if not in player mode (audiobook)
                if !crate::editor_manager::is_current_audiobook(state.parent) {
                    if let Err(e) =
                        PostMessageW(state.parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0))
                    {
                        crate::log_debug(&format!("Failed to post WM_FOCUS_EDITOR: {}", e));
                    }
                    if let Some(hwnd_edit) = get_active_edit(state.parent) {
                        NotifyWinEvent(EVENT_OBJECT_FOCUS, hwnd_edit, OBJID_CLIENT, CHILDID_SELF);
                    }
                }
            })
            .is_none()
            {
                crate::log_debug("Failed to access import state");
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ImportState;
            if !ptr.is_null() {
                let _unused_box = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_import_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut ImportState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ImportState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

struct LoadResult {
    transcripts: Vec<YtSubtitleOption>,
    error: Option<ImportError>,
}

struct TextLoadResult {
    text: String,
    include_timestamps: bool,
    error: Option<ImportError>,
}

fn start_load_languages(hwnd: HWND) -> bool {
    let mut language = Language::English;
    let mut url = String::new();
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);
    let mut already_loading = false;
    let mut cancelled_flag: Option<Arc<AtomicBool>> = None;

    if unsafe {
        with_import_state(hwnd, |state| {
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            language = state.language;
            url = read_edit_text(state.url_edit);
            already_loading = state.loading;
            state.loading = true;
            state.cancelled.store(false, Ordering::SeqCst);
            cancelled_flag = Some(state.cancelled.clone());
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access import state at L641");
    }
    if already_loading {
        return false;
    }

    let labels_data = labels(language);
    unsafe {
        if let Err(e) = SetWindowTextW(
            status,
            PCWSTR(to_wide(&labels_data.loading_languages).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set status text: {}", e));
        }
        ShowWindow(status, SW_SHOW);
        SetFocus(status);
        EnableWindow(edit, false);
        EnableWindow(load_button, false);
        EnableWindow(combo, false);
        EnableWindow(timestamp, false);
        EnableWindow(ok_button, false);
    }

    let cancelled = cancelled_flag.unwrap_or_else(|| Arc::new(AtomicBool::new(false)));

    std::thread::spawn(move || {
        let result = std::panic::catch_unwind(|| {
            if cancelled.load(Ordering::Relaxed) {
                return LoadResult {
                    transcripts: Vec::new(),
                    error: None,
                };
            }

            let download_url = match normalize_youtube_input_for_download(&url) {
                Some(value) => value,
                None => {
                    return LoadResult {
                        transcripts: Vec::new(),
                        error: Some(ImportError::InvalidUrl),
                    };
                }
            };

            // Extract video ID for InnerTube API
            let video_id = extract_video_id(&url);

            // Try InnerTube API first (faster and cleaner)
            if let Some(ref vid) = video_id {
                if cancelled.load(Ordering::Relaxed) {
                    return LoadResult {
                        transcripts: Vec::new(),
                        error: None,
                    };
                }
                crate::log_debug("Trying InnerTube API for transcript list...");
                if let Some(transcripts) = fetch_transcript_list_innertube(vid)
                    && !transcripts.is_empty()
                {
                    crate::log_debug(&format!(
                        "InnerTube API success: found {} transcripts",
                        transcripts.len()
                    ));
                    return LoadResult {
                        transcripts,
                        error: None,
                    };
                }
                crate::log_debug("InnerTube API failed, falling back to yt-dlp...");
            }

            // Fallback to yt-dlp
            let ytdlp_path = match ensure_ytdlp_available(hwnd, language, &labels_data) {
                Ok(Some(path)) => path,
                Ok(None) => {
                    return LoadResult {
                        transcripts: Vec::new(),
                        error: None,
                    };
                }
                Err(err) => {
                    crate::log_debug(&format!("yt-dlp download failed: {}", err));
                    unsafe {
                        show_error(hwnd, language, &labels_data.ytdlp_download_failed);
                    }
                    return LoadResult {
                        transcripts: Vec::new(),
                        error: None,
                    };
                }
            };

            let fetch_result =
                fetch_transcript_list_with_retry(&ytdlp_path, &download_url, &cancelled);

            if cancelled.load(Ordering::Relaxed) {
                return LoadResult {
                    transcripts: Vec::new(),
                    error: None,
                };
            }
            match fetch_result {
                Ok(transcripts) => {
                    if transcripts.is_empty() {
                        LoadResult {
                            transcripts: Vec::new(),
                            error: Some(ImportError::NoTranscript),
                        }
                    } else {
                        LoadResult {
                            transcripts,
                            error: None,
                        }
                    }
                }
                Err(err) => LoadResult {
                    transcripts: Vec::new(),
                    error: Some(err),
                },
            }
        })
        .unwrap_or_else(|e| {
            crate::log_debug(&format!("YouTube transcript thread panicked: {:?}", e));
            LoadResult {
                transcripts: Vec::new(),
                error: Some(ImportError::Other("Thread panicked".to_string())),
            }
        });

        if cancelled.load(Ordering::Relaxed) {
            return;
        }

        unsafe {
            if !IsWindow(hwnd).as_bool() {
                return;
            }
            if result.error.is_none() && result.transcripts.is_empty() {
                if let Err(e) = PostMessageW(hwnd, WM_YT_LOAD_CANCEL, WPARAM(0), LPARAM(0)) {
                    crate::log_debug(&format!("Failed to post WM_YT_LOAD_CANCEL: {}", e));
                }
            } else if let Err(e) = PostMessageW(
                hwnd,
                WM_YT_LOAD_COMPLETE,
                WPARAM(0),
                LPARAM(Box::into_raw(Box::new(result)) as isize),
            ) {
                crate::log_debug(&format!("Failed to post WM_YT_LOAD_COMPLETE: {}", e));
            }
        }
    });
    true
}

fn finish_load_languages(hwnd: HWND, result: LoadResult) {
    let mut language = Language::English;
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);

    let state_ok = unsafe {
        with_import_state(hwnd, |state| {
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            language = state.language;
            state.loading = false;
        })
    }
    .is_some();

    if !state_ok {
        crate::log_debug("Failed to access import state in finish_load_languages");
        return;
    }

    let labels_data = labels(language);
    unsafe {
        if let Err(_e) = SetWindowTextW(status, PCWSTR(to_wide("").as_ptr())) {
            crate::log_debug(&format!("Failed to set status text: {:?}", _e));
        }
        EnableWindow(edit, true);
        EnableWindow(load_button, true);
        EnableWindow(combo, true);
        EnableWindow(timestamp, true);
        EnableWindow(ok_button, true);
    }

    if let Some(err) = result.error {
        crate::log_debug(&format!("YouTube transcript error: {:?}", err));
        unsafe {
            show_error(hwnd, language, &error_message(language, &err));
            SetForegroundWindow(hwnd);
            if edit.0 != 0 {
                SetFocus(edit);
            }
        }
        return;
    }

    unsafe {
        SendMessageW(combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for transcript in result.transcripts.iter() {
            let mut label = format!("{} ({})", transcript.language, transcript.code);
            if transcript.is_generated {
                label.push_str(&format!(" - {}", labels_data.auto));
            }
            let wide = to_wide(&label);
            SendMessageW(
                combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(wide.as_ptr() as isize),
            );
        }
        SendMessageW(combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        SetFocus(combo);
    }

    if unsafe {
        with_import_state(hwnd, |state| {
            state.transcripts = result.transcripts;
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access import state at L786");
    }
}

fn reset_languages_loading_state(hwnd: HWND) {
    let mut language = Language::English;
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);

    let state_ok = unsafe {
        with_import_state(hwnd, |state| {
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            language = state.language;
            state.loading = false;
        })
    }
    .is_some();

    if !state_ok {
        crate::log_debug("Failed to access import state in reset_languages_loading_state");
        return;
    }

    unsafe {
        if let Err(_e) = SetWindowTextW(status, PCWSTR(to_wide("").as_ptr())) {
            crate::log_debug(&format!("Failed to set status text: {:?}", _e));
        }
        ShowWindow(status, SW_HIDE);
        EnableWindow(edit, true);
        EnableWindow(load_button, true);
        EnableWindow(combo, true);
        EnableWindow(timestamp, true);
        EnableWindow(ok_button, true);
        if edit.0 != 0 {
            SetFocus(edit);
        }
    }
}

fn reset_text_loading_state(hwnd: HWND) {
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);

    let state_ok = unsafe {
        with_import_state(hwnd, |state| {
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            state.loading = false;
        })
    }
    .is_some();

    if !state_ok {
        crate::log_debug("Failed to access import state in reset_text_loading_state");
        return;
    }

    unsafe {
        if let Err(_e) = SetWindowTextW(status, PCWSTR(to_wide("").as_ptr())) {
            crate::log_debug(&format!("Failed to set status text: {:?}", _e));
        }
        ShowWindow(status, SW_HIDE);
        EnableWindow(edit, true);
        EnableWindow(load_button, true);
        EnableWindow(combo, true);
        EnableWindow(timestamp, true);
        EnableWindow(ok_button, true);
    }
}

fn start_load_transcript_text(
    hwnd: HWND,
    transcript: YtSubtitleOption,
    include_timestamps: bool,
    url: String,
) -> bool {
    let mut language = Language::English;
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);
    let mut already_loading = false;
    let mut cancelled_flag: Option<Arc<AtomicBool>> = None;

    if unsafe {
        with_import_state(hwnd, |state| {
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            language = state.language;
            already_loading = state.loading;
            state.loading = true;
            state.cancelled.store(false, Ordering::SeqCst);
            cancelled_flag = Some(state.cancelled.clone());
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access import state at start_load_transcript_text");
    }
    if already_loading {
        return false;
    }

    let labels_data = labels(language);
    unsafe {
        if let Err(e) = SetWindowTextW(
            status,
            PCWSTR(to_wide(&labels_data.loading_transcript).as_ptr()),
        ) {
            crate::log_debug(&format!("Failed to set status text: {}", e));
        }
        ShowWindow(status, SW_SHOW);
        SetFocus(status);
        EnableWindow(edit, false);
        EnableWindow(load_button, false);
        EnableWindow(combo, false);
        EnableWindow(timestamp, false);
        EnableWindow(ok_button, false);
    }

    let cancelled = cancelled_flag.unwrap_or_else(|| Arc::new(AtomicBool::new(false)));
    let download_url = match normalize_youtube_input_for_download(&url) {
        Some(value) => value,
        None => {
            finish_load_text(
                hwnd,
                TextLoadResult {
                    text: String::new(),
                    include_timestamps,
                    error: Some(ImportError::InvalidUrl),
                },
            );
            return false;
        }
    };

    std::thread::spawn(move || {
        let result = std::panic::catch_unwind(|| {
            if cancelled.load(Ordering::Relaxed) {
                return TextLoadResult {
                    text: String::new(),
                    include_timestamps,
                    error: None,
                };
            }

            let ytdlp_path = match ensure_ytdlp_available(hwnd, language, &labels_data) {
                Ok(Some(path)) => path,
                Ok(None) => {
                    return TextLoadResult {
                        text: String::new(),
                        include_timestamps,
                        error: None,
                    };
                }
                Err(err) => {
                    crate::log_debug(&format!("yt-dlp download failed: {}", err));
                    unsafe {
                        show_error(hwnd, language, &labels_data.ytdlp_download_failed);
                    }
                    return TextLoadResult {
                        text: String::new(),
                        include_timestamps,
                        error: None,
                    };
                }
            };

            let fetch_result = fetch_transcript_text_with_retry(
                &transcript,
                include_timestamps,
                &cancelled,
                Some(&ytdlp_path),
                Some(&download_url),
            );

            if cancelled.load(Ordering::Relaxed) {
                return TextLoadResult {
                    text: String::new(),
                    include_timestamps,
                    error: None,
                };
            }

            match fetch_result {
                Ok(text) => TextLoadResult {
                    text,
                    include_timestamps,
                    error: None,
                },
                Err(err) => TextLoadResult {
                    text: String::new(),
                    include_timestamps,
                    error: Some(err),
                },
            }
        })
        .unwrap_or_else(|e| {
            crate::log_debug(&format!("YouTube transcript text thread panicked: {:?}", e));
            TextLoadResult {
                text: String::new(),
                include_timestamps,
                error: Some(ImportError::Other("Thread panicked".to_string())),
            }
        });

        if cancelled.load(Ordering::Relaxed) {
            return;
        }

        unsafe {
            if !IsWindow(hwnd).as_bool() {
                return;
            }
            if result.error.is_none() && result.text.is_empty() {
                if let Err(e) = PostMessageW(hwnd, WM_YT_TEXT_CANCEL, WPARAM(0), LPARAM(0)) {
                    crate::log_debug(&format!("Failed to post WM_YT_TEXT_CANCEL: {}", e));
                }
            } else if let Err(e) = PostMessageW(
                hwnd,
                WM_YT_TEXT_COMPLETE,
                WPARAM(0),
                LPARAM(Box::into_raw(Box::new(result)) as isize),
            ) {
                crate::log_debug(&format!("Failed to post WM_YT_TEXT_COMPLETE: {}", e));
            }
        }
    });
    true
}

fn finish_load_text(hwnd: HWND, result: TextLoadResult) {
    let mut language = Language::English;
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);

    let state_ok = unsafe {
        with_import_state(hwnd, |state| {
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            language = state.language;
            state.loading = false;
        })
    }
    .is_some();

    if !state_ok {
        crate::log_debug("Failed to access import state in finish_load_text");
        return;
    }

    unsafe {
        if let Err(_e) = SetWindowTextW(status, PCWSTR(to_wide("").as_ptr())) {
            crate::log_debug(&format!("Failed to set status text: {:?}", _e));
        }
        ShowWindow(status, SW_HIDE);
        EnableWindow(edit, true);
        EnableWindow(load_button, true);
        EnableWindow(combo, true);
        EnableWindow(timestamp, true);
        EnableWindow(ok_button, true);
    }

    if let Some(err) = result.error {
        crate::log_debug(&format!("YouTube transcript text error: {:?}", err));
        unsafe {
            show_error(hwnd, language, &error_message(language, &err));
            SetForegroundWindow(hwnd);
        }
        return;
    }

    if unsafe {
        with_import_state(hwnd, |state| {
            *state.result.lock().unwrap_or_else(|e| e.into_inner()) = Some(ImportResult {
                text: result.text,
                include_timestamps: result.include_timestamps,
            });
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access import state at finish_load_text");
    }

    unsafe {
        if let Err(_e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
            crate::log_debug(&format!("Error closing window: {:?}", _e));
        }
    }
}

fn read_edit_text(hwnd: HWND) -> String {
    if hwnd.0 == 0 {
        return String::new();
    }
    let len = unsafe { windows::Win32::UI::WindowsAndMessaging::GetWindowTextLengthW(hwnd) };
    if len == 0 {
        return String::new();
    }
    let mut buf = vec![0u16; (len + 1) as usize];
    unsafe {
        windows::Win32::UI::WindowsAndMessaging::GetWindowTextW(hwnd, &mut buf);
    }
    String::from_utf16_lossy(&buf[..len as usize])
}

fn extract_video_id(input: &str) -> Option<String> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return None;
    }
    if trimmed.len() == 11
        && trimmed
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '_')
    {
        return Some(trimmed.to_string());
    }

    let candidate = if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        trimmed.to_string()
    } else {
        format!("https://{trimmed}")
    };
    let url = Url::parse(&candidate).ok()?;
    let host = url.host_str()?.to_lowercase();

    if host.ends_with("youtu.be") {
        let path = url.path().trim_matches('/');
        if !path.is_empty() {
            return Some(path.split('/').next()?.to_string());
        }
    }

    if host.contains("youtube.com") {
        if let Some((_, value)) = url.query_pairs().find(|(key, _)| key == "v") {
            return Some(value.to_string());
        }
        let path = url.path().trim_matches('/');
        if let Some(id) = path
            .strip_prefix("shorts/")
            .or_else(|| path.strip_prefix("embed/"))
            .or_else(|| path.strip_prefix("live/"))
        {
            let id = id.split('/').next().unwrap_or("");
            if !id.is_empty() {
                return Some(id.to_string());
            }
        }
    }

    None
}

fn normalize_youtube_input_for_download(input: &str) -> Option<String> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Some(id) = extract_video_id(trimmed) {
        return Some(format!("https://www.youtube.com/watch?v={id}"));
    }
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        Some(trimmed.to_string())
    } else {
        Some(format!("https://{trimmed}"))
    }
}

fn ensure_ytdlp_available(
    hwnd: HWND,
    language: Language,
    labels: &Labels,
) -> Result<Option<PathBuf>, String> {
    let path = settings_dir().join(YTDLP_EXE_NAME);
    if path.exists() {
        check_ytdlp_update(hwnd, language, labels, &path);
        return Ok(Some(path));
    }
    let title = to_wide(&confirm_title(language));
    let message = to_wide(&labels.ytdlp_prompt_download);
    let response = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    if response != IDYES {
        return Ok(None);
    }
    match download_ytdlp(&path) {
        Ok(()) => Ok(Some(path)),
        Err(err) => Err(err),
    }
}

fn check_ytdlp_update(hwnd: HWND, language: Language, labels: &Labels, path: &Path) {
    if YTDLP_UPDATE_CHECKED.swap(true, Ordering::SeqCst) {
        return;
    }
    let installed = match ytdlp_installed_version(path) {
        Ok(version) => version,
        Err(err) => {
            crate::log_debug(&format!("yt-dlp version read failed: {}", err));
            return;
        }
    };
    let latest = match fetch_latest_ytdlp_version() {
        Ok(version) => version,
        Err(err) => {
            crate::log_debug(&format!("yt-dlp latest version fetch failed: {}", err));
            return;
        }
    };
    if compare_versions(&installed, &latest) != Some(CmpOrdering::Less) {
        return;
    }

    let prompt = i18n::tr_f(
        language,
        "youtube.ytdlp_update_prompt",
        &[("current", &installed), ("latest", &latest)],
    );
    let title = to_wide(&confirm_title(language));
    let message = to_wide(&prompt);
    let response = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    if response != IDYES {
        return;
    }
    if let Err(err) = download_ytdlp(path) {
        crate::log_debug(&format!("yt-dlp update failed: {}", err));
        unsafe {
            show_error(hwnd, language, &labels.ytdlp_update_failed);
        }
    }
}

fn ytdlp_installed_version(path: &Path) -> Result<String, String> {
    let output = ytdlp_command(path)
        .arg("--version")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|err| err.to_string())?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(stderr.trim().to_string());
    }
    let stdout = String::from_utf8_lossy(&output.stdout);
    let version = stdout.trim();
    if version.is_empty() {
        return Err("Empty yt-dlp version".to_string());
    }
    Ok(version.to_string())
}

fn fetch_latest_ytdlp_version() -> Result<String, String> {
    let client = reqwest::blocking::Client::builder()
        .user_agent(YTDLP_USER_AGENT)
        .timeout(std::time::Duration::from_secs(30))
        .build()
        .map_err(|err| err.to_string())?;
    let resp = client
        .get(YTDLP_LATEST_API_URL)
        .send()
        .map_err(|err| err.to_string())?;
    let resp = resp.error_for_status().map_err(|err| err.to_string())?;
    let value: serde_json::Value = resp.json().map_err(|err| err.to_string())?;
    let tag = value
        .get("tag_name")
        .and_then(|v| v.as_str())
        .unwrap_or("")
        .trim();
    let version = tag.trim_start_matches('v');
    if version.is_empty() {
        return Err("Missing yt-dlp tag_name".to_string());
    }
    Ok(version.to_string())
}

fn compare_versions(current: &str, latest: &str) -> Option<CmpOrdering> {
    let parse = |input: &str| {
        input
            .split('.')
            .map(|part| part.trim().parse::<u32>())
            .collect::<Result<Vec<_>, _>>()
            .ok()
    };
    let a = parse(current)?;
    let b = parse(latest)?;
    let max_len = a.len().max(b.len());
    for idx in 0..max_len {
        let left = *a.get(idx).unwrap_or(&0);
        let right = *b.get(idx).unwrap_or(&0);
        match left.cmp(&right) {
            CmpOrdering::Equal => {}
            other => return Some(other),
        }
    }
    Some(CmpOrdering::Equal)
}

fn download_ytdlp(target: &Path) -> Result<(), String> {
    if let Some(parent) = target.parent() {
        std::fs::create_dir_all(parent).map_err(|err| err.to_string())?;
    }
    let temp_path = target.with_extension("download");
    let client = reqwest::blocking::Client::builder()
        .user_agent(YTDLP_USER_AGENT)
        .timeout(std::time::Duration::from_secs(300))
        .build()
        .map_err(|err| err.to_string())?;
    let mut resp = client
        .get(YTDLP_DOWNLOAD_URL)
        .send()
        .map_err(|err| err.to_string())?;
    resp = resp.error_for_status().map_err(|err| err.to_string())?;
    let expected_len = resp.content_length();
    let mut file = std::fs::File::create(&temp_path).map_err(|err| err.to_string())?;
    let written = std::io::copy(&mut resp, &mut file).map_err(|err| err.to_string())?;
    file.sync_all().map_err(|err| err.to_string())?;
    if let Some(expected) = expected_len
        && written != expected
    {
        crate::log_if_err!(std::fs::remove_file(&temp_path));
        return Err("yt-dlp download incomplete".to_string());
    }
    if target.exists()
        && let Err(err) = std::fs::remove_file(target)
    {
        crate::log_if_err!(std::fs::remove_file(&temp_path));
        return Err(err.to_string());
    }
    std::fs::rename(&temp_path, target).map_err(|err| err.to_string())?;
    Ok(())
}

fn fetch_transcript_list_with_ytdlp(
    ytdlp_path: &Path,
    url: &str,
) -> Result<Vec<YtSubtitleOption>, ImportError> {
    let output = ytdlp_command(ytdlp_path)
        .arg("--no-playlist")
        .arg("--socket-timeout")
        .arg(YTDLP_SOCKET_TIMEOUT_SECS)
        .arg("--skip-download")
        .arg("--list-subs")
        .arg("--no-warnings")
        .arg("--force-ipv4")
        .arg(url)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|err| ImportError::Other(err.to_string()))?;

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    let combined = format!("{stdout}\n{stderr}");

    if !output.status.success() {
        if combined.to_lowercase().contains("no subtitles") {
            return Err(ImportError::NoTranscript);
        }
        return Err(ImportError::Other(combined.trim().to_string()));
    }

    let mut items = parse_ytdlp_subtitle_list(&combined);
    if items.is_empty() && combined.to_lowercase().contains("no subtitles") {
        return Err(ImportError::NoTranscript);
    }

    items.sort_by(|a, b| {
        a.is_generated
            .cmp(&b.is_generated)
            .then(a.language.cmp(&b.language))
            .then(a.code.cmp(&b.code))
    });
    Ok(items)
}

fn parse_ytdlp_subtitle_list(output: &str) -> Vec<YtSubtitleOption> {
    let mut items = Vec::new();
    let mut mode_generated = None;
    for raw in output.lines() {
        let line = raw.trim();
        if line.is_empty() {
            mode_generated = None;
            continue;
        }
        let lower = line.to_lowercase();
        if lower.contains("available subtitles") {
            mode_generated = Some(false);
            continue;
        }
        if lower.contains("available automatic captions") {
            mode_generated = Some(true);
            continue;
        }
        let Some(is_generated) = mode_generated else {
            continue;
        };
        if line.starts_with("Language") {
            continue;
        }
        let tokens: Vec<&str> = line.split_whitespace().collect();
        if tokens.is_empty() {
            continue;
        }
        let code = tokens[0].trim();
        if code.is_empty() || code.eq_ignore_ascii_case("language") {
            continue;
        }
        let mut name_parts = Vec::new();
        for token in tokens.iter().skip(1) {
            if is_subtitle_format_token(token) {
                break;
            }
            name_parts.push(*token);
        }
        let language = if name_parts.is_empty() {
            code.to_string()
        } else {
            name_parts.join(" ")
        };
        items.push(YtSubtitleOption {
            language,
            code: code.to_string(),
            is_generated,
        });
    }
    items
}

fn is_subtitle_format_token(token: &str) -> bool {
    let trimmed = token.trim_matches(|c: char| c == ',' || c == ';');
    let lower = trimmed.to_lowercase();
    matches!(
        lower.as_str(),
        "vtt" | "srt" | "ttml" | "srv3" | "srv2" | "srv1" | "json3" | "ass" | "ssa" | "lrc"
    ) || lower.contains(',')
}

#[derive(Debug)]
enum ImportError {
    InvalidUrl,
    NoTranscript,
    Other(String),
}

fn fetch_transcript_text_with_fallback(
    transcript: &YtSubtitleOption,
    include_timestamps: bool,
    ytdlp_path: Option<&Path>,
    download_url: Option<&str>,
) -> Result<String, ImportError> {
    // If this is an InnerTube transcript, use the direct API
    if is_innertube_transcript(transcript) {
        crate::log_debug("Fetching transcript via InnerTube API...");
        if let Some(text) = fetch_transcript_text_innertube(transcript, include_timestamps)
            && !text.is_empty()
        {
            crate::log_debug("InnerTube transcript fetch successful");
            return Ok(text);
        }
        crate::log_debug("InnerTube transcript fetch failed, falling back to yt-dlp...");
    }

    // Fallback to yt-dlp
    let ytdlp_path =
        ytdlp_path.ok_or_else(|| ImportError::Other("yt-dlp not available".to_string()))?;
    let download_url = download_url.ok_or(ImportError::InvalidUrl)?;
    let language_code = get_transcript_language_code(transcript);
    fetch_transcript_text_with_ytdlp(
        ytdlp_path,
        download_url,
        language_code,
        transcript.is_generated,
        include_timestamps,
    )
}

fn fetch_transcript_text_with_ytdlp(
    ytdlp_path: &Path,
    url: &str,
    language_code: &str,
    is_generated: bool,
    include_timestamps: bool,
) -> Result<String, ImportError> {
    let temp_dir = std::env::temp_dir().join(format!("sonarpad_ytdlp_{}", std::process::id()));
    if let Err(err) = std::fs::create_dir_all(&temp_dir) {
        return Err(ImportError::Other(format!(
            "Failed to create temp dir: {err}"
        )));
    }
    let output_template = temp_dir.join("yt_transcript_%(id)s.%(ext)s");
    let output_template_str = output_template.to_string_lossy().to_string();
    let mut cmd = ytdlp_command(ytdlp_path);
    cmd.arg("--no-playlist")
        .arg("--socket-timeout")
        .arg(YTDLP_SOCKET_TIMEOUT_SECS)
        .arg("--skip-download");

    // Use --write-sub for manual subtitles, --write-auto-sub for auto-generated
    if is_generated {
        cmd.arg("--write-auto-sub");
    } else {
        cmd.arg("--write-sub");
    }

    cmd.arg("--sub-lang")
        .arg(language_code)
        .arg("--sub-format")
        .arg("vtt/srt")
        .arg("--no-warnings")
        .arg("--force-ipv4")
        .arg("-o")
        .arg(output_template_str)
        .arg(url);

    let output = cmd
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|err| ImportError::Other(err.to_string()))?;

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    if !output.status.success() {
        let combined = format!("{stdout}\n{stderr}");
        if combined.contains("No subtitles") || combined.contains("no subtitles") {
            cleanup_ytdlp_temp_dir(&temp_dir);
            return Err(ImportError::NoTranscript);
        }
        cleanup_ytdlp_temp_dir(&temp_dir);
        return Err(ImportError::Other(combined.trim().to_string()));
    }

    let mut subtitle_path = parse_ytdlp_subtitle_path(&stdout)
        .or_else(|| parse_ytdlp_subtitle_path(&stderr))
        .filter(|path| path.exists());
    if subtitle_path.is_none() {
        subtitle_path = find_ytdlp_subtitle_in_dir(&temp_dir, language_code);
    }

    let subtitle_path = match subtitle_path {
        Some(path) => path,
        None => {
            cleanup_ytdlp_temp_dir(&temp_dir);
            return Err(ImportError::NoTranscript);
        }
    };

    let content = std::fs::read_to_string(&subtitle_path)
        .map_err(|err| ImportError::Other(err.to_string()))?;
    let text = parse_subtitle_text(&content, include_timestamps);
    if text.trim().is_empty() {
        cleanup_ytdlp_temp_dir(&temp_dir);
        return Err(ImportError::NoTranscript);
    }
    cleanup_ytdlp_temp_dir(&temp_dir);
    Ok(text)
}

fn error_message(language: Language, err: &ImportError) -> String {
    let labels = labels(language);
    match err {
        ImportError::InvalidUrl => labels.invalid_url,
        ImportError::NoTranscript => labels.no_transcript,
        ImportError::Other(msg) => {
            crate::log_debug(&format!("YouTube transcript import error: {}", msg));
            labels.import_error
        }
    }
}

fn clean_transcript_text(text: &str) -> String {
    let trimmed = text.trim_start();
    let cleaned = trimmed.strip_prefix(">>").unwrap_or(trimmed).trim_start();
    let stripped = strip_vtt_inline_tags(cleaned);
    let decoded = decode_html_entities(&stripped);
    decoded.trim().to_string()
}

fn strip_vtt_inline_tags(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    let mut in_tag = false;
    for ch in input.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ => {
                if !in_tag {
                    out.push(ch);
                }
            }
        }
    }
    out
}

fn decode_html_entities(input: &str) -> String {
    input
        .replace("&gt;", ">")
        .replace("&lt;", "<")
        .replace("&amp;", "&")
        .replace("&nbsp;", " ")
}

fn format_timestamp(seconds: f64) -> String {
    let total = seconds.max(0.0).floor() as u64;
    let hours = total / 3600;
    let minutes = (total % 3600) / 60;
    let secs = total % 60;
    if hours > 0 {
        format!("{hours:02}:{minutes:02}:{secs:02}")
    } else {
        format!("{minutes:02}:{secs:02}")
    }
}

fn parse_ytdlp_subtitle_path(output: &str) -> Option<PathBuf> {
    for line in output.lines() {
        if let Some(idx) = line.find("Writing video subtitles to:") {
            let path = line[(idx + "Writing video subtitles to:".len())..].trim();
            if !path.is_empty() {
                return Some(PathBuf::from(path));
            }
        }
    }
    None
}

fn find_ytdlp_subtitle_in_dir(dir: &Path, language_code: &str) -> Option<PathBuf> {
    let entries = std::fs::read_dir(dir).ok()?;
    for entry in entries {
        let entry = match entry {
            Ok(entry) => entry,
            Err(err) => {
                crate::log_debug(&format!("yt-dlp temp dir entry error: {}", err));
                continue;
            }
        };
        let path = entry.path();
        if !path.is_file() {
            continue;
        }
        let name = match path.file_name().and_then(|s| s.to_str()) {
            Some(name) => name,
            None => continue,
        };
        if !name.contains(language_code) {
            continue;
        }
        if name.ends_with(".vtt") || name.ends_with(".srt") {
            return Some(path);
        }
    }
    None
}

fn parse_subtitle_text(content: &str, include_timestamps: bool) -> String {
    // First pass: collect all unique text segments with their timestamps
    let mut segments: Vec<(Option<String>, String)> = Vec::new();
    let mut current_stamp: Option<String> = None;
    let mut seen_texts: std::collections::HashSet<String> = std::collections::HashSet::new();

    for raw in content.lines() {
        let line = raw.trim();
        if line.is_empty() {
            continue;
        }
        if line.starts_with("WEBVTT")
            || line.starts_with("NOTE")
            || line.starts_with("STYLE")
            || line.starts_with("REGION")
            || line.starts_with("Kind:")
            || line.starts_with("Language:")
            || line.starts_with("X-TIMESTAMP-MAP")
        {
            continue;
        }
        if line.contains("-->") {
            current_stamp = parse_subtitle_timestamp(line);
            continue;
        }
        if line.chars().all(|c| c.is_ascii_digit()) {
            continue;
        }
        let preferred = select_vtt_payload_line(line);
        let cleaned = clean_transcript_text(preferred);
        if cleaned.is_empty() {
            continue;
        }

        // Skip if we've already seen this exact text
        if seen_texts.contains(&cleaned) {
            continue;
        }
        seen_texts.insert(cleaned.clone());
        segments.push((current_stamp.clone(), cleaned));
    }

    // Build output, merging consecutive segments without repeating text
    let mut lines_out: Vec<String> = Vec::new();
    let mut current_text = String::new();
    let mut current_ts: Option<String> = None;

    for (stamp, text) in segments {
        // Check if this text is a continuation (previous text is prefix of current)
        if !current_text.is_empty() && text.starts_with(&current_text) {
            // This is a continuation, extract only the new part
            let new_part = text[current_text.len()..].trim();
            if !new_part.is_empty() {
                current_text = text;
            }
        } else if !current_text.is_empty() && current_text.starts_with(&text) {
            // Current text already contains this, skip
            continue;
        } else {
            // New segment, flush previous
            if !current_text.is_empty() {
                if include_timestamps {
                    if let Some(ts) = &current_ts {
                        lines_out.push(format!("[{ts}] {current_text}"));
                    } else {
                        lines_out.push(current_text.clone());
                    }
                } else {
                    lines_out.push(current_text.clone());
                }
            }
            current_text = text;
            current_ts = stamp;
        }
    }

    // Flush last segment
    if !current_text.is_empty() {
        if include_timestamps {
            if let Some(ts) = &current_ts {
                lines_out.push(format!("[{ts}] {current_text}"));
            } else {
                lines_out.push(current_text);
            }
        } else {
            lines_out.push(current_text);
        }
    }

    if include_timestamps {
        collapse_repeated_lines(&lines_out).join("\n")
    } else {
        collapse_repeated_phrases(&lines_out.join(" "))
    }
}

/// Remove consecutive duplicate lines (ignoring timestamps for comparison)
fn collapse_repeated_lines(lines: &[String]) -> Vec<String> {
    if lines.is_empty() {
        return Vec::new();
    }
    let mut result = Vec::with_capacity(lines.len());
    let mut last_text = String::new();

    for line in lines {
        // Extract text without timestamp for comparison
        let text = if line.starts_with('[') {
            line.find("] ").map(|i| &line[i + 2..]).unwrap_or(line)
        } else {
            line.as_str()
        };

        if text != last_text {
            result.push(line.clone());
            last_text = text.to_string();
        }
    }
    result
}

fn select_vtt_payload_line(line: &str) -> &str {
    if let Some(idx) = line.rfind('>') {
        let tail = line[idx + 1..].trim();
        if !tail.is_empty() {
            return tail;
        }
    }
    line
}

fn parse_subtitle_timestamp(line: &str) -> Option<String> {
    let start = line.split("-->").next()?.trim();
    let time = start.split_whitespace().next()?.trim();
    let time = time.split(',').next().unwrap_or(time);
    let parts: Vec<&str> = time.split(':').collect();
    let (hours, minutes, seconds) = match parts.len() {
        3 => {
            let hours = parts[0].parse::<u64>().ok()?;
            let minutes = parts[1].parse::<u64>().ok()?;
            let seconds = parse_seconds(parts[2])?;
            (hours, minutes, seconds)
        }
        2 => {
            let minutes = parts[0].parse::<u64>().ok()?;
            let seconds = parse_seconds(parts[1])?;
            (0, minutes, seconds)
        }
        _ => return None,
    };
    let total = (hours * 3600) + (minutes * 60) + seconds;
    Some(format_timestamp(total as f64))
}

fn parse_seconds(part: &str) -> Option<u64> {
    let main = part.split('.').next().unwrap_or(part);
    main.parse::<u64>().ok()
}

fn collapse_repeated_phrases(text: &str) -> String {
    let tokens: Vec<&str> = text.split_whitespace().collect();
    if tokens.len() < 2 {
        return text.to_string();
    }
    let normalized: Vec<String> = tokens.iter().map(|t| normalize_token(t)).collect();
    let mut out: Vec<&str> = Vec::with_capacity(tokens.len());
    let mut i = 0;
    let max_k = 20usize;
    while i < tokens.len() {
        let mut matched = false;
        let max_len = (tokens.len() - i) / 2;
        let upper = max_k.min(max_len);
        // Check phrases from longest to shortest, including single words (k >= 1)
        for k in (1..=upper).rev() {
            if i + 2 * k <= tokens.len() && normalized[i..i + k] == normalized[i + k..i + 2 * k] {
                out.extend_from_slice(&tokens[i..i + k]);
                i += k;
                // Skip all subsequent repetitions
                while i + k <= tokens.len() && normalized[i - k..i] == normalized[i..i + k] {
                    i += k;
                }
                matched = true;
                break;
            }
        }
        if !matched {
            out.push(tokens[i]);
            i += 1;
        }
    }
    out.join(" ")
}

fn normalize_token(token: &str) -> String {
    let mut out = String::with_capacity(token.len());
    for ch in token.chars() {
        if ch.is_alphanumeric() {
            for lower in ch.to_lowercase() {
                out.push(lower);
            }
        }
    }
    out
}

fn cleanup_ytdlp_temp_dir(dir: &Path) {
    if !dir.exists() {
        return;
    }
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries {
            let entry = match entry {
                Ok(entry) => entry,
                Err(err) => {
                    crate::log_debug(&format!("yt-dlp temp cleanup entry error: {}", err));
                    continue;
                }
            };
            let path = entry.path();
            if path.is_file() {
                crate::log_if_err!(std::fs::remove_file(&path));
            }
        }
    }
    crate::log_if_err!(std::fs::remove_dir(dir));
}

fn fetch_transcript_list_with_retry(
    ytdlp_path: &Path,
    url: &str,
    cancelled: &Arc<AtomicBool>,
) -> Result<Vec<YtSubtitleOption>, ImportError> {
    let mut attempts = 0;
    loop {
        if cancelled.load(Ordering::Relaxed) {
            return Err(ImportError::Other("Cancelled".to_string()));
        }
        match fetch_transcript_list_with_ytdlp(ytdlp_path, url) {
            Ok(list) => return Ok(list),
            Err(e) => {
                attempts += 1;
                if attempts >= 5 {
                    return Err(e);
                }
                match e {
                    ImportError::Other(_) => {
                        std::thread::sleep(std::time::Duration::from_millis(
                            500 * 2u64.pow(attempts - 1),
                        ));
                    }
                    _ => return Err(e),
                }
            }
        }
    }
}

fn fetch_transcript_text_with_retry(
    transcript: &YtSubtitleOption,
    include_timestamps: bool,
    cancelled: &Arc<AtomicBool>,
    ytdlp_path: Option<&Path>,
    download_url: Option<&str>,
) -> Result<String, ImportError> {
    let mut attempts = 0;
    loop {
        if cancelled.load(Ordering::Relaxed) {
            return Err(ImportError::Other("Cancelled".to_string()));
        }
        match fetch_transcript_text_with_fallback(
            transcript,
            include_timestamps,
            ytdlp_path,
            download_url,
        ) {
            Ok(text) => return Ok(text),
            Err(e) => {
                attempts += 1;
                if attempts >= 5 {
                    return Err(e);
                }
                match e {
                    ImportError::Other(_) => {
                        std::thread::sleep(std::time::Duration::from_millis(
                            500 * 2u64.pow(attempts - 1),
                        ));
                    }
                    _ => return Err(e),
                }
            }
        }
    }
}

// ============================================================================
// YouTube InnerTube API (direct access - faster and cleaner transcripts)
// ============================================================================

/// Fetch transcript list using YouTube's InnerTube API directly
/// Returns None if the API fails (will fallback to yt-dlp)
fn fetch_transcript_list_innertube(video_id: &str) -> Option<Vec<YtSubtitleOption>> {
    let client = match reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(15))
        .build()
    {
        Ok(c) => c,
        Err(e) => {
            crate::log_debug(&format!("InnerTube: failed to create client: {}", e));
            return None;
        }
    };

    let payload = serde_json::json!({
        "context": {
            "client": {
                "clientName": INNERTUBE_CLIENT_NAME,
                "clientVersion": INNERTUBE_CLIENT_VERSION,
                "hl": "en",
                "gl": "US"
            }
        },
        "videoId": video_id
    });

    let response = match client
        .post(INNERTUBE_API_URL)
        .json(&payload)
        .header("Content-Type", "application/json")
        .header("User-Agent", INNERTUBE_USER_AGENT)
        .send()
    {
        Ok(r) => r,
        Err(e) => {
            crate::log_debug(&format!("InnerTube: request failed: {}", e));
            return None;
        }
    };

    if !response.status().is_success() {
        crate::log_debug(&format!("InnerTube: HTTP error {}", response.status()));
        return None;
    }

    let data: serde_json::Value = match response.json() {
        Ok(d) => d,
        Err(e) => {
            crate::log_debug(&format!("InnerTube: JSON parse error: {}", e));
            return None;
        }
    };

    // Check for playability status
    if let Some(status) = data.get("playabilityStatus")
        && let Some(reason) = status.get("reason").and_then(|r| r.as_str())
    {
        crate::log_debug(&format!("InnerTube: playability error: {}", reason));
        return None;
    }

    // Extract caption tracks
    let captions = match data.get("captions") {
        Some(c) => c,
        None => {
            crate::log_debug("InnerTube: no captions field in response");
            return None;
        }
    };

    let renderer = match captions.get("playerCaptionsTracklistRenderer") {
        Some(r) => r,
        None => {
            crate::log_debug("InnerTube: no playerCaptionsTracklistRenderer");
            return None;
        }
    };

    let caption_tracks = match renderer.get("captionTracks").and_then(|t| t.as_array()) {
        Some(t) => t,
        None => {
            crate::log_debug("InnerTube: no captionTracks array");
            return None;
        }
    };

    let mut items = Vec::new();
    for track in caption_tracks {
        // Get required fields - skip track if missing
        let language_code = match track.get("languageCode").and_then(|c| c.as_str()) {
            Some(c) => c,
            None => continue,
        };

        let base_url = match track.get("baseUrl").and_then(|u| u.as_str()) {
            Some(u) => u,
            None => continue,
        };

        // Filter out non-subtitle tracks (like live_chat)
        let vss_id = track.get("vssId").and_then(|v| v.as_str()).unwrap_or("");
        if vss_id.contains("live_chat") || language_code.contains("live_chat") {
            crate::log_debug(&format!(
                "InnerTube: skipping non-subtitle track: {}",
                vss_id
            ));
            continue;
        }

        let is_generated = track.get("kind").and_then(|k| k.as_str()) == Some("asr");

        // Extract language name
        let language = track
            .get("name")
            .and_then(|n| {
                n.get("runs")
                    .and_then(|r| r.as_array())
                    .and_then(|arr| arr.first())
                    .and_then(|first| first.get("text"))
                    .and_then(|t| t.as_str())
                    .or_else(|| n.get("simpleText").and_then(|s| s.as_str()))
            })
            .unwrap_or(language_code);

        // Filter out non-language entries
        let lang_lower = language.to_lowercase();
        if lang_lower.contains("live chat")
            || lang_lower.contains("livechat")
            || lang_lower.contains("json")
        {
            crate::log_debug(&format!(
                "InnerTube: skipping non-subtitle language: {}",
                language
            ));
            continue;
        }

        items.push(YtSubtitleOption {
            language: language.to_string(),
            code: format!("{}|{}", language_code, base_url), // Store both code and URL
            is_generated,
        });
    }

    if items.is_empty() {
        crate::log_debug("InnerTube: no valid subtitle tracks found");
        return None;
    }

    crate::log_debug(&format!("InnerTube: found {} subtitle tracks", items.len()));

    // Sort: manual first, then by language
    items.sort_by(|a, b| {
        a.is_generated
            .cmp(&b.is_generated)
            .then(a.language.cmp(&b.language))
    });

    Some(items)
}

/// Fetch transcript text using YouTube's InnerTube API directly
fn fetch_transcript_text_innertube(
    transcript: &YtSubtitleOption,
    include_timestamps: bool,
) -> Option<String> {
    // Extract the base URL from the code field
    let parts: Vec<&str> = transcript.code.splitn(2, '|').collect();
    if parts.len() != 2 {
        return None;
    }
    let base_url = parts[1];

    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(30))
        .build()
        .ok()?;

    let response = client
        .get(base_url)
        .header("User-Agent", "Mozilla/5.0")
        .send()
        .ok()?;

    if !response.status().is_success() {
        return None;
    }

    let xml = response.text().ok()?;
    Some(parse_innertube_transcript(&xml, include_timestamps))
}

/// Parse the srv3 XML format from InnerTube API
fn parse_innertube_transcript(xml: &str, include_timestamps: bool) -> String {
    let mut segments: Vec<(f64, String)> = Vec::new();

    for line in xml.lines() {
        let line = line.trim();
        if !line.starts_with("<p ") || !line.contains("t=") {
            continue;
        }

        // Extract timestamp (in milliseconds)
        let start_ms = extract_xml_attr(line, "t").unwrap_or(0.0);
        let start_sec = start_ms / 1000.0;

        // Extract text from <s> tags or directly from <p>
        let text = extract_transcript_text(line);
        let text = decode_xml_entities(&text);
        let text = text.trim().to_string();

        if !text.is_empty() {
            segments.push((start_sec, text));
        }
    }

    if segments.is_empty() {
        return String::new();
    }

    if include_timestamps {
        segments
            .iter()
            .map(|(start, text)| format!("[{}] {}", format_timestamp(*start), text))
            .collect::<Vec<_>>()
            .join("\n")
    } else {
        let full_text = segments
            .iter()
            .map(|(_, t)| t.as_str())
            .collect::<Vec<_>>()
            .join(" ");
        collapse_repeated_phrases(&full_text)
    }
}

/// Extract text content from a <p> line (handles <s> sub-elements)
fn extract_transcript_text(line: &str) -> String {
    // Find the first '>' which ends the opening <p ...> tag
    let start = match line.find('>') {
        Some(pos) => pos + 1,
        None => return String::new(),
    };

    // Find </p> to get the end
    let end = line.rfind("</p>").unwrap_or(line.len());

    if start >= end {
        return String::new();
    }

    let content = &line[start..end];

    // Remove all tags and keep only text
    let mut result = String::new();
    let mut in_tag = false;

    for c in content.chars() {
        match c {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => result.push(c),
            _ => {}
        }
    }

    result
}

fn extract_xml_attr(line: &str, attr: &str) -> Option<f64> {
    let pattern = format!("{}=\"", attr);
    let start = line.find(&pattern)?;
    let after = &line[start + pattern.len()..];
    let end = after.find('"')?;
    after[..end].parse().ok()
}

fn decode_xml_entities(text: &str) -> String {
    text.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
        .replace("&nbsp;", " ")
        .replace("&apos;", "'")
}

/// Check if a YtSubtitleOption was created from InnerTube API (has URL in code)
fn is_innertube_transcript(transcript: &YtSubtitleOption) -> bool {
    transcript.code.contains('|')
}

/// Get the language code from a transcript (handles both yt-dlp and InnerTube formats)
fn get_transcript_language_code(transcript: &YtSubtitleOption) -> &str {
    if is_innertube_transcript(transcript) {
        transcript
            .code
            .split('|')
            .next()
            .unwrap_or(&transcript.code)
    } else {
        &transcript.code
    }
}
