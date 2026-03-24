use std::cmp::Ordering as CmpOrdering;
use std::collections::HashSet;
use std::io::{BufRead, BufReader};
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Mutex};

use encoding_rs::WINDOWS_1252;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
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
    AppendMenuW, BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, CB_ADDSTRING,
    CB_GETCURSEL, CB_RESETCONTENT, CB_SETCURSEL, CBN_SELCHANGE, CBS_DROPDOWNLIST, CREATESTRUCTW,
    CW_USEDEFAULT, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DispatchMessageW,
    EN_CHANGE, ES_AUTOHSCROLL, ES_MULTILINE, ES_READONLY, GWLP_USERDATA, GetCursorPos,
    GetForegroundWindow, GetWindowLongPtrW, HMENU, HWND_TOPMOST, IDC_ARROW, IDYES, IsChild,
    IsDialogMessageW, IsWindow, LoadCursorW, MB_ICONQUESTION, MB_YESNO, MF_STRING, MSG, PM_REMOVE,
    PeekMessageW, PostMessageW, RegisterClassW, SW_HIDE, SW_SHOW, SWP_NOMOVE, SWP_NOSIZE,
    SWP_SHOWWINDOW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos,
    SetWindowTextW, ShowWindow, TPM_NONOTIFY, TPM_RETURNCMD, TrackPopupMenu, TranslateMessage,
    WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU, WM_CREATE, WM_DESTROY, WM_KEYDOWN,
    WM_NCDESTROY, WM_SETFONT, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::{EM_REPLACESEL, screen_reader_speak, to_wide, to_wide_normalized};
use crate::editor_manager::get_edit_text;
use crate::i18n;
use crate::settings::{Language, StreamFavorite, confirm_title, save_settings, settings_dir};
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
const STREAM_ID_URL: usize = 9311;
const STREAM_ID_FORMAT: usize = 9312;
const STREAM_ID_OK: usize = 9313;
const STREAM_ID_CANCEL: usize = 9314;
const STREAM_ID_DIRECT_PLAY: usize = 9315;
const STREAM_ID_FAVORITES: usize = 9316;

#[inline]
fn ignore_bool(_value: bool) {}
const STREAM_ID_QUALITY: usize = 9317;
const STREAM_TRACK_ID_COMBO: usize = 9321;
const STREAM_TRACK_ID_OK: usize = 9322;
const STREAM_TRACK_ID_CANCEL: usize = 9323;
const STREAM_DIALOG_CLASS_NAME: &str = "SonarpadStreamAudio";
const STREAM_TRACK_DIALOG_CLASS_NAME: &str = "SonarpadStreamAudioTrack";
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
const STREAM_DOWNLOAD_STALL_SECS: u64 = 180;
const STREAM_POST_100_GRACE_SECS: u64 = 25;
const STREAM_RETRY_TIMEOUT_SECS: u64 = 120;
static YTDLP_UPDATE_CHECKED: AtomicBool = AtomicBool::new(false);
static PENDING_STREAM_REOPEN_CONTEXT: Mutex<Option<crate::YouTubeReturnContext>> = Mutex::new(None);

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
    ytdlp_found_in_path: String,
    ytdlp_path_update_download_local: String,
    ytdlp_path_update_keep_system: String,
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
        ytdlp_found_in_path: i18n::tr(language, "youtube.ytdlp_found_in_path"),
        ytdlp_path_update_download_local: i18n::tr(
            language,
            "youtube.ytdlp_path_update_download_local",
        ),
        ytdlp_path_update_keep_system: i18n::tr(language, "youtube.ytdlp_path_update_keep_system"),
    }
}

fn ytdlp_command(path: &Path) -> Command {
    let mut cmd = Command::new(path);
    cmd.stdin(Stdio::null());
    cmd.creation_flags(CREATE_NO_WINDOW.0);
    cmd
}

fn youtube_ui_language_code(language: Language) -> &'static str {
    match language {
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
        Language::Russian => "ru",
        Language::Chinese => "zh",
    }
}

fn ytdlp_debug_enabled() -> bool {
    std::env::var("SONARPAD_YTDLP_DEBUG")
        .map(|v| {
            let v = v.trim().to_ascii_lowercase();
            v == "1" || v == "true" || v == "yes" || v == "on"
        })
        .unwrap_or(false)
}

fn post_focus_editor(parent: HWND) {
    unsafe {
        if let Err(e) = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0)) {
            crate::log_debug(&format!("Failed to post WM_FOCUS_EDITOR: {}", e));
        }
    }
}

pub fn import_youtube_transcript(parent: HWND) {
    let (language, include_timestamps) = {
        with_state(parent, |state| {
            (
                state.settings.language,
                state.settings.youtube_include_timestamps,
            )
        })
        .unwrap_or((Language::Italian, true))
    };
    let Some(result) = show_import_dialog(parent, language, include_timestamps) else {
        post_focus_editor(parent);
        return;
    };
    {
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
            post_focus_editor(parent);
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
        post_focus_editor(parent);
    }
}

fn show_import_dialog(
    parent: HWND,
    language: Language,
    include_timestamps: bool,
) -> Option<ImportResult> {
    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
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
    crate::register_class_w_safe(&wc);

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
    crate::watchdog::enter_modal_dialog();
    loop {
        if !crate::is_window_handle_valid(hwnd) {
            break;
        }
        let res = crate::get_message_w_safe(&mut msg, HWND(0), 0, 0);
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
    unsafe {
        crate::panic_guard::guard(
            "import_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || import_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn import_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
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
                        let idx =
                            SendMessageW(state.lang_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        if idx < 0 || idx as usize >= state.transcripts.len() {
                            return;
                        }
                        let include_timestamps =
                            SendMessageW(state.timestamp_check, BM_GETCHECK, WPARAM(0), LPARAM(0))
                                .0
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
                let result = Box::from_raw(lparam.0 as *mut LoadResult);
                finish_load_languages(hwnd, *result);
                LRESULT(0)
            }
            WM_YT_LOAD_CANCEL => {
                reset_languages_loading_state(hwnd);
                LRESULT(0)
            }
            WM_YT_TEXT_COMPLETE => {
                let result = Box::from_raw(lparam.0 as *mut TextLoadResult);
                finish_load_text(hwnd, *result);
                LRESULT(0)
            }
            WM_YT_TEXT_CANCEL => {
                reset_text_loading_state(hwnd);
                LRESULT(0)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    if let Err(_e) = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_CANCEL), LPARAM(0))
                    {
                        crate::log_debug(&format!("Error: {:?}", _e));
                    }
                    return LRESULT(0);
                }
                if wparam.0 as u32 == VK_RETURN.0 as u32 {
                    let focus = GetFocus();
                    let url_edit =
                        with_import_state(hwnd, |state| state.url_edit).unwrap_or(HWND(0));
                    if focus == url_edit {
                        if let Err(_e) =
                            PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_LOAD), LPARAM(0))
                        {
                            crate::log_debug(&format!("Error: {:?}", _e));
                        }
                        return LRESULT(0);
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
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
                            NotifyWinEvent(
                                EVENT_OBJECT_FOCUS,
                                hwnd_edit,
                                OBJID_CLIENT,
                                CHILDID_SELF,
                            );
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
}

fn with_import_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut ImportState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut ImportState;
    crate::with_raw_mut_ptr_safe(ptr, f)
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

    if with_import_state(hwnd, |state| {
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
            let ytdlp_path = match ensure_ytdlp_available(hwnd, language, &labels_data, None) {
                Ok(Some(path)) => path,
                Ok(None) => {
                    return LoadResult {
                        transcripts: Vec::new(),
                        error: None,
                    };
                }
                Err(err) => {
                    crate::log_debug(&format!("yt-dlp download failed: {}", err));
                    show_error(hwnd, language, &labels_data.ytdlp_download_failed);
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

    let state_ok = with_import_state(hwnd, |state| {
        edit = state.url_edit;
        ok_button = state.ok_button;
        load_button = state.load_button;
        combo = state.lang_combo;
        timestamp = state.timestamp_check;
        status = state.status_label;
        language = state.language;
        state.loading = false;
    })
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

    if with_import_state(hwnd, |state| {
        state.transcripts = result.transcripts;
    })
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

    let state_ok = with_import_state(hwnd, |state| {
        edit = state.url_edit;
        ok_button = state.ok_button;
        load_button = state.load_button;
        combo = state.lang_combo;
        timestamp = state.timestamp_check;
        status = state.status_label;
        language = state.language;
        state.loading = false;
    })
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

    let state_ok = with_import_state(hwnd, |state| {
        edit = state.url_edit;
        ok_button = state.ok_button;
        load_button = state.load_button;
        combo = state.lang_combo;
        timestamp = state.timestamp_check;
        status = state.status_label;
        state.loading = false;
    })
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

    if with_import_state(hwnd, |state| {
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

            let mut fetch_result = fetch_transcript_text_with_retry(
                &transcript,
                include_timestamps,
                &cancelled,
                None,
                Some(&download_url),
            );
            if matches!(fetch_result, Err(ImportError::YtdlpUnavailable)) {
                let ytdlp_path = match ensure_ytdlp_available(hwnd, language, &labels_data, None) {
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
                        show_error(hwnd, language, &labels_data.ytdlp_download_failed);
                        return TextLoadResult {
                            text: String::new(),
                            include_timestamps,
                            error: None,
                        };
                    }
                };
                fetch_result = fetch_transcript_text_with_retry(
                    &transcript,
                    include_timestamps,
                    &cancelled,
                    Some(&ytdlp_path),
                    Some(&download_url),
                );
            }

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

    let state_ok = with_import_state(hwnd, |state| {
        edit = state.url_edit;
        ok_button = state.ok_button;
        load_button = state.load_button;
        combo = state.lang_combo;
        timestamp = state.timestamp_check;
        status = state.status_label;
        language = state.language;
        state.loading = false;
    })
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

    if with_import_state(hwnd, |state| {
        *state.result.lock().unwrap_or_else(|e| e.into_inner()) = Some(ImportResult {
            text: result.text,
            include_timestamps: result.include_timestamps,
        });
    })
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

fn tr_or(language: Language, key: &str, fallback: &str) -> String {
    let translated = i18n::tr(language, key);
    if translated == key {
        fallback.to_string()
    } else {
        translated
    }
}

fn load_stream_favorites(parent: HWND) -> Vec<StreamFavorite> {
    with_state(parent, |state| state.settings.stream_favorites.clone()).unwrap_or_default()
}

fn refill_stream_favorites_combo(state: &StreamDialogState, selection: Option<usize>) {
    unsafe {
        SendMessageW(state.favorites_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for favorite in &state.favorites {
            let display = if favorite.label.trim().is_empty() {
                favorite.url.clone()
            } else {
                favorite.label.clone()
            };
            let wide = to_wide(&display);
            SendMessageW(
                state.favorites_combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(wide.as_ptr() as isize),
            );
        }
        if !state.favorites.is_empty() {
            let max_index = state.favorites.len() - 1;
            let selected = selection.unwrap_or(0).min(max_index);
            SendMessageW(
                state.favorites_combo,
                CB_SETCURSEL,
                WPARAM(selected),
                LPARAM(0),
            );
        }
    }
}

fn add_stream_favorite(parent: HWND, label: String, url: String) {
    let normalized_url = normalize_youtube_collection_url(&url).unwrap_or(url);
    if !is_youtube_collection_url(&normalized_url) {
        return;
    }
    let favorite = StreamFavorite {
        label: label.trim().to_string(),
        url: normalized_url,
    };
    if favorite.url.is_empty() {
        return;
    }
    if with_state(parent, |state| {
        state
            .settings
            .stream_favorites
            .retain(|existing| !existing.url.eq_ignore_ascii_case(&favorite.url));
        state.settings.stream_favorites.insert(0, favorite);
        save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to save stream favorite");
    }
}

fn remove_stream_favorite(parent: HWND, url: &str) {
    if with_state(parent, |state| {
        state
            .settings
            .stream_favorites
            .retain(|favorite| !favorite.url.eq_ignore_ascii_case(url));
        save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to remove stream favorite");
    }
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

    // URL field: treat input as URL and strip hidden/newline whitespace that may split long links.
    let compact: String = trimmed.chars().filter(|ch| !ch.is_whitespace()).collect();

    let normalized = compact
        .replace("&amp;", "&")
        .replace("\\?", "?")
        .replace("\\=", "=")
        .trim()
        .trim_matches('"')
        .trim_matches('\'')
        .to_string();
    if normalized.is_empty() {
        return None;
    }

    // Unwrap common redirect/wrapper query params first (q=, url=, u=).
    if let Ok(url) = Url::parse(&normalized) {
        for (key, value) in url.query_pairs() {
            let k = key.as_ref();
            if k.eq_ignore_ascii_case("url")
                || k.eq_ignore_ascii_case("u")
                || k.eq_ignore_ascii_case("q")
            {
                let v = value.trim();
                if v.starts_with("http://") || v.starts_with("https://") {
                    if let Some(unwrapped) = normalize_youtube_input_for_download(v) {
                        return Some(unwrapped);
                    }
                } else if v.starts_with('/') {
                    // e.g. youtube attribution links: u=/watch?v=...
                    let joined = format!("https://www.youtube.com{v}");
                    if let Some(unwrapped) = normalize_youtube_input_for_download(&joined) {
                        return Some(unwrapped);
                    }
                }
            }
        }
    }

    // Only normalize to watch?v= for actual YouTube hosts.
    if let Ok(url) = Url::parse(&normalized)
        && let Some(host) = url.host_str()
        && (host.eq_ignore_ascii_case("youtu.be")
            || host.eq_ignore_ascii_case("youtube.com")
            || host.ends_with(".youtube.com"))
        && let Some(id) = extract_video_id(&normalized)
    {
        return Some(format!("https://www.youtube.com/watch?v={id}"));
    }
    if normalized.starts_with("http://") || normalized.starts_with("https://") {
        Some(normalized)
    } else {
        Some(format!("https://{normalized}"))
    }
}

fn normalize_youtube_collection_url(input: &str) -> Option<String> {
    let normalized = normalize_youtube_input_for_download(input)?;
    let Ok(mut url) = Url::parse(&normalized) else {
        return Some(normalized);
    };
    let Some(host) = url.host_str() else {
        return Some(normalized);
    };
    if !(host.eq_ignore_ascii_case("youtube.com") || host.ends_with(".youtube.com")) {
        return Some(normalized);
    }
    let path = url.path().trim_end_matches('/').to_string();
    let segments: Vec<&str> = path
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect();
    let should_force_videos = match segments.as_slice() {
        ["channel", _] | ["user", _] | ["c", _] => true,
        [segment] => segment.starts_with('@'),
        _ => false,
    };
    if should_force_videos && !path.ends_with("/videos") {
        url.set_path(&format!("{path}/videos"));
    }
    Some(url.to_string())
}

fn is_youtube_collection_url(input: &str) -> bool {
    let Some(normalized) = normalize_youtube_collection_url(input) else {
        return false;
    };
    let Ok(url) = Url::parse(&normalized) else {
        return false;
    };
    let Some(host) = url.host_str() else {
        return false;
    };
    if !(host.eq_ignore_ascii_case("youtube.com") || host.ends_with(".youtube.com")) {
        return false;
    }
    let path = url.path().trim_end_matches('/');
    path.starts_with("/channel/")
        || path.starts_with("/user/")
        || path.starts_with("/c/")
        || path.starts_with("/@")
        || path == "/playlist"
}

fn is_youtube_channel_url(input: &str) -> bool {
    let Some(normalized) = normalize_youtube_collection_url(input) else {
        return false;
    };
    let Ok(url) = Url::parse(&normalized) else {
        return false;
    };
    let Some(host) = url.host_str() else {
        return false;
    };
    if !(host.eq_ignore_ascii_case("youtube.com") || host.ends_with(".youtube.com")) {
        return false;
    }
    let path = url.path().trim_end_matches('/');
    path.starts_with("/channel/")
        || path.starts_with("/user/")
        || path.starts_with("/c/")
        || path.starts_with("/@")
}

#[derive(Clone)]
struct StreamCollectionEntry {
    label: String,
    url: String,
}

struct ResolvedStreamSelection {
    url: String,
    collection_url: Option<String>,
    collection_page: Option<usize>,
}

const STREAM_SELECTION_PAGE_SIZE: usize = 20;
const STREAM_SELECTION_LOAD_MORE_KEY: &str = "stream_audio.load_more_videos";
const STREAM_SELECTION_PREVIOUS_KEY: &str = "stream_audio.previous_results";

fn stream_entry_url(entry: &serde_json::Value) -> Option<String> {
    entry
        .get("webpage_url")
        .and_then(|value| value.as_str())
        .map(|value| value.to_string())
        .or_else(|| {
            entry
                .get("url")
                .and_then(|value| value.as_str())
                .map(|value| {
                    if value.starts_with("http://") || value.starts_with("https://") {
                        value.to_string()
                    } else if value.starts_with('@')
                        || value.starts_with("channel/")
                        || value.starts_with("user/")
                        || value.starts_with("c/")
                        || value.starts_with("playlist")
                    {
                        format!("https://www.youtube.com/{value}")
                    } else {
                        format!("https://www.youtube.com/watch?v={value}")
                    }
                })
        })
}

fn collect_stream_collection_entries(
    entries: &[serde_json::Value],
    language: Language,
) -> Vec<StreamCollectionEntry> {
    let mut out = Vec::new();
    for entry in entries.iter() {
        let Some(video_url) = stream_entry_url(entry) else {
            continue;
        };
        out.push(StreamCollectionEntry {
            label: format_stream_entry_label(entry, language),
            url: video_url,
        });
    }
    out
}

fn probe_youtube_collection_entries(
    ytdlp_path: &Path,
    url: &str,
    language: Language,
    page: usize,
) -> Result<(Vec<StreamCollectionEntry>, bool), String> {
    let start = page * STREAM_SELECTION_PAGE_SIZE + 1;
    let end = start + STREAM_SELECTION_PAGE_SIZE;
    let target_url = normalize_youtube_collection_url(url).unwrap_or_else(|| url.to_string());
    let output = ytdlp_command(ytdlp_path)
        .arg("--flat-playlist")
        .arg("--dump-single-json")
        .arg("--playlist-start")
        .arg(start.to_string())
        .arg("--playlist-end")
        .arg(end.to_string())
        .arg("--no-warnings")
        .arg("--skip-download")
        .arg("--extractor-args")
        .arg(format!(
            "youtube:lang={}",
            youtube_ui_language_code(language)
        ))
        .arg("--")
        .arg(&target_url)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|e| e.to_string())?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if stderr.is_empty() {
            "yt-dlp collection probe failed".to_string()
        } else {
            stderr
        });
    }

    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).map_err(|e| e.to_string())?;
    let Some(entries) = json.get("entries").and_then(|value| value.as_array()) else {
        return Ok((Vec::new(), false));
    };

    let mut out = collect_stream_collection_entries(entries, language);
    let has_more = out.len() > STREAM_SELECTION_PAGE_SIZE;
    if has_more {
        out.truncate(STREAM_SELECTION_PAGE_SIZE);
    }
    Ok((out, has_more))
}

fn probe_youtube_search_entries(
    ytdlp_path: &Path,
    query: &str,
    language: Language,
    page: usize,
) -> Result<(Vec<StreamCollectionEntry>, bool), String> {
    let limit = (page + 1) * STREAM_SELECTION_PAGE_SIZE + 1;
    let encoded_query: String = url::form_urlencoded::byte_serialize(query.as_bytes()).collect();
    let search_url = format!("https://www.youtube.com/results?search_query={encoded_query}");
    let output = ytdlp_command(ytdlp_path)
        .arg("--flat-playlist")
        .arg("--dump-single-json")
        .arg("--playlist-end")
        .arg(limit.to_string())
        .arg("--no-warnings")
        .arg("--skip-download")
        .arg("--extractor-args")
        .arg(format!(
            "youtube:lang={}",
            youtube_ui_language_code(language)
        ))
        .arg("--")
        .arg(&search_url)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|e| e.to_string())?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if stderr.is_empty() {
            "yt-dlp search probe failed".to_string()
        } else {
            stderr
        });
    }

    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).map_err(|e| e.to_string())?;
    let Some(entries) = json.get("entries").and_then(|value| value.as_array()) else {
        return Ok((Vec::new(), false));
    };

    let skip = page * STREAM_SELECTION_PAGE_SIZE;
    let has_more = entries.len() > skip + STREAM_SELECTION_PAGE_SIZE;
    let page_entries: Vec<serde_json::Value> = entries
        .iter()
        .skip(skip)
        .take(STREAM_SELECTION_PAGE_SIZE)
        .cloned()
        .collect();
    let mut collected = collect_stream_collection_entries(&page_entries, language);
    collected.sort_by_key(|entry| !is_youtube_collection_url(&entry.url));
    Ok((collected, has_more))
}

fn choose_stream_collection_entry_page(
    parent: HWND,
    language: Language,
    entries: &[StreamCollectionEntry],
    has_previous: bool,
    has_more: bool,
) -> Option<String> {
    let mut labels =
        Vec::with_capacity(entries.len() + usize::from(has_previous) + usize::from(has_more));
    if has_previous {
        labels.push(i18n::tr(language, STREAM_SELECTION_PREVIOUS_KEY));
    }
    labels.extend(entries.iter().map(|entry| entry.label.clone()));
    if has_more {
        labels.push(i18n::tr(language, STREAM_SELECTION_LOAD_MORE_KEY));
    }
    let favorite_candidates = Arc::new(entries.to_vec());
    let context_action_enabled = {
        let favorite_candidates = Arc::clone(&favorite_candidates);
        Arc::new(move |selected: &str| {
            favorite_candidates
                .iter()
                .find(|entry| entry.label == selected)
                .map(|entry| is_youtube_collection_url(&entry.url))
                .unwrap_or(false)
        })
    };
    let context_action_handler = {
        let favorite_candidates = Arc::clone(&favorite_candidates);
        Arc::new(move |selected: String| {
            if let Some(entry) = favorite_candidates
                .iter()
                .find(|entry| entry.label == selected && is_youtube_collection_url(&entry.url))
            {
                add_stream_favorite(parent, entry.label.clone(), entry.url.clone());
            }
        })
    };
    crate::app_windows::interpreter_select_window::select_interpreter_with_context_action_without_parent_restore_on_accept(
        parent,
        labels,
        language,
        i18n::tr(language, "stream_audio.prompt_title"),
        None,
        crate::app_windows::interpreter_select_window::InterpreterContextAction {
            label: tr_or(
                language,
                "stream_audio.add_to_favorites",
                "Add to favorites",
            ),
            enabled: context_action_enabled,
            handler: context_action_handler,
        },
    )
}

fn choose_youtube_collection_entry(
    parent: HWND,
    language: Language,
    ytdlp_path: &Path,
    url: &str,
    initial_page: Option<usize>,
    initial_progress: Option<HWND>,
) -> Result<Option<ResolvedStreamSelection>, String> {
    if !is_youtube_collection_url(url) {
        return Ok(Some(ResolvedStreamSelection {
            url: url.to_string(),
            collection_url: None,
            collection_page: None,
        }));
    }

    let mut page = initial_page.unwrap_or(0);
    let mut shared_progress = initial_progress;
    loop {
        let progress = if let Some(existing) = shared_progress.take() {
            report_progress_status(existing, &i18n::tr(language, "podcasts.loading"));
            keep_stream_progress_focus(existing);
            existing
        } else {
            open_progress_dialog(
                parent,
                language,
                "stream_audio.progress_title",
                "podcasts.loading",
                false,
            )
        };
        let ytdlp = ytdlp_path.to_path_buf();
        let url_owned = url.to_string();
        let worker = std::thread::spawn(move || {
            probe_youtube_collection_entries(&ytdlp, &url_owned, language, page)
        });
        let mut last_focus_keepalive = std::time::Instant::now();
        while !worker.is_finished() {
            ignore_bool(pump_messages_detect_stream_cancel(parent, progress));
            if last_focus_keepalive.elapsed() > std::time::Duration::from_millis(300) {
                keep_stream_progress_focus(progress);
                last_focus_keepalive = std::time::Instant::now();
            }
            std::thread::sleep(std::time::Duration::from_millis(30));
        }
        let (entries, has_more) = match worker.join() {
            Ok(result) => result?,
            Err(_) => return Err("yt-dlp collection probe worker failed".to_string()),
        };
        crate::log_debug(&format!(
            "stream transition [collection_probe.completed]: page={} entries={} has_more={}",
            page,
            entries.len(),
            has_more
        ));
        if !entries.is_empty() {
            crate::app_windows::podcast_save_window::suppress_parent_restore_on_close(progress);
        }
        close_progress_dialog(progress);
        if entries.is_empty() {
            crate::log_debug(&format!(
                "stream transition [collection_probe.empty]: page={} url={}",
                page, url
            ));
            return if page == 0 {
                Err(i18n::tr(language, "stream_audio.no_matching_videos"))
            } else {
                Ok(None)
            };
        }

        let Some(selected) =
            choose_stream_collection_entry_page(parent, language, &entries, page > 0, has_more)
        else {
            return Ok(None);
        };
        if selected == i18n::tr(language, STREAM_SELECTION_LOAD_MORE_KEY) {
            crate::log_debug(&format!(
                "stream transition [collection_probe.load_more]: current_page={} next_page={}",
                page,
                page + 1
            ));
            page += 1;
            continue;
        }
        if selected == i18n::tr(language, STREAM_SELECTION_PREVIOUS_KEY) {
            crate::log_debug(&format!(
                "stream transition [collection_probe.previous]: current_page={} next_page={}",
                page,
                page.saturating_sub(1)
            ));
            page = page.saturating_sub(1);
            continue;
        }
        let selected_url = entries
            .into_iter()
            .find(|entry| entry.label == selected)
            .map(|entry| entry.url);
        if let Some(selected_url) = selected_url {
            if is_youtube_collection_url(&selected_url) {
                crate::log_debug(&format!(
                    "stream transition [collection_probe.selected_collection]: url={}",
                    selected_url
                ));
                return choose_youtube_collection_entry(
                    parent,
                    language,
                    ytdlp_path,
                    &selected_url,
                    None,
                    None,
                );
            }
            crate::log_debug(&format!(
                "stream transition [collection_probe.selected_video]: page={} url={}",
                page, selected_url
            ));
            return Ok(Some(ResolvedStreamSelection {
                url: selected_url,
                collection_url: Some(url.to_string()),
                collection_page: Some(page),
            }));
        }
        return Ok(None);
    }
}

fn choose_youtube_search_entry(
    parent: HWND,
    language: Language,
    ytdlp_path: &Path,
    query: &str,
    initial_progress: Option<HWND>,
) -> Result<Option<ResolvedStreamSelection>, String> {
    let mut page = 0usize;
    let mut shared_progress = initial_progress;
    loop {
        let progress = if let Some(existing) = shared_progress.take() {
            report_progress_status(existing, &i18n::tr(language, "podcasts.loading"));
            keep_stream_progress_focus(existing);
            existing
        } else {
            open_progress_dialog(
                parent,
                language,
                "stream_audio.progress_title",
                "podcasts.loading",
                false,
            )
        };
        let ytdlp = ytdlp_path.to_path_buf();
        let query_owned = query.to_string();
        let worker = std::thread::spawn(move || {
            probe_youtube_search_entries(&ytdlp, &query_owned, language, page)
        });
        let mut last_focus_keepalive = std::time::Instant::now();
        while !worker.is_finished() {
            ignore_bool(pump_messages_detect_stream_cancel(parent, progress));
            if last_focus_keepalive.elapsed() > std::time::Duration::from_millis(300) {
                keep_stream_progress_focus(progress);
                last_focus_keepalive = std::time::Instant::now();
            }
            std::thread::sleep(std::time::Duration::from_millis(30));
        }
        let (entries, has_more) = match worker.join() {
            Ok(result) => result?,
            Err(_) => return Err("yt-dlp search probe worker failed".to_string()),
        };
        crate::log_debug(&format!(
            "stream transition [search_probe.completed]: page={} entries={} has_more={}",
            page,
            entries.len(),
            has_more
        ));
        if !entries.is_empty() {
            crate::app_windows::podcast_save_window::suppress_parent_restore_on_close(progress);
        }
        close_progress_dialog(progress);
        if entries.is_empty() {
            return if page == 0 {
                Err(i18n::tr(language, "stream_audio.no_matching_videos"))
            } else {
                Ok(None)
            };
        }

        let Some(selected) =
            choose_stream_collection_entry_page(parent, language, &entries, page > 0, has_more)
        else {
            return Ok(None);
        };
        if selected == i18n::tr(language, STREAM_SELECTION_LOAD_MORE_KEY) {
            page += 1;
            continue;
        }
        if selected == i18n::tr(language, STREAM_SELECTION_PREVIOUS_KEY) {
            page = page.saturating_sub(1);
            continue;
        }
        let selected_url = entries
            .into_iter()
            .find(|entry| entry.label == selected)
            .map(|entry| entry.url);
        if let Some(selected_url) = selected_url {
            if is_youtube_collection_url(&selected_url) {
                crate::log_debug(&format!(
                    "stream transition [search_probe.selected_collection]: url={}",
                    selected_url
                ));
                return choose_youtube_collection_entry(
                    parent,
                    language,
                    ytdlp_path,
                    &selected_url,
                    None,
                    None,
                );
            }
            return Ok(Some(ResolvedStreamSelection {
                url: selected_url,
                collection_url: None,
                collection_page: None,
            }));
        }
        return Ok(None);
    }
}
fn resolve_stream_input_url(
    parent: HWND,
    language: Language,
    ytdlp_path: &Path,
    input: &str,
    initial_collection_page: Option<usize>,
    initial_progress: Option<HWND>,
) -> Result<Option<ResolvedStreamSelection>, String> {
    if looks_like_valid_stream_url(input) {
        return choose_youtube_collection_entry(
            parent,
            language,
            ytdlp_path,
            input,
            initial_collection_page,
            initial_progress,
        );
    }
    choose_youtube_search_entry(parent, language, ytdlp_path, input, initial_progress)
}

fn is_members_only_stream_error(err: &str) -> bool {
    let err_lc = err.to_ascii_lowercase();
    err_lc.contains("members-only")
        || err_lc.contains("members only")
        || err_lc.contains("join this channel to get access to members-only content")
}

fn members_only_stream_message(language: Language) -> String {
    format!(
        "{} {}",
        i18n::tr(language, "stream_audio.members_only_video"),
        i18n::tr(language, "stream_audio.choose_another_video")
    )
}
fn looks_like_valid_stream_url(input: &str) -> bool {
    let Some(normalized) = normalize_youtube_input_for_download(input) else {
        return false;
    };
    let Ok(url) = Url::parse(&normalized) else {
        return false;
    };
    let scheme = url.scheme();
    if scheme != "http" && scheme != "https" {
        return false;
    }
    let Some(host) = url.host_str() else {
        return false;
    };
    let host_lc = host.to_ascii_lowercase();
    host_lc == "localhost"
        || host_lc.contains('.')
        || host_lc.chars().all(|c| c.is_ascii_digit() || c == '.')
        || (host_lc.starts_with('[') && host_lc.ends_with(']'))
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum StreamOutputFormat {
    Auto,
    Mp3,
    M4a,
    Opus,
    Ogg,
    Wav,
    Flac,
    Mp4,
}

#[derive(Clone, Copy, Debug)]
enum StreamQualitySelection {
    Original,
    BitrateKbps(u32),
    Mp4High,
    Mp4Medium,
    Mp4Low,
}

#[derive(Clone)]
struct StreamDialogResult {
    url: String,
    format: StreamOutputFormat,
    quality: StreamQualitySelection,
    direct_play: bool,
    reopen_collection_page: Option<usize>,
}

#[derive(Clone)]
struct StreamAudioTrack {
    format_id: String,
    label: String,
}

struct StreamDialogInit {
    parent: HWND,
    language: Language,
    default_format: StreamOutputFormat,
    result: Arc<Mutex<Option<StreamDialogResult>>>,
}

struct StreamDialogState {
    parent: HWND,
    language: Language,
    url_edit: HWND,
    favorites_combo: HWND,
    favorites: Vec<StreamFavorite>,
    format_combo: HWND,
    quality_combo: HWND,
    direct_play_check: HWND,
    ok_button: HWND,
    result: Arc<Mutex<Option<StreamDialogResult>>>,
}

struct StreamTrackDialogInit {
    parent: HWND,
    language: Language,
    tracks: Vec<StreamAudioTrack>,
    result: Arc<Mutex<Option<Option<String>>>>,
}

struct StreamTrackDialogState {
    parent: HWND,
    language: Language,
    combo: HWND,
    ok_button: HWND,
    tracks: Vec<StreamAudioTrack>,
    result: Arc<Mutex<Option<Option<String>>>>,
}

impl StreamOutputFormat {
    fn combo_items(language: Language) -> Vec<(String, StreamOutputFormat)> {
        vec![
            (
                i18n::tr(language, "stream_audio.format.auto"),
                StreamOutputFormat::Auto,
            ),
            ("mp4".to_string(), StreamOutputFormat::Mp4),
            ("mp3".to_string(), StreamOutputFormat::Mp3),
            ("m4a".to_string(), StreamOutputFormat::M4a),
            ("opus".to_string(), StreamOutputFormat::Opus),
            ("ogg".to_string(), StreamOutputFormat::Ogg),
            ("wav".to_string(), StreamOutputFormat::Wav),
            ("flac".to_string(), StreamOutputFormat::Flac),
        ]
    }

    fn as_audio_convert_settings(
        self,
        quality: StreamQualitySelection,
    ) -> Option<crate::ffmpeg_export::ConvertAudioSettings> {
        use crate::ffmpeg_export::{
            ConvertAudioFormat as F, ConvertAudioQuality as Q, ConvertAudioSettings,
        };
        match self {
            StreamOutputFormat::Mp3 => Some(ConvertAudioSettings {
                format: F::Mp3,
                quality: match quality {
                    StreamQualitySelection::BitrateKbps(kbps) => Q::BitrateKbps(kbps),
                    StreamQualitySelection::Original => Q::None,
                    _ => Q::BitrateKbps(192),
                },
            }),
            StreamOutputFormat::M4a => Some(ConvertAudioSettings {
                format: F::Aac,
                quality: Q::BitrateKbps(192),
            }),
            StreamOutputFormat::Opus => Some(ConvertAudioSettings {
                format: F::Opus,
                quality: Q::BitrateKbps(160),
            }),
            StreamOutputFormat::Ogg => Some(ConvertAudioSettings {
                format: F::Ogg,
                quality: Q::OggQuality(6),
            }),
            StreamOutputFormat::Flac => Some(ConvertAudioSettings {
                format: F::Flac,
                quality: Q::FlacCompression(5),
            }),
            StreamOutputFormat::Wav => Some(ConvertAudioSettings {
                format: F::Wav,
                quality: Q::None,
            }),
            StreamOutputFormat::Auto | StreamOutputFormat::Mp4 => None,
        }
    }

    fn extension(self) -> Option<&'static str> {
        match self {
            StreamOutputFormat::Mp3 => Some("mp3"),
            StreamOutputFormat::M4a => Some("m4a"),
            StreamOutputFormat::Opus => Some("opus"),
            StreamOutputFormat::Ogg => Some("ogg"),
            StreamOutputFormat::Wav => Some("wav"),
            StreamOutputFormat::Flac => Some("flac"),
            StreamOutputFormat::Mp4 => Some("mp4"),
            StreamOutputFormat::Auto => None,
        }
    }

    fn from_settings_value(value: &str) -> Self {
        match value.trim().to_ascii_lowercase().as_str() {
            "mp4" => StreamOutputFormat::Mp4,
            "mp3" => StreamOutputFormat::Mp3,
            "m4a" => StreamOutputFormat::M4a,
            "opus" => StreamOutputFormat::Opus,
            "ogg" => StreamOutputFormat::Ogg,
            "wav" => StreamOutputFormat::Wav,
            "flac" => StreamOutputFormat::Flac,
            _ => StreamOutputFormat::Auto,
        }
    }

    fn settings_value(self) -> &'static str {
        match self {
            StreamOutputFormat::Auto => "auto",
            StreamOutputFormat::Mp4 => "mp4",
            StreamOutputFormat::Mp3 => "mp3",
            StreamOutputFormat::M4a => "m4a",
            StreamOutputFormat::Opus => "opus",
            StreamOutputFormat::Ogg => "ogg",
            StreamOutputFormat::Wav => "wav",
            StreamOutputFormat::Flac => "flac",
        }
    }
}

fn stream_quality_items(
    language: Language,
    format: StreamOutputFormat,
) -> Vec<(String, StreamQualitySelection)> {
    match format {
        StreamOutputFormat::Mp3 => {
            let mut items = vec![(
                i18n::tr(language, "stream_audio.quality.original"),
                StreamQualitySelection::Original,
            )];
            for kbps in [64u32, 80, 96, 128, 160, 192, 196, 224, 250, 256, 320] {
                items.push((
                    format!("{kbps} kbps"),
                    StreamQualitySelection::BitrateKbps(kbps),
                ));
            }
            items
        }
        StreamOutputFormat::Mp4 => vec![
            (
                i18n::tr(language, "stream_audio.quality.original"),
                StreamQualitySelection::Original,
            ),
            (
                i18n::tr(language, "stream_audio.quality.high"),
                StreamQualitySelection::Mp4High,
            ),
            (
                i18n::tr(language, "stream_audio.quality.medium"),
                StreamQualitySelection::Mp4Medium,
            ),
            (
                i18n::tr(language, "stream_audio.quality.low"),
                StreamQualitySelection::Mp4Low,
            ),
        ],
        _ => vec![(
            i18n::tr(language, "stream_audio.quality.original"),
            StreamQualitySelection::Original,
        )],
    }
}

fn current_stream_format(state: &StreamDialogState) -> StreamOutputFormat {
    let format_idx =
        crate::send_message_w_safe(state.format_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    StreamOutputFormat::combo_items(state.language)
        .get(format_idx.max(0) as usize)
        .map(|(_, f)| *f)
        .unwrap_or(StreamOutputFormat::Auto)
}

fn refill_stream_quality_combo(state: &StreamDialogState, keep_selection: bool) {
    let prev_selection = if keep_selection {
        crate::send_message_w_safe(state.quality_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0
    } else {
        -1
    };
    crate::send_message_w_safe(state.quality_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let format = current_stream_format(state);
    let items = stream_quality_items(state.language, format);
    for (label, _) in &items {
        let wide = to_wide(label);
        crate::send_message_w_safe(
            state.quality_combo,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(wide.as_ptr() as isize),
        );
    }
    let selected_idx = if prev_selection >= 0 && (prev_selection as usize) < items.len() {
        prev_selection
    } else {
        0
    };
    crate::send_message_w_safe(
        state.quality_combo,
        CB_SETCURSEL,
        WPARAM(selected_idx as usize),
        LPARAM(0),
    );
}

fn with_stream_dialog_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut StreamDialogState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut StreamDialogState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

fn with_stream_track_dialog_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut StreamTrackDialogState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut StreamTrackDialogState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

unsafe extern "system" fn stream_dialog_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "stream_dialog_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || stream_dialog_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn stream_dialog_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create_struct = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*create_struct).lpCreateParams as *mut StreamDialogInit;
                if init_ptr.is_null() {
                    return LRESULT(0);
                }
                let init = Box::from_raw(init_ptr);
                let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

                let url_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&i18n::tr(init.language, "stream_audio.url_label")).as_ptr()),
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
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    110,
                    16,
                    330,
                    22,
                    hwnd,
                    HMENU(STREAM_ID_URL as isize),
                    HINSTANCE(0),
                    None,
                );
                let favorites_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(
                        to_wide(&tr_or(
                            init.language,
                            "stream_audio.favorites_label",
                            "Favorites:",
                        ))
                        .as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    50,
                    90,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let favorites_combo = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    110,
                    48,
                    330,
                    180,
                    hwnd,
                    HMENU(STREAM_ID_FAVORITES as isize),
                    HINSTANCE(0),
                    None,
                );
                let direct_play_check = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "stream_audio.direct_play")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                    110,
                    82,
                    330,
                    22,
                    hwnd,
                    HMENU(STREAM_ID_DIRECT_PLAY as isize),
                    HINSTANCE(0),
                    None,
                );
                let format_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&i18n::tr(init.language, "stream_audio.format_label")).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    116,
                    90,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let format_combo = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    110,
                    114,
                    210,
                    180,
                    hwnd,
                    HMENU(STREAM_ID_FORMAT as isize),
                    HINSTANCE(0),
                    None,
                );
                let quality_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(
                        to_wide(&i18n::tr(init.language, "stream_audio.quality_label")).as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    148,
                    90,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let quality_combo = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    110,
                    146,
                    210,
                    180,
                    hwnd,
                    HMENU(STREAM_ID_QUALITY as isize),
                    HINSTANCE(0),
                    None,
                );
                let ok_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "youtube.ok")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    350,
                    178,
                    90,
                    28,
                    hwnd,
                    HMENU(STREAM_ID_OK as isize),
                    HINSTANCE(0),
                    None,
                );
                let cancel_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "youtube.cancel")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    350,
                    210,
                    90,
                    28,
                    hwnd,
                    HMENU(STREAM_ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                );

                for control in [
                    url_label,
                    url_edit,
                    favorites_label,
                    favorites_combo,
                    format_label,
                    format_combo,
                    quality_label,
                    quality_combo,
                    direct_play_check,
                    ok_button,
                    cancel_button,
                ] {
                    if control.0 != 0 && hfont.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                }

                let items = StreamOutputFormat::combo_items(init.language);
                for (label, _) in &items {
                    let wide = to_wide(label);
                    SendMessageW(
                        format_combo,
                        CB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(wide.as_ptr() as isize),
                    );
                }
                let default_idx = items
                    .iter()
                    .position(|(_, format)| *format == init.default_format)
                    .unwrap_or(0);
                SendMessageW(format_combo, CB_SETCURSEL, WPARAM(default_idx), LPARAM(0));
                let favorites = load_stream_favorites(init.parent);

                let state = Box::new(StreamDialogState {
                    parent: init.parent,
                    language: init.language,
                    url_edit,
                    favorites_combo,
                    favorites,
                    format_combo,
                    quality_combo,
                    direct_play_check,
                    ok_button,
                    result: init.result.clone(),
                });
                refill_stream_quality_combo(&state, false);
                refill_stream_favorites_combo(&state, Some(0));
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                SetFocus(url_edit);
                LRESULT(0)
            }
            WM_CONTEXTMENU => {
                let target = HWND(wparam.0 as isize);
                let handled = with_stream_dialog_state(hwnd, |state| {
                    if target.0 != 0 && target != state.favorites_combo && target != hwnd {
                        return false;
                    }
                    let selected =
                        SendMessageW(state.favorites_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                    if selected < 0 || selected as usize >= state.favorites.len() {
                        return false;
                    }
                    let menu = match CreatePopupMenu() {
                        Ok(menu) => menu,
                        Err(err) => {
                            crate::log_debug(&format!(
                                "Failed to create stream favorites context menu: {}",
                                err
                            ));
                            return false;
                        }
                    };
                    let label =
                        to_wide(&i18n::tr(state.language, "voice_panel.remove_favorite"));
                    if let Err(err) = AppendMenuW(menu, MF_STRING, 1, PCWSTR(label.as_ptr())) {
                        crate::log_debug(&format!(
                            "Failed to append stream favorites context menu item: {}",
                            err
                        ));
                        crate::log_if_err!(DestroyMenu(menu));
                        return false;
                    }
                    let point = if lparam.0 == -1 {
                        let mut pt = POINT::default();
                        if let Err(err) = GetCursorPos(&mut pt) {
                            crate::log_debug(&format!(
                                "Failed to query cursor position for stream favorites context menu: {}",
                                err
                            ));
                            crate::log_if_err!(DestroyMenu(menu));
                            return false;
                        }
                        pt
                    } else {
                        POINT {
                            x: (lparam.0 as u32 & 0xFFFF) as i16 as i32,
                            y: ((lparam.0 as u32 >> 16) & 0xFFFF) as i16 as i32,
                        }
                    };
                    let command = TrackPopupMenu(
                        menu,
                        TPM_RETURNCMD | TPM_NONOTIFY,
                        point.x,
                        point.y,
                        0,
                        hwnd,
                        None,
                    );
                    crate::log_if_err!(DestroyMenu(menu));
                    if command.0 != 1 {
                        return true;
                    }
                    let removed_url = state.favorites[selected as usize].url.clone();
                    state.favorites.remove(selected as usize);
                    remove_stream_favorite(state.parent, &removed_url);
                    let next_selection = if state.favorites.is_empty() {
                        None
                    } else if selected as usize >= state.favorites.len() {
                        Some(state.favorites.len() - 1)
                    } else {
                        Some(selected as usize)
                    };
                    refill_stream_favorites_combo(state, next_selection);
                    true
                })
                .unwrap_or(false);
                if handled {
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_COMMAND => {
                let cmd_id = wparam.0 & 0xffff;
                let notify_code = (wparam.0 >> 16) & 0xffff;
                if cmd_id == STREAM_ID_FORMAT && notify_code == CBN_SELCHANGE as usize {
                    if with_stream_dialog_state(hwnd, |state| {
                        refill_stream_quality_combo(state, false);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to refresh stream quality combo");
                    }
                    return LRESULT(0);
                }
                if cmd_id == STREAM_ID_OK || cmd_id == 1 {
                    if with_stream_dialog_state(hwnd, |state| {
                        let msg = i18n::tr(state.language, "stream_audio.progress_downloading");
                        if !screen_reader_speak(&msg) {
                            crate::log_debug("Screen reader speak failed");
                        }
                        let mut url = read_edit_text(state.url_edit);
                        if url.trim().is_empty() {
                            let favorite_idx = SendMessageW(
                                state.favorites_combo,
                                CB_GETCURSEL,
                                WPARAM(0),
                                LPARAM(0),
                            )
                            .0;
                            if favorite_idx >= 0
                                && let Some(favorite) = state.favorites.get(favorite_idx as usize)
                            {
                                url = favorite.url.clone();
                            }
                        }
                        let format_idx =
                            SendMessageW(state.format_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        let format = StreamOutputFormat::combo_items(state.language)
                            .get(format_idx.max(0) as usize)
                            .map(|(_, f)| *f)
                            .unwrap_or(StreamOutputFormat::Auto);
                        let quality_idx =
                            SendMessageW(state.quality_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        let quality = stream_quality_items(state.language, format)
                            .get(quality_idx.max(0) as usize)
                            .map(|(_, q)| *q)
                            .unwrap_or(StreamQualitySelection::Original);
                        let direct_play = SendMessageW(
                            state.direct_play_check,
                            BM_GETCHECK,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0 == BST_CHECKED.0 as isize;
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) =
                            Some(StreamDialogResult {
                                url,
                                format,
                                quality,
                                direct_play,
                                reopen_collection_page: None,
                            });
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access stream dialog state");
                    }
                    crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
                    return LRESULT(0);
                }
                if cmd_id == STREAM_ID_CANCEL || cmd_id == 2 {
                    if with_stream_dialog_state(hwnd, |state| {
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) = None;
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access stream dialog state");
                    }
                    crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    crate::log_if_err!(PostMessageW(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(STREAM_ID_CANCEL),
                        LPARAM(0)
                    ));
                    return LRESULT(0);
                }
                if wparam.0 as u32 == VK_RETURN.0 as u32 {
                    let ok =
                        with_stream_dialog_state(hwnd, |state| state.ok_button).unwrap_or(HWND(0));
                    let favorites_combo =
                        with_stream_dialog_state(hwnd, |state| state.favorites_combo)
                            .unwrap_or(HWND(0));
                    if GetFocus() == ok || GetFocus() == favorites_combo {
                        crate::log_if_err!(PostMessageW(
                            hwnd,
                            WM_COMMAND,
                            WPARAM(STREAM_ID_OK),
                            LPARAM(0)
                        ));
                        return LRESULT(0);
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => {
                if with_stream_dialog_state(hwnd, |state| {
                    EnableWindow(state.parent, true);
                    let accepted = state
                        .result
                        .lock()
                        .unwrap_or_else(|e| e.into_inner())
                        .is_some();
                    if !accepted {
                        SetForegroundWindow(state.parent);
                    }
                })
                .is_none()
                {
                    crate::log_debug("Failed to access stream dialog state");
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut StreamDialogState;
                if !ptr.is_null() {
                    let _unused = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn show_stream_dialog(
    parent: HWND,
    language: Language,
    default_format: StreamOutputFormat,
) -> Option<StreamDialogResult> {
    if let Some(context) = take_pending_stream_reopen_context() {
        let input = context.input.unwrap_or_default();
        crate::log_debug(&format!(
            "stream transition [reopen_stream_selection]: input={} page={:?}",
            input, context.collection_page
        ));
        return Some(StreamDialogResult {
            url: input,
            format: default_format,
            quality: StreamQualitySelection::Original,
            direct_play: false,
            reopen_collection_page: context.collection_page,
        });
    }
    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
    let class_name = to_wide(STREAM_DIALOG_CLASS_NAME);
    let wc = windows::Win32::UI::WindowsAndMessaging::WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(stream_dialog_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe {
        RegisterClassW(&wc);
    }

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(StreamDialogInit {
        parent,
        language,
        default_format,
        result: Arc::clone(&result),
    });
    let title = to_wide(&i18n::tr(language, "stream_audio.prompt_title"));
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            470,
            300,
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
    }
    pin_stream_modal_window(hwnd);

    let mut msg = MSG::default();
    loop {
        if !crate::is_window_handle_valid(hwnd) {
            break;
        }
        let res = crate::get_message_w_safe(&mut msg, HWND(0), 0, 0);
        if res.0 == 0 || res.0 == -1 {
            break;
        }
        unsafe {
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    crate::watchdog::exit_modal_dialog();
    let result_value = result.lock().unwrap_or_else(|e| e.into_inner()).clone();
    unsafe {
        if result_value.is_none() {
            EnableWindow(parent, true);
            SetForegroundWindow(parent);
        } else {
            crate::log_debug(
                "stream track selection accepted: keeping parent disabled to avoid focus bounce",
            );
        }
    }
    result_value
}

fn set_pending_stream_reopen_context(context: Option<crate::YouTubeReturnContext>) {
    let mut pending = PENDING_STREAM_REOPEN_CONTEXT
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    *pending = context;
}

fn take_pending_stream_reopen_context() -> Option<crate::YouTubeReturnContext> {
    let mut pending = PENDING_STREAM_REOPEN_CONTEXT
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    pending.take()
}

fn update_youtube_return_context_from_selection(
    parent: HWND,
    original_input: &str,
    collection_url: Option<&str>,
    collection_page: Option<usize>,
) {
    let return_input = collection_url
        .map(ToOwned::to_owned)
        .unwrap_or_else(|| original_input.to_string());
    crate::log_debug(&format!(
        "stream transition [update_return_context_from_selection]: input={} collection_page={:?}",
        return_input, collection_page
    ));
    crate::set_active_youtube_return_context(parent, Some(return_input), collection_page);
}

pub(crate) fn reopen_stream_selection(parent: HWND, context: crate::YouTubeReturnContext) {
    set_pending_stream_reopen_context(Some(context));
    play_streaming_audio_from_url(parent);
}

fn open_progress_dialog(
    parent: HWND,
    language: Language,
    title_key: &str,
    status_key: &str,
    show_cancel: bool,
) -> HWND {
    let labels = crate::app_windows::podcast_save_window::SaveDialogLabels {
        title: i18n::tr(language, title_key),
        in_progress: i18n::tr(language, status_key),
        cancel: i18n::tr(language, "podcast.save.cancel"),
        cancel_confirm: i18n::tr(language, "podcast.cancel_confirm"),
    };
    let dialog = crate::app_windows::podcast_save_window::open_with_labels(
        parent,
        language,
        labels,
        show_cancel,
    );
    crate::log_debug(&format!(
        "stream progress dialog opened: parent={:?} dialog={:?} title_key={} status_key={} show_cancel={}",
        parent, dialog, title_key, status_key, show_cancel
    ));
    if dialog.0 != 0 {
        pin_stream_modal_window(dialog);
        crate::app_windows::podcast_save_window::focus_cancel_button(dialog);
        keep_stream_progress_focus(dialog);
        std::thread::sleep(std::time::Duration::from_millis(15));
        pin_stream_modal_window(dialog);
        keep_stream_progress_focus(dialog);
        log_stream_focus_snapshot("open_progress_dialog.after_focus", dialog);
    }
    dialog
}

fn restore_stream_parent_focus(hwnd: HWND) {
    if hwnd.0 == 0 {
        return;
    }
    crate::set_foreground_window_safe(hwnd);
}

fn log_stream_transition(tag: &str, parent: HWND) {
    crate::log_debug(&format!(
        "stream transition [{}]: parent={}",
        tag,
        describe_window_handle(parent)
    ));
    log_stream_focus_snapshot(tag, parent);
}

fn reclaim_stream_modal_parent_foreground(parent: HWND, tag: &str) {
    if parent.0 == 0 {
        return;
    }
    crate::log_debug(&format!(
        "stream parent foreground reclaim [{}]: before={}",
        tag,
        describe_window_handle(crate::get_foreground_window_safe())
    ));
    crate::bring_window_to_foreground(parent);
    log_stream_transition(tag, parent);
}

fn pin_stream_modal_window(hwnd: HWND) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        if let Err(err) = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
        ) {
            crate::log_debug(&format!(
                "Failed to pin stream modal window {:?} as topmost: {}",
                hwnd, err
            ));
        }
    }
    crate::set_foreground_window_safe(hwnd);
}

fn close_progress_dialog(dialog: HWND) {
    if dialog.0 == 0 {
        return;
    }
    crate::send_message_w_safe(
        dialog,
        crate::app_windows::podcast_save_window::WM_PODCAST_SAVE_DONE,
        WPARAM(0),
        LPARAM(0),
    );
}

fn report_progress(dialog: HWND, pct: u32) {
    if dialog.0 == 0 {
        return;
    }
    crate::log_debug(&format!(
        "stream progress update: dialog={:?} pct={}",
        dialog, pct
    ));
    crate::send_message_w_safe(
        dialog,
        crate::app_windows::podcast_save_window::WM_PODCAST_SAVE_PROGRESS,
        WPARAM(pct.min(100) as usize),
        LPARAM(0),
    );
}

fn report_progress_status(dialog: HWND, text: &str) {
    if dialog.0 == 0 {
        return;
    }
    crate::log_debug(&format!(
        "stream progress status: dialog={:?} text={}",
        dialog, text
    ));
    crate::app_windows::podcast_save_window::set_status_text(dialog, text);
}

fn pump_messages_detect_stream_cancel(parent: HWND, dialog: HWND) -> bool {
    let mut cancelled = false;
    let mut msg = MSG::default();
    unsafe {
        while PeekMessageW(&mut msg, HWND(0), 0, 0, PM_REMOVE).as_bool() {
            if msg.hwnd == parent
                && msg.message == crate::app_windows::podcast_save_window::WM_PODCAST_SAVE_CANCEL
            {
                cancelled = true;
                continue;
            }
            if dialog.0 != 0 && IsDialogMessageW(dialog, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    cancelled
}

fn keep_stream_progress_focus(dialog: HWND) {
    unsafe {
        if dialog.0 == 0 || !IsWindow(dialog).as_bool() {
            return;
        }
        let fg = GetForegroundWindow();
        let dialog_is_foreground = fg == dialog || (fg.0 != 0 && IsChild(dialog, fg).as_bool());
        if !dialog_is_foreground {
            crate::log_debug(&format!(
                "stream focus keepalive: restoring foreground dialog={:?} current_fg={}",
                dialog,
                describe_window_handle(fg)
            ));
            SetForegroundWindow(dialog);
        }
        let focus = GetFocus();
        let dialog_has_focus =
            focus == dialog || (focus.0 != 0 && IsChild(dialog, focus).as_bool());
        if !dialog_has_focus {
            crate::log_debug(&format!(
                "stream focus keepalive: restoring keyboard focus dialog={:?} current_focus={}",
                dialog,
                describe_window_handle(focus)
            ));
            crate::app_windows::podcast_save_window::focus_cancel_button(dialog);
            log_stream_focus_snapshot("keep_stream_progress_focus.after_restore", dialog);
        }
    }
}

fn get_window_text_for_log(hwnd: HWND, max_len: usize) -> String {
    if hwnd.0 == 0 {
        return String::new();
    }
    let len = crate::get_window_text_length_w_safe(hwnd);
    if len <= 0 {
        return String::new();
    }
    let mut buf = vec![0u16; len as usize + 1];
    let read = crate::get_window_text_w_safe(hwnd, &mut buf);
    if read <= 0 {
        return String::new();
    }
    let mut text = String::from_utf16_lossy(&buf[..read as usize]);
    if text.chars().count() > max_len {
        text = text.chars().take(max_len).collect();
        text.push_str("...");
    }
    text
}

fn get_window_class_for_log(hwnd: HWND) -> String {
    if hwnd.0 == 0 {
        return String::new();
    }
    let mut class_buf = [0u16; 128];
    let len = crate::get_class_name_w_safe(hwnd, &mut class_buf);
    if len <= 0 {
        return String::new();
    }
    String::from_utf16_lossy(&class_buf[..len as usize])
}

fn describe_window_handle(hwnd: HWND) -> String {
    if hwnd.0 == 0 {
        return "HWND(0)".to_string();
    }
    format!(
        "HWND({}) class='{}' text='{}'",
        hwnd.0,
        get_window_class_for_log(hwnd),
        get_window_text_for_log(hwnd, 120)
    )
}

fn log_stream_focus_snapshot(tag: &str, dialog: HWND) {
    let fg = crate::get_foreground_window_safe();
    let focus = crate::get_focus_safe();
    crate::log_debug(&format!(
        "stream focus snapshot [{}]: dialog={} foreground={} focus={}",
        tag,
        describe_window_handle(dialog),
        describe_window_handle(fg),
        describe_window_handle(focus)
    ));
}

fn parse_ytdlp_progress_pct(line: &str) -> Option<u32> {
    let marker = line.find('%')?;
    let before = &line[..marker];
    let mut start = before.len();
    for (idx, ch) in before.char_indices().rev() {
        if ch.is_ascii_digit() || ch == '.' {
            start = idx;
        } else {
            break;
        }
    }
    if start >= before.len() {
        return None;
    }
    let value = before[start..].trim().parse::<f32>().ok()?;
    Some(value.clamp(0.0, 100.0).round() as u32)
}

fn truncate_debug_text(text: &str, max_chars: usize) -> String {
    let mut out = String::new();
    for ch in text.chars().take(max_chars) {
        out.push(ch);
    }
    if text.chars().count() > max_chars {
        out.push_str("...(truncated)");
    }
    out
}

fn log_stream_cache_snapshot(cache_dir: &Path, prefix: &str, context: &str) {
    let mut items: Vec<String> = Vec::new();
    let Ok(entries) = std::fs::read_dir(cache_dir) else {
        crate::log_debug(&format!(
            "yt-dlp cache snapshot [{}]: read_dir failed for {}",
            context,
            cache_dir.to_string_lossy()
        ));
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if !path.is_file() {
            continue;
        }
        let Some(name) = path.file_name().and_then(|s| s.to_str()) else {
            continue;
        };
        if !name.starts_with(prefix) {
            continue;
        }
        let len = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        items.push(format!("{name} ({len} bytes)"));
    }
    items.sort_unstable();
    if items.is_empty() {
        crate::log_debug(&format!(
            "yt-dlp cache snapshot [{}]: no files for prefix={}",
            context, prefix
        ));
    } else {
        crate::log_debug(&format!(
            "yt-dlp cache snapshot [{}]: {}",
            context,
            items.join(", ")
        ));
    }
}

fn find_latest_downloaded_stream_file(cache_dir: &Path, prefix: &str) -> Option<PathBuf> {
    let mut best: Option<(std::time::SystemTime, PathBuf)> = None;
    const ALLOWED_EXTS: &[&str] = &[
        "mp3", "m4a", "aac", "opus", "ogg", "wav", "flac", "mp4", "webm", "mkv", "ts",
    ];
    let entries = std::fs::read_dir(cache_dir).ok()?;
    for entry in entries.flatten() {
        let path = entry.path();
        if !path.is_file() {
            continue;
        }
        let Some(name) = path.file_name().and_then(|s| s.to_str()) else {
            continue;
        };
        if !name.starts_with(prefix) {
            continue;
        }
        let name_lc = name.to_ascii_lowercase();
        // Exclude temporary/fragment artifacts (e.g. .part, .ytdl, .part-FragNNN).
        if name_lc.ends_with(".part")
            || name_lc.ends_with(".ytdl")
            || name_lc.contains(".part-")
            || name_lc.contains(".ytdl-")
            || name_lc.contains("-frag")
            || name_lc.contains(".frag")
        {
            continue;
        }
        let Some(ext) = path.extension().and_then(|e| e.to_str()) else {
            continue;
        };
        let ext = ext.to_ascii_lowercase();
        if !ALLOWED_EXTS.contains(&ext.as_str())
            || ext == "part"
            || ext == "ytdl"
            || ext == "tmp"
            || ext == "temp"
        {
            continue;
        }
        if std::fs::metadata(&path)
            .map(|m| m.len() == 0)
            .unwrap_or(false)
        {
            continue;
        }
        let modified = std::fs::metadata(&path)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        match &best {
            Some((best_time, _)) if modified <= *best_time => {}
            _ => best = Some((modified, path)),
        }
    }
    best.map(|(_, p)| p)
}

fn plain_label(text: &str) -> String {
    text.replace('&', "")
}

fn format_stream_entry_upload_date(entry: &serde_json::Value) -> Option<String> {
    let raw = entry.get("upload_date")?.as_str()?.trim();
    if raw.len() == 8 && raw.chars().all(|ch| ch.is_ascii_digit()) {
        return Some(format!("{}-{}-{}", &raw[0..4], &raw[4..6], &raw[6..8]));
    }
    if raw.is_empty() {
        None
    } else {
        Some(raw.to_string())
    }
}

fn format_stream_entry_view_count(entry: &serde_json::Value, language: Language) -> Option<String> {
    let count = entry.get("view_count")?.as_u64()?;
    let formatted = format_stream_count(count);
    Some(format!(
        "{} {}",
        i18n::tr(language, "stream_audio.views_label"),
        formatted
    ))
}

fn format_stream_count(count: u64) -> String {
    let digits = count.to_string();
    let mut formatted = String::with_capacity(digits.len() + digits.len() / 3);
    for (idx, ch) in digits.chars().rev().enumerate() {
        if idx != 0 && idx % 3 == 0 {
            formatted.push('.');
        }
        formatted.push(ch);
    }
    formatted.chars().rev().collect()
}

fn format_stream_entry_subscriber_count(
    entry: &serde_json::Value,
    language: Language,
) -> Option<String> {
    let count = entry
        .get("channel_follower_count")
        .and_then(|value| {
            value
                .as_u64()
                .or_else(|| value.as_f64().map(|v| v.round() as u64))
        })
        .or_else(|| {
            entry.get("subscriber_count").and_then(|value| {
                value
                    .as_u64()
                    .or_else(|| value.as_f64().map(|v| v.round() as u64))
            })
        })?;
    let formatted = format_stream_count(count);
    Some(format!(
        "{} {}",
        i18n::tr(language, "stream_audio.subscribers_label"),
        formatted
    ))
}

fn format_stream_entry_channel(entry: &serde_json::Value) -> Option<String> {
    entry
        .get("channel")
        .and_then(|value| value.as_str())
        .or_else(|| entry.get("uploader").and_then(|value| value.as_str()))
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(plain_label)
}

fn format_stream_entry_duration(entry: &serde_json::Value) -> Option<String> {
    if let Some(text) = entry
        .get("duration_string")
        .and_then(|value| value.as_str())
        .map(str::trim)
        .filter(|value| !value.is_empty())
    {
        return Some(text.to_string());
    }
    let seconds = entry.get("duration").and_then(|value| {
        value
            .as_u64()
            .or_else(|| value.as_f64().map(|v| v.round() as u64))
    })?;
    let hours = seconds / 3600;
    let minutes = (seconds % 3600) / 60;
    let secs = seconds % 60;
    if hours > 0 {
        Some(format!("{hours}:{minutes:02}:{secs:02}"))
    } else {
        Some(format!("{minutes}:{secs:02}"))
    }
}
fn format_stream_entry_label(entry: &serde_json::Value, language: Language) -> String {
    let title = entry
        .get("title")
        .and_then(|value| value.as_str())
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .unwrap_or("Untitled");
    let mut parts = vec![plain_label(title)];
    if let Some(duration) = format_stream_entry_duration(entry) {
        parts.push(format!(
            "{} {}",
            i18n::tr(language, "stream_audio.duration_label"),
            duration
        ));
    }
    if let Some(channel) = format_stream_entry_channel(entry) {
        parts.push(format!(
            "{} {}",
            i18n::tr(language, "stream_audio.channel_label"),
            channel
        ));
    }
    if stream_entry_url(entry)
        .as_deref()
        .map(is_youtube_channel_url)
        .unwrap_or(false)
        && let Some(subscriber_count) = format_stream_entry_subscriber_count(entry, language)
    {
        parts.push(subscriber_count);
    }
    if let Some(date) = format_stream_entry_upload_date(entry) {
        parts.push(date);
    }
    if let Some(view_count) = format_stream_entry_view_count(entry, language) {
        parts.push(view_count);
    }
    parts.join(" - ")
}
fn probe_stream_media_title(ytdlp_path: &Path, url: &str) -> Option<String> {
    let output = ytdlp_command(ytdlp_path)
        .arg("--no-playlist")
        .arg("--no-warnings")
        .arg("--skip-download")
        .arg("--print")
        .arg("title")
        .arg("--")
        .arg(url)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }
    let raw = match String::from_utf8(output.stdout.clone()) {
        Ok(text) => text,
        Err(_) => {
            let (decoded, _, _) = WINDOWS_1252.decode(&output.stdout);
            decoded.into_owned()
        }
    };
    let first_non_empty = raw.lines().map(str::trim).find(|line| !line.is_empty())?;
    let repaired = repair_stream_title_mojibake(first_non_empty);
    let sanitized = crate::sanitize_filename(&repaired);
    if sanitized.is_empty() {
        None
    } else {
        Some(sanitized)
    }
}

fn repair_stream_title_mojibake(text: &str) -> String {
    let repaired = text
        .replace("â€™", "’")
        .replace("â€˜", "‘")
        .replace("â€œ", "“")
        .replace("â€", "”")
        .replace("â€“", "–")
        .replace("â€”", "—")
        .replace("â€¦", "…")
        .replace("Â ", " ")
        .replace("Ã ", "à")
        .replace("Ã¨", "è")
        .replace("Ã©", "é")
        .replace("Ã¬", "ì")
        .replace("Ã²", "ò")
        .replace("Ã¹", "ù")
        .replace("Ã€", "À")
        .replace("Ãˆ", "È")
        .replace("Ã‰", "É")
        .replace("ÃŒ", "Ì")
        .replace("Ã’", "Ò")
        .replace("Ã™", "Ù");
    repaired.replace('\u{FFFD}', "")
}

fn unique_stream_named_path(dir: &Path, title: &str, ext: &str) -> PathBuf {
    let first = dir.join(format!("{title}.{ext}"));
    if !first.exists() {
        return first;
    }
    for n in 2..=999 {
        let candidate = dir.join(format!("{title} ({n}).{ext}"));
        if !candidate.exists() {
            return candidate;
        }
    }
    let stamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0);
    dir.join(format!("{title}_{stamp}.{ext}"))
}

fn probe_stream_audio_tracks(
    ytdlp_path: &Path,
    url: &str,
) -> Result<Vec<StreamAudioTrack>, String> {
    let output = ytdlp_command(ytdlp_path)
        .arg("--dump-single-json")
        .arg("--no-playlist")
        .arg("--no-warnings")
        .arg("--force-ipv4")
        .arg("--")
        .arg(url)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|e| e.to_string())?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if stderr.is_empty() {
            "yt-dlp probe failed".to_string()
        } else {
            stderr
        });
    }
    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).map_err(|e| e.to_string())?;
    let mut seen = HashSet::new();
    let mut tracks = Vec::new();
    if let Some(formats) = json.get("formats").and_then(|v| v.as_array()) {
        for fmt in formats {
            let format_id = fmt
                .get("format_id")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .trim();
            if format_id.is_empty() || !seen.insert(format_id.to_string()) {
                continue;
            }
            let acodec = fmt
                .get("acodec")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .trim();
            if acodec.is_empty() || acodec.eq_ignore_ascii_case("none") {
                continue;
            }
            let vcodec = fmt
                .get("vcodec")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .trim();
            if !vcodec.is_empty() && !vcodec.eq_ignore_ascii_case("none") {
                continue;
            }
            let lang = fmt
                .get("language")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .trim();
            let note = fmt
                .get("format_note")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .trim();
            let abr = fmt.get("abr").and_then(|v| v.as_f64()).unwrap_or(0.0);
            let mut label = String::new();
            if !lang.is_empty() {
                label.push_str(lang);
            } else if !note.is_empty() {
                label.push_str(note);
            } else {
                label.push_str("audio");
            }
            if abr > 0.0 {
                label.push_str(&format!(" {}k", abr.round() as i64));
            }
            label.push_str(&format!(" ({format_id})"));
            tracks.push(StreamAudioTrack {
                format_id: format_id.to_string(),
                label,
            });
        }
    }
    Ok(tracks)
}

fn probe_stream_audio_tracks_responsive(
    parent: HWND,
    language: Language,
    ytdlp_path: &Path,
    url: &str,
) -> Result<Vec<StreamAudioTrack>, String> {
    let progress = open_progress_dialog(
        parent,
        language,
        "stream_audio.progress_title",
        "podcasts.loading",
        false,
    );
    let ytdlp = ytdlp_path.to_path_buf();
    let url = url.to_string();
    let worker = std::thread::spawn(move || probe_stream_audio_tracks(&ytdlp, &url));
    let mut last_focus_keepalive = std::time::Instant::now();
    while !worker.is_finished() {
        ignore_bool(pump_messages_detect_stream_cancel(parent, progress));
        if last_focus_keepalive.elapsed() > std::time::Duration::from_millis(300) {
            keep_stream_progress_focus(progress);
            last_focus_keepalive = std::time::Instant::now();
        }
        std::thread::sleep(std::time::Duration::from_millis(30));
    }
    let result = match worker.join() {
        Ok(result) => result,
        Err(_) => Err("yt-dlp track probe worker failed".to_string()),
    };
    crate::log_debug(&format!(
        "stream transition [track_probe.completed]: ok={} track_count={}",
        result.is_ok(),
        result.as_ref().map(|tracks| tracks.len()).unwrap_or(0)
    ));
    if result.is_ok() {
        crate::app_windows::podcast_save_window::suppress_parent_restore_on_close(progress);
    }
    close_progress_dialog(progress);
    result
}

unsafe extern "system" fn stream_track_dialog_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "stream_track_dialog_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || stream_track_dialog_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn stream_track_dialog_wndproc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create_struct = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*create_struct).lpCreateParams as *mut StreamTrackDialogInit;
                if init_ptr.is_null() {
                    return LRESULT(0);
                }
                let init = Box::from_raw(init_ptr);
                let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

                let label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(
                        to_wide(&plain_label(&i18n::tr(
                            init.language,
                            "playback.audio_track",
                        )))
                        .as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    18,
                    120,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let combo = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    140,
                    16,
                    300,
                    220,
                    hwnd,
                    HMENU(STREAM_TRACK_ID_COMBO as isize),
                    HINSTANCE(0),
                    None,
                );
                let ok_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "youtube.ok")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    260,
                    56,
                    88,
                    28,
                    hwnd,
                    HMENU(STREAM_TRACK_ID_OK as isize),
                    HINSTANCE(0),
                    None,
                );
                let cancel_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(init.language, "youtube.cancel")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    352,
                    56,
                    88,
                    28,
                    hwnd,
                    HMENU(STREAM_TRACK_ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                );

                for control in [label, combo, ok_button, cancel_button] {
                    if control.0 != 0 && hfont.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                }

                let auto_label = i18n::tr(init.language, "stream_audio.track.auto");
                let auto_w = to_wide(&auto_label);
                SendMessageW(
                    combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(auto_w.as_ptr() as isize),
                );
                for track in &init.tracks {
                    let w = to_wide(&track.label);
                    SendMessageW(combo, CB_ADDSTRING, WPARAM(0), LPARAM(w.as_ptr() as isize));
                }
                SendMessageW(combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));

                let state = Box::new(StreamTrackDialogState {
                    parent: init.parent,
                    language: init.language,
                    combo,
                    ok_button,
                    tracks: init.tracks,
                    result: init.result,
                });
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                SetFocus(combo);
                LRESULT(0)
            }
            WM_COMMAND => {
                let cmd_id = wparam.0 & 0xffff;
                if cmd_id == STREAM_TRACK_ID_OK || cmd_id == 1 {
                    if with_stream_track_dialog_state(hwnd, |state| {
                        let msg = i18n::tr(state.language, "podcasts.loading");
                        if !screen_reader_speak(&msg) {
                            crate::log_debug("Screen reader speak failed");
                        }
                        let idx = SendMessageW(state.combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        let selected: Option<Option<String>> = if idx <= 0 {
                            Some(None)
                        } else {
                            state
                                .tracks
                                .get((idx - 1) as usize)
                                .map(|t| Some(t.format_id.clone()))
                        };
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) = selected;
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access stream track dialog state");
                    }
                    crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
                    return LRESULT(0);
                }
                if cmd_id == STREAM_TRACK_ID_CANCEL || cmd_id == 2 {
                    if with_stream_track_dialog_state(hwnd, |state| {
                        *state.result.lock().unwrap_or_else(|e| e.into_inner()) = None;
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access stream track dialog state");
                    }
                    crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    crate::log_if_err!(PostMessageW(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(STREAM_TRACK_ID_CANCEL),
                        LPARAM(0)
                    ));
                    return LRESULT(0);
                }
                if wparam.0 as u32 == VK_RETURN.0 as u32 {
                    let ok = with_stream_track_dialog_state(hwnd, |state| state.ok_button)
                        .unwrap_or(HWND(0));
                    if GetFocus() == ok {
                        crate::log_if_err!(PostMessageW(
                            hwnd,
                            WM_COMMAND,
                            WPARAM(STREAM_TRACK_ID_OK),
                            LPARAM(0)
                        ));
                        return LRESULT(0);
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => {
                if with_stream_track_dialog_state(hwnd, |state| {
                    EnableWindow(state.parent, true);
                    let accepted = state
                        .result
                        .lock()
                        .unwrap_or_else(|e| e.into_inner())
                        .is_some();
                    if !accepted {
                        SetForegroundWindow(state.parent);
                    }
                })
                .is_none()
                {
                    crate::log_debug("Failed to access stream track dialog state");
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut StreamTrackDialogState;
                if !ptr.is_null() {
                    let _unused = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn choose_stream_audio_track(
    parent: HWND,
    language: Language,
    tracks: Vec<StreamAudioTrack>,
) -> Option<Option<String>> {
    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
    let class_name = to_wide(STREAM_TRACK_DIALOG_CLASS_NAME);
    let wc = windows::Win32::UI::WindowsAndMessaging::WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(stream_track_dialog_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe {
        RegisterClassW(&wc);
    }
    let result = Arc::new(Mutex::new(None));
    let init = Box::new(StreamTrackDialogInit {
        parent,
        language,
        tracks,
        result: Arc::clone(&result),
    });
    let title = to_wide(&plain_label(&i18n::tr(language, "playback.audio_track")));
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            470,
            140,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };
    if hwnd.0 == 0 {
        return Some(None);
    }
    unsafe {
        EnableWindow(parent, false);
    }
    pin_stream_modal_window(hwnd);
    let mut msg = MSG::default();
    loop {
        if !crate::is_window_handle_valid(hwnd) {
            break;
        }
        let res = crate::get_message_w_safe(&mut msg, HWND(0), 0, 0);
        if res.0 == 0 || res.0 == -1 {
            break;
        }
        unsafe {
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    let result_value = result.lock().unwrap_or_else(|e| e.into_inner()).clone();
    unsafe {
        EnableWindow(parent, true);
        if result_value.is_none() {
            SetForegroundWindow(parent);
        }
    }
    result_value
}

pub fn play_streaming_audio_from_url(parent: HWND) {
    log_stream_transition("play_streaming_audio.start", parent);
    let (language, default_format) = {
        with_state(parent, |state| {
            (
                state.settings.language,
                StreamOutputFormat::from_settings_value(
                    &state.settings.stream_audio_default_format,
                ),
            )
        })
    }
    .unwrap_or((Language::default(), StreamOutputFormat::Auto));
    let Some(dialog_data) = show_stream_dialog(parent, language, default_format) else {
        post_focus_editor(parent);
        return;
    };
    let saved = {
        with_state(parent, |state| {
            let new_value = dialog_data.format.settings_value().to_string();
            if state.settings.stream_audio_default_format != new_value {
                state.settings.stream_audio_default_format = new_value;
                save_settings(state.settings.clone());
            }
        })
    };
    if saved.is_none() {
        crate::log_debug("Failed to persist stream output format preference");
    }
    let input = dialog_data.url.trim().to_string();
    if input.is_empty() {
        show_error(
            parent,
            language,
            &i18n::tr(language, "stream_audio.invalid_url"),
        );
        return;
    }

    let labels_data = labels(language);
    let ytdlp_debug = ytdlp_debug_enabled();
    if ytdlp_debug {
        crate::log_debug("yt-dlp debug mode enabled for streaming (SONARPAD_YTDLP_DEBUG=1)");
    }

    let needs_ytdlp_selection =
        !looks_like_valid_stream_url(&input) || is_youtube_collection_url(&input);
    if needs_ytdlp_selection {
        crate::set_active_youtube_return_context(parent, Some(input.clone()), None);
    } else {
        crate::set_active_youtube_return_context(parent, None, None);
    }
    let mut url = input.clone();
    let mut collection_url: Option<String> = None;
    let mut collection_page: Option<usize> = None;
    let mut ytdlp_path = None;
    if needs_ytdlp_selection {
        let bootstrap_progress = open_progress_dialog(
            parent,
            language,
            "stream_audio.progress_title",
            "podcasts.loading",
            false,
        );
        let path = match ensure_ytdlp_available(
            parent,
            language,
            &labels_data,
            Some(bootstrap_progress),
        ) {
            Ok(Some(path)) => path,
            Ok(None) => {
                post_focus_editor(parent);
                return;
            }
            Err(err) => {
                let message =
                    i18n::tr_f(language, "stream_audio.download_failed", &[("err", &err)]);
                show_error(parent, language, &message);
                return;
            }
        };
        let resolved = match resolve_stream_input_url(
            parent,
            language,
            &path,
            &input,
            dialog_data.reopen_collection_page,
            Some(bootstrap_progress),
        ) {
            Ok(Some(selection)) => selection,
            Ok(None) => {
                post_focus_editor(parent);
                return;
            }
            Err(err) => {
                let message =
                    i18n::tr_f(language, "stream_audio.download_failed", &[("err", &err)]);
                show_error(parent, language, &message);
                return;
            }
        };
        url = resolved.url;
        collection_url = resolved.collection_url;
        collection_page = resolved.collection_page;
        crate::log_debug(&format!(
            "stream transition [resolved_selection]: url={} collection_url={} collection_page={:?}",
            url,
            collection_url.as_deref().unwrap_or(""),
            collection_page
        ));
        update_youtube_return_context_from_selection(
            parent,
            &input,
            collection_url.as_deref(),
            collection_page,
        );
        ytdlp_path = Some(path);
    }
    if collection_page.is_none()
        && let Some(initial_page) = dialog_data.reopen_collection_page
        && collection_url
            .as_deref()
            .map(is_youtube_collection_url)
            .unwrap_or(false)
    {
        collection_page = Some(initial_page);
        update_youtube_return_context_from_selection(
            parent,
            &input,
            collection_url.as_deref(),
            collection_page,
        );
    }

    if dialog_data.direct_play {
        let stream_path = PathBuf::from(&url);
        crate::queue_audio_files_and_play(parent, vec![stream_path.clone()]);
        crate::editor_manager::mark_current_document_from_rss(parent, true);
        let episode_title = stream_path
            .file_name()
            .and_then(|s| s.to_str())
            .filter(|s| !s.trim().is_empty())
            .map(|s| s.to_string())
            .or_else(|| Some(url.clone()));
        crate::set_active_podcast_episode_info(parent, Some(url), episode_title, Some(stream_path));
        let return_input = collection_url.clone().unwrap_or_else(|| input.clone());
        crate::set_active_youtube_return_context(parent, Some(return_input), collection_page);
        crate::menu::update_playback_menu(parent, true);
        return;
    }

    let ytdlp_path = match ytdlp_path {
        Some(path) => path,
        None => match ensure_ytdlp_available(parent, language, &labels_data, None) {
            Ok(Some(path)) => path,
            Ok(None) => {
                post_focus_editor(parent);
                return;
            }
            Err(err) => {
                let message =
                    i18n::tr_f(language, "stream_audio.download_failed", &[("err", &err)]);
                show_error(parent, language, &message);
                return;
            }
        },
    };
    'select_video: loop {
        let selected_audio_format =
            match probe_stream_audio_tracks_responsive(parent, language, &ytdlp_path, &url) {
                Ok(tracks) if tracks.len() > 1 => {
                    match choose_stream_audio_track(parent, language, tracks) {
                        Some(chosen) => chosen,
                        None => {
                            post_focus_editor(parent);
                            return;
                        }
                    }
                }
                Ok(_) => None,
                Err(err) => {
                    crate::log_debug(&format!("Stream audio track probe failed: {}", err));
                    if is_members_only_stream_error(&err)
                        && let Some(source_collection_url) = collection_url.as_deref()
                    {
                        show_error(parent, language, &members_only_stream_message(language));
                        let next_selection = match choose_youtube_collection_entry(
                            parent,
                            language,
                            &ytdlp_path,
                            source_collection_url,
                            collection_page,
                            None,
                        ) {
                            Ok(Some(selection)) => selection,
                            Ok(None) => {
                                post_focus_editor(parent);
                                return;
                            }
                            Err(next_err) => {
                                let message = i18n::tr_f(
                                    language,
                                    "stream_audio.download_failed",
                                    &[("err", &next_err)],
                                );
                                show_error(parent, language, &message);
                                return;
                            }
                        };
                        url = next_selection.url;
                        collection_url = next_selection.collection_url;
                        collection_page = next_selection.collection_page;
                        update_youtube_return_context_from_selection(
                            parent,
                            &input,
                            collection_url.as_deref(),
                            collection_page,
                        );
                        continue 'select_video;
                    }
                    None
                }
            };

        reclaim_stream_modal_parent_foreground(
            parent,
            "play_streaming_audio.before_final_progress",
        );
        let cache_dir = settings_dir().join("podcast cache");
        if let Err(err) = std::fs::create_dir_all(&cache_dir) {
            let message = i18n::tr_f(
                language,
                "stream_audio.download_failed",
                &[("err", &err.to_string())],
            );
            show_error(parent, language, &message);
            return;
        }
        let stamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_millis())
            .unwrap_or(0);
        let prefix = format!("stream_{}_{}", std::process::id(), stamp);
        let output_template = cache_dir.join(format!("{prefix}.%(ext)s"));
        let progress = open_progress_dialog(
            parent,
            language,
            "stream_audio.progress_title",
            "stream_audio.progress_downloading",
            true,
        );
        keep_stream_progress_focus(progress);
        std::thread::sleep(std::time::Duration::from_millis(15));
        keep_stream_progress_focus(progress);
        log_stream_focus_snapshot("play_streaming_audio.progress_opened", progress);
        let stream_title = probe_stream_media_title(&ytdlp_path, &url);

        let mut cmd = ytdlp_command(&ytdlp_path);
        cmd.arg("--no-playlist")
            .arg("--socket-timeout")
            .arg(YTDLP_SOCKET_TIMEOUT_SECS)
            .arg("--no-warnings")
            .arg("--newline")
            .arg("--verbose");

        if let Some(format_id) = selected_audio_format.as_ref() {
            match dialog_data.format {
                StreamOutputFormat::Mp4 => {
                    cmd.arg("-f").arg(format!(
                        "bestvideo[ext=mp4]+{id}/bestvideo+{id}/best[ext=mp4]/best",
                        id = format_id
                    ));
                }
                StreamOutputFormat::Auto => {
                    cmd.arg("-f").arg(format!(
                        "bestvideo[ext=webm]+{id}/bestvideo+{id}/best",
                        id = format_id
                    ));
                }
                _ => {
                    cmd.arg("-f").arg(format_id);
                }
            }
        } else {
            match dialog_data.format {
                StreamOutputFormat::Mp4 => {
                    let mp4_profile = match dialog_data.quality {
                        StreamQualitySelection::Mp4High => {
                            "bestvideo[ext=mp4]+bestaudio/best[ext=mp4]/best"
                        }
                        StreamQualitySelection::Mp4Medium => {
                            "best[ext=mp4][height<=720]/best[height<=720]/best"
                        }
                        StreamQualitySelection::Mp4Low => {
                            "best[ext=mp4][height<=480]/best[height<=480]/worst"
                        }
                        _ => "best[ext=mp4]/best",
                    };
                    cmd.arg("-f").arg(mp4_profile);
                }
                StreamOutputFormat::Auto => {
                    cmd.arg("-f")
                        .arg("bestvideo[ext=webm]+bestaudio[ext=webm]/bestvideo+bestaudio/best");
                }
                _ => {
                    cmd.arg("-f").arg("bestaudio/best");
                }
            }
        }
        cmd.arg("-o")
            .arg(output_template.to_string_lossy().to_string())
            .arg("--")
            .arg(&url)
            .stdout(Stdio::null())
            .stderr(Stdio::piped());
        if ytdlp_debug {
            let format_for_log = selected_audio_format.as_deref().map_or_else(
                || match dialog_data.format {
                    StreamOutputFormat::Auto => i18n::tr(language, "stream_audio.format.auto"),
                    StreamOutputFormat::Mp3 => "mp3".to_string(),
                    StreamOutputFormat::M4a => "m4a".to_string(),
                    StreamOutputFormat::Opus => "opus".to_string(),
                    StreamOutputFormat::Ogg => "ogg".to_string(),
                    StreamOutputFormat::Wav => "wav".to_string(),
                    StreamOutputFormat::Flac => "flac".to_string(),
                    StreamOutputFormat::Mp4 => "mp4".to_string(),
                },
                |f| f.to_string(),
            );
            crate::log_debug(&format!(
                "yt-dlp stream start: url={} output_template={} format={}",
                url,
                output_template.to_string_lossy(),
                format_for_log
            ));
        }

        let mut child = match cmd.spawn() {
            Ok(c) => c,
            Err(err) => {
                close_progress_dialog(progress);
                let message = i18n::tr_f(
                    language,
                    "stream_audio.download_failed",
                    &[("err", &err.to_string())],
                );
                show_error(parent, language, &message);
                return;
            }
        };
        keep_stream_progress_focus(progress);

        let stderr_pipe = match child.stderr.take() {
            Some(stderr) => stderr,
            None => {
                close_progress_dialog(progress);
                show_error(
                    parent,
                    language,
                    &i18n::tr_f(
                        language,
                        "stream_audio.download_failed",
                        &[("err", "yt-dlp stderr unavailable")],
                    ),
                );
                return;
            }
        };

        let progress_shared = Arc::new(AtomicU32::new(0));
        let activity_shared = Arc::new(AtomicU32::new(0));
        let stderr_shared = Arc::new(Mutex::new(String::new()));
        let progress_reader = Arc::clone(&progress_shared);
        let activity_reader = Arc::clone(&activity_shared);
        let stderr_reader = Arc::clone(&stderr_shared);
        let reader_thread = std::thread::spawn(move || {
            for line_result in BufReader::new(stderr_pipe).lines() {
                let line = match line_result {
                    Ok(line) => line,
                    Err(err) => {
                        crate::log_debug(&format!("yt-dlp stderr read failed: {}", err));
                        continue;
                    }
                };
                if let Ok(mut captured) = stderr_reader.lock()
                    && captured.len() < 16_384
                {
                    if !captured.is_empty() {
                        captured.push('\n');
                    }
                    captured.push_str(&line);
                }
                if ytdlp_debug {
                    crate::log_debug(&format!("yt-dlp stderr: {}", line));
                }
                activity_reader.fetch_add(1, Ordering::Relaxed);
                if let Some(pct) = parse_ytdlp_progress_pct(&line) {
                    progress_reader.fetch_max(pct, Ordering::Relaxed);
                }
            }
        });

        let allow_early_finalize = !matches!(
            dialog_data.format,
            StreamOutputFormat::Auto | StreamOutputFormat::Mp4
        );

        let mut last_pct = 0u32;
        let mut ui_pct = 0u32;
        let mut ui_target_pct = 0u32;
        let mut last_activity = 0u32;
        let mut stalled = false;
        let mut last_progress_at = std::time::Instant::now();
        let mut reached_100_at: Option<std::time::Instant> = None;
        let mut last_focus_keepalive = std::time::Instant::now();
        crate::log_debug(&format!(
            "stream download loop start: progress_dialog={:?} url={} cache_prefix={} allow_early_finalize={}",
            progress, url, prefix, allow_early_finalize
        ));
        let _status = loop {
            match child.try_wait() {
                Ok(Some(status)) => {
                    crate::log_debug(&format!(
                        "stream download loop exit: success={} code={:?}",
                        status.success(),
                        status.code()
                    ));
                    break Some(status);
                }
                Ok(None) => {
                    if pump_messages_detect_stream_cancel(parent, progress) {
                        crate::log_debug("stream download loop: cancel detected while downloading");
                        close_progress_dialog(progress);
                        if let Err(err) = child.kill() {
                            crate::log_debug(&format!(
                                "Failed to kill cancelled yt-dlp process: {}",
                                err
                            ));
                        }
                        return;
                    }
                    let pct = progress_shared.load(Ordering::Relaxed);
                    if pct > last_pct {
                        crate::log_debug(&format!(
                            "stream download raw progress: previous={} new_raw={} ui_target_before={} ui_pct={}",
                            last_pct, pct, ui_target_pct, ui_pct
                        ));
                        last_pct = pct;
                        last_progress_at = std::time::Instant::now();
                        // Keep room for post-download finalization so 100% appears right before playback.
                        ui_target_pct = pct.min(95);
                        crate::log_debug(&format!(
                            "stream download target progress adjusted: raw={} ui_target_after={}",
                            pct, ui_target_pct
                        ));
                        if last_pct >= 100 && reached_100_at.is_none() {
                            reached_100_at = Some(std::time::Instant::now());
                            crate::log_debug("stream download raw progress reached 100%");
                        }
                    }
                    if ui_pct < ui_target_pct {
                        ui_pct = (ui_pct + 1).min(ui_target_pct);
                        report_progress(progress, ui_pct);
                    }
                    if last_focus_keepalive.elapsed() > std::time::Duration::from_millis(300) {
                        keep_stream_progress_focus(progress);
                        last_focus_keepalive = std::time::Instant::now();
                    }
                    if allow_early_finalize
                        && reached_100_at
                            .map(|t| {
                                t.elapsed()
                                    > std::time::Duration::from_secs(STREAM_POST_100_GRACE_SECS)
                            })
                            .unwrap_or(false)
                        && find_latest_downloaded_stream_file(&cache_dir, &prefix).is_some()
                    {
                        crate::log_debug(
                            "Stream download reached 100% with output present: finalizing early",
                        );
                        log_stream_focus_snapshot("stream_download.finalize_early", progress);
                        if let Err(err) = child.kill() {
                            crate::log_debug(&format!(
                                "Failed to kill finalized yt-dlp process: {}",
                                err
                            ));
                        }
                        break None;
                    }
                    let activity = activity_shared.load(Ordering::Relaxed);
                    if activity != last_activity {
                        crate::log_debug(&format!(
                            "stream download activity heartbeat: previous={} current={}",
                            last_activity, activity
                        ));
                        last_activity = activity;
                        last_progress_at = std::time::Instant::now();
                    }
                    if last_progress_at.elapsed()
                        > std::time::Duration::from_secs(STREAM_DOWNLOAD_STALL_SECS)
                    {
                        stalled = true;
                        crate::log_debug("Stream download stalled: terminating yt-dlp process");
                        log_stream_focus_snapshot("stream_download.stalled", progress);
                        if let Err(err) = child.kill() {
                            crate::log_debug(&format!(
                                "Failed to kill stalled yt-dlp process: {}",
                                err
                            ));
                        }
                        break None;
                    }
                    std::thread::sleep(std::time::Duration::from_millis(30));
                }
                Err(err) => {
                    crate::log_debug(&format!("stream download loop wait error: {}", err));
                    close_progress_dialog(progress);
                    show_error(
                        parent,
                        language,
                        &i18n::tr_f(
                            language,
                            "stream_audio.download_failed",
                            &[("err", &err.to_string())],
                        ),
                    );
                    return;
                }
            }
        };
        if _status.is_some() {
            if let Err(err) = reader_thread.join() {
                crate::log_debug(&format!("yt-dlp stderr thread join failed: {:?}", err));
            }
        } else {
            // Avoid UI hangs when yt-dlp is terminated early (finalization/stall paths).
            // The stderr reader thread will exit on its own once pipes are closed.
            crate::log_debug("Skipping yt-dlp stderr thread join after early termination");
        }
        let stderr_capture = stderr_shared
            .lock()
            .map(|s| s.clone())
            .unwrap_or_else(|_| String::new());
        if ytdlp_debug && !stderr_capture.trim().is_empty() {
            crate::log_debug(&format!(
                "yt-dlp combined stderr: {}",
                truncate_debug_text(stderr_capture.trim(), 5000)
            ));
        }

        report_progress_status(progress, &i18n::tr(language, "podcasts.loading"));
        log_stream_focus_snapshot("stream_download.after_download_before_finalize", progress);
        let primary_path = find_latest_downloaded_stream_file(&cache_dir, &prefix);
        let downloaded_path = if primary_path.is_some() {
            primary_path
        } else {
            let err = if stderr_capture.trim().is_empty() {
                "yt-dlp failed".to_string()
            } else {
                stderr_capture.trim().to_string()
            };
            let err_lc = err.to_ascii_lowercase();
            let should_retry = stalled
                || err_lc.contains("downloaded file is empty")
                || err_lc.contains("http error 404")
                || err_lc.contains("unable to download video data")
                || err_lc.contains("requested format is not available");
            if ytdlp_debug {
                crate::log_debug(&format!(
                    "yt-dlp primary output missing: stalled={} should_retry={} err={}",
                    stalled,
                    should_retry,
                    truncate_debug_text(&err, 1000)
                ));
                log_stream_cache_snapshot(&cache_dir, &prefix, "primary-missing");
            }
            if should_retry {
                crate::log_debug("Stream download failed/stalled: retrying with fallback profiles");
                report_progress(progress, 95);
                let retry_profiles: [(&str, Option<&str>); 2] = [
                    ("audio-autoip", Some("bestaudio/best")),
                    ("auto-autoip", None),
                ];
                let mut retry_path: Option<PathBuf> = None;
                let mut retry_error = String::new();
                'retry_rounds: for round in 0..2 {
                    for (idx, (profile_name, profile)) in retry_profiles.iter().enumerate() {
                        let retry_prefix = format!("{prefix}_retry{}_{}", round + 1, idx + 1);
                        let retry_template = cache_dir.join(format!("{retry_prefix}.%(ext)s"));
                        let mut retry = ytdlp_command(&ytdlp_path);
                        retry
                            .arg("--no-playlist")
                            .arg("--socket-timeout")
                            .arg(YTDLP_SOCKET_TIMEOUT_SECS)
                            .arg("--no-warnings")
                            .arg("--verbose")
                            .arg("--extractor-retries")
                            .arg("3")
                            .arg("--fragment-retries")
                            .arg("3");
                        if let Some(profile) = profile {
                            retry.arg("-f").arg(profile);
                        }
                        retry
                            .arg("-o")
                            .arg(retry_template.to_string_lossy().to_string())
                            .arg("--")
                            .arg(&url)
                            .stdout(Stdio::null())
                            .stderr(Stdio::null());
                        if ytdlp_debug {
                            crate::log_debug(&format!(
                                "yt-dlp retry start: round={} profile={} format={:?} output_template={}",
                                round + 1,
                                profile_name,
                                profile,
                                retry_template.to_string_lossy()
                            ));
                        }
                        match retry.spawn() {
                            Ok(mut retry_child) => {
                                let retry_start = std::time::Instant::now();
                                let mut status_opt = None;
                                loop {
                                    match retry_child.try_wait() {
                                        Ok(Some(status)) => {
                                            status_opt = Some(status);
                                            break;
                                        }
                                        Ok(None) => {
                                            if pump_messages_detect_stream_cancel(parent, progress)
                                            {
                                                close_progress_dialog(progress);
                                                let _unused = retry_child.kill();
                                                return;
                                            }
                                            if last_focus_keepalive.elapsed()
                                                > std::time::Duration::from_millis(300)
                                            {
                                                keep_stream_progress_focus(progress);
                                                last_focus_keepalive = std::time::Instant::now();
                                            }
                                            if retry_start.elapsed()
                                                > std::time::Duration::from_secs(
                                                    STREAM_RETRY_TIMEOUT_SECS,
                                                )
                                            {
                                                let _unused = retry_child.kill();
                                                retry_error = format!(
                                                    "yt-dlp retry timeout (round={}, profile={}) after {}s",
                                                    round + 1,
                                                    profile_name,
                                                    STREAM_RETRY_TIMEOUT_SECS
                                                );
                                                crate::log_debug(&retry_error);
                                                break;
                                            }
                                            std::thread::sleep(std::time::Duration::from_millis(
                                                30,
                                            ));
                                        }
                                        Err(e) => {
                                            retry_error = format!(
                                                "yt-dlp retry wait error (round={}, profile={}): {}",
                                                round + 1,
                                                profile_name,
                                                e
                                            );
                                            crate::log_debug(&retry_error);
                                            break;
                                        }
                                    }
                                }
                                let status = status_opt;
                                if ytdlp_debug {
                                    crate::log_debug(&format!(
                                        "yt-dlp retry status: round={} profile={} format={:?} success={} code={:?}",
                                        round + 1,
                                        profile_name,
                                        profile,
                                        status.map(|s| s.success()).unwrap_or(false),
                                        status.and_then(|s| s.code())
                                    ));
                                }
                                if let Some(found) =
                                    find_latest_downloaded_stream_file(&cache_dir, &retry_prefix)
                                {
                                    if ytdlp_debug {
                                        log_stream_cache_snapshot(
                                            &cache_dir,
                                            &retry_prefix,
                                            "retry-found",
                                        );
                                    }
                                    retry_path = Some(found);
                                    break 'retry_rounds;
                                }
                                if let Some(status) = status
                                    && !status.success()
                                {
                                    retry_error = format!(
                                        "yt-dlp retry failed (round={}, profile={}) exit_code={:?}",
                                        round + 1,
                                        profile_name,
                                        status.code()
                                    );
                                    crate::log_debug(&retry_error);
                                }
                            }
                            Err(e) => {
                                crate::log_debug(&format!(
                                    "yt-dlp retry spawn/output error (round={}, profile={}): {}",
                                    round + 1,
                                    profile_name,
                                    e
                                ));
                                if ytdlp_debug {
                                    crate::log_debug(&format!(
                                        "yt-dlp retry spawn/output error (round={}, profile={}): {}",
                                        round + 1,
                                        profile_name,
                                        e
                                    ));
                                }
                                retry_error = e.to_string();
                            }
                        }
                    }
                }
                if retry_path.is_none() {
                    if ytdlp_debug {
                        log_stream_cache_snapshot(&cache_dir, &prefix, "retry-missing");
                    }
                    let msg = if retry_error.trim().is_empty() {
                        err
                    } else {
                        retry_error
                    };
                    let message =
                        i18n::tr_f(language, "stream_audio.download_failed", &[("err", &msg)]);
                    close_progress_dialog(progress);
                    show_error(parent, language, &message);
                    return;
                }
                retry_path
            } else {
                let message =
                    i18n::tr_f(language, "stream_audio.download_failed", &[("err", &err)]);
                close_progress_dialog(progress);
                show_error(parent, language, &message);
                return;
            }
        };

        let Some(downloaded_path) = downloaded_path else {
            close_progress_dialog(progress);
            show_error(
                parent,
                language,
                &i18n::tr(language, "stream_audio.no_output"),
            );
            return;
        };
        let final_path = if let Some(convert_settings) = dialog_data
            .format
            .as_audio_convert_settings(dialog_data.quality)
        {
            let target_ext = dialog_data.format.extension().unwrap_or("mp3");
            let same_extension = downloaded_path
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.eq_ignore_ascii_case(target_ext))
                .unwrap_or(false);
            let must_reencode_mp3 = matches!(
                (dialog_data.format, dialog_data.quality),
                (
                    StreamOutputFormat::Mp3,
                    StreamQualitySelection::BitrateKbps(_)
                )
            );
            if same_extension && !must_reencode_mp3 {
                crate::log_debug(&format!(
                    "stream conversion skipped: downloaded_path={} target_ext={} same_extension={} must_reencode_mp3={}",
                    downloaded_path.to_string_lossy(),
                    target_ext,
                    same_extension,
                    must_reencode_mp3
                ));
                close_progress_dialog(progress);
                downloaded_path.clone()
            } else {
                let converted_path = downloaded_path.with_extension(target_ext);
                let convert_cancel = Arc::new(AtomicBool::new(false));
                crate::log_debug(&format!(
                    "stream conversion start: input={} output={} target_ext={} quality={:?} format={:?}",
                    downloaded_path.to_string_lossy(),
                    converted_path.to_string_lossy(),
                    target_ext,
                    dialog_data.quality,
                    dialog_data.format
                ));
                report_progress_status(
                    progress,
                    &i18n::tr(language, "stream_audio.progress_converting"),
                );
                report_progress(progress, 0);
                log_stream_focus_snapshot("stream_conversion.started", progress);
                let mut last_pump = std::time::Instant::now();
                let mut last_reported_pct = 0u32;
                let convert_cancel_for_cb = Arc::clone(&convert_cancel);
                let mut progress_cb = |pct: u32| {
                    let normalized = (pct / 100).min(100);
                    crate::log_debug(&format!(
                        "stream conversion progress callback: raw_pct={} normalized_pct={} last_reported_pct={}",
                        pct, normalized, last_reported_pct
                    ));
                    if normalized > last_reported_pct {
                        last_reported_pct = normalized;
                        report_progress(progress, normalized);
                    }
                    // Keep window responsive during in-process conversion on slower machines.
                    if last_pump.elapsed() >= std::time::Duration::from_millis(50) {
                        let cancelled = pump_messages_detect_stream_cancel(parent, progress);
                        crate::log_debug(&format!(
                            "stream conversion ui pump: cancelled={} elapsed_ms={}",
                            cancelled,
                            last_pump.elapsed().as_millis()
                        ));
                        log_stream_focus_snapshot("stream_conversion.ui_pump", progress);
                        if cancelled {
                            convert_cancel_for_cb.store(true, Ordering::Relaxed);
                        }
                        last_pump = std::time::Instant::now();
                    }
                };
                let convert_result = crate::ffmpeg_export::convert_audio_file(
                    &downloaded_path,
                    &converted_path,
                    &convert_settings,
                    Some(Arc::clone(&convert_cancel)),
                    Some(&mut progress_cb),
                );
                crate::log_debug(&format!(
                    "stream conversion result: success={} output_exists={} output={}",
                    convert_result.is_ok(),
                    converted_path.exists(),
                    converted_path.to_string_lossy()
                ));
                report_progress(progress, 100);
                close_progress_dialog(progress);
                match convert_result {
                    Ok(()) => {
                        crate::log_if_err!(std::fs::remove_file(&downloaded_path));
                        converted_path
                    }
                    Err(err) => {
                        if convert_cancel.load(Ordering::Relaxed) || err == "Conversion canceled." {
                            crate::log_debug("stream conversion cancelled by user");
                            crate::log_if_err!(std::fs::remove_file(&converted_path));
                            return;
                        }
                        show_error(
                            parent,
                            language,
                            &i18n::tr_f(language, "stream_audio.convert_failed", &[("err", &err)]),
                        );
                        return;
                    }
                }
            }
        } else {
            report_progress(progress, 100);
            close_progress_dialog(progress);
            downloaded_path
        };

        let playback_path = if let Some(title) = stream_title.as_ref() {
            let ext = final_path
                .extension()
                .and_then(|e| e.to_str())
                .filter(|e| !e.trim().is_empty())
                .unwrap_or("mp3");
            let target = unique_stream_named_path(&cache_dir, title, ext);
            if target != final_path {
                match std::fs::rename(&final_path, &target) {
                    Ok(()) => target,
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Stream title rename failed ({} -> {}): {}",
                            final_path.to_string_lossy(),
                            target.to_string_lossy(),
                            err
                        ));
                        final_path.clone()
                    }
                }
            } else {
                final_path.clone()
            }
        } else {
            final_path.clone()
        };

        crate::queue_audio_files_and_play(parent, vec![playback_path.clone()]);
        crate::editor_manager::mark_current_document_from_rss(parent, true);
        let episode_title = stream_title.or_else(|| {
            playback_path
                .file_stem()
                .and_then(|s| s.to_str())
                .map(|s| s.to_string())
        });
        crate::set_active_podcast_episode_info(
            parent,
            Some(url),
            episode_title,
            Some(playback_path),
        );
        update_youtube_return_context_from_selection(
            parent,
            &input,
            collection_url.as_deref(),
            collection_page,
        );
        crate::menu::update_playback_menu(parent, true);
        return;
    }
}

fn ensure_ytdlp_available(
    hwnd: HWND,
    language: Language,
    labels: &Labels,
    initial_progress: Option<HWND>,
) -> Result<Option<PathBuf>, String> {
    restore_stream_parent_focus(hwnd);
    let local_path = settings_dir().join(YTDLP_EXE_NAME);
    if local_path.exists() {
        let progress = initial_progress;
        if let Some(progress) = progress {
            report_progress_status(progress, &i18n::tr(language, "podcasts.loading"));
            keep_stream_progress_focus(progress);
        }
        check_ytdlp_update(hwnd, language, labels, &local_path);
        if let Some(progress) = progress {
            keep_stream_progress_focus(progress);
        }
        restore_stream_parent_focus(hwnd);
        return Ok(Some(local_path));
    }
    if let Some(path_in_system) = find_ytdlp_in_path() {
        let progress = initial_progress;
        if let Some(progress) = progress {
            report_progress_status(progress, &i18n::tr(language, "podcasts.loading"));
            keep_stream_progress_focus(progress);
        }
        crate::log_debug(&format!(
            "{} {}",
            labels.ytdlp_found_in_path,
            path_in_system.display()
        ));
        let selected_path =
            check_ytdlp_path_update(hwnd, language, labels, &path_in_system, &local_path);
        if let Some(progress) = progress {
            keep_stream_progress_focus(progress);
        }
        restore_stream_parent_focus(hwnd);
        return Ok(Some(selected_path));
    }
    if let Some(progress) = initial_progress {
        close_progress_dialog(progress);
    }
    let title = to_wide(&confirm_title(language));
    let message = to_wide(&labels.ytdlp_prompt_download);
    let response = crate::message_box_w_safe(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_YESNO | MB_ICONQUESTION,
    );
    restore_stream_parent_focus(hwnd);
    if response != IDYES {
        return Ok(None);
    }
    let progress = open_progress_dialog(
        hwnd,
        language,
        "stream_audio.ytdlp_progress_title",
        "stream_audio.ytdlp_progress_downloading",
        false,
    );
    let download_result =
        download_ytdlp_with_progress(&local_path, |pct| report_progress(progress, pct));
    close_progress_dialog(progress);
    restore_stream_parent_focus(hwnd);
    match download_result {
        Ok(()) => Ok(Some(local_path)),
        Err(err) => Err(err),
    }
}

fn find_ytdlp_in_path() -> Option<PathBuf> {
    let path_var = std::env::var_os("PATH")?;
    for dir in std::env::split_paths(&path_var) {
        let candidate = dir.join(YTDLP_EXE_NAME);
        if candidate.is_file() {
            return Some(candidate);
        }
    }
    None
}

fn check_ytdlp_path_update(
    hwnd: HWND,
    language: Language,
    labels: &Labels,
    path_in_system: &Path,
    local_path: &Path,
) -> PathBuf {
    restore_stream_parent_focus(hwnd);
    if YTDLP_UPDATE_CHECKED.swap(true, Ordering::SeqCst) {
        return path_in_system.to_path_buf();
    }
    let installed = match ytdlp_installed_version(path_in_system) {
        Ok(version) => version,
        Err(err) => {
            crate::log_debug(&format!("yt-dlp PATH version read failed: {}", err));
            return path_in_system.to_path_buf();
        }
    };
    let latest = match fetch_latest_ytdlp_version() {
        Ok(version) => version,
        Err(err) => {
            crate::log_debug(&format!("yt-dlp latest version fetch failed: {}", err));
            return path_in_system.to_path_buf();
        }
    };
    if compare_versions(&installed, &latest) != Some(CmpOrdering::Less) {
        return path_in_system.to_path_buf();
    }

    let prompt = i18n::tr_f(
        language,
        "youtube.ytdlp_path_update_prompt",
        &[("current", &installed), ("latest", &latest)],
    );
    crate::log_debug(&format!(
        "{} / {}",
        labels.ytdlp_path_update_download_local, labels.ytdlp_path_update_keep_system
    ));
    let title = to_wide(&confirm_title(language));
    let message = to_wide(&prompt);
    let response = crate::message_box_w_safe(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_YESNO | MB_ICONQUESTION,
    );
    restore_stream_parent_focus(hwnd);
    if response != IDYES {
        return path_in_system.to_path_buf();
    }

    let progress = open_progress_dialog(
        hwnd,
        language,
        "stream_audio.ytdlp_progress_title",
        "stream_audio.ytdlp_progress_downloading",
        false,
    );
    let result = download_ytdlp_with_progress(local_path, |pct| report_progress(progress, pct));
    close_progress_dialog(progress);
    restore_stream_parent_focus(hwnd);
    if let Err(err) = result {
        crate::log_debug(&format!(
            "yt-dlp local download from PATH prompt failed: {}",
            err
        ));
        show_error(hwnd, language, &labels.ytdlp_update_failed);
        return path_in_system.to_path_buf();
    }
    local_path.to_path_buf()
}

fn check_ytdlp_update(hwnd: HWND, language: Language, labels: &Labels, path: &Path) {
    restore_stream_parent_focus(hwnd);
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
    let response = crate::message_box_w_safe(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_YESNO | MB_ICONQUESTION,
    );
    restore_stream_parent_focus(hwnd);
    if response != IDYES {
        return;
    }
    let progress = open_progress_dialog(
        hwnd,
        language,
        "stream_audio.ytdlp_progress_title",
        "stream_audio.ytdlp_progress_downloading",
        false,
    );
    let result = download_ytdlp_with_progress(path, |pct| report_progress(progress, pct));
    close_progress_dialog(progress);
    restore_stream_parent_focus(hwnd);
    if let Err(err) = result {
        crate::log_debug(&format!("yt-dlp update failed: {}", err));
        show_error(hwnd, language, &labels.ytdlp_update_failed);
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

fn download_ytdlp_with_progress<F: FnMut(u32)>(
    target: &Path,
    mut progress_cb: F,
) -> Result<(), String> {
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
    let mut written = 0u64;
    let mut buf = [0u8; 64 * 1024];
    let mut last_reported = 0u32;
    loop {
        let read = std::io::Read::read(&mut resp, &mut buf).map_err(|err| err.to_string())?;
        if read == 0 {
            break;
        }
        std::io::Write::write_all(&mut file, &buf[..read]).map_err(|err| err.to_string())?;
        written = written.saturating_add(read as u64);
        if let Some(expected) = expected_len
            && expected > 0
        {
            let pct = ((written.saturating_mul(100)) / expected).min(100) as u32;
            if pct > last_reported {
                last_reported = pct;
                progress_cb(pct);
            }
        }
    }
    file.sync_all().map_err(|err| err.to_string())?;
    if last_reported < 100 {
        progress_cb(100);
    }
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
    YtdlpUnavailable,
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
    let ytdlp_path = ytdlp_path.ok_or(ImportError::YtdlpUnavailable)?;
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
        ImportError::YtdlpUnavailable => labels.ytdlp_download_failed,
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
