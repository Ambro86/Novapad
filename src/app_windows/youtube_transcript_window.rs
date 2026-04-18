use std::cmp::Ordering as CmpOrdering;
use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, Read};
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Mutex};

use encoding_rs::WINDOWS_1252;
use windows::Win32::Foundation::{BOOL, HINSTANCE, HWND, LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, COLOR_HIGHLIGHT, COLOR_WINDOW, DT_CALCRECT, DT_NOPREFIX, DT_WORDBREAK, DrawTextW,
    EndPaint, FillRect, GetDC, HBRUSH, HFONT, HGDIOBJ, InvalidateRect, PAINTSTRUCT, ReleaseDC,
    SelectObject, SetBkMode, TRANSPARENT,
};
use windows::Win32::System::Threading::CREATE_NO_WINDOW;
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::RichEdit::{CHARRANGE, EM_EXSETSEL};
use windows::Win32::UI::Controls::{
    BST_CHECKED, EM_SCROLLCARET, EM_SETSEL, SetScrollInfo, ShowScrollBar, WC_BUTTON, WC_COMBOBOXW,
    WC_LISTBOXW, WC_STATIC,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, GetKeyState, SetFocus, VK_BACK, VK_CONTROL, VK_DOWN, VK_END, VK_ESCAPE,
    VK_HOME, VK_LEFT, VK_NEXT, VK_PRIOR, VK_RETURN, VK_RIGHT, VK_TAB, VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, CB_ADDSTRING,
    CB_GETCURSEL, CB_RESETCONTENT, CB_SETCURSEL, CBN_SELCHANGE, CBS_DROPDOWNLIST, CREATESTRUCTW,
    CW_USEDEFAULT, CreatePopupMenu, CreateWindowExW, DLGC_WANTARROWS, DefWindowProcW, DestroyMenu,
    DispatchMessageW, EN_CHANGE, ES_AUTOHSCROLL, ES_MULTILINE, ES_READONLY, GWLP_USERDATA,
    GetCursorPos, GetForegroundWindow, GetParent, GetScrollInfo, GetWindowLongPtrW, HMENU,
    HWND_TOPMOST, IDC_ARROW, IDYES, IsChild, IsDialogMessageW, IsWindow, LB_ADDSTRING,
    LB_GETCURSEL, LB_RESETCONTENT, LB_SETCURSEL, LBN_SELCHANGE, LBS_HASSTRINGS, LBS_NOTIFY,
    LoadCursorW, MB_ICONQUESTION, MB_YESNO, MF_STRING, MSG, PM_REMOVE, PeekMessageW, PostMessageW,
    RegisterClassW, SB_BOTTOM, SB_LINEDOWN, SB_LINEUP, SB_PAGEDOWN, SB_PAGEUP, SB_THUMBPOSITION,
    SB_THUMBTRACK, SB_TOP, SB_VERT, SCROLLINFO, SIF_PAGE, SIF_POS, SIF_RANGE, SIF_TRACKPOS,
    SW_HIDE, SW_SHOW, SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, SetWindowPos, SetWindowTextW, ShowWindow, TPM_NONOTIFY, TPM_RETURNCMD,
    TrackPopupMenu, TranslateMessage, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU,
    WM_CREATE, WM_DESTROY, WM_GETDLGCODE, WM_KEYDOWN, WM_LBUTTONDOWN, WM_MOUSEWHEEL, WM_NCDESTROY,
    WM_PAINT, WM_SETFOCUS, WM_SETFONT, WM_SIZE, WM_VSCROLL, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE,
    WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::{EM_REPLACESEL, screen_reader_speak, to_wide, to_wide_normalized};
use crate::app_windows::prompt_window;
use crate::editor_manager::get_edit_text;
use crate::i18n;
use crate::settings::{
    Language, StreamFavorite, clear_ytdlp_site_credentials, confirm_title,
    get_ytdlp_site_credentials, save_settings, set_ytdlp_site_credentials, settings_dir,
};
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
const STREAM_ID_OPEN_COMMENTS: usize = 9317;

#[inline]
fn ignore_bool(_value: bool) {}
const STREAM_ID_QUALITY: usize = 9318;
const STREAM_TRACK_ID_COMBO: usize = 9321;
const STREAM_TRACK_ID_OK: usize = 9322;
const STREAM_TRACK_ID_CANCEL: usize = 9323;
const STREAM_DIALOG_CLASS_NAME: &str = "SonarpadStreamAudio";
const STREAM_TRACK_DIALOG_CLASS_NAME: &str = "SonarpadStreamAudioTrack";
const YOUTUBE_COMMENTS_DIALOG_CLASS_NAME: &str = "SonarpadYouTubeComments";
const YOUTUBE_COMMENTS_VIEW_CLASS_NAME: &str = "SonarpadYouTubeCommentsView";
const YOUTUBE_COMMENTS_INITIAL_PARENT_LIMIT: usize = 50;
const YOUTUBE_COMMENTS_ID_VIEW: usize = 9331;
const YOUTUBE_COMMENTS_ID_ACCESSIBILITY_PROXY: usize = 9332;
const YOUTUBE_COMMENTS_ID_CLOSE: usize = 9333;
const YOUTUBE_COMMENTS_ID_SEARCH_EDIT: usize = 9334;
const YOUTUBE_COMMENTS_ID_SEARCH_BUTTON: usize = 9335;
const YOUTUBE_COMMENTS_ID_SEARCH_LABEL: usize = 9336;
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
const YTDLP_PROGRESS_TEMPLATE_PREFIX: &str = "SONARPAD_PROGRESS";
const FORCE_YTDLP_AUTH_PROMPT_FOR_TESTING: bool = false;
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
    rate_limited: String,
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
    let rate_limited = i18n::tr(language, "youtube.rate_limited");
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
        rate_limited: if rate_limited == "youtube.rate_limited" {
            i18n::tr(language, "youtube.import_error")
        } else {
            rate_limited
        },
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
    crate::log_debug(&format!(
        "post_focus_editor: parent={:?} foreground_before={:?} focus_before={:?}",
        parent,
        crate::get_foreground_window_safe(),
        crate::get_focus_safe()
    ));
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
            WM_SETFOCUS => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ImportState;
                if ptr.is_null() {
                    return LRESULT(0);
                }
                let state = &mut *ptr;
                SetFocus(state.url_edit);
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
    let mut parent = HWND(0);
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);

    let state_ok = with_import_state(hwnd, |state| {
        parent = state.parent;
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
            let owner = if parent.0 != 0 { parent } else { hwnd };
            show_error(owner, language, &error_message(language, &err));
            ShowWindow(hwnd, SW_SHOW);
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
    let mut parent = HWND(0);
    let mut edit = HWND(0);
    let mut ok_button = HWND(0);
    let mut load_button = HWND(0);
    let mut combo = HWND(0);
    let mut timestamp = HWND(0);
    let mut status = HWND(0);

    let state_ok = with_import_state(hwnd, |state| {
        parent = state.parent;
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
            let owner = if parent.0 != 0 { parent } else { hwnd };
            show_error(owner, language, &error_message(language, &err));
            ShowWindow(hwnd, SW_SHOW);
            SetForegroundWindow(hwnd);
            if edit.0 != 0 {
                SetFocus(edit);
            } else if combo.0 != 0 {
                SetFocus(combo);
            }
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

#[derive(Clone)]
struct YoutubeComment {
    id: String,
    parent: String,
    author: String,
    text: String,
    time_text: String,
}

struct YoutubeCommentsDialogInit {
    parent: HWND,
    language: Language,
    title: String,
    comments: Vec<YoutubeComment>,
    show_load_all_action: bool,
    action_result: Arc<Mutex<YoutubeCommentsDialogAction>>,
    flat_selection: Option<YoutubeFlatSelectionInit>,
    flat_search: Option<YoutubeFlatSearchInit>,
    flat_close_button_label: Option<String>,
    flat_context_action:
        Option<crate::app_windows::interpreter_select_window::InterpreterContextAction>,
    right_arrow_accepts_selection: bool,
    left_arrow_closes: bool,
}

#[derive(Default)]
struct YoutubeCommentsDialogMode {
    show_load_all_action: bool,
    flat_selection: Option<YoutubeFlatSelectionInit>,
    flat_search: Option<YoutubeFlatSearchInit>,
    flat_close_button_label: Option<String>,
    flat_context_action:
        Option<crate::app_windows::interpreter_select_window::InterpreterContextAction>,
    right_arrow_accepts_selection: bool,
    left_arrow_closes: bool,
}

#[derive(Clone, Copy, Default, Eq, PartialEq)]
enum YoutubeCommentsDialogAction {
    #[default]
    None,
    LoadAll,
}

#[derive(Clone)]
enum YoutubeCommentRowKind {
    Comment { comment_index: usize },
    LoadAllComments,
}

#[derive(Clone)]
struct YoutubeCommentRow {
    kind: YoutubeCommentRowKind,
    depth: usize,
    height: i32,
    has_children: bool,
    expanded: bool,
}

struct YoutubeCommentsDialogState {
    language: Language,
    title_label: HWND,
    search_label: HWND,
    view: HWND,
    accessibility_proxy: HWND,
    close_button: HWND,
    search_edit: HWND,
    search_button: HWND,
    font: HFONT,
    comments: Vec<YoutubeComment>,
    root_indices: Vec<usize>,
    children_by_parent: HashMap<String, Vec<usize>>,
    expanded_comment_ids: HashSet<String>,
    show_load_all_action: bool,
    action_result: Arc<Mutex<YoutubeCommentsDialogAction>>,
    rows: Vec<YoutubeCommentRow>,
    selected_row: usize,
    scroll_offset: i32,
    content_height: i32,
    flat_list_mode: bool,
    flat_selection_result: Option<Arc<Mutex<Option<String>>>>,
    flat_search_result: Option<Arc<Mutex<Option<String>>>>,
    flat_context_action:
        Option<crate::app_windows::interpreter_select_window::InterpreterContextAction>,
    right_arrow_accepts_selection: bool,
    left_arrow_closes: bool,
}

#[derive(Clone)]
pub(crate) struct MultilineSelectionItem {
    pub(crate) id: String,
    pub(crate) title: String,
    pub(crate) description: Option<String>,
}

pub(crate) struct MultilineSearchOptions {
    pub(crate) initial_query: String,
    pub(crate) search_button_label: String,
    pub(crate) show_search_edit: bool,
    pub(crate) context_action:
        Option<crate::app_windows::interpreter_select_window::InterpreterContextAction>,
    pub(crate) right_arrow_accepts_selection: bool,
    pub(crate) left_arrow_closes: bool,
}

struct YoutubeFlatSelectionInit {
    initial_selected_id: Option<String>,
    result: Arc<Mutex<Option<String>>>,
}

struct YoutubeFlatSearchInit {
    initial_query: String,
    button_label: String,
    result: Arc<Mutex<Option<String>>>,
    show_edit: bool,
}

pub(crate) enum MultilineSelectionResult {
    Cancelled,
    Selected(String),
    Search(String),
}

struct ResolvedStreamSelection {
    url: String,
    collection_url: Option<String>,
    collection_page: Option<usize>,
    selected_label: Option<String>,
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
    ytdlp_path: &Path,
    entries: &[StreamCollectionEntry],
    has_previous: bool,
    has_more: bool,
    initial_selected_label: Option<&str>,
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
    let ytdlp_path = Arc::new(ytdlp_path.to_path_buf());
    let add_to_favorites_action =
        crate::app_windows::interpreter_select_window::InterpreterContextAction {
            label: tr_or(
                language,
                "stream_audio.add_to_favorites",
                "Add to favorites",
            ),
            ctrl_c_shortcut: false,
            enabled: {
                let favorite_candidates = Arc::clone(&favorite_candidates);
                Arc::new(move |selected: &str| {
                    favorite_candidates
                        .iter()
                        .find(|entry| entry.label == selected)
                        .map(|entry| is_youtube_collection_url(&entry.url))
                        .unwrap_or(false)
                })
            },
            handler: {
                let favorite_candidates = Arc::clone(&favorite_candidates);
                Arc::new(move |selected: String| {
                    if let Some(entry) = favorite_candidates.iter().find(|entry| {
                        entry.label == selected && is_youtube_collection_url(&entry.url)
                    }) {
                        add_stream_favorite(parent, entry.label.clone(), entry.url.clone());
                    }
                })
            },
        };
    let view_comments_action =
        crate::app_windows::interpreter_select_window::InterpreterContextAction {
            label: tr_or(language, "stream_audio.view_comments", "View comments"),
            ctrl_c_shortcut: false,
            enabled: {
                let favorite_candidates = Arc::clone(&favorite_candidates);
                Arc::new(move |selected: &str| {
                    favorite_candidates
                        .iter()
                        .find(|entry| entry.label == selected)
                        .map(|entry| {
                            !is_youtube_collection_url(&entry.url)
                                && extract_video_id(&entry.url).is_some()
                        })
                        .unwrap_or(false)
                })
            },
            handler: {
                let favorite_candidates = Arc::clone(&favorite_candidates);
                let ytdlp_path = Arc::clone(&ytdlp_path);
                Arc::new(move |selected: String| {
                    if let Some(entry) = favorite_candidates
                        .iter()
                        .find(|entry| entry.label == selected)
                    {
                        show_youtube_comments_for_stream_entry(
                            parent,
                            language,
                            &ytdlp_path,
                            entry,
                        );
                    }
                })
            },
        };
    crate::app_windows::interpreter_select_window::select_interpreter_with_context_actions_without_parent_restore_on_accept(
        parent,
        labels,
        language,
        i18n::tr(language, "stream_audio.prompt_title"),
        initial_selected_label.map(ToOwned::to_owned),
        vec![view_comments_action, add_to_favorites_action],
    )
}

fn show_youtube_comments_for_stream_entry(
    parent: HWND,
    language: Language,
    ytdlp_path: &Path,
    entry: &StreamCollectionEntry,
) {
    let foreground = crate::get_foreground_window_safe();
    let ui_parent = {
        let mut class_buf = [0u16; 128];
        let class_len = if crate::is_window_handle_valid(foreground) {
            crate::get_class_name_w_safe(foreground, &mut class_buf)
        } else {
            0
        };
        let foreground_is_stream_selector = class_len > 0
            && String::from_utf16_lossy(&class_buf[..class_len as usize])
                == "SonarpadInterpreterSelect";
        if foreground_is_stream_selector {
            foreground
        } else {
            parent
        }
    };
    crate::log_debug(&format!(
        "YT comments choose parent entry_parent={:?} ui_parent={:?} foreground={:?} focus={:?}",
        parent,
        ui_parent,
        foreground,
        crate::get_focus_safe()
    ));
    let restore_video_selection = || {
        crate::log_debug(&format!(
            "YT comments restore selection parent={:?} ui_parent={:?} focus_before={:?}",
            parent,
            ui_parent,
            crate::get_focus_safe()
        ));
        if crate::is_window_handle_valid(ui_parent) {
            crate::bring_window_to_foreground(ui_parent);
            if !crate::app_windows::interpreter_select_window::restore_interpreter_select_focus(
                ui_parent,
            ) {
                restore_stream_dialog_focus(ui_parent);
            }
        } else {
            restore_stream_dialog_focus(parent);
        }
    };
    let ytdlp_path = ytdlp_path.to_path_buf();
    let entry_url = entry.url.clone();
    let mut max_parent_comments = Some(YOUTUBE_COMMENTS_INITIAL_PARENT_LIMIT);
    loop {
        let progress = open_progress_dialog(
            ui_parent,
            language,
            "stream_audio.progress_title",
            "stream_audio.comments_loading",
            false,
        );
        let worker_ytdlp_path = ytdlp_path.clone();
        let worker_entry_url = entry_url.clone();
        let worker_max_parent_comments = max_parent_comments;
        let worker = std::thread::spawn(move || {
            fetch_youtube_comments_with_ytdlp(
                &worker_ytdlp_path,
                &worker_entry_url,
                worker_max_parent_comments,
            )
        });
        while !worker.is_finished() {
            ignore_bool(pump_messages_detect_stream_cancel(parent, progress));
            std::thread::sleep(std::time::Duration::from_millis(30));
        }
        close_progress_dialog(progress);
        let comments = match worker.join() {
            Ok(Ok(comments)) => comments,
            Ok(Err(err)) => {
                show_error(
                    parent,
                    language,
                    &i18n::tr_f(language, "stream_audio.download_failed", &[("err", &err)]),
                );
                restore_video_selection();
                return;
            }
            Err(_) => {
                show_error(
                    parent,
                    language,
                    &tr_or(
                        language,
                        "stream_audio.comments_load_failed",
                        "Failed to load video comments.",
                    ),
                );
                restore_video_selection();
                return;
            }
        };
        if comments.is_empty() {
            show_error(
                parent,
                language,
                &tr_or(
                    language,
                    "stream_audio.no_comments",
                    "No comments are available for this video.",
                ),
            );
            restore_video_selection();
            return;
        }
        let root_comment_count = comments
            .iter()
            .filter(|comment| comment.parent.eq_ignore_ascii_case("root"))
            .count();
        let show_load_all_action = max_parent_comments
            .map(|limit| root_comment_count >= limit)
            .unwrap_or(false);
        let action = open_youtube_comments_window_with_mode(
            ui_parent,
            language,
            entry.label.clone(),
            comments,
            YoutubeCommentsDialogMode {
                show_load_all_action,
                ..Default::default()
            },
        );
        if action == YoutubeCommentsDialogAction::LoadAll && max_parent_comments.is_some() {
            max_parent_comments = None;
            continue;
        }
        break;
    }
}

pub(crate) fn select_multiline_items_with_search(
    parent: HWND,
    language: Language,
    title: String,
    items: Vec<MultilineSelectionItem>,
    initial_selected_id: Option<String>,
    search_options: MultilineSearchOptions,
) -> MultilineSelectionResult {
    let comments = items
        .into_iter()
        .map(|item| {
            let title = normalize_comment_text(&item.title);
            let description = item
                .description
                .as_deref()
                .map(normalize_comment_text)
                .unwrap_or_default();
            YoutubeComment {
                id: item.id,
                parent: "root".to_string(),
                author: if title.is_empty() {
                    "Elemento".to_string()
                } else {
                    title
                },
                text: description,
                time_text: String::new(),
            }
        })
        .collect::<Vec<_>>();
    let selection_result = Arc::new(Mutex::new(None));
    let search_result = Arc::new(Mutex::new(None));
    open_youtube_comments_window_with_mode(
        parent,
        language,
        title,
        comments,
        YoutubeCommentsDialogMode {
            show_load_all_action: false,
            flat_selection: Some(YoutubeFlatSelectionInit {
                initial_selected_id,
                result: Arc::clone(&selection_result),
            }),
            flat_search: Some(YoutubeFlatSearchInit {
                initial_query: search_options.initial_query,
                button_label: search_options.search_button_label,
                result: Arc::clone(&search_result),
                show_edit: search_options.show_search_edit,
            }),
            flat_close_button_label: None,
            flat_context_action: search_options.context_action,
            right_arrow_accepts_selection: search_options.right_arrow_accepts_selection,
            left_arrow_closes: search_options.left_arrow_closes,
        },
    );
    if let Some(value) = selection_result
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .clone()
    {
        MultilineSelectionResult::Selected(value)
    } else if let Some(value) = search_result
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .clone()
    {
        MultilineSelectionResult::Search(value)
    } else {
        MultilineSelectionResult::Cancelled
    }
}

fn fetch_youtube_comments_with_ytdlp(
    ytdlp_path: &Path,
    url: &str,
    max_parent_comments: Option<usize>,
) -> Result<Vec<YoutubeComment>, String> {
    let mut cmd = ytdlp_command(ytdlp_path);
    cmd.arg("--no-playlist")
        .arg("--skip-download")
        .arg("--write-comments")
        .arg("--dump-single-json")
        .arg("--no-warnings");
    if let Some(limit) = max_parent_comments {
        cmd.arg("--extractor-args")
            .arg(format!("youtube:max_comments=all,{limit},all,all"));
    }
    let output = cmd
        .arg("--")
        .arg(url)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|e| e.to_string())?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if stderr.is_empty() {
            "yt-dlp comments extraction failed".to_string()
        } else {
            stderr
        });
    }
    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).map_err(|e| e.to_string())?;
    let Some(items) = json.get("comments").and_then(|value| value.as_array()) else {
        return Ok(Vec::new());
    };
    let mut comments = Vec::with_capacity(items.len());
    for item in items {
        let Some(id) = item.get("id").and_then(|value| value.as_str()) else {
            continue;
        };
        let text = item
            .get("text")
            .and_then(|value| value.as_str())
            .map(str::trim)
            .unwrap_or("");
        if text.is_empty() {
            continue;
        }
        let parent = item
            .get("parent")
            .and_then(|value| value.as_str())
            .filter(|value| !value.trim().is_empty())
            .unwrap_or("root");
        let author = item
            .get("author")
            .and_then(|value| value.as_str())
            .map(str::trim)
            .filter(|value| !value.is_empty())
            .unwrap_or("Unknown");
        let time_text = item
            .get("_time_text")
            .and_then(|value| value.as_str())
            .map(str::trim)
            .unwrap_or("");
        comments.push(YoutubeComment {
            id: id.to_string(),
            parent: parent.to_string(),
            author: author.to_string(),
            text: text.to_string(),
            time_text: time_text.to_string(),
        });
    }
    Ok(comments)
}

fn open_youtube_comments_window_with_mode(
    parent: HWND,
    language: Language,
    title: String,
    comments: Vec<YoutubeComment>,
    mode: YoutubeCommentsDialogMode,
) -> YoutubeCommentsDialogAction {
    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
    let class_name = to_wide(YOUTUBE_COMMENTS_DIALOG_CLASS_NAME);
    let view_class_name = to_wide(YOUTUBE_COMMENTS_VIEW_CLASS_NAME);
    let wc = windows::Win32::UI::WindowsAndMessaging::WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(youtube_comments_dialog_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    crate::register_class_w_safe(&wc);
    let view_wc = windows::Win32::UI::WindowsAndMessaging::WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(view_class_name.as_ptr()),
        lpfnWndProc: Some(youtube_comments_view_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    crate::register_class_w_safe(&view_wc);

    let action_result = Arc::new(Mutex::new(YoutubeCommentsDialogAction::None));
    let init = Box::new(YoutubeCommentsDialogInit {
        parent,
        language,
        title: plain_label(&title),
        comments,
        show_load_all_action: mode.show_load_all_action,
        action_result: Arc::clone(&action_result),
        flat_selection: mode.flat_selection,
        flat_search: mode.flat_search,
        flat_close_button_label: mode.flat_close_button_label,
        flat_context_action: mode.flat_context_action,
        right_arrow_accepts_selection: mode.right_arrow_accepts_selection,
        left_arrow_closes: mode.left_arrow_closes,
    });
    let window_title = if init.flat_selection.is_some() {
        init.title.clone()
    } else {
        tr_or(
            language,
            "stream_audio.comments_window_title",
            "Video comments",
        )
    };
    let title = to_wide(&window_title);
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            820,
            560,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };
    if hwnd.0 == 0 {
        return YoutubeCommentsDialogAction::None;
    }

    unsafe {
        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);
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
            if msg.message == WM_KEYDOWN {
                let focus = GetFocus();
                crate::log_debug(&format!(
                    "YT comments loop WM_KEYDOWN key={} focus={:?} hwnd={:?} msg_hwnd={:?}",
                    msg.wParam.0, focus, hwnd, msg.hwnd
                ));
                let search_enter_handled = with_youtube_comments_state(hwnd, |state| {
                    focus == state.search_edit && msg.wParam.0 as u32 == VK_RETURN.0 as u32
                })
                .unwrap_or(false);
                if search_enter_handled && trigger_youtube_comments_search(hwnd) {
                    continue;
                }
                let ctrl_c_handled = with_youtube_comments_state(hwnd, |state| {
                    msg.wParam.0 as u32 == 'C' as u32
                        && (GetKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000) != 0
                        && focus != state.search_edit
                        && trigger_youtube_comments_context_action(hwnd)
                })
                .unwrap_or(false);
                if ctrl_c_handled {
                    continue;
                }
                let tab_to_list_handled = with_youtube_comments_state(hwnd, |state| {
                    if msg.wParam.0 as u32 != VK_TAB.0 as u32
                        || focus != state.close_button
                        || !state.flat_list_mode
                    {
                        return false;
                    }
                    crate::log_debug(&format!(
                        "YT comments loop rerouting TAB from close button to proxy selected_row={}",
                        state.selected_row
                    ));
                    sync_youtube_comments_accessibility_proxy_selection(state, true);
                    announce_selected_youtube_comment(state);
                    true
                })
                .unwrap_or(false);
                if tab_to_list_handled {
                    continue;
                }
                let back_to_previous_handled = with_youtube_comments_state(hwnd, |state| {
                    msg.wParam.0 as u32 == VK_BACK.0 as u32
                        && state.flat_list_mode
                        && focus != state.search_edit
                        && focus != state.search_button
                        && focus != state.close_button
                })
                .unwrap_or(false);
                if back_to_previous_handled {
                    crate::log_debug("YT comments loop handling BACKSPACE as close");
                    crate::send_message_w_safe(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(YOUTUBE_COMMENTS_ID_CLOSE),
                        LPARAM(0),
                    );
                    continue;
                }
                let left_arrow_close_handled = with_youtube_comments_state(hwnd, |state| {
                    msg.wParam.0 as u32 == VK_LEFT.0 as u32
                        && state.flat_list_mode
                        && state.left_arrow_closes
                        && focus != state.search_edit
                        && focus != state.search_button
                        && focus != state.close_button
                })
                .unwrap_or(false);
                if left_arrow_close_handled {
                    crate::log_debug("YT comments loop handling LEFT as close");
                    crate::send_message_w_safe(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(YOUTUBE_COMMENTS_ID_CLOSE),
                        LPARAM(0),
                    );
                    continue;
                }
                if msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                    crate::log_debug("YT comments loop handling ESC as close");
                    crate::send_message_w_safe(
                        hwnd,
                        WM_COMMAND,
                        WPARAM(YOUTUBE_COMMENTS_ID_CLOSE),
                        LPARAM(0),
                    );
                    continue;
                }
                let (should_capture, handled) = with_youtube_comments_state(hwnd, |state| {
                    let search_edit_has_focus = focus == state.search_edit;
                    let search_controls_have_focus =
                        focus == state.search_edit || focus == state.search_button;
                    let keep_edit_navigation = search_edit_has_focus
                        && matches!(
                            msg.wParam.0 as u32,
                            k if k == VK_HOME.0 as u32
                                || k == VK_END.0 as u32
                                || k == VK_LEFT.0 as u32
                                || k == VK_RIGHT.0 as u32
                        );
                    let keep_search_controls_navigation = search_controls_have_focus
                        && matches!(
                            msg.wParam.0 as u32,
                            k if k == VK_UP.0 as u32 || k == VK_DOWN.0 as u32
                        );
                    let should_capture = focus != state.close_button
                        && !keep_edit_navigation
                        && !keep_search_controls_navigation
                        && is_youtube_comments_dialog_navigation_key(state, msg.wParam.0 as u32);
                    crate::log_debug(&format!(
                        "YT comments loop state before route view={:?} proxy={:?} close={:?} rows={} selected_row={} scroll_offset={} content_height={} should_capture={} keep_edit_navigation={} keep_search_controls_navigation={}",
                        state.view,
                        state.accessibility_proxy,
                        state.close_button,
                        state.rows.len(),
                        state.selected_row,
                        state.scroll_offset,
                        state.content_height,
                        should_capture,
                        keep_edit_navigation,
                        keep_search_controls_navigation
                    ));
                    if should_capture
                        && state.flat_list_mode
                        && msg.wParam.0 as u32 == VK_RETURN.0 as u32
                    {
                        if GetFocus() == state.search_edit {
                            return (false, false);
                        }
                        return (true, false);
                    }
                    let handled = should_capture
                        && crate::is_window_handle_valid(state.view)
                        && handle_youtube_comments_keydown_for_state(
                            state,
                            msg.wParam.0 as u32,
                            true,
                        );
                    (should_capture, handled)
                })
                .unwrap_or((false, false));
                crate::log_debug(&format!(
                    "YT comments loop WM_KEYDOWN should_capture={} handled={} key={}",
                    should_capture, handled, msg.wParam.0
                ));
                if should_capture
                    && msg.wParam.0 as u32 == VK_RETURN.0 as u32
                    && accept_youtube_comments_flat_selection(hwnd)
                {
                    continue;
                }
                let right_arrow_accept_handled = with_youtube_comments_state(hwnd, |state| {
                    let focus_is_flat_list =
                        focus == state.view || focus == state.accessibility_proxy;
                    should_capture
                        && msg.wParam.0 as u32 == VK_RIGHT.0 as u32
                        && state.flat_list_mode
                        && state.right_arrow_accepts_selection
                        && focus_is_flat_list
                })
                .unwrap_or(false);
                if right_arrow_accept_handled && accept_youtube_comments_flat_selection(hwnd) {
                    continue;
                }
                if should_capture {
                    continue;
                }
            }
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                crate::log_debug(&format!(
                    "YT comments IsDialogMessageW handled msg={} wparam={} hwnd={:?}",
                    msg.message, msg.wParam.0, msg.hwnd
                ));
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    unsafe {
        EnableWindow(parent, true);
    }
    if crate::is_window_handle_valid(parent) {
        crate::log_debug(&format!(
            "YT comments closing restore start parent={:?} focus_before_fg_restore={:?}",
            parent,
            crate::get_focus_safe()
        ));
        crate::set_foreground_window_safe(parent);
    }
    restore_stream_dialog_focus(parent);
    *action_result.lock().unwrap_or_else(|e| e.into_inner())
}

unsafe extern "system" fn youtube_comments_dialog_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "youtube_comments_dialog_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || youtube_comments_dialog_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn youtube_comments_dialog_wndproc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr =
                unsafe { (*create_struct).lpCreateParams as *mut YoutubeCommentsDialogInit };
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = crate::box_from_raw_safe(init_ptr);
            let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));
            let view_class_name = to_wide(YOUTUBE_COMMENTS_VIEW_CLASS_NAME);

            let title_label = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&init.title).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    10,
                    10,
                    780,
                    22,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };
            let view = unsafe {
                CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    PCWSTR(view_class_name.as_ptr()),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL,
                    10,
                    40,
                    780,
                    450,
                    hwnd,
                    HMENU(YOUTUBE_COMMENTS_ID_VIEW as isize),
                    HINSTANCE(0),
                    None,
                )
            };
            let accessibility_proxy = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_LISTBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WINDOW_STYLE((LBS_NOTIFY | LBS_HASSTRINGS) as u32),
                    0,
                    0,
                    1,
                    1,
                    hwnd,
                    HMENU(YOUTUBE_COMMENTS_ID_ACCESSIBILITY_PROXY as isize),
                    HINSTANCE(0),
                    None,
                )
            };
            let (search_label, search_edit, search_button) =
                if let Some(search) = init.flat_search.as_ref() {
                    let search_label = if search.show_edit {
                        unsafe {
                            CreateWindowExW(
                                Default::default(),
                                WC_STATIC,
                                PCWSTR(to_wide("Digita la stringa di ricerca").as_ptr()),
                                WS_CHILD | WS_VISIBLE,
                                10,
                                40,
                                300,
                                18,
                                hwnd,
                                HMENU(YOUTUBE_COMMENTS_ID_SEARCH_LABEL as isize),
                                HINSTANCE(0),
                                None,
                            )
                        }
                    } else {
                        HWND(0)
                    };
                    let search_edit = if search.show_edit {
                        unsafe {
                            CreateWindowExW(
                                WS_EX_CLIENTEDGE,
                                w!("EDIT"),
                                PCWSTR(to_wide(&search.initial_query).as_ptr()),
                                WS_CHILD
                                    | WS_VISIBLE
                                    | WS_TABSTOP
                                    | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                                10,
                                40,
                                680,
                                26,
                                hwnd,
                                HMENU(YOUTUBE_COMMENTS_ID_SEARCH_EDIT as isize),
                                HINSTANCE(0),
                                None,
                            )
                        }
                    } else {
                        HWND(0)
                    };
                    let search_button = unsafe {
                        CreateWindowExW(
                            Default::default(),
                            WC_BUTTON,
                            PCWSTR(to_wide(&search.button_label).as_ptr()),
                            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                            700,
                            40,
                            90,
                            26,
                            hwnd,
                            HMENU(YOUTUBE_COMMENTS_ID_SEARCH_BUTTON as isize),
                            HINSTANCE(0),
                            None,
                        )
                    };
                    if search.show_edit {
                        let search_text = to_wide(&search.initial_query);
                        crate::log_if_err!(crate::set_window_text_w_safe(
                            search_edit,
                            PCWSTR(search_text.as_ptr())
                        ));
                    }
                    (search_label, search_edit, search_button)
                } else {
                    (HWND(0), HWND(0), HWND(0))
                };
            let close_button = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(
                        to_wide(
                            init.flat_close_button_label
                                .as_deref()
                                .unwrap_or(&i18n::tr(init.language, "youtube.ok")),
                        )
                        .as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    700,
                    500,
                    90,
                    28,
                    hwnd,
                    HMENU(YOUTUBE_COMMENTS_ID_CLOSE as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            if hfont.0 != 0 {
                unsafe {
                    for control in [
                        title_label,
                        search_label,
                        view,
                        accessibility_proxy,
                        close_button,
                        search_edit,
                        search_button,
                    ] {
                        if control.0 == 0 {
                            continue;
                        }
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                }
            }

            let (root_indices, children_by_parent) = build_youtube_comment_threads(&init.comments);
            let expanded_comment_ids = HashSet::new();
            let flat_selection = init.flat_selection;
            crate::log_debug(&format!(
                "YT comments WM_CREATE title='{}' comments={} roots={} parents_with_children={}",
                init.title,
                init.comments.len(),
                root_indices.len(),
                children_by_parent.len()
            ));
            let mut state = Box::new(YoutubeCommentsDialogState {
                language: init.language,
                title_label,
                search_label,
                view,
                accessibility_proxy,
                close_button,
                search_edit,
                search_button,
                font: hfont,
                comments: init.comments,
                root_indices,
                children_by_parent,
                expanded_comment_ids,
                show_load_all_action: init.show_load_all_action,
                action_result: Arc::clone(&init.action_result),
                rows: Vec::new(),
                selected_row: 0,
                scroll_offset: 0,
                content_height: 0,
                flat_list_mode: flat_selection.is_some(),
                flat_selection_result: flat_selection
                    .as_ref()
                    .map(|selection| Arc::clone(&selection.result)),
                flat_search_result: init
                    .flat_search
                    .as_ref()
                    .map(|search| Arc::clone(&search.result)),
                flat_context_action: init.flat_context_action,
                right_arrow_accepts_selection: init.right_arrow_accepts_selection,
                left_arrow_closes: init.left_arrow_closes,
            });
            if let Some(selection) = flat_selection.as_ref()
                && let Some(index) = state.comments.iter().position(|comment| {
                    Some(comment.id.as_str()) == selection.initial_selected_id.as_deref()
                })
            {
                state.selected_row = index;
            }
            crate::set_window_long_ptr_w_safe(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            relayout_youtube_comments_dialog(hwnd);
            unsafe {
                SetFocus(accessibility_proxy);
            }
            crate::log_debug(&format!(
                "YT comments WM_CREATE completed hwnd={:?} view={:?} proxy={:?} close_button={:?}",
                hwnd, view, accessibility_proxy, close_button
            ));
            LRESULT(0)
        }
        WM_SIZE => {
            crate::log_debug(&format!(
                "YT comments WM_SIZE hwnd={:?} width={} height={}",
                hwnd,
                (lparam.0 as u32 & 0xffff) as i16 as i32,
                ((lparam.0 as u32 >> 16) & 0xffff) as i16 as i32
            ));
            relayout_youtube_comments_dialog(hwnd);
            LRESULT(0)
        }
        WM_SETFOCUS => {
            with_youtube_comments_state(hwnd, |state| {
                let target = if crate::is_window_handle_valid(state.accessibility_proxy) {
                    state.accessibility_proxy
                } else if crate::is_window_handle_valid(state.search_edit) {
                    state.search_edit
                } else if crate::is_window_handle_valid(state.view) {
                    state.view
                } else {
                    HWND(0)
                };
                if target.0 != 0 {
                    unsafe {
                        SetFocus(target);
                    }
                }
            });
            LRESULT(0)
        }
        WM_KEYDOWN => {
            crate::log_debug(&format!(
                "YT comments dialog WM_KEYDOWN key={} focus={:?}",
                wparam.0,
                unsafe { GetFocus() }
            ));
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                crate::log_debug("YT comments dialog handling ESC as close");
                crate::send_message_w_safe(
                    hwnd,
                    WM_COMMAND,
                    WPARAM(YOUTUBE_COMMENTS_ID_CLOSE),
                    LPARAM(0),
                );
                return LRESULT(0);
            }
            let handled = with_youtube_comments_state(hwnd, |state| {
                let focus = unsafe { GetFocus() };
                let keep_search_controls_navigation =
                    (focus == state.search_edit || focus == state.search_button)
                        && matches!(
                            wparam.0 as u32,
                            k if k == VK_UP.0 as u32 || k == VK_DOWN.0 as u32
                        );
                if focus == state.close_button
                    || keep_search_controls_navigation
                    || !is_youtube_comments_dialog_navigation_key(state, wparam.0 as u32)
                {
                    crate::log_debug(&format!(
                        "YT comments dialog WM_KEYDOWN ignored key={} focus={:?} close_button={:?} keep_search_controls_navigation={}",
                        wparam.0, focus, state.close_button, keep_search_controls_navigation
                    ));
                    return false;
                }
                if state.flat_list_mode && wparam.0 as u32 == VK_RETURN.0 as u32 {
                    if focus == state.search_edit || focus == state.search_button {
                        return false;
                    }
                    return accept_youtube_comments_flat_selection(hwnd);
                }
                if state.flat_list_mode
                    && state.right_arrow_accepts_selection
                    && wparam.0 as u32 == VK_RIGHT.0 as u32
                {
                    let focus_is_flat_list =
                        focus == state.view || focus == state.accessibility_proxy;
                    if !focus_is_flat_list {
                        return false;
                    }
                    return accept_youtube_comments_flat_selection(hwnd);
                }
                handle_youtube_comments_keydown_for_state(state, wparam.0 as u32, true)
            })
            .unwrap_or(false);
            crate::log_debug(&format!(
                "YT comments dialog WM_KEYDOWN handled={} key={}",
                handled, wparam.0
            ));
            if handled
                || with_youtube_comments_state(hwnd, |state| {
                    is_youtube_comments_dialog_navigation_key(state, wparam.0 as u32)
                })
                .unwrap_or(false)
            {
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            let notify = ((wparam.0 >> 16) & 0xffff) as u16;
            if cmd_id == YOUTUBE_COMMENTS_ID_ACCESSIBILITY_PROXY && notify == LBN_SELCHANGE as u16 {
                with_youtube_comments_state(hwnd, |state| {
                    let sel = crate::send_message_w_safe(
                        state.accessibility_proxy,
                        LB_GETCURSEL,
                        WPARAM(0),
                        LPARAM(0),
                    )
                    .0 as i32;
                    crate::log_debug(&format!(
                        "YT comments proxy LBN_SELCHANGE sel={} previous_selected={} rows={}",
                        sel,
                        state.selected_row,
                        state.rows.len()
                    ));
                    if sel >= 0 {
                        let selected_row = sel as usize;
                        if selected_row < state.rows.len() && selected_row != state.selected_row {
                            state.selected_row = selected_row;
                            let mut rect = windows::Win32::Foundation::RECT::default();
                            if crate::get_client_rect_safe(state.view, &mut rect).is_ok() {
                                let viewport_height = rect.bottom - rect.top;
                                ensure_selected_youtube_comment_visible(state, viewport_height);
                                update_youtube_comments_scrollbar(state, viewport_height);
                            }
                            if !unsafe { InvalidateRect(state.view, None, BOOL(1)) }.as_bool() {
                                crate::log_debug(
                                    "InvalidateRect failed after proxy selection change",
                                );
                            }
                        }
                    }
                });
                return LRESULT(0);
            }
            if (cmd_id == YOUTUBE_COMMENTS_ID_SEARCH_BUTTON
                || (cmd_id == YOUTUBE_COMMENTS_ID_SEARCH_EDIT && notify == 0))
                && trigger_youtube_comments_search(hwnd)
            {
                return LRESULT(0);
            }
            if cmd_id == YOUTUBE_COMMENTS_ID_CLOSE || cmd_id == 1 {
                if lparam.0 != 0 && accept_youtube_comments_flat_selection(hwnd) {
                    return LRESULT(0);
                }
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_CONTEXTMENU => {
            if show_youtube_comments_context_menu(hwnd, wparam, lparam) {
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            crate::log_if_err!(crate::destroy_window_safe(hwnd));
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA)
                as *mut YoutubeCommentsDialogState;
            if !ptr.is_null() {
                let _unused_box = crate::box_from_raw_safe(ptr);
            }
            LRESULT(0)
        }
        _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
    }
}

fn with_youtube_comments_state<R>(
    hwnd: HWND,
    f: impl FnOnce(&mut YoutubeCommentsDialogState) -> R,
) -> Option<R> {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut YoutubeCommentsDialogState;
    if ptr.is_null() {
        return None;
    }
    Some(f(unsafe { &mut *ptr }))
}

fn youtube_comments_row_comment<'a>(
    state: &'a YoutubeCommentsDialogState,
    row: &'a YoutubeCommentRow,
) -> Option<&'a YoutubeComment> {
    match &row.kind {
        YoutubeCommentRowKind::Comment { comment_index } => state.comments.get(*comment_index),
        YoutubeCommentRowKind::LoadAllComments => None,
    }
}

fn youtube_comments_load_all_label(language: Language) -> String {
    i18n::tr(language, "stream_audio.load_all_comments")
}

fn request_youtube_comments_load_all(state: &YoutubeCommentsDialogState) -> bool {
    if !state.show_load_all_action {
        return false;
    }
    let mut action = state
        .action_result
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    if *action == YoutubeCommentsDialogAction::LoadAll {
        return false;
    }
    *action = YoutubeCommentsDialogAction::LoadAll;
    true
}

fn accept_youtube_comments_flat_selection(hwnd: HWND) -> bool {
    let accepted = with_youtube_comments_state(hwnd, |state| {
        if !state.flat_list_mode {
            return false;
        }
        let Some(row) = state.rows.get(state.selected_row) else {
            return false;
        };
        let Some(comment) = youtube_comments_row_comment(state, row) else {
            return false;
        };
        let Some(result) = &state.flat_selection_result else {
            return false;
        };
        *result.lock().unwrap_or_else(|e| e.into_inner()) = Some(comment.id.clone());
        true
    })
    .unwrap_or(false);
    if accepted {
        crate::log_if_err!(crate::destroy_window_safe(hwnd));
    }
    accepted
}

fn trigger_youtube_comments_search(hwnd: HWND) -> bool {
    let query = with_youtube_comments_state(hwnd, |state| {
        let Some(result) = &state.flat_search_result else {
            return None;
        };
        if !crate::is_window_handle_valid(state.search_edit) {
            *result.lock().unwrap_or_else(|e| e.into_inner()) = Some(String::new());
            return Some(String::new());
        }
        let text = read_edit_text(state.search_edit).trim().to_string();
        if text.is_empty() {
            return None;
        }
        *result.lock().unwrap_or_else(|e| e.into_inner()) = Some(text.clone());
        Some(text)
    })
    .flatten();
    if query.is_some() {
        crate::log_if_err!(crate::destroy_window_safe(hwnd));
        return true;
    }
    false
}

fn build_youtube_comment_threads(
    comments: &[YoutubeComment],
) -> (Vec<usize>, HashMap<String, Vec<usize>>) {
    let known_ids: HashSet<&str> = comments.iter().map(|comment| comment.id.as_str()).collect();
    let mut root_indices = Vec::new();
    let mut children_by_parent: HashMap<String, Vec<usize>> = HashMap::new();
    for (index, comment) in comments.iter().enumerate() {
        let parent_id = comment.parent.trim();
        if parent_id.eq_ignore_ascii_case("root") || !known_ids.contains(parent_id) {
            root_indices.push(index);
        } else {
            children_by_parent
                .entry(parent_id.to_string())
                .or_default()
                .push(index);
        }
    }
    crate::log_debug(&format!(
        "YT comments build threads comments={} roots={} parent_buckets={}",
        comments.len(),
        root_indices.len(),
        children_by_parent.len()
    ));
    (root_indices, children_by_parent)
}

const YOUTUBE_COMMENTS_MARGIN: i32 = 10;
const YOUTUBE_COMMENTS_VIEW_TOP: i32 = 40;
const YOUTUBE_COMMENTS_SEARCH_LABEL_HEIGHT: i32 = 18;
const YOUTUBE_COMMENTS_SEARCH_HEIGHT: i32 = 26;
const YOUTUBE_COMMENTS_SEARCH_GAP: i32 = 8;
const YOUTUBE_COMMENTS_CLOSE_WIDTH: i32 = 90;
const YOUTUBE_COMMENTS_CLOSE_HEIGHT: i32 = 28;
const YOUTUBE_COMMENTS_CLOSE_BOTTOM_MARGIN: i32 = 12;
const YOUTUBE_COMMENTS_VIEW_BOTTOM_GAP: i32 = 10;
const YOUTUBE_COMMENTS_ROW_PADDING_X: i32 = 8;
const YOUTUBE_COMMENTS_ROW_PADDING_Y: i32 = 6;
const YOUTUBE_COMMENTS_ROW_GAP: i32 = 4;
const YOUTUBE_COMMENTS_INDENT_WIDTH: i32 = 22;
const YOUTUBE_COMMENTS_TOGGLE_WIDTH: i32 = 22;
const YOUTUBE_COMMENTS_SCROLL_LINE: i32 = 40;
const YOUTUBE_COMMENTS_WHEEL_DELTA: i32 = 120;

fn relayout_youtube_comments_dialog(hwnd: HWND) {
    let mut client = windows::Win32::Foundation::RECT::default();
    if crate::get_client_rect_safe(hwnd, &mut client).is_err() {
        crate::log_debug("Failed to query YouTube comments dialog client rect");
        return;
    }
    with_youtube_comments_state(hwnd, |state| unsafe {
        let width = (client.right - client.left).max(220);
        let height = (client.bottom - client.top).max(180);
        let close_x = width - YOUTUBE_COMMENTS_MARGIN - YOUTUBE_COMMENTS_CLOSE_WIDTH;
        let close_y = height - YOUTUBE_COMMENTS_CLOSE_BOTTOM_MARGIN - YOUTUBE_COMMENTS_CLOSE_HEIGHT;
        let search_row_height = if crate::is_window_handle_valid(state.search_edit) {
            YOUTUBE_COMMENTS_SEARCH_LABEL_HEIGHT
                + YOUTUBE_COMMENTS_SEARCH_GAP
                + YOUTUBE_COMMENTS_SEARCH_HEIGHT
                + YOUTUBE_COMMENTS_SEARCH_GAP
        } else {
            0
        };
        let view_top = YOUTUBE_COMMENTS_VIEW_TOP + search_row_height;
        let view_height = (close_y - view_top - YOUTUBE_COMMENTS_VIEW_BOTTOM_GAP).max(80);
        crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.title_label,
            YOUTUBE_COMMENTS_MARGIN,
            YOUTUBE_COMMENTS_MARGIN,
            width - (YOUTUBE_COMMENTS_MARGIN * 2),
            22,
            true
        ));
        crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.view,
            YOUTUBE_COMMENTS_MARGIN,
            view_top,
            width - (YOUTUBE_COMMENTS_MARGIN * 2),
            view_height,
            true
        ));
        if crate::is_window_handle_valid(state.search_edit) {
            let search_button_width = 90;
            let search_button_x = width - YOUTUBE_COMMENTS_MARGIN - search_button_width;
            let search_edit_width =
                (search_button_x - YOUTUBE_COMMENTS_MARGIN - YOUTUBE_COMMENTS_SEARCH_GAP).max(120);
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
                state.search_label,
                YOUTUBE_COMMENTS_MARGIN,
                YOUTUBE_COMMENTS_VIEW_TOP,
                width - (YOUTUBE_COMMENTS_MARGIN * 2),
                YOUTUBE_COMMENTS_SEARCH_LABEL_HEIGHT,
                true
            ));
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
                state.search_edit,
                YOUTUBE_COMMENTS_MARGIN,
                YOUTUBE_COMMENTS_VIEW_TOP
                    + YOUTUBE_COMMENTS_SEARCH_LABEL_HEIGHT
                    + YOUTUBE_COMMENTS_SEARCH_GAP,
                search_edit_width,
                YOUTUBE_COMMENTS_SEARCH_HEIGHT,
                true
            ));
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
                state.search_button,
                search_button_x,
                YOUTUBE_COMMENTS_VIEW_TOP
                    + YOUTUBE_COMMENTS_SEARCH_LABEL_HEIGHT
                    + YOUTUBE_COMMENTS_SEARCH_GAP,
                search_button_width,
                YOUTUBE_COMMENTS_SEARCH_HEIGHT,
                true
            ));
        } else if crate::is_window_handle_valid(state.search_button) {
            let search_button_x =
                close_x - YOUTUBE_COMMENTS_SEARCH_GAP - YOUTUBE_COMMENTS_CLOSE_WIDTH;
            crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
                state.search_button,
                search_button_x,
                close_y,
                YOUTUBE_COMMENTS_CLOSE_WIDTH,
                YOUTUBE_COMMENTS_CLOSE_HEIGHT,
                true
            ));
        }
        crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.close_button,
            close_x,
            close_y,
            YOUTUBE_COMMENTS_CLOSE_WIDTH,
            YOUTUBE_COMMENTS_CLOSE_HEIGHT,
            true
        ));
        crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.accessibility_proxy,
            close_x - 2,
            close_y + YOUTUBE_COMMENTS_CLOSE_HEIGHT - 2,
            1,
            1,
            true
        ));
    });
    rebuild_youtube_comment_rows(hwnd);
}

fn rebuild_youtube_comment_rows(hwnd: HWND) {
    with_youtube_comments_state(hwnd, |state| {
        let mut rect = windows::Win32::Foundation::RECT::default();
        if crate::get_client_rect_safe(state.view, &mut rect).is_err() {
            crate::log_debug("Failed to query comment custom view rect");
            return;
        }
        let view_width = (rect.right - rect.left).max(120);
        let viewport_height = (rect.bottom - rect.top).max(1);
        let hdc = unsafe { GetDC(state.view) };
        if hdc.0 == 0 {
            crate::log_debug("Failed to acquire DC for YouTube comments view");
            return;
        }
        let old_font = if state.font.0 != 0 {
            unsafe { SelectObject(hdc, HGDIOBJ(state.font.0)) }
        } else {
            HGDIOBJ(0)
        };
        let mut rows = Vec::new();
        for &comment_index in &state.root_indices {
            append_youtube_comment_rows(state, hdc, view_width, comment_index, 0, &mut rows);
        }
        if state.show_load_all_action {
            let height = measure_youtube_comment_row_height(
                hdc,
                view_width,
                0,
                &youtube_comments_load_all_label(state.language),
            );
            rows.push(YoutubeCommentRow {
                kind: YoutubeCommentRowKind::LoadAllComments,
                depth: 0,
                height,
                has_children: false,
                expanded: false,
            });
        }
        if old_font.0 != 0 {
            unsafe {
                let _prev = SelectObject(hdc, old_font);
            }
        }
        let _released = unsafe { ReleaseDC(state.view, hdc) };

        state.rows = rows;
        state.content_height = state
            .rows
            .iter()
            .fold(0, |acc, row| acc + row.height + YOUTUBE_COMMENTS_ROW_GAP)
            .saturating_sub(YOUTUBE_COMMENTS_ROW_GAP);
        if !state.rows.is_empty() {
            state.selected_row = state.selected_row.min(state.rows.len().saturating_sub(1));
        } else {
            state.selected_row = 0;
        }
        clamp_youtube_comment_scroll(state, viewport_height);
        ensure_selected_youtube_comment_visible(state, viewport_height);
        update_youtube_comments_scrollbar(state, viewport_height);
        crate::log_debug(&format!(
            "YT comments rebuild rows comments={} roots={} rows={} selected_row={} scroll_offset={} content_height={} viewport_height={} view_width={}",
            state.comments.len(),
            state.root_indices.len(),
            state.rows.len(),
            state.selected_row,
            state.scroll_offset,
            state.content_height,
            viewport_height,
            view_width
        ));
        rebuild_youtube_comments_accessibility_proxy(state);
        if !unsafe { InvalidateRect(state.view, None, BOOL(1)) }.as_bool() {
            crate::log_debug("InvalidateRect failed for YouTube comments custom view");
        }
    });
}

fn append_youtube_comment_rows(
    state: &YoutubeCommentsDialogState,
    hdc: windows::Win32::Graphics::Gdi::HDC,
    view_width: i32,
    comment_index: usize,
    depth: usize,
    rows: &mut Vec<YoutubeCommentRow>,
) {
    let Some(comment) = state.comments.get(comment_index) else {
        return;
    };
    let has_children = state
        .children_by_parent
        .get(comment.id.as_str())
        .map(|items| !items.is_empty())
        .unwrap_or(false);
    let effective_depth = if state.flat_list_mode { 0 } else { depth };
    let expanded = has_children && state.expanded_comment_ids.contains(comment.id.as_str());
    let height = measure_youtube_comment_row_height(
        hdc,
        view_width,
        effective_depth,
        &format_youtube_comment_tree_label(state, comment),
    );
    rows.push(YoutubeCommentRow {
        kind: YoutubeCommentRowKind::Comment { comment_index },
        depth: effective_depth,
        height,
        has_children,
        expanded,
    });
    crate::log_debug(&format!(
        "YT comments append row comment_index={} id={} depth={} has_children={} expanded={} current_rows={}",
        comment_index,
        comment.id,
        depth,
        has_children,
        expanded,
        rows.len()
    ));
    if expanded && let Some(children) = state.children_by_parent.get(comment.id.as_str()) {
        for &child_index in children {
            append_youtube_comment_rows(state, hdc, view_width, child_index, depth + 1, rows);
        }
    }
}

fn measure_youtube_comment_row_height(
    hdc: windows::Win32::Graphics::Gdi::HDC,
    view_width: i32,
    depth: usize,
    text: &str,
) -> i32 {
    let indent = (depth as i32 * YOUTUBE_COMMENTS_INDENT_WIDTH) + YOUTUBE_COMMENTS_TOGGLE_WIDTH;
    let available_width = (view_width
        - (YOUTUBE_COMMENTS_ROW_PADDING_X * 2)
        - indent
        - YOUTUBE_COMMENTS_ROW_PADDING_X)
        .max(80);
    let mut rect = windows::Win32::Foundation::RECT {
        left: 0,
        top: 0,
        right: available_width,
        bottom: 0,
    };
    let mut wide = to_wide(text);
    unsafe {
        let _drawn = DrawTextW(
            hdc,
            wide.as_mut_slice(),
            &mut rect,
            DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX,
        );
    }
    (rect.bottom - rect.top + (YOUTUBE_COMMENTS_ROW_PADDING_Y * 2)).max(24)
}

fn youtube_comment_reply_position(
    state: &YoutubeCommentsDialogState,
    comment: &YoutubeComment,
    depth: usize,
) -> Option<(usize, usize)> {
    if depth == 0 {
        return None;
    }
    let siblings = state.children_by_parent.get(comment.parent.as_str())?;
    let total = siblings.len();
    let index = siblings.iter().position(|&comment_index| {
        state
            .comments
            .get(comment_index)
            .map(|item| item.id.as_str())
            == Some(comment.id.as_str())
    })?;
    Some((index + 1, total))
}

fn format_youtube_comment_tree_label(
    state: &YoutubeCommentsDialogState,
    comment: &YoutubeComment,
) -> String {
    if state.flat_list_mode {
        let title = normalize_comment_text(&comment.author);
        let description = normalize_comment_text(&comment.text);
        if title.is_empty() {
            return description;
        }
        if description.is_empty() {
            return title;
        }
        return format!("{title}\r\n{description}");
    }
    let language = state.language;
    let mut parts = vec![comment.author.clone()];
    if !comment.time_text.is_empty() {
        parts.push(localize_comment_time_text(language, &comment.time_text));
    }
    let mut label = parts.join(" - ");
    let text = normalize_comment_text(&comment.text);
    if !text.is_empty() {
        label.push_str(" - ");
        label.push_str(&text);
    }
    label
}

fn format_youtube_comment_accessibility_label(
    state: &YoutubeCommentsDialogState,
    comment: &YoutubeComment,
    depth: usize,
) -> String {
    if state.flat_list_mode {
        let title = normalize_comment_text(&comment.author);
        let description = normalize_comment_text(&comment.text);
        if title.is_empty() {
            return description;
        }
        if description.is_empty() {
            return title;
        }
        return format!("{title} - {description}");
    }
    let language = state.language;
    let mut parts = Vec::new();
    if let Some((index, total)) = youtube_comment_reply_position(state, comment, depth) {
        parts.push(i18n::tr_f(
            language,
            "stream_audio.comment_reply_position",
            &[("index", &index.to_string()), ("total", &total.to_string())],
        ));
    }
    parts.push(comment.author.clone());
    if !comment.time_text.is_empty() {
        parts.push(localize_comment_time_text(language, &comment.time_text));
    }
    let mut label = parts.join(" - ");
    let text = normalize_comment_text(&comment.text);
    if !text.is_empty() {
        label.push_str(" - ");
        label.push_str(&text);
    }
    label
}

fn normalize_comment_text(text: &str) -> String {
    text.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn is_youtube_comments_dialog_navigation_key(state: &YoutubeCommentsDialogState, key: u32) -> bool {
    if state.flat_list_mode {
        return matches!(
            key,
            k if k == VK_UP.0 as u32
                || k == VK_DOWN.0 as u32
                || k == VK_HOME.0 as u32
                || k == VK_END.0 as u32
                || k == VK_PRIOR.0 as u32
                || k == VK_NEXT.0 as u32
                || (state.right_arrow_accepts_selection && k == VK_RIGHT.0 as u32)
                || k == VK_RETURN.0 as u32
        );
    }
    matches!(
        key,
        k if k == VK_UP.0 as u32
            || k == VK_DOWN.0 as u32
            || k == VK_LEFT.0 as u32
            || k == VK_RIGHT.0 as u32
            || k == VK_HOME.0 as u32
            || k == VK_END.0 as u32
            || k == VK_PRIOR.0 as u32
            || k == VK_NEXT.0 as u32
            || k == VK_RETURN.0 as u32
    )
}

fn youtube_comment_row_accessibility_label(
    state: &YoutubeCommentsDialogState,
    row: &YoutubeCommentRow,
) -> Option<String> {
    if matches!(&row.kind, YoutubeCommentRowKind::LoadAllComments) {
        return Some(youtube_comments_load_all_label(state.language));
    }
    let comment = youtube_comments_row_comment(state, row)?;
    let mut label = String::new();
    if row.depth > 0 {
        label.push_str(&"  ".repeat(row.depth));
    }
    label.push_str(&format_youtube_comment_accessibility_label(
        state, comment, row.depth,
    ));
    if row.has_children {
        let reply_count = state
            .children_by_parent
            .get(comment.id.as_str())
            .map(|children| children.len())
            .unwrap_or(0);
        let reply_count_text = if reply_count == 1 {
            tr_or(
                state.language,
                "stream_audio.comment_reply_count_one",
                "1 reply",
            )
        } else {
            i18n::tr_f(
                state.language,
                "stream_audio.comment_reply_count_many",
                &[("count", &reply_count.to_string())],
            )
        };
        let state_text = if row.expanded {
            tr_or(
                state.language,
                "stream_audio.comment_expandable_expanded",
                "expanded",
            )
        } else {
            tr_or(
                state.language,
                "stream_audio.comment_expandable_collapsed",
                "not expanded",
            )
        };
        label.push_str(". ");
        label.push_str(&reply_count_text);
        label.push_str(". ");
        label.push_str(&state_text);
    }
    Some(label)
}

fn rebuild_youtube_comments_accessibility_proxy(state: &YoutubeCommentsDialogState) {
    crate::log_debug(&format!(
        "YT comments rebuild proxy rows={} selected_row={} proxy={:?}",
        state.rows.len(),
        state.selected_row,
        state.accessibility_proxy
    ));
    crate::send_message_w_safe(
        state.accessibility_proxy,
        LB_RESETCONTENT,
        WPARAM(0),
        LPARAM(0),
    );
    for row in &state.rows {
        let Some(label) = youtube_comment_row_accessibility_label(state, row) else {
            continue;
        };
        crate::send_message_w_safe(
            state.accessibility_proxy,
            LB_ADDSTRING,
            WPARAM(0),
            LPARAM(to_wide(&label).as_ptr() as isize),
        );
    }
    if !state.rows.is_empty() {
        crate::send_message_w_safe(
            state.accessibility_proxy,
            LB_SETCURSEL,
            WPARAM(state.selected_row),
            LPARAM(0),
        );
    }
}

fn sync_youtube_comments_accessibility_proxy_selection(
    state: &YoutubeCommentsDialogState,
    focus_proxy: bool,
) {
    crate::log_debug(&format!(
        "YT comments sync proxy selection selected_row={} focus_proxy={} proxy={:?}",
        state.selected_row, focus_proxy, state.accessibility_proxy
    ));
    crate::send_message_w_safe(
        state.accessibility_proxy,
        LB_SETCURSEL,
        WPARAM(state.selected_row),
        LPARAM(0),
    );
    if focus_proxy {
        unsafe {
            SetFocus(state.accessibility_proxy);
        }
    }
}

fn selected_youtube_comment_id(state: &YoutubeCommentsDialogState) -> Option<String> {
    let row = state.rows.get(state.selected_row)?;
    let comment = youtube_comments_row_comment(state, row)?;
    Some(comment.id.clone())
}

fn trigger_youtube_comments_context_action(hwnd: HWND) -> bool {
    with_youtube_comments_state(hwnd, |state| {
        let Some(action) = state.flat_context_action.as_ref() else {
            return false;
        };
        let Some(selected_id) = selected_youtube_comment_id(state) else {
            return false;
        };
        if !(action.enabled)(&selected_id) {
            return false;
        }
        (action.handler)(selected_id);
        true
    })
    .unwrap_or(false)
}

fn show_youtube_comments_context_menu(hwnd: HWND, wparam: WPARAM, lparam: LPARAM) -> bool {
    with_youtube_comments_state(hwnd, |state| unsafe {
        let Some(action) = state.flat_context_action.as_ref() else {
            return false;
        };
        let target = HWND(wparam.0 as isize);
        if target.0 != 0
            && target != hwnd
            && target != state.view
            && target != state.accessibility_proxy
        {
            return false;
        }
        let Some(selected_id) = selected_youtube_comment_id(state) else {
            return false;
        };
        if !(action.enabled)(&selected_id) {
            return false;
        }

        let menu = match CreatePopupMenu() {
            Ok(menu) => menu,
            Err(err) => {
                crate::log_debug(&format!("Failed to create multiline context menu: {}", err));
                return false;
            }
        };
        let label = to_wide(&action.label);
        if let Err(err) = AppendMenuW(menu, MF_STRING, 1, PCWSTR(label.as_ptr())) {
            crate::log_debug(&format!(
                "Failed to append multiline context menu item: {}",
                err
            ));
            crate::log_if_err!(DestroyMenu(menu));
            return false;
        }
        let point = if lparam.0 == -1 {
            let mut pt = POINT::default();
            if let Err(err) = GetCursorPos(&mut pt) {
                crate::log_debug(&format!(
                    "Failed to query cursor position for multiline context menu: {}",
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
        if command.0 == 1 {
            (action.handler)(selected_id);
        }
        true
    })
    .unwrap_or(false)
}

unsafe extern "system" fn youtube_comments_view_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "youtube_comments_view_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || youtube_comments_view_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn youtube_comments_view_wndproc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_GETDLGCODE => {
            crate::log_debug(&format!("YT comments view WM_GETDLGCODE hwnd={:?}", hwnd));
            LRESULT(DLGC_WANTARROWS as isize)
        }
        WM_SETFOCUS => {
            with_youtube_comments_view_state(hwnd, |state| {
                if state.rows.is_empty() {
                    crate::log_debug("YT comments view WM_SETFOCUS with no rows");
                    return;
                }
                state.selected_row = state.selected_row.min(state.rows.len().saturating_sub(1));
                crate::log_debug(&format!(
                    "YT comments view WM_SETFOCUS selected_row={} rows={} flat_list_mode={}",
                    state.selected_row,
                    state.rows.len(),
                    state.flat_list_mode
                ));
                sync_youtube_comments_accessibility_proxy_selection(state, false);
                announce_selected_youtube_comment(state);
            });
            LRESULT(0)
        }
        WM_PAINT => {
            paint_youtube_comments_view(hwnd);
            LRESULT(0)
        }
        WM_VSCROLL => {
            crate::log_debug(&format!(
                "YT comments view WM_VSCROLL wparam={} hwnd={:?}",
                wparam.0, hwnd
            ));
            if handle_youtube_comments_vscroll(hwnd, wparam) {
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_MOUSEWHEEL => {
            crate::log_debug(&format!(
                "YT comments view WM_MOUSEWHEEL wparam={} hwnd={:?}",
                wparam.0, hwnd
            ));
            if handle_youtube_comments_mouse_wheel(hwnd, wparam) {
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        WM_LBUTTONDOWN => {
            crate::log_debug(&format!(
                "YT comments view WM_LBUTTONDOWN lparam={} hwnd={:?}",
                lparam.0, hwnd
            ));
            activate_youtube_comment_row_from_click(hwnd, lparam);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            crate::log_debug(&format!(
                "YT comments view WM_KEYDOWN key={} hwnd={:?}",
                wparam.0, hwnd
            ));
            if handle_youtube_comments_keydown(hwnd, wparam.0 as u32) {
                return LRESULT(0);
            }
            crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
        }
        _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
    }
}

fn with_youtube_comments_view_state<R>(
    hwnd: HWND,
    f: impl FnOnce(&mut YoutubeCommentsDialogState) -> R,
) -> Option<R> {
    let parent = unsafe { GetParent(hwnd) };
    if parent.0 == 0 {
        return None;
    }
    with_youtube_comments_state(parent, f)
}

fn paint_youtube_comments_view(hwnd: HWND) {
    let mut ps = PAINTSTRUCT::default();
    let hdc = unsafe { BeginPaint(hwnd, &mut ps) };
    if hdc.0 == 0 {
        return;
    }
    with_youtube_comments_view_state(hwnd, |state| {
        let mut client = windows::Win32::Foundation::RECT::default();
        if crate::get_client_rect_safe(hwnd, &mut client).is_err() {
            return;
        }
        unsafe {
            let _filled = FillRect(hdc, &client, HBRUSH((COLOR_WINDOW.0 + 1) as isize));
        }
        let old_font = if state.font.0 != 0 {
            unsafe { SelectObject(hdc, HGDIOBJ(state.font.0)) }
        } else {
            HGDIOBJ(0)
        };
        unsafe {
            let _mode = SetBkMode(hdc, TRANSPARENT);
        }
        let mut y = -state.scroll_offset;
        for (row_index, row) in state.rows.iter().enumerate() {
            let top = y;
            let bottom = y + row.height;
            if bottom >= 0 && top <= client.bottom {
                let row_rect = windows::Win32::Foundation::RECT {
                    left: 0,
                    top,
                    right: client.right,
                    bottom,
                };
                if row_index == state.selected_row {
                    unsafe {
                        let _filled =
                            FillRect(hdc, &row_rect, HBRUSH((COLOR_HIGHLIGHT.0 + 1) as isize));
                    }
                }
                let indent_left = YOUTUBE_COMMENTS_ROW_PADDING_X
                    + (row.depth as i32 * YOUTUBE_COMMENTS_INDENT_WIDTH);
                if row.has_children {
                    let marker = if row.expanded { "[-]" } else { "[+]" };
                    let mut marker_rect = windows::Win32::Foundation::RECT {
                        left: indent_left,
                        top: top + YOUTUBE_COMMENTS_ROW_PADDING_Y,
                        right: indent_left + YOUTUBE_COMMENTS_TOGGLE_WIDTH,
                        bottom: bottom - YOUTUBE_COMMENTS_ROW_PADDING_Y,
                    };
                    let mut marker_w = to_wide(marker);
                    unsafe {
                        let _drawn =
                            DrawTextW(hdc, marker_w.as_mut_slice(), &mut marker_rect, DT_NOPREFIX);
                    }
                }
                let mut text_rect = windows::Win32::Foundation::RECT {
                    left: indent_left + YOUTUBE_COMMENTS_TOGGLE_WIDTH,
                    top: top + YOUTUBE_COMMENTS_ROW_PADDING_Y,
                    right: client.right - YOUTUBE_COMMENTS_ROW_PADDING_X,
                    bottom: bottom - YOUTUBE_COMMENTS_ROW_PADDING_Y,
                };
                let label = match &row.kind {
                    YoutubeCommentRowKind::Comment { .. } => {
                        youtube_comments_row_comment(state, row)
                            .map(|comment| format_youtube_comment_tree_label(state, comment))
                    }
                    YoutubeCommentRowKind::LoadAllComments => {
                        Some(youtube_comments_load_all_label(state.language))
                    }
                };
                if let Some(label) = label {
                    let mut wide = to_wide(&label);
                    unsafe {
                        let _drawn = DrawTextW(
                            hdc,
                            wide.as_mut_slice(),
                            &mut text_rect,
                            DT_WORDBREAK | DT_NOPREFIX,
                        );
                    }
                }
            }
            y = bottom + YOUTUBE_COMMENTS_ROW_GAP;
        }
        if old_font.0 != 0 {
            unsafe {
                let _prev = SelectObject(hdc, old_font);
            }
        }
    });
    unsafe {
        let _ended = EndPaint(hwnd, &ps);
    }
}

fn clamp_youtube_comment_scroll(state: &mut YoutubeCommentsDialogState, viewport_height: i32) {
    let max_offset = (state.content_height - viewport_height.max(1)).max(0);
    state.scroll_offset = state.scroll_offset.clamp(0, max_offset);
}

fn update_youtube_comments_scrollbar(state: &YoutubeCommentsDialogState, viewport_height: i32) {
    let max_offset = (state.content_height - viewport_height.max(1)).max(0);
    let info = SCROLLINFO {
        cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
        fMask: SIF_RANGE | SIF_PAGE | SIF_POS,
        nMin: 0,
        nMax: max_offset,
        nPage: viewport_height.max(1) as u32,
        nPos: state.scroll_offset,
        ..Default::default()
    };
    if unsafe { SetScrollInfo(state.view, SB_VERT, &info, true) } == 0 {
        crate::log_debug("SetScrollInfo failed for YouTube comments custom view");
    }
    if unsafe { ShowScrollBar(state.view, SB_VERT, max_offset > 0) }.is_err() {
        crate::log_debug("ShowScrollBar failed for YouTube comments custom view");
    }
}

fn youtube_comment_row_bounds(
    state: &YoutubeCommentsDialogState,
    row_index: usize,
) -> Option<(i32, i32)> {
    let mut y = 0i32;
    for (index, row) in state.rows.iter().enumerate() {
        let bottom = y + row.height;
        if index == row_index {
            return Some((y, bottom));
        }
        y = bottom + YOUTUBE_COMMENTS_ROW_GAP;
    }
    None
}

fn ensure_selected_youtube_comment_visible(
    state: &mut YoutubeCommentsDialogState,
    viewport_height: i32,
) {
    let Some((top, bottom)) = youtube_comment_row_bounds(state, state.selected_row) else {
        return;
    };
    if top < state.scroll_offset {
        state.scroll_offset = top;
    } else if bottom > state.scroll_offset + viewport_height {
        state.scroll_offset = (bottom - viewport_height).max(0);
    }
    clamp_youtube_comment_scroll(state, viewport_height);
}

fn handle_youtube_comments_mouse_wheel(hwnd: HWND, wparam: WPARAM) -> bool {
    let delta = (((wparam.0 >> 16) & 0xffff) as i16) as i32;
    if delta == 0 {
        return false;
    }
    let changed = with_youtube_comments_view_state(hwnd, |state| {
        let mut rect = windows::Win32::Foundation::RECT::default();
        if crate::get_client_rect_safe(hwnd, &mut rect).is_err() {
            return false;
        }
        let viewport_height = rect.bottom - rect.top;
        let steps = (delta.abs() / YOUTUBE_COMMENTS_WHEEL_DELTA).max(1);
        let delta_pixels = steps * YOUTUBE_COMMENTS_SCROLL_LINE;
        let old = state.scroll_offset;
        state.scroll_offset = if delta > 0 {
            old - delta_pixels
        } else {
            old + delta_pixels
        };
        clamp_youtube_comment_scroll(state, viewport_height);
        if state.scroll_offset != old {
            update_youtube_comments_scrollbar(state, viewport_height);
            true
        } else {
            false
        }
    })
    .unwrap_or(false);
    crate::log_debug(&format!(
        "YT comments mouse wheel delta={} changed={}",
        delta, changed
    ));
    if changed && !unsafe { InvalidateRect(hwnd, None, BOOL(1)) }.as_bool() {
        crate::log_debug("InvalidateRect failed after mouse wheel in comment custom view");
    }
    changed
}

fn handle_youtube_comments_vscroll(hwnd: HWND, wparam: WPARAM) -> bool {
    let command = (wparam.0 & 0xffff) as u32;
    let thumb_pos = ((wparam.0 >> 16) & 0xffff) as i32;
    let changed = with_youtube_comments_view_state(hwnd, |state| {
        let mut rect = windows::Win32::Foundation::RECT::default();
        if crate::get_client_rect_safe(hwnd, &mut rect).is_err() {
            return false;
        }
        let viewport_height = rect.bottom - rect.top;
        let max_offset = (state.content_height - viewport_height.max(1)).max(0);
        let old = state.scroll_offset;
        let page_step =
            (viewport_height - YOUTUBE_COMMENTS_SCROLL_LINE).max(YOUTUBE_COMMENTS_SCROLL_LINE);
        let mut new = old;
        match command {
            c if c == SB_LINEUP.0 as u32 => new -= YOUTUBE_COMMENTS_SCROLL_LINE,
            c if c == SB_LINEDOWN.0 as u32 => new += YOUTUBE_COMMENTS_SCROLL_LINE,
            c if c == SB_PAGEUP.0 as u32 => new -= page_step,
            c if c == SB_PAGEDOWN.0 as u32 => new += page_step,
            c if c == SB_TOP.0 as u32 => new = 0,
            c if c == SB_BOTTOM.0 as u32 => new = max_offset,
            c if c == SB_THUMBPOSITION.0 as u32 || c == SB_THUMBTRACK.0 as u32 => {
                let mut info = SCROLLINFO {
                    cbSize: std::mem::size_of::<SCROLLINFO>() as u32,
                    fMask: SIF_TRACKPOS,
                    ..Default::default()
                };
                if unsafe { GetScrollInfo(hwnd, SB_VERT, &mut info) }.is_ok() {
                    new = info.nTrackPos;
                } else {
                    new = thumb_pos;
                }
            }
            _ => {}
        }
        state.scroll_offset = new.clamp(0, max_offset);
        if state.scroll_offset != old {
            update_youtube_comments_scrollbar(state, viewport_height);
            true
        } else {
            false
        }
    })
    .unwrap_or(false);
    crate::log_debug(&format!(
        "YT comments vscroll command={} thumb_pos={} changed={}",
        command, thumb_pos, changed
    ));
    if changed && !unsafe { InvalidateRect(hwnd, None, BOOL(1)) }.as_bool() {
        crate::log_debug("InvalidateRect failed after vscroll in comment custom view");
    }
    changed
}

fn activate_youtube_comment_row_from_click(hwnd: HWND, lparam: LPARAM) {
    let click_x = (lparam.0 as u32 & 0xffff) as i16 as i32;
    let click_y = ((lparam.0 as u32 >> 16) & 0xffff) as i16 as i32;
    crate::log_debug(&format!(
        "YT comments click x={} y={} hwnd={:?}",
        click_x, click_y, hwnd
    ));
    let mut invalidate = false;
    with_youtube_comments_view_state(hwnd, |state| {
        if let Some(row_index) = hit_test_youtube_comment_row(state, click_y) {
            crate::log_debug(&format!(
                "YT comments click hit row_index={} previous_selected={}",
                row_index, state.selected_row
            ));
            state.selected_row = row_index;
            sync_youtube_comments_accessibility_proxy_selection(state, true);
            if let Some(row) = state.rows.get(row_index) {
                let toggle_left = YOUTUBE_COMMENTS_ROW_PADDING_X
                    + (row.depth as i32 * YOUTUBE_COMMENTS_INDENT_WIDTH);
                let toggle_right = toggle_left + YOUTUBE_COMMENTS_TOGGLE_WIDTH;
                if matches!(&row.kind, YoutubeCommentRowKind::LoadAllComments) {
                    if request_youtube_comments_load_all(state) {
                        let parent_hwnd =
                            unsafe { windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd) };
                        if parent_hwnd.0 != 0 {
                            crate::log_if_err!(crate::destroy_window_safe(parent_hwnd));
                        }
                        return;
                    }
                } else if row.has_children && click_x >= toggle_left && click_x <= toggle_right {
                    toggle_youtube_comment_row_expansion(state, row_index);
                } else {
                    let mut rect = windows::Win32::Foundation::RECT::default();
                    if crate::get_client_rect_safe(hwnd, &mut rect).is_ok() {
                        ensure_selected_youtube_comment_visible(state, rect.bottom - rect.top);
                        update_youtube_comments_scrollbar(state, rect.bottom - rect.top);
                    }
                }
                invalidate = true;
            }
        }
    });
    if invalidate && !unsafe { InvalidateRect(hwnd, None, BOOL(1)) }.as_bool() {
        crate::log_debug("InvalidateRect failed after comment row click");
    }
}

fn hit_test_youtube_comment_row(state: &YoutubeCommentsDialogState, click_y: i32) -> Option<usize> {
    let mut y = -state.scroll_offset;
    for (index, row) in state.rows.iter().enumerate() {
        let bottom = y + row.height;
        if click_y >= y && click_y < bottom {
            return Some(index);
        }
        y = bottom + YOUTUBE_COMMENTS_ROW_GAP;
    }
    None
}

fn announce_selected_youtube_comment(state: &YoutubeCommentsDialogState) {
    let Some(announcement) = selected_youtube_comment_announcement(state) else {
        return;
    };
    crate::log_debug(&format!(
        "YT comments announce selected via proxy selected_row={} text='{}'",
        state.selected_row, announcement
    ));
    let wide = to_wide(&announcement);
    crate::log_if_err!(crate::set_window_text_w_safe(
        state.accessibility_proxy,
        PCWSTR(wide.as_ptr())
    ));
}

fn announce_selected_youtube_comment_expand_state(
    state: &YoutubeCommentsDialogState,
    expanded: bool,
) {
    let key = if expanded {
        "stream_audio.comment_expandable_expanded"
    } else {
        "stream_audio.comment_expandable_collapsed"
    };
    let fallback = if expanded { "expanded" } else { "not expanded" };
    screen_reader_speak(&tr_or(state.language, key, fallback));
}

fn selected_youtube_comment_announcement(state: &YoutubeCommentsDialogState) -> Option<String> {
    let row = state.rows.get(state.selected_row)?;
    if matches!(&row.kind, YoutubeCommentRowKind::LoadAllComments) {
        return Some(youtube_comments_load_all_label(state.language));
    }
    let comment = youtube_comments_row_comment(state, row)?;
    let mut announcement = format_youtube_comment_accessibility_label(state, comment, row.depth);
    if row.has_children {
        let reply_count = state
            .children_by_parent
            .get(comment.id.as_str())
            .map(|children| children.len())
            .unwrap_or(0);
        let reply_count_text = if reply_count == 1 {
            tr_or(
                state.language,
                "stream_audio.comment_reply_count_one",
                "1 reply",
            )
        } else {
            i18n::tr_f(
                state.language,
                "stream_audio.comment_reply_count_many",
                &[("count", &reply_count.to_string())],
            )
        };
        let state_text = if row.expanded {
            tr_or(
                state.language,
                "stream_audio.comment_expandable_expanded",
                "expanded",
            )
        } else {
            tr_or(
                state.language,
                "stream_audio.comment_expandable_collapsed",
                "not expanded",
            )
        };
        announcement.push_str(". ");
        announcement.push_str(&reply_count_text);
        announcement.push_str(". ");
        announcement.push_str(&state_text);
    }
    Some(announcement)
}

fn handle_youtube_comments_keydown(hwnd: HWND, key: u32) -> bool {
    crate::log_debug(&format!(
        "YT comments handle keydown entry hwnd={:?} key={}",
        hwnd, key
    ));
    let handled = with_youtube_comments_view_state(hwnd, |state| {
        let mut rect = windows::Win32::Foundation::RECT::default();
        if crate::get_client_rect_safe(hwnd, &mut rect).is_err() {
            crate::log_debug("YT comments handle keydown: get_client_rect failed");
            return false;
        }
        handle_youtube_comments_keydown_inner(state, key, rect.bottom - rect.top)
    })
    .unwrap_or(false);
    let load_all_requested = with_youtube_comments_view_state(hwnd, |state| {
        *state
            .action_result
            .lock()
            .unwrap_or_else(|e| e.into_inner())
            == YoutubeCommentsDialogAction::LoadAll
    })
    .unwrap_or(false);
    if handled && load_all_requested {
        let parent_hwnd = unsafe { windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd) };
        if parent_hwnd.0 != 0 {
            crate::log_if_err!(crate::destroy_window_safe(parent_hwnd));
        }
        return true;
    }
    if handled && !unsafe { InvalidateRect(hwnd, None, BOOL(1)) }.as_bool() {
        crate::log_debug("InvalidateRect failed after comment keydown");
    }
    handled
}

fn handle_youtube_comments_keydown_for_state(
    state: &mut YoutubeCommentsDialogState,
    key: u32,
    focus_proxy: bool,
) -> bool {
    crate::log_debug(&format!(
        "YT comments keydown for state key={} focus_proxy={} view={:?} proxy={:?} close_button={:?} rows={} selected_row={} scroll_offset={}",
        key,
        focus_proxy,
        state.view,
        state.accessibility_proxy,
        state.close_button,
        state.rows.len(),
        state.selected_row,
        state.scroll_offset
    ));
    if !crate::is_window_handle_valid(state.view)
        || !crate::is_window_handle_valid(state.accessibility_proxy)
    {
        crate::log_debug("YT comments keydown for state aborted: invalid view or proxy handle");
        return false;
    }
    let mut rect = windows::Win32::Foundation::RECT::default();
    if crate::get_client_rect_safe(state.view, &mut rect).is_err() {
        crate::log_debug("YT comments keydown for state aborted: failed get_client_rect");
        return false;
    }
    if focus_proxy {
        unsafe {
            SetFocus(state.accessibility_proxy);
        }
        crate::log_debug(&format!(
            "YT comments keydown for state forced focus to proxy={:?}",
            state.accessibility_proxy
        ));
    }
    let handled = handle_youtube_comments_keydown_inner(state, key, rect.bottom - rect.top);
    crate::log_debug(&format!(
        "YT comments keydown for state result handled={} key={} selected_row={} scroll_offset={}",
        handled, key, state.selected_row, state.scroll_offset
    ));
    if handled
        && *state
            .action_result
            .lock()
            .unwrap_or_else(|e| e.into_inner())
            == YoutubeCommentsDialogAction::LoadAll
    {
        let parent_hwnd = unsafe { windows::Win32::UI::WindowsAndMessaging::GetParent(state.view) };
        if parent_hwnd.0 != 0 {
            crate::log_if_err!(crate::destroy_window_safe(parent_hwnd));
        }
        return true;
    }
    if handled && !unsafe { InvalidateRect(state.view, None, BOOL(1)) }.as_bool() {
        crate::log_debug("InvalidateRect failed after dialog-routed comment keydown");
    }
    handled
}

fn handle_youtube_comments_keydown_inner(
    state: &mut YoutubeCommentsDialogState,
    key: u32,
    viewport_height: i32,
) -> bool {
    if state.rows.is_empty() {
        crate::log_debug("YT comments keydown inner aborted: rows is empty");
        return false;
    }
    let previous_selected_row = state.selected_row;
    let previous_scroll_offset = state.scroll_offset;
    let selected_row_before = state.rows.get(state.selected_row).cloned();
    match key {
        k if k == VK_UP.0 as u32 => {
            if state.selected_row > 0 {
                state.selected_row -= 1;
            }
        }
        k if k == VK_DOWN.0 as u32 => {
            if state.selected_row + 1 < state.rows.len() {
                state.selected_row += 1;
            }
        }
        k if k == VK_HOME.0 as u32 => {
            state.selected_row = 0;
        }
        k if k == VK_END.0 as u32 => {
            state.selected_row = state.rows.len().saturating_sub(1);
        }
        k if k == VK_PRIOR.0 as u32 => {
            move_youtube_comment_selection_by_page(state, -1);
        }
        k if k == VK_NEXT.0 as u32 => {
            move_youtube_comment_selection_by_page(state, 1);
        }
        k if k == VK_LEFT.0 as u32 => {
            if !collapse_or_select_parent_youtube_comment(state) {
                return false;
            }
        }
        k if k == VK_RIGHT.0 as u32 => {
            if !expand_youtube_comment(state) {
                return false;
            }
        }
        k if k == VK_RETURN.0 as u32 => {
            if state
                .rows
                .get(state.selected_row)
                .map(|row| matches!(&row.kind, YoutubeCommentRowKind::LoadAllComments))
                .unwrap_or(false)
            {
                return request_youtube_comments_load_all(state);
            }
            if !toggle_youtube_comment_expansion(state) {
                return false;
            }
        }
        _ => return false,
    }
    ensure_selected_youtube_comment_visible(state, viewport_height);
    update_youtube_comments_scrollbar(state, viewport_height);
    if state.selected_row != previous_selected_row {
        sync_youtube_comments_accessibility_proxy_selection(state, false);
        announce_selected_youtube_comment(state);
    } else if let Some(row_before) = selected_row_before {
        if key == VK_LEFT.0 as u32 && row_before.has_children && row_before.expanded {
            announce_selected_youtube_comment_expand_state(state, false);
        } else if (key == VK_RIGHT.0 as u32 || key == VK_RETURN.0 as u32)
            && row_before.has_children
            && !row_before.expanded
        {
            announce_selected_youtube_comment_expand_state(state, true);
        }
    }
    crate::log_debug(&format!(
        "YT comments keydown inner key={} rows={} selected_row {}->{} scroll_offset {}->{} viewport_height={}",
        key,
        state.rows.len(),
        previous_selected_row,
        state.selected_row,
        previous_scroll_offset,
        state.scroll_offset,
        viewport_height
    ));
    true
}

fn move_youtube_comment_selection_by_page(state: &mut YoutubeCommentsDialogState, direction: i32) {
    let step = 8usize;
    if direction < 0 {
        state.selected_row = state.selected_row.saturating_sub(step);
    } else if !state.rows.is_empty() {
        state.selected_row = (state.selected_row + step).min(state.rows.len().saturating_sub(1));
    }
}

fn collapse_or_select_parent_youtube_comment(state: &mut YoutubeCommentsDialogState) -> bool {
    let Some(row) = state.rows.get(state.selected_row).cloned() else {
        return false;
    };
    let YoutubeCommentRowKind::Comment { comment_index } = row.kind else {
        return false;
    };
    if row.has_children && row.expanded {
        toggle_youtube_comment_row_expansion(state, state.selected_row);
        return true;
    }
    if row.depth == 0 {
        return false;
    }
    let Some(comment) = state.comments.get(comment_index) else {
        return false;
    };
    let parent_id = comment.parent.as_str();
    if let Some(parent_row_index) = state.rows.iter().position(|candidate| {
        youtube_comments_row_comment(state, candidate)
            .map(|candidate_comment| candidate_comment.id.as_str() == parent_id)
            .unwrap_or(false)
    }) {
        if let Some(parent_row) = state.rows.get(parent_row_index)
            && parent_row.has_children
            && parent_row.expanded
        {
            toggle_youtube_comment_row_expansion(state, parent_row_index);
            state.selected_row = parent_row_index.min(state.rows.len().saturating_sub(1));
            return true;
        }
        state.selected_row = parent_row_index;
        return true;
    }
    false
}

fn expand_youtube_comment(state: &mut YoutubeCommentsDialogState) -> bool {
    let Some(row) = state.rows.get(state.selected_row).cloned() else {
        return false;
    };
    if !matches!(row.kind, YoutubeCommentRowKind::Comment { .. }) {
        return false;
    }
    if !row.has_children || row.expanded {
        return false;
    }
    toggle_youtube_comment_row_expansion(state, state.selected_row);
    true
}

fn toggle_youtube_comment_expansion(state: &mut YoutubeCommentsDialogState) -> bool {
    let Some(row) = state.rows.get(state.selected_row).cloned() else {
        return false;
    };
    if !matches!(row.kind, YoutubeCommentRowKind::Comment { .. }) {
        return false;
    }
    if !row.has_children {
        return false;
    }
    toggle_youtube_comment_row_expansion(state, state.selected_row);
    true
}

fn toggle_youtube_comment_row_expansion(state: &mut YoutubeCommentsDialogState, row_index: usize) {
    let Some(row) = state.rows.get(row_index).cloned() else {
        return;
    };
    let YoutubeCommentRowKind::Comment { comment_index } = row.kind else {
        return;
    };
    let Some(comment) = state.comments.get(comment_index) else {
        return;
    };
    crate::log_debug(&format!(
        "YT comments toggle expansion row_index={} comment_id={} expanded_before={} rows={} selected_row={}",
        row_index,
        comment.id,
        row.expanded,
        state.rows.len(),
        state.selected_row
    ));
    if row.expanded {
        state.expanded_comment_ids.remove(comment.id.as_str());
    } else {
        state.expanded_comment_ids.insert(comment.id.clone());
    }
    let parent_hwnd = unsafe { windows::Win32::UI::WindowsAndMessaging::GetParent(state.view) };
    if parent_hwnd.0 != 0 {
        rebuild_youtube_comment_rows(parent_hwnd);
    }
}

fn localize_comment_time_text(language: Language, raw: &str) -> String {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return String::new();
    }
    let raw_edited_suffix = " (edited)";
    let localized_edited_suffix = format!(
        " {}",
        tr_or(language, "stream_audio.comment_edited_suffix", "(edited)")
    );
    let (base_text, edited) = if let Some(base) = trimmed.strip_suffix(raw_edited_suffix) {
        (base.trim_end(), true)
    } else {
        (trimmed, false)
    };
    let normalized = base_text.to_ascii_lowercase();
    if normalized == "just now" {
        let mut localized = i18n::tr(language, "stream_audio.comment_time_just_now");
        if edited {
            localized.push_str(&localized_edited_suffix);
        }
        return localized;
    }

    let parts: Vec<&str> = normalized.split_whitespace().collect();
    if parts.len() < 3 || parts.last() != Some(&"ago") {
        return trimmed.to_string();
    }

    let count = match parts[0] {
        "a" | "an" => 1u64,
        value => match value.parse::<u64>() {
            Ok(value) => value,
            Err(_) => return trimmed.to_string(),
        },
    };
    let count_string = count.to_string();
    let args = [("count", count_string.as_str())];

    let localized = match parts[1] {
        "minute" | "minutes" => {
            if count == 1 {
                i18n::tr(language, "stream_audio.comment_time_minute_ago")
            } else {
                i18n::tr_f(language, "stream_audio.comment_time_minutes_ago", &args)
            }
        }
        "hour" | "hours" => {
            if count == 1 {
                i18n::tr(language, "stream_audio.comment_time_hour_ago")
            } else {
                i18n::tr_f(language, "stream_audio.comment_time_hours_ago", &args)
            }
        }
        "day" | "days" => {
            if count == 1 {
                i18n::tr(language, "stream_audio.comment_time_day_ago")
            } else {
                i18n::tr_f(language, "stream_audio.comment_time_days_ago", &args)
            }
        }
        "week" | "weeks" => {
            if count == 1 {
                i18n::tr(language, "stream_audio.comment_time_week_ago")
            } else {
                i18n::tr_f(language, "stream_audio.comment_time_weeks_ago", &args)
            }
        }
        "month" | "months" => {
            if count == 1 {
                i18n::tr(language, "stream_audio.comment_time_month_ago")
            } else {
                i18n::tr_f(language, "stream_audio.comment_time_months_ago", &args)
            }
        }
        "year" | "years" => {
            if count == 1 {
                i18n::tr(language, "stream_audio.comment_time_year_ago")
            } else {
                i18n::tr_f(language, "stream_audio.comment_time_years_ago", &args)
            }
        }
        _ => return trimmed.to_string(),
    };

    if edited {
        format!("{localized}{localized_edited_suffix}")
    } else {
        localized
    }
}

fn choose_youtube_collection_entry(
    parent: HWND,
    language: Language,
    ytdlp_path: &Path,
    url: &str,
    initial_page: Option<usize>,
    initial_selected_label: Option<&str>,
    initial_progress: Option<HWND>,
) -> Result<Option<ResolvedStreamSelection>, String> {
    if !is_youtube_collection_url(url) {
        return Ok(Some(ResolvedStreamSelection {
            url: url.to_string(),
            collection_url: None,
            collection_page: None,
            selected_label: None,
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

        let Some(selected) = choose_stream_collection_entry_page(
            parent,
            language,
            ytdlp_path,
            &entries,
            page > 0,
            has_more,
            initial_selected_label.filter(|_| page == initial_page.unwrap_or(0)),
        ) else {
            restore_stream_parent_after_selection_cancel(parent);
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
                selected_label: Some(selected),
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
    initial_selected_label: Option<&str>,
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

        let Some(selected) = choose_stream_collection_entry_page(
            parent,
            language,
            ytdlp_path,
            &entries,
            page > 0,
            has_more,
            initial_selected_label.filter(|_| page == 0),
        ) else {
            restore_stream_parent_after_selection_cancel(parent);
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
                    None,
                );
            }
            return Ok(Some(ResolvedStreamSelection {
                url: selected_url,
                collection_url: None,
                collection_page: None,
                selected_label: Some(selected),
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
    initial_selected_label: Option<&str>,
    initial_progress: Option<HWND>,
) -> Result<Option<ResolvedStreamSelection>, String> {
    if looks_like_valid_stream_url(input) {
        return choose_youtube_collection_entry(
            parent,
            language,
            ytdlp_path,
            input,
            initial_collection_page,
            initial_selected_label,
            initial_progress,
        );
    }
    choose_youtube_search_entry(
        parent,
        language,
        ytdlp_path,
        input,
        initial_selected_label,
        initial_progress,
    )
}

fn is_members_only_stream_error(err: &str) -> bool {
    let err_lc = err.to_ascii_lowercase();
    err_lc.contains("members-only")
        || err_lc.contains("members only")
        || err_lc.contains("join this channel to get access to members-only content")
}

fn is_drm_not_supported_stream_error(err: &str) -> bool {
    let err_lc = err.to_ascii_lowercase();
    err_lc.contains("known to use drm protection")
        || err_lc.contains("uses drm protection")
        || err_lc.contains("not be supported")
        || err_lc.contains("[drm]")
}

fn members_only_stream_message(language: Language) -> String {
    format!(
        "{} {}",
        i18n::tr(language, "stream_audio.members_only_video"),
        i18n::tr(language, "stream_audio.choose_another_video")
    )
}

fn drm_not_supported_stream_message(language: Language) -> String {
    i18n::tr(language, "stream_audio.drm_not_supported")
}

fn is_login_required_stream_error(err: &str) -> bool {
    if FORCE_YTDLP_AUTH_PROMPT_FOR_TESTING {
        return true;
    }
    let err_lc = err.to_ascii_lowercase();
    err_lc.contains("use --username and --password") || err_lc.contains("login required")
}

fn stream_auth_site_key(url: &str) -> Option<String> {
    Url::parse(url)
        .ok()
        .and_then(|parsed| parsed.host_str().map(|host| host.to_ascii_lowercase()))
}

fn load_saved_stream_site_credentials(
    parent: HWND,
    site_key: &str,
) -> Option<YtdlpAuthCredentials> {
    with_state(parent, |state| {
        get_ytdlp_site_credentials(&state.settings, site_key)
            .map(|(username, password)| YtdlpAuthCredentials { username, password })
    })
    .flatten()
}

fn save_stream_site_credentials(parent: HWND, site_key: &str, credentials: &YtdlpAuthCredentials) {
    let Some(settings_snapshot) = with_state(parent, |state| {
        if !set_ytdlp_site_credentials(
            &mut state.settings,
            site_key,
            &credentials.username,
            &credentials.password,
        ) {
            return None;
        }
        Some(state.settings.clone())
    })
    .flatten() else {
        return;
    };
    save_settings(settings_snapshot);
}

fn clear_stream_site_credentials(parent: HWND, site_key: &str) {
    let Some(settings_snapshot) = with_state(parent, |state| {
        if !clear_ytdlp_site_credentials(&mut state.settings, site_key) {
            return None;
        }
        Some(state.settings.clone())
    })
    .flatten() else {
        return;
    };
    save_settings(settings_snapshot);
}

struct PromptedYtdlpCredentials {
    credentials: YtdlpAuthCredentials,
    save_credentials: bool,
}

fn prompt_ytdlp_credentials(
    parent: HWND,
    language: Language,
    username_default: &str,
    save_credentials_default: bool,
    saved_credentials_failed: bool,
) -> Option<PromptedYtdlpCredentials> {
    let title = tr_or(
        language,
        "stream_audio.auth_required_title",
        "Authentication required",
    );
    let (body_key, body_fallback) = if saved_credentials_failed {
        (
            "stream_audio.saved_credentials_failed",
            "Saved credentials did not work. Enter the credentials for this site again:",
        )
    } else {
        (
            "stream_audio.auth_prompt",
            "Enter the credentials required by this site:",
        )
    };
    let body = tr_or(language, body_key, body_fallback);
    let result = prompt_window::prompt_credentials(
        parent,
        &title,
        &body,
        username_default,
        save_credentials_default,
        language,
    )?;
    let username = result.username.trim().to_string();
    if username.is_empty() {
        return None;
    }
    let password = result.password.trim().to_string();
    if password.is_empty() {
        return None;
    }
    Some(PromptedYtdlpCredentials {
        credentials: YtdlpAuthCredentials { username, password },
        save_credentials: result.save_credentials,
    })
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
    reopen_selected_label: Option<String>,
}

#[derive(Clone)]
struct StreamAudioTrack {
    format_id: String,
    label: String,
}

#[derive(Clone)]
struct YtdlpAuthCredentials {
    username: String,
    password: String,
}

struct YtdlpDownloadAttempt {
    primary_path: Option<PathBuf>,
    stderr_capture: String,
    stalled: bool,
}

struct YtdlpDownloadRequest<'a> {
    parent: HWND,
    progress: HWND,
    language: Language,
    ytdlp_path: &'a Path,
    url: &'a str,
    cache_dir: &'a Path,
    prefix: &'a str,
    output_template: &'a Path,
    dialog_data: &'a StreamDialogResult,
    selected_audio_format: Option<&'a str>,
    ytdlp_debug: bool,
    credentials: Option<&'a YtdlpAuthCredentials>,
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
    open_comments_button: HWND,
    ok_button: HWND,
    result: Arc<Mutex<Option<StreamDialogResult>>>,
}

struct StreamTrackDialogInit {
    parent: HWND,
    language: Language,
    label: String,
    options: Vec<String>,
    default_selection: usize,
    result: Arc<Mutex<Option<usize>>>,
}

struct StreamTrackDialogState {
    parent: HWND,
    language: Language,
    combo: HWND,
    ok_button: HWND,
    options: Vec<String>,
    result: Arc<Mutex<Option<usize>>>,
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

fn current_stream_dialog_input(state: &StreamDialogState) -> String {
    let url = read_edit_text(state.url_edit);
    if !url.trim().is_empty() {
        return url;
    }
    let favorite_idx =
        crate::send_message_w_safe(state.favorites_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if favorite_idx >= 0
        && let Some(favorite) = state.favorites.get(favorite_idx as usize)
    {
        return favorite.url.clone();
    }
    String::new()
}

fn update_stream_open_comments_button(state: &StreamDialogState) {
    let input = current_stream_dialog_input(state);
    let is_enabled = extract_video_id(&input).is_some() && !is_youtube_collection_url(&input);
    crate::enable_window_safe(state.open_comments_button, is_enabled);
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
                let open_comments_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(
                        to_wide(&i18n::tr(init.language, "stream_audio.open_comments")).as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    110,
                    178,
                    130,
                    28,
                    hwnd,
                    HMENU(STREAM_ID_OPEN_COMMENTS as isize),
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
                    open_comments_button,
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
                    open_comments_button,
                    ok_button,
                    result: init.result.clone(),
                });
                refill_stream_quality_combo(&state, false);
                refill_stream_favorites_combo(&state, Some(0));
                update_stream_open_comments_button(&state);
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                SetFocus(url_edit);
                LRESULT(0)
            }
            WM_SETFOCUS => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut StreamDialogState;
                if ptr.is_null() {
                    return LRESULT(0);
                }
                let state = &mut *ptr;
                SetFocus(state.url_edit);
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
                if cmd_id == STREAM_ID_URL && notify_code == EN_CHANGE as usize {
                    if with_stream_dialog_state(hwnd, |state| {
                        update_stream_open_comments_button(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to update stream open comments button");
                    }
                    return LRESULT(0);
                }
                if cmd_id == STREAM_ID_FAVORITES && notify_code == CBN_SELCHANGE as usize {
                    if with_stream_dialog_state(hwnd, |state| {
                        update_stream_open_comments_button(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to update stream open comments button");
                    }
                    return LRESULT(0);
                }
                if cmd_id == STREAM_ID_OPEN_COMMENTS {
                    let Some((language, input)) = with_stream_dialog_state(hwnd, |state| {
                        (state.language, current_stream_dialog_input(state))
                    }) else {
                        crate::log_debug("Failed to access stream dialog state");
                        return LRESULT(0);
                    };
                    let Some(normalized_url) = normalize_youtube_input_for_download(&input) else {
                        return LRESULT(0);
                    };
                    if extract_video_id(&normalized_url).is_none()
                        || is_youtube_collection_url(&normalized_url)
                    {
                        return LRESULT(0);
                    }

                    let labels_data = labels(language);
                    let progress = open_progress_dialog(
                        hwnd,
                        language,
                        "stream_audio.progress_title",
                        "stream_audio.comments_loading",
                        false,
                    );
                    let ytdlp_path = match ensure_ytdlp_available(
                        hwnd,
                        language,
                        &labels_data,
                        Some(progress),
                    ) {
                        Ok(Some(path)) => path,
                        Ok(None) => return LRESULT(0),
                        Err(err) => {
                            close_progress_dialog(progress);
                            let message = i18n::tr_f(
                                language,
                                "stream_audio.download_failed",
                                &[("err", &err)],
                            );
                            show_error(hwnd, language, &message);
                            return LRESULT(0);
                        }
                    };
                    close_progress_dialog(progress);
                    let entry = StreamCollectionEntry {
                        label: normalized_url.clone(),
                        url: normalized_url,
                    };
                    show_youtube_comments_for_stream_entry(hwnd, language, &ytdlp_path, &entry);
                    if with_stream_dialog_state(hwnd, |state| {
                        update_stream_open_comments_button(state);
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to refresh stream open comments button");
                    }
                    return LRESULT(0);
                }
                if cmd_id == STREAM_ID_OK || cmd_id == 1 {
                    if with_stream_dialog_state(hwnd, |state| {
                        let msg = i18n::tr(state.language, "stream_audio.progress_downloading");
                        if !screen_reader_speak(&msg) {
                            crate::log_debug("Screen reader speak failed");
                        }
                        let url = current_stream_dialog_input(state);
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
                                reopen_selected_label: None,
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
            reopen_selected_label: context.selected_label,
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
    selected_label: Option<&str>,
) {
    let return_input = collection_url
        .map(ToOwned::to_owned)
        .unwrap_or_else(|| original_input.to_string());
    crate::log_debug(&format!(
        "stream transition [update_return_context_from_selection]: input={} collection_page={:?} selected_label={}",
        return_input,
        collection_page,
        selected_label.unwrap_or("")
    ));
    if crate::with_state(parent, |state| {
        state.active_youtube_return_context.input = Some(return_input);
        state.active_youtube_return_context.collection_page = collection_page;
        state.active_youtube_return_context.selected_label = selected_label.map(ToOwned::to_owned);
    })
    .is_none()
    {
        crate::log_debug("Failed to update YouTube return context from selection");
    }
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
        crate::app_windows::podcast_save_window::disable_fake_progress(dialog);
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

fn restore_stream_parent_after_selection_cancel(hwnd: HWND) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        EnableWindow(hwnd, true);
    }
    restore_stream_parent_focus(hwnd);
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

fn restore_stream_dialog_focus(parent: HWND) {
    crate::log_debug(&format!(
        "stream focus restore after comments start parent={:?} focus_before={:?}",
        parent,
        crate::get_focus_safe()
    ));
    if crate::app_windows::interpreter_select_window::restore_interpreter_select_focus(parent) {
        crate::log_debug(&format!(
            "stream focus restore after comments: restored interpreter selection parent={:?}",
            parent
        ));
        log_stream_focus_snapshot("youtube_comments.closed.after_restore", parent);
        return;
    }
    let restored = with_stream_dialog_state(parent, |state| {
        let target = if crate::is_window_handle_valid(state.url_edit) {
            state.url_edit
        } else if crate::is_window_handle_valid(state.favorites_combo) {
            state.favorites_combo
        } else if crate::is_window_handle_valid(state.ok_button) {
            state.ok_button
        } else {
            HWND(0)
        };
        crate::log_debug(&format!(
            "stream focus restore after comments fallback: parent={:?} target={:?} url_edit={:?} favorites_combo={:?} ok_button={:?} focus_before_target={:?}",
            parent,
            target,
            state.url_edit,
            state.favorites_combo,
            state.ok_button,
            crate::get_focus_safe()
        ));
        if target.0 != 0 {
            unsafe {
                SetFocus(target);
            }
            crate::log_debug(&format!(
                "stream focus restore after comments fallback target applied: parent={:?} target={:?} focus_after={:?}",
                parent,
                target,
                crate::get_focus_safe()
            ));
            true
        } else {
            false
        }
    })
    .unwrap_or(false);
    if !restored {
        crate::log_debug(&format!(
            "stream focus restore after comments: no stream dialog state for parent={:?}",
            parent
        ));
        post_focus_editor(parent);
        crate::schedule_italiaonline_close_focus_debug_snapshots(parent);
        crate::log_debug(&format!(
            "stream focus restore after comments: posted WM_FOCUS_EDITOR parent={:?} foreground_after_post={:?} focus_after_post={:?}",
            parent,
            crate::get_foreground_window_safe(),
            crate::get_focus_safe()
        ));
    }
    log_stream_focus_snapshot("youtube_comments.closed.after_restore", parent);
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
    if let Some(pct) = parse_ytdlp_progress_template_pct(line) {
        return Some(pct);
    }
    let lower = line.to_ascii_lowercase();
    let has_download_progress_prefix = lower.contains("[download]");
    let has_named_progress_fields = lower.contains("downloaded_bytes")
        || lower.contains("total_bytes")
        || lower.contains("total_bytes_estimate");
    if !has_download_progress_prefix && !has_named_progress_fields {
        return None;
    }
    if let Some(pct) = parse_ytdlp_progress_bytes_pct(line) {
        return Some(pct);
    }
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

fn parse_ytdlp_progress_template_pct(line: &str) -> Option<u32> {
    let marker = format!("{YTDLP_PROGRESS_TEMPLATE_PREFIX}:");
    let start = line.find(&marker)? + marker.len();
    let mut parts = line[start..].split(':');
    let downloaded = parts
        .next()?
        .trim()
        .parse::<u64>()
        .ok()
        .filter(|v| *v > 0)?;
    let total = parts
        .next()
        .and_then(|part| part.trim().parse::<u64>().ok())
        .filter(|v| *v > 0)
        .or_else(|| {
            parts
                .next()
                .and_then(|part| part.trim().parse::<u64>().ok())
                .filter(|v| *v > 0)
        })?;
    Some(
        ((downloaded as f64 / total as f64) * 100.0)
            .clamp(0.0, 100.0)
            .round() as u32,
    )
}

fn parse_ytdlp_progress_bytes_pct(line: &str) -> Option<u32> {
    let downloaded = parse_ytdlp_named_u64(line, "downloaded_bytes")
        .or_else(|| parse_ytdlp_named_u64(line, "downloaded"))
        .filter(|value| *value > 0)?;
    let total = parse_ytdlp_named_u64(line, "total_bytes")
        .or_else(|| parse_ytdlp_named_u64(line, "total_bytes_estimate"))
        .or_else(|| parse_ytdlp_named_u64(line, "total"))
        .filter(|value| *value > 0)?;
    Some(
        ((downloaded as f64 / total as f64) * 100.0)
            .clamp(0.0, 100.0)
            .round() as u32,
    )
}

fn ytdlp_download_progress_template() -> String {
    format!(
        "download:{prefix}:%(progress.downloaded_bytes)s:%(progress.total_bytes)s:%(progress.total_bytes_estimate)s",
        prefix = YTDLP_PROGRESS_TEMPLATE_PREFIX
    )
}

fn configure_ytdlp_stream_download_command(
    cmd: &mut Command,
    url: &str,
    output_template: &Path,
    dialog_data: &StreamDialogResult,
    selected_audio_format: Option<&str>,
    credentials: Option<&YtdlpAuthCredentials>,
) {
    cmd.arg("--no-playlist")
        .arg("--socket-timeout")
        .arg(YTDLP_SOCKET_TIMEOUT_SECS)
        .arg("--no-warnings")
        .arg("--newline")
        .arg("--progress-template")
        .arg(ytdlp_download_progress_template());

    if let Some(credentials) = credentials {
        crate::log_debug(&format!(
            "yt-dlp auth args enabled username={} password_arg=true",
            credentials.username
        ));
        cmd.arg("--username")
            .arg(&credentials.username)
            .arg("--password")
            .arg(&credentials.password);
    }

    if let Some(format_id) = selected_audio_format {
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
                cmd.arg("-f").arg("bestaudio/best");
            }
            StreamOutputFormat::Mp3
                if matches!(dialog_data.quality, StreamQualitySelection::Original) =>
            {
                cmd.arg("-f").arg("bestaudio[ext=mp3]/bestaudio/best");
            }
            StreamOutputFormat::M4a => {
                cmd.arg("-f").arg("m4a/bestaudio[ext=m4a]/bestaudio/best");
            }
            _ => {
                cmd.arg("-f").arg("bestaudio/best");
            }
        }
    }

    cmd.arg("-o")
        .arg(output_template.to_string_lossy().to_string())
        .arg("--")
        .arg(url)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
}

fn handle_ytdlp_progress_line(
    line: String,
    ytdlp_debug: bool,
    output_label: &str,
    stderr_capture: &Arc<Mutex<String>>,
    progress: &Arc<AtomicU32>,
    activity: &Arc<AtomicU32>,
) {
    if let Ok(mut captured) = stderr_capture.lock()
        && captured.len() < 16_384
    {
        if !captured.is_empty() {
            captured.push('\n');
        }
        captured.push_str(&line);
    }
    if ytdlp_debug {
        crate::log_debug(&format!("yt-dlp {}: {}", output_label, line));
    }
    activity.fetch_add(1, Ordering::Relaxed);
    if let Some(pct) = parse_ytdlp_progress_pct(&line) {
        progress.fetch_max(pct, Ordering::Relaxed);
    }
}

fn spawn_ytdlp_output_reader<R: Read + Send + 'static>(
    pipe: R,
    ytdlp_debug: bool,
    output_label: &'static str,
    stderr_capture: Arc<Mutex<String>>,
    progress: Arc<AtomicU32>,
    activity: Arc<AtomicU32>,
) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        for line_result in BufReader::new(pipe).lines() {
            let line = match line_result {
                Ok(line) => line,
                Err(err) => {
                    crate::log_debug(&format!("yt-dlp {} read failed: {}", output_label, err));
                    continue;
                }
            };
            handle_ytdlp_progress_line(
                line,
                ytdlp_debug,
                output_label,
                &stderr_capture,
                &progress,
                &activity,
            );
        }
    })
}

fn run_ytdlp_stream_download_attempt(
    req: YtdlpDownloadRequest<'_>,
) -> Result<YtdlpDownloadAttempt, String> {
    let mut cmd = ytdlp_command(req.ytdlp_path);
    configure_ytdlp_stream_download_command(
        &mut cmd,
        req.url,
        req.output_template,
        req.dialog_data,
        req.selected_audio_format,
        req.credentials,
    );
    if req.ytdlp_debug {
        let format_for_log = req.selected_audio_format.map_or_else(
            || match req.dialog_data.format {
                StreamOutputFormat::Auto => i18n::tr(req.language, "stream_audio.format.auto"),
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
            "yt-dlp stream start: url={} output_template={} format={} auth={}",
            req.url,
            req.output_template.to_string_lossy(),
            format_for_log,
            req.credentials.is_some()
        ));
    }
    let mut child = cmd.spawn().map_err(|err| {
        i18n::tr_f(
            req.language,
            "stream_audio.download_failed",
            &[("err", &err.to_string())],
        )
    })?;
    keep_stream_progress_focus(req.progress);

    let stderr_pipe = child.stderr.take().ok_or_else(|| {
        i18n::tr_f(
            req.language,
            "stream_audio.download_failed",
            &[("err", "yt-dlp stderr unavailable")],
        )
    })?;
    let stdout_pipe = child.stdout.take().ok_or_else(|| {
        i18n::tr_f(
            req.language,
            "stream_audio.download_failed",
            &[("err", "yt-dlp stdout unavailable")],
        )
    })?;

    let progress_shared = Arc::new(AtomicU32::new(0));
    let activity_shared = Arc::new(AtomicU32::new(0));
    let stderr_shared = Arc::new(Mutex::new(String::new()));
    let stderr_thread = spawn_ytdlp_output_reader(
        stderr_pipe,
        req.ytdlp_debug,
        "stderr",
        Arc::clone(&stderr_shared),
        Arc::clone(&progress_shared),
        Arc::clone(&activity_shared),
    );
    let stdout_thread = spawn_ytdlp_output_reader(
        stdout_pipe,
        req.ytdlp_debug,
        "stdout",
        Arc::clone(&stderr_shared),
        Arc::clone(&progress_shared),
        Arc::clone(&activity_shared),
    );

    let allow_early_finalize = !matches!(
        req.dialog_data.format,
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
    let status = loop {
        match child.try_wait() {
            Ok(Some(status)) => break Some(status),
            Ok(None) => {
                if pump_messages_detect_stream_cancel(req.parent, req.progress) {
                    close_progress_dialog(req.progress);
                    if let Err(err) = child.kill() {
                        crate::log_debug(&format!(
                            "Failed to kill cancelled yt-dlp process: {}",
                            err
                        ));
                    }
                    return Err(String::new());
                }
                let pct = progress_shared.load(Ordering::Relaxed);
                if pct > last_pct {
                    last_pct = pct;
                    last_progress_at = std::time::Instant::now();
                    ui_target_pct = pct.min(99);
                    if last_pct >= 100 && reached_100_at.is_none() {
                        reached_100_at = Some(std::time::Instant::now());
                    }
                }
                if ui_pct < ui_target_pct {
                    ui_pct = ui_target_pct;
                    report_progress(req.progress, ui_pct);
                }
                if last_focus_keepalive.elapsed() > std::time::Duration::from_millis(300) {
                    keep_stream_progress_focus(req.progress);
                    last_focus_keepalive = std::time::Instant::now();
                }
                if allow_early_finalize
                    && reached_100_at
                        .map(|t| {
                            t.elapsed() > std::time::Duration::from_secs(STREAM_POST_100_GRACE_SECS)
                        })
                        .unwrap_or(false)
                    && find_latest_downloaded_stream_file(req.cache_dir, req.prefix).is_some()
                {
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
                    last_activity = activity;
                    last_progress_at = std::time::Instant::now();
                }
                if last_progress_at.elapsed()
                    > std::time::Duration::from_secs(STREAM_DOWNLOAD_STALL_SECS)
                {
                    stalled = true;
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
                return Err(i18n::tr_f(
                    req.language,
                    "stream_audio.download_failed",
                    &[("err", &err.to_string())],
                ));
            }
        }
    };

    if status.is_some() {
        if let Err(err) = stderr_thread.join() {
            crate::log_debug(&format!("yt-dlp stderr thread join failed: {:?}", err));
        }
        if let Err(err) = stdout_thread.join() {
            crate::log_debug(&format!("yt-dlp stdout thread join failed: {:?}", err));
        }
    }

    let stderr_capture = stderr_shared
        .lock()
        .map(|s| s.clone())
        .unwrap_or_else(|_| String::new());

    Ok(YtdlpDownloadAttempt {
        primary_path: find_latest_downloaded_stream_file(req.cache_dir, req.prefix),
        stderr_capture,
        stalled,
    })
}

fn parse_ytdlp_named_u64(line: &str, key: &str) -> Option<u64> {
    for separator in ['=', ':'] {
        let marker = format!("{key}{separator}");
        let start = line.find(&marker)? + marker.len();
        let digits: String = line[start..]
            .chars()
            .skip_while(|ch| ch.is_whitespace())
            .take_while(|ch| ch.is_ascii_digit())
            .collect();
        if let Ok(value) = digits.parse::<u64>() {
            return Some(value);
        }
    }
    None
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
                    PCWSTR(to_wide(&plain_label(&init.label)).as_ptr()),
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

                for option in &init.options {
                    let w = to_wide(option);
                    SendMessageW(combo, CB_ADDSTRING, WPARAM(0), LPARAM(w.as_ptr() as isize));
                }
                let default_selection = init
                    .default_selection
                    .min(init.options.len().saturating_sub(1));
                SendMessageW(combo, CB_SETCURSEL, WPARAM(default_selection), LPARAM(0));

                let state = Box::new(StreamTrackDialogState {
                    parent: init.parent,
                    language: init.language,
                    combo,
                    ok_button,
                    options: init.options,
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
                        let selected = (idx >= 0)
                            .then_some(idx as usize)
                            .filter(|idx| *idx < state.options.len());
                        if let Some(selected) = selected {
                            *state.result.lock().unwrap_or_else(|e| e.into_inner()) =
                                Some(selected);
                        }
                    })
                    .is_none()
                    {
                        crate::log_debug("Failed to access stream track dialog state");
                    }
                    crate::log_if_err!(PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)));
                    return LRESULT(0);
                }
                if cmd_id == STREAM_TRACK_ID_CANCEL || cmd_id == 2 {
                    if with_stream_track_dialog_state(hwnd, |_state| {}).is_none() {
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

pub(crate) fn choose_combo_option_dialog(
    parent: HWND,
    language: Language,
    title: String,
    label: String,
    options: Vec<String>,
    default_selection: usize,
) -> Option<usize> {
    if options.is_empty() {
        return None;
    }
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
        label,
        options,
        default_selection,
        result: Arc::clone(&result),
    });
    let title = to_wide(&plain_label(&title));
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
    let result_value = *result.lock().unwrap_or_else(|e| e.into_inner());
    unsafe {
        EnableWindow(parent, true);
        if result_value.is_none() {
            SetForegroundWindow(parent);
        }
    }
    result_value
}

fn choose_stream_audio_track(
    parent: HWND,
    language: Language,
    tracks: Vec<StreamAudioTrack>,
) -> Option<Option<String>> {
    let mut options = Vec::with_capacity(tracks.len() + 1);
    options.push(i18n::tr(language, "stream_audio.track.auto"));
    options.extend(tracks.iter().map(|track| track.label.clone()));
    let selected = choose_combo_option_dialog(
        parent,
        language,
        i18n::tr(language, "playback.audio_track"),
        i18n::tr(language, "playback.audio_track"),
        options,
        0,
    );
    match selected {
        Some(0) => Some(None),
        Some(index) => tracks
            .get(index - 1)
            .map(|track| Some(track.format_id.clone())),
        None => None,
    }
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
    let mut selected_label = dialog_data.reopen_selected_label.clone();
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
            dialog_data.reopen_selected_label.as_deref(),
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
        selected_label = resolved.selected_label;
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
            selected_label.as_deref(),
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
            selected_label.as_deref(),
        );
    }
    let should_reopen_selection =
        needs_ytdlp_selection || collection_url.is_some() || collection_page.is_some();

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
        crate::set_active_podcast_episode_info(
            parent,
            Some(url),
            None,
            episode_title,
            None,
            Some(stream_path),
        );
        if should_reopen_selection {
            update_youtube_return_context_from_selection(
                parent,
                &input,
                collection_url.as_deref(),
                collection_page,
                selected_label.as_deref(),
            );
        } else {
            crate::set_active_youtube_return_context(parent, None, None);
        }
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
                        selected_label = next_selection.selected_label;
                        update_youtube_return_context_from_selection(
                            parent,
                            &input,
                            collection_url.as_deref(),
                            collection_page,
                            selected_label.as_deref(),
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
        let downloading_status = i18n::tr(language, "stream_audio.progress_downloading");
        report_progress_status(progress, &downloading_status);
        crate::screen_reader_speak(&downloading_status);
        let stream_title = probe_stream_media_title(&ytdlp_path, &url);

        let mut progress = progress;
        let site_key = stream_auth_site_key(&url);
        let saved_credentials = site_key
            .as_deref()
            .and_then(|site| load_saved_stream_site_credentials(parent, site));
        let forced_credentials = if FORCE_YTDLP_AUTH_PROMPT_FOR_TESTING {
            if let Some(credentials) = saved_credentials.clone() {
                Some(credentials)
            } else {
                let Some(prompted) =
                    prompt_ytdlp_credentials(parent, language, "", site_key.is_some(), false)
                else {
                    close_progress_dialog(progress);
                    post_focus_editor(parent);
                    return;
                };
                if let Some(site) = site_key.as_ref() {
                    if prompted.save_credentials {
                        save_stream_site_credentials(parent, site, &prompted.credentials);
                    } else {
                        clear_stream_site_credentials(parent, site);
                    }
                }
                Some(prompted.credentials)
            }
        } else {
            None
        };
        let mut attempt = match run_ytdlp_stream_download_attempt(YtdlpDownloadRequest {
            parent,
            progress,
            language,
            ytdlp_path: &ytdlp_path,
            url: &url,
            cache_dir: &cache_dir,
            prefix: &prefix,
            output_template: &output_template,
            dialog_data: &dialog_data,
            selected_audio_format: selected_audio_format.as_deref(),
            ytdlp_debug,
            credentials: forced_credentials.as_ref().or(saved_credentials.as_ref()),
        }) {
            Ok(attempt) => attempt,
            Err(message) => {
                if !message.is_empty() {
                    close_progress_dialog(progress);
                    show_error(parent, language, &message);
                }
                return;
            }
        };
        let mut primary_path = attempt.primary_path.take();
        let mut stderr_capture = attempt.stderr_capture;
        let mut stalled = attempt.stalled;
        if primary_path.is_none() {
            let err = if stderr_capture.trim().is_empty() {
                "yt-dlp failed".to_string()
            } else {
                stderr_capture.trim().to_string()
            };
            if is_login_required_stream_error(&err) {
                let saved_credentials_failed = saved_credentials.is_some();
                let saved_username = saved_credentials
                    .as_ref()
                    .map(|credentials| credentials.username.clone())
                    .unwrap_or_default();
                if let Some(site) = site_key.as_deref()
                    && saved_credentials_failed
                {
                    clear_stream_site_credentials(parent, site);
                }
                close_progress_dialog(progress);
                let Some(prompted) = prompt_ytdlp_credentials(
                    parent,
                    language,
                    &saved_username,
                    site_key.is_some(),
                    saved_credentials_failed,
                ) else {
                    post_focus_editor(parent);
                    return;
                };
                if let Some(site) = site_key.as_ref() {
                    if prompted.save_credentials {
                        save_stream_site_credentials(parent, site, &prompted.credentials);
                    } else {
                        clear_stream_site_credentials(parent, site);
                    }
                }
                let auth_prefix = format!("{prefix}_auth");
                let auth_output_template = cache_dir.join(format!("{auth_prefix}.%(ext)s"));
                progress = open_progress_dialog(
                    parent,
                    language,
                    "stream_audio.progress_title",
                    "stream_audio.progress_downloading",
                    true,
                );
                keep_stream_progress_focus(progress);
                attempt = match run_ytdlp_stream_download_attempt(YtdlpDownloadRequest {
                    parent,
                    progress,
                    language,
                    ytdlp_path: &ytdlp_path,
                    url: &url,
                    cache_dir: &cache_dir,
                    prefix: &auth_prefix,
                    output_template: &auth_output_template,
                    dialog_data: &dialog_data,
                    selected_audio_format: selected_audio_format.as_deref(),
                    ytdlp_debug,
                    credentials: Some(&prompted.credentials),
                }) {
                    Ok(attempt) => attempt,
                    Err(message) => {
                        if !message.is_empty() {
                            close_progress_dialog(progress);
                            show_error(parent, language, &message);
                        }
                        return;
                    }
                };
                primary_path = attempt.primary_path.take();
                stderr_capture = attempt.stderr_capture;
                stalled = attempt.stalled;
            }
        }
        if ytdlp_debug && !stderr_capture.trim().is_empty() {
            crate::log_debug(&format!(
                "yt-dlp combined stderr: {}",
                truncate_debug_text(stderr_capture.trim(), 5000)
            ));
        }

        report_progress_status(progress, &i18n::tr(language, "podcasts.loading"));
        log_stream_focus_snapshot("stream_download.after_download_before_finalize", progress);
        let downloaded_path = if primary_path.is_some() {
            primary_path
        } else {
            let err = if stderr_capture.trim().is_empty() {
                "yt-dlp failed".to_string()
            } else {
                stderr_capture.trim().to_string()
            };
            if is_drm_not_supported_stream_error(&err) {
                close_progress_dialog(progress);
                show_error(
                    parent,
                    language,
                    &drm_not_supported_stream_message(language),
                );
                return;
            }
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
                            .arg("--progress-template")
                            .arg(ytdlp_download_progress_template())
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
                                let mut retry_focus_keepalive = std::time::Instant::now();
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
                                            if retry_focus_keepalive.elapsed()
                                                > std::time::Duration::from_millis(300)
                                            {
                                                keep_stream_progress_focus(progress);
                                                retry_focus_keepalive = std::time::Instant::now();
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

        let cache_limit_bytes = with_state(parent, |state| {
            state.settings.podcast_cache_limit_mb as u64 * 1024 * 1024
        })
        .unwrap_or(500 * 1024 * 1024);
        crate::app_windows::podcasts_window::enforce_podcast_cache_limit(
            &cache_dir,
            cache_limit_bytes,
            Some(&playback_path),
        );

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
            None,
            episode_title,
            None,
            Some(playback_path),
        );
        if should_reopen_selection {
            update_youtube_return_context_from_selection(
                parent,
                &input,
                collection_url.as_deref(),
                collection_page,
                selected_label.as_deref(),
            );
        } else {
            crate::set_active_youtube_return_context(parent, None, None);
        }
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
    RateLimited,
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
        if combined.contains("HTTP Error 429")
            || combined.contains("Too Many Requests")
            || combined.contains("too many requests")
        {
            cleanup_ytdlp_temp_dir(&temp_dir);
            return Err(ImportError::RateLimited);
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
        ImportError::RateLimited => labels.rate_limited,
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
