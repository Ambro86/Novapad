#![allow(clippy::bool_comparison)]
use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
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
    DefWindowProcW, DestroyWindow, DispatchMessageW, EN_CHANGE, GWLP_USERDATA, GetMessageW,
    GetWindowLongPtrW, HMENU, IDC_ARROW, IsDialogMessageW, IsWindow, LoadCursorW, MSG,
    PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW,
    SetWindowTextW, TranslateMessage, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE,
    WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE,
    WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

use url::Url;
use yt_transcript_rs::errors::CouldNotRetrieveTranscriptReason;
use yt_transcript_rs::{Transcript, TranscriptList, YouTubeTranscriptApi};

use crate::accessibility::{EM_REPLACESEL, from_wide, to_wide, to_wide_normalized};
use crate::editor_manager::get_edit_text;
use crate::i18n;
use crate::settings::{Language, save_settings};
use crate::with_state;
use crate::{WM_FOCUS_EDITOR, get_active_edit, show_error};

const YT_IMPORT_CLASS_NAME: &str = "NovapadYouTubeTranscript";
const YT_ID_URL: usize = 9301;
const YT_ID_LOAD: usize = 9302;
const YT_ID_LANG: usize = 9303;
const YT_ID_TIMESTAMP: usize = 9304;
const YT_ID_OK: usize = 9305;
const YT_ID_CANCEL: usize = 9306;
const WM_YT_LOAD_COMPLETE: u32 = WM_APP + 40;

#[derive(Clone)]
struct ImportResult {
    transcript: Transcript,
    include_timestamps: bool,
}

struct ImportInit {
    parent: HWND,
    language: Language,
    include_timestamps: bool,
    result: Arc<Mutex<Option<ImportResult>>>,
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
    transcripts: Vec<Transcript>,
    result: Arc<Mutex<Option<ImportResult>>>,
}

struct Labels {
    title: String,
    url: String,
    load: String,
    language: String,
    include_timestamps: String,
    loading: String,
    ok: String,
    cancel: String,
    auto: String,
    invalid_url: String,
    no_transcript: String,
    network_error: String,
    import_error: String,
    no_document: String,
}

fn labels(language: Language) -> Labels {
    Labels {
        title: i18n::tr(language, "youtube.title"),
        url: i18n::tr(language, "youtube.url"),
        load: i18n::tr(language, "youtube.load"),
        language: i18n::tr(language, "youtube.language"),
        include_timestamps: i18n::tr(language, "youtube.include_timestamps"),
        loading: i18n::tr(language, "youtube.loading"),
        ok: i18n::tr(language, "youtube.ok"),
        cancel: i18n::tr(language, "youtube.cancel"),
        auto: i18n::tr(language, "youtube.auto"),
        invalid_url: i18n::tr(language, "youtube.invalid_url"),
        no_transcript: i18n::tr(language, "youtube.no_transcript"),
        network_error: i18n::tr(language, "youtube.network_error"),
        import_error: i18n::tr(language, "youtube.import_error"),
        no_document: i18n::tr(language, "youtube.no_document"),
    }
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
            let _ = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
        }
        return;
    };
    let _ = unsafe {
        with_state(parent, |state| {
            state.settings.youtube_include_timestamps = result.include_timestamps;
            save_settings(state.settings.clone());
        })
    };

    let text = match fetch_transcript_text(&result.transcript, result.include_timestamps) {
        Ok(text) => text,
        Err(err) => {
            unsafe {
                show_error(parent, language, &error_message(language, &err));
                let _ = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
            }
            return;
        }
    };

    unsafe {
        let Some(hwnd_edit) = get_active_edit(parent) else {
            show_error(parent, language, &labels(language).no_document);
            let _ = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
            return;
        };
        let _ = SetFocus(hwnd_edit);
        let existing = get_edit_text(hwnd_edit);
        let combined = if existing.is_empty() {
            text
        } else {
            format!("{text}\n\n{existing}")
        };
        let wide = to_wide_normalized(&combined);
        let _ = SendMessageW(hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(-1));
        let _ = SendMessageW(
            hwnd_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(wide.as_ptr() as isize),
        );
        let end = combined.len() as i32;
        let _ = SendMessageW(
            hwnd_edit,
            EM_SETSEL,
            WPARAM(end as usize),
            LPARAM(end as isize),
        );
        let cr = CHARRANGE { cpMin: 0, cpMax: 0 };
        let _ = SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&cr as *const _ as isize),
        );
        let _ = SendMessageW(hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        let _ = PostMessageW(parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
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
        if unsafe { IsWindow(hwnd).as_bool() } == false {
            break;
        }
        let res = unsafe { GetMessageW(&mut msg, HWND(0), 0, 0) };
        if res.0 == 0 {
            break;
        }
        unsafe {
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_CANCEL as usize), LPARAM(0));
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let ok = with_import_state(hwnd, |state| state.ok_button).unwrap_or(HWND(0));
                if GetFocus() == ok {
                    let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_OK as usize), LPARAM(0));
                    continue;
                }
            }
            if IsDialogMessageW(hwnd, &mut msg).as_bool() {
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

    result.lock().unwrap().clone()
}

unsafe extern "system" fn import_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
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
                Default::default(),
                WC_STATIC,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                110,
                86,
                390,
                18,
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
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

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
                transcripts: Vec::new(),
                result: init.result.clone(),
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            let initial_check = if init.include_timestamps {
                BST_CHECKED.0
            } else {
                0
            };
            let _ = SendMessageW(
                timestamp_check,
                BM_SETCHECK,
                WPARAM(initial_check as usize),
                LPARAM(0),
            );
            let _ = SetFocus(url_edit);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            let notification = ((wparam.0 >> 16) & 0xffff) as u16;
            if cmd_id == YT_ID_LOAD {
                start_load_languages(hwnd);
                LRESULT(0)
            } else if cmd_id == YT_ID_OK {
                let mut should_close = false;
                let _ = with_import_state(hwnd, |state| {
                    if state.loading {
                        return;
                    }
                    if state.transcripts.is_empty() {
                        let _ = start_load_languages(hwnd);
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
                    *state.result.lock().unwrap() = Some(ImportResult {
                        transcript,
                        include_timestamps,
                    });
                    should_close = true;
                });
                if should_close {
                    let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                }
                LRESULT(0)
            } else if cmd_id == YT_ID_CANCEL {
                let _ = with_import_state(hwnd, |state| {
                    *state.result.lock().unwrap() = None;
                });
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                LRESULT(0)
            } else if cmd_id == YT_ID_URL && notification as u32 == EN_CHANGE {
                let _ = with_import_state(hwnd, |state| {
                    if state.loading {
                        return;
                    }
                    state.transcripts.clear();
                    let _ = SendMessageW(state.lang_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
                });
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
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_CANCEL as usize), LPARAM(0));
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let url_edit = with_import_state(hwnd, |state| state.url_edit).unwrap_or(HWND(0));
                if focus == url_edit {
                    let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(YT_ID_LOAD as usize), LPARAM(0));
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let _ = with_import_state(hwnd, |state| {
                EnableWindow(state.parent, true);
                SetForegroundWindow(state.parent);
                let _ = PostMessageW(state.parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
            });
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ImportState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
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
    transcripts: Vec<Transcript>,
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

    let _ = unsafe {
        with_import_state(hwnd, |state| {
            already_loading = state.loading;
            if already_loading {
                return;
            }
            language = state.language;
            url = read_edit_text(state.url_edit);
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            state.loading = true;
        })
    };
    if already_loading {
        return false;
    }

    let labels_data = labels(language);
    unsafe {
        let _ = SetWindowTextW(status, PCWSTR(to_wide(&labels_data.loading).as_ptr()));
        let _ = EnableWindow(edit, false);
        let _ = EnableWindow(load_button, false);
        let _ = EnableWindow(combo, false);
        let _ = EnableWindow(timestamp, false);
        let _ = EnableWindow(ok_button, false);
    }

    std::thread::spawn(move || {
        let result = if let Some(video_id) = extract_video_id(&url) {
            match fetch_transcript_list(&video_id) {
                Ok(list) => {
                    let transcripts = collect_transcripts(list);
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
        } else {
            LoadResult {
                transcripts: Vec::new(),
                error: Some(ImportError::InvalidUrl),
            }
        };
        unsafe {
            let _ = PostMessageW(
                hwnd,
                WM_YT_LOAD_COMPLETE,
                WPARAM(0),
                LPARAM(Box::into_raw(Box::new(result)) as isize),
            );
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

    let _ = unsafe {
        with_import_state(hwnd, |state| {
            language = state.language;
            edit = state.url_edit;
            ok_button = state.ok_button;
            load_button = state.load_button;
            combo = state.lang_combo;
            timestamp = state.timestamp_check;
            status = state.status_label;
            state.loading = false;
        })
    };

    let labels_data = labels(language);
    unsafe {
        let _ = SetWindowTextW(status, PCWSTR(to_wide("").as_ptr()));
        let _ = EnableWindow(edit, true);
        let _ = EnableWindow(load_button, true);
        let _ = EnableWindow(combo, true);
        let _ = EnableWindow(timestamp, true);
        let _ = EnableWindow(ok_button, true);
    }

    if let Some(err) = result.error {
        unsafe {
            show_error(hwnd, language, &error_message(language, &err));
            let _ = SetFocus(edit);
        }
        return;
    }

    unsafe {
        let _ = SendMessageW(combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for transcript in result.transcripts.iter() {
            let mut label = format!("{} ({})", transcript.language(), transcript.language_code());
            if transcript.is_generated() {
                label.push_str(&format!(" - {}", labels_data.auto));
            }
            let wide = to_wide(&label);
            let _ = SendMessageW(
                combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(wide.as_ptr() as isize),
            );
        }
        let _ = SendMessageW(combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        let _ = SetFocus(combo);
    }

    let _ = unsafe {
        with_import_state(hwnd, |state| {
            state.transcripts = result.transcripts;
        })
    };
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
        let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowTextW(hwnd, &mut buf);
    }
    unsafe { from_wide(buf.as_ptr()) }
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

fn collect_transcripts(list: TranscriptList) -> Vec<Transcript> {
    let mut manual: Vec<Transcript> = list
        .manually_created_transcripts
        .values()
        .cloned()
        .collect();
    let mut generated: Vec<Transcript> = list.generated_transcripts.values().cloned().collect();
    manual.sort_by(|a, b| {
        a.language()
            .cmp(b.language())
            .then(a.language_code().cmp(b.language_code()))
    });
    generated.sort_by(|a, b| {
        a.language()
            .cmp(b.language())
            .then(a.language_code().cmp(b.language_code()))
    });
    manual.extend(generated);
    manual
}

#[derive(Debug)]
enum ImportError {
    InvalidUrl,
    NoTranscript,
    Network,
    Other,
}

fn fetch_transcript_list(video_id: &str) -> Result<TranscriptList, ImportError> {
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .map_err(|_| ImportError::Other)?;
    rt.block_on(async {
        let api = YouTubeTranscriptApi::new(None, None, None).map_err(|_| ImportError::Other)?;
        api.list_transcripts(video_id)
            .await
            .map_err(map_transcript_error)
    })
}

fn fetch_transcript_text(
    transcript: &Transcript,
    include_timestamps: bool,
) -> Result<String, ImportError> {
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .map_err(|_| ImportError::Other)?;
    rt.block_on(async {
        let client = reqwest::Client::new();
        let fetched = transcript
            .fetch(&client, false)
            .await
            .map_err(map_transcript_error)?;
        if include_timestamps {
            Ok(format_with_timestamps(&fetched))
        } else {
            Ok(format_without_timestamps(&fetched))
        }
    })
}

fn map_transcript_error(err: yt_transcript_rs::errors::CouldNotRetrieveTranscript) -> ImportError {
    match err.reason {
        Some(CouldNotRetrieveTranscriptReason::InvalidVideoId) => ImportError::InvalidUrl,
        Some(CouldNotRetrieveTranscriptReason::TranscriptsDisabled)
        | Some(CouldNotRetrieveTranscriptReason::NoTranscriptFound { .. })
        | Some(CouldNotRetrieveTranscriptReason::VideoUnavailable)
        | Some(CouldNotRetrieveTranscriptReason::VideoUnplayable { .. }) => {
            ImportError::NoTranscript
        }
        Some(CouldNotRetrieveTranscriptReason::YouTubeRequestFailed(_))
        | Some(CouldNotRetrieveTranscriptReason::RequestBlocked(_))
        | Some(CouldNotRetrieveTranscriptReason::IpBlocked(_)) => ImportError::Network,
        _ => ImportError::Other,
    }
}

fn error_message(language: Language, err: &ImportError) -> String {
    let labels = labels(language);
    match err {
        ImportError::InvalidUrl => labels.invalid_url,
        ImportError::NoTranscript => labels.no_transcript,
        ImportError::Network => labels.network_error,
        ImportError::Other => labels.import_error,
    }
}

fn format_with_timestamps(fetched: &yt_transcript_rs::FetchedTranscript) -> String {
    let mut lines = Vec::new();
    for part in fetched.parts() {
        let stamp = format_timestamp(part.start);
        let text = clean_transcript_text(&part.text);
        if text.is_empty() {
            continue;
        }
        lines.push(format!("[{stamp}] {text}"));
    }
    lines.join("\n")
}

fn format_without_timestamps(fetched: &yt_transcript_rs::FetchedTranscript) -> String {
    let mut parts = Vec::new();
    for part in fetched.parts() {
        let text = clean_transcript_text(&part.text);
        if text.is_empty() {
            continue;
        }
        parts.push(text);
    }
    parts.join(" ")
}

fn clean_transcript_text(text: &str) -> String {
    let trimmed = text.trim_start();
    let cleaned = trimmed.strip_prefix(">>").unwrap_or(trimmed).trim_start();
    cleaned.to_string()
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
