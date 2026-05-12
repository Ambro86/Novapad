use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{PBM_SETPOS, PBM_SETRANGE, PROGRESS_CLASSW};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus, VK_ESCAPE};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, GWLP_USERDATA,
    GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW, IsWindow,
    LoadCursorW, RegisterClassW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW,
    SetWindowTextW, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN,
    WM_NCDESTROY, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

use crate::accessibility::{handle_accessibility, to_wide};
use crate::i18n;
use crate::settings::Language;
use crate::{log_debug, show_error, show_info, with_state};

const CLASS_NAME: &str = "SonarpadMediaSplit";
const ID_INPUT: usize = 9701;
const ID_START: usize = 9702;
const ID_CLOSE: usize = 9703;
const WM_SPLIT_PROGRESS: u32 = WM_APP + 180;
const WM_SPLIT_DONE: u32 = WM_APP + 181;
const WM_SPLIT_CLOSE_AFTER_SUCCESS: u32 = WM_APP + 182;

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum MediaSplitMode {
    Parts,
    Time,
}

struct SplitDone {
    result: Result<PathBuf, SplitFailure>,
}

struct CompletionDialog {
    message: String,
    language: Language,
    success: bool,
    close_after: bool,
    owner: HWND,
}

enum SplitFailure {
    Message(String),
    TimeTooLong {
        part_duration: String,
        media_duration: String,
    },
}

struct SplitWindowState {
    parent: HWND,
    input_path: PathBuf,
    mode: MediaSplitMode,
    input: HWND,
    progress: HWND,
    status: HWND,
    start_button: HWND,
    close_button: HWND,
    cancel: Arc<AtomicBool>,
    running: bool,
    language: Language,
}

pub fn open(parent: HWND, mode: MediaSplitMode) {
    let existing = with_state(parent, |state| state.media_split_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        unsafe {
            if IsWindow(existing).as_bool() {
                SetForegroundWindow(existing);
                return;
            }
        }
        with_state(parent, |state| state.media_split_window = HWND(0));
    }

    let Some(input_path) = crate::current_local_playback_media_path(parent) else {
        let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
        show_error(
            parent,
            language,
            &i18n::tr(language, "media_split.error.no_local_media"),
        );
        return;
    };

    unsafe {
        let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
        let class_name = to_wide(CLASS_NAME);
        let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
        let title = match mode {
            MediaSplitMode::Parts => i18n::tr(language, "media_split.parts.title"),
            MediaSplitMode::Time => i18n::tr(language, "media_split.time.title"),
        };
        let title_w = to_wide(&title);

        let wc = WNDCLASSW {
            hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
                LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
            ),
            hInstance: hinstance,
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(split_wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let state = Box::new(SplitWindowState {
            parent,
            input_path,
            mode,
            input: HWND(0),
            progress: HWND(0),
            status: HWND(0),
            start_button: HWND(0),
            close_button: HWND(0),
            cancel: Arc::new(AtomicBool::new(false)),
            running: false,
            language,
        });
        let state_ptr = Box::into_raw(state);
        let hwnd = CreateWindowExW(
            WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            440,
            210,
            parent,
            HMENU(0),
            hinstance,
            Some(state_ptr as *const _),
        );
        if hwnd.0 == 0 {
            let _unused = Box::from_raw(state_ptr);
            return;
        }
        with_state(parent, |state| state.media_split_window = hwnd);
        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);
    }
}

pub fn handle_navigation(hwnd: HWND, msg: &windows::Win32::UI::WindowsAndMessaging::MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
        crate::log_if_err!(crate::post_message_w_safe(
            hwnd,
            WM_CLOSE,
            WPARAM(0),
            LPARAM(0),
        ));
        return true;
    }
    handle_accessibility(hwnd, msg)
}

unsafe extern "system" fn split_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "media_split_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || split_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn split_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let cs = lparam.0 as *const CREATESTRUCTW;
                let state_ptr = (*cs).lpCreateParams as *mut SplitWindowState;
                if state_ptr.is_null() {
                    return LRESULT(0);
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, state_ptr as isize);
                let state = &mut *state_ptr;
                let language = state.language;
                let label = match state.mode {
                    MediaSplitMode::Parts => i18n::tr(language, "media_split.parts.label"),
                    MediaSplitMode::Time => i18n::tr(language, "media_split.time.label"),
                };
                let hint = match state.mode {
                    MediaSplitMode::Parts => i18n::tr(language, "media_split.parts.hint"),
                    MediaSplitMode::Time => i18n::tr(language, "media_split.time.hint"),
                };
                let start = i18n::tr(language, "media_split.start");
                let close = i18n::tr(language, "media_split.close");
                let ready = i18n::tr(language, "media_split.status.ready");
                let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);

                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&label).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    12,
                    14,
                    400,
                    18,
                    hwnd,
                    HMENU(0),
                    hinstance,
                    None,
                );
                let input = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    12,
                    36,
                    140,
                    24,
                    hwnd,
                    HMENU(ID_INPUT as isize),
                    hinstance,
                    None,
                );
                let default_value = match state.mode {
                    MediaSplitMode::Parts => "2".to_string(),
                    MediaSplitMode::Time => "00:30:00".to_string(),
                };
                crate::log_if_err!(SetWindowTextW(
                    input,
                    PCWSTR(to_wide(&default_value).as_ptr()),
                ));
                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&hint).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    12,
                    66,
                    400,
                    18,
                    hwnd,
                    HMENU(0),
                    hinstance,
                    None,
                );
                let progress = CreateWindowExW(
                    Default::default(),
                    PROGRESS_CLASSW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE,
                    12,
                    92,
                    400,
                    22,
                    hwnd,
                    HMENU(0),
                    hinstance,
                    None,
                );
                SendMessageW(progress, PBM_SETRANGE, WPARAM(0), LPARAM(100isize << 16));
                SendMessageW(progress, PBM_SETPOS, WPARAM(0), LPARAM(0));
                let status = CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&ready).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    12,
                    120,
                    400,
                    18,
                    hwnd,
                    HMENU(0),
                    hinstance,
                    None,
                );
                let start_button = CreateWindowExW(
                    Default::default(),
                    w!("BUTTON"),
                    PCWSTR(to_wide(&start).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    160,
                    145,
                    120,
                    28,
                    hwnd,
                    HMENU(ID_START as isize),
                    hinstance,
                    None,
                );
                let close_button = CreateWindowExW(
                    Default::default(),
                    w!("BUTTON"),
                    PCWSTR(to_wide(&close).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    292,
                    145,
                    120,
                    28,
                    hwnd,
                    HMENU(ID_CLOSE as isize),
                    hinstance,
                    None,
                );
                state.input = input;
                state.progress = progress;
                state.status = status;
                state.start_button = start_button;
                state.close_button = close_button;
                SetFocus(input);
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                match id {
                    ID_START => start_split(hwnd),
                    ID_CLOSE => {
                        crate::log_if_err!(crate::post_message_w_safe(
                            hwnd,
                            WM_CLOSE,
                            WPARAM(0),
                            LPARAM(0),
                        ));
                    }
                    _ => {}
                }
                LRESULT(0)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    crate::log_if_err!(crate::post_message_w_safe(
                        hwnd,
                        WM_CLOSE,
                        WPARAM(0),
                        LPARAM(0),
                    ));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_SPLIT_PROGRESS => {
                with_split_state(hwnd, |state| {
                    let pct = wparam.0.min(100);
                    SendMessageW(state.progress, PBM_SETPOS, WPARAM(pct), LPARAM(0));
                    let text = i18n::tr_f(
                        state.language,
                        "media_split.status.progress",
                        &[("pct", &pct.to_string())],
                    );
                    crate::log_if_err!(SetWindowTextW(
                        state.status,
                        PCWSTR(to_wide(&text).as_ptr()),
                    ));
                });
                LRESULT(0)
            }
            WM_SPLIT_DONE => {
                let done_ptr = lparam.0 as *mut SplitDone;
                if !done_ptr.is_null() {
                    log_debug("Media split: WM_SPLIT_DONE received");
                    let done = Box::from_raw(done_ptr);
                    let dialog = with_split_state(hwnd, |state| {
                        state.running = false;
                        EnableWindow(state.start_button, true);
                        match &done.result {
                            Ok(folder) => {
                                SendMessageW(state.progress, PBM_SETPOS, WPARAM(100), LPARAM(0));
                                let msg = i18n::tr_f(
                                    state.language,
                                    "media_split.success",
                                    &[("folder", &folder.to_string_lossy())],
                                );
                                crate::log_if_err!(SetWindowTextW(
                                    state.status,
                                    PCWSTR(to_wide(&msg).as_ptr()),
                                ));
                                Some(CompletionDialog {
                                    message: msg,
                                    language: state.language,
                                    success: true,
                                    close_after: true,
                                    owner: state.parent,
                                })
                            }
                            Err(SplitFailure::Message(err)) => {
                                let msg = i18n::tr_f(
                                    state.language,
                                    "media_split.error.failed",
                                    &[("err", err)],
                                );
                                crate::log_if_err!(SetWindowTextW(
                                    state.status,
                                    PCWSTR(to_wide(&msg).as_ptr()),
                                ));
                                Some(CompletionDialog {
                                    message: msg,
                                    language: state.language,
                                    success: false,
                                    close_after: false,
                                    owner: state.parent,
                                })
                            }
                            Err(SplitFailure::TimeTooLong {
                                part_duration,
                                media_duration,
                            }) => {
                                let msg = i18n::tr_f(
                                    state.language,
                                    "media_split.error.time_too_long",
                                    &[
                                        ("part_duration", part_duration),
                                        ("media_duration", media_duration),
                                    ],
                                );
                                crate::log_if_err!(SetWindowTextW(
                                    state.status,
                                    PCWSTR(to_wide(&msg).as_ptr()),
                                ));
                                Some(CompletionDialog {
                                    message: msg,
                                    language: state.language,
                                    success: false,
                                    close_after: true,
                                    owner: state.parent,
                                })
                            }
                        }
                    })
                    .flatten();
                    if let Some(dialog) = dialog {
                        if dialog.success {
                            show_info(dialog.owner, dialog.language, &dialog.message);
                            log_debug("Media split: completion dialog closed, posting close");
                            crate::log_if_err!(crate::post_message_w_safe(
                                hwnd,
                                WM_SPLIT_CLOSE_AFTER_SUCCESS,
                                WPARAM(0),
                                LPARAM(0),
                            ));
                        } else {
                            show_error(dialog.owner, dialog.language, &dialog.message);
                            if dialog.close_after {
                                log_debug("Media split: error dialog closed, posting close");
                                crate::log_if_err!(crate::post_message_w_safe(
                                    hwnd,
                                    WM_SPLIT_CLOSE_AFTER_SUCCESS,
                                    WPARAM(0),
                                    LPARAM(0),
                                ));
                            }
                        }
                    }
                }
                LRESULT(0)
            }
            WM_SPLIT_CLOSE_AFTER_SUCCESS => {
                log_debug("Media split: close-after-success received");
                if with_split_state(hwnd, |state| state.running).unwrap_or(false) {
                    log_debug("Media split: close-after-success skipped while running");
                } else {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                }
                LRESULT(0)
            }
            WM_CLOSE => {
                let can_close = with_split_state(hwnd, |state| {
                    if state.running {
                        state.cancel.store(true, Ordering::Relaxed);
                        let msg = i18n::tr(state.language, "media_split.status.canceling");
                        crate::log_if_err!(SetWindowTextW(
                            state.status,
                            PCWSTR(to_wide(&msg).as_ptr()),
                        ));
                        false
                    } else {
                        true
                    }
                })
                .unwrap_or(true);
                if can_close {
                    crate::log_if_err!(windows::Win32::UI::WindowsAndMessaging::DestroyWindow(
                        hwnd
                    ));
                }
                LRESULT(0)
            }
            WM_DESTROY => {
                log_debug("Media split: WM_DESTROY");
                with_split_state(hwnd, |state| {
                    with_state(state.parent, |app_state| {
                        app_state.media_split_window = HWND(0)
                    });
                    EnableWindow(state.parent, true);
                    if IsWindow(state.parent).as_bool() {
                        SetForegroundWindow(state.parent);
                        log_debug("Media split: parent re-enabled and foreground restored");
                    }
                });
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SplitWindowState;
                if !ptr.is_null() {
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                    let _unused = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn start_split(hwnd: HWND) {
    let params = with_split_state(hwnd, |state| {
        if state.running {
            return None;
        }
        let input_text = get_window_text(state.input).trim().to_string();
        let value = match state.mode {
            MediaSplitMode::Parts => match input_text.parse::<u32>() {
                Ok(parts) if (2..=999).contains(&parts) => SplitValue::Parts(parts),
                _ => {
                    show_error(
                        hwnd,
                        state.language,
                        &i18n::tr(state.language, "media_split.error.invalid_parts"),
                    );
                    return None;
                }
            },
            MediaSplitMode::Time => match crate::audio_player::parse_time_input(&input_text) {
                Ok(seconds) if seconds > 0 => SplitValue::Seconds(seconds),
                _ => {
                    show_error(
                        hwnd,
                        state.language,
                        &i18n::tr(state.language, "media_split.error.invalid_time"),
                    );
                    return None;
                }
            },
        };
        state.running = true;
        state.cancel.store(false, Ordering::Relaxed);
        unsafe {
            EnableWindow(state.start_button, false);
            SendMessageW(state.progress, PBM_SETPOS, WPARAM(0), LPARAM(0));
            crate::log_if_err!(SetWindowTextW(
                state.status,
                PCWSTR(to_wide(&i18n::tr(state.language, "media_split.status.running")).as_ptr()),
            ));
        }
        Some((state.input_path.clone(), value, state.cancel.clone()))
    })
    .flatten();

    let Some((input_path, value, cancel)) = params else {
        return;
    };

    let hwnd_value = hwnd.0;
    thread::spawn(move || {
        let hwnd = HWND(hwnd_value);
        let result = split_media_file(&input_path, value, cancel, |pct| {
            crate::log_if_err!(crate::post_message_w_safe(
                hwnd,
                WM_SPLIT_PROGRESS,
                WPARAM(pct as usize),
                LPARAM(0),
            ));
        });
        let payload = Box::into_raw(Box::new(SplitDone { result }));
        if let Err(err) =
            crate::post_message_w_safe(hwnd, WM_SPLIT_DONE, WPARAM(0), LPARAM(payload as isize))
        {
            log_debug(&format!("Media split: failed to post done message: {err}"));
            let _unused = unsafe { Box::from_raw(payload) };
        }
    });
}

#[derive(Clone, Copy)]
enum SplitValue {
    Parts(u32),
    Seconds(u64),
}

fn split_media_file<F>(
    input_path: &Path,
    value: SplitValue,
    cancel: Arc<AtomicBool>,
    mut progress: F,
) -> Result<PathBuf, SplitFailure>
where
    F: FnMut(u32),
{
    if cancel.load(Ordering::Relaxed) {
        return Err(SplitFailure::Message("Canceled".to_string()));
    }
    let duration = crate::ffmpeg_export::media_duration_secs(input_path)
        .or_else(|| crate::audio_player::audiobook_duration_secs(input_path))
        .ok_or_else(|| {
            SplitFailure::Message("FFmpeg: unable to detect media duration".to_string())
        })?;
    let segment_seconds = match value {
        SplitValue::Parts(parts) => duration.div_ceil(parts as u64).max(1),
        SplitValue::Seconds(seconds) => {
            if seconds >= duration {
                return Err(SplitFailure::TimeTooLong {
                    part_duration: format_duration(seconds),
                    media_duration: format_duration(duration),
                });
            }
            seconds.max(1)
        }
    };
    let segment_seconds = segment_seconds.min(u32::MAX as u64) as u32;
    let folder = output_folder(input_path)?;
    std::fs::create_dir_all(&folder).map_err(|e| SplitFailure::Message(e.to_string()))?;
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .filter(|s| !s.trim().is_empty())
        .unwrap_or("media");
    let ext = input_path
        .extension()
        .and_then(|s| s.to_str())
        .filter(|s| !s.trim().is_empty())
        .unwrap_or("mp3");
    let pattern = folder.join(format!("{}_part_%03d.{}", sanitize_name(stem), ext));
    let mut cb = |pct: u32| {
        if !cancel.load(Ordering::Relaxed) {
            progress(pct);
        }
    };
    crate::ffmpeg_export::segment_media_file(
        input_path,
        &pattern,
        segment_seconds,
        1,
        Some(&mut cb),
    )
    .map_err(SplitFailure::Message)?;
    if cancel.load(Ordering::Relaxed) {
        return Err(SplitFailure::Message("Canceled".to_string()));
    }
    progress(100);
    Ok(folder)
}

fn output_folder(input_path: &Path) -> Result<PathBuf, SplitFailure> {
    let parent = input_path
        .parent()
        .ok_or_else(|| SplitFailure::Message("Invalid input folder".to_string()))?;
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .filter(|s| !s.trim().is_empty())
        .unwrap_or("media");
    Ok(parent.join(format!("{}_parts", sanitize_name(stem))))
}

fn format_duration(seconds: u64) -> String {
    let hours = seconds / 3600;
    let minutes = (seconds % 3600) / 60;
    let seconds = seconds % 60;
    if hours > 0 {
        format!("{hours:02}:{minutes:02}:{seconds:02}")
    } else {
        format!("{minutes:02}:{seconds:02}")
    }
}

fn sanitize_name(name: &str) -> String {
    let mut out = String::new();
    for ch in name.chars() {
        if matches!(ch, '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*') {
            out.push('_');
        } else {
            out.push(ch);
        }
    }
    if out.trim().is_empty() {
        "media".to_string()
    } else {
        out
    }
}

fn with_split_state<R>(hwnd: HWND, f: impl FnOnce(&mut SplitWindowState) -> R) -> Option<R> {
    unsafe {
        let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SplitWindowState;
        if ptr.is_null() {
            None
        } else {
            Some(f(&mut *ptr))
        }
    }
}

fn get_window_text(hwnd: HWND) -> String {
    unsafe {
        let len = GetWindowTextLengthW(hwnd);
        let mut buf = vec![0u16; len as usize + 1];
        let read = GetWindowTextW(hwnd, &mut buf);
        String::from_utf16_lossy(&buf[..read as usize])
    }
}
