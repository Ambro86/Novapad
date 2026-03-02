use std::path::{Path, PathBuf};
use std::sync::{
    Arc,
    atomic::{AtomicBool, Ordering},
};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, GetSaveFileNameW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY,
    OFN_OVERWRITEPROMPT, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{WC_BUTTON, WC_COMBOBOXW, WC_EDIT, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus, VK_ESCAPE};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL, CB_RESETCONTENT, CB_SETCURSEL, CBS_DROPDOWN,
    CBS_DROPDOWNLIST, CREATESTRUCTW, CreateWindowExW, DefWindowProcW, DestroyWindow,
    ES_AUTOHSCROLL, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW,
    IsWindow, LoadCursorW, PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow,
    SetWindowLongPtrW, SetWindowTextW, ShowWindow, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE,
    WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, PWSTR};

use crate::accessibility::to_wide;
use crate::app_windows::podcast_save_window::{self, SaveDialogLabels};
use crate::ffmpeg_export::{
    ConvertAudioFormat, ConvertAudioQuality, ConvertAudioSettings, build_ffmpeg_args,
    convert_audio_file, validate_mp3_bitrate,
};
use crate::i18n;
use crate::settings::{FileFormat, Language};
use crate::{log_debug, show_error, show_info, with_state};

const CONVERT_CLASS_NAME: &str = "SonarpadConvertAudio";

const CONVERT_ID_INPUT_EDIT: usize = 9601;
const CONVERT_ID_INPUT_BROWSE: usize = 9602;
const CONVERT_ID_OUTPUT_EDIT: usize = 9603;
const CONVERT_ID_OUTPUT_BROWSE: usize = 9604;
const CONVERT_ID_FORMAT: usize = 9605;
const CONVERT_ID_QUALITY_EDIT: usize = 9606;
const CONVERT_ID_QUALITY_COMBO: usize = 9607;
const CONVERT_ID_CONVERT: usize = 9608;
const CONVERT_ID_CLOSE: usize = 9609;

const WM_CONVERT_DONE: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 140;

#[derive(Clone, Copy, PartialEq, Eq)]
enum AudioFormat {
    Mp3,
    Aac,
    Opus,
    Ogg,
    Flac,
    Wav,
    Aiff,
}

struct ConvertLabels {
    title: String,
    input: String,
    output: String,
    browse: String,
    format: String,
    format_mp3: String,
    format_aac: String,
    format_opus: String,
    format_ogg: String,
    format_flac: String,
    format_wav: String,
    format_aiff: String,
    quality_bitrate: String,
    quality_ogg: String,
    quality_flac: String,
    status_ready: String,
    status_running: String,
    status_done: String,
    button_convert: String,
    button_close: String,
    error_no_input: String,
    error_no_output: String,
    error_invalid_bitrate: String,
    error_same_path: String,
    error_failed: String,
    success: String,
    open_title: String,
    save_title: String,
}

struct ConvertWindowState {
    hwnd: HWND,
    parent: HWND,
    input_edit: HWND,
    output_edit: HWND,
    format_combo: HWND,
    quality_label: HWND,
    quality_edit_combo: HWND,
    quality_combo: HWND,
    convert_button: HWND,
    close_button: HWND,
    status_label: HWND,
    language: Language,
    status_dialog: HWND,
    cancel_flag: Option<Arc<AtomicBool>>,
    running: bool,
}

fn labels(language: Language) -> ConvertLabels {
    ConvertLabels {
        title: i18n::tr(language, "convert_audio.title"),
        input: i18n::tr(language, "convert_audio.input"),
        output: i18n::tr(language, "convert_audio.output"),
        browse: i18n::tr(language, "convert_audio.browse"),
        format: i18n::tr(language, "convert_audio.format"),
        format_mp3: i18n::tr(language, "convert_audio.format.mp3"),
        format_aac: i18n::tr(language, "convert_audio.format.aac"),
        format_opus: i18n::tr(language, "convert_audio.format.opus"),
        format_ogg: i18n::tr(language, "convert_audio.format.ogg"),
        format_flac: i18n::tr(language, "convert_audio.format.flac"),
        format_wav: i18n::tr(language, "convert_audio.format.wav"),
        format_aiff: i18n::tr(language, "convert_audio.format.aiff"),
        quality_bitrate: i18n::tr(language, "convert_audio.quality.bitrate"),
        quality_ogg: i18n::tr(language, "convert_audio.quality.ogg"),
        quality_flac: i18n::tr(language, "convert_audio.quality.flac"),
        status_ready: i18n::tr(language, "convert_audio.status.ready"),
        status_running: i18n::tr(language, "convert_audio.status.running"),
        status_done: i18n::tr(language, "convert_audio.status.done"),
        button_convert: i18n::tr(language, "convert_audio.button.convert"),
        button_close: i18n::tr(language, "convert_audio.button.close"),
        error_no_input: i18n::tr(language, "convert_audio.error.no_input"),
        error_no_output: i18n::tr(language, "convert_audio.error.no_output"),
        error_invalid_bitrate: i18n::tr(language, "convert_audio.error.invalid_bitrate"),
        error_same_path: i18n::tr(language, "convert_audio.error.same_path"),
        error_failed: i18n::tr(language, "convert_audio.error.failed"),
        success: i18n::tr(language, "convert_audio.success"),
        open_title: i18n::tr(language, "convert_audio.open_title"),
        save_title: i18n::tr(language, "convert_audio.save_title"),
    }
}

pub fn handle_navigation(hwnd: HWND, msg: &windows::Win32::UI::WindowsAndMessaging::MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
        if let Err(e) = unsafe { PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) } {
            crate::log_debug(&format!("Failed to post WM_CLOSE: {}", e));
        }
        return true;
    }
    false
}

pub fn open(parent: HWND) {
    let existing = { with_state(parent, |state| state.convert_audio_window).unwrap_or(HWND(0)) };
    if existing.0 != 0 {
        if unsafe { !IsWindow(existing).as_bool() } {
            let result = {
                with_state(parent, |state| {
                    state.convert_audio_window = HWND(0);
                })
            };
            if result.is_none() {
                crate::log_debug("Failed to access convert window state");
            }
        } else {
            unsafe {
                SetForegroundWindow(existing);
            }
            return;
        }
    }

    let language = { with_state(parent, |state| state.settings.language) }.unwrap_or_default();
    let class_name = to_wide(CONVERT_CLASS_NAME);
    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(convert_wndproc),
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
            140,
            140,
            520,
            260,
            parent,
            HMENU(0),
            hinstance,
            None,
        )
    };
    if hwnd.0 == 0 {
        return;
    }

    if {
        with_state(parent, |state| {
            state.convert_audio_window = hwnd;
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access convert window state");
    }
}
unsafe extern "system" fn convert_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "convert_wndproc",
        || DefWindowProcW(hwnd, msg, wparam, lparam),
        || convert_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn convert_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let cs = &*(lparam.0 as *const CREATESTRUCTW);
                let parent = cs.hwndParent;
                let language =
                    with_state(parent, |state| state.settings.language).unwrap_or_default();
                let labels = labels(language);
                let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

                let input_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.input).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    18,
                    120,
                    18,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let input_edit = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_EDIT,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    16,
                    38,
                    380,
                    24,
                    hwnd,
                    HMENU(CONVERT_ID_INPUT_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                let input_browse = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.browse).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    404,
                    38,
                    90,
                    24,
                    hwnd,
                    HMENU(CONVERT_ID_INPUT_BROWSE as isize),
                    HINSTANCE(0),
                    None,
                );

                let output_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.output).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    70,
                    120,
                    18,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let output_edit = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_EDIT,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    16,
                    90,
                    380,
                    24,
                    hwnd,
                    HMENU(CONVERT_ID_OUTPUT_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                let output_browse = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.browse).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    404,
                    90,
                    90,
                    24,
                    hwnd,
                    HMENU(CONVERT_ID_OUTPUT_BROWSE as isize),
                    HINSTANCE(0),
                    None,
                );

                let format_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.format).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    122,
                    120,
                    18,
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
                    140,
                    118,
                    160,
                    200,
                    hwnd,
                    HMENU(CONVERT_ID_FORMAT as isize),
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
                    LPARAM(to_wide(&labels.format_aac).as_ptr() as isize),
                );
                SendMessageW(
                    format_combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&labels.format_opus).as_ptr() as isize),
                );
                SendMessageW(
                    format_combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&labels.format_ogg).as_ptr() as isize),
                );
                SendMessageW(
                    format_combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&labels.format_flac).as_ptr() as isize),
                );
                SendMessageW(
                    format_combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&labels.format_wav).as_ptr() as isize),
                );
                SendMessageW(
                    format_combo,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&labels.format_aiff).as_ptr() as isize),
                );
                SendMessageW(format_combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));

                let quality_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.quality_bitrate).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    154,
                    160,
                    18,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let quality_edit_combo = CreateWindowExW(
                    Default::default(),
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWN as u32),
                    180,
                    150,
                    120,
                    200,
                    hwnd,
                    HMENU(CONVERT_ID_QUALITY_EDIT as isize),
                    HINSTANCE(0),
                    None,
                );
                let quality_combo = CreateWindowExW(
                    Default::default(),
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    180,
                    150,
                    120,
                    200,
                    hwnd,
                    HMENU(CONVERT_ID_QUALITY_COMBO as isize),
                    HINSTANCE(0),
                    None,
                );

                let status_label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&labels.status_ready).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    186,
                    300,
                    18,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let convert_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.button_convert).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    300,
                    180,
                    96,
                    28,
                    hwnd,
                    HMENU(CONVERT_ID_CONVERT as isize),
                    HINSTANCE(0),
                    None,
                );
                let close_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&labels.button_close).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    402,
                    180,
                    92,
                    28,
                    hwnd,
                    HMENU(CONVERT_ID_CLOSE as isize),
                    HINSTANCE(0),
                    None,
                );

                for control in [
                    input_label,
                    input_edit,
                    input_browse,
                    output_label,
                    output_edit,
                    output_browse,
                    format_label,
                    format_combo,
                    quality_label,
                    quality_edit_combo,
                    quality_combo,
                    status_label,
                    convert_button,
                    close_button,
                ] {
                    if control.0 != 0 && hfont.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    }
                }

                let mut state = Box::new(ConvertWindowState {
                    hwnd,
                    parent,
                    input_edit,
                    output_edit,
                    format_combo,
                    quality_label,
                    quality_edit_combo,
                    quality_combo,
                    convert_button,
                    close_button,
                    status_label,
                    language,
                    status_dialog: HWND(0),
                    cancel_flag: None,
                    running: false,
                });

                update_quality_controls(&mut state, AudioFormat::Mp3, &labels);
                ShowWindow(
                    state.quality_combo,
                    windows::Win32::UI::WindowsAndMessaging::SW_HIDE,
                );
                if let Some(path) = current_media_path(parent) {
                    set_edit_text(state.input_edit, &path);
                    if get_edit_text(state.output_edit).is_empty()
                        && let Some(suggested) =
                            build_default_output_path(&path, current_format(state.format_combo))
                    {
                        set_edit_text(state.output_edit, &suggested);
                    }
                }

                SetWindowLongPtrW(
                    hwnd,
                    windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                    Box::into_raw(state) as isize,
                );
                SetFocus(input_edit);
                LRESULT(0)
            }
            WM_COMMAND => {
                let cmd_id = wparam.0 & 0xFFFF;
                let notify = (wparam.0 >> 16) as u16;
                let ptr =
                    GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                        as *mut ConvertWindowState;
                if ptr.is_null() {
                    return LRESULT(0);
                }
                let state = &mut *ptr;
                let labels = labels(state.language);

                match cmd_id {
                    CONVERT_ID_INPUT_BROWSE => {
                        if state.running {
                            return LRESULT(0);
                        }
                        if let Some(path) = open_input_dialog(state.hwnd, state.language, &labels) {
                            set_edit_text(state.input_edit, &path);
                            if get_edit_text(state.output_edit).is_empty()
                                && let Some(suggested) = build_default_output_path(
                                    &path,
                                    current_format(state.format_combo),
                                )
                            {
                                set_edit_text(state.output_edit, &suggested);
                            }
                            SetForegroundWindow(state.hwnd);
                            SetFocus(state.input_edit);
                        }
                    }
                    CONVERT_ID_OUTPUT_BROWSE => {
                        if state.running {
                            return LRESULT(0);
                        }
                        let format = current_format(state.format_combo);
                        let initial = get_edit_text(state.output_edit);
                        let initial_path = if initial.is_empty() {
                            None
                        } else {
                            Some(PathBuf::from(initial))
                        };
                        if let Some(path) = open_output_dialog(
                            state.hwnd,
                            state.language,
                            &labels,
                            format,
                            initial_path.as_ref(),
                        ) {
                            set_edit_text(state.output_edit, &path);
                            SetForegroundWindow(state.hwnd);
                            SetFocus(state.output_edit);
                        }
                    }
                    CONVERT_ID_FORMAT => {
                        if notify as u32 == windows::Win32::UI::WindowsAndMessaging::CBN_SELCHANGE {
                            let format = current_format(state.format_combo);
                            update_quality_controls(state, format, &labels);
                            let current_output = get_edit_text(state.output_edit);
                            if !current_output.is_empty() {
                                let mut path = PathBuf::from(current_output);
                                path.set_extension(extension_for_format(format));
                                set_edit_text(state.output_edit, &path);
                            }
                        }
                    }
                    CONVERT_ID_CONVERT => {
                        if state.running {
                            return LRESULT(0);
                        }
                        let input_text = get_edit_text(state.input_edit);
                        if input_text.is_empty() {
                            show_error(state.parent, state.language, &labels.error_no_input);
                            return LRESULT(0);
                        }
                        let output_text = get_edit_text(state.output_edit);
                        if output_text.is_empty() {
                            show_error(state.parent, state.language, &labels.error_no_output);
                            return LRESULT(0);
                        }
                        if input_text == output_text {
                            show_error(state.parent, state.language, &labels.error_same_path);
                            return LRESULT(0);
                        }

                        let format = current_format(state.format_combo);
                        let quality = match read_quality(state, format, &labels) {
                            Ok(q) => q,
                            Err(err) => {
                                show_error(state.parent, state.language, &err);
                                return LRESULT(0);
                            }
                        };

                        let settings = ConvertAudioSettings {
                            format: map_format(format),
                            quality,
                        };

                        let args = build_ffmpeg_args(&settings);
                        log_debug(&format!("Convert audio args: {}", args.join(" ")));

                        state.running = true;
                        set_status(state, &labels.status_running);
                        set_controls_enabled(state, false);
                        let cancel_flag = Arc::new(AtomicBool::new(false));
                        state.cancel_flag = Some(cancel_flag.clone());
                        if state.status_dialog.0 == 0 {
                            let dialog_labels = SaveDialogLabels {
                                title: labels.title.clone(),
                                in_progress: labels.status_running.clone(),
                                cancel: i18n::tr(state.language, "podcast.save.cancel"),
                            };
                            let dialog = podcast_save_window::open_with_labels(
                                state.hwnd,
                                state.language,
                                dialog_labels,
                                true,
                            );
                            if dialog.0 != 0 {
                                state.status_dialog = dialog;
                            }
                        }

                        let input = PathBuf::from(input_text);
                        let output = PathBuf::from(output_text);
                        let hwnd_target = state.hwnd;

                        std::thread::spawn(move || {
                            let result = convert_audio_file(
                                &input,
                                &output,
                                &settings,
                                Some(cancel_flag),
                                None,
                            )
                            .map(|_| output);
                            let boxed = Box::new(result);
                            if IsWindow(hwnd_target).as_bool() {
                                let raw = Box::into_raw(boxed);
                                if let Err(err) = PostMessageW(
                                    hwnd_target,
                                    WM_CONVERT_DONE,
                                    WPARAM(0),
                                    LPARAM(raw as isize),
                                ) {
                                    log_debug(&format!("Failed to post convert result: {}", err));
                                    let _cleanup = Box::from_raw(raw);
                                }
                            } else {
                                let _boxed = boxed;
                            }
                        });
                    }
                    CONVERT_ID_CLOSE => {
                        if let Err(e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                            log_debug(&format!("Failed to post WM_CLOSE: {}", e));
                        }
                    }
                    _ => {}
                }
                LRESULT(0)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    if let Err(e) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                        log_debug(&format!("Failed to post WM_CLOSE: {}", e));
                    }
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CONVERT_DONE => {
                let ptr =
                    GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                        as *mut ConvertWindowState;
                if ptr.is_null() {
                    return LRESULT(0);
                }
                let state = &mut *ptr;
                let labels = labels(state.language);

                let boxed = Box::from_raw(lparam.0 as *mut Result<PathBuf, String>);
                state.running = false;
                set_controls_enabled(state, true);
                state.cancel_flag = None;
                if state.status_dialog.0 != 0 {
                    let dialog = state.status_dialog;
                    state.status_dialog = HWND(0);
                    if IsWindow(dialog).as_bool()
                        && let Err(e) = PostMessageW(
                            dialog,
                            podcast_save_window::WM_PODCAST_SAVE_DONE,
                            WPARAM(0),
                            LPARAM(0),
                        )
                    {
                        log_debug(&format!("Failed to close status dialog: {}", e));
                    }
                }

                match *boxed {
                    Ok(path) => {
                        set_status(state, &labels.status_done);
                        set_edit_text(state.output_edit, &path);
                        show_info(state.parent, state.language, &labels.success);
                    }
                    Err(err) => {
                        if err == "Conversion canceled." {
                            log_debug("Convert audio canceled by user");
                            set_status(state, &labels.status_ready);
                            show_info(
                                state.parent,
                                state.language,
                                &i18n::tr(state.language, "tts.cancelled"),
                            );
                        } else {
                            log_debug(&format!("Convert audio failed: {}", err));
                            set_status(state, &labels.status_ready);
                            show_error(state.parent, state.language, &labels.error_failed);
                        }
                    }
                }
                LRESULT(0)
            }
            podcast_save_window::WM_PODCAST_SAVE_CANCEL => {
                let ptr =
                    GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                        as *mut ConvertWindowState;
                if ptr.is_null() {
                    return LRESULT(0);
                }
                let state = &mut *ptr;
                if let Some(flag) = state.cancel_flag.as_ref() {
                    flag.store(true, Ordering::Relaxed);
                }
                LRESULT(0)
            }
            WM_CLOSE => {
                crate::log_if_err!(DestroyWindow(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => {
                let ptr =
                    GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                        as *mut ConvertWindowState;
                if !ptr.is_null() {
                    let state = &mut *ptr;
                    if state.status_dialog.0 != 0 {
                        crate::log_if_err!(DestroyWindow(state.status_dialog));
                        state.status_dialog = HWND(0);
                    }
                }
                let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
                if with_state(parent, |state| {
                    state.convert_audio_window = HWND(0);
                })
                .is_none()
                {
                    log_debug("Failed to update convert window state");
                }
                crate::focus_editor(parent);
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr =
                    GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                        as *mut ConvertWindowState;
                if !ptr.is_null() {
                    SetWindowLongPtrW(
                        hwnd,
                        windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                        0,
                    );
                    let _unused_box = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn current_format(combo: HWND) -> AudioFormat {
    let sel = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    match sel {
        1 => AudioFormat::Aac,
        2 => AudioFormat::Opus,
        3 => AudioFormat::Ogg,
        4 => AudioFormat::Flac,
        5 => AudioFormat::Wav,
        6 => AudioFormat::Aiff,
        _ => AudioFormat::Mp3,
    }
}

fn map_format(format: AudioFormat) -> ConvertAudioFormat {
    match format {
        AudioFormat::Mp3 => ConvertAudioFormat::Mp3,
        AudioFormat::Aac => ConvertAudioFormat::Aac,
        AudioFormat::Opus => ConvertAudioFormat::Opus,
        AudioFormat::Ogg => ConvertAudioFormat::Ogg,
        AudioFormat::Flac => ConvertAudioFormat::Flac,
        AudioFormat::Wav => ConvertAudioFormat::Wav,
        AudioFormat::Aiff => ConvertAudioFormat::Aiff,
    }
}

fn current_media_path(parent: HWND) -> Option<PathBuf> {
    {
        with_state(parent, |state| {
            if let Some(player) = &state.active_audiobook {
                return Some(player.path.clone());
            }
            if let Some(doc) = state.docs.get(state.current)
                && matches!(doc.format, FileFormat::Audiobook)
            {
                return doc.path.clone();
            }
            None
        })
        .unwrap_or(None)
    }
}

fn extension_for_format(format: AudioFormat) -> &'static str {
    match format {
        AudioFormat::Mp3 => "mp3",
        AudioFormat::Aac => "m4a",
        AudioFormat::Opus => "opus",
        AudioFormat::Ogg => "ogg",
        AudioFormat::Flac => "flac",
        AudioFormat::Wav => "wav",
        AudioFormat::Aiff => "aiff",
    }
}

fn build_default_output_path(input: &Path, format: AudioFormat) -> Option<PathBuf> {
    let parent = input.parent()?;
    let stem = input
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audio");
    let file_name = format!("{stem}_converted.{}", extension_for_format(format));
    Some(parent.join(file_name))
}

fn read_quality(
    state: &ConvertWindowState,
    format: AudioFormat,
    labels: &ConvertLabels,
) -> Result<ConvertAudioQuality, String> {
    match format {
        AudioFormat::Mp3 => {
            let text = get_combo_text(state.quality_edit_combo);
            let value = text
                .trim()
                .parse::<i32>()
                .map_err(|_| labels.error_invalid_bitrate.clone())?;
            validate_mp3_bitrate(value)
                .map(ConvertAudioQuality::BitrateKbps)
                .map_err(|_| labels.error_invalid_bitrate.clone())
        }
        AudioFormat::Aac | AudioFormat::Opus => {
            let text = get_combo_text(state.quality_combo);
            let value = text
                .trim()
                .parse::<u32>()
                .map_err(|_| labels.error_invalid_bitrate.clone())?;
            Ok(ConvertAudioQuality::BitrateKbps(value))
        }
        AudioFormat::Ogg => {
            let text = get_combo_text(state.quality_combo);
            let q = text.trim().trim_start_matches('q');
            let value = q
                .parse::<u8>()
                .map_err(|_| labels.error_invalid_bitrate.clone())?;
            Ok(ConvertAudioQuality::OggQuality(value))
        }
        AudioFormat::Flac => {
            let text = get_combo_text(state.quality_combo);
            let value = text
                .trim()
                .parse::<u8>()
                .map_err(|_| labels.error_invalid_bitrate.clone())?;
            Ok(ConvertAudioQuality::FlacCompression(value))
        }
        AudioFormat::Wav | AudioFormat::Aiff => Ok(ConvertAudioQuality::None),
    }
}

fn set_status(state: &ConvertWindowState, text: &str) {
    let wide = to_wide(text);
    if let Err(e) = unsafe { SetWindowTextW(state.status_label, PCWSTR(wide.as_ptr())) } {
        log_debug(&format!("Failed to set status text: {}", e));
    }
}

fn set_controls_enabled(state: &ConvertWindowState, enabled: bool) {
    unsafe {
        EnableWindow(state.input_edit, enabled);
        EnableWindow(state.output_edit, enabled);
        EnableWindow(state.format_combo, enabled);
        EnableWindow(state.quality_edit_combo, enabled);
        EnableWindow(state.quality_combo, enabled);
        EnableWindow(state.convert_button, enabled);
        EnableWindow(state.close_button, enabled);
    }
}

fn update_quality_controls(
    state: &mut ConvertWindowState,
    format: AudioFormat,
    labels: &ConvertLabels,
) {
    let (label_text, show_edit, show_combo) = match format {
        AudioFormat::Mp3 => (&labels.quality_bitrate, true, false),
        AudioFormat::Aac | AudioFormat::Opus => (&labels.quality_bitrate, false, true),
        AudioFormat::Ogg => (&labels.quality_ogg, false, true),
        AudioFormat::Flac => (&labels.quality_flac, false, true),
        AudioFormat::Wav | AudioFormat::Aiff => (&labels.quality_bitrate, false, false),
    };

    let label_wide = to_wide(label_text);
    if let Err(e) = unsafe { SetWindowTextW(state.quality_label, PCWSTR(label_wide.as_ptr())) } {
        log_debug(&format!("Failed to set quality label: {}", e));
    }

    unsafe {
        ShowWindow(
            state.quality_label,
            if show_edit || show_combo {
                windows::Win32::UI::WindowsAndMessaging::SW_SHOW
            } else {
                windows::Win32::UI::WindowsAndMessaging::SW_HIDE
            },
        );
        ShowWindow(
            state.quality_edit_combo,
            if show_edit {
                windows::Win32::UI::WindowsAndMessaging::SW_SHOW
            } else {
                windows::Win32::UI::WindowsAndMessaging::SW_HIDE
            },
        );
        ShowWindow(
            state.quality_combo,
            if show_combo {
                windows::Win32::UI::WindowsAndMessaging::SW_SHOW
            } else {
                windows::Win32::UI::WindowsAndMessaging::SW_HIDE
            },
        );
    }

    if show_edit {
        set_combo_items(
            state.quality_edit_combo,
            &[
                "64", "80", "96", "112", "128", "160", "192", "224", "256", "320",
            ],
        );
        if get_combo_text(state.quality_edit_combo).is_empty() {
            set_combo_text(state.quality_edit_combo, "192");
        }
    }
    if show_combo {
        let items: Vec<&str> = match format {
            AudioFormat::Aac => vec![
                "64", "80", "96", "112", "128", "160", "192", "224", "256", "320",
            ],
            AudioFormat::Opus => vec!["64", "96", "128", "160"],
            AudioFormat::Ogg => vec!["q3", "q4", "q5", "q6", "q7", "q8"],
            AudioFormat::Flac => vec!["0", "1", "2", "3", "4", "5", "6", "7", "8"],
            _ => Vec::new(),
        };
        set_combo_items(state.quality_combo, &items);
        let default_index = match format {
            AudioFormat::Aac => 4,
            AudioFormat::Opus => 2,
            AudioFormat::Ogg => 2,
            AudioFormat::Flac => 5,
            _ => 0,
        };
        unsafe {
            SendMessageW(
                state.quality_combo,
                CB_SETCURSEL,
                WPARAM(default_index),
                LPARAM(0),
            );
        }
    }
}

fn set_combo_items(hwnd: HWND, items: &[&str]) {
    unsafe {
        SendMessageW(hwnd, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for item in items {
            SendMessageW(
                hwnd,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(to_wide(item).as_ptr() as isize),
            );
        }
    }
}

fn set_combo_text(hwnd: HWND, text: &str) {
    let wide = to_wide(text);
    if let Err(e) = unsafe { SetWindowTextW(hwnd, PCWSTR(wide.as_ptr())) } {
        log_debug(&format!("Failed to set combo text: {}", e));
    }
}

fn get_combo_text(hwnd: HWND) -> String {
    let len = unsafe { GetWindowTextLengthW(hwnd) } as usize;
    if len == 0 {
        return String::new();
    }
    let mut buf = vec![0u16; len + 1];
    let read = unsafe { GetWindowTextW(hwnd, &mut buf) } as usize;
    let used = read.min(len);
    String::from_utf16_lossy(&buf[..used])
}

fn get_edit_text(hwnd: HWND) -> String {
    let len = unsafe { GetWindowTextLengthW(hwnd) } as usize;
    if len == 0 {
        return String::new();
    }
    let mut buf = vec![0u16; len + 1];
    let read = unsafe { GetWindowTextW(hwnd, &mut buf) } as usize;
    let used = read.min(len);
    String::from_utf16_lossy(&buf[..used])
}

fn set_edit_text(hwnd: HWND, path: &Path) {
    let wide = to_wide(&path.to_string_lossy());
    if let Err(e) = unsafe { SetWindowTextW(hwnd, PCWSTR(wide.as_ptr())) } {
        log_debug(&format!("Failed to set edit text: {}", e));
    }
}

fn open_input_dialog(parent: HWND, language: Language, labels: &ConvertLabels) -> Option<PathBuf> {
    unsafe {
        let filter_raw = i18n::tr(language, "podcasts.download_filter");
        let filter = to_wide(&filter_raw.replace("\\0", "\0"));
        let title = to_wide(&labels.open_title);
        let mut buffer = [0u16; 1024];
        let mut ofn = OPENFILENAMEW {
            lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
            hwndOwner: parent,
            lpstrFile: PWSTR(buffer.as_mut_ptr()),
            nMaxFile: buffer.len() as u32,
            lpstrFilter: PCWSTR(filter.as_ptr()),
            lpstrTitle: PCWSTR(title.as_ptr()),
            Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
            ..Default::default()
        };
        if !GetOpenFileNameW(&mut ofn).as_bool() {
            return None;
        }
        let len = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
        if len == 0 {
            return None;
        }
        Some(PathBuf::from(String::from_utf16_lossy(&buffer[..len])))
    }
}
fn open_output_dialog(
    parent: HWND,
    language: Language,
    labels: &ConvertLabels,
    format: AudioFormat,
    initial: Option<&PathBuf>,
) -> Option<PathBuf> {
    unsafe {
        let all_label = i18n::tr(language, "dialog.all_files");
        let filter = format!("{} (*.*)\0*.*\0\0", all_label);
        let filter_wide = to_wide(&filter);
        let title = to_wide(&labels.save_title);
        let mut buffer = [0u16; 1024];
        if let Some(path) = initial {
            let path_w = to_wide(&path.to_string_lossy());
            let copy_len = path_w.len().min(buffer.len() - 1);
            buffer[..copy_len].copy_from_slice(&path_w[..copy_len]);
        }
        let ext_wide = to_wide(extension_for_format(format));
        let mut ofn = OPENFILENAMEW {
            lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
            hwndOwner: parent,
            lpstrFile: PWSTR(buffer.as_mut_ptr()),
            nMaxFile: buffer.len() as u32,
            lpstrFilter: PCWSTR(filter_wide.as_ptr()),
            lpstrTitle: PCWSTR(title.as_ptr()),
            lpstrDefExt: PCWSTR(ext_wide.as_ptr()),
            Flags: OFN_EXPLORER | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST,
            ..Default::default()
        };
        if !GetSaveFileNameW(&mut ofn).as_bool() {
            return None;
        }
        let len = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
        if len == 0 {
            return None;
        }
        let mut path = PathBuf::from(String::from_utf16_lossy(&buffer[..len]));
        if path.extension().is_none() {
            path.set_extension(extension_for_format(format));
        }
        Some(path)
    }
}
