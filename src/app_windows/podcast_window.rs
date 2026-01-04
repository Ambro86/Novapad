use crate::accessibility::{handle_accessibility, to_wide};
use crate::app_windows::podcast_save_window;
use crate::podcast_recorder::{
    AudioDevice, RecorderConfig, RecorderHandle, RecorderStatus, default_output_folder,
    list_input_devices, list_output_devices, probe_device, start_recording,
};
use crate::settings::{
    AppSettings, Language, PODCAST_DEVICE_DEFAULT, PodcastFormat, confirm_title,
    default_podcast_save_folder, save_settings,
};
use crate::{i18n, show_error, with_state};
use std::path::PathBuf;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    BST_CHECKED, BST_UNCHECKED, WC_BUTTON, WC_COMBOBOXW, WC_EDIT, WC_STATIC,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetKeyState, SetFocus, VK_CONTROL, VK_ESCAPE, VK_SHIFT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BM_CLICK, BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_DEFPUSHBUTTON, BS_GROUPBOX,
    CB_ADDSTRING, CB_GETCURSEL, CB_RESETCONTENT, CB_SETCURSEL, CBN_SELCHANGE, CBS_DROPDOWNLIST,
    CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow, EN_CHANGE,
    GWLP_USERDATA, GetDlgItem, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU,
    IDC_ARROW, IDOK, LoadCursorW, MB_ICONERROR, MB_ICONINFORMATION, MB_ICONWARNING, MB_OK,
    MB_OKCANCEL, MSG, MessageBoxW, PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow,
    SetTimer, SetWindowLongPtrW, SetWindowTextW, ShowWindow, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WM_TIMER, WNDCLASSW, WS_CAPTION,
    WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP,
    WS_VISIBLE,
};
use windows::core::PCWSTR;

const PODCAST_CLASS_NAME: &str = "NovapadPodcast";
const PODCAST_TIMER_ID: usize = 1;

const PODCAST_ID_INCLUDE_MIC: usize = 11001;
const PODCAST_ID_MIC_DEVICE: usize = 11002;
const PODCAST_ID_INCLUDE_SYSTEM: usize = 11003;
const PODCAST_ID_SYSTEM_DEVICE: usize = 11004;
const PODCAST_ID_FORMAT: usize = 11005;
const PODCAST_ID_BITRATE: usize = 11006;
const PODCAST_ID_SAVE_PATH: usize = 11007;
const PODCAST_ID_BROWSE: usize = 11008;
const PODCAST_ID_FILENAME_PREVIEW: usize = 11009;
const PODCAST_ID_START: usize = 11010;
const PODCAST_ID_PAUSE: usize = 11011;
const PODCAST_ID_RESUME: usize = 11012;
const PODCAST_ID_STOP: usize = 11013;
const PODCAST_ID_CLOSE: usize = 11014;
const PODCAST_ID_STATUS: usize = 11015;
const PODCAST_ID_ELAPSED: usize = 11016;
const PODCAST_ID_LEVEL_MIC: usize = 11017;
const PODCAST_ID_LEVEL_SYSTEM: usize = 11018;
const PODCAST_ID_HINT: usize = 11019;
const PODCAST_ID_SYSTEM_UNAVAILABLE: usize = 11020;
const WM_PODCAST_SAVE_RESULT: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 74;

struct PodcastSaveResult {
    success: bool,
    message: String,
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN {
        let ctrl = (GetKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000) != 0;
        let shift = (GetKeyState(VK_SHIFT.0 as i32) as u16 & 0x8000) != 0;
        let key = msg.wParam.0 as u32;

        if key == VK_ESCAPE.0 as u32 {
            let _ = SendMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
            return true;
        }
        if ctrl && !shift {
            if key == 'S' as u32 {
                click_button(hwnd, PODCAST_ID_START);
                return true;
            }
            if key == 'P' as u32 {
                click_button(hwnd, PODCAST_ID_PAUSE);
                return true;
            }
            if key == 'T' as u32 {
                click_button(hwnd, PODCAST_ID_STOP);
                return true;
            }
        }
    }
    handle_accessibility(hwnd, msg)
}

struct PodcastState {
    parent: HWND,
    language: Language,
    include_mic: HWND,
    mic_device: HWND,
    include_system: HWND,
    system_device: HWND,
    format_combo: HWND,
    bitrate_combo: HWND,
    save_path: HWND,
    filename_preview: HWND,
    start_button: HWND,
    pause_button: HWND,
    resume_button: HWND,
    stop_button: HWND,
    status_text: HWND,
    elapsed_text: HWND,
    level_mic_text: HWND,
    level_system_text: HWND,
    hint_text: HWND,
    system_unavailable_text: HWND,
    mic_devices: Vec<AudioDevice>,
    system_devices: Vec<AudioDevice>,
    recorder: Option<RecorderHandle>,
    system_available: bool,
    saving_dialog: HWND,
    save_cancel: Option<Arc<AtomicBool>>,
}

struct PodcastLabels {
    title: String,
    input_group: String,
    output_group: String,
    controls_group: String,
    status_group: String,
    include_mic: String,
    mic_device: String,
    include_system: String,
    system_device: String,
    system_unavailable: String,
    format: String,
    bitrate: String,
    save_path: String,
    browse: String,
    filename: String,
    start: String,
    pause: String,
    resume: String,
    stop: String,
    close: String,
    status_label: String,
    elapsed_label: String,
    level_mic: String,
    level_system: String,
    status_idle: String,
    status_recording: String,
    status_paused: String,
    status_saving: String,
    status_error: String,
    default_device: String,
    hint_select_source: String,
    confirm_close_recording: String,
    error_system_audio: String,
    error_microphone: String,
}
fn labels(language: Language) -> PodcastLabels {
    PodcastLabels {
        title: i18n::tr(language, "podcast.title"),
        input_group: i18n::tr(language, "podcast.group.input"),
        output_group: i18n::tr(language, "podcast.group.output"),
        controls_group: i18n::tr(language, "podcast.group.controls"),
        status_group: i18n::tr(language, "podcast.group.status"),
        include_mic: i18n::tr(language, "podcast.include_mic"),
        mic_device: i18n::tr(language, "podcast.mic_device"),
        include_system: i18n::tr(language, "podcast.include_system"),
        system_device: i18n::tr(language, "podcast.system_device"),
        system_unavailable: i18n::tr(language, "podcast.system_unavailable"),
        format: i18n::tr(language, "podcast.format"),
        bitrate: i18n::tr(language, "podcast.bitrate"),
        save_path: i18n::tr(language, "podcast.save_path"),
        browse: i18n::tr(language, "podcast.browse"),
        filename: i18n::tr(language, "podcast.filename_preview"),
        start: i18n::tr(language, "podcast.start"),
        pause: i18n::tr(language, "podcast.pause"),
        resume: i18n::tr(language, "podcast.resume"),
        stop: i18n::tr(language, "podcast.stop"),
        close: i18n::tr(language, "podcast.close"),
        status_label: i18n::tr(language, "podcast.status_label"),
        elapsed_label: i18n::tr(language, "podcast.elapsed_label"),
        level_mic: i18n::tr(language, "podcast.level_mic"),
        level_system: i18n::tr(language, "podcast.level_system"),
        status_idle: i18n::tr(language, "podcast.status.idle"),
        status_recording: i18n::tr(language, "podcast.status.recording"),
        status_paused: i18n::tr(language, "podcast.status.paused"),
        status_saving: i18n::tr(language, "podcast.status.saving"),
        status_error: i18n::tr(language, "podcast.status.error"),
        default_device: i18n::tr(language, "podcast.device.default"),
        hint_select_source: i18n::tr(language, "podcast.hint.select_source"),
        confirm_close_recording: i18n::tr(language, "podcast.confirm_close_recording"),
        error_system_audio: i18n::tr(language, "podcast.error.system_audio"),
        error_microphone: i18n::tr(language, "podcast.error.microphone"),
    }
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.podcast_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(PODCAST_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(podcast_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let title = to_wide(&labels(language).title);

    let window = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        640,
        620,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if window.0 != 0 {
        let _ = with_state(parent, |state| {
            state.podcast_window = window;
        });
        EnableWindow(parent, false);
        SetForegroundWindow(window);
    }
}
unsafe extern "system" fn podcast_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = labels(language);
            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

            let group_input = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.input_group).as_ptr()),
                WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_GROUPBOX as u32),
                10,
                10,
                600,
                140,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let include_mic = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.include_mic).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                20,
                35,
                220,
                22,
                hwnd,
                HMENU(PODCAST_ID_INCLUDE_MIC as isize),
                None,
                None,
            );

            let label_mic = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.mic_device).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                40,
                62,
                180,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let mic_device = CreateWindowExW(
                Default::default(),
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                230,
                58,
                350,
                200,
                hwnd,
                HMENU(PODCAST_ID_MIC_DEVICE as isize),
                None,
                None,
            );

            let include_system = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.include_system).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                20,
                88,
                220,
                22,
                hwnd,
                HMENU(PODCAST_ID_INCLUDE_SYSTEM as isize),
                None,
                None,
            );

            let label_system = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.system_device).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                40,
                115,
                180,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let system_device = CreateWindowExW(
                Default::default(),
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                230,
                111,
                350,
                200,
                hwnd,
                HMENU(PODCAST_ID_SYSTEM_DEVICE as isize),
                None,
                None,
            );

            let system_unavailable_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.system_unavailable).as_ptr()),
                WS_CHILD,
                40,
                135,
                540,
                18,
                hwnd,
                HMENU(PODCAST_ID_SYSTEM_UNAVAILABLE as isize),
                None,
                None,
            );

            let group_output = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.output_group).as_ptr()),
                WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_GROUPBOX as u32),
                10,
                160,
                600,
                170,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let label_format = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.format).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                190,
                100,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let format_combo = CreateWindowExW(
                Default::default(),
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                140,
                186,
                140,
                200,
                hwnd,
                HMENU(PODCAST_ID_FORMAT as isize),
                None,
                None,
            );

            let label_bitrate = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.bitrate).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                300,
                190,
                100,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let bitrate_combo = CreateWindowExW(
                Default::default(),
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                420,
                186,
                140,
                200,
                hwnd,
                HMENU(PODCAST_ID_BITRATE as isize),
                None,
                None,
            );

            let label_save = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.save_path).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                220,
                220,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let save_path = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                20,
                242,
                430,
                24,
                hwnd,
                HMENU(PODCAST_ID_SAVE_PATH as isize),
                None,
                None,
            );

            let browse_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.browse).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                460,
                240,
                100,
                26,
                hwnd,
                HMENU(PODCAST_ID_BROWSE as isize),
                None,
                None,
            );

            let label_filename = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.filename).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                272,
                220,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let filename_preview = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                20,
                294,
                540,
                18,
                hwnd,
                HMENU(PODCAST_ID_FILENAME_PREVIEW as isize),
                None,
                None,
            );

            let group_controls = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.controls_group).as_ptr()),
                WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_GROUPBOX as u32),
                10,
                340,
                600,
                100,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let start_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.start).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                20,
                365,
                90,
                28,
                hwnd,
                HMENU(PODCAST_ID_START as isize),
                None,
                None,
            );

            let pause_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.pause).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                120,
                365,
                90,
                28,
                hwnd,
                HMENU(PODCAST_ID_PAUSE as isize),
                None,
                None,
            );

            let resume_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.resume).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                220,
                365,
                90,
                28,
                hwnd,
                HMENU(PODCAST_ID_RESUME as isize),
                None,
                None,
            );

            let stop_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.stop).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                320,
                365,
                110,
                28,
                hwnd,
                HMENU(PODCAST_ID_STOP as isize),
                None,
                None,
            );

            let close_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.close).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                440,
                365,
                90,
                28,
                hwnd,
                HMENU(PODCAST_ID_CLOSE as isize),
                None,
                None,
            );

            let group_status = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.status_group).as_ptr()),
                WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_GROUPBOX as u32),
                10,
                450,
                600,
                130,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let status_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.status_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                475,
                80,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let status_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.status_idle).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                110,
                475,
                180,
                18,
                hwnd,
                HMENU(PODCAST_ID_STATUS as isize),
                None,
                None,
            );

            let elapsed_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.elapsed_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                300,
                475,
                120,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let elapsed_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide("00:00:00").as_ptr()),
                WS_CHILD | WS_VISIBLE,
                430,
                475,
                120,
                18,
                hwnd,
                HMENU(PODCAST_ID_ELAPSED as isize),
                None,
                None,
            );

            let level_mic_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.level_mic).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20,
                505,
                150,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let level_mic_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide("0").as_ptr()),
                WS_CHILD | WS_VISIBLE,
                180,
                505,
                80,
                18,
                hwnd,
                HMENU(PODCAST_ID_LEVEL_MIC as isize),
                None,
                None,
            );

            let level_system_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.level_system).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                300,
                505,
                150,
                18,
                hwnd,
                HMENU(0),
                None,
                None,
            );

            let level_system_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide("0").as_ptr()),
                WS_CHILD | WS_VISIBLE,
                460,
                505,
                80,
                18,
                hwnd,
                HMENU(PODCAST_ID_LEVEL_SYSTEM as isize),
                None,
                None,
            );

            let hint_text = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.hint_select_source).as_ptr()),
                WS_CHILD,
                20,
                528,
                540,
                18,
                hwnd,
                HMENU(PODCAST_ID_HINT as isize),
                None,
                None,
            );

            let controls = [
                group_input,
                include_mic,
                label_mic,
                mic_device,
                include_system,
                label_system,
                system_device,
                system_unavailable_text,
                group_output,
                label_format,
                format_combo,
                label_bitrate,
                bitrate_combo,
                label_save,
                save_path,
                browse_button,
                label_filename,
                filename_preview,
                group_controls,
                start_button,
                pause_button,
                resume_button,
                stop_button,
                close_button,
                group_status,
                status_label,
                status_text,
                elapsed_label,
                elapsed_text,
                level_mic_label,
                level_mic_text,
                level_system_label,
                level_system_text,
                hint_text,
            ];
            for control in controls {
                let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }

            populate_combos(
                format_combo,
                bitrate_combo,
                mic_device,
                system_device,
                language,
            );

            let (mic_devices, system_devices, system_available) = load_devices(language);
            let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
            let mut state = PodcastState {
                parent,
                language,
                include_mic,
                mic_device,
                include_system,
                system_device,
                format_combo,
                bitrate_combo,
                save_path,
                filename_preview,
                start_button,
                pause_button,
                resume_button,
                stop_button,
                status_text,
                elapsed_text,
                level_mic_text,
                level_system_text,
                hint_text,
                system_unavailable_text,
                mic_devices,
                system_devices,
                recorder: None,
                system_available,
                saving_dialog: HWND(0),
                save_cancel: None,
            };

            apply_settings_to_ui(&mut state, &settings);
            update_source_controls(&state);
            update_format_controls(&state);
            update_filename_preview(&state);
            update_recording_controls(&state);
            update_status_text(&state, RecorderStatus::Idle);

            let boxed = Box::new(state);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(boxed) as isize);

            let _ = SetTimer(hwnd, PODCAST_TIMER_ID, 500, None);
            SetFocus(include_mic);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            let code = (wparam.0 >> 16) as u16;
            let mut handled = false;
            with_podcast_state(hwnd, |state| match id {
                PODCAST_ID_INCLUDE_MIC | PODCAST_ID_INCLUDE_SYSTEM => {
                    update_source_controls(state);
                    update_recording_controls(state);
                    persist_settings(state);
                    handled = true;
                }
                PODCAST_ID_MIC_DEVICE | PODCAST_ID_SYSTEM_DEVICE => {
                    if code == CBN_SELCHANGE as u16 {
                        persist_settings(state);
                    }
                    handled = true;
                }
                PODCAST_ID_FORMAT => {
                    if code == CBN_SELCHANGE as u16 {
                        update_format_controls(state);
                        update_filename_preview(state);
                        persist_settings(state);
                    }
                    handled = true;
                }
                PODCAST_ID_BITRATE => {
                    if code == CBN_SELCHANGE as u16 {
                        persist_settings(state);
                    }
                    handled = true;
                }
                PODCAST_ID_SAVE_PATH => {
                    if code == EN_CHANGE as u16 {
                        update_filename_preview(state);
                        persist_settings(state);
                    }
                    handled = true;
                }
                PODCAST_ID_BROWSE => {
                    if let Some(folder) = browse_for_folder(hwnd, state.language) {
                        let path = folder.to_string_lossy().to_string();
                        let wide = to_wide(&path);
                        let _ = SetWindowTextW(state.save_path, PCWSTR(wide.as_ptr()));
                        update_filename_preview(state);
                        persist_settings(state);
                    }
                    handled = true;
                }
                PODCAST_ID_START => {
                    handled = true;
                    start_recording_action(state);
                }
                PODCAST_ID_PAUSE => {
                    if let Some(recorder) = state.recorder.as_ref() {
                        if recorder.status() == RecorderStatus::Recording {
                            recorder.pause();
                            update_recording_controls(state);
                            update_status_text(state, RecorderStatus::Paused);
                        } else if recorder.status() == RecorderStatus::Paused {
                            recorder.resume();
                            update_recording_controls(state);
                            update_status_text(state, RecorderStatus::Recording);
                        }
                    }
                    handled = true;
                }
                PODCAST_ID_RESUME => {
                    if let Some(recorder) = state.recorder.as_ref() {
                        recorder.resume();
                        update_recording_controls(state);
                        update_status_text(state, RecorderStatus::Recording);
                    }
                    handled = true;
                }
                PODCAST_ID_STOP => {
                    handled = true;
                    stop_recording_action(state, hwnd);
                }
                PODCAST_ID_CLOSE => {
                    let _ = SendMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                    handled = true;
                }
                _ => {}
            });
            if handled {
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_TIMER => {
            if wparam.0 == PODCAST_TIMER_ID {
                with_podcast_state(hwnd, |state| {
                    update_status_from_recorder(state);
                });
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        podcast_save_window::WM_PODCAST_SAVE_CLOSED => {
            with_podcast_state(hwnd, |state| {
                state.saving_dialog = HWND(0);
                state.save_cancel = None;
                let _ = with_state(state.parent, |app| {
                    app.podcast_save_window = HWND(0);
                });
                if state.start_button.0 != 0 {
                    SetFocus(state.start_button);
                }
            });
            LRESULT(0)
        }
        podcast_save_window::WM_PODCAST_SAVE_CANCEL => {
            with_podcast_state(hwnd, |state| {
                if let Some(cancel) = state.save_cancel.as_ref() {
                    cancel.store(true, Ordering::Relaxed);
                }
            });
            LRESULT(0)
        }
        WM_PODCAST_SAVE_RESULT => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let result = unsafe { Box::from_raw(lparam.0 as *mut PodcastSaveResult) };
            with_podcast_state(hwnd, |state| {
                let title = if result.success {
                    i18n::tr(state.language, "podcast.done_title")
                } else {
                    crate::settings::error_title(state.language)
                };
                let title_w = to_wide(&title);
                let msg_w = to_wide(&result.message);
                let flags = if result.success {
                    MB_OK | MB_ICONINFORMATION
                } else {
                    MB_OK | MB_ICONERROR
                };
                unsafe {
                    MessageBoxW(
                        hwnd,
                        PCWSTR(msg_w.as_ptr()),
                        PCWSTR(title_w.as_ptr()),
                        flags,
                    );
                    SetFocus(state.start_button);
                }
            });
            LRESULT(0)
        }
        WM_CLOSE => {
            let mut should_close = true;
            with_podcast_state(hwnd, |state| {
                if let Some(recorder) = state.recorder.as_ref() {
                    if matches!(
                        recorder.status(),
                        RecorderStatus::Recording | RecorderStatus::Paused
                    ) {
                        let labels = labels(state.language);
                        let text = to_wide(&labels.confirm_close_recording);
                        let title = to_wide(&confirm_title(state.language));
                        let result = windows::Win32::UI::WindowsAndMessaging::MessageBoxW(
                            hwnd,
                            PCWSTR(text.as_ptr()),
                            PCWSTR(title.as_ptr()),
                            MB_OKCANCEL | MB_ICONWARNING,
                        );
                        if result == IDOK {
                            stop_recording_action(state, hwnd);
                        } else {
                            should_close = false;
                        }
                    }
                }
            });
            if should_close {
                let _ = DestroyWindow(hwnd);
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            with_podcast_state(hwnd, |state| {
                if let Some(recorder) = state.recorder.take() {
                    let _ = recorder.stop();
                }
                if state.saving_dialog.0 != 0 {
                    let _ = DestroyWindow(state.saving_dialog);
                    state.saving_dialog = HWND(0);
                    state.save_cancel = None;
                    let _ = with_state(state.parent, |app| {
                        app.podcast_save_window = HWND(0);
                    });
                }
                EnableWindow(state.parent, true);
                unsafe {
                    let _ =
                        PostMessageW(state.parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
                }
            });
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let parent = with_podcast_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                let _ = with_state(parent, |state| {
                    state.podcast_window = HWND(0);
                });
            }
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
            if ptr != 0 {
                unsafe {
                    let _ = Box::from_raw(ptr as *mut PodcastState);
                }
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

fn with_podcast_state<T>(hwnd: HWND, f: impl FnOnce(&mut PodcastState) -> T) -> Option<T> {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) };
    if ptr == 0 {
        return None;
    }
    let state = unsafe { &mut *(ptr as *mut PodcastState) };
    Some(f(state))
}
fn populate_combos(
    format_combo: HWND,
    bitrate_combo: HWND,
    mic_combo: HWND,
    system_combo: HWND,
    language: Language,
) {
    unsafe {
        let _ = SendMessageW(format_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let mp3 = to_wide("MP3");
        let wav = to_wide("WAV");
        let _ = SendMessageW(
            format_combo,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(mp3.as_ptr() as isize),
        );
        let _ = SendMessageW(
            format_combo,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(wav.as_ptr() as isize),
        );

        let _ = SendMessageW(bitrate_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for bitrate in ["128 kbps", "192 kbps", "256 kbps"] {
            let text = to_wide(bitrate);
            let _ = SendMessageW(
                bitrate_combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(text.as_ptr() as isize),
            );
        }

        let _ = SendMessageW(mic_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(system_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));

        let default_text = to_wide(&labels(language).default_device);
        let _ = SendMessageW(
            mic_combo,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(default_text.as_ptr() as isize),
        );
        let _ = SendMessageW(
            system_combo,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(default_text.as_ptr() as isize),
        );
    }
}

fn load_devices(language: Language) -> (Vec<AudioDevice>, Vec<AudioDevice>, bool) {
    let mut mic_devices = vec![AudioDevice {
        id: PODCAST_DEVICE_DEFAULT.to_string(),
        name: labels(language).default_device,
    }];
    if let Ok(list) = list_input_devices() {
        mic_devices.extend(list);
    }
    let mut system_devices = vec![AudioDevice {
        id: PODCAST_DEVICE_DEFAULT.to_string(),
        name: labels(language).default_device,
    }];
    let mut system_available = false;
    if let Ok(list) = list_output_devices() {
        system_available = !list.is_empty();
        system_devices.extend(list);
    }
    (mic_devices, system_devices, system_available)
}

fn apply_settings_to_ui(state: &mut PodcastState, settings: &AppSettings) {
    unsafe {
        let _ = SendMessageW(
            state.include_mic,
            BM_SETCHECK,
            WPARAM(if settings.podcast_include_microphone {
                BST_CHECKED.0 as usize
            } else {
                BST_UNCHECKED.0 as usize
            }),
            LPARAM(0),
        );
        let _ = SendMessageW(
            state.include_system,
            BM_SETCHECK,
            WPARAM(if settings.podcast_include_system_audio {
                BST_CHECKED.0 as usize
            } else {
                BST_UNCHECKED.0 as usize
            }),
            LPARAM(0),
        );

        for (index, device) in state.mic_devices.iter().enumerate() {
            let name = to_wide(&device.name);
            let _ = SendMessageW(
                state.mic_device,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(name.as_ptr() as isize),
            );
            if device.id == settings.podcast_microphone_device_id {
                let _ = SendMessageW(state.mic_device, CB_SETCURSEL, WPARAM(index), LPARAM(0));
            }
        }
        if SendMessageW(state.mic_device, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 == -1 {
            let _ = SendMessageW(state.mic_device, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        }

        for (index, device) in state.system_devices.iter().enumerate() {
            let name = to_wide(&device.name);
            let _ = SendMessageW(
                state.system_device,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(name.as_ptr() as isize),
            );
            if device.id == settings.podcast_system_device_id {
                let _ = SendMessageW(state.system_device, CB_SETCURSEL, WPARAM(index), LPARAM(0));
            }
        }
        if SendMessageW(state.system_device, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 == -1 {
            let _ = SendMessageW(state.system_device, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        }

        let format_index = match settings.podcast_output_format {
            PodcastFormat::Mp3 => 0,
            PodcastFormat::Wav => 1,
        };
        let _ = SendMessageW(
            state.format_combo,
            CB_SETCURSEL,
            WPARAM(format_index),
            LPARAM(0),
        );

        let bitrate_index = match settings.podcast_mp3_bitrate {
            192 => 1,
            256 => 2,
            _ => 0,
        };
        let _ = SendMessageW(
            state.bitrate_combo,
            CB_SETCURSEL,
            WPARAM(bitrate_index),
            LPARAM(0),
        );

        let path = if settings.podcast_save_folder.trim().is_empty() {
            default_podcast_save_folder()
        } else {
            settings.podcast_save_folder.clone()
        };
        let path_w = to_wide(&path);
        let _ = SetWindowTextW(state.save_path, PCWSTR(path_w.as_ptr()));
    }
}

fn update_source_controls(state: &PodcastState) {
    unsafe {
        let mic_checked = is_checked(state.include_mic);
        let system_checked = is_checked(state.include_system);
        EnableWindow(state.mic_device, mic_checked);
        EnableWindow(state.system_device, system_checked);

        if !state.system_available {
            EnableWindow(state.include_system, false);
            EnableWindow(state.system_device, false);
            ShowWindow(
                state.system_unavailable_text,
                windows::Win32::UI::WindowsAndMessaging::SW_SHOW,
            );
        } else {
            ShowWindow(
                state.system_unavailable_text,
                windows::Win32::UI::WindowsAndMessaging::SW_HIDE,
            );
        }

        let hint = if !mic_checked && !system_checked {
            windows::Win32::UI::WindowsAndMessaging::SW_SHOW
        } else {
            windows::Win32::UI::WindowsAndMessaging::SW_HIDE
        };
        ShowWindow(state.hint_text, hint);
    }
}

fn update_format_controls(state: &PodcastState) {
    unsafe {
        let format = selected_format(state);
        EnableWindow(state.bitrate_combo, format == PodcastFormat::Mp3);
    }
}

fn update_recording_controls(state: &PodcastState) {
    unsafe {
        let has_sources = is_checked(state.include_mic) || is_checked(state.include_system);
        let status = state
            .recorder
            .as_ref()
            .map(|recorder| recorder.status())
            .unwrap_or(RecorderStatus::Idle);
        let recording = matches!(status, RecorderStatus::Recording);
        let paused = matches!(status, RecorderStatus::Paused);

        EnableWindow(state.start_button, has_sources && !recording && !paused);
        EnableWindow(state.pause_button, recording || paused);
        EnableWindow(state.resume_button, paused);
        EnableWindow(state.stop_button, recording || paused);
    }
}

fn update_status_text(state: &PodcastState, status: RecorderStatus) {
    let labels = labels(state.language);
    let text = match status {
        RecorderStatus::Idle => labels.status_idle,
        RecorderStatus::Recording => labels.status_recording,
        RecorderStatus::Paused => labels.status_paused,
        RecorderStatus::Saving => labels.status_saving,
        RecorderStatus::Error => labels.status_error,
    };
    let wide = to_wide(&text);
    unsafe {
        let _ = SetWindowTextW(state.status_text, PCWSTR(wide.as_ptr()));
    }
}

fn update_status_from_recorder(state: &mut PodcastState) {
    if let Some(recorder) = state.recorder.as_ref() {
        let status = recorder.status();
        update_status_text(state, status);
        let elapsed = recorder.elapsed();
        let total_secs = elapsed.as_secs();
        let hours = total_secs / 3600;
        let mins = (total_secs / 60) % 60;
        let secs = total_secs % 60;
        let time_text = format!("{:02}:{:02}:{:02}", hours, mins, secs);
        let time_w = to_wide(&time_text);
        unsafe {
            let _ = SetWindowTextW(state.elapsed_text, PCWSTR(time_w.as_ptr()));
        }
        let levels = recorder.levels();
        let mic_text = levels.mic_peak.to_string();
        let sys_text = levels.system_peak.to_string();
        unsafe {
            let mic_w = to_wide(&mic_text);
            let sys_w = to_wide(&sys_text);
            let _ = SetWindowTextW(state.level_mic_text, PCWSTR(mic_w.as_ptr()));
            let _ = SetWindowTextW(state.level_system_text, PCWSTR(sys_w.as_ptr()));
        }
        if let Some(err) = recorder.take_error() {
            unsafe {
                show_error(state.parent, state.language, &err);
            }
        }
    } else {
        update_status_text(state, RecorderStatus::Idle);
    }
}

fn update_filename_preview(state: &PodcastState) {
    let format = selected_format(state);
    let ext = match format {
        PodcastFormat::Mp3 => "mp3",
        PodcastFormat::Wav => "wav",
    };
    let timestamp = chrono::Local::now().format("%Y-%m-%d_%H-%M-%S");
    let name = format!("Podcast_{timestamp}.{ext}");
    let wide = to_wide(&name);
    unsafe {
        let _ = SetWindowTextW(state.filename_preview, PCWSTR(wide.as_ptr()));
    }
}
fn start_recording_action(state: &mut PodcastState) {
    if state.recorder.is_some() {
        return;
    }
    let labels = labels(state.language);
    let include_mic = is_checked(state.include_mic);
    let include_system = is_checked(state.include_system);
    if !include_mic && !include_system {
        return;
    }

    let mic_device_id = selected_device_id(state, true);
    let system_device_id = selected_device_id(state, false);
    if include_mic {
        if let Err(err) = probe_device(&mic_device_id, false) {
            unsafe {
                show_error(
                    state.parent,
                    state.language,
                    &format!("{} {}", labels.error_microphone, err),
                );
            }
            if include_system {
                unsafe {
                    let _ = SendMessageW(
                        state.include_mic,
                        BM_SETCHECK,
                        WPARAM(BST_UNCHECKED.0 as usize),
                        LPARAM(0),
                    );
                }
            }
        }
    }
    if include_system {
        if let Err(err) = probe_device(&system_device_id, true) {
            unsafe {
                show_error(
                    state.parent,
                    state.language,
                    &format!("{} {}", labels.error_system_audio, err),
                );
            }
            if include_mic {
                unsafe {
                    let _ = SendMessageW(
                        state.include_system,
                        BM_SETCHECK,
                        WPARAM(BST_UNCHECKED.0 as usize),
                        LPARAM(0),
                    );
                }
            }
        }
    }

    let include_system = is_checked(state.include_system);
    let include_mic = is_checked(state.include_mic);
    if !include_mic && !include_system {
        update_recording_controls(state);
        update_source_controls(state);
        return;
    }

    let config = RecorderConfig {
        include_mic,
        mic_device_id: selected_device_id(state, true),
        include_system,
        system_device_id: selected_device_id(state, false),
        output_format: selected_format(state),
        mp3_bitrate: selected_bitrate(state),
        save_folder: selected_save_folder(state),
    };
    match start_recording(config) {
        Ok(recorder) => {
            state.recorder = Some(recorder);
            update_recording_controls(state);
            update_status_text(state, RecorderStatus::Recording);
            unsafe {
                SetFocus(state.pause_button);
            }
        }
        Err(err) => {
            unsafe {
                show_error(state.parent, state.language, &err);
            }
            update_recording_controls(state);
        }
    }
}

fn stop_recording_action(state: &mut PodcastState, hwnd: HWND) {
    if state.recorder.is_none() {
        return;
    }
    if state.saving_dialog.0 == 0 {
        let dialog = unsafe { podcast_save_window::open(hwnd) };
        if dialog.0 != 0 {
            state.saving_dialog = dialog;
            unsafe {
                let _ = with_state(state.parent, |app| {
                    app.podcast_save_window = dialog;
                });
            }
        }
    }
    let cancel_flag = Arc::new(AtomicBool::new(false));
    state.save_cancel = Some(cancel_flag.clone());
    if let Some(recorder) = state.recorder.take() {
        update_status_text(state, RecorderStatus::Saving);
        let language = state.language;
        let dialog = state.saving_dialog;
        let cancel = cancel_flag;
        std::thread::spawn(move || {
            let result = recorder.stop_with_progress(
                |pct| {
                    if dialog.0 != 0 {
                        unsafe {
                            let _ = PostMessageW(
                                dialog,
                                podcast_save_window::WM_PODCAST_SAVE_PROGRESS,
                                WPARAM(pct as usize),
                                LPARAM(0),
                            );
                        }
                    }
                },
                Some(cancel.clone()),
            );
            let cancelled = cancel.load(Ordering::Relaxed);
            let mut notify = None;
            if let Err(err) = result {
                if !(cancelled && err == "Saving canceled.") {
                    notify = Some(PodcastSaveResult {
                        success: false,
                        message: err,
                    });
                }
            } else {
                notify = Some(PodcastSaveResult {
                    success: true,
                    message: i18n::tr(language, "podcast.saved"),
                });
            }
            if dialog.0 != 0 {
                unsafe {
                    let _ = PostMessageW(
                        dialog,
                        podcast_save_window::WM_PODCAST_SAVE_DONE,
                        WPARAM(0),
                        LPARAM(0),
                    );
                }
            }
            if let Some(payload) = notify {
                let _ = unsafe {
                    PostMessageW(
                        hwnd,
                        WM_PODCAST_SAVE_RESULT,
                        WPARAM(0),
                        LPARAM(Box::into_raw(Box::new(payload)) as isize),
                    )
                };
            }
        });
    }
    update_recording_controls(state);
}

fn persist_settings(state: &PodcastState) {
    let include_mic = is_checked(state.include_mic);
    let include_system = is_checked(state.include_system);
    let mic_device_id = selected_device_id(state, true);
    let system_device_id = selected_device_id(state, false);
    let output_format = selected_format(state);
    let bitrate = selected_bitrate(state);
    let save_folder = selected_save_folder(state).to_string_lossy().to_string();
    unsafe {
        let _ = with_state(state.parent, |app| {
            app.settings.podcast_include_microphone = include_mic;
            app.settings.podcast_microphone_device_id = mic_device_id;
            app.settings.podcast_include_system_audio = include_system;
            app.settings.podcast_system_device_id = system_device_id;
            app.settings.podcast_output_format = output_format;
            app.settings.podcast_mp3_bitrate = bitrate;
            app.settings.podcast_save_folder = save_folder;
            save_settings(app.settings.clone());
        });
    }
}

fn selected_format(state: &PodcastState) -> PodcastFormat {
    let sel = unsafe { SendMessageW(state.format_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if sel == 1 {
        PodcastFormat::Wav
    } else {
        PodcastFormat::Mp3
    }
}

fn selected_bitrate(state: &PodcastState) -> u32 {
    let sel = unsafe { SendMessageW(state.bitrate_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    match sel {
        1 => 192,
        2 => 256,
        _ => 128,
    }
}

fn selected_device_id(state: &PodcastState, mic: bool) -> String {
    let combo = if mic {
        state.mic_device
    } else {
        state.system_device
    };
    let list = if mic {
        &state.mic_devices
    } else {
        &state.system_devices
    };
    let sel = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    let index = if sel < 0 { 0 } else { sel as usize };
    list.get(index)
        .map(|d| d.id.clone())
        .unwrap_or_else(|| PODCAST_DEVICE_DEFAULT.to_string())
}

fn selected_save_folder(state: &PodcastState) -> PathBuf {
    unsafe {
        let len = GetWindowTextLengthW(state.save_path) as usize;
        let mut buf = vec![0u16; len + 1];
        let read = GetWindowTextW(state.save_path, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        if text.trim().is_empty() {
            default_output_folder()
        } else {
            PathBuf::from(text)
        }
    }
}

fn is_checked(hwnd: HWND) -> bool {
    unsafe { SendMessageW(hwnd, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 == BST_CHECKED.0 as isize }
}

fn click_button(hwnd: HWND, id: usize) {
    unsafe {
        let button = GetDlgItem(hwnd, id as i32);
        if button.0 != 0 {
            let _ = SendMessageW(button, BM_CLICK, WPARAM(0), LPARAM(0));
        }
    }
}

fn browse_for_folder(owner: HWND, language: Language) -> Option<PathBuf> {
    crate::app_windows::find_in_files_window::browse_for_folder(owner, language)
}
