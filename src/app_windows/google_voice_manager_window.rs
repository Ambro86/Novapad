use crate::accessibility::to_wide;
use crate::google_tts::GoogleVoicePackageStatus;
use crate::settings::Language;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::UI::Controls::{WC_BUTTON, WC_LISTBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, VK_ESCAPE};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW,
    IsDialogMessageW, LB_ADDSTRING, LB_GETCURSEL, LB_RESETCONTENT, LB_SETCURSEL, LBN_SELCHANGE,
    LBS_NOTIFY, LoadCursorW, MB_ICONQUESTION, MB_OK, MB_YESNO, MSG, MessageBoxW, PostMessageW,
    RegisterClassW, SW_SHOW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW,
    ShowWindow, TranslateMessage, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::PCWSTR;

const CLASS_NAME: &str = "SonarpadGoogleVoiceManager";
// WM_COMMAND carries control identifiers in the low 16 bits of WPARAM.
// Keep every identifier below 65536 or button/list notifications will never match.
const ID_LIST: usize = 65001;
const ID_STATUS: usize = 65002;
const ID_DOWNLOAD: usize = 65003;
const ID_REMOVE: usize = 65004;
const ID_CLOSE: usize = 65005;
const WM_GOOGLE_DOWNLOAD_PROGRESS: u32 = WM_APP + 160;
const WM_GOOGLE_DOWNLOAD_DONE: u32 = WM_APP + 161;

struct ManagerInit {
    parent: HWND,
    language: Language,
    font: HFONT,
}

struct ManagerState {
    parent: HWND,
    language: Language,
    list: HWND,
    status: HWND,
    download_button: HWND,
    remove_button: HWND,
    packages: Vec<GoogleVoicePackageStatus>,
    downloading: bool,
    cancel: Arc<AtomicBool>,
    last_announced_progress: i32,
}

struct DownloadResult {
    package_id: String,
    result: Result<(), String>,
}

fn tr(language: Language, key: &str) -> String {
    crate::i18n::tr(language, key)
}

fn package_label(language: Language, status: &GoogleVoicePackageStatus) -> String {
    let voices = status
        .package
        .speakers
        .iter()
        .filter_map(|speaker| {
            let name = if speaker.name.trim().is_empty() {
                speaker.speaker.trim()
            } else {
                speaker.name.trim()
            };
            if name.is_empty() {
                None
            } else {
                Some(name.to_string())
            }
        })
        .collect::<Vec<_>>()
        .join(", ");
    let state = if status.installed {
        tr(language, "google_tts.voices.installed")
    } else {
        tr(language, "google_tts.voices.not_installed")
    };
    let size_mb = status.package.compressed_size as f64 / 1_048_576.0;
    if voices.is_empty() {
        format!(
            "{} — {} — {:.1} MB — {}",
            status.language, status.package.id, size_mb, state
        )
    } else {
        format!(
            "{} — {} — {:.1} MB — {}",
            status.language, voices, size_mb, state
        )
    }
}

unsafe fn state_mut(hwnd: HWND) -> Option<&'static mut ManagerState> {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ManagerState;
    if ptr.is_null() {
        None
    } else {
        Some(unsafe { &mut *ptr })
    }
}

fn selected_index(state: &ManagerState) -> Option<usize> {
    let selected = unsafe { SendMessageW(state.list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if selected < 0 {
        None
    } else {
        Some(selected as usize)
    }
}

fn set_status(state: &ManagerState, text: &str) {
    let wide = to_wide(text);
    unsafe {
        if let Err(err) = SetWindowTextW(state.status, PCWSTR(wide.as_ptr())) {
            crate::log_debug(&format!("Google TTS manager status update failed: {err}"));
        }
    }
}

fn update_buttons(state: &ManagerState) {
    let selected = selected_index(state).and_then(|index| state.packages.get(index));
    let can_download = selected.is_some_and(|item| !item.installed) && !state.downloading;
    let can_remove = selected.is_some_and(|item| item.installed) && !state.downloading;
    unsafe {
        EnableWindow(state.download_button, can_download);
        EnableWindow(state.remove_button, can_remove);
        // Keep the list enabled while downloading. Disabling the focused button and
        // the list at the same time leaves screen readers on a dead focus object.
        EnableWindow(state.list, true);
    }
}

fn focus_voice_list(hwnd: HWND, state: &ManagerState, reason: &str) {
    unsafe {
        SetForegroundWindow(hwnd);
    }
    crate::set_focus_safe(state.list);
    crate::log_debug(&format!(
        "Google TTS manager: focus restored to voice list ({reason})"
    ));
}

fn refresh_list(state: &mut ManagerState, preferred_id: Option<&str>) {
    state.packages = crate::google_tts::catalog_packages();
    unsafe {
        SendMessageW(state.list, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
    }
    let mut selected = 0usize;
    for (index, package) in state.packages.iter().enumerate() {
        let label = package_label(state.language, package);
        let wide = to_wide(&label);
        unsafe {
            SendMessageW(
                state.list,
                LB_ADDSTRING,
                WPARAM(0),
                LPARAM(wide.as_ptr() as isize),
            );
        }
        if preferred_id.is_some_and(|id| id == package.package.id) {
            selected = index;
        }
    }
    if !state.packages.is_empty() {
        unsafe {
            SendMessageW(state.list, LB_SETCURSEL, WPARAM(selected), LPARAM(0));
        }
    }
    update_buttons(state);
}

fn start_download(hwnd: HWND, state: &mut ManagerState) {
    let Some(index) = selected_index(state) else {
        return;
    };
    let Some(package) = state.packages.get(index) else {
        return;
    };
    if package.installed || state.downloading {
        return;
    }
    let package_id = package.package.id.clone();
    crate::log_debug(&format!(
        "Google TTS manager: download requested for {package_id}"
    ));
    state.downloading = true;
    state.cancel = Arc::new(AtomicBool::new(false));
    state.last_announced_progress = -1;
    update_buttons(state);
    let downloading_message = tr(state.language, "google_tts.voices.downloading");
    set_status(state, &downloading_message);
    crate::accessibility::screen_reader_speak(&downloading_message);
    // The Download button is disabled now. Move focus immediately back to the
    // list so NVDA remains inside the dialog while the worker runs.
    focus_voice_list(hwnd, state, "download_started");
    let cancel = state.cancel.clone();
    let _download_thread = thread::spawn(move || {
        crate::log_debug(&format!(
            "Google TTS manager: worker started for {package_id}"
        ));
        let result = crate::google_tts::download_package(&package_id, &cancel, |percentage| {
            if let Err(err) = unsafe {
                PostMessageW(
                    hwnd,
                    WM_GOOGLE_DOWNLOAD_PROGRESS,
                    WPARAM(percentage.max(0) as usize),
                    LPARAM(0),
                )
            } {
                crate::log_debug(&format!("Google TTS progress post failed: {err}"));
            }
        });
        match &result {
            Ok(()) => crate::log_debug(&format!(
                "Google TTS manager: download completed for {package_id}"
            )),
            Err(err) => crate::log_debug(&format!(
                "Google TTS manager: download failed for {package_id}: {err}"
            )),
        }
        let payload = Box::new(DownloadResult { package_id, result });
        let raw = Box::into_raw(payload);
        if let Err(err) = unsafe {
            PostMessageW(
                hwnd,
                WM_GOOGLE_DOWNLOAD_DONE,
                WPARAM(0),
                LPARAM(raw as isize),
            )
        } {
            crate::log_debug(&format!("Google TTS completion post failed: {err}"));
            let _unused = unsafe { Box::from_raw(raw) };
        }
    });
}

fn remove_selected(hwnd: HWND, state: &mut ManagerState) {
    let Some(index) = selected_index(state) else {
        return;
    };
    let Some(package) = state.packages.get(index) else {
        return;
    };
    if !package.installed || state.downloading {
        return;
    }
    let question = tr(state.language, "google_tts.voices.remove_confirm");
    let title = tr(state.language, "google_tts.voices.title");
    let answer = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(to_wide(&question).as_ptr()),
            PCWSTR(to_wide(&title).as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    if answer.0 != 6 {
        return;
    }
    let package_id = package.package.id.clone();
    match crate::google_tts::remove_package(&package_id) {
        Ok(()) => {
            refresh_list(state, Some(&package_id));
            set_status(state, &tr(state.language, "google_tts.voices.removed"));
            focus_voice_list(hwnd, state, "package_removed");
        }
        Err(err) => unsafe {
            MessageBoxW(
                hwnd,
                PCWSTR(to_wide(&err).as_ptr()),
                PCWSTR(to_wide(&tr(state.language, "app.error_title")).as_ptr()),
                MB_OK,
            );
        },
    }
}

unsafe extern "system" fn manager_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "google_voice_manager_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || manager_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn manager_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*create).lpCreateParams as *mut ManagerInit;
                if init_ptr.is_null() {
                    return LRESULT(-1);
                }
                let init = Box::from_raw(init_ptr);
                let font = init.font;
                let list = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_LISTBOXW,
                    PCWSTR::null(),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WS_VSCROLL
                        | WINDOW_STYLE(LBS_NOTIFY as u32),
                    16,
                    16,
                    760,
                    360,
                    hwnd,
                    HMENU(ID_LIST as isize),
                    HINSTANCE(0),
                    None,
                );
                let status = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&tr(init.language, "google_tts.voices.ready")).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    384,
                    760,
                    26,
                    hwnd,
                    HMENU(ID_STATUS as isize),
                    HINSTANCE(0),
                    None,
                );
                let download_button = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&tr(init.language, "google_tts.voices.download")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    400,
                    420,
                    120,
                    30,
                    hwnd,
                    HMENU(ID_DOWNLOAD as isize),
                    HINSTANCE(0),
                    None,
                );
                let remove_button = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&tr(init.language, "google_tts.voices.remove")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    528,
                    420,
                    120,
                    30,
                    hwnd,
                    HMENU(ID_REMOVE as isize),
                    HINSTANCE(0),
                    None,
                );
                let close_button = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&tr(init.language, "google_tts.voices.close")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    656,
                    420,
                    120,
                    30,
                    hwnd,
                    HMENU(ID_CLOSE as isize),
                    HINSTANCE(0),
                    None,
                );
                for control in [list, status, download_button, remove_button, close_button] {
                    if control.0 != 0 && font.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(font.0 as usize), LPARAM(1));
                    }
                }
                let mut state = Box::new(ManagerState {
                    parent: init.parent,
                    language: init.language,
                    list,
                    status,
                    download_button,
                    remove_button,
                    packages: Vec::new(),
                    downloading: false,
                    cancel: Arc::new(AtomicBool::new(false)),
                    last_announced_progress: -1,
                });
                refresh_list(&mut state, None);
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                SetForegroundWindow(hwnd);
                crate::set_focus_safe(list);
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                let notification = ((wparam.0 >> 16) & 0xffff) as u32;
                let Some(state) = state_mut(hwnd) else {
                    return DefWindowProcW(hwnd, msg, wparam, lparam);
                };
                match id {
                    ID_LIST if notification == LBN_SELCHANGE => {
                        update_buttons(state);
                        LRESULT(0)
                    }
                    ID_DOWNLOAD => {
                        crate::log_debug("Google TTS manager: Download button activated");
                        start_download(hwnd, state);
                        LRESULT(0)
                    }
                    ID_REMOVE => {
                        crate::log_debug("Google TTS manager: Remove button activated");
                        remove_selected(hwnd, state);
                        LRESULT(0)
                    }
                    ID_CLOSE | 2 => {
                        crate::log_debug("Google TTS manager: Close or Escape activated");
                        if let Err(err) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                            crate::log_debug(&format!(
                                "Google TTS manager close post failed: {err}"
                            ));
                        }
                        LRESULT(0)
                    }
                    _ => DefWindowProcW(hwnd, msg, wparam, lparam),
                }
            }
            WM_GOOGLE_DOWNLOAD_PROGRESS => {
                if let Some(state) = state_mut(hwnd) {
                    let message = tr(state.language, "google_tts.voices.progress")
                        .replace("{pct}", &wparam.0.min(100).to_string());
                    set_status(state, &message);
                    let percentage = wparam.0.min(100) as i32;
                    if percentage == 0
                        || percentage == 100
                        || percentage >= state.last_announced_progress.saturating_add(5)
                    {
                        state.last_announced_progress = percentage;
                        crate::accessibility::screen_reader_speak(&message);
                    }
                }
                LRESULT(0)
            }
            WM_GOOGLE_DOWNLOAD_DONE => {
                let payload = lparam.0 as *mut DownloadResult;
                if payload.is_null() {
                    return LRESULT(0);
                }
                let payload = Box::from_raw(payload);
                if let Some(state) = state_mut(hwnd) {
                    state.downloading = false;
                    refresh_list(state, Some(&payload.package_id));
                    match &payload.result {
                        Ok(()) => {
                            let message = tr(state.language, "google_tts.voices.download_complete");
                            set_status(state, &message);
                            crate::accessibility::screen_reader_speak(&message);
                            focus_voice_list(hwnd, state, "download_complete");
                        }
                        Err(err) if err == "cancelled" => {
                            set_status(state, &tr(state.language, "google_tts.voices.cancelled"));
                            focus_voice_list(hwnd, state, "download_cancelled");
                        }
                        Err(err) => {
                            set_status(state, err);
                            MessageBoxW(
                                hwnd,
                                PCWSTR(to_wide(err).as_ptr()),
                                PCWSTR(to_wide(&tr(state.language, "app.error_title")).as_ptr()),
                                MB_OK,
                            );
                            focus_voice_list(hwnd, state, "download_error");
                        }
                    }
                }
                LRESULT(0)
            }
            WM_KEYDOWN => {
                if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                    if let Err(err) = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0)) {
                        crate::log_debug(&format!("Google TTS manager escape close failed: {err}"));
                    }
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                if let Some(state) = state_mut(hwnd) {
                    state.cancel.store(true, Ordering::Relaxed);
                }
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_DESTROY => {
                if let Some(state) = state_mut(hwnd) {
                    EnableWindow(state.parent, true);
                    SetForegroundWindow(state.parent);
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ManagerState;
                if !ptr.is_null() {
                    let _unused = Box::from_raw(ptr);
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn open_modal(parent: HWND, language: Language, font: HFONT) {
    unsafe {
        let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
        let class_name = to_wide(CLASS_NAME);
        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hInstance: hinstance,
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(manager_wndproc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
            ..Default::default()
        };
        RegisterClassW(&wc);
        let init = Box::new(ManagerInit {
            parent,
            language,
            font,
        });
        let title = to_wide(&tr(language, "google_tts.voices.title"));
        let hwnd = CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            812,
            510,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        );
        if hwnd.0 == 0 {
            crate::log_debug("Google TTS manager: window creation failed");
            return;
        }
        crate::watchdog::enter_modal_dialog();
        EnableWindow(parent, false);
        ShowWindow(hwnd, SW_SHOW);
        SetForegroundWindow(hwnd);
        let mut message = MSG::default();
        loop {
            if !crate::is_window_handle_valid(hwnd) {
                break;
            }
            let result = GetMessageW(&mut message, HWND(0), 0, 0);
            if result.0 == 0 || result.0 == -1 {
                break;
            }
            if crate::app_windows::calendar_window::handle_reminder_alert_message(&message) {
                continue;
            }
            if IsDialogMessageW(hwnd, &message).as_bool() {
                continue;
            }
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        EnableWindow(parent, true);
        SetForegroundWindow(parent);
        crate::watchdog::exit_modal_dialog();
    }
}

pub fn open_with_language(parent: HWND, language: Language, font: HFONT) {
    open_modal(parent, language, font);
}
