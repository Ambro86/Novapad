use crate::accessibility::to_wide;
use crate::google_tts::GoogleVoicePackageStatus;
use crate::settings::Language;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::{WC_BUTTON, WC_COMBOBOXW, WC_LISTBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, VK_ESCAPE};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL, CB_RESETCONTENT, CB_SETCURSEL, CBN_SELCHANGE,
    CBS_DROPDOWNLIST, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HMENU, IDC_ARROW,
    IsDialogMessageW, LB_ADDSTRING, LB_GETCURSEL, LB_RESETCONTENT, LB_SETCURSEL, LBN_SELCHANGE,
    LBS_NOTIFY, LoadCursorW, MB_ICONQUESTION, MB_OK, MB_YESNO, MSG, MessageBoxW, PostMessageW,
    RegisterClassW, SW_SHOW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW,
    ShowWindow, TranslateMessage, WINDOW_EX_STYLE, WINDOW_STYLE, WM_ACTIVATE, WM_ACTIVATEAPP,
    WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL,
    WM_SETFOCUS, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE,
    WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
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
const ID_LANGUAGE_FILTER: usize = 65006;
const WM_GOOGLE_DOWNLOAD_PROGRESS: u32 = WM_APP + 160;
const WM_GOOGLE_DOWNLOAD_DONE: u32 = WM_APP + 161;
const WM_GOOGLE_RESTORE_LANGUAGE_FOCUS: u32 = WM_APP + 162;

struct ManagerInit {
    parent: HWND,
    language: Language,
    font: HFONT,
}

struct ManagerState {
    parent: HWND,
    language: Language,
    language_filter: HWND,
    list: HWND,
    status: HWND,
    download_button: HWND,
    remove_button: HWND,
    all_packages: Vec<GoogleVoicePackageStatus>,
    packages: Vec<GoogleVoicePackageStatus>,
    language_codes: Vec<String>,
    downloading: bool,
    cancel: Arc<AtomicBool>,
    last_announced_progress: i32,
    focus_restore_pending: bool,
}

struct DownloadResult {
    package_id: String,
    result: Result<(), String>,
}

fn tr(language: Language, key: &str) -> String {
    crate::i18n::tr(language, key)
}

fn localized_language_name(language: Language, locale: &str) -> String {
    let mut parts = locale.split(['-', '_']);
    let root = parts.next().unwrap_or(locale).trim().to_ascii_lowercase();
    let region = parts.next().unwrap_or_default().trim().to_ascii_lowercase();

    let language_key = format!("voice.lang.{root}");
    let language_name = {
        let value = tr(language, &language_key);
        if value == language_key {
            root.to_ascii_uppercase()
        } else {
            value
        }
    };
    if region.is_empty() || region == "xa" {
        return language_name;
    }

    let country_key = format!("options.podcast_country.{region}");
    let country_name = tr(language, &country_key);
    if country_name == country_key {
        format!("{language_name} ({})", region.to_ascii_uppercase())
    } else {
        format!("{language_name} ({country_name})")
    }
}

fn collect_language_codes(
    language: Language,
    packages: &[GoogleVoicePackageStatus],
) -> Vec<String> {
    let mut codes = packages
        .iter()
        .map(|package| package.language.clone())
        .collect::<Vec<_>>();
    codes.sort_by(|left, right| {
        localized_language_name(language, left)
            .to_lowercase()
            .cmp(&localized_language_name(language, right).to_lowercase())
            .then_with(|| left.cmp(right))
    });
    codes.dedup_by(|left, right| left.eq_ignore_ascii_case(right));
    codes
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
            localized_language_name(language, &status.language),
            status.package.id,
            size_mb,
            state
        )
    } else {
        format!(
            "{} — {} — {:.1} MB — {}",
            localized_language_name(language, &status.language),
            voices,
            size_mb,
            state
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

fn focus_language_filter(hwnd: HWND, state: &ManagerState, reason: &str) {
    unsafe {
        SetForegroundWindow(hwnd);
        SendMessageW(
            hwnd,
            WM_NEXTDLGCTL,
            WPARAM(state.language_filter.0 as usize),
            LPARAM(1),
        );
    }
    crate::set_focus_safe(state.language_filter);
    unsafe {
        NotifyWinEvent(
            crate::EVENT_OBJECT_FOCUS,
            state.language_filter,
            crate::OBJID_CLIENT.0,
            windows::Win32::UI::WindowsAndMessaging::CHILDID_SELF as i32,
        );
    }
    crate::log_debug(&format!(
        "Google TTS manager: focus moved to language filter ({reason})"
    ));
}

fn post_language_filter_focus(hwnd: HWND, state: &mut ManagerState, reason: &str) {
    if state.focus_restore_pending {
        return;
    }
    state.focus_restore_pending = true;
    crate::log_debug(&format!(
        "Google TTS manager: queued language filter focus ({reason})"
    ));
    if let Err(error) =
        unsafe { PostMessageW(hwnd, WM_GOOGLE_RESTORE_LANGUAGE_FOCUS, WPARAM(0), LPARAM(0)) }
    {
        state.focus_restore_pending = false;
        crate::log_debug(&format!(
            "Google TTS manager: unable to queue language filter focus: {error:?}"
        ));
    }
}

fn selected_language_code(state: &ManagerState) -> Option<&str> {
    let selected =
        unsafe { SendMessageW(state.language_filter, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if selected <= 0 {
        None
    } else {
        state
            .language_codes
            .get(selected as usize - 1)
            .map(String::as_str)
    }
}

fn fill_language_filter(state: &mut ManagerState, preferred_code: Option<&str>) {
    state.language_codes = collect_language_codes(state.language, &state.all_packages);
    unsafe {
        SendMessageW(state.language_filter, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let all_label = to_wide(&tr(state.language, "google_tts.voices.all_languages"));
        SendMessageW(
            state.language_filter,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(all_label.as_ptr() as isize),
        );
    }

    let mut selected = 0usize;
    for (index, code) in state.language_codes.iter().enumerate() {
        let label = localized_language_name(state.language, code);
        let wide = to_wide(&label);
        unsafe {
            SendMessageW(
                state.language_filter,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(wide.as_ptr() as isize),
            );
        }
        if preferred_code.is_some_and(|preferred| preferred.eq_ignore_ascii_case(code)) {
            selected = index + 1;
        }
    }
    unsafe {
        SendMessageW(
            state.language_filter,
            CB_SETCURSEL,
            WPARAM(selected),
            LPARAM(0),
        );
    }
}

fn populate_package_list(state: &mut ManagerState, preferred_id: Option<&str>) {
    let selected_language = selected_language_code(state).map(str::to_string);
    state.packages = state
        .all_packages
        .iter()
        .filter(|package| {
            selected_language
                .as_deref()
                .is_none_or(|code| package.language.eq_ignore_ascii_case(code))
        })
        .cloned()
        .collect();

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

fn refresh_list(state: &mut ManagerState, preferred_id: Option<&str>) {
    let selected_language = selected_language_code(state).map(str::to_string);
    state.all_packages = crate::google_tts::catalog_packages();
    fill_language_filter(state, selected_language.as_deref());
    populate_package_list(state, preferred_id);
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
    match crate::google_tts::remove_package(&package_id, state.language) {
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
                let filter_label = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_STATIC,
                    PCWSTR(
                        to_wide(&tr(init.language, "google_tts.voices.filter_language")).as_ptr(),
                    ),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    18,
                    190,
                    24,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let language_filter = CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                    210,
                    14,
                    566,
                    300,
                    hwnd,
                    HMENU(ID_LANGUAGE_FILTER as isize),
                    HINSTANCE(0),
                    None,
                );
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
                    52,
                    760,
                    324,
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
                for control in [
                    filter_label,
                    language_filter,
                    list,
                    status,
                    download_button,
                    remove_button,
                    close_button,
                ] {
                    if control.0 != 0 && font.0 != 0 {
                        SendMessageW(control, WM_SETFONT, WPARAM(font.0 as usize), LPARAM(1));
                    }
                }
                let mut state = Box::new(ManagerState {
                    parent: init.parent,
                    language: init.language,
                    language_filter,
                    list,
                    status,
                    download_button,
                    remove_button,
                    all_packages: Vec::new(),
                    packages: Vec::new(),
                    language_codes: Vec::new(),
                    downloading: false,
                    cancel: Arc::new(AtomicBool::new(false)),
                    last_announced_progress: -1,
                    focus_restore_pending: false,
                });
                refresh_list(&mut state, None);
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                if let Some(state) = state_mut(hwnd) {
                    focus_language_filter(hwnd, state, "dialog_created");
                }
                LRESULT(0)
            }
            WM_ACTIVATE => {
                if wparam.0 & 0xffff != 0
                    && let Some(state) = state_mut(hwnd)
                {
                    post_language_filter_focus(hwnd, state, "window_activated");
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_ACTIVATEAPP => {
                if wparam.0 != 0
                    && let Some(state) = state_mut(hwnd)
                {
                    post_language_filter_focus(hwnd, state, "application_reactivated");
                }
                LRESULT(0)
            }
            WM_SETFOCUS => {
                if let Some(state) = state_mut(hwnd) {
                    post_language_filter_focus(hwnd, state, "window_received_focus");
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_GOOGLE_RESTORE_LANGUAGE_FOCUS => {
                if let Some(state) = state_mut(hwnd) {
                    state.focus_restore_pending = false;
                    focus_language_filter(hwnd, state, "posted_after_activation");
                }
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                let notification = ((wparam.0 >> 16) & 0xffff) as u32;
                let Some(state) = state_mut(hwnd) else {
                    return DefWindowProcW(hwnd, msg, wparam, lparam);
                };
                match id {
                    ID_LANGUAGE_FILTER if notification == CBN_SELCHANGE => {
                        populate_package_list(state, None);
                        LRESULT(0)
                    }
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
        if let Some(state) = state_mut(hwnd) {
            focus_language_filter(hwnd, state, "dialog_shown");
        }
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
