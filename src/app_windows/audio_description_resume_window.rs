use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_PATHMUSTEXIST,
    OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{WC_BUTTON, WC_COMBOBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus, VK_ESCAPE};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL, CB_SETCURSEL, CBS_DROPDOWNLIST, CREATESTRUCTW,
    CW_USEDEFAULT, CreateWindowExW, DestroyWindow, DispatchMessageW, GWLP_USERDATA,
    GetWindowLongPtrW, HMENU, IDC_ARROW, IsDialogMessageW, LoadCursorW, MSG, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, PWSTR};

use crate::accessibility::to_wide;
use crate::audio_description::load_audio_description_resume_settings;
use crate::i18n;
use crate::settings::{Language, default_audio_description_save_folder, save_settings};
use crate::with_state;

const CLASS_NAME: &str = "SonarpadAudioDescriptionResumeSelector";
const ID_PROJECT: usize = 9681;
const ID_CONTINUE: usize = 9682;
const ID_BROWSE: usize = 9683;
const ID_CANCEL: usize = 9684;
const CHECKPOINT_SUFFIX: &str = ".sonarpad-ad.partial.json";
const MAX_RECENT_FOLDERS: usize = 8;

#[derive(Clone)]
struct ResumeCandidate {
    path: PathBuf,
    label: String,
    modified: SystemTime,
}

struct ResumeSelectorInit {
    app_parent: HWND,
    language: Language,
    candidates: Vec<ResumeCandidate>,
    result: Arc<Mutex<Option<PathBuf>>>,
}

struct ResumeSelectorState {
    app_parent: HWND,
    language: Language,
    combo: HWND,
    candidates: Vec<ResumeCandidate>,
    result: Arc<Mutex<Option<PathBuf>>>,
}

fn candidate_project_name(path: &Path) -> String {
    let filename = path
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or_default();
    let without_checkpoint = filename.strip_suffix(CHECKPOINT_SUFFIX).unwrap_or(filename);
    let without_audio_description = without_checkpoint
        .strip_suffix("_audiodescritto")
        .unwrap_or(without_checkpoint);
    without_audio_description.replace('_', " ")
}

fn candidate_label(path: &Path) -> Option<String> {
    let resume = load_audio_description_resume_settings(path).ok()?;
    let name = candidate_project_name(path);
    let folder = path
        .parent()
        .and_then(Path::file_name)
        .and_then(|value| value.to_str())
        .unwrap_or_default();
    let progress = format!("{}/{}", resume.completed_chunks, resume.total_chunks);
    Some(if folder.is_empty() {
        format!("{name} — {progress}")
    } else {
        format!("{name} — {progress} — {folder}")
    })
}

fn is_checkpoint_path(path: &Path) -> bool {
    path.is_file()
        && path
            .file_name()
            .and_then(|value| value.to_str())
            .is_some_and(|name| name.to_ascii_lowercase().ends_with(CHECKPOINT_SUFFIX))
}

fn candidate_directories(app_parent: HWND) -> Vec<PathBuf> {
    let (configured, recent) = with_state(app_parent, |state| {
        (
            state.settings.audio_description_save_folder.clone(),
            state
                .settings
                .audio_description_recent_project_folders
                .clone(),
        )
    })
    .unwrap_or_else(|| (default_audio_description_save_folder(), Vec::new()));

    let mut directories = Vec::new();
    for folder in std::iter::once(configured).chain(recent) {
        let folder = folder.trim();
        if folder.is_empty() {
            continue;
        }
        let path = PathBuf::from(folder);
        let key = path.to_string_lossy();
        if directories
            .iter()
            .any(|known: &PathBuf| known.to_string_lossy().eq_ignore_ascii_case(&key))
        {
            continue;
        }
        directories.push(path);
    }
    directories
}

fn discover_resume_candidates(app_parent: HWND) -> Vec<ResumeCandidate> {
    let mut candidates = Vec::new();
    let mut seen = Vec::<String>::new();

    for directory in candidate_directories(app_parent) {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if !is_checkpoint_path(&path) {
                continue;
            }
            let key = path.to_string_lossy().to_string();
            if seen.iter().any(|known| known.eq_ignore_ascii_case(&key)) {
                continue;
            }
            let Some(label) = candidate_label(&path) else {
                crate::log_debug(&format!(
                    "Audio description resume selector: ignoring invalid checkpoint {}",
                    path.display()
                ));
                continue;
            };
            let modified = entry
                .metadata()
                .ok()
                .and_then(|metadata| metadata.modified().ok())
                .unwrap_or(UNIX_EPOCH);
            seen.push(key);
            candidates.push(ResumeCandidate {
                path,
                label,
                modified,
            });
        }
    }

    candidates.sort_by(|left, right| {
        right
            .modified
            .cmp(&left.modified)
            .then_with(|| left.label.to_lowercase().cmp(&right.label.to_lowercase()))
    });
    candidates
}

pub(crate) fn remember_project_folder(app_parent: HWND, path: &Path) {
    let Some(folder) = path
        .parent()
        .filter(|folder| !folder.as_os_str().is_empty())
    else {
        return;
    };
    let folder = folder.to_string_lossy().trim().to_string();
    if folder.is_empty() {
        return;
    }

    if with_state(app_parent, |state| {
        state
            .settings
            .audio_description_recent_project_folders
            .retain(|known| !known.eq_ignore_ascii_case(&folder));
        state
            .settings
            .audio_description_recent_project_folders
            .insert(0, folder);
        state
            .settings
            .audio_description_recent_project_folders
            .truncate(MAX_RECENT_FOLDERS);
        save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug(
            "Audio description resume selector: app state unavailable while remembering project folder",
        );
    }
}

fn browse_checkpoint(parent: HWND, language: Language) -> Option<PathBuf> {
    let filter = to_wide(
        "Sonarpad audio-description checkpoint (*.sonarpad-ad.partial.json)\0*.sonarpad-ad.partial.json\0\0",
    );
    let title = to_wide(&i18n::tr(language, "audio_description.resume.open_title"));
    let mut buffer = [0_u16; 2048];
    let mut dialog = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: parent,
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrTitle: PCWSTR(title.as_ptr()),
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST,
        ..Default::default()
    };
    if !unsafe { GetOpenFileNameW(&mut dialog).as_bool() } {
        return None;
    }
    let len = buffer.iter().position(|value| *value == 0)?;
    let path = PathBuf::from(String::from_utf16_lossy(&buffer[..len]));
    is_checkpoint_path(&path).then_some(path)
}

pub(crate) fn choose_resume_checkpoint(
    parent: HWND,
    app_parent: HWND,
    language: Language,
) -> Option<PathBuf> {
    let candidates = discover_resume_candidates(app_parent);
    let hinstance = HINSTANCE(crate::get_module_handle_raw_default());
    let class_name = to_wide(CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(window_proc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    crate::register_class_w_safe(&wc);

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(ResumeSelectorInit {
        app_parent,
        language,
        candidates,
        result: result.clone(),
    });
    let title = to_wide(&i18n::tr(language, "audio_description.resume.open_title"));
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            650,
            230,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };
    if hwnd.0 == 0 {
        let selected = browse_checkpoint(parent, language);
        if let Some(path) = selected.as_ref() {
            remember_project_folder(app_parent, path);
        }
        return selected;
    }

    unsafe {
        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);
    }

    let mut message = MSG::default();
    loop {
        if !crate::is_window_handle_valid(hwnd) {
            break;
        }
        let get_message = crate::get_message_w_safe(&mut message, HWND(0), 0, 0);
        if get_message.0 == 0 {
            break;
        }
        unsafe {
            if message.message == WM_KEYDOWN && message.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _destroy_result = DestroyWindow(hwnd);
                continue;
            }
            if IsDialogMessageW(hwnd, &message).as_bool() {
                continue;
            }
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    unsafe {
        EnableWindow(parent, true);
        SetForegroundWindow(parent);
    }

    let selected = result
        .lock()
        .unwrap_or_else(|error| error.into_inner())
        .clone();
    if let Some(path) = selected.as_ref() {
        remember_project_folder(app_parent, path);
    }
    selected
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "audio_description_resume_selector_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || window_proc_inner(hwnd, msg, wparam, lparam),
    )
}

fn window_proc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let create = &*(lparam.0 as *const CREATESTRUCTW);
                let init_ptr = create.lpCreateParams as *mut ResumeSelectorInit;
                if init_ptr.is_null() {
                    return LRESULT(0);
                }
                let init = Box::from_raw(init_ptr);
                let hfont = with_state(init.app_parent, |state| state.hfont).unwrap_or(HFONT(0));
                let has_candidates = !init.candidates.is_empty();
                let hint_text = if has_candidates {
                    i18n::tr(init.language, "audio_description.resume.choose_hint")
                } else {
                    i18n::tr(init.language, "audio_description.resume.none_found")
                };
                let label_text = i18n::tr(init.language, "audio_description.resume.choose_label");
                let browse_text = i18n::tr(init.language, "audio_description.resume.browse_other");
                let continue_text = i18n::tr(init.language, "audio_description.resume.start");
                let cancel_text = i18n::tr(init.language, "common.cancel");

                let hint = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&hint_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    14,
                    600,
                    36,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                );
                let label = CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&label_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    56,
                    600,
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
                    16,
                    78,
                    600,
                    220,
                    hwnd,
                    HMENU(ID_PROJECT as isize),
                    HINSTANCE(0),
                    None,
                );
                for candidate in &init.candidates {
                    let wide = to_wide(&candidate.label);
                    SendMessageW(
                        combo,
                        CB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(wide.as_ptr() as isize),
                    );
                }
                if has_candidates {
                    SendMessageW(combo, CB_SETCURSEL, WPARAM(0), LPARAM(0));
                }

                let continue_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&continue_text).as_ptr()),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WINDOW_STYLE(if has_candidates {
                            BS_DEFPUSHBUTTON as u32
                        } else {
                            0
                        }),
                    16,
                    126,
                    180,
                    30,
                    hwnd,
                    HMENU(ID_CONTINUE as isize),
                    HINSTANCE(0),
                    None,
                );
                let browse_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&browse_text).as_ptr()),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WINDOW_STYLE(if has_candidates {
                            0
                        } else {
                            BS_DEFPUSHBUTTON as u32
                        }),
                    206,
                    126,
                    250,
                    30,
                    hwnd,
                    HMENU(ID_BROWSE as isize),
                    HINSTANCE(0),
                    None,
                );
                let cancel_button = CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&cancel_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    466,
                    126,
                    150,
                    30,
                    hwnd,
                    HMENU(ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                );

                for control in [
                    hint,
                    label,
                    combo,
                    continue_button,
                    browse_button,
                    cancel_button,
                ] {
                    SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
                EnableWindow(combo, has_candidates);
                EnableWindow(continue_button, has_candidates);

                let state = Box::new(ResumeSelectorState {
                    app_parent: init.app_parent,
                    language: init.language,
                    combo,
                    candidates: init.candidates,
                    result: init.result.clone(),
                });
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
                SetFocus(if has_candidates { combo } else { browse_button });
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                let pointer = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ResumeSelectorState;
                if pointer.is_null() {
                    return LRESULT(0);
                }
                let state = &mut *pointer;
                match id {
                    ID_CONTINUE => {
                        let index = SendMessageW(state.combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        if index >= 0
                            && let Some(candidate) = state.candidates.get(index as usize)
                        {
                            if let Ok(mut result) = state.result.lock() {
                                *result = Some(candidate.path.clone());
                            }
                            let _destroy_result = DestroyWindow(hwnd);
                        }
                    }
                    ID_BROWSE => {
                        if let Some(path) = browse_checkpoint(hwnd, state.language) {
                            remember_project_folder(state.app_parent, &path);
                            if let Ok(mut result) = state.result.lock() {
                                *result = Some(path);
                            }
                            let _destroy_result = DestroyWindow(hwnd);
                        }
                    }
                    ID_CANCEL => {
                        let _destroy_result = DestroyWindow(hwnd);
                    }
                    _ => {}
                }
                LRESULT(0)
            }
            WM_CLOSE => {
                let _destroy_result = DestroyWindow(hwnd);
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let pointer = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ResumeSelectorState;
                if !pointer.is_null() {
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                    let _boxed_state = Box::from_raw(pointer);
                }
                crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam)
            }
            _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        }
    }
}
