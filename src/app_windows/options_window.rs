use windows::core::{PCWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM, LRESULT, HINSTANCE};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, RegisterClassW,
    SendMessageW, SetWindowLongPtrW, SetForegroundWindow,
    GWLP_USERDATA, WM_CREATE, WM_DESTROY, WM_NCDESTROY, WM_CLOSE, WM_COMMAND,
    WM_KEYDOWN, WM_SETFONT, WM_APP, MessageBoxW,
    WS_CAPTION, WS_SYSMENU, WS_VISIBLE, WS_CHILD, WS_TABSTOP, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME,
    CW_USEDEFAULT, HMENU, WNDCLASSW, WS_EX_CLIENTEDGE,
    BS_DEFPUSHBUTTON, BS_AUTOCHECKBOX, BM_SETCHECK, BM_GETCHECK,
    CB_ADDSTRING, CB_RESETCONTENT, CB_SETCURSEL, CB_GETCURSEL, CB_GETITEMDATA, CB_SETITEMDATA, CBS_DROPDOWNLIST,
    MB_OK, MB_ICONERROR, CB_GETDROPPEDSTATE, GetParent, MSG,
    CREATESTRUCTW, LoadCursorW, IDC_ARROW, WINDOW_STYLE
};
use windows::Win32::UI::Controls::{WC_BUTTON, WC_STATIC, WC_COMBOBOXW, BST_CHECKED};
use windows::Win32::Graphics::Gdi::{HBRUSH, COLOR_WINDOW, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, SetFocus, VK_RETURN, VK_ESCAPE, EnableWindow};
use crate::{with_state, rebuild_menus, apply_word_wrap_to_all_edits};
use crate::settings::{Language, OpenBehavior, VoiceInfo, save_settings, TRUSTED_CLIENT_TOKEN, VOICE_LIST_URL};
use crate::accessibility::{to_wide, handle_accessibility};
use std::thread;

const OPTIONS_CLASS_NAME: &str = "NovapadOptions";
const OPTIONS_ID_LANG: usize = 6001;
const OPTIONS_ID_OPEN: usize = 6002;
const OPTIONS_ID_VOICE: usize = 6003;
const OPTIONS_ID_MULTILINGUAL: usize = 6004;
const OPTIONS_ID_SPLIT_ON_NEWLINE: usize = 6007;
const OPTIONS_ID_WORD_WRAP: usize = 6008;
const OPTIONS_ID_MOVE_CURSOR: usize = 6009;
const OPTIONS_ID_AUDIO_SKIP: usize = 6010;
const OPTIONS_ID_OK: usize = 6005;
const OPTIONS_ID_CANCEL: usize = 6006;

const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = GetFocus();
        if GetParent(focus) == hwnd {
             let dropped = SendMessageW(focus, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 != 0;
             if !dropped {
                 let _ = with_options_state(hwnd, |state| {
                     let _ = SendMessageW(hwnd, WM_COMMAND, WPARAM(OPTIONS_ID_OK | (0 << 16)), LPARAM(state.ok_button.0));
                 });
                 return true;
             }
        }
    }
    handle_accessibility(hwnd, msg)
}

struct OptionsDialogState {
    parent: HWND,
    combo_lang: HWND,
    combo_open: HWND,
    combo_voice: HWND,
    combo_audio_skip: HWND,
    checkbox_multilingual: HWND,
    checkbox_split_on_newline: HWND,
    checkbox_word_wrap: HWND,
    checkbox_move_cursor: HWND,
    ok_button: HWND,
}

struct OptionsLabels {
    title: &'static str,
    label_language: &'static str,
    label_open: &'static str,
    label_voice: &'static str,
    label_multilingual: &'static str,
    label_split_on_newline: &'static str,
    label_word_wrap: &'static str,
    label_move_cursor: &'static str,
    label_audio_skip: &'static str,
    lang_it: &'static str,
    lang_en: &'static str,
    open_new_tab: &'static str,
    open_new_window: &'static str,
    ok: &'static str,
    cancel: &'static str,
    voices_loading: &'static str,
    voices_empty: &'static str,
}

fn options_labels(language: Language) -> OptionsLabels {
    match language {
        Language::Italian => OptionsLabels {
            title: "Opzioni",
            label_language: "Lingua interfaccia:",
            label_open: "Apertura file:",
            label_voice: "Voce:",
            label_multilingual: "Mostra solo voci multilingua",
            label_split_on_newline: "Spezza la lettura quando si va a capo",
            label_word_wrap: "A capo automatico nella finestra",
            label_move_cursor: "Sposta il cursore durante la lettura",
            label_audio_skip: "Spostamento MP3 (frecce):",
            lang_it: "Italiano",
            lang_en: "Inglese",
            open_new_tab: "Apri file in nuovo tab",
            open_new_window: "Apri file in nuova finestra",
            ok: "OK",
            cancel: "Annulla",
            voices_loading: "Caricamento voci...",
            voices_empty: "Nessuna voce disponibile",
        },
        Language::English => OptionsLabels {
            title: "Options",
            label_language: "Interface language:",
            label_open: "Open behavior:",
            label_voice: "Voice:",
            label_multilingual: "Show only multilingual voices",
            label_split_on_newline: "Split reading on new lines",
            label_word_wrap: "Word wrap in editor",
            label_move_cursor: "Move cursor during reading",
            label_audio_skip: "MP3 skip interval:",
            lang_it: "Italian",
            lang_en: "English",
            open_new_tab: "Open files in new tab",
            open_new_window: "Open files in new window",
            ok: "OK",
            cancel: "Cancel",
            voices_loading: "Loading voices...",
            voices_empty: "No voices available",
        },
    }
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.options_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(OPTIONS_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(LoadCursorW(None, IDC_ARROW).unwrap_or_default().0),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(options_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = options_labels(language);
    let title = to_wide(labels.title);

    let dialog = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        520,
        400,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if dialog.0 != 0 {
        let _ = with_state(parent, |state| {
            state.options_dialog = dialog;
        });
        EnableWindow(parent, true);
        SetForegroundWindow(dialog);
        ensure_voice_list_loaded(parent, language);
    }
}

pub unsafe fn refresh_voices(hwnd: HWND) {
    let (parent, combo_voice, checkbox) = match with_options_state(hwnd, |state| {
        (state.parent, state.combo_voice, state.checkbox_multilingual)
    }) {
        Some(values) => values,
        None => return,
    };
    let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let voices = with_state(parent, |state| state.voice_list.clone()).unwrap_or_default();
    let labels = options_labels(settings.language);
    let only_multilingual = SendMessageW(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
        == BST_CHECKED.0;

    populate_voice_combo(combo_voice, &voices, &settings.tts_voice, only_multilingual, &labels);
}

unsafe extern "system" fn options_wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = options_labels(language);

            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));
            
            let label_lang = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_language).as_ptr()), WS_CHILD | WS_VISIBLE, 20, 20, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_lang = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, 18, 300, 120, hwnd, HMENU(OPTIONS_ID_LANG as isize), HINSTANCE(0), None);

            let label_open = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_open).as_ptr()), WS_CHILD | WS_VISIBLE, 20, 60, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_open = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, 58, 300, 120, hwnd, HMENU(OPTIONS_ID_OPEN as isize), HINSTANCE(0), None);

            let label_voice = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_voice).as_ptr()), WS_CHILD | WS_VISIBLE, 20, 100, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_voice = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, 98, 300, 140, hwnd, HMENU(OPTIONS_ID_VOICE as isize), HINSTANCE(0), None);

            let checkbox_multilingual = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_multilingual).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, 138, 300, 20, hwnd, HMENU(OPTIONS_ID_MULTILINGUAL as isize), HINSTANCE(0), None);

            let label_audio_skip = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_audio_skip).as_ptr()), WS_CHILD | WS_VISIBLE, 20, 170, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_audio_skip = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, 168, 300, 140, hwnd, HMENU(OPTIONS_ID_AUDIO_SKIP as isize), HINSTANCE(0), None);

            let checkbox_split_on_newline = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_split_on_newline).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, 202, 300, 20, hwnd, HMENU(OPTIONS_ID_SPLIT_ON_NEWLINE as isize), HINSTANCE(0), None);

            let checkbox_word_wrap = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_word_wrap).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, 226, 300, 20, hwnd, HMENU(OPTIONS_ID_WORD_WRAP as isize), HINSTANCE(0), None);

            let checkbox_move_cursor = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_move_cursor).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, 250, 300, 20, hwnd, HMENU(OPTIONS_ID_MOVE_CURSOR as isize), HINSTANCE(0), None);

            let ok_button = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.ok).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32), 280, 300, 90, 28, hwnd, HMENU(OPTIONS_ID_OK as isize), HINSTANCE(0), None);
            let cancel_button = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.cancel).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 380, 300, 90, 28, hwnd, HMENU(OPTIONS_ID_CANCEL as isize), HINSTANCE(0), None);

            for control in [label_lang, combo_lang, label_open, combo_open, label_voice, combo_voice, label_audio_skip, combo_audio_skip, checkbox_multilingual, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor, ok_button, cancel_button] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let dialog_state = Box::new(OptionsDialogState {
                parent, combo_lang, combo_open, combo_voice, combo_audio_skip, checkbox_multilingual, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor, ok_button
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(dialog_state) as isize);
            initialize_options_dialog(hwnd);
            SetFocus(combo_lang);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            match cmd_id {
                OPTIONS_ID_OK => {
                    apply_options_dialog(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_CANCEL | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_MULTILINGUAL => {
                    refresh_voices(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let is_voice = with_options_state(hwnd, |state| focus == state.combo_voice).unwrap_or(false);
                if is_voice {
                    apply_options_dialog(hwnd);
                    return LRESULT(0);
                }
            } else if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_DESTROY => {
            let parent = with_options_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                EnableWindow(parent, true);
                SetForegroundWindow(parent);
                let _ = with_state(parent, |state| {
                    state.options_dialog = HWND(0);
                });
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_options_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut OptionsDialogState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe fn initialize_options_dialog(hwnd: HWND) {
    let (parent, combo_lang, combo_open, _combo_voice, combo_audio_skip, checkbox_multilingual, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor) = match with_options_state(hwnd, |state| {
        (
            state.parent, state.combo_lang, state.combo_open, state.combo_voice, state.combo_audio_skip,
            state.checkbox_multilingual, state.checkbox_split_on_newline, state.checkbox_word_wrap, state.checkbox_move_cursor
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let labels = options_labels(settings.language);

    let _ = SendMessageW(combo_lang, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(combo_lang, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.lang_it).as_ptr() as isize));
    let _ = SendMessageW(combo_lang, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.lang_en).as_ptr() as isize));
    let lang_index = match settings.language {
        Language::Italian => 0,
        Language::English => 1,
    };
    let _ = SendMessageW(combo_lang, CB_SETCURSEL, WPARAM(lang_index), LPARAM(0));

    let _ = SendMessageW(combo_open, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(combo_open, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.open_new_tab).as_ptr() as isize));
    let _ = SendMessageW(combo_open, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.open_new_window).as_ptr() as isize));
    let open_index = match settings.open_behavior {
        OpenBehavior::NewTab => 0,
        OpenBehavior::NewWindow => 1,
    };
    let _ = SendMessageW(combo_open, CB_SETCURSEL, WPARAM(open_index), LPARAM(0));

    let _ = SendMessageW(checkbox_multilingual, BM_SETCHECK, WPARAM(if settings.tts_only_multilingual { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_split_on_newline, BM_SETCHECK, WPARAM(if settings.split_on_newline { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_word_wrap, BM_SETCHECK, WPARAM(if settings.word_wrap { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_move_cursor, BM_SETCHECK, WPARAM(if settings.move_cursor_during_reading { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));

    let _ = SendMessageW(combo_audio_skip, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let skip_options = [(10, "10 s"), (30, "30 s"), (60, "1 m"), (120, "2 m"), (300, "5 m")];
    let mut selected_idx = 2;
    for (secs, label) in skip_options.iter() {
        let idx = SendMessageW(combo_audio_skip, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(label).as_ptr() as isize)).0 as usize;
        let _ = SendMessageW(combo_audio_skip, CB_SETITEMDATA, WPARAM(idx), LPARAM(*secs as isize));
        if *secs == settings.audiobook_skip_seconds {
            selected_idx = idx;
        }
    }
    let _ = SendMessageW(combo_audio_skip, CB_SETCURSEL, WPARAM(selected_idx), LPARAM(0));

    refresh_voices(hwnd);
}

unsafe fn populate_voice_combo(combo_voice: HWND, voices: &[VoiceInfo], selected: &str, only_multilingual: bool, labels: &OptionsLabels) {
    let _ = SendMessageW(combo_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    if voices.is_empty() {
        let label = if crate::with_state(windows::Win32::UI::WindowsAndMessaging::GetParent(combo_voice), |s| s.voice_list.is_empty()).unwrap_or(true) {
            labels.voices_loading
        } else {
            labels.voices_empty
        };
        let _ = SendMessageW(combo_voice, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(label).as_ptr() as isize));
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        return;
    }
    let mut selected_index: Option<usize> = None;
    let mut combo_index = 0usize;

    for (voice_index, voice) in voices.iter().enumerate() {
        if only_multilingual && !voice.is_multilingual {
            continue;
        }
        let label = format!("{} ({})", voice.short_name, voice.locale);
        let wide = to_wide(&label);
        let idx = SendMessageW(combo_voice, CB_ADDSTRING, WPARAM(0), LPARAM(wide.as_ptr() as isize)).0;
        if idx >= 0 {
            let _ = SendMessageW(combo_voice, CB_SETITEMDATA, WPARAM(idx as usize), LPARAM(voice_index as isize));
            if voice.short_name == selected {
                selected_index = Some(combo_index);
            }
            combo_index += 1;
        }
    }
    
    if let Some(idx) = selected_index {
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
    } else if combo_index > 0 {
         let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
    }
}

unsafe fn apply_options_dialog(hwnd: HWND) {
    let (parent, combo_lang, combo_open, combo_voice, combo_audio_skip, checkbox_multilingual, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor) = match with_options_state(hwnd, |state| {
        (state.parent, state.combo_lang, state.combo_open, state.combo_voice, state.combo_audio_skip, state.checkbox_multilingual, state.checkbox_split_on_newline, state.checkbox_word_wrap, state.checkbox_move_cursor)
    }) { Some(values) => values, None => return };

    let mut settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let old_language = settings.language;
    let old_word_wrap = settings.word_wrap;

    let lang_sel = SendMessageW(combo_lang, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.language = if lang_sel == 1 { Language::English } else { Language::Italian };

    let open_sel = SendMessageW(combo_open, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.open_behavior = if open_sel == 1 { OpenBehavior::NewWindow } else { OpenBehavior::NewTab };

    settings.tts_only_multilingual = SendMessageW(checkbox_multilingual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.split_on_newline = SendMessageW(checkbox_split_on_newline, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.word_wrap = SendMessageW(checkbox_word_wrap, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.move_cursor_during_reading = SendMessageW(checkbox_move_cursor, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;

    let voices = with_state(parent, |state| state.voice_list.clone()).unwrap_or_default();
    let voice_sel = SendMessageW(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel >= 0 {
        let voice_index = SendMessageW(combo_voice, CB_GETITEMDATA, WPARAM(voice_sel as usize), LPARAM(0)).0 as usize;
        if voice_index < voices.len() {
            settings.tts_voice = voices[voice_index].short_name.clone();
        }
    }
    
    let skip_sel = SendMessageW(combo_audio_skip, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if skip_sel >= 0 {
        let skip_secs = SendMessageW(combo_audio_skip, CB_GETITEMDATA, WPARAM(skip_sel as usize), LPARAM(0)).0;
        settings.audiobook_skip_seconds = skip_secs as u32;
    }

    let _ = with_state(parent, |state| { state.settings = settings.clone(); });
    let new_language = settings.language;
    save_settings(settings.clone());

    if old_language != new_language {
        rebuild_menus(parent);
    }
    if old_word_wrap != settings.word_wrap {
        apply_word_wrap_to_all_edits(parent, settings.word_wrap);
    }
    let _ = DestroyWindow(hwnd);
}

fn ensure_voice_list_loaded(hwnd: HWND, language: Language) {
    let has_list = unsafe { with_state(hwnd, |state| !state.voice_list.is_empty()) }.unwrap_or(false);
    if has_list { return; }
    
    thread::spawn(move || {
        match fetch_voice_list() {
            Ok(list) => {
                let payload = Box::new(list);
                let _ = unsafe {
                    windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                        hwnd,
                        WM_TTS_VOICES_LOADED,
                        WPARAM(0),
                        LPARAM(Box::into_raw(payload) as isize),
                    )
                };
            }
            Err(err) => unsafe {
                let msg_str = match language {
                    Language::Italian => format!("Errore nel caricamento delle voci: {err}"),
                    Language::English => format!("Failed to load voices: {err}"),
                };
                let title_str = match language {
                    Language::Italian => "Errore",
                    Language::English => "Error",
                };
                let message = to_wide(&msg_str);
                let title = to_wide(title_str);
                MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONERROR);
            },
        }
    });
}

fn fetch_voice_list() -> Result<Vec<VoiceInfo>, String> {
    let url = format!("{}?trustedclienttoken={}", VOICE_LIST_URL, TRUSTED_CLIENT_TOKEN);
    let resp = reqwest::blocking::get(url).map_err(|err| err.to_string())?;
    let value: serde_json::Value = resp.json().map_err(|err| err.to_string())?;
    let Some(voices) = value.as_array() else {
        return Err("Risposta non valida".to_string());
    };

    let mut results = Vec::new();
    for voice in voices {
        let short_name = voice["ShortName"].as_str().unwrap_or("").to_string();
        if short_name.is_empty() { continue; }
        let locale = voice["Locale"].as_str().unwrap_or("").to_string();
        let is_multilingual = short_name.contains("Multilingual");
        results.push(VoiceInfo { short_name, locale, is_multilingual });
    }
    results.sort_by(|a, b| a.short_name.cmp(&b.short_name));
    Ok(results)
}