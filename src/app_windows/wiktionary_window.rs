use crate::accessibility::{handle_accessibility, to_wide};
use crate::i18n;
use crate::log_debug;
use crate::settings::Language;
use crate::update_dictionary_cache;
use crate::wiktionary;
use crate::with_state;
use std::sync::atomic::{AtomicUsize, Ordering};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, SetFocus, VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    ES_AUTOHSCROLL, ES_AUTOVSCROLL, ES_MULTILINE, ES_READONLY, GWLP_USERDATA, GetDlgCtrlID,
    GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW, IsWindow,
    LoadCursorW, MSG, PostMessageW, RegisterClassW, SetForegroundWindow, SetWindowLongPtrW,
    SetWindowTextW, WINDOW_STYLE, WM_APP, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN,
    WM_NCDESTROY, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_POPUP, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const WIKTIONARY_CLASS_NAME: &str = "NovapadWiktionary";
const WIKTIONARY_INPUT_ID: usize = 9401;
const WIKTIONARY_SEARCH_ID: usize = 9402;
const WIKTIONARY_OUTPUT_ID: usize = 9403;
const WIKTIONARY_CLOSE_ID: usize = 9404;

const WM_LOOKUP_DONE: u32 = WM_APP + 100;
static LOOKUP_GENERATION: AtomicUsize = AtomicUsize::new(0);

struct WiktionaryWindowState {
    parent: HWND,
    input: HWND,
    output: HWND,
    search: HWND,
    close: HWND,
}

struct WiktionaryLabels {
    title: String,
    word: String,
    search: String,
    results: String,
    close: String,
}

fn wiktionary_labels(language: crate::settings::Language) -> WiktionaryLabels {
    WiktionaryLabels {
        title: i18n::tr(language, "dictionary.lookup.title"),
        word: i18n::tr(language, "dictionary.lookup.word"),
        search: i18n::tr(language, "dictionary.lookup.search"),
        results: i18n::tr(language, "dictionary.lookup.results"),
        close: i18n::tr(language, "dictionary.lookup.close"),
    }
}

fn dictionary_cache_key(language: Language, pref: &str, word: &str) -> String {
    let lang = match language {
        Language::Italian => "it",
        Language::English => "en",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::Vietnamese => "vi",
    };
    format!(
        "{}|{}|{}",
        lang,
        pref.trim().to_ascii_lowercase(),
        word.trim().to_ascii_lowercase()
    )
}

fn translation_label_parts(language: Language) -> (String, String) {
    let sample = i18n::tr_f(language, "dictionary.translation_label", &[("lang", "XX")]);
    if let Some(pos) = sample.find("XX") {
        let prefix = sample[..pos].to_string();
        let suffix = sample[pos + 2..].to_string();
        (prefix, suffix)
    } else {
        (sample, String::new())
    }
}

fn format_cached_output(language: Language, lines: &[String]) -> String {
    let definitions_label = i18n::tr(language, "dictionary.definitions");
    let synonyms_label = i18n::tr(language, "dictionary.synonyms");
    let (trans_prefix, trans_suffix) = translation_label_parts(language);

    let mut out = String::new();
    let mut section = "title";
    let mut def_index = 1usize;
    let mut trans_index = 1usize;

    for line in lines {
        if line == &definitions_label {
            if !out.is_empty() {
                out.push('\n');
            }
            out.push_str(&definitions_label);
            out.push_str(":\n");
            section = "definitions";
            def_index = 1;
            continue;
        }
        if line == &synonyms_label {
            if !out.is_empty() {
                out.push('\n');
            }
            out.push_str(&synonyms_label);
            out.push_str(":\n");
            section = "synonyms";
            continue;
        }
        let is_translation = line.starts_with(&trans_prefix) && line.ends_with(&trans_suffix);
        if is_translation {
            if !out.is_empty() {
                out.push('\n');
            }
            out.push_str(line);
            out.push('\n');
            section = "translation";
            trans_index = 1;
            continue;
        }

        match section {
            "definitions" => {
                out.push_str(&format!("{def_index}. {line}\n"));
                def_index += 1;
            }
            "translation" => {
                out.push_str(&format!("{trans_index}. {line}\n"));
                trans_index += 1;
            }
            "synonyms" => {
                out.push_str(line);
                out.push('\n');
            }
            _ => {
                out.push_str(line);
                out.push_str("\n\n");
                section = "body";
            }
        }
    }

    out.trim_end().to_string()
}

fn to_windows_newlines(text: &str) -> String {
    text.replace("\n", "\r\n")
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    handle_accessibility(hwnd, msg)
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.wiktionary_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(WIKTIONARY_CLASS_NAME);
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = wiktionary_labels(language);
    let title = to_wide(&labels.title);

    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(wiktionary_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let state = Box::new(WiktionaryWindowState {
        parent,
        input: HWND(0),
        output: HWND(0),
        search: HWND(0),
        close: HWND(0),
    });
    let state_ptr = Box::into_raw(state);
    let hwnd = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        520,
        420,
        parent,
        HMENU(0),
        hinstance,
        Some(state_ptr as *const _),
    );
    if hwnd.0 == 0 {
        drop(Box::from_raw(state_ptr));
        return;
    }
    with_state(parent, |state| state.wiktionary_window = hwnd);
}

unsafe extern "system" fn wiktionary_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*cs).lpCreateParams as *mut WiktionaryWindowState;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, init_ptr as isize);

            let parent = (*init_ptr).parent;
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = wiktionary_labels(language);

            let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&labels.word).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                12,
                14,
                120,
                18,
                hwnd,
                HMENU(1),
                hinstance,
                None,
            );
            let input = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                12,
                34,
                360,
                24,
                hwnd,
                HMENU(WIKTIONARY_INPUT_ID as isize),
                hinstance,
                None,
            );
            let search = CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&labels.search).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                382,
                32,
                110,
                28,
                hwnd,
                HMENU(WIKTIONARY_SEARCH_ID as isize),
                hinstance,
                None,
            );
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&labels.results).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                12,
                68,
                120,
                18,
                hwnd,
                HMENU(2),
                hinstance,
                None,
            );
            let output = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WINDOW_STYLE(ES_MULTILINE as u32)
                    | WINDOW_STYLE(ES_AUTOVSCROLL as u32)
                    | WINDOW_STYLE(ES_READONLY as u32),
                12,
                88,
                480,
                260,
                hwnd,
                HMENU(WIKTIONARY_OUTPUT_ID as isize),
                hinstance,
                None,
            );
            let close = CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&labels.close).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                412,
                360,
                80,
                28,
                hwnd,
                HMENU(WIKTIONARY_CLOSE_ID as isize),
                hinstance,
                None,
            );

            (*init_ptr).input = input;
            (*init_ptr).output = output;
            (*init_ptr).search = search;
            (*init_ptr).close = close;
            let proc_ptr = tab_subclass_proc as usize;
            for control in [input, search, output, close] {
                let prev = SetWindowLongPtrW(
                    control,
                    windows::Win32::UI::WindowsAndMessaging::GWLP_WNDPROC,
                    proc_ptr as isize,
                );
                SetWindowLongPtrW(control, GWLP_USERDATA, prev);
            }
            SetFocus(input);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                crate::log_if_err!(DestroyWindow(hwnd));
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let Some((search, close)) =
                    with_window_state(hwnd, |state| (state.search, state.close))
                else {
                    return DefWindowProcW(hwnd, msg, wparam, lparam);
                };
                if focus == close {
                    crate::log_if_err!(DestroyWindow(hwnd));
                    return LRESULT(0);
                }
                if focus == search {
                    run_lookup(hwnd);
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_COMMAND => {
            let id = wparam.0 & 0xffff;
            if id == WIKTIONARY_CLOSE_ID || id == 2 {
                crate::log_if_err!(DestroyWindow(hwnd));
                return LRESULT(0);
            }
            if id == WIKTIONARY_SEARCH_ID || id == 1 {
                run_lookup(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_LOOKUP_DONE => {
            let generation = wparam.0;
            let current_gen = LOOKUP_GENERATION.load(Ordering::SeqCst);
            if generation != current_gen {
                return LRESULT(0);
            }
            let result_ptr = lparam.0 as *mut String;
            if !result_ptr.is_null() {
                let text = Box::from_raw(result_ptr);
                if let Some(output) = with_window_state(hwnd, |state| state.output)
                    && let Err(_e) = SetWindowTextW(output, PCWSTR(to_wide(&text).as_ptr()))
                {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
            }
            LRESULT(0)
        }
        WM_DESTROY => LRESULT(0),
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WiktionaryWindowState;
            if !ptr.is_null() {
                let state = Box::from_raw(ptr);
                with_state(state.parent, |s| s.wiktionary_window = HWND(0));
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn handle_enter_key(hwnd: HWND) -> bool {
    let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
    if parent.0 == 0 {
        return false;
    }
    let id = GetDlgCtrlID(hwnd) as usize;
    if id == WIKTIONARY_CLOSE_ID {
        crate::log_if_err!(DestroyWindow(parent));
        return true;
    }
    if id == WIKTIONARY_SEARCH_ID {
        run_lookup(parent);
        return true;
    }
    if id == WIKTIONARY_INPUT_ID {
        run_lookup(parent);
        return true;
    }
    false
}

unsafe extern "system" fn tab_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_KEYDOWN {
        let key = wparam.0 as u16;
        if key == windows::Win32::UI::Input::KeyboardAndMouse::VK_TAB.0 {
            let shift_down = windows::Win32::UI::Input::KeyboardAndMouse::GetKeyState(
                windows::Win32::UI::Input::KeyboardAndMouse::VK_SHIFT.0 as i32,
            ) & 0x8000u16 as i16
                != 0;
            let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            if parent.0 != 0 {
                focus_next_control(parent, hwnd, shift_down);
                return LRESULT(0);
            }
        }
        if key == windows::Win32::UI::Input::KeyboardAndMouse::VK_RETURN.0 && handle_enter_key(hwnd)
        {
            return LRESULT(0);
        }
    }
    if msg == windows::Win32::UI::WindowsAndMessaging::WM_CHAR
        && wparam.0 == 13
        && handle_enter_key(hwnd)
    {
        return LRESULT(0);
    }
    if msg == windows::Win32::UI::WindowsAndMessaging::WM_GETDLGCODE {
        let id = GetDlgCtrlID(hwnd) as usize;
        if id == WIKTIONARY_CLOSE_ID || id == WIKTIONARY_SEARCH_ID {
            return LRESULT(windows::Win32::UI::WindowsAndMessaging::DLGC_WANTALLKEYS as isize);
        }
    }
    let prev = windows::Win32::UI::WindowsAndMessaging::GetWindowLongPtrW(
        hwnd,
        windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
    );
    if prev == 0 {
        return DefWindowProcW(hwnd, msg, wparam, lparam);
    }
    windows::Win32::UI::WindowsAndMessaging::CallWindowProcW(
        Some(std::mem::transmute::<
            isize,
            unsafe extern "system" fn(HWND, u32, WPARAM, LPARAM) -> LRESULT,
        >(prev)),
        hwnd,
        msg,
        wparam,
        lparam,
    )
}

unsafe fn focus_next_control(parent: HWND, current: HWND, shift_down: bool) {
    let order = with_window_state(parent, |state| {
        vec![state.input, state.search, state.output, state.close]
    })
    .unwrap_or_default();
    if order.is_empty() {
        return;
    }
    let Some(pos) = order.iter().position(|hwnd| *hwnd == current) else {
        return;
    };
    let next_index = if shift_down {
        if pos == 0 { order.len() - 1 } else { pos - 1 }
    } else if pos + 1 >= order.len() {
        0
    } else {
        pos + 1
    };
    let target = order[next_index];
    if target.0 != 0 {
        SetFocus(target);
    }
}

unsafe fn with_window_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut WiktionaryWindowState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut WiktionaryWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe fn run_lookup(hwnd: HWND) {
    let Some((parent, input, output)) =
        with_window_state(hwnd, |state| (state.parent, state.input, state.output))
    else {
        return;
    };
    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    let pref = with_state(parent, |s| {
        s.settings.dictionary_translation_language.clone()
    })
    .unwrap_or_else(|| "auto".to_string());

    let len = GetWindowTextLengthW(input);
    if len <= 0 {
        let msg = i18n::tr(language, "dictionary.no_word");
        if let Err(_e) =
            SetWindowTextW(output, PCWSTR(to_wide(&to_windows_newlines(&msg)).as_ptr()))
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        return;
    }
    let mut buf = vec![0u16; (len + 1) as usize];
    let _read = GetWindowTextW(input, &mut buf);
    let word = String::from_utf16_lossy(&buf[..len as usize]);
    let trimmed = word.trim().to_string();
    if trimmed.is_empty() {
        let msg = i18n::tr(language, "dictionary.no_word");
        if let Err(_e) =
            SetWindowTextW(output, PCWSTR(to_wide(&to_windows_newlines(&msg)).as_ptr()))
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        return;
    }

    SetFocus(output);

    let key = dictionary_cache_key(language, &pref, &trimmed);
    let cached_lines =
        with_state(parent, |state| state.dictionary_cache.get(&key).cloned()).unwrap_or(None);
    if let Some(lines) = cached_lines {
        let text = format_cached_output(language, &lines);
        if let Err(_e) = SetWindowTextW(
            output,
            PCWSTR(to_wide(&to_windows_newlines(&text)).as_ptr()),
        ) {}
        return;
    }

    let generation = LOOKUP_GENERATION
        .fetch_add(1, Ordering::SeqCst)
        .wrapping_add(1);
    LOOKUP_GENERATION.store(generation, Ordering::SeqCst);

    let loading_msg = i18n::tr(language, "dictionary.loading");
    if let Err(_e) = SetWindowTextW(
        output,
        PCWSTR(to_wide(&to_windows_newlines(&loading_msg)).as_ptr()),
    ) {}

    let hwnd_val = hwnd.0;
    let parent_hwnd = parent;
    let cache_key = key.clone();
    std::thread::spawn(move || {
        let result = wiktionary::lookup_for_language_with_meta(&trimmed, language, &pref);
        let text = match result {
            Ok((entry, _is_large)) => {
                let lines = wiktionary::format_menu_lines(language, &entry);
                update_dictionary_cache(parent_hwnd, cache_key.clone(), lines.clone());
                wiktionary::format_output_text(language, &entry)
            }
            Err(wiktionary::LookupError::NotFound { .. }) => {
                i18n::tr(language, "dictionary.not_found")
            }
            Err(err) => {
                log_debug(&format!("Dictionary lookup failed: {err}"));
                i18n::tr(language, "dictionary.not_found")
            }
        };
        let text = to_windows_newlines(&text);

        let text_ptr = Box::into_raw(Box::new(text));
        let hwnd = HWND(hwnd_val);
        unsafe {
            if IsWindow(hwnd).as_bool() {
                if let Err(e) = PostMessageW(
                    hwnd,
                    WM_LOOKUP_DONE,
                    WPARAM(generation),
                    LPARAM(text_ptr as isize),
                ) {
                    crate::log_debug(&format!("Failed to post WM_LOOKUP_DONE: {}", e));
                }
            } else {
                drop(Box::from_raw(text_ptr));
            }
        }
    });
}
