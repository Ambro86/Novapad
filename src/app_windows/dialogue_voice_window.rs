use crate::dialogue_voice::DialogueVoiceConfig;
use crate::settings::{Language, TtsEngine, VoiceInfo};
use crate::{i18n, to_wide};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    BST_CHECKED, BST_UNCHECKED, WC_BUTTON, WC_COMBOBOXW, WC_EDIT, WC_STATIC,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus};
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, CB_ADDSTRING, CB_GETCURSEL, CB_GETITEMDATA,
    CB_RESETCONTENT, CB_SETCURSEL, CB_SETITEMDATA, CBN_SELCHANGE, CBS_DROPDOWN, CBS_DROPDOWNLIST,
    CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, GWLP_USERDATA,
    GetDlgItem, GetMessageW, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU,
    IDC_ARROW, IsDialogMessageW, IsWindow, MSG, RegisterClassW, SW_HIDE, SW_SHOW, SendMessageW,
    SetWindowLongPtrW, ShowWindow, TranslateMessage, WINDOW_EX_STYLE, WINDOW_STYLE, WM_CLOSE,
    WM_COMMAND, WM_CREATE, WM_KEYDOWN, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::PCWSTR;

const ID_ENGINE: i32 = 1001;
const ID_LANGUAGE_LABEL: i32 = 1101;
const ID_LANGUAGE: i32 = 1002;
const ID_VOICE: i32 = 1003;
const ID_ONLY_MULTILINGUAL: i32 = 1004;
const ID_RATE: i32 = 1005;
const ID_PITCH: i32 = 1006;
const ID_VOLUME: i32 = 1007;
const ID_OPEN_QUOTE: i32 = 1008;
const ID_CLOSE_QUOTE: i32 = 1009;
const ID_ALLOW_MULTILINE: i32 = 1010;
const ID_OK: i32 = 1;
const ID_CANCEL: i32 = 2;

struct DialogueVoiceDialogData {
    language: Language,
    edge_voices: Vec<VoiceInfo>,
    sapi5_voices: Vec<VoiceInfo>,
    sapi4_voices: Vec<VoiceInfo>,
    default_engine: TtsEngine,
    default_voice: String,
    default_rate: i32,
    default_pitch: i32,
    default_volume: i32,
    default_open_quote: String,
    default_close_quote: String,
    default_allow_multiline: bool,
    edge_language_codes: Vec<String>,
    result: Option<DialogueVoiceConfig>,
}

fn read_control_text(hwnd: HWND) -> String {
    unsafe {
        let len = GetWindowTextLengthW(hwnd);
        if len <= 0 {
            return String::new();
        }
        let mut buf = vec![0u16; (len + 1) as usize];
        let read = GetWindowTextW(hwnd, &mut buf);
        String::from_utf16_lossy(&buf[..read as usize])
    }
}

fn voice_locale_language_code(locale: &str) -> Option<String> {
    let base = locale.split(['-', '_']).next()?.trim();
    if base.is_empty() {
        return None;
    }
    Some(base.to_ascii_lowercase())
}

fn localized_voice_language_name(language: Language, code: &str) -> String {
    let key = format!("voice.lang.{}", code);
    let localized = i18n::tr(language, &key);
    if localized != key {
        return localized;
    }
    code.to_ascii_uppercase()
}

fn collect_voice_language_codes(voices: &[VoiceInfo]) -> Vec<String> {
    let mut codes: Vec<String> = voices
        .iter()
        .filter_map(|v| voice_locale_language_code(&v.locale))
        .collect();
    codes.sort();
    codes.dedup();
    codes
}

fn fill_value_combo(hwnd_combo: HWND, items: &[(String, i32)], selected: i32) {
    unsafe {
        SendMessageW(hwnd_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let mut selected_idx = 0usize;
        let mut best_distance = i32::MAX;
        for (idx, (label, value)) in items.iter().enumerate() {
            let w = to_wide(label);
            SendMessageW(
                hwnd_combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(w.as_ptr() as isize),
            );
            SendMessageW(
                hwnd_combo,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(*value as isize),
            );
            let distance = (*value - selected).abs();
            if distance < best_distance {
                best_distance = distance;
                selected_idx = idx;
            }
        }
        if !items.is_empty() {
            SendMessageW(hwnd_combo, CB_SETCURSEL, WPARAM(selected_idx), LPARAM(0));
        }
    }
}

fn selected_combo_value(hwnd_combo: HWND, fallback: i32) -> i32 {
    unsafe {
        let sel = SendMessageW(hwnd_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if sel < 0 {
            return fallback;
        }
        let data = SendMessageW(hwnd_combo, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0;
        if data == -1 { fallback } else { data as i32 }
    }
}

fn fill_engine_combo(hwnd_combo: HWND, language: Language, selected: TtsEngine) {
    unsafe {
        SendMessageW(hwnd_combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let items = [
            (TtsEngine::Edge, i18n::tr(language, "options.engine.edge")),
            (TtsEngine::Sapi5, i18n::tr(language, "options.engine.sapi5")),
            (TtsEngine::Sapi4, i18n::tr(language, "options.engine.sapi4")),
        ];
        let mut selected_idx = 0usize;
        for (i, (engine, label)) in items.iter().enumerate() {
            let w = to_wide(label);
            let idx = SendMessageW(
                hwnd_combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(w.as_ptr() as isize),
            )
            .0 as usize;
            SendMessageW(
                hwnd_combo,
                CB_SETITEMDATA,
                WPARAM(idx),
                LPARAM(match engine {
                    TtsEngine::Edge => 0,
                    TtsEngine::Sapi5 => 1,
                    TtsEngine::Sapi4 => 2,
                }),
            );
            if *engine == selected {
                selected_idx = i;
            }
        }
        SendMessageW(hwnd_combo, CB_SETCURSEL, WPARAM(selected_idx), LPARAM(0));
    }
}

fn selected_engine(hwnd_combo: HWND) -> TtsEngine {
    unsafe {
        let sel = SendMessageW(hwnd_combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if sel < 0 {
            return TtsEngine::Edge;
        }
        let data = SendMessageW(hwnd_combo, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0;
        match data {
            1 => TtsEngine::Sapi5,
            2 => TtsEngine::Sapi4,
            _ => TtsEngine::Edge,
        }
    }
}

fn fill_voice_combo(hwnd_dialog: HWND, preferred_voice: &str) {
    let combo = unsafe { GetDlgItem(hwnd_dialog, ID_VOICE) };
    if combo.0 == 0 {
        return;
    }
    let ptr = unsafe { GetWindowLongPtrW(hwnd_dialog, GWLP_USERDATA) as *mut DialogueVoiceDialogData };
    if ptr.is_null() {
        return;
    }
    let data = unsafe { &mut *ptr };
    let engine_combo = unsafe { GetDlgItem(hwnd_dialog, ID_ENGINE) };
    let engine = selected_engine(engine_combo);
    let voices = match engine {
        TtsEngine::Edge => &data.edge_voices,
        TtsEngine::Sapi5 => &data.sapi5_voices,
        TtsEngine::Sapi4 => &data.sapi4_voices,
    };
    let only_multilingual = unsafe {
        SendMessageW(
        unsafe { GetDlgItem(hwnd_dialog, ID_ONLY_MULTILINGUAL) },
        BM_GETCHECK,
        WPARAM(0),
        LPARAM(0),
    )
    .0 as u32
    }
        == BST_CHECKED.0;
    let language_filter = if engine == TtsEngine::Edge && !only_multilingual {
        let sel = unsafe {
            SendMessageW(
            unsafe { GetDlgItem(hwnd_dialog, ID_LANGUAGE) },
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
        .0
        };
        if sel >= 0 {
            data.edge_language_codes.get(sel as usize).cloned()
        } else {
            None
        }
    } else {
        None
    };
    unsafe { SendMessageW(combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0)); }
    let mut selected_idx = 0usize;
    let mut combo_index = 0usize;
    for (i, voice) in voices.iter().enumerate() {
        if engine == TtsEngine::Edge && only_multilingual && !voice.is_multilingual {
            continue;
        }
        if let Some(filter) = language_filter.as_deref() {
            let Some(code) = voice_locale_language_code(&voice.locale) else {
                continue;
            };
            if code != filter {
                continue;
            }
        }
        let label = if voice.locale.trim().is_empty() {
            voice.short_name.clone()
        } else {
            format!("{} ({})", voice.short_name, voice.locale)
        };
        let w = to_wide(&label);
        let idx = unsafe {
            SendMessageW(combo, CB_ADDSTRING, WPARAM(0), LPARAM(w.as_ptr() as isize)).0 as usize
        };
        unsafe { SendMessageW(combo, CB_SETITEMDATA, WPARAM(idx), LPARAM(i as isize)); }
        if voice.short_name.eq_ignore_ascii_case(preferred_voice) {
            selected_idx = combo_index;
        }
        combo_index += 1;
    }
    if combo_index > 0 {
        unsafe { SendMessageW(combo, CB_SETCURSEL, WPARAM(selected_idx), LPARAM(0)); }
    }
}

fn selected_voice(hwnd_dialog: HWND) -> String {
    let combo = unsafe { GetDlgItem(hwnd_dialog, ID_VOICE) };
    if combo.0 == 0 {
        return String::new();
    }
    let ptr = unsafe { GetWindowLongPtrW(hwnd_dialog, GWLP_USERDATA) as *mut DialogueVoiceDialogData };
    if ptr.is_null() {
        return String::new();
    }
    let data = unsafe { &*ptr };
    let engine = selected_engine(unsafe { GetDlgItem(hwnd_dialog, ID_ENGINE) });
    let voices = match engine {
        TtsEngine::Edge => &data.edge_voices,
        TtsEngine::Sapi5 => &data.sapi5_voices,
        TtsEngine::Sapi4 => &data.sapi4_voices,
    };
    let sel = unsafe { SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
    if sel < 0 {
        return String::new();
    }
    let idx = unsafe { SendMessageW(combo, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as usize };
    voices
        .get(idx)
        .map(|v| v.short_name.clone())
        .unwrap_or_default()
}

fn refresh_edge_controls(hwnd_dialog: HWND, preferred_voice: &str) {
    let ptr = unsafe { GetWindowLongPtrW(hwnd_dialog, GWLP_USERDATA) as *mut DialogueVoiceDialogData };
    if ptr.is_null() {
        return;
    }
    let data = unsafe { &mut *ptr };
    let engine = selected_engine(unsafe { GetDlgItem(hwnd_dialog, ID_ENGINE) });
    let label_language = unsafe { GetDlgItem(hwnd_dialog, ID_LANGUAGE_LABEL) };
    let combo_language = unsafe { GetDlgItem(hwnd_dialog, ID_LANGUAGE) };
    let check_multilingual = unsafe { GetDlgItem(hwnd_dialog, ID_ONLY_MULTILINGUAL) };
    let is_edge = engine == TtsEngine::Edge;
    unsafe {
        ShowWindow(check_multilingual, if is_edge { SW_SHOW } else { SW_HIDE });
        EnableWindow(check_multilingual, is_edge);
    }

    let only_multilingual = unsafe {
        SendMessageW(check_multilingual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
    }
        == BST_CHECKED.0;
    let show_language = is_edge && !only_multilingual;
    unsafe {
        ShowWindow(
            label_language,
            if show_language { SW_SHOW } else { SW_HIDE },
        );
        ShowWindow(
            combo_language,
            if show_language { SW_SHOW } else { SW_HIDE },
        );
        EnableWindow(combo_language, show_language);
    }

    if show_language {
        let previous_selection = {
            let sel = unsafe { SendMessageW(combo_language, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0 };
            if sel >= 0 {
                data.edge_language_codes.get(sel as usize).cloned()
            } else {
                None
            }
        };
        let mut codes = collect_voice_language_codes(&data.edge_voices);
        if !codes.is_empty() {
            let selected_from_voice = data
                .edge_voices
                .iter()
                .find(|v| v.short_name == preferred_voice)
                .and_then(|v| voice_locale_language_code(&v.locale));
            let selected_code = previous_selection
                .filter(|code| codes.contains(code))
                .or(selected_from_voice.filter(|code| codes.contains(code)))
                .unwrap_or_else(|| codes[0].clone());
            SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
            let mut selected_idx = 0usize;
            for (idx, code) in codes.iter().enumerate() {
                let label = localized_voice_language_name(data.language, code);
                SendMessageW(
                    combo_language,
                    CB_ADDSTRING,
                    WPARAM(0),
                    LPARAM(to_wide(&label).as_ptr() as isize),
                );
                if *code == selected_code {
                    selected_idx = idx;
                }
            }
            SendMessageW(
                combo_language,
                CB_SETCURSEL,
                WPARAM(selected_idx),
                LPARAM(0),
            );
        } else {
            SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        }
        data.edge_language_codes = std::mem::take(&mut codes);
    } else {
        SendMessageW(combo_language, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        data.edge_language_codes.clear();
    }

    fill_voice_combo(hwnd_dialog, preferred_voice);
}

fn tts_rate_items(language: Language) -> Vec<(String, i32)> {
    vec![
        (i18n::tr(language, "tts_tuning.speed.extremely_slow"), -100),
        (i18n::tr(language, "tts_tuning.speed.very_slow"), -60),
        (i18n::tr(language, "tts_tuning.speed.slow"), -35),
        (i18n::tr(language, "tts_tuning.speed.a_bit_slow"), -20),
        (i18n::tr(language, "tts_tuning.speed.slightly_slow"), -10),
        (i18n::tr(language, "tts_tuning.speed.normal"), 0),
        (i18n::tr(language, "tts_tuning.speed.slightly_fast"), 10),
        (i18n::tr(language, "tts_tuning.speed.a_bit_fast"), 20),
        (i18n::tr(language, "tts_tuning.speed.fast"), 35),
        (i18n::tr(language, "tts_tuning.speed.very_fast"), 50),
        (i18n::tr(language, "tts_tuning.speed.super_fast"), 100),
    ]
}

fn tts_pitch_items(language: Language) -> Vec<(String, i32)> {
    vec![
        (i18n::tr(language, "tts_tuning.pitch.very_low"), -12),
        (i18n::tr(language, "tts_tuning.pitch.low"), -10),
        (i18n::tr(language, "tts_tuning.pitch.a_bit_low"), -7),
        (i18n::tr(language, "tts_tuning.pitch.slightly_low"), -5),
        (i18n::tr(language, "tts_tuning.pitch.a_little_lower"), -2),
        (i18n::tr(language, "tts_tuning.pitch.normal"), 0),
        (i18n::tr(language, "tts_tuning.pitch.a_little_higher"), 2),
        (i18n::tr(language, "tts_tuning.pitch.slightly_high"), 5),
        (i18n::tr(language, "tts_tuning.pitch.a_bit_high"), 7),
        (i18n::tr(language, "tts_tuning.pitch.high"), 9),
        (i18n::tr(language, "tts_tuning.pitch.very_high"), 12),
    ]
}

fn tts_volume_items(language: Language) -> Vec<(String, i32)> {
    vec![
        (i18n::tr(language, "tts_tuning.volume.very_low"), 25),
        (i18n::tr(language, "tts_tuning.volume.low"), 40),
        (i18n::tr(language, "tts_tuning.volume.a_bit_low"), 55),
        (i18n::tr(language, "tts_tuning.volume.medium_low"), 70),
        (i18n::tr(language, "tts_tuning.volume.slightly_low"), 85),
        (i18n::tr(language, "tts_tuning.volume.normal"), 100),
        (i18n::tr(language, "tts_tuning.volume.slightly_high"), 115),
        (i18n::tr(language, "tts_tuning.volume.medium_high"), 130),
        (i18n::tr(language, "tts_tuning.volume.a_bit_high"), 145),
        (i18n::tr(language, "tts_tuning.volume.high"), 160),
        (i18n::tr(language, "tts_tuning.volume.very_high"), 180),
        (i18n::tr(language, "tts_tuning.volume.maximum"), 200),
    ]
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    crate::panic_guard::guard(
        "dialogue_voice_window_wndproc",
        || DefWindowProcW(hwnd, msg, wparam, lparam),
        || wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const windows::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*cs).lpCreateParams as *mut DialogueVoiceDialogData;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            if ptr.is_null() {
                return LRESULT(0);
            }
            let data = &*ptr;
            let lang = data.language;
            let y0 = 16;
            let lx = 16;
            let cx = 210;
            let wlabel = 180;
            let wcombo = 320;
            let h = 24;
            let gap = 34;
            let mut y = y0;

            let mk_label = |text: String, yy: i32| {
                CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    WC_STATIC,
                    PCWSTR(to_wide(&text).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    lx,
                    yy,
                    wlabel,
                    h,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };
            let mk_combo = |id: i32, yy: i32, style: WINDOW_STYLE| {
                CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_COMBOBOXW,
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | style,
                    cx,
                    yy,
                    wcombo,
                    300,
                    hwnd,
                    HMENU(id as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            mk_label(i18n::tr(lang, "options.label.tts_engine"), y);
            mk_combo(ID_ENGINE, y, WINDOW_STYLE(CBS_DROPDOWNLIST as u32));
            y += gap;
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                WC_STATIC,
                PCWSTR(to_wide(&i18n::tr(lang, "options.label.voice_language")).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                lx,
                y,
                wlabel,
                h,
                hwnd,
                HMENU(ID_LANGUAGE_LABEL as isize),
                HINSTANCE(0),
                None,
            );
            mk_combo(ID_LANGUAGE, y, WINDOW_STYLE(CBS_DROPDOWNLIST as u32));
            y += gap;
            mk_label(i18n::tr(lang, "options.label.dialogue_voice"), y);
            mk_combo(ID_VOICE, y, WINDOW_STYLE(CBS_DROPDOWNLIST as u32));
            y += gap;
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(lang, "options.label.multilingual")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                cx,
                y,
                wcombo,
                h,
                hwnd,
                HMENU(ID_ONLY_MULTILINGUAL as isize),
                HINSTANCE(0),
                None,
            );
            y += gap;
            mk_label(i18n::tr(lang, "tts_tuning.label_speed"), y);
            mk_combo(ID_RATE, y, WINDOW_STYLE(CBS_DROPDOWN as u32));
            y += gap;
            mk_label(i18n::tr(lang, "tts_tuning.label_pitch"), y);
            mk_combo(ID_PITCH, y, WINDOW_STYLE(CBS_DROPDOWN as u32));
            y += gap;
            mk_label(i18n::tr(lang, "tts_tuning.label_volume"), y);
            mk_combo(ID_VOLUME, y, WINDOW_STYLE(CBS_DROPDOWN as u32));
            y += gap;
            mk_label(i18n::tr(lang, "dialogue_voice.apply.open_quote_title"), y);
            CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR(to_wide(&data.default_open_quote).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                cx,
                y,
                wcombo,
                h,
                hwnd,
                HMENU(ID_OPEN_QUOTE as isize),
                HINSTANCE(0),
                None,
            );
            y += gap;
            mk_label(i18n::tr(lang, "dialogue_voice.apply.close_quote_title"), y);
            CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR(to_wide(&data.default_close_quote).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                cx,
                y,
                wcombo,
                h,
                hwnd,
                HMENU(ID_CLOSE_QUOTE as isize),
                HINSTANCE(0),
                None,
            );
            y += gap;
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(lang, "dialogue_voice.apply.multiline_body")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                cx,
                y,
                wcombo,
                h,
                hwnd,
                HMENU(ID_ALLOW_MULTILINE as isize),
                HINSTANCE(0),
                None,
            );
            y += gap + 8;
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(lang, "options.ok")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                cx + wcombo - 190,
                y,
                90,
                28,
                hwnd,
                HMENU(ID_OK as isize),
                HINSTANCE(0),
                None,
            );
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                WC_BUTTON,
                PCWSTR(to_wide(&i18n::tr(lang, "options.cancel")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                cx + wcombo - 95,
                y,
                90,
                28,
                hwnd,
                HMENU(ID_CANCEL as isize),
                HINSTANCE(0),
                None,
            );

            let engine_combo = GetDlgItem(hwnd, ID_ENGINE);
            fill_engine_combo(engine_combo, lang, data.default_engine);
            SendMessageW(
                GetDlgItem(hwnd, ID_ONLY_MULTILINGUAL),
                BM_SETCHECK,
                WPARAM(BST_UNCHECKED.0 as usize),
                LPARAM(0),
            );
            refresh_edge_controls(hwnd, &data.default_voice);
            fill_value_combo(
                GetDlgItem(hwnd, ID_RATE),
                &tts_rate_items(lang),
                data.default_rate,
            );
            fill_value_combo(
                GetDlgItem(hwnd, ID_PITCH),
                &tts_pitch_items(lang),
                data.default_pitch,
            );
            fill_value_combo(
                GetDlgItem(hwnd, ID_VOLUME),
                &tts_volume_items(lang),
                data.default_volume,
            );
            SendMessageW(
                GetDlgItem(hwnd, ID_ALLOW_MULTILINE),
                BM_SETCHECK,
                WPARAM(if data.default_allow_multiline {
                    BST_CHECKED.0 as usize
                } else {
                    BST_UNCHECKED.0 as usize
                }),
                LPARAM(0),
            );
            SetFocus(engine_combo);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as i32;
            let code = ((wparam.0 >> 16) & 0xffff) as u16;
            if id == ID_ENGINE && code as u32 == CBN_SELCHANGE {
                refresh_edge_controls(hwnd, "");
                return LRESULT(0);
            }
            if id == ID_LANGUAGE && code as u32 == CBN_SELCHANGE {
                fill_voice_combo(hwnd, "");
                return LRESULT(0);
            }
            if id == ID_ONLY_MULTILINGUAL {
                refresh_edge_controls(hwnd, "");
                return LRESULT(0);
            }
            if id == ID_OK {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DialogueVoiceDialogData;
                if !ptr.is_null() {
                    let data = &mut *ptr;
                    let engine = selected_engine(GetDlgItem(hwnd, ID_ENGINE));
                    let voice = selected_voice(hwnd);
                    let rate = selected_combo_value(GetDlgItem(hwnd, ID_RATE), data.default_rate);
                    let pitch =
                        selected_combo_value(GetDlgItem(hwnd, ID_PITCH), data.default_pitch);
                    let volume =
                        selected_combo_value(GetDlgItem(hwnd, ID_VOLUME), data.default_volume);
                    let open_quote = read_control_text(GetDlgItem(hwnd, ID_OPEN_QUOTE));
                    let close_quote = read_control_text(GetDlgItem(hwnd, ID_CLOSE_QUOTE));
                    let allow_multiline = SendMessageW(
                        GetDlgItem(hwnd, ID_ALLOW_MULTILINE),
                        BM_GETCHECK,
                        WPARAM(0),
                        LPARAM(0),
                    )
                    .0 as u32
                        == BST_CHECKED.0;
                    if !voice.trim().is_empty() && !open_quote.is_empty() && !close_quote.is_empty()
                    {
                        data.result = Some(DialogueVoiceConfig {
                            engine,
                            voice,
                            use_secondary_voice: false,
                            secondary_voice: String::new(),
                            secondary_engine: TtsEngine::Edge,
                            secondary_rate: 0,
                            secondary_pitch: 0,
                            secondary_volume: 100,
                            rate,
                            pitch,
                            volume,
                            opening_quote: open_quote,
                            closing_quote: close_quote,
                            allow_multiline,
                        });
                    }
                }
                crate::log_if_err!(DestroyWindow(hwnd));
                return LRESULT(0);
            }
            if id == ID_CANCEL {
                crate::log_if_err!(DestroyWindow(hwnd));
                return LRESULT(0);
            }
            LRESULT(0)
        }
        WM_KEYDOWN => {
            let key = wparam.0 as u32;
            if key == windows::Win32::UI::Input::KeyboardAndMouse::VK_RETURN.0 as u32 {
                SendMessageW(hwnd, WM_COMMAND, WPARAM(ID_OK as usize), LPARAM(0));
                return LRESULT(0);
            }
            if key == windows::Win32::UI::Input::KeyboardAndMouse::VK_ESCAPE.0 as u32 {
                SendMessageW(hwnd, WM_COMMAND, WPARAM(ID_CANCEL as usize), LPARAM(0));
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            crate::log_if_err!(DestroyWindow(hwnd));
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

pub fn open_dialog(
    parent: HWND,
    language: Language,
    edge_voices: Vec<VoiceInfo>,
    sapi5_voices: Vec<VoiceInfo>,
    sapi4_voices: Vec<VoiceInfo>,
    default_cfg: DialogueVoiceConfig,
) -> Option<DialogueVoiceConfig> {
    unsafe {
        let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
        let class_name = to_wide("SonarpadDialogueVoiceDialog");
        let title = i18n::tr(language, "dialogue_voice.apply.title");

        static ONCE: std::sync::Once = std::sync::Once::new();
        ONCE.call_once(|| {
            let wc = WNDCLASSW {
                hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
                    windows::Win32::UI::WindowsAndMessaging::LoadCursorW(None, IDC_ARROW)
                        .unwrap_or_default()
                        .0,
                ),
                hInstance: hinstance,
                lpszClassName: PCWSTR(class_name.as_ptr()),
                lpfnWndProc: Some(wndproc),
                ..Default::default()
            };
            RegisterClassW(&wc);
        });

        let mut data = DialogueVoiceDialogData {
            language,
            edge_voices,
            sapi5_voices,
            sapi4_voices,
            default_engine: default_cfg.engine,
            default_voice: default_cfg.voice,
            default_rate: default_cfg.rate,
            default_pitch: default_cfg.pitch,
            default_volume: default_cfg.volume,
            default_open_quote: default_cfg.opening_quote,
            default_close_quote: default_cfg.closing_quote,
            default_allow_multiline: default_cfg.allow_multiline,
            edge_language_codes: Vec::new(),
            result: None,
        };

        let hwnd = CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(to_wide(&title).as_ptr()),
            WS_CAPTION | WS_SYSMENU,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            580,
            460,
            parent,
            HMENU(0),
            hinstance,
            Some(&mut data as *mut _ as *const std::ffi::c_void),
        );
        if hwnd.0 == 0 {
            return None;
        }

        windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow(parent, false);
        ShowWindow(hwnd, SW_SHOW);

        let mut msg = MSG::default();
        while IsWindow(hwnd).as_bool() && GetMessageW(&mut msg, HWND(0), 0, 0).into() {
            if !IsDialogMessageW(hwnd, &msg).as_bool() {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
        windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow(parent, true);
        windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow(parent);

        data.result
    }
}
