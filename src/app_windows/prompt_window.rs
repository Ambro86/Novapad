use crate::accessibility::{EM_GETSEL, EM_REPLACESEL, EM_SCROLLCARET, to_wide};
use crate::conpty::{ConPtySession, ConPtySpawn};
use crate::settings::{Language, confirm_title, save_settings};
use crate::{i18n, log_debug, show_error, with_state};
use std::collections::VecDeque;
use std::process::Command;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::{Duration, SystemTime, UNIX_EPOCH};
use windows::Win32::Foundation::{HANDLE, HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetDC, GetTextMetricsW, ReleaseDC, TEXTMETRICW};
use windows::Win32::Storage::FileSystem::ReadFile;
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
};
use windows::Win32::System::Diagnostics::Debug::MessageBeep;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Memory::{GMEM_MOVEABLE, GlobalAlloc, GlobalLock, GlobalUnlock};
use windows::Win32::System::Power::{
    ES_CONTINUOUS, ES_SYSTEM_REQUIRED, EXECUTION_STATE, SetThreadExecutionState,
};

use windows::Win32::UI::Controls::{WC_BUTTON, WC_EDIT, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetFocus, GetKeyState, SetFocus, VK_CONTROL, VK_ESCAPE, VK_RETURN, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, CreateWindowExW, DefWindowProcW, DestroyWindow,
    ES_AUTOHSCROLL, ES_AUTOVSCROLL, ES_MULTILINE, ES_READONLY, GetClientRect, GetParent,
    GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW, KillTimer,
    LoadCursorW, MB_ICONQUESTION, MB_OKCANCEL, MESSAGEBOX_STYLE, MSG, MessageBoxW, PostMessageW,
    RegisterClassW, SendMessageW, SetForegroundWindow, SetTimer, SetWindowLongPtrW, SetWindowTextW,
    WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY,
    WM_SETFOCUS, WM_SETFONT, WM_SIZE, WM_SYSKEYDOWN, WM_TIMER, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SIZEBOX, WS_SYSMENU, WS_TABSTOP,
    WS_VISIBLE, WS_VSCROLL,
};
use windows::core::PCWSTR;

const PROMPT_CLASS_NAME: &str = "NovapadPrompt";

const PROMPT_ID_INPUT: usize = 9301;
const PROMPT_ID_OUTPUT: usize = 9302;
const PROMPT_ID_AUTOSCROLL: usize = 9303;
const PROMPT_ID_STRIP_ANSI: usize = 9304;
const PROMPT_ID_ANNOUNCE_LINES: usize = 9305;
const PROMPT_ID_BEEP_ON_IDLE: usize = 9306;
const PROMPT_ID_PREVENT_SLEEP: usize = 9307;

const WM_PROMPT_OUTPUT: u32 = WM_APP + 60;
const EM_SETSEL: u32 = 0x00B1;
const EM_LIMITTEXT: u32 = 0x00C5;
const EM_SETREADONLY: u32 = 0x00CF;
const PROMPT_OUTPUT_LIMIT: usize = 40_000;
const PROMPT_OUTPUT_KEEP: usize = 10_000;
const PROMPT_OUTPUT_TIMER_ID: usize = 3;
const PROMPT_OUTPUT_FLUSH_CHARS: usize = 2048;

struct PromptLabels {
    title: String,
    input: String,
    output: String,
    autoscroll: String,
    strip_ansi: String,
    announce_lines: String,
    beep_on_idle: String,
    prevent_sleep: String,
    clear_confirm: String,
}

struct AnsiStripper {
    state: AnsiState,
    chars_consumed: usize,
}

enum AnsiState {
    Normal,
    Esc,
    Csi,
    Osc,
    OscEsc, // Met ESC while in OSC, waiting for backslash
}

impl AnsiStripper {
    fn new() -> Self {
        Self {
            state: AnsiState::Normal,
            chars_consumed: 0,
        }
    }

    fn process(&mut self, input: &str) -> String {
        let mut out = String::with_capacity(input.len());
        for ch in input.chars() {
            // Safety valve: prevent getting stuck in a state for too long
            if !matches!(self.state, AnsiState::Normal) {
                self.chars_consumed += 1;
                if self.chars_consumed > 1000 {
                    self.state = AnsiState::Normal;
                    self.chars_consumed = 0;
                }
            }

            match self.state {
                AnsiState::Normal => {
                    if ch == '\x1B' {
                        self.state = AnsiState::Esc;
                        self.chars_consumed = 0;
                    } else {
                        out.push(ch);
                    }
                }
                AnsiState::Esc => {
                    match ch {
                        '[' => self.state = AnsiState::Csi,
                        ']' => self.state = AnsiState::Osc,
                        '(' | ')' | '>' | '=' | '7' | '8' | 'M' | 'E' | 'D' | 'H' | 'Z' => {
                            // Simple sequences, just consume and return to normal
                            self.state = AnsiState::Normal;
                        }
                        _ => {
                            // Unknown sequence or simple ESC
                            self.state = AnsiState::Normal;
                            out.push(ch);
                        }
                    }
                }
                AnsiState::Csi => {
                    // CSI sequences MUST be ASCII.
                    // Parameter bytes: 0x30–0x3F (0-9:;<=>?)
                    // Intermediate bytes: 0x20–0x2F (space !"#$%&'()*+,-./)
                    // Final bytes: 0x40–0x7E (@A-Z[\]^_`a-z{|}~)
                    if ch.is_ascii() {
                        let b = ch as u8;
                        if (0x40..=0x7E).contains(&b) {
                            // Valid terminator
                            self.state = AnsiState::Normal;
                        } else if (0x20..=0x3F).contains(&b) {
                            // Valid parameter/intermediate, consume
                        } else {
                            // Invalid ASCII char for CSI (e.g. control char < 0x20)
                            // Abort sequence and output character
                            self.state = AnsiState::Normal;
                            out.push(ch);
                        }
                    } else {
                        // Non-ASCII character (e.g. UTF-8 box drawing).
                        // Definitely not part of a standard CSI sequence.
                        // Abort sequence and output character.
                        self.state = AnsiState::Normal;
                        out.push(ch);
                    }
                }
                AnsiState::Osc => {
                    if ch == '\x07' {
                        // BEL terminates OSC
                        self.state = AnsiState::Normal;
                    } else if ch == '\x1B' {
                        // Check for ST (ESC \)
                        self.state = AnsiState::OscEsc;
                    }
                    // Otherwise consume content of OSC
                }
                AnsiState::OscEsc => {
                    if ch == '\\' {
                        self.state = AnsiState::Normal;
                    } else {
                        // Not a backslash, so it wasn't an ST terminator.
                        // We are still in OSC mode.
                        self.state = AnsiState::Osc;
                    }
                }
            }
        }
        out
    }
}

struct PromptState {
    parent: HWND,
    label_input: HWND,
    input: HWND,
    label_output: HWND,
    output: HWND,

    checkbox_autoscroll: HWND,
    checkbox_strip_ansi: HWND,
    checkbox_announce_lines: HWND,
    checkbox_beep_on_idle: HWND,
    checkbox_prevent_sleep: HWND,
    auto_scroll: bool,
    strip_ansi: bool,
    announce_lines: bool,
    beep_on_idle: bool,
    prevent_sleep: bool,
    buffer: String,
    buffer_utf16_len: usize,
    line_start_byte: usize,
    line_start_utf16: usize,
    line_has_content: bool,
    blank_line_streak: u8,
    pending_ws: String,
    program_is_codex: bool,
    last_announced_line: String,
    beep_state: Arc<PromptBeepState>,
    session: Option<ConPtySession>,
    reader_cancel: Arc<AtomicBool>,
    ansi_stripper: AnsiStripper,
    output_queue: VecDeque<String>,
    output_flush_active: bool,
}

fn prompt_labels(language: Language) -> PromptLabels {
    PromptLabels {
        title: i18n::tr(language, "prompt.title"),
        input: i18n::tr(language, "prompt.input"),
        output: i18n::tr(language, "prompt.output"),
        autoscroll: i18n::tr(language, "prompt.autoscroll"),
        strip_ansi: i18n::tr(language, "prompt.strip_ansi"),
        announce_lines: i18n::tr(language, "prompt.announce_lines"),
        beep_on_idle: i18n::tr(language, "prompt.beep_on_idle"),
        prevent_sleep: i18n::tr(language, "prompt.prevent_sleep"),
        clear_confirm: i18n::tr(language, "prompt.clear_confirm"),
    }
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.prompt_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(PROMPT_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(prompt_wndproc),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = prompt_labels(language);
    let title = to_wide(&labels.title);

    let hwnd = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_SIZEBOX | WS_VISIBLE,
        140,
        140,
        720,
        520,
        None,
        HMENU(0),
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if hwnd.0 == 0 {
        return;
    }

    let _ = with_state(parent, |state| {
        state.prompt_window = hwnd;
    });
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    let focus = GetFocus();
    if focus.0 == 0 {
        return false;
    }
    let focus_parent = GetParent(focus);
    if focus != hwnd && focus_parent != hwnd {
        return false;
    }

    if msg.message == WM_SYSKEYDOWN {
        if msg.wParam.0 as u32 == 'I' as u32 {
            let _ = with_prompt_state(hwnd, |state| {
                SetFocus(state.input);
            });
            return true;
        }
        if msg.wParam.0 as u32 == 'O' as u32 {
            let _ = with_prompt_state(hwnd, |state| {
                SetFocus(state.output);
            });
            return true;
        }
        return false;
    }

    if msg.message != WM_KEYDOWN {
        return false;
    }

    if msg.wParam.0 as u32 == VK_TAB.0 as u32 {
        let shift_down = (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
        let _ = with_prompt_state(hwnd, |state| {
            let order = [
                state.input,
                state.output,
                state.checkbox_autoscroll,
                state.checkbox_strip_ansi,
                state.checkbox_announce_lines,
                state.checkbox_beep_on_idle,
                state.checkbox_prevent_sleep,
            ];
            let mut idx = order.iter().position(|&h| h == focus).unwrap_or(0);
            if shift_down {
                idx = if idx == 0 { order.len() - 1 } else { idx - 1 };
            } else {
                idx = (idx + 1) % order.len();
            }
            SetFocus(order[idx]);
        });
        return true;
    }

    if msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let _ = with_prompt_state(hwnd, |state| {
            if focus == state.input {
                send_input_to_pty(state);
            }
        });
        return true;
    }

    let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
    if ctrl_down && msg.wParam.0 as u32 == 'C' as u32 {
        let _ = with_prompt_state(hwnd, |state| {
            if focus == state.output {
                copy_output_selection(state.output);
            } else if let Some(session) = state.session.as_ref() {
                let _ = session.send_ctrl_c();
            }
        });
        return true;
    }
    if ctrl_down && msg.wParam.0 as u32 == 'L' as u32 {
        let _ = with_prompt_state(hwnd, |state| {
            if confirm_clear_output(hwnd, state.parent) {
                clear_output(state);
            }
        });
        return true;
    }

    false
}

unsafe extern "system" fn prompt_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct =
                lparam.0 as *const windows::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = prompt_labels(language);
            let hfont = with_state(parent, |state| state.hfont).unwrap_or_default();
            let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();

            let label_input = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.input).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                16,
                80,
                18,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let input = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                100,
                14,
                580,
                22,
                hwnd,
                HMENU(PROMPT_ID_INPUT as isize),
                HINSTANCE(0),
                None,
            );

            let label_output = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.output).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                50,
                80,
                18,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let output = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WINDOW_STYLE((ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY) as u32),
                16,
                70,
                664,
                360,
                hwnd,
                HMENU(PROMPT_ID_OUTPUT as isize),
                HINSTANCE(0),
                None,
            );
            let _ = SendMessageW(output, EM_LIMITTEXT, WPARAM(0x7FFFFFFE), LPARAM(0));

            let checkbox_autoscroll = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.autoscroll).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                16,
                440,
                200,
                20,
                hwnd,
                HMENU(PROMPT_ID_AUTOSCROLL as isize),
                HINSTANCE(0),
                None,
            );
            let checkbox_strip_ansi = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.strip_ansi).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                230,
                440,
                220,
                20,
                hwnd,
                HMENU(PROMPT_ID_STRIP_ANSI as isize),
                HINSTANCE(0),
                None,
            );
            let checkbox_announce_lines = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.announce_lines).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                16,
                464,
                260,
                20,
                hwnd,
                HMENU(PROMPT_ID_ANNOUNCE_LINES as isize),
                HINSTANCE(0),
                None,
            );
            let checkbox_beep_on_idle = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.beep_on_idle).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                290,
                464,
                240,
                20,
                hwnd,
                HMENU(PROMPT_ID_BEEP_ON_IDLE as isize),
                HINSTANCE(0),
                None,
            );
            let checkbox_prevent_sleep = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.prevent_sleep).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
                16,
                488,
                320,
                20,
                hwnd,
                HMENU(PROMPT_ID_PREVENT_SLEEP as isize),
                HINSTANCE(0),
                None,
            );

            for control in [
                label_input,
                input,
                label_output,
                output,
                checkbox_autoscroll,
                checkbox_strip_ansi,
                checkbox_announce_lines,
                checkbox_beep_on_idle,
                checkbox_prevent_sleep,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let auto_scroll = settings.prompt_auto_scroll;
            let strip_ansi = settings.prompt_strip_ansi;
            let announce_lines = settings.prompt_announce_lines;
            let beep_on_idle = settings.prompt_beep_on_idle;
            let prevent_sleep = settings.prompt_prevent_sleep;
            let program_lower = settings.prompt_program.to_ascii_lowercase();
            let program_is_codex =
                program_lower.contains("codex") || program_lower.contains("claude");
            let _ = SendMessageW(
                checkbox_autoscroll,
                BM_SETCHECK,
                WPARAM(if auto_scroll { 1 } else { 0 }),
                LPARAM(0),
            );
            let _ = SendMessageW(
                checkbox_strip_ansi,
                BM_SETCHECK,
                WPARAM(if strip_ansi { 1 } else { 0 }),
                LPARAM(0),
            );
            let _ = SendMessageW(
                checkbox_announce_lines,
                BM_SETCHECK,
                WPARAM(if announce_lines { 1 } else { 0 }),
                LPARAM(0),
            );
            let _ = SendMessageW(
                checkbox_beep_on_idle,
                BM_SETCHECK,
                WPARAM(if beep_on_idle { 1 } else { 0 }),
                LPARAM(0),
            );
            let _ = SendMessageW(
                checkbox_prevent_sleep,
                BM_SETCHECK,
                WPARAM(if prevent_sleep { 1 } else { 0 }),
                LPARAM(0),
            );

            let reader_cancel = Arc::new(AtomicBool::new(false));
            let beep_state = Arc::new(PromptBeepState::new(beep_on_idle, prevent_sleep));
            let mut state = PromptState {
                parent,
                label_input,
                input,
                label_output,
                output,

                checkbox_autoscroll,
                checkbox_strip_ansi,
                checkbox_announce_lines,
                checkbox_beep_on_idle,
                checkbox_prevent_sleep,
                auto_scroll,
                strip_ansi,
                announce_lines,
                beep_on_idle,
                prevent_sleep,
                buffer: String::new(),
                buffer_utf16_len: 0,
                line_start_byte: 0,
                line_start_utf16: 0,
                line_has_content: false,
                blank_line_streak: 0,
                pending_ws: String::new(),
                program_is_codex,
                last_announced_line: String::new(),
                beep_state: beep_state.clone(),
                session: None,
                reader_cancel: reader_cancel.clone(),
                ansi_stripper: AnsiStripper::new(),
                output_queue: VecDeque::new(),
                output_flush_active: false,
            };

            layout_prompt(hwnd, &state);

            if let Some(spawn) = start_prompt_session(hwnd, &settings.prompt_program, &state) {
                state.session = Some(spawn.session);
                start_output_reader(hwnd, spawn.output_read, reader_cancel, beep_state);
            }

            SetWindowLongPtrW(
                hwnd,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                Box::into_raw(Box::new(state)) as isize,
            );
            SetFocus(input);
            LRESULT(0)
        }
        WM_SIZE => {
            let _ = with_prompt_state(hwnd, |state| {
                layout_prompt(hwnd, state);
                if let Some(session) = state.session.as_ref() {
                    if let Some((cols, rows)) = output_cells(state.output) {
                        let _ = session.resize(cols, rows);
                    }
                }
            });
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            match cmd_id {
                PROMPT_ID_AUTOSCROLL => {
                    let _ = with_prompt_state(hwnd, |state| {
                        let checked = SendMessageW(
                            state.checkbox_autoscroll,
                            BM_GETCHECK,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0 != 0;
                        state.auto_scroll = checked;
                        update_prompt_settings(state.parent, |settings| {
                            settings.prompt_auto_scroll = checked;
                        });
                    });
                    LRESULT(0)
                }
                PROMPT_ID_STRIP_ANSI => {
                    let _ = with_prompt_state(hwnd, |state| {
                        let checked = SendMessageW(
                            state.checkbox_strip_ansi,
                            BM_GETCHECK,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0 != 0;
                        state.strip_ansi = checked;
                        update_prompt_settings(state.parent, |settings| {
                            settings.prompt_strip_ansi = checked;
                        });
                    });
                    LRESULT(0)
                }
                PROMPT_ID_ANNOUNCE_LINES => {
                    let _ = with_prompt_state(hwnd, |state| {
                        let checked = SendMessageW(
                            state.checkbox_announce_lines,
                            BM_GETCHECK,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0 != 0;
                        state.announce_lines = checked;
                        update_prompt_settings(state.parent, |settings| {
                            settings.prompt_announce_lines = checked;
                        });
                    });
                    LRESULT(0)
                }
                PROMPT_ID_BEEP_ON_IDLE => {
                    let _ = with_prompt_state(hwnd, |state| {
                        let checked = SendMessageW(
                            state.checkbox_beep_on_idle,
                            BM_GETCHECK,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0 != 0;
                        state.beep_on_idle = checked;
                        state.beep_state.enabled.store(checked, Ordering::Relaxed);
                        update_prompt_settings(state.parent, |settings| {
                            settings.prompt_beep_on_idle = checked;
                        });
                    });
                    LRESULT(0)
                }
                PROMPT_ID_PREVENT_SLEEP => {
                    let _ = with_prompt_state(hwnd, |state| {
                        let checked = SendMessageW(
                            state.checkbox_prevent_sleep,
                            BM_GETCHECK,
                            WPARAM(0),
                            LPARAM(0),
                        )
                        .0 != 0;
                        state.prevent_sleep = checked;
                        state
                            .beep_state
                            .sleep_enabled
                            .store(checked, Ordering::Relaxed);
                        if !checked && state.beep_state.sleep_active.load(Ordering::Relaxed) {
                            let _ = apply_prevent_sleep(false);
                            state
                                .beep_state
                                .sleep_active
                                .store(false, Ordering::Relaxed);
                        }
                        update_prompt_settings(state.parent, |settings| {
                            settings.prompt_prevent_sleep = checked;
                        });
                    });
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let _ = with_prompt_state(hwnd, |state| {
                    if focus == state.input {
                        send_input_to_pty(state);
                        return;
                    }
                });
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            let ctrl_down = (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
            if ctrl_down && wparam.0 as u32 == 'C' as u32 {
                let _ = with_prompt_state(hwnd, |state| {
                    let focus = GetFocus();
                    if focus == state.output {
                        copy_output_selection(state.output);
                    } else if let Some(session) = state.session.as_ref() {
                        let _ = session.send_ctrl_c();
                    }
                });
                return LRESULT(0);
            }
            if ctrl_down && wparam.0 as u32 == 'L' as u32 {
                let _ = with_prompt_state(hwnd, |state| {
                    if confirm_clear_output(hwnd, state.parent) {
                        clear_output(state);
                    }
                });
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_SYSKEYDOWN => {
            if wparam.0 as u32 == 'I' as u32 {
                let _ = with_prompt_state(hwnd, |state| {
                    SetFocus(state.input);
                });
                return LRESULT(0);
            }
            if wparam.0 as u32 == 'O' as u32 {
                let _ = with_prompt_state(hwnd, |state| {
                    SetFocus(state.output);
                });
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_SETFOCUS => {
            let _ = with_prompt_state(hwnd, |state| {
                if state.input.0 != 0 {
                    SetFocus(state.input);
                }
            });
            LRESULT(0)
        }
        WM_PROMPT_OUTPUT => {
            if lparam.0 == 0 {
                return LRESULT(0);
            }
            let payload = unsafe { Box::from_raw(lparam.0 as *mut String) };
            let _ = with_prompt_state(hwnd, |state| {
                state.output_queue.push_back(*payload);
                if !state.output_flush_active {
                    state.output_flush_active = true;
                    let _ = SetTimer(hwnd, PROMPT_OUTPUT_TIMER_ID, 20, None);
                }
            });
            LRESULT(0)
        }
        WM_TIMER => {
            if wparam.0 as usize == PROMPT_OUTPUT_TIMER_ID {
                let _ = with_prompt_state(hwnd, |state| {
                    let mut budget = PROMPT_OUTPUT_FLUSH_CHARS;
                    while budget > 0 {
                        let Some(chunk) = state.output_queue.pop_front() else {
                            break;
                        };
                        budget = budget.saturating_sub(chunk.len());
                        append_output(state, &chunk);
                    }
                    if state.output_queue.is_empty() {
                        state.output_flush_active = false;
                        let _ = KillTimer(hwnd, PROMPT_OUTPUT_TIMER_ID);
                    }
                });
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_DESTROY => {
            let _ = with_prompt_state(hwnd, |state| {
                state.reader_cancel.store(true, Ordering::Relaxed);
                state.output_queue.clear();
                state.output_flush_active = false;
                let _ = KillTimer(hwnd, PROMPT_OUTPUT_TIMER_ID);
                if state.beep_state.sleep_active.load(Ordering::Relaxed) {
                    let _ = apply_prevent_sleep(false);
                    state
                        .beep_state
                        .sleep_active
                        .store(false, Ordering::Relaxed);
                }
                if let Some(mut session) = state.session.take() {
                    session.close();
                }
            });
            let parent = with_prompt_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            let _ = with_state(parent, |state| {
                state.prompt_window = HWND(0);
            });
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr =
                GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                    as *mut PromptState;
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

unsafe fn with_prompt_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut PromptState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
        as *mut PromptState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

fn copy_output_selection(hwnd_output: HWND) {
    unsafe {
        const CF_UNICODETEXT: u32 = 13;
        let mut start: u32 = 0;
        let mut end: u32 = 0;
        let _ = SendMessageW(
            hwnd_output,
            EM_GETSEL,
            WPARAM(&mut start as *mut u32 as usize),
            LPARAM(&mut end as *mut u32 as isize),
        );
        if end <= start {
            return;
        }
        let len = GetWindowTextLengthW(hwnd_output);
        if len <= 0 {
            return;
        }
        let mut buf = vec![0u16; (len + 1) as usize];
        let read = GetWindowTextW(hwnd_output, &mut buf) as usize;
        if read == 0 {
            return;
        }
        let start = (start as usize).min(read);
        let end = (end as usize).min(read);
        if end <= start {
            return;
        }
        let mut selection = buf[start..end].to_vec();
        selection.push(0);
        if OpenClipboard(hwnd_output).is_err() {
            return;
        }
        let _ = EmptyClipboard();
        let size = selection.len() * std::mem::size_of::<u16>();
        let handle = match GlobalAlloc(GMEM_MOVEABLE, size) {
            Ok(handle) => handle,
            Err(_) => {
                let _ = CloseClipboard();
                return;
            }
        };
        if handle.0.is_null() {
            let _ = CloseClipboard();
            return;
        }
        let ptr = GlobalLock(handle) as *mut u16;
        if ptr.is_null() {
            let _ = CloseClipboard();
            return;
        }
        std::ptr::copy_nonoverlapping(selection.as_ptr(), ptr, selection.len());
        let _ = GlobalUnlock(handle);
        let _ = SetClipboardData(CF_UNICODETEXT, HANDLE(handle.0 as isize));
        let _ = CloseClipboard();
    }
}

fn start_prompt_session(hwnd: HWND, program: &str, state: &PromptState) -> Option<ConPtySpawn> {
    let (cols, rows) = output_cells(state.output).unwrap_or((80, 24));
    match ConPtySession::spawn(program, cols, rows) {
        Ok(spawn) => Some(spawn),
        Err(err) => {
            log_debug(&format!("Prompt spawn failed: {err}"));
            unsafe {
                let language =
                    with_state(state.parent, |state| state.settings.language).unwrap_or_default();
                show_error(hwnd, language, &format!("Prompt error: {err}"));
            }
            None
        }
    }
}

fn start_output_reader(
    hwnd: HWND,
    output_read: windows::Win32::Foundation::HANDLE,
    cancel: Arc<AtomicBool>,
    beep_state: Arc<PromptBeepState>,
) {
    let beep_cancel = cancel.clone();
    let beep_state_clone = beep_state.clone();
    std::thread::spawn(move || {
        loop {
            if beep_cancel.load(Ordering::Relaxed) {
                break;
            }
            let last = beep_state_clone.last_output_ms.load(Ordering::Relaxed);
            if last != 0 {
                let now = now_ms();
                if now.saturating_sub(last) >= 1_000
                    && beep_state_clone.enabled.load(Ordering::Relaxed)
                    && !beep_state_clone.beeped.swap(true, Ordering::Relaxed)
                {
                    unsafe {
                        let _ = MessageBeep(MESSAGEBOX_STYLE(0));
                    }
                }
                if now.saturating_sub(last) >= 1_000
                    && beep_state_clone.sleep_enabled.load(Ordering::Relaxed)
                    && beep_state_clone.sleep_active.load(Ordering::Relaxed)
                {
                    let _ = apply_prevent_sleep(false);
                    beep_state_clone
                        .sleep_active
                        .store(false, Ordering::Relaxed);
                }
            }
            std::thread::sleep(Duration::from_millis(200));
        }
    });
    std::thread::spawn(move || {
        let mut buffer = [0u8; 4096];
        let mut total_read = 0usize;
        loop {
            if cancel.load(Ordering::Relaxed) {
                break;
            }
            let mut read = 0u32;
            log_debug("Prompt: Calling ReadFile...");
            let ok =
                unsafe { ReadFile(output_read, Some(&mut buffer), Some(&mut read), None).is_ok() };
            if !ok {
                let err = unsafe { windows::Win32::Foundation::GetLastError() };
                log_debug(&format!(
                    "Prompt: ReadFile returned false, error code: {:?}, total bytes read so far: {}",
                    err, total_read
                ));
                break;
            }
            if read == 0 {
                log_debug(&format!(
                    "Prompt: ReadFile read 0 bytes (EOF), total bytes read: {}",
                    total_read
                ));
                break;
            }
            total_read += read as usize;
            log_debug(&format!(
                "Prompt: ReadFile read {} bytes, total: {}",
                read, total_read
            ));
            beep_state.last_output_ms.store(now_ms(), Ordering::Relaxed);
            beep_state.beeped.store(false, Ordering::Relaxed);
            if beep_state.sleep_enabled.load(Ordering::Relaxed)
                && !beep_state.sleep_active.swap(true, Ordering::Relaxed)
            {
                let _ = apply_prevent_sleep(true);
            }
            let chunk = String::from_utf8_lossy(&buffer[..read as usize]).to_string();
            let payload = Box::new(chunk);
            unsafe {
                let payload_ptr = Box::into_raw(payload);
                if PostMessageW(
                    hwnd,
                    WM_PROMPT_OUTPUT,
                    WPARAM(0),
                    LPARAM(payload_ptr as isize),
                )
                .is_err()
                {
                    let _ = Box::from_raw(payload_ptr);
                    break;
                }
            }
        }
        unsafe {
            let _ = windows::Win32::Foundation::CloseHandle(output_read);
        }
    });
}

fn update_prompt_settings<F>(parent: HWND, update: F)
where
    F: FnOnce(&mut crate::settings::AppSettings),
{
    let settings = unsafe {
        with_state(parent, |state| {
            update(&mut state.settings);
            state.settings.clone()
        })
    }
    .unwrap_or_default();
    save_settings(settings);
}

unsafe fn send_input_to_pty(state: &mut PromptState) {
    if state.input.0 == 0 {
        return;
    }
    let len = GetWindowTextLengthW(state.input);
    if len < 0 {
        return;
    }
    let mut buffer = vec![0u16; (len + 1) as usize];
    let read = GetWindowTextW(state.input, &mut buffer);
    let text = String::from_utf16_lossy(&buffer[..read as usize]);
    if state.program_is_codex && is_codex_approvals_command(&text) {
        spawn_codex_approvals();
        let _ = SetWindowTextW(state.input, PCWSTR::null());
        return;
    }
    if let Some(session) = state.session.as_ref() {
        let newline = if state.program_is_codex { "\n" } else { "\r\n" };
        let payload = format!("{text}{newline}");
        let _ = session.write_input(&payload);
    }
    let _ = SetWindowTextW(state.input, PCWSTR::null());
}

unsafe fn clear_output(state: &mut PromptState) {
    state.buffer.clear();
    state.buffer_utf16_len = 0;
    state.line_start_byte = 0;
    state.line_start_utf16 = 0;
    state.line_has_content = false;
    state.blank_line_streak = 0;
    state.pending_ws.clear();
    state.last_announced_line.clear();
    let _ = SetWindowTextW(state.output, PCWSTR::null());
}

unsafe fn trim_output_keep_last(state: &mut PromptState) {
    if state.buffer_utf16_len <= PROMPT_OUTPUT_KEEP {
        return;
    }
    let excess = state.buffer_utf16_len - PROMPT_OUTPUT_KEEP;
    let mut units_removed = 0usize;
    let mut cut_idx = 0usize;
    for (byte_idx, ch) in state.buffer.char_indices() {
        units_removed += ch.len_utf16();
        cut_idx = byte_idx + ch.len_utf8();
        if units_removed >= excess {
            break;
        }
    }
    if cut_idx == 0 {
        return;
    }
    state.buffer.drain(..cut_idx);
    state.buffer_utf16_len -= units_removed;
    state.line_start_byte = state.buffer.len();
    state.line_start_utf16 = state.buffer_utf16_len;
    state.line_has_content = false;
    state.blank_line_streak = 0;
    state.pending_ws.clear();
    state.last_announced_line.clear();
    let wide = to_wide(&state.buffer);
    let _ = SendMessageW(state.output, EM_SETREADONLY, WPARAM(0), LPARAM(0));
    let _ = SetWindowTextW(state.output, PCWSTR(wide.as_ptr()));
    let _ = SendMessageW(state.output, EM_SETREADONLY, WPARAM(1), LPARAM(0));
    let _ = SendMessageW(
        state.output,
        EM_SETSEL,
        WPARAM(state.buffer_utf16_len),
        LPARAM(state.buffer_utf16_len as isize),
    );
    let _ = SendMessageW(state.output, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
}

fn apply_prevent_sleep(enabled: bool) -> bool {
    let flags = if enabled {
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    } else {
        ES_CONTINUOUS
    };
    unsafe { SetThreadExecutionState(flags) != EXECUTION_STATE(0) }
}

fn confirm_clear_output(hwnd: HWND, parent: HWND) -> bool {
    let language =
        unsafe { with_state(parent, |state| state.settings.language).unwrap_or_default() };
    let labels = prompt_labels(language);
    let title = to_wide(&confirm_title(language));
    let message = to_wide(&labels.clear_confirm);
    unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OKCANCEL | MB_ICONQUESTION,
        )
        .0 == 1
    }
}

fn append_output(state: &mut PromptState, text: &str) {
    let filtered = if state.strip_ansi {
        let stripped = state.ansi_stripper.process(text);
        filter_context_left_lines(&stripped)
    } else {
        text.to_string()
    };

    let filtered_units = filtered.encode_utf16().count();
    if state.buffer_utf16_len + filtered_units > PROMPT_OUTPUT_LIMIT {
        unsafe {
            trim_output_keep_last(state);
        }
    }

    let prev_len = state.buffer_utf16_len;
    let prev_line_start_utf16 = state.line_start_utf16;
    let prev_line_start_byte = state.line_start_byte;
    let mut had_cr = false;
    let mut delta = String::new();
    let mut newline_appended = false;
    let mut lines_to_announce: Vec<String> = Vec::new();
    let mut chars = filtered.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\r' {
            if matches!(chars.peek(), Some(&'\n')) {
                let _ = chars.next();
                if !state.line_has_content {
                    if state.blank_line_streak >= 1 {
                        continue;
                    }
                    state.blank_line_streak = 1;
                } else {
                    state.blank_line_streak = 0;
                }
                append_newline(
                    state,
                    &mut delta,
                    &mut newline_appended,
                    &mut lines_to_announce,
                );
                state.line_has_content = false;
                state.pending_ws.clear();
            } else {
                had_cr = true;
                state.buffer.truncate(state.line_start_byte);
                state.buffer_utf16_len = state.line_start_utf16;
                delta.clear();
                state.line_has_content = false;
                state.blank_line_streak = 0;
                state.pending_ws.clear();
            }
            continue;
        }
        if ch == '\n' {
            if !state.line_has_content {
                if state.blank_line_streak >= 1 {
                    continue;
                }
                state.blank_line_streak = 1;
            } else {
                state.blank_line_streak = 0;
            }
            append_newline(
                state,
                &mut delta,
                &mut newline_appended,
                &mut lines_to_announce,
            );
            state.line_has_content = false;
            state.pending_ws.clear();
            continue;
        }
        if matches!(ch, ' ' | '\t') && !state.line_has_content {
            state.pending_ws.push(ch);
            continue;
        }
        if !state.pending_ws.is_empty() {
            state.buffer.push_str(&state.pending_ws);
            state.buffer_utf16_len += state.pending_ws.encode_utf16().count();
            delta.push_str(&state.pending_ws);
            state.pending_ws.clear();
        }
        state.buffer.push(ch);
        state.buffer_utf16_len += ch.len_utf16();
        delta.push(ch);
        if !ch.is_whitespace() {
            state.line_has_content = true;
        }
        state.blank_line_streak = 0;
    }

    let output = state.output;
    let focus = unsafe { GetFocus() };
    let mut sel_start = 0u32;
    let mut sel_end = 0u32;
    if focus == output {
        unsafe {
            let _ = SendMessageW(
                output,
                EM_GETSEL,
                WPARAM(&mut sel_start as *mut _ as usize),
                LPARAM(&mut sel_end as *mut _ as isize),
            );
        }
    }
    let should_scroll = state.auto_scroll && (focus != output || sel_end as usize == prev_len);

    let replace_start = if had_cr {
        prev_line_start_utf16
    } else {
        prev_len
    };
    let replace_end = prev_len;
    let replace_text = if had_cr {
        state.buffer[prev_line_start_byte..].to_string()
    } else {
        delta.clone() // Clone for debug log
    };

    let wide = to_wide(&replace_text);
    log_debug(&format!(
        "Prompt: appending output '{}'",
        replace_text.trim()
    ));
    unsafe {
        let _ = SendMessageW(
            output,
            EM_SETREADONLY,
            WPARAM(0), // False
            LPARAM(0),
        );
        let _ = SendMessageW(
            output,
            EM_SETSEL,
            WPARAM(replace_start),
            LPARAM(replace_end as isize),
        );
        let _ = SendMessageW(
            output,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(wide.as_ptr() as isize),
        );
        let _ = SendMessageW(
            output,
            EM_SETREADONLY,
            WPARAM(1), // True
            LPARAM(0),
        );
    }
    if state.announce_lines && newline_appended {
        for line in lines_to_announce {
            announce_line(&line);
            state.last_announced_line = line;
        }
    }
    if state.announce_lines {
        let current_line = state.buffer[state.line_start_byte..].to_string();
        if !current_line.is_empty()
            && current_line != state.last_announced_line
            && looks_like_prompt(&current_line)
        {
            announce_line(&current_line);
            state.last_announced_line = current_line;
        }
    }

    if should_scroll {
        unsafe {
            let _ = SendMessageW(
                output,
                EM_SETSEL,
                WPARAM(state.buffer_utf16_len),
                LPARAM(state.buffer_utf16_len as isize),
            );
            let _ = SendMessageW(output, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        }
    } else if focus == output {
        let max = state.buffer_utf16_len as u32;
        let restore_start = sel_start.min(max);
        let restore_end = sel_end.min(max);
        unsafe {
            let _ = SendMessageW(
                output,
                EM_SETSEL,
                WPARAM(restore_start as usize),
                LPARAM(restore_end as isize),
            );
        }
    }
}

fn append_newline(
    state: &mut PromptState,
    delta: &mut String,
    newline_appended: &mut bool,
    lines_to_announce: &mut Vec<String>,
) {
    let line = state.buffer[state.line_start_byte..].to_string();
    if !line.is_empty() {
        lines_to_announce.push(line);
    }
    state.buffer.push('\r');
    state.buffer.push('\n');
    state.buffer_utf16_len += 2;
    state.line_start_byte = state.buffer.len();
    state.line_start_utf16 = state.buffer_utf16_len;
    delta.push('\r');
    delta.push('\n');
    *newline_appended = true;
}

fn announce_line(line: &str) {
    if line.is_empty() {
        return;
    }
    crate::accessibility::nvda_speak(line);
}

fn looks_like_prompt(line: &str) -> bool {
    let trimmed = line.trim_end();
    if trimmed.is_empty() {
        return false;
    }
    let last = trimmed.chars().last().unwrap_or(' ');
    matches!(last, '>' | '$' | '#')
}

fn filter_context_left_lines(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    let bytes = input.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let line_start = i;
        while i < bytes.len() && bytes[i] != b'\n' && bytes[i] != b'\r' {
            i += 1;
        }
        let line = &input[line_start..i];
        let mut line_end = "";
        if i < bytes.len() {
            if bytes[i] == b'\r' {
                if i + 1 < bytes.len() && bytes[i + 1] == b'\n' {
                    line_end = "\r\n";
                    i += 2;
                } else {
                    line_end = "\r";
                    i += 1;
                }
            } else {
                line_end = "\n";
                i += 1;
            }
        }
        let line = if is_whitespace_only_line(line) {
            ""
        } else {
            line
        };
        if !is_context_left_line(line) && !is_interrupt_hint_line(line) {
            out.push_str(line);
            out.push_str(line_end);
        }
    }
    out
}

fn is_codex_approvals_command(text: &str) -> bool {
    text.trim().eq_ignore_ascii_case("/approvals")
}

fn spawn_codex_approvals() {
    let spawn = Command::new("cmd")
        .args(["/c", "start", "", "codex", "/approvals"])
        .spawn();
    if let Err(err) = spawn {
        log_debug(&format!("Prompt approvals spawn failed: {err}"));
    }
}

fn is_whitespace_only_line(line: &str) -> bool {
    !line.is_empty() && line.chars().all(|ch| ch == ' ' || ch == '\t')
}

fn is_context_left_line(line: &str) -> bool {
    let trimmed = line.trim();
    if trimmed.is_empty() {
        return false;
    }
    let lower = trimmed.to_ascii_lowercase();
    if lower.contains("context left") && lower.contains("shortcuts") {
        return true;
    }
    let Some(before_suffix) = trimmed.strip_suffix("context left") else {
        return false;
    };
    let before = before_suffix.trim_end();
    let Some(num_part) = before.strip_suffix('%') else {
        return false;
    };
    !num_part.is_empty() && num_part.chars().all(|c| c.is_ascii_digit())
}

fn is_interrupt_hint_line(line: &str) -> bool {
    line.to_ascii_lowercase().contains("esc to interrupt")
}

fn output_cells(hwnd_output: HWND) -> Option<(i16, i16)> {
    let mut rect = windows::Win32::Foundation::RECT::default();
    unsafe {
        if GetClientRect(hwnd_output, &mut rect).is_err() {
            return None;
        }
    }
    let width = (rect.right - rect.left).max(1);
    let height = (rect.bottom - rect.top).max(1);
    let (char_w, char_h) = text_metrics(hwnd_output).unwrap_or((8, 16));
    let cols = (width / char_w).max(1) as i16;
    let rows = (height / char_h).max(1) as i16;
    Some((cols, rows))
}

fn text_metrics(hwnd: HWND) -> Option<(i32, i32)> {
    unsafe {
        let hdc = GetDC(hwnd);
        if hdc.0 == 0 {
            return None;
        }
        let mut tm = TEXTMETRICW::default();
        let ok = GetTextMetricsW(hdc, &mut tm).as_bool();
        let _ = ReleaseDC(hwnd, hdc);
        if ok {
            Some((tm.tmAveCharWidth.max(1), tm.tmHeight.max(1)))
        } else {
            None
        }
    }
}

fn client_size(hwnd: HWND) -> Option<(i32, i32)> {
    let mut rect = windows::Win32::Foundation::RECT::default();
    unsafe {
        if GetClientRect(hwnd, &mut rect).is_err() {
            return None;
        }
    }
    Some((rect.right - rect.left, rect.bottom - rect.top))
}

fn layout_prompt(hwnd: HWND, state: &PromptState) {
    let Some((width, height)) = client_size(hwnd) else {
        return;
    };
    let margin = 16;
    let label_width = 80;
    let input_height = 22;
    let label_height = 18;
    let checkbox_height = 20;
    let spacing = 8;

    let mut y = margin;
    unsafe {
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.label_input,
            margin,
            y,
            label_width,
            label_height,
            true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.input,
            margin + label_width + spacing,
            y - 2,
            (width - margin * 2 - label_width - spacing).max(120),
            input_height,
            true,
        );
    }
    y += input_height + spacing;

    let output_height =
        (height - y - label_height - checkbox_height * 3 - spacing * 2 - margin).max(120);
    unsafe {
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.label_output,
            margin,
            y,
            label_width,
            label_height,
            true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.output,
            margin,
            y + label_height,
            (width - margin * 2).max(120),
            output_height,
            true,
        );
    }
    let output_bottom = y + label_height + output_height;
    let checkbox_y = output_bottom + spacing;
    let checkbox_y2 = checkbox_y + checkbox_height + spacing;
    let checkbox_y3 = checkbox_y2 + checkbox_height + spacing;
    unsafe {
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.checkbox_autoscroll,
            margin,
            checkbox_y,
            200,
            checkbox_height,
            true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.checkbox_strip_ansi,
            margin + 210,
            checkbox_y,
            220,
            checkbox_height,
            true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.checkbox_announce_lines,
            margin,
            checkbox_y2,
            260,
            checkbox_height,
            true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.checkbox_beep_on_idle,
            margin + 270,
            checkbox_y2,
            240,
            checkbox_height,
            true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.checkbox_prevent_sleep,
            margin,
            checkbox_y3,
            320,
            checkbox_height,
            true,
        );
    }
}
struct PromptBeepState {
    last_output_ms: AtomicU64,
    beeped: AtomicBool,
    enabled: AtomicBool,
    sleep_enabled: AtomicBool,
    sleep_active: AtomicBool,
}

impl PromptBeepState {
    fn new(beep_enabled: bool, sleep_enabled: bool) -> Self {
        Self {
            last_output_ms: AtomicU64::new(0),
            beeped: AtomicBool::new(false),
            enabled: AtomicBool::new(beep_enabled),
            sleep_enabled: AtomicBool::new(sleep_enabled),
            sleep_active: AtomicBool::new(false),
        }
    }
}

fn now_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis() as u64
}
