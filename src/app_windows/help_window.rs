#![allow(clippy::if_same_then_else, clippy::collapsible_else_if)]
use crate::accessibility::{normalize_to_crlf, to_wide};
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::WC_BUTTON;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetFocus, GetKeyState, SetFocus, VK_RETURN, VK_SHIFT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    ES_AUTOVSCROLL, ES_MULTILINE, ES_WANTRETURN, GWLP_USERDATA, GetWindowLongPtrW, HMENU,
    IDC_ARROW, IDCANCEL, LoadCursorW, MoveWindow, RegisterClassW, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_SETFOCUS, WM_SETFONT, WM_SIZE, WNDCLASSW,
    WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_OVERLAPPEDWINDOW, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const HELP_CLASS_NAME: &str = "NovapadHelp";
const HELP_ID_OK: usize = 7003;
const DONATIONS_IT: &str = include_str!("../../donations_it.txt");
const DONATIONS_EN: &str = include_str!("../../donations_en.txt");
const DONATIONS_ES: &str = include_str!("../../donations_es.txt");
const DONATIONS_PT: &str = include_str!("../../donations_pt.txt");

#[derive(Clone, Copy)]
enum HelpWindowKind {
    Guide,
    Changelog,
    Donations,
}

struct HelpWindowInit {
    parent: HWND,
    kind: HelpWindowKind,
    language: Language,
}

struct HelpWindowState {
    parent: HWND,
    edit: HWND,
    ok_button: HWND,
    kind: HelpWindowKind,
}

pub unsafe fn open(parent: HWND) {
    open_window(parent, HelpWindowKind::Guide);
}

pub unsafe fn open_changelog(parent: HWND) {
    open_window(parent, HelpWindowKind::Changelog);
}

pub unsafe fn open_donations(parent: HWND) {
    open_window(parent, HelpWindowKind::Donations);
}

unsafe fn open_window(parent: HWND, kind: HelpWindowKind) {
    let existing = with_state(parent, |state| match kind {
        HelpWindowKind::Guide => state.help_window,
        HelpWindowKind::Changelog => state.changelog_window,
        HelpWindowKind::Donations => state.donations_window,
    })
    .unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(HELP_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(help_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let title = to_wide(&help_title(language, kind));
    let init = Box::new(HelpWindowInit {
        parent,
        kind,
        language,
    });
    let init_ptr = Box::into_raw(init);
    let window = CreateWindowExW(
        WS_EX_CONTROLPARENT,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        640,
        520,
        parent,
        None,
        hinstance,
        Some(init_ptr as *const std::ffi::c_void),
    );

    if window.0 != 0 {
        let _ = with_state(parent, |state| match kind {
            HelpWindowKind::Guide => state.help_window = window,
            HelpWindowKind::Changelog => state.changelog_window = window,
            HelpWindowKind::Donations => state.donations_window = window,
        });
        SetForegroundWindow(window);
    } else if !init_ptr.is_null() {
        let _ = Box::from_raw(init_ptr);
    }
}

pub unsafe fn handle_tab(hwnd: HWND) {
    let _ = with_help_state(hwnd, |state| {
        let focus = GetFocus();
        let shift_down = (GetKeyState(VK_SHIFT.0 as i32) as u16) & 0x8000 != 0;

        if shift_down {
            if focus == state.edit {
                SetFocus(state.ok_button);
            } else {
                SetFocus(state.edit);
            }
        } else {
            if focus == state.edit {
                SetFocus(state.ok_button);
            } else {
                SetFocus(state.edit);
            }
        }
    });
}

unsafe extern "system" fn help_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*create_struct).lpCreateParams as *mut HelpWindowInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = Box::from_raw(init_ptr);
            let parent = init.parent;
            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));

            let edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_VSCROLL
                    | WINDOW_STYLE(ES_MULTILINE as u32)
                    | WINDOW_STYLE(ES_AUTOVSCROLL as u32)
                    | WINDOW_STYLE(ES_WANTRETURN as u32)
                    | WS_TABSTOP,
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            SendMessageW(
                edit,
                windows::Win32::UI::Controls::EM_SETREADONLY,
                WPARAM(1),
                LPARAM(0),
            );
            if hfont.0 != 0 {
                let _ = SendMessageW(edit, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }

            let ok_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                w!("OK"),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                0,
                0,
                0,
                0,
                hwnd,
                HMENU(HELP_ID_OK as isize),
                HINSTANCE(0),
                None,
            );
            if hfont.0 != 0 && ok_button.0 != 0 {
                let _ = SendMessageW(ok_button, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }

            let content = match init.kind {
                HelpWindowKind::Guide => match init.language {
                    Language::Italian => include_str!("../../guida.txt").to_string(),
                    Language::English => include_str!("../../guida_en.txt").to_string(),
                    Language::Spanish => include_str!("../../guida_es.txt").to_string(),
                    Language::Portuguese => include_str!("../../guida_pt.txt").to_string(),
                },
                HelpWindowKind::Changelog => match init.language {
                    Language::Italian => include_str!("../../CHANGELOG_IT.md").to_string(),
                    Language::English => include_str!("../../CHANGELOG.md").to_string(),
                    Language::Spanish => include_str!("../../CHANGELOG_ES.md").to_string(),
                    Language::Portuguese => include_str!("../../CHANGELOG_PT.md").to_string(),
                },
                HelpWindowKind::Donations => donations_content(init.language),
            };
            let content = normalize_to_crlf(&content);
            let content_wide = to_wide(&content);
            let _ = SetWindowTextW(edit, PCWSTR(content_wide.as_ptr()));
            SetFocus(edit);

            let state = Box::new(HelpWindowState {
                parent,
                edit,
                ok_button,
                kind: init.kind,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            LRESULT(0)
        }
        WM_SETFOCUS => {
            let _ = with_help_state(hwnd, |state| {
                SetFocus(state.edit);
            });
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            if cmd_id == HELP_ID_OK || cmd_id == IDCANCEL.0 as usize {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_SIZE => {
            let width = (lparam.0 & 0xffff) as i32;
            let height = ((lparam.0 >> 16) & 0xffff) as i32;
            let _ = with_help_state(hwnd, |state| {
                let button_width = 90;
                let button_height = 28;
                let margin = 12;
                let edit_height = (height - button_height - (margin * 2)).max(0);
                let _ = MoveWindow(state.edit, 0, 0, width, edit_height, true);
                let _ = MoveWindow(
                    state.ok_button,
                    width - button_width - margin,
                    edit_height + margin,
                    button_width,
                    button_height,
                    true,
                );
            });
            LRESULT(0)
        }
        WM_DESTROY => {
            let (parent, kind) = with_help_state(hwnd, |state| (state.parent, state.kind))
                .unwrap_or((HWND(0), HelpWindowKind::Guide));
            if parent.0 != 0 {
                let _ = with_state(parent, |state| match kind {
                    HelpWindowKind::Guide => state.help_window = HWND(0),
                    HelpWindowKind::Changelog => state.changelog_window = HWND(0),
                    HelpWindowKind::Donations => state.donations_window = HWND(0),
                });
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut HelpWindowState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let _ = with_help_state(hwnd, |state| {
                    if GetFocus() == state.ok_button {
                        let _ = DestroyWindow(hwnd);
                    }
                });
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_help_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut HelpWindowState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut HelpWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

fn help_title(language: Language, kind: HelpWindowKind) -> String {
    match kind {
        HelpWindowKind::Guide => i18n::tr(language, "help.window.guide"),
        HelpWindowKind::Changelog => i18n::tr(language, "help.window.changelog"),
        HelpWindowKind::Donations => i18n::tr(language, "help.window.donations"),
    }
}

fn donations_content(language: Language) -> String {
    match language {
        Language::Italian => DONATIONS_IT.to_string(),
        Language::English => DONATIONS_EN.to_string(),
        Language::Spanish => DONATIONS_ES.to_string(),
        Language::Portuguese => DONATIONS_PT.to_string(),
    }
}
