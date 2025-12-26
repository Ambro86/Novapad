use windows::core::{PCWSTR, w};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM, LRESULT, HINSTANCE};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, RegisterClassW,
    SendMessageW, SetWindowLongPtrW, SetWindowTextW, SetForegroundWindow, MoveWindow,
    GWLP_USERDATA, WM_CREATE, WM_DESTROY, WM_NCDESTROY, WM_CLOSE, WM_SIZE, WM_COMMAND, WM_SETFOCUS,
    WM_KEYDOWN, WM_SETFONT,
    WS_OVERLAPPEDWINDOW, WS_VISIBLE, WS_CHILD, WS_TABSTOP, WS_VSCROLL, WS_EX_CONTROLPARENT,
    WS_EX_CLIENTEDGE, CW_USEDEFAULT, HMENU, WNDCLASSW,
    IDCANCEL, BS_DEFPUSHBUTTON, ES_MULTILINE, ES_AUTOVSCROLL, ES_WANTRETURN,
    CREATESTRUCTW, WINDOW_STYLE, LoadCursorW, IDC_ARROW
};
use windows::Win32::UI::Controls::{WC_BUTTON};
use windows::Win32::Graphics::Gdi::{HBRUSH, COLOR_WINDOW, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, SetFocus, VK_RETURN, VK_SHIFT, GetKeyState};
use crate::{with_state};
use crate::settings::Language;
use crate::accessibility::{to_wide, normalize_to_crlf};

const HELP_CLASS_NAME: &str = "NovapadHelp";
const HELP_ID_OK: usize = 7003;

struct HelpWindowState {
    parent: HWND,
    edit: HWND,
    ok_button: HWND,
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.help_window).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(HELP_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(LoadCursorW(None, IDC_ARROW).unwrap_or_default().0),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(help_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let title = to_wide(help_title(language));
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
        Some(parent.0 as *const std::ffi::c_void),
    );

    if window.0 != 0 {
        let _ = with_state(parent, |state| {
            state.help_window = window;
        });
        SetForegroundWindow(window);
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

unsafe extern "system" fn help_wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
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
            SendMessageW(edit, windows::Win32::UI::Controls::EM_SETREADONLY, WPARAM(1), LPARAM(0));
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

            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let guide_content = match language {
                Language::Italian => include_str!("../../guida.txt"),
                Language::English => include_str!("../../guida_en.txt"),
            };
            let guide = normalize_to_crlf(guide_content);
            let guide_wide = to_wide(&guide);
            let _ = SetWindowTextW(edit, PCWSTR(guide_wide.as_ptr()));
            SetFocus(edit);

            let state = Box::new(HelpWindowState { parent, edit, ok_button });
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
            let parent = with_help_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                let _ = with_state(parent, |state| {
                    state.help_window = HWND(0);
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

fn help_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Guida",
        Language::English => "Guide",
    }
}
