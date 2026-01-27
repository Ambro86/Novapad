//! Text input dialog for getting a single line of text from the user.
//!
//! This module provides a modal dialog with a text input field.
//! It replicates the accessibility behavior of the original novapad implementation:
//! - EDIT with WS_TABSTOP for keyboard navigation
//! - SetFocus on EDIT at creation (NVDA announces the field)
//! - IsDialogMessageW for Tab navigation
//! - WS_EX_CONTROLPARENT for proper child navigation
//! - ESC to cancel, ENTER to confirm

use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::WC_BUTTON;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, GetWindowTextLengthW,
    GetWindowTextW, HCURSOR, HMENU, IDC_ARROW, IsDialogMessageW, IsWindow, LoadCursorW, MSG,
    PostMessageW, RegisterClassW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW,
    TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_KEYDOWN, WM_NCDESTROY,
    WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_POPUP, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

use crate::{FontHandle, PlatformError, WindowHandle, to_wide};

const TEXT_INPUT_CLASS_NAME: &str = "PlatformWindowsTextInput";
const ID_EDIT: usize = 1801;
const ID_OK: usize = 1802;
const ID_CANCEL: usize = 1803;

/// Parameters for the text input dialog.
pub struct TextInputDialogParams<'a> {
    pub title: &'a str,
    pub label: &'a str,
    pub hint: &'a str,
    pub ok_label: &'a str,
    pub cancel_label: &'a str,
    pub font: Option<FontHandle>,
}

struct TextInputInit {
    label: String,
    hint: String,
    ok_label: String,
    cancel_label: String,
    font: HFONT,
    result: Arc<Mutex<Option<String>>>,
}

struct TextInputState {
    edit: HWND,
    ok_btn: HWND,
    result: Arc<Mutex<Option<String>>>,
}

/// Shows a modal text input dialog.
///
/// Returns:
/// - `Ok(Some(text))` if user entered text and pressed OK
/// - `Ok(None)` if user cancelled
/// - `Err(PlatformError)` if dialog creation failed
///
/// # Accessibility
/// - Focus is set to the EDIT field on creation
/// - Tab navigates between EDIT and buttons
/// - ESC cancels, ENTER confirms
pub fn show_text_input_dialog(
    parent: Option<WindowHandle>,
    params: TextInputDialogParams,
) -> Result<Option<String>, PlatformError> {
    let parent_hwnd = parent.map(|h| h.raw()).unwrap_or(HWND(0));

    // SAFETY: GetModuleHandleW with None returns the current module handle
    let hinstance = HINSTANCE(unsafe {
        GetModuleHandleW(None)
            .map_err(|e| PlatformError::Win32Error(format!("GetModuleHandleW failed: {}", e)))?
            .0
    });

    let class_name = to_wide(TEXT_INPUT_CLASS_NAME);

    let wc = WNDCLASSW {
        hCursor: HCURSOR(
            unsafe { LoadCursorW(None, IDC_ARROW) }
                .unwrap_or_default()
                .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(text_input_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };

    // SAFETY: RegisterClassW is safe to call, may fail if class already registered (ignored)
    unsafe { RegisterClassW(&wc) };

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(TextInputInit {
        label: params.label.to_string(),
        hint: params.hint.to_string(),
        ok_label: params.ok_label.to_string(),
        cancel_label: params.cancel_label.to_string(),
        font: params.font.map(|f| f.raw()).unwrap_or(HFONT(0)),
        result: result.clone(),
    });

    let title_wide = to_wide(params.title);

    // SAFETY: CreateWindowExW with valid parameters
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            360,
            180,
            parent_hwnd,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };

    if hwnd.0 == 0 {
        return Err(PlatformError::Win32Error(
            "Failed to create text input dialog".to_string(),
        ));
    }

    // Make dialog modal: disable parent and bring dialog to foreground
    if parent_hwnd.0 != 0 {
        // SAFETY: EnableWindow is safe with valid HWND
        unsafe { EnableWindow(parent_hwnd, false) };
    }
    // SAFETY: SetForegroundWindow is safe with valid HWND
    unsafe { SetForegroundWindow(hwnd) };

    // Message loop with IsDialogMessageW for accessibility
    let mut msg = MSG::default();
    loop {
        // SAFETY: IsWindow checks if window still exists
        if !unsafe { IsWindow(hwnd).as_bool() } {
            break;
        }

        // SAFETY: GetMessageW blocks until message available
        let res = unsafe { GetMessageW(&mut msg, HWND(0), 0, 0) };
        if res.0 == 0 {
            break;
        }

        // SAFETY: Message handling
        unsafe {
            // ESC closes dialog (cancel)
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(ID_CANCEL), LPARAM(0));
                continue;
            }

            // ENTER confirms (when focus is on EDIT or OK button)
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let (edit, ok_btn) =
                    with_text_input_state(hwnd, |state| (state.edit, state.ok_btn))
                        .unwrap_or((HWND(0), HWND(0)));
                let focus = GetFocus();
                if focus == edit || focus == ok_btn {
                    let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(ID_OK), LPARAM(0));
                    continue;
                }
            }

            // IsDialogMessageW handles Tab navigation and control interaction
            if IsDialogMessageW(hwnd, &msg).as_bool() {
                continue;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    // Restore parent window
    if parent_hwnd.0 != 0 {
        // SAFETY: EnableWindow and SetForegroundWindow are safe with valid HWND
        unsafe {
            EnableWindow(parent_hwnd, true);
            SetForegroundWindow(parent_hwnd);
        }
    }

    // Extract result
    let guard = result.lock().unwrap_or_else(|e| e.into_inner());
    Ok(guard.clone())
}

/// Window procedure for the text input dialog.
///
/// # Safety
/// This is a Win32 callback - must handle all messages safely.
unsafe extern "system" fn text_input_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            // SAFETY: lparam points to CREATESTRUCTW during WM_CREATE
            let create_struct = unsafe { &*(lparam.0 as *const CREATESTRUCTW) };
            let init_ptr = create_struct.lpCreateParams as *mut TextInputInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }

            // SAFETY: We own this box, take ownership
            let init = unsafe { Box::from_raw(init_ptr) };

            // Create label (STATIC)
            // SAFETY: CreateWindowExW with valid parameters
            let label = unsafe {
                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&init.label).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    10,
                    12,
                    330,
                    16,
                    hwnd,
                    HMENU(1),
                    HINSTANCE(0),
                    None,
                )
            };

            // Create EDIT input with WS_TABSTOP and WS_EX_CLIENTEDGE for 3D border
            // SAFETY: CreateWindowExW with valid parameters
            let edit = unsafe {
                CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    10,
                    30,
                    160,
                    24,
                    hwnd,
                    HMENU(ID_EDIT as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            // Create hint (STATIC)
            // SAFETY: CreateWindowExW with valid parameters
            let hint = unsafe {
                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&init.hint).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    10,
                    58,
                    330,
                    16,
                    hwnd,
                    HMENU(2),
                    HINSTANCE(0),
                    None,
                )
            };

            // Create OK button with BS_DEFPUSHBUTTON
            // SAFETY: CreateWindowExW with valid parameters
            let ok_btn = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&init.ok_label).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    170,
                    110,
                    80,
                    26,
                    hwnd,
                    HMENU(ID_OK as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            // Create Cancel button
            // SAFETY: CreateWindowExW with valid parameters
            let cancel_btn = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&init.cancel_label).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    260,
                    110,
                    80,
                    26,
                    hwnd,
                    HMENU(ID_CANCEL as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            // Apply font if provided
            if init.font.0 != 0 {
                // SAFETY: SendMessageW with valid handles
                unsafe {
                    SendMessageW(label, WM_SETFONT, WPARAM(init.font.0 as usize), LPARAM(1));
                    SendMessageW(edit, WM_SETFONT, WPARAM(init.font.0 as usize), LPARAM(1));
                    SendMessageW(hint, WM_SETFONT, WPARAM(init.font.0 as usize), LPARAM(1));
                    SendMessageW(ok_btn, WM_SETFONT, WPARAM(init.font.0 as usize), LPARAM(1));
                    SendMessageW(
                        cancel_btn,
                        WM_SETFONT,
                        WPARAM(init.font.0 as usize),
                        LPARAM(1),
                    );
                }
            }

            // ACCESSIBILITY CRITICAL: Set focus to EDIT field
            // SAFETY: SetFocus with valid handle
            unsafe { SetFocus(edit) };

            // Store state for later access
            let state = Box::new(TextInputState {
                edit,
                ok_btn,
                result: init.result.clone(),
            });

            // SAFETY: SetWindowLongPtrW stores our state pointer
            unsafe { SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize) };

            LRESULT(0)
        }

        WM_COMMAND => {
            let id = wparam.0 & 0xffff;
            match id {
                ID_OK => {
                    // Get text from EDIT and store in result
                    with_text_input_state(hwnd, |state| {
                        // SAFETY: GetWindowTextLengthW with valid handle
                        let len = unsafe { GetWindowTextLengthW(state.edit) };
                        if len > 0 {
                            let mut buf = vec![0u16; (len + 1) as usize];
                            // SAFETY: GetWindowTextW with valid handle and sized buffer
                            let actual_len = unsafe { GetWindowTextW(state.edit, &mut buf) };
                            if actual_len > 0 {
                                let text = String::from_utf16_lossy(&buf[..actual_len as usize]);
                                if let Ok(mut guard) = state.result.lock() {
                                    *guard = Some(text);
                                }
                            }
                        } else {
                            // Empty input - store empty string
                            if let Ok(mut guard) = state.result.lock() {
                                *guard = Some(String::new());
                            }
                        }
                    });
                    // SAFETY: DestroyWindow with valid handle
                    let _ = unsafe { DestroyWindow(hwnd) };
                    LRESULT(0)
                }
                ID_CANCEL => {
                    // SAFETY: DestroyWindow with valid handle
                    let _ = unsafe { DestroyWindow(hwnd) };
                    LRESULT(0)
                }
                _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
            }
        }

        WM_CLOSE => {
            // SAFETY: DestroyWindow with valid handle
            let _ = unsafe { DestroyWindow(hwnd) };
            LRESULT(0)
        }

        WM_NCDESTROY => {
            // Clean up state
            // SAFETY: GetWindowLongPtrW retrieves our stored pointer
            let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut TextInputState;
            if !ptr.is_null() {
                // SAFETY: We own this box, drop it
                let _ = unsafe { Box::from_raw(ptr) };
            }
            LRESULT(0)
        }

        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

/// Helper to access dialog state safely.
fn with_text_input_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut TextInputState) -> R,
{
    // SAFETY: GetWindowLongPtrW retrieves our stored pointer
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut TextInputState;
    if ptr.is_null() {
        None
    } else {
        // SAFETY: Pointer is valid if not null (we set it in WM_CREATE)
        Some(f(unsafe { &mut *ptr }))
    }
}
