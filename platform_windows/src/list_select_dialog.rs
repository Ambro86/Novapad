//! List selection dialog for choosing an item from a list.
//!
//! This module provides a modal dialog with a listbox for selecting one item.
//! It replicates the accessibility behavior of the original novapad implementation:
//! - LISTBOX with WS_TABSTOP for keyboard navigation
//! - LB_SETCURSEL(0) to pre-select first item (NVDA announces it)
//! - SetFocus on listbox at creation
//! - IsDialogMessageW for Tab navigation
//! - WS_EX_CONTROLPARENT for proper child navigation
//! - ESC to cancel, ENTER to confirm

use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::WC_BUTTON;
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus, VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, DestroyWindow,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, GetWindowLongPtrW, HCURSOR, HMENU, IDC_ARROW,
    IsDialogMessageW, IsWindow, LB_ADDSTRING, LB_GETCURSEL, LB_GETTEXT, LB_GETTEXTLEN,
    LB_SETCURSEL, LBS_NOTIFY, LoadCursorW, MSG, PostMessageW, RegisterClassW, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_KEYDOWN, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

use crate::{FontHandle, PlatformError, WindowHandle, to_wide};

const LIST_SELECT_CLASS_NAME: &str = "PlatformWindowsListSelect";
const ID_LIST: usize = 9201;
const ID_OK: usize = 9202;
const ID_CANCEL: usize = 9203;

/// Parameters for the list selection dialog.
pub struct ListSelectDialogParams<'a> {
    pub title: &'a str,
    pub items: &'a [String],
    pub ok_label: &'a str,
    pub cancel_label: &'a str,
    pub font: Option<FontHandle>,
}

struct ListSelectInit {
    items: Vec<String>,
    ok_label: String,
    cancel_label: String,
    font: HFONT,
    result: Arc<Mutex<Option<String>>>,
}

struct ListSelectState {
    list: HWND,
    result: Arc<Mutex<Option<String>>>,
}

/// Shows a modal list selection dialog.
///
/// Returns:
/// - `Ok(Some(selected_item))` if user selected an item and pressed OK
/// - `Ok(None)` if user cancelled
/// - `Err(PlatformError)` if dialog creation failed
///
/// # Accessibility
/// - Focus is set to the listbox on creation
/// - First item is pre-selected so NVDA announces it
/// - Tab navigates between listbox and buttons
/// - ESC cancels, ENTER confirms
pub fn show_list_select_dialog(
    parent: Option<WindowHandle>,
    params: ListSelectDialogParams,
) -> Result<Option<String>, PlatformError> {
    if params.items.is_empty() {
        return Ok(None);
    }

    let parent_hwnd = parent.map(|h| h.raw()).unwrap_or(HWND(0));

    // SAFETY: GetModuleHandleW with None returns the current module handle
    let hinstance = HINSTANCE(unsafe {
        GetModuleHandleW(None)
            .map_err(|e| PlatformError::Win32Error(format!("GetModuleHandleW failed: {}", e)))?
            .0
    });

    let class_name = to_wide(LIST_SELECT_CLASS_NAME);

    let wc = WNDCLASSW {
        hCursor: HCURSOR(
            unsafe { LoadCursorW(None, IDC_ARROW) }
                .unwrap_or_default()
                .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(list_select_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };

    // SAFETY: RegisterClassW is safe to call, may fail if class already registered (ignored)
    unsafe { RegisterClassW(&wc) };

    let result = Arc::new(Mutex::new(None));
    let init = Box::new(ListSelectInit {
        items: params.items.to_vec(),
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
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            500,
            300,
            parent_hwnd,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };

    if hwnd.0 == 0 {
        return Err(PlatformError::Win32Error(
            "Failed to create list select dialog".to_string(),
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
            // ESC closes dialog
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                continue;
            }

            // ENTER confirms selection
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(ID_OK), LPARAM(0));
                continue;
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

/// Window procedure for the list select dialog.
///
/// # Safety
/// This is a Win32 callback - must handle all messages safely.
unsafe extern "system" fn list_select_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            // SAFETY: lparam points to CREATESTRUCTW during WM_CREATE
            let create_struct = unsafe { &*(lparam.0 as *const CREATESTRUCTW) };
            let init_ptr = create_struct.lpCreateParams as *mut ListSelectInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }

            // SAFETY: We own this box, take ownership
            let init = unsafe { Box::from_raw(init_ptr) };

            // Create listbox with accessibility-critical styles:
            // - WS_TABSTOP: allows Tab navigation
            // - LBS_NOTIFY: sends notification messages
            // - WS_VSCROLL: scrollbar for long lists
            // SAFETY: CreateWindowExW with valid parameters
            let list = unsafe {
                CreateWindowExW(
                    Default::default(),
                    w!("LISTBOX"),
                    PCWSTR::null(),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_TABSTOP
                        | WS_VSCROLL
                        | WINDOW_STYLE(LBS_NOTIFY as u32),
                    10,
                    10,
                    460,
                    200,
                    hwnd,
                    HMENU(ID_LIST as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            // Add items to listbox
            for item in init.items.iter() {
                let item_wide = to_wide(item);
                // SAFETY: SendMessageW with valid listbox handle
                unsafe {
                    SendMessageW(
                        list,
                        LB_ADDSTRING,
                        WPARAM(0),
                        LPARAM(item_wide.as_ptr() as isize),
                    );
                }
            }

            // ACCESSIBILITY CRITICAL: Pre-select first item so NVDA announces it
            // SAFETY: SendMessageW with valid listbox handle
            unsafe { SendMessageW(list, LB_SETCURSEL, WPARAM(0), LPARAM(0)) };

            // ACCESSIBILITY CRITICAL: Set focus to listbox
            // SAFETY: SetFocus with valid handle
            unsafe { SetFocus(list) };

            // Create OK button with BS_DEFPUSHBUTTON (Enter activates it)
            // SAFETY: CreateWindowExW with valid parameters
            let ok_btn = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&init.ok_label).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    280,
                    220,
                    90,
                    28,
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
                    380,
                    220,
                    90,
                    28,
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
                    SendMessageW(list, WM_SETFONT, WPARAM(init.font.0 as usize), LPARAM(1));
                    SendMessageW(ok_btn, WM_SETFONT, WPARAM(init.font.0 as usize), LPARAM(1));
                    SendMessageW(
                        cancel_btn,
                        WM_SETFONT,
                        WPARAM(init.font.0 as usize),
                        LPARAM(1),
                    );
                }
            }

            // Store state for later access
            let state = Box::new(ListSelectState {
                list,
                result: init.result.clone(),
            });

            // SAFETY: SetWindowLongPtrW stores our state pointer
            unsafe { SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize) };

            LRESULT(0)
        }

        WM_COMMAND => {
            let id = wparam.0;
            match id {
                ID_OK => {
                    // Get selected item and store in result
                    with_list_state(hwnd, |state| {
                        // SAFETY: SendMessageW with valid listbox handle
                        let sel =
                            unsafe { SendMessageW(state.list, LB_GETCURSEL, WPARAM(0), LPARAM(0)) }
                                .0;
                        if sel >= 0 {
                            let len = unsafe {
                                SendMessageW(
                                    state.list,
                                    LB_GETTEXTLEN,
                                    WPARAM(sel as usize),
                                    LPARAM(0),
                                )
                            }
                            .0;
                            if len >= 0 {
                                let mut buf = vec![0u16; (len + 1) as usize];
                                // SAFETY: Buffer is correctly sized
                                unsafe {
                                    SendMessageW(
                                        state.list,
                                        LB_GETTEXT,
                                        WPARAM(sel as usize),
                                        LPARAM(buf.as_mut_ptr() as isize),
                                    );
                                }
                                let text = String::from_utf16_lossy(&buf[..len as usize]);
                                if let Ok(mut guard) = state.result.lock() {
                                    *guard = Some(text);
                                }
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
            let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ListSelectState;
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
fn with_list_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut ListSelectState) -> R,
{
    // SAFETY: GetWindowLongPtrW retrieves our stored pointer
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut ListSelectState;
    if ptr.is_null() {
        None
    } else {
        // SAFETY: Pointer is valid if not null (we set it in WM_CREATE)
        Some(f(unsafe { &mut *ptr }))
    }
}
