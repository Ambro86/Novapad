use crate::{PlatformError, WindowHandle, to_wide};
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{PCWSTR, w};

const SS_CENTER: u32 = 0x1;
const ABOUT_CLASS_NAME: &str = "NovapadAbout";

struct AboutState {
    message: String,
}

pub fn open(parent: Option<WindowHandle>, title: &str, message: &str) -> Result<(), PlatformError> {
    // SAFETY: GetModuleHandleW is safe with None.
    let hinstance =
        unsafe { GetModuleHandleW(None).map_err(|e| PlatformError::Win32Error(e.to_string()))? };

    let class_name = to_wide(ABOUT_CLASS_NAME);

    // SAFETY: Registering a window class is a standard Win32 operation.
    unsafe {
        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hInstance: hinstance.into(),
            lpszClassName: PCWSTR(class_name.as_ptr()),
            lpfnWndProc: Some(about_wndproc),
            ..Default::default()
        };
        RegisterClassW(&wc);
    }

    let state = Box::new(AboutState {
        message: message.to_string(),
    });
    let state_ptr = Box::into_raw(state);

    let parent_hwnd = parent.map(|h| h.raw()).unwrap_or(HWND(0));
    let title_wide = to_wide(title);

    // SAFETY: CreateWindowExW is safe when called with valid parameters.
    // We pass ownership of the state Box to the window via lpCreateParams.
    unsafe {
        let hwnd = CreateWindowExW(
            WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            400,
            250,
            parent_hwnd,
            HMENU(0),
            hinstance,
            Some(state_ptr as *const _),
        );

        if hwnd.0 == 0 {
            // Re-take ownership to drop it if window creation failed
            let _ = Box::from_raw(state_ptr);
            return Err(PlatformError::Win32Error(
                "Failed to create about window".to_string(),
            ));
        }
    }

    Ok(())
}

unsafe extern "system" fn about_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            // SAFETY: lpCreateParams is guaranteed to be the state_ptr we passed.
            unsafe {
                let cs = lparam.0 as *const CREATESTRUCTW;
                let state_ptr = (*cs).lpCreateParams as *mut AboutState;
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, state_ptr as isize);

                let state = &*state_ptr;
                let hinstance = GetModuleHandleW(None).unwrap_or_default();

                // Create message static control
                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&state.message).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WINDOW_STYLE(SS_CENTER),
                    20,
                    40,
                    360,
                    100,
                    hwnd,
                    HMENU(0),
                    hinstance,
                    None,
                );

                // Create OK button
                CreateWindowExW(
                    Default::default(),
                    w!("BUTTON"),
                    w!("OK"),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    160,
                    160,
                    80,
                    30,
                    hwnd,
                    HMENU(IDOK.0 as isize),
                    hinstance,
                    None,
                );
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            if wparam.0 & 0xffff == IDOK.0 as usize {
                // SAFETY: DestroyWindow is safe for own windows.
                unsafe {
                    let _ = DestroyWindow(hwnd);
                }
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            // SAFETY: We retrieve the Box we created earlier and let it drop.
            unsafe {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AboutState;
                if !ptr.is_null() {
                    let _ = Box::from_raw(ptr);
                }
            }
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}
