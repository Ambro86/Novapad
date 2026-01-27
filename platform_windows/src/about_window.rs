use crate::{PlatformError, WindowHandle, to_wide};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{MB_ICONINFORMATION, MB_OK, MessageBoxW};
use windows::core::PCWSTR;

pub fn open(parent: Option<WindowHandle>, title: &str, message: &str) -> Result<(), PlatformError> {
    let parent_hwnd = parent.map(|h| h.raw()).unwrap_or(HWND(0));
    let title_wide = to_wide(title);
    let message_wide = to_wide(message);

    // SAFETY: MessageBoxW is a standard, safe Win32 call when provided with valid UTF-16 strings.
    unsafe {
        MessageBoxW(
            parent_hwnd,
            PCWSTR(message_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            MB_OK | MB_ICONINFORMATION,
        );
    }

    Ok(())
}
