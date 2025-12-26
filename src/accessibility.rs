use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{MSG, IsDialogMessageW, WM_KEYDOWN};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    VK_SPACE, VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_HOME, VK_END, VK_PRIOR, VK_NEXT
};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;

pub enum PlayerAction {
    TogglePause,
    Seek(i64),
    Volume(f32),
    BlockNavigation,
    None,
}

pub const EM_GETSEL: u32 = 0x00B0;
pub const EM_EXSETSEL: u32 = 0x0400 + 55;
pub const EM_SCROLLCARET: u32 = 0x00B7;
pub const EM_REPLACESEL: u32 = 0x00C2;

pub const ES_CENTER: u32 = 0x0001;
pub const ES_READONLY: u32 = 0x0800;

/// Converts a Rust string slice to a wide string (UTF-16) null-terminated vector.
/// Essential for all Win32 UI calls.
pub fn to_wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Converts a Rust string slice to a wide string (UTF-16) null-terminated vector,
/// normalizing newlines to CRLF.
pub fn to_wide_normalized(text: &str) -> Vec<u16> {
    let normalized = text.replace("\r\n", "\n").replace('\n', "\r\n");
    to_wide(&normalized)
}

pub fn normalize_to_crlf(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\n', "\r\n")
}

/// Converts a wide string pointer (UTF-16) to a Rust String.
pub unsafe fn from_wide(ptr: *const u16) -> String {
    let mut len = 0;
    while *ptr.add(len) != 0 {
        len += 1;
    }
    let slice = std::slice::from_raw_parts(ptr, len);
    OsString::from_wide(slice).to_string_lossy().into_owned()
}

/// The "Golden Rule" accessibility handler.
/// This function ensures that standard navigation keys (TAB, ENTER, SPACE, ARROWS)
/// function correctly across all dialogs and windows.
///
/// Returns `true` if the message was handled and should be skipped by the main loop.
pub unsafe fn handle_accessibility(hwnd: HWND, msg: &MSG) -> bool {
    // 1. Standard Dialog Navigation (TAB, Arrows, Enter, Space on controls)
    // IsDialogMessageW handles the vast majority of accessibility rules automatically.
    if IsDialogMessageW(hwnd, msg).as_bool() {
        return true;
    }

    false
}

/// Handles keyboard accessibility for the player window (Audiobook).
/// Returns a PlayerAction indicating what the application should do.
pub fn handle_player_keyboard(msg: &MSG, skip_seconds: u32) -> PlayerAction {
    if msg.message == WM_KEYDOWN {
        match msg.wParam.0 as u32 {
            vk if vk == VK_SPACE.0 as u32 => PlayerAction::TogglePause,
            vk if vk == VK_LEFT.0 as u32 => PlayerAction::Seek(-(skip_seconds as i64)),
            vk if vk == VK_RIGHT.0 as u32 => PlayerAction::Seek(skip_seconds as i64),
            vk if vk == VK_UP.0 as u32 => PlayerAction::Volume(0.1),
            vk if vk == VK_DOWN.0 as u32 => PlayerAction::Volume(-0.1),
            // Block navigation to prevent screen reader noise
            vk if vk == VK_HOME.0 as u32 || vk == VK_END.0 as u32 ||
                  vk == VK_PRIOR.0 as u32 || vk == VK_NEXT.0 as u32 => PlayerAction::BlockNavigation,
            _ => PlayerAction::None,
        }
    } else {
        PlayerAction::None
    }
}
