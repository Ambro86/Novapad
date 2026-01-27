#![deny(unsafe_op_in_unsafe_fn)]

use windows::Win32::Foundation::HWND;

pub mod about_window;
mod com_guard;
pub mod dialogs;

pub use com_guard::ComGuard;
pub use dialogs::{DialogResult, DialogStyle, show_message_box};

#[derive(Debug)]
pub enum PlatformError {
    ComError(String),
    Win32Error(String),
    Other(String),
}

impl std::fmt::Display for PlatformError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ComError(s) => write!(f, "COM Error: {}", s),
            Self::Win32Error(s) => write!(f, "Win32 Error: {}", s),
            Self::Other(s) => write!(f, "Platform Error: {}", s),
        }
    }
}

impl std::error::Error for PlatformError {}

/// A safe wrapper around a Win32 HWND.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct WindowHandle(HWND);

impl WindowHandle {
    /// Create from a raw handle value.
    /// This is used to bridge existing Win32 code during migration.
    pub fn from_isize(handle: isize) -> Self {
        Self(HWND(handle))
    }

    /// Internal constructor for creating from raw Win32 HWND.
    /// Restricted to the crate.
    #[allow(dead_code)]
    pub(crate) fn from_raw(hwnd: HWND) -> Self {
        Self(hwnd)
    }

    /// Access the raw HWND. Restricted to the crate.
    pub(crate) fn raw(&self) -> HWND {
        self.0
    }

    pub fn is_valid(&self) -> bool {
        self.0.0 != 0
    }
}

/// Converts a Rust string slice to a wide string (UTF-16) null-terminated vector.
pub fn to_wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

pub fn init() -> Result<(), String> {
    Ok(())
}
