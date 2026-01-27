#![deny(unsafe_op_in_unsafe_fn)]

use windows::Win32::Foundation::HWND;

mod com_guard;
pub use com_guard::ComGuard;

#[derive(Debug)]
pub enum PlatformError {
    ComError(String),
    Other(String),
}

impl std::fmt::Display for PlatformError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ComError(s) => write!(f, "COM Error: {}", s),
            Self::Other(s) => write!(f, "Platform Error: {}", s),
        }
    }
}

impl std::error::Error for PlatformError {}

/// A safe wrapper around a Win32 HWND.
/// In this phase, it does not implement Send/Sync unless necessary.
pub struct SafeHwnd(HWND);

impl SafeHwnd {
    /// Internal constructor for creating from raw Win32 HWND.
    /// This is NOT exported to novapad directly.
    #[allow(dead_code)]
    pub(crate) fn from_raw(hwnd: HWND) -> Self {
        Self(hwnd)
    }

    /// Access the raw HWND. Restricted to the crate.
    #[allow(dead_code)]
    pub(crate) fn raw(&self) -> HWND {
        self.0
    }

    pub fn is_valid(&self) -> bool {
        self.0.0 != 0
    }
}

pub fn init() -> Result<(), String> {
    Ok(())
}
