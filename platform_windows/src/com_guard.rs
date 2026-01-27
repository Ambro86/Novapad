//! RAII guard for COM initialization.

use crate::PlatformError;
use windows::Win32::Foundation::RPC_E_CHANGED_MODE;
use windows::Win32::System::Com::{
    COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, CoInitializeEx, CoUninitialize,
};

/// RAII guard for COM initialization.
/// Automatically calls CoUninitialize when dropped.
pub struct ComGuard {
    should_uninit: bool,
}

impl ComGuard {
    /// Initialize COM with apartment-threaded model (STA).
    pub fn new_sta() -> Result<Self, PlatformError> {
        Self::init(COINIT_APARTMENTTHREADED)
    }

    /// Initialize COM with multi-threaded model (MTA).
    pub fn new_mta() -> Result<Self, PlatformError> {
        Self::init(COINIT_MULTITHREADED)
    }

    fn init(coinit: windows::Win32::System::Com::COINIT) -> Result<Self, PlatformError> {
        // SAFETY: CoInitializeEx is a standard Win32 call.
        // We properly handle the pairing with CoUninitialize via the Drop trait.
        let result = unsafe { CoInitializeEx(None, coinit) };

        if result.is_ok() {
            return Ok(Self {
                should_uninit: true,
            });
        }

        // S_FALSE (HRESULT 1) = COM already initialized on this thread with same model
        if result == windows::core::HRESULT(1) {
            return Ok(Self {
                should_uninit: true,
            });
        }

        // RPC_E_CHANGED_MODE = COM already initialized with different model
        // We can still use COM, but shouldn't call CoUninitialize
        if let Err(ref e) = result.ok()
            && e.code() == RPC_E_CHANGED_MODE
        {
            return Ok(Self {
                should_uninit: false,
            });
        }

        result
            .ok()
            .map_err(|e| PlatformError::ComError(e.to_string()))?;

        Ok(Self {
            should_uninit: false,
        })
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        if self.should_uninit {
            // SAFETY: Pairing with CoInitializeEx is guaranteed by RAII and should_uninit flag.
            unsafe { CoUninitialize() };
        }
    }
}
