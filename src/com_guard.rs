//! RAII guard for COM initialization.
//!
//! This module provides a thread-safe COM initialization guard that ensures
//! proper pairing of CoInitializeEx/CoUninitialize calls.

use windows::Win32::Foundation::RPC_E_CHANGED_MODE;
use windows::Win32::System::Com::{
    COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, CoInitializeEx, CoUninitialize,
};

/// RAII guard for COM apartment-threaded initialization.
/// Automatically calls CoUninitialize when dropped.
pub struct ComGuard {
    should_uninit: bool,
}

impl ComGuard {
    /// Initialize COM with apartment-threaded model (STA).
    /// Use this for UI threads and most SAPI operations.
    pub fn new_sta() -> Result<Self, windows::core::Error> {
        Self::init(COINIT_APARTMENTTHREADED)
    }

    /// Initialize COM with multi-threaded model (MTA).
    /// Use this for background worker threads.
    pub fn new_mta() -> Result<Self, windows::core::Error> {
        Self::init(COINIT_MULTITHREADED)
    }

    fn init(coinit: windows::Win32::System::Com::COINIT) -> Result<Self, windows::core::Error> {
        let result = unsafe { CoInitializeEx(None, coinit) };

        // S_OK = success, we initialized COM
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
        if let Err(ref e) = result.ok() {
            if e.code() == RPC_E_CHANGED_MODE {
                return Ok(Self {
                    should_uninit: false,
                });
            }
        }

        // Other errors - propagate
        result.ok()?;
        Ok(Self {
            should_uninit: false,
        })
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        if self.should_uninit {
            unsafe { CoUninitialize() };
        }
    }
}
