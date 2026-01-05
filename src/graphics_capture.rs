//! Windows Graphics Capture API wrapper for screen recording
//!
//! This module provides monitor enumeration and screen capture functionality
//! using Windows.Graphics.Capture (WinRT) APIs.

use std::sync::Arc;
use windows::Graphics::Capture::{
    Direct3D11CaptureFramePool, GraphicsCaptureItem, GraphicsCaptureSession,
};
use windows::Graphics::DirectX::Direct3D11::IDirect3DDevice;
use windows::Graphics::DirectX::DirectXPixelFormat;
use windows::Win32::Foundation::{BOOL, LPARAM, RECT, TRUE};
use windows::Win32::Graphics::Direct3D::{D3D_DRIVER_TYPE_HARDWARE, D3D_FEATURE_LEVEL_11_0};
use windows::Win32::Graphics::Direct3D11::{
    D3D11_CREATE_DEVICE_BGRA_SUPPORT, D3D11_SDK_VERSION, D3D11CreateDevice, ID3D11Device,
    ID3D11Texture2D,
};
use windows::Win32::Graphics::Dxgi::IDXGIDevice;
use windows::Win32::Graphics::Gdi::{
    EnumDisplayMonitors, GetMonitorInfoW, HDC, HMONITOR, MONITORINFOEXW,
};
use windows::Win32::System::WinRT::Direct3D11::{
    CreateDirect3D11DeviceFromDXGIDevice, IDirect3DDxgiInterfaceAccess,
};
use windows::Win32::System::WinRT::Graphics::Capture::IGraphicsCaptureItemInterop;
use windows::Win32::UI::WindowsAndMessaging::MONITORINFOF_PRIMARY;
use windows::core::{Interface, Result as WinResult};

/// Information about a display monitor
#[derive(Clone, Debug)]
pub struct MonitorInfo {
    pub id: String,
    pub name: String,
    pub is_primary: bool,
    pub width: u32,
    pub height: u32,
    pub hmonitor: isize, // Store HMONITOR value for CreateForMonitor
}

/// Enumerate all available monitors
pub fn list_monitors() -> Result<Vec<MonitorInfo>, String> {
    unsafe {
        let mut monitors = Vec::new();

        let callback_data = &mut monitors as *mut Vec<MonitorInfo>;

        let success = EnumDisplayMonitors(
            None,
            None,
            Some(enum_monitors_callback),
            LPARAM(callback_data as isize),
        );

        if !success.as_bool() {
            return Err("EnumDisplayMonitors failed".to_string());
        }

        if monitors.is_empty() {
            return Err("No monitors found".to_string());
        }

        Ok(monitors)
    }
}

unsafe extern "system" fn enum_monitors_callback(
    hmonitor: HMONITOR,
    _hdc: HDC,
    _rect: *mut RECT,
    lparam: LPARAM,
) -> BOOL {
    let monitors = &mut *(lparam.0 as *mut Vec<MonitorInfo>);

    let mut info = MONITORINFOEXW {
        monitorInfo: windows::Win32::Graphics::Gdi::MONITORINFO {
            cbSize: std::mem::size_of::<MONITORINFOEXW>() as u32,
            ..Default::default()
        },
        ..Default::default()
    };

    if GetMonitorInfoW(hmonitor, &mut info.monitorInfo as *mut _ as *mut _).as_bool() {
        let device_name = String::from_utf16_lossy(&info.szDevice)
            .trim_end_matches('\0')
            .to_string();

        let rect = &info.monitorInfo.rcMonitor;
        let width = (rect.right - rect.left) as u32;
        let height = (rect.bottom - rect.top) as u32;
        let is_primary = (info.monitorInfo.dwFlags & MONITORINFOF_PRIMARY) != 0;

        let name = if is_primary {
            format!("{} ({}x{}) - Primary", device_name, width, height)
        } else {
            format!("{} ({}x{})", device_name, width, height)
        };

        monitors.push(MonitorInfo {
            id: device_name.clone(),
            name,
            is_primary,
            width,
            height,
            hmonitor: hmonitor.0,
        });
    }

    TRUE
}

/// Graphics Capture session wrapper
pub struct CaptureSession {
    session: GraphicsCaptureSession,
    frame_pool: Direct3D11CaptureFramePool,
}

// SAFETY: The underlying WinRT objects are thread-safe for cross-thread use
unsafe impl Send for CaptureSession {}
unsafe impl Sync for CaptureSession {}

impl CaptureSession {
    /// Create a new capture session for the specified monitor
    pub fn new_for_monitor(monitor_info: &MonitorInfo) -> Result<Arc<Self>, String> {
        unsafe {
            // Create D3D11 device
            let mut device: Option<ID3D11Device> = None;
            let feature_levels = [D3D_FEATURE_LEVEL_11_0];

            D3D11CreateDevice(
                None, // Use default adapter
                D3D_DRIVER_TYPE_HARDWARE,
                None,
                D3D11_CREATE_DEVICE_BGRA_SUPPORT, // Required for Graphics.Capture
                Some(&feature_levels),
                D3D11_SDK_VERSION,
                Some(&mut device),
                None,
                None,
            )
            .map_err(|e| format!("D3D11CreateDevice failed: {}", e))?;

            let device = device.ok_or("D3D11Device is None")?;

            // Convert ID3D11Device to IDirect3DDevice (WinRT)
            let dxgi_device: IDXGIDevice = device
                .cast()
                .map_err(|e| format!("Cast to IDXGIDevice failed: {}", e))?;

            let inspectable = CreateDirect3D11DeviceFromDXGIDevice(&dxgi_device)
                .map_err(|e| format!("CreateDirect3D11DeviceFromDXGIDevice failed: {}", e))?;

            let d3d_device: IDirect3DDevice = inspectable
                .cast()
                .map_err(|e| format!("Cast to IDirect3DDevice failed: {}", e))?;

            // Create GraphicsCaptureItem for the monitor using interop
            let hmonitor = HMONITOR(monitor_info.hmonitor);
            let interop =
                windows::core::factory::<GraphicsCaptureItem, IGraphicsCaptureItemInterop>()
                    .map_err(|e| {
                        format!("Failed to get IGraphicsCaptureItemInterop factory: {}", e)
                    })?;
            let item: GraphicsCaptureItem = interop.CreateForMonitor(hmonitor).map_err(|e| {
                format!(
                    "CreateForMonitor failed: {}. Monitor may have been disconnected.",
                    e
                )
            })?;

            let size = item
                .Size()
                .map_err(|e| format!("Failed to get item size: {}", e))?;

            // Create frame pool
            let frame_pool = Direct3D11CaptureFramePool::Create(
                &d3d_device,
                DirectXPixelFormat::B8G8R8A8UIntNormalized,
                2, // Keep 2 frames buffered
                size,
            )
            .map_err(|e| format!("Failed to create frame pool: {}", e))?;

            // Create capture session
            let session = frame_pool
                .CreateCaptureSession(&item)
                .map_err(|e| format!("Failed to create capture session: {}", e))?;

            // Enable cursor capture
            session
                .SetIsCursorCaptureEnabled(true)
                .map_err(|e| format!("Failed to set cursor capture: {}", e))?;

            Ok(Arc::new(CaptureSession {
                session,
                frame_pool,
            }))
        }
    }

    /// Start capturing frames
    pub fn start(&self) -> Result<(), String> {
        self.session
            .StartCapture()
            .map_err(|e| format!("StartCapture failed: {}", e))
    }

    /// Try to get the next captured frame
    pub fn try_get_next_frame(&self) -> WinResult<Option<ID3D11Texture2D>> {
        unsafe {
            let frame = match self.frame_pool.TryGetNextFrame() {
                Ok(f) => f,
                Err(_) => return Ok(None),
            };

            let surface = frame.Surface()?;

            // Convert WinRT surface to D3D11 texture
            let access: IDirect3DDxgiInterfaceAccess = surface.cast()?;
            let texture: ID3D11Texture2D = access.GetInterface()?;

            Ok(Some(texture))
        }
    }

    /// Stop capturing
    pub fn stop(&self) {
        let _ = self.session.Close();
        let _ = self.frame_pool.Close();
    }
}

impl Drop for CaptureSession {
    fn drop(&mut self) {
        self.stop();
    }
}
