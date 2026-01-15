use crate::settings;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetKeyState, VK_ADD, VK_CONTROL, VK_DOWN, VK_END, VK_ESCAPE, VK_HOME, VK_LEFT, VK_MENU,
    VK_NEXT, VK_OEM_MINUS, VK_OEM_PERIOD, VK_OEM_PLUS, VK_PRIOR, VK_RIGHT, VK_SHIFT, VK_SPACE,
    VK_SUBTRACT, VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{IsDialogMessageW, MSG, WM_KEYDOWN};

pub enum PlayerCommand {
    TogglePause,
    Stop,
    Seek(i64),
    Volume(f32),
    Speed(f32),
    MuteToggle,
    GoToTime,
    AnnounceTime,
    ChapterPrev,
    ChapterNext,
    ChapterList,
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
/// Returns a PlayerCommand indicating what the application should do.
pub fn handle_player_keyboard(msg: &MSG, skip_seconds: u32) -> PlayerCommand {
    if msg.message == WM_KEYDOWN {
        let ctrl_down = unsafe { (GetKeyState(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0 };
        let alt_down = unsafe { (GetKeyState(VK_MENU.0 as i32) & (0x8000u16 as i16)) != 0 };
        let shift_down = unsafe { (GetKeyState(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0 };
        match msg.wParam.0 as u32 {
            vk if alt_down && shift_down && vk == 'P' as u32 => PlayerCommand::ChapterPrev,
            vk if alt_down && shift_down && vk == 'N' as u32 => PlayerCommand::ChapterNext,
            vk if alt_down && shift_down && vk == 'L' as u32 => PlayerCommand::ChapterList,
            vk if ctrl_down && vk == 'T' as u32 => PlayerCommand::GoToTime,
            vk if ctrl_down && vk == 'I' as u32 => PlayerCommand::AnnounceTime,
            vk if vk == VK_SPACE.0 as u32 => PlayerCommand::TogglePause,
            vk if vk == VK_LEFT.0 as u32 => PlayerCommand::Seek(-(skip_seconds as i64)),
            vk if vk == VK_RIGHT.0 as u32 => PlayerCommand::Seek(skip_seconds as i64),
            vk if vk == VK_UP.0 as u32 => PlayerCommand::Volume(0.1),
            vk if vk == VK_DOWN.0 as u32 => PlayerCommand::Volume(-0.1),
            vk if vk == VK_OEM_PLUS.0 as u32 || vk == VK_ADD.0 as u32 => PlayerCommand::Speed(0.1),
            vk if vk == VK_OEM_MINUS.0 as u32 || vk == VK_SUBTRACT.0 as u32 => {
                PlayerCommand::Speed(-0.1)
            }
            vk if vk == VK_OEM_PERIOD.0 as u32 => PlayerCommand::Stop,
            vk if vk == VK_ESCAPE.0 as u32 => PlayerCommand::Stop,
            vk if vk == 'M' as u32 => PlayerCommand::MuteToggle,
            // Block navigation to prevent screen reader noise
            vk if vk == VK_HOME.0 as u32
                || vk == VK_END.0 as u32
                || vk == VK_PRIOR.0 as u32
                || vk == VK_NEXT.0 as u32 =>
            {
                PlayerCommand::BlockNavigation
            }
            _ => PlayerCommand::None,
        }
    } else {
        PlayerCommand::None
    }
}

static mut NVDA_HANDLE: isize = 0;
static mut NVDA_SPEAK_ADDR: usize = 0;

/// Attempts to speak text using the NVDA Controller Client DLL.
/// Returns true if the DLL was loaded and the function called, false otherwise.
pub fn nvda_speak(text: &str) -> bool {
    use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
    use windows::core::{PCSTR, PCWSTR};

    unsafe {
        if NVDA_HANDLE == 0 {
            let dll_name = if cfg!(target_arch = "x86_64") {
                "nvdaControllerClient64.dll"
            } else {
                "nvdaControllerClient32.dll"
            };
            let dll_path = settings::settings_dir().join(dll_name);
            let dll_path_wide = to_wide(&dll_path.to_string_lossy());
            if let Ok(h) = LoadLibraryW(PCWSTR(dll_path_wide.as_ptr())) {
                NVDA_HANDLE = h.0;
                let proc_name = std::ffi::CString::new("nvdaController_speakText").unwrap();
                if let Some(addr) = GetProcAddress(
                    windows::Win32::Foundation::HMODULE(NVDA_HANDLE),
                    PCSTR(proc_name.as_ptr() as *const u8),
                ) {
                    NVDA_SPEAK_ADDR = addr as usize;
                }
            }
        }

        if NVDA_SPEAK_ADDR != 0 {
            let func: unsafe extern "system" fn(*const u16) -> i32 =
                std::mem::transmute(NVDA_SPEAK_ADDR);
            let wide = to_wide(text);
            let _ = func(wide.as_ptr());
            return true;
        }
    }
    false
}

pub fn ensure_nvda_controller_client() {
    let dll_name = if cfg!(target_arch = "x86_64") {
        "nvdaControllerClient64.dll"
    } else {
        "nvdaControllerClient32.dll"
    };

    let dll_path = settings::settings_dir().join(dll_name);

    if dll_path.exists() {
        return;
    }

    let url = format!(
        "https://raw.githubusercontent.com/Ambro86/Novapad/master/dll/{}",
        dll_name
    );

    // Run download in a separate thread to not block startup
    std::thread::spawn(move || {
        if let Ok(response) = reqwest::blocking::get(&url) {
            if response.status().is_success() {
                if let Ok(bytes) = response.bytes() {
                    // Unique temporary file to avoid partial reads if multiple processes try this
                    // (though unlikely in this context, good practice)
                    let tmp_path = dll_path.with_extension("tmp");
                    if let Ok(mut file) = std::fs::File::create(&tmp_path) {
                        use std::io::Write;
                        if file.write_all(&bytes).is_ok() {
                            let _ = std::fs::rename(tmp_path, dll_path);
                        }
                    }
                }
            }
        }
    });
}

pub fn ensure_soundtouch_dll() {
    let dll_name = if cfg!(target_arch = "x86_64") {
        "SoundTouch64.dll"
    } else {
        "SoundTouch64.dll"
    };

    let dll_path = settings::settings_dir().join(dll_name);
    if dll_path.exists() {
        return;
    }

    let url = format!(
        "https://raw.githubusercontent.com/Ambro86/Novapad/master/dll/{}",
        dll_name
    );

    std::thread::spawn(move || {
        if let Ok(response) = reqwest::blocking::get(&url) {
            if response.status().is_success() {
                if let Ok(bytes) = response.bytes() {
                    let tmp_path = dll_path.with_extension("tmp");
                    if let Ok(mut file) = std::fs::File::create(&tmp_path) {
                        use std::io::Write;
                        if file.write_all(&bytes).is_ok() {
                            let _ = std::fs::rename(tmp_path, dll_path);
                        }
                    }
                }
            }
        }
    });
}
