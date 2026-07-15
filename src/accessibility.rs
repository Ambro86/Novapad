use crate::settings;
use std::sync::{OnceLock, mpsc};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{
    CLSCTX_ALL, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoUninitialize, DISPATCH_METHOD,
    DISPPARAMS, EXCEPINFO, IDispatch,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    VK_ADD, VK_CONTROL, VK_DOWN, VK_END, VK_ESCAPE, VK_HOME, VK_LEFT, VK_MENU, VK_NEXT,
    VK_OEM_MINUS, VK_OEM_PERIOD, VK_OEM_PLUS, VK_PRIOR, VK_RIGHT, VK_SHIFT, VK_SPACE, VK_SUBTRACT,
    VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{MSG, WM_KEYDOWN};
use windows::core::{BSTR, GUID, PCWSTR, VARIANT};

pub const EM_GETSEL: u32 = 0x00B0;
pub const EM_EXSETSEL: u32 = 0x0400 + 55;
pub const EM_SCROLLCARET: u32 = 0x00B7;
pub const EM_REPLACESEL: u32 = 0x00C2;

pub const ES_CENTER: u32 = 0x0001;
pub const ES_READONLY: u32 = 0x0800;

/// Converts a Rust string slice to a wide string (UTF-16) null-terminated vector.
/// Embedded NUL characters are removed because Win32 treats them as end-of-string markers.
/// Essential for all Win32 UI calls.
pub fn to_wide(value: &str) -> Vec<u16> {
    value
        .encode_utf16()
        .filter(|unit| *unit != 0)
        .chain(std::iter::once(0))
        .collect()
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

/// The "Golden Rule" accessibility handler.
/// This function ensures that standard navigation keys (TAB, ENTER, SPACE, ARROWS)
/// function correctly across all dialogs and windows.
///
/// Returns `true` if the message was handled and should be skipped by the main loop.
pub fn handle_accessibility(hwnd: HWND, msg: &MSG) -> bool {
    // Only process messages that belong to this window or its children.
    // This prevents IsDialogMessageW from interfering with messages destined
    // for other windows (e.g., the main editor while a secondary window is open).
    if msg.hwnd != hwnd && !crate::is_child_safe(hwnd, msg.hwnd) {
        return false;
    }

    // 1. Standard Dialog Navigation (TAB, Arrows, Enter, Space on controls)
    // IsDialogMessageW handles the vast majority of accessibility rules automatically.
    if crate::is_dialog_message_w_safe(hwnd, msg).as_bool() {
        return true;
    }

    false
}

/// Handles keyboard accessibility for the player window (Audiobook).
/// Returns a PlayerCommand indicating what the application should do.
pub fn handle_player_keyboard(msg: &MSG, skip_seconds: u32) -> PlayerCommand {
    if msg.message == WM_KEYDOWN {
        let ctrl_down = (crate::get_key_state_safe(VK_CONTROL.0 as i32) & (0x8000u16 as i16)) != 0;
        let alt_down = (crate::get_key_state_safe(VK_MENU.0 as i32) & (0x8000u16 as i16)) != 0;
        let shift_down = (crate::get_key_state_safe(VK_SHIFT.0 as i32) & (0x8000u16 as i16)) != 0;
        let vk = msg.wParam.0 as u32;
        let base_skip = skip_seconds.max(1) as i64;
        let small_skip = (base_skip / 3).max(1);
        let large_skip = base_skip.saturating_mul(3);
        let seek_step = if ctrl_down {
            large_skip
        } else if shift_down {
            small_skip
        } else {
            base_skip
        };

        let cmd = match vk {
            vk if alt_down && shift_down && vk == 'P' as u32 => PlayerCommand::ChapterPrev,
            vk if alt_down && shift_down && vk == 'N' as u32 => PlayerCommand::ChapterNext,
            vk if alt_down && shift_down && vk == 'L' as u32 => PlayerCommand::ChapterList,
            vk if ctrl_down && !alt_down && !shift_down && vk == VK_PRIOR.0 as u32 => {
                PlayerCommand::TrackPrev
            }
            vk if ctrl_down && !alt_down && !shift_down && vk == VK_NEXT.0 as u32 => {
                PlayerCommand::TrackNext
            }
            vk if ctrl_down && alt_down && !shift_down && vk == VK_PRIOR.0 as u32 => {
                PlayerCommand::ChapterPrev
            }
            vk if ctrl_down && alt_down && !shift_down && vk == VK_NEXT.0 as u32 => {
                PlayerCommand::ChapterNext
            }
            vk if ctrl_down && !alt_down && !shift_down && vk == 'T' as u32 => {
                PlayerCommand::GoToTime
            }
            vk if ctrl_down && vk == 'I' as u32 => PlayerCommand::AnnounceTime,
            vk if vk == VK_SPACE.0 as u32 => PlayerCommand::TogglePause,
            vk if !ctrl_down && !alt_down && !shift_down && vk == VK_HOME.0 as u32 => {
                PlayerCommand::SeekToStart
            }
            vk if !ctrl_down && !alt_down && !shift_down && vk == VK_END.0 as u32 => {
                PlayerCommand::SeekToEnd
            }
            vk if vk == VK_LEFT.0 as u32 => PlayerCommand::Seek(-seek_step),
            vk if vk == VK_RIGHT.0 as u32 => PlayerCommand::Seek(seek_step),
            vk if vk == VK_UP.0 as u32 => PlayerCommand::Volume(0.1),
            vk if vk == VK_DOWN.0 as u32 => PlayerCommand::Volume(-0.1),
            vk if vk == VK_OEM_PLUS.0 as u32 || vk == VK_ADD.0 as u32 => {
                if shift_down {
                    PlayerCommand::Pitch(1.0)
                } else {
                    PlayerCommand::Speed(0.1)
                }
            }
            vk if vk == VK_OEM_MINUS.0 as u32 || vk == VK_SUBTRACT.0 as u32 => {
                if shift_down {
                    PlayerCommand::Pitch(-1.0)
                } else {
                    PlayerCommand::Speed(-0.1)
                }
            }
            vk if vk == VK_OEM_PERIOD.0 as u32 => PlayerCommand::StopOnly,
            vk if vk == VK_ESCAPE.0 as u32 => PlayerCommand::Stop,
            vk if vk == 'M' as u32 => PlayerCommand::MuteToggle,
            // Block navigation to prevent screen reader noise
            vk if !ctrl_down
                && !alt_down
                && !shift_down
                && (vk == VK_PRIOR.0 as u32 || vk == VK_NEXT.0 as u32) =>
            {
                PlayerCommand::BlockNavigation
            }
            _ => PlayerCommand::None,
        };

        if !matches!(cmd, PlayerCommand::None) {
            crate::log_debug(&format!(
                "Keyboard: VK 0x{:X} triggered PlayerCommand::{:?}",
                vk, cmd
            ));
        }
        cmd
    } else {
        PlayerCommand::None
    }
}

#[derive(Debug)]
pub enum PlayerCommand {
    TogglePause,
    Stop,
    StopOnly,
    Seek(i64),
    Volume(f32),
    VolumeReset,
    Speed(f32),
    Pitch(f32),
    SpeedReset,
    PitchReset,
    MuteToggle,
    GoToTime,
    AnnounceTime,
    SeekToStart,
    SeekToEnd,
    ChapterPrev,
    ChapterNext,
    ChapterList,
    TrackPrev,
    TrackNext,
    BlockNavigation,
    None,
}

static NVDA_LIB: OnceLock<libloading::Library> = OnceLock::new();
static NVDA_SPEAK: OnceLock<Option<unsafe extern "C" fn(*const u16) -> i32>> = OnceLock::new();
static NVDA_BRAILLE: OnceLock<Option<unsafe extern "C" fn(*const u16) -> i32>> = OnceLock::new();
static NVDA_TEST: OnceLock<Option<unsafe extern "C" fn() -> i32>> = OnceLock::new();
static JAWS_PROBE_LOGGED: OnceLock<()> = OnceLock::new();
static BRAILLE_WORKER: OnceLock<mpsc::Sender<String>> = OnceLock::new();

fn nvda_init() -> bool {
    NVDA_SPEAK
        .get_or_init(|| {
            let dll_name = if cfg!(target_arch = "x86_64") {
                "nvdaControllerClient64.dll"
            } else {
                "nvdaControllerClient32.dll"
            };
            let dll_path = settings::settings_dir().join(dll_name);
            let lib = unsafe { libloading::Library::new(&dll_path).ok()? };
            let func = unsafe {
                let symbol: libloading::Symbol<unsafe extern "C" fn(*const u16) -> i32> =
                    lib.get(b"nvdaController_speakText\0").ok()?;
                *symbol
            };
            let braille = unsafe {
                lib.get(b"nvdaController_brailleMessage\0").ok().map(
                    |symbol: libloading::Symbol<unsafe extern "C" fn(*const u16) -> i32>| *symbol,
                )
            };
            let test = unsafe {
                let symbol: libloading::Symbol<unsafe extern "C" fn() -> i32> =
                    lib.get(b"nvdaController_testIfRunning\0").ok()?;
                *symbol
            };
            if NVDA_TEST.set(Some(test)).is_err() {
                return None;
            }
            if NVDA_BRAILLE.set(braille).is_err() {
                return None;
            }
            if NVDA_LIB.set(lib).is_err() {
                return None;
            }
            Some(func)
        })
        .is_some()
}

fn nvda_speak_raw(text: &str) -> bool {
    if !nvda_init() {
        return false;
    }
    let speak = match NVDA_SPEAK.get() {
        Some(Some(func)) => *func,
        _ => return false,
    };

    let wide = to_wide(text);
    unsafe {
        speak(wide.as_ptr());
    }
    true
}

fn nvda_braille_raw(text: &str) -> bool {
    if !nvda_init() {
        return false;
    }
    let braille = match NVDA_BRAILLE.get() {
        Some(Some(func)) => *func,
        _ => return false,
    };

    let wide = to_wide(text);
    unsafe {
        braille(wide.as_ptr());
    }
    true
}

fn braille_worker_sender() -> &'static mpsc::Sender<String> {
    BRAILLE_WORKER.get_or_init(|| {
        let (tx, rx) = mpsc::channel::<String>();
        std::thread::spawn(move || {
            while let Ok(text) = rx.recv() {
                if nvda_is_active() {
                    nvda_braille_raw(&text);
                }
                jaws_braille(&text);
            }
        });
        tx
    })
}

fn enqueue_braille(text: &str) {
    let send_res = braille_worker_sender().send(text.to_owned());
    if let Err(e) = send_res {
        crate::log_debug(&format!("Braille worker send failed: {e}"));
    }
}

fn nvda_is_active() -> bool {
    if !nvda_init() {
        return false;
    }
    match NVDA_TEST.get() {
        Some(Some(test)) => unsafe { test() == 0 },
        _ => false,
    }
}

fn jaws_invoke_saystring(
    text: &str,
    interrupt: bool,
    log_failures: bool,
) -> windows::core::Result<()> {
    let init_res = crate::com_guard::co_initialize_ex_safe(None, COINIT_APARTMENTTHREADED);
    if let Err(e) = init_res.ok() {
        if log_failures {
            crate::log_debug(&format!("JAWS CoInitializeEx failed: {e}"));
        }
        return Err(e);
    }
    let should_uninit = init_res.is_ok();

    let result = (|| {
        let prog_id = to_wide("FreedomSci.JawsApi");
        let clsid = match unsafe { CLSIDFromProgID(PCWSTR(prog_id.as_ptr())) } {
            Ok(id) => id,
            Err(e) => {
                if log_failures {
                    crate::log_debug(&format!("JAWS CLSIDFromProgID failed: {e}"));
                }
                return Err(e);
            }
        };

        let disp: IDispatch = match crate::co_create_instance_safe(&clsid, None, CLSCTX_ALL) {
            Ok(obj) => obj,
            Err(e) => {
                if log_failures {
                    crate::log_debug(&format!("JAWS CoCreateInstance failed: {e}"));
                }
                return Err(e);
            }
        };

        let mut dispid: i32 = 0;
        let name = to_wide("SayString");
        let names = [PCWSTR(name.as_ptr())];
        let getid_res =
            unsafe { disp.GetIDsOfNames(&GUID::default(), names.as_ptr(), 1, 0, &mut dispid) };
        if let Err(e) = getid_res {
            if log_failures {
                crate::log_debug(&format!("JAWS GetIDsOfNames failed: {e}"));
            }
            return Err(e);
        }

        let text_with_pad = format!("      {text}");
        let mut args = [
            VARIANT::from(if interrupt { 1i32 } else { 0i32 }),
            VARIANT::from(BSTR::from(text_with_pad)),
        ];
        let disp_params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: args.len() as u32,
            cNamedArgs: 0,
        };
        let mut excep = EXCEPINFO::default();
        let mut arg_err: u32 = 0;
        let invoke_res = unsafe {
            disp.Invoke(
                dispid,
                &GUID::default(),
                0,
                DISPATCH_METHOD,
                &disp_params,
                None,
                Some(&mut excep),
                Some(&mut arg_err),
            )
        };
        if log_failures && let Err(ref e) = invoke_res {
            crate::log_debug(&format!("JAWS Invoke failed: {e} (arg err: {arg_err})"));
        }
        invoke_res.map(|_| ())
    })();

    if should_uninit {
        unsafe { CoUninitialize() };
    }

    result
}

fn jaws_escape_function_string_arg(text: &str) -> String {
    text.replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace("\r\n", " ")
        .replace(['\r', '\n'], " ")
}

fn jaws_invoke_braille_message(text: &str, log_failures: bool) -> windows::core::Result<()> {
    let init_res = crate::com_guard::co_initialize_ex_safe(None, COINIT_APARTMENTTHREADED);
    if let Err(e) = init_res.ok() {
        if log_failures {
            crate::log_debug(&format!("JAWS CoInitializeEx failed: {e}"));
        }
        return Err(e);
    }
    let should_uninit = init_res.is_ok();

    let result = (|| {
        let prog_id = to_wide("FreedomSci.JawsApi");
        let clsid = match unsafe { CLSIDFromProgID(PCWSTR(prog_id.as_ptr())) } {
            Ok(id) => id,
            Err(e) => {
                if log_failures {
                    crate::log_debug(&format!("JAWS CLSIDFromProgID failed: {e}"));
                }
                return Err(e);
            }
        };

        let disp: IDispatch = match crate::co_create_instance_safe(&clsid, None, CLSCTX_ALL) {
            Ok(obj) => obj,
            Err(e) => {
                if log_failures {
                    crate::log_debug(&format!("JAWS CoCreateInstance failed: {e}"));
                }
                return Err(e);
            }
        };

        let mut dispid: i32 = 0;
        let name = to_wide("RunFunction");
        let names = [PCWSTR(name.as_ptr())];
        let getid_res =
            unsafe { disp.GetIDsOfNames(&GUID::default(), names.as_ptr(), 1, 0, &mut dispid) };
        if let Err(e) = getid_res {
            if log_failures {
                crate::log_debug(&format!("JAWS GetIDsOfNames failed: {e}"));
            }
            return Err(e);
        }

        let function_call = format!(
            "BrailleMessage(\"{}\")",
            jaws_escape_function_string_arg(text)
        );
        let mut args = [VARIANT::from(BSTR::from(function_call))];
        let disp_params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: args.len() as u32,
            cNamedArgs: 0,
        };
        let mut excep = EXCEPINFO::default();
        let mut arg_err: u32 = 0;
        let invoke_res = unsafe {
            disp.Invoke(
                dispid,
                &GUID::default(),
                0,
                DISPATCH_METHOD,
                &disp_params,
                None,
                Some(&mut excep),
                Some(&mut arg_err),
            )
        };
        if log_failures && let Err(ref e) = invoke_res {
            crate::log_debug(&format!("JAWS Invoke failed: {e} (arg err: {arg_err})"));
        }
        invoke_res.map(|_| ())
    })();

    if should_uninit {
        unsafe { CoUninitialize() };
    }

    result
}

fn jaws_is_active() -> bool {
    match jaws_invoke_saystring("", false, false) {
        Ok(()) => true,
        Err(e) => {
            if JAWS_PROBE_LOGGED.set(()).is_ok() {
                crate::log_debug(&format!("JAWS probe via COM failed: {e}"));
            }
            let class = to_wide("JFWUI2");
            let title = to_wide("JAWS");
            crate::find_window_w_safe(PCWSTR(class.as_ptr()), PCWSTR(title.as_ptr())).0 != 0
        }
    }
}

fn jaws_speak(text: &str, interrupt: bool) -> bool {
    if !jaws_is_active() {
        return false;
    }

    jaws_invoke_saystring(text, interrupt, true).is_ok()
}

fn jaws_braille(text: &str) -> bool {
    if !jaws_is_active() {
        return false;
    }

    jaws_invoke_braille_message(text, true).is_ok()
}

/// Speaks text through available screen readers (NVDA and JAWS).
/// Returns true if any screen reader accepted the message.
pub fn screen_reader_speak(text: &str) -> bool {
    let mut ok = false;
    if nvda_is_active() {
        ok |= nvda_speak_raw(text);
    }
    ok |= jaws_speak(text, false);
    enqueue_braille(text);
    ok
}

/// Attempts to speak text using the NVDA Controller Client DLL.
/// Also mirrors output to JAWS when available.
pub fn nvda_speak(text: &str) -> bool {
    screen_reader_speak(text)
}

// Le DLL nvdaControllerClient64.dll e SoundTouch64.dll sono ora embedded
// nell'exe e estratte automaticamente da embedded_deps::extract_all()

#[cfg(test)]
mod tests {
    use super::{to_wide, to_wide_normalized};

    #[test]
    fn wide_conversion_removes_embedded_nuls_and_keeps_one_terminator() {
        let wide = to_wide("testo prima\0testo dopo");
        let expected: Vec<u16> = "testo primatesto dopo"
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        assert_eq!(wide, expected);
        assert_eq!(wide.iter().filter(|unit| **unit == 0).count(), 1);
    }

    #[test]
    fn normalized_wide_conversion_preserves_text_after_embedded_nul() {
        let wide = to_wide_normalized("prima\0dopo\nseconda riga");
        let expected: Vec<u16> = "primadopo\r\nseconda riga"
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        assert_eq!(wide, expected);
    }
}
