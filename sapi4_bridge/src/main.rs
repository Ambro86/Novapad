use std::ptr;
use std::env;
use std::io::{self, Read, BufRead};
use std::path::Path;
use std::sync::mpsc;
use std::sync::atomic::{AtomicBool, Ordering};
use winapi::um::combaseapi::{CoInitializeEx, CoCreateInstance, CLSCTX_ALL};
use winapi::shared::guiddef::{GUID, REFIID};
use winapi::um::unknwnbase::IUnknown;
use winapi::um::winuser::{MSG, PeekMessageW, TranslateMessage, DispatchMessageW, PM_REMOVE};
use winapi::shared::winerror::{S_OK, E_NOINTERFACE};
use winapi::shared::minwindef::{DWORD, FILETIME, WORD};
use widestring::U16CString;

// SAPI4 GUIDs
const CLSID_TTSENUMERATOR: GUID = GUID { Data1: 0xD67C0280, Data2: 0xC743, Data3: 0x11CD, Data4: [0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50] };
const IID_ITTSENUM: GUID = GUID { Data1: 0x6B837B20, Data2: 0x4A47, Data3: 0x101B, Data4: [0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F] };
const IID_ITTSFIND: GUID = GUID { Data1: 0x7AA42960, Data2: 0x4A47, Data3: 0x101B, Data4: [0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F] };

// CLSID for MMAudioDest (default audio output)
const CLSID_MMAUDIODEST: GUID = GUID { Data1: 0xCB96B400, Data2: 0xC743, Data3: 0x11CD, Data4: [0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50] };
const IID_IAUDIO: GUID = GUID { Data1: 0xF546B340, Data2: 0xC743, Data3: 0x11CD, Data4: [0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50] };
const IID_IUNKNOWN: GUID = GUID { Data1: 0x00000000, Data2: 0x0000, Data3: 0x0000, Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46] };
const CLSID_AUDIODESTFILE: GUID = GUID { Data1: 0xD4623720, Data2: 0xE4B9, Data3: 0x11CF, Data4: [0x8D, 0x56, 0x00, 0xA0, 0xC9, 0x03, 0x4A, 0x7E] };
const IID_IAUDIOFILE: GUID = GUID { Data1: 0xFD7C2320, Data2: 0x3D6D, Data3: 0x11B9, Data4: [0xC0, 0x00, 0xFE, 0xD6, 0xCB, 0xA3, 0xB1, 0xA9] };
const IID_ITTSNOTIFYSINKW: GUID = GUID { Data1: 0xC0FA8F40, Data2: 0x4A46, Data3: 0x101B, Data4: [0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F] };
const IID_IAUDIOFILENOTIFYSINK: GUID = GUID { Data1: 0x492FE490, Data2: 0x51E7, Data3: 0x11B9, Data4: [0xC0, 0x00, 0xFE, 0xD6, 0xCB, 0xA3, 0xB1, 0xA9] };
const TTSDATAFLAG_TAGGED: DWORD = 1;

// ITTSEnum vtable
#[repr(C)]
struct ITTSEnumVtbl {
    query_interface: unsafe extern "system" fn(*mut ITTSEnum, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut ITTSEnum) -> u32,
    release: unsafe extern "system" fn(*mut ITTSEnum) -> u32,
    next: unsafe extern "system" fn(*mut ITTSEnum, u32, *mut TTSMODEINFO, *mut u32) -> i32,
    skip: unsafe extern "system" fn(*mut ITTSEnum, u32) -> i32,
    reset: unsafe extern "system" fn(*mut ITTSEnum) -> i32,
    clone: unsafe extern "system" fn(*mut ITTSEnum, *mut *mut ITTSEnum) -> i32,
    select: unsafe extern "system" fn(*mut ITTSEnum, GUID, *mut *mut ITTSCentral, *mut IUnknown) -> i32,
}

#[repr(C)]
struct ITTSEnum {
    lpVtbl: *const ITTSEnumVtbl,
}

// ITTSFind vtable - this is different from ITTSEnum!
#[repr(C)]
struct ITTSFindVtbl {
    query_interface: unsafe extern "system" fn(*mut ITTSFind, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut ITTSFind) -> u32,
    release: unsafe extern "system" fn(*mut ITTSFind) -> u32,
    find: unsafe extern "system" fn(*mut ITTSFind, *const TTSMODEINFO, *const TTSMODEINFO, *mut TTSMODEINFO) -> i32,
    select: unsafe extern "system" fn(*mut ITTSFind, GUID, *mut *mut ITTSCentral, *mut IUnknown) -> i32,
}

#[repr(C)]
struct ITTSFind {
    lpVtbl: *const ITTSFindVtbl,
}

// ITTSCentral vtable
#[repr(C)]
struct ITTSCentralVtbl {
    query_interface: unsafe extern "system" fn(*mut ITTSCentral, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut ITTSCentral) -> u32,
    release: unsafe extern "system" fn(*mut ITTSCentral) -> u32,
    inject: unsafe extern "system" fn(*mut ITTSCentral, *const u16) -> i32,
    mode_get: unsafe extern "system" fn(*mut ITTSCentral, *mut TTSMODEINFO) -> i32,
    phoneme: unsafe extern "system" fn(*mut ITTSCentral, u32, DWORD, SDATA, *mut SDATA) -> i32,
    posn_get: unsafe extern "system" fn(*mut ITTSCentral, *mut QWORD) -> i32,
    text_data: unsafe extern "system" fn(*mut ITTSCentral, u32, DWORD, SDATA, *mut std::ffi::c_void, GUID) -> i32,
    to_file_time: unsafe extern "system" fn(*mut ITTSCentral, *mut QWORD, *mut FILETIME) -> i32,
    audio_pause: unsafe extern "system" fn(*mut ITTSCentral) -> i32,
    audio_resume: unsafe extern "system" fn(*mut ITTSCentral) -> i32,
    audio_reset: unsafe extern "system" fn(*mut ITTSCentral) -> i32,
    register: unsafe extern "system" fn(*mut ITTSCentral, *mut std::ffi::c_void, GUID, *mut DWORD) -> i32,
    un_register: unsafe extern "system" fn(*mut ITTSCentral, DWORD) -> i32,
}

#[repr(C)]
struct ITTSCentral {
    lpVtbl: *const ITTSCentralVtbl,
}

// IAudioFile vtable
#[repr(C)]
struct IAudioFileVtbl {
    query_interface: unsafe extern "system" fn(*mut IAudioFile, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut IAudioFile) -> u32,
    release: unsafe extern "system" fn(*mut IAudioFile) -> u32,
    register: unsafe extern "system" fn(*mut IAudioFile, *mut std::ffi::c_void) -> i32,
    set: unsafe extern "system" fn(*mut IAudioFile, *const u16, DWORD) -> i32,
    add: unsafe extern "system" fn(*mut IAudioFile, *const u16, DWORD) -> i32,
    flush: unsafe extern "system" fn(*mut IAudioFile) -> i32,
    real_time_set: unsafe extern "system" fn(*mut IAudioFile, WORD) -> i32,
    real_time_get: unsafe extern "system" fn(*mut IAudioFile, *mut WORD) -> i32,
}

#[repr(C)]
struct IAudioFile {
    lpVtbl: *const IAudioFileVtbl,
}

// ITTSNotifySink vtable
#[repr(C)]
struct ITTSNotifySinkVtbl {
    query_interface: unsafe extern "system" fn(*mut ITTSNotifySink, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut ITTSNotifySink) -> u32,
    release: unsafe extern "system" fn(*mut ITTSNotifySink) -> u32,
    attrib_changed: unsafe extern "system" fn(*mut ITTSNotifySink, DWORD) -> i32,
    audio_start: unsafe extern "system" fn(*mut ITTSNotifySink, QWORD) -> i32,
    audio_stop: unsafe extern "system" fn(*mut ITTSNotifySink, QWORD) -> i32,
    visual: unsafe extern "system" fn(*mut ITTSNotifySink, QWORD, u16, u16, DWORD, *mut std::ffi::c_void) -> i32,
}

#[repr(C)]
struct ITTSNotifySink {
    lpVtbl: *const ITTSNotifySinkVtbl,
    done: *const AtomicBool,
}

// IAudioFileNotifySink vtable
#[repr(C)]
struct IAudioFileNotifySinkVtbl {
    query_interface: unsafe extern "system" fn(*mut IAudioFileNotifySink, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut IAudioFileNotifySink) -> u32,
    release: unsafe extern "system" fn(*mut IAudioFileNotifySink) -> u32,
    file_begin: unsafe extern "system" fn(*mut IAudioFileNotifySink, DWORD) -> i32,
    file_end: unsafe extern "system" fn(*mut IAudioFileNotifySink, DWORD) -> i32,
    queue_empty: unsafe extern "system" fn(*mut IAudioFileNotifySink) -> i32,
    posn: unsafe extern "system" fn(*mut IAudioFileNotifySink, QWORD, QWORD) -> i32,
}

#[repr(C)]
struct IAudioFileNotifySink {
    lpVtbl: *const IAudioFileNotifySinkVtbl,
    done: *const AtomicBool,
}

type QWORD = u64;

#[repr(C)]
#[derive(Clone, Copy)]
struct SDATA {
    data: *mut u8,
    size: DWORD,
}

const TTSI_NAMELEN: usize = 262;
const TTSI_STYLELEN: usize = 262;
const LANG_LEN: usize = 64;

#[repr(C)]
#[derive(Clone, Copy)]
struct LANGUAGEW {
    language_id: WORD,
    dialect: [u16; LANG_LEN],
}

// TTSMODEINFOW - matches NVDA SAPI4 layout
#[repr(C)]
#[derive(Clone, Copy)]
struct TTSMODEINFO {
    engine_id: GUID,
    manufacturer: [u16; TTSI_NAMELEN],
    product_name: [u16; TTSI_NAMELEN],
    mode_id: GUID,
    mode_name: [u16; TTSI_NAMELEN],
    language: LANGUAGEW,
    speaker: [u16; TTSI_NAMELEN],
    style: [u16; TTSI_STYLELEN],
    gender: WORD,
    age: WORD,
    features: DWORD,
    interfaces: DWORD,
    engine_features: DWORD,
}

impl TTSMODEINFO {
    fn mode_id(&self) -> GUID {
        self.mode_id
    }

    fn mode_name(&self) -> String {
        String::from_utf16_lossy(&self.mode_name)
            .trim_end_matches('\0')
            .to_string()
    }
}

enum ServerCommand {
    Speak(String),
    Pause,
    Resume,
    Stop,
    Quit,
}

fn guid_eq(a: REFIID, b: &GUID) -> bool {
    unsafe {
        let a = &*a;
        a.Data1 == b.Data1
            && a.Data2 == b.Data2
            && a.Data3 == b.Data3
            && a.Data4 == b.Data4
    }
}

unsafe extern "system" fn notify_query_interface(
    this: *mut ITTSNotifySink,
    riid: REFIID,
    ppv: *mut *mut std::ffi::c_void,
) -> i32 {
    if ppv.is_null() {
        return E_NOINTERFACE;
    }
    *ppv = ptr::null_mut();
    if guid_eq(riid, &IID_IUNKNOWN) || guid_eq(riid, &IID_ITTSNOTIFYSINKW) {
        *ppv = this as *mut std::ffi::c_void;
        return S_OK;
    }
    E_NOINTERFACE
}

unsafe extern "system" fn notify_add_ref(_this: *mut ITTSNotifySink) -> u32 {
    1
}

unsafe extern "system" fn notify_release(_this: *mut ITTSNotifySink) -> u32 {
    1
}

unsafe extern "system" fn notify_attrib_changed(_this: *mut ITTSNotifySink, _attr: DWORD) -> i32 {
    S_OK
}

unsafe extern "system" fn notify_audio_start(_this: *mut ITTSNotifySink, _ts: QWORD) -> i32 {
    S_OK
}

unsafe extern "system" fn notify_audio_stop(this: *mut ITTSNotifySink, _ts: QWORD) -> i32 {
    let sink = &*(this as *mut ITTSNotifySink);
    if !sink.done.is_null() {
        (*sink.done).store(true, Ordering::Release);
    }
    S_OK
}

unsafe extern "system" fn notify_visual(
    _this: *mut ITTSNotifySink,
    _ts: QWORD,
    _ipa: u16,
    _engine: u16,
    _hints: DWORD,
    _mouth: *mut std::ffi::c_void,
) -> i32 {
    S_OK
}

unsafe extern "system" fn audio_file_notify_query_interface(
    this: *mut IAudioFileNotifySink,
    riid: REFIID,
    ppv: *mut *mut std::ffi::c_void,
) -> i32 {
    if ppv.is_null() {
        return E_NOINTERFACE;
    }
    *ppv = ptr::null_mut();
    if guid_eq(riid, &IID_IUNKNOWN) || guid_eq(riid, &IID_IAUDIOFILENOTIFYSINK) {
        *ppv = this as *mut std::ffi::c_void;
        return S_OK;
    }
    E_NOINTERFACE
}

unsafe extern "system" fn audio_file_notify_add_ref(_this: *mut IAudioFileNotifySink) -> u32 {
    1
}

unsafe extern "system" fn audio_file_notify_release(_this: *mut IAudioFileNotifySink) -> u32 {
    1
}

unsafe extern "system" fn audio_file_notify_file_begin(_this: *mut IAudioFileNotifySink, _id: DWORD) -> i32 {
    S_OK
}

unsafe extern "system" fn audio_file_notify_file_end(this: *mut IAudioFileNotifySink, _id: DWORD) -> i32 {
    let sink = &*(this as *mut IAudioFileNotifySink);
    if !sink.done.is_null() {
        (*sink.done).store(true, Ordering::Release);
    }
    S_OK
}

unsafe extern "system" fn audio_file_notify_queue_empty(this: *mut IAudioFileNotifySink) -> i32 {
    let sink = &*(this as *mut IAudioFileNotifySink);
    if !sink.done.is_null() {
        (*sink.done).store(true, Ordering::Release);
    }
    S_OK
}

unsafe extern "system" fn audio_file_notify_posn(
    _this: *mut IAudioFileNotifySink,
    _posn: QWORD,
    _time: QWORD,
) -> i32 {
    S_OK
}

static NOTIFY_VTBL: ITTSNotifySinkVtbl = ITTSNotifySinkVtbl {
    query_interface: notify_query_interface,
    add_ref: notify_add_ref,
    release: notify_release,
    attrib_changed: notify_attrib_changed,
    audio_start: notify_audio_start,
    audio_stop: notify_audio_stop,
    visual: notify_visual,
};

static AUDIO_FILE_NOTIFY_VTBL: IAudioFileNotifySinkVtbl = IAudioFileNotifySinkVtbl {
    query_interface: audio_file_notify_query_interface,
    add_ref: audio_file_notify_add_ref,
    release: audio_file_notify_release,
    file_begin: audio_file_notify_file_begin,
    file_end: audio_file_notify_file_end,
    queue_empty: audio_file_notify_queue_empty,
    posn: audio_file_notify_posn,
};

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.iter().any(|a| a == "--list") {
        list_voices();
        return;
    }

    let target_idx = args.iter().position(|a| a == "--voice")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse::<u32>().ok())
        .unwrap_or(1);

    let output_path = args.iter().position(|a| a == "--output")
        .and_then(|i| args.get(i + 1))
        .map(|s| s.to_string());

    if let Some(path) = output_path {
        speak_to_file(target_idx, &path);
        return;
    }

    if args.iter().any(|a| a == "--server") {
        run_server(target_idx);
        return;
    }

    speak_with_voice(target_idx);
}

fn get_mode_name(info: &TTSMODEINFO) -> String {
    info.mode_name()
}

fn list_voices() {
    unsafe {
        CoInitializeEx(ptr::null_mut(), 0x2);

        let mut enum_ptr: *mut ITTSEnum = ptr::null_mut();
        let hr = CoCreateInstance(
            &CLSID_TTSENUMERATOR,
            ptr::null_mut(),
            CLSCTX_ALL,
            &IID_ITTSENUM,
            &mut enum_ptr as *mut _ as *mut _
        );

        if hr != S_OK || enum_ptr.is_null() {
            eprintln!("Failed to create TTSEnumerator, hr={:#x}", hr);
            return;
        }

        let tts_enum = &*enum_ptr;
        let vtbl = &*tts_enum.lpVtbl;

        (vtbl.reset)(enum_ptr);

        eprintln!("sizeof TTSMODEINFO = {}", std::mem::size_of::<TTSMODEINFO>());

        let mut idx = 1u32;
        loop {
            let mut mode_info: TTSMODEINFO = std::mem::zeroed();
            let mut fetched: u32 = 0;

            let hr = (vtbl.next)(enum_ptr, 1, &mut mode_info, &mut fetched);
            if hr != S_OK || fetched == 0 {
                break;
            }

            let name = get_mode_name(&mode_info);
            println!("VOICE:{}|{}", idx, name);
            idx += 1;
        }

        (vtbl.release)(enum_ptr);
    }
}

fn spawn_command_reader() -> mpsc::Receiver<ServerCommand> {
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let stdin = io::stdin();
        let mut reader = io::BufReader::new(stdin.lock());
        loop {
            let mut line = String::new();
            match reader.read_line(&mut line) {
                Ok(0) => {
                    let _ = tx.send(ServerCommand::Quit);
                    break;
                }
                Ok(_) => {
                    let cmd = line.trim_end_matches(&['\r', '\n'][..]);
                    if cmd.starts_with("SPEAK ") {
                        let len_str = cmd.trim_start_matches("SPEAK ").trim();
                        if let Ok(len) = len_str.parse::<usize>() {
                            let mut buf = vec![0u8; len];
                            if reader.read_exact(&mut buf).is_ok() {
                                let text = String::from_utf8_lossy(&buf).to_string();
                                let _ = tx.send(ServerCommand::Speak(text));
                                continue;
                            }
                        }
                        let _ = tx.send(ServerCommand::Quit);
                        break;
                    }
                    match cmd {
                        "PAUSE" => { let _ = tx.send(ServerCommand::Pause); }
                        "RESUME" => { let _ = tx.send(ServerCommand::Resume); }
                        "STOP" => { let _ = tx.send(ServerCommand::Stop); }
                        "QUIT" => {
                            let _ = tx.send(ServerCommand::Quit);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(_) => {
                    let _ = tx.send(ServerCommand::Quit);
                    break;
                }
            }
        }
    });
    rx
}

unsafe fn init_central(target_idx: u32) -> Option<(*mut ITTSEnum, *mut ITTSCentral)> {
    init_central_with_audio(target_idx, ptr::null_mut())
}

unsafe fn init_central_with_audio(target_idx: u32, mut audio_ptr: *mut IUnknown) -> Option<(*mut ITTSEnum, *mut ITTSCentral)> {
    CoInitializeEx(ptr::null_mut(), 0x2);

    let mut enum_ptr: *mut ITTSEnum = ptr::null_mut();
    let hr = CoCreateInstance(
        &CLSID_TTSENUMERATOR,
        ptr::null_mut(),
        CLSCTX_ALL,
        &IID_ITTSENUM,
        &mut enum_ptr as *mut _ as *mut _
    );

    if hr != S_OK || enum_ptr.is_null() {
        eprintln!("Failed to create TTSEnumerator, hr={:#x}", hr);
        return None;
    }

    let tts_enum = &*enum_ptr;
    let vtbl = &*tts_enum.lpVtbl;

    (vtbl.reset)(enum_ptr);

    let mut target_mode: TTSMODEINFO = std::mem::zeroed();
    let mut found = false;
    let mut idx = 1u32;

    loop {
        let mut mode_info: TTSMODEINFO = std::mem::zeroed();
        let mut fetched: u32 = 0;

        let hr = (vtbl.next)(enum_ptr, 1, &mut mode_info, &mut fetched);
        if hr != S_OK || fetched == 0 {
            break;
        }

        if idx == target_idx {
            target_mode = mode_info;
            found = true;
            let name = get_mode_name(&mode_info);
            eprintln!("Found voice {}: {}", idx, name);
            break;
        }
        idx += 1;
    }

    if !found {
        eprintln!("Voice {} not found", target_idx);
        (vtbl.release)(enum_ptr);
        return None;
    }

    if audio_ptr.is_null() {
        let hr = CoCreateInstance(
            &CLSID_MMAUDIODEST,
            ptr::null_mut(),
            CLSCTX_ALL,
            &IID_IAUDIO,
            &mut audio_ptr as *mut _ as *mut _
        );

        if hr != S_OK || audio_ptr.is_null() {
            eprintln!("Failed to create MMAudioDest, hr={:#x}. Trying without it.", hr);
            audio_ptr = ptr::null_mut();
        }
    }

    let mut central_ptr: *mut ITTSCentral = ptr::null_mut();
    let mode_guid = target_mode.mode_id();
    eprintln!(
        "ModeID GUID: {:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        mode_guid.Data1, mode_guid.Data2, mode_guid.Data3,
        mode_guid.Data4[0], mode_guid.Data4[1], mode_guid.Data4[2], mode_guid.Data4[3],
        mode_guid.Data4[4], mode_guid.Data4[5], mode_guid.Data4[6], mode_guid.Data4[7]
    );

    let hr = (vtbl.select)(enum_ptr, mode_guid, &mut central_ptr, audio_ptr);
    if hr != S_OK || central_ptr.is_null() {
        eprintln!("ITTSEnum::Select failed, hr={:#x}", hr);
        (vtbl.release)(enum_ptr);
        return None;
    }

    eprintln!("ITTSEnum::Select succeeded!");
    Some((enum_ptr, central_ptr))
}

unsafe fn speak_text_with_flags(central_ptr: *mut ITTSCentral, text: &str, flags: DWORD) -> i32 {
    let tts_central = &*central_ptr;
    let central_vtbl = &*tts_central.lpVtbl;

    let text_u16 = U16CString::from_str(text).unwrap();
    let text_bytes = text_u16.as_slice_with_nul();

    let sdata = SDATA {
        data: text_bytes.as_ptr() as *mut u8,
        size: (text_bytes.len() * 2) as DWORD,
    };

    (central_vtbl.text_data)(
        central_ptr,
        0,
        flags,
        sdata,
        ptr::null_mut(),
        std::mem::zeroed()
    )
}

unsafe fn speak_text(central_ptr: *mut ITTSCentral, text: &str) -> i32 {
    speak_text_with_flags(central_ptr, text, 0)
}

fn run_server(target_idx: u32) {
    unsafe {
        let Some((enum_ptr, central_ptr)) = init_central(target_idx) else {
            return;
        };
        let tts_enum = &*enum_ptr;
        let vtbl = &*tts_enum.lpVtbl;
        let tts_central = &*central_ptr;
        let central_vtbl = &*tts_central.lpVtbl;

        let rx = spawn_command_reader();
        let mut running = true;

        while running {
            let mut msg: MSG = std::mem::zeroed();
            while PeekMessageW(&mut msg, ptr::null_mut(), 0, 0, PM_REMOVE) != 0 {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }

            match rx.try_recv() {
                Ok(ServerCommand::Speak(text)) => {
                    let hr = speak_text(central_ptr, &text);
                    if hr != S_OK {
                        eprintln!("TextData failed, hr={:#x}", hr);
                    }
                }
                Ok(ServerCommand::Pause) => { let _ = (central_vtbl.audio_pause)(central_ptr); }
                Ok(ServerCommand::Resume) => { let _ = (central_vtbl.audio_resume)(central_ptr); }
                Ok(ServerCommand::Stop) => {
                    let _ = (central_vtbl.audio_reset)(central_ptr);
                    running = false;
                }
                Ok(ServerCommand::Quit) => { running = false; }
                Err(mpsc::TryRecvError::Disconnected) => { running = false; }
                Err(mpsc::TryRecvError::Empty) => {}
            }

            std::thread::sleep(std::time::Duration::from_millis(10));
        }

        (central_vtbl.release)(central_ptr);
        (vtbl.release)(enum_ptr);
    }
}

fn speak_to_file(target_idx: u32, output_path: &str) {
    unsafe {
        CoInitializeEx(ptr::null_mut(), 0x2);

        let mut audio_file_ptr: *mut IAudioFile = ptr::null_mut();
        let hr = CoCreateInstance(
            &CLSID_AUDIODESTFILE,
            ptr::null_mut(),
            CLSCTX_ALL,
            &IID_IAUDIOFILE,
            &mut audio_file_ptr as *mut _ as *mut _
        );
        if hr != S_OK || audio_file_ptr.is_null() {
            eprintln!("Failed to create AudioDestFile, hr={:#x}", hr);
            return;
        }

        let audio_vtbl = &*(*audio_file_ptr).lpVtbl;
        let out_wide = match U16CString::from_str(output_path) {
            Ok(v) => v,
            Err(_) => {
                eprintln!("Invalid output path");
                (audio_vtbl.release)(audio_file_ptr);
                return;
            }
        };
        let hr = (audio_vtbl.set)(audio_file_ptr, out_wide.as_ptr(), 1);
        if hr != S_OK {
            eprintln!("AudioFile::Set failed, hr={:#x}", hr);
            (audio_vtbl.release)(audio_file_ptr);
            return;
        }

        let _ = (audio_vtbl.real_time_set)(audio_file_ptr, 0x100);

        let audio_unknown = audio_file_ptr as *mut IUnknown;
        let Some((enum_ptr, central_ptr)) = init_central_with_audio(target_idx, audio_unknown) else {
            (audio_vtbl.release)(audio_file_ptr);
            return;
        };

        let tts_enum = &*enum_ptr;
        let enum_vtbl = &*tts_enum.lpVtbl;
        let tts_central = &*central_ptr;
        let central_vtbl = &*tts_central.lpVtbl;

        let done = AtomicBool::new(false);
        let notify = ITTSNotifySink {
            lpVtbl: &NOTIFY_VTBL,
            done: &done,
        };
        let audio_notify = IAudioFileNotifySink {
            lpVtbl: &AUDIO_FILE_NOTIFY_VTBL,
            done: &done,
        };
        let _ = (audio_vtbl.register)(
            audio_file_ptr,
            &audio_notify as *const _ as *mut std::ffi::c_void,
        );
        let mut reg_key: DWORD = 0;
        let _ = (central_vtbl.register)(
            central_ptr,
            &notify as *const _ as *mut std::ffi::c_void,
            IID_ITTSNOTIFYSINKW,
            &mut reg_key,
        );

        let mut text = String::new();
        let _ = io::stdin().read_to_string(&mut text);
        if text.is_empty() {
            (central_vtbl.release)(central_ptr);
            (enum_vtbl.release)(enum_ptr);
            (audio_vtbl.release)(audio_file_ptr);
            return;
        }

        let hr = speak_text_with_flags(central_ptr, &text, TTSDATAFLAG_TAGGED);
        if hr != S_OK {
            eprintln!("TextData failed, hr={:#x}", hr);
        } else {
            let output_path = Path::new(output_path);
            let mut last_size = std::fs::metadata(output_path).map(|m| m.len()).unwrap_or(0);
            let mut last_change = std::time::Instant::now();
            let mut msg: MSG = std::mem::zeroed();
            while !done.load(Ordering::Acquire) {
                while PeekMessageW(&mut msg, ptr::null_mut(), 0, 0, PM_REMOVE) != 0 {
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
                let size = std::fs::metadata(output_path).map(|m| m.len()).unwrap_or(0);
                if size != last_size {
                    last_size = size;
                    last_change = std::time::Instant::now();
                } else if size > 44 && last_change.elapsed().as_secs() >= 1 {
                    break;
                }
                std::thread::sleep(std::time::Duration::from_millis(10));
            }
        }

        let _ = (audio_vtbl.flush)(audio_file_ptr);
        (central_vtbl.release)(central_ptr);
        (enum_vtbl.release)(enum_ptr);
        (audio_vtbl.release)(audio_file_ptr);
    }
}

fn speak_with_voice(target_idx: u32) {
    unsafe {
        let Some((enum_ptr, central_ptr)) = init_central(target_idx) else {
            return;
        };
        let tts_enum = &*enum_ptr;
        let vtbl = &*tts_enum.lpVtbl;
        let tts_central = &*central_ptr;
        let central_vtbl = &*tts_central.lpVtbl;

        // Read text from stdin
        let mut text = String::new();
        let _ = io::stdin().read_to_string(&mut text);
        if text.is_empty() {
            text = "Test di sintesi vocale. Questa è la voce italiana.".to_string();
        }

        let hr = speak_text(central_ptr, &text);

        if hr != S_OK {
            eprintln!("TextData failed, hr={:#x}", hr);
        } else {
            eprintln!("TextData succeeded!");
        }

        // Message loop
        let mut msg: MSG = std::mem::zeroed();
        let start = std::time::Instant::now();
        while start.elapsed().as_secs() < 30 {
            while PeekMessageW(&mut msg, ptr::null_mut(), 0, 0, PM_REMOVE) != 0 {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            std::thread::sleep(std::time::Duration::from_millis(10));
        }

        (central_vtbl.release)(central_ptr);
        (vtbl.release)(enum_ptr);
    }
}
