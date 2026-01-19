#![allow(non_snake_case)]
#![allow(clippy::upper_case_acronyms)]

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
use windows::core::PCWSTR;
use windows::Win32::Media::MediaFoundation::{
    IMFMediaType, IMFSinkWriter, IMFSourceReader, MFCreateMediaType,
    MFCreateSinkWriterFromURL, MFCreateSourceReaderFromURL, MFShutdown, MFStartup,
    MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT,
    MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE,
    MF_SOURCE_READERF_ENDOFSTREAM, MF_SOURCE_READER_FIRST_AUDIO_STREAM, MF_VERSION,
    MFAudioFormat_MP3, MFAudioFormat_PCM, MFMediaType_Audio,
};

// SAPI4 GUIDs
const CLSID_TTSENUMERATOR: GUID = GUID { Data1: 0xD67C0280, Data2: 0xC743, Data3: 0x11CD, Data4: [0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50] };
const IID_ITTSENUM: GUID = GUID { Data1: 0x6B837B20, Data2: 0x4A47, Data3: 0x101B, Data4: [0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F] };

// CLSID for MMAudioDest (default audio output)
const CLSID_MMAUDIODEST: GUID = GUID { Data1: 0xCB96B400, Data2: 0xC743, Data3: 0x11CD, Data4: [0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50] };
const IID_IAUDIO: GUID = GUID { Data1: 0xF546B340, Data2: 0xC743, Data3: 0x11CD, Data4: [0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50] };
const IID_IUNKNOWN: GUID = GUID { Data1: 0x00000000, Data2: 0x0000, Data3: 0x0000, Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46] };
const CLSID_AUDIODESTFILE: GUID = GUID { Data1: 0xD4623720, Data2: 0xE4B9, Data3: 0x11CF, Data4: [0x8D, 0x56, 0x00, 0xA0, 0xC9, 0x03, 0x4A, 0x7E] };
const IID_IAUDIOFILE: GUID = GUID { Data1: 0xFD7C2320, Data2: 0x3D6D, Data3: 0x11B9, Data4: [0xC0, 0x00, 0xFE, 0xD6, 0xCB, 0xA3, 0xB1, 0xA9] };
const IID_ITTSNOTIFYSINKW: GUID = GUID { Data1: 0xC0FA8F40, Data2: 0x4A46, Data3: 0x101B, Data4: [0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F] };
const IID_IAUDIOFILENOTIFYSINK: GUID = GUID { Data1: 0x492FE490, Data2: 0x51E7, Data3: 0x11B9, Data4: [0xC0, 0x00, 0xFE, 0xD6, 0xCB, 0xA3, 0xB1, 0xA9] };
const TTSDATAFLAG_TAGGED: DWORD = 1;
const IID_ITTSATTRIBUTESW: GUID = GUID { Data1: 0x1287A280, Data2: 0x4A47, Data3: 0x101B, Data4: [0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F] };
const TTSATTR_MINSPEED: DWORD = 0;
const TTSATTR_MAXSPEED: DWORD = 0xFFFF_FFFF;
const TTSATTR_MINPITCH: WORD = 0;
const TTSATTR_MAXPITCH: WORD = 0xFFFF;
const TTSATTR_MINVOLUME: DWORD = 0;
const TTSATTR_MAXVOLUME: DWORD = 0xFFFF_FFFF;

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

// ITTSAttributes vtable
#[repr(C)]
struct ITTSAttributesVtbl {
    query_interface: unsafe extern "system" fn(*mut ITTSAttributes, REFIID, *mut *mut std::ffi::c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut ITTSAttributes) -> u32,
    release: unsafe extern "system" fn(*mut ITTSAttributes) -> u32,
    pitch_get: unsafe extern "system" fn(*mut ITTSAttributes, *mut WORD) -> i32,
    pitch_set: unsafe extern "system" fn(*mut ITTSAttributes, WORD) -> i32,
    real_time_get: unsafe extern "system" fn(*mut ITTSAttributes, *mut DWORD) -> i32,
    real_time_set: unsafe extern "system" fn(*mut ITTSAttributes, DWORD) -> i32,
    speed_get: unsafe extern "system" fn(*mut ITTSAttributes, *mut DWORD) -> i32,
    speed_set: unsafe extern "system" fn(*mut ITTSAttributes, DWORD) -> i32,
    volume_get: unsafe extern "system" fn(*mut ITTSAttributes, *mut DWORD) -> i32,
    volume_set: unsafe extern "system" fn(*mut ITTSAttributes, DWORD) -> i32,
}

#[repr(C)]
struct ITTSAttributes {
    lpVtbl: *const ITTSAttributesVtbl,
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
    let sink = &*this;
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
    let sink = &*this;
    if !sink.done.is_null() {
        (*sink.done).store(true, Ordering::Release);
    }
    S_OK
}

unsafe extern "system" fn audio_file_notify_queue_empty(this: *mut IAudioFileNotifySink) -> i32 {
    let sink = &*this;
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

struct MfGuard;

impl MfGuard {
    fn start() -> Result<Self, String> {
        unsafe {
            if let Err(e) = MFStartup(MF_VERSION, 0) {
                return Err(format!("MFStartup failed: {}", e));
            }
        }
        Ok(MfGuard)
    }
}

impl Drop for MfGuard {
    fn drop(&mut self) {
        unsafe {
            if let Err(e) = MFShutdown() {
                eprintln!("MFShutdown failed: {}", e);
            }
        }
    }
}

fn encode_wav_to_mp3(wav_path: &Path, mp3_path: &Path) -> Result<(), String> {
    unsafe {
        let _guard = MfGuard::start()?;

        let wav_str = wav_path.to_str().ok_or("Invalid wav path")?;
        let mp3_str = mp3_path.to_str().ok_or("Invalid mp3 path")?;
        let wav_wide = U16CString::from_str(wav_str).map_err(|_| "Invalid wav path")?;
        let mp3_wide = U16CString::from_str(mp3_str).map_err(|_| "Invalid mp3 path")?;

        let reader: IMFSourceReader =
            MFCreateSourceReaderFromURL(PCWSTR(wav_wide.as_ptr()), None)
                .map_err(|e| format!("MFCreateSourceReaderFromURL failed: {}", e))?;
        let pcm_type: IMFMediaType =
            MFCreateMediaType().map_err(|e| format!("MFCreateMediaType failed: {}", e))?;
        pcm_type
            .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
            .map_err(|e| format!("SetGUID major type failed: {}", e))?;
        pcm_type
            .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)
            .map_err(|e| format!("SetGUID subtype PCM failed: {}", e))?;
        let sample_rate = 44100u32;
        let channels = 2u32;
        let bits_per_sample = 16u32;
        let block_align = channels * (bits_per_sample / 8);
        let avg_bytes = sample_rate * block_align;
        pcm_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, sample_rate)
            .map_err(|e| format!("Set sample rate failed: {}", e))?;
        pcm_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, channels)
            .map_err(|e| format!("Set channels failed: {}", e))?;
        pcm_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, bits_per_sample)
            .map_err(|e| format!("Set bits per sample failed: {}", e))?;
        pcm_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, block_align)
            .map_err(|e| format!("Set block alignment failed: {}", e))?;
        pcm_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, avg_bytes)
            .map_err(|e| format!("Set avg bytes failed: {}", e))?;
        reader
            .SetCurrentMediaType(
                MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32,
                None,
                &pcm_type,
            )
            .map_err(|e| format!("SetCurrentMediaType failed: {}", e))?;

        let in_type = reader
            .GetCurrentMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32)
            .map_err(|e| format!("GetCurrentMediaType failed: {}", e))?;
        let out_type: IMFMediaType =
            MFCreateMediaType().map_err(|e| format!("MFCreateMediaType failed: {}", e))?;
        out_type
            .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
            .map_err(|e| format!("SetGUID major type failed: {}", e))?;
        out_type
            .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_MP3)
            .map_err(|e| format!("SetGUID subtype MP3 failed: {}", e))?;
        out_type
            .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, channels)
            .map_err(|e| format!("Set channels failed: {}", e))?;
        out_type
            .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, sample_rate)
            .map_err(|e| format!("Set sample rate failed: {}", e))?;
        let mp3_avg_bytes = 128_000u32 / 8;
        out_type
            .SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, mp3_avg_bytes)
            .map_err(|e| format!("Set mp3 bitrate failed: {}", e))?;

        let writer: IMFSinkWriter =
            MFCreateSinkWriterFromURL(PCWSTR(mp3_wide.as_ptr()), None, None)
                .map_err(|e| format!("MFCreateSinkWriterFromURL failed: {}", e))?;
        let stream_index = writer
            .AddStream(&out_type)
            .map_err(|e| format!("SinkWriter AddStream failed: {}", e))?;
        if let Err(e) = writer.SetInputMediaType(stream_index, &in_type, None) {
            return Err(format!("SinkWriter SetInputMediaType failed: {}", e));
        }
        writer
            .BeginWriting()
            .map_err(|e| format!("SinkWriter BeginWriting failed: {}", e))?;

        loop {
            let mut read_stream = 0u32;
            let mut flags = 0u32;
            let mut _timestamp = 0i64;
            let mut sample = None;
            reader
                .ReadSample(
                    MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32,
                    0,
                    Some(&mut read_stream),
                    Some(&mut flags),
                    Some(&mut _timestamp),
                    Some(&mut sample),
                )
                .map_err(|e| format!("ReadSample failed: {}", e))?;

            if flags & (MF_SOURCE_READERF_ENDOFSTREAM.0 as u32) != 0 {
                break;
            }
            if let Some(sample) = sample {
                writer
                    .WriteSample(stream_index, &sample)
                    .map_err(|e| format!("WriteSample failed: {}", e))?;
            }
        }

        writer
            .Finalize()
            .map_err(|e| format!("SinkWriter Finalize failed: {}", e))?;
        Ok(())
    }
}

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

    let rate = args.iter().position(|a| a == "--rate")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse::<i32>().ok());
    let pitch = args.iter().position(|a| a == "--pitch")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse::<i32>().ok());
    let volume = args.iter().position(|a| a == "--volume")
        .and_then(|i| args.get(i + 1))
        .and_then(|s| s.parse::<i32>().ok());

    let output_path = args.iter().position(|a| a == "--output")
        .and_then(|i| args.get(i + 1))
        .map(|s| s.to_string());

    if let Some(path) = output_path {
        speak_to_file(target_idx, &path, rate, pitch, volume);
        return;
    }

    if args.iter().any(|a| a == "--server") {
        run_server(target_idx, rate, pitch, volume);
        return;
    }

    speak_with_voice(target_idx, rate, pitch, volume);
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
                    if let Err(e) = tx.send(ServerCommand::Quit) { eprintln!("Send Quit failed: {}", e); }
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
                                if let Err(e) = tx.send(ServerCommand::Speak(text)) { eprintln!("Send Speak failed: {}", e); }
                                continue;
                            }
                        }
                        if let Err(e) = tx.send(ServerCommand::Quit) { eprintln!("Send Quit failed: {}", e); }
                        break;
                    }
                    match cmd {
                        "PAUSE" => { if let Err(e) = tx.send(ServerCommand::Pause) { eprintln!("Send Pause failed: {}", e); } }
                        "RESUME" => { if let Err(e) = tx.send(ServerCommand::Resume) { eprintln!("Send Resume failed: {}", e); } }
                        "STOP" => { if let Err(e) = tx.send(ServerCommand::Stop) { eprintln!("Send Stop failed: {}", e); } }
                        "QUIT" => {
                            if let Err(e) = tx.send(ServerCommand::Quit) { eprintln!("Send Quit failed: {}", e); }
                            break;
                        }
                        _ => {}
                    }
                }
                Err(_) => {
                    if let Err(e) = tx.send(ServerCommand::Quit) { eprintln!("Send Quit failed: {}", e); }
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

fn run_server(target_idx: u32, rate: Option<i32>, pitch: Option<i32>, volume: Option<i32>) {
    unsafe {
        let Some((enum_ptr, central_ptr)) = init_central(target_idx) else {
            return;
        };
        let tts_enum = &*enum_ptr;
        let vtbl = &*tts_enum.lpVtbl;
        let tts_central = &*central_ptr;
        let central_vtbl = &*tts_central.lpVtbl;

        apply_tts_attributes(central_ptr, rate, pitch, volume);
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
                Ok(ServerCommand::Pause) => {
                    let hr = (central_vtbl.audio_pause)(central_ptr);
                    if hr != S_OK { eprintln!("AudioPause failed: {:#x}", hr); }
                }
                Ok(ServerCommand::Resume) => {
                    let hr = (central_vtbl.audio_resume)(central_ptr);
                    if hr != S_OK { eprintln!("AudioResume failed: {:#x}", hr); }
                }
                Ok(ServerCommand::Stop) => {
                    let hr = (central_vtbl.audio_reset)(central_ptr);
                    if hr != S_OK { eprintln!("AudioReset failed: {:#x}", hr); }
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

fn speak_to_file(target_idx: u32, output_path: &str, rate: Option<i32>, pitch: Option<i32>, volume: Option<i32>) {
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

        let output_path = Path::new(output_path);
        let is_mp3 = output_path
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.eq_ignore_ascii_case("mp3"))
            .unwrap_or(false);
        let wav_path = if is_mp3 {
            output_path.with_extension("wav")
        } else {
            output_path.to_path_buf()
        };

        let audio_vtbl = &*(*audio_file_ptr).lpVtbl;
        let wav_str = match wav_path.to_str() {
            Some(s) => s,
            None => {
                eprintln!("Invalid output path");
                (audio_vtbl.release)(audio_file_ptr);
                return;
            }
        };
        let out_wide = match U16CString::from_str(wav_str) {
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

        let hr = (audio_vtbl.real_time_set)(audio_file_ptr, 0x100);
        if hr != S_OK { eprintln!("AudioFile RealTimeSet failed: {:#x}", hr); }

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
        let hr = (audio_vtbl.register)(
            audio_file_ptr,
            &audio_notify as *const _ as *mut std::ffi::c_void,
        );
        if hr != S_OK { eprintln!("AudioFile Register failed: {:#x}", hr); }
        let mut reg_key: DWORD = 0;
        let hr = (central_vtbl.register)(
            central_ptr,
            &notify as *const _ as *mut std::ffi::c_void,
            IID_ITTSNOTIFYSINKW,
            &mut reg_key,
        );
        if hr != S_OK { eprintln!("ITTSCentral Register failed: {:#x}", hr); }

        apply_tts_attributes(central_ptr, rate, pitch, volume);

        let mut text = String::new();
        if io::stdin().read_to_string(&mut text).is_err() {
            eprintln!("Failed to read text from stdin");
        }
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
            let output_path = wav_path.as_path();
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

        let hr = (audio_vtbl.flush)(audio_file_ptr);
        if hr != S_OK { eprintln!("AudioFile Flush failed: {:#x}", hr); }
        (central_vtbl.release)(central_ptr);
        (enum_vtbl.release)(enum_ptr);
        (audio_vtbl.release)(audio_file_ptr);

        if is_mp3 {
            if let Err(err) = encode_wav_to_mp3(&wav_path, output_path) {
                eprintln!("WAV->MP3 failed: {}", err);
            } else if let Err(e) = std::fs::remove_file(&wav_path) {
                eprintln!("Failed to remove temp wav: {}", e);
            }
        }
    }
}

fn speak_with_voice(target_idx: u32, rate: Option<i32>, pitch: Option<i32>, volume: Option<i32>) {
    unsafe {
        let Some((enum_ptr, central_ptr)) = init_central(target_idx) else {
            return;
        };
        let tts_enum = &*enum_ptr;
        let vtbl = &*tts_enum.lpVtbl;
        let tts_central = &*central_ptr;
        let central_vtbl = &*tts_central.lpVtbl;

        apply_tts_attributes(central_ptr, rate, pitch, volume);

        // Read text from stdin
        let mut text = String::new();
        if io::stdin().read_to_string(&mut text).is_err() {
            eprintln!("Failed to read text from stdin");
        }
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

unsafe fn apply_tts_attributes(
    central_ptr: *mut ITTSCentral,
    rate: Option<i32>,
    pitch: Option<i32>,
    volume: Option<i32>,
) {
    if rate.is_none() && pitch.is_none() && volume.is_none() {
        return;
    }
    let tts_central = &*central_ptr;
    let central_vtbl = &*tts_central.lpVtbl;
    let mut attr_ptr: *mut ITTSAttributes = ptr::null_mut();
    let hr = (central_vtbl.query_interface)(
        central_ptr,
        &IID_ITTSATTRIBUTESW,
        &mut attr_ptr as *mut _ as *mut _,
    );
    if hr != S_OK || attr_ptr.is_null() {
        return;
    }
    let attr = &*attr_ptr;
    let attr_vtbl = &*attr.lpVtbl;

    let mut speed_percent: Option<f64> = None;
    let mut pitch_percent: Option<f64> = None;
    let mut volume_percent: Option<f64> = None;

    if let Some(rate) = rate {
        let rate = rate.clamp(-100, 100);
        speed_percent = Some(((rate as f64 + 100.0) / 2.0).clamp(0.0, 100.0));
    }

    if let Some(pitch) = pitch {
        let pitch = pitch.clamp(-12, 12);
        pitch_percent = Some(((pitch as f64 + 12.0) / 24.0 * 100.0).clamp(0.0, 100.0));
    }

    if let Some(volume) = volume {
        let volume = volume.clamp(0, 100);
        volume_percent = Some(volume as f64);
    }

    let scale_with_default = |percent: f64, min: f64, max: f64, default: f64| -> f64 {
        let percent = percent.clamp(0.0, 100.0);
        if percent <= 50.0 {
            let span = (default - min).max(0.0);
            min + span * (percent / 50.0)
        } else {
            let span = (max - default).max(0.0);
            default + span * ((percent - 50.0) / 50.0)
        }
    };

    if let Some(percent) = speed_percent {
        let mut old_val: DWORD = 0;
        if (attr_vtbl.speed_get)(attr_ptr, &mut old_val) == S_OK {
            if (attr_vtbl.speed_set)(attr_ptr, TTSATTR_MINSPEED) != S_OK { eprintln!("Failed to set min speed"); }
            let mut min_val: DWORD = 0;
            if (attr_vtbl.speed_get)(attr_ptr, &mut min_val) == S_OK {
                if (attr_vtbl.speed_set)(attr_ptr, TTSATTR_MAXSPEED) != S_OK { eprintln!("Failed to set max speed"); }
                let mut max_val: DWORD = 0;
                if (attr_vtbl.speed_get)(attr_ptr, &mut max_val) == S_OK {
                    let max_val = max_val.saturating_sub(1);
                    if max_val > min_val {
                        let default_val = old_val.clamp(min_val, max_val);
                        let scaled = scale_with_default(
                            percent,
                            min_val as f64,
                            max_val as f64,
                            default_val as f64,
                        );
                        if (attr_vtbl.speed_set)(attr_ptr, scaled as u32) != S_OK {
                            eprintln!("Failed to set speed");
                        }
                    } else if (attr_vtbl.speed_set)(attr_ptr, old_val) != S_OK {
                        eprintln!("Failed to restore speed");
                    }
                }
            }
        }
    }

    if let Some(percent) = pitch_percent {
        let mut old_val: WORD = 0;
        if (attr_vtbl.pitch_get)(attr_ptr, &mut old_val) == S_OK {
            if (attr_vtbl.pitch_set)(attr_ptr, TTSATTR_MINPITCH) != S_OK { eprintln!("Failed to set min pitch"); }
            let mut min_val: WORD = 0;
            if (attr_vtbl.pitch_get)(attr_ptr, &mut min_val) == S_OK {
                if (attr_vtbl.pitch_set)(attr_ptr, TTSATTR_MAXPITCH) != S_OK { eprintln!("Failed to set max pitch"); }
                let mut max_val: WORD = 0;
                if (attr_vtbl.pitch_get)(attr_ptr, &mut max_val) == S_OK {
                    if max_val > min_val {
                        let default_val = old_val.clamp(min_val, max_val);
                        let scaled = scale_with_default(
                            percent,
                            min_val as f64,
                            max_val as f64,
                            default_val as f64,
                        );
                        if (attr_vtbl.pitch_set)(attr_ptr, scaled as u16) != S_OK {
                            eprintln!("Failed to set pitch");
                        }
                    } else if (attr_vtbl.pitch_set)(attr_ptr, old_val) != S_OK {
                        eprintln!("Failed to restore pitch");
                    }
                }
            }
        }
    }

    if let Some(percent) = volume_percent {
        if (attr_vtbl.volume_set)(attr_ptr, TTSATTR_MINVOLUME) != S_OK { eprintln!("Failed to set min volume"); }
        let mut min_val: DWORD = 0;
        if (attr_vtbl.volume_get)(attr_ptr, &mut min_val) == S_OK {
            let min_val = min_val & 0xFFFF;
            if (attr_vtbl.volume_set)(attr_ptr, TTSATTR_MAXVOLUME) != S_OK { eprintln!("Failed to set max volume"); }
            let mut max_val: DWORD = 0;
            if (attr_vtbl.volume_get)(attr_ptr, &mut max_val) == S_OK {
                let max_val = max_val & 0xFFFF;
                if max_val > min_val {
                    let scaled = min_val as f64 + (max_val - min_val) as f64 * (percent / 100.0);
                    let val = scaled as u32;
                    let val = (val & 0xFFFF) | (val << 16);
                    if (attr_vtbl.volume_set)(attr_ptr, val) != S_OK {
                        eprintln!("Failed to set volume");
                    }
                }
            }
        }
    }

    (attr_vtbl.release)(attr_ptr);
}
