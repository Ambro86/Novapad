use crate::accessibility::from_wide;
use crate::mf_encoder;
use crate::settings;
use crate::settings::{PODCAST_DEVICE_DEFAULT, PodcastFormat};
use chrono::Local;
use std::collections::VecDeque;
use std::fs::{File, OpenOptions};
use std::io::{Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::{Duration, Instant};
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::RPC_E_CHANGED_MODE;
use windows::Win32::Media::Audio::{
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK,
    DEVICE_STATE_ACTIVE, EDataFlow, IAudioCaptureClient, IAudioClient, IMMDevice,
    IMMDeviceCollection, IMMDeviceEnumerator, MMDeviceEnumerator, WAVEFORMATEX,
    WAVEFORMATEXTENSIBLE, eCapture, eConsole, eRender,
};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT};
use windows::Win32::System::Com::StructuredStorage::PropVariantToStringAlloc;
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_MULTITHREADED, CoCreateInstance, CoInitializeEx, CoTaskMemFree,
    CoUninitialize, STGM_READ,
};
use windows::Win32::System::Power::{ES_CONTINUOUS, ES_SYSTEM_REQUIRED, SetThreadExecutionState};
use windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore;
use windows::core::PCWSTR;

const TARGET_SAMPLE_RATE: u32 = 44100;
const TARGET_CHANNELS: u16 = 2;
const TARGET_BITS: u16 = 16;
const MIX_CHUNK_FRAMES: usize = 512;

#[derive(Clone)]
pub struct AudioDevice {
    pub id: String,
    pub name: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SampleFormat {
    I16,
    F32,
}

struct DeviceEnumerator {
    _init: ComInit,
    inner: IMMDeviceEnumerator,
}

impl DeviceEnumerator {
    fn new() -> Result<Self, String> {
        let init = ComInit::new()?;
        let inner: IMMDeviceEnumerator = unsafe {
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
                .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
        };
        Ok(Self { _init: init, inner })
    }
}

struct ComInit {
    should_uninit: bool,
}

impl ComInit {
    fn new() -> Result<Self, String> {
        unsafe {
            let result = CoInitializeEx(None, COINIT_MULTITHREADED);
            if let Err(err) = result.ok() {
                if err.code() == RPC_E_CHANGED_MODE {
                    return Ok(Self {
                        should_uninit: false,
                    });
                }
                return Err(format!("CoInitializeEx failed: {err}"));
            }
        }
        Ok(Self {
            should_uninit: true,
        })
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        if self.should_uninit {
            unsafe {
                CoUninitialize();
            }
        }
    }
}

pub fn list_input_devices() -> Result<Vec<AudioDevice>, String> {
    list_devices(eCapture)
}

pub fn list_output_devices() -> Result<Vec<AudioDevice>, String> {
    list_devices(eRender)
}

pub fn probe_device(device_id: &str, loopback: bool) -> Result<(), String> {
    let _com = ComInit::new()?;
    let device = resolve_device(device_id, loopback)?;
    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };
    let mix_format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };
    let mut stream_flags = 0;
    if loopback {
        stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
    }
    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                stream_flags,
                10_000_000,
                0,
                mix_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
        CoTaskMemFree(Some(mix_format as *const _));
    }
    Ok(())
}

fn list_devices(flow: EDataFlow) -> Result<Vec<AudioDevice>, String> {
    let enumerator = DeviceEnumerator::new()?;
    let collection: IMMDeviceCollection = unsafe {
        enumerator
            .inner
            .EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE)
            .map_err(|e| format!("EnumAudioEndpoints failed: {e}"))?
    };
    let count = unsafe {
        collection
            .GetCount()
            .map_err(|e| format!("GetCount failed: {e}"))?
    };
    let mut devices = Vec::new();
    for index in 0..count {
        let device: IMMDevice = unsafe {
            collection
                .Item(index)
                .map_err(|e| format!("Device Item failed: {e}"))?
        };
        if let Some(info) = device_info(&device) {
            devices.push(info);
        }
    }
    Ok(devices)
}

fn device_id(device: &IMMDevice) -> Option<String> {
    unsafe {
        let id = device.GetId().ok()?;
        if id.is_null() {
            return None;
        }
        let value = from_wide(id.0);
        CoTaskMemFree(Some(id.0 as *const _));
        if value.is_empty() { None } else { Some(value) }
    }
}

fn device_info(device: &IMMDevice) -> Option<AudioDevice> {
    let id = device_id(device)?;
    let name = device_friendly_name(device).unwrap_or_else(|| id.clone());
    Some(AudioDevice { id, name })
}

fn device_friendly_name(device: &IMMDevice) -> Option<String> {
    unsafe {
        let store: IPropertyStore = device.OpenPropertyStore(STGM_READ).ok()?;
        let value = store.GetValue(&PKEY_Device_FriendlyName).ok()?;
        let name_ptr = PropVariantToStringAlloc(&value).ok()?;
        if name_ptr.is_null() {
            return None;
        }
        let name = from_wide(name_ptr.0);
        CoTaskMemFree(Some(name_ptr.0 as *const _));
        if name.is_empty() { None } else { Some(name) }
    }
}

#[derive(Clone)]
pub struct RecorderConfig {
    pub include_mic: bool,
    pub mic_device_id: String,
    pub include_system: bool,
    pub system_device_id: String,
    pub output_format: PodcastFormat,
    pub mp3_bitrate: u32,
    pub save_folder: PathBuf,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum RecorderStatus {
    Idle,
    Recording,
    Paused,
    Saving,
    Error,
}

pub struct RecorderHandle {
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
    threads: Vec<JoinHandle<Result<(), String>>>,
    output_path: PathBuf,
    temp_wav: PathBuf,
    temp_mp3: PathBuf,
    format: PodcastFormat,
    mp3_bitrate: u32,
}

struct SharedState {
    status: Mutex<RecorderStatus>,
    last_error: Mutex<Option<String>>,
    started_at: Mutex<Option<Instant>>,
    paused_at: Mutex<Option<Instant>>,
    paused_total: Mutex<Duration>,
    mic_peak: AtomicU32,
    system_peak: AtomicU32,
    include_mic: bool,
    include_system: bool,
}

impl SharedState {
    fn new(include_mic: bool, include_system: bool) -> Self {
        SharedState {
            status: Mutex::new(RecorderStatus::Idle),
            last_error: Mutex::new(None),
            started_at: Mutex::new(None),
            paused_at: Mutex::new(None),
            paused_total: Mutex::new(Duration::ZERO),
            mic_peak: AtomicU32::new(0),
            system_peak: AtomicU32::new(0),
            include_mic,
            include_system,
        }
    }
}

pub struct LevelSnapshot {
    pub mic_peak: u32,
    pub system_peak: u32,
}

pub fn start_recording(config: RecorderConfig) -> Result<RecorderHandle, String> {
    if !config.include_mic && !config.include_system {
        return Err("No audio sources selected.".to_string());
    }

    let output_folder = if config.save_folder.as_os_str().is_empty() {
        PathBuf::from(settings::default_podcast_save_folder())
    } else {
        config.save_folder.clone()
    };
    if let Some(parent) = output_folder.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let _ = std::fs::create_dir_all(&output_folder);

    let timestamp = Local::now().format("%Y-%m-%d_%H-%M-%S").to_string();
    let base_name = format!("Podcast_{timestamp}");
    let output_path = output_folder.join(format!(
        "{}.{}",
        base_name,
        match config.output_format {
            PodcastFormat::Mp3 => "mp3",
            PodcastFormat::Wav => "wav",
        }
    ));
    let temp_wav = output_folder.join(format!("{base_name}.wav.tmp"));
    let temp_mp3 = output_folder.join(format!("{base_name}_tmp.mp3"));

    let shared = Arc::new(SharedState::new(config.include_mic, config.include_system));
    *shared.status.lock().unwrap() = RecorderStatus::Recording;
    *shared.started_at.lock().unwrap() = Some(Instant::now());

    let stop = Arc::new(AtomicBool::new(false));
    let paused = Arc::new(AtomicBool::new(false));

    let mix_buffer = Arc::new(MixBuffer::new());
    let mut threads = Vec::new();

    if config.include_mic {
        let buffer = mix_buffer.clone();
        let shared_state = shared.clone();
        let stop_flag = stop.clone();
        let paused_flag = paused.clone();
        let device_id = config.mic_device_id.clone();
        threads.push(thread::spawn(move || {
            let result = capture_source(
                SourceKind::Microphone,
                &device_id,
                false,
                buffer,
                shared_state.clone(),
                stop_flag.clone(),
                paused_flag,
            );
            if let Err(err) = &result {
                if let Ok(mut error) = shared_state.last_error.lock() {
                    *error = Some(err.clone());
                }
                if let Ok(mut status) = shared_state.status.lock() {
                    *status = RecorderStatus::Error;
                }
                stop_flag.store(true, Ordering::SeqCst);
            }
            result
        }));
    }

    if config.include_system {
        let buffer = mix_buffer.clone();
        let shared_state = shared.clone();
        let stop_flag = stop.clone();
        let paused_flag = paused.clone();
        let device_id = config.system_device_id.clone();
        threads.push(thread::spawn(move || {
            let result = capture_source(
                SourceKind::System,
                &device_id,
                true,
                buffer,
                shared_state.clone(),
                stop_flag.clone(),
                paused_flag,
            );
            if let Err(err) = &result {
                if let Ok(mut error) = shared_state.last_error.lock() {
                    *error = Some(err.clone());
                }
                if let Ok(mut status) = shared_state.status.lock() {
                    *status = RecorderStatus::Error;
                }
                stop_flag.store(true, Ordering::SeqCst);
            }
            result
        }));
    }

    let keep_awake_stop = stop.clone();
    threads.push(thread::spawn(move || keep_awake_loop(keep_awake_stop)));

    let writer_buffer = mix_buffer.clone();
    let writer_shared = shared.clone();
    let writer_stop = stop.clone();
    let writer_paused = paused.clone();
    let writer_path = temp_wav.clone();
    threads.push(thread::spawn(move || {
        let result = write_mixed_audio(
            writer_path,
            writer_buffer,
            writer_shared.clone(),
            writer_stop.clone(),
            writer_paused,
        );
        if let Err(err) = &result {
            if let Ok(mut error) = writer_shared.last_error.lock() {
                *error = Some(err.clone());
            }
            if let Ok(mut status) = writer_shared.status.lock() {
                *status = RecorderStatus::Error;
            }
            writer_stop.store(true, Ordering::SeqCst);
        }
        result
    }));

    Ok(RecorderHandle {
        shared,
        stop,
        paused,
        threads,
        output_path,
        temp_wav,
        temp_mp3,
        format: config.output_format,
        mp3_bitrate: config.mp3_bitrate,
    })
}

impl RecorderHandle {
    pub fn pause(&self) {
        if !self.paused.swap(true, Ordering::SeqCst) {
            if let Ok(mut paused_at) = self.shared.paused_at.lock() {
                *paused_at = Some(Instant::now());
            }
            if let Ok(mut status) = self.shared.status.lock() {
                *status = RecorderStatus::Paused;
            }
        }
    }

    pub fn resume(&self) {
        if self.paused.swap(false, Ordering::SeqCst) {
            let now = Instant::now();
            if let Ok(mut paused_at) = self.shared.paused_at.lock() {
                if let Some(start) = paused_at.take() {
                    if let Ok(mut total) = self.shared.paused_total.lock() {
                        *total += now.saturating_duration_since(start);
                    }
                }
            }
            if let Ok(mut status) = self.shared.status.lock() {
                *status = RecorderStatus::Recording;
            }
        }
    }

    pub fn stop(self) -> Result<PathBuf, String> {
        self.stop_with_progress(|_| {}, None)
    }

    pub fn stop_with_progress<F>(
        mut self,
        mut progress: F,
        cancel: Option<Arc<AtomicBool>>,
    ) -> Result<PathBuf, String>
    where
        F: FnMut(u32),
    {
        self.stop.store(true, Ordering::SeqCst);
        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Saving;
        }
        let threads = std::mem::take(&mut self.threads);
        for handle in threads {
            match handle.join() {
                Ok(Ok(())) => {}
                Ok(Err(err)) => {
                    self.set_error(&err);
                    return Err(err);
                }
                Err(_) => {
                    let err = "Recording thread panicked.".to_string();
                    self.set_error(&err);
                    return Err(err);
                }
            }
        }

        if let Some(cancel) = cancel.as_ref() {
            if cancel.load(Ordering::Relaxed) {
                let _ = std::fs::remove_file(&self.temp_wav);
                let _ = std::fs::remove_file(&self.temp_mp3);
                return Err("Saving canceled.".to_string());
            }
        }

        if self.format == PodcastFormat::Mp3 {
            if let Err(err) = mf_encoder::encode_wav_to_mp3_with_bitrate_progress(
                &self.temp_wav,
                &self.temp_mp3,
                self.mp3_bitrate,
                &mut progress,
                cancel.as_deref(),
            ) {
                let _ = std::fs::remove_file(&self.temp_wav);
                let _ = std::fs::remove_file(&self.temp_mp3);
                self.set_error(&err);
                return Err(err);
            }
            let _ = std::fs::remove_file(&self.temp_wav);
            if let Err(err) = rename_atomic(&self.temp_mp3, &self.output_path) {
                self.set_error(&err);
                return Err(err);
            }
        } else {
            progress(100);
            if let Err(err) = rename_atomic(&self.temp_wav, &self.output_path) {
                self.set_error(&err);
                return Err(err);
            }
        }

        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Idle;
        }
        Ok(self.output_path.clone())
    }

    pub fn status(&self) -> RecorderStatus {
        self.shared
            .status
            .lock()
            .map(|status| *status)
            .unwrap_or(RecorderStatus::Error)
    }

    pub fn levels(&self) -> LevelSnapshot {
        LevelSnapshot {
            mic_peak: self.shared.mic_peak.load(Ordering::Relaxed),
            system_peak: self.shared.system_peak.load(Ordering::Relaxed),
        }
    }

    pub fn elapsed(&self) -> Duration {
        let start = self.shared.started_at.lock().ok().and_then(|s| *s);
        let start = match start {
            Some(value) => value,
            None => return Duration::ZERO,
        };
        let paused_total = self
            .shared
            .paused_total
            .lock()
            .map(|v| *v)
            .unwrap_or(Duration::ZERO);
        let paused_at = self.shared.paused_at.lock().ok().and_then(|s| *s);
        let now = Instant::now();
        let mut elapsed = now.saturating_duration_since(start);
        if let Some(paused_at) = paused_at {
            elapsed = paused_at.saturating_duration_since(start);
        }
        elapsed.saturating_sub(paused_total)
    }

    pub fn take_error(&self) -> Option<String> {
        self.shared.last_error.lock().ok()?.take()
    }

    fn set_error(&self, message: &str) {
        if let Ok(mut err) = self.shared.last_error.lock() {
            *err = Some(message.to_string());
        }
        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Error;
        }
    }
}

fn rename_atomic(src: &Path, dest: &Path) -> Result<(), String> {
    if dest.exists() {
        let _ = std::fs::remove_file(dest);
    }
    std::fs::rename(src, dest).map_err(|e| e.to_string())
}

fn keep_awake_loop(stop: Arc<AtomicBool>) -> Result<(), String> {
    unsafe {
        let _ = SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
    }
    while !stop.load(Ordering::SeqCst) {
        thread::sleep(Duration::from_secs(30));
        unsafe {
            let _ = SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
        }
    }
    unsafe {
        let _ = SetThreadExecutionState(ES_CONTINUOUS);
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum SourceKind {
    Microphone,
    System,
}

struct MixBuffer {
    inner: Mutex<MixQueues>,
    condvar: Condvar,
}

struct MixQueues {
    mic: VecDeque<f32>,
    system: VecDeque<f32>,
}

impl MixBuffer {
    fn new() -> Self {
        MixBuffer {
            inner: Mutex::new(MixQueues {
                mic: VecDeque::new(),
                system: VecDeque::new(),
            }),
            condvar: Condvar::new(),
        }
    }

    fn push(&self, source: SourceKind, samples: Vec<f32>) {
        let mut inner = self.inner.lock().unwrap_or_else(|e| e.into_inner());
        match source {
            SourceKind::Microphone => inner.mic.extend(samples),
            SourceKind::System => inner.system.extend(samples),
        }
        self.condvar.notify_one();
    }
}

fn write_mixed_audio(
    path: PathBuf,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let mut writer = WavWriter::create(&path)?;
    let mut last_write = Instant::now();
    loop {
        if stop.load(Ordering::SeqCst) {
            break;
        }
        if paused.load(Ordering::SeqCst) {
            thread::sleep(Duration::from_millis(30));
            continue;
        }

        let mut inner = buffer.inner.lock().unwrap_or_else(|e| e.into_inner());
        let (need_mic, need_sys) = (shared.include_mic, shared.include_system);
        let available_mic = inner.mic.len() / TARGET_CHANNELS as usize;
        let available_sys = inner.system.len() / TARGET_CHANNELS as usize;
        let can_mix = if need_mic && need_sys {
            available_mic >= MIX_CHUNK_FRAMES && available_sys >= MIX_CHUNK_FRAMES
        } else if need_mic {
            available_mic >= MIX_CHUNK_FRAMES
        } else {
            available_sys >= MIX_CHUNK_FRAMES
        };

        if !can_mix {
            let _ = buffer
                .condvar
                .wait_timeout(inner, Duration::from_millis(40));
            continue;
        }

        let frames = MIX_CHUNK_FRAMES;
        let mut mixed = Vec::with_capacity(frames * TARGET_CHANNELS as usize);
        for _ in 0..frames {
            let mut left = 0.0f32;
            let mut right = 0.0f32;
            if need_mic {
                left += inner.mic.pop_front().unwrap_or(0.0);
                right += inner.mic.pop_front().unwrap_or(0.0);
            }
            if need_sys {
                left += inner.system.pop_front().unwrap_or(0.0);
                right += inner.system.pop_front().unwrap_or(0.0);
            }
            if need_mic && need_sys {
                left *= 0.5;
                right *= 0.5;
            }
            mixed.push(left.clamp(-1.0, 1.0));
            mixed.push(right.clamp(-1.0, 1.0));
        }
        drop(inner);

        writer.write_f32(&mixed)?;

        let elapsed = last_write.elapsed();
        if elapsed < Duration::from_millis(10) {
            thread::sleep(Duration::from_millis(5));
        }
        last_write = Instant::now();
    }
    writer.finalize()?;
    Ok(())
}

fn capture_source(
    kind: SourceKind,
    device_id: &str,
    loopback: bool,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let _com = ComInit::new()?;
    let device = resolve_device(device_id, loopback)?;
    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };

    let mix_format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };
    let (input_rate, input_channels, input_format) = parse_format(unsafe { &*mix_format });

    let mut stream_flags = 0;
    if loopback {
        stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
    }
    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                stream_flags,
                10_000_000,
                0,
                mix_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
    }
    unsafe {
        CoTaskMemFree(Some(mix_format as *const _));
    }

    let capture: IAudioCaptureClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService capture failed: {e}"))?
    };
    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }

    let mut resampler =
        LinearResampler::new(input_rate, TARGET_SAMPLE_RATE, input_channels as usize);

    loop {
        if stop.load(Ordering::SeqCst) {
            break;
        }
        if paused.load(Ordering::SeqCst) {
            thread::sleep(Duration::from_millis(15));
            continue;
        }

        let mut packet_len = unsafe {
            capture
                .GetNextPacketSize()
                .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
        };
        while packet_len > 0 {
            let mut data_ptr: *mut u8 = std::ptr::null_mut();
            let mut frames = 0u32;
            let mut flags = 0u32;
            unsafe {
                capture
                    .GetBuffer(&mut data_ptr, &mut frames, &mut flags, None, None)
                    .map_err(|e| format!("GetBuffer failed: {e}"))?;
            }
            let samples = if flags & (AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0 {
                vec![0f32; frames as usize * input_channels as usize]
            } else {
                read_samples(data_ptr, frames, input_channels, input_format)
            };
            unsafe {
                capture
                    .ReleaseBuffer(frames)
                    .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
            }

            update_peak(&shared, &kind, &samples);
            let resampled = resampler.push(&samples);
            let stereo = to_stereo(&resampled, input_channels as usize);
            buffer.push(kind, stereo);
            packet_len = unsafe {
                capture
                    .GetNextPacketSize()
                    .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
            };
        }
        thread::sleep(Duration::from_millis(10));
    }

    unsafe {
        let _ = client.Stop();
    }
    Ok(())
}

fn resolve_device(device_id: &str, loopback: bool) -> Result<IMMDevice, String> {
    let enumerator = DeviceEnumerator::new()?;
    if device_id.is_empty() || device_id == PODCAST_DEVICE_DEFAULT {
        let flow = if loopback { eRender } else { eCapture };
        return unsafe {
            enumerator
                .inner
                .GetDefaultAudioEndpoint(flow, eConsole)
                .map_err(|e| format!("GetDefaultAudioEndpoint failed: {e}"))
        };
    }
    let wide = crate::accessibility::to_wide(device_id);
    unsafe {
        enumerator
            .inner
            .GetDevice(PCWSTR(wide.as_ptr()))
            .map_err(|e| format!("GetDevice failed: {e}"))
    }
}

fn parse_format(fmt: &WAVEFORMATEX) -> (u32, u16, SampleFormat) {
    let channels = fmt.nChannels;
    let rate = fmt.nSamplesPerSec;
    let mut format = match fmt.wFormatTag as u32 {
        WAVE_FORMAT_IEEE_FLOAT => SampleFormat::F32,
        _ => SampleFormat::I16,
    };
    if fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE {
        let ext = unsafe { &*(fmt as *const _ as *const WAVEFORMATEXTENSIBLE) };
        let subformat = unsafe { std::ptr::read_unaligned(std::ptr::addr_of!(ext.SubFormat)) };
        if subformat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
            format = SampleFormat::F32;
        } else {
            format = SampleFormat::I16;
        }
    }
    (rate, channels, format)
}

fn read_samples(ptr: *mut u8, frames: u32, channels: u16, format: SampleFormat) -> Vec<f32> {
    let sample_count = frames as usize * channels as usize;
    if ptr.is_null() || sample_count == 0 {
        return Vec::new();
    }
    unsafe {
        match format {
            SampleFormat::F32 => {
                let slice = std::slice::from_raw_parts(ptr as *const f32, sample_count);
                slice.to_vec()
            }
            SampleFormat::I16 => {
                let slice = std::slice::from_raw_parts(ptr as *const i16, sample_count);
                slice.iter().map(|s| *s as f32 / i16::MAX as f32).collect()
            }
        }
    }
}

fn update_peak(shared: &SharedState, kind: &SourceKind, samples: &[f32]) {
    let mut peak = 0f32;
    for sample in samples {
        let abs = sample.abs();
        if abs > peak {
            peak = abs;
        }
    }
    let value = (peak * i16::MAX as f32) as u32;
    match kind {
        SourceKind::Microphone => {
            shared.mic_peak.store(value, Ordering::Relaxed);
        }
        SourceKind::System => {
            shared.system_peak.store(value, Ordering::Relaxed);
        }
    }
}

fn to_stereo(samples: &[f32], channels: usize) -> Vec<f32> {
    if channels == TARGET_CHANNELS as usize {
        return samples.to_vec();
    }
    let frames = samples.len() / channels;
    let mut out = Vec::with_capacity(frames * TARGET_CHANNELS as usize);
    for frame in 0..frames {
        let base = frame * channels;
        let left = samples[base];
        let right = if channels > 1 {
            samples[base + 1]
        } else {
            left
        };
        out.push(left);
        out.push(right);
    }
    out
}

struct LinearResampler {
    input_rate: u32,
    output_rate: u32,
    channels: usize,
    pos: f64,
    buffer: Vec<f32>,
}

impl LinearResampler {
    fn new(input_rate: u32, output_rate: u32, channels: usize) -> Self {
        LinearResampler {
            input_rate,
            output_rate,
            channels,
            pos: 0.0,
            buffer: Vec::new(),
        }
    }

    fn push(&mut self, samples: &[f32]) -> Vec<f32> {
        self.buffer.extend_from_slice(samples);
        if self.input_rate == 0 || self.output_rate == 0 || self.channels == 0 {
            return Vec::new();
        }
        let step = self.input_rate as f64 / self.output_rate as f64;
        let frames_available = self.buffer.len() / self.channels;
        let mut out = Vec::new();
        while self.pos + 1.0 < frames_available as f64 {
            let i0 = self.pos.floor() as usize;
            let i1 = i0 + 1;
            let frac = self.pos - i0 as f64;
            for ch in 0..self.channels {
                let s0 = self.buffer[i0 * self.channels + ch];
                let s1 = self.buffer[i1 * self.channels + ch];
                out.push((1.0 - frac as f32) * s0 + (frac as f32) * s1);
            }
            self.pos += step;
        }
        let drop_frames = self.pos.floor() as usize;
        if drop_frames > 0 {
            let drop_samples = drop_frames * self.channels;
            self.buffer.drain(0..drop_samples);
            self.pos -= drop_frames as f64;
        }
        out
    }
}

struct WavWriter {
    file: File,
    data_size: u32,
}

impl WavWriter {
    fn create(path: &Path) -> Result<Self, String> {
        let file = OpenOptions::new()
            .create(true)
            .truncate(true)
            .write(true)
            .open(path)
            .map_err(|e| e.to_string())?;
        let mut writer = WavWriter { file, data_size: 0 };
        writer.write_header_placeholder()?;
        Ok(writer)
    }

    fn write_header_placeholder(&mut self) -> Result<(), String> {
        self.file.write_all(b"RIFF").map_err(|e| e.to_string())?;
        self.file
            .write_all(&0u32.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file.write_all(b"WAVE").map_err(|e| e.to_string())?;
        self.file.write_all(b"fmt ").map_err(|e| e.to_string())?;
        self.file
            .write_all(&16u32.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&1u16.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&TARGET_CHANNELS.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&TARGET_SAMPLE_RATE.to_le_bytes())
            .map_err(|e| e.to_string())?;
        let byte_rate = TARGET_SAMPLE_RATE * TARGET_CHANNELS as u32 * (TARGET_BITS as u32 / 8);
        let block_align = TARGET_CHANNELS * (TARGET_BITS / 8);
        self.file
            .write_all(&byte_rate.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&block_align.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&TARGET_BITS.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file.write_all(b"data").map_err(|e| e.to_string())?;
        self.file
            .write_all(&0u32.to_le_bytes())
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn write_f32(&mut self, samples: &[f32]) -> Result<(), String> {
        let mut buf = Vec::with_capacity(samples.len() * 2);
        for sample in samples {
            let clamped = sample.clamp(-1.0, 1.0);
            let v = (clamped * i16::MAX as f32) as i16;
            buf.extend_from_slice(&v.to_le_bytes());
        }
        self.file.write_all(&buf).map_err(|e| e.to_string())?;
        self.data_size = self.data_size.saturating_add(buf.len() as u32);
        Ok(())
    }

    fn finalize(&mut self) -> Result<(), String> {
        let riff_size = 36u32.saturating_add(self.data_size);
        self.file
            .seek(SeekFrom::Start(4))
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&riff_size.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .seek(SeekFrom::Start(40))
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&self.data_size.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file.flush().map_err(|e| e.to_string())?;
        Ok(())
    }
}

pub fn default_output_folder() -> PathBuf {
    PathBuf::from(settings::default_podcast_save_folder())
}
