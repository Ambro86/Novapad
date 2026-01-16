use crate::accessibility::from_wide;
use crate::audio_capture::{self, AudioRecorderHandle as AudioRecorder};
use crate::audio_utils;
use crate::mf_encoder;
use crate::settings;
use crate::settings::{PODCAST_DEVICE_DEFAULT, PodcastFormat};
use chrono::Local;
use std::collections::VecDeque;
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
    pub mic_gain: f32,
    pub include_system: bool,
    pub system_device_id: String,
    pub system_gain: f32,
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
        return Err("No sources selected.".to_string());
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

    // Audio-only path
    let shared = Arc::new(SharedState::new(config.include_mic, config.include_system));
    *shared.status.lock().unwrap() = RecorderStatus::Recording;
    *shared.started_at.lock().unwrap() = Some(Instant::now());

    let stop = Arc::new(AtomicBool::new(false));
    let paused = Arc::new(AtomicBool::new(false));

    let mix_buffer = Arc::new(MixBuffer::new());
    let mut threads = Vec::new();

    if config.include_mic {
        crate::log_debug("Starting microphone capture thread");
        let buffer = mix_buffer.clone();
        let shared_state = shared.clone();
        let stop_flag = stop.clone();
        let paused_flag = paused.clone();
        let device_id = config.mic_device_id.clone();
        let mic_gain = config.mic_gain;
        threads.push(thread::spawn(move || {
            crate::log_debug("Microphone capture thread started");
            let result = capture_source(
                SourceKind::Microphone,
                &device_id,
                false,
                mic_gain,
                buffer,
                shared_state.clone(),
                stop_flag.clone(),
                paused_flag,
            );
            if let Err(err) = &result {
                crate::log_debug(&format!("Microphone capture error: {}", err));
                if let Ok(mut error) = shared_state.last_error.lock() {
                    *error = Some(err.clone());
                }
                if let Ok(mut status) = shared_state.status.lock() {
                    *status = RecorderStatus::Error;
                }
                stop_flag.store(true, Ordering::SeqCst);
            } else {
                crate::log_debug("Microphone capture thread stopped normally");
            }
            result
        }));
    }

    if config.include_system {
        crate::log_debug("Starting system audio capture thread");
        let buffer = mix_buffer.clone();
        let shared_state = shared.clone();
        let stop_flag = stop.clone();
        let paused_flag = paused.clone();
        let device_id = config.system_device_id.clone();
        let system_gain = config.system_gain;
        threads.push(thread::spawn(move || {
            crate::log_debug("System audio capture thread started");
            let result = capture_source(
                SourceKind::System,
                &device_id,
                true,
                system_gain,
                buffer,
                shared_state.clone(),
                stop_flag.clone(),
                paused_flag,
            );
            if let Err(err) = &result {
                crate::log_debug(&format!("System audio capture error: {}", err));
                if let Ok(mut error) = shared_state.last_error.lock() {
                    *error = Some(err.clone());
                }
                if let Ok(mut status) = shared_state.status.lock() {
                    *status = RecorderStatus::Error;
                }
                stop_flag.store(true, Ordering::SeqCst);
            } else {
                crate::log_debug("System audio capture thread stopped normally");
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
    let writer_path = match config.output_format {
        PodcastFormat::Mp3 => temp_mp3.clone(),
        PodcastFormat::Wav => temp_wav.clone(),
    };
    let writer_format = config.output_format;
    let writer_bitrate = config.mp3_bitrate;
    threads.push(thread::spawn(move || {
        let result = write_mixed_audio(
            writer_path,
            writer_format,
            writer_bitrate,
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
    })
}

/// Start microphone audio recording using WASAPI (non-loopback)
#[allow(dead_code)]
fn start_mic_audio_recording(device_id: &str) -> Result<AudioRecorder, String> {
    use crate::audio_capture::AudioQueue;

    let audio_queue = Arc::new(AudioQueue::new(3000));
    let stop = Arc::new(AtomicBool::new(false));

    // Get mic format
    let _com = ComInit::new()?;
    let device = resolve_device(device_id, false)?; // false = not loopback

    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };

    let format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };

    let sample_rate = unsafe { (*format).nSamplesPerSec };
    let channels = unsafe { (*format).nChannels };

    crate::log_debug(&format!(
        "Mic audio capture will use: {} Hz, {} channels",
        sample_rate, channels
    ));

    let audio_queue_clone = Arc::clone(&audio_queue);
    let stop_clone = Arc::clone(&stop);
    let device_id_clone = device_id.to_string();

    let thread =
        thread::spawn(move || mic_capture_loop(audio_queue_clone, stop_clone, &device_id_clone));

    Ok(audio_capture::create_audio_recorder_handle(
        stop,
        thread,
        audio_queue,
        sample_rate,
        channels,
    ))
}

/// Microphone capture loop - similar to audio_capture but for mic
#[allow(dead_code)]
fn mic_capture_loop(
    audio_queue: Arc<audio_capture::AudioQueue>,
    stop: Arc<AtomicBool>,
    device_id: &str,
) -> Result<(), String> {
    use crate::audio_capture::AudioSample;

    let _com = ComInit::new()?;
    let device = resolve_device(device_id, false)?; // false = not loopback

    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };

    let format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };

    let sample_rate = unsafe { (*format).nSamplesPerSec };
    let channels = unsafe { (*format).nChannels };

    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                0, // No loopback for mic
                10_000_000,
                0,
                format,
                None,
            )
            .map_err(|e| format!("Initialize failed: {e}"))?;
    }

    let capture_client: IAudioCaptureClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService failed: {e}"))?
    };

    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }

    crate::log_debug("Mic capture loop started");

    while !stop.load(Ordering::SeqCst) {
        thread::sleep(Duration::from_millis(10));

        let packet_length = unsafe {
            capture_client
                .GetNextPacketSize()
                .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
        };

        if packet_length == 0 {
            continue;
        }

        let mut buffer: *mut u8 = std::ptr::null_mut();
        let mut num_frames = 0u32;
        let mut flags = 0u32;

        unsafe {
            capture_client
                .GetBuffer(&mut buffer, &mut num_frames, &mut flags, None, None)
                .map_err(|e| format!("GetBuffer failed: {e}"))?;
        }

        if num_frames > 0 {
            let buffer_size = (num_frames * channels as u32 * 2) as usize;
            let audio_data =
                unsafe { std::slice::from_raw_parts(buffer as *const i16, buffer_size / 2) };

            let sample = AudioSample {
                data: audio_data.to_vec(),
                sample_rate,
                channels,
            };

            audio_queue.push(sample);

            unsafe {
                capture_client
                    .ReleaseBuffer(num_frames)
                    .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
            }
        }
    }

    unsafe {
        client.Stop().ok();
    }

    crate::log_debug("Mic capture loop stopped");
    Ok(())
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
        crate::log_debug("Stopping podcast recording");

        // Signal encoder to stop (so it knows to drain queues and exit)
        self.stop.store(true, Ordering::SeqCst);
        crate::log_debug("Signaled encoder to stop");

        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Saving;
        }

        // Wait for all threads to finish
        let threads = std::mem::take(&mut self.threads);
        for handle in threads {
            match handle.join() {
                Ok(Ok(())) => {}
                Ok(Err(err)) => {
                    crate::log_debug(&format!("Thread error: {}", err));
                    self.set_error(&err);
                    return Err(err);
                }
                Err(_) => {
                    let err = "Recording thread panicked.".to_string();
                    crate::log_debug(&err);
                    self.set_error(&err);
                    return Err(err);
                }
            }
        }
        crate::log_debug("All threads stopped");

        if let Some(cancel) = cancel.as_ref() {
            if cancel.load(Ordering::Relaxed) {
                let _ = std::fs::remove_file(&self.temp_wav);
                let _ = std::fs::remove_file(&self.temp_mp3);
                return Err("Saving canceled.".to_string());
            }
        }

        if self.format == PodcastFormat::Mp3 {
            progress(100);
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

    #[allow(dead_code)]
    fn pop_chunk(&self, timeout: Duration) -> Option<Vec<i16>> {
        let mut inner = self.inner.lock().unwrap_or_else(|e| e.into_inner());

        // Wait for data if queues are empty
        if inner.mic.is_empty() && inner.system.is_empty() {
            let result = self.condvar.wait_timeout(inner, timeout).unwrap();
            inner = result.0;
        }

        // Determine how many samples we can mix
        let mic_len = inner.mic.len();
        let system_len = inner.system.len();

        if mic_len == 0 && system_len == 0 {
            return None;
        }

        // Take up to MIX_CHUNK_FRAMES samples
        let count = mic_len.max(system_len).min(MIX_CHUNK_FRAMES);
        let mut output = Vec::with_capacity(count);

        for _ in 0..count {
            let mic_sample = inner.mic.pop_front().unwrap_or(0.0);
            let system_sample = inner.system.pop_front().unwrap_or(0.0);

            // Mix and convert to i16
            let mixed = mic_sample + system_sample;
            let clamped = mixed.clamp(-1.0, 1.0);
            let sample_i16 = (clamped * i16::MAX as f32) as i16;
            output.push(sample_i16);
        }

        Some(output)
    }

    #[allow(dead_code)]
    fn is_empty(&self) -> bool {
        let inner = self.inner.lock().unwrap_or_else(|e| e.into_inner());
        inner.mic.is_empty() && inner.system.is_empty()
    }
}

fn write_mixed_audio(
    path: PathBuf,
    format: PodcastFormat,
    mp3_bitrate: u32,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    match format {
        PodcastFormat::Mp3 => {
            write_mixed_audio_mp3(path, mp3_bitrate, buffer, shared, stop, paused)
        }
        PodcastFormat::Wav => write_mixed_audio_wav(path, buffer, shared, stop, paused),
    }
}

fn write_mixed_audio_wav(
    path: PathBuf,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let mut writer =
        audio_utils::WavWriter::create(&path, TARGET_SAMPLE_RATE, TARGET_CHANNELS, TARGET_BITS)
            .map_err(|e| e.to_string())?;

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

        writer
            .write_samples_f32(&mixed)
            .map_err(|e| e.to_string())?;

        let elapsed = last_write.elapsed();
        if elapsed < Duration::from_millis(10) {
            thread::sleep(Duration::from_millis(5));
        }
        last_write = Instant::now();
    }
    writer.finalize().map_err(|e| e.to_string())?;
    Ok(())
}

fn write_mixed_audio_mp3(
    path: PathBuf,
    mp3_bitrate: u32,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let mut writer = mf_encoder::Mp3StreamWriter::create(
        &path,
        mp3_bitrate,
        TARGET_SAMPLE_RATE,
        TARGET_CHANNELS,
    )?;
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

        let mut pcm = Vec::with_capacity(mixed.len());
        for sample in mixed {
            let v = (sample.clamp(-1.0, 1.0) * i16::MAX as f32) as i16;
            pcm.push(v);
        }
        writer.write_i16(&pcm)?;

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
    gain: f32,
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
            let mut stereo = to_stereo(&resampled, input_channels as usize);

            // Apply gain
            if gain != 1.0 {
                for sample in stereo.iter_mut() {
                    *sample = (*sample * gain).clamp(-1.0, 1.0);
                }
            }

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
    // Note: COM must already be initialized by the caller and kept alive
    // for the lifetime of the returned device.
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };

    if device_id.is_empty() || device_id == PODCAST_DEVICE_DEFAULT {
        let flow = if loopback { eRender } else { eCapture };
        return unsafe {
            enumerator
                .GetDefaultAudioEndpoint(flow, eConsole)
                .map_err(|e| format!("GetDefaultAudioEndpoint failed: {e}"))
        };
    }

    let wide = crate::accessibility::to_wide(device_id);
    unsafe {
        enumerator
            .GetDevice(PCWSTR(wide.as_ptr()))
            .map_err(|e| format!("GetDevice({}) failed: {e}", device_id))
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

pub fn default_output_folder() -> PathBuf {
    PathBuf::from(settings::default_podcast_save_folder())
}
