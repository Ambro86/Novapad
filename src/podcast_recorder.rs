use crate::audio_utils;
use crate::com_guard::ComGuard;
use crate::mf_encoder;
use crate::settings;
use crate::settings::{PODCAST_DEVICE_DEFAULT, PodcastFormat};
use chrono::Local;
use std::collections::VecDeque;
use std::ffi::OsString;
use std::mem::ManuallyDrop;
use std::mem::size_of;
use std::os::windows::ffi::OsStringExt;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::{Duration, Instant};
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::CloseHandle;
use windows::Win32::Media::Audio::{
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
    AUDCLNT_STREAMFLAGS_LOOPBACK, AUDIOCLIENT_ACTIVATION_PARAMS, AUDIOCLIENT_ACTIVATION_PARAMS_0,
    AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK, AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS,
    ActivateAudioInterfaceAsync, AudioSessionStateActive, DEVICE_STATE_ACTIVE, EDataFlow,
    IActivateAudioInterfaceAsyncOperation, IActivateAudioInterfaceCompletionHandler,
    IAudioCaptureClient, IAudioClient, IAudioSessionControl, IAudioSessionControl2,
    IAudioSessionEnumerator, IAudioSessionManager2, IMMDevice, IMMDeviceCollection,
    IMMDeviceEnumerator, MMDeviceEnumerator, PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
    VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, WAVEFORMATEX, eCapture, eConsole, eRender,
};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT};
use windows::Win32::System::Com::StructuredStorage::PropVariantToStringAlloc;
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree, STGM_READ};
use windows::Win32::System::Power::{ES_CONTINUOUS, ES_SYSTEM_REQUIRED, SetThreadExecutionState};
use windows::Win32::System::Threading::{
    OpenProcess, PROCESS_NAME_FORMAT, PROCESS_QUERY_LIMITED_INFORMATION, QueryFullProcessImageNameW,
};
use windows::Win32::System::Variant::VT_BLOB;
use windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore;
use windows::core::{HRESULT, Interface, PCWSTR, PROPVARIANT, PWSTR, implement};

const TARGET_SAMPLE_RATE: u32 = 44100;
const TARGET_CHANNELS: u16 = 2;
const TARGET_BITS: u16 = 16;
const MIX_CHUNK_FRAMES: usize = 512;

#[derive(Clone)]
pub struct AudioDevice {
    pub id: String,
    pub name: String,
}

#[derive(Clone)]
pub struct AudioApp {
    pub pid: u32,
    pub display_name: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SampleFormat {
    I16,
    F32,
}

struct DeviceEnumerator {
    _init: ComGuard,
    inner: IMMDeviceEnumerator,
}

impl DeviceEnumerator {
    fn new() -> Result<Self, String> {
        let init = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;
        let inner: IMMDeviceEnumerator = unsafe {
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
                .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
        };
        Ok(Self { _init: init, inner })
    }
}

pub fn list_input_devices() -> Result<Vec<AudioDevice>, String> {
    list_devices(eCapture)
}

pub fn list_output_devices() -> Result<Vec<AudioDevice>, String> {
    list_devices(eRender)
}

pub fn list_audio_apps() -> Result<Vec<AudioApp>, String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };
    let device = unsafe {
        enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole)
            .map_err(|e| format!("GetDefaultAudioEndpoint(render) failed: {e}"))?
    };
    let session_manager: IAudioSessionManager2 = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("IAudioSessionManager2 activate failed: {e}"))?
    };
    let sessions: IAudioSessionEnumerator = unsafe {
        session_manager
            .GetSessionEnumerator()
            .map_err(|e| format!("GetSessionEnumerator failed: {e}"))?
    };
    let count = unsafe {
        sessions
            .GetCount()
            .map_err(|e| format!("GetCount failed: {e}"))?
    };

    let mut apps = Vec::new();
    for index in 0..count {
        let control: IAudioSessionControl = unsafe {
            sessions
                .GetSession(index)
                .map_err(|e| format!("GetSession({index}) failed: {e}"))?
        };
        let state = unsafe {
            control
                .GetState()
                .map_err(|e| format!("GetState({index}) failed: {e}"))?
        };
        if state != AudioSessionStateActive {
            continue;
        }
        let session_control2: IAudioSessionControl2 = control
            .cast()
            .map_err(|e| format!("IAudioSessionControl2 cast failed: {e}"))?;
        let pid = unsafe {
            session_control2
                .GetProcessId()
                .map_err(|e| format!("GetProcessId({index}) failed: {e}"))?
        };
        if pid == 0 || pid == std::process::id() {
            continue;
        }
        let display_name = process_display_name(pid);
        if !display_name.is_empty() {
            apps.push(AudioApp { pid, display_name });
        }
    }
    apps.sort_by(|a, b| {
        a.display_name
            .to_lowercase()
            .cmp(&b.display_name.to_lowercase())
    });
    apps.dedup_by(|a, b| a.pid == b.pid);
    Ok(apps)
}

pub fn probe_device_with_name(
    device_id: &str,
    device_name: &str,
    loopback: bool,
) -> Result<(), String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;
    let device = resolve_device_with_name(device_id, device_name, loopback)?;
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
                stream_flags | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
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

pub fn probe_process_loopback(process_id: u32) -> Result<(), String> {
    if process_id == 0 {
        return Err("Invalid target process id.".to_string());
    }
    crate::log_debug(&format!(
        "Process loopback probe: preparing audio client for PID {}",
        process_id
    ));
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;
    let client = activate_process_loopback_client(process_id)?;
    let wave_format = process_loopback_wave_format();
    crate::log_debug(&format!(
        "Process loopback probe: initializing shared client for PID {}",
        process_id
    ));
    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                0,
                0,
                &wave_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
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
        let value = id.to_string().unwrap_or_default();
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
        let name = name_ptr.to_string().unwrap_or_default();
        CoTaskMemFree(Some(name_ptr.0 as *const _));
        if name.is_empty() { None } else { Some(name) }
    }
}

#[derive(Clone)]
pub struct RecorderConfig {
    pub include_mic: bool,
    pub mic_device_id: String,
    pub mic_device_name: String,
    pub mic_gain: f32,
    pub include_system: bool,
    pub system_device_id: String,
    pub system_device_name: String,
    pub system_gain: f32,
    pub single_app_process_id: Option<u32>,
    pub selected_app_process_ids: Vec<u32>,
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
        crate::log_if_err!(std::fs::create_dir_all(parent));
    }
    crate::log_if_err!(std::fs::create_dir_all(&output_folder));

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
    *shared.status.lock().unwrap_or_else(|e| e.into_inner()) = RecorderStatus::Recording;
    *shared.started_at.lock().unwrap_or_else(|e| e.into_inner()) = Some(Instant::now());

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
        let device_name = config.mic_device_name.clone();
        let mic_gain = config.mic_gain;
        threads.push(thread::spawn(move || {
            crate::log_debug("Microphone capture thread started");
            let result = capture_source(CaptureOptions {
                kind: SourceKind::Microphone,
                device_id,
                device_name,
                loopback: false,
                gain: mic_gain,
                target_process_id: None,
                buffer,
                shared: shared_state.clone(),
                stop: stop_flag.clone(),
                paused: paused_flag,
            });
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
        let mut target_process_ids = config.selected_app_process_ids.clone();
        if target_process_ids.is_empty() {
            if let Some(pid) = config.single_app_process_id {
                target_process_ids.push(pid);
            }
        } else {
            target_process_ids.sort_unstable();
            target_process_ids.dedup();
        }

        let capture_targets = if target_process_ids.is_empty() {
            vec![None]
        } else {
            target_process_ids.into_iter().map(Some).collect()
        };

        for target_process_id in capture_targets {
            let buffer = mix_buffer.clone();
            let shared_state = shared.clone();
            let stop_flag = stop.clone();
            let paused_flag = paused.clone();
            let device_id = config.system_device_id.clone();
            let device_name = config.system_device_name.clone();
            let system_gain = config.system_gain;
            threads.push(thread::spawn(move || {
                crate::log_debug("System audio capture thread started");
                let result = capture_source(CaptureOptions {
                    kind: SourceKind::System,
                    device_id,
                    device_name,
                    loopback: true,
                    gain: system_gain,
                    target_process_id,
                    buffer,
                    shared: shared_state.clone(),
                    stop: stop_flag.clone(),
                    paused: paused_flag,
                });
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
    }

    let keep_awake_stop = stop.clone();
    threads.push(thread::spawn(move || keep_awake_loop(keep_awake_stop)));

    let writer_buffer = mix_buffer.clone();
    let writer_shared = shared.clone();
    let writer_stop = stop.clone();
    let writer_paused = paused.clone();
    let (writer_path, writer_format) = match config.output_format {
        PodcastFormat::Mp3 => (temp_mp3.clone(), PodcastFormat::Mp3),
        PodcastFormat::Wav => (temp_wav.clone(), PodcastFormat::Wav),
    };
    let writer_bitrate = config.mp3_bitrate;
    let writer_config = WriterConfig {
        path: writer_path,
        format: writer_format,
        mp3_bitrate: writer_bitrate,
    };
    threads.push(thread::spawn(move || {
        let result = write_mixed_audio(
            writer_config,
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
            if let Ok(mut paused_at) = self.shared.paused_at.lock()
                && let Some(start) = paused_at.take()
                && let Ok(mut total) = self.shared.paused_total.lock()
            {
                *total += now.saturating_duration_since(start);
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

        if let Some(cancel) = cancel.as_ref()
            && cancel.load(Ordering::Relaxed)
        {
            crate::log_if_err!(std::fs::remove_file(&self.temp_wav));
            crate::log_if_err!(std::fs::remove_file(&self.temp_mp3));
            return Err("Saving canceled.".to_string());
        }

        if self.format == PodcastFormat::Mp3 {
            if let Some(cancel) = cancel.as_ref()
                && cancel.load(Ordering::Relaxed)
            {
                crate::log_if_err!(std::fs::remove_file(&self.temp_wav));
                crate::log_if_err!(std::fs::remove_file(&self.temp_mp3));
                return Err("Saving canceled.".to_string());
            }
            progress(100);
            if let Err(err) = rename_atomic(&self.temp_mp3, &self.output_path) {
                crate::log_debug(&format!("MP3 final rename failed: {}", err));
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
        crate::log_if_err!(std::fs::remove_file(dest));
    }
    std::fs::rename(src, dest).map_err(|e| e.to_string())
}

fn keep_awake_loop(stop: Arc<AtomicBool>) -> Result<(), String> {
    const KEEP_AWAKE_REFRESH: Duration = Duration::from_secs(30);
    const KEEP_AWAKE_POLL: Duration = Duration::from_millis(200);

    unsafe {
        SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
    }
    let mut last_refresh = Instant::now();
    while !stop.load(Ordering::SeqCst) {
        if last_refresh.elapsed() >= KEEP_AWAKE_REFRESH {
            unsafe {
                SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
            }
            last_refresh = Instant::now();
        }
        thread::sleep(KEEP_AWAKE_POLL);
    }
    unsafe {
        SetThreadExecutionState(ES_CONTINUOUS);
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

struct WriterConfig {
    path: PathBuf,
    format: PodcastFormat,
    mp3_bitrate: u32,
}

fn write_mixed_audio(
    config: WriterConfig,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    match config.format {
        PodcastFormat::Mp3 => write_mixed_audio_mp3(
            config.path,
            config.mp3_bitrate,
            buffer,
            shared,
            stop,
            paused,
        ),
        PodcastFormat::Wav => write_mixed_audio_wav(config.path, buffer, shared, stop, paused),
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

        let mixed = {
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
                crate::log_if_err!(
                    buffer
                        .condvar
                        .wait_timeout(inner, Duration::from_millis(40))
                );
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
            mixed
        };

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

        let mixed = {
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
                crate::log_if_err!(
                    buffer
                        .condvar
                        .wait_timeout(inner, Duration::from_millis(40))
                );
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
            mixed
        };

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

struct CaptureOptions {
    kind: SourceKind,
    device_id: String,
    device_name: String,
    loopback: bool,
    gain: f32,
    target_process_id: Option<u32>,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
}

fn capture_source(options: CaptureOptions) -> Result<(), String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;
    crate::log_debug(&format!(
        "capture_source: kind={:?}, device_id='{}', name='{}', loopback={}",
        match options.kind {
            SourceKind::Microphone => "Microphone",
            SourceKind::System => "System",
        },
        options.device_id,
        options.device_name,
        options.loopback
    ));
    if let Some(target_process_id) = options.target_process_id {
        crate::log_debug(&format!(
            "capture_source: using process loopback for PID {}",
            target_process_id
        ));
    }
    let client: IAudioClient = if let Some(target_process_id) = options.target_process_id {
        activate_process_loopback_client(target_process_id)?
    } else {
        let device =
            resolve_device_with_name(&options.device_id, &options.device_name, options.loopback)?;
        if matches!(options.kind, SourceKind::Microphone) {
            crate::log_debug("Microphone capture: device resolved");
        }
        unsafe {
            device
                .Activate(CLSCTX_ALL, None)
                .map_err(|e| format!("AudioClient activate failed: {e}"))?
        }
    };

    let mut stream_flags = 0;
    if options.loopback {
        stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
    }
    let (input_rate, input_channels, input_format) = if options.target_process_id.is_some() {
        let wave_format = process_loopback_wave_format();
        let parsed = parse_format(&wave_format);
        unsafe {
            client
                .Initialize(
                    AUDCLNT_SHAREMODE_SHARED,
                    stream_flags | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                    0,
                    0,
                    &wave_format,
                    None,
                )
                .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
        }
        parsed
    } else {
        let mix_format = unsafe {
            client
                .GetMixFormat()
                .map_err(|e| format!("GetMixFormat failed: {e}"))?
        };
        let parsed = parse_mix_format_ptr(mix_format)?;
        unsafe {
            client
                .Initialize(
                    AUDCLNT_SHAREMODE_SHARED,
                    stream_flags | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                    10_000_000,
                    0,
                    mix_format,
                    None,
                )
                .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
            CoTaskMemFree(Some(mix_format as *const _));
        }
        parsed
    };
    if matches!(options.kind, SourceKind::Microphone) {
        crate::log_debug("Microphone capture: client initialized");
    }

    let capture: IAudioCaptureClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService capture failed: {e}"))?
    };
    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }
    if matches!(options.kind, SourceKind::Microphone) {
        crate::log_debug("Microphone capture: client started");
    }

    let mut resampler =
        LinearResampler::new(input_rate, TARGET_SAMPLE_RATE, input_channels as usize);

    loop {
        if options.stop.load(Ordering::SeqCst) {
            break;
        }
        if options.paused.load(Ordering::SeqCst) {
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

            update_peak(&options.shared, &options.kind, &samples);
            let resampled = resampler.push(&samples);
            let mut stereo = to_stereo(&resampled, input_channels as usize);

            // Apply gain
            if options.gain != 1.0 {
                for sample in stereo.iter_mut() {
                    *sample = (*sample * options.gain).clamp(-1.0, 1.0);
                }
            }

            options.buffer.push(options.kind, stereo);
            packet_len = unsafe {
                capture
                    .GetNextPacketSize()
                    .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
            };
        }
        thread::sleep(Duration::from_millis(10));
    }

    unsafe {
        crate::log_if_err!(client.Stop());
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
        crate::log_debug(&format!(
            "resolve_device: using default device (loopback={})",
            loopback
        ));
        return unsafe {
            enumerator
                .GetDefaultAudioEndpoint(flow, eConsole)
                .map_err(|e| format!("GetDefaultAudioEndpoint failed: {e}"))
        };
    }

    crate::log_debug(&format!(
        "resolve_device: looking for device_id='{}'",
        device_id
    ));
    let wide = crate::accessibility::to_wide(device_id);
    let result = unsafe {
        enumerator
            .GetDevice(PCWSTR(wide.as_ptr()))
            .map_err(|e| format!("GetDevice({}) failed: {e}", device_id))
    };
    if result.is_ok() {
        crate::log_debug("resolve_device: device found successfully");
    }
    result
}

fn resolve_device_with_name(
    device_id: &str,
    device_name: &str,
    loopback: bool,
) -> Result<IMMDevice, String> {
    let name = device_name.trim();

    // Attempt to resolve by ID first
    match resolve_device(device_id, loopback) {
        Ok(device) => Ok(device),
        Err(err) => {
            if name.is_empty() {
                return Err(err);
            }
            let flow = if loopback { eRender } else { eCapture };
            let enumerator: IMMDeviceEnumerator = unsafe {
                CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
                    .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
            };
            let collection: IMMDeviceCollection = unsafe {
                enumerator
                    .EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE)
                    .map_err(|e| format!("EnumAudioEndpoints failed: {e}"))?
            };
            let count = unsafe {
                collection
                    .GetCount()
                    .map_err(|e| format!("GetCount failed: {e}"))?
            };
            let needle = name.to_lowercase();
            for index in 0..count {
                let device: IMMDevice = unsafe {
                    collection
                        .Item(index)
                        .map_err(|e| format!("Device Item failed: {e}"))?
                };
                if let Some(render_name) = device_friendly_name(&device)
                    && (render_name.to_lowercase().contains(&needle)
                        || needle.contains(&render_name.to_lowercase()))
                {
                    crate::log_debug(&format!(
                        "resolve_device: matched by name '{}'",
                        render_name
                    ));
                    return Ok(device);
                }
            }

            Err(err)
        }
    }
}

fn parse_format(fmt: &WAVEFORMATEX) -> (u32, u16, SampleFormat) {
    let channels = fmt.nChannels;
    let rate = fmt.nSamplesPerSec;
    if channels < 1 {
        crate::log_debug(&format!(
            "Recorder: invalid channel count {} in mix format",
            channels
        ));
    }
    let mut format = match fmt.wFormatTag as u32 {
        WAVE_FORMAT_IEEE_FLOAT => SampleFormat::F32,
        _ => SampleFormat::I16,
    };
    if fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE {
        let ext = crate::wave_format_extensible_ref_safe(fmt);
        let subformat = crate::read_unaligned_safe(std::ptr::addr_of!(ext.SubFormat));
        if subformat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
            format = SampleFormat::F32;
        } else {
            format = SampleFormat::I16;
        }
    }
    (rate, channels, format)
}

fn parse_mix_format_ptr(mix_format: *mut WAVEFORMATEX) -> Result<(u32, u16, SampleFormat), String> {
    crate::with_raw_mut_ptr_safe(mix_format, |fmt| parse_format(fmt))
        .ok_or_else(|| "GetMixFormat returned null pointer".to_string())
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
            if value > 0 && value.is_multiple_of(5000) {
                crate::log_debug(&format!("Recorder mic peak={}", value));
            }
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

pub(crate) fn process_loopback_wave_format() -> WAVEFORMATEX {
    let bits_per_sample = 16u16;
    let block_align = TARGET_CHANNELS * (bits_per_sample / 8);
    WAVEFORMATEX {
        wFormatTag: 1,
        nChannels: TARGET_CHANNELS,
        nSamplesPerSec: TARGET_SAMPLE_RATE,
        nAvgBytesPerSec: TARGET_SAMPLE_RATE * block_align as u32,
        nBlockAlign: block_align,
        wBitsPerSample: bits_per_sample,
        cbSize: 0,
    }
}

pub(crate) fn activate_process_loopback_client(process_id: u32) -> Result<IAudioClient, String> {
    crate::log_debug(&format!(
        "Process loopback activation: requesting async activation for PID {}",
        process_id
    ));
    let params = AUDIOCLIENT_ACTIVATION_PARAMS {
        ActivationType: AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK,
        Anonymous: AUDIOCLIENT_ACTIVATION_PARAMS_0 {
            ProcessLoopbackParams: AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS {
                TargetProcessId: process_id,
                ProcessLoopbackMode: PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
            },
        },
    };

    let mut raw: windows::core::imp::PROPVARIANT = unsafe { std::mem::zeroed() };
    raw.Anonymous.Anonymous.vt = VT_BLOB.0;
    raw.Anonymous.Anonymous.Anonymous.blob = windows::core::imp::BLOB {
        cbSize: size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>() as u32,
        pBlobData: (&params as *const AUDIOCLIENT_ACTIVATION_PARAMS)
            .cast_mut()
            .cast(),
    };
    let prop_variant = ManuallyDrop::new(unsafe { PROPVARIANT::from_raw(raw) });

    let state = Arc::new(ActivationState::default());
    let handler: IActivateAudioInterfaceCompletionHandler =
        ActivateAudioCompletionHandler::new(state.clone()).into();
    let _operation = unsafe {
        ActivateAudioInterfaceAsync(
            VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
            &IAudioClient::IID,
            Some((&*prop_variant) as *const PROPVARIANT),
            &handler,
        )
        .map_err(|e| format!("ActivateAudioInterfaceAsync failed: {e}"))?
    };

    let mut guard = state.result.lock().unwrap_or_else(|e| e.into_inner());
    while guard.is_none() {
        guard = state.condvar.wait(guard).unwrap_or_else(|e| e.into_inner());
    }
    let raw = guard
        .take()
        .ok_or_else(|| "Activation completed without result.".to_string())??;
    crate::log_debug(&format!(
        "Process loopback activation: async activation completed for PID {}",
        process_id
    ));
    unsafe {
        Ok(IAudioClient::from_raw(
            (raw as *mut core::ffi::c_void).cast(),
        ))
    }
}

#[derive(Default)]
struct ActivationState {
    result: Mutex<Option<Result<usize, String>>>,
    condvar: Condvar,
}

#[implement(IActivateAudioInterfaceCompletionHandler)]
struct ActivateAudioCompletionHandler {
    state: Arc<ActivationState>,
}

impl ActivateAudioCompletionHandler {
    fn new(state: Arc<ActivationState>) -> Self {
        Self { state }
    }
}

impl windows::Win32::Media::Audio::IActivateAudioInterfaceCompletionHandler_Impl
    for ActivateAudioCompletionHandler
{
    fn ActivateCompleted(
        &self,
        activateoperation: Option<&IActivateAudioInterfaceAsyncOperation>,
    ) -> windows::core::Result<()> {
        crate::log_debug("Process loopback activation: callback entered");
        let result = match activateoperation {
            Some(operation) => {
                let mut activation_hr = HRESULT(0);
                let mut activated = None;
                unsafe {
                    operation.GetActivateResult(&mut activation_hr, &mut activated)?;
                }
                if let Err(err) = activation_hr.ok() {
                    crate::log_debug(&format!(
                        "Process loopback activation: GetActivateResult returned error {}",
                        err
                    ));
                    Err(format!("GetActivateResult failed: {err}"))
                } else if let Some(activated) = activated {
                    match activated.cast::<IAudioClient>() {
                        Ok(client) => {
                            crate::log_debug("Process loopback activation: received IAudioClient");
                            Ok(client.into_raw() as usize)
                        }
                        Err(err) => {
                            crate::log_debug(&format!(
                                "Process loopback activation: IAudioClient cast failed {}",
                                err
                            ));
                            Err(format!("IAudioClient cast failed: {err}"))
                        }
                    }
                } else {
                    crate::log_debug(
                        "Process loopback activation: callback returned no activated interface",
                    );
                    Err("Activation returned no audio client.".to_string())
                }
            }
            None => {
                crate::log_debug("Process loopback activation: callback received no operation");
                Err("Audio activation callback received no operation.".to_string())
            }
        };

        let mut guard = self.state.result.lock().unwrap_or_else(|e| e.into_inner());
        *guard = Some(result);
        self.state.condvar.notify_one();
        Ok(())
    }
}

fn process_display_name(process_id: u32) -> String {
    let base_name = process_image_name(process_id)
        .and_then(|path| {
            PathBuf::from(path)
                .file_stem()
                .map(|stem| stem.to_string_lossy().to_string())
        })
        .filter(|name| !name.trim().is_empty())
        .unwrap_or_else(|| format!("PID {process_id}"));
    if let Some(window_title) = crate::find_process_window_title(process_id)
        && !window_title.is_empty()
    {
        format!("{base_name} - {window_title} (PID {process_id})")
    } else {
        format!("{base_name} (PID {process_id})")
    }
}

fn process_image_name(process_id: u32) -> Option<String> {
    unsafe {
        let handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, process_id).ok()?;
        let mut buffer = vec![0u16; 1024];
        let mut len = buffer.len() as u32;
        let result = QueryFullProcessImageNameW(
            handle,
            PROCESS_NAME_FORMAT(0),
            PWSTR(buffer.as_mut_ptr()),
            &mut len,
        );
        crate::log_if_err!(CloseHandle(handle));
        if result.is_err() || len == 0 {
            return None;
        }
        Some(
            OsString::from_wide(&buffer[..len as usize])
                .to_string_lossy()
                .to_string(),
        )
    }
}
