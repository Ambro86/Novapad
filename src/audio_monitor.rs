//! Audio monitoring using WASAPI directly (same as podcast recorder)
//! This ensures device IDs match between the UI list and actual capture.

use crate::com_guard::ComGuard;
use crate::podcast_recorder::{activate_process_loopback_client, process_loopback_wave_format};
use crate::settings::PODCAST_DEVICE_DEFAULT;
use std::collections::VecDeque;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::thread::{self, JoinHandle};
use std::time::Duration;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Media::Audio::{
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
    AUDCLNT_STREAMFLAGS_LOOPBACK, DEVICE_STATE_ACTIVE, IAudioCaptureClient, IAudioClient,
    IAudioRenderClient, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, MMDeviceEnumerator,
    WAVEFORMATEX, eCapture, eConsole, eRender,
};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT};
use windows::Win32::System::Com::StructuredStorage::PropVariantToStringAlloc;
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree, STGM_READ};
use windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore;
use windows::core::PCWSTR;

pub struct MonitorHandle {
    stop: Arc<AtomicBool>,
    capture_threads: Vec<JoinHandle<()>>,
    playback_thread: Option<JoinHandle<()>>,
}

impl MonitorHandle {
    pub fn stop(mut self) {
        self.stop.store(true, Ordering::SeqCst);
        for handle in self.capture_threads.drain(..) {
            crate::log_if_err!(handle.join());
        }
        if let Some(handle) = self.playback_thread.take() {
            crate::log_if_err!(handle.join());
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SampleFormat {
    I16,
    F32,
}

/// Starts audio monitoring with the specified device and gain.
///
/// # Arguments
/// * `device_id` - The Windows MMDevice ID to use (or PODCAST_DEVICE_DEFAULT for default)
/// * `_device_name` - The device name (unused, kept for API compatibility)
/// * `gain` - Volume multiplier (1.0 = normal, 0.5 = half, 2.0 = double)
pub fn start_monitoring(
    device_id: String,
    _device_name: String,
    gain: f32,
) -> Result<MonitorHandle, String> {
    start_monitoring_source(MonitorSource::InputDevice(device_id), gain)
}

pub fn start_system_monitoring(
    device_id: String,
    device_name: String,
    gain: f32,
) -> Result<MonitorHandle, String> {
    start_monitoring_source(MonitorSource::SystemLoopback(device_id, device_name), gain)
}

pub fn start_process_monitoring(process_id: u32, gain: f32) -> Result<MonitorHandle, String> {
    start_monitoring_source(MonitorSource::ProcessLoopback(process_id), gain)
}

pub fn start_processes_monitoring(process_ids: &[u32], gain: f32) -> Result<MonitorHandle, String> {
    let mut sources = Vec::new();
    for process_id in process_ids {
        if *process_id != 0 {
            sources.push(MonitorSource::ProcessLoopback(*process_id));
        }
    }
    if sources.is_empty() {
        return Err("No target processes selected.".to_string());
    }
    start_monitoring_sources(sources, gain)
}

#[derive(Clone)]
enum MonitorSource {
    InputDevice(String),
    SystemLoopback(String, String),
    ProcessLoopback(u32),
}

struct MonitorBuffer {
    queues: Vec<Mutex<VecDeque<f32>>>,
}

impl MonitorBuffer {
    fn new(source_count: usize) -> Self {
        let mut queues = Vec::with_capacity(source_count);
        for _ in 0..source_count {
            queues.push(Mutex::new(VecDeque::with_capacity(24000)));
        }
        Self { queues }
    }

    fn push(&self, source_index: usize, mono: Vec<f32>) {
        let Some(queue) = self.queues.get(source_index) else {
            return;
        };
        if let Ok(mut q) = queue.lock() {
            if q.len() > 24000 {
                q.clear();
            }
            q.extend(mono);
        }
    }

    fn pop_mixed(
        &self,
        frames_to_write: usize,
        process_loopback_only: bool,
        output_rate: u32,
    ) -> Vec<f32> {
        if self.queues.is_empty() {
            return Vec::new();
        }

        if self.queues.len() == 1 {
            let frames_to_pop = if process_loopback_only {
                let ratio = 44_100f32 / output_rate as f32;
                ((frames_to_write as f32 * ratio).ceil() as usize).max(1)
            } else {
                frames_to_write
            };
            if let Ok(mut q) = self.queues[0].lock() {
                let mut mono = Vec::with_capacity(frames_to_pop);
                while mono.len() < frames_to_pop && !q.is_empty() {
                    mono.push(q.pop_front().unwrap_or(0.0));
                }
                return mono;
            }
            return Vec::new();
        }

        let mut locked = Vec::with_capacity(self.queues.len());
        for queue in &self.queues {
            let Ok(guard) = queue.lock() else {
                return Vec::new();
            };
            locked.push(guard);
        }

        let target_frames = if process_loopback_only {
            let ratio = 44_100f32 / output_rate as f32;
            ((frames_to_write as f32 * ratio).ceil() as usize).max(1)
        } else {
            frames_to_write
        };

        let Some(frames_available) = locked.iter().map(|q| q.len()).max() else {
            return Vec::new();
        };
        let frames_to_mix = frames_available.min(target_frames);
        if frames_to_mix == 0 {
            return Vec::new();
        }

        let mut mono = Vec::with_capacity(frames_to_mix);
        for _ in 0..frames_to_mix {
            let mut sum = 0.0f32;
            let mut contributors = 0usize;
            for queue in &mut locked {
                if let Some(sample) = queue.pop_front() {
                    sum += sample;
                    contributors += 1;
                }
            }
            if contributors == 0 {
                mono.push(0.0);
            } else {
                mono.push((sum / contributors as f32).clamp(-1.0, 1.0));
            }
        }
        mono
    }
}

fn start_monitoring_source(source: MonitorSource, gain: f32) -> Result<MonitorHandle, String> {
    start_monitoring_sources(vec![source], gain)
}

fn start_monitoring_sources(
    sources: Vec<MonitorSource>,
    gain: f32,
) -> Result<MonitorHandle, String> {
    let stop = Arc::new(AtomicBool::new(false));
    let process_loopback_only = sources
        .iter()
        .all(|source| matches!(source, MonitorSource::ProcessLoopback(_)));

    // Keep one queue per source; playback mixes them in sync.
    let buffer = Arc::new(MonitorBuffer::new(sources.len()));

    let mut capture_threads = Vec::new();
    for (source_index, source) in sources.into_iter().enumerate() {
        let capture_stop = stop.clone();
        let capture_buffer = buffer.clone();
        capture_threads.push(thread::spawn(move || {
            if let Err(e) = capture_loop(source, source_index, gain, capture_buffer, capture_stop) {
                crate::log_debug(&format!("Monitor capture error: {}", e));
            }
        }));
    }

    // Start playback thread
    let playback_stop = stop.clone();
    let playback_buffer = buffer;
    // Don't prefer any specific device, use system default
    let playback_name = String::new();
    let playback_thread = thread::spawn(move || {
        if let Err(e) = playback_loop(
            playback_buffer,
            playback_stop,
            playback_name,
            process_loopback_only,
        ) {
            crate::log_debug(&format!("Monitor playback error: {}", e));
        }
    });

    Ok(MonitorHandle {
        stop,
        capture_threads,
        playback_thread: Some(playback_thread),
    })
}

fn resolve_input_device(device_id: &str) -> Result<IMMDevice, String> {
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };

    if device_id.is_empty() || device_id == PODCAST_DEVICE_DEFAULT {
        crate::log_debug("Monitor: using default input device");
        return unsafe {
            enumerator
                .GetDefaultAudioEndpoint(eCapture, eConsole)
                .map_err(|e| format!("GetDefaultAudioEndpoint(capture) failed: {e}"))
        };
    }

    crate::log_debug(&format!("Monitor: looking for device_id='{}'", device_id));
    let wide = crate::accessibility::to_wide(device_id);
    let result = unsafe {
        enumerator
            .GetDevice(PCWSTR(wide.as_ptr()))
            .map_err(|e| format!("GetDevice({}) failed: {e}", device_id))
    };
    if result.is_ok() {
        crate::log_debug("Monitor: input device found successfully");
    }
    result
}

fn resolve_output_loopback_device(device_id: &str, device_name: &str) -> Result<IMMDevice, String> {
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };

    if device_id.is_empty() || device_id == PODCAST_DEVICE_DEFAULT {
        crate::log_debug("Monitor: using default output device for loopback");
        return unsafe {
            enumerator
                .GetDefaultAudioEndpoint(eRender, eConsole)
                .map_err(|e| format!("GetDefaultAudioEndpoint(render) failed: {e}"))
        };
    }

    crate::log_debug(&format!(
        "Monitor: looking for output device_id='{}' for loopback",
        device_id
    ));
    let wide = crate::accessibility::to_wide(device_id);
    let by_id = unsafe {
        enumerator
            .GetDevice(PCWSTR(wide.as_ptr()))
            .map_err(|e| format!("GetDevice({}) failed: {e}", device_id))
    };
    if by_id.is_ok() {
        crate::log_debug("Monitor: output loopback device found successfully by id");
        return by_id;
    }

    let needle = device_name.trim().to_lowercase();
    if needle.is_empty() {
        return by_id;
    }

    let collection: IMMDeviceCollection = unsafe {
        enumerator
            .EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)
            .map_err(|e| format!("EnumAudioEndpoints(render) failed: {e}"))?
    };
    let count = unsafe {
        collection
            .GetCount()
            .map_err(|e| format!("GetCount(render) failed: {e}"))?
    };
    for index in 0..count {
        let device: IMMDevice = unsafe {
            collection
                .Item(index)
                .map_err(|e| format!("Device Item(render) failed: {e}"))?
        };
        if let Some(render_name) = device_friendly_name(&device)
            && (render_name.to_lowercase().contains(&needle)
                || needle.contains(&render_name.to_lowercase()))
        {
            crate::log_debug(&format!(
                "Monitor: matched output loopback device by name '{}'",
                render_name
            ));
            return Ok(device);
        }
    }

    by_id
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

fn resolve_output_device(preferred_name: Option<&str>) -> Result<IMMDevice, String> {
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };

    if let Some(name) = preferred_name {
        let needle = name.trim();
        if !needle.is_empty() {
            let needle_lower = needle.to_lowercase();
            let collection: IMMDeviceCollection = unsafe {
                enumerator
                    .EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)
                    .map_err(|e| format!("EnumAudioEndpoints(render) failed: {e}"))?
            };
            let count = unsafe {
                collection
                    .GetCount()
                    .map_err(|e| format!("GetCount(render) failed: {e}"))?
            };
            for index in 0..count {
                let device: IMMDevice = unsafe {
                    collection
                        .Item(index)
                        .map_err(|e| format!("Device Item(render) failed: {e}"))?
                };
                if let Some(render_name) = device_friendly_name(&device)
                    && (render_name.to_lowercase().contains(&needle_lower)
                        || needle_lower.contains(&render_name.to_lowercase()))
                {
                    crate::log_debug(&format!(
                        "Monitor: matched output device by name '{}'",
                        render_name
                    ));
                    return Ok(device);
                }
            }
        }
    }

    unsafe {
        let device = enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole)
            .map_err(|e| format!("GetDefaultAudioEndpoint(render) failed: {e}"))?;
        if let Ok(id) = device.GetId() {
            if id.is_null() {
                crate::log_debug("Monitor: output device id is null");
            } else {
                let value = id.to_string().unwrap_or_default();
                CoTaskMemFree(Some(id.0 as *const _));
                crate::log_debug(&format!("Monitor: output device id='{}'", value));
            }
        }
        Ok(device)
    }
}

fn parse_format(fmt: &WAVEFORMATEX) -> (u32, u16, SampleFormat) {
    let channels = fmt.nChannels;
    let rate = fmt.nSamplesPerSec;
    if channels < 1 {
        crate::log_debug(&format!(
            "Monitor: invalid channel count {} in mix format",
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

fn to_mono_f32(samples: &[f32], channels: usize, gain: f32) -> Vec<f32> {
    let frames = samples.len() / channels;
    let mut out = Vec::with_capacity(frames);
    for frame in 0..frames {
        let base = frame * channels;
        let sum: f32 = samples[base..base + channels].iter().sum();
        let mono = sum / channels as f32;
        out.push((mono * gain).clamp(-1.0, 1.0));
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
        Self {
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

use windows::Win32::Media::Audio::Endpoints::IAudioEndpointVolume;

fn capture_loop(
    source: MonitorSource,
    source_index: usize,
    gain: f32,
    buffer: Arc<MonitorBuffer>,
    stop: Arc<AtomicBool>,
) -> Result<(), String> {
    match &source {
        MonitorSource::InputDevice(device_id) => {
            crate::log_debug(&format!(
                "Monitor capture_loop started for device_id='{}'",
                device_id
            ));
        }
        MonitorSource::SystemLoopback(device_id, device_name) => {
            crate::log_debug(&format!(
                "Monitor capture_loop started for system loopback device_id='{}' name='{}'",
                device_id, device_name
            ));
        }
        MonitorSource::ProcessLoopback(process_id) => {
            crate::log_debug(&format!(
                "Monitor capture_loop started for process_id='{}'",
                process_id
            ));
        }
    }
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;
    let (client, input_rate, input_channels, input_format) = match &source {
        MonitorSource::InputDevice(device_id) => {
            let device = resolve_input_device(device_id)?;

            unsafe {
                if let Ok(endpoint_volume) =
                    device.Activate::<IAudioEndpointVolume>(CLSCTX_ALL, None)
                {
                    let level = endpoint_volume.GetMasterVolumeLevelScalar().unwrap_or(0.0);
                    let mute = endpoint_volume
                        .GetMute()
                        .unwrap_or(windows::Win32::Foundation::BOOL(0));
                    crate::log_debug(&format!(
                        "Monitor: Device System Volume: {:.0}% Muted: {:?}",
                        level * 100.0,
                        mute.as_bool()
                    ));
                }
            }

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
            let parsed = parse_mix_format_ptr(mix_format)?;
            unsafe {
                client
                    .Initialize(
                        AUDCLNT_SHAREMODE_SHARED,
                        AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                        1_000_000,
                        0,
                        mix_format,
                        None,
                    )
                    .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
                CoTaskMemFree(Some(mix_format as *const _));
            }
            (client, parsed.0, parsed.1, parsed.2)
        }
        MonitorSource::SystemLoopback(device_id, device_name) => {
            let device = resolve_output_loopback_device(device_id, device_name)?;
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
            let parsed = parse_mix_format_ptr(mix_format)?;
            unsafe {
                client
                    .Initialize(
                        AUDCLNT_SHAREMODE_SHARED,
                        AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                        1_000_000,
                        0,
                        mix_format,
                        None,
                    )
                    .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
                CoTaskMemFree(Some(mix_format as *const _));
            }
            (client, parsed.0, parsed.1, parsed.2)
        }
        MonitorSource::ProcessLoopback(process_id) => {
            let client = activate_process_loopback_client(*process_id)?;
            let wave_format = process_loopback_wave_format();
            let parsed = parse_format(&wave_format);
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
            (client, parsed.0, parsed.1, parsed.2)
        }
    };
    crate::log_debug(&format!(
        "Monitor capture: rate={}, channels={}, format={:?}",
        input_rate,
        input_channels,
        match input_format {
            SampleFormat::F32 => "F32",
            SampleFormat::I16 => "I16",
        }
    ));

    let capture: IAudioCaptureClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService capture failed: {e}"))?
    };

    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }

    while !stop.load(Ordering::SeqCst) {
        let packet_len = unsafe {
            capture
                .GetNextPacketSize()
                .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
        };

        if packet_len == 0 {
            thread::sleep(Duration::from_millis(5));
            continue;
        }

        let mut data_ptr: *mut u8 = std::ptr::null_mut();
        let mut frames = 0u32;
        let mut flags = 0u32;
        unsafe {
            capture
                .GetBuffer(&mut data_ptr, &mut frames, &mut flags, None, None)
                .map_err(|e| format!("GetBuffer failed: {e}"))?;
        }

        let is_silent = flags & (AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0;

        let samples = if is_silent {
            vec![0f32; frames as usize * input_channels as usize]
        } else {
            read_samples(data_ptr, frames, input_channels, input_format)
        };

        unsafe {
            capture
                .ReleaseBuffer(frames)
                .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
        }

        // Convert to mono and apply gain
        let mono = to_mono_f32(&samples, input_channels as usize, gain);

        buffer.push(source_index, mono);
    }

    unsafe {
        crate::log_if_err!(client.Stop());
    }
    Ok(())
}

fn playback_loop(
    buffer: Arc<MonitorBuffer>,
    stop: Arc<AtomicBool>,
    preferred_output_name: String,
    process_loopback_only: bool,
) -> Result<(), String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;

    let device = resolve_output_device(Some(&preferred_output_name))?;
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
    let (output_rate, output_channels, output_format) = parse_mix_format_ptr(mix_format)?;
    crate::log_debug(&format!(
        "Monitor playback: rate={}, channels={}, format={:?}, process_loopback_only={}",
        output_rate,
        output_channels,
        match output_format {
            SampleFormat::F32 => "F32",
            SampleFormat::I16 => "I16",
        },
        process_loopback_only
    ));

    let mut process_resampler = if process_loopback_only && output_rate != 44_100 {
        Some(LinearResampler::new(44_100, output_rate, 1))
    } else {
        None
    };

    let buffer_size = unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                0,
                1_000_000, // 100ms buffer
                0,
                mix_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;

        client
            .GetBufferSize()
            .map_err(|e| format!("GetBufferSize failed: {e}"))?
    };

    let render: IAudioRenderClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService render failed: {e}"))?
    };

    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }

    while !stop.load(Ordering::SeqCst) {
        let padding = unsafe {
            client
                .GetCurrentPadding()
                .map_err(|e| format!("GetCurrentPadding failed: {e}"))?
        };

        let frames_available = buffer_size - padding;
        if frames_available == 0 {
            thread::sleep(Duration::from_millis(5));
            continue;
        }

        let frames_to_write = frames_available.min(1024); // Write in small chunks

        let data_ptr = unsafe {
            render
                .GetBuffer(frames_to_write)
                .map_err(|e| format!("GetBuffer failed: {e}"))?
        };

        // Get samples from shared buffer
        let mut output_samples =
            Vec::with_capacity(frames_to_write as usize * output_channels as usize);

        let mono_samples =
            buffer.pop_mixed(frames_to_write as usize, process_loopback_only, output_rate);

        let playback_mono = if let Some(resampler) = process_resampler.as_mut() {
            resampler.push(&mono_samples)
        } else {
            mono_samples
        };

        let mono_frames_to_write = playback_mono.len().min(frames_to_write as usize);
        for mono in playback_mono.into_iter().take(mono_frames_to_write) {
            for _ in 0..output_channels {
                output_samples.push(mono);
            }
        }

        // Fill remaining with silence
        output_samples.resize(frames_to_write as usize * output_channels as usize, 0.0);

        // Write to output buffer
        unsafe {
            match output_format {
                SampleFormat::F32 => {
                    let out_slice =
                        std::slice::from_raw_parts_mut(data_ptr as *mut f32, output_samples.len());
                    out_slice.copy_from_slice(&output_samples);
                }
                SampleFormat::I16 => {
                    let out_slice =
                        std::slice::from_raw_parts_mut(data_ptr as *mut i16, output_samples.len());
                    for (i, &sample) in output_samples.iter().enumerate() {
                        out_slice[i] = (sample.clamp(-1.0, 1.0) * i16::MAX as f32) as i16;
                    }
                }
            }

            render
                .ReleaseBuffer(frames_to_write, 0)
                .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
        }

        thread::sleep(Duration::from_millis(5));
    }

    unsafe {
        crate::log_if_err!(client.Stop());
        CoTaskMemFree(Some(mix_format as *const _));
    }
    Ok(())
}
