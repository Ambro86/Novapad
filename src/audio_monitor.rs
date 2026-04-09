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
    capture_thread: Option<JoinHandle<()>>,
    playback_thread: Option<JoinHandle<()>>,
}

impl MonitorHandle {
    pub fn stop(mut self) {
        self.stop.store(true, Ordering::SeqCst);
        if let Some(handle) = self.capture_thread.take() {
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

pub fn start_process_monitoring(process_id: u32, gain: f32) -> Result<MonitorHandle, String> {
    start_monitoring_source(MonitorSource::ProcessLoopback(process_id), gain)
}

#[derive(Clone)]
enum MonitorSource {
    InputDevice(String),
    ProcessLoopback(u32),
}

fn start_monitoring_source(source: MonitorSource, gain: f32) -> Result<MonitorHandle, String> {
    let stop = Arc::new(AtomicBool::new(false));

    // Shared buffer between capture and playback threads (small for low latency)
    let buffer = Arc::new(Mutex::new(VecDeque::<f32>::with_capacity(24000)));

    // Start capture thread
    let capture_stop = stop.clone();
    let capture_buffer = buffer.clone();
    let capture_thread = thread::spawn(move || {
        if let Err(e) = capture_loop(source, gain, capture_buffer, capture_stop) {
            crate::log_debug(&format!("Monitor capture error: {}", e));
        }
    });

    // Start playback thread
    let playback_stop = stop.clone();
    let playback_buffer = buffer;
    // Don't prefer any specific device, use system default
    let playback_name = String::new();
    let playback_thread = thread::spawn(move || {
        if let Err(e) = playback_loop(playback_buffer, playback_stop, playback_name) {
            crate::log_debug(&format!("Monitor playback error: {}", e));
        }
    });

    Ok(MonitorHandle {
        stop,
        capture_thread: Some(capture_thread),
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

use windows::Win32::Media::Audio::Endpoints::IAudioEndpointVolume;

fn capture_loop(
    source: MonitorSource,
    gain: f32,
    buffer: Arc<Mutex<VecDeque<f32>>>,
    stop: Arc<AtomicBool>,
) -> Result<(), String> {
    match &source {
        MonitorSource::InputDevice(device_id) => {
            crate::log_debug(&format!(
                "Monitor capture_loop started for device_id='{}'",
                device_id
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

        // Push to shared buffer
        if let Ok(mut q) = buffer.lock() {
            // Prevent buffer from growing too large (keep latency low)
            if q.len() > 24000 {
                q.clear(); // Drop all old samples to reset latency
            }
            q.extend(mono);
        }
    }

    unsafe {
        crate::log_if_err!(client.Stop());
    }
    Ok(())
}

fn playback_loop(
    buffer: Arc<Mutex<VecDeque<f32>>>,
    stop: Arc<AtomicBool>,
    preferred_output_name: String,
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
        "Monitor playback: rate={}, channels={}, format={:?}",
        output_rate,
        output_channels,
        match output_format {
            SampleFormat::F32 => "F32",
            SampleFormat::I16 => "I16",
        }
    ));

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
        let mut samples_needed = frames_to_write as usize * output_channels as usize;
        let mut output_samples = Vec::with_capacity(samples_needed);

        if let Ok(mut q) = buffer.lock() {
            while samples_needed > 0 && !q.is_empty() {
                let mono = q.pop_front().unwrap_or(0.0);
                // Expand mono to all output channels
                for _ in 0..output_channels {
                    output_samples.push(mono);
                    samples_needed = samples_needed.saturating_sub(1);
                }
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
