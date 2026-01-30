//! Audio monitoring using WASAPI directly (same as podcast recorder)
//! This ensures device IDs match between the UI list and actual capture.

use crate::com_guard::ComGuard;
use crate::settings::PODCAST_DEVICE_DEFAULT;
use std::collections::VecDeque;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::thread::{self, JoinHandle};
use std::time::Duration;
use windows::Win32::Media::Audio::{
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, IAudioCaptureClient, IAudioClient,
    IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, WAVEFORMATEX,
    WAVEFORMATEXTENSIBLE, eCapture, eConsole, eRender,
};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
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
    let stop = Arc::new(AtomicBool::new(false));

    // Shared buffer between capture and playback threads (small for low latency)
    let buffer = Arc::new(Mutex::new(VecDeque::<f32>::with_capacity(4800)));

    // Start capture thread
    let capture_stop = stop.clone();
    let capture_buffer = buffer.clone();
    let capture_device_id = device_id.clone();
    let capture_thread = thread::spawn(move || {
        if let Err(e) = capture_loop(capture_device_id, gain, capture_buffer, capture_stop) {
            crate::log_debug(&format!("Monitor capture error: {}", e));
        }
    });

    // Start playback thread
    let playback_stop = stop.clone();
    let playback_buffer = buffer;
    let playback_thread = thread::spawn(move || {
        if let Err(e) = playback_loop(playback_buffer, playback_stop) {
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

fn resolve_output_device() -> Result<IMMDevice, String> {
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };

    unsafe {
        enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole)
            .map_err(|e| format!("GetDefaultAudioEndpoint(render) failed: {e}"))
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

fn capture_loop(
    device_id: String,
    gain: f32,
    buffer: Arc<Mutex<VecDeque<f32>>>,
    stop: Arc<AtomicBool>,
) -> Result<(), String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;

    let device = resolve_input_device(&device_id)?;
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
    crate::log_debug(&format!(
        "Monitor capture: rate={}, channels={}, format={:?}",
        input_rate,
        input_channels,
        match input_format {
            SampleFormat::F32 => "F32",
            SampleFormat::I16 => "I16",
        }
    ));

    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                0,
                500_000, // 50ms buffer for low latency
                0,
                mix_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
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

        // Convert to mono and apply gain
        let mono = to_mono_f32(&samples, input_channels as usize, gain);

        // Push to shared buffer
        if let Ok(mut q) = buffer.lock() {
            // Prevent buffer from growing too large (keep latency low)
            if q.len() > 4800 {
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

fn playback_loop(buffer: Arc<Mutex<VecDeque<f32>>>, stop: Arc<AtomicBool>) -> Result<(), String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;

    let device = resolve_output_device()?;
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
    let (output_rate, output_channels, output_format) = parse_format(unsafe { &*mix_format });
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
                500_000, // 50ms buffer for low latency
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
