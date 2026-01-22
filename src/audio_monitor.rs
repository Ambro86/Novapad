use crate::com_guard::ComGuard;
use crate::settings::PODCAST_DEVICE_DEFAULT;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread::{self, JoinHandle};
use std::time::Duration;
use windows::Win32::Media::Audio::{
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, IAudioCaptureClient, IAudioClient,
    IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, ISimpleAudioVolume, MMDeviceEnumerator,
    WAVEFORMATEX, WAVEFORMATEXTENSIBLE, eCapture, eMultimedia, eRender,
};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::core::PCWSTR;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum SampleFormat {
    I16,
    F32,
}

pub struct MonitorHandle {
    stop: Arc<AtomicBool>,
    thread: Option<JoinHandle<()>>,
}

impl MonitorHandle {
    pub fn stop(mut self) {
        self.stop.store(true, Ordering::SeqCst);
        if let Some(t) = self.thread.take() {
            let _ = t.join();
        }
    }
}

pub fn start_monitoring(device_id: String) -> Result<MonitorHandle, String> {
    let stop = Arc::new(AtomicBool::new(false));
    let stop_clone = stop.clone();

    let thread = thread::spawn(move || {
        if let Err(e) = monitor_loop(device_id, stop_clone) {
            crate::log_debug(&format!("Monitor loop error: {}", e));
        }
    });

    Ok(MonitorHandle {
        stop,
        thread: Some(thread),
    })
}

fn monitor_loop(device_id: String, stop: Arc<AtomicBool>) -> Result<(), String> {
    let _com = ComGuard::new_mta().map_err(|e| format!("CoInitializeEx failed: {e}"))?;

    // 1. Setup Render (Output) first to know target format
    let enumerator: IMMDeviceEnumerator = unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
    };

    let render_device: IMMDevice = unsafe {
        enumerator
            .GetDefaultAudioEndpoint(eRender, eMultimedia)
            .map_err(|e| format!("GetDefaultAudioEndpoint failed: {e}"))?
    };

    let render_client: IAudioClient = unsafe {
        render_device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("Render Activate failed: {e}"))?
    };

    let render_format_ptr = unsafe {
        render_client
            .GetMixFormat()
            .map_err(|e| format!("Render GetMixFormat failed: {e}"))?
    };
    let render_format = unsafe { *render_format_ptr };
    let (out_rate, out_channels, out_fmt) = parse_format(&render_format);

    unsafe {
        render_client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                0, // No loopback flag for rendering
                10_000_000,
                0,
                render_format_ptr,
                None,
            )
            .map_err(|e| format!("Render Initialize failed: {e}"))?;
    }

    let render_service: IAudioRenderClient = unsafe {
        render_client
            .GetService()
            .map_err(|e| format!("Render GetService failed: {e}"))?
    };

    let volume_control: ISimpleAudioVolume = unsafe {
        render_client
            .GetService()
            .map_err(|e| format!("Render GetService Volume failed: {e}"))?
    };
    unsafe {
        volume_control
            .SetMasterVolume(1.0, std::ptr::null())
            .map_err(|e| format!("SetMasterVolume failed: {e}"))?;
        volume_control
            .SetMute(false, std::ptr::null())
            .map_err(|e| format!("SetMute failed: {e}"))?;
    }

    // Pre-roll silence
    let buffer_size = unsafe {
        render_client
            .GetBufferSize()
            .map_err(|e| format!("GetBufferSize failed: {e}"))?
    };
    let pre_roll = buffer_size / 2;
    unsafe {
        let buffer_ptr = render_service
            .GetBuffer(pre_roll)
            .map_err(|e| format!("Pre-roll GetBuffer failed: {e}"))?;
        let frame_size = (out_channels as u32)
            * if matches!(out_fmt, SampleFormat::F32) {
                4
            } else {
                2
            };
        std::ptr::write_bytes(buffer_ptr, 0, (pre_roll * frame_size) as usize);
        render_service
            .ReleaseBuffer(pre_roll, AUDCLNT_BUFFERFLAGS_SILENT.0 as u32)
            .map_err(|e| format!("Pre-roll ReleaseBuffer failed: {e}"))?;
    }

    // 2. Setup Capture (Input)
    let capture_device = resolve_device(&enumerator, &device_id)?;
    let capture_client_interface: IAudioClient = unsafe {
        capture_device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("Capture Activate failed: {e}"))?
    };

    let capture_format_ptr = unsafe {
        capture_client_interface
            .GetMixFormat()
            .map_err(|e| format!("Capture GetMixFormat failed: {e}"))?
    };
    let capture_format = unsafe { *capture_format_ptr };
    let (in_rate, in_channels, in_fmt) = parse_format(&capture_format);

    crate::log_debug(&format!(
        "Monitor: Capture {} Hz {} ch {:?}, Render {} Hz {} ch {:?}",
        in_rate, in_channels, in_fmt, out_rate, out_channels, out_fmt
    ));

    unsafe {
        capture_client_interface
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                0,
                10_000_000,
                0,
                capture_format_ptr,
                None,
            )
            .map_err(|e| format!("Capture Initialize failed: {e}"))?;
    }

    let capture_service: IAudioCaptureClient = unsafe {
        capture_client_interface
            .GetService()
            .map_err(|e| format!("Capture GetService failed: {e}"))?
    };

    // 3. Start
    unsafe {
        render_client
            .Start()
            .map_err(|e| format!("Render Start failed: {e}"))?;
        capture_client_interface
            .Start()
            .map_err(|e| format!("Capture Start failed: {e}"))?;
    }

    let mut resampler = LinearResampler::new(in_rate, out_rate, in_channels as usize);
    let mut first_packet = true;
    let mut debug_counter = 0;
    let mut silence_counter = 0;

    let mut debug_wav = crate::audio_utils::WavWriter::create(
        std::path::Path::new("monitor_debug.wav"),
        out_rate,
        out_channels,
        16,
    )
    .ok();

    // Main loop
    while !stop.load(Ordering::Relaxed) {
        let mut packet_len = unsafe {
            capture_service
                .GetNextPacketSize()
                .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
        };

        while packet_len > 0 {
            if first_packet {
                crate::log_debug("Monitor: First packet received");
                first_packet = false;
            }
            let mut data_ptr: *mut u8 = std::ptr::null_mut();
            let mut frames_available = 0u32;
            let mut flags = 0u32;

            unsafe {
                capture_service
                    .GetBuffer(&mut data_ptr, &mut frames_available, &mut flags, None, None)
                    .map_err(|e| format!("GetBuffer failed: {e}"))?;
            }

            let samples = if flags & (AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0 {
                silence_counter += 1;
                if silence_counter % 100 == 0 {
                    crate::log_debug("Monitor: Silence detected");
                }
                vec![0f32; frames_available as usize * in_channels as usize]
            } else {
                read_samples(data_ptr, frames_available, in_channels, in_fmt)
            };

            unsafe {
                capture_service
                    .ReleaseBuffer(frames_available)
                    .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
            }

            // Resample
            let resampled = if in_rate == out_rate {
                samples
            } else {
                resampler.push(&samples)
            };

            // Channel mix (naive) if needed
            let output_samples = if in_channels == out_channels {
                resampled
            } else {
                mix_channels(&resampled, in_channels as usize, out_channels as usize)
            };

            if let Some(w) = debug_wav.as_mut() {
                let _ = w.write_samples_f32(&output_samples);
            }

            // Write to Render
            let frames_to_write = output_samples.len() / out_channels as usize;
            if frames_to_write > 0 {
                debug_counter += 1;
                if debug_counter % 50 == 0 {
                    crate::log_debug(&format!("Monitor: Writing {} frames", frames_to_write));
                }
                unsafe {
                    // Check padding
                    let padding = render_client.GetCurrentPadding().unwrap_or(0);
                    let buffer_size = render_client.GetBufferSize().unwrap_or(0);
                    let available = buffer_size.saturating_sub(padding);

                    if available >= frames_to_write as u32 {
                        let buffer_ptr = render_service
                            .GetBuffer(frames_to_write as u32)
                            .map_err(|e| format!("Render GetBuffer failed: {e}"))?;

                        write_samples(buffer_ptr, &output_samples, out_fmt);

                        render_service
                            .ReleaseBuffer(frames_to_write as u32, 0)
                            .map_err(|e| format!("Render ReleaseBuffer failed: {e}"))?;
                    } else {
                        crate::log_debug(&format!(
                            "Monitor: Dropping {} frames (available {})",
                            frames_to_write, available
                        ));
                    }
                }
            }

            packet_len = unsafe {
                capture_service
                    .GetNextPacketSize()
                    .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
            };
        }

        thread::sleep(Duration::from_millis(5));
    }

    unsafe {
        let _ = render_client.Stop();
        let _ = capture_client_interface.Stop();
        CoTaskMemFree(Some(render_format_ptr as *const _));
        CoTaskMemFree(Some(capture_format_ptr as *const _));
    }

    if let Some(mut w) = debug_wav {
        let _ = w.finalize();
    }

    Ok(())
}

fn resolve_device(enumerator: &IMMDeviceEnumerator, device_id: &str) -> Result<IMMDevice, String> {
    if device_id.is_empty() || device_id == PODCAST_DEVICE_DEFAULT {
        unsafe {
            enumerator
                .GetDefaultAudioEndpoint(eCapture, eMultimedia)
                .map_err(|e| format!("GetDefaultAudioEndpoint failed: {e}"))
        }
    } else {
        let wide = crate::accessibility::to_wide(device_id);
        unsafe {
            enumerator
                .GetDevice(PCWSTR(wide.as_ptr()))
                .map_err(|e| format!("GetDevice failed: {e}"))
        }
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

fn write_samples(ptr: *mut u8, samples: &[f32], format: SampleFormat) {
    unsafe {
        match format {
            SampleFormat::F32 => {
                let slice = std::slice::from_raw_parts_mut(ptr as *mut f32, samples.len());
                slice.copy_from_slice(samples);
            }
            SampleFormat::I16 => {
                let slice = std::slice::from_raw_parts_mut(ptr as *mut i16, samples.len());
                for (i, sample) in samples.iter().enumerate() {
                    let v = (sample.clamp(-1.0, 1.0) * i16::MAX as f32) as i16;
                    slice[i] = v;
                }
            }
        }
    }
}

fn mix_channels(samples: &[f32], in_channels: usize, out_channels: usize) -> Vec<f32> {
    if in_channels == out_channels {
        return samples.to_vec();
    }
    let frames = samples.len() / in_channels;
    let mut out = Vec::with_capacity(frames * out_channels);
    for i in 0..frames {
        let base = i * in_channels;
        // Naive mix: take first channel or average
        let mono = if in_channels >= 2 {
            (samples[base] + samples[base + 1]) * 0.5
        } else {
            samples[base]
        };

        for _ in 0..out_channels {
            out.push(mono);
        }
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
