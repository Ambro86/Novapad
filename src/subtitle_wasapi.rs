//! Synchronized WASAPI output for subtitle audio playback.
//!
//! This module provides a dedicated WASAPI output stream for subtitle TTS that
//! synchronizes precisely with the main audio clock, enabling sub-millisecond
//! accurate subtitle timing through sample-level scheduling.

use crate::com_guard::ComGuard;
use crate::log_debug;
use std::collections::VecDeque;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::thread;
use windows::Win32::Foundation::{CloseHandle, HANDLE, WAIT_OBJECT_0, WAIT_TIMEOUT};
use windows::Win32::Media::Audio::{
    AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, IAudioClient, IAudioRenderClient,
    IMMDeviceEnumerator, MMDeviceEnumerator, WAVEFORMATEX, WAVEFORMATEXTENSIBLE, eConsole, eRender,
};
use windows::Win32::Media::Multimedia::KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::Win32::System::Performance::QueryPerformanceCounter;
use windows::Win32::System::Threading::{CreateEventW, WaitForSingleObject};
use windows::core::GUID;

const WAVE_FORMAT_PCM: u16 = 0x0001;
const WAVE_FORMAT_IEEE_FLOAT: u16 = 0x0003;
const WAVE_FORMAT_EXTENSIBLE: u16 = 0xFFFE;
const KSDATAFORMAT_SUBTYPE_PCM: GUID = GUID::from_u128(0x00000001_0000_0010_8000_00AA00389B71);

/// Desired buffer duration in 100-nanosecond units (5ms for low latency).
const BUFFER_DURATION_100NS: i64 = 50_000;

#[derive(Clone, Copy, Debug)]
enum WasapiSampleFormat {
    Pcm16,
    Float32,
}

#[derive(Clone, Copy)]
enum MixFormat {
    Standard(WAVEFORMATEX),
    Extensible(WAVEFORMATEXTENSIBLE),
}

/// A scheduled audio chunk to be played at a specific time.
#[derive(Clone)]
pub struct ScheduledChunk {
    /// PCM samples (interleaved f32, resampled to device sample rate)
    pub samples: Arc<[f32]>,
    /// Target playback time in seconds (main audio clock time)
    pub target_secs: f64,
}

enum SubtitleWasapiCommand {
    Schedule(ScheduledChunk),
    Clear,
    Stop,
}

struct HandleGuard(HANDLE);

impl Drop for HandleGuard {
    fn drop(&mut self) {
        if self.0.0 != 0 {
            let _ = unsafe { CloseHandle(self.0) };
        }
    }
}

/// Main audio clock reference for synchronization.
pub struct MainAudioClock {
    position_units: Arc<AtomicU64>,
    clock_freq: Arc<AtomicU64>,
    base_micros: Arc<AtomicU64>,
    last_qpc: Arc<AtomicU64>,
    qpc_freq: Arc<AtomicU64>,
    paused: Arc<AtomicBool>,
}

impl MainAudioClock {
    /// Create from WasapiOutput atomics.
    pub fn new(
        position_units: Arc<AtomicU64>,
        clock_freq: Arc<AtomicU64>,
        base_micros: Arc<AtomicU64>,
        last_qpc: Arc<AtomicU64>,
        qpc_freq: Arc<AtomicU64>,
        paused: Arc<AtomicBool>,
    ) -> Self {
        Self {
            position_units,
            clock_freq,
            base_micros,
            last_qpc,
            qpc_freq,
            paused,
        }
    }

    /// Get current playback position in seconds with QPC interpolation.
    /// This provides sub-millisecond accuracy between WASAPI clock updates.
    pub fn position_secs(&self) -> Option<f64> {
        let clock_freq = self.clock_freq.load(Ordering::Acquire);
        if clock_freq == 0 {
            return None;
        }

        // Read all values with proper ordering
        let base = self.base_micros.load(Ordering::Acquire);
        let units = self.position_units.load(Ordering::Acquire);
        let last_qpc = self.last_qpc.load(Ordering::Acquire);
        let qpc_freq = self.qpc_freq.load(Ordering::Acquire);
        let is_paused = self.paused.load(Ordering::Acquire);

        let mut secs = units as f64 / clock_freq as f64;

        // Interpolate using QPC for sub-period accuracy (only when playing)
        if !is_paused && last_qpc > 0 && qpc_freq > 0 {
            let mut now_qpc: i64 = 0;
            if unsafe { QueryPerformanceCounter(&mut now_qpc) }.is_ok() {
                let delta_qpc = now_qpc.saturating_sub(last_qpc as i64);
                // Clamp interpolation to reasonable range (max 50ms ahead)
                let max_delta = qpc_freq as i64 / 20;
                let clamped_delta = delta_qpc.min(max_delta);
                secs += clamped_delta as f64 / qpc_freq as f64;
            }
        }

        Some((base as f64 / 1_000_000.0) + secs)
    }

    pub fn is_paused(&self) -> bool {
        self.paused.load(Ordering::Acquire)
    }
}

/// Synchronized WASAPI output for subtitle audio.
pub struct SubtitleWasapiOutput {
    tx: std::sync::mpsc::Sender<SubtitleWasapiCommand>,
    device_sample_rate: Arc<AtomicU32>,
    device_channels: Arc<AtomicU32>,
}

impl SubtitleWasapiOutput {
    /// Start the subtitle WASAPI output thread.
    pub fn start(clock: Arc<MainAudioClock>, volume: f32) -> Result<Arc<Self>, String> {
        let (tx, rx) = std::sync::mpsc::channel();
        let device_sample_rate = Arc::new(AtomicU32::new(0));
        let device_channels = Arc::new(AtomicU32::new(0));
        let volume_bits = Arc::new(AtomicU32::new(volume.to_bits()));

        let device_sample_rate_thread = device_sample_rate.clone();
        let device_channels_thread = device_channels.clone();
        let volume_bits_thread = volume_bits.clone();

        thread::spawn(move || {
            let _com = match ComGuard::new_mta() {
                Ok(guard) => guard,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: COM init failed: {}", e));
                    return;
                }
            };

            let enumerator: IMMDeviceEnumerator =
                match unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) } {
                    Ok(e) => e,
                    Err(e) => {
                        log_debug(&format!("SubtitleWASAPI: device enumerator failed: {}", e));
                        return;
                    }
                };

            let device = match unsafe { enumerator.GetDefaultAudioEndpoint(eRender, eConsole) } {
                Ok(d) => d,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: default device failed: {}", e));
                    return;
                }
            };

            let client: IAudioClient = match unsafe { device.Activate(CLSCTX_ALL, None) } {
                Ok(c) => c,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: activate failed: {}", e));
                    return;
                }
            };

            let mix_ptr = match unsafe { client.GetMixFormat() } {
                Ok(ptr) => ptr,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: GetMixFormat failed: {}", e));
                    return;
                }
            };

            let (mix_format, sample_format) = unsafe { parse_mix_format(mix_ptr) };
            unsafe {
                CoTaskMemFree(Some(mix_ptr as *const _));
            }
            let Some((mix_format, sample_format)) = mix_format.zip(sample_format) else {
                log_debug("SubtitleWASAPI: unsupported mix format.");
                return;
            };

            let format_ptr = mix_format_ptr(&mix_format);
            let format_ref = mix_format_ref(&mix_format);

            // Request low-latency buffer
            if let Err(e) = unsafe {
                client.Initialize(
                    AUDCLNT_SHAREMODE_SHARED,
                    AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                    BUFFER_DURATION_100NS,
                    0,
                    format_ptr,
                    None,
                )
            } {
                log_debug(&format!("SubtitleWASAPI: Initialize failed: {}", e));
                return;
            }

            let buffer_frames = match unsafe { client.GetBufferSize() } {
                Ok(f) => f,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: GetBufferSize failed: {}", e));
                    return;
                }
            };

            let event_handle = match unsafe { CreateEventW(None, false, false, None) } {
                Ok(h) => h,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: CreateEvent failed: {}", e));
                    return;
                }
            };
            let _event_guard = HandleGuard(event_handle);

            if let Err(e) = unsafe { client.SetEventHandle(event_handle) } {
                log_debug(&format!("SubtitleWASAPI: SetEventHandle failed: {}", e));
                return;
            }

            // Get stream latency for compensation
            let stream_latency_frames = match unsafe { client.GetStreamLatency() } {
                Ok(latency_100ns) => {
                    let latency_secs = latency_100ns as f64 / 10_000_000.0;
                    let frames = (latency_secs * format_ref.nSamplesPerSec as f64) as u32;
                    log_debug(&format!(
                        "SubtitleWASAPI: stream latency {:.3}ms ({} frames)",
                        latency_secs * 1000.0,
                        frames
                    ));
                    frames
                }
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: GetStreamLatency failed: {}", e));
                    0
                }
            };

            let render: IAudioRenderClient = match unsafe { client.GetService() } {
                Ok(r) => r,
                Err(e) => {
                    log_debug(&format!("SubtitleWASAPI: GetService(Render) failed: {}", e));
                    return;
                }
            };

            let channels = format_ref.nChannels;
            let sample_rate = format_ref.nSamplesPerSec;
            device_sample_rate_thread.store(sample_rate, Ordering::Release);
            device_channels_thread.store(channels as u32, Ordering::Release);

            let bytes_per_sample = match sample_format {
                WasapiSampleFormat::Pcm16 => 2usize,
                WasapiSampleFormat::Float32 => 4usize,
            };

            log_debug(&format!(
                "SubtitleWASAPI: initialized {}ch @ {} Hz, buffer {} frames, total latency {} frames ({:.1}ms)",
                channels,
                sample_rate,
                buffer_frames,
                stream_latency_frames + buffer_frames,
                (stream_latency_frames + buffer_frames) as f64 / sample_rate as f64 * 1000.0
            ));

            if let Err(e) = unsafe { client.Start() } {
                log_debug(&format!("SubtitleWASAPI: Start failed: {}", e));
                return;
            }

            // Chunk with sample-level offset tracking
            struct PlayingChunk {
                samples: Arc<[f32]>,
                read_offset: usize,
                /// Samples of silence to insert before this chunk starts
                silence_samples: usize,
            }

            let mut pending: VecDeque<(ScheduledChunk, bool)> = VecDeque::new(); // (chunk, scheduled)
            let mut active: Option<PlayingChunk> = None;

            loop {
                // Process commands
                while let Ok(cmd) = rx.try_recv() {
                    match cmd {
                        SubtitleWasapiCommand::Schedule(chunk) => {
                            pending.push_back((chunk, false));
                        }
                        SubtitleWasapiCommand::Clear => {
                            pending.clear();
                            active = None;
                        }
                        SubtitleWasapiCommand::Stop => {
                            let _ = unsafe { client.Stop() };
                            return;
                        }
                    }
                }

                let wait_res = unsafe { WaitForSingleObject(event_handle, 100) };
                if wait_res == WAIT_TIMEOUT {
                    continue;
                } else if wait_res != WAIT_OBJECT_0 {
                    log_debug("SubtitleWASAPI: WaitForSingleObject failed.");
                    return;
                }

                let padding = match unsafe { client.GetCurrentPadding() } {
                    Ok(p) => p,
                    Err(e) => {
                        log_debug(&format!("SubtitleWASAPI: GetCurrentPadding failed: {}", e));
                        return;
                    }
                };

                let available = buffer_frames.saturating_sub(padding);
                if available == 0 {
                    continue;
                }

                let buf_ptr = match unsafe { render.GetBuffer(available) } {
                    Ok(ptr) => ptr,
                    Err(e) => {
                        log_debug(&format!("SubtitleWASAPI: GetBuffer failed: {}", e));
                        return;
                    }
                };

                let volume = f32::from_bits(volume_bits_thread.load(Ordering::Relaxed));
                let frame_count = available as usize;
                let sample_count = frame_count * channels as usize;
                let is_paused = clock.is_paused();

                // Get current main audio position
                let current_time = clock.position_secs().unwrap_or(0.0);

                // Calculate when samples we're about to write will actually play
                // (current position + buffer latency)
                let total_latency_frames = padding + stream_latency_frames;
                let latency_secs = total_latency_frames as f64 / sample_rate as f64;
                let playback_time = current_time + latency_secs;

                // Schedule pending chunks with sample-accurate timing
                if active.is_none() && !is_paused {
                    while let Some((chunk, scheduled)) = pending.front_mut() {
                        if !*scheduled {
                            // Calculate how many silence samples we need before this chunk
                            let time_until_target = chunk.target_secs - playback_time;

                            if time_until_target <= 0.0 {
                                // Target time has passed, play immediately
                                active = Some(PlayingChunk {
                                    samples: chunk.samples.clone(),
                                    read_offset: 0,
                                    silence_samples: 0,
                                });
                                pending.pop_front();
                                break;
                            } else if time_until_target < 1.0 {
                                // Within 1 second, schedule with precise silence padding
                                let silence_frames =
                                    (time_until_target * sample_rate as f64) as usize;
                                let silence_samples = silence_frames * channels as usize;
                                active = Some(PlayingChunk {
                                    samples: chunk.samples.clone(),
                                    read_offset: 0,
                                    silence_samples,
                                });
                                pending.pop_front();
                                break;
                            } else {
                                // Too far in the future, wait
                                break;
                            }
                        } else {
                            pending.pop_front();
                        }
                    }
                }

                // Write samples to buffer
                unsafe {
                    let byte_len = sample_count * bytes_per_sample;
                    let out = std::slice::from_raw_parts_mut(buf_ptr, byte_len);

                    // Helper to get next sample
                    let mut get_sample = || -> f32 {
                        if is_paused {
                            return 0.0;
                        }
                        if let Some(ref mut chunk) = active {
                            if chunk.silence_samples > 0 {
                                chunk.silence_samples -= 1;
                                return 0.0;
                            }
                            if chunk.read_offset < chunk.samples.len() {
                                let s = chunk.samples[chunk.read_offset];
                                chunk.read_offset += 1;
                                return s;
                            }
                            active = None;
                        }
                        0.0
                    };

                    match sample_format {
                        WasapiSampleFormat::Pcm16 => {
                            let mut byte_offset = 0usize;
                            for _ in 0..sample_count {
                                let sample = get_sample();
                                let scaled = (sample * volume).clamp(-1.0, 1.0);
                                let val = (scaled * i16::MAX as f32) as i16;
                                let bytes = val.to_le_bytes();
                                out[byte_offset] = bytes[0];
                                out[byte_offset + 1] = bytes[1];
                                byte_offset += 2;
                            }
                        }
                        WasapiSampleFormat::Float32 => {
                            let mut byte_offset = 0usize;
                            for _ in 0..sample_count {
                                let sample = get_sample();
                                let scaled = sample * volume;
                                let bytes = scaled.to_le_bytes();
                                out[byte_offset..byte_offset + 4].copy_from_slice(&bytes);
                                byte_offset += 4;
                            }
                        }
                    }
                }

                if let Err(e) = unsafe { render.ReleaseBuffer(available, 0) } {
                    log_debug(&format!("SubtitleWASAPI: ReleaseBuffer failed: {}", e));
                    return;
                }
            }
        });

        Ok(Arc::new(Self {
            tx,
            device_sample_rate,
            device_channels,
        }))
    }

    /// Schedule a chunk to play at a specific time.
    pub fn schedule(&self, chunk: ScheduledChunk) {
        if let Err(e) = self.tx.send(SubtitleWasapiCommand::Schedule(chunk)) {
            log_debug(&format!("SubtitleWASAPI: schedule failed: {}", e));
        }
    }

    /// Clear all pending and active chunks.
    pub fn clear(&self) {
        if let Err(e) = self.tx.send(SubtitleWasapiCommand::Clear) {
            log_debug(&format!("SubtitleWASAPI: clear failed: {}", e));
        }
    }

    /// Stop the output thread.
    pub fn stop(&self) {
        let _ = self.tx.send(SubtitleWasapiCommand::Stop);
    }

    /// Get the device sample rate.
    pub fn device_sample_rate(&self) -> u32 {
        self.device_sample_rate.load(Ordering::Acquire)
    }

    /// Get the device channel count.
    pub fn device_channels(&self) -> u32 {
        self.device_channels.load(Ordering::Acquire)
    }

    /// Check if the output is ready (device initialized).
    pub fn is_ready(&self) -> bool {
        self.device_sample_rate.load(Ordering::Acquire) > 0
    }
}

unsafe fn parse_mix_format(
    mix_ptr: *mut WAVEFORMATEX,
) -> (Option<MixFormat>, Option<WasapiSampleFormat>) {
    let fmt_ptr = mix_ptr as *const WAVEFORMATEX;
    if fmt_ptr.is_null() {
        return (None, None);
    }
    let fmt = *fmt_ptr;
    if fmt.wFormatTag == WAVE_FORMAT_EXTENSIBLE {
        let ext_ptr = mix_ptr as *const WAVEFORMATEXTENSIBLE;
        if ext_ptr.is_null() {
            return (None, None);
        }
        let ext = *ext_ptr;
        let sub = ext.SubFormat;
        if sub == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
            (
                Some(MixFormat::Extensible(ext)),
                Some(WasapiSampleFormat::Float32),
            )
        } else if sub == KSDATAFORMAT_SUBTYPE_PCM {
            (
                Some(MixFormat::Extensible(ext)),
                Some(WasapiSampleFormat::Pcm16),
            )
        } else {
            (None, None)
        }
    } else if fmt.wFormatTag == WAVE_FORMAT_IEEE_FLOAT {
        (
            Some(MixFormat::Standard(fmt)),
            Some(WasapiSampleFormat::Float32),
        )
    } else if fmt.wFormatTag == WAVE_FORMAT_PCM {
        (
            Some(MixFormat::Standard(fmt)),
            Some(WasapiSampleFormat::Pcm16),
        )
    } else {
        (None, None)
    }
}

fn mix_format_ptr(format: &MixFormat) -> *const WAVEFORMATEX {
    match format {
        MixFormat::Standard(fmt) => fmt as *const WAVEFORMATEX,
        MixFormat::Extensible(fmt) => fmt as *const WAVEFORMATEXTENSIBLE as *const WAVEFORMATEX,
    }
}

fn mix_format_ref(format: &MixFormat) -> &WAVEFORMATEX {
    match format {
        MixFormat::Standard(fmt) => fmt,
        MixFormat::Extensible(fmt) => &fmt.Format,
    }
}

/// Decode MP3 bytes to PCM f32 samples using rodio/symphonia.
pub fn decode_mp3_to_pcm(mp3_data: &[u8]) -> Result<(Vec<f32>, u32, u16), String> {
    use rodio::{Decoder, Source};
    use std::io::Cursor;

    let cursor = Cursor::new(mp3_data.to_vec());
    let decoder = Decoder::new(cursor).map_err(|e| format!("Decoder error: {}", e))?;

    let sample_rate = decoder.sample_rate();
    let channels = decoder.channels();

    let samples: Vec<f32> = decoder.collect();

    Ok((samples, sample_rate, channels))
}

/// Resample PCM data to target sample rate and channels using linear interpolation.
pub fn resample_pcm(
    samples: &[f32],
    source_rate: u32,
    source_channels: u16,
    target_rate: u32,
    target_channels: u16,
) -> Vec<f32> {
    if source_rate == target_rate && source_channels == target_channels {
        return samples.to_vec();
    }

    let src_ch = source_channels as usize;
    let tgt_ch = target_channels as usize;
    let src_frames = samples.len() / src_ch;
    let ratio = source_rate as f64 / target_rate as f64;
    let tgt_frames = (src_frames as f64 / ratio).ceil() as usize;

    let mut output = Vec::with_capacity(tgt_frames * tgt_ch);

    for tgt_frame in 0..tgt_frames {
        let src_pos = tgt_frame as f64 * ratio;
        let src_frame = src_pos.floor() as usize;
        let frac = src_pos.fract() as f32;

        for tgt_c in 0..tgt_ch {
            let src_c = tgt_c % src_ch;

            let sample = if src_frame + 1 < src_frames {
                let s0 = samples[src_frame * src_ch + src_c];
                let s1 = samples[(src_frame + 1) * src_ch + src_c];
                s0 + (s1 - s0) * frac
            } else if src_frame < src_frames {
                samples[src_frame * src_ch + src_c]
            } else {
                0.0
            };

            output.push(sample);
        }
    }

    output
}
