use crate::com_guard::ComGuard;
use crate::log_debug;
use rodio::Source;
use rodio::source::UniformSourceIterator;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc;
use std::thread;
use std::time::Duration;
use windows::Win32::Media::Audio::{
    AUDCLNT_SHAREMODE_SHARED, IAudioClient, IAudioClock, IAudioRenderClient, IMMDeviceEnumerator,
    WAVEFORMATEX, WAVEFORMATEXTENSIBLE,
};
use windows::Win32::Media::Audio::{MMDeviceEnumerator, eConsole, eRender};
use windows::Win32::Media::Multimedia::KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::core::GUID;

const WAVE_FORMAT_PCM: u16 = 0x0001;
const WAVE_FORMAT_IEEE_FLOAT: u16 = 0x0003;
const WAVE_FORMAT_EXTENSIBLE: u16 = 0xFFFE;
const KSDATAFORMAT_SUBTYPE_PCM: GUID = GUID::from_u128(0x00000001_0000_0010_8000_00AA00389B71);

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

enum WasapiCommand {
    Play,
    Pause,
    Stop,
}

pub struct WasapiOutput {
    tx: mpsc::Sender<WasapiCommand>,
    position_micros: Arc<AtomicU64>,
    base_micros: Arc<AtomicU64>,
    latency_micros: Arc<AtomicU64>,
    padding_frames: Arc<AtomicU32>,
    sample_rate_hz: Arc<AtomicU32>,
    volume_bits: Arc<AtomicU32>,
    stopped: Arc<AtomicBool>,
}

impl WasapiOutput {
    pub fn start(
        source: Box<dyn Source<Item = f32> + Send>,
        start_paused: bool,
        volume: f32,
        base_secs: f64,
    ) -> Result<Arc<Self>, String> {
        let (tx, rx) = mpsc::channel();
        let position_micros = Arc::new(AtomicU64::new(0));
        let base_micros = Arc::new(AtomicU64::new((base_secs.max(0.0) * 1_000_000.0) as u64));
        let latency_micros = Arc::new(AtomicU64::new(0));
        let padding_frames = Arc::new(AtomicU32::new(0));
        let sample_rate_hz = Arc::new(AtomicU32::new(0));
        let volume_bits = Arc::new(AtomicU32::new(volume.to_bits()));
        let stopped = Arc::new(AtomicBool::new(false));

        let position_micros_thread = position_micros.clone();
        let latency_micros_thread = latency_micros.clone();
        let padding_frames_thread = padding_frames.clone();
        let sample_rate_hz_thread = sample_rate_hz.clone();
        let volume_bits_thread = volume_bits.clone();
        let stopped_thread = stopped.clone();

        thread::spawn(move || {
            let _com = match ComGuard::new_mta() {
                Ok(guard) => guard,
                Err(e) => {
                    log_debug(&format!("WASAPI: COM init failed: {}", e));
                    return;
                }
            };

            let enumerator: IMMDeviceEnumerator =
                match unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) } {
                    Ok(enumerator) => enumerator,
                    Err(e) => {
                        log_debug(&format!("WASAPI: device enumerator failed: {}", e));
                        return;
                    }
                };

            let device = match unsafe { enumerator.GetDefaultAudioEndpoint(eRender, eConsole) } {
                Ok(device) => device,
                Err(e) => {
                    log_debug(&format!("WASAPI: default device failed: {}", e));
                    return;
                }
            };

            let client: IAudioClient = match unsafe { device.Activate(CLSCTX_ALL, None) } {
                Ok(client) => client,
                Err(e) => {
                    log_debug(&format!("WASAPI: activate audio client failed: {}", e));
                    return;
                }
            };

            let mix_ptr = match unsafe { client.GetMixFormat() } {
                Ok(ptr) => ptr,
                Err(e) => {
                    log_debug(&format!("WASAPI: GetMixFormat failed: {}", e));
                    return;
                }
            };

            let (mix_format, sample_format) = unsafe { parse_mix_format(mix_ptr) };
            unsafe {
                CoTaskMemFree(Some(mix_ptr as *const _));
            }
            let Some((mix_format, sample_format)) = mix_format.zip(sample_format) else {
                log_debug("WASAPI: unsupported mix format.");
                return;
            };

            let buffer_duration = 0i64;
            let format_ptr = mix_format_ptr(&mix_format);
            let format_ref = mix_format_ref(&mix_format);
            if let Err(e) = unsafe {
                client.Initialize(
                    AUDCLNT_SHAREMODE_SHARED,
                    0,
                    buffer_duration,
                    0,
                    format_ptr,
                    None,
                )
            } {
                log_debug(&format!("WASAPI: Initialize failed: {}", e));
                return;
            }

            let buffer_frames = match unsafe { client.GetBufferSize() } {
                Ok(frames) => frames,
                Err(e) => {
                    log_debug(&format!("WASAPI: GetBufferSize failed: {}", e));
                    return;
                }
            };
            let mut default_period: i64 = 0;
            let mut min_period: i64 = 0;
            if let Err(e) =
                unsafe { client.GetDevicePeriod(Some(&mut default_period), Some(&mut min_period)) }
            {
                log_debug(&format!("WASAPI: GetDevicePeriod failed: {}", e));
            }
            match unsafe { client.GetStreamLatency() } {
                Ok(latency) => {
                    let micros = (latency / 10).max(0);
                    latency_micros_thread.store(micros as u64, Ordering::Relaxed);
                    log_debug(&format!(
                        "WASAPI: Stream latency {:.3} ms",
                        micros as f64 / 1000.0
                    ));
                }
                Err(e) => {
                    log_debug(&format!("WASAPI: GetStreamLatency failed: {}", e));
                }
            };

            let render: IAudioRenderClient = match unsafe { client.GetService() } {
                Ok(render) => render,
                Err(e) => {
                    log_debug(&format!("WASAPI: GetService(Render) failed: {}", e));
                    return;
                }
            };

            let clock: IAudioClock = match unsafe { client.GetService() } {
                Ok(clock) => clock,
                Err(e) => {
                    log_debug(&format!("WASAPI: GetService(Clock) failed: {}", e));
                    return;
                }
            };

            let channels = format_ref.nChannels;
            let sample_rate = format_ref.nSamplesPerSec;
            sample_rate_hz_thread.store(sample_rate, Ordering::Relaxed);
            let bytes_per_sample = match sample_format {
                WasapiSampleFormat::Pcm16 => 2usize,
                WasapiSampleFormat::Float32 => 4usize,
            };
            let clock_freq = match unsafe { clock.GetFrequency() } {
                Ok(freq) => freq,
                Err(e) => {
                    log_debug(&format!("WASAPI: GetFrequency failed: {}", e));
                    sample_rate as u64
                }
            };
            if default_period > 0 || min_period > 0 {
                log_debug(&format!(
                    "WASAPI: format {}ch @ {} Hz, buffer {} frames, clock {} Hz, device period {:.3} ms (min {:.3} ms)",
                    channels,
                    sample_rate,
                    buffer_frames,
                    clock_freq,
                    default_period as f64 / 10_000.0,
                    min_period as f64 / 10_000.0
                ));
            } else {
                log_debug(&format!(
                    "WASAPI: format {}ch @ {} Hz, buffer {} frames, clock {} Hz",
                    channels, sample_rate, buffer_frames, clock_freq
                ));
            }
            let mut source = UniformSourceIterator::new(source, channels, sample_rate);
            let mut playing = !start_paused;
            if playing && let Err(e) = unsafe { client.Start() } {
                log_debug(&format!("WASAPI: Start failed: {}", e));
                return;
            }

            loop {
                while let Ok(cmd) = rx.try_recv() {
                    match cmd {
                        WasapiCommand::Play => {
                            if !playing {
                                if let Err(e) = unsafe { client.Start() } {
                                    log_debug(&format!("WASAPI: Start failed: {}", e));
                                } else {
                                    playing = true;
                                }
                            }
                        }
                        WasapiCommand::Pause => {
                            if playing {
                                if let Err(e) = unsafe { client.Stop() } {
                                    log_debug(&format!("WASAPI: Stop failed: {}", e));
                                } else {
                                    playing = false;
                                }
                            }
                        }
                        WasapiCommand::Stop => {
                            stopped_thread.store(true, Ordering::Relaxed);
                            if let Err(e) = unsafe { client.Stop() } {
                                log_debug(&format!("WASAPI: Stop failed: {}", e));
                            }
                            return;
                        }
                    }
                }

                if stopped_thread.load(Ordering::Relaxed) {
                    let _ = unsafe { client.Stop() };
                    return;
                }

                if !playing {
                    thread::sleep(Duration::from_millis(20));
                    continue;
                }

                let padding = match unsafe { client.GetCurrentPadding() } {
                    Ok(padding) => padding,
                    Err(e) => {
                        log_debug(&format!("WASAPI: GetCurrentPadding failed: {}", e));
                        return;
                    }
                };
                padding_frames_thread.store(padding, Ordering::Relaxed);
                let available = buffer_frames.saturating_sub(padding);
                if available == 0 {
                    update_clock(&clock, clock_freq, &position_micros_thread);
                    thread::sleep(Duration::from_millis(5));
                    continue;
                }

                let buf_ptr = match unsafe { render.GetBuffer(available) } {
                    Ok(ptr) => ptr,
                    Err(e) => {
                        log_debug(&format!("WASAPI: GetBuffer failed: {}", e));
                        return;
                    }
                };

                let mut finished = false;
                let volume = f32::from_bits(volume_bits_thread.load(Ordering::Relaxed));
                let sample_count = available as usize * channels as usize;

                unsafe {
                    let byte_len = sample_count * bytes_per_sample;
                    let out = std::slice::from_raw_parts_mut(buf_ptr, byte_len);
                    match sample_format {
                        WasapiSampleFormat::Pcm16 => {
                            let mut offset = 0usize;
                            for _ in 0..sample_count {
                                let sample = match source.next() {
                                    Some(v) => v,
                                    None => {
                                        finished = true;
                                        0.0
                                    }
                                };
                                let scaled = (sample * volume).clamp(-1.0, 1.0);
                                let val = (scaled * i16::MAX as f32) as i16;
                                let bytes = val.to_le_bytes();
                                out[offset] = bytes[0];
                                out[offset + 1] = bytes[1];
                                offset += 2;
                            }
                        }
                        WasapiSampleFormat::Float32 => {
                            let mut offset = 0usize;
                            for _ in 0..sample_count {
                                let sample = match source.next() {
                                    Some(v) => v,
                                    None => {
                                        finished = true;
                                        0.0
                                    }
                                };
                                let scaled = sample * volume;
                                let bytes = scaled.to_le_bytes();
                                out[offset..offset + 4].copy_from_slice(&bytes);
                                offset += 4;
                            }
                        }
                    }
                }

                if let Err(e) = unsafe { render.ReleaseBuffer(available, 0) } {
                    log_debug(&format!("WASAPI: ReleaseBuffer failed: {}", e));
                    return;
                }

                update_clock(&clock, clock_freq, &position_micros_thread);

                if finished {
                    stopped_thread.store(true, Ordering::Relaxed);
                    let _ = unsafe { client.Stop() };
                    return;
                }
            }
        });

        Ok(Arc::new(Self {
            tx,
            position_micros,
            base_micros,
            latency_micros,
            padding_frames,
            sample_rate_hz,
            volume_bits,
            stopped,
        }))
    }

    pub fn play(&self) {
        if let Err(e) = self.tx.send(WasapiCommand::Play) {
            log_debug(&format!("WASAPI: play command failed: {}", e));
        }
    }

    pub fn pause(&self) {
        if let Err(e) = self.tx.send(WasapiCommand::Pause) {
            log_debug(&format!("WASAPI: pause command failed: {}", e));
        }
    }

    pub fn stop(&self) {
        if let Err(e) = self.tx.send(WasapiCommand::Stop) {
            log_debug(&format!("WASAPI: stop command failed: {}", e));
        }
    }

    pub fn set_volume(&self, volume: f32) {
        self.volume_bits.store(volume.to_bits(), Ordering::Relaxed);
    }

    pub fn position_secs(&self) -> Option<f64> {
        if self.stopped.load(Ordering::Relaxed) {
            return None;
        }
        let base = self.base_micros.load(Ordering::Relaxed);
        let micros = self.position_micros.load(Ordering::Relaxed);
        let sample_rate = self.sample_rate_hz.load(Ordering::Relaxed);
        let padding = self.padding_frames.load(Ordering::Relaxed);
        let padding_secs = if sample_rate > 0 {
            padding as f64 / sample_rate as f64
        } else {
            0.0
        };
        Some((((base + micros) as f64 / 1_000_000.0) - padding_secs).max(0.0))
    }

    pub fn latency_secs(&self) -> f64 {
        let micros = self.latency_micros.load(Ordering::Relaxed);
        micros as f64 / 1_000_000.0
    }

    pub fn padding_secs(&self) -> f64 {
        let sample_rate = self.sample_rate_hz.load(Ordering::Relaxed);
        let padding = self.padding_frames.load(Ordering::Relaxed);
        if sample_rate > 0 {
            padding as f64 / sample_rate as f64
        } else {
            0.0
        }
    }

    pub fn sample_rate_hz(&self) -> u32 {
        self.sample_rate_hz.load(Ordering::Relaxed)
    }
}

fn update_clock(clock: &IAudioClock, clock_freq: u64, position_micros: &AtomicU64) {
    let mut position: u64 = 0;
    let mut _qpc: u64 = 0;
    if unsafe { clock.GetPosition(&mut position, Some(&mut _qpc)) }.is_ok() {
        let secs = position as f64 / clock_freq as f64;
        position_micros.store((secs * 1_000_000.0) as u64, Ordering::Relaxed);
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
