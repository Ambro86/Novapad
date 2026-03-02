use crate::com_guard::ComGuard;
use crate::log_debug;
use rodio::Source;
use rodio::source::UniformSourceIterator;
use std::collections::VecDeque;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicI64, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc;
use std::thread;
use std::time::Duration;
use windows::Win32::Foundation::{CloseHandle, HANDLE, WAIT_OBJECT_0, WAIT_TIMEOUT};
use windows::Win32::Media::Audio::{
    AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, IAudioClient, IAudioClock,
    IAudioRenderClient, IMMDeviceEnumerator, WAVEFORMATEX, WAVEFORMATEXTENSIBLE,
};
use windows::Win32::Media::Audio::{MMDeviceEnumerator, eConsole, eRender};
use windows::Win32::Media::Multimedia::KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::Win32::System::Performance::{QueryPerformanceCounter, QueryPerformanceFrequency};
use windows::Win32::System::Threading::{CreateEventW, WaitForSingleObject};
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

/// A subtitle audio chunk scheduled at an exact sample position.
#[derive(Clone)]
pub struct ScheduledSubtitle {
    /// PCM samples (interleaved f32, must match device sample rate and channels)
    pub samples: Arc<[f32]>,
    /// Target playback time in seconds (absolute position in media)
    pub target_secs: f64,
    /// Volume multiplier for this subtitle (0.0 to 1.0)
    pub volume: f32,
}

enum WasapiCommand {
    Play,
    Pause,
    Stop,
    ScheduleSubtitle(ScheduledSubtitle),
    ClearSubtitles,
}

struct HandleGuard(HANDLE);

impl Drop for HandleGuard {
    fn drop(&mut self) {
        if self.0.0 != 0
            && let Err(e) = unsafe { CloseHandle(self.0) }
        {
            log_debug(&format!("WASAPI: CloseHandle failed: {}", e));
        }
    }
}

pub struct WasapiOutput {
    tx: mpsc::Sender<WasapiCommand>,
    position_units: Arc<AtomicU64>,
    last_qpc: Arc<AtomicU64>,
    qpc_freq: Arc<AtomicU64>,
    clock_freq: Arc<AtomicU64>,
    base_micros: Arc<AtomicU64>,
    padding_frames: Arc<AtomicU32>,
    sample_rate_hz: Arc<AtomicU32>,
    device_channels: Arc<AtomicU32>,
    volume_bits: Arc<AtomicU32>,
    stopped: Arc<AtomicBool>,
    last_written_end_pts_us: Arc<AtomicI64>,
}

impl WasapiOutput {
    pub fn start(
        source: Box<dyn Source<Item = f32> + Send>,
        start_paused: bool,
        volume: f32,
        base_secs: f64,
        source_pts_us: Option<Arc<AtomicI64>>,
    ) -> Result<Arc<Self>, String> {
        let (tx, rx) = mpsc::channel();
        let position_units = Arc::new(AtomicU64::new(0));
        let last_qpc = Arc::new(AtomicU64::new(0));
        let qpc_freq = Arc::new(AtomicU64::new(0));
        let clock_freq = Arc::new(AtomicU64::new(0));
        let base_micros = Arc::new(AtomicU64::new((base_secs.max(0.0) * 1_000_000.0) as u64));
        let padding_frames = Arc::new(AtomicU32::new(0));
        let sample_rate_hz = Arc::new(AtomicU32::new(0));
        let device_channels = Arc::new(AtomicU32::new(0));
        let volume_bits = Arc::new(AtomicU32::new(volume.to_bits()));
        let stopped = Arc::new(AtomicBool::new(false));
        let last_written_end_pts_us = Arc::new(AtomicI64::new(0));

        let position_units_thread = position_units.clone();
        let last_qpc_thread = last_qpc.clone();
        let qpc_freq_thread = qpc_freq.clone();
        let clock_freq_thread = clock_freq.clone();
        let base_micros_thread = base_micros.clone();
        let pts_offset_us_thread = Arc::new(AtomicI64::new(i64::MIN));
        let padding_frames_thread = padding_frames.clone();
        let sample_rate_hz_thread = sample_rate_hz.clone();
        let device_channels_thread = device_channels.clone();
        let volume_bits_thread = volume_bits.clone();
        let stopped_thread = stopped.clone();
        let last_written_end_pts_us_thread = last_written_end_pts_us.clone();
        let paused_thread = Arc::new(AtomicBool::new(start_paused));
        let source_pts_us_thread = source_pts_us.clone();

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
                    AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
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
            let event_handle = match unsafe { CreateEventW(None, false, false, None) } {
                Ok(handle) => handle,
                Err(e) => {
                    log_debug(&format!("WASAPI: CreateEvent failed: {}", e));
                    return;
                }
            };
            let _event_guard = HandleGuard(event_handle);
            if let Err(e) = unsafe { client.SetEventHandle(event_handle) } {
                log_debug(&format!("WASAPI: SetEventHandle failed: {}", e));
                return;
            }
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
            device_channels_thread.store(channels as u32, Ordering::Relaxed);
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
            clock_freq_thread.store(clock_freq, Ordering::Relaxed);
            let mut qpc_freq_val: i64 = 0;
            if unsafe { QueryPerformanceFrequency(&mut qpc_freq_val) }.is_ok() {
                qpc_freq_thread.store(qpc_freq_val as u64, Ordering::Relaxed);
            } else {
                log_debug("WASAPI: QueryPerformanceFrequency failed.");
            }
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

            // Subtitle mixing state
            struct ActiveSubtitle {
                samples: Arc<[f32]>,
                read_offset: usize,
                volume: f32,
            }
            let mut pending_subtitles: VecDeque<ScheduledSubtitle> = VecDeque::new();
            let mut active_subtitle: Option<ActiveSubtitle> = None;
            // Track position for subtitle timing
            // We track samples written to the buffer, which represents when the audio
            // will actually be heard (accounting for buffer latency)
            let base_samples: u64 = (base_secs * sample_rate as f64 * channels as f64) as u64;
            let mut samples_written_since_start: u64 = 0;

            loop {
                while let Ok(cmd) = rx.try_recv() {
                    match cmd {
                        WasapiCommand::Play => {
                            if !playing {
                                if let Err(e) = unsafe { client.Start() } {
                                    log_debug(&format!("WASAPI: Start failed: {}", e));
                                } else {
                                    playing = true;
                                    paused_thread.store(false, Ordering::Relaxed);
                                }
                            }
                        }
                        WasapiCommand::Pause => {
                            if playing {
                                if let Err(e) = unsafe { client.Stop() } {
                                    log_debug(&format!("WASAPI: Stop failed: {}", e));
                                } else {
                                    playing = false;
                                    paused_thread.store(true, Ordering::Relaxed);
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
                        WasapiCommand::ScheduleSubtitle(sub) => {
                            pending_subtitles.push_back(sub);
                        }
                        WasapiCommand::ClearSubtitles => {
                            pending_subtitles.clear();
                            active_subtitle = None;
                        }
                    }
                }

                if stopped_thread.load(Ordering::Relaxed) {
        // Stop and cleanup
        unsafe {
            if let Err(e) = client.Stop() {
                crate::log_debug(&format!("WASAPI: client.Stop failed: {:?}", e));
            }
        }
                    return;
                }

                if !playing {
                    thread::sleep(Duration::from_millis(20));
                    continue;
                }

                let wait_res = unsafe { WaitForSingleObject(event_handle, 2000) };
                if wait_res == WAIT_TIMEOUT {
                    update_clock(&clock, &position_units_thread, &last_qpc_thread);
                    continue;
                } else if wait_res != WAIT_OBJECT_0 {
                    log_debug("WASAPI: WaitForSingleObject failed.");
                    return;
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
                    update_clock(&clock, &position_units_thread, &last_qpc_thread);
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
                let mut buffer_start_pts_us: Option<i64> = None;

                // Helper to get subtitle sample for current position
                // `written_in_buffer` is samples written so far in this buffer fill
                let get_subtitle_sample = |sample_idx: usize,
                                           pending: &mut VecDeque<ScheduledSubtitle>,
                                           active: &mut Option<ActiveSubtitle>,
                                           base: u64,
                                           written_since_start: u64|
                 -> f32 {
                    // Current absolute sample position (when this sample will be heard)
                    // = base position + samples written since start + current index in buffer
                    let current_sample = base + written_since_start + sample_idx as u64;

                    // Check if we should start a pending subtitle
                    if active.is_none()
                        && let Some(sub) = pending.front()
                    {
                        let offset_us = pts_offset_us_thread.load(Ordering::Acquire);
                        let offset_secs = if offset_us == i64::MIN {
                            0.0
                        } else {
                            offset_us as f64 / 1_000_000.0
                        };
                        let effective_target_secs = (sub.target_secs - offset_secs).max(0.0);
                        let target_sample =
                            (effective_target_secs * sample_rate as f64 * channels as f64) as u64;
                        if current_sample >= target_sample {
                            // Time to start this subtitle
                            if let Some(sub) = pending.pop_front() {
                                *active = Some(ActiveSubtitle {
                                    samples: sub.samples,
                                    read_offset: 0,
                                    volume: sub.volume,
                                });
                            }
                        }
                        // Otherwise wait until the target time
                    }

                    // Mix active subtitle if any
                    if let Some(sub) = active {
                        if sub.read_offset < sub.samples.len() {
                            let s = sub.samples[sub.read_offset] * sub.volume;
                            sub.read_offset += 1;
                            return s;
                        } else {
                            // Subtitle finished
                            *active = None;
                        }
                    }
                    0.0
                };

                unsafe {
                    let byte_len = sample_count * bytes_per_sample;
                    let out = std::slice::from_raw_parts_mut(buf_ptr, byte_len);
                    match sample_format {
                        WasapiSampleFormat::Pcm16 => {
                            let mut offset = 0usize;
                            for i in 0..sample_count {
                                let audio_sample = match source.next() {
                                    Some(v) => v,
                                    None => {
                                        finished = true;
                                        0.0
                                    }
                                };
                                if i == 0
                                    && buffer_start_pts_us.is_none()
                                    && let Some(ref pts_clock) = source_pts_us_thread
                                {
                                    let start_pts = pts_clock.load(Ordering::Acquire);
                                    buffer_start_pts_us = Some(start_pts);
                                    if pts_offset_us_thread.load(Ordering::Acquire) == i64::MIN {
                                        let base =
                                            base_micros_thread.load(Ordering::Relaxed) as i64;
                                        let offset = start_pts.saturating_sub(base);
                                        if pts_offset_us_thread
                                            .compare_exchange(
                                                i64::MIN,
                                                offset,
                                                Ordering::AcqRel,
                                                Ordering::Acquire,
                                            )
                                            .is_ok()
                                        {
                                            log_debug(&format!(
                                                "SubtitleClock: PTS offset set to {:.3}s (pts {} us, base {} us)",
                                                offset as f64 / 1_000_000.0,
                                                start_pts,
                                                base
                                            ));
                                        }
                                    }
                                }
                                let subtitle_sample = get_subtitle_sample(
                                    i,
                                    &mut pending_subtitles,
                                    &mut active_subtitle,
                                    base_samples,
                                    samples_written_since_start,
                                );
                                // Mix audio + subtitle
                                let mixed =
                                    (audio_sample * volume + subtitle_sample).clamp(-1.0, 1.0);
                                let val = (mixed * i16::MAX as f32) as i16;
                                let bytes = val.to_le_bytes();
                                out[offset] = bytes[0];
                                out[offset + 1] = bytes[1];
                                offset += 2;
                            }
                        }
                        WasapiSampleFormat::Float32 => {
                            let mut offset = 0usize;
                            for i in 0..sample_count {
                                let audio_sample = match source.next() {
                                    Some(v) => v,
                                    None => {
                                        finished = true;
                                        0.0
                                    }
                                };
                                if i == 0
                                    && buffer_start_pts_us.is_none()
                                    && let Some(ref pts_clock) = source_pts_us_thread
                                {
                                    let start_pts = pts_clock.load(Ordering::Acquire);
                                    buffer_start_pts_us = Some(start_pts);
                                    if pts_offset_us_thread.load(Ordering::Acquire) == i64::MIN {
                                        let base =
                                            base_micros_thread.load(Ordering::Relaxed) as i64;
                                        let offset = start_pts.saturating_sub(base);
                                        if pts_offset_us_thread
                                            .compare_exchange(
                                                i64::MIN,
                                                offset,
                                                Ordering::AcqRel,
                                                Ordering::Acquire,
                                            )
                                            .is_ok()
                                        {
                                            log_debug(&format!(
                                                "SubtitleClock: PTS offset set to {:.3}s (pts {} us, base {} us)",
                                                offset as f64 / 1_000_000.0,
                                                start_pts,
                                                base
                                            ));
                                        }
                                    }
                                }
                                let subtitle_sample = get_subtitle_sample(
                                    i,
                                    &mut pending_subtitles,
                                    &mut active_subtitle,
                                    base_samples,
                                    samples_written_since_start,
                                );
                                // Mix audio + subtitle
                                let mixed = audio_sample * volume + subtitle_sample;
                                let bytes = mixed.to_le_bytes();
                                out[offset..offset + 4].copy_from_slice(&bytes);
                                offset += 4;
                            }
                        }
                    }
                }

                // Update sample counter
                samples_written_since_start += sample_count as u64;

                if let Err(e) = unsafe { render.ReleaseBuffer(available, 0) } {
                    log_debug(&format!("WASAPI: ReleaseBuffer failed: {}", e));
                    return;
                }

                if let Some(start_pts) = buffer_start_pts_us
                    && sample_rate > 0
                    && channels > 0
                {
                    let frames_written = sample_count / channels as usize;
                    let end_pts = (start_pts as i128).saturating_add(
                        (frames_written as i128)
                            .saturating_mul(1_000_000)
                            .saturating_div(sample_rate as i128),
                    ) as i64;
                    let mut candidate = end_pts;
                    loop {
                        let prev = last_written_end_pts_us_thread.load(Ordering::Acquire);
                        if candidate < prev {
                            candidate = prev;
                        }
                        match last_written_end_pts_us_thread.compare_exchange(
                            prev,
                            candidate,
                            Ordering::AcqRel,
                            Ordering::Acquire,
                        ) {
                            Ok(_) => break,
                            Err(actual) => {
                                if candidate < actual {
                                    candidate = actual;
                                }
                            }
                        }
                    }
                }

                update_clock(&clock, &position_units_thread, &last_qpc_thread);

                if finished {
                    stopped_thread.store(true, Ordering::Relaxed);
        // Stop and cleanup
        unsafe {
            if let Err(e) = client.Stop() {
                crate::log_debug(&format!("WASAPI: client.Stop failed: {:?}", e));
            }
        }
                    return;
                }
            }
        });

        Ok(Arc::new(Self {
            tx,
            position_units,
            last_qpc,
            qpc_freq,
            clock_freq,
            base_micros,
            padding_frames,
            sample_rate_hz,
            device_channels,
            volume_bits,
            stopped,
            last_written_end_pts_us,
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

    /// Schedule a subtitle to be mixed at an exact time.
    /// The subtitle audio will be mixed with the main audio at sample-accurate timing.
    pub fn schedule_subtitle(&self, subtitle: ScheduledSubtitle) {
        if let Err(e) = self.tx.send(WasapiCommand::ScheduleSubtitle(subtitle)) {
            log_debug(&format!("WASAPI: schedule_subtitle failed: {}", e));
        }
    }

    /// Clear all pending and active subtitles (e.g., on seek).
    pub fn clear_subtitles(&self) {
        if let Err(e) = self.tx.send(WasapiCommand::ClearSubtitles) {
            log_debug(&format!("WASAPI: clear_subtitles failed: {}", e));
        }
    }

    /// Get the device sample rate (for resampling subtitle audio).
    pub fn device_sample_rate(&self) -> u32 {
        self.sample_rate_hz.load(Ordering::Relaxed)
    }

    /// Get the device channel count (for resampling subtitle audio).
    pub fn device_channels(&self) -> u32 {
        self.device_channels.load(Ordering::Relaxed)
    }

    pub fn position_secs(&self) -> Option<f64> {
        if self.stopped.load(Ordering::Relaxed) {
            return None;
        }
        let clock_freq = self.clock_freq.load(Ordering::Relaxed);
        let sample_rate = self.sample_rate_hz.load(Ordering::Relaxed);
        // WASAPI not ready yet - clock not initialized
        if clock_freq == 0 || sample_rate == 0 {
            return None;
        }
        let base = self.base_micros.load(Ordering::Relaxed);
        let units = self.position_units.load(Ordering::Relaxed);
        let mut secs = units as f64 / clock_freq as f64;
        let last_qpc = self.last_qpc.load(Ordering::Relaxed);
        let qpc_freq = self.qpc_freq.load(Ordering::Relaxed);
        if last_qpc > 0 && qpc_freq > 0 {
            let mut now_qpc: i64 = 0;
            if unsafe { QueryPerformanceCounter(&mut now_qpc) }.is_ok() {
                let delta = now_qpc.saturating_sub(last_qpc as i64);
                secs += delta as f64 / qpc_freq as f64;
            }
        }
        let padding = self.padding_frames.load(Ordering::Relaxed);
        let padding_secs = padding as f64 / sample_rate as f64;
        Some(((base as f64 / 1_000_000.0) + secs - padding_secs).max(0.0))
    }

    pub fn audible_time_us(&self) -> Option<i64> {
        if self.stopped.load(Ordering::Relaxed) {
            return None;
        }
        let end_pts = self.last_written_end_pts_us.load(Ordering::Acquire);
        if end_pts <= 0 {
            return None;
        }
        let sample_rate = self.sample_rate_hz.load(Ordering::Relaxed);
        if sample_rate == 0 {
            return None;
        }
        let padding = self.padding_frames.load(Ordering::Relaxed);
        let padding_us = (padding as i128)
            .saturating_mul(1_000_000)
            .saturating_div(sample_rate as i128);
        let audible = (end_pts as i128).saturating_sub(padding_us);
        Some(audible.max(0) as i64)
    }

    pub fn subtitle_timing_debug(&self) -> Option<(u32, i64, i64, u32)> {
        if self.stopped.load(Ordering::Relaxed) {
            return None;
        }
        let last_end = self.last_written_end_pts_us.load(Ordering::Acquire);
        if last_end <= 0 {
            return None;
        }
        let sample_rate = self.sample_rate_hz.load(Ordering::Relaxed);
        if sample_rate == 0 {
            return None;
        }
        let padding = self.padding_frames.load(Ordering::Relaxed);
        let padding_us = (padding as i128)
            .saturating_mul(1_000_000)
            .saturating_div(sample_rate as i128)
            .max(0) as i64;
        Some((padding, padding_us, last_end, sample_rate))
    }

}

fn update_clock(clock: &IAudioClock, position_units: &AtomicU64, last_qpc: &AtomicU64) {
    let mut position: u64 = 0;
    let mut _qpc: u64 = 0;
    if unsafe { clock.GetPosition(&mut position, Some(&mut _qpc)) }.is_ok() {
        position_units.store(position, Ordering::Relaxed);
        last_qpc.store(_qpc, Ordering::Relaxed);
    }
}

fn parse_mix_format(
    mix_ptr: *mut WAVEFORMATEX,
) -> (Option<MixFormat>, Option<WasapiSampleFormat>) {
    unsafe {
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
