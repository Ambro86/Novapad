use crate::accessibility::to_wide;
use crate::log_debug;
use crate::miniaudio_sys::*;
use rodio::Source;
use std::collections::VecDeque;
use std::ffi::c_void;
use std::mem::MaybeUninit;
use std::path::Path;
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, mpsc};
use std::time::Duration;
use std::time::Instant;

const MA_SUCCESS_CODE: ma_result = 0;
const MA_AT_END_CODE: ma_result = -17;
const MA_DEVICE_TYPE_PLAYBACK_CODE: ma_device_type = 1;
const MA_FORMAT_F32_CODE: ma_format = 5;

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

enum MiniaudioCommand {
    ScheduleSubtitle(ScheduledSubtitle),
    ClearSubtitles,
}

struct ActiveSubtitle {
    samples: Arc<[f32]>,
    read_offset: usize,
    volume: f32,
}

struct MiniaudioOutputState {
    frames_written: AtomicU64,
    base_frames: AtomicU64,
    sample_rate_hz: AtomicU32,
    device_channels: AtomicU32,
    volume_bits: AtomicU32,
    stopped: AtomicBool,
    paused: AtomicBool,
    start_instant: Instant,
    last_callback_ns: AtomicU64,
    output_latency_frames: AtomicU64,
}

struct MiniaudioOutputInner {
    source: Box<dyn Source<Item = f32> + Send>,
    rx: mpsc::Receiver<MiniaudioCommand>,
    pending_subtitles: VecDeque<ScheduledSubtitle>,
    active_subtitle: Option<ActiveSubtitle>,
    base_samples: u64,
    samples_written: u64,
    channels: u16,
    sample_rate: u32,
    state: Arc<MiniaudioOutputState>,
}

unsafe extern "C" fn miniaudio_data_callback(
    device: *mut ma_device,
    output: *mut c_void,
    _input: *const c_void,
    frame_count: ma_uint32,
) {
    if device.is_null() || output.is_null() {
        return;
    }
    let inner_ptr = unsafe { (*device).pUserData } as *mut MiniaudioOutputInner;
    if inner_ptr.is_null() {
        return;
    }
    let inner = unsafe { &mut *inner_ptr };
    inner.render(device, output as *mut f32, frame_count as usize);
}

impl MiniaudioOutputInner {
    fn handle_commands(&mut self) {
        while let Ok(cmd) = self.rx.try_recv() {
            match cmd {
                MiniaudioCommand::ScheduleSubtitle(sub) => {
                    self.pending_subtitles.push_back(sub);
                }
                MiniaudioCommand::ClearSubtitles => {
                    self.pending_subtitles.clear();
                    self.active_subtitle = None;
                }
            }
        }
    }

    fn render(&mut self, device: *mut ma_device, output: *mut f32, frame_count: usize) {
        if self.state.stopped.load(Ordering::Relaxed) {
            return;
        }
        if !device.is_null() {
            let device_channels = unsafe { (*device).playback.channels };
            if device_channels > 0 {
                self.state
                    .device_channels
                    .store(device_channels, Ordering::Relaxed);
            }
        }
        self.handle_commands();

        let sample_count = frame_count.saturating_mul(self.channels as usize);
        let out = unsafe { std::slice::from_raw_parts_mut(output, sample_count) };

        if self.state.paused.load(Ordering::Relaxed) {
            out.fill(0.0);
            return;
        }

        let volume = f32::from_bits(self.state.volume_bits.load(Ordering::Relaxed));
        let mut finished = false;
        let sample_rate = self.state.sample_rate_hz.load(Ordering::Relaxed) as u64;
        let max_latency = (sample_rate * 2).max(1); // clamp to 2s
        let raw_latency = if !device.is_null() {
            let period = unsafe { (*device).playback.internalPeriodSizeInFrames } as u64;
            let periods = unsafe { (*device).playback.internalPeriods } as u64;
            let internal_rate = unsafe { (*device).playback.internalSampleRate } as u64;
            let mut latency_frames = if period > 0 && periods > 0 {
                period.saturating_mul(periods)
            } else {
                frame_count as u64
            };
            if internal_rate > 0 && internal_rate != sample_rate {
                latency_frames = latency_frames
                    .saturating_mul(sample_rate)
                    .saturating_div(internal_rate);
            }
            latency_frames.min(max_latency)
        } else {
            (frame_count as u64).min(max_latency)
        };
        let prev = self.state.output_latency_frames.load(Ordering::Relaxed);
        let blended = if prev == 0 {
            raw_latency
        } else {
            (prev * 9 + raw_latency) / 10
        };
        self.state
            .output_latency_frames
            .store(blended, Ordering::Release);
        for (i, out_sample) in out.iter_mut().enumerate() {
            let audio_sample = match self.source.next() {
                Some(v) => v,
                None => {
                    finished = true;
                    0.0
                }
            };

            let subtitle_sample = self.get_subtitle_sample(i);
            *out_sample = (audio_sample * volume + subtitle_sample).clamp(-1.0, 1.0);
        }

        self.samples_written = self.samples_written.saturating_add(sample_count as u64);
        let frames_written = self.samples_written / self.channels as u64;
        self.state
            .frames_written
            .store(frames_written, Ordering::Release);
        let now_ns = self.state.start_instant.elapsed().as_nanos();
        if now_ns <= u64::MAX as u128 {
            self.state
                .last_callback_ns
                .store(now_ns as u64, Ordering::Release);
        }

        if finished {
            self.state.stopped.store(true, Ordering::Release);
            let stop_result = unsafe { ma_device_stop(device) };
            if stop_result != MA_SUCCESS_CODE {
                log_debug(&format!("miniaudio: stop failed: {}", stop_result));
            }
        }
    }

    fn get_subtitle_sample(&mut self, sample_idx: usize) -> f32 {
        let current_sample = self.base_samples + self.samples_written + sample_idx as u64;

        if self.active_subtitle.is_none()
            && let Some(sub) = self.pending_subtitles.front()
        {
            let target_sample =
                (sub.target_secs * self.sample_rate as f64 * self.channels as f64) as u64;
            if current_sample >= target_sample {
                let sub = self.pending_subtitles.pop_front().unwrap();
                self.active_subtitle = Some(ActiveSubtitle {
                    samples: sub.samples,
                    read_offset: 0,
                    volume: sub.volume,
                });
            }
        }

        if let Some(sub) = &mut self.active_subtitle {
            if sub.read_offset < sub.samples.len() {
                let s = sub.samples[sub.read_offset] * sub.volume;
                sub.read_offset += 1;
                return s;
            }
            self.active_subtitle = None;
        }
        0.0
    }
}

pub struct MiniaudioOutput {
    device: Box<ma_device>,
    inner_ptr: *mut MiniaudioOutputInner,
    tx: mpsc::Sender<MiniaudioCommand>,
    state: Arc<MiniaudioOutputState>,
}

// SAFETY: MiniaudioOutput is shared across threads; it uses atomics and channels for shared state,
// and the underlying miniaudio device APIs are thread-safe for start/stop/volume operations.
unsafe impl Send for MiniaudioOutput {}
unsafe impl Sync for MiniaudioOutput {}

impl MiniaudioOutput {
    pub fn start(
        source: Box<dyn Source<Item = f32> + Send>,
        start_paused: bool,
        volume: f32,
        base_secs: f64,
    ) -> Result<Arc<Self>, String> {
        let channels = source.channels();
        let sample_rate = source.sample_rate();
        if channels == 0 || sample_rate == 0 {
            return Err("miniaudio: invalid source format".to_string());
        }

        let base_frames = (base_secs.max(0.0) * sample_rate as f64) as u64;
        let base_samples = base_frames.saturating_mul(channels as u64);

        let state = Arc::new(MiniaudioOutputState {
            frames_written: AtomicU64::new(0),
            base_frames: AtomicU64::new(base_frames),
            sample_rate_hz: AtomicU32::new(sample_rate),
            device_channels: AtomicU32::new(channels as u32),
            volume_bits: AtomicU32::new(volume.to_bits()),
            stopped: AtomicBool::new(false),
            paused: AtomicBool::new(start_paused),
            start_instant: Instant::now(),
            last_callback_ns: AtomicU64::new(0),
            output_latency_frames: AtomicU64::new(0),
        });

        let (tx, rx) = mpsc::channel();
        let inner = Box::new(MiniaudioOutputInner {
            source,
            rx,
            pending_subtitles: VecDeque::new(),
            active_subtitle: None,
            base_samples,
            samples_written: 0,
            channels,
            sample_rate,
            state: Arc::clone(&state),
        });
        let inner_ptr = Box::into_raw(inner) as *mut c_void;

        let mut config = unsafe { ma_device_config_init(MA_DEVICE_TYPE_PLAYBACK_CODE) };
        config.playback.format = MA_FORMAT_F32_CODE;
        config.playback.channels = channels as ma_uint32;
        config.sampleRate = sample_rate as ma_uint32;
        config.dataCallback = Some(miniaudio_data_callback);
        config.pUserData = inner_ptr;

        let mut device = Box::new(unsafe { MaybeUninit::<ma_device>::zeroed().assume_init() });
        let init_result = unsafe { ma_device_init(ptr::null_mut(), &config, device.as_mut()) };
        if init_result != MA_SUCCESS_CODE {
            unsafe {
                drop(Box::from_raw(inner_ptr as *mut MiniaudioOutputInner));
            }
            return Err(format!("miniaudio: device init failed: {}", init_result));
        }

        if !start_paused {
            let start_result = unsafe { ma_device_start(device.as_mut()) };
            if start_result != MA_SUCCESS_CODE {
                unsafe { ma_device_uninit(device.as_mut()) };
                unsafe {
                    drop(Box::from_raw(inner_ptr as *mut MiniaudioOutputInner));
                }
                return Err(format!("miniaudio: device start failed: {}", start_result));
            }
        }

        Ok(Arc::new(Self {
            device,
            inner_ptr: inner_ptr as *mut MiniaudioOutputInner,
            tx,
            state,
        }))
    }

    pub fn play(&self) {
        if self.state.stopped.load(Ordering::Relaxed) {
            return;
        }
        let result = unsafe { ma_device_start(self.device.as_ref() as *const _ as *mut _) };
        if result != MA_SUCCESS_CODE {
            log_debug(&format!("miniaudio: start failed: {}", result));
            return;
        }
        self.state.paused.store(false, Ordering::Relaxed);
    }

    pub fn pause(&self) {
        if self.state.stopped.load(Ordering::Relaxed) {
            return;
        }
        let result = unsafe { ma_device_stop(self.device.as_ref() as *const _ as *mut _) };
        if result != MA_SUCCESS_CODE {
            log_debug(&format!("miniaudio: stop failed: {}", result));
            return;
        }
        self.state.paused.store(true, Ordering::Relaxed);
    }

    pub fn stop(&self) {
        if self.state.stopped.swap(true, Ordering::AcqRel) {
            return;
        }
        let result = unsafe { ma_device_stop(self.device.as_ref() as *const _ as *mut _) };
        if result != MA_SUCCESS_CODE {
            log_debug(&format!("miniaudio: stop failed: {}", result));
        }
        unsafe {
            ma_device_uninit(self.device.as_ref() as *const _ as *mut _);
        }
    }

    pub fn set_volume(&self, volume: f32) {
        self.state
            .volume_bits
            .store(volume.to_bits(), Ordering::Relaxed);
    }

    pub fn schedule_subtitle(&self, subtitle: ScheduledSubtitle) {
        if let Err(e) = self.tx.send(MiniaudioCommand::ScheduleSubtitle(subtitle)) {
            log_debug(&format!("miniaudio: schedule_subtitle failed: {}", e));
        }
    }

    pub fn clear_subtitles(&self) {
        if let Err(e) = self.tx.send(MiniaudioCommand::ClearSubtitles) {
            log_debug(&format!("miniaudio: clear_subtitles failed: {}", e));
        }
    }

    pub fn device_sample_rate(&self) -> u32 {
        self.state.sample_rate_hz.load(Ordering::Relaxed)
    }

    pub fn device_channels(&self) -> u32 {
        self.state.device_channels.load(Ordering::Relaxed)
    }

    pub fn position_secs(&self) -> Option<f64> {
        if self.state.stopped.load(Ordering::Relaxed) {
            return None;
        }
        let sample_rate = self.state.sample_rate_hz.load(Ordering::Relaxed);
        if sample_rate == 0 {
            return None;
        }
        let base = self.state.base_frames.load(Ordering::Relaxed);
        let frames = self.state.frames_written.load(Ordering::Relaxed);
        let latency = self.state.output_latency_frames.load(Ordering::Acquire);
        if self.state.paused.load(Ordering::Relaxed) {
            let corrected = base.saturating_add(frames).saturating_sub(latency);
            return Some((corrected as f64 / sample_rate as f64).max(0.0));
        }
        let last_ns = self.state.last_callback_ns.load(Ordering::Acquire);
        let mut interp_frames = 0u64;
        if last_ns > 0 {
            let now_ns = self.state.start_instant.elapsed().as_nanos();
            if now_ns >= last_ns as u128 {
                let delta_ns = (now_ns - last_ns as u128).min(u64::MAX as u128) as u64;
                let delta_frames = (delta_ns as u128 * sample_rate as u128) / 1_000_000_000u128;
                // Clamp interpolation to avoid runaway if callbacks stall.
                interp_frames = delta_frames.min(sample_rate as u128) as u64;
            }
        }
        let corrected = base
            .saturating_add(frames)
            .saturating_add(interp_frames)
            .saturating_sub(latency);
        Some((corrected as f64 / sample_rate as f64).max(0.0))
    }
}

impl Drop for MiniaudioOutput {
    fn drop(&mut self) {
        self.stop();
        if !self.inner_ptr.is_null() {
            unsafe {
                drop(Box::from_raw(self.inner_ptr));
            }
            self.inner_ptr = ptr::null_mut();
        }
    }
}

pub struct MiniaudioDecoderSource {
    decoder: ma_decoder,
    sample_rate: u32,
    channels: u16,
    buffer: Vec<f32>,
    buffer_pos: usize,
    total_duration: Option<Duration>,
    finished: bool,
    cursor_frames: Arc<AtomicU64>,
}

// SAFETY: decoder is used only on the playback thread.
unsafe impl Send for MiniaudioDecoderSource {}

impl MiniaudioDecoderSource {
    pub fn try_new(path: &Path, start_seconds: u64) -> Result<Self, String> {
        let path_str = path.to_string_lossy();
        let mut decoder = MaybeUninit::<ma_decoder>::uninit();
        let mut config = unsafe { ma_decoder_config_init(MA_FORMAT_F32_CODE, 0, 0) };
        config.format = MA_FORMAT_F32_CODE;

        let path_wide = to_wide(path_str.as_ref());
        let init_result = unsafe {
            ma_decoder_init_file_w(
                path_wide.as_ptr() as *const wchar_t,
                &config,
                decoder.as_mut_ptr(),
            )
        };
        if init_result != MA_SUCCESS_CODE {
            return Err(format!("miniaudio: decoder init failed: {}", init_result));
        }

        let mut decoder = unsafe { decoder.assume_init() };
        let mut format = MA_FORMAT_F32_CODE;
        let mut channels: ma_uint32 = 0;
        let mut sample_rate: ma_uint32 = 0;
        let format_result = unsafe {
            ma_decoder_get_data_format(
                &mut decoder,
                &mut format,
                &mut channels,
                &mut sample_rate,
                ptr::null_mut(),
                0,
            )
        };
        if format_result != MA_SUCCESS_CODE {
            unsafe { ma_decoder_uninit(&mut decoder) };
            return Err(format!(
                "miniaudio: decoder format failed: {}",
                format_result
            ));
        }

        if start_seconds > 0 && sample_rate > 0 {
            let target_frame = (start_seconds as u128)
                .saturating_mul(sample_rate as u128)
                .min(u64::MAX as u128) as u64;
            let seek_result = unsafe { ma_decoder_seek_to_pcm_frame(&mut decoder, target_frame) };
            if seek_result != MA_SUCCESS_CODE {
                log_debug(&format!("miniaudio: seek failed: {}", seek_result));
            }
        }

        let mut total_duration = None;
        let mut total_frames: ma_uint64 = 0;
        let length_result =
            unsafe { ma_decoder_get_length_in_pcm_frames(&mut decoder, &mut total_frames) };
        if length_result == MA_SUCCESS_CODE && sample_rate > 0 {
            let secs = total_frames as f64 / sample_rate as f64;
            total_duration = Some(Duration::from_secs_f64(secs));
        }

        Ok(Self {
            decoder,
            sample_rate: sample_rate as u32,
            channels: channels as u16,
            buffer: Vec::new(),
            buffer_pos: 0,
            total_duration,
            finished: false,
            cursor_frames: Arc::new(AtomicU64::new(0)),
        })
    }

    pub fn cursor_frames_handle(&self) -> Arc<AtomicU64> {
        Arc::clone(&self.cursor_frames)
    }

    pub fn sample_rate_hz(&self) -> u32 {
        self.sample_rate
    }
}

impl Iterator for MiniaudioDecoderSource {
    type Item = f32;

    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }
        if self.buffer_pos >= self.buffer.len() {
            let frames_to_read = 1024u64;
            let channels = self.channels as usize;
            let mut temp = vec![0.0f32; frames_to_read as usize * channels];
            let mut frames_read: ma_uint64 = 0;
            let result = unsafe {
                ma_decoder_read_pcm_frames(
                    &mut self.decoder,
                    temp.as_mut_ptr() as *mut c_void,
                    frames_to_read,
                    &mut frames_read,
                )
            };
            if frames_read == 0 {
                self.finished = true;
                if result != MA_SUCCESS_CODE && result != MA_AT_END_CODE {
                    log_debug(&format!("miniaudio: read failed: {}", result));
                }
                return None;
            }
            let mut cursor: ma_uint64 = 0;
            let cursor_result =
                unsafe { ma_decoder_get_cursor_in_pcm_frames(&mut self.decoder, &mut cursor) };
            if cursor_result == MA_SUCCESS_CODE {
                self.cursor_frames.store(cursor, Ordering::Release);
            } else {
                log_debug(&format!(
                    "miniaudio: cursor query failed: {}",
                    cursor_result
                ));
            }
            temp.truncate(frames_read as usize * channels);
            self.buffer = temp;
            self.buffer_pos = 0;
        }
        let sample = self.buffer[self.buffer_pos];
        self.buffer_pos += 1;
        Some(sample)
    }
}

impl Source for MiniaudioDecoderSource {
    fn current_span_len(&self) -> Option<usize> {
        None
    }

    fn channels(&self) -> u16 {
        self.channels
    }

    fn sample_rate(&self) -> u32 {
        self.sample_rate
    }

    fn total_duration(&self) -> Option<Duration> {
        self.total_duration
    }
}

impl Drop for MiniaudioDecoderSource {
    fn drop(&mut self) {
        unsafe {
            ma_decoder_uninit(&mut self.decoder);
        }
    }
}
