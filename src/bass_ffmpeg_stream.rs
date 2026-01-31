//! Streaming FFmpeg audio through BASS using a callback.
//! This avoids writing to a WAV file and provides instant playback.

use crate::bass_sys::{
    BASS_SAMPLE_FLOAT, BASS_STREAM_DECODE, BASS_STREAMPROC_END, Dword, Hstream, bass_api,
    bass_error,
};
use crate::ffmpeg_source::FfmpegSource;
use crate::log_debug;
use rodio::Source;
use std::collections::VecDeque;
use std::ffi::c_void;
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};

/// Shared state between FFmpeg decoder thread and BASS callback
struct StreamState {
    /// Audio samples buffer (interleaved f32)
    buffer: Mutex<VecDeque<f32>>,
    /// Signal when new data is available
    data_ready: Condvar,
    /// True when FFmpeg has finished decoding
    eof: AtomicBool,
    /// True when stream should stop
    stop: AtomicBool,
    /// Sample rate
    sample_rate: u32,
    /// Number of channels
    channels: u16,
}

/// Handle for FFmpeg streaming through BASS
pub struct FfmpegBassStream {
    state: Arc<StreamState>,
    decoder_thread: Option<JoinHandle<()>>,
}

impl FfmpegBassStream {
    /// Create a new FFmpeg->BASS stream for the given file
    pub fn new(
        path: &Path,
        start_seconds: u64,
        stream_index: Option<i32>,
    ) -> Result<(Self, Hstream), String> {
        let api = bass_api()?;

        // Create FFmpeg source to get format info
        let source = FfmpegSource::try_new(path, start_seconds, None, stream_index)?;
        let sample_rate = source.sample_rate();
        let channels = source.channels();

        if sample_rate == 0 || channels == 0 {
            return Err("FFmpeg: invalid format".to_string());
        }

        log_debug(&format!(
            "FFmpeg stream: rate={} channels={}",
            sample_rate, channels
        ));

        // Create shared state
        let state = Arc::new(StreamState {
            buffer: Mutex::new(VecDeque::with_capacity(
                sample_rate as usize * channels as usize,
            )),
            data_ready: Condvar::new(),
            eof: AtomicBool::new(false),
            stop: AtomicBool::new(false),
            sample_rate,
            channels,
        });

        // Create BASS stream with callback
        let state_ptr = Arc::into_raw(state.clone()) as *mut c_void;
        let handle = unsafe {
            (api.stream_create)(
                sample_rate,
                channels as u32,
                BASS_SAMPLE_FLOAT | BASS_STREAM_DECODE,
                stream_callback,
                state_ptr,
            )
        };

        if handle == 0 {
            // Clean up the Arc we leaked
            unsafe {
                drop(Arc::from_raw(state_ptr as *const StreamState));
            }
            return Err(format!(
                "BASS_StreamCreate failed (error {})",
                bass_error(api)
            ));
        }

        // Start decoder thread
        let decoder_state = state.clone();
        let decoder_thread = thread::spawn(move || {
            decode_loop(source, decoder_state);
        });

        Ok((
            FfmpegBassStream {
                state,
                decoder_thread: Some(decoder_thread),
            },
            handle,
        ))
    }

    /// Stop the stream and clean up
    pub fn stop(&mut self) {
        self.state.stop.store(true, Ordering::SeqCst);
        self.state.data_ready.notify_all();
        if let Some(thread) = self.decoder_thread.take() {
            crate::log_if_err!(thread.join());
        }
    }
}

impl Drop for FfmpegBassStream {
    fn drop(&mut self) {
        self.stop();
    }
}

/// BASS stream callback - called when BASS needs more audio data
unsafe extern "C" fn stream_callback(
    _handle: Hstream,
    buffer: *mut c_void,
    length: Dword,
    user: *mut c_void,
) -> Dword {
    if user.is_null() {
        return BASS_STREAMPROC_END;
    }

    // Get state from user pointer (don't take ownership)
    let state = &*(user as *const StreamState);

    // Calculate how many samples we need (length is in bytes, f32 = 4 bytes)
    let samples_needed = length as usize / 4;

    let mut written = 0usize;
    let out_buffer = std::slice::from_raw_parts_mut(buffer as *mut f32, samples_needed);

    // Try to get samples from buffer
    {
        let mut buf = state.buffer.lock().unwrap_or_else(|e| e.into_inner());

        while written < samples_needed {
            if let Some(sample) = buf.pop_front() {
                out_buffer[written] = sample;
                written += 1;
            } else if state.eof.load(Ordering::SeqCst) {
                // No more data and EOF reached
                break;
            } else {
                // Buffer empty but not EOF - fill with silence and wait for more
                // This prevents blocking the audio thread
                break;
            }
        }
    }

    // Notify decoder we consumed data
    state.data_ready.notify_one();

    // Fill remaining with silence if needed
    for sample in out_buffer.iter_mut().take(samples_needed).skip(written) {
        *sample = 0.0;
    }

    if written == 0 && state.eof.load(Ordering::SeqCst) {
        BASS_STREAMPROC_END
    } else {
        length // Return requested length (we filled with silence if needed)
    }
}

/// Decoder thread loop - reads from FFmpeg and fills buffer
fn decode_loop(mut source: FfmpegSource, state: Arc<StreamState>) {
    let max_buffer_samples = state.sample_rate as usize * state.channels as usize * 2; // 2 seconds
    let mut sample_count: u64 = 0;

    for sample in source.by_ref() {
        if state.stop.load(Ordering::SeqCst) {
            log_debug(&format!(
                "FFmpeg decode_loop: stopped after {} samples",
                sample_count
            ));
            break;
        }
        sample_count += 1;
        if sample_count == 1 {
            log_debug("FFmpeg decode_loop: first sample received");
        }

        // Wait if buffer is full
        {
            let mut buf = state.buffer.lock().unwrap_or_else(|e| e.into_inner());
            while buf.len() >= max_buffer_samples && !state.stop.load(Ordering::SeqCst) {
                buf = state
                    .data_ready
                    .wait(buf)
                    .unwrap_or_else(|e| e.into_inner());
            }

            if state.stop.load(Ordering::SeqCst) {
                break;
            }

            buf.push_back(sample);
        }
    }

    state.eof.store(true, Ordering::SeqCst);
    log_debug(&format!(
        "FFmpeg stream: decode complete ({} samples)",
        sample_count
    ));
}
