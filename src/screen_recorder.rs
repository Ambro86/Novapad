//! Integrated screen recorder with video + audio + encoding
//!
//! This module coordinates video capture, audio capture, and MP4 encoding

use crate::audio_capture::{self, AudioRecorderHandle};
use crate::graphics_capture::MonitorInfo;
use crate::mf_encoder::Mp4StreamWriter;
use crate::video_recorder::{self, VideoRecorderHandle};
use std::path::PathBuf;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread::{self, JoinHandle};
use std::time::Duration;

/// Integrated screen recording session
pub struct ScreenRecorder {
    video_recorder: VideoRecorderHandle,
    audio_recorder: AudioRecorderHandle,
    encoder_stop: Arc<AtomicBool>,
    encoder_thread: Option<JoinHandle<Result<(), String>>>,
}

impl ScreenRecorder {
    /// Start a new screen recording session
    pub fn start(monitor: &MonitorInfo, output_path: PathBuf) -> Result<Self, String> {
        crate::log_debug(&format!("Starting screen recording: {:?}", output_path));

        // Start video capture
        let video_recorder = video_recorder::start_video_recording(monitor)?;
        crate::log_debug("Video recorder started");

        // Start audio capture
        let audio_recorder = audio_capture::start_audio_recording()?;
        crate::log_debug(&format!(
            "Audio recorder started: {} Hz, {} channels",
            audio_recorder.sample_rate, audio_recorder.channels
        ));

        // Create MP4 writer with audio sample rate from capture
        let writer = Mp4StreamWriter::create(
            &output_path,
            monitor.width,
            monitor.height,
            audio_recorder.sample_rate,
        )?;
        crate::log_debug("MP4 writer created");

        // Start encoder thread
        let encoder_stop = Arc::new(AtomicBool::new(false));
        let encoder_stop_clone = Arc::clone(&encoder_stop);
        let video_queue = Arc::clone(&video_recorder.frame_queue);
        let audio_queue = Arc::clone(&audio_recorder.audio_queue);

        let encoder_thread = thread::spawn(move || {
            encoder_loop(writer, video_queue, audio_queue, encoder_stop_clone)
        });

        crate::log_debug("Encoder thread started");

        Ok(ScreenRecorder {
            video_recorder,
            audio_recorder,
            encoder_stop,
            encoder_thread: Some(encoder_thread),
        })
    }

    /// Stop recording and finalize the video file
    pub fn stop(mut self) -> Result<(), String> {
        crate::log_debug("Stopping screen recording");

        // FIRST: Signal encoder to stop (so it knows to drain queues and exit)
        self.encoder_stop.store(true, Ordering::SeqCst);
        crate::log_debug("Signaled encoder to stop");

        // SECOND: Signal recorders to stop producing new frames/audio
        // (but don't wait for them yet - let them push final samples)
        self.video_recorder.signal_stop();
        self.audio_recorder.signal_stop();
        crate::log_debug("Signaled recorders to stop");

        // THIRD: Wait for encoder to finish (it will drain remaining frames while D3D11 resources are still valid)
        if let Some(thread) = self.encoder_thread.take() {
            crate::log_debug("Waiting for encoder to finish...");
            thread
                .join()
                .map_err(|_| "Encoder thread panicked".to_string())??;
            crate::log_debug("Encoder stopped");
        }

        // FOURTH: Now wait for recorders to fully shut down
        self.video_recorder.join()?;
        crate::log_debug("Video recorder stopped");

        self.audio_recorder.join()?;
        crate::log_debug("Audio recorder stopped");

        crate::log_debug("Screen recording stopped successfully");
        Ok(())
    }
}

/// Encoder loop that reads from video and audio queues and writes to MP4
fn encoder_loop(
    mut writer: Mp4StreamWriter,
    video_queue: Arc<video_recorder::FrameQueue>,
    audio_queue: Arc<audio_capture::AudioQueue>,
    stop: Arc<AtomicBool>,
) -> Result<(), String> {
    crate::log_debug("Encoder loop started");

    let mut last_video_ts = 0i64;
    let mut frames_encoded = 0u64;
    let mut audio_samples_encoded = 0u64;

    loop {
        let should_stop = stop.load(Ordering::SeqCst);

        // If stop signal received, check if we can exit
        if should_stop {
            let video_empty = video_queue.is_empty();
            let audio_empty = audio_queue.is_empty();

            if video_empty && audio_empty {
                crate::log_debug(
                    "Stop signal received and queues empty, waiting 100ms for final samples...",
                );
                thread::sleep(Duration::from_millis(100));

                // Final check
                if video_queue.is_empty() && audio_queue.is_empty() {
                    crate::log_debug("Exiting encoder loop - no more data");
                    break;
                }
            } else {
                crate::log_debug(&format!(
                    "Stop signal received - draining queues (video_empty: {}, audio_empty: {})",
                    video_empty, audio_empty
                ));
                // Continue processing - don't exit until BOTH queues are empty
            }
        }

        // Process video frames with adaptive timeout
        // Use shorter timeout if not stopping, to process audio faster
        let video_timeout = if should_stop {
            Duration::from_millis(10)
        } else {
            Duration::from_millis(30) // Reduced from 100ms to prevent audio backlog
        };
        if let Some(frame) = video_queue.pop(video_timeout) {
            if frame.timestamp > last_video_ts {
                match writer.write_video_frame(&frame) {
                    Ok(()) => {
                        last_video_ts = frame.timestamp;
                        frames_encoded += 1;

                        if frames_encoded % 30 == 0 {
                            crate::log_debug(&format!(
                                "Encoded {} video frames, {} audio samples",
                                frames_encoded, audio_samples_encoded
                            ));
                        }
                    }
                    Err(e) => {
                        crate::log_debug(&format!("Error writing video frame: {}", e));
                    }
                }
            }
        }

        // Process audio samples with shorter timeout when stopping
        let audio_timeout = if should_stop {
            Duration::from_millis(5)
        } else {
            Duration::from_millis(10)
        };

        // Check if we should write more audio
        // Don't write if audio is too far behind video (prevents MF blocking)
        let audio_timestamp = writer.get_audio_timestamp();
        let max_audio_lag = 150_000_000; // 15 seconds max lag in 100-nanosecond units (increased for slow encoding)
        let audio_behind_video = last_video_ts.saturating_sub(audio_timestamp);
        let can_write_audio = !should_stop
            || (audio_timestamp <= last_video_ts && audio_behind_video < max_audio_lag);

        if should_stop && !audio_queue.is_empty() {
            if can_write_audio {
                crate::log_debug(&format!(
                    "Attempting to drain audio queue (audio_ts: {}, video_ts: {}, lag: {} ms)...",
                    audio_timestamp,
                    last_video_ts,
                    audio_behind_video / 10_000
                ));
            } else {
                let lag_seconds = audio_behind_video as f64 / 10_000_000.0;
                crate::log_debug(&format!(
                    "Audio too far behind video ({:.2}s lag), discarding remaining audio to prevent MF blocking",
                    lag_seconds
                ));
                // Drain queue without writing
                let mut discarded = 0;
                while let Some(_) = audio_queue.pop(Duration::from_millis(1)) {
                    discarded += 1;
                }
                crate::log_debug(&format!(
                    "Discarded {} audio samples (audio was at {:.2}s, video at {:.2}s)",
                    discarded,
                    audio_timestamp as f64 / 10_000_000.0,
                    last_video_ts as f64 / 10_000_000.0
                ));
            }
        }

        if can_write_audio {
            match audio_queue.pop(audio_timeout) {
                Some(audio_sample) => {
                    if should_stop {
                        crate::log_debug(&format!(
                            "Popped audio sample: {} bytes",
                            audio_sample.data.len()
                        ));
                    }

                    if should_stop {
                        crate::log_debug("Writing audio sample...");
                    }

                    match writer.write_audio_samples(&audio_sample.data) {
                        Ok(()) => {
                            audio_samples_encoded += audio_sample.data.len() as u64;

                            // Log every second of audio written
                            let audio_duration = writer.get_audio_duration_seconds();
                            if (audio_duration as u64) % 1 == 0 && audio_duration > 0.0 {
                                let prev_duration = ((audio_duration - 0.1) as u64) % 1;
                                if prev_duration != 0 {
                                    crate::log_debug(&format!(
                                        "Audio written: {:.2}s (queue size: {})",
                                        audio_duration,
                                        audio_queue.len()
                                    ));
                                }
                            }

                            if should_stop {
                                crate::log_debug("Audio sample written successfully");
                            }
                        }
                        Err(e) => {
                            crate::log_debug(&format!("Error writing audio: {}", e));
                        }
                    }
                }
                None => {
                    if should_stop && !audio_queue.is_empty() {
                        crate::log_debug(
                            "WARNING: audio_queue.pop() returned None but queue not empty!",
                        );
                    }
                }
            }
        } else if !should_stop {
            // Log when audio writing is blocked during recording (shouldn't happen normally)
            crate::log_debug(&format!(
                "WARNING: Audio writing blocked during recording! audio_ts={}, video_ts={}, lag={}s",
                writer.get_audio_timestamp(),
                last_video_ts,
                (last_video_ts - writer.get_audio_timestamp()) as f64 / 10_000_000.0
            ));
        }
    }

    // Calculate actual durations
    let video_duration_seconds = if frames_encoded > 0 {
        (last_video_ts as f64) / 10_000_000.0
    } else {
        0.0
    };

    let audio_duration_seconds = writer.get_audio_duration_seconds();

    crate::log_debug(&format!(
        "=== RECORDING STATISTICS ===\n\
         Video: {} frames, duration: {:.2} seconds (last_ts: {})\n\
         Audio: {} samples, duration: {:.2} seconds\n\
         ============================",
        frames_encoded,
        video_duration_seconds,
        last_video_ts,
        audio_samples_encoded,
        audio_duration_seconds
    ));

    // Check if any data was written
    if frames_encoded == 0 {
        crate::log_debug(&format!(
            "Warning: No video frames were encoded. Recording may have been too short. \
             Audio samples: {}",
            audio_samples_encoded
        ));
        // Don't return error - allow finalization to attempt completion
    }

    // Finalize the MP4 file
    crate::log_debug("Calling writer.finalize()...");
    writer.finalize()?;
    crate::log_debug("MP4 file finalized successfully");

    Ok(())
}
