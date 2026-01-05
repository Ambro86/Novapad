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
        crate::log_debug("Audio recorder started");

        // Create MP4 writer
        let writer = Mp4StreamWriter::create(&output_path, monitor.width, monitor.height)?;
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

        // Signal encoder to stop
        self.encoder_stop.store(true, Ordering::SeqCst);

        // Stop video capture first
        self.video_recorder.stop()?;
        crate::log_debug("Video recorder stopped");

        // Stop audio capture
        self.audio_recorder.stop()?;
        crate::log_debug("Audio recorder stopped");

        // Wait for encoder thread to finish
        if let Some(thread) = self.encoder_thread.take() {
            thread
                .join()
                .map_err(|_| "Encoder thread panicked".to_string())??;
            crate::log_debug("Encoder thread stopped");
        }

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

    let timeout = Duration::from_millis(100);
    let mut last_video_ts = 0i64;
    let mut last_audio_ts = 0i64;
    let mut frames_encoded = 0u64;
    let mut audio_samples_encoded = 0u64;

    loop {
        let should_stop = stop.load(Ordering::SeqCst);

        // Try to get video frame
        if let Some(frame) = video_queue.pop(timeout) {
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

        // Try to get audio samples
        if let Some(audio_sample) = audio_queue.pop(Duration::from_millis(10)) {
            if audio_sample.timestamp > last_audio_ts {
                match writer.write_audio_samples(&audio_sample.data, audio_sample.timestamp) {
                    Ok(()) => {
                        last_audio_ts = audio_sample.timestamp;
                        audio_samples_encoded += audio_sample.data.len() as u64;
                    }
                    Err(e) => {
                        crate::log_debug(&format!("Error writing audio: {}", e));
                    }
                }
            }
        }

        // If stop signal received and queues are empty, exit
        if should_stop && video_queue.is_empty() && audio_queue.is_empty() {
            // Give some time for final samples to arrive
            thread::sleep(Duration::from_millis(200));

            // Final check
            if video_queue.is_empty() && audio_queue.is_empty() {
                break;
            }
        }
    }

    crate::log_debug(&format!(
        "Encoder loop finished: {} video frames, {} audio samples",
        frames_encoded, audio_samples_encoded
    ));

    // Check if any data was written
    if frames_encoded == 0 {
        return Err(format!(
            "No video frames were captured! Video capture may have failed. \
             Audio samples: {}. Check if screen capture permissions are enabled.",
            audio_samples_encoded
        ));
    }

    // Finalize the MP4 file
    writer.finalize()?;
    crate::log_debug("MP4 file finalized");

    Ok(())
}
