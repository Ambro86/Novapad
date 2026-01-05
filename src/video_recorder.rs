//! Video recording module for screen capture
//!
//! This module handles video frame capture, queuing, and coordination with
//! audio recording for synchronized A/V output.

use crate::graphics_capture::{CaptureSession, MonitorInfo};
use std::collections::VecDeque;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::{Duration, Instant};
use windows::Win32::Graphics::Direct3D11::ID3D11Texture2D;
use windows::Win32::System::Performance::{QueryPerformanceCounter, QueryPerformanceFrequency};

/// A single captured video frame with timestamp
#[derive(Clone)]
pub struct VideoFrame {
    pub texture: ID3D11Texture2D,
    pub timestamp: i64, // 100-nanosecond units
    pub width: u32,
    pub height: u32,
}

/// Thread-safe queue for video frames
pub struct FrameQueue {
    inner: Mutex<VecDeque<VideoFrame>>,
    condvar: Condvar,
    max_frames: usize,
}

impl FrameQueue {
    pub fn new(max_frames: usize) -> Self {
        FrameQueue {
            inner: Mutex::new(VecDeque::with_capacity(max_frames)),
            condvar: Condvar::new(),
            max_frames,
        }
    }

    /// Push a frame to the queue (drops oldest if full)
    pub fn push(&self, frame: VideoFrame) {
        let mut queue = self.inner.lock().unwrap();

        if queue.len() >= self.max_frames {
            let dropped = queue.pop_front();
            if let Some(f) = dropped {
                crate::log_debug(&format!(
                    "Frame queue overflow - dropped frame at ts={}",
                    f.timestamp
                ));
            }
        }

        queue.push_back(frame);
        self.condvar.notify_one();
    }

    /// Try to pop a frame from the queue (with timeout)
    pub fn pop(&self, timeout: Duration) -> Option<VideoFrame> {
        let mut queue = self.inner.lock().unwrap();

        if queue.is_empty() {
            let result = self.condvar.wait_timeout(queue, timeout).unwrap();
            queue = result.0;
        }

        queue.pop_front()
    }

    /// Check if queue is empty
    pub fn is_empty(&self) -> bool {
        self.inner.lock().unwrap().is_empty()
    }

    /// Get current queue length
    pub fn len(&self) -> usize {
        self.inner.lock().unwrap().len()
    }
}

/// Handle to control video recording
pub struct VideoRecorderHandle {
    stop: Arc<AtomicBool>,
    threads: Vec<JoinHandle<Result<(), String>>>,
    pub frame_queue: Arc<FrameQueue>,
}

impl VideoRecorderHandle {
    /// Stop video recording and wait for threads to finish
    pub fn stop(self) -> Result<(), String> {
        self.stop.store(true, Ordering::SeqCst);

        for thread in self.threads {
            thread
                .join()
                .map_err(|_| "Video capture thread panicked".to_string())??;
        }

        Ok(())
    }
}

/// Start video recording for the specified monitor
pub fn start_video_recording(monitor_info: &MonitorInfo) -> Result<VideoRecorderHandle, String> {
    let frame_queue = Arc::new(FrameQueue::new(60)); // Max 2 seconds at 30 FPS
    let stop = Arc::new(AtomicBool::new(false));

    // Create capture session
    let session = CaptureSession::new_for_monitor(monitor_info)?;

    // Start capture
    session.start()?;

    // Spawn capture thread
    let capture_thread = {
        let session = Arc::clone(&session);
        let frame_queue = Arc::clone(&frame_queue);
        let stop = Arc::clone(&stop);

        thread::spawn(move || capture_loop(session, frame_queue, stop))
    };

    Ok(VideoRecorderHandle {
        stop,
        threads: vec![capture_thread],
        frame_queue,
    })
}

/// Main capture loop (runs in dedicated thread)
fn capture_loop(
    session: Arc<CaptureSession>,
    frame_queue: Arc<FrameQueue>,
    stop: Arc<AtomicBool>,
) -> Result<(), String> {
    unsafe {
        let start_qpc = query_performance_counter()?;
        let qpc_freq = query_performance_frequency()?;

        let target_interval = Duration::from_nanos(33_333_333); // ~30 FPS
        let mut last_capture = Instant::now();

        crate::log_debug("Video capture loop started");

        while !stop.load(Ordering::SeqCst) {
            let now = Instant::now();

            // Capture frame if enough time elapsed
            if now.duration_since(last_capture) >= target_interval {
                match session.try_get_next_frame() {
                    Ok(Some(texture)) => {
                        let qpc = query_performance_counter()?;
                        let timestamp = ((qpc - start_qpc) * 10_000_000) / qpc_freq;

                        let mut desc =
                            windows::Win32::Graphics::Direct3D11::D3D11_TEXTURE2D_DESC::default();
                        texture.GetDesc(&mut desc);

                        let frame = VideoFrame {
                            texture,
                            timestamp,
                            width: desc.Width,
                            height: desc.Height,
                        };

                        frame_queue.push(frame);
                        last_capture = now;
                    }
                    Ok(None) => {
                        // No frame available yet, continue
                    }
                    Err(e) => {
                        crate::log_debug(&format!("Frame capture error: {}", e));
                        // Continue trying
                    }
                }
            }

            // Small sleep to avoid busy-wait
            thread::sleep(Duration::from_millis(5));
        }

        crate::log_debug("Video capture loop stopped");
        Ok(())
    }
}

/// Query performance counter (for timestamps)
unsafe fn query_performance_counter() -> Result<i64, String> {
    let mut counter = 0i64;
    QueryPerformanceCounter(&mut counter)
        .map_err(|e| format!("QueryPerformanceCounter failed: {}", e))?;
    Ok(counter)
}

/// Query performance frequency
unsafe fn query_performance_frequency() -> Result<i64, String> {
    let mut freq = 0i64;
    QueryPerformanceFrequency(&mut freq)
        .map_err(|e| format!("QueryPerformanceFrequency failed: {}", e))?;
    Ok(freq)
}
