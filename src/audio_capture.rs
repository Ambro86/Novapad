//! Audio capture module using WASAPI (Windows Audio Session API)
//!
//! This module provides audio capture functionality for screen recording

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::Duration;
use windows::Win32::Media::Audio::{
    AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, IAudioCaptureClient, IAudioClient,
    IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, eConsole, eRender,
};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
};

/// Audio sample with timestamp
#[derive(Clone)]
pub struct AudioSample {
    pub data: Vec<i16>, // 16-bit PCM stereo samples
    pub timestamp: i64, // 100-nanosecond units
    pub sample_rate: u32,
    pub channels: u16,
}

/// Thread-safe queue for audio samples
pub struct AudioQueue {
    inner: Mutex<Vec<AudioSample>>,
    condvar: Condvar,
    max_samples: usize,
}

impl AudioQueue {
    pub fn new(max_samples: usize) -> Self {
        AudioQueue {
            inner: Mutex::new(Vec::with_capacity(max_samples)),
            condvar: Condvar::new(),
            max_samples,
        }
    }

    pub fn push(&self, sample: AudioSample) {
        let mut queue = self.inner.lock().unwrap();

        if queue.len() >= self.max_samples {
            queue.remove(0); // Drop oldest
        }

        queue.push(sample);
        self.condvar.notify_one();
    }

    pub fn pop(&self, timeout: Duration) -> Option<AudioSample> {
        let mut queue = self.inner.lock().unwrap();

        if queue.is_empty() {
            let result = self.condvar.wait_timeout(queue, timeout).unwrap();
            queue = result.0;
        }

        if !queue.is_empty() {
            Some(queue.remove(0))
        } else {
            None
        }
    }

    pub fn len(&self) -> usize {
        self.inner.lock().unwrap().len()
    }

    pub fn is_empty(&self) -> bool {
        self.inner.lock().unwrap().is_empty()
    }
}

/// Handle to control audio recording
pub struct AudioRecorderHandle {
    stop: Arc<AtomicBool>,
    thread: Option<JoinHandle<Result<(), String>>>,
    pub audio_queue: Arc<AudioQueue>,
}

impl AudioRecorderHandle {
    pub fn stop(mut self) -> Result<(), String> {
        self.stop.store(true, Ordering::SeqCst);

        if let Some(thread) = self.thread.take() {
            thread
                .join()
                .map_err(|_| "Audio capture thread panicked".to_string())??;
        }

        Ok(())
    }
}

/// Start audio recording using WASAPI loopback
pub fn start_audio_recording() -> Result<AudioRecorderHandle, String> {
    let audio_queue = Arc::new(AudioQueue::new(200)); // ~4-5 seconds buffer at 44.1kHz
    let stop = Arc::new(AtomicBool::new(false));

    let audio_queue_clone = Arc::clone(&audio_queue);
    let stop_clone = Arc::clone(&stop);

    let thread = thread::spawn(move || audio_capture_loop(audio_queue_clone, stop_clone));

    Ok(AudioRecorderHandle {
        stop,
        thread: Some(thread),
        audio_queue,
    })
}

fn audio_capture_loop(audio_queue: Arc<AudioQueue>, stop: Arc<AtomicBool>) -> Result<(), String> {
    unsafe {
        // Initialize COM for this thread
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)
            .ok()
            .map_err(|e| format!("CoInitializeEx failed: {:?}", e))?;

        let result = audio_capture_loop_impl(audio_queue, stop);

        CoUninitialize();
        result
    }
}

unsafe fn audio_capture_loop_impl(
    audio_queue: Arc<AudioQueue>,
    stop: Arc<AtomicBool>,
) -> Result<(), String> {
    crate::log_debug("Audio capture loop started");

    // Get default audio endpoint (speakers/headphones for loopback)
    let enumerator: IMMDeviceEnumerator =
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
            .map_err(|e| format!("Failed to create device enumerator: {}", e))?;

    let device: IMMDevice = enumerator
        .GetDefaultAudioEndpoint(eRender, eConsole)
        .map_err(|e| format!("Failed to get default audio endpoint: {}", e))?;

    let audio_client: IAudioClient = device
        .Activate(CLSCTX_ALL, None)
        .map_err(|e| format!("Failed to activate audio client: {}", e))?;

    // Get mix format
    let format_ptr = audio_client
        .GetMixFormat()
        .map_err(|e| format!("GetMixFormat failed: {}", e))?;
    let format = &*format_ptr;

    // Initialize audio client for loopback capture
    let buffer_duration = 10_000_000; // 1 second in 100ns units
    audio_client
        .Initialize(
            AUDCLNT_SHAREMODE_SHARED,
            AUDCLNT_STREAMFLAGS_LOOPBACK,
            buffer_duration,
            0,
            format,
            None,
        )
        .map_err(|e| format!("Initialize failed: {}", e))?;

    let capture_client: IAudioCaptureClient = audio_client
        .GetService()
        .map_err(|e| format!("GetService failed: {}", e))?;

    // Start capture
    audio_client
        .Start()
        .map_err(|e| format!("Start failed: {}", e))?;

    // Copy packed struct fields to avoid unaligned reference errors
    let sample_rate = format.nSamplesPerSec;
    let channels = format.nChannels;
    let bits_per_sample = format.wBitsPerSample;
    let bytes_per_sample = (bits_per_sample / 8) as usize;

    crate::log_debug(&format!(
        "Audio format: {} Hz, {} channels, {} bits",
        sample_rate, channels, bits_per_sample
    ));

    let mut qpc_freq = 0i64;
    windows::Win32::System::Performance::QueryPerformanceFrequency(&mut qpc_freq)
        .ok()
        .ok_or("QueryPerformanceFrequency failed")?;

    let mut start_qpc = 0i64;
    windows::Win32::System::Performance::QueryPerformanceCounter(&mut start_qpc)
        .ok()
        .ok_or("QueryPerformanceCounter failed")?;

    // Capture loop
    while !stop.load(Ordering::SeqCst) {
        thread::sleep(Duration::from_millis(10));

        let packet_length = match capture_client.GetNextPacketSize() {
            Ok(len) => len,
            Err(_) => continue,
        };

        let mut current_packet_length = packet_length;
        while current_packet_length > 0 {
            let mut buffer_ptr: *mut u8 = std::ptr::null_mut();
            let mut num_frames = 0u32;
            let mut flags = 0u32;

            if capture_client
                .GetBuffer(&mut buffer_ptr, &mut num_frames, &mut flags, None, None)
                .is_err()
            {
                break;
            }

            if num_frames > 0 && !buffer_ptr.is_null() {
                // Calculate timestamp
                let mut current_qpc = 0i64;
                let _ =
                    windows::Win32::System::Performance::QueryPerformanceCounter(&mut current_qpc);
                let timestamp = ((current_qpc - start_qpc) * 10_000_000) / qpc_freq;

                // Convert to 16-bit PCM stereo
                let frame_size = (channels as usize) * bytes_per_sample;
                let total_bytes = (num_frames as usize) * frame_size;
                let buffer_slice = std::slice::from_raw_parts(buffer_ptr, total_bytes);

                let mut samples = Vec::with_capacity(num_frames as usize * 2); // stereo

                if bits_per_sample == 16 {
                    // Already 16-bit PCM
                    for chunk in buffer_slice.chunks_exact(2) {
                        let sample = i16::from_le_bytes([chunk[0], chunk[1]]);
                        samples.push(sample);
                    }
                } else if bits_per_sample == 32 {
                    // Convert 32-bit float to 16-bit PCM
                    for chunk in buffer_slice.chunks_exact(4) {
                        let float_val =
                            f32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
                        let sample = (float_val.clamp(-1.0, 1.0) * 32767.0) as i16;
                        samples.push(sample);
                    }
                }

                // If mono, duplicate to stereo
                if channels == 1 {
                    let mut stereo_samples = Vec::with_capacity(samples.len() * 2);
                    for sample in samples {
                        stereo_samples.push(sample);
                        stereo_samples.push(sample);
                    }
                    samples = stereo_samples;
                }

                let audio_sample = AudioSample {
                    data: samples,
                    timestamp,
                    sample_rate,
                    channels: 2, // Always output stereo
                };

                audio_queue.push(audio_sample);
            }

            let _ = capture_client.ReleaseBuffer(num_frames);

            current_packet_length = match capture_client.GetNextPacketSize() {
                Ok(len) => len,
                Err(_) => break,
            };
        }
    }

    audio_client
        .Stop()
        .map_err(|e| format!("Stop failed: {}", e))?;

    crate::log_debug("Audio capture loop stopped");
    Ok(())
}
