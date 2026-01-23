use cpal::traits::{DeviceTrait, HostTrait, StreamTrait};
use cpal::{Stream, SupportedStreamConfig};
use std::collections::VecDeque;
use std::sync::{Arc, Mutex};

pub struct MonitorHandle {
    _input_stream: Stream,
    _output_stream: Stream,
}

impl MonitorHandle {
    pub fn stop(self) {
        // Dropping the streams stops them
    }
}

/// Starts audio monitoring with the specified device and gain.
///
/// # Arguments
/// * `device_id` - The device ID to use (or PODCAST_DEVICE_DEFAULT for default)
/// * `device_name` - The device name to match (used as fallback)
/// * `gain` - Volume multiplier (1.0 = normal, 0.5 = half, 2.0 = double)
pub fn start_monitoring(
    device_id: String,
    device_name: String,
    gain: f32,
) -> Result<MonitorHandle, String> {
    let host = cpal::default_host();

    let input_device =
        if device_id.is_empty() || device_id == crate::settings::PODCAST_DEVICE_DEFAULT {
            host.default_input_device()
        } else {
            let mut found = None;
            if let Ok(devices) = host.input_devices() {
                for device in devices {
                    if let Ok(desc) = device.description() {
                        // Try exact match on description (cpal uses description as identifier)
                        if desc.name() == device_name {
                            found = Some(device);
                            break;
                        }
                    }
                }
            }
            found.or_else(|| host.default_input_device())
        }
        .ok_or("No input device available")?;

    let output_device = host
        .default_output_device()
        .ok_or("No output device available")?;

    let input_config = input_device
        .default_input_config()
        .map_err(|e| format!("Failed to get input config: {}", e))?;
    let output_config = output_device
        .default_output_config()
        .map_err(|e| format!("Failed to get output config: {}", e))?;

    let in_sample_rate = input_config.sample_rate();
    let out_sample_rate = output_config.sample_rate();
    let in_channels = input_config.channels() as usize;
    let out_channels = output_config.channels() as usize;

    // Buffer holds mono samples at input sample rate, will be resampled and expanded to stereo on output
    let buffer = Arc::new(Mutex::new(VecDeque::<f32>::with_capacity(48000 * 2)));
    let buffer_producer = buffer.clone();
    let buffer_consumer = buffer.clone();

    let input_stream = match input_config.sample_format() {
        cpal::SampleFormat::F32 => build_input_stream::<f32>(
            &input_device,
            &input_config,
            buffer_producer,
            in_channels,
            gain,
        ),
        cpal::SampleFormat::I16 => build_input_stream::<i16>(
            &input_device,
            &input_config,
            buffer_producer,
            in_channels,
            gain,
        ),
        cpal::SampleFormat::U16 => build_input_stream::<u16>(
            &input_device,
            &input_config,
            buffer_producer,
            in_channels,
            gain,
        ),
        _ => Err("Unsupported input format".to_string()),
    }?;

    let output_stream = match output_config.sample_format() {
        cpal::SampleFormat::F32 => build_output_stream::<f32>(
            &output_device,
            &output_config,
            buffer_consumer,
            out_channels,
            in_sample_rate,
            out_sample_rate,
        ),
        cpal::SampleFormat::I16 => build_output_stream::<i16>(
            &output_device,
            &output_config,
            buffer_consumer,
            out_channels,
            in_sample_rate,
            out_sample_rate,
        ),
        cpal::SampleFormat::U16 => build_output_stream::<u16>(
            &output_device,
            &output_config,
            buffer_consumer,
            out_channels,
            in_sample_rate,
            out_sample_rate,
        ),
        _ => Err("Unsupported output format".to_string()),
    }?;

    input_stream
        .play()
        .map_err(|e| format!("Failed to play input stream: {}", e))?;
    output_stream
        .play()
        .map_err(|e| format!("Failed to play output stream: {}", e))?;

    Ok(MonitorHandle {
        _input_stream: input_stream,
        _output_stream: output_stream,
    })
}

fn build_input_stream<T>(
    device: &cpal::Device,
    config: &SupportedStreamConfig,
    buffer: Arc<Mutex<VecDeque<f32>>>,
    channels: usize,
    gain: f32,
) -> Result<Stream, String>
where
    T: cpal::Sample + cpal::SizedSample,
    f32: From<T>,
{
    let err_fn = |err| crate::log_debug(&format!("Monitor Input Error: {}", err));
    let stream = device
        .build_input_stream(
            &config.clone().into(),
            move |data: &[T], _: &_| {
                let mut q = buffer.lock().unwrap_or_else(|e| e.into_inner());
                // If buffer too full, drop oldest to prevent memory growth
                if q.len() > 96000 {
                    let drop = data.len().min(q.len());
                    q.drain(0..drop);
                }
                // Convert to mono f32 with gain, mix channels if stereo
                for frame in data.chunks(channels) {
                    let mono: f32 = if channels == 1 {
                        f32::from(frame[0])
                    } else {
                        // Mix all channels to mono
                        let sum: f32 = frame.iter().map(|&s| f32::from(s)).sum();
                        sum / channels as f32
                    };
                    // Apply gain and clamp to avoid clipping
                    let gained = (mono * gain).clamp(-1.0, 1.0);
                    q.push_back(gained);
                }
            },
            err_fn,
            None,
        )
        .map_err(|e| e.to_string())?;
    Ok(stream)
}

fn build_output_stream<T>(
    device: &cpal::Device,
    config: &SupportedStreamConfig,
    buffer: Arc<Mutex<VecDeque<f32>>>,
    out_channels: usize,
    in_sample_rate: u32,
    out_sample_rate: u32,
) -> Result<Stream, String>
where
    T: cpal::Sample + cpal::SizedSample + cpal::FromSample<f32>,
{
    let err_fn = |err| crate::log_debug(&format!("Monitor Output Error: {}", err));

    // Calculate resampling ratio (how many input samples per output sample)
    let resample_ratio = in_sample_rate as f64 / out_sample_rate as f64;

    // State for linear interpolation resampling
    let resample_state = Arc::new(Mutex::new(ResampleState {
        position: 0.0,
        last_sample: 0.0,
    }));

    let stream = device
        .build_output_stream(
            &config.clone().into(),
            move |data: &mut [T], _: &_| {
                let mut q = buffer.lock().unwrap_or_else(|e| e.into_inner());
                let mut rs = resample_state.lock().unwrap_or_else(|e| e.into_inner());

                // Buffer contains mono samples at input rate
                // We need to output stereo samples at output rate
                let frames_needed = data.len() / out_channels;

                for frame_idx in 0..frames_needed {
                    // Get the resampled mono sample
                    let mono_sample =
                        if (in_sample_rate == out_sample_rate) || resample_ratio == 1.0 {
                            // No resampling needed
                            q.pop_front().unwrap_or(0.0)
                        } else {
                            // Linear interpolation resampling
                            resample_sample(&mut q, &mut rs, resample_ratio)
                        };

                    // Write mono sample to all output channels (expand to stereo)
                    for ch in 0..out_channels {
                        data[frame_idx * out_channels + ch] = T::from_sample(mono_sample);
                    }
                }
            },
            err_fn,
            None,
        )
        .map_err(|e| e.to_string())?;
    Ok(stream)
}

struct ResampleState {
    position: f64,
    last_sample: f32,
}

/// Simple linear interpolation resampling
fn resample_sample(buffer: &mut VecDeque<f32>, state: &mut ResampleState, ratio: f64) -> f32 {
    // Advance position by the ratio
    state.position += ratio;

    // Consume samples from buffer as needed
    while state.position >= 1.0 {
        state.last_sample = buffer.pop_front().unwrap_or(state.last_sample);
        state.position -= 1.0;
    }

    // Get next sample for interpolation
    let next_sample = buffer.front().copied().unwrap_or(state.last_sample);

    // Linear interpolation between last and next sample
    let t = state.position as f32;
    state.last_sample * (1.0 - t) + next_sample * t
}
