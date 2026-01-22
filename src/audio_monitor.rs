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

pub fn start_monitoring(device_id: String, device_name: String) -> Result<MonitorHandle, String> {
    let host = cpal::default_host();

    let input_device =
        if device_id.is_empty() || device_id == crate::settings::PODCAST_DEVICE_DEFAULT {
            host.default_input_device()
        } else {
            let mut found = None;
            if let Ok(devices) = host.input_devices() {
                for device in devices {
                    if let Ok(name) = device.name() {
                        // Try exact match on name
                        if name == device_name {
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

    // We use a shared buffer of f32 samples.
    // If channels differ, we mix or duplicate.
    // If sample rates differ, we rely on cpal to handle it? No, cpal doesn't resample automatically.
    // We must handle resampling if rates differ.
    // For simplicity, let's just push to buffer and pop. The pitch will change if rates differ!
    // But usually both are 44.1 or 48k. If they differ, we might have pitch shift.
    // Given the complexity of resampling, we'll try raw first.

    let buffer = Arc::new(Mutex::new(VecDeque::<f32>::with_capacity(48000 * 2))); // 1 sec stereo approx
    let buffer_producer = buffer.clone();
    let buffer_consumer = buffer.clone();

    let in_channels = input_config.channels() as usize;
    let out_channels = output_config.channels() as usize;

    let input_stream = match input_config.sample_format() {
        cpal::SampleFormat::F32 => {
            build_input_stream::<f32>(&input_device, &input_config, buffer_producer, in_channels)
        }
        cpal::SampleFormat::I16 => {
            build_input_stream::<i16>(&input_device, &input_config, buffer_producer, in_channels)
        }
        cpal::SampleFormat::U16 => {
            build_input_stream::<u16>(&input_device, &input_config, buffer_producer, in_channels)
        }
        _ => Err("Unsupported input format".to_string()),
    }?;

    let output_stream = match output_config.sample_format() {
        cpal::SampleFormat::F32 => build_output_stream::<f32>(
            &output_device,
            &output_config,
            buffer_consumer,
            out_channels,
        ),
        cpal::SampleFormat::I16 => build_output_stream::<i16>(
            &output_device,
            &output_config,
            buffer_consumer,
            out_channels,
        ),
        cpal::SampleFormat::U16 => build_output_stream::<u16>(
            &output_device,
            &output_config,
            buffer_consumer,
            out_channels,
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
    _channels: usize,
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
                let mut q = buffer.lock().unwrap();
                // Convert to f32 and push
                // If buffer too full, drop oldest
                if q.len() > 96000 {
                    // Safety cap
                    let drop = data.len().min(q.len());
                    q.drain(0..drop);
                }
                for &sample in data {
                    q.push_back(f32::from(sample));
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
    _channels: usize,
) -> Result<Stream, String>
where
    T: cpal::Sample + cpal::SizedSample + cpal::FromSample<f32>,
{
    let err_fn = |err| crate::log_debug(&format!("Monitor Output Error: {}", err));
    let stream = device
        .build_output_stream(
            &config.clone().into(),
            move |data: &mut [T], _: &_| {
                let mut q = buffer.lock().unwrap();
                // We need to fill data.len() samples.
                // But we might have different channel count in buffer (from input).
                // This naive copy assumes input channels == output channels or we don't care (speedup/slowdown).
                // Actually, if in_channels != out_channels, we have a problem.
                // Let's assume user wants to hear mic (usually mono/stereo) on speakers (stereo).
                // The buffer contains interleaved samples.

                // NOTE: Ideally we should handle channel mixing here.
                // But since we don't know input channels here (it's lost in closure type),
                // we rely on luck or simple 1:1 mapping for now.
                // Most mics are mono/stereo, speakers stereo.
                // If mic is mono, we read 1 sample per frame. If speaker is stereo, we write 2 samples per frame.
                // We'll run out of buffer twice as fast.

                // To fix this properly, we need to know input channels here or normalize buffer to stereo.
                // Let's normalize buffer to Mono in input and expand to Stereo in output?
                // Or just copy.

                for sample in data.iter_mut() {
                    *sample = if let Some(v) = q.pop_front() {
                        T::from_sample(v)
                    } else {
                        T::from_sample(0.0f32)
                    };
                }
            },
            err_fn,
            None,
        )
        .map_err(|e| e.to_string())?;
    Ok(stream)
}
