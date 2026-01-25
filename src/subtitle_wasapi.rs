//! Audio decoding and resampling utilities for subtitle playback.

/// Decode MP3 bytes to PCM f32 samples using rodio/symphonia.
pub fn decode_mp3_to_pcm(mp3_data: &[u8]) -> Result<(Vec<f32>, u32, u16), String> {
    use rodio::{Decoder, Source};
    use std::io::Cursor;

    let cursor = Cursor::new(mp3_data.to_vec());
    let decoder = Decoder::new(cursor).map_err(|e| format!("Decoder error: {}", e))?;

    let sample_rate = decoder.sample_rate();
    let channels = decoder.channels();

    let samples: Vec<f32> = decoder.collect();

    Ok((samples, sample_rate, channels))
}

/// Resample PCM data to target sample rate and channels using linear interpolation.
pub fn resample_pcm(
    samples: &[f32],
    source_rate: u32,
    source_channels: u16,
    target_rate: u32,
    target_channels: u16,
) -> Vec<f32> {
    if source_rate == target_rate && source_channels == target_channels {
        return samples.to_vec();
    }

    let src_ch = source_channels as usize;
    let tgt_ch = target_channels as usize;
    let src_frames = samples.len() / src_ch;
    let ratio = source_rate as f64 / target_rate as f64;
    let tgt_frames = (src_frames as f64 / ratio).ceil() as usize;

    let mut output = Vec::with_capacity(tgt_frames * tgt_ch);

    for tgt_frame in 0..tgt_frames {
        let src_pos = tgt_frame as f64 * ratio;
        let src_frame = src_pos.floor() as usize;
        let frac = src_pos.fract() as f32;

        for tgt_c in 0..tgt_ch {
            let src_c = tgt_c % src_ch;

            let sample = if src_frame + 1 < src_frames {
                let s0 = samples[src_frame * src_ch + src_c];
                let s1 = samples[(src_frame + 1) * src_ch + src_c];
                s0 + (s1 - s0) * frac
            } else if src_frame < src_frames {
                samples[src_frame * src_ch + src_c]
            } else {
                0.0
            };

            output.push(sample);
        }
    }

    output
}
