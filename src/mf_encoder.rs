#![allow(clippy::seek_from_current)]
use crate::accessibility::to_wide;
use crate::video_recorder::VideoFrame;
use std::ffi::c_void;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;
use windows::Win32::Graphics::Direct3D11::{
    D3D11_MAP_READ, D3D11_MAPPED_SUBRESOURCE, ID3D11Texture2D,
};
use windows::Win32::Media::MediaFoundation::{
    IMFMediaBuffer, IMFMediaType, IMFSample, IMFSinkWriter, IMFSourceReader,
    MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT,
    MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_AVG_BITRATE,
    MF_MT_FIXED_SIZE_SAMPLES, MF_MT_FRAME_RATE, MF_MT_FRAME_SIZE, MF_MT_INTERLACE_MODE,
    MF_MT_MAJOR_TYPE, MF_MT_PIXEL_ASPECT_RATIO, MF_MT_SAMPLE_SIZE, MF_MT_SUBTYPE,
    MF_SOURCE_READER_FIRST_AUDIO_STREAM, MF_SOURCE_READERF_ENDOFSTREAM, MF_VERSION,
    MFAudioFormat_AAC, MFAudioFormat_MP3, MFAudioFormat_PCM, MFCreateMediaType,
    MFCreateMemoryBuffer, MFCreateSample, MFCreateSinkWriterFromURL, MFCreateSourceReaderFromURL,
    MFMediaType_Audio, MFMediaType_Video, MFShutdown, MFStartup, MFVideoFormat_H264,
    MFVideoFormat_RGB32,
};
use windows::core::PCWSTR;

struct MfGuard;

impl MfGuard {
    fn start() -> Result<Self, String> {
        unsafe {
            if let Err(e) = MFStartup(MF_VERSION, 0) {
                return Err(format!(
                    "Media Foundation not available. Install Media Feature Pack on Windows N/KN. ({})",
                    e
                ));
            }
        }
        Ok(MfGuard)
    }
}

impl Drop for MfGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = MFShutdown();
        }
    }
}

pub struct Mp3StreamWriter {
    _guard: MfGuard,
    writer: IMFSinkWriter,
    stream_index: u32,
    sample_time: i64,
    sample_rate: u32,
    bytes_per_frame: u32,
}

impl Mp3StreamWriter {
    pub fn create(
        mp3_path: &Path,
        bitrate_kbps: u32,
        sample_rate: u32,
        channels: u16,
    ) -> Result<Self, String> {
        unsafe {
            let bitrate_kbps = match bitrate_kbps {
                192 => 192,
                256 => 256,
                _ => 128,
            };
            crate::log_debug(&format!(
                "MF: streaming mp3 writer. mp3={:?} bitrate_kbps={} rate={} ch={}",
                mp3_path, bitrate_kbps, sample_rate, channels
            ));
            let guard = MfGuard::start()?;

            let mp3_wide = to_wide(mp3_path.to_str().ok_or("Invalid mp3 path")?);
            let writer: IMFSinkWriter =
                MFCreateSinkWriterFromURL(PCWSTR(mp3_wide.as_ptr()), None, None)
                    .map_err(|e| format!("MFCreateSinkWriterFromURL failed: {}", e))?;

            let pcm_type: IMFMediaType = MFCreateMediaType()
                .map_err(|e| format!("MFCreateMediaType (pcm) failed: {}", e))?;
            pcm_type
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
                .map_err(|e| format!("SetGUID major type failed: {}", e))?;
            pcm_type
                .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)
                .map_err(|e| format!("SetGUID subtype PCM failed: {}", e))?;
            let requested_bits = 16u32;
            let requested_channels = channels as u32;
            let block_align = requested_channels * (requested_bits / 8);
            let avg_bytes = sample_rate * block_align;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, sample_rate)
                .map_err(|e| format!("Set sample rate failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, requested_channels)
                .map_err(|e| format!("Set channels failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, requested_bits)
                .map_err(|e| format!("Set bits failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, block_align)
                .map_err(|e| format!("Set block alignment failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, avg_bytes)
                .map_err(|e| format!("Set avg bytes failed: {}", e))?;
            let _ = pcm_type.SetUINT32(&MF_MT_FIXED_SIZE_SAMPLES, 1);
            let _ = pcm_type.SetUINT32(&MF_MT_SAMPLE_SIZE, block_align);

            let out_type: IMFMediaType = MFCreateMediaType()
                .map_err(|e| format!("MFCreateMediaType (mp3) failed: {}", e))?;
            out_type
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
                .map_err(|e| format!("SetGUID major type (out) failed: {}", e))?;
            out_type
                .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_MP3)
                .map_err(|e| format!("SetGUID subtype MP3 failed: {}", e))?;
            out_type
                .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, requested_channels)
                .map_err(|e| format!("Set channels (out) failed: {}", e))?;
            out_type
                .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, sample_rate)
                .map_err(|e| format!("Set sample rate (out) failed: {}", e))?;
            let mp3_avg_bytes = (bitrate_kbps * 1000) / 8;
            out_type
                .SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, mp3_avg_bytes)
                .map_err(|e| format!("Set mp3 bitrate failed: {}", e))?;

            let stream_index = writer
                .AddStream(&out_type)
                .map_err(|e| format!("SinkWriter AddStream failed: {}", e))?;
            if let Err(e) = writer.SetInputMediaType(stream_index, &pcm_type, None) {
                crate::log_debug(&format!("MF: SetInputMediaType failed: {}", e));
                return Err(format!("SinkWriter SetInputMediaType failed: {}", e));
            }
            writer
                .BeginWriting()
                .map_err(|e| format!("SinkWriter BeginWriting failed: {}", e))?;

            Ok(Mp3StreamWriter {
                _guard: guard,
                writer,
                stream_index,
                sample_time: 0,
                sample_rate,
                bytes_per_frame: block_align,
            })
        }
    }

    pub fn write_i16(&mut self, samples: &[i16]) -> Result<(), String> {
        if samples.is_empty() {
            return Ok(());
        }
        let byte_len = (samples.len() * 2) as u32;
        let frames = byte_len / self.bytes_per_frame;
        if frames == 0 {
            return Ok(());
        }
        let duration = (frames as i64 * 10_000_000i64) / self.sample_rate as i64;
        unsafe {
            let buffer: IMFMediaBuffer = MFCreateMemoryBuffer(byte_len)
                .map_err(|e| format!("MFCreateMemoryBuffer failed: {}", e))?;
            let mut data_ptr = std::ptr::null_mut();
            let mut max_len = 0u32;
            buffer
                .Lock(&mut data_ptr, Some(&mut max_len), None)
                .map_err(|e| format!("IMFMediaBuffer::Lock failed: {}", e))?;
            if !data_ptr.is_null() {
                std::ptr::copy_nonoverlapping(
                    samples.as_ptr() as *const u8,
                    data_ptr,
                    byte_len as usize,
                );
            }
            buffer
                .Unlock()
                .map_err(|e| format!("IMFMediaBuffer::Unlock failed: {}", e))?;
            buffer
                .SetCurrentLength(byte_len)
                .map_err(|e| format!("IMFMediaBuffer::SetCurrentLength failed: {}", e))?;

            let sample: IMFSample =
                MFCreateSample().map_err(|e| format!("MFCreateSample failed: {}", e))?;
            sample
                .AddBuffer(&buffer)
                .map_err(|e| format!("IMFSample::AddBuffer failed: {}", e))?;
            sample
                .SetSampleTime(self.sample_time)
                .map_err(|e| format!("IMFSample::SetSampleTime failed: {}", e))?;
            sample
                .SetSampleDuration(duration)
                .map_err(|e| format!("IMFSample::SetSampleDuration failed: {}", e))?;

            self.writer
                .WriteSample(self.stream_index, &sample)
                .map_err(|e| format!("WriteSample failed: {}", e))?;
        }
        self.sample_time = self.sample_time.saturating_add(duration);
        Ok(())
    }

    pub fn finalize(self) -> Result<(), String> {
        unsafe {
            self.writer
                .Finalize()
                .map_err(|e| format!("SinkWriter Finalize failed: {}", e))?;
        }
        Ok(())
    }
}

fn read_wav_data_info(path: &Path) -> Result<(u64, u32, i16), String> {
    let mut file = File::open(path).map_err(|e| e.to_string())?;
    let mut riff_header = [0u8; 12];
    file.read_exact(&mut riff_header)
        .map_err(|e| e.to_string())?;
    if &riff_header[0..4] != b"RIFF" || &riff_header[8..12] != b"WAVE" {
        return Err("Invalid WAV header".to_string());
    }

    loop {
        let mut chunk_header = [0u8; 8];
        if file.read_exact(&mut chunk_header).is_err() {
            break;
        }
        let chunk_id = &chunk_header[0..4];
        let chunk_size = u32::from_le_bytes(chunk_header[4..8].try_into().unwrap());

        if chunk_id == b"data" {
            let data_offset = file.seek(SeekFrom::Current(0)).map_err(|e| e.to_string())?;
            return Ok((data_offset, chunk_size, 0));
        } else {
            file.seek(SeekFrom::Current(chunk_size as i64))
                .map_err(|e| e.to_string())?;
        }

        if chunk_size % 2 == 1 {
            file.seek(SeekFrom::Current(1)).map_err(|e| e.to_string())?;
        }
    }
    Err("WAV data chunk not found".to_string())
}

pub fn encode_wav_to_mp3(wav_path: &Path, mp3_path: &Path) -> Result<(), String> {
    encode_wav_to_mp3_with_bitrate(wav_path, mp3_path, 128)
}

pub fn encode_wav_to_mp3_with_bitrate(
    wav_path: &Path,
    mp3_path: &Path,
    bitrate_kbps: u32,
) -> Result<(), String> {
    encode_wav_to_mp3_with_bitrate_progress(wav_path, mp3_path, bitrate_kbps, |_| {}, None)
}

pub fn encode_wav_to_mp3_with_bitrate_progress<F>(
    wav_path: &Path,
    mp3_path: &Path,
    bitrate_kbps: u32,
    mut progress: F,
    cancel: Option<&std::sync::atomic::AtomicBool>,
) -> Result<(), String>
where
    F: FnMut(u32),
{
    unsafe {
        let bitrate_kbps = match bitrate_kbps {
            192 => 192,
            256 => 256,
            _ => 128,
        };
        crate::log_debug(&format!(
            "MF: encode wav to mp3. wav={:?} mp3={:?} bitrate_kbps={}",
            wav_path, mp3_path, bitrate_kbps
        ));
        let _guard = MfGuard::start()?;

        let wav_wide = to_wide(wav_path.to_str().ok_or("Invalid wav path")?);
        let mp3_wide = to_wide(mp3_path.to_str().ok_or("Invalid mp3 path")?);

        let reader: IMFSourceReader = MFCreateSourceReaderFromURL(PCWSTR(wav_wide.as_ptr()), None)
            .map_err(|e| format!("MFCreateSourceReaderFromURL failed: {}", e))?;

        let pcm_type: IMFMediaType =
            MFCreateMediaType().map_err(|e| format!("MFCreateMediaType (pcm) failed: {}", e))?;
        pcm_type
            .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
            .map_err(|e| format!("SetGUID major type failed: {}", e))?;
        pcm_type
            .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)
            .map_err(|e| format!("SetGUID subtype PCM failed: {}", e))?;
        let requested_rate = 44100u32;
        let requested_channels = 2u32;
        let requested_bits = 16u32;
        let requested_block_align = requested_channels * (requested_bits / 8);
        let requested_avg_bytes = requested_rate * requested_block_align;
        let _ = pcm_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, requested_rate);
        let _ = pcm_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, requested_channels);
        let _ = pcm_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, requested_bits);
        let _ = pcm_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, requested_block_align);
        let _ = pcm_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, requested_avg_bytes);
        reader
            .SetCurrentMediaType(
                MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32,
                None,
                &pcm_type,
            )
            .map_err(|e| format!("SetCurrentMediaType failed: {}", e))?;

        let in_type = reader
            .GetCurrentMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32)
            .map_err(|e| format!("GetCurrentMediaType failed: {}", e))?;

        let mut data_size = 0u64;
        if let Ok((data_offset, size, peak)) = read_wav_data_info(wav_path) {
            data_size = size as u64;
            crate::log_debug(&format!(
                "MF: wav data offset={} size={} peak={}",
                data_offset, size, peak
            ));
        }
        let mut sample_rate = 0u32;
        let mut channels = 0u32;
        let mut bits_per_sample = 0u32;
        let mut block_align = 0u32;
        let mut avg_bytes_in = 0u32;
        if let Ok(val) = in_type.GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND) {
            sample_rate = val;
        }
        if let Ok(val) = in_type.GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS) {
            channels = val;
        }
        if let Ok(val) = in_type.GetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE) {
            bits_per_sample = val;
        }
        if let Ok(val) = in_type.GetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT) {
            block_align = val;
        }
        if let Ok(val) = in_type.GetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND) {
            avg_bytes_in = val;
        }

        crate::log_debug(&format!(
            "MF: input wfx rate={} ch={} bits={} block_align={} avg_bytes={}",
            sample_rate, channels, bits_per_sample, block_align, avg_bytes_in
        ));
        crate::log_debug(&format!(
            "MF: requested rate={} ch={} bits={}",
            requested_rate, requested_channels, requested_bits
        ));
        if sample_rate == 0 || channels == 0 {
            return Err("MF: invalid input audio format".to_string());
        }

        let input_type = in_type;
        let _ = input_type.SetUINT32(&MF_MT_FIXED_SIZE_SAMPLES, 1);
        if block_align != 0 {
            let _ = input_type.SetUINT32(&MF_MT_SAMPLE_SIZE, block_align);
        }

        let out_type: IMFMediaType =
            MFCreateMediaType().map_err(|e| format!("MFCreateMediaType (mp3) failed: {}", e))?;
        out_type
            .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
            .map_err(|e| format!("SetGUID major type (out) failed: {}", e))?;
        out_type
            .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_MP3)
            .map_err(|e| format!("SetGUID subtype MP3 failed: {}", e))?;
        out_type
            .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, requested_channels)
            .map_err(|e| format!("Set channels failed: {}", e))?;
        out_type
            .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, requested_rate)
            .map_err(|e| format!("Set sample rate failed: {}", e))?;
        let mp3_avg_bytes = (bitrate_kbps * 1000) / 8;
        out_type
            .SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, mp3_avg_bytes)
            .map_err(|e| format!("Set mp3 bitrate failed: {}", e))?;
        crate::log_debug(&format!(
            "MF: output mp3 rate={} ch={} avg_bytes={}",
            requested_rate, requested_channels, mp3_avg_bytes
        ));

        let writer: IMFSinkWriter =
            MFCreateSinkWriterFromURL(PCWSTR(mp3_wide.as_ptr()), None, None)
                .map_err(|e| format!("MFCreateSinkWriterFromURL failed: {}", e))?;
        let stream_index = writer
            .AddStream(&out_type)
            .map_err(|e| format!("SinkWriter AddStream failed: {}", e))?;
        if let Err(e) = writer.SetInputMediaType(stream_index, &input_type, None) {
            crate::log_debug(&format!("MF: SetInputMediaType failed: {}", e));
            return Err(format!("SinkWriter SetInputMediaType failed: {}", e));
        }
        writer
            .BeginWriting()
            .map_err(|e| format!("SinkWriter BeginWriting failed: {}", e))?;

        let mut sample_count: u64 = 0;
        let mut total_bytes: u64 = 0;
        let mut last_pct: u32 = 0;
        loop {
            if let Some(cancel) = cancel {
                if cancel.load(std::sync::atomic::Ordering::Relaxed) {
                    return Err("Saving canceled.".to_string());
                }
            }
            let mut read_stream = 0u32;
            let mut flags = 0u32;
            let mut _timestamp = 0i64;
            let mut sample = None;
            reader
                .ReadSample(
                    MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32,
                    0,
                    Some(&mut read_stream),
                    Some(&mut flags),
                    Some(&mut _timestamp),
                    Some(&mut sample),
                )
                .map_err(|e| format!("ReadSample failed: {}", e))?;

            if flags & (MF_SOURCE_READERF_ENDOFSTREAM.0 as u32) != 0 {
                break;
            }
            if let Some(sample) = sample {
                sample_count = sample_count.saturating_add(1);
                if let Ok(len) = sample.GetTotalLength() {
                    total_bytes = total_bytes.saturating_add(len as u64);
                }
                writer
                    .WriteSample(stream_index, &sample)
                    .map_err(|e| format!("WriteSample failed: {}", e))?;
                if data_size > 0 {
                    let pct = ((total_bytes.saturating_mul(100)) / data_size).min(100) as u32;
                    if pct > last_pct {
                        last_pct = pct;
                        progress(pct);
                    }
                }
            }
        }

        if let Some(cancel) = cancel {
            if cancel.load(std::sync::atomic::Ordering::Relaxed) {
                return Err("Saving canceled.".to_string());
            }
        }
        if last_pct < 100 {
            progress(100);
        }
        writer
            .Finalize()
            .map_err(|e| format!("SinkWriter Finalize failed: {}", e))?;
        crate::log_debug(&format!(
            "MF: samples_written={} total_bytes={}",
            sample_count, total_bytes
        ));
        if let Ok(size) = std::fs::metadata(mp3_path).map(|m| m.len()) {
            crate::log_debug(&format!("MF: encode completed. mp3_size={}", size));
        } else {
            crate::log_debug("MF: encode completed.");
        }
        Ok(())
    }
}

/// MP4 stream writer for video (H.264) + audio (AAC) recording
pub struct Mp4StreamWriter {
    _guard: MfGuard,
    writer: IMFSinkWriter,
    video_stream_index: u32,
    audio_stream_index: u32,
    audio_sample_rate: u32,
    audio_timestamp: i64, // Current audio timestamp in 100-nanosecond units
    staging_texture: Option<ID3D11Texture2D>, // Reusable staging texture for video frames
    staging_width: u32,
    staging_height: u32,
}

// SAFETY: IMFSinkWriter is thread-safe for writing from a single thread
unsafe impl Send for Mp4StreamWriter {}

impl Mp4StreamWriter {
    pub fn create(
        path: &Path,
        width: u32,
        height: u32,
        audio_sample_rate: u32,
    ) -> Result<Self, String> {
        unsafe {
            crate::log_debug(&format!(
                "MF: creating MP4 writer. path={:?} size={}x{} audio={}Hz",
                path, width, height, audio_sample_rate
            ));

            let guard = MfGuard::start()?;

            let mp4_wide = to_wide(path.to_str().ok_or("Invalid MP4 path")?);
            let writer: IMFSinkWriter =
                MFCreateSinkWriterFromURL(PCWSTR(mp4_wide.as_ptr()), None, None)
                    .map_err(|e| format!("MFCreateSinkWriterFromURL failed: {}", e))?;

            // Configure video stream (H.264)
            let video_out = MFCreateMediaType()
                .map_err(|e| format!("MFCreateMediaType (video_out) failed: {}", e))?;
            video_out
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Video)
                .map_err(|e| format!("SetGUID video major type failed: {}", e))?;
            video_out
                .SetGUID(&MF_MT_SUBTYPE, &MFVideoFormat_H264)
                .map_err(|e| format!("SetGUID H264 subtype failed: {}", e))?;
            video_out
                .SetUINT64(&MF_MT_FRAME_SIZE, ((width as u64) << 32) | (height as u64))
                .map_err(|e| format!("SetUINT64 frame size failed: {}", e))?;
            video_out
                .SetUINT64(&MF_MT_FRAME_RATE, (30u64 << 32) | 1)
                .map_err(|e| format!("SetUINT64 frame rate failed: {}", e))?; // 30 FPS
            video_out
                .SetUINT32(&MF_MT_AVG_BITRATE, 2_000_000)
                .map_err(|e| format!("SetUINT32 avg bitrate failed: {}", e))?; // 2 Mbps (reduced for faster encoding)
            video_out
                .SetUINT32(&MF_MT_INTERLACE_MODE, 2)
                .map_err(|e| format!("SetUINT32 interlace mode failed: {}", e))?; // Progressive
            video_out
                .SetUINT64(&MF_MT_PIXEL_ASPECT_RATIO, (1u64 << 32) | 1)
                .map_err(|e| format!("SetUINT64 pixel aspect ratio failed: {}", e))?; // 1:1

            let video_in = MFCreateMediaType()
                .map_err(|e| format!("MFCreateMediaType (video_in) failed: {}", e))?;
            video_in
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Video)
                .map_err(|e| format!("SetGUID video major type (in) failed: {}", e))?;
            video_in
                .SetGUID(&MF_MT_SUBTYPE, &MFVideoFormat_RGB32)
                .map_err(|e| format!("SetGUID RGB32 subtype failed: {}", e))?; // BGRA from capture
            video_in
                .SetUINT64(&MF_MT_FRAME_SIZE, ((width as u64) << 32) | (height as u64))
                .map_err(|e| format!("SetUINT64 frame size (in) failed: {}", e))?;
            video_in
                .SetUINT64(&MF_MT_FRAME_RATE, (30u64 << 32) | 1)
                .map_err(|e| format!("SetUINT64 frame rate (in) failed: {}", e))?;

            let video_stream_index = writer
                .AddStream(&video_out)
                .map_err(|e| format!("AddStream video failed: {}", e))?;
            writer
                .SetInputMediaType(video_stream_index, &video_in, None)
                .map_err(|e| format!("SetInputMediaType video failed: {}", e))?;

            // Configure audio stream (AAC)
            let audio_out = MFCreateMediaType()
                .map_err(|e| format!("MFCreateMediaType (audio_out) failed: {}", e))?;
            audio_out
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
                .map_err(|e| format!("SetGUID audio major type failed: {}", e))?;
            audio_out
                .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)
                .map_err(|e| format!("SetGUID AAC subtype failed: {}", e))?;
            audio_out
                .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, 2)
                .map_err(|e| format!("SetUINT32 audio channels failed: {}", e))?;
            audio_out
                .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, audio_sample_rate)
                .map_err(|e| format!("SetUINT32 audio sample rate failed: {}", e))?;
            audio_out
                .SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, 24000)
                .map_err(|e| format!("SetUINT32 audio avg bytes failed: {}", e))?; // 192 kbps

            let audio_in = MFCreateMediaType()
                .map_err(|e| format!("MFCreateMediaType (audio_in) failed: {}", e))?;
            audio_in
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
                .map_err(|e| format!("SetGUID audio major type (in) failed: {}", e))?;
            audio_in
                .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)
                .map_err(|e| format!("SetGUID PCM subtype failed: {}", e))?;
            audio_in
                .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, audio_sample_rate)
                .map_err(|e| format!("SetUINT32 audio sample rate (in) failed: {}", e))?;
            audio_in
                .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, 2)
                .map_err(|e| format!("SetUINT32 audio channels (in) failed: {}", e))?;
            audio_in
                .SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)
                .map_err(|e| format!("SetUINT32 audio bits per sample failed: {}", e))?;
            audio_in
                .SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, 4)
                .map_err(|e| format!("SetUINT32 audio block alignment failed: {}", e))?;

            let audio_avg_bytes_in = audio_sample_rate * 4; // 16-bit stereo = 4 bytes per frame
            audio_in
                .SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, audio_avg_bytes_in)
                .map_err(|e| format!("SetUINT32 audio avg bytes (in) failed: {}", e))?;

            let audio_stream_index = writer
                .AddStream(&audio_out)
                .map_err(|e| format!("AddStream audio failed: {}", e))?;
            writer
                .SetInputMediaType(audio_stream_index, &audio_in, None)
                .map_err(|e| format!("SetInputMediaType audio failed: {}", e))?;

            writer
                .BeginWriting()
                .map_err(|e| format!("BeginWriting failed: {}", e))?;

            crate::log_debug("MP4 writer initialized successfully");

            Ok(Mp4StreamWriter {
                _guard: guard,
                writer,
                video_stream_index,
                audio_stream_index,
                audio_sample_rate,
                audio_timestamp: 0,
                staging_texture: None,
                staging_width: 0,
                staging_height: 0,
            })
        }
    }

    pub fn write_video_frame(&mut self, frame: &VideoFrame) -> Result<(), String> {
        unsafe {
            use windows::Win32::Graphics::Direct3D11::{
                D3D11_CPU_ACCESS_READ, D3D11_TEXTURE2D_DESC, D3D11_USAGE_STAGING, ID3D11Texture2D,
            };

            // Get D3D11 device and context
            let device = frame
                .texture
                .GetDevice()
                .map_err(|e| format!("GetDevice failed: {}", e))?;

            let context = device
                .GetImmediateContext()
                .map_err(|e| format!("GetImmediateContext failed: {}", e))?;

            // Check if we need to create/recreate staging texture
            let need_new_staging = self.staging_texture.is_none()
                || self.staging_width != frame.width
                || self.staging_height != frame.height;

            if need_new_staging {
                // Get source texture description
                let mut src_desc = D3D11_TEXTURE2D_DESC::default();
                frame.texture.GetDesc(&mut src_desc);

                // Create staging texture for CPU access
                let staging_desc = D3D11_TEXTURE2D_DESC {
                    Width: src_desc.Width,
                    Height: src_desc.Height,
                    MipLevels: 1,
                    ArraySize: 1,
                    Format: src_desc.Format,
                    SampleDesc: windows::Win32::Graphics::Dxgi::Common::DXGI_SAMPLE_DESC {
                        Count: 1,
                        Quality: 0,
                    },
                    Usage: D3D11_USAGE_STAGING,
                    BindFlags: 0,
                    CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32,
                    MiscFlags: Default::default(),
                };

                let mut staging_texture: Option<ID3D11Texture2D> = None;
                device
                    .CreateTexture2D(&staging_desc, None, Some(&mut staging_texture))
                    .map_err(|e| format!("CreateTexture2D (staging) failed: {}", e))?;

                self.staging_texture = staging_texture;
                self.staging_width = frame.width;
                self.staging_height = frame.height;

                crate::log_debug(&format!(
                    "Created reusable staging texture: {}x{}",
                    frame.width, frame.height
                ));
            }

            let staging_texture = self.staging_texture.as_ref().unwrap();

            // Copy from GPU texture to staging texture
            context.CopyResource(staging_texture, &frame.texture);

            // Map staging texture to CPU memory
            let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
            context
                .Map(staging_texture, 0, D3D11_MAP_READ, 0, Some(&mut mapped))
                .map_err(|e| format!("Map staging texture failed: {}", e))?;

            // Calculate buffer size
            let byte_count = (frame.width * frame.height * 4) as u32; // BGRA = 4 bytes per pixel

            // Create Media Foundation buffer
            let buffer = MFCreateMemoryBuffer(byte_count)
                .map_err(|e| format!("MFCreateMemoryBuffer failed: {}", e))?;

            let mut data_ptr: *mut c_void = std::ptr::null_mut();
            buffer
                .Lock(
                    &mut data_ptr as *mut *mut c_void as *mut *mut u8,
                    None,
                    None,
                )
                .map_err(|e| format!("Buffer Lock failed: {}", e))?;

            if !data_ptr.is_null() && !mapped.pData.is_null() {
                // Handle row pitch - mapped data might have padding
                let src_pitch = mapped.RowPitch as usize;
                let dst_pitch = (frame.width * 4) as usize;

                if src_pitch == dst_pitch {
                    // No padding, direct copy
                    std::ptr::copy_nonoverlapping(mapped.pData, data_ptr, byte_count as usize);
                } else {
                    // Copy row by row to handle padding
                    for row in 0..frame.height {
                        let src_offset = row as usize * src_pitch;
                        let dst_offset = row as usize * dst_pitch;
                        std::ptr::copy_nonoverlapping(
                            (mapped.pData as *const u8).add(src_offset),
                            (data_ptr as *mut u8).add(dst_offset),
                            dst_pitch,
                        );
                    }
                }
            }

            buffer
                .Unlock()
                .map_err(|e| format!("Buffer Unlock failed: {}", e))?;
            buffer
                .SetCurrentLength(byte_count)
                .map_err(|e| format!("SetCurrentLength failed: {}", e))?;

            context.Unmap(staging_texture, 0);

            // Create sample
            let sample = MFCreateSample().map_err(|e| format!("MFCreateSample failed: {}", e))?;
            sample
                .AddBuffer(&buffer)
                .map_err(|e| format!("AddBuffer failed: {}", e))?;
            sample
                .SetSampleTime(frame.timestamp)
                .map_err(|e| format!("SetSampleTime failed: {}", e))?;
            sample
                .SetSampleDuration(333333)
                .map_err(|e| format!("SetSampleDuration failed: {}", e))?; // ~33.33ms @ 30fps

            self.writer
                .WriteSample(self.video_stream_index, &sample)
                .map_err(|e| format!("WriteSample failed: {}", e))?;

            Ok(())
        }
    }

    /// Get current audio duration in seconds
    pub fn get_audio_duration_seconds(&self) -> f64 {
        (self.audio_timestamp as f64) / 10_000_000.0
    }

    /// Get current audio timestamp in 100-nanosecond units
    pub fn get_audio_timestamp(&self) -> i64 {
        self.audio_timestamp
    }

    pub fn write_audio_samples(&mut self, samples: &[i16]) -> Result<(), String> {
        if samples.is_empty() {
            return Ok(());
        }

        unsafe {
            let byte_len = (samples.len() * 2) as u32;
            let buffer = MFCreateMemoryBuffer(byte_len)
                .map_err(|e| format!("MFCreateMemoryBuffer failed: {}", e))?;

            let mut data_ptr = std::ptr::null_mut();
            buffer
                .Lock(&mut data_ptr, None, None)
                .map_err(|e| format!("Buffer Lock failed: {}", e))?;

            if !data_ptr.is_null() {
                std::ptr::copy_nonoverlapping(
                    samples.as_ptr() as *const u8,
                    data_ptr,
                    byte_len as usize,
                );
            }

            buffer
                .Unlock()
                .map_err(|e| format!("Buffer Unlock failed: {}", e))?;
            buffer
                .SetCurrentLength(byte_len)
                .map_err(|e| format!("SetCurrentLength failed: {}", e))?;

            let sample = MFCreateSample().map_err(|e| format!("MFCreateSample failed: {}", e))?;
            sample
                .AddBuffer(&buffer)
                .map_err(|e| format!("AddBuffer failed: {}", e))?;

            // Use incrementing timestamp based on samples written
            sample
                .SetSampleTime(self.audio_timestamp)
                .map_err(|e| format!("SetSampleTime failed: {}", e))?;

            // Calculate duration based on actual sample rate
            let frames = (samples.len() / 2) as i64; // Stereo: 2 channels
            let duration = (frames * 10_000_000) / self.audio_sample_rate as i64;
            sample
                .SetSampleDuration(duration)
                .map_err(|e| format!("SetSampleDuration failed: {}", e))?;

            self.writer
                .WriteSample(self.audio_stream_index, &sample)
                .map_err(|e| format!("WriteSample failed: {}", e))?;

            // Increment timestamp for next sample
            self.audio_timestamp += duration;

            Ok(())
        }
    }

    pub fn finalize(self) -> Result<(), String> {
        unsafe {
            crate::log_debug("Finalizing MP4 writer...");
            self.writer
                .Finalize()
                .map_err(|e| format!("Finalize failed: {}", e))?;
            crate::log_debug("MP4 writer finalized successfully");
            Ok(())
        }
    }
}
