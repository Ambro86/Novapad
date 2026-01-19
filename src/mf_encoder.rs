use crate::accessibility::to_wide;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;
use windows::Win32::Media::MediaFoundation::{
    IMFMediaBuffer, IMFMediaType, IMFSample, IMFSinkWriter, IMFSourceReader,
    MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT,
    MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_FIXED_SIZE_SAMPLES,
    MF_MT_MAJOR_TYPE, MF_MT_SAMPLE_SIZE, MF_MT_SUBTYPE, MF_SOURCE_READER_FIRST_AUDIO_STREAM,
    MF_SOURCE_READERF_ENDOFSTREAM, MF_VERSION, MFAudioFormat_MP3, MFAudioFormat_PCM,
    MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFCreateSinkWriterFromURL,
    MFCreateSourceReaderFromURL, MFMediaType_Audio, MFShutdown, MFStartup,
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
            crate::log_if_err!(MFShutdown());
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
            if let Err(e) = pcm_type.SetUINT32(&MF_MT_FIXED_SIZE_SAMPLES, 1) {
                crate::log_debug(&format!("Failed to set fixed size samples: {}", e));
            }
            if let Err(_e) = pcm_type.SetUINT32(&MF_MT_SAMPLE_SIZE, block_align) {
                crate::log_debug(&format!("Error: {:?}", _e));
            }

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
        let chunk_size = u32::from_le_bytes(chunk_header[4..8].try_into().unwrap_or([0, 0, 0, 0]));
        if chunk_id == b"data" {
            let data_offset = file.stream_position().map_err(|e| e.to_string())?;
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
        if let Err(e) = pcm_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, requested_rate) {
            crate::log_debug(&format!("Failed to set audio samples per second: {}", e));
        }
        if let Err(_e) = pcm_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, requested_channels) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = pcm_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, requested_bits) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = pcm_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, requested_block_align) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if let Err(_e) = pcm_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, requested_avg_bytes)
        {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
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
        if let Err(_e) = input_type.SetUINT32(&MF_MT_FIXED_SIZE_SAMPLES, 1) {
            crate::log_debug(&format!("Error: {:?}", _e));
        }
        if block_align != 0
            && let Err(_e) = input_type.SetUINT32(&MF_MT_SAMPLE_SIZE, block_align)
        {
            crate::log_debug(&format!("Error: {:?}", _e));
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
            if let Some(cancel) = cancel
                && cancel.load(std::sync::atomic::Ordering::Relaxed)
            {
                return Err("Saving canceled.".to_string());
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

        if let Some(cancel) = cancel
            && cancel.load(std::sync::atomic::Ordering::Relaxed)
        {
            return Err("Saving canceled.".to_string());
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
