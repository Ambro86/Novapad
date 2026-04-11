use crate::accessibility::to_wide;
use std::path::Path;
use std::sync::Once;
use windows::Win32::Media::MediaFoundation::{
    IMFCollection, IMFMediaBuffer, IMFMediaType, IMFSample, IMFSinkWriter, IMFSourceReader,
    MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT,
    MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_FIXED_SIZE_SAMPLES,
    MF_MT_MAJOR_TYPE, MF_MT_SAMPLE_SIZE, MF_MT_SUBTYPE, MF_PD_DURATION, MF_VERSION,
    MFAudioFormat_MP3, MFAudioFormat_PCM, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample,
    MFCreateSinkWriterFromURL, MFCreateSourceReaderFromURL, MFMediaType_Audio, MFStartup,
    MFT_ENUM_FLAG_ALL, MFTranscodeGetAudioOutputAvailableTypes,
};
use windows::core::{Interface, PCWSTR};

static MF_STARTUP: Once = Once::new();

pub struct MfGuard;

impl MfGuard {
    pub fn start() -> Result<Self, String> {
        let mut res = Ok(());
        MF_STARTUP.call_once(|| {
            crate::log_debug("MF: startup begin");
            unsafe {
                if let Err(e) = MFStartup(MF_VERSION, 0) {
                    res = Err(format!(
                        "Media Foundation not available. Install Media Feature Pack on Windows N/KN. ({})",
                        e
                    ));
                } else {
                    crate::log_debug("MF: startup ok");
                }
            }
        });
        res.map(|_| MfGuard)
    }
}

impl Drop for MfGuard {
    fn drop(&mut self) {
        // We keep MF running once started to avoid overhead of continuous startup/shutdown
    }
}

pub fn get_audio_duration_mf(path: &std::path::Path) -> Result<u64, String> {
    unsafe {
        let _guard = MfGuard::start()?;

        let path_wide = crate::accessibility::to_wide(path.to_str().ok_or("Invalid path")?);

        let reader: IMFSourceReader = MFCreateSourceReaderFromURL(PCWSTR(path_wide.as_ptr()), None)
            .map_err(|e| format!("MFCreateSourceReaderFromURL failed: {}", e))?;

        // Use MF_SOURCE_READER_MEDIASOURCE (0xFFFFFFFF) for presentation-wide attributes like duration

        let var = reader
            .GetPresentationAttribute(0xffffffff, &MF_PD_DURATION)
            .map_err(|e| format!("GetPresentationAttribute failed: {}", e))?;

        // MF_PD_DURATION is a VT_I8 (long long) representing duration in 100-nanosecond units.

        let hns = windows::Win32::System::Com::StructuredStorage::PropVariantToInt64(&var)
            .map_err(|e| format!("PropVariantToInt64 failed: {}", e))?;
        let secs = (hns / 10_000_000) as u64;
        Ok(secs)
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
                64 => 64,
                80 => 80,
                96 => 96,
                112 => 112,
                128 => 128,
                160 => 160,
                192 => 192,
                224 => 224,
                256 => 256,
                320 => 320,
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
            if let Err(e) = pcm_type.SetUINT32(&MF_MT_SAMPLE_SIZE, block_align) {
                crate::log_debug(&format!("MF: set sample size failed: {}", e));
            }

            let out_type = select_mp3_output_type(sample_rate, requested_channels, bitrate_kbps)?;

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

fn select_mp3_output_type(
    sample_rate: u32,
    channels: u32,
    bitrate_kbps: u32,
) -> Result<IMFMediaType, String> {
    unsafe {
        let requested_avg_bytes = (bitrate_kbps * 1000) / 8;
        let available_types: IMFCollection = MFTranscodeGetAudioOutputAvailableTypes(
            &MFAudioFormat_MP3,
            MFT_ENUM_FLAG_ALL.0 as u32,
            None,
        )
        .map_err(|e| format!("MFTranscodeGetAudioOutputAvailableTypes failed: {}", e))?;
        let count = available_types
            .GetElementCount()
            .map_err(|e| format!("IMFCollection::GetElementCount failed: {}", e))?;

        let mut best_candidate: Option<(u32, IMFMediaType)> = None;
        for index in 0..count {
            let candidate: IMFMediaType = available_types
                .GetElement(index)
                .map_err(|e| format!("IMFCollection::GetElement failed: {}", e))?
                .cast()
                .map_err(|e| format!("IMFCollection element cast failed: {}", e))?;

            let candidate_rate = candidate
                .GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND)
                .map_err(|e| format!("MF_MT_AUDIO_SAMPLES_PER_SECOND missing: {}", e))?;
            let candidate_channels = candidate
                .GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS)
                .map_err(|e| format!("MF_MT_AUDIO_NUM_CHANNELS missing: {}", e))?;
            let candidate_avg_bytes = candidate
                .GetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND)
                .map_err(|e| format!("MF_MT_AUDIO_AVG_BYTES_PER_SECOND missing: {}", e))?;

            let score = mp3_output_type_score(
                sample_rate,
                channels,
                requested_avg_bytes,
                candidate_rate,
                candidate_channels,
                candidate_avg_bytes,
            );

            if best_candidate
                .as_ref()
                .map(|(best_score, _)| score < *best_score)
                .unwrap_or(true)
            {
                best_candidate = Some((score, candidate));
            }
        }

        let Some((_, selected)) = best_candidate else {
            return Err("No Media Foundation MP3 output types available.".to_string());
        };

        let selected_rate = selected
            .GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND)
            .map_err(|e| format!("Selected MP3 rate unavailable: {}", e))?;
        let selected_channels = selected
            .GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS)
            .map_err(|e| format!("Selected MP3 channels unavailable: {}", e))?;
        let selected_avg_bytes = selected
            .GetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND)
            .map_err(|e| format!("Selected MP3 bitrate unavailable: {}", e))?;
        crate::log_debug(&format!(
            "MF: selected mp3 output type request={}kbps/{}Hz/{}ch actual={}kbps/{}Hz/{}ch",
            bitrate_kbps,
            sample_rate,
            channels,
            (selected_avg_bytes * 8) / 1000,
            selected_rate,
            selected_channels
        ));

        Ok(selected)
    }
}

fn mp3_output_type_score(
    requested_rate: u32,
    requested_channels: u32,
    requested_avg_bytes: u32,
    candidate_rate: u32,
    candidate_channels: u32,
    candidate_avg_bytes: u32,
) -> u32 {
    let bitrate_delta = candidate_avg_bytes.abs_diff(requested_avg_bytes);
    let rate_delta = candidate_rate.abs_diff(requested_rate);
    let channel_delta = candidate_channels.abs_diff(requested_channels);

    channel_delta
        .saturating_mul(1_000_000)
        .saturating_add(bitrate_delta.saturating_mul(100))
        .saturating_add(rate_delta)
}
