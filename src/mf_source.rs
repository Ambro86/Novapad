use crate::accessibility::to_wide;
use crate::mf_encoder::MfGuard;
use rodio::Source;
use std::time::Duration;
use windows::Win32::Media::MediaFoundation::{
    IMFAttributes, IMFSourceReader, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_NUM_CHANNELS,
    MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE, MF_PD_DURATION,
    MF_SOURCE_READER_ALL_STREAMS, MF_SOURCE_READER_FIRST_AUDIO_STREAM,
    MF_SOURCE_READER_MEDIASOURCE, MF_SOURCE_READERF_ENDOFSTREAM, MFAudioFormat_PCM,
    MFCreateAttributes, MFCreateMediaType, MFCreateSourceReaderFromURL, MFMediaType_Audio,
};
use windows::Win32::System::Com::StructuredStorage::PropVariantToInt64;
use windows::core::{GUID, PCWSTR, PROPVARIANT};

// GUID for MF_READWRITE_ENABLE_FORMAT_CONVERSION: 17445a34-22f1-4fd7-a07d-87660c21be13
const MF_READWRITE_ENABLE_FORMAT_CONVERSION: GUID =
    GUID::from_u128(0x17445a34_22f1_4fd7_a07d_87660c21be13);

pub struct MfSource {
    _guard: MfGuard,
    reader: IMFSourceReader,
    channels: u16,
    sample_rate: u32,
    duration: Option<Duration>,
    buffer: Vec<f32>,
    index: usize,
    eof: bool,
}

unsafe impl Send for MfSource {}
unsafe impl Sync for MfSource {}

impl MfSource {
    pub fn try_new(path: &std::path::Path) -> Result<Self, String> {
        unsafe {
            let guard = MfGuard::start()?;
            let path_wide = to_wide(path.to_str().ok_or("Invalid path")?);

            // Create attributes to enable format conversion and hardware/system decoders
            let mut attributes: Option<IMFAttributes> = None;
            MFCreateAttributes(&mut attributes, 2)
                .map_err(|e| format!("MFCreateAttributes failed: {}", e))?;
            let attributes = attributes.ok_or("Failed to create attributes")?;

            attributes
                .SetUINT32(&MF_READWRITE_ENABLE_FORMAT_CONVERSION, 1)
                .map_err(|e| format!("Failed to enable format conversion: {}", e))?;

            // Allow all decoders (hardware and system)
            attributes
                .SetUINT32(&windows::Win32::Media::MediaFoundation::MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, 1)
                .map_err(|e| e.to_string())?;

            let reader: IMFSourceReader =
                MFCreateSourceReaderFromURL(PCWSTR(path_wide.as_ptr()), Some(&attributes))
                    .map_err(|e| format!("MFCreateSourceReaderFromURL failed: {}", e))?;

            // Disable all streams and try to find the first audio stream
            reader
                .SetStreamSelection(MF_SOURCE_READER_ALL_STREAMS.0 as u32, false)
                .ok();
            reader
                .SetStreamSelection(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32, true)
                .map_err(|e| format!("Enable audio stream failed: {}", e))?;

            // Get native format to match sample rate and channels
            let native_type = reader
                .GetNativeMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32, 0)
                .map_err(|e| format!("GetNativeMediaType failed: {}", e))?;

            let native_channels = native_type
                .GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS)
                .unwrap_or(2);
            let native_rate = native_type
                .GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND)
                .unwrap_or(44100);

            // Configure output to PCM matching native parameters
            let pcm_type =
                MFCreateMediaType().map_err(|e| format!("MFCreateMediaType failed: {}", e))?;
            pcm_type
                .SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)
                .map_err(|e| format!("SetGUID major failed: {}", e))?;
            pcm_type
                .SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)
                .map_err(|e| format!("SetGUID subtype failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)
                .map_err(|e| format!("Set bits failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, native_channels)
                .map_err(|e| format!("Set channels failed: {}", e))?;
            pcm_type
                .SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, native_rate)
                .map_err(|e| format!("Set rate failed: {}", e))?;

            reader
                .SetCurrentMediaType(
                    MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32,
                    None,
                    &pcm_type,
                )
                .map_err(|e| format!("SetCurrentMediaType failed: {}", e))?;

            let current_type = reader
                .GetCurrentMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32)
                .map_err(|e| format!("GetCurrentMediaType failed: {}", e))?;

            let channels = current_type
                .GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS)
                .map_err(|e| format!("Get channels (final) failed: {}", e))?
                as u16;
            let sample_rate = current_type
                .GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND)
                .map_err(|e| format!("Get sample rate (final) failed: {}", e))?;

            let duration = get_reader_duration(&reader).map(Duration::from_secs);

            Ok(Self {
                _guard: guard,
                reader,
                channels,
                sample_rate,
                duration,
                buffer: Vec::new(),
                index: 0,
                eof: false,
            })
        }
    }

    pub fn seek(&mut self, pos: Duration) -> Result<(), String> {
        unsafe {
            let hns = (pos.as_nanos() / 100) as i64;
            let var = PROPVARIANT::from(hns);

            self.reader
                .SetCurrentPosition(&GUID::default(), &var)
                .map_err(|e| format!("SetCurrentPosition failed: {}", e))?;
            self.buffer.clear();
            self.index = 0;
            self.eof = false;
            Ok(())
        }
    }

    fn refill(&mut self) -> bool {
        if self.eof {
            return false;
        }
        unsafe {
            let mut flags = 0u32;
            let mut sample = None;
            let res = self.reader.ReadSample(
                MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32,
                0,
                None,
                Some(&mut flags),
                None,
                Some(&mut sample),
            );

            if res.is_err() || (flags & MF_SOURCE_READERF_ENDOFSTREAM.0 as u32) != 0 {
                self.eof = true;
                return false;
            }

            if let Some(sample) = sample {
                let buffer = sample.ConvertToContiguousBuffer().ok();
                if let Some(buffer) = buffer {
                    let mut data_ptr = std::ptr::null_mut();
                    let mut current_len = 0u32;
                    if buffer
                        .Lock(&mut data_ptr, None, Some(&mut current_len))
                        .is_ok()
                    {
                        let samples_count = current_len as usize / 2; // Assuming 16-bit PCM
                        let raw_samples =
                            std::slice::from_raw_parts(data_ptr as *const i16, samples_count);
                        self.buffer = raw_samples.iter().map(|&s| s as f32 / 32768.0).collect();
                        self.index = 0;
                        buffer.Unlock().ok();
                        return !self.buffer.is_empty();
                    }
                }
            }
        }
        false
    }
}

impl Iterator for MfSource {
    type Item = f32;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.buffer.len() && !self.refill() {
            return None;
        }
        let sample = self.buffer[self.index];
        self.index += 1;
        Some(sample)
    }
}

impl Source for MfSource {
    fn current_span_len(&self) -> Option<usize> {
        None
    }

    fn channels(&self) -> u16 {
        self.channels
    }

    fn sample_rate(&self) -> u32 {
        self.sample_rate
    }

    fn total_duration(&self) -> Option<Duration> {
        self.duration
    }
}

fn get_reader_duration(reader: &IMFSourceReader) -> Option<u64> {
    unsafe {
        // Use MF_SOURCE_READER_MEDIASOURCE for presentation-wide attributes
        let var = reader
            .GetPresentationAttribute(MF_SOURCE_READER_MEDIASOURCE.0 as u32, &MF_PD_DURATION)
            .ok()?;

        let hns = PropVariantToInt64(&var).ok()?;

        Some((hns / 10_000_000) as u64)
    }
}
