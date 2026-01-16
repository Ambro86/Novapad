use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

/// Errors that can occur during audio operations
#[derive(Debug)]
pub enum AudioError {
    Io(std::io::Error),
    InvalidFormat(String),
}

impl std::fmt::Display for AudioError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            AudioError::Io(err) => write!(f, "IO error: {}", err),
            AudioError::InvalidFormat(msg) => write!(f, "Invalid format: {}", msg),
        }
    }
}

impl From<std::io::Error> for AudioError {
    fn from(err: std::io::Error) -> Self {
        AudioError::Io(err)
    }
}

/// Helper to write WAV files safely
pub struct WavWriter {
    file: File,
    data_size: u32,
    sample_rate: u32,
    channels: u16,
    bits_per_sample: u16,
}

impl WavWriter {
    pub fn create(
        path: &Path,
        sample_rate: u32,
        channels: u16,
        bits_per_sample: u16,
    ) -> Result<Self, AudioError> {
        let file = std::fs::OpenOptions::new()
            .create(true)
            .truncate(true)
            .write(true)
            .open(path)?;

        let mut writer = WavWriter {
            file,
            data_size: 0,
            sample_rate,
            channels,
            bits_per_sample,
        };
        writer.write_header_placeholder()?;
        Ok(writer)
    }

    fn write_header_placeholder(&mut self) -> Result<(), AudioError> {
        // RIFF header
        self.file.write_all(b"RIFF")?;
        self.file.write_all(&0u32.to_le_bytes())?; // Placeholder for file size
        self.file.write_all(b"WAVE")?;

        // fmt chunk
        self.file.write_all(b"fmt ")?;
        self.file.write_all(&16u32.to_le_bytes())?; // Chunk size
        self.file.write_all(&1u16.to_le_bytes())?; // PCM format
        self.file.write_all(&self.channels.to_le_bytes())?;
        self.file.write_all(&self.sample_rate.to_le_bytes())?;

        let byte_rate = self.sample_rate * self.channels as u32 * (self.bits_per_sample as u32 / 8);
        let block_align = self.channels * (self.bits_per_sample / 8);

        self.file.write_all(&byte_rate.to_le_bytes())?;
        self.file.write_all(&block_align.to_le_bytes())?;
        self.file.write_all(&self.bits_per_sample.to_le_bytes())?;

        // data chunk
        self.file.write_all(b"data")?;
        self.file.write_all(&0u32.to_le_bytes())?; // Placeholder for data size

        Ok(())
    }

    #[allow(dead_code)]
    pub fn write_samples_i16(&mut self, samples: &[i16]) -> Result<(), AudioError> {
        let mut buf = Vec::with_capacity(samples.len() * 2);
        for sample in samples {
            buf.extend_from_slice(&sample.to_le_bytes());
        }
        self.file.write_all(&buf)?;
        self.data_size = self.data_size.saturating_add(buf.len() as u32);
        Ok(())
    }

    pub fn write_samples_f32(&mut self, samples: &[f32]) -> Result<(), AudioError> {
        // Convert f32 samples (-1.0 to 1.0) to i16
        let mut buf = Vec::with_capacity(samples.len() * 2);
        for sample in samples {
            let clamped = sample.clamp(-1.0, 1.0);
            let v = (clamped * i16::MAX as f32) as i16;
            buf.extend_from_slice(&v.to_le_bytes());
        }
        self.file.write_all(&buf)?;
        self.data_size = self.data_size.saturating_add(buf.len() as u32);
        Ok(())
    }

    pub fn write_silence_ms(&mut self, duration_ms: u32) -> Result<(), AudioError> {
        let bytes_per_sample = (self.bits_per_sample / 8) as u32;
        let samples = self.sample_rate.saturating_mul(duration_ms) / 1000;
        let total_samples = samples.saturating_mul(self.channels as u32);
        let byte_count = total_samples.saturating_mul(bytes_per_sample);

        let zeros = vec![0u8; 4096];
        let mut remaining = byte_count as usize;
        while remaining > 0 {
            let chunk = remaining.min(zeros.len());
            self.file.write_all(&zeros[..chunk])?;
            remaining -= chunk;
        }
        self.data_size = self.data_size.saturating_add(byte_count);
        Ok(())
    }

    pub fn finalize(&mut self) -> Result<(), AudioError> {
        let riff_size = 36u32.saturating_add(self.data_size);

        // Update RIFF size
        self.file.seek(SeekFrom::Start(4))?;
        self.file.write_all(&riff_size.to_le_bytes())?;

        // Update data chunk size
        self.file.seek(SeekFrom::Start(40))?;
        self.file.write_all(&self.data_size.to_le_bytes())?;

        self.file.flush()?;
        Ok(())
    }
}

/// Get the size of the data chunk in a WAV file
pub fn get_wav_data_size(path: &Path) -> Result<u32, AudioError> {
    let mut file = File::open(path)?;
    let mut header = [0u8; 12];
    file.read_exact(&mut header)?;

    if &header[0..4] != b"RIFF" || &header[8..12] != b"WAVE" {
        return Err(AudioError::InvalidFormat("Invalid WAV header".to_string()));
    }

    let mut buffer = [0u8; 8];
    while file.read_exact(&mut buffer).is_ok() {
        let chunk_id = &buffer[0..4];
        let chunk_size = u32::from_le_bytes([buffer[4], buffer[5], buffer[6], buffer[7]]);

        if chunk_id == b"data" {
            return Ok(chunk_size);
        }

        // Skip chunk (must be even-aligned)
        let skip = if chunk_size % 2 == 1 {
            chunk_size + 1
        } else {
            chunk_size
        };
        file.seek(SeekFrom::Current(skip as i64))?;
    }
    Err(AudioError::InvalidFormat(
        "WAV data chunk not found".to_string(),
    ))
}

/// Write a simple silence WAV file (utility function)
pub fn write_silence_file(
    path: &Path,
    sample_rate: u32,
    channels: u16,
    bits_per_sample: u16,
    duration_ms: u32,
) -> Result<(), AudioError> {
    let mut writer = WavWriter::create(path, sample_rate, channels, bits_per_sample)?;
    writer.write_silence_ms(duration_ms)?;
    writer.finalize()?;
    Ok(())
}
