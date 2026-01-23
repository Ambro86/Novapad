use crate::accessibility::to_wide;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Com::StructuredStorage::PropVariantClear;
use windows::Win32::System::Com::{COINIT_MULTITHREADED, CoInitializeEx, CoUninitialize};
use windows::Win32::UI::Controls::RichEdit::{EDITSTREAM, EM_STREAMOUT, SF_RTF};
use windows::Win32::UI::Shell::PropertiesSystem::{
    GPS_READWRITE, IPropertyStore, PROPERTYKEY, SHGetPropertyStoreFromParsingName,
};
use windows::Win32::UI::Shell::{SHStrDupW, ShellExecuteW};
use windows::Win32::UI::WindowsAndMessaging::SendMessageW;
use windows::core::{GUID, PCWSTR, PROPVARIANT, PWSTR, w};

const VT_LPWSTR: u16 = 31;

// System.Title: {F29F85E0-4FF9-1068-AB91-08002B27B3D9} 2
const PKEY_TITLE: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xF29F85E0_4FF9_1068_AB91_08002B27B3D9),
    pid: 2,
};

// System.Author: {F29F85E0-4FF9-1068-AB91-08002B27B3D9} 4
const PKEY_AUTHOR: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xF29F85E0_4FF9_1068_AB91_08002B27B3D9),
    pid: 4,
};

// System.Comment: {F29F85E0-4FF9-1068-AB91-08002B27B3D9} 6
const PKEY_COMMENT: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xF29F85E0_4FF9_1068_AB91_08002B27B3D9),
    pid: 6,
};

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

/// Joins multiple WAV files into a single one
pub fn join_wav_files(inputs: &[std::path::PathBuf], output: &Path) -> Result<(), AudioError> {
    if inputs.is_empty() {
        return Ok(());
    }

    // 1. Read the first file to get audio format details (sample rate, channels, etc.)
    let mut fmt_chunk = [0u8; 16];
    let mut sample_rate = 0u32;
    let mut channels = 0u16;
    let mut bits_per_sample = 0u16;

    {
        let mut first_file = File::open(&inputs[0])?;
        let mut header = [0u8; 12];
        first_file.read_exact(&mut header)?;
        if &header[0..4] != b"RIFF" || &header[8..12] != b"WAVE" {
            return Err(AudioError::InvalidFormat(
                "Not a valid RIFF/WAVE file".to_string(),
            ));
        }

        let mut found_fmt = false;
        let mut buffer = [0u8; 8];
        while first_file.read_exact(&mut buffer).is_ok() {
            let chunk_id = &buffer[0..4];
            let chunk_size = u32::from_le_bytes([buffer[4], buffer[5], buffer[6], buffer[7]]);
            if chunk_id == b"fmt " {
                let to_read = chunk_size.min(16) as usize;
                first_file.read_exact(&mut fmt_chunk[..to_read])?;
                if chunk_size > 16 {
                    first_file.seek(SeekFrom::Current((chunk_size - 16) as i64))?;
                }
                channels = u16::from_le_bytes([fmt_chunk[2], fmt_chunk[3]]);
                sample_rate =
                    u32::from_le_bytes([fmt_chunk[4], fmt_chunk[5], fmt_chunk[6], fmt_chunk[7]]);
                bits_per_sample = u16::from_le_bytes([fmt_chunk[14], fmt_chunk[15]]);
                found_fmt = true;
                break;
            } else {
                let skip = if chunk_size % 2 == 1 {
                    chunk_size + 1
                } else {
                    chunk_size
                };
                first_file.seek(SeekFrom::Current(skip as i64))?;
            }
        }
        if !found_fmt {
            return Err(AudioError::InvalidFormat("fmt chunk not found".to_string()));
        }
    }

    // 2. Create the output file with a standard 44-byte header
    let mut writer = WavWriter::create(output, sample_rate, channels, bits_per_sample)?;
    let mut total_data_size = 0u32;

    // 3. Append data from each file
    for input_path in inputs {
        let mut file = File::open(input_path)?;
        let mut header = [0u8; 12];
        file.read_exact(&mut header)?;

        let mut buffer = [0u8; 8];
        while file.read_exact(&mut buffer).is_ok() {
            let chunk_id = &buffer[0..4];
            let chunk_size = u32::from_le_bytes([buffer[4], buffer[5], buffer[6], buffer[7]]);

            if chunk_id == b"data" {
                let mut data_buf = vec![0u8; 1024 * 64];
                let mut remaining = chunk_size as usize;
                while remaining > 0 {
                    let to_read = remaining.min(data_buf.len());
                    file.read_exact(&mut data_buf[..to_read])?;
                    writer.file.write_all(&data_buf[..to_read])?;
                    remaining -= to_read;
                }
                total_data_size = total_data_size.saturating_add(chunk_size);
                break;
            } else {
                let skip = if chunk_size % 2 == 1 {
                    chunk_size + 1
                } else {
                    chunk_size
                };
                file.seek(SeekFrom::Current(skip as i64))?;
            }
        }
    }

    // 4. Finalize the output header with the correct total size
    writer.data_size = total_data_size;
    writer.finalize()?;

    Ok(())
}

pub fn open_url_in_browser(url: &str) -> Result<(), String> {
    let wide = to_wide(url);
    unsafe {
        let result = ShellExecuteW(
            HWND(0),
            w!("open"),
            PCWSTR(wide.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            windows::Win32::UI::WindowsAndMessaging::SW_SHOW,
        );
        if result.0 as isize <= 32 {
            return Err(format!("ShellExecute failed: {}", result.0));
        }
    }
    Ok(())
}

pub fn set_file_metadata(
    path: &Path,
    title: Option<&str>,
    author: Option<&str>,
    comment: Option<&str>,
) -> Result<(), String> {
    unsafe {
        let _hr = CoInitializeEx(None, COINIT_MULTITHREADED);

        let path_wide = to_wide(path.to_str().ok_or("Invalid path")?);

        let store: IPropertyStore =
            SHGetPropertyStoreFromParsingName(PCWSTR(path_wide.as_ptr()), None, GPS_READWRITE)
                .map_err(|e| format!("SHGetPropertyStoreFromParsingName failed: {}", e))?;

        if let Some(t) = title {
            set_prop(&store, &PKEY_TITLE, t)?;
        }
        if let Some(a) = author {
            set_prop(&store, &PKEY_AUTHOR, a)?;
        }
        if let Some(c) = comment {
            set_prop(&store, &PKEY_COMMENT, c)?;
        }

        store
            .Commit()
            .map_err(|e| format!("IPropertyStore::Commit failed: {}", e))?;
        CoUninitialize();
    }
    Ok(())
}

#[repr(C)]
#[allow(non_snake_case)]
struct MyPropVariant {
    vt: u16,
    w_reserved1: u16,
    w_reserved2: u16,
    w_reserved3: u16,
    pwsz_val: PWSTR,
    _padding: [u8; 8],
}

unsafe fn set_prop(store: &IPropertyStore, key: &PROPERTYKEY, value: &str) -> Result<(), String> {
    let wide = to_wide(value);
    let psz = SHStrDupW(PCWSTR(wide.as_ptr())).map_err(|e: windows::core::Error| e.to_string())?;

    let mut pv = MyPropVariant {
        vt: VT_LPWSTR,
        w_reserved1: 0,
        w_reserved2: 0,
        w_reserved3: 0,
        pwsz_val: psz,
        _padding: [0; 8],
    };

    let res = store.SetValue(key, &pv as *const MyPropVariant as *const PROPVARIANT);
    unsafe {
        PropVariantClear(&mut pv as *mut MyPropVariant as *mut PROPVARIANT).ok();
    }

    res.map_err(|e| format!("IPropertyStore::SetValue failed: {}", e))
}

pub unsafe fn write_rtf_text(path: &Path, hwnd_edit: HWND) -> Result<(), String> {
    let mut file = File::create(path).map_err(|e| e.to_string())?;

    unsafe extern "system" fn stream_out_callback(
        dw_cookie: usize,
        pb_buff: *mut u8,
        cb: i32,
        pcb: *mut i32,
    ) -> u32 {
        let file = &mut *(dw_cookie as *mut File);
        let data = std::slice::from_raw_parts(pb_buff, cb as usize);
        match file.write_all(data) {
            Ok(_) => {
                *pcb = cb;
                0
            }
            Err(_) => 1,
        }
    }

    let mut es = EDITSTREAM {
        dwCookie: &mut file as *mut _ as usize,
        dwError: 0,
        pfnCallback: Some(stream_out_callback),
    };

    SendMessageW(
        hwnd_edit,
        EM_STREAMOUT,
        WPARAM(SF_RTF as usize),
        LPARAM(&mut es as *mut _ as isize),
    );

    if es.dwError != 0 {
        return Err(format!("RTF stream out failed: {}", es.dwError));
    }

    crate::log_debug(&format!("RTF: Successfully saved to {:?}", path));
    Ok(())
}
