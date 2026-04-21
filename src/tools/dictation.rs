use crate::podcast_recorder::{self, RecorderHandle};
use crate::settings::{PODCAST_DEVICE_DEFAULT, PodcastFormat};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};

pub struct DictationRecordingConfig {
    pub mic_device_id: String,
    pub mic_gain: f32,
}

fn dictation_temp_dir() -> PathBuf {
    std::env::temp_dir().join("SonarpadDictation")
}

pub fn wav_duration_seconds(path: &Path) -> Result<f64, String> {
    let mut file = File::open(path).map_err(|err| format!("open wav failed: {err}"))?;
    let mut riff = [0_u8; 12];
    file.read_exact(&mut riff)
        .map_err(|err| format!("read wav header failed: {err}"))?;
    if &riff[0..4] != b"RIFF" || &riff[8..12] != b"WAVE" {
        return Err("not a RIFF/WAVE file".to_string());
    }

    let mut sample_rate = None;
    let mut block_align = None;
    let mut data_size = None;

    loop {
        let mut chunk_header = [0_u8; 8];
        match file.read_exact(&mut chunk_header) {
            Ok(()) => {}
            Err(err) if err.kind() == std::io::ErrorKind::UnexpectedEof => break,
            Err(err) => return Err(format!("read wav chunk header failed: {err}")),
        }

        let chunk_size = u32::from_le_bytes([
            chunk_header[4],
            chunk_header[5],
            chunk_header[6],
            chunk_header[7],
        ]);
        let chunk_id = &chunk_header[0..4];

        if chunk_id == b"fmt " {
            let mut fmt = vec![0_u8; chunk_size as usize];
            file.read_exact(&mut fmt)
                .map_err(|err| format!("read fmt chunk failed: {err}"))?;
            if fmt.len() >= 16 {
                sample_rate = Some(u32::from_le_bytes([fmt[4], fmt[5], fmt[6], fmt[7]]));
                block_align = Some(u16::from_le_bytes([fmt[12], fmt[13]]));
            }
        } else if chunk_id == b"data" {
            data_size = Some(chunk_size);
            file.seek(SeekFrom::Current(i64::from(chunk_size)))
                .map_err(|err| format!("skip data chunk failed: {err}"))?;
        } else {
            file.seek(SeekFrom::Current(i64::from(chunk_size)))
                .map_err(|err| format!("skip wav chunk failed: {err}"))?;
        }

        if chunk_size % 2 != 0 {
            file.seek(SeekFrom::Current(1))
                .map_err(|err| format!("skip wav padding failed: {err}"))?;
        }
    }

    let sample_rate = sample_rate.ok_or_else(|| "missing wav sample rate".to_string())?;
    let block_align = block_align.ok_or_else(|| "missing wav block align".to_string())?;
    let data_size = data_size.ok_or_else(|| "missing wav data chunk".to_string())?;
    if sample_rate == 0 || block_align == 0 {
        return Err("invalid wav fmt values".to_string());
    }

    Ok(data_size as f64 / f64::from(sample_rate) / f64::from(block_align))
}

pub fn start_recording(config: &DictationRecordingConfig) -> Result<RecorderHandle, String> {
    let device_name =
        if config.mic_device_id.is_empty() || config.mic_device_id == PODCAST_DEVICE_DEFAULT {
            String::new()
        } else {
            podcast_recorder::list_input_devices()?
                .into_iter()
                .find(|device| device.id == config.mic_device_id)
                .map(|device| device.name)
                .unwrap_or_default()
        };

    crate::log_debug(&format!(
        "Dictation: start recording mic_device_id='{}' gain={:.2} temp_dir={}",
        config.mic_device_id,
        config.mic_gain,
        dictation_temp_dir().display()
    ));

    podcast_recorder::start_recording(podcast_recorder::RecorderConfig {
        include_mic: true,
        mic_device_id: config.mic_device_id.clone(),
        mic_device_name: device_name,
        mic_gain: config.mic_gain,
        include_system: false,
        system_device_id: String::new(),
        system_device_name: String::new(),
        system_gain: 0.0,
        single_app_process_id: None,
        selected_app_process_ids: Vec::new(),
        output_format: PodcastFormat::Wav,
        mp3_bitrate: 64,
        save_folder: dictation_temp_dir(),
    })
}
