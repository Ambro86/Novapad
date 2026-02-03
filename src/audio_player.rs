use crate::accessibility::{nvda_speak, to_wide};
use crate::bass_output::BassOutput;
use crate::ffmpeg_export::{MixExportOptions, export_mixed_media, is_mixed_output};
use crate::ffmpeg_source::FfmpegSource;
use crate::i18n;
use crate::log_debug;
use crate::settings::{FileFormat, SubtitleReadMode, confirm_title, settings_dir};
use crate::subtitles::{
    SUBTITLE_EXTENSIONS, SubtitleCue, clear_subtitle_override, find_subtitle_for_media,
    load_subtitles, set_subtitle_override,
};
use crate::tts_engine;
use crate::tts_engine::TtsCommand;
use crate::with_state;
use futures_util::StreamExt;
use rodio::{Decoder, Source};
use sha2::Digest;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::Mutex;
use std::sync::OnceLock;
use std::sync::atomic::AtomicUsize;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;
use std::time::UNIX_EPOCH;
use uuid::Uuid;
use windows::Win32::Foundation::{CloseHandle, HANDLE, HWND, WAIT_OBJECT_0};
use windows::Win32::System::Threading::{
    CreateWaitableTimerW, GetCurrentThread, SetThreadPriority, SetWaitableTimer,
    THREAD_PRIORITY_HIGHEST, WaitForSingleObject,
};
use windows::Win32::UI::WindowsAndMessaging::{IDYES, MB_ICONQUESTION, MB_YESNO, MessageBoxW};
use windows::core::PCWSTR;

type SubtitleSpeechCancel = Arc<Mutex<Option<Arc<AtomicBool>>>>;
type SubtitleSpeechCommand = Arc<Mutex<Option<tokio::sync::mpsc::UnboundedSender<TtsCommand>>>>;
type SubtitleSpeechHandles = (SubtitleSpeechCancel, SubtitleSpeechCommand);

pub struct AudiobookPlayer {
    pub path: PathBuf,
    output: Arc<BassOutput>,
    pub is_paused: bool,
    pub start_instant: std::time::Instant,
    pub accumulated_seconds: u64,
    pub volume: f32,
    pub muted: bool,
    pub prev_volume: f32,
    pub speed: f32,
    pub pitch: f32,
    pub subtitle_cancel: Arc<AtomicBool>,
    pub subtitle_hold: bool,
    pub subtitle_speech_cancel: SubtitleSpeechCancel,
    pub subtitle_speech_command: SubtitleSpeechCommand,
    pub session_id: u64,
}

impl AudiobookPlayer {
    fn play(&self) {
        self.output.play();
    }

    fn pause(&self) {
        self.output.pause();
    }

    fn stop(&self) {
        self.output.stop();
    }

    fn set_volume(&self, volume: f32) {
        self.output.set_volume(volume);
    }

    pub(crate) fn position_secs(&self) -> Option<f64> {
        self.output.position_secs()
    }
}

pub unsafe fn set_audiobook_subtitle_override(
    hwnd: HWND,
    subtitle_path: &Path,
) -> Result<(), String> {
    let ext_ok = subtitle_path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| {
            SUBTITLE_EXTENSIONS
                .iter()
                .any(|ext| s.eq_ignore_ascii_case(ext))
        })
        .unwrap_or(false);
    if !ext_ok {
        return Err("invalid_subtitle".to_string());
    }

    let mut active = None;
    let mut current_media = None;
    let mut new_cancel = None;
    let mut old_cancel = None;
    let mut speech_handles = None;
    let mut subtitle_mode = SubtitleReadMode::Off;
    let mut settings_snapshot = None;
    let mut resume_info: Option<(PathBuf, u64, AudiobookPlaybackOptions)> = None;
    let state_ok = with_state(hwnd, |state| {
        subtitle_mode = state.settings.subtitle_read_mode;
        settings_snapshot = Some(state.settings.clone());
        if let Some(player) = &mut state.active_audiobook {
            let seconds = audiobook_position_secs(player).max(0.0).floor() as u64;
            let options = AudiobookPlaybackOptions {
                speed: player.speed,
                pitch: player.pitch,
                paused: player.is_paused,
                volume: player.volume,
                muted: player.muted,
                prev_volume: player.prev_volume,
                mix_export: false,
                audio_track: None,
            };
            resume_info = Some((player.path.clone(), seconds, options));
            current_media = Some(player.path.clone());
            old_cancel = Some(player.subtitle_cancel.clone());
            let fresh_cancel = Arc::new(AtomicBool::new(false));
            player.subtitle_cancel = fresh_cancel.clone();
            new_cancel = Some(fresh_cancel);
            speech_handles = Some((
                player.subtitle_speech_cancel.clone(),
                player.subtitle_speech_command.clone(),
            ));
            active = Some(player.path.clone());
            return;
        }
        if let Some(doc) = state.docs.get(state.current)
            && matches!(doc.format, FileFormat::Audiobook)
        {
            current_media = doc.path.clone();
        }
    })
    .is_some();
    if !state_ok {
        return Err("state_unavailable".to_string());
    }

    let Some(media_path) = current_media else {
        return Err("no_media".to_string());
    };

    set_subtitle_override(&media_path, subtitle_path.to_path_buf());

    if let Some(cancel) = old_cancel {
        cancel.store(true, Ordering::Relaxed);
    }
    if let Some(handles) = speech_handles {
        stop_shared_subtitle_speech(&handles.0, &handles.1, "subtitle_override");
    }
    if let (Some(active_path), Some(cancel)) = (active, new_cancel) {
        if subtitle_mode == SubtitleReadMode::Record {
            if let Some(settings) = settings_snapshot {
                if resume_info.is_some() {
                    stop_audiobook_playback(hwnd);
                }
                let path_clone = active_path.clone();
                let resume_info = resume_info.clone();
                std::thread::spawn(move || {
                    if !confirm_edge_subtitle_download(hwnd, &path_clone, &settings) {
                        log_debug("Subtitle: Edge download not confirmed for record mode.");
                        if let Some((resume_path, seconds, options)) = resume_info {
                            start_audiobook_at_with_options(hwnd, resume_path, seconds, options);
                        }
                        return;
                    }
                    let mix_opts = MixExportOptions {
                        ducking: settings.subtitle_mix_ducking,
                    };
                    match export_mixed_media(&path_clone, &settings, &mix_opts) {
                        Ok(out_path) => {
                            log_debug(&format!(
                                "Subtitle: mixed file created at {}",
                                out_path.display()
                            ));
                            if let Some((_, seconds, mut options)) = resume_info {
                                options.mix_export = false;
                                start_audiobook_at_with_options(hwnd, out_path, seconds, options);
                            }
                        }
                        Err(err) => {
                            log_debug(&format!("Subtitle: mix export failed: {}", err));
                            if let Some((resume_path, seconds, options)) = resume_info {
                                start_audiobook_at_with_options(
                                    hwnd,
                                    resume_path,
                                    seconds,
                                    options,
                                );
                            }
                        }
                    }
                });
            } else {
                log_debug("Subtitle: settings unavailable for record mode.");
            }
        } else {
            start_subtitle_reader(hwnd, active_path, cancel, None);
        }
    }

    Ok(())
}

pub unsafe fn clear_audiobook_subtitle_override(hwnd: HWND) -> Result<(), String> {
    let mut active = None;
    let mut current_media = None;
    let mut new_cancel = None;
    let mut old_cancel = None;
    let mut speech_handles = None;
    let state_ok = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            current_media = Some(player.path.clone());
            old_cancel = Some(player.subtitle_cancel.clone());
            let fresh_cancel = Arc::new(AtomicBool::new(false));
            player.subtitle_cancel = fresh_cancel.clone();
            new_cancel = Some(fresh_cancel);
            speech_handles = Some((
                player.subtitle_speech_cancel.clone(),
                player.subtitle_speech_command.clone(),
            ));
            active = Some(player.path.clone());
            return;
        }
        if let Some(doc) = state.docs.get(state.current)
            && matches!(doc.format, FileFormat::Audiobook)
        {
            current_media = doc.path.clone();
        }
    })
    .is_some();
    if !state_ok {
        return Err("state_unavailable".to_string());
    }

    let Some(media_path) = current_media else {
        return Err("no_media".to_string());
    };

    clear_subtitle_override(&media_path);

    if let Some(cancel) = old_cancel {
        cancel.store(true, Ordering::Relaxed);
    }
    if let Some(handles) = speech_handles {
        stop_shared_subtitle_speech(&handles.0, &handles.1, "subtitle_override_clear");
    }
    if let (Some(active_path), Some(cancel)) = (active, new_cancel) {
        start_subtitle_reader(hwnd, active_path, cancel, None);
    }

    Ok(())
}

fn codec_name(codec: symphonia::core::codecs::CodecType) -> &'static str {
    use symphonia::core::codecs::{
        CODEC_TYPE_AAC, CODEC_TYPE_FLAC, CODEC_TYPE_MP3, CODEC_TYPE_OPUS, CODEC_TYPE_PCM_S16BE,
        CODEC_TYPE_PCM_S16LE, CODEC_TYPE_PCM_U8, CODEC_TYPE_VORBIS, CODEC_TYPE_WMA,
    };

    if codec == CODEC_TYPE_OPUS {
        "opus"
    } else if codec == CODEC_TYPE_AAC {
        "aac"
    } else if codec == CODEC_TYPE_VORBIS {
        "vorbis"
    } else if codec == CODEC_TYPE_MP3 {
        "mp3"
    } else if codec == CODEC_TYPE_FLAC {
        "flac"
    } else if codec == CODEC_TYPE_WMA {
        "wma"
    } else if codec == CODEC_TYPE_PCM_S16LE {
        "pcm_s16le"
    } else if codec == CODEC_TYPE_PCM_S16BE {
        "pcm_s16be"
    } else if codec == CODEC_TYPE_PCM_U8 {
        "pcm_u8"
    } else {
        "unknown"
    }
}

fn log_mkv_probe_once(path: &Path) {
    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    if extension != "mkv" {
        return;
    }

    static LOGGED: OnceLock<Mutex<HashSet<PathBuf>>> = OnceLock::new();
    let logged = LOGGED.get_or_init(|| Mutex::new(HashSet::new()));
    {
        let mut guard = match logged.lock() {
            Ok(guard) => guard,
            Err(_) => {
                log_debug("Audio probe: failed to lock log set.");
                return;
            }
        };
        if !guard.insert(path.to_path_buf()) {
            return;
        }
    }

    let file = match std::fs::File::open(path) {
        Ok(file) => file,
        Err(e) => {
            log_debug(&format!("Audio probe: failed to open file: {}", e));
            return;
        }
    };
    let mss = symphonia::core::io::MediaSourceStream::new(
        Box::new(file),
        symphonia::core::io::MediaSourceStreamOptions::default(),
    );
    let mut hint = symphonia::core::probe::Hint::new();
    hint.with_extension("mkv");

    let probed = symphonia::default::get_probe().format(
        &hint,
        mss,
        &symphonia::core::formats::FormatOptions::default(),
        &symphonia::core::meta::MetadataOptions::default(),
    );

    match probed {
        Ok(probed) => {
            let format = probed.format;
            for track in format.tracks() {
                let params = &track.codec_params;
                let channels = params.channels.map(|c| c.count());
                log_debug(&format!(
                    "Audio probe: track={} codec={} ({}) rate={:?} ch={:?}",
                    track.id,
                    codec_name(params.codec),
                    params.codec,
                    params.sample_rate,
                    channels
                ));
            }
        }
        Err(err) => {
            log_debug(&format!("Audio probe: failed to parse MKV: {}", err));
        }
    }
}

fn ffmpeg_cache_key(path: &Path) -> String {
    let mut hasher = sha2::Sha256::new();
    hasher.update(path.to_string_lossy().as_bytes());
    if let Ok(meta) = std::fs::metadata(path) {
        hasher.update(meta.len().to_le_bytes());
        if let Ok(modified) = meta.modified()
            && let Ok(duration) = modified.duration_since(UNIX_EPOCH)
        {
            hasher.update(duration.as_secs().to_le_bytes());
            hasher.update(duration.subsec_nanos().to_le_bytes());
        }
    }
    hex::encode(hasher.finalize())
}

fn write_wav_header<W: Write + Seek>(
    writer: &mut W,
    data_bytes: u64,
    channels: u16,
    sample_rate: u32,
) -> std::io::Result<()> {
    let data_bytes = data_bytes.min(u32::MAX as u64) as u32;
    let chunk_size = 36u32.saturating_add(data_bytes);
    let byte_rate = sample_rate
        .saturating_mul(channels as u32)
        .saturating_mul(2);
    let block_align = channels.saturating_mul(2);
    let bits_per_sample = 16u16;

    writer.seek(SeekFrom::Start(0))?;
    writer.write_all(b"RIFF")?;
    writer.write_all(&chunk_size.to_le_bytes())?;
    writer.write_all(b"WAVE")?;
    writer.write_all(b"fmt ")?;
    writer.write_all(&16u32.to_le_bytes())?;
    writer.write_all(&1u16.to_le_bytes())?;
    writer.write_all(&channels.to_le_bytes())?;
    writer.write_all(&sample_rate.to_le_bytes())?;
    writer.write_all(&byte_rate.to_le_bytes())?;
    writer.write_all(&block_align.to_le_bytes())?;
    writer.write_all(&bits_per_sample.to_le_bytes())?;
    writer.write_all(b"data")?;
    writer.write_all(&data_bytes.to_le_bytes())?;
    Ok(())
}

fn decode_ffmpeg_to_wav(path: &Path, stream_index: Option<i32>) -> Result<PathBuf, String> {
    let cache_dir = settings_dir().join("ffmpeg_cache");
    std::fs::create_dir_all(&cache_dir)
        .map_err(|e| format!("FFmpeg cache dir create failed: {}", e))?;
    // Include stream_index in cache key to separate different audio tracks
    let mut key = ffmpeg_cache_key(path);
    if let Some(idx) = stream_index {
        key.push_str(&format!("_s{}", idx));
    }
    let wav_path = cache_dir.join(format!("{}.wav", &key[..key.len().min(20)]));
    if wav_path.exists() {
        log_debug(&format!("FFmpeg: using cached WAV {}", wav_path.display()));
        return Ok(wav_path);
    }

    log_debug(&format!(
        "FFmpeg: decoding {} (stream {:?}) to WAV",
        path.display(),
        stream_index
    ));
    let mut source = FfmpegSource::try_new(path, 0, None, stream_index)?;
    let sample_rate = source.sample_rate();
    let channels = source.channels();
    log_debug(&format!(
        "FFmpeg: source rate={} channels={}",
        sample_rate, channels
    ));
    if sample_rate == 0 || channels == 0 {
        return Err("FFmpeg: invalid output format".to_string());
    }

    let file = File::create(&wav_path).map_err(|e| format!("FFmpeg: wav create failed: {}", e))?;
    let mut writer = std::io::BufWriter::with_capacity(65536, file);
    write_wav_header(&mut writer, 0, channels, sample_rate)
        .map_err(|e| format!("FFmpeg: wav header failed: {}", e))?;

    let mut data_bytes: u64 = 0;
    let mut buffer = Vec::with_capacity(8192);
    for sample in source.by_ref() {
        let scaled = (sample * i16::MAX as f32).round();
        let clamped = scaled.clamp(i16::MIN as f32, i16::MAX as f32) as i16;
        buffer.extend_from_slice(&clamped.to_le_bytes());
        data_bytes = data_bytes.saturating_add(2);

        // Write in chunks for better performance
        if buffer.len() >= 8192 {
            writer
                .write_all(&buffer)
                .map_err(|e| format!("FFmpeg: wav write failed: {}", e))?;
            buffer.clear();
        }
    }
    // Write remaining samples
    if !buffer.is_empty() {
        writer
            .write_all(&buffer)
            .map_err(|e| format!("FFmpeg: wav write failed: {}", e))?;
    }
    writer
        .flush()
        .map_err(|e| format!("FFmpeg: wav flush failed: {}", e))?;

    log_debug(&format!("FFmpeg: decoded {} bytes to WAV", data_bytes));
    if data_bytes == 0 {
        // Remove empty WAV file
        crate::log_if_err!(std::fs::remove_file(&wav_path));
        return Err("FFmpeg: no audio samples decoded".to_string());
    }

    let mut file = writer
        .into_inner()
        .map_err(|e| format!("FFmpeg: into_inner failed: {}", e))?;
    write_wav_header(&mut file, data_bytes, channels, sample_rate)
        .map_err(|e| format!("FFmpeg: wav finalize failed: {}", e))?;
    Ok(wav_path)
}

pub fn parse_time_input(input: &str) -> Result<u64, String> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err("empty".to_string());
    }
    if trimmed.chars().all(|c| c.is_ascii_digit()) {
        return trimmed.parse::<u64>().map_err(|_| "invalid".to_string());
    }
    if trimmed.contains(':') {
        let parts: Vec<&str> = trimmed.split(':').collect();
        if parts.len() == 2 || parts.len() == 3 {
            let mut nums = Vec::with_capacity(parts.len());
            for part in parts {
                let part = part.trim();
                if part.is_empty() || !part.chars().all(|c| c.is_ascii_digit()) {
                    return Err("invalid".to_string());
                }
                nums.push(part.parse::<u64>().map_err(|_| "invalid".to_string())?);
            }
            if nums.len() == 2 {
                let minutes = nums[0];
                let seconds = nums[1];
                if seconds >= 60 {
                    return Err("invalid".to_string());
                }
                return Ok(minutes * 60 + seconds);
            }
            let hours = nums[0];
            let minutes = nums[1];
            let seconds = nums[2];
            if minutes >= 60 || seconds >= 60 {
                return Err("invalid".to_string());
            }
            return Ok(hours * 3600 + minutes * 60 + seconds);
        }
    }
    Err("invalid".to_string())
}

pub fn audiobook_duration_secs(path: &Path) -> Option<u64> {
    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    if extension == "mp3"
        && let Ok(d) = mp3_duration::from_path(path)
    {
        return Some(d.as_secs());
    }

    // Use alias if needed for MF
    let mut final_path = path.to_path_buf();
    if (extension == "mp4" || extension == "aac" || extension == "mp3")
        && !path.to_string_lossy().contains("podcast cache")
    {
        let cache_dir = settings_dir().join("podcast cache");
        let mut hasher = sha2::Sha256::new();
        use sha2::Digest;
        hasher.update(path.to_string_lossy().as_bytes());
        let hash = hex::encode(hasher.finalize());
        let linked_path = cache_dir.join(format!("link_{}.m4a", &hash[..16]));
        if linked_path.exists() {
            final_path = linked_path;
        }
    }

    if let Ok(dur) = crate::mf_encoder::get_audio_duration_mf(&final_path) {
        return Some(dur);
    }
    let file = std::fs::File::open(path).ok()?;
    let source: Decoder<_> = Decoder::new(std::io::BufReader::new(file)).ok()?;
    if let Some(dur) = source.total_duration() {
        return Some(dur.as_secs());
    }
    if extension != "mp3" {
        // Already tried for mp3
        mp3_duration::from_path(path).ok().map(|d| d.as_secs())
    } else {
        None
    }
}

#[derive(Clone, Copy)]
struct AudiobookPlaybackOptions {
    speed: f32,
    pitch: f32,
    paused: bool,
    volume: f32,
    muted: bool,
    prev_volume: f32,
    mix_export: bool,
    audio_track: Option<i32>,
}

fn start_audiobook_at_with_options(
    hwnd: HWND,
    path: PathBuf,
    seconds: u64,
    options: AudiobookPlaybackOptions,
) {
    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let settings =
            unsafe { with_state(hwnd_main, |state| state.settings.clone()) }.unwrap_or_default();
        let want_mix = (settings.subtitle_read_mode == SubtitleReadMode::Record
            || options.mix_export)
            && crate::subtitles::find_subtitle_for_media(&path).is_some();
        if want_mix
            && !is_mixed_output(&path)
            && settings.subtitle_read_mode != SubtitleReadMode::Off
        {
            let path_clone = path.clone();
            let opts = options;
            std::thread::spawn(move || {
                if settings.subtitle_read_mode == SubtitleReadMode::Record
                    && !confirm_edge_subtitle_download(hwnd_main, &path_clone, &settings)
                {
                    log_debug("Subtitle: Edge download not confirmed for record mode.");
                    return;
                }
                log_debug(&format!(
                    "Subtitle: offline mix requested for {}",
                    path_clone.display()
                ));
                let mix_opts = MixExportOptions {
                    ducking: settings.subtitle_mix_ducking,
                };
                match export_mixed_media(&path_clone, &settings, &mix_opts) {
                    Ok(mixed_path) => {
                        start_audiobook_at_with_options(
                            hwnd_main,
                            mixed_path,
                            seconds,
                            AudiobookPlaybackOptions {
                                mix_export: false,
                                ..opts
                            },
                        );
                    }
                    Err(err) => {
                        log_debug(&format!("Subtitle: mix export failed: {}", err));
                        start_audiobook_at_with_options(
                            hwnd_main,
                            path_clone,
                            seconds,
                            AudiobookPlaybackOptions {
                                mix_export: false,
                                ..opts
                            },
                        );
                    }
                }
            });
            return;
        }

        let subtitle_hold = should_hold_for_edge_subtitles(hwnd_main, &path);
        let effective_paused = options.paused || subtitle_hold;
        let subtitle_cancel = Arc::new(AtomicBool::new(false));
        let subtitle_speech_cancel = Arc::new(Mutex::new(None));
        let subtitle_speech_command = Arc::new(Mutex::new(None));
        let subtitle_path = path.clone();
        let requested_speed = options.speed;
        log_debug(&format!(
            "Audio player: Thread started for {}",
            path.display()
        ));
        log_debug(&format!("Audio player: Opening file {}", path.display()));
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        let subtitle_mode =
            unsafe { with_state(hwnd_main, |state| state.settings.subtitle_read_mode) }
                .unwrap_or(SubtitleReadMode::Off);
        let effective_subtitle_mode = if subtitle_mode == SubtitleReadMode::Record {
            SubtitleReadMode::Off
        } else {
            subtitle_mode
        };
        let cached_subtitles = get_or_load_subtitles(&path, effective_subtitle_mode);
        let subtitles_available = cached_subtitles
            .as_ref()
            .map(|entry| !entry.cues.is_empty())
            .unwrap_or(false);
        let _subtitles_active =
            effective_subtitle_mode != SubtitleReadMode::Off && subtitles_available;
        let effective_speed = requested_speed;
        log_mkv_probe_once(&path);

        // Force M4A extension for Media Foundation to avoid indexing hangs on .mp4 video files
        // and to speed up .mp3/.aac opening.
        let mut final_path = path.clone();
        if (extension == "mp4" || extension == "aac" || extension == "mp3")
            && !is_mixed_output(&path)
            && !path.to_string_lossy().contains("podcast cache")
        {
            let cache_dir = settings_dir().join("podcast cache");
            std::fs::create_dir_all(&cache_dir).ok();
            let mut hasher = sha2::Sha256::new();
            hasher.update(path.to_string_lossy().as_bytes());
            let hash = hex::encode(hasher.finalize());
            let linked_path = cache_dir.join(format!("link_{}.m4a", &hash[..16]));

            if !linked_path.exists() {
                // Try hard link first (instant, no space)
                if let Err(e) = std::fs::hard_link(&path, &linked_path)
                    && e.kind() != std::io::ErrorKind::AlreadyExists
                {
                    // Fallback to copy if on different volume (slow, but only once)
                    if std::fs::copy(&path, &linked_path).is_err() {}
                }
            }

            if linked_path.exists() {
                final_path = linked_path;
            }
        }

        log_debug("Audio player: Opening with BASS...");
        let initial_volume = if options.muted { 0.0 } else { options.volume };
        let output = match BassOutput::start(
            &final_path,
            seconds,
            effective_speed,
            options.pitch,
            initial_volume,
            effective_paused,
        ) {
            Ok(output) => output,
            Err(err) => {
                log_debug(&format!("Audio player: BASS open failed: {}", err));
                // Try FFmpeg streaming first (instant playback)
                match BassOutput::start_with_ffmpeg(
                    &final_path,
                    seconds,
                    effective_speed,
                    options.pitch,
                    initial_volume,
                    effective_paused,
                    options.audio_track,
                ) {
                    Ok(output) => {
                        log_debug("Audio player: using FFmpeg streaming");
                        output
                    }
                    Err(stream_err) => {
                        log_debug(&format!(
                            "Audio player: FFmpeg streaming failed: {}",
                            stream_err
                        ));
                        // Fallback to WAV decode (slower but reliable)
                        match decode_ffmpeg_to_wav(&final_path, options.audio_track) {
                            Ok(wav_path) => match BassOutput::start(
                                &wav_path,
                                seconds,
                                effective_speed,
                                options.pitch,
                                initial_volume,
                                effective_paused,
                            ) {
                                Ok(output) => output,
                                Err(err) => {
                                    log_debug(&format!(
                                        "Audio player: BASS fallback failed: {}",
                                        err
                                    ));
                                    return;
                                }
                            },
                            Err(err) => {
                                log_debug(&format!(
                                    "Audio player: FFmpeg fallback failed: {}",
                                    err
                                ));
                                return;
                            }
                        }
                    }
                }
            }
        };

        log_debug("Audio player: Playback started");

        let session_id = unsafe {
            with_state(hwnd_main, |state| {
                state.audiobook_session_id = state.audiobook_session_id.wrapping_add(1);
                state.audiobook_session_id
            })
            .unwrap_or(0)
        };
        // Speed check removed: allowing BASS tempo to handle speed with subtitles.
        /*
        if subtitles_active && (requested_speed - 1.0).abs() > f32::EPSILON {
            log_debug(
                "Subtitles active: forcing speed=1.0 (time-stretch disabled) for accurate sync.",
            );
        }
        */

        let player = AudiobookPlayer {
            path,
            output,
            is_paused: effective_paused,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume: options.volume,
            muted: options.muted,
            prev_volume: options.prev_volume,
            speed: effective_speed,
            pitch: options.pitch,
            subtitle_cancel: subtitle_cancel.clone(),
            subtitle_hold,
            subtitle_speech_cancel: subtitle_speech_cancel.clone(),
            subtitle_speech_command: subtitle_speech_command.clone(),
            session_id,
        };

        unsafe {
            if with_state(hwnd_main, |state| {
                state.active_audiobook = Some(player);
            })
            .is_none()
            {
                crate::log_debug("Failed to access audio player state");
            }
        }
        start_subtitle_reader(hwnd_main, subtitle_path, subtitle_cancel, cached_subtitles);
    });
}

pub unsafe fn start_audiobook_playback(hwnd: HWND, path: &Path) {
    // Record telemetry
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("audio");
    crate::telemetry::record_action("audio_play", ext);
    crate::telemetry::set_audio_playing(true);

    crate::log_debug(&format!(
        "Audio player: start_audiobook_playback called for {}",
        path.display()
    ));
    crate::reset_active_podcast_chapters_for_playback(hwnd);
    let path_buf = path.to_path_buf();

    // List available audio tracks and store them in state
    let audio_tracks = match crate::ffmpeg_source::list_audio_streams(path) {
        Ok(tracks) => {
            log_debug(&format!(
                "Audio player: found {} audio tracks",
                tracks.len()
            ));
            for track in &tracks {
                log_debug(&format!(
                    "  Track {}: {:?} {:?} {} {}ch {}Hz default={}",
                    track.index,
                    track.language,
                    track.title,
                    track.codec,
                    track.channels,
                    track.sample_rate,
                    track.is_default
                ));
            }
            tracks
        }
        Err(e) => {
            log_debug(&format!("Audio player: failed to list audio tracks: {}", e));
            Vec::new()
        }
    };

    // Store tracks in state and clear any previous selection for new files
    with_state(hwnd, |state| {
        state.available_audio_tracks = audio_tracks;
        state.selected_audio_track = None;
    });

    let (bookmark_pos, speed, pitch, volume, mix_export) = with_state(hwnd, |state| {
        let pos = state
            .bookmarks
            .files
            .get(&path_buf.to_string_lossy().to_string())
            .and_then(|list| list.last())
            .map(|bm| bm.position)
            .unwrap_or(0);
        (
            pos,
            state.settings.audiobook_playback_speed,
            state.settings.audiobook_playback_pitch,
            state.settings.audiobook_playback_volume,
            state.settings.subtitle_read_mode == SubtitleReadMode::Record,
        )
    })
    .unwrap_or((0, 1.0, 0.0, 1.0, false));

    start_audiobook_at_with_options(
        hwnd,
        path_buf,
        bookmark_pos as u64,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused: false,
            volume,
            muted: false,
            prev_volume: volume,
            mix_export,
            audio_track: None,
        },
    );
}

pub unsafe fn toggle_audiobook_pause(hwnd: HWND) {
    crate::log_debug("Audio player: toggle_audiobook_pause triggered");
    let start_action = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.is_paused {
                crate::log_debug("Audio player: Resuming playback");
                player.play();
                player.is_paused = false;
                player.start_instant = std::time::Instant::now();
            } else {
                crate::log_debug("Audio player: Pausing playback");
                player.pause();
                player.is_paused = true;
                let pos = audiobook_position_secs(player);
                player.accumulated_seconds = pos.floor() as u64;
            }
            return None;
        }

        let doc = state.docs.get(state.current)?;
        if !matches!(doc.format, FileFormat::Audiobook) {
            return None;
        }
        let path = doc.path.clone()?;
        let from_start = state
            .last_stopped_audiobook
            .as_ref()
            .map(|p| p == &path)
            .unwrap_or(false);
        if from_start {
            state.last_stopped_audiobook = None;
        }
        Some((path, from_start))
    })
    .flatten();

    if let Some((path, from_start)) = start_action {
        if from_start {
            start_audiobook_at(hwnd, &path, 0);
        } else {
            start_audiobook_playback(hwnd, &path);
        }
    }
}

pub unsafe fn seek_audiobook(hwnd: HWND, seconds: i64) {
    enum SeekAction {
        Restart {
            path: PathBuf,
            current_pos: u64,
            duration: Option<u64>,
            speed: f32,
            pitch: f32,
            paused: bool,
            volume: f32,
            muted: bool,
            prev_volume: f32,
        },
    }

    let result = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            stop_shared_subtitle_speech(
                &player.subtitle_speech_cancel,
                &player.subtitle_speech_command,
                "seek",
            );
            let current_pos = audiobook_position_secs(player);
            if !player.is_paused {
                player.start_instant = std::time::Instant::now();
            }
            let new_pos = (current_pos as i64 + seconds).max(0);
            // Output doesn't support direct seek, always restart
            player.accumulated_seconds = new_pos as u64;
            Some(SeekAction::Restart {
                path: player.path.clone(),
                current_pos: new_pos as u64,
                duration: audiobook_duration_secs(&player.path),
                speed: player.speed,
                pitch: player.pitch,
                paused: player.is_paused,
                volume: player.volume,
                muted: player.muted,
                prev_volume: player.prev_volume,
            })
        } else {
            None
        }
    })
    .flatten();

    let action = match result {
        Some(v) => v,
        None => return,
    };

    let SeekAction::Restart {
        path,
        current_pos,
        duration,
        speed,
        pitch,
        paused,
        volume,
        muted,
        prev_volume,
    } = action;

    if let Some(duration) = duration
        && current_pos >= duration
    {
        stop_audiobook_playback(hwnd);
        return;
    }

    let audio_track = with_state(hwnd, |state| state.selected_audio_track).flatten();
    stop_audiobook_playback(hwnd);
    start_audiobook_at_with_options(
        hwnd,
        path,
        current_pos,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track,
        },
    );
}

pub unsafe fn seek_audiobook_to(hwnd: HWND, seconds: u64) -> Result<(), String> {
    let path = with_state(hwnd, |state| {
        if let Some(player) = &state.active_audiobook {
            stop_shared_subtitle_speech(
                &player.subtitle_speech_cancel,
                &player.subtitle_speech_command,
                "seek",
            );
            Some(player.path.clone())
        } else {
            None
        }
    })
    .flatten()
    .ok_or_else(|| "No active audiobook".to_string())?;

    if let Some(duration) = audiobook_duration_secs(&path)
        && seconds >= duration
    {
        stop_audiobook_playback(hwnd);
        return Ok(());
    }

    // Output doesn't support direct seek, always restart
    start_audiobook_at(hwnd, &path, seconds);
    Ok(())
}

pub unsafe fn stop_audiobook_playback(hwnd: HWND) {
    crate::telemetry::set_audio_playing(false);
    crate::log_debug("Audio player: stop_audiobook_playback called");
    if with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            crate::log_debug(&format!(
                "Audio player: Stopping and removing player for {}",
                player.path.display()
            ));
            stop_shared_subtitle_speech(
                &player.subtitle_speech_cancel,
                &player.subtitle_speech_command,
                "stop",
            );
            state.last_stopped_audiobook = Some(player.path.clone());
            player.subtitle_cancel.store(true, Ordering::Relaxed);
            player.stop();
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }
}

pub unsafe fn start_audiobook_at(hwnd: HWND, path: &Path, seconds: u64) {
    crate::log_debug(&format!(
        "Audio player: start_audiobook_at called for {} at {}s",
        path.display(),
        seconds
    ));
    let (speed, pitch, volume, muted, prev_volume) = with_state(hwnd, |state| {
        if let Some(player) = &state.active_audiobook {
            (
                player.speed,
                player.pitch,
                player.volume,
                player.muted,
                player.prev_volume,
            )
        } else {
            (1.0, 0.0, 1.0, false, 1.0)
        }
    })
    .unwrap_or((1.0, 0.0, 1.0, false, 1.0));

    let audio_track = with_state(hwnd, |state| state.selected_audio_track).flatten();
    stop_audiobook_playback(hwnd);
    let path_buf = path.to_path_buf();
    start_audiobook_at_with_options(
        hwnd,
        path_buf,
        seconds,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused: false,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track,
        },
    );
}

pub unsafe fn change_audiobook_volume(hwnd: HWND, delta: f32) {
    let new_volume = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.muted {
                player.prev_volume = (player.prev_volume + delta).clamp(0.0, 6.0);
                return None;
            }
            player.volume = (player.volume + delta).clamp(0.0, 6.0);
            player.set_volume(player.volume);
            Some(player.volume)
        } else {
            None
        }
    })
    .flatten();

    if let Some(volume) = new_volume
        && with_state(hwnd, |state| {
            state.settings.audiobook_playback_volume = volume;
            crate::settings::save_settings(state.settings.clone());
        })
        .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }
}

pub unsafe fn change_audiobook_speed(hwnd: HWND, delta: f32) -> Option<f32> {
    // We allow speed change even with subtitles now, relying on BASS tempo.
    let result = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            let current = audiobook_position_secs(&player).floor() as u64;
            let new_speed = (player.speed + delta).clamp(0.5, 3.0);
            player.stop();
            Some((
                player.path,
                current,
                new_speed,
                player.pitch,
                player.is_paused,
                player.volume,
                player.muted,
                player.prev_volume,
            ))
        } else {
            None
        }
    })
    .flatten();

    let (path, current, speed, pitch, paused, volume, muted, prev_volume) = result?;
    let audio_track = with_state(hwnd, |state| state.selected_audio_track).flatten();

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track,
        },
    );

    // Save speed to settings
    if with_state(hwnd, |state| {
        state.settings.audiobook_playback_speed = speed;
        crate::settings::save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }

    Some(speed)
}

pub unsafe fn change_audiobook_pitch(hwnd: HWND, delta: f32) -> Option<f32> {
    // Pitch change via BASS tempo
    let result = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            let current = audiobook_position_secs(&player).floor() as u64;
            let new_pitch = (player.pitch + delta).clamp(-12.0, 12.0);
            player.stop();
            Some((
                player.path,
                current,
                player.speed,
                new_pitch,
                player.is_paused,
                player.volume,
                player.muted,
                player.prev_volume,
            ))
        } else {
            None
        }
    })
    .flatten();

    let (path, current, speed, pitch, paused, volume, muted, prev_volume) = result?;
    let audio_track = with_state(hwnd, |state| state.selected_audio_track).flatten();

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track,
        },
    );

    // Save pitch to settings
    if with_state(hwnd, |state| {
        state.settings.audiobook_playback_pitch = pitch;
        crate::settings::save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }

    Some(pitch)
}

pub unsafe fn reset_audiobook_speed(hwnd: HWND) -> Option<f32> {
    let result = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            let current = audiobook_position_secs(&player).floor() as u64;
            let new_speed = 1.0;
            player.stop();
            Some((
                player.path,
                current,
                new_speed,
                player.pitch,
                player.is_paused,
                player.volume,
                player.muted,
                player.prev_volume,
            ))
        } else {
            None
        }
    })
    .flatten();

    let (path, current, speed, pitch, paused, volume, muted, prev_volume) = result?;
    let audio_track = with_state(hwnd, |state| state.selected_audio_track).flatten();

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track,
        },
    );

    // Save speed to settings
    if with_state(hwnd, |state| {
        state.settings.audiobook_playback_speed = speed;
        crate::settings::save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }

    Some(speed)
}

pub unsafe fn reset_audiobook_pitch(hwnd: HWND) -> Option<f32> {
    let result = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            let current = audiobook_position_secs(&player).floor() as u64;
            let new_pitch = 0.0;
            player.stop();
            Some((
                player.path,
                current,
                player.speed,
                new_pitch,
                player.is_paused,
                player.volume,
                player.muted,
                player.prev_volume,
            ))
        } else {
            None
        }
    })
    .flatten();

    let (path, current, speed, pitch, paused, volume, muted, prev_volume) = result?;
    let audio_track = with_state(hwnd, |state| state.selected_audio_track).flatten();

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track,
        },
    );

    // Save pitch to settings
    if with_state(hwnd, |state| {
        state.settings.audiobook_playback_pitch = pitch;
        crate::settings::save_settings(state.settings.clone());
    })
    .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }

    Some(pitch)
}

pub unsafe fn audiobook_volume_level(hwnd: HWND) -> Option<f32> {
    with_state(hwnd, |state| {
        state
            .active_audiobook
            .as_ref()
            .map(|player| if player.muted { 0.0 } else { player.volume })
    })
    .flatten()
}

/// Switch to a different audio track and restart playback.
pub unsafe fn switch_audio_track(hwnd: HWND, track_index: i32) {
    let result = with_state(hwnd, |state| {
        // Verify the track exists
        let valid = state
            .available_audio_tracks
            .iter()
            .any(|t| t.index == track_index);
        if !valid {
            return None;
        }
        state.selected_audio_track = Some(track_index);

        if let Some(player) = state.active_audiobook.take() {
            let current = audiobook_position_secs(&player).floor() as u64;
            player.stop();
            Some((
                player.path,
                current,
                player.speed,
                player.pitch,
                player.is_paused,
                player.volume,
                player.muted,
                player.prev_volume,
            ))
        } else {
            None
        }
    })
    .flatten();

    let Some((path, current, speed, pitch, paused, volume, muted, prev_volume)) = result else {
        return;
    };

    log_debug(&format!(
        "Audio player: switching to track {} at {}s",
        track_index, current
    ));

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            pitch,
            paused,
            volume,
            muted,
            prev_volume,
            mix_export: false,
            audio_track: Some(track_index),
        },
    );

    // Update the playback menu to reflect the new selection
    crate::menu::update_playback_menu(hwnd, true);
}

pub unsafe fn toggle_audiobook_mute(hwnd: HWND) {
    if with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.muted {
                let restored = if player.prev_volume > 0.0 {
                    player.prev_volume
                } else {
                    1.0
                };
                player.volume = restored;
                player.muted = false;
                player.set_volume(player.volume);
            } else {
                if player.volume > 0.0 {
                    player.prev_volume = player.volume;
                }
                player.volume = 0.0;
                player.muted = true;
                player.set_volume(0.0);
            }
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access audio player state");
    }
}

struct SubtitlePlaybackState {
    paused: bool,
    position_secs: f64,
    session_id: u64,
}

struct WaitableTimer {
    handle: HANDLE,
}

impl WaitableTimer {
    fn new() -> Option<Self> {
        match unsafe { CreateWaitableTimerW(None, true, None) } {
            Ok(handle) => {
                if handle.is_invalid() {
                    None
                } else {
                    Some(Self { handle })
                }
            }
            Err(e) => {
                log_debug(&format!("Subtitle: CreateWaitableTimerW failed: {}", e));
                None
            }
        }
    }
}

impl Drop for WaitableTimer {
    fn drop(&mut self) {
        unsafe {
            if let Err(e) = CloseHandle(self.handle) {
                log_debug(&format!("CloseHandle failed: {:?}", e));
            }
        }
    }
}

fn sleep_precise(timer: &WaitableTimer, duration: Duration) -> bool {
    if duration.is_zero() {
        return true;
    }
    let nanos = duration.as_nanos();
    let mut due = -((nanos / 100) as i64);
    if due == 0 {
        due = -1;
    }
    if let Err(e) = unsafe { SetWaitableTimer(timer.handle, &due, 0, None, None, false) } {
        log_debug(&format!("Subtitle: SetWaitableTimer failed: {}", e));
        return false;
    }
    let wait = unsafe { WaitForSingleObject(timer.handle, u32::MAX) };
    wait == WAIT_OBJECT_0
}

static SUBTITLE_EDGE_CONFIRMED: OnceLock<Mutex<HashSet<String>>> = OnceLock::new();
static SUBTITLE_CACHE: OnceLock<Mutex<HashMap<String, SubtitleCacheEntry>>> = OnceLock::new();
const SUBTITLE_PRELOAD_MAX_BYTES: Option<usize> = None;
const SUBTITLE_SCHEDULE_AHEAD_SECS: f64 = 10.0;

#[derive(Clone)]
struct SubtitleCacheEntry {
    subtitle_path: PathBuf,
    stamp: (u128, u64),
    cues: Vec<SubtitleCue>,
}

fn subtitle_mode_key(mode: SubtitleReadMode) -> &'static str {
    match mode {
        SubtitleReadMode::Off => "off",
        SubtitleReadMode::Nvda => "nvda",
        SubtitleReadMode::User => "user",
        SubtitleReadMode::Sapi5 => "sapi5",
        SubtitleReadMode::Sapi4 => "sapi4",
        SubtitleReadMode::Edge => "edge",
        SubtitleReadMode::Record => "record",
    }
}

fn subtitle_cache_key(media_path: &Path, mode: SubtitleReadMode) -> String {
    format!(
        "{}|{}",
        media_path.to_string_lossy(),
        subtitle_mode_key(mode)
    )
}

fn subtitle_file_stamp(path: &Path) -> Option<(u128, u64)> {
    let meta = std::fs::metadata(path).ok()?;
    let len = meta.len();
    let modified = meta.modified().ok()?;
    let duration = modified.duration_since(UNIX_EPOCH).ok()?;
    Some((duration.as_nanos(), len))
}

fn get_or_load_subtitles(media_path: &Path, mode: SubtitleReadMode) -> Option<SubtitleCacheEntry> {
    if mode == SubtitleReadMode::Off || mode == SubtitleReadMode::Record {
        return None;
    }
    let subtitle_path = find_subtitle_for_media(media_path)?;
    let stamp = subtitle_file_stamp(&subtitle_path)?;
    let key = subtitle_cache_key(media_path, mode);
    let cache = SUBTITLE_CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    if let Ok(map) = cache.lock()
        && let Some(entry) = map.get(&key)
        && entry.subtitle_path == subtitle_path
        && entry.stamp == stamp
    {
        return Some(entry.clone());
    }

    let cues = match load_subtitles(&subtitle_path) {
        Ok(cues) => cues,
        Err(err) => {
            log_debug(&format!("Subtitle: precheck failed: {}", err));
            return None;
        }
    };
    if cues.is_empty() {
        return None;
    }
    let entry = SubtitleCacheEntry {
        subtitle_path,
        stamp,
        cues,
    };
    if let Ok(mut map) = cache.lock() {
        map.insert(key, entry.clone());
    }
    Some(entry)
}

fn compute_audiobook_position_secs(
    output_pos: Option<f64>,
    is_paused: bool,
    accumulated_seconds: u64,
    elapsed: Duration,
    speed: f32,
) -> f64 {
    if let Some(pos) = output_pos {
        return pos.max(0.0);
    }
    if is_paused {
        accumulated_seconds as f64
    } else {
        accumulated_seconds as f64 + elapsed.as_secs_f64() * speed as f64
    }
}

pub(crate) fn audiobook_position_secs(player: &AudiobookPlayer) -> f64 {
    compute_audiobook_position_secs(
        player.output.position_secs(),
        player.is_paused,
        player.accumulated_seconds,
        player.start_instant.elapsed(),
        player.speed,
    )
}

fn subtitle_playback_state(hwnd: HWND, path: &Path) -> Option<SubtitlePlaybackState> {
    unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.path.as_path() != path {
                return None;
            }
            Some(SubtitlePlaybackState {
                paused: player.is_paused,
                position_secs: audiobook_position_secs(player),
                session_id: player.session_id,
            })
        })
        .flatten()
    }
}

/// Get the main audio output for subtitle mixing.
fn get_main_audio_output(hwnd: HWND, path: &Path) -> Option<Arc<BassOutput>> {
    unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.path.as_path() != path {
                return None;
            }
            Some(player.output.clone())
        })
        .flatten()
    }
}

fn subtitle_speech_handles(hwnd: HWND, path: &Path) -> Option<SubtitleSpeechHandles> {
    unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.path.as_path() != path {
                return None;
            }
            Some((
                Arc::clone(&player.subtitle_speech_cancel),
                Arc::clone(&player.subtitle_speech_command),
            ))
        })
        .flatten()
    }
}

fn stop_shared_subtitle_speech(
    cancel_store: &SubtitleSpeechCancel,
    command_store: &SubtitleSpeechCommand,
    reason: &str,
) {
    let cancel = cancel_store.lock().ok().and_then(|mut guard| guard.take());
    if let Some(cancel) = cancel {
        cancel.store(true, Ordering::SeqCst);
    }
    let command = command_store.lock().ok().and_then(|mut guard| guard.take());
    if let Some(tx) = command
        && let Err(err) = tx.send(TtsCommand::Stop)
    {
        log_debug(&format!("Subtitle: stop command failed: {}", err));
    }
    if !reason.is_empty() {
        log_debug(&format!("Subtitle: stopped active speech ({})", reason));
    }
}

fn subtitle_hold_state(hwnd: HWND, path: &Path) -> Option<bool> {
    unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.path.as_path() != path {
                return None;
            }
            Some(player.subtitle_hold)
        })
        .flatten()
    }
}

fn clear_subtitle_hold(hwnd: HWND, path: &Path) -> bool {
    unsafe {
        with_state(hwnd, |state| {
            let player = match state.active_audiobook.as_mut() {
                Some(player) => player,
                None => return false,
            };
            if player.path.as_path() != path {
                return false;
            }
            player.subtitle_hold = false;
            player.is_paused = false;
            player.start_instant = std::time::Instant::now();
            player.play();
            true
        })
        .unwrap_or(false)
    }
}

fn pause_active_backend(hwnd: HWND, path: &Path) -> bool {
    unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.path.as_path() != path {
                return None;
            }
            player.pause();
            Some(true)
        })
        .flatten()
        .unwrap_or(false)
    }
}

fn resume_active_backend(hwnd: HWND, path: &Path) -> bool {
    unsafe {
        with_state(hwnd, |state| {
            let player = state.active_audiobook.as_ref()?;
            if player.path.as_path() != path {
                return None;
            }
            player.play();
            Some(true)
        })
        .flatten()
        .unwrap_or(false)
    }
}

fn parse_sapi4_voice_index(voice: &str) -> Option<i32> {
    let rest = voice.strip_prefix("SAPI4#")?;
    let idx = rest.split('|').next()?;
    idx.parse::<i32>().ok()
}

fn should_hold_for_edge_subtitles(hwnd: HWND, media_path: &Path) -> bool {
    let settings = match unsafe { with_state(hwnd, |state| state.settings.clone()) } {
        Some(settings) => settings,
        None => return false,
    };
    if settings.subtitle_read_mode != SubtitleReadMode::User {
        return false;
    }
    if settings.tts_engine != crate::settings::TtsEngine::Edge {
        return false;
    }
    let _subtitle_path = match find_subtitle_for_media(media_path) {
        Some(path) => path,
        None => return false,
    };
    // subtitle_path will be dropped at end of scope
    let base_cache_dir = settings_dir().join("subtitle_cache");
    let mut hasher = sha2::Sha256::new();
    hasher.update(media_path.to_string_lossy().as_bytes());
    hasher.update(settings.tts_voice.as_bytes());
    let hash = hex::encode(hasher.finalize());
    let dir = base_cache_dir.join(&hash[..16]);
    if !dir.exists() {
        return true;
    }
    let cache_ready = edge_cache_has_valid_mp3(&dir);
    if !cache_ready {
        return true;
    }
    false
}

fn edge_subtitle_key(media_path: &Path, settings: &crate::settings::AppSettings) -> String {
    format!(
        "{}|{}|{}|{}|{}",
        media_path.to_string_lossy(),
        settings.tts_voice,
        settings.tts_rate,
        settings.tts_pitch,
        settings.tts_volume
    )
}

fn edge_cache_has_valid_mp3(dir: &Path) -> bool {
    let entries = match std::fs::read_dir(dir) {
        Ok(entries) => entries,
        Err(_) => return false,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        let is_mp3 = path
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.eq_ignore_ascii_case("mp3"))
            .unwrap_or(false);
        if !is_mp3 {
            continue;
        }
        let size = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        if size > 0 {
            return true;
        }
    }
    false
}

fn confirm_edge_subtitle_download(
    hwnd: HWND,
    media_path: &Path,
    settings: &crate::settings::AppSettings,
) -> bool {
    if settings.tts_engine != crate::settings::TtsEngine::Edge {
        return true;
    }
    let edge_key = edge_subtitle_key(media_path, settings);
    if is_edge_confirmed(&edge_key) {
        return true;
    }
    let msg = i18n::tr(settings.language, "subtitles.edge_confirm");
    let title = confirm_title(settings.language);
    let msg_w = to_wide(&msg);
    let title_w = to_wide(&title);
    let response = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(msg_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    if response != IDYES {
        return false;
    }
    mark_edge_confirmed(&edge_key);
    true
}

fn mark_edge_confirmed(key: &str) {
    let confirmed = SUBTITLE_EDGE_CONFIRMED.get_or_init(|| Mutex::new(HashSet::new()));
    if let Ok(mut set) = confirmed.lock() {
        set.insert(key.to_string());
    }
}

fn is_edge_confirmed(key: &str) -> bool {
    let confirmed = SUBTITLE_EDGE_CONFIRMED.get_or_init(|| Mutex::new(HashSet::new()));
    confirmed
        .lock()
        .map(|set| set.contains(key))
        .unwrap_or(false)
}

fn start_subtitle_reader(
    hwnd: HWND,
    media_path: PathBuf,
    cancel: Arc<AtomicBool>,
    cached: Option<SubtitleCacheEntry>,
) {
    std::thread::spawn(move || {
        if let Err(e) = unsafe { SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_HIGHEST) } {
            log_debug(&format!("Subtitle: SetThreadPriority failed: {}", e));
        }
        let settings = match unsafe { with_state(hwnd, |state| state.settings.clone()) } {
            Some(settings) => settings,
            None => {
                log_debug("Subtitle: Failed to access settings.");
                return;
            }
        };
        let mode = settings.subtitle_read_mode;
        if mode == SubtitleReadMode::Off {
            return;
        }
        let session_id = match subtitle_playback_state(hwnd, &media_path) {
            Some(state) => state.session_id,
            None => return,
        };
        let effective_mode = match mode {
            SubtitleReadMode::Off => SubtitleReadMode::Off,
            SubtitleReadMode::Nvda => SubtitleReadMode::Nvda,
            SubtitleReadMode::Record => SubtitleReadMode::Off,
            _ => SubtitleReadMode::User,
        };
        let (subtitle_path, mut cues) = if let Some(entry) = cached {
            (entry.subtitle_path.clone(), entry.cues.clone())
        } else {
            let subtitle_path = match find_subtitle_for_media(&media_path) {
                Some(path) => path,
                None => return,
            };
            let cues = match load_subtitles(&subtitle_path) {
                Ok(cues) => cues,
                Err(err) => {
                    log_debug(&format!("Subtitle: {}", err));
                    return;
                }
            };
            (subtitle_path, cues)
        };
        if cues.is_empty() {
            return;
        }
        if let (Some(first), Some(last)) = (cues.first(), cues.last()) {
            log_debug(&format!(
                "Subtitle: loaded {} cues from {} (first {:.3}-{:.3}s, last {:.3}-{:.3}s)",
                cues.len(),
                subtitle_path.display(),
                first.start.as_secs_f64(),
                first.end.as_secs_f64(),
                last.start.as_secs_f64(),
                last.end.as_secs_f64()
            ));
        }

        let main_output = get_main_audio_output(hwnd, &media_path);
        let speech_handles = subtitle_speech_handles(hwnd, &media_path);

        let mut paused_for_download = subtitle_hold_state(hwnd, &media_path).unwrap_or(false);
        if effective_mode == SubtitleReadMode::User
            && settings.tts_engine == crate::settings::TtsEngine::Edge
        {
            log_debug(&format!(
                "Subtitle: Edge TTS enabled for {} (settings_dir={})",
                media_path.display(),
                settings_dir().display()
            ));
            let edge_key = edge_subtitle_key(&media_path, &settings);
            let msg = i18n::tr(settings.language, "subtitles.edge_confirm");
            let title = confirm_title(settings.language);
            let msg_w = to_wide(&msg);
            let title_w = to_wide(&title);
            let mut paused_for_prompt = false;
            if !paused_for_download
                && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                && !state.paused
                && pause_active_backend(hwnd, &media_path)
            {
                paused_for_download = true;
                paused_for_prompt = true;
            }
            if !is_edge_confirmed(&edge_key) {
                let response = unsafe {
                    MessageBoxW(
                        hwnd,
                        PCWSTR(msg_w.as_ptr()),
                        PCWSTR(title_w.as_ptr()),
                        MB_YESNO | MB_ICONQUESTION,
                    )
                };
                if response != IDYES {
                    if paused_for_prompt
                        && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                        && !state.paused
                    {
                        resume_active_backend(hwnd, &media_path);
                    }
                    return;
                }
                mark_edge_confirmed(&edge_key);
            }

            let base_cache_dir = settings_dir().join("subtitle_cache");
            if let Err(e) = std::fs::create_dir_all(&base_cache_dir) {
                log_debug(&format!("Subtitle: cache dir create failed: {}", e));
                return;
            }
            let mut hasher = sha2::Sha256::new();
            hasher.update(media_path.to_string_lossy().as_bytes());
            hasher.update(settings.tts_voice.as_bytes());
            let hash = hex::encode(hasher.finalize());
            let dir = base_cache_dir.join(&hash[..16]);
            let cache_ready = dir.exists() && edge_cache_has_valid_mp3(&dir);
            log_debug(&format!(
                "Subtitle: Edge cache status ready={} dir={} voice={}",
                cache_ready,
                dir.display(),
                settings.tts_voice
            ));
            if !cache_ready {
                if dir.exists()
                    && let Err(e) = std::fs::remove_dir_all(&dir)
                {
                    log_debug(&format!("Subtitle: cache cleanup failed: {}", e));
                }
                if let Err(e) = std::fs::create_dir_all(&dir) {
                    log_debug(&format!("Subtitle: cache dir create failed: {}", e));
                    return;
                }
            }
            if !cache_ready {
                log_debug(&format!(
                    "Subtitle: predownload async start ({} cues) for {}",
                    cues.len(),
                    subtitle_path.display()
                ));
                if let Some(state) = subtitle_playback_state(hwnd, &media_path)
                    && !state.paused
                    && pause_active_backend(hwnd, &media_path)
                {
                    paused_for_download = true;
                }

                let first_idx = cues
                    .iter()
                    .position(|cue| !cue.text.trim().is_empty())
                    .unwrap_or(0);
                let (first_tx, first_rx) = std::sync::mpsc::channel::<()>();
                let first_sent = Arc::new(AtomicBool::new(false));

                let mut jobs = Vec::new();
                for (idx, cue) in cues.iter_mut().enumerate() {
                    let text = cue.text.trim();
                    if text.is_empty() {
                        continue;
                    }
                    let path = dir.join(format!("cue_{:04}.mp3", idx));
                    cue.audio_path = Some(path.clone());
                    jobs.push((idx, cue.text.clone(), path));
                }
                let total_jobs = jobs.len();

                let cancel_bg = cancel.clone();
                let voice = settings.tts_voice.clone();
                let language = settings.language;
                let rate = settings.tts_rate;
                let pitch = settings.tts_pitch;
                let volume = settings.tts_volume;
                let record_mode = mode == SubtitleReadMode::Record;
                let completed = Arc::new(AtomicUsize::new(0));
                let first_sent_bg = first_sent.clone();
                let first_tx_bg = first_tx.clone();
                let completed_bg = completed.clone();
                std::thread::spawn(move || {
                    let rt = match tokio::runtime::Runtime::new() {
                        Ok(rt) => rt,
                        Err(e) => {
                            log_debug(&format!("Subtitle: failed to create runtime: {}", e));
                            return;
                        }
                    };
                    rt.block_on(async move {
                        const CUE_CONCURRENCY: usize = 30;
                        let cancel_loop = cancel_bg.clone();
                        let tasks = jobs.into_iter().map(|(idx, text, path)| {
                            let cancel = cancel_bg.clone();
                            let voice = voice.clone();
                            let first_sent_bg = first_sent_bg.clone();
                            let first_tx_bg = first_tx_bg.clone();
                            let completed_bg = completed_bg.clone();
                            async move {
                            let mut empty_attempts = 0u64;
                            let mut attempts = 0u64;
                            loop {
                                if cancel.load(Ordering::Relaxed) {
                                    return;
                                }
                                attempts = attempts.saturating_add(1);
                                if attempts == 1 {
                                    log_debug(&format!(
                                        "Subtitle: downloading Edge cue {} -> {}",
                                        idx,
                                        path.display()
                                    ));
                                }
                                let request_id = Uuid::new_v4().simple().to_string();
                                match tts_engine::download_audio_chunk(
                                    text.trim(),
                                    &voice,
                                    &request_id,
                                    rate,
                                    pitch,
                                    volume,
                                    language,
                                )
                                .await
                                {
                                    Ok(bytes) => {
                                        if bytes.is_empty() {
                                            empty_attempts = empty_attempts.saturating_add(1);
                                            log_debug(&format!(
                                                "Subtitle: empty Edge audio (attempt {}) for cue {}",
                                                empty_attempts, idx
                                            ));
                                            tokio::time::sleep(Duration::from_millis(200)).await;
                                            continue;
                                        }
                                        match std::fs::write(&path, bytes) {
                                            Ok(()) => {
                                                let size = std::fs::metadata(&path)
                                                    .map(|m| m.len())
                                                    .unwrap_or(0);
                                                log_debug(&format!(
                                                    "Subtitle: wrote Edge cue {} ({} bytes)",
                                                    idx, size
                                                ));
                                                if size == 0 {
                                                    empty_attempts = empty_attempts.saturating_add(1);
                                                    log_debug(&format!(
                                                        "Subtitle: zero-byte Edge audio (attempt {}) for {}",
                                                        empty_attempts,
                                                        path.display()
                                                    ));
                                                    if let Err(e) = std::fs::remove_file(&path) {
                                                        log_debug(&format!(
                                                            "Subtitle: failed to remove empty file: {}",
                                                            e
                                                        ));
                                                    }
                                                    tokio::time::sleep(Duration::from_millis(200)).await;
                                                    continue;
                                                }
                                                if idx == first_idx
                                                    && !first_sent_bg.load(Ordering::Relaxed)
                                                {
                                                    first_sent_bg.store(true, Ordering::Relaxed);
                                                    let _unused = first_tx_bg.send(());
                                                }
                                                let done = completed_bg.fetch_add(1, Ordering::Relaxed) + 1;
                                                if record_mode
                                                    && (done == total_jobs
                                                        || done.is_multiple_of(10))
                                                {
                                                    let msg = format!(
                                                        "{} sottotitoli di {} creati",
                                                        done, total_jobs
                                                    );
                                                    let _ = nvda_speak(&msg);
                                                }
                                            }
                                            Err(e) => {
                                                log_debug(&format!(
                                                    "Subtitle: failed to write audio chunk: {}",
                                                    e
                                                ));
                                            }
                                        }
                                        break;
                                    }
                                    Err(err) => {
                                        log_debug(&format!("Subtitle: download failed: {}", err));
                                        if attempts.is_multiple_of(5) {
                                            log_debug(&format!(
                                                "Subtitle: download retrying (attempt {}) for cue {}",
                                                attempts, idx
                                            ));
                                        }
                                        tokio::time::sleep(Duration::from_millis(500)).await;
                                    }
                                }
                            }
                            }
                        });
                        let mut stream =
                            futures_util::stream::iter(tasks).buffered(CUE_CONCURRENCY);
                        while let Some(()) = stream.next().await {
                            if cancel_loop.load(Ordering::Relaxed) {
                                break;
                            }
                        }
                    });
                });

                let start_wait = std::time::Instant::now();
                let mut first_ready = false;
                while !first_ready {
                    if cancel.load(Ordering::Relaxed) {
                        return;
                    }
                    if first_rx.try_recv().is_ok() {
                        first_ready = true;
                        break;
                    }
                    if start_wait.elapsed() >= Duration::from_secs(10) {
                        break;
                    }
                    std::thread::sleep(Duration::from_millis(100));
                }
                log_debug(&format!(
                    "Subtitle: first cue ready={} wait_ms={}",
                    first_ready,
                    start_wait.elapsed().as_millis()
                ));
            } else {
                for (idx, cue) in cues.iter_mut().enumerate() {
                    let path = dir.join(format!("cue_{:04}.mp3", idx));
                    if path.exists() {
                        cue.audio_path = Some(path);
                    }
                }
                let available = cues.iter().filter(|cue| cue.audio_path.is_some()).count();
                log_debug(&format!(
                    "Subtitle: cache ready ({}/{}) for {}",
                    available,
                    cues.len(),
                    subtitle_path.display()
                ));
                log_debug(&format!(
                    "Subtitle: cache dir {} (base {})",
                    dir.display(),
                    base_cache_dir.display()
                ));
                for (idx, cue) in cues.iter().enumerate().take(3) {
                    if let Some(path) = cue.audio_path.as_ref() {
                        let size = std::fs::metadata(path).map(|m| m.len()).unwrap_or(0);
                        log_debug(&format!(
                            "Subtitle: cache cue {} -> {} ({} bytes)",
                            idx,
                            path.display(),
                            size
                        ));
                    }
                }
            }

            let mut preload_bytes: usize = 0;
            let mut preload_count: usize = 0;
            let mut preload_skipped: usize = 0;
            for cue in cues.iter_mut() {
                let Some(path) = cue.audio_path.as_ref() else {
                    continue;
                };
                if let Some(max_bytes) = SUBTITLE_PRELOAD_MAX_BYTES
                    && preload_bytes >= max_bytes
                {
                    preload_skipped += 1;
                    continue;
                }
                match std::fs::read(path) {
                    Ok(bytes) => {
                        preload_bytes = preload_bytes.saturating_add(bytes.len());
                        cue.audio_data = Some(Arc::from(bytes));
                        preload_count += 1;
                    }
                    Err(e) => {
                        log_debug(&format!(
                            "Subtitle: preload failed for {}: {}",
                            path.display(),
                            e
                        ));
                    }
                }
            }
            log_debug(&format!(
                "Subtitle: preload complete ({}/{}) cached={} skipped={}",
                preload_count,
                cues.len(),
                preload_count,
                preload_skipped
            ));

            if cancel.load(Ordering::Relaxed) {
                return;
            }

            if paused_for_download
                && let Some(held) = subtitle_hold_state(hwnd, &media_path)
                && held
            {
                clear_subtitle_hold(hwnd, &media_path);
            } else if paused_for_download
                && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                && !state.paused
            {
                resume_active_backend(hwnd, &media_path);
            }
        }

        let offset_secs = settings.subtitle_offset_ms as f64 / 1000.0;
        let (mut index, mut last_position, mut last_paused) =
            if let Some(state) = subtitle_playback_state(hwnd, &media_path) {
                let raw_pos = state.position_secs;
                let index = cues
                    .iter()
                    .position(|cue| cue.end.as_secs_f64() >= raw_pos)
                    .unwrap_or(cues.len());
                (index, raw_pos, state.paused)
            } else {
                return;
            };
        let mut last_spoken_index: Option<usize> = None;
        let mut last_spoken_pos = 0.0f64;
        let mut last_spoken_text: Option<String> = None;
        let mut spoken_once: HashSet<usize> = HashSet::new();

        let timer = WaitableTimer::new();
        if timer.is_none() {
            log_debug("Subtitle: high-precision timer unavailable, falling back to sleep.");
        }

        fn try_read_edge_audio(path: &Path) -> Option<Vec<u8>> {
            if let Ok(bytes) = std::fs::read(path) {
                return Some(bytes);
            }
            for _ in 0..5 {
                std::thread::sleep(Duration::from_millis(100));
                if let Ok(bytes) = std::fs::read(path) {
                    return Some(bytes);
                }
            }
            None
        }

        let wait_until_target = |mut current_pos: f64, target: f64| -> Option<f64> {
            if current_pos >= target {
                return Some(current_pos);
            }
            loop {
                if cancel.load(Ordering::Relaxed) {
                    return None;
                }
                let wait_state = subtitle_playback_state(hwnd, &media_path)?;
                if wait_state.session_id != session_id {
                    return None;
                }
                if wait_state.paused {
                    std::thread::sleep(Duration::from_millis(20));
                    continue;
                }
                current_pos = wait_state.position_secs;
                if current_pos >= target {
                    return Some(current_pos);
                }
                let remaining = target - current_pos;
                if remaining > 0.01 {
                    let mut sleep_secs = remaining - 0.001;
                    sleep_secs = sleep_secs.clamp(0.0, 0.01);
                    if sleep_secs > 0.0 {
                        if let Some(ref timer) = timer {
                            if !sleep_precise(timer, Duration::from_secs_f64(sleep_secs)) {
                                std::thread::sleep(Duration::from_secs_f64(sleep_secs));
                            }
                        } else {
                            std::thread::sleep(Duration::from_secs_f64(sleep_secs));
                        }
                    } else {
                        std::thread::sleep(Duration::from_millis(5));
                    }
                } else {
                    std::hint::spin_loop();
                }
            }
        };

        loop {
            if cancel.load(Ordering::Relaxed) {
                break;
            }
            let state = match subtitle_playback_state(hwnd, &media_path) {
                Some(state) => state,
                None => break,
            };
            if state.session_id != session_id {
                break;
            }

            if state.paused {
                last_paused = true;
                std::thread::sleep(Duration::from_millis(80));
                continue;
            } else if last_paused {
                last_paused = false;
            }

            let mut raw_pos = state.position_secs;
            let seek_delta = raw_pos - last_position;
            if seek_delta.abs() > 1.0 {
                log_debug(&format!(
                    "SubtitleSync: seek detected at {:.3}s (prev {:.3}s) for {}",
                    raw_pos,
                    last_position,
                    media_path.display()
                ));
                // Seek interrupts are handled by the main thread.
                if let Some(pos) = cues.iter().position(|cue| cue.end.as_secs_f64() >= raw_pos) {
                    index = pos;
                } else {
                    index = cues.len();
                }
                last_spoken_index = None;
                spoken_once.clear();
                // Clear pending audio on seek
                if let Some(ref output) = main_output {
                    output.clear_subtitles();
                }
            }
            last_position = raw_pos;
            if raw_pos + 0.5 < last_spoken_pos {
                last_spoken_index = None;
                spoken_once.clear();
                last_spoken_text = None;
            }

            while index < cues.len() {
                let cue = cues[index].clone();
                let cue_start = cue.start.as_secs_f64();
                let cue_end = cue.end.as_secs_f64();

                // Simple unified timing for all backends:
                // Schedule/speak when we're within a window ahead of cue start
                // Target is exactly cue_start (+ user offset)
                let target = cue_start + offset_secs;
                let schedule_ready = raw_pos + SUBTITLE_SCHEDULE_AHEAD_SECS >= target;

                if !schedule_ready {
                    break;
                }

                // If we're already past the cue end, skip it.
                if raw_pos > cue_end {
                    index += 1;
                    continue;
                }
                if let Some(last) = last_spoken_index
                    && last == index
                    && (raw_pos - last_spoken_pos).abs() < 0.5
                {
                    index += 1;
                    continue;
                }
                if let Some(ref last_text) = last_spoken_text
                    && last_text == &cue.text
                    && (raw_pos - last_spoken_pos).abs() < 1.0
                {
                    index += 1;
                    continue;
                }
                if spoken_once.contains(&index) {
                    index += 1;
                    continue;
                }

                let delta_from_start = raw_pos - cue_start;
                let mut preview = cue.text.replace('\n', " ");
                if preview.len() > 80 {
                    preview.truncate(80);
                    preview.push_str("...");
                }
                log_debug(&format!(
                    "SubtitleSync: idx={} start={:.3}s raw={:.3}s delta={:.3}s offset={:.3}s text='{}'",
                    index, cue_start, raw_pos, delta_from_start, offset_secs, preview
                ));
                let mut did_emit = false;
                match effective_mode {
                    SubtitleReadMode::Off | SubtitleReadMode::Record => {}
                    SubtitleReadMode::Nvda => {
                        if let Some(pos) = wait_until_target(raw_pos, target) {
                            raw_pos = pos;
                        } else {
                            return;
                        }
                        if !nvda_speak(&cue.text) {
                            log_debug("Subtitle: NVDA speak failed.");
                        }
                        did_emit = true;
                    }
                    SubtitleReadMode::User
                    | SubtitleReadMode::Sapi5
                    | SubtitleReadMode::Sapi4
                    | SubtitleReadMode::Edge => match settings.tts_engine {
                        crate::settings::TtsEngine::Edge => {
                            if main_output.is_some() {
                                if let Some(pos) = wait_until_target(raw_pos, target) {
                                    raw_pos = pos;
                                } else {
                                    return;
                                }
                                if let Some(ref path) = cue.audio_path {
                                    if let Some(bytes) = try_read_edge_audio(path) {
                                        let cancel = crate::tts_engine::play_edge_bytes_async(
                                            bytes,
                                            settings.tts_volume,
                                        );
                                        if let Some((ref cancel_store, ref command_store)) =
                                            speech_handles
                                        {
                                            if let Ok(mut guard) = cancel_store.lock() {
                                                *guard = Some(cancel);
                                            }
                                            if let Ok(mut guard) = command_store.lock() {
                                                *guard = None;
                                            }
                                        }
                                        did_emit = true;
                                    } else {
                                        log_debug("Subtitle: Edge audio missing, skipping cue.");
                                    }
                                } else if let Some(ref audio) = cue.audio_data {
                                    let cancel = crate::tts_engine::play_edge_bytes_async(
                                        audio.to_vec(),
                                        settings.tts_volume,
                                    );
                                    if let Some((ref cancel_store, ref command_store)) =
                                        speech_handles
                                    {
                                        if let Ok(mut guard) = cancel_store.lock() {
                                            *guard = Some(cancel);
                                        }
                                        if let Ok(mut guard) = command_store.lock() {
                                            *guard = None;
                                        }
                                    }
                                    did_emit = true;
                                } else {
                                    log_debug("Subtitle: Edge audio missing, skipping cue.");
                                }
                            } else {
                                log_debug("Subtitle: main output missing, skipping cue.");
                            }
                        }
                        crate::settings::TtsEngine::Sapi5 => {
                            if let Some(pos) = wait_until_target(raw_pos, target) {
                                raw_pos = pos;
                            } else {
                                return;
                            }
                            let cancel_flag = Arc::new(AtomicBool::new(false));
                            let (command_tx, command_rx) = tokio::sync::mpsc::unbounded_channel();
                            if let Err(err) = crate::sapi5_engine::play_sapi(
                                vec![cue.text.clone()],
                                settings.tts_voice.clone(),
                                settings.tts_rate,
                                settings.tts_pitch,
                                settings.tts_volume,
                                cancel_flag.clone(),
                                command_rx,
                            ) {
                                log_debug(&format!("Subtitle: SAPI5 failed: {}", err));
                            } else if let Some((ref cancel_store, ref command_store)) =
                                speech_handles
                            {
                                if let Ok(mut guard) = cancel_store.lock() {
                                    *guard = Some(cancel_flag);
                                }
                                if let Ok(mut guard) = command_store.lock() {
                                    *guard = Some(command_tx);
                                }
                            }
                            did_emit = true;
                        }
                        crate::settings::TtsEngine::Sapi4 => {
                            if let Some(pos) = wait_until_target(raw_pos, target) {
                                raw_pos = pos;
                            } else {
                                return;
                            }
                            let voice_index = match parse_sapi4_voice_index(&settings.tts_voice) {
                                Some(idx) => idx,
                                None => {
                                    log_debug("Subtitle: invalid SAPI4 voice, defaulting to 0.");
                                    0
                                }
                            };
                            log_debug(&format!(
                                "Subtitle: SAPI4 speak voice='{}' idx={}",
                                settings.tts_voice, voice_index
                            ));
                            let cancel_flag = Arc::new(AtomicBool::new(false));
                            let (command_tx, command_rx) = tokio::sync::mpsc::unbounded_channel();
                            crate::sapi4_engine::play_sapi4(
                                voice_index,
                                cue.text.clone(),
                                settings.tts_rate,
                                settings.tts_pitch,
                                settings.tts_volume,
                                cancel_flag.clone(),
                                command_rx,
                            );
                            if let Some((ref cancel_store, ref command_store)) = speech_handles {
                                if let Ok(mut guard) = cancel_store.lock() {
                                    *guard = Some(cancel_flag.clone());
                                }
                                if let Ok(mut guard) = command_store.lock() {
                                    *guard = Some(command_tx.clone());
                                }
                            }
                            let stop_after = (cue_end - cue_start).max(0.5);
                            std::thread::spawn(move || {
                                std::thread::sleep(Duration::from_secs_f64(stop_after));
                                if let Err(err) = command_tx.send(TtsCommand::Stop) {
                                    log_debug(&format!("Subtitle: SAPI4 stop failed: {}", err));
                                }
                            });
                            did_emit = true;
                        }
                    },
                }
                if did_emit {
                    last_spoken_index = Some(index);
                    last_spoken_pos = raw_pos;
                    last_spoken_text = Some(cue.text.clone());
                    spoken_once.insert(index);
                }
                index += 1;
            }

            // Reduced polling for better timing precision
            std::thread::sleep(Duration::from_millis(10));
        }

        // Cleanup: clear any pending subtitles
        if let Some(ref output) = main_output {
            output.clear_subtitles();
        }
        if let Some((cancel_store, command_store)) = speech_handles {
            if let Ok(mut guard) = cancel_store.lock() {
                *guard = None;
            }
            if let Ok(mut guard) = command_store.lock() {
                *guard = None;
            }
        }
    });
}

#[cfg(test)]
mod tests {
    use super::{compute_audiobook_position_secs, parse_time_input};
    use std::time::Duration;

    #[test]
    fn parse_seconds() {
        assert_eq!(parse_time_input("90").unwrap(), 90);
    }

    #[test]
    fn parse_mm_ss() {
        assert_eq!(parse_time_input("01:30").unwrap(), 90);
        assert_eq!(parse_time_input("10:00").unwrap(), 600);
    }

    #[test]
    fn parse_hh_mm_ss() {
        assert_eq!(parse_time_input("00:01:30").unwrap(), 90);
    }

    #[test]
    fn parse_invalid() {
        assert!(parse_time_input("").is_err());
        assert!(parse_time_input("abc").is_err());
        assert!(parse_time_input("1:99").is_err());
        assert!(parse_time_input("1:2:99").is_err());
        assert!(parse_time_input("1:2:3:4").is_err());
    }

    #[test]
    fn audiobook_position_uses_output_when_available() {
        let pos =
            compute_audiobook_position_secs(Some(12.5), false, 10, Duration::from_secs(2), 2.0);
        assert!((pos - 12.5).abs() < f64::EPSILON);
    }

    #[test]
    fn audiobook_position_clamps_output_to_zero() {
        let pos =
            compute_audiobook_position_secs(Some(-3.0), false, 10, Duration::from_secs(2), 2.0);
        assert!((pos - 0.0).abs() < f64::EPSILON);
    }

    #[test]
    fn audiobook_position_respects_speed_when_playing() {
        let pos = compute_audiobook_position_secs(None, false, 10, Duration::from_secs(2), 1.5);
        assert!((pos - 13.0).abs() < f64::EPSILON);
    }

    #[test]
    fn audiobook_position_ignores_elapsed_when_paused() {
        let pos = compute_audiobook_position_secs(None, true, 42, Duration::from_secs(99), 2.0);
        assert!((pos - 42.0).abs() < f64::EPSILON);
    }
}
