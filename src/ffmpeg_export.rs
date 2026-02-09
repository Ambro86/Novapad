use crate::ffmpeg_dyn::*;
use crate::ffmpeg_source::{FfmpegApi, FfmpegSource, ffmpeg_api, list_audio_streams};
use crate::log_debug;
use crate::sapi4_engine;
use crate::sapi5_engine;
use crate::settings::AppSettings;
use crate::subtitle_wasapi::{decode_mp3_to_pcm, resample_pcm};
use crate::subtitles::{SubtitleCue, find_subtitle_for_media, load_subtitles};
use crate::tts_engine;
use rodio::Source;
use sha2::Digest;
use std::collections::VecDeque;
use std::ffi::CString;
use std::path::{Path, PathBuf};
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicI64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use windows::Win32::Foundation::HWND;

const MIX_SUFFIX: &str = "_tts_mix_";
const FILM_BASE_VOLUME: f32 = 0.7;
const DESC_VOLUME: f32 = 0.7;
const DUCK_MULTIPLIER: f32 = 0.35;
const DUCK_PAD_BEFORE_SEC: f32 = 0.0;
const DUCK_PAD_AFTER_SEC: f32 = 0.0;
const DUCK_END_TRIM_SEC: f32 = 0.90;
const DUCK_START_DELAY_SEC: f32 = 0.28;

static MIX_OUTPUTS: OnceLock<Mutex<Vec<PathBuf>>> = OnceLock::new();
const DEFAULT_AAC_BITRATE: i64 = 192_000;
const AVIO_FLAG_WRITE: i32 = 2;
const AV_CODEC_FLAG_QSCALE_FALLBACK: i32 = 1 << 1;
const FF_QP2LAMBDA_FALLBACK: i32 = 118;

#[derive(Clone)]
pub struct MixExportOptions {
    pub ducking: bool,
}

struct CueAudio {
    start_sample: u64,
    samples: Arc<[f32]>,
}

struct ActiveCue {
    samples: Arc<[f32]>,
    read_offset: usize,
}

fn tts_cache_dir(media_path: &Path, settings: &AppSettings) -> Result<PathBuf, String> {
    let base_dir = crate::settings::settings_dir().join("subtitle_cache");
    let mut hasher = sha2::Sha256::new();
    hasher.update(media_path.to_string_lossy().as_bytes());
    let engine_key = match settings.tts_engine {
        crate::settings::TtsEngine::Edge => "edge",
        crate::settings::TtsEngine::Sapi5 => "sapi5",
        crate::settings::TtsEngine::Sapi4 => "sapi4",
    };
    hasher.update(engine_key.as_bytes());
    hasher.update(settings.tts_voice.as_bytes());
    let hash = hex::encode(hasher.finalize());
    Ok(base_dir.join(&hash[..16]))
}

fn parse_sapi4_voice_index(voice: &str) -> i32 {
    if let Some(hash_pos) = voice.find('#') {
        let rest = &voice[hash_pos + 1..];
        if let Some(pipe_pos) = rest.find('|') {
            return rest[..pipe_pos].parse::<i32>().unwrap_or(0);
        }
        return rest.parse::<i32>().unwrap_or(0);
    }
    0
}

fn ensure_tts_cache(
    media_path: &Path,
    subtitle_path: &Path,
    settings: &AppSettings,
) -> Result<Vec<SubtitleCue>, String> {
    let cues = load_subtitles(subtitle_path)?;
    if cues.is_empty() {
        return Err("Subtitle: no cues to export".to_string());
    }
    let cache_dir = tts_cache_dir(media_path, settings)?;
    if !cache_dir.exists() {
        std::fs::create_dir_all(&cache_dir)
            .map_err(|e| format!("Subtitle: cache dir create failed: {}", e))?;
    }
    for (idx, cue) in cues.iter().enumerate() {
        let ext = match settings.tts_engine {
            crate::settings::TtsEngine::Edge => "mp3",
            _ => "wav",
        };
        let path = cache_dir.join(format!("cue_{:04}.{}", idx, ext));
        if path.exists() {
            continue;
        }
        let text = cue.text.trim();
        if text.is_empty() {
            continue;
        }
        let chunks = tts_engine::split_into_tts_chunks(text, false, &[], settings.tts_engine);
        let has_overrides = chunks.iter().any(|chunk| chunk.override_voice.is_some());

        match settings.tts_engine {
            crate::settings::TtsEngine::Edge => {
                if has_overrides {
                    let cancel = Arc::new(std::sync::atomic::AtomicBool::new(false));
                    let options = tts_engine::AudiobookCommonOptions {
                        voice: &settings.tts_voice,
                        output: &path,
                        progress_hwnd: HWND(0),
                        cancel,
                        language: settings.language,
                        audiobook_bitrate_kbps: settings.audiobook_m4b_bitrate,
                        rate: settings.tts_rate,
                        pitch: settings.tts_pitch,
                        volume: settings.tts_volume,
                        sapi4_threads: None,
                    };
                    let config = tts_engine::MixedAudiobookConfig {
                        main_engine: settings.tts_engine,
                    };
                    let mut progress = 0usize;
                    tts_engine::render_mixed_audiobook_part(
                        &chunks,
                        &mut progress,
                        &path,
                        &options,
                        &config,
                    )?;
                } else {
                    let rt = tokio::runtime::Runtime::new()
                        .map_err(|e| format!("Subtitle: failed to create runtime: {}", e))?;
                    let request_id = uuid::Uuid::new_v4().simple().to_string();
                    match rt.block_on(tts_engine::download_audio_chunk(
                        text,
                        &settings.tts_voice,
                        &request_id,
                        settings.tts_rate,
                        settings.tts_pitch,
                        settings.tts_volume,
                        settings.language,
                    )) {
                        Ok(bytes) => {
                            std::fs::write(&path, bytes).map_err(|e| {
                                format!("Subtitle: failed to write audio chunk: {}", e)
                            })?;
                        }
                        Err(err) => {
                            return Err(format!("Subtitle: download failed: {}", err));
                        }
                    }
                }
            }
            crate::settings::TtsEngine::Sapi5 => {
                if has_overrides {
                    let cancel = Arc::new(std::sync::atomic::AtomicBool::new(false));
                    let options = tts_engine::AudiobookCommonOptions {
                        voice: &settings.tts_voice,
                        output: &path,
                        progress_hwnd: HWND(0),
                        cancel,
                        language: settings.language,
                        audiobook_bitrate_kbps: settings.audiobook_m4b_bitrate,
                        rate: settings.tts_rate,
                        pitch: settings.tts_pitch,
                        volume: settings.tts_volume,
                        sapi4_threads: None,
                    };
                    let config = tts_engine::MixedAudiobookConfig {
                        main_engine: settings.tts_engine,
                    };
                    let mut progress = 0usize;
                    tts_engine::render_mixed_audiobook_part(
                        &chunks,
                        &mut progress,
                        &path,
                        &options,
                        &config,
                    )?;
                } else {
                    let cancel = Arc::new(std::sync::atomic::AtomicBool::new(false));
                    let options = sapi5_engine::SapiExportOptions {
                        chunks: &[text.to_string()],
                        voice_name: &settings.tts_voice,
                        output_path: &path,
                        language: settings.language,
                        rate: settings.tts_rate,
                        pitch: settings.tts_pitch,
                        volume: settings.tts_volume,
                        cancel,
                    };
                    sapi5_engine::speak_sapi_to_file(options, |_| {})?;
                }
            }
            crate::settings::TtsEngine::Sapi4 => {
                if has_overrides {
                    let cancel = Arc::new(std::sync::atomic::AtomicBool::new(false));
                    let options = tts_engine::AudiobookCommonOptions {
                        voice: &settings.tts_voice,
                        output: &path,
                        progress_hwnd: HWND(0),
                        cancel,
                        language: settings.language,
                        audiobook_bitrate_kbps: settings.audiobook_m4b_bitrate,
                        rate: settings.tts_rate,
                        pitch: settings.tts_pitch,
                        volume: settings.tts_volume,
                        sapi4_threads: None,
                    };
                    let config = tts_engine::MixedAudiobookConfig {
                        main_engine: settings.tts_engine,
                    };
                    let mut progress = 0usize;
                    tts_engine::render_mixed_audiobook_part(
                        &chunks,
                        &mut progress,
                        &path,
                        &options,
                        &config,
                    )?;
                } else {
                    let cancel = Arc::new(std::sync::atomic::AtomicBool::new(false));
                    let voice_idx = parse_sapi4_voice_index(&settings.tts_voice);
                    let options = sapi4_engine::Sapi4Options {
                        rate: settings.tts_rate,
                        pitch: settings.tts_pitch,
                        volume: settings.tts_volume,
                        cancel,
                    };
                    sapi4_engine::speak_sapi4_to_file(
                        &[text.to_string()],
                        voice_idx,
                        &path,
                        options,
                        |_| {},
                    )?;
                }
            }
        }
    }
    Ok(cues)
}

fn collect_tts_audio(
    cues: &[SubtitleCue],
    media_path: &Path,
    settings: &AppSettings,
    sample_rate: u32,
    channels: u16,
    offset_secs: f64,
) -> Result<VecDeque<CueAudio>, String> {
    let cache_dir = tts_cache_dir(media_path, settings)?;
    let mut pending = VecDeque::new();
    for (idx, cue) in cues.iter().enumerate() {
        let mp3_path = cache_dir.join(format!("cue_{:04}.mp3", idx));
        let wav_path = cache_dir.join(format!("cue_{:04}.wav", idx));
        let path = if mp3_path.exists() {
            mp3_path
        } else if wav_path.exists() {
            wav_path
        } else {
            continue;
        };
        let bytes =
            std::fs::read(&path).map_err(|e| format!("Subtitle: failed to read cache: {}", e))?;
        let (samples, src_rate, src_ch) = decode_mp3_to_pcm(&bytes)?;
        let resampled = resample_pcm(&samples, src_rate, src_ch, sample_rate, channels);
        let start_secs = (cue.start.as_secs_f64() + offset_secs).max(0.0);
        let start_sample = (start_secs * sample_rate as f64 * channels as f64) as u64;
        pending.push_back(CueAudio {
            start_sample,
            samples: Arc::from(resampled),
        });
    }
    Ok(pending)
}

fn build_duck_intervals(
    pending: &VecDeque<CueAudio>,
    sample_rate: u32,
    channels: u16,
) -> Vec<(u64, u64)> {
    if pending.is_empty() {
        return Vec::new();
    }
    let samples_per_sec = sample_rate as f32 * channels as f32;
    let pad_before = (DUCK_PAD_BEFORE_SEC * samples_per_sec) as i64;
    let pad_after = (DUCK_PAD_AFTER_SEC * samples_per_sec) as i64;
    let start_delay = (DUCK_START_DELAY_SEC * samples_per_sec) as i64;
    let end_trim = (DUCK_END_TRIM_SEC * samples_per_sec) as i64;

    let mut intervals = Vec::new();
    for cue in pending.iter() {
        let start = cue.start_sample as i64;
        let mut duck_start = start - pad_before + start_delay;
        let duck_end = start + cue.samples.len() as i64 + pad_after - end_trim;
        if duck_end <= duck_start {
            continue;
        }
        if duck_start < 0 {
            duck_start = 0;
        }
        if duck_end < 0 {
            continue;
        }
        intervals.push((duck_start as u64, duck_end as u64));
    }
    if intervals.is_empty() {
        return intervals;
    }
    intervals.sort_by_key(|(start, _)| *start);
    let mut merged: Vec<(u64, u64)> = Vec::with_capacity(intervals.len());
    for (start, end) in intervals {
        if let Some(last) = merged.last_mut()
            && start <= last.1
        {
            last.1 = last.1.max(end);
            continue;
        }
        merged.push((start, end));
    }
    merged
}

fn register_mix_output(path: &Path) {
    let list = MIX_OUTPUTS.get_or_init(|| Mutex::new(Vec::new()));
    if let Ok(mut outputs) = list.lock()
        && !outputs.iter().any(|p| p == path)
    {
        outputs.push(path.to_path_buf());
    }
}

fn dict_set_str(
    api: &FfmpegApi,
    dict: &mut *mut AVDictionary,
    key: &str,
    value: &str,
) -> Result<(), String> {
    let key_c = CString::new(key).map_err(|_| "FFmpeg: invalid dict key".to_string())?;
    let value_c = CString::new(value).map_err(|_| "FFmpeg: invalid dict value".to_string())?;
    let ret = unsafe { (api.av_dict_set)(dict, key_c.as_ptr(), value_c.as_ptr(), 0) };
    if ret < 0 {
        return Err(format!("FFmpeg: failed to set {} (code {})", key, ret));
    }
    Ok(())
}

fn segment_format_from_path(path: &Path) -> Option<&'static str> {
    let ext = path.extension()?.to_str()?.to_ascii_lowercase();
    match ext.as_str() {
        "mp3" => Some("mp3"),
        "m4a" | "m4b" | "mp4" => Some("ipod"),
        "aac" => Some("adts"),
        "ogg" => Some("ogg"),
        "opus" => Some("opus"),
        "flac" => Some("flac"),
        "wav" => Some("wav"),
        _ => None,
    }
}

pub fn segment_audio_file(
    input_path: &Path,
    output_pattern: &Path,
    segment_seconds: u32,
    start_number: u32,
) -> Result<(), String> {
    let api = ffmpeg_api()?;
    let input_c = CString::new(input_path.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid input path".to_string())?;
    let output_c = CString::new(output_pattern.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid output pattern".to_string())?;
    let format_c =
        CString::new("segment").map_err(|_| "FFmpeg: invalid segment format".to_string())?;

    let mut in_ctx: *mut AVFormatContext = ptr::null_mut();
    let mut out_ctx: *mut AVFormatContext = ptr::null_mut();

    let open_ret = unsafe {
        (api.avformat_open_input)(
            &mut in_ctx,
            input_c.as_ptr(),
            ptr::null_mut(),
            ptr::null_mut(),
        )
    };
    if open_ret < 0 || in_ctx.is_null() {
        return Err("FFmpeg: failed to open input".to_string());
    }
    if unsafe { (api.avformat_find_stream_info)(in_ctx, ptr::null_mut()) } < 0 {
        unsafe { (api.avformat_close_input)(&mut in_ctx) };
        return Err("FFmpeg: input stream info failed".to_string());
    }

    let audio_stream_idx = unsafe {
        (api.av_find_best_stream)(
            in_ctx,
            AVMediaType_AVMEDIA_TYPE_AUDIO,
            -1,
            -1,
            ptr::null_mut(),
            0,
        )
    };
    if audio_stream_idx < 0 {
        unsafe { (api.avformat_close_input)(&mut in_ctx) };
        return Err("FFmpeg: audio stream not found".to_string());
    }

    let alloc_ret = unsafe {
        (api.avformat_alloc_output_context2)(
            &mut out_ctx,
            ptr::null_mut(),
            format_c.as_ptr(),
            output_c.as_ptr(),
        )
    };
    if alloc_ret < 0 || out_ctx.is_null() {
        unsafe { (api.avformat_close_input)(&mut in_ctx) };
        return Err("FFmpeg: failed to allocate segment output".to_string());
    }

    let in_stream = unsafe { *(*in_ctx).streams.add(audio_stream_idx as usize) };
    let out_stream = unsafe { (api.avformat_new_stream)(out_ctx, ptr::null()) };
    if out_stream.is_null() {
        unsafe {
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_ctx);
        }
        return Err("FFmpeg: failed to create output stream".to_string());
    }
    unsafe {
        (api.avcodec_parameters_copy)((*out_stream).codecpar, (*in_stream).codecpar);
        (*out_stream).time_base = (*in_stream).time_base;
    }

    let mut dict: *mut AVDictionary = ptr::null_mut();
    if segment_seconds == 0 {
        unsafe {
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_ctx);
        }
        return Err("FFmpeg: segment time must be > 0".to_string());
    }
    dict_set_str(api, &mut dict, "segment_time", &segment_seconds.to_string())?;
    let start_number = start_number.max(1);
    dict_set_str(
        api,
        &mut dict,
        "segment_start_number",
        &start_number.to_string(),
    )?;
    dict_set_str(api, &mut dict, "reset_timestamps", "1")?;
    if let Some(fmt) = segment_format_from_path(output_pattern) {
        dict_set_str(api, &mut dict, "segment_format", fmt)?;
    }

    let header_ret = unsafe { (api.avformat_write_header)(out_ctx, &mut dict) };
    unsafe {
        (api.av_dict_free)(&mut dict);
    }
    if header_ret < 0 {
        unsafe {
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_ctx);
        }
        return Err("FFmpeg: failed to write segment header".to_string());
    }

    let mut pkt = unsafe { (api.av_packet_alloc)() };
    if pkt.is_null() {
        unsafe {
            (api.av_write_trailer)(out_ctx);
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_ctx);
        }
        return Err("FFmpeg: packet alloc failed".to_string());
    }

    loop {
        let read_ret = unsafe { (api.av_read_frame)(in_ctx, pkt) };
        if read_ret < 0 {
            break;
        }
        if unsafe { (*pkt).stream_index } != audio_stream_idx {
            unsafe { (api.av_packet_unref)(pkt) };
            continue;
        }
        unsafe {
            (*pkt).stream_index = (*out_stream).index;
            (api.av_packet_rescale_ts)(pkt, (*in_stream).time_base, (*out_stream).time_base);
            let write_ret = (api.av_interleaved_write_frame)(out_ctx, pkt);
            if write_ret < 0 {
                log_debug(&format!("FFmpeg: segment write failed: {}", write_ret));
            }
            (api.av_packet_unref)(pkt);
        }
    }

    unsafe {
        let trailer_ret = (api.av_write_trailer)(out_ctx);
        if trailer_ret < 0 {
            log_debug(&format!("FFmpeg: av_write_trailer failed: {}", trailer_ret));
        }
        (api.av_packet_free)(&mut pkt);
        (api.avformat_free_context)(out_ctx);
        (api.avformat_close_input)(&mut in_ctx);
    }
    Ok(())
}

pub fn cleanup_tts_artifacts() -> Result<(), String> {
    let mut errors = Vec::new();
    let outputs = MIX_OUTPUTS.get_or_init(|| Mutex::new(Vec::new()));
    if let Ok(mut list) = outputs.lock() {
        for path in list.drain(..) {
            if path.exists()
                && let Err(e) = std::fs::remove_file(&path)
            {
                errors.push(format!(
                    "Subtitle: failed to delete mixed output {}: {}",
                    path.display(),
                    e
                ));
            }
        }
    }

    let cache_dir = crate::settings::settings_dir().join("subtitle_cache");
    if cache_dir.exists()
        && let Err(e) = std::fs::remove_dir_all(&cache_dir)
    {
        errors.push(format!(
            "Subtitle: failed to delete cache {}: {}",
            cache_dir.display(),
            e
        ));
    }

    if errors.is_empty() {
        Ok(())
    } else {
        Err(errors.join(" | "))
    }
}

fn pick_encoder_sample_fmt(codec: *const AVCodec) -> AVSampleFormat {
    unsafe {
        if codec.is_null() {
            return AVSampleFormat_AV_SAMPLE_FMT_FLTP;
        }
        let fmts = (*codec).sample_fmts;
        if fmts.is_null() {
            return AVSampleFormat_AV_SAMPLE_FMT_FLTP;
        }
        let mut idx = 0;
        loop {
            let fmt = *fmts.add(idx);
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_NONE {
                break;
            }
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_FLTP {
                return fmt;
            }
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_FLT {
                return fmt;
            }
            idx += 1;
        }
    }
    AVSampleFormat_AV_SAMPLE_FMT_FLTP
}

fn pick_encoder_sample_fmt_with_preference(
    codec: *const AVCodec,
    preferred: Option<AVSampleFormat>,
) -> AVSampleFormat {
    unsafe {
        if codec.is_null() {
            return preferred.unwrap_or(AVSampleFormat_AV_SAMPLE_FMT_FLTP);
        }
        let fmts = (*codec).sample_fmts;
        if fmts.is_null() {
            return preferred.unwrap_or(AVSampleFormat_AV_SAMPLE_FMT_FLTP);
        }
        let mut idx = 0;
        let mut first = None;
        loop {
            let fmt = *fmts.add(idx);
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_NONE {
                break;
            }
            if first.is_none() {
                first = Some(fmt);
            }
            if let Some(pref) = preferred
                && fmt == pref
            {
                return fmt;
            }
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_FLTP {
                return fmt;
            }
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_FLT {
                return fmt;
            }
            if fmt == AVSampleFormat_AV_SAMPLE_FMT_S16 {
                return fmt;
            }
            idx += 1;
        }
        first.unwrap_or(AVSampleFormat_AV_SAMPLE_FMT_FLTP)
    }
}

fn rescale_q(value: i64, src: AVRational, dst: AVRational) -> i64 {
    if src.den == 0 || dst.num == 0 {
        return 0;
    }
    let num = value as i128 * src.num as i128 * dst.den as i128;
    let den = src.den as i128 * dst.num as i128;
    if den == 0 { 0 } else { (num / den) as i64 }
}

fn read_next_packet_for_stream(
    api: &FfmpegApi,
    ctx: *mut AVFormatContext,
    pkt: *mut AVPacket,
    stream_idx: i32,
) -> bool {
    loop {
        let ret = unsafe { (api.av_read_frame)(ctx, pkt) };
        if ret < 0 {
            return false;
        }
        let idx = unsafe { (*pkt).stream_index };
        if idx == stream_idx {
            return true;
        }
        unsafe {
            (api.av_packet_unref)(pkt);
        }
    }
}

fn encode_mixed_audio_to_m4a(
    api: &FfmpegApi,
    input_path: &Path,
    subtitle_path: &Path,
    settings: &AppSettings,
    options: &MixExportOptions,
    out_audio_path: &Path,
) -> Result<(), String> {
    let cues = ensure_tts_cache(input_path, subtitle_path, settings)?;

    let mut source = FfmpegSource::try_new(input_path, 0, None, None)?;
    let sample_rate = source.sample_rate();
    let channels = source.channels();

    let offset_secs = settings.subtitle_offset_ms as f64 / 1000.0;
    let mut pending = collect_tts_audio(
        &cues,
        input_path,
        settings,
        sample_rate,
        channels,
        offset_secs,
    )?;
    let duck_intervals = if options.ducking {
        build_duck_intervals(&pending, sample_rate, channels)
    } else {
        Vec::new()
    };
    let ducking_enabled = options.ducking && !duck_intervals.is_empty();
    let mut duck_idx = 0usize;
    let mut duck_end = 0u64;
    let mut duck_active = false;
    let mut duck_next_change = u64::MAX;
    if let Some((start, end)) = duck_intervals.first() {
        duck_end = *end;
        duck_next_change = *start;
    }
    let mut active: Vec<ActiveCue> = Vec::new();

    let out_path_c = CString::new(out_audio_path.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid output path".to_string())?;
    let mut out_ctx: *mut AVFormatContext = ptr::null_mut();
    let alloc_ret = unsafe {
        (api.avformat_alloc_output_context2)(
            &mut out_ctx,
            ptr::null_mut(),
            ptr::null(),
            out_path_c.as_ptr(),
        )
    };
    if alloc_ret < 0 || out_ctx.is_null() {
        return Err("FFmpeg: failed to allocate output context".to_string());
    }

    let codec = unsafe { (api.avcodec_find_encoder)(AVCodecID_AV_CODEC_ID_AAC) };
    if codec.is_null() {
        unsafe { (api.avformat_free_context)(out_ctx) };
        return Err("FFmpeg: AAC encoder not found".to_string());
    }
    let stream = unsafe { (api.avformat_new_stream)(out_ctx, codec) };
    if stream.is_null() {
        unsafe { (api.avformat_free_context)(out_ctx) };
        return Err("FFmpeg: failed to create output stream".to_string());
    }
    let mut codec_ctx = unsafe { (api.avcodec_alloc_context3)(codec) };
    if codec_ctx.is_null() {
        unsafe { (api.avformat_free_context)(out_ctx) };
        return Err("FFmpeg: failed to allocate encoder context".to_string());
    }
    let enc_fmt = pick_encoder_sample_fmt(codec);
    unsafe {
        (*codec_ctx).sample_rate = sample_rate as i32;
        (*codec_ctx).bit_rate = DEFAULT_AAC_BITRATE;
        (*codec_ctx).sample_fmt = enc_fmt;
        (*codec_ctx).time_base = AVRational {
            num: 1,
            den: sample_rate as i32,
        };
        let mut out_layout: AVChannelLayout = std::mem::zeroed();
        (api.av_channel_layout_default)(&mut out_layout, channels as i32);
        (*codec_ctx).ch_layout = out_layout;
    }
    let open_ret = unsafe { (api.avcodec_open2)(codec_ctx, codec, ptr::null_mut()) };
    if open_ret < 0 {
        unsafe {
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to open encoder".to_string());
    }
    unsafe {
        (*stream).time_base = (*codec_ctx).time_base;
        let par_ret = (api.avcodec_parameters_from_context)((*stream).codecpar, codec_ctx);
        if par_ret < 0 {
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            return Err("FFmpeg: failed to copy encoder parameters".to_string());
        }
    }

    let mut io: *mut AVIOContext = ptr::null_mut();
    let open_io = unsafe { (api.avio_open)(&mut io, out_path_c.as_ptr(), AVIO_FLAG_WRITE) };
    if open_io < 0 {
        unsafe {
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to open output file".to_string());
    }
    unsafe {
        (*out_ctx).pb = io;
    }
    let header_ret = unsafe { (api.avformat_write_header)(out_ctx, ptr::null_mut()) };
    if header_ret < 0 {
        unsafe {
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to write output header".to_string());
    }

    let frame_size = unsafe { (*codec_ctx).frame_size }.max(1024) as usize;
    let mut frame = unsafe { (api.av_frame_alloc)() };
    if frame.is_null() {
        unsafe {
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to allocate frame".to_string());
    }
    unsafe {
        (*frame).nb_samples = frame_size as i32;
        (*frame).format = enc_fmt;
        (*frame).sample_rate = sample_rate as i32;
        let mut frame_layout: AVChannelLayout = std::mem::zeroed();
        (api.av_channel_layout_default)(&mut frame_layout, channels as i32);
        (*frame).ch_layout = frame_layout;
        if (api.av_frame_get_buffer)(frame, 0) < 0 {
            (api.av_frame_free)(&mut frame);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            return Err("FFmpeg: failed to allocate frame buffer".to_string());
        }
    }

    let mut in_layout: AVChannelLayout = unsafe { std::mem::zeroed() };
    let mut out_layout: AVChannelLayout = unsafe { std::mem::zeroed() };
    unsafe {
        (api.av_channel_layout_default)(&mut in_layout, channels as i32);
        (api.av_channel_layout_default)(&mut out_layout, channels as i32);
    }
    let mut swr_ctx: *mut SwrContext = ptr::null_mut();
    let swr_ret = unsafe {
        (api.swr_alloc_set_opts2)(
            &mut swr_ctx,
            &out_layout,
            enc_fmt,
            sample_rate as i32,
            &in_layout,
            AVSampleFormat_AV_SAMPLE_FMT_FLT,
            sample_rate as i32,
            0,
            ptr::null_mut(),
        )
    };
    if swr_ret < 0 || swr_ctx.is_null() {
        unsafe {
            (api.av_frame_free)(&mut frame);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to init resampler".to_string());
    }
    if unsafe { (api.swr_init)(swr_ctx) } < 0 {
        unsafe {
            (api.swr_free)(&mut swr_ctx);
            (api.av_frame_free)(&mut frame);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to init resampler".to_string());
    }

    let mut mixed_buf: VecDeque<f32> = VecDeque::new();
    let mut sample_index: u64 = 0;
    let mut next_pts: i64 = 0;
    let channel_count = channels as usize;
    let base_film_gain = if options.ducking {
        FILM_BASE_VOLUME
    } else {
        1.0
    };
    let ducked_film_gain = FILM_BASE_VOLUME * DUCK_MULTIPLIER;
    let desc_gain = if options.ducking { DESC_VOLUME } else { 1.0 };

    let read_next_sample = |pending: &mut VecDeque<CueAudio>,
                            active: &mut Vec<ActiveCue>,
                            current_sample: u64|
     -> f32 {
        while let Some(front) = pending.front() {
            if current_sample >= front.start_sample {
                if let Some(cue) = pending.pop_front() {
                    active.push(ActiveCue {
                        samples: cue.samples,
                        read_offset: 0,
                    });
                }
            } else {
                break;
            }
        }
        let mut tts_sample = 0.0f32;
        active.retain_mut(|cue| {
            if cue.read_offset < cue.samples.len() {
                tts_sample += cue.samples[cue.read_offset];
                cue.read_offset += 1;
                true
            } else {
                false
            }
        });
        tts_sample
    };

    loop {
        while mixed_buf.len() < frame_size * channel_count {
            match source.next() {
                Some(orig) => {
                    let tts = read_next_sample(&mut pending, &mut active, sample_index);
                    let mut film_gain = base_film_gain;
                    if ducking_enabled && sample_index >= duck_next_change {
                        if duck_active {
                            duck_active = false;
                            duck_idx = duck_idx.saturating_add(1);
                            if let Some((start, end)) = duck_intervals.get(duck_idx) {
                                duck_end = *end;
                                duck_next_change = *start;
                            } else {
                                duck_next_change = u64::MAX;
                            }
                        } else {
                            duck_active = true;
                            duck_next_change = duck_end;
                        }
                    }
                    if duck_active {
                        film_gain = ducked_film_gain;
                    }
                    let mixed = orig * film_gain + tts * desc_gain;
                    mixed_buf.push_back(mixed.clamp(-1.0, 1.0));
                    sample_index = sample_index.saturating_add(1);
                }
                None => break,
            }
        }
        if mixed_buf.is_empty() {
            break;
        }
        let mut input_frame = Vec::with_capacity(frame_size * channel_count);
        while input_frame.len() < frame_size * channel_count {
            if let Some(v) = mixed_buf.pop_front() {
                input_frame.push(v);
            } else {
                input_frame.push(0.0);
            }
        }
        unsafe {
            if (api.av_frame_make_writable)(frame) < 0 {
                return Err("FFmpeg: frame not writable".to_string());
            }
        }
        let mut in_ptr = input_frame.as_ptr() as *const u8;
        let in_ptrs = &mut in_ptr as *mut *const u8;
        let out_count = unsafe { (*frame).nb_samples };
        let conv = unsafe {
            (api.swr_convert)(
                swr_ctx,
                (*frame).data.as_mut_ptr(),
                out_count,
                in_ptrs,
                out_count,
            )
        };
        if conv < 0 {
            return Err("FFmpeg: resample failed".to_string());
        }
        unsafe {
            (*frame).pts = next_pts;
        }
        next_pts += out_count as i64;
        let send_ret = unsafe { (api.avcodec_send_frame)(codec_ctx, frame) };
        if send_ret < 0 {
            return Err("FFmpeg: send_frame failed".to_string());
        }
        loop {
            let mut pkt = unsafe { (api.av_packet_alloc)() };
            if pkt.is_null() {
                return Err("FFmpeg: packet alloc failed".to_string());
            }
            let recv = unsafe { (api.avcodec_receive_packet)(codec_ctx, pkt) };
            if recv == 0 {
                unsafe {
                    (*pkt).stream_index = (*stream).index;
                    (api.av_packet_rescale_ts)(pkt, (*codec_ctx).time_base, (*stream).time_base);
                    let wret = (api.av_interleaved_write_frame)(out_ctx, pkt);
                    (api.av_packet_unref)(pkt);
                    (api.av_packet_free)(&mut pkt);
                    if wret < 0 {
                        return Err("FFmpeg: write audio packet failed".to_string());
                    }
                }
            } else {
                unsafe {
                    (api.av_packet_free)(&mut pkt);
                }
                break;
            }
        }
    }

    unsafe {
        (api.avcodec_send_frame)(codec_ctx, ptr::null());
        loop {
            let mut pkt = (api.av_packet_alloc)();
            if pkt.is_null() {
                break;
            }
            let recv = (api.avcodec_receive_packet)(codec_ctx, pkt);
            if recv == 0 {
                (*pkt).stream_index = (*stream).index;
                (api.av_packet_rescale_ts)(pkt, (*codec_ctx).time_base, (*stream).time_base);
                (api.av_interleaved_write_frame)(out_ctx, pkt);
                (api.av_packet_unref)(pkt);
                (api.av_packet_free)(&mut pkt);
            } else {
                (api.av_packet_free)(&mut pkt);
                break;
            }
        }
        let trailer_ret = (api.av_write_trailer)(out_ctx);
        if trailer_ret < 0 {
            log_debug(&format!("FFmpeg: av_write_trailer failed: {}", trailer_ret));
        }
        (api.swr_free)(&mut swr_ctx);
        (api.av_frame_free)(&mut frame);
        (api.avio_closep)(&mut io);
        (api.avcodec_free_context)(&mut codec_ctx);
        (api.avformat_free_context)(out_ctx);
        (api.av_channel_layout_uninit)(&mut in_layout);
        (api.av_channel_layout_uninit)(&mut out_layout);
    }

    Ok(())
}

fn mux_video_with_audio(
    api: &FfmpegApi,
    video_path: &Path,
    audio_path: &Path,
    out_path: &Path,
) -> Result<(), String> {
    let video_c = CString::new(video_path.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid video path".to_string())?;
    let audio_c = CString::new(audio_path.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid audio path".to_string())?;
    let out_c = CString::new(out_path.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid output path".to_string())?;

    let mut in_video: *mut AVFormatContext = ptr::null_mut();
    let mut in_audio: *mut AVFormatContext = ptr::null_mut();
    let mut out_ctx: *mut AVFormatContext = ptr::null_mut();

    let open_vid = unsafe {
        (api.avformat_open_input)(
            &mut in_video,
            video_c.as_ptr(),
            ptr::null_mut(),
            ptr::null_mut(),
        )
    };
    if open_vid < 0 || in_video.is_null() {
        return Err("FFmpeg: failed to open video input".to_string());
    }
    let open_aud = unsafe {
        (api.avformat_open_input)(
            &mut in_audio,
            audio_c.as_ptr(),
            ptr::null_mut(),
            ptr::null_mut(),
        )
    };
    if open_aud < 0 || in_audio.is_null() {
        unsafe { (api.avformat_close_input)(&mut in_video) };
        return Err("FFmpeg: failed to open audio input".to_string());
    }
    if unsafe { (api.avformat_find_stream_info)(in_video, ptr::null_mut()) } < 0 {
        unsafe {
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: video stream info failed".to_string());
    }
    if unsafe { (api.avformat_find_stream_info)(in_audio, ptr::null_mut()) } < 0 {
        unsafe {
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: audio stream info failed".to_string());
    }

    let video_stream_idx = unsafe {
        (api.av_find_best_stream)(
            in_video,
            AVMediaType_AVMEDIA_TYPE_VIDEO,
            -1,
            -1,
            ptr::null_mut(),
            0,
        )
    };
    if video_stream_idx < 0 {
        unsafe {
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: video stream not found".to_string());
    }
    let audio_stream_idx = unsafe {
        (api.av_find_best_stream)(
            in_audio,
            AVMediaType_AVMEDIA_TYPE_AUDIO,
            -1,
            -1,
            ptr::null_mut(),
            0,
        )
    };
    if audio_stream_idx < 0 {
        unsafe {
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: audio stream not found".to_string());
    }

    let alloc_ret = unsafe {
        (api.avformat_alloc_output_context2)(
            &mut out_ctx,
            ptr::null_mut(),
            ptr::null(),
            out_c.as_ptr(),
        )
    };
    if alloc_ret < 0 || out_ctx.is_null() {
        unsafe {
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: failed to alloc output context".to_string());
    }

    let in_v_stream = unsafe { *(*in_video).streams.add(video_stream_idx as usize) };
    let in_a_stream = unsafe { *(*in_audio).streams.add(audio_stream_idx as usize) };
    let out_v_stream = unsafe { (api.avformat_new_stream)(out_ctx, ptr::null()) };
    let out_a_stream = unsafe { (api.avformat_new_stream)(out_ctx, ptr::null()) };
    if out_v_stream.is_null() || out_a_stream.is_null() {
        unsafe {
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: failed to create output streams".to_string());
    }
    unsafe {
        (api.avcodec_parameters_copy)((*out_v_stream).codecpar, (*in_v_stream).codecpar);
        (api.avcodec_parameters_copy)((*out_a_stream).codecpar, (*in_a_stream).codecpar);
        (*out_v_stream).time_base = (*in_v_stream).time_base;
        (*out_a_stream).time_base = (*in_a_stream).time_base;
    }

    let mut io: *mut AVIOContext = ptr::null_mut();
    let open_io = unsafe { (api.avio_open)(&mut io, out_c.as_ptr(), AVIO_FLAG_WRITE) };
    if open_io < 0 {
        unsafe {
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: failed to open output file".to_string());
    }
    unsafe {
        (*out_ctx).pb = io;
    }
    let header_ret = unsafe { (api.avformat_write_header)(out_ctx, ptr::null_mut()) };
    if header_ret < 0 {
        unsafe {
            (api.avio_closep)(&mut io);
            (api.avformat_free_context)(out_ctx);
            (api.avformat_close_input)(&mut in_audio);
            (api.avformat_close_input)(&mut in_video);
        }
        return Err("FFmpeg: failed to write header".to_string());
    }

    let mut pkt_v = unsafe { (api.av_packet_alloc)() };
    let mut pkt_a = unsafe { (api.av_packet_alloc)() };
    let mut has_v = read_next_packet_for_stream(api, in_video, pkt_v, video_stream_idx);
    let mut has_a = read_next_packet_for_stream(api, in_audio, pkt_a, audio_stream_idx);

    while has_v || has_a {
        let write_video = if !has_a {
            true
        } else if !has_v {
            false
        } else {
            let vts = unsafe { (*pkt_v).pts };
            let ats = unsafe { (*pkt_a).pts };
            let v_cmp = rescale_q(vts, unsafe { (*in_v_stream).time_base }, unsafe {
                (*out_v_stream).time_base
            });
            let a_cmp = rescale_q(ats, unsafe { (*in_a_stream).time_base }, unsafe {
                (*out_a_stream).time_base
            });
            v_cmp <= a_cmp
        };

        if write_video {
            unsafe {
                (*pkt_v).stream_index = (*out_v_stream).index;
                (api.av_packet_rescale_ts)(
                    pkt_v,
                    (*in_v_stream).time_base,
                    (*out_v_stream).time_base,
                );
                let write_ret = (api.av_interleaved_write_frame)(out_ctx, pkt_v);
                if write_ret < 0 {
                    log_debug(&format!(
                        "FFmpeg: av_interleaved_write_frame (V) failed: {}",
                        write_ret
                    ));
                }
                (api.av_packet_unref)(pkt_v);
            }
            has_v = read_next_packet_for_stream(api, in_video, pkt_v, video_stream_idx);
        } else {
            unsafe {
                (*pkt_a).stream_index = (*out_a_stream).index;
                (api.av_packet_rescale_ts)(
                    pkt_a,
                    (*in_a_stream).time_base,
                    (*out_a_stream).time_base,
                );
                let write_ret = (api.av_interleaved_write_frame)(out_ctx, pkt_a);
                if write_ret < 0 {
                    log_debug(&format!(
                        "FFmpeg: av_interleaved_write_frame (A) failed: {}",
                        write_ret
                    ));
                }
                (api.av_packet_unref)(pkt_a);
            }
            has_a = read_next_packet_for_stream(api, in_audio, pkt_a, audio_stream_idx);
        }
    }

    unsafe {
        let trailer_ret = (api.av_write_trailer)(out_ctx);
        if trailer_ret < 0 {
            log_debug(&format!("FFmpeg: av_write_trailer failed: {}", trailer_ret));
        }
        if !pkt_v.is_null() {
            (api.av_packet_free)(&mut pkt_v);
        }
        if !pkt_a.is_null() {
            (api.av_packet_free)(&mut pkt_a);
        }
        (api.avio_closep)(&mut io);
        (api.avformat_free_context)(out_ctx);
        (api.avformat_close_input)(&mut in_audio);
        (api.avformat_close_input)(&mut in_video);
    }
    Ok(())
}

pub fn export_mixed_media(
    media_path: &Path,
    settings: &AppSettings,
    options: &MixExportOptions,
) -> Result<PathBuf, String> {
    let subtitle_path = find_subtitle_for_media(media_path)
        .ok_or_else(|| "Subtitle: not found for media".to_string())?;
    let api = ffmpeg_api()?;

    let mut hasher = sha2::Sha256::new();
    hasher.update(settings.tts_voice.as_bytes());
    hasher.update(settings.tts_rate.to_string().as_bytes());
    hasher.update(settings.tts_pitch.to_string().as_bytes());
    hasher.update(settings.tts_volume.to_string().as_bytes());
    hasher.update(settings.subtitle_offset_ms.to_string().as_bytes());
    hasher.update(if options.ducking { b"duck1" } else { b"duck0" });
    let mix_hash = hex::encode(hasher.finalize());

    let stem = media_path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("media");

    // Output directory is the same as the source media file
    let output_dir = media_path
        .parent()
        .ok_or_else(|| "Subtitle: media path has no parent directory".to_string())?;

    // Cache directory for temporary files
    let cache_dir = crate::settings::settings_dir()
        .join("subtitle_cache")
        .join("mixed");
    if let Err(e) = std::fs::create_dir_all(&cache_dir) {
        return Err(format!("Subtitle: failed to create cache dir: {}", e));
    }

    // Final output goes next to the original media file
    let out_mp4 = output_dir.join(format!("{stem}{MIX_SUFFIX}{}.mp4", &mix_hash[..8]));
    // Temporary audio file in cache
    let out_audio = cache_dir.join(format!("{stem}{MIX_SUFFIX}{}.m4a", &mix_hash[..8]));

    log_debug(&format!(
        "Subtitle: output path will be {}",
        out_mp4.display()
    ));

    if out_mp4.exists() {
        log_debug(&format!(
            "Subtitle: mixed file already exists at {}",
            out_mp4.display()
        ));
        return Ok(out_mp4);
    }

    log_debug(&format!(
        "Subtitle: exporting mixed audio for {}",
        media_path.display()
    ));
    encode_mixed_audio_to_m4a(
        api,
        media_path,
        &subtitle_path,
        settings,
        options,
        &out_audio,
    )?;
    log_debug(&format!(
        "Subtitle: muxing video+audio to {}",
        out_mp4.display()
    ));
    mux_video_with_audio(api, media_path, &out_audio, &out_mp4)?;
    if let Err(e) = std::fs::remove_file(&out_audio) {
        crate::log_debug(&format!(
            "Subtitle: failed to delete temp audio {}: {}",
            out_audio.display(),
            e
        ));
    }
    register_mix_output(&out_mp4);
    log_debug(&format!(
        "Subtitle: mixed file created successfully at {}",
        out_mp4.display()
    ));
    Ok(out_mp4)
}

pub fn is_mixed_output(path: &Path) -> bool {
    path.file_stem()
        .and_then(|s| s.to_str())
        .map(|s| s.contains(MIX_SUFFIX))
        .unwrap_or(false)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConvertAudioFormat {
    Mp3,
    Aac,
    Opus,
    Ogg,
    Flac,
    Wav,
    Aiff,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConvertAudioQuality {
    BitrateKbps(u32),
    OggQuality(u8),
    FlacCompression(u8),
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConvertAudioSettings {
    pub format: ConvertAudioFormat,
    pub quality: ConvertAudioQuality,
}

pub fn build_ffmpeg_args(settings: &ConvertAudioSettings) -> Vec<String> {
    match settings.format {
        ConvertAudioFormat::Mp3 => match settings.quality {
            ConvertAudioQuality::BitrateKbps(bitrate) => vec![
                "-c:a".to_string(),
                "libmp3lame".to_string(),
                "-b:a".to_string(),
                format!("{bitrate}k"),
            ],
            _ => vec!["-c:a".to_string(), "libmp3lame".to_string()],
        },
        ConvertAudioFormat::Aac => match settings.quality {
            ConvertAudioQuality::BitrateKbps(bitrate) => vec![
                "-c:a".to_string(),
                "aac".to_string(),
                "-b:a".to_string(),
                format!("{bitrate}k"),
            ],
            _ => vec!["-c:a".to_string(), "aac".to_string()],
        },
        ConvertAudioFormat::Opus => match settings.quality {
            ConvertAudioQuality::BitrateKbps(bitrate) => vec![
                "-c:a".to_string(),
                "libopus".to_string(),
                "-b:a".to_string(),
                format!("{bitrate}k"),
            ],
            _ => vec!["-c:a".to_string(), "libopus".to_string()],
        },
        ConvertAudioFormat::Ogg => match settings.quality {
            ConvertAudioQuality::OggQuality(q) => vec![
                "-c:a".to_string(),
                "libvorbis".to_string(),
                "-q:a".to_string(),
                q.to_string(),
            ],
            _ => vec!["-c:a".to_string(), "libvorbis".to_string()],
        },
        ConvertAudioFormat::Flac => match settings.quality {
            ConvertAudioQuality::FlacCompression(level) => vec![
                "-c:a".to_string(),
                "flac".to_string(),
                "-compression_level".to_string(),
                level.to_string(),
            ],
            _ => vec!["-c:a".to_string(), "flac".to_string()],
        },
        ConvertAudioFormat::Wav | ConvertAudioFormat::Aiff => {
            vec!["-c:a".to_string(), "pcm_s16le".to_string()]
        }
    }
}

pub fn validate_mp3_bitrate(value: i32) -> Result<u32, String> {
    if (96..=320).contains(&value) {
        Ok(value as u32)
    } else {
        Err("MP3 bitrate must be between 96 and 320 kbps".to_string())
    }
}

pub fn convert_audio_file(
    input_path: &Path,
    output_path: &Path,
    settings: &ConvertAudioSettings,
    cancel: Option<Arc<AtomicBool>>,
    mut progress: Option<&mut dyn FnMut(u32)>,
) -> Result<(), String> {
    let api = ffmpeg_api()?;
    let args = build_ffmpeg_args(settings);
    log_debug(&format!(
        "FFmpeg convert start: input={} output={} args={}",
        input_path.display(),
        output_path.display(),
        args.join(" ")
    ));

    let streams = list_audio_streams(input_path)?;
    let preferred_stream = streams
        .iter()
        .find(|s| s.is_default)
        .or_else(|| streams.first())
        .map(|s| s.index);
    let Some(stream_index) = preferred_stream else {
        return Err("FFmpeg: no audio streams found".to_string());
    };

    let pts_clock = Arc::new(AtomicI64::new(0));
    let mut source =
        FfmpegSource::try_new(input_path, 0, Some(pts_clock.clone()), Some(stream_index))?;
    let in_sample_rate = source.sample_rate();
    let channels = source.channels();
    let total_duration = source.total_duration();
    let total_duration_us = total_duration.and_then(|d| {
        let micros = d.as_micros();
        if micros == 0 {
            None
        } else {
            Some(micros.min(i64::MAX as u128) as i64)
        }
    });
    let total_frames = total_duration.and_then(|d| {
        if d.is_zero() {
            None
        } else {
            let frames = (d.as_secs_f64() * in_sample_rate as f64).round();
            if frames.is_finite() && frames > 0.0 {
                Some(frames as u64)
            } else {
                None
            }
        }
    });
    let out_sample_rate = match settings.format {
        ConvertAudioFormat::Opus => 48_000,
        _ => in_sample_rate,
    };

    let codec_id = match settings.format {
        ConvertAudioFormat::Mp3 => AVCodecID_AV_CODEC_ID_MP3,
        ConvertAudioFormat::Aac => AVCodecID_AV_CODEC_ID_AAC,
        ConvertAudioFormat::Opus => AVCodecID_AV_CODEC_ID_OPUS,
        ConvertAudioFormat::Ogg => AVCodecID_AV_CODEC_ID_VORBIS,
        ConvertAudioFormat::Flac => AVCodecID_AV_CODEC_ID_FLAC,
        ConvertAudioFormat::Wav | ConvertAudioFormat::Aiff => AVCodecID_AV_CODEC_ID_PCM_S16LE,
    };

    let out_path_c = CString::new(output_path.to_string_lossy().as_bytes())
        .map_err(|_| "FFmpeg: invalid output path".to_string())?;
    let mut out_ctx: *mut AVFormatContext = ptr::null_mut();
    let alloc_ret = unsafe {
        (api.avformat_alloc_output_context2)(
            &mut out_ctx,
            ptr::null_mut(),
            ptr::null(),
            out_path_c.as_ptr(),
        )
    };
    if alloc_ret < 0 || out_ctx.is_null() {
        return Err("FFmpeg: failed to allocate output context".to_string());
    }

    let codec = unsafe { (api.avcodec_find_encoder)(codec_id) };
    if codec.is_null() {
        unsafe { (api.avformat_free_context)(out_ctx) };
        return Err("FFmpeg: encoder not found".to_string());
    }
    let stream = unsafe { (api.avformat_new_stream)(out_ctx, codec) };
    if stream.is_null() {
        unsafe { (api.avformat_free_context)(out_ctx) };
        return Err("FFmpeg: failed to create output stream".to_string());
    }

    let mut codec_ctx = unsafe { (api.avcodec_alloc_context3)(codec) };
    if codec_ctx.is_null() {
        unsafe { (api.avformat_free_context)(out_ctx) };
        return Err("FFmpeg: failed to allocate encoder context".to_string());
    }

    let preferred_fmt = match settings.format {
        ConvertAudioFormat::Wav | ConvertAudioFormat::Aiff => {
            Some(AVSampleFormat_AV_SAMPLE_FMT_S16)
        }
        _ => None,
    };
    let enc_fmt = pick_encoder_sample_fmt_with_preference(codec, preferred_fmt);
    unsafe {
        (*codec_ctx).sample_rate = out_sample_rate as i32;
        (*codec_ctx).sample_fmt = enc_fmt;
        (*codec_ctx).time_base = AVRational {
            num: 1,
            den: out_sample_rate as i32,
        };
        let mut out_layout: AVChannelLayout = std::mem::zeroed();
        (api.av_channel_layout_default)(&mut out_layout, channels as i32);
        (*codec_ctx).ch_layout = out_layout;
    }

    match (settings.format, settings.quality) {
        (ConvertAudioFormat::Mp3, ConvertAudioQuality::BitrateKbps(bitrate)) => unsafe {
            let target = (bitrate as i64).saturating_mul(1000);
            (*codec_ctx).bit_rate = target;
            (*codec_ctx).rc_min_rate = target;
            (*codec_ctx).rc_max_rate = target;
            (*codec_ctx).bit_rate_tolerance = 0;
        },
        (
            ConvertAudioFormat::Aac | ConvertAudioFormat::Opus,
            ConvertAudioQuality::BitrateKbps(bitrate),
        ) => unsafe {
            (*codec_ctx).bit_rate = (bitrate as i64).saturating_mul(1000);
        },
        (ConvertAudioFormat::Ogg, ConvertAudioQuality::OggQuality(q)) => unsafe {
            (*codec_ctx).flags |= AV_CODEC_FLAG_QSCALE_FALLBACK;
            (*codec_ctx).global_quality = (q as i32).saturating_mul(FF_QP2LAMBDA_FALLBACK);
        },
        (ConvertAudioFormat::Flac, ConvertAudioQuality::FlacCompression(level)) => unsafe {
            (*codec_ctx).compression_level = level as i32;
        },
        _ => {}
    }

    let open_ret = unsafe { (api.avcodec_open2)(codec_ctx, codec, ptr::null_mut()) };
    if open_ret < 0 {
        unsafe {
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to open encoder".to_string());
    }

    unsafe {
        (*stream).time_base = (*codec_ctx).time_base;
        let par_ret = (api.avcodec_parameters_from_context)((*stream).codecpar, codec_ctx);
        if par_ret < 0 {
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            return Err("FFmpeg: failed to copy encoder parameters".to_string());
        }
    }

    let mut io: *mut AVIOContext = ptr::null_mut();
    let open_io = unsafe { (api.avio_open)(&mut io, out_path_c.as_ptr(), AVIO_FLAG_WRITE) };
    if open_io < 0 {
        unsafe {
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to open output file".to_string());
    }
    unsafe {
        (*out_ctx).pb = io;
    }
    let header_ret = unsafe { (api.avformat_write_header)(out_ctx, ptr::null_mut()) };
    if header_ret < 0 {
        unsafe {
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
        }
        return Err("FFmpeg: failed to write output header".to_string());
    }

    let in_frame_size = unsafe { (*codec_ctx).frame_size }.max(1024) as usize;
    let mut in_layout: AVChannelLayout = unsafe { std::mem::zeroed() };
    let mut out_layout: AVChannelLayout = unsafe { std::mem::zeroed() };
    unsafe {
        (api.av_channel_layout_default)(&mut in_layout, channels as i32);
        (api.av_channel_layout_default)(&mut out_layout, channels as i32);
    }
    let mut swr_ctx: *mut SwrContext = ptr::null_mut();
    let swr_ret = unsafe {
        (api.swr_alloc_set_opts2)(
            &mut swr_ctx,
            &out_layout,
            enc_fmt,
            out_sample_rate as i32,
            &in_layout,
            AVSampleFormat_AV_SAMPLE_FMT_FLT,
            in_sample_rate as i32,
            0,
            ptr::null_mut(),
        )
    };
    if swr_ret < 0 || swr_ctx.is_null() {
        unsafe {
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            (api.av_channel_layout_uninit)(&mut in_layout);
            (api.av_channel_layout_uninit)(&mut out_layout);
        }
        return Err("FFmpeg: failed to init resampler".to_string());
    }
    if unsafe { (api.swr_init)(swr_ctx) } < 0 {
        unsafe {
            (api.swr_free)(&mut swr_ctx);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            (api.av_channel_layout_uninit)(&mut in_layout);
            (api.av_channel_layout_uninit)(&mut out_layout);
        }
        return Err("FFmpeg: failed to init resampler".to_string());
    }

    let mut frame = unsafe { (api.av_frame_alloc)() };
    if frame.is_null() {
        unsafe {
            (api.swr_free)(&mut swr_ctx);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            (api.av_channel_layout_uninit)(&mut in_layout);
            (api.av_channel_layout_uninit)(&mut out_layout);
        }
        return Err("FFmpeg: failed to allocate frame".to_string());
    }
    let mut out_capacity = unsafe { (api.swr_get_out_samples)(swr_ctx, in_frame_size as i32) };
    if out_capacity <= 0 {
        out_capacity = in_frame_size as i32;
    }
    unsafe {
        (*frame).nb_samples = out_capacity;
        (*frame).format = enc_fmt;
        (*frame).sample_rate = out_sample_rate as i32;
        let mut frame_layout: AVChannelLayout = std::mem::zeroed();
        (api.av_channel_layout_default)(&mut frame_layout, channels as i32);
        (*frame).ch_layout = frame_layout;
        if (api.av_frame_get_buffer)(frame, 0) < 0 {
            (api.av_frame_free)(&mut frame);
            (api.swr_free)(&mut swr_ctx);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            (api.av_channel_layout_uninit)(&mut in_layout);
            (api.av_channel_layout_uninit)(&mut out_layout);
            return Err("FFmpeg: failed to allocate frame buffer".to_string());
        }
    }

    let channel_count = channels as usize;
    let mut next_pts: i64 = 0;
    let mut canceled = false;
    let mut processed_frames: u64 = 0;
    let mut last_pct: u32 = 0;
    let mut last_pts_us: i64 = 0;
    if let Some(cb) = progress.as_mut() {
        cb(0);
    }
    let is_canceled = |flag: &Option<Arc<AtomicBool>>| -> bool {
        flag.as_ref()
            .map(|cancel| cancel.load(Ordering::Relaxed))
            .unwrap_or(false)
    };
    let drain_packets = || -> Result<(), String> {
        loop {
            let mut pkt = unsafe { (api.av_packet_alloc)() };
            if pkt.is_null() {
                return Err("FFmpeg: packet alloc failed".to_string());
            }
            let recv = unsafe { (api.avcodec_receive_packet)(codec_ctx, pkt) };
            if recv == 0 {
                unsafe {
                    (*pkt).stream_index = (*stream).index;
                    (api.av_packet_rescale_ts)(pkt, (*codec_ctx).time_base, (*stream).time_base);
                    let wret = (api.av_interleaved_write_frame)(out_ctx, pkt);
                    (api.av_packet_unref)(pkt);
                    (api.av_packet_free)(&mut pkt);
                    if wret < 0 {
                        return Err("FFmpeg: write audio packet failed".to_string());
                    }
                }
            } else {
                unsafe {
                    (api.av_packet_free)(&mut pkt);
                }
                break;
            }
        }
        Ok(())
    };

    loop {
        if is_canceled(&cancel) {
            canceled = true;
            break;
        }
        let mut input_frame = Vec::with_capacity(in_frame_size * channel_count);
        while input_frame.len() < in_frame_size * channel_count {
            if is_canceled(&cancel) {
                canceled = true;
                break;
            }
            match source.next() {
                Some(sample) => input_frame.push(sample),
                None => break,
            }
        }
        if canceled {
            break;
        }
        if input_frame.is_empty() {
            break;
        }
        if input_frame.len() < in_frame_size * channel_count {
            input_frame.resize(in_frame_size * channel_count, 0.0);
        }
        processed_frames =
            processed_frames.saturating_add((input_frame.len() / channel_count) as u64);
        if let (Some(total_us), Some(cb)) = (total_duration_us, progress.as_mut())
            && total_us > 0
        {
            let mut pts_us = pts_clock.load(Ordering::Acquire);
            if pts_us < 0 {
                pts_us = 0;
            }
            if pts_us < last_pts_us {
                pts_us = last_pts_us;
            } else {
                last_pts_us = pts_us;
            }
            let pct = ((pts_us as u128 * 10000) / total_us as u128).min(10000) as u32;
            if pct > last_pct {
                last_pct = pct;
                cb(pct);
            }
        } else if let (Some(total), Some(cb)) = (total_frames, progress.as_mut())
            && total > 0
        {
            let pct = ((processed_frames.saturating_mul(10000)) / total).min(10000) as u32;
            if pct > last_pct {
                last_pct = pct;
                cb(pct);
            }
        }

        unsafe {
            if (api.av_frame_make_writable)(frame) < 0 {
                return Err("FFmpeg: frame not writable".to_string());
            }
        }
        let mut in_ptr = input_frame.as_ptr() as *const u8;
        let in_ptrs = &mut in_ptr as *mut *const u8;
        let out_count = unsafe { (*frame).nb_samples };
        let converted = unsafe {
            (api.swr_convert)(
                swr_ctx,
                (*frame).data.as_mut_ptr(),
                out_count,
                in_ptrs,
                in_frame_size as i32,
            )
        };
        if converted < 0 {
            return Err("FFmpeg: resample failed".to_string());
        }
        unsafe {
            (*frame).nb_samples = converted;
            (*frame).pts = next_pts;
        }
        next_pts += converted as i64;

        let send_ret = unsafe { (api.avcodec_send_frame)(codec_ctx, frame) };
        if send_ret < 0 {
            return Err("FFmpeg: send_frame failed".to_string());
        }
        drain_packets()?;
    }

    if !canceled {
        loop {
            if is_canceled(&cancel) {
                canceled = true;
                break;
            }
            let pending = unsafe { (api.swr_get_out_samples)(swr_ctx, 0) };
            if pending <= 0 {
                break;
            }
            unsafe {
                if (api.av_frame_make_writable)(frame) < 0 {
                    return Err("FFmpeg: frame not writable".to_string());
                }
            }
            let out_count = unsafe { (*frame).nb_samples };
            let converted = unsafe {
                (api.swr_convert)(
                    swr_ctx,
                    (*frame).data.as_mut_ptr(),
                    out_count,
                    ptr::null(),
                    0,
                )
            };
            if converted <= 0 {
                break;
            }
            unsafe {
                (*frame).nb_samples = converted;
                (*frame).pts = next_pts;
            }
            next_pts += converted as i64;
            let send_ret = unsafe { (api.avcodec_send_frame)(codec_ctx, frame) };
            if send_ret < 0 {
                return Err("FFmpeg: send_frame failed".to_string());
            }
            drain_packets()?;
        }
    }

    if canceled {
        unsafe {
            (api.swr_free)(&mut swr_ctx);
            (api.av_frame_free)(&mut frame);
            (api.avio_closep)(&mut io);
            (api.avcodec_free_context)(&mut codec_ctx);
            (api.avformat_free_context)(out_ctx);
            (api.av_channel_layout_uninit)(&mut in_layout);
            (api.av_channel_layout_uninit)(&mut out_layout);
        }
        return Err("Conversion canceled.".to_string());
    }

    if let Some(cb) = progress.as_mut() {
        cb(10000);
    }

    unsafe {
        (api.avcodec_send_frame)(codec_ctx, ptr::null());
        loop {
            let mut pkt = (api.av_packet_alloc)();
            if pkt.is_null() {
                break;
            }
            let recv = (api.avcodec_receive_packet)(codec_ctx, pkt);
            if recv == 0 {
                (*pkt).stream_index = (*stream).index;
                (api.av_packet_rescale_ts)(pkt, (*codec_ctx).time_base, (*stream).time_base);
                (api.av_interleaved_write_frame)(out_ctx, pkt);
                (api.av_packet_unref)(pkt);
                (api.av_packet_free)(&mut pkt);
            } else {
                (api.av_packet_free)(&mut pkt);
                break;
            }
        }
        let trailer_ret = (api.av_write_trailer)(out_ctx);
        if trailer_ret < 0 {
            log_debug(&format!("FFmpeg: av_write_trailer failed: {}", trailer_ret));
        }
        (api.swr_free)(&mut swr_ctx);
        (api.av_frame_free)(&mut frame);
        (api.avio_closep)(&mut io);
        (api.avcodec_free_context)(&mut codec_ctx);
        (api.avformat_free_context)(out_ctx);
        (api.av_channel_layout_uninit)(&mut in_layout);
        (api.av_channel_layout_uninit)(&mut out_layout);
    }

    Ok(())
}

#[cfg(test)]
mod convert_tests {
    use super::*;

    #[test]
    fn test_validate_mp3_bitrate() {
        assert!(validate_mp3_bitrate(95).is_err());
        assert_eq!(validate_mp3_bitrate(96).unwrap(), 96);
        assert_eq!(validate_mp3_bitrate(320).unwrap(), 320);
        assert!(validate_mp3_bitrate(321).is_err());
    }

    #[test]
    fn test_build_ffmpeg_args_mp3() {
        let settings = ConvertAudioSettings {
            format: ConvertAudioFormat::Mp3,
            quality: ConvertAudioQuality::BitrateKbps(192),
        };
        let args = build_ffmpeg_args(&settings);
        assert_eq!(
            args,
            vec![
                "-c:a".to_string(),
                "libmp3lame".to_string(),
                "-b:a".to_string(),
                "192k".to_string(),
            ]
        );
    }
}
