use crate::accessibility::{nvda_speak, to_wide};
use crate::ffmpeg_source::FfmpegSource;
use crate::i18n;
use crate::log_debug;
use crate::settings::{FileFormat, SubtitleReadMode, confirm_title, settings_dir};
use crate::subtitles::{find_subtitle_for_media, load_subtitles};
use crate::tts_engine;
use crate::with_state;
use libloading::Library;
use rodio::{Decoder, OutputStream, OutputStreamBuilder, Sink, Source};
use sha2::Digest;
use std::collections::HashSet;
use std::ffi::c_void;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::Mutex;
use std::sync::OnceLock;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;
use uuid::Uuid;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{IDYES, MB_ICONQUESTION, MB_YESNO, MessageBoxW};
use windows::core::PCWSTR;

pub struct AudiobookPlayer {
    pub path: PathBuf,
    pub sink: Arc<Sink>,
    pub _stream: OutputStream, // Deve essere mantenuto in vita
    pub is_paused: bool,
    pub start_instant: std::time::Instant,
    pub accumulated_seconds: u64,
    pub volume: f32,
    pub muted: bool,
    pub prev_volume: f32,
    pub speed: f32,
    pub subtitle_cancel: Arc<AtomicBool>,
    pub subtitle_hold: bool,
}

type SoundTouchHandle = *mut c_void;
type SoundTouchCreate = unsafe extern "C" fn() -> SoundTouchHandle;
type SoundTouchDestroy = unsafe extern "C" fn(SoundTouchHandle);
type SoundTouchSetSampleRate = unsafe extern "C" fn(SoundTouchHandle, u32);
type SoundTouchSetChannels = unsafe extern "C" fn(SoundTouchHandle, u32);
type SoundTouchSetTempo = unsafe extern "C" fn(SoundTouchHandle, f32);
type SoundTouchPutSamples = unsafe extern "C" fn(SoundTouchHandle, *const f32, u32);
type SoundTouchReceiveSamples = unsafe extern "C" fn(SoundTouchHandle, *mut f32, u32) -> u32;
type SoundTouchFlush = unsafe extern "C" fn(SoundTouchHandle);
type SoundTouchClear = unsafe extern "C" fn(SoundTouchHandle);

struct SoundTouchApi {
    _lib: Library,
    create: SoundTouchCreate,
    destroy: SoundTouchDestroy,
    set_sample_rate: SoundTouchSetSampleRate,
    set_channels: SoundTouchSetChannels,
    set_tempo: SoundTouchSetTempo,
    put_samples: SoundTouchPutSamples,
    receive_samples: SoundTouchReceiveSamples,
    flush: SoundTouchFlush,
    clear: SoundTouchClear,
}

fn load_symbol<T: Copy>(lib: &Library, names: &[&str]) -> Option<T> {
    for name in names {
        let mut symbol_name = Vec::with_capacity(name.len() + 1);
        symbol_name.extend_from_slice(name.as_bytes());
        symbol_name.push(0);
        if let Ok(symbol) = unsafe { lib.get::<T>(&symbol_name) } {
            return Some(*symbol);
        }
    }
    log_debug(&format!("SoundTouch symbol missing: {:?}", names));
    None
}

fn load_soundtouch_api() -> Option<&'static SoundTouchApi> {
    static SOUND_TOUCH: OnceLock<Option<SoundTouchApi>> = OnceLock::new();
    SOUND_TOUCH
        .get_or_init(|| {
            if !cfg!(target_arch = "x86_64") {
                log_debug("SoundTouch DLL not available for this architecture.");
                return None;
            }
            let dll_name = "SoundTouch64.dll";
            let mut candidates = Vec::new();
            candidates.push(settings_dir().join(dll_name));
            if let Ok(appdata) = std::env::var("APPDATA") {
                candidates.push(PathBuf::from(appdata).join("Novapad").join(dll_name));
            }
            if let Ok(exe) = std::env::current_exe()
                && let Some(dir) = exe.parent()
            {
                candidates.push(dir.join("dll").join(dll_name));
                candidates.push(dir.join(dll_name));
            }
            if let Ok(dir) = std::env::current_dir() {
                candidates.push(dir.join("dll").join(dll_name));
                candidates.push(dir.join(dll_name));
            }

            let mut lib = None;
            for path in candidates {
                match unsafe { Library::new(&path) } {
                    Ok(loaded) => {
                        lib = Some(loaded);
                        log_debug(&format!("SoundTouch loaded: {}", path.to_string_lossy()));
                        break;
                    }
                    Err(_) => {
                        log_debug(&format!(
                            "SoundTouch load failed: {}",
                            path.to_string_lossy()
                        ));
                    }
                }
            }
            let lib = lib?;
            let create = load_symbol::<SoundTouchCreate>(
                &lib,
                &[
                    "soundtouch_createInstance",
                    "_soundtouch_createInstance",
                    "soundtouch_createInstance@0",
                ],
            )?;
            let destroy = load_symbol::<SoundTouchDestroy>(
                &lib,
                &[
                    "soundtouch_destroyInstance",
                    "_soundtouch_destroyInstance",
                    "soundtouch_destroyInstance@4",
                ],
            )?;
            let set_sample_rate = load_symbol::<SoundTouchSetSampleRate>(
                &lib,
                &[
                    "soundtouch_setSampleRate",
                    "_soundtouch_setSampleRate",
                    "soundtouch_setSampleRate@8",
                ],
            )?;
            let set_channels = load_symbol::<SoundTouchSetChannels>(
                &lib,
                &[
                    "soundtouch_setChannels",
                    "_soundtouch_setChannels",
                    "soundtouch_setChannels@8",
                ],
            )?;
            let set_tempo = load_symbol::<SoundTouchSetTempo>(
                &lib,
                &[
                    "soundtouch_setTempo",
                    "_soundtouch_setTempo",
                    "soundtouch_setTempo@8",
                ],
            )?;
            let put_samples = load_symbol::<SoundTouchPutSamples>(
                &lib,
                &[
                    "soundtouch_putSamples",
                    "_soundtouch_putSamples",
                    "soundtouch_putSamples@12",
                ],
            )?;
            let receive_samples = load_symbol::<SoundTouchReceiveSamples>(
                &lib,
                &[
                    "soundtouch_receiveSamples",
                    "_soundtouch_receiveSamples",
                    "soundtouch_receiveSamples@12",
                ],
            )?;
            let flush = load_symbol::<SoundTouchFlush>(
                &lib,
                &[
                    "soundtouch_flush",
                    "_soundtouch_flush",
                    "soundtouch_flush@4",
                ],
            )?;
            let clear = load_symbol::<SoundTouchClear>(
                &lib,
                &[
                    "soundtouch_clear",
                    "_soundtouch_clear",
                    "soundtouch_clear@4",
                ],
            )?;
            Some(SoundTouchApi {
                _lib: lib,
                create,
                destroy,
                set_sample_rate,
                set_channels,
                set_tempo,
                put_samples,
                receive_samples,
                flush,
                clear,
            })
        })
        .as_ref()
}

struct SoundTouch {
    api: &'static SoundTouchApi,
    handle: SoundTouchHandle,
    channels: u16,
}

unsafe impl Send for SoundTouch {}

impl SoundTouch {
    fn new(sample_rate: u32, channels: u16, tempo: f32) -> Option<Self> {
        let api = load_soundtouch_api()?;
        unsafe {
            let handle = (api.create)();
            if handle.is_null() {
                return None;
            }
            (api.set_sample_rate)(handle, sample_rate);
            (api.set_channels)(handle, channels as u32);
            (api.set_tempo)(handle, tempo);
            Some(Self {
                api,
                handle,
                channels,
            })
        }
    }

    fn put_samples(&self, samples: &[f32], frames: u32) {
        unsafe {
            (self.api.put_samples)(self.handle, samples.as_ptr(), frames);
        }
    }

    fn receive_samples(&self, out: &mut [f32], max_frames: u32) -> u32 {
        unsafe { (self.api.receive_samples)(self.handle, out.as_mut_ptr(), max_frames) }
    }

    fn flush(&self) {
        unsafe {
            (self.api.flush)(self.handle);
        }
    }
}

impl Drop for SoundTouch {
    fn drop(&mut self) {
        unsafe {
            (self.api.clear)(self.handle);
            (self.api.destroy)(self.handle);
        }
    }
}

struct SoundTouchSource<S>
where
    S: Source<Item = f32>,
{
    input: S,
    st: SoundTouch,
    buffer: Vec<f32>,
    index: usize,
    finished: bool,
}

unsafe impl<S> Send for SoundTouchSource<S> where S: Source<Item = f32> + Send {}

impl<S> SoundTouchSource<S>
where
    S: Source<Item = f32>,
{
    fn try_new(input: S, tempo: f32) -> Result<Self, S> {
        let channels = input.channels();
        let sample_rate = input.sample_rate();
        let st = match SoundTouch::new(sample_rate, channels, tempo) {
            Some(st) => st,
            None => return Err(input),
        };
        Ok(Self {
            input,
            st,
            buffer: Vec::new(),
            index: 0,
            finished: false,
        })
    }

    fn refill(&mut self) -> bool {
        const INPUT_FRAMES: usize = 2048;
        const OUTPUT_FRAMES: usize = 4096;
        let channels = self.st.channels as usize;

        self.buffer.clear();
        self.index = 0;
        let mut produced = false;
        let mut attempts = 0;

        while !produced && attempts < 8 {
            attempts += 1;
            if !self.finished {
                let mut input_samples = Vec::with_capacity(INPUT_FRAMES * channels);
                while input_samples.len() < INPUT_FRAMES * channels {
                    if let Some(sample) = self.input.next() {
                        input_samples.push(sample);
                    } else {
                        break;
                    }
                }
                let frames = input_samples.len() / channels;
                if frames > 0 {
                    self.st.put_samples(&input_samples, frames as u32);
                } else {
                    self.st.flush();
                    self.finished = true;
                }
            } else {
                self.st.flush();
            }

            let mut out = vec![0.0f32; OUTPUT_FRAMES * channels];
            loop {
                let received = self.st.receive_samples(&mut out, OUTPUT_FRAMES as u32);
                if received == 0 {
                    break;
                }
                produced = true;
                let count = received as usize * channels;
                self.buffer.extend_from_slice(&out[..count]);
            }
        }

        !self.buffer.is_empty()
    }
}

impl<S> Iterator for SoundTouchSource<S>
where
    S: Source<Item = f32>,
{
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

impl<S> Source for SoundTouchSource<S>
where
    S: Source<Item = f32>,
{
    fn current_span_len(&self) -> Option<usize> {
        None
    }

    fn channels(&self) -> u16 {
        self.st.channels
    }

    fn sample_rate(&self) -> u32 {
        self.input.sample_rate()
    }

    fn total_duration(&self) -> Option<Duration> {
        None
    }
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
    drop(guard);

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

struct AudiobookPlaybackOptions {
    speed: f32,
    paused: bool,
    volume: f32,
    muted: bool,
    prev_volume: f32,
}

fn start_audiobook_at_with_options(
    hwnd: HWND,
    path: PathBuf,
    seconds: u64,
    options: AudiobookPlaybackOptions,
) {
    let subtitle_hold = should_hold_for_edge_subtitles(hwnd, &path);
    let effective_paused = options.paused || subtitle_hold;
    let subtitle_cancel = Arc::new(AtomicBool::new(false));
    let subtitle_path = path.clone();
    let effective_speed =
        if (options.speed - 1.0).abs() > f32::EPSILON && load_soundtouch_api().is_some() {
            options.speed
        } else {
            1.0
        };
    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        log_debug(&format!(
            "Audio player: Thread started for {}",
            path.display()
        ));
        let stream_handle = match OutputStreamBuilder::open_default_stream() {
            Ok(v) => v,
            Err(e) => {
                log_debug(&format!(
                    "Audio player: Failed to get default output stream: {}",
                    e
                ));
                return;
            }
        };
        let sink: Arc<Sink> = Arc::new(Sink::connect_new(stream_handle.mixer()));

        log_debug(&format!("Audio player: Opening file {}", path.display()));
        let file_size = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        let use_ffmpeg = matches!(
            extension.as_str(),
            "mkv"
                | "avi"
                | "mov"
                | "m4v"
                | "webm"
                | "mpg"
                | "mpeg"
                | "ts"
                | "m2ts"
                | "mts"
                | "wmv"
                | "asf"
                | "flv"
                | "vob"
                | "3gp"
                | "flac"
                | "ogg"
                | "opus"
                | "wma"
                | "aiff"
                | "m4b"
        );
        let prefer_streaming = use_ffmpeg;
        log_mkv_probe_once(&path);

        // Force M4A extension for Media Foundation to avoid indexing hangs on .mp4 video files
        // and to speed up .mp3/.aac opening.
        let mut final_path = path.clone();
        if (extension == "mp4" || extension == "aac" || extension == "mp3")
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

        let extension = final_path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_mf_favored =
            extension == "m4a" || extension == "mp4" || extension == "aac" || extension == "mp3";

        log_debug("Audio player: Creating decoder...");
        let source: Box<dyn Source<Item = f32> + Send> = if use_ffmpeg {
            log_debug(&format!("Audio player: Using FFmpeg for {}", extension));
            match FfmpegSource::try_new(&final_path, seconds) {
                Ok(src) => Box::new(src),
                Err(err) => {
                    log_debug(&format!(
                        "Audio player: FFmpeg failed for {}, falling back to rodio: {}",
                        extension, err
                    ));
                    let file = match std::fs::File::open(&path) {
                        Ok(file) => file,
                        Err(e) => {
                            log_debug(&format!(
                                "Audio player: Rodio fallback failed to open file: {}",
                                e
                            ));
                            return;
                        }
                    };
                    let reader = std::io::BufReader::with_capacity(8 * 1024 * 1024, file);
                    let mut builder = Decoder::builder().with_data(reader);
                    if file_size > 0 {
                        builder = builder.with_byte_len(file_size);
                    }
                    if !extension.is_empty() {
                        builder = builder.with_hint(&extension);
                    }
                    match builder.build() {
                        Ok(d) => Box::new(d.skip_duration(std::time::Duration::from_secs(seconds))),
                        Err(err) => {
                            log_debug(&format!(
                                "Audio player: Rodio fallback decoder failed: {}",
                                err
                            ));
                            return;
                        }
                    }
                }
            }
        } else if is_mf_favored {
            log_debug(&format!(
                "Audio player: Using Media Foundation for {}",
                extension
            ));
            match crate::mf_source::MfSource::try_new(&final_path) {
                Ok(mut mfs) => {
                    if seconds > 0 {
                        log_debug(&format!("Audio player: Efficient seek to {}s", seconds));
                        if let Err(e) = mfs.seek(std::time::Duration::from_secs(seconds)) {
                            log_debug(&format!("Audio player: MfSource seek failed: {}", e));
                        }
                    }
                    Box::new(mfs)
                }
                Err(e) => {
                    log_debug(&format!(
                        "Audio player: MfSource failed for {}, falling back to rodio: {}",
                        extension, e
                    ));
                    // Fallback to rodio decoder
                    let file = match std::fs::File::open(&path) {
                        Ok(file) => file,
                        Err(e) => {
                            log_debug(&format!(
                                "Audio player: Rodio fallback failed to open file: {}",
                                e
                            ));
                            return;
                        }
                    };
                    let reader = std::io::BufReader::with_capacity(8 * 1024 * 1024, file);
                    let mut builder = Decoder::builder().with_data(reader);
                    if file_size > 0 {
                        builder = builder.with_byte_len(file_size);
                    }
                    if !extension.is_empty() {
                        builder = builder.with_hint(&extension);
                    }
                    if prefer_streaming {
                        builder = builder.with_gapless(false);
                    }
                    match builder.build() {
                        Ok(d) => Box::new(d.skip_duration(std::time::Duration::from_secs(seconds))),
                        Err(err) => {
                            log_debug(&format!(
                                "Audio player: Rodio fallback decoder failed: {}",
                                err
                            ));
                            return;
                        }
                    }
                }
            }
        } else if prefer_streaming || file_size > 100 * 1024 * 1024 {
            if prefer_streaming {
                log_debug("Audio player: Streaming decode path selected.");
            } else {
                log_debug(
                    "Audio player: Large file detected (>100MB). Indexing may take a few seconds...",
                );
            }
            let file = match std::fs::File::open(&path) {
                Ok(f) => f,
                Err(e) => {
                    log_debug(&format!("Audio player: failed to open large file: {}", e));
                    return;
                }
            };
            let reader = std::io::BufReader::with_capacity(8 * 1024 * 1024, file);
            let mut builder = Decoder::builder().with_data(reader);
            if file_size > 0 {
                builder = builder.with_byte_len(file_size);
            }
            if !extension.is_empty() {
                builder = builder.with_hint(&extension);
            }
            if prefer_streaming {
                builder = builder.with_gapless(false);
            }
            match builder.build() {
                Ok(d) => {
                    if seconds > 0 {
                        log_debug(&format!("Audio player: Skipping to {} seconds", seconds));
                        Box::new(d.skip_duration(std::time::Duration::from_secs(seconds)))
                    } else {
                        Box::new(d)
                    }
                }
                Err(e) => {
                    log_debug(&format!("Audio player: Failed to create decoder: {}", e));
                    return;
                }
            }
        } else {
            match std::fs::read(&path) {
                Ok(bytes) => {
                    log_debug(&format!(
                        "Audio player: Read {} bytes into memory",
                        bytes.len()
                    ));
                    let bytes_len = bytes.len() as u64;
                    let mut builder = Decoder::builder().with_data(std::io::Cursor::new(bytes));
                    if bytes_len > 0 {
                        builder = builder.with_byte_len(bytes_len);
                    }
                    if !extension.is_empty() {
                        builder = builder.with_hint(&extension);
                    }
                    if prefer_streaming {
                        builder = builder.with_gapless(false);
                    }
                    match builder.build() {
                        Ok(d) => {
                            if seconds > 0 {
                                log_debug(&format!(
                                    "Audio player: Skipping to {} seconds",
                                    seconds
                                ));
                                Box::new(d.skip_duration(std::time::Duration::from_secs(seconds)))
                            } else {
                                Box::new(d)
                            }
                        }
                        Err(e) => {
                            log_debug(&format!(
                                "Audio player: Failed to create memory decoder: {}",
                                e
                            ));
                            return;
                        }
                    }
                }
                Err(e) => {
                    log_debug(&format!(
                        "Audio player: Failed to read file into memory: {}",
                        e
                    ));
                    let file = match std::fs::File::open(&path) {
                        Ok(f) => f,
                        Err(e) => {
                            log_debug(&format!(
                                "Audio player: failed to open file for fallback: {}",
                                e
                            ));
                            return;
                        }
                    };
                    let reader = std::io::BufReader::new(file);
                    let mut builder = Decoder::builder().with_data(reader);
                    if file_size > 0 {
                        builder = builder.with_byte_len(file_size);
                    }
                    if !extension.is_empty() {
                        builder = builder.with_hint(&extension);
                    }
                    if prefer_streaming {
                        builder = builder.with_gapless(false);
                    }
                    match builder.build() {
                        Ok(d) => {
                            if seconds > 0 {
                                log_debug(&format!(
                                    "Audio player: Skipping to {} seconds",
                                    seconds
                                ));
                                Box::new(d.skip_duration(std::time::Duration::from_secs(seconds)))
                            } else {
                                Box::new(d)
                            }
                        }
                        Err(err) => {
                            log_debug(&format!(
                                "Audio player: Final decoder attempt failed: {}",
                                err
                            ));
                            return;
                        }
                    }
                }
            }
        };

        log_debug(&format!(
            "Audio player: Source format {}ch @ {} Hz",
            source.channels(),
            source.sample_rate()
        ));

        if (effective_speed - 1.0).abs() > f32::EPSILON {
            log_debug(&format!(
                "Audio player: Applying speed factor {}",
                effective_speed
            ));
            match SoundTouchSource::try_new(source, effective_speed) {
                Ok(st_source) => sink.append(st_source),
                Err(source) => sink.append(source),
            }
        } else {
            log_debug("Audio player: Appending source to sink");
            sink.append(source);
        }

        if options.muted {
            sink.set_volume(0.0);
        } else {
            sink.set_volume(options.volume);
        }
        if effective_paused {
            sink.pause();
        }
        log_debug("Audio player: Playback started");

        let player = AudiobookPlayer {
            path,
            sink: sink.clone(),
            _stream: stream_handle,
            is_paused: effective_paused,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume: options.volume,
            muted: options.muted,
            prev_volume: options.prev_volume,
            speed: effective_speed,
            subtitle_cancel: subtitle_cancel.clone(),
            subtitle_hold,
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
        start_subtitle_reader(hwnd_main, subtitle_path, subtitle_cancel);
    });
}

pub unsafe fn start_audiobook_playback(hwnd: HWND, path: &Path) {
    crate::log_debug(&format!(
        "Audio player: start_audiobook_playback called for {}",
        path.display()
    ));
    crate::reset_active_podcast_chapters_for_playback(hwnd);
    let path_buf = path.to_path_buf();

    let (bookmark_pos, speed, volume) = with_state(hwnd, |state| {
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
            state.settings.audiobook_playback_volume,
        )
    })
    .unwrap_or((0, 1.0, 1.0));

    start_audiobook_at_with_options(
        hwnd,
        path_buf,
        bookmark_pos as u64,
        AudiobookPlaybackOptions {
            speed,
            paused: false,
            volume,
            muted: false,
            prev_volume: volume,
        },
    );
}

pub unsafe fn toggle_audiobook_pause(hwnd: HWND) {
    crate::log_debug("Audio player: toggle_audiobook_pause triggered");
    let start_action = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.is_paused {
                crate::log_debug("Audio player: Resuming playback");
                player.sink.play();
                player.is_paused = false;
                player.start_instant = std::time::Instant::now();
            } else {
                crate::log_debug("Audio player: Pausing playback");
                player.sink.pause();
                player.is_paused = true;
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
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
        Applied,
        Restart {
            path: PathBuf,
            current_pos: u64,
            speed: f32,
            paused: bool,
            volume: f32,
            muted: bool,
            prev_volume: f32,
        },
    }

    let result = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if !player.is_paused {
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
                player.start_instant = std::time::Instant::now();
            }
            let new_pos = (player.accumulated_seconds as i64 + seconds).max(0);
            if (player.speed - 1.0).abs() <= f32::EPSILON
                && player
                    .sink
                    .try_seek(std::time::Duration::from_secs(new_pos as u64))
                    .is_ok()
            {
                player.accumulated_seconds = new_pos as u64;
                if !player.is_paused {
                    player.start_instant = std::time::Instant::now();
                }
                return Some(SeekAction::Applied);
            }
            player.accumulated_seconds = new_pos as u64;
            Some(SeekAction::Restart {
                path: player.path.clone(),
                current_pos: new_pos as u64,
                speed: player.speed,
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

    if matches!(action, SeekAction::Applied) {
        return;
    }

    let SeekAction::Restart {
        path,
        current_pos,
        speed,
        paused,
        volume,
        muted,
        prev_volume,
    } = action
    else {
        return;
    };

    stop_audiobook_playback(hwnd);
    start_audiobook_at_with_options(
        hwnd,
        path,
        current_pos,
        AudiobookPlaybackOptions {
            speed,
            paused,
            volume,
            muted,
            prev_volume,
        },
    );
}

pub unsafe fn seek_audiobook_to(hwnd: HWND, seconds: u64) -> Result<(), String> {
    enum SeekToAction {
        Applied,
        Restart(PathBuf),
    }

    let action = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if (player.speed - 1.0).abs() <= f32::EPSILON
                && player
                    .sink
                    .try_seek(std::time::Duration::from_secs(seconds))
                    .is_ok()
            {
                player.accumulated_seconds = seconds;
                if !player.is_paused {
                    player.start_instant = std::time::Instant::now();
                }
                return Some(SeekToAction::Applied);
            }
            return Some(SeekToAction::Restart(player.path.clone()));
        }
        None
    })
    .flatten()
    .ok_or_else(|| "No active audiobook".to_string())?;

    let path = match action {
        SeekToAction::Applied => return Ok(()),
        SeekToAction::Restart(path) => path,
    };

    start_audiobook_at(hwnd, &path, seconds);
    Ok(())
}

pub unsafe fn stop_audiobook_playback(hwnd: HWND) {
    crate::log_debug("Audio player: stop_audiobook_playback called");
    if with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            crate::log_debug(&format!(
                "Audio player: Stopping and removing player for {}",
                player.path.display()
            ));
            state.last_stopped_audiobook = Some(player.path.clone());
            player.subtitle_cancel.store(true, Ordering::Relaxed);
            player.sink.stop();
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
    let (speed, volume, muted, prev_volume) = with_state(hwnd, |state| {
        if let Some(player) = &state.active_audiobook {
            (
                player.speed,
                player.volume,
                player.muted,
                player.prev_volume,
            )
        } else {
            (1.0, 1.0, false, 1.0)
        }
    })
    .unwrap_or((1.0, 1.0, false, 1.0));

    stop_audiobook_playback(hwnd);
    let path_buf = path.to_path_buf();
    start_audiobook_at_with_options(
        hwnd,
        path_buf,
        seconds,
        AudiobookPlaybackOptions {
            speed,
            paused: false,
            volume,
            muted,
            prev_volume,
        },
    );
}

pub unsafe fn change_audiobook_volume(hwnd: HWND, delta: f32) {
    let new_volume = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.muted {
                player.prev_volume = (player.prev_volume + delta).clamp(0.0, 3.0);
                return None;
            }
            player.volume = (player.volume + delta).clamp(0.0, 3.0);
            player.sink.set_volume(player.volume);
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
    load_soundtouch_api()?;
    let result = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            let current = if player.is_paused {
                player.accumulated_seconds
            } else {
                player.accumulated_seconds + player.start_instant.elapsed().as_secs()
            };
            let new_speed = (player.speed + delta).clamp(0.5, 3.0);
            player.sink.stop();
            Some((
                player.path,
                current,
                new_speed,
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

    let (path, current, speed, paused, volume, muted, prev_volume) = result?;

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            paused,
            volume,
            muted,
            prev_volume,
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

pub unsafe fn reset_audiobook_speed(hwnd: HWND) -> Option<f32> {
    load_soundtouch_api()?;
    let result = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            let current = if player.is_paused {
                player.accumulated_seconds
            } else {
                player.accumulated_seconds + player.start_instant.elapsed().as_secs()
            };
            let new_speed = 1.0;
            player.sink.stop();
            Some((
                player.path,
                current,
                new_speed,
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

    let (path, current, speed, paused, volume, muted, prev_volume) = result?;

    start_audiobook_at_with_options(
        hwnd,
        path,
        current,
        AudiobookPlaybackOptions {
            speed,
            paused,
            volume,
            muted,
            prev_volume,
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

pub unsafe fn audiobook_volume_level(hwnd: HWND) -> Option<f32> {
    with_state(hwnd, |state| {
        state
            .active_audiobook
            .as_ref()
            .map(|player| if player.muted { 0.0 } else { player.volume })
    })
    .flatten()
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
                player.sink.set_volume(player.volume);
            } else {
                if player.volume > 0.0 {
                    player.prev_volume = player.volume;
                }
                player.volume = 0.0;
                player.muted = true;
                player.sink.set_volume(0.0);
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
}

static SUBTITLE_EDGE_CONFIRMED: OnceLock<Mutex<HashSet<String>>> = OnceLock::new();
const SUBTITLE_AUDIO_LATENCY_SECS: f64 = 0.28;

fn audiobook_position_secs(player: &AudiobookPlayer) -> f64 {
    if player.is_paused {
        player.accumulated_seconds as f64
    } else {
        player.accumulated_seconds as f64
            + player.start_instant.elapsed().as_secs_f64() * player.speed as f64
    }
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
            })
        })
        .flatten()
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
            player.sink.play();
            true
        })
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
    let subtitle_path = match find_subtitle_for_media(media_path) {
        Some(path) => path,
        None => return false,
    };
    let _ = subtitle_path;
    let base_cache_dir = settings_dir().join("subtitle_cache");
    let mut hasher = sha2::Sha256::new();
    hasher.update(media_path.to_string_lossy().as_bytes());
    hasher.update(settings.tts_voice.as_bytes());
    let hash = hex::encode(hasher.finalize());
    let dir = base_cache_dir.join(&hash[..16]);
    if !dir.exists() {
        return true;
    }
    let cache_ready = std::fs::read_dir(&dir)
        .map(|mut it| {
            it.any(|entry| {
                entry
                    .ok()
                    .and_then(|e| {
                        e.path()
                            .extension()
                            .and_then(|s| s.to_str())
                            .map(|s| s.eq_ignore_ascii_case("mp3"))
                    })
                    .unwrap_or(false)
            })
        })
        .unwrap_or(false);
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

fn start_subtitle_reader(hwnd: HWND, media_path: PathBuf, cancel: Arc<AtomicBool>) {
    std::thread::spawn(move || {
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
        let effective_mode = match mode {
            SubtitleReadMode::Off => SubtitleReadMode::Off,
            SubtitleReadMode::Nvda => SubtitleReadMode::Nvda,
            _ => SubtitleReadMode::User,
        };
        let subtitle_path = match find_subtitle_for_media(&media_path) {
            Some(path) => path,
            None => return,
        };
        let mut cues = match load_subtitles(&subtitle_path) {
            Ok(cues) => cues,
            Err(err) => {
                log_debug(&format!("Subtitle: {}", err));
                return;
            }
        };
        if cues.is_empty() {
            return;
        }

        let mut subtitle_stream: Option<OutputStream> = None;
        let mut subtitle_sink: Option<Sink> = None;

        let mut paused_for_download = subtitle_hold_state(hwnd, &media_path).unwrap_or(false);
        if effective_mode == SubtitleReadMode::User
            && settings.tts_engine == crate::settings::TtsEngine::Edge
        {
            let edge_key = edge_subtitle_key(&media_path, &settings);
            let msg = i18n::tr(settings.language, "subtitles.edge_confirm");
            let title = confirm_title(settings.language);
            let msg_w = to_wide(&msg);
            let title_w = to_wide(&title);
            let mut paused_for_prompt = false;
            if !paused_for_download
                && let Some(player) = unsafe {
                    with_state(hwnd, |state| {
                        state.active_audiobook.as_ref().map(|p| p.sink.clone())
                    })
                }
                .flatten()
                && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                && !state.paused
            {
                player.pause();
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
                        && let Some(player) = unsafe {
                            with_state(hwnd, |state| {
                                state.active_audiobook.as_ref().map(|p| p.sink.clone())
                            })
                        }
                        .flatten()
                        && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                        && !state.paused
                    {
                        player.play();
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
            let cache_ready = dir.exists()
                && std::fs::read_dir(&dir)
                    .map(|mut it| {
                        it.any(|entry| {
                            entry
                                .ok()
                                .and_then(|e| {
                                    e.path()
                                        .extension()
                                        .and_then(|s| s.to_str())
                                        .map(|s| s.eq_ignore_ascii_case("mp3"))
                                })
                                .unwrap_or(false)
                        })
                    })
                    .unwrap_or(false);
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
            let rt = match tokio::runtime::Runtime::new() {
                Ok(rt) => rt,
                Err(e) => {
                    log_debug(&format!("Subtitle: failed to create runtime: {}", e));
                    return;
                }
            };

            if !cache_ready {
                if let Some(player) = unsafe {
                    with_state(hwnd, |state| {
                        state.active_audiobook.as_ref().map(|p| p.sink.clone())
                    })
                }
                .flatten()
                    && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                    && !state.paused
                {
                    player.pause();
                    paused_for_download = true;
                }
                for (idx, cue) in cues.iter_mut().enumerate() {
                    if cancel.load(Ordering::Relaxed) {
                        return;
                    }
                    let text = cue.text.trim();
                    if text.is_empty() {
                        continue;
                    }
                    let request_id = Uuid::new_v4().simple().to_string();
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
                            let path = dir.join(format!("cue_{:04}.mp3", idx));
                            match std::fs::write(&path, bytes) {
                                Ok(()) => cue.audio_path = Some(path),
                                Err(e) => {
                                    log_debug(&format!(
                                        "Subtitle: failed to write audio chunk: {}",
                                        e
                                    ));
                                }
                            }
                        }
                        Err(err) => {
                            log_debug(&format!("Subtitle: download failed: {}", err));
                        }
                    }
                }
            } else {
                for (idx, cue) in cues.iter_mut().enumerate() {
                    let path = dir.join(format!("cue_{:04}.mp3", idx));
                    if path.exists() {
                        cue.audio_path = Some(path);
                    }
                }
            }

            if cancel.load(Ordering::Relaxed) {
                return;
            }
            let stream_handle = match OutputStreamBuilder::open_default_stream() {
                Ok(v) => v,
                Err(e) => {
                    log_debug(&format!("Subtitle: Failed to get output stream: {}", e));
                    return;
                }
            };
            let sink = Sink::connect_new(stream_handle.mixer());
            subtitle_stream = Some(stream_handle);
            subtitle_sink = Some(sink);
            if paused_for_download
                && let Some(held) = subtitle_hold_state(hwnd, &media_path)
                && held
            {
                clear_subtitle_hold(hwnd, &media_path);
            } else if paused_for_download
                && let Some(player) = unsafe {
                    with_state(hwnd, |state| {
                        state.active_audiobook.as_ref().map(|p| p.sink.clone())
                    })
                }
                .flatten()
                && let Some(state) = subtitle_playback_state(hwnd, &media_path)
                && !state.paused
            {
                player.play();
            }
        }

        let (mut index, mut last_position, mut last_paused) =
            if let Some(state) = subtitle_playback_state(hwnd, &media_path) {
                let adjusted_pos = (state.position_secs - SUBTITLE_AUDIO_LATENCY_SECS).max(0.0);
                let index = cues
                    .iter()
                    .position(|cue| cue.end.as_secs_f64() >= adjusted_pos)
                    .unwrap_or(cues.len());
                let last_paused = state.paused;
                if last_paused && let Some(sink) = subtitle_sink.as_ref() {
                    sink.pause();
                }
                (index, adjusted_pos, last_paused)
            } else {
                return;
            };

        loop {
            if cancel.load(Ordering::Relaxed) {
                break;
            }
            let state = match subtitle_playback_state(hwnd, &media_path) {
                Some(state) => state,
                None => break,
            };

            if state.paused {
                if !last_paused && let Some(sink) = subtitle_sink.as_ref() {
                    sink.pause();
                }
                last_paused = true;
                std::thread::sleep(Duration::from_millis(80));
                continue;
            } else if last_paused {
                if let Some(sink) = subtitle_sink.as_ref() {
                    sink.play();
                }
                last_paused = false;
            }

            let adjusted_pos = (state.position_secs - SUBTITLE_AUDIO_LATENCY_SECS).max(0.0);
            if adjusted_pos + 0.05 < last_position {
                log_debug(&format!(
                    "SubtitleSync: seek detected at {:.3}s (prev {:.3}s) for {}",
                    adjusted_pos,
                    last_position,
                    media_path.display()
                ));
                if let Some(pos) = cues
                    .iter()
                    .position(|cue| cue.end.as_secs_f64() >= adjusted_pos)
                {
                    index = pos;
                } else {
                    index = cues.len();
                }
            }
            last_position = adjusted_pos;

            while index < cues.len() && adjusted_pos >= cues[index].start.as_secs_f64() {
                let cue = cues[index].clone();
                let mut preview = cue.text.replace('\n', " ");
                if preview.len() > 80 {
                    preview.truncate(80);
                    preview.push_str("...");
                }
                log_debug(&format!(
                    "SubtitleSync: speak idx={} start={:.3}s end={:.3}s pos={:.3}s speed={:.2} text='{}'",
                    index,
                    cue.start.as_secs_f64(),
                    cue.end.as_secs_f64(),
                    adjusted_pos,
                    settings.audiobook_playback_speed,
                    preview
                ));
                match effective_mode {
                    SubtitleReadMode::Off => {}
                    SubtitleReadMode::Nvda => {
                        if !nvda_speak(&cue.text) {
                            log_debug("Subtitle: NVDA speak failed.");
                        }
                    }
                    SubtitleReadMode::User
                    | SubtitleReadMode::Sapi5
                    | SubtitleReadMode::Sapi4
                    | SubtitleReadMode::Edge => match settings.tts_engine {
                        crate::settings::TtsEngine::Edge => {
                            if let (Some(path), Some(sink)) =
                                (cue.audio_path.as_ref(), subtitle_sink.as_ref())
                            {
                                match std::fs::File::open(path) {
                                    Ok(file) => match Decoder::new(std::io::BufReader::new(file)) {
                                        Ok(source) => sink.append(source),
                                        Err(e) => log_debug(&format!(
                                            "Subtitle: failed to decode audio chunk: {}",
                                            e
                                        )),
                                    },
                                    Err(e) => {
                                        log_debug(&format!(
                                            "Subtitle: failed to open chunk: {}",
                                            e
                                        ));
                                    }
                                }
                            }
                        }
                        crate::settings::TtsEngine::Sapi5 => {
                            let cancel_flag = Arc::new(AtomicBool::new(false));
                            let (command_tx, command_rx) = tokio::sync::mpsc::unbounded_channel();
                            drop(command_tx);
                            if let Err(err) = crate::sapi5_engine::play_sapi(
                                vec![cue.text.clone()],
                                settings.tts_voice.clone(),
                                settings.tts_rate,
                                settings.tts_pitch,
                                settings.tts_volume,
                                cancel_flag,
                                command_rx,
                            ) {
                                log_debug(&format!("Subtitle: SAPI5 failed: {}", err));
                            }
                        }
                        crate::settings::TtsEngine::Sapi4 => {
                            let voice_index = match parse_sapi4_voice_index(&settings.tts_voice) {
                                Some(idx) => idx,
                                None => {
                                    log_debug("Subtitle: invalid SAPI4 voice, defaulting to 0.");
                                    0
                                }
                            };
                            let cancel_flag = Arc::new(AtomicBool::new(false));
                            let (_command_tx, command_rx) = tokio::sync::mpsc::unbounded_channel();
                            crate::sapi4_engine::play_sapi4(
                                voice_index,
                                cue.text.clone(),
                                settings.tts_rate,
                                settings.tts_pitch,
                                settings.tts_volume,
                                cancel_flag,
                                command_rx,
                            );
                        }
                    },
                }
                index += 1;
            }

            std::thread::sleep(Duration::from_millis(80));
        }

        drop(subtitle_sink);
        drop(subtitle_stream);
    });
}

#[cfg(test)]
mod tests {
    use super::parse_time_input;

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
}
