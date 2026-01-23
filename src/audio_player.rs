use crate::log_debug;
use crate::settings::{FileFormat, settings_dir};
use crate::with_state;
use libloading::Library;
use rodio::{Decoder, OutputStream, Sink, Source};
use sha2::Digest;
use std::ffi::c_void;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::OnceLock;
use std::time::Duration;
use windows::Win32::Foundation::HWND;

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
    fn current_frame_len(&self) -> Option<usize> {
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
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(e) => {
                log_debug(&format!(
                    "Audio player: Failed to get default output stream: {}",
                    e
                ));
                return;
            }
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(e) => {
                log_debug(&format!("Audio player: Failed to create sink: {}", e));
                return;
            }
        };

        log_debug(&format!("Audio player: Opening file {}", path.display()));
        let file_size = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();

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
        let source: Box<dyn Source<Item = f32> + Send> = if is_mf_favored {
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
                    let file = std::fs::File::open(&path).map_err(|e| e.to_string()).ok();
                    if let Some(file) = file {
                        let reader = std::io::BufReader::with_capacity(8 * 1024 * 1024, file);
                        match Decoder::new(reader) {
                            Ok(d) => Box::new(
                                d.convert_samples()
                                    .skip_duration(std::time::Duration::from_secs(seconds)),
                            ),
                            Err(err) => {
                                log_debug(&format!(
                                    "Audio player: Rodio fallback decoder failed: {}",
                                    err
                                ));
                                return;
                            }
                        }
                    } else {
                        return;
                    }
                }
            }
        } else if file_size > 100 * 1024 * 1024 {
            log_debug(
                "Audio player: Large file detected (>100MB). Indexing may take a few seconds...",
            );
            let file = match std::fs::File::open(&path) {
                Ok(f) => f,
                Err(e) => {
                    log_debug(&format!("Audio player: failed to open large file: {}", e));
                    return;
                }
            };
            let reader = std::io::BufReader::with_capacity(8 * 1024 * 1024, file);
            match Decoder::new(reader) {
                Ok(d) => {
                    let src = d.convert_samples();
                    if seconds > 0 {
                        log_debug(&format!("Audio player: Skipping to {} seconds", seconds));
                        Box::new(src.skip_duration(std::time::Duration::from_secs(seconds)))
                    } else {
                        Box::new(src)
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
                    match Decoder::new(std::io::Cursor::new(bytes)) {
                        Ok(d) => {
                            let src = d.convert_samples();
                            if seconds > 0 {
                                log_debug(&format!(
                                    "Audio player: Skipping to {} seconds",
                                    seconds
                                ));
                                Box::new(src.skip_duration(std::time::Duration::from_secs(seconds)))
                            } else {
                                Box::new(src)
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
                    match Decoder::new(std::io::BufReader::new(file)) {
                        Ok(d) => {
                            let src = d.convert_samples();
                            if seconds > 0 {
                                log_debug(&format!(
                                    "Audio player: Skipping to {} seconds",
                                    seconds
                                ));
                                Box::new(src.skip_duration(std::time::Duration::from_secs(seconds)))
                            } else {
                                Box::new(src)
                            }
                        }
                        err => {
                            log_debug(&format!(
                                "Audio player: Final decoder attempt failed: {:?}",
                                err.is_err()
                            ));
                            return;
                        }
                    }
                }
            }
        };

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
        if options.paused {
            sink.pause();
        }
        log_debug("Audio player: Playback started");

        let player = AudiobookPlayer {
            path,
            sink: sink.clone(),
            _stream,
            is_paused: options.paused,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume: options.volume,
            muted: options.muted,
            prev_volume: options.prev_volume,
            speed: effective_speed,
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
    let result = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if !player.is_paused {
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
                player.start_instant = std::time::Instant::now();
            }
            let new_pos = (player.accumulated_seconds as i64 + seconds).max(0);
            player.accumulated_seconds = new_pos as u64;
            Some((
                player.path.clone(),
                new_pos as u64,
                player.speed,
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

    let (path, current_pos, speed, paused, volume, muted, prev_volume) = match result {
        Some(v) => v,
        None => return,
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
    let path = with_state(hwnd, |state| {
        state
            .active_audiobook
            .as_ref()
            .map(|player| player.path.clone())
    })
    .flatten()
    .ok_or_else(|| "No active audiobook".to_string())?;

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
