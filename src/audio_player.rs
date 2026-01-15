use crate::accessibility::to_wide;
use crate::log_debug;
use crate::settings::{FileFormat, settings_dir};
use crate::with_state;
use rodio::{Decoder, OutputStream, Sink, Source};
use std::ffi::c_void;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::OnceLock;
use std::time::Duration;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
use windows::core::{PCSTR, PCWSTR};

fn read_u16_le(buf: &[u8], offset: usize) -> Option<u16> {
    let bytes = buf.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32_le(buf: &[u8], offset: usize) -> Option<u32> {
    let bytes = buf.get(offset..offset + 4)?;
    Some(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn rva_to_offset(sections: &[(u32, u32, u32)], rva: u32) -> Option<usize> {
    for (virt_addr, virt_size, raw_ptr) in sections {
        if rva >= *virt_addr && rva < (*virt_addr).saturating_add(*virt_size) {
            let offset = rva - *virt_addr;
            return Some((*raw_ptr).saturating_add(offset) as usize);
        }
    }
    None
}

fn read_export_names(path: &Path) -> Vec<String> {
    let Ok(buf) = std::fs::read(path) else {
        return Vec::new();
    };
    if buf.len() < 0x40 || &buf[0..2] != b"MZ" {
        return Vec::new();
    }
    let e_lfanew = match read_u32_le(&buf, 0x3c) {
        Some(v) => v as usize,
        None => return Vec::new(),
    };
    if buf.len() < e_lfanew + 24 || &buf[e_lfanew..e_lfanew + 4] != b"PE\0\0" {
        return Vec::new();
    }
    let file_header_offset = e_lfanew + 4;
    let number_of_sections = match read_u16_le(&buf, file_header_offset + 2) {
        Some(v) => v as usize,
        None => return Vec::new(),
    };
    let size_of_optional_header = match read_u16_le(&buf, file_header_offset + 16) {
        Some(v) => v as usize,
        None => return Vec::new(),
    };
    let optional_offset = file_header_offset + 20;
    if buf.len() < optional_offset + size_of_optional_header {
        return Vec::new();
    }
    let magic = read_u16_le(&buf, optional_offset).unwrap_or(0);
    let data_dir_offset = match magic {
        0x10b => optional_offset + 0x60,
        0x20b => optional_offset + 0x70,
        _ => return Vec::new(),
    };
    let export_rva = read_u32_le(&buf, data_dir_offset).unwrap_or(0);
    let export_size = read_u32_le(&buf, data_dir_offset + 4).unwrap_or(0);
    if export_rva == 0 || export_size == 0 {
        return Vec::new();
    }
    let section_offset = optional_offset + size_of_optional_header;
    let mut sections = Vec::new();
    for i in 0..number_of_sections {
        let base = section_offset + i * 40;
        if buf.len() < base + 40 {
            break;
        }
        let virtual_size = read_u32_le(&buf, base + 8).unwrap_or(0);
        let virtual_address = read_u32_le(&buf, base + 12).unwrap_or(0);
        let raw_size = read_u32_le(&buf, base + 16).unwrap_or(0);
        let raw_ptr = read_u32_le(&buf, base + 20).unwrap_or(0);
        let size = std::cmp::max(virtual_size, raw_size);
        sections.push((virtual_address, size, raw_ptr));
    }
    let export_offset = match rva_to_offset(&sections, export_rva) {
        Some(v) => v,
        None => return Vec::new(),
    };
    if buf.len() < export_offset + 40 {
        return Vec::new();
    }
    let number_of_names = read_u32_le(&buf, export_offset + 24).unwrap_or(0) as usize;
    let address_of_names = read_u32_le(&buf, export_offset + 32).unwrap_or(0);
    let names_offset = match rva_to_offset(&sections, address_of_names) {
        Some(v) => v,
        None => return Vec::new(),
    };
    let mut names = Vec::new();
    for i in 0..number_of_names {
        let name_rva = match read_u32_le(&buf, names_offset + i * 4) {
            Some(v) => v,
            None => break,
        };
        let name_offset = match rva_to_offset(&sections, name_rva) {
            Some(v) => v,
            None => continue,
        };
        let mut end = name_offset;
        while end < buf.len() && buf[end] != 0 {
            end += 1;
        }
        if end > name_offset && end <= buf.len() {
            if let Ok(s) = std::str::from_utf8(&buf[name_offset..end]) {
                names.push(s.to_string());
            }
        }
    }
    names
}

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
    _handle: windows::Win32::Foundation::HMODULE,
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

fn load_soundtouch_api() -> Option<&'static SoundTouchApi> {
    static SOUND_TOUCH: OnceLock<Option<SoundTouchApi>> = OnceLock::new();
    SOUND_TOUCH
        .get_or_init(|| unsafe {
            let dll_name = if cfg!(target_arch = "x86_64") {
                "SoundTouch64.dll"
            } else {
                "SoundTouch64.dll"
            };
            let mut candidates = Vec::new();
            candidates.push(settings_dir().join(dll_name));
            if let Ok(appdata) = std::env::var("APPDATA") {
                candidates.push(PathBuf::from(appdata).join("Novapad").join(dll_name));
            }
            if let Ok(exe) = std::env::current_exe() {
                if let Some(dir) = exe.parent() {
                    candidates.push(dir.join("dll").join(dll_name));
                    candidates.push(dir.join(dll_name));
                }
            }
            if let Ok(dir) = std::env::current_dir() {
                candidates.push(dir.join("dll").join(dll_name));
                candidates.push(dir.join(dll_name));
            }

            let mut h = None;
            let mut loaded_path = None;
            for path in candidates {
                let dll_path_wide = to_wide(&path.to_string_lossy());
                if let Ok(handle) = LoadLibraryW(PCWSTR(dll_path_wide.as_ptr())) {
                    h = Some(handle);
                    loaded_path = Some(path);
                    break;
                } else {
                    log_debug(&format!(
                        "SoundTouch load failed: {}",
                        path.to_string_lossy()
                    ));
                }
            }
            let h = h?;
            if let Some(path) = loaded_path {
                let exports = read_export_names(&path);
                let mut filtered: Vec<String> = exports
                    .into_iter()
                    .filter(|name| {
                        let lower = name.to_lowercase();
                        lower.contains("soundtouch")
                            || lower.contains("tempo")
                            || lower.contains("sample")
                    })
                    .collect();
                filtered.sort();
                if !filtered.is_empty() {
                    let preview = if filtered.len() > 40 {
                        filtered[..40].join(", ")
                    } else {
                        filtered.join(", ")
                    };
                    log_debug(&format!("SoundTouch exports: {}", preview));
                }
            }
            let proc = |names: &[&str]| {
                for name in names {
                    if let Ok(cstr) = std::ffi::CString::new(*name) {
                        if let Some(addr) = GetProcAddress(h, PCSTR(cstr.as_ptr() as *const u8)) {
                            return Some(addr);
                        }
                    }
                }
                log_debug(&format!("SoundTouch symbol missing: {:?}", names));
                None
            };
            Some(SoundTouchApi {
                _handle: h,
                create: std::mem::transmute(proc(&[
                    "soundtouch_createInstance",
                    "_soundtouch_createInstance",
                    "soundtouch_createInstance@0",
                ])?),
                destroy: std::mem::transmute(proc(&[
                    "soundtouch_destroyInstance",
                    "_soundtouch_destroyInstance",
                    "soundtouch_destroyInstance@4",
                ])?),
                set_sample_rate: std::mem::transmute(proc(&[
                    "soundtouch_setSampleRate",
                    "_soundtouch_setSampleRate",
                    "soundtouch_setSampleRate@8",
                ])?),
                set_channels: std::mem::transmute(proc(&[
                    "soundtouch_setChannels",
                    "_soundtouch_setChannels",
                    "soundtouch_setChannels@8",
                ])?),
                set_tempo: std::mem::transmute(proc(&[
                    "soundtouch_setTempo",
                    "_soundtouch_setTempo",
                    "soundtouch_setTempo@8",
                ])?),
                put_samples: std::mem::transmute(proc(&[
                    "soundtouch_putSamples",
                    "_soundtouch_putSamples",
                    "soundtouch_putSamples@12",
                ])?),
                receive_samples: std::mem::transmute(proc(&[
                    "soundtouch_receiveSamples",
                    "_soundtouch_receiveSamples",
                    "soundtouch_receiveSamples@12",
                ])?),
                flush: std::mem::transmute(proc(&[
                    "soundtouch_flush",
                    "_soundtouch_flush",
                    "soundtouch_flush@4",
                ])?),
                clear: std::mem::transmute(proc(&[
                    "soundtouch_clear",
                    "_soundtouch_clear",
                    "soundtouch_clear@4",
                ])?),
            })
        })
        .as_ref()
}

struct SoundTouch {
    api: SoundTouchApi,
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
                api: SoundTouchApi {
                    _handle: api._handle,
                    create: api.create,
                    destroy: api.destroy,
                    set_sample_rate: api.set_sample_rate,
                    set_channels: api.set_channels,
                    set_tempo: api.set_tempo,
                    put_samples: api.put_samples,
                    receive_samples: api.receive_samples,
                    flush: api.flush,
                    clear: api.clear,
                },
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
        if self.index >= self.buffer.len() {
            if !self.refill() {
                return None;
            }
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
    let file = std::fs::File::open(path).ok()?;
    let source: Decoder<_> = Decoder::new(std::io::BufReader::new(file)).ok()?;
    if let Some(dur) = source.total_duration() {
        return Some(dur.as_secs());
    }
    mp3_duration::from_path(path).ok().map(|d| d.as_secs())
}

pub unsafe fn start_audiobook_playback(hwnd: HWND, path: &Path) {
    crate::reset_active_podcast_chapters_for_playback(hwnd);
    let path_buf = path.to_path_buf();

    let (bookmark_pos, speed, volume) = with_state(hwnd, |state| {
        let pos = state
            .bookmarks
            .files
            .get(&path_buf.to_string_lossy().to_string())
            .and_then(|list| list.last()) // Usa l'ultimo segnalibro per l'audio
            .map(|bm| bm.position)
            .unwrap_or(0);
        (
            pos,
            state.settings.audiobook_playback_speed,
            state.settings.audiobook_playback_volume,
        )
    })
    .unwrap_or((0, 1.0, 1.0));

    start_audiobook_at_with_speed(
        hwnd,
        path_buf,
        bookmark_pos as u64,
        speed,
        false,
        volume,
        false,
        volume,
    );
}

pub unsafe fn toggle_audiobook_pause(hwnd: HWND) {
    let start_action = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.is_paused {
                player.sink.play();
                player.is_paused = false;
                player.start_instant = std::time::Instant::now();
            } else {
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
    start_audiobook_at_with_speed(
        hwnd,
        path,
        current_pos,
        speed,
        paused,
        volume,
        muted,
        prev_volume,
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
    let _ = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            state.last_stopped_audiobook = Some(player.path.clone());
            player.sink.stop();
        }
    });
}

pub unsafe fn start_audiobook_at(hwnd: HWND, path: &Path, seconds: u64) {
    // Preserve current speed and volume settings
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
    start_audiobook_at_with_speed(
        hwnd,
        path_buf,
        seconds,
        speed,
        false,
        volume,
        muted,
        prev_volume,
    );
}

fn start_audiobook_at_with_speed(
    hwnd: HWND,
    path: PathBuf,
    seconds: u64,
    speed: f32,
    paused: bool,
    volume: f32,
    muted: bool,
    prev_volume: f32,
) {
    let effective_speed = if (speed - 1.0).abs() > f32::EPSILON && load_soundtouch_api().is_some() {
        speed
    } else {
        1.0
    };
    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path) {
            Ok(f) => f,
            Err(_) => return,
        };

        let base: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        let source: Box<dyn Source<Item = f32> + Send> = if seconds > 0 {
            Box::new(
                base.skip_duration(std::time::Duration::from_secs(seconds))
                    .convert_samples(),
            )
        } else {
            Box::new(base.convert_samples())
        };

        if (effective_speed - 1.0).abs() > f32::EPSILON {
            match SoundTouchSource::try_new(source, effective_speed) {
                Ok(st_source) => sink.append(st_source),
                Err(source) => sink.append(source),
            }
        } else {
            sink.append(source);
        }

        if muted {
            sink.set_volume(0.0);
        } else {
            sink.set_volume(volume);
        }
        if paused {
            sink.pause();
        }

        let player = AudiobookPlayer {
            path,
            sink: sink.clone(),
            _stream,
            is_paused: paused,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume,
            muted,
            prev_volume,
            speed: effective_speed,
        };

        let _ = unsafe {
            with_state(hwnd_main, |state| {
                state.active_audiobook = Some(player);
            })
        };
    });
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

    // Save volume to settings
    if let Some(volume) = new_volume {
        with_state(hwnd, |state| {
            state.settings.audiobook_playback_volume = volume;
            crate::settings::save_settings(state.settings.clone());
        });
    }
}

pub unsafe fn change_audiobook_speed(hwnd: HWND, delta: f32) -> Option<f32> {
    if load_soundtouch_api().is_none() {
        return None;
    }
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

    let (path, current, speed, paused, volume, muted, prev_volume) = match result {
        Some(v) => v,
        None => return None,
    };

    start_audiobook_at_with_speed(
        hwnd,
        path,
        current,
        speed,
        paused,
        volume,
        muted,
        prev_volume,
    );

    // Save speed to settings
    with_state(hwnd, |state| {
        state.settings.audiobook_playback_speed = speed;
        crate::settings::save_settings(state.settings.clone());
    });

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
    let _ = with_state(hwnd, |state| {
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
