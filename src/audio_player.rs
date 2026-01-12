use crate::accessibility::to_wide;
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
                "SoundTouch-6766862dc3e61fe695b186dfb40dc6b5.dll"
            } else {
                "SoundTouch-6766862dc3e61fe695b186dfb40dc6b5.dll"
            };
            let dll_path = settings_dir().join(dll_name);
            let dll_path_wide = to_wide(&dll_path.to_string_lossy());
            let h = LoadLibraryW(PCWSTR(dll_path_wide.as_ptr())).ok()?;
            let proc = |name: &str| {
                let cstr = std::ffi::CString::new(name).ok()?;
                let addr = GetProcAddress(h, PCSTR(cstr.as_ptr() as *const u8))?;
                Some(addr)
            };
            Some(SoundTouchApi {
                _handle: h,
                create: std::mem::transmute(proc("soundtouch_createInstance")?),
                destroy: std::mem::transmute(proc("soundtouch_destroyInstance")?),
                set_sample_rate: std::mem::transmute(proc("soundtouch_setSampleRate")?),
                set_channels: std::mem::transmute(proc("soundtouch_setChannels")?),
                set_tempo: std::mem::transmute(proc("soundtouch_setTempo")?),
                put_samples: std::mem::transmute(proc("soundtouch_putSamples")?),
                receive_samples: std::mem::transmute(proc("soundtouch_receiveSamples")?),
                flush: std::mem::transmute(proc("soundtouch_flush")?),
                clear: std::mem::transmute(proc("soundtouch_clear")?),
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
            let count = received as usize * channels;
            self.buffer.extend_from_slice(&out[..count]);
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
    let path_buf = path.to_path_buf();

    let bookmark_pos = with_state(hwnd, |state| {
        state
            .bookmarks
            .files
            .get(&path_buf.to_string_lossy().to_string())
            .and_then(|list| list.last()) // Usa l'ultimo segnalibro per l'audio
            .map(|bm| bm.position)
            .unwrap_or(0)
    })
    .unwrap_or(0);

    start_audiobook_at_with_speed(
        hwnd,
        path_buf,
        bookmark_pos as u64,
        1.0,
        false,
        1.0,
        false,
        1.0,
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
    stop_audiobook_playback(hwnd);
    let path_buf = path.to_path_buf();
    start_audiobook_at_with_speed(hwnd, path_buf, seconds, 1.0, false, 1.0, false, 1.0);
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
    let _ = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.muted {
                player.prev_volume = (player.prev_volume + delta).clamp(0.0, 3.0);
                return;
            }
            player.volume = (player.volume + delta).clamp(0.0, 3.0);
            player.sink.set_volume(player.volume);
        }
    });
}

pub unsafe fn change_audiobook_speed(hwnd: HWND, delta: f32) {
    if load_soundtouch_api().is_none() {
        return;
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
        None => return,
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
}

pub unsafe fn audiobook_speed_level(hwnd: HWND) -> Option<f32> {
    with_state(hwnd, |state| {
        state.active_audiobook.as_ref().map(|player| player.speed)
    })
    .flatten()
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
