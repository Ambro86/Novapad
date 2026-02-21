use crate::accessibility::to_wide;
use crate::bass_ffmpeg_stream::FfmpegBassStream;
use crate::bass_sys::{
    BASS_ATTRIB_TEMPO, BASS_ATTRIB_TEMPO_PITCH, BASS_ATTRIB_VOL, BASS_FX_FREESOURCE, BASS_POS_BYTE,
    BASS_SAMPLE_FLOAT, BASS_STREAM_DECODE, BASS_STREAM_PRESCAN, BASS_UNICODE, BassApi, Dword,
    Hplugin, Hstream, bass_api, bass_error, bass_fx_api, log_bass_error,
};
use crate::embedded_deps;
use crate::log_debug;
use std::ffi::c_void;
use std::path::Path;
use std::ptr;
use std::sync::{Arc, Mutex, OnceLock};

pub struct BassOutput {
    api: &'static BassApi,
    handle: Mutex<Hstream>,
    /// Keep FFmpeg stream alive while playing
    _ffmpeg_stream: Option<FfmpegBassStream>,
    /// Offset in seconds for FFmpeg streams (seek position)
    start_offset_secs: f64,
}

static BASS_INIT: OnceLock<Result<(), String>> = OnceLock::new();
static BASS_PLUGINS: OnceLock<Result<Vec<Hplugin>, String>> = OnceLock::new();

fn init_bass_once() -> Result<(), String> {
    BASS_INIT
        .get_or_init(|| {
            let api = bass_api()?;
            let init_ok = unsafe { (api.init)(-1, 44_100, 0, ptr::null_mut(), ptr::null()) };
            if init_ok == 0 {
                return Err(format!("BASS_Init failed (error {})", bass_error(api)));
            }
            Ok(())
        })
        .clone()
}

fn load_plugins_once(api: &BassApi) -> Result<Vec<Hplugin>, String> {
    BASS_PLUGINS
        .get_or_init(|| {
            let plugin_names = [
                "bass_aac.dll",
                "bass_alac.dll",
                "bassflac.dll",
                "bassopus.dll",
                "basswma.dll",
            ];
            let mut handles = Vec::new();
            for name in plugin_names {
                let path = embedded_deps::get_dep_path(name);
                if !path.exists() {
                    log_debug(&format!("BASS: plugin missing: {}", path.display()));
                    continue;
                }
                let wide = to_wide(path.to_string_lossy().as_ref());
                let handle =
                    unsafe { (api.plugin_load)(wide.as_ptr() as *const c_void, BASS_UNICODE) };
                if handle == 0 {
                    log_bass_error(api, &format!("PluginLoad {}", name));
                } else {
                    handles.push(handle);
                }
            }
            Ok(handles)
        })
        .clone()
}

fn create_stream_from_path(api: &BassApi, path: &Path, flags: Dword) -> Result<Hstream, String> {
    let wide = to_wide(path.to_string_lossy().as_ref());
    let handle = unsafe {
        (api.stream_create_file)(
            0,
            wide.as_ptr() as *const c_void,
            0,
            0,
            flags | BASS_UNICODE,
        )
    };
    if handle == 0 {
        return Err(format!(
            "BASS_StreamCreateFile failed (error {})",
            bass_error(api)
        ));
    }
    Ok(handle)
}

impl BassOutput {
    pub fn start(
        path: &Path,
        start_seconds: u64,
        speed: f32,
        pitch: f32,
        volume: f32,
        paused: bool,
    ) -> Result<Arc<Self>, String> {
        init_bass_once()?;
        let api = bass_api()?;
        let fx_api = bass_fx_api().ok();
        if let Err(err) = load_plugins_once(api) {
            log_debug(&format!("BASS: plugin load failed: {}", err));
        }

        let want_tempo = (speed != 1.0 || pitch != 0.0) && fx_api.is_some();
        let mut flags = BASS_STREAM_PRESCAN | BASS_SAMPLE_FLOAT;
        if want_tempo {
            flags |= BASS_STREAM_DECODE;
        }

        let mut source = create_stream_from_path(api, path, flags)?;
        let handle = if want_tempo {
            if let Some(fx_api) = fx_api {
                let tempo_handle = unsafe {
                    (fx_api.tempo_create)(source, BASS_FX_FREESOURCE | BASS_SAMPLE_FLOAT)
                };
                if tempo_handle == 0 {
                    log_bass_error(api, "BASS_FX_TempoCreate");
                    let free_ok = unsafe { (api.stream_free)(source) };
                    if free_ok == 0 {
                        log_bass_error(api, "BASS_StreamFree");
                    }
                    source = create_stream_from_path(api, path, BASS_STREAM_PRESCAN)?;
                    source
                } else {
                    let tempo = ((speed as f64 - 1.0) * 100.0) as f32;
                    let tempo = tempo.clamp(-95.0, 5000.0);
                    let set_ok = unsafe {
                        (api.channel_set_attribute)(tempo_handle, BASS_ATTRIB_TEMPO, tempo)
                    };
                    if set_ok == 0 {
                        log_bass_error(api, "BASS_ChannelSetAttribute tempo");
                    }
                    let pitch_clamped = pitch.clamp(-60.0, 60.0);
                    let set_pitch_ok = unsafe {
                        (api.channel_set_attribute)(
                            tempo_handle,
                            BASS_ATTRIB_TEMPO_PITCH,
                            pitch_clamped,
                        )
                    };
                    if set_pitch_ok == 0 {
                        log_bass_error(api, "BASS_ChannelSetAttribute pitch");
                    }
                    tempo_handle
                }
            } else {
                source
            }
        } else {
            if speed != 1.0 {
                log_debug("BASS: bass_fx unavailable, forcing speed=1.0.");
            }
            source
        };

        let volume = volume.clamp(0.0, 6.0);
        let set_ok = unsafe { (api.channel_set_attribute)(handle, BASS_ATTRIB_VOL, volume) };
        if set_ok == 0 {
            log_bass_error(api, "BASS_ChannelSetAttribute volume");
        }

        if start_seconds > 0 {
            let pos = unsafe { (api.channel_seconds2bytes)(handle, start_seconds as f64) };
            let seek_ok = unsafe { (api.channel_set_position)(handle, pos, BASS_POS_BYTE) };
            if seek_ok == 0 {
                log_bass_error(api, "BASS_ChannelSetPosition");
            }
        }

        if !paused {
            let play_ok = unsafe { (api.channel_play)(handle, 0) };
            if play_ok == 0 {
                log_bass_error(api, "BASS_ChannelPlay");
            }
        }

        Ok(Arc::new(Self {
            api,
            handle: Mutex::new(handle),
            _ffmpeg_stream: None,
            start_offset_secs: 0.0,
        }))
    }

    /// Start playback using FFmpeg streaming (no intermediate WAV file)
    pub fn start_with_ffmpeg(
        path: &Path,
        start_seconds: u64,
        speed: f32,
        pitch: f32,
        volume: f32,
        paused: bool,
        stream_index: Option<i32>,
    ) -> Result<Arc<Self>, String> {
        init_bass_once()?;
        let api = bass_api()?;
        let fx_api = bass_fx_api().ok();

        // Create FFmpeg stream with BASS callback
        let (ffmpeg_stream, source_handle) =
            FfmpegBassStream::new(path, start_seconds, stream_index)?;

        let want_tempo = (speed != 1.0 || pitch != 0.0) && fx_api.is_some();
        let handle = if want_tempo {
            if let Some(fx_api) = fx_api {
                let tempo_handle = unsafe {
                    (fx_api.tempo_create)(source_handle, BASS_FX_FREESOURCE | BASS_SAMPLE_FLOAT)
                };
                if tempo_handle == 0 {
                    log_bass_error(api, "BASS_FX_TempoCreate (ffmpeg)");
                    source_handle
                } else {
                    let tempo = ((speed as f64 - 1.0) * 100.0) as f32;
                    let tempo = tempo.clamp(-95.0, 5000.0);
                    let set_ok = unsafe {
                        (api.channel_set_attribute)(tempo_handle, BASS_ATTRIB_TEMPO, tempo)
                    };
                    if set_ok == 0 {
                        log_bass_error(api, "BASS_ChannelSetAttribute tempo (ffmpeg)");
                    }
                    let pitch_clamped = pitch.clamp(-60.0, 60.0);
                    let set_pitch_ok = unsafe {
                        (api.channel_set_attribute)(
                            tempo_handle,
                            BASS_ATTRIB_TEMPO_PITCH,
                            pitch_clamped,
                        )
                    };
                    if set_pitch_ok == 0 {
                        log_bass_error(api, "BASS_ChannelSetAttribute pitch (ffmpeg)");
                    }
                    tempo_handle
                }
            } else {
                source_handle
            }
        } else {
            if speed != 1.0 {
                log_debug("BASS: bass_fx unavailable, forcing speed=1.0.");
            }
            source_handle
        };

        let volume = volume.clamp(0.0, 6.0);
        let set_ok = unsafe { (api.channel_set_attribute)(handle, BASS_ATTRIB_VOL, volume) };
        if set_ok == 0 {
            log_bass_error(api, "BASS_ChannelSetAttribute volume (ffmpeg)");
        }

        if !paused {
            let play_ok = unsafe { (api.channel_play)(handle, 0) };
            if play_ok == 0 {
                log_bass_error(api, "BASS_ChannelPlay (ffmpeg)");
            }
        }

        log_debug("BASS: FFmpeg streaming started");
        Ok(Arc::new(Self {
            api,
            handle: Mutex::new(handle),
            _ffmpeg_stream: Some(ffmpeg_stream),
            start_offset_secs: start_seconds as f64,
        }))
    }

    pub fn play(&self) -> bool {
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let ok = unsafe { (self.api.channel_play)(handle, 0) };
        if ok == 0 {
            log_bass_error(self.api, "BASS_ChannelPlay");
            return false;
        }
        true
    }

    pub fn pause(&self) -> bool {
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let ok = unsafe { (self.api.channel_pause)(handle) };
        if ok == 0 {
            log_bass_error(self.api, "BASS_ChannelPause");
            return false;
        }
        true
    }

    pub fn stop(&self) {
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let ok = unsafe { (self.api.channel_stop)(handle) };
        if ok == 0 {
            log_bass_error(self.api, "BASS_ChannelStop");
        }
        let free_ok = unsafe { (self.api.stream_free)(handle) };
        if free_ok == 0 {
            log_bass_error(self.api, "BASS_StreamFree");
        }
    }

    pub fn set_volume(&self, volume: f32) {
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let volume = volume.clamp(0.0, 6.0);
        let ok = unsafe { (self.api.channel_set_attribute)(handle, BASS_ATTRIB_VOL, volume) };
        if ok == 0 {
            log_bass_error(self.api, "BASS_ChannelSetAttribute volume");
        }
    }

    pub fn position_secs(&self) -> Option<f64> {
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let pos = unsafe { (self.api.channel_get_position)(handle, BASS_POS_BYTE) };
        if pos == 0 {
            let err = bass_error(self.api);
            if err != 0 {
                log_debug(&format!("BASS: position failed (error {})", err));
                return None;
            }
        }
        let bass_pos = unsafe { (self.api.channel_bytes2seconds)(handle, pos) }.max(0.0);
        // Add start offset for FFmpeg streams that were seeked
        Some(bass_pos + self.start_offset_secs)
    }

    pub fn seek_to_seconds(&self, absolute_seconds: f64) -> bool {
        // FFmpeg streaming starts from `start_offset_secs`. If caller asks to seek
        // before this offset, we must fail here so the caller can reopen at the
        // real absolute position.
        if self.start_offset_secs > 0.0 && absolute_seconds < self.start_offset_secs {
            return false;
        }
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let relative = (absolute_seconds - self.start_offset_secs).max(0.0);
        let pos = unsafe { (self.api.channel_seconds2bytes)(handle, relative) };
        let ok = unsafe { (self.api.channel_set_position)(handle, pos, BASS_POS_BYTE) };
        if ok == 0 {
            log_bass_error(self.api, "BASS_ChannelSetPosition");
            return false;
        }
        true
    }

    pub fn is_stopped(&self) -> bool {
        const BASS_ACTIVE_STOPPED: Dword = 0;
        let handle = *self.handle.lock().unwrap_or_else(|e| e.into_inner());
        let state = unsafe { (self.api.channel_is_active)(handle) };
        state == BASS_ACTIVE_STOPPED
    }

    pub fn clear_subtitles(&self) {
        // Edge subtitle audio is played through the TTS engine path (rodio).
    }
}
