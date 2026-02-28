// src/telemetry.rs
//
// Lightweight telemetry for hang diagnostics.
// Tracks user actions, maintains a ring buffer of logs, and provides context for Sentry.
// Designed to be low-overhead and privacy-respecting.

use std::collections::VecDeque;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, OnceLock};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

/// Maximum number of log lines to keep in memory
const MAX_LOG_LINES: usize = 200;

/// Maximum number of recent actions to track
const MAX_ACTIONS: usize = 20;

/// Minimum interval between hang reports (5 minutes)
const MIN_REPORT_INTERVAL_SECS: u64 = 300;

/// Global telemetry state
static TELEMETRY: OnceLock<Telemetry> = OnceLock::new();

/// Last action recorded (for quick access without lock)
static LAST_ACTION: OnceLock<Mutex<LastAction>> = OnceLock::new();

/// Last hang report timestamp (milliseconds from UNIX epoch)
static LAST_HANG_REPORT_MS: AtomicU64 = AtomicU64::new(0);

/// Last action that was reported (to detect if something changed)
static LAST_REPORTED_ACTION: OnceLock<Mutex<String>> = OnceLock::new();

/// Represents the last recorded action
struct LastAction {
    action: &'static str,
    details: String,
    timestamp: Instant,
}

impl Default for LastAction {
    fn default() -> Self {
        Self {
            action: "startup",
            details: String::new(),
            timestamp: Instant::now(),
        }
    }
}

/// Telemetry state with ring buffers
struct Telemetry {
    /// Ring buffer of recent log lines
    logs: Mutex<VecDeque<String>>,
    /// Ring buffer of recent actions
    actions: Mutex<VecDeque<ActionRecord>>,
    /// Application state
    app_state: Mutex<AppState>,
}

/// Record of a user action
struct ActionRecord {
    action: &'static str,
    details: String,
    timestamp_ms: u64,
}

/// Current application state (for hang context)
#[derive(Default)]
struct AppState {
    current_window: String,
    audio_playing: bool,
    updater_state: String,
    tts_active: bool,
}

impl Telemetry {
    fn new() -> Self {
        Self {
            logs: Mutex::new(VecDeque::with_capacity(MAX_LOG_LINES)),
            actions: Mutex::new(VecDeque::with_capacity(MAX_ACTIONS)),
            app_state: Mutex::new(AppState::default()),
        }
    }
}

/// Initialize telemetry. Call once at startup.
pub fn init() {
    TELEMETRY.get_or_init(Telemetry::new);
    LAST_ACTION.get_or_init(|| Mutex::new(LastAction::default()));
    LAST_REPORTED_ACTION.get_or_init(|| Mutex::new(String::new()));
}

/// Record a user action. Adds a breadcrumb and updates last_action.
/// This is the main entry point for tracking what the user is doing.
///
/// # Arguments
/// * `action` - Static string identifying the action type (e.g., "file_open", "audio_play")
/// * `details` - Dynamic details about the action (will be sanitized)
pub fn record_action(action: &'static str, details: impl std::fmt::Display) {
    let details_str = details.to_string();
    let sanitized = sanitize_details(&details_str);
    let now_ms = current_time_ms();

    // Update last action (quick atomic-ish update)
    if let Some(last) = LAST_ACTION.get()
        && let Ok(mut guard) = last.lock()
    {
        guard.action = action;
        guard.details = sanitized.clone();
        guard.timestamp = Instant::now();
    }

    // Add to actions ring buffer
    if let Some(telemetry) = TELEMETRY.get()
        && let Ok(mut actions) = telemetry.actions.lock()
    {
        if actions.len() >= MAX_ACTIONS {
            actions.pop_front();
        }
        actions.push_back(ActionRecord {
            action,
            details: sanitized.clone(),
            timestamp_ms: now_ms,
        });
    }

    // Add Sentry breadcrumb (if enabled)
    if crate::sentry_integration::is_enabled() {
        crate::sentry_integration::add_breadcrumb(action, &sanitized);
    }
}

/// Push a log line to the ring buffer.
/// Called from the logging system.
pub fn push_log_line(line: impl Into<String>) {
    if let Some(telemetry) = TELEMETRY.get()
        && let Ok(mut logs) = telemetry.logs.lock()
    {
        if logs.len() >= MAX_LOG_LINES {
            logs.pop_front();
        }
        logs.push_back(line.into());
    }
}

/// Update audio playback state.
pub fn set_audio_playing(playing: bool) {
    if let Some(telemetry) = TELEMETRY.get()
        && let Ok(mut state) = telemetry.app_state.lock()
    {
        state.audio_playing = playing;
    }
}

/// Update TTS state.
pub fn set_tts_active(active: bool) {
    if let Some(telemetry) = TELEMETRY.get()
        && let Ok(mut state) = telemetry.app_state.lock()
    {
        state.tts_active = active;
    }
}

/// Context snapshot for hang reports
pub struct HangContext {
    pub last_action: String,
    pub last_action_details: String,
    pub action_age_secs: u64,
    pub current_window: String,
    pub audio_playing: bool,
    pub tts_active: bool,
    pub updater_state: String,
    pub recent_actions: Vec<String>,
    pub recent_logs: String,
}

/// Take a snapshot of the current context for hang reporting.
pub fn snapshot_context() -> HangContext {
    let (last_action, last_action_details, action_age_secs) = if let Some(last) = LAST_ACTION.get()
    {
        if let Ok(guard) = last.lock() {
            (
                guard.action.to_string(),
                guard.details.clone(),
                guard.timestamp.elapsed().as_secs(),
            )
        } else {
            ("unknown".to_string(), String::new(), 0)
        }
    } else {
        ("unknown".to_string(), String::new(), 0)
    };

    let (current_window, audio_playing, tts_active, updater_state) =
        if let Some(telemetry) = TELEMETRY.get() {
            if let Ok(state) = telemetry.app_state.lock() {
                (
                    state.current_window.clone(),
                    state.audio_playing,
                    state.tts_active,
                    state.updater_state.clone(),
                )
            } else {
                (String::new(), false, false, String::new())
            }
        } else {
            (String::new(), false, false, String::new())
        };

    let recent_actions = if let Some(telemetry) = TELEMETRY.get() {
        if let Ok(actions) = telemetry.actions.lock() {
            actions
                .iter()
                .map(|a| format!("[{}] {}: {}", a.timestamp_ms, a.action, a.details))
                .collect()
        } else {
            Vec::new()
        }
    } else {
        Vec::new()
    };

    let recent_logs = if let Some(telemetry) = TELEMETRY.get() {
        if let Ok(logs) = telemetry.logs.lock() {
            // Join last 50 lines for Sentry (to avoid huge payloads)
            logs.iter()
                .rev()
                .take(50)
                .collect::<Vec<_>>()
                .into_iter()
                .rev()
                .cloned()
                .collect::<Vec<_>>()
                .join("\n")
        } else {
            String::new()
        }
    } else {
        String::new()
    };

    HangContext {
        last_action,
        last_action_details,
        action_age_secs,
        current_window,
        audio_playing,
        tts_active,
        updater_state,
        recent_actions,
        recent_logs,
    }
}

/// Check if we should send a hang report (rate limiting).
/// Returns true if we should send, false if rate limited.
pub fn should_report_hang() -> bool {
    let now_ms = current_time_ms();
    let last_ms = LAST_HANG_REPORT_MS.load(Ordering::Relaxed);
    let min_interval_ms = MIN_REPORT_INTERVAL_SECS * 1000;

    // Check time-based rate limit
    if now_ms.saturating_sub(last_ms) < min_interval_ms {
        // Check if action changed (allow report if different action)
        let current_action = if let Some(last) = LAST_ACTION.get() {
            if let Ok(guard) = last.lock() {
                format!("{}:{}", guard.action, guard.details)
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        let last_reported = if let Some(last) = LAST_REPORTED_ACTION.get() {
            if let Ok(guard) = last.lock() {
                guard.clone()
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        if current_action == last_reported {
            return false; // Same action, rate limited
        }
    }

    true
}

/// Mark that we sent a hang report (for rate limiting).
pub fn mark_hang_reported() {
    LAST_HANG_REPORT_MS.store(current_time_ms(), Ordering::Relaxed);

    // Store current action as last reported
    if let Some(last) = LAST_ACTION.get()
        && let Ok(guard) = last.lock()
    {
        let action_str = format!("{}:{}", guard.action, guard.details);
        if let Some(last_reported) = LAST_REPORTED_ACTION.get()
            && let Ok(mut reported) = last_reported.lock()
        {
            *reported = action_str;
        }
    }
}

/// Sanitize details to remove sensitive information (paths, etc.)
fn sanitize_details(details: &str) -> String {
    // Remove Windows paths
    let re = fancy_regex::Regex::new(r#"(?i)([A-Z]:\\[^\s"'<>|]*|\\\\[^\s"'<>|]+)"#).ok();
    match re {
        Some(regex) => regex.replace_all(details, "<path>").to_string(),
        None => details.to_string(),
    }
}

/// Get current time in milliseconds from UNIX epoch
fn current_time_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

// ============================================================================
// Convenience macros for recording actions
// ============================================================================

/// Record a file operation action.
#[macro_export]
macro_rules! telemetry_file {
    ($op:expr) => {
        $crate::telemetry::record_action(concat!("file_", $op), "")
    };
    ($op:expr, $details:expr) => {
        $crate::telemetry::record_action(concat!("file_", $op), $details)
    };
}

/// Record an audio operation action.
#[macro_export]
macro_rules! telemetry_audio {
    ($op:expr) => {
        $crate::telemetry::record_action(concat!("audio_", $op), "")
    };
    ($op:expr, $details:expr) => {
        $crate::telemetry::record_action(concat!("audio_", $op), $details)
    };
}

/// Record a UI action.
#[macro_export]
macro_rules! telemetry_ui {
    ($op:expr) => {
        $crate::telemetry::record_action(concat!("ui_", $op), "")
    };
    ($op:expr, $details:expr) => {
        $crate::telemetry::record_action(concat!("ui_", $op), $details)
    };
}
