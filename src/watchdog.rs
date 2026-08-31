// src/watchdog.rs
//
// Watchdog thread per rilevare freeze/hang dell'applicazione.
// Monitora un heartbeat che deve essere aggiornato dal message loop principale.

use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, OnceLock};
use std::time::{Duration, SystemTime, UNIX_EPOCH};

/// Handle globale del watchdog per accesso da qualsiasi punto dell'app
static GLOBAL_WATCHDOG: OnceLock<Arc<WatchdogState>> = OnceLock::new();

/// Chiamare quando si apre un dialog modale per sospendere il rilevamento hang.
pub fn enter_modal_dialog() {
    if let Some(state) = GLOBAL_WATCHDOG.get() {
        state.enter_modal();
    }
}

/// Chiamare quando si chiude un dialog modale per riprendere il rilevamento hang.
pub fn exit_modal_dialog() {
    if let Some(state) = GLOBAL_WATCHDOG.get() {
        state.exit_modal();
    }
}

/// Aggiorna il watchdog da un message loop UI annidato o da una finestra di
/// avanzamento che sta elaborando messaggi. Se il watchdog non è ancora
/// inizializzato, non fa nulla.
pub fn heartbeat() {
    if let Some(state) = GLOBAL_WATCHDOG.get() {
        state.heartbeat();
    }
}

/// Stato condiviso del watchdog
pub struct WatchdogState {
    /// Timestamp dell'ultimo heartbeat (millisecondi da UNIX epoch)
    last_heartbeat_ms: AtomicU64,
    /// Flag per fermare il watchdog
    should_stop: AtomicBool,
    /// Flag che indica se un hang è già stato riportato (evita spam)
    hang_reported: AtomicBool,
    /// Contatore di dialogs modali aperti (quando > 0, il watchdog è sospeso)
    modal_dialog_count: AtomicU64,
}

impl WatchdogState {
    fn new() -> Self {
        Self {
            last_heartbeat_ms: AtomicU64::new(current_time_ms()),
            should_stop: AtomicBool::new(false),
            hang_reported: AtomicBool::new(false),
            modal_dialog_count: AtomicU64::new(0),
        }
    }

    /// Aggiorna il heartbeat. Chiamare dal message loop principale.
    pub fn heartbeat(&self) {
        self.last_heartbeat_ms
            .store(current_time_ms(), Ordering::Relaxed);
        // Reset hang flag quando l'app risponde
        self.hang_reported.store(false, Ordering::Relaxed);
    }

    /// Segnala al watchdog di fermarsi.
    pub fn stop(&self) {
        self.should_stop.store(true, Ordering::Relaxed);
    }

    /// Incrementa il contatore di dialogs modali (sospende il watchdog)
    pub fn enter_modal(&self) {
        self.modal_dialog_count.fetch_add(1, Ordering::Relaxed);
    }

    /// Decrementa il contatore di dialogs modali e aggiorna heartbeat
    pub fn exit_modal(&self) {
        self.modal_dialog_count.fetch_sub(1, Ordering::Relaxed);
        // Aggiorna heartbeat all'uscita dal dialog per evitare falsi positivi
        self.heartbeat();
    }

    fn should_stop(&self) -> bool {
        self.should_stop.load(Ordering::Relaxed)
    }

    fn last_heartbeat_ms(&self) -> u64 {
        self.last_heartbeat_ms.load(Ordering::Relaxed)
    }

    fn is_in_modal(&self) -> bool {
        self.modal_dialog_count.load(Ordering::Relaxed) > 0
    }

    fn mark_hang_reported(&self) -> bool {
        // Ritorna true se era già reported (per evitare duplicati)
        self.hang_reported.swap(true, Ordering::Relaxed)
    }
}

/// Handle per controllare il watchdog
pub struct WatchdogHandle {
    state: Arc<WatchdogState>,
    _thread: Option<std::thread::JoinHandle<()>>,
}

impl WatchdogHandle {
    /// Aggiorna il heartbeat. Chiamare periodicamente dal message loop.
    pub fn heartbeat(&self) {
        self.state.heartbeat();
    }

    /// Ferma il watchdog gracefully.
    pub fn stop(&self) {
        self.state.stop();
    }
}

impl Drop for WatchdogHandle {
    fn drop(&mut self) {
        self.state.stop();
        // Non facciamo join per evitare blocchi
    }
}

/// Configurazione del watchdog
pub struct WatchdogConfig {
    /// Intervallo di check in secondi
    pub check_interval_secs: u64,
    /// Soglia di hang in secondi (se heartbeat > questa soglia, è hang)
    pub hang_threshold_secs: u64,
    /// Se true, logga ogni heartbeat check (verbose)
    pub log_heartbeats: bool,
    /// Se true, invia eventi Sentry per hang
    pub report_to_sentry: bool,
}

impl Default for WatchdogConfig {
    fn default() -> Self {
        Self {
            check_interval_secs: 5,
            hang_threshold_secs: 30,
            log_heartbeats: false,
            report_to_sentry: true,
        }
    }
}

/// Avvia il watchdog thread.
/// Ritorna un handle per controllare il watchdog e aggiornare l'heartbeat.
pub fn start_watchdog(config: WatchdogConfig) -> WatchdogHandle {
    let state = Arc::new(WatchdogState::new());
    let state_clone = Arc::clone(&state);

    // Salva lo state globalmente per accesso dai dialogs modali
    // Ignoriamo l'errore se già impostato (può succedere solo in test)
    if GLOBAL_WATCHDOG.set(Arc::clone(&state)).is_err() {
        crate::log_debug("Watchdog: global state already set");
    }

    let thread = std::thread::Builder::new()
        .name("watchdog".to_string())
        .spawn(move || {
            watchdog_loop(state_clone, config);
        });

    let thread_handle = match thread {
        Ok(handle) => Some(handle),
        Err(e) => {
            crate::log_debug(&format!("Watchdog: failed to start thread: {}", e));
            None
        }
    };

    WatchdogHandle {
        state,
        _thread: thread_handle,
    }
}

/// Loop principale del watchdog
fn watchdog_loop(state: Arc<WatchdogState>, config: WatchdogConfig) {
    let check_interval = Duration::from_secs(config.check_interval_secs);
    let hang_threshold_ms = config.hang_threshold_secs * 1000;

    crate::log_debug(&format!(
        "Watchdog: started (check every {}s, hang threshold {}s)",
        config.check_interval_secs, config.hang_threshold_secs
    ));

    loop {
        std::thread::sleep(check_interval);

        if state.should_stop() {
            crate::log_debug("Watchdog: stopping");
            break;
        }

        let now_ms = current_time_ms();
        let last_heartbeat = state.last_heartbeat_ms();
        let elapsed_ms = now_ms.saturating_sub(last_heartbeat);

        if config.log_heartbeats {
            crate::log_debug(&format!("Watchdog: heartbeat age = {}ms", elapsed_ms));
        }

        // Rileva hang (ma ignora se siamo in un dialog modale)
        if elapsed_ms > hang_threshold_ms && !state.is_in_modal() {
            // Evita di riportare lo stesso hang più volte
            if state.mark_hang_reported() {
                continue;
            }

            let elapsed_secs = elapsed_ms / 1000;
            let message = format!(
                "Possible application hang detected: no heartbeat for {} seconds",
                elapsed_secs
            );

            crate::log_debug(&format!("Watchdog: {}", message));

            // Salva su file per diagnostica
            save_hang_report(elapsed_secs);

            // Invia a Sentry se abilitato
            if config.report_to_sentry && crate::sentry_integration::is_enabled() {
                report_hang_to_sentry(elapsed_secs);
            }
        }
    }

    crate::log_debug("Watchdog: stopped");
}

/// Riporta hang a Sentry con contesto ricco
fn report_hang_to_sentry(elapsed_secs: u64) {
    use crate::sentry_integration;
    use crate::telemetry;

    // Rate limiting: evita spam
    if !telemetry::should_report_hang() {
        crate::log_debug("Watchdog: hang report rate-limited");
        return;
    }

    // Prendi snapshot del contesto
    let ctx = telemetry::snapshot_context();

    sentry::with_scope(
        |scope| {
            // Tags per filtraggio
            scope.set_tag("err.type", "hang");
            scope.set_tag("err.severity", "warning");
            scope.set_tag("hang.last_action", &ctx.last_action);

            if !ctx.current_window.is_empty() {
                scope.set_tag("hang.window", &ctx.current_window);
            }

            // Extra data
            scope.set_extra(
                "hang.duration_secs",
                sentry::protocol::Value::from(elapsed_secs),
            );
            scope.set_extra(
                "hang.last_action_age_secs",
                sentry::protocol::Value::from(ctx.action_age_secs),
            );
            scope.set_extra(
                "hang.last_action_details",
                sentry::protocol::Value::from(ctx.last_action_details.clone()),
            );
            scope.set_extra(
                "hang.audio_playing",
                sentry::protocol::Value::from(ctx.audio_playing),
            );
            scope.set_extra(
                "hang.tts_active",
                sentry::protocol::Value::from(ctx.tts_active),
            );

            if !ctx.updater_state.is_empty() {
                scope.set_extra(
                    "hang.updater_state",
                    sentry::protocol::Value::from(ctx.updater_state.clone()),
                );
            }

            // Recent actions as a list
            if !ctx.recent_actions.is_empty() {
                scope.set_extra(
                    "hang.recent_actions",
                    sentry::protocol::Value::from(ctx.recent_actions.join("\n")),
                );
            }

            // Recent logs (truncated to avoid huge payloads)
            if !ctx.recent_logs.is_empty() {
                let logs_truncated = if ctx.recent_logs.chars().count() > 8000 {
                    format!("...(truncated)...\n{}", tail_chars(&ctx.recent_logs, 8000))
                } else {
                    ctx.recent_logs.clone()
                };
                scope.set_extra(
                    "hang.recent_logs",
                    sentry::protocol::Value::from(logs_truncated),
                );
            }
        },
        || {
            let message = format!(
                "Possible hang: UI unresponsive for {} seconds (last action: {})",
                elapsed_secs, ctx.last_action
            );
            let event_id = sentry::capture_message(&message, sentry::Level::Warning);

            // Salva l'event ID per diagnostica
            sentry_integration::save_last_fatal(&message, Some(&event_id));
        },
    );

    // Marca come riportato (per rate limiting)
    telemetry::mark_hang_reported();

    // Flush per assicurarsi che l'evento sia inviato
    if let Some(client) = sentry::Hub::current().client() {
        client.flush(Some(Duration::from_secs(2)));
    }
}

fn tail_chars(text: &str, max_chars: usize) -> String {
    if max_chars == 0 {
        return String::new();
    }
    let total = text.chars().count();
    if total <= max_chars {
        return text.to_string();
    }
    text.chars().skip(total - max_chars).collect()
}

/// Salva report di hang su file
fn save_hang_report(elapsed_secs: u64) {
    let settings_dir = crate::settings::settings_dir();
    let path = settings_dir.join("last_hang.txt");

    let timestamp = chrono::Local::now().format("%Y-%m-%d %H:%M:%S");
    let content = format!(
        "Timestamp: {}\n\
         Type: Possible hang\n\
         Duration: {} seconds without heartbeat\n\
         Note: This may indicate the application was busy or frozen\n",
        timestamp, elapsed_secs
    );

    if let Err(e) = std::fs::write(&path, content) {
        crate::log_debug(&format!("Watchdog: failed to save hang report: {}", e));
    }
}

/// Ottiene il tempo corrente in millisecondi da UNIX epoch
fn current_time_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}
