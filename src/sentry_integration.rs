// src/sentry_integration.rs
//
// Integrazione Sentry per crash reporting opt-in.
// Non invia dati sensibili (path sanitizzati).

// Alcune funzioni sono API pubbliche per uso futuro
#![allow(dead_code)]

use std::panic::PanicHookInfo;
use std::sync::OnceLock;

// DSN di default per le release ufficiali (non è PII).
// Nota: evita di stamparlo nei log.
const DEFAULT_SENTRY_DSN: &str = "https://1ca59fd9ca0412302237d0c576eccec5@o4510784617316352.ingest.de.sentry.io/4510784634159184";

/// ID univoco dell'evento Sentry (re-export di Uuid)
pub type EventId = uuid::Uuid;

static SENTRY_GUARD: OnceLock<sentry::ClientInitGuard> = OnceLock::new();

// Regex per path Windows (compilata una volta sola)
static PATH_REGEX: OnceLock<Option<fancy_regex::Regex>> = OnceLock::new();

fn get_path_regex() -> Option<&'static fancy_regex::Regex> {
    PATH_REGEX
        .get_or_init(|| {
            // Match Windows paths: C:\..., D:\..., \\server\..., etc.
            fancy_regex::Regex::new(r#"(?i)([A-Z]:\\[^\s"'<>|]*|\\\\[^\s"'<>|]+)"#).ok()
        })
        .as_ref()
}

/// Rimuove informazioni sensibili dal messaggio (path, etc.)
fn sanitize_message(message: &str) -> String {
    match get_path_regex() {
        Some(re) => re.replace_all(message, "<path>").to_string(),
        None => message.to_string(),
    }
}

/// Inizializza Sentry se abilitato nelle impostazioni.
/// Deve essere chiamato una sola volta all'avvio, prima di qualsiasi altra operazione.
/// Il guard viene conservato per tutta la durata dell'applicazione.
pub fn init(send_crash_reports: bool, dsn: Option<&str>) {
    if !send_crash_reports {
        crate::log_debug("Sentry: crash reporting disabled by user settings");
        return;
    }

    let dsn = dsn
        .filter(|s| !s.trim().is_empty())
        .unwrap_or(DEFAULT_SENTRY_DSN);

    let guard = sentry::init((
        dsn,
        sentry::ClientOptions {
            release: Some(env!("CARGO_PKG_VERSION").into()),
            environment: Some(
                if cfg!(debug_assertions) {
                    "development"
                } else {
                    "production"
                }
                .into(),
            ),
            sample_rate: 0.5,
            attach_stacktrace: true,
            default_integrations: true,
            before_send: Some(std::sync::Arc::new(|mut event| {
                // Sanitizza il messaggio principale
                if let Some(ref message) = event.message {
                    event.message = Some(sanitize_message(message));
                }
                // Sanitizza le eccezioni
                for exc in event.exception.values.iter_mut() {
                    if let Some(ref value) = exc.value {
                        exc.value = Some(sanitize_message(value));
                    }
                }
                Some(event)
            })),
            ..Default::default()
        },
    ));

    if guard.is_enabled() {
        if SENTRY_GUARD.set(guard).is_err() {
            crate::log_debug("Sentry: already initialized");
        } else {
            // Configura tag globali (zero PII)
            configure_global_tags();
            crate::log_debug("Sentry: initialized successfully");
        }
    } else {
        crate::log_debug("Sentry: initialization failed (guard not enabled)");
    }
}

/// Configura tag e extra globali sullo scope Sentry.
/// Questi tag vengono aggiunti a tutti gli eventi.
/// Non contiene PII.
fn configure_global_tags() {
    sentry::configure_scope(|scope| {
        // App mode: standalone vs normal
        #[cfg(feature = "standalone")]
        scope.set_tag("app.mode", "standalone");
        #[cfg(not(feature = "standalone"))]
        scope.set_tag("app.mode", "normal");

        // App version
        scope.set_tag("app.version", env!("CARGO_PKG_VERSION"));

        // Build type
        scope.set_tag(
            "build.debug",
            if cfg!(debug_assertions) {
                "true"
            } else {
                "false"
            },
        );

        // Feature flags (compile-time)
        scope.set_tag("feature.updater", "true"); // Updater sempre compilato

        // Audio backend
        scope.set_tag("audio.backend", "bass");

        // Build target info (no PII)
        scope.set_extra(
            "build.target",
            sentry::protocol::Value::from(std::env::consts::ARCH),
        );
        scope.set_extra(
            "build.rustc",
            sentry::protocol::Value::from(built_info::rustc_version()),
        );
        scope.set_extra(
            "build.os",
            sentry::protocol::Value::from(std::env::consts::OS),
        );
    });
}

/// Informazioni di build (senza PII)
mod built_info {
    /// Versione del compilatore Rust usato per la build.
    /// Fallback a "unknown" se non disponibile.
    pub fn rustc_version() -> &'static str {
        option_env!("RUSTC_VERSION").unwrap_or("unknown")
    }
}

/// Cattura un errore fatale come stringa e lo invia a Sentry.
/// Usare per errori critici che non sono panic.
/// Ritorna l'EventId se l'evento è stato inviato.
pub fn capture_fatal_string(context: &str, message: &str) -> Option<EventId> {
    if !is_enabled() {
        return None;
    }

    let sanitized_context = sanitize_message(context);
    let sanitized_message = sanitize_message(message);
    let full_message = format!("{}: {}", sanitized_context, sanitized_message);

    sentry::with_scope(
        |scope| {
            scope.set_tag("err.severity", "fatal");
            scope.set_tag("err.context", sanitized_context);
        },
        || Some(sentry::capture_message(&full_message, sentry::Level::Error)),
    )
}

/// Cattura un errore non fatale.
/// Ritorna l'EventId se l'evento è stato inviato.
pub fn capture_error_string(context: &str, message: &str) -> Option<EventId> {
    if !is_enabled() {
        return None;
    }

    let sanitized_context = sanitize_message(context);
    let sanitized_message = sanitize_message(message);
    let full_message = format!("{}: {}", sanitized_context, sanitized_message);

    sentry::with_scope(
        |scope| {
            scope.set_tag("err.context", sanitized_context);
        },
        || Some(sentry::capture_message(&full_message, sentry::Level::Error)),
    )
}

/// Cattura un errore Windows con contesto.
/// Invia HRESULT code e messaggio come extra/tags.
/// Ritorna l'EventId se l'evento è stato inviato.
pub fn capture_windows_error(context: &str, err: &windows::core::Error) -> Option<EventId> {
    if !is_enabled() {
        return None;
    }

    let sanitized_context = sanitize_message(context);
    let error_code = err.code().0;
    let error_message = sanitize_message(&err.message());

    sentry::with_scope(
        |scope| {
            scope.set_tag("err.type", "windows");
            scope.set_tag("err.context", &sanitized_context);
            scope.set_tag("err.hresult", format!("0x{:08X}", error_code as u32));
            scope.set_extra("windows.hresult_decimal", error_code.into());
            scope.set_extra(
                "windows.error_message",
                sentry::protocol::Value::from(error_message.clone()),
            );
        },
        || {
            let message = format!("{}: HRESULT 0x{:08X}", sanitized_context, error_code as u32);
            Some(sentry::capture_message(&message, sentry::Level::Error))
        },
    )
}

/// Cattura un errore Windows fatale con contesto.
/// Ritorna l'EventId se l'evento è stato inviato.
pub fn capture_fatal_windows_error(context: &str, err: &windows::core::Error) -> Option<EventId> {
    if !is_enabled() {
        return None;
    }

    let sanitized_context = sanitize_message(context);
    let error_code = err.code().0;
    let error_message = sanitize_message(&err.message());

    sentry::with_scope(
        |scope| {
            scope.set_tag("err.severity", "fatal");
            scope.set_tag("err.type", "windows");
            scope.set_tag("err.context", &sanitized_context);
            scope.set_tag("err.hresult", format!("0x{:08X}", error_code as u32));
            scope.set_extra("windows.hresult_decimal", error_code.into());
            scope.set_extra(
                "windows.error_message",
                sentry::protocol::Value::from(error_message.clone()),
            );
        },
        || {
            let message = format!("{}: HRESULT 0x{:08X}", sanitized_context, error_code as u32);
            Some(sentry::capture_message(&message, sentry::Level::Error))
        },
    )
}

/// Aggiunge contesto utente anonimo (opzionale, per debugging).
/// Non inviare mai dati identificabili reali.
pub fn set_user_context(anonymous_id: Option<&str>) {
    if !is_enabled() {
        return;
    }

    sentry::configure_scope(|scope| {
        if let Some(id) = anonymous_id {
            scope.set_user(Some(sentry::User {
                id: Some(id.to_string()),
                ..Default::default()
            }));
        } else {
            scope.set_user(None);
        }
    });
}

/// Aggiunge un breadcrumb per tracciare il flusso dell'applicazione.
/// I messaggi vengono sanitizzati automaticamente.
pub fn add_breadcrumb(category: &str, message: &str) {
    add_breadcrumb_with_level(category, message, BreadcrumbLevel::Info);
}

/// Livelli di breadcrumb supportati
#[derive(Clone, Copy)]
pub enum BreadcrumbLevel {
    Debug,
    Info,
    Warning,
    Error,
}

impl From<BreadcrumbLevel> for sentry::Level {
    fn from(level: BreadcrumbLevel) -> Self {
        match level {
            BreadcrumbLevel::Debug => sentry::Level::Debug,
            BreadcrumbLevel::Info => sentry::Level::Info,
            BreadcrumbLevel::Warning => sentry::Level::Warning,
            BreadcrumbLevel::Error => sentry::Level::Error,
        }
    }
}

/// Aggiunge un breadcrumb con livello specifico.
/// I messaggi vengono sanitizzati automaticamente.
/// Non alloca se Sentry non è abilitato.
pub fn add_breadcrumb_with_level(category: &str, message: &str, level: BreadcrumbLevel) {
    if !is_enabled() {
        return;
    }

    let sanitized = sanitize_message(message);

    sentry::add_breadcrumb(sentry::Breadcrumb {
        category: Some(category.to_string()),
        message: Some(sanitized),
        level: level.into(),
        ..Default::default()
    });
}

/// Macro per aggiungere breadcrumb in modo leggero.
/// Non alloca se Sentry non è abilitato.
///
/// # Forme supportate:
/// ```ignore
/// breadcrumb!("category", "message");
/// breadcrumb!("category", "format {} string", arg);
/// breadcrumb!("category", level: Warning, "message");
/// breadcrumb!("category", level: Error, "format {}", arg);
/// ```
#[macro_export]
macro_rules! breadcrumb {
    // Forma semplice: breadcrumb!("category", "message")
    ($category:expr, $message:expr) => {
        if $crate::sentry_integration::is_enabled() {
            $crate::sentry_integration::add_breadcrumb($category, $message);
        }
    };
    // Forma con format: breadcrumb!("category", "format {}", arg)
    ($category:expr, $($arg:tt)+) => {
        if $crate::sentry_integration::is_enabled() {
            $crate::sentry_integration::add_breadcrumb($category, &format!($($arg)+));
        }
    };
}

/// Macro per breadcrumb con livello specifico.
///
/// # Forme supportate:
/// ```ignore
/// breadcrumb_level!(Debug, "category", "message");
/// breadcrumb_level!(Warning, "category", "format {}", arg);
/// ```
#[macro_export]
macro_rules! breadcrumb_level {
    ($level:ident, $category:expr, $message:expr) => {
        if $crate::sentry_integration::is_enabled() {
            $crate::sentry_integration::add_breadcrumb_with_level(
                $category,
                $message,
                $crate::sentry_integration::BreadcrumbLevel::$level,
            );
        }
    };
    ($level:ident, $category:expr, $($arg:tt)+) => {
        if $crate::sentry_integration::is_enabled() {
            $crate::sentry_integration::add_breadcrumb_with_level(
                $category,
                &format!($($arg)+),
                $crate::sentry_integration::BreadcrumbLevel::$level,
            );
        }
    };
}

/// Flush esplicito prima della chiusura.
/// Il drop del guard lo fa automaticamente, ma questo permette di specificare un timeout.
pub fn flush(timeout_secs: u64) {
    if let Some(guard) = SENTRY_GUARD.get()
        && guard.is_enabled()
        && let Some(client) = sentry::Hub::current().client()
    {
        client.flush(Some(std::time::Duration::from_secs(timeout_secs)));
    }
}

/// Verifica se Sentry è attivo.
pub fn is_enabled() -> bool {
    SENTRY_GUARD.get().map(|g| g.is_enabled()).unwrap_or(false)
}

/// Formatta un EventId per la visualizzazione all'utente.
/// Ritorna una stringa nel formato "Error ID: <short_id>" o stringa vuota se None.
pub fn format_event_id(event_id: Option<&EventId>) -> String {
    match event_id {
        Some(id) => format!("\n\nError ID: {}", id),
        None => String::new(),
    }
}

/// Salva l'ultimo errore fatale su file per inclusione nei diagnostici.
/// Il messaggio viene sanitizzato per rimuovere path sensibili.
pub fn save_last_fatal(message: &str, event_id: Option<&EventId>) {
    let settings_dir = crate::settings::settings_dir();
    let path = settings_dir.join("last_error.txt");

    let timestamp = chrono::Local::now().format("%Y-%m-%d %H:%M:%S");
    let sanitized = sanitize_message(message);
    let event_id_str = event_id
        .map(|id| format!("Event ID: {}\n", id))
        .unwrap_or_default();

    let content = format!(
        "Timestamp: {}\n\
         {}\
         Message: {}\n",
        timestamp, event_id_str, sanitized
    );

    if let Err(e) = std::fs::write(&path, content) {
        crate::log_debug(&format!("Failed to save last fatal error: {}", e));
    }
}

/// Installa un panic hook che:
/// 1. Logga su file
/// 2. Invia a Sentry (se abilitato)
/// 3. Chiama il panic hook precedente
pub fn install_panic_hook() {
    let previous_hook = std::panic::take_hook();

    std::panic::set_hook(Box::new(move |info: &PanicHookInfo<'_>| {
        // 1. Costruisci il messaggio di panic
        let location = info
            .location()
            .map(|loc| format!("{}:{}:{}", loc.file(), loc.line(), loc.column()))
            .unwrap_or_else(|| "<unknown>".to_string());

        let payload = if let Some(s) = info.payload().downcast_ref::<&str>() {
            (*s).to_string()
        } else if let Some(s) = info.payload().downcast_ref::<String>() {
            s.clone()
        } else {
            "unknown panic payload".to_string()
        };

        let panic_message = format!("PANIC at {}: {}", location, payload);

        // 2. Logga sempre su file (prima di tutto)
        crate::log_debug(&panic_message);

        // 3. Invia a Sentry se abilitato
        if is_enabled() {
            let sanitized = sanitize_message(&panic_message);
            sentry::with_scope(
                |scope| {
                    scope.set_tag("err.severity", "fatal");
                    scope.set_tag("err.type", "panic");
                },
                || {
                    sentry::capture_message(&sanitized, sentry::Level::Error);
                },
            );
            // Flush sincrono per assicurarsi che l'evento sia inviato
            if let Some(client) = sentry::Hub::current().client() {
                client.flush(Some(std::time::Duration::from_secs(2)));
            }
        }

        // 4. Chiama il panic hook precedente (per backtrace, ecc.)
        previous_hook(info);
    }));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sanitize_simple_path() {
        let input = "Failed to open C:\\Users\\John\\Documents\\file.txt";
        let result = sanitize_message(input);
        assert_eq!(result, "Failed to open <path>");
    }

    #[test]
    fn test_sanitize_multiple_paths() {
        let input = "Copy from C:\\source\\file.txt to D:\\dest\\file.txt failed";
        let result = sanitize_message(input);
        assert_eq!(result, "Copy from <path> to <path> failed");
    }

    #[test]
    fn test_sanitize_unc_path() {
        let input = "Cannot access \\\\server\\share\\folder\\file.doc";
        let result = sanitize_message(input);
        assert_eq!(result, "Cannot access <path>");
    }

    #[test]
    fn test_sanitize_no_path() {
        let input = "Generic error without path";
        let result = sanitize_message(input);
        assert_eq!(result, "Generic error without path");
    }

    #[test]
    fn test_sanitize_lowercase_drive() {
        let input = "Error reading c:\\temp\\data.bin";
        let result = sanitize_message(input);
        assert_eq!(result, "Error reading <path>");
    }
}
