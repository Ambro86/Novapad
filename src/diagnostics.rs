// src/diagnostics.rs
//
// Export diagnostics for troubleshooting.
// Creates a zip file with logs, settings, and version info.
// All paths are sanitized to avoid PII leakage.

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;
use zip::ZipWriter;
use zip::write::FileOptions;

use crate::sentry_integration;
use crate::settings;

/// Sanitizza i path nel contenuto testuale per evitare PII.
/// Sostituisce path Windows (C:\Users\...) con <user_path>.
fn sanitize_content(content: &str) -> String {
    // Pattern per path utente Windows
    let re = fancy_regex::Regex::new(r"(?i)C:\\Users\\[^\\]+").ok();
    match re {
        Some(regex) => regex.replace_all(content, "<user_path>").to_string(),
        None => content.to_string(),
    }
}

/// Genera il contenuto di version.txt
fn generate_version_info() -> String {
    let version = env!("CARGO_PKG_VERSION");

    #[cfg(feature = "standalone")]
    let mode = "standalone";
    #[cfg(not(feature = "standalone"))]
    let mode = "normal";

    let debug = if cfg!(debug_assertions) {
        "debug"
    } else {
        "release"
    };

    let commit = option_env!("GIT_COMMIT_HASH").unwrap_or("unknown");
    let build_date = option_env!("BUILD_DATE").unwrap_or("unknown");
    let rustc = option_env!("RUSTC_VERSION").unwrap_or("unknown");
    let arch = std::env::consts::ARCH;
    let os = std::env::consts::OS;

    format!(
        "Sonarpad Diagnostics\n\
         ==================\n\n\
         Version: {}\n\
         Mode: {}\n\
         Build: {}\n\
         Commit: {}\n\
         Build date: {}\n\
         Rustc: {}\n\
         Arch: {}\n\
         OS: {}\n\
         Sentry enabled: {}\n",
        version,
        mode,
        debug,
        commit,
        build_date,
        rustc,
        arch,
        os,
        sentry_integration::is_enabled()
    )
}

/// Aggiunge un file al zip, gestendo errori gracefully.
fn add_file_to_zip<W: Write + std::io::Seek>(
    zip: &mut ZipWriter<W>,
    source_path: &Path,
    archive_name: &str,
    sanitize: bool,
) -> Result<(), String> {
    let mut file = match File::open(source_path) {
        Ok(f) => f,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            // File non esiste, skip silently
            return Ok(());
        }
        Err(e) => {
            return Err(format!("Failed to open {}: {}", archive_name, e));
        }
    };

    let mut content = String::new();
    if file.read_to_string(&mut content).is_err() {
        // Prova a leggere come binario se non è testo valido
        let mut binary_content = Vec::new();
        let mut file = match File::open(source_path) {
            Ok(f) => f,
            Err(e) => {
                return Err(format!("Failed to reopen {}: {}", archive_name, e));
            }
        };
        if let Err(e) = file.read_to_end(&mut binary_content) {
            return Err(format!("Failed to read {}: {}", archive_name, e));
        }

        let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        if let Err(e) = zip.start_file(archive_name, options) {
            return Err(format!("Failed to create {} in zip: {}", archive_name, e));
        }
        if let Err(e) = zip.write_all(&binary_content) {
            return Err(format!("Failed to write {} to zip: {}", archive_name, e));
        }
        return Ok(());
    }

    let final_content = if sanitize {
        sanitize_content(&content)
    } else {
        content
    };

    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    if let Err(e) = zip.start_file(archive_name, options) {
        return Err(format!("Failed to create {} in zip: {}", archive_name, e));
    }
    if let Err(e) = zip.write_all(final_content.as_bytes()) {
        return Err(format!("Failed to write {} to zip: {}", archive_name, e));
    }

    Ok(())
}

/// Aggiunge contenuto testuale direttamente al zip.
fn add_text_to_zip<W: Write + std::io::Seek>(
    zip: &mut ZipWriter<W>,
    archive_name: &str,
    content: &str,
) -> Result<(), String> {
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    if let Err(e) = zip.start_file(archive_name, options) {
        return Err(format!("Failed to create {} in zip: {}", archive_name, e));
    }
    if let Err(e) = zip.write_all(content.as_bytes()) {
        return Err(format!("Failed to write {} to zip: {}", archive_name, e));
    }
    Ok(())
}

/// Esporta i file diagnostici in un archivio zip.
///
/// Include:
/// - version.txt: informazioni sulla versione e build
/// - Sonarpad.log: file di log (sanitizzato)
/// - settings.json: configurazione (sanitizzata)
/// - last_error.txt: ultimo errore fatale (se esiste)
/// - sentry_last_event_id.txt: ultimo event ID Sentry (se esiste)
pub fn export_diagnostics_zip(dest_path: &Path) -> Result<(), String> {
    let settings_dir = settings::settings_dir();

    // Crea il file zip
    let file = File::create(dest_path).map_err(|e| format!("Failed to create zip file: {}", e))?;
    let mut zip = ZipWriter::new(file);

    // 1. version.txt - sempre generato
    let version_info = generate_version_info();
    add_text_to_zip(&mut zip, "version.txt", &version_info)?;

    // 2. Sonarpad.log - sanitizzato
    let log_path = settings_dir.join("Sonarpad.log");
    add_file_to_zip(&mut zip, &log_path, "Sonarpad.log", true)?;

    // 3. settings.json - sanitizzato
    let settings_path = settings_dir.join("settings.json");
    add_file_to_zip(&mut zip, &settings_path, "settings.json", true)?;

    // 4. last_error.txt - se esiste (sanitizzato)
    let last_error_path = settings_dir.join("last_error.txt");
    add_file_to_zip(&mut zip, &last_error_path, "last_error.txt", true)?;

    // 5. sentry_last_event_id.txt - se esiste
    let sentry_id_path = settings_dir.join("sentry_last_event_id.txt");
    add_file_to_zip(&mut zip, &sentry_id_path, "sentry_last_event_id.txt", false)?;

    // 6. last_hang.txt - se esiste (info su possibili freeze)
    let last_hang_path = settings_dir.join("last_hang.txt");
    add_file_to_zip(&mut zip, &last_hang_path, "last_hang.txt", false)?;

    // Finalizza il zip
    zip.finish()
        .map_err(|e| format!("Failed to finalize zip: {}", e))?;

    Ok(())
}
