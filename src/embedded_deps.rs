//! Modulo per l'estrazione delle dipendenze embedded nell'exe.
//! Le DLL vengono estratte in %APPDATA%/novapad/ al primo avvio
//! e aggiornate solo se cambiano.

use std::fs;
use std::io::Write;
use std::path::PathBuf;

// Embed delle DLL direttamente nell'exe
const LIBCURL_DLL: &[u8] = include_bytes!("../dll/libcurl.dll");
const ZLIB_DLL: &[u8] = include_bytes!("../dll/zlib.dll");
const CACERT_PEM: &[u8] = include_bytes!("../dll/cacert.pem");
const SOUNDTOUCH_DLL: &[u8] = include_bytes!("../dll/SoundTouch64.dll");
const NVDA_CLIENT_DLL: &[u8] = include_bytes!("../dll/nvdaControllerClient64.dll");

/// Calcola un hash semplice per verificare se il file è cambiato
#[allow(dead_code)]
fn simple_hash(data: &[u8]) -> u64 {
    use std::collections::hash_map::DefaultHasher;
    use std::hash::{Hash, Hasher};
    let mut hasher = DefaultHasher::new();
    data.hash(&mut hasher);
    hasher.finish()
}

/// Scrive un file solo se non esiste o se l'hash è diverso
fn write_if_changed(path: &PathBuf, data: &[u8]) -> std::io::Result<bool> {
    if path.exists() {
        // Controlla dimensione come quick check
        if let Ok(metadata) = fs::metadata(path) {
            if metadata.len() == data.len() as u64 {
                // Stessa dimensione, probabilmente uguale - skip
                return Ok(false);
            }
        }
    }

    // Scrivi in file temporaneo poi rinomina (atomico)
    let tmp_path = path.with_extension("tmp");
    let mut file = fs::File::create(&tmp_path)?;
    file.write_all(data)?;
    file.sync_all()?;
    drop(file);

    // Rimuovi vecchio file se esiste
    let _ = fs::remove_file(path);
    fs::rename(&tmp_path, path)?;

    Ok(true)
}

/// Estrae tutte le dipendenze in %APPDATA%/novapad/
/// Ritorna il path della cartella delle dipendenze
pub fn extract_all() -> std::io::Result<PathBuf> {
    let deps_dir = crate::settings::settings_dir();
    fs::create_dir_all(&deps_dir)?;

    // Lista delle dipendenze da estrarre
    let deps: &[(&str, &[u8])] = &[
        ("libcurl.dll", LIBCURL_DLL),
        ("zlib.dll", ZLIB_DLL),
        ("cacert.pem", CACERT_PEM),
        ("SoundTouch64.dll", SOUNDTOUCH_DLL),
        ("nvdaControllerClient64.dll", NVDA_CLIENT_DLL),
    ];

    for (name, data) in deps {
        let path = deps_dir.join(name);
        match write_if_changed(&path, data) {
            Ok(true) => {
                // File scritto/aggiornato
                #[cfg(debug_assertions)]
                eprintln!("[embedded_deps] Extracted: {}", name);
            }
            Ok(false) => {
                // File già presente e uguale
            }
            Err(e) => {
                eprintln!("[embedded_deps] Failed to extract {}: {}", name, e);
            }
        }
    }

    Ok(deps_dir)
}

/// Ritorna il path di una dipendenza specifica
pub fn get_dep_path(name: &str) -> PathBuf {
    crate::settings::settings_dir().join(name)
}

/// Ritorna il path di libcurl.dll
#[allow(dead_code)]
pub fn libcurl_path() -> PathBuf {
    get_dep_path("libcurl.dll")
}

/// Ritorna il path di cacert.pem
pub fn cacert_path() -> PathBuf {
    get_dep_path("cacert.pem")
}

/// Ritorna il path di SoundTouch64.dll
#[allow(dead_code)]
pub fn soundtouch_path() -> PathBuf {
    get_dep_path("SoundTouch64.dll")
}

/// Ritorna il path di nvdaControllerClient64.dll
#[allow(dead_code)]
pub fn nvda_client_path() -> PathBuf {
    get_dep_path("nvdaControllerClient64.dll")
}
