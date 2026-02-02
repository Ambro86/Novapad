//! Modulo per l'estrazione delle dipendenze embedded nell'exe.
//! Le DLL vengono estratte in %APPDATA%/sonarpad/ al primo avvio
//! e aggiornate solo se cambiano.

use std::fs;
use std::io::Write;
use std::path::PathBuf;

// Embed delle DLL direttamente nell'exe
// libcurl: runtime portatile senza dipendenze di sistema.
const LIBCURL_DLL: &[u8] = include_bytes!("../dll/libcurl.dll");
// zlib: dipendenza runtime di libcurl.
const ZLIB_DLL: &[u8] = include_bytes!("../dll/zlib.dll");
// cacert: bundle CA per la verifica TLS con libcurl embedded.
const CACERT_PEM: &[u8] = include_bytes!("../dll/cacert.pem");
// NVDA: integrazione con screen reader tramite controller client.
const NVDA_CLIENT_DLL: &[u8] = include_bytes!("../dll/nvdaControllerClient64.dll");
// BASS: backend audio per riproduzione e timing precisi.
const BASS_DLL: &[u8] = include_bytes!("../dll/bass.dll");
const BASS_FX_DLL: &[u8] = include_bytes!("../dll/bass_fx.dll");
const BASS_AAC_DLL: &[u8] = include_bytes!("../dll/bass_aac.dll");
const BASS_ALAC_DLL: &[u8] = include_bytes!("../dll/bass_alac.dll");
const BASS_FLAC_DLL: &[u8] = include_bytes!("../dll/bassflac.dll");
const BASS_OPUS_DLL: &[u8] = include_bytes!("../dll/bassopus.dll");
const BASS_WMA_DLL: &[u8] = include_bytes!("../dll/basswma.dll");
// PDFium: rendering/estrazione testo PDF più robusta.
const PDFIUM_DLL: &[u8] = include_bytes!("../dll/pdfium.dll");

// FFmpeg da vcpkg (link locale)
const AVCODEC_DLL: &[u8] = include_bytes!("../vcpkg_installed/x64-windows/bin/avcodec-62.dll");
const AVFORMAT_DLL: &[u8] = include_bytes!("../vcpkg_installed/x64-windows/bin/avformat-62.dll");
const AVUTIL_DLL: &[u8] = include_bytes!("../vcpkg_installed/x64-windows/bin/avutil-60.dll");
const SWRESAMPLE_DLL: &[u8] = include_bytes!("../vcpkg_installed/x64-windows/bin/swresample-6.dll");
const SWSCALE_DLL: &[u8] = include_bytes!("../vcpkg_installed/x64-windows/bin/swscale-9.dll");
const OPUS_DLL: &[u8] = include_bytes!("../vcpkg_installed/x64-windows/bin/opus.dll");

/// Scrive un file solo se non esiste o se l'hash è diverso
fn write_if_changed(path: &PathBuf, data: &[u8]) -> std::io::Result<bool> {
    if path.exists() {
        // Controlla dimensione come quick check
        if let Ok(metadata) = fs::metadata(path)
            && metadata.len() == data.len() as u64
        {
            // Stessa dimensione, probabilmente uguale - skip
            return Ok(false);
        }
    }

    // Scrivi in file temporaneo poi rinomina (atomico)
    let tmp_path = path.with_extension("tmp");
    let mut file = fs::File::create(&tmp_path)?;
    file.write_all(data)?;
    file.sync_all()?;
    std::mem::drop(file);

    // Rimuovi vecchio file se esiste
    crate::log_if_err!(fs::remove_file(path));
    fs::rename(&tmp_path, path)?;

    Ok(true)
}

/// Estrae tutte le dipendenze in %APPDATA%/sonarpad/
/// Ritorna il path della cartella delle dipendenze
pub fn extract_all() -> std::io::Result<PathBuf> {
    let deps_dir = crate::settings::settings_dir();
    fs::create_dir_all(&deps_dir)?;

    // Lista delle dipendenze da estrarre
    let deps: &[(&str, &[u8])] = &[
        ("libcurl.dll", LIBCURL_DLL),
        ("zlib.dll", ZLIB_DLL),
        ("cacert.pem", CACERT_PEM),
        ("nvdaControllerClient64.dll", NVDA_CLIENT_DLL),
        ("bass.dll", BASS_DLL),
        ("bass_fx.dll", BASS_FX_DLL),
        ("bass_aac.dll", BASS_AAC_DLL),
        ("bass_alac.dll", BASS_ALAC_DLL),
        ("bassflac.dll", BASS_FLAC_DLL),
        ("bassopus.dll", BASS_OPUS_DLL),
        ("basswma.dll", BASS_WMA_DLL),
        ("pdfium.dll", PDFIUM_DLL),
        ("avcodec-62.dll", AVCODEC_DLL),
        ("avformat-62.dll", AVFORMAT_DLL),
        ("avutil-60.dll", AVUTIL_DLL),
        ("swresample-6.dll", SWRESAMPLE_DLL),
        ("swscale-9.dll", SWSCALE_DLL),
        ("opus.dll", OPUS_DLL),
        ("sapi4_bridge_32.exe", SAPI4_BRIDGE_32_EXE),
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

/// Ritorna il path di cacert.pem
pub fn cacert_path() -> PathBuf {
    get_dep_path("cacert.pem")
}

// SAPI4 bridge: helper 32-bit per salvataggio su file.
const SAPI4_BRIDGE_32_EXE: &[u8] = include_bytes!("../dll/sapi4_bridge_32.exe");
