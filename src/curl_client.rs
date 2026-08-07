use curl::easy::{Easy, List};
use std::ffi::CString;
use std::fs::OpenOptions;
use std::io::{Error, Seek, SeekFrom, Write};
use std::path::Path;
use std::time::Duration;

fn log_profile(_profile: &str, _url: &str, _status: &str) {
    #[cfg(debug_assertions)]
    if let Ok(mut exe_path) = std::env::current_exe() {
        exe_path.set_file_name("debug_curl_profile.log");
        use std::io::Write;
        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&exe_path)
        {
            crate::log_if_err!(writeln!(file, "[{}] {} - {}", _profile, _status, _url));
        }
    }
}

pub const CURLOPT_SSL_ENABLE_ALPS: i32 = 1002;
pub const CURLOPT_SSL_CERT_COMPRESSION: i32 = 1003;
pub const CURLOPT_SSL_ENABLE_TICKET: i32 = 1004;
pub const CURLOPT_HTTP2_PSEUDO_HEADERS_ORDER: i32 = 1005;
pub const CURLOPT_HTTP2_SETTINGS: i32 = 1006;
pub const CURLOPT_SSL_PERMUTE_EXTENSIONS: i32 = 1007;
pub const CURLOPT_TLS_GREASE: i32 = 1011;
pub const CURLOPT_TLS_EXTENSION_ORDER: i32 = 1012;

pub struct CurlClient;

fn parse_content_length_header(header_line: &[u8]) -> Option<u64> {
    let text = std::str::from_utf8(header_line).ok()?.trim();
    let (name, value) = text.split_once(':')?;
    if !name.eq_ignore_ascii_case("content-length") {
        return None;
    }
    value.trim().parse::<u64>().ok().filter(|v| *v > 0)
}

fn parse_http_status_header(header_line: &[u8]) -> Option<u32> {
    let text = std::str::from_utf8(header_line).ok()?.trim();
    let mut parts = text.split_whitespace();
    let version = parts.next()?;
    if !version.starts_with("HTTP/") {
        return None;
    }
    parts.next()?.parse::<u32>().ok()
}

fn parse_content_range_total(header_line: &[u8]) -> Option<u64> {
    let text = std::str::from_utf8(header_line).ok()?.trim();
    let (name, value) = text.split_once(':')?;
    if !name.eq_ignore_ascii_case("content-range") {
        return None;
    }
    let (_, total) = value.trim().split_once('/')?;
    total.trim().parse::<u64>().ok().filter(|v| *v > 0)
}

fn apply_tls_ca(easy: &mut Easy) -> anyhow::Result<()> {
    // Prima scelta: CA bundle embedded in memoria (evita problemi di path Unicode/permessi su cacert.pem).
    let (blob_applied, blob_rc) = unsafe {
        let handle = easy.raw();
        let cacert = crate::embedded_deps::cacert_bytes();
        let mut blob = curl_sys::curl_blob {
            data: cacert.as_ptr() as *mut _,
            len: cacert.len(),
            flags: curl_sys::CURL_BLOB_COPY,
        };
        let rc = curl_sys::curl_easy_setopt(handle, curl_sys::CURLOPT_CAINFO_BLOB, &mut blob);
        (rc == curl_sys::CURLE_OK, rc)
    };
    if blob_applied {
        return Ok(());
    }

    crate::log_debug(&format!(
        "Curl: CURLOPT_CAINFO_BLOB unavailable/failed (rc={:?}), falling back to cacert.pem path",
        blob_rc
    ));

    // Fallback legacy: file CA in AppData.
    let cacert_path = crate::embedded_deps::cacert_path();
    if cacert_path.exists() {
        easy.cainfo(cacert_path.to_string_lossy().as_ref())?;
    } else {
        // Ultimo fallback compatibilità: disabilita verify come già faceva il codice precedente.
        easy.ssl_verify_peer(false)?;
        easy.ssl_verify_host(false)?;
    }
    Ok(())
}

impl CurlClient {
    pub fn fetch_url_impersonated(url: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        Self::fetch_url_impersonated_with_progress(url, |_| {})
    }

    pub fn fetch_url_iphone_impersonated(url: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        log_profile("IPHONE_SAFARI", url, "forced");
        Self::fetch_iphone(url, |_| {})
    }

    pub fn fetch_url_impersonated_with_progress<F: FnMut(u32)>(
        url: &str,
        mut progress_cb: F,
    ) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        // WSJ, Dow Jones e Podbean: vai diretto con iPhone
        let use_iphone_direct = url.contains("wsj.com")
            || url.contains("dowjones.com")
            || url.contains("barrons.com")
            || url.contains("podbean.com");

        if use_iphone_direct {
            log_profile("IPHONE_SAFARI", url, "direct (Known Host)");
            return Self::fetch_iphone(url, progress_cb);
        }

        // PRIMA: proviamo con il profilo Chrome dettagliato (TLS fingerprinting avanzato)
        log_profile("CHROME_ADVANCED", url, "attempting");
        match fetch_url_chrome_advanced(url, &mut progress_cb) {
            Ok(bytes) => {
                let check = String::from_utf8_lossy(&bytes).to_lowercase();
                // Se non è bloccato, ritorna il risultato
                if !check.contains("just a moment")
                    && !check.contains("dd-captcha")
                    && bytes.len() >= 3000
                {
                    log_profile("CHROME_ADVANCED", url, "success");
                    return Ok(bytes);
                }
                // Altrimenti, fallback su iPhone
                log_profile(
                    "CHROME_ADVANCED",
                    url,
                    &format!("blocked (len={})", bytes.len()),
                );
            }
            Err(e) => {
                // Se fallisce, fallback su iPhone
                log_profile("CHROME_ADVANCED", url, &format!("error: {}", e));
            }
        }

        // FALLBACK: iPhone Safari
        log_profile("IPHONE_SAFARI", url, "attempting fallback");
        let result = Self::fetch_iphone(url, progress_cb);
        match &result {
            Ok(bytes) => log_profile("IPHONE_SAFARI", url, &format!("done (len={})", bytes.len())),
            Err(e) => log_profile("IPHONE_SAFARI", url, &format!("error: {}", e)),
        }
        result
    }

    pub fn fetch_url_to_file_with_progress<F: FnMut(u32)>(
        url: &str,
        destination: &Path,
        resume_from: u64,
        progress_cb: F,
    ) -> Result<u64, Box<dyn std::error::Error>> {
        Self::fetch_iphone_to_file(url, destination, resume_from, progress_cb)
    }

    pub fn post_form_impersonated(
        url: &str,
        body: &str,
        headers: &[&str],
    ) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut easy = Easy::new();
        easy.url(url)?;
        easy.follow_location(true)?;
        easy.timeout(Duration::from_secs(30))?;
        easy.connect_timeout(Duration::from_secs(30))?;
        easy.accept_encoding("gzip, deflate, br")?;
        easy.cookie_file("")?;
        easy.post(true)?;
        easy.post_fields_copy(body.as_bytes())?;

        apply_tls_ca(&mut easy)?;

        let mut list = List::new();
        for header in headers {
            list.append(header)?;
        }
        easy.http_headers(list)?;

        let mut data = Vec::new();
        {
            let mut transfer = easy.transfer();
            transfer.write_function(|new_data| {
                data.extend_from_slice(new_data);
                Ok(new_data.len())
            })?;
            transfer.perform()?;
        }
        Ok(data)
    }

    pub fn resolve_final_url_iphone_impersonated(
        url: &str,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let mut easy = Easy::new();
        easy.url(url)?;
        easy.follow_location(true)?;
        easy.max_redirections(10)?;
        easy.timeout(Duration::from_secs(30))?;
        easy.connect_timeout(Duration::from_secs(30))?;
        easy.accept_encoding("gzip, deflate, br")?;
        easy.pipewait(true)?;
        easy.cookie_file("")?;
        // Request only the first byte when possible to avoid downloading the full media.
        easy.range("0-0")?;

        apply_tls_ca(&mut easy)?;

        easy.ssl_cipher_list("ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305")?;

        let mut list = List::new();
        list.append("User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1")?;
        list.append("Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")?;
        list.append("Accept-Language: it-IT,it;q=0.9,en-US;q=0.8")?;
        list.append("Upgrade-Insecure-Requests: 1")?;
        list.append("Connection: keep-alive")?;
        easy.http_headers(list)?;

        {
            let mut transfer = easy.transfer();
            transfer.write_function(|new_data| Ok(new_data.len()))?;
            transfer.perform()?;
        }

        if let Some(effective_url) = easy.effective_url()?
            && !effective_url.trim().is_empty()
        {
            return Ok(effective_url.to_string());
        }

        Ok(url.to_string())
    }

    fn fetch_iphone<F: FnMut(u32)>(
        url: &str,
        mut progress_cb: F,
    ) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        crate::log_debug(&format!("Curl: fetch_iphone starting for {}", url));
        let mut easy = Easy::new();
        easy.url(url)?;
        easy.follow_location(true)?;
        easy.timeout(Duration::from_secs(600))?; // Aumentato a 10m per file grandi
        easy.connect_timeout(Duration::from_secs(30))?;
        easy.accept_encoding("gzip, deflate, br")?;
        easy.pipewait(true)?;
        easy.cookie_file("")?;
        easy.progress(true)?;

        apply_tls_ca(&mut easy)?;

        // Cipher list compatibile con curl/OpenSSL
        easy.ssl_cipher_list("ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305")?;

        let mut list = List::new();
        list.append("User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1")?;
        list.append("Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")?;
        list.append("Accept-Language: it-IT,it;q=0.9,en-US;q=0.8")?;
        list.append("Upgrade-Insecure-Requests: 1")?;
        list.append("Connection: keep-alive")?;

        easy.http_headers(list)?;

        let mut data = Vec::new();
        crate::log_debug("Curl: starting perform...");
        let mut last_log_mb = 0;
        let mut last_pct = 0u32;
        let header_content_length = std::cell::Cell::new(0.0f64);
        {
            let mut transfer = easy.transfer();
            transfer.header_function(|header| {
                if let Some(len) = parse_content_length_header(header) {
                    header_content_length.set(len as f64);
                }
                true
            })?;
            transfer.write_function(|new_data| {
                data.extend_from_slice(new_data);
                let current_mb = data.len() / (1024 * 1024);
                if current_mb > last_log_mb && current_mb % 5 == 0 {
                    crate::log_debug(&format!("Curl: downloaded {} MB...", current_mb));
                    last_log_mb = current_mb;
                }
                Ok(new_data.len())
            })?;
            transfer.progress_function(|dltotal, dlnow, _, _| {
                let total = if dltotal > 0.0 {
                    dltotal
                } else {
                    header_content_length.get()
                };
                if total > 0.0 {
                    let pct = (dlnow / total * 100.0) as u32;
                    if pct > last_pct {
                        last_pct = pct;
                        progress_cb(pct);
                    }
                }
                true
            })?;
            transfer.perform()?;
        }
        crate::log_debug(&format!(
            "Curl: perform finished, downloaded {} bytes",
            data.len()
        ));
        Ok(data)
    }

    fn fetch_iphone_to_file<F: FnMut(u32)>(
        url: &str,
        destination: &Path,
        resume_from: u64,
        mut progress_cb: F,
    ) -> Result<u64, Box<dyn std::error::Error>> {
        crate::log_debug(&format!(
            "Curl: fetch_iphone_to_file starting for {} resume_from={}",
            url, resume_from
        ));
        let mut easy = Easy::new();
        easy.url(url)?;
        easy.follow_location(true)?;
        easy.timeout(Duration::from_secs(600))?;
        easy.connect_timeout(Duration::from_secs(30))?;
        easy.accept_encoding("gzip, deflate, br")?;
        easy.pipewait(true)?;
        easy.cookie_file("")?;
        easy.progress(true)?;
        if resume_from > 0 {
            easy.range(&format!("{resume_from}-"))?;
        }

        apply_tls_ca(&mut easy)?;

        easy.ssl_cipher_list("ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305")?;

        let mut list = List::new();
        list.append("User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1")?;
        list.append("Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")?;
        list.append("Accept-Language: it-IT,it;q=0.9,en-US;q=0.8")?;
        list.append("Upgrade-Insecure-Requests: 1")?;
        list.append("Connection: keep-alive")?;
        easy.http_headers(list)?;

        let file = OpenOptions::new()
            .create(true)
            .truncate(false)
            .read(true)
            .write(true)
            .open(destination)?;
        let file = std::cell::RefCell::new(file);
        if resume_from > 0 {
            file.borrow_mut().seek(SeekFrom::Start(resume_from))?;
        }

        crate::log_debug("Curl: starting perform...");
        let mut last_log_mb = 0;
        let mut last_pct = 0u32;
        let status_code = std::cell::Cell::new(0u32);
        let content_length = std::cell::Cell::new(0u64);
        let content_range_total = std::cell::Cell::new(0u64);
        let effective_resume = std::cell::Cell::new(resume_from);
        let downloaded_this_request = std::cell::Cell::new(0u64);
        let reset_to_zero = std::cell::Cell::new(false);
        let write_error = std::cell::RefCell::new(None::<String>);
        {
            let mut transfer = easy.transfer();
            transfer.header_function(|header| {
                if let Some(status) = parse_http_status_header(header) {
                    status_code.set(status);
                }
                if let Some(len) = parse_content_length_header(header) {
                    content_length.set(len);
                }
                if let Some(total) = parse_content_range_total(header) {
                    content_range_total.set(total);
                }
                true
            })?;
            transfer.write_function(|new_data| {
                if resume_from > 0 && !reset_to_zero.get() && status_code.get() == 200 {
                    let reset_result = {
                        let mut writer = file.borrow_mut();
                        writer
                            .set_len(0)
                            .and_then(|_| writer.seek(SeekFrom::Start(0)).map(|_| ()))
                    };
                    if let Err(err) = reset_result {
                        *write_error.borrow_mut() = Some(err.to_string());
                        return Err(curl::easy::WriteError::Pause);
                    }
                    effective_resume.set(0);
                    downloaded_this_request.set(0);
                    reset_to_zero.set(true);
                    crate::log_debug(
                        "Curl: resume request ignored by server, restarting file from zero",
                    );
                }

                {
                    let mut writer = file.borrow_mut();
                    if let Err(err) = writer.write_all(new_data) {
                        *write_error.borrow_mut() = Some(err.to_string());
                        return Err(curl::easy::WriteError::Pause);
                    }
                }

                let downloaded = downloaded_this_request
                    .get()
                    .saturating_add(new_data.len() as u64);
                downloaded_this_request.set(downloaded);
                let written_total = effective_resume.get().saturating_add(downloaded);
                let current_mb = (written_total / (1024 * 1024)) as usize;
                if current_mb > last_log_mb && current_mb.is_multiple_of(5) {
                    crate::log_debug(&format!("Curl: downloaded {} MB...", current_mb));
                    last_log_mb = current_mb;
                }
                Ok(new_data.len())
            })?;
            transfer.progress_function(|dltotal, dlnow, _, _| {
                let total = if content_range_total.get() > 0 {
                    content_range_total.get() as f64
                } else if dltotal > 0.0 {
                    dltotal + effective_resume.get() as f64
                } else if content_length.get() > 0 {
                    content_length.get() as f64 + effective_resume.get() as f64
                } else {
                    0.0
                };

                if total > 0.0 {
                    let downloaded_now = if status_code.get() == 200 {
                        dlnow
                    } else {
                        dlnow + effective_resume.get() as f64
                    };
                    let pct = (downloaded_now / total * 100.0) as u32;
                    if pct > last_pct {
                        last_pct = pct;
                        progress_cb(pct);
                    }
                }
                true
            })?;
            if let Err(err) = transfer.perform() {
                if let Some(message) = write_error.borrow_mut().take() {
                    return Err(Box::new(Error::other(message)));
                }
                return Err(err.into());
            }
        }

        file.borrow_mut().flush()?;
        let final_len = file.borrow().metadata()?.len();
        crate::log_debug(&format!(
            "Curl: perform finished, wrote {} bytes to {}",
            final_len,
            destination.display()
        ));
        Ok(final_len)
    }
}

/// Vecchio profilo Chrome dettagliato con TLS fingerprinting avanzato
/// (dal commit c5e5842)
fn fetch_url_chrome_advanced<F: FnMut(u32)>(
    url: &str,
    mut progress_cb: F,
) -> anyhow::Result<Vec<u8>> {
    let mut easy = Easy::new();
    easy.url(url)?;

    // Abilita il motore dei cookie (fondamentale per evitare "cookie absent")
    easy.cookie_file("")?;

    easy.accept_encoding("")?;
    easy.follow_location(true)?;
    easy.max_redirections(10)?;
    easy.connect_timeout(std::time::Duration::from_secs(30))?;
    easy.timeout(std::time::Duration::from_secs(600))?; // 10m per file grandi
    easy.progress(true)?;

    apply_tls_ca(&mut easy)?;

    unsafe {
        let handle = easy.raw();
        curl_sys::curl_easy_setopt(handle, CURLOPT_TLS_GREASE, 1);
        curl_sys::curl_easy_setopt(handle, CURLOPT_SSL_PERMUTE_EXTENSIONS, 1);
        curl_sys::curl_easy_setopt(handle, CURLOPT_SSL_ENABLE_TICKET, 1);
        curl_sys::curl_easy_setopt(handle, CURLOPT_SSL_ENABLE_ALPS, 1);

        let tls_exts = CString::new(
            "grease,server_name,extended_master_secret,renegotiation_info,supported_groups,ec_point_formats,session_ticket,application_layer_protocol_negotiation,status_request,signature_algorithms,signed_certificate_timestamp,compress_certificate,application_settings,key_share,psk_key_exchange_modes,supported_versions",
        )?;
        curl_sys::curl_easy_setopt(handle, CURLOPT_TLS_EXTENSION_ORDER, tls_exts.as_ptr());

        let h2_order = CString::new("m,a,s,p")?;
        curl_sys::curl_easy_setopt(
            handle,
            CURLOPT_HTTP2_PSEUDO_HEADERS_ORDER,
            h2_order.as_ptr(),
        );

        let h2_settings = CString::new("1:65536;3:1000;4:6291456;6:262144")?;
        curl_sys::curl_easy_setopt(handle, CURLOPT_HTTP2_SETTINGS, h2_settings.as_ptr());

        let cert_comp = CString::new("brotli")?;
        curl_sys::curl_easy_setopt(handle, CURLOPT_SSL_CERT_COMPRESSION, cert_comp.as_ptr());
    }

    let mut list = List::new();
    list.append("User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")?;
    list.append("Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")?;
    list.append("Accept-Language: it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7")?;
    list.append("Cache-Control: max-age=0")?;
    list.append(
        "Sec-Ch-Ua: \"Google Chrome\";v=\"131\", \"Chromium\";v=\"131\", \"Not_A Brand\";v=\"24\"",
    )?;
    list.append("Sec-Ch-Ua-Mobile: ?0")?;
    list.append("Sec-Ch-Ua-Platform: \"Windows\"")?;
    list.append("Upgrade-Insecure-Requests: 1")?;
    list.append("Sec-Fetch-Dest: document")?;
    list.append("Sec-Fetch-Mode: navigate")?;
    list.append("Sec-Fetch-Site: none")?;
    list.append("Sec-Fetch-User: ?1")?;

    // Aggiungiamo un Referer credibile
    list.append("Referer: https://www.google.com/")?;

    easy.http_headers(list)?;

    let mut data = Vec::new();
    let mut last_pct = 0u32;
    let header_content_length = std::cell::Cell::new(0.0f64);
    {
        let mut transfer = easy.transfer();
        transfer.header_function(|header| {
            if let Some(len) = parse_content_length_header(header) {
                header_content_length.set(len as f64);
            }
            true
        })?;
        transfer.write_function(|new_data| {
            data.extend_from_slice(new_data);
            Ok(new_data.len())
        })?;
        transfer.progress_function(|dltotal, dlnow, _, _| {
            let total = if dltotal > 0.0 {
                dltotal
            } else {
                header_content_length.get()
            };
            if total > 0.0 {
                let pct = (dlnow / total * 100.0) as u32;
                if pct > last_pct {
                    last_pct = pct;
                    progress_cb(pct);
                }
            }
            true
        })?;
        transfer.perform()?;
    }

    Ok(data)
}
