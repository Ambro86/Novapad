use curl::easy::{Easy, List};
use std::ffi::CString;
use std::time::Duration;

fn log_profile(profile: &str, url: &str, status: &str) {
    if let Ok(mut exe_path) = std::env::current_exe() {
        exe_path.set_file_name("debug_curl_profile.log");
        use std::io::Write;
        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&exe_path)
        {
            let _ = writeln!(file, "[{}] {} - {}", profile, status, url);
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

impl CurlClient {
    pub fn fetch_url_impersonated(url: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        // WSJ e Dow Jones: vai diretto con iPhone
        let is_wsj_or_dowjones =
            url.contains("wsj.com") || url.contains("dowjones.com") || url.contains("barrons.com");

        if is_wsj_or_dowjones {
            log_profile("IPHONE_SAFARI", url, "direct (WSJ/DowJones)");
            return Self::fetch_iphone(url);
        }

        // PRIMA: proviamo con il profilo Chrome dettagliato (TLS fingerprinting avanzato)
        log_profile("CHROME_ADVANCED", url, "attempting");
        match fetch_url_chrome_advanced(url) {
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
        let result = Self::fetch_iphone(url);
        match &result {
            Ok(bytes) => log_profile("IPHONE_SAFARI", url, &format!("done (len={})", bytes.len())),
            Err(e) => log_profile("IPHONE_SAFARI", url, &format!("error: {}", e)),
        }
        result
    }

    fn fetch_iphone(url: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut easy = Easy::new();
        easy.url(url)?;
        easy.follow_location(true)?;
        easy.timeout(Duration::from_secs(25))?;
        easy.accept_encoding("gzip, deflate, br")?;
        easy.pipewait(true)?;
        easy.cookie_file("")?;

        // Verifica certificati CA da APPDATA (estratti da embedded_deps)
        let cacert_path = crate::embedded_deps::cacert_path();
        if cacert_path.exists() {
            easy.cainfo(cacert_path.to_string_lossy().as_ref())?;
        } else {
            easy.ssl_verify_peer(false)?;
            easy.ssl_verify_host(false)?;
        }

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
}

/// Vecchio profilo Chrome dettagliato con TLS fingerprinting avanzato
/// (dal commit c5e5842)
fn fetch_url_chrome_advanced(url: &str) -> anyhow::Result<Vec<u8>> {
    let mut easy = Easy::new();
    easy.url(url)?;

    // Abilita il motore dei cookie (fondamentale per evitare "cookie absent")
    easy.cookie_file("")?;

    easy.accept_encoding("")?;
    easy.follow_location(true)?;
    easy.max_redirections(10)?;
    easy.connect_timeout(std::time::Duration::from_secs(10))?;
    easy.timeout(std::time::Duration::from_secs(30))?;

    // Verifica certificati CA da APPDATA (estratti da embedded_deps)
    let cacert_path = crate::embedded_deps::cacert_path();
    if cacert_path.exists() {
        easy.cainfo(cacert_path.to_string_lossy().as_ref())?;
    } else {
        easy.ssl_verify_peer(false)?;
        easy.ssl_verify_host(false)?;
    }

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
