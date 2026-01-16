use curl::easy::{Easy, List};
use std::ffi::CString;
use std::path::Path;

pub const CURLOPT_SSL_ENABLE_ALPS: i32 = 1002;
pub const CURLOPT_SSL_CERT_COMPRESSION: i32 = 1003;
pub const CURLOPT_SSL_ENABLE_TICKET: i32 = 1004;
pub const CURLOPT_HTTP2_PSEUDO_HEADERS_ORDER: i32 = 1005;
pub const CURLOPT_HTTP2_SETTINGS: i32 = 1006;
pub const CURLOPT_SSL_PERMUTE_EXTENSIONS: i32 = 1007;
pub const CURLOPT_TLS_GREASE: i32 = 1011;
pub const CURLOPT_TLS_EXTENSION_ORDER: i32 = 1012;

pub fn fetch_url_impersonated(url: &str) -> anyhow::Result<Vec<u8>> {
    let mut easy = Easy::new();
    easy.url(url)?;

    // Abilita il motore dei cookie (fondamentale per evitare "cookie absent")
    easy.cookie_file("")?;

    easy.accept_encoding("")?;
    easy.follow_location(true)?;
    easy.max_redirections(10)?;
    easy.connect_timeout(std::time::Duration::from_secs(10))?;
    easy.timeout(std::time::Duration::from_secs(30))?;

    // Verifica certificati CA
    if Path::new("cacert.pem").exists() {
        easy.cainfo("cacert.pem")?;
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
