use curl::easy::{Easy, List};
use std::time::Duration;

pub struct CurlClient;

impl CurlClient {
    pub fn fetch_url_impersonated(url: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        // Se è un sito del gruppo Dow Jones, andiamo diretti di iPhone che sappiamo funzionare
        let is_dow_jones =
            url.contains("dowjones.com") || url.contains("wsj.com") || url.contains("barrons.com");

        if is_dow_jones {
            return Self::execute_fetch(url, false);
        }

        // Altrimenti proviamo prima Chrome 131 (ottimo per Science.org)
        let res = Self::execute_fetch(url, true)?;
        let check = String::from_utf8_lossy(&res).to_lowercase();

        // Se Chrome viene bloccato, fallback su iPhone
        if check.contains("just a moment") || check.contains("dd-captcha") || check.len() < 3000 {
            return Self::execute_fetch(url, false);
        }

        Ok(res)
    }

    fn execute_fetch(url: &str, use_chrome: bool) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut easy = Easy::new();
        easy.url(url)?;
        easy.follow_location(true)?;
        easy.timeout(Duration::from_secs(25))?;
        easy.accept_encoding("gzip, deflate, br")?;
        easy.pipewait(true)?;

        let mut list = List::new();

        if use_chrome {
            // --- PROFILO CHROME 131 (Windows) ---
            easy.ssl_cipher_list("TLS_AES_128_GCM_SHA256:TLS_AES_256_GCM_SHA384:TLS_CHACHA20_POLY1305_SHA256:ECDHE-ECDSA-AES_128_GCM_SHA256:ECDHE-RSA-AES_128_GCM_SHA256:ECDHE-ECDSA-AES_256_GCM_SHA384:ECDHE-RSA-AES_256_GCM_SHA384")?;

            list.append("User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")?;
            list.append("Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")?;
            list.append("sec-ch-ua: \"Google Chrome\";v=\"131\", \"Chromium\";v=\"131\", \"Not_A Brand\";v=\"24\"")?;
            list.append("sec-ch-ua-mobile: ?0")?;
            list.append("sec-ch-ua-platform: \"Windows\"")?;
        } else {
            // --- PROFILO IPHONE SAFARI (iOS 17.5) ---
            easy.ssl_cipher_list("TLS_AES_128_GCM_SHA256:TLS_AES_256_GCM_SHA384:TLS_CHACHA20_POLY1305_SHA256:ECDHE-ECDSA-AES_128_GCM_SHA256:ECDHE-RSA-AES_128_GCM_SHA256:ECDHE-ECDSA-AES_256_GCM_SHA384:ECDHE-RSA-AES_256_GCM_SHA384")?;

            unsafe {
                let handle = easy.raw();
                curl_sys::curl_easy_setopt(handle, 1012, 1012); // Impersonate Safari TLS
                curl_sys::curl_easy_setopt(handle, 1005, 1005); // Impersonate Safari H2
            }

            list.append("User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1")?;
            list.append("Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")?;
        }

        list.append("Accept-Language: it-IT,it;q=0.9,en-US;q=0.8")?;
        list.append("Upgrade-Insecure-Requests: 1")?;
        list.append("Connection: keep-alive")?;

        easy.http_headers(list)?;
        easy.cookie_file("")?; // Cookie solo in memoria

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

    pub fn fetch(url: &str) -> Result<String, Box<dyn std::error::Error>> {
        let bytes = Self::fetch_url_impersonated(url)?;
        Ok(String::from_utf8_lossy(&bytes).to_string())
    }
}
