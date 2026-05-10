use flate2::read::GzDecoder;
use rand::Rng;
use rand::seq::SliceRandom;
use reqwest::Client;
use reqwest::header::{
    ACCEPT, ACCEPT_ENCODING, ACCEPT_LANGUAGE, AUTHORIZATION, CONTENT_TYPE, HeaderMap, HeaderName,
    HeaderValue, ORIGIN, REFERER, USER_AGENT,
};
use serde_json::{Value, json};
use std::fmt;
use std::io::Read;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::{Duration, SystemTime, UNIX_EPOCH};

pub struct TranslatorDeepLFree {
    pub source_lang: Option<String>,
    pub target_lang: String,
    endpoint: String,
    client: Client,
}

#[derive(Debug)]
pub enum TranslatorError {
    HttpClientBuild(String),
    Network(String),
    HttpStatus { status: u16, body: String },
    ReadResponse(String),
    DecodeGzip(String),
    DecodeUtf8(String),
    ParseJson(String),
    MissingTranslatedText,
    Cancelled,
}

impl fmt::Display for TranslatorError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TranslatorError::HttpClientBuild(err) => {
                write!(f, "Failed to build HTTP client: {err}")
            }
            TranslatorError::Network(err) => {
                write!(f, "Network error: {err}")
            }
            TranslatorError::HttpStatus { status, body } => {
                write!(f, "HTTP error {status}: {body}")
            }
            TranslatorError::ReadResponse(err) => {
                write!(f, "Failed to read response body: {err}")
            }
            TranslatorError::DecodeGzip(err) => {
                write!(f, "Failed to decode gzip response: {err}")
            }
            TranslatorError::DecodeUtf8(err) => {
                write!(f, "Failed to decode response as UTF-8: {err}")
            }
            TranslatorError::ParseJson(err) => {
                write!(f, "Failed to parse JSON response: {err}")
            }
            TranslatorError::MissingTranslatedText => {
                write!(f, "Translated text field was not found in the response")
            }
            TranslatorError::Cancelled => {
                write!(f, "Translation canceled")
            }
        }
    }
}

impl std::error::Error for TranslatorError {}

pub struct TranslatorGoogleFree {
    pub source_lang: String,
    pub target_lang: String,
    client: Client,
}

impl TranslatorGoogleFree {
    pub fn new(
        target_lang: impl Into<String>,
        source_lang: impl Into<String>,
    ) -> Result<Self, TranslatorError> {
        let client = Client::builder()
            .timeout(Duration::from_secs(30))
            .build()
            .map_err(|err| TranslatorError::HttpClientBuild(err.to_string()))?;

        Ok(Self {
            source_lang: source_lang.into(),
            target_lang: target_lang.into(),
            client,
        })
    }

    fn endpoints() -> &'static [&'static str] {
        &[
            "https://translate.googleapis.com/translate_a/single",
            "https://translate.googleapis.mirror.nvdadr.com/translate_a/single",
            "https://translate.google.com/translate_a/single",
            "https://translate.google.co.in/translate_a/single",
            "https://translate.google.co.uk/translate_a/single",
            "https://translate.google.com.au/translate_a/single",
            "https://translate.google.ca/translate_a/single",
            "https://translate.google.de/translate_a/single",
            "https://translate.google.es/translate_a/single",
            "https://translate.google.fr/translate_a/single",
            "https://translate.google.it/translate_a/single",
            "https://translate.google.nl/translate_a/single",
            "https://translate.google.pt/translate_a/single",
            "https://translate.google.ru/translate_a/single",
        ]
    }

    fn endpoints_random_order() -> Vec<&'static str> {
        let mut endpoints = Self::endpoints().to_vec();
        endpoints.shuffle(&mut rand::thread_rng());
        endpoints
    }

    fn normalized_google_lang(lang: &str) -> String {
        match lang.trim().to_ascii_lowercase().as_str() {
            "iw" => "he".to_string(),
            "jw" => "jv".to_string(),
            "zh-cn" | "zh_tw" | "zh-tw" => "zh".to_string(),
            "" => "auto".to_string(),
            other => other.to_string(),
        }
    }

    fn parse_result(response: Value) -> Result<String, TranslatorError> {
        let sentences = response["sentences"]
            .as_array()
            .ok_or(TranslatorError::MissingTranslatedText)?;
        let mut translated = String::new();

        for sentence in sentences {
            if let Some(text) = sentence["trans"].as_str() {
                translated.push_str(text);
            }
        }

        if translated.is_empty() && !sentences.is_empty() {
            return Err(TranslatorError::MissingTranslatedText);
        }

        Ok(translated)
    }

    async fn translate_with_endpoint(
        &self,
        endpoint: &str,
        text: &str,
    ) -> Result<String, TranslatorError> {
        let source_lang = Self::normalized_google_lang(&self.source_lang);
        let target_lang = Self::normalized_google_lang(&self.target_lang);

        let response = self
            .client
            .get(endpoint)
            .header(USER_AGENT, HeaderValue::from_static("Mozilla/5.0"))
            .query(&[
                ("client", "gtx"),
                ("sl", source_lang.as_str()),
                ("tl", target_lang.as_str()),
                ("dt", "t"),
                ("dj", "1"),
                ("q", text),
            ])
            .send()
            .await
            .map_err(|err| TranslatorError::Network(err.to_string()))?;

        let status = response.status();
        let body = response
            .text()
            .await
            .map_err(|err| TranslatorError::ReadResponse(err.to_string()))?;

        if !status.is_success() {
            return Err(TranslatorError::HttpStatus {
                status: status.as_u16(),
                body,
            });
        }

        let response_json: Value = serde_json::from_str(&body)
            .map_err(|err| TranslatorError::ParseJson(err.to_string()))?;

        Self::parse_result(response_json)
    }

    pub async fn translate(&self, text: &str) -> Result<String, TranslatorError> {
        let mut last_error = None;

        for endpoint in Self::endpoints_random_order() {
            match self.translate_with_endpoint(endpoint, text).await {
                Ok(translated) => return Ok(translated),
                Err(err) => last_error = Some(err),
            }
        }

        Err(last_error.unwrap_or(TranslatorError::MissingTranslatedText))
    }

    pub async fn translate_chunked_cancellable(
        &self,
        text: &str,
        cancel: Option<&AtomicBool>,
    ) -> Result<String, TranslatorError> {
        const MAX_CHUNK_CHARS: usize = 1_500;
        const MAX_CHUNK_ATTEMPTS: usize = 3;

        let chunks = TranslatorDeepLFree::split_translation_chunks(text, MAX_CHUNK_CHARS);
        if chunks.len() == 1 {
            if cancel.is_some_and(|flag| flag.load(Ordering::Relaxed)) {
                return Err(TranslatorError::Cancelled);
            }
            return self.translate(text).await;
        }

        let mut translated_text = String::new();

        for (index, chunk) in chunks.iter().enumerate() {
            if cancel.is_some_and(|flag| flag.load(Ordering::Relaxed)) {
                return Err(TranslatorError::Cancelled);
            }
            if index > 0 {
                let delay_seconds = rand::thread_rng().gen_range(1u64..=3u64);
                tokio::time::sleep(Duration::from_secs(delay_seconds)).await;
            }

            if chunk.trim().is_empty() {
                translated_text.push_str(chunk);
                continue;
            }

            let mut last_error = None;
            for attempt in 1..=MAX_CHUNK_ATTEMPTS {
                if cancel.is_some_and(|flag| flag.load(Ordering::Relaxed)) {
                    return Err(TranslatorError::Cancelled);
                }
                match self.translate(chunk).await {
                    Ok(translated_chunk) => {
                        translated_text.push_str(&translated_chunk);
                        last_error = None;
                        break;
                    }
                    Err(err) => {
                        last_error = Some(err);
                        if attempt < MAX_CHUNK_ATTEMPTS {
                            let delay_seconds = rand::thread_rng().gen_range(1u64..=3u64);
                            tokio::time::sleep(Duration::from_secs(delay_seconds)).await;
                        }
                    }
                }
            }

            if let Some(err) = last_error {
                return Err(err);
            }
        }

        Ok(translated_text)
    }
}

impl TranslatorDeepLFree {
    pub fn new(target_lang: impl Into<String>) -> Result<Self, TranslatorError> {
        let client = Client::builder()
            .no_gzip()
            .http1_only()
            .timeout(Duration::from_secs(10))
            .build()
            .map_err(|err| TranslatorError::HttpClientBuild(err.to_string()))?;

        Ok(Self {
            source_lang: None,
            target_lang: target_lang.into(),
            endpoint: "https://www2.deepl.com/jsonrpc?client=chrome-extension,1.5.1".to_string(),
            client,
        })
    }

    pub fn with_source_lang(
        target_lang: impl Into<String>,
        source_lang: impl Into<String>,
    ) -> Result<Self, TranslatorError> {
        let mut translator = Self::new(target_lang)?;
        translator.source_lang = Some(source_lang.into());
        Ok(translator)
    }

    fn vars(&self, text: &str) -> (u64, u64) {
        let uid = rand::thread_rng().gen_range(1_000_000_000u64..=9_999_999_999u64);

        let mut count_i = text.matches('i').count() as u64;

        let mut ts = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis() as u64;

        if count_i > 0 {
            count_i += 1;
            ts = ts - ts % count_i + count_i;
        }

        (uid, ts)
    }

    fn headers(&self) -> HeaderMap {
        let mut headers = HeaderMap::new();

        headers.insert(ACCEPT, HeaderValue::from_static("*/*"));
        headers.insert(ACCEPT_ENCODING, HeaderValue::from_static("gzip, deflate"));
        headers.insert(ACCEPT_LANGUAGE, HeaderValue::from_static("en-US,en;q=0.9"));
        headers.insert(AUTHORIZATION, HeaderValue::from_static("None"));

        headers.insert(
            HeaderName::from_static("authority"),
            HeaderValue::from_static("www2.deepl.com"),
        );

        headers.insert(
            CONTENT_TYPE,
            HeaderValue::from_static("application/json; charset=utf-8"),
        );

        headers.insert(
            USER_AGENT,
            HeaderValue::from_static(
                "DeepLBrowserExtension/1.5.1 Mozilla/5.0 (Macintosh; \
                 Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, \
                 like Gecko) Chrome/114.0.0.0 Safari/537.36",
            ),
        );

        headers.insert(
            ORIGIN,
            HeaderValue::from_static("chrome-extension://cofdbpoegempjloogbagkncekinflcnj"),
        );

        headers.insert(REFERER, HeaderValue::from_static("https://www.deepl.com/"));

        headers
    }

    fn source_code(&self) -> String {
        self.source_lang
            .clone()
            .unwrap_or_else(|| "auto".to_string())
    }

    fn build_body(&self, text: &str) -> String {
        let mut target_lang = self.target_lang.clone();

        let common_job_params = if target_lang.contains('-') {
            let portions: Vec<&str> = target_lang.split('-').collect();

            let variant = format!("{}-{}", portions[0].to_lowercase(), portions[1]);

            target_lang = portions[0].to_string();

            json!({
                "regionalVariant": variant
            })
        } else {
            json!({})
        };

        let (uid, ts) = self.vars(text);

        let body = json!({
            "jsonrpc": "2.0",
            "method": "LMT_handle_texts",
            "params": {
                "commonJobParams": common_job_params,
                "texts": [
                    {
                        "text": text
                    }
                ],
                "splitting": "newlines",
                "lang": {
                    "source_lang_user_selected": self.source_code(),
                    "target_lang": target_lang
                },
                "timestamp": ts
            },
            "id": uid
        });

        let body = body.to_string();

        if (uid + 3) % 13 == 0 || (uid + 5) % 29 == 0 {
            body.replace("\"method\":\"", "\"method\" : \"")
        } else {
            body.replace("\"method\":\"", "\"method\": \"")
        }
    }

    fn is_strong_chunk_boundary(ch: char) -> bool {
        matches!(
            ch,
            '\n' | '\r'
                | '.'
                | ','
                | ';'
                | ':'
                | '!'
                | '?'
                | '。'
                | '，'
                | '；'
                | '：'
                | '！'
                | '？'
        )
    }

    fn is_soft_chunk_boundary(ch: char) -> bool {
        ch.is_whitespace()
    }

    fn next_chunk_end(text: &str, start: usize, max_chars: usize) -> usize {
        let mut char_count = 0usize;
        let mut hard_end = text.len();
        let mut last_strong_boundary = None;
        let mut last_soft_boundary = None;

        for (offset, ch) in text[start..].char_indices() {
            char_count += 1;
            let end = start + offset + ch.len_utf8();

            if Self::is_strong_chunk_boundary(ch) {
                last_strong_boundary = Some(end);
            } else if Self::is_soft_chunk_boundary(ch) {
                last_soft_boundary = Some(end);
            }

            if char_count >= max_chars {
                hard_end = end;
                break;
            }
        }

        if hard_end == text.len() {
            return text.len();
        }

        last_strong_boundary
            .or(last_soft_boundary)
            .filter(|end| *end > start)
            .unwrap_or(hard_end)
    }

    fn split_translation_chunks(text: &str, max_chars: usize) -> Vec<String> {
        if text.chars().count() <= max_chars {
            return vec![text.to_string()];
        }

        let mut chunks = Vec::new();
        let mut start = 0usize;

        while start < text.len() {
            let end = Self::next_chunk_end(text, start, max_chars);
            chunks.push(text[start..end].to_string());
            start = end;
        }

        chunks
    }

    fn parse_result(response: &str) -> Result<String, TranslatorError> {
        let response_json: Value = serde_json::from_str(response)
            .map_err(|err| TranslatorError::ParseJson(err.to_string()))?;

        response_json["result"]["texts"][0]["text"]
            .as_str()
            .map(|text| text.to_string())
            .ok_or(TranslatorError::MissingTranslatedText)
    }

    pub async fn translate(&self, text: &str) -> Result<String, TranslatorError> {
        let response = self
            .client
            .post(&self.endpoint)
            .headers(self.headers())
            .body(self.build_body(text))
            .send()
            .await
            .map_err(|err| TranslatorError::Network(err.to_string()))?;

        let status = response.status();
        let is_gzip = response
            .headers()
            .get("Content-Encoding")
            .and_then(|value| value.to_str().ok())
            == Some("gzip");

        let bytes = response
            .bytes()
            .await
            .map_err(|err| TranslatorError::ReadResponse(err.to_string()))?;

        if !status.is_success() {
            let body = String::from_utf8_lossy(&bytes).to_string();

            return Err(TranslatorError::HttpStatus {
                status: status.as_u16(),
                body,
            });
        }

        let response_text = if is_gzip {
            let mut decoder = GzDecoder::new(&bytes[..]);
            let mut decoded = String::new();

            decoder
                .read_to_string(&mut decoded)
                .map_err(|err| TranslatorError::DecodeGzip(err.to_string()))?;

            decoded
        } else {
            String::from_utf8(bytes.to_vec())
                .map_err(|err| TranslatorError::DecodeUtf8(err.to_string()))?
        };

        Self::parse_result(&response_text)
    }

    pub async fn translate_chunked_cancellable(
        &self,
        text: &str,
        cancel: Option<&AtomicBool>,
    ) -> Result<String, TranslatorError> {
        const MAX_CHUNK_CHARS: usize = 3_000;

        let chunks = Self::split_translation_chunks(text, MAX_CHUNK_CHARS);
        if chunks.len() == 1 {
            if cancel.is_some_and(|flag| flag.load(Ordering::Relaxed)) {
                return Err(TranslatorError::Cancelled);
            }
            return self.translate(text).await;
        }

        let mut translated_text = String::new();

        for (index, chunk) in chunks.iter().enumerate() {
            if cancel.is_some_and(|flag| flag.load(Ordering::Relaxed)) {
                return Err(TranslatorError::Cancelled);
            }
            if index > 0 {
                let delay_seconds = rand::thread_rng().gen_range(1u64..=3u64);
                tokio::time::sleep(Duration::from_secs(delay_seconds)).await;
            }

            if chunk.trim().is_empty() {
                translated_text.push_str(chunk);
                continue;
            }

            translated_text.push_str(&self.translate(chunk).await?);
        }

        Ok(translated_text)
    }
}
