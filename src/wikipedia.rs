use crate::settings::Language;
use reqwest::Url;
use reqwest::blocking::Client;
use serde::Deserialize;
use serde_json::Value;
use std::fmt;
use std::time::Duration;

#[derive(Debug, Clone)]
pub struct SearchResult {
    pub pageid: i64,
    pub title: String,
}

#[derive(Debug, Clone)]
pub struct ExtractResult {
    pub extract: String,
    pub url: String,
}

#[derive(Debug)]
pub enum WikipediaError {
    NotFound,
    Api { code: String, info: String },
    Other(String),
}

impl fmt::Display for WikipediaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            WikipediaError::NotFound => write!(f, "Wikipedia page not found"),
            WikipediaError::Api { code, info } => write!(f, "MediaWiki API error ({code}): {info}"),
            WikipediaError::Other(message) => write!(f, "{message}"),
        }
    }
}

impl std::error::Error for WikipediaError {}

#[derive(Debug, Deserialize)]
struct MwErrorEnvelope {
    error: MwError,
}

#[derive(Debug, Deserialize)]
struct MwError {
    code: String,
    info: String,
}

#[derive(Debug, Deserialize)]
struct SearchResponse {
    query: SearchQuery,
}

#[derive(Debug, Deserialize)]
struct SearchQuery {
    search: Vec<SearchHit>,
}

#[derive(Debug, Deserialize)]
struct SearchHit {
    pageid: i64,
    title: String,
}

#[derive(Debug, Deserialize)]
struct ExtractResponse {
    query: ExtractQuery,
}

#[derive(Debug, Deserialize)]
struct ExtractQuery {
    pages: Vec<ExtractPage>,
}

#[derive(Debug, Deserialize)]
struct ExtractPage {
    title: Option<String>,
    extract: Option<String>,
    missing: Option<bool>,
}

fn validate_lang_subdomain(lang: &str) -> Result<(), WikipediaError> {
    if lang.is_empty() || !lang.chars().all(|c| c.is_ascii_alphanumeric() || c == '-') {
        return Err(WikipediaError::Other(format!(
            "Invalid Wikipedia language code: {lang}"
        )));
    }
    Ok(())
}

fn build_search_url(lang: &str, query: &str, limit: usize) -> Result<Url, WikipediaError> {
    validate_lang_subdomain(lang)?;
    let base = format!("https://{lang}.wikipedia.org/w/api.php");
    let mut url = Url::parse(&base).map_err(|err| WikipediaError::Other(err.to_string()))?;
    url.query_pairs_mut()
        .append_pair("action", "query")
        .append_pair("list", "search")
        .append_pair("srsearch", query)
        .append_pair("srlimit", &limit.to_string())
        .append_pair("format", "json")
        .append_pair("formatversion", "2");
    Ok(url)
}

fn build_extract_url(lang: &str, pageid: i64) -> Result<Url, WikipediaError> {
    validate_lang_subdomain(lang)?;
    let base = format!("https://{lang}.wikipedia.org/w/api.php");
    let mut url = Url::parse(&base).map_err(|err| WikipediaError::Other(err.to_string()))?;
    url.query_pairs_mut()
        .append_pair("action", "query")
        .append_pair("prop", "extracts")
        .append_pair("explaintext", "1")
        .append_pair("pageids", &pageid.to_string())
        .append_pair("redirects", "1")
        .append_pair("format", "json")
        .append_pair("formatversion", "2");
    Ok(url)
}

fn build_article_url(lang: &str, title: &str) -> Result<String, WikipediaError> {
    validate_lang_subdomain(lang)?;
    let base = format!("https://{lang}.wikipedia.org/wiki/");
    let mut url = Url::parse(&base).map_err(|err| WikipediaError::Other(err.to_string()))?;
    let mut path = String::from("/wiki/");
    let normalized = title.replace(' ', "_");
    let encoded: String = url::form_urlencoded::byte_serialize(normalized.as_bytes()).collect();
    path.push_str(&encoded);
    url.set_path(&path);
    Ok(url.to_string())
}

fn parse_or_error<T: for<'de> Deserialize<'de>>(value: Value) -> Result<T, WikipediaError> {
    if value.get("error").is_some() {
        let err: MwErrorEnvelope =
            serde_json::from_value(value).map_err(|e| WikipediaError::Other(e.to_string()))?;
        return Err(WikipediaError::Api {
            code: err.error.code,
            info: err.error.info,
        });
    }
    serde_json::from_value(value).map_err(|e| WikipediaError::Other(e.to_string()))
}

fn http_client() -> Result<Client, WikipediaError> {
    Client::builder()
        .timeout(Duration::from_secs(10))
        .user_agent("Novapad/0.5 (Wikipedia import)")
        .build()
        .map_err(|e| WikipediaError::Other(e.to_string()))
}

pub fn language_to_code(language: Language) -> &'static str {
    match language {
        Language::Italian => "it",
        Language::English => "en",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::Vietnamese => "vi",
    }
}

pub fn resolve_language_code(language: Language, preference: &str) -> String {
    let pref = preference.trim().to_ascii_lowercase();
    if pref.is_empty() || pref == "auto" {
        return language_to_code(language).to_string();
    }
    match pref.as_str() {
        "it" | "en" | "es" | "pt" | "vi" => pref,
        _ => language_to_code(language).to_string(),
    }
}

pub fn search_articles(
    lang: &str,
    query: &str,
    limit: usize,
) -> Result<Vec<SearchResult>, WikipediaError> {
    let trimmed = query.trim();
    if trimmed.is_empty() {
        return Ok(Vec::new());
    }
    let url = build_search_url(lang, trimmed, limit)?;
    let client = http_client()?;
    let resp = client
        .get(url)
        .send()
        .map_err(|e| WikipediaError::Other(e.to_string()))?;
    let value: Value = resp
        .json()
        .map_err(|e| WikipediaError::Other(e.to_string()))?;
    let parsed: SearchResponse = parse_or_error(value)?;
    let results = parsed
        .query
        .search
        .into_iter()
        .filter(|hit| !hit.title.trim().is_empty())
        .map(|hit| SearchResult {
            pageid: hit.pageid,
            title: hit.title,
        })
        .collect();
    Ok(results)
}

pub fn fetch_extract(lang: &str, pageid: i64) -> Result<ExtractResult, WikipediaError> {
    let url = build_extract_url(lang, pageid)?;
    let client = http_client()?;
    let resp = client
        .get(url)
        .send()
        .map_err(|e| WikipediaError::Other(e.to_string()))?;
    let value: Value = resp
        .json()
        .map_err(|e| WikipediaError::Other(e.to_string()))?;
    let parsed: ExtractResponse = parse_or_error(value)?;
    let Some(page) = parsed.query.pages.into_iter().next() else {
        return Err(WikipediaError::NotFound);
    };
    if page.missing.unwrap_or(false) {
        return Err(WikipediaError::NotFound);
    }
    let title = page.title.unwrap_or_default();
    let extract = page.extract.unwrap_or_default();
    if title.trim().is_empty() || extract.trim().is_empty() {
        return Err(WikipediaError::NotFound);
    }
    let url = build_article_url(lang, &title)?;
    Ok(ExtractResult { extract, url })
}
