use crate::settings::Language;
use reqwest::Url;
use reqwest::blocking::Client;
use scraper::{ElementRef, Html, Selector};
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
struct ParseResponse {
    parse: ParsePage,
}

#[derive(Debug, Deserialize)]
struct ParsePage {
    title: String,
    text: String,
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

fn build_parse_url(lang: &str, pageid: i64) -> Result<Url, WikipediaError> {
    validate_lang_subdomain(lang)?;
    let base = format!("https://{lang}.wikipedia.org/w/api.php");
    let mut url = Url::parse(&base).map_err(|err| WikipediaError::Other(err.to_string()))?;
    url.query_pairs_mut()
        .append_pair("action", "parse")
        .append_pair("pageid", &pageid.to_string())
        .append_pair("prop", "text")
        .append_pair("disableeditsection", "1")
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
        .user_agent("Sonarpad/0.5 (Wikipedia import)")
        .build()
        .map_err(|e| WikipediaError::Other(e.to_string()))
}

pub fn language_to_code(language: Language) -> &'static str {
    match language {
        Language::Italian => "it",
        Language::Ukrainian | Language::English => "en",
        Language::Lithuanian => "lt",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::Swedish => "sv",
        Language::Vietnamese => "vi",
        Language::Czech => "cs",
        Language::Polish => "pl",
        Language::French => "fr",
        Language::Serbian => "sr",
        Language::Russian => "ru",
        Language::Chinese => "zh",
        Language::Hindi => "hi",
    }
}

pub fn resolve_language_code(language: Language, preference: &str) -> String {
    let pref = preference.trim().to_ascii_lowercase();
    if pref.is_empty() || pref == "auto" {
        return language_to_code(language).to_string();
    }
    match pref.as_str() {
        "it" | "en" | "es" | "pt" | "sv" | "vi" | "cs" | "pl" | "fr" | "ru" | "zh" => pref,
        _ => language_to_code(language).to_string(),
    }
}

fn html_fragment_to_text(html: &str) -> String {
    let mut out = String::new();
    let mut inside = false;
    let mut tag = String::new();
    let mut last_newline = false;
    let mut skip_stack: Vec<String> = Vec::new();
    let mut in_comment = false;

    for ch in html.chars() {
        if in_comment {
            tag.push(ch);
            if tag.ends_with("-->") {
                in_comment = false;
                tag.clear();
            }
            continue;
        }

        if inside {
            if ch == '>' {
                inside = false;
                let tag_trimmed = tag.trim();
                if tag_trimmed.starts_with("!--") {
                    if !tag_trimmed.ends_with("--") {
                        in_comment = true;
                    }
                    tag.clear();
                    continue;
                }

                let tag_name = tag_trimmed
                    .trim()
                    .trim_start_matches('/')
                    .split_whitespace()
                    .next()
                    .unwrap_or("")
                    .to_ascii_lowercase();
                let is_closing = tag_trimmed.starts_with('/');

                if matches!(
                    tag_name.as_str(),
                    "head"
                        | "style"
                        | "script"
                        | "title"
                        | "sup"
                        | "table"
                        | "figure"
                        | "figcaption"
                        | "noscript"
                ) {
                    if is_closing {
                        if let Some(pos) = skip_stack.iter().rposition(|t| t == &tag_name) {
                            skip_stack.truncate(pos);
                        }
                    } else {
                        skip_stack.push(tag_name.clone());
                    }
                    tag.clear();
                    continue;
                }
                if matches!(
                    tag_name.as_str(),
                    "br" | "p"
                        | "div"
                        | "li"
                        | "tr"
                        | "hr"
                        | "ul"
                        | "ol"
                        | "table"
                        | "blockquote"
                        | "dl"
                        | "dt"
                        | "dd"
                        | "h1"
                        | "h2"
                        | "h3"
                        | "h4"
                        | "h5"
                        | "h6"
                ) && skip_stack.is_empty()
                    && !last_newline
                    && !out.is_empty()
                {
                    out.push('\n');
                    last_newline = true;
                }
                tag.clear();
            } else {
                tag.push(ch);
            }
            continue;
        }
        if ch == '<' {
            inside = true;
            continue;
        }
        if !skip_stack.is_empty() {
            continue;
        }
        out.push(ch);
        last_newline = ch == '\n';
    }

    out.replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}

fn normalize_wikipedia_text_block(text: &str) -> String {
    let mut out = String::new();
    let mut blank_run = 0usize;
    for line in text.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            blank_run += 1;
            if blank_run <= 1 && !out.is_empty() {
                out.push('\n');
            }
            continue;
        }
        blank_run = 0;
        if !out.is_empty() && !out.ends_with('\n') {
            out.push('\n');
        }
        out.push_str(trimmed);
    }
    out.trim().to_string()
}

fn should_skip_parse_element(element: &ElementRef<'_>) -> bool {
    let name = element.value().name();
    if matches!(
        name,
        "table" | "style" | "script" | "figure" | "figcaption" | "noscript"
    ) {
        return true;
    }

    let classes = element.value().classes().collect::<Vec<_>>();
    classes.iter().any(|class_name| {
        matches!(
            *class_name,
            "mw-editsection"
                | "reference"
                | "reflist"
                | "navbox"
                | "vertical-navbox"
                | "authority-control"
                | "metadata"
                | "infobox"
                | "sinottico"
                | "thumb"
                | "tright"
                | "tleft"
                | "toc"
                | "hatnote"
                | "ambox"
                | "sistersitebox"
                | "mw-empty-elt"
        )
    })
}

fn heading_text(name: &str, text: &str) -> String {
    let marks = match name {
        "h2" => "==",
        "h3" => "===",
        "h4" => "====",
        "h5" => "=====",
        "h6" => "======",
        _ => "",
    };
    if marks.is_empty() {
        text.to_string()
    } else {
        format!("{marks} {text} {marks}")
    }
}

fn parse_article_html_to_text(html: &str) -> String {
    let document = Html::parse_fragment(html);
    let selector = match Selector::parse("div.mw-parser-output") {
        Ok(selector) => selector,
        Err(_) => return normalize_wikipedia_text_block(&html_fragment_to_text(html)),
    };
    let Some(container) = document.select(&selector).next() else {
        return normalize_wikipedia_text_block(&html_fragment_to_text(html));
    };

    let mut blocks = Vec::new();
    for child in container.children() {
        let Some(element) = ElementRef::wrap(child) else {
            continue;
        };
        if should_skip_parse_element(&element) {
            continue;
        }

        let name = element.value().name();
        if !matches!(
            name,
            "p" | "ul" | "ol" | "dl" | "div" | "blockquote" | "h2" | "h3" | "h4" | "h5" | "h6"
        ) {
            continue;
        }

        let text = normalize_wikipedia_text_block(&html_fragment_to_text(&element.html()));
        if text.is_empty() {
            continue;
        }

        if name.starts_with('h') {
            blocks.push(heading_text(name, &text));
        } else {
            blocks.push(text);
        }
    }

    if blocks.is_empty() {
        normalize_wikipedia_text_block(&html_fragment_to_text(html))
    } else {
        blocks.join("\n\n")
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
    let url = build_parse_url(lang, pageid)?;
    let client = http_client()?;
    let resp = client
        .get(url)
        .send()
        .map_err(|e| WikipediaError::Other(e.to_string()))?;
    let value: Value = resp
        .json()
        .map_err(|e| WikipediaError::Other(e.to_string()))?;
    let parsed: ParseResponse = parse_or_error(value)?;
    let title = parsed.parse.title;
    let extract = parse_article_html_to_text(&parsed.parse.text);
    if title.trim().is_empty() || extract.trim().is_empty() {
        return Err(WikipediaError::NotFound);
    }
    let url = build_article_url(lang, &title)?;
    Ok(ExtractResult { extract, url })
}

#[cfg(test)]
mod tests {
    use super::parse_article_html_to_text;

    #[test]
    fn wikipedia_html_parser_keeps_quote_after_intro_line() {
        let html = r#"
        <div class="mw-parser-output">
          <p>Pertini incrociò Ginzburg mentre lo riportavano in cella dopo un feroce pestaggio, e in quell'occasione quegli trovò la forza di sussurrargli:</p>
          <div class="itwiki-template-citazione">
            <div class="itwiki-template-citazione-singola"><p>«Guai se alla fine della guerra dovessimo incolpare tutto il popolo tedesco per la malvagità di pochi.»</p></div>
          </div>
          <p>Anche don Morosini fu visto da Pertini dopo un interrogatorio delle SS.</p>
        </div>
        "#;

        let text = parse_article_html_to_text(html);

        assert!(text.contains("trovò la forza di sussurrargli:"));
        assert!(text.contains("«Guai se alla fine della guerra dovessimo incolpare tutto il popolo tedesco per la malvagità di pochi.»"));
        assert!(
            text.contains(
                "Anche don Morosini fu visto da Pertini dopo un interrogatorio delle SS."
            )
        );
    }
}
