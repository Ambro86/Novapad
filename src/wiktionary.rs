use crate::i18n;
use crate::settings::Language;
use reqwest::Url;
use reqwest::blocking::Client;
use serde::Deserialize;
use serde_json::Value;
use std::fmt;
use std::time::Duration;

const MAX_PAGE_LENGTH: usize = 30000;
const LARGE_PAGE_MAX_DEFS: usize = 5;
const LARGE_PAGE_MAX_LINES: usize = 200;
const MAX_CHARS_PER_DEF: usize = 500;

#[derive(Debug, Clone)]
pub struct LookupOutput {
    pub lang: String,
    pub word: String,
    pub definitions: Vec<String>,
    pub synonyms: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct DictionaryAndTranslation {
    pub dictionary: LookupOutput,
    pub translation: Option<LookupOutput>,
}

#[derive(Debug)]
pub enum LookupError {
    NotFound { lang: String, word: String },
    Api { code: String, info: String },
    Other(String),
}

impl fmt::Display for LookupError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            LookupError::NotFound { lang, word } => {
                write!(f, "Word not found: {word} (lang={lang})")
            }
            LookupError::Api { code, info } => {
                write!(f, "MediaWiki API error ({code}): {info}")
            }
            LookupError::Other(message) => write!(f, "{message}"),
        }
    }
}

impl std::error::Error for LookupError {}

#[derive(Debug, Deserialize)]
struct MwParseResponse {
    parse: MwParse,
}

#[derive(Debug, Deserialize)]
struct MwParse {
    wikitext: String,
}

#[derive(Debug, Deserialize)]
struct MwInfoResponse {
    query: MwInfoQuery,
}

#[derive(Debug, Deserialize)]
struct MwInfoQuery {
    pages: Vec<MwInfoPage>,
}

#[derive(Debug, Deserialize)]
struct MwInfoPage {
    pageid: Option<i64>,
    missing: Option<bool>,
    length: Option<usize>,
}

#[derive(Debug, Deserialize)]
struct MwErrorEnvelope {
    error: MwError,
}

#[derive(Debug, Deserialize)]
struct MwError {
    code: String,
    info: String,
}

fn validate_lang_subdomain(lang: &str) -> Result<(), LookupError> {
    if lang.is_empty() || !lang.chars().all(|c| c.is_ascii_alphanumeric() || c == '-') {
        return Err(LookupError::Other(format!(
            "Invalid Wiktionary language code: {lang}"
        )));
    }
    Ok(())
}

fn build_parse_wikitext_url(lang: &str, word: &str) -> Result<Url, LookupError> {
    validate_lang_subdomain(lang)?;
    let base = format!("https://{lang}.wiktionary.org/w/api.php");
    let mut url = Url::parse(&base).map_err(|err| LookupError::Other(err.to_string()))?;
    url.query_pairs_mut()
        .append_pair("action", "parse")
        .append_pair("page", word)
        .append_pair("prop", "wikitext")
        .append_pair("section", "1")
        .append_pair("format", "json")
        .append_pair("formatversion", "2");
    Ok(url)
}

fn build_page_info_url(lang: &str, word: &str) -> Result<Url, LookupError> {
    validate_lang_subdomain(lang)?;
    let base = format!("https://{lang}.wiktionary.org/w/api.php");
    let mut url = Url::parse(&base).map_err(|err| LookupError::Other(err.to_string()))?;
    url.query_pairs_mut()
        .append_pair("action", "query")
        .append_pair("prop", "info")
        .append_pair("titles", word)
        .append_pair("format", "json")
        .append_pair("formatversion", "2");
    Ok(url)
}
fn strip_links(mut s: String) -> String {
    while let Some(start) = s.find("[[") {
        let Some(end) = s[start + 2..].find("]]").map(|i| i + start + 2) else {
            break;
        };
        let inner = &s[start + 2..end];
        let replacement = inner.split('|').next_back().unwrap_or(inner).to_string();
        s.replace_range(start..end + 2, &replacement);
    }
    s
}

fn strip_external_links(mut s: String) -> String {
    while let Some(start) = s.find('[') {
        if s[start..].starts_with("[[") {
            let Some(end) = s[start + 2..].find("]]").map(|i| i + start + 2) else {
                break;
            };
            let inner = &s[start..end + 2];
            let rest = &s[end + 2..];
            let mut out = String::with_capacity(s.len());
            out.push_str(&s[..start]);
            out.push_str(inner);
            out.push_str(rest);
            s = out;
            continue;
        }
        let Some(end) = s[start + 1..].find(']').map(|i| i + start + 1) else {
            break;
        };
        let inner = &s[start + 1..end];
        let is_url = inner.starts_with("http://")
            || inner.starts_with("https://")
            || inner.starts_with("www.");
        let replacement = if is_url {
            inner
                .split_whitespace()
                .skip(1)
                .collect::<Vec<_>>()
                .join(" ")
        } else {
            inner.to_string()
        };
        s.replace_range(start..end + 1, &replacement);
    }
    s
}

fn strip_html_comments(mut s: String) -> String {
    while let Some(start) = s.find("<!--") {
        let Some(end) = s[start + 4..].find("-->").map(|i| i + start + 4) else {
            break;
        };
        s.replace_range(start..end + 3, "");
    }
    s
}

fn strip_templates(s: String) -> String {
    if !s.contains("{{") {
        return s;
    }
    let mut out = String::with_capacity(s.len());
    let mut depth = 0usize;
    let mut i = 0usize;
    while i < s.len() {
        if s[i..].starts_with("{{") {
            depth += 1;
            i += 2;
            continue;
        }
        if s[i..].starts_with("}}") && depth > 0 {
            depth -= 1;
            i += 2;
            continue;
        }
        let ch = s[i..].chars().next().unwrap_or('\0');
        if depth == 0 && ch != '\0' {
            out.push(ch);
        }
        i += ch.len_utf8().max(1);
    }
    out
}

fn clean_wikitext_line(s: String) -> String {
    let mut x = strip_html_comments(s);
    x = strip_links(x);
    x = strip_external_links(x);
    x = strip_templates(x);
    x = x.replace("''", "");
    x.split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .trim()
        .to_string()
}

fn extract_definitions_with_subpoints(
    wikitext: &str,
    max_main_defs: usize,
    max_total_lines: usize,
) -> Vec<String> {
    let cleaned_text = strip_templates(wikitext.to_string());
    let mut out = Vec::new();
    let mut main_count = 0;

    for line in cleaned_text.lines() {
        if out.len() >= max_total_lines {
            break;
        }
        let l = line.trim_start();

        if l.starts_with("# ") && main_count < max_main_defs {
            let cleaned = clean_wikitext_line(l[2..].trim().to_string());
            let truncated = cleaned.chars().take(MAX_CHARS_PER_DEF).collect::<String>();
            if !truncated.is_empty() {
                out.push(truncated);
                main_count += 1;
            }
            if main_count >= max_main_defs {
                break;
            }
        }
    }

    out
}

fn extract_synonyms(wikitext: &str, max_syns: usize) -> Vec<String> {
    let mut out = Vec::new();
    let start_pos = match wikitext.find("{{-sin-}}") {
        Some(p) => p,
        None => return out,
    };
    let after_start = &wikitext[start_pos + "{{-sin-}}".len()..];
    let end_pos = after_start.find("\n{{-").unwrap_or(after_start.len());
    let block = &after_start[..end_pos];
    let block = strip_templates(block.to_string());

    for line in block.lines() {
        if line.trim_start().starts_with('*') {
            let cleaned = clean_wikitext_line(line.trim_start()[1..].trim().to_string());
            if !cleaned.is_empty() {
                out.push(cleaned);
                if out.len() >= max_syns {
                    break;
                }
            }
        }
    }

    out
}

pub struct WiktionaryService {
    client: Client,
}

impl WiktionaryService {
    pub fn new() -> Result<Self, LookupError> {
        let client = Client::builder()
            .user_agent("Novapad/0.5 (Wiktionary dictionary)")
            .timeout(Duration::from_secs(4))
            .build()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        Ok(Self { client })
    }

    fn fetch_page_length(&self, lang: &str, word: &str) -> Result<usize, LookupError> {
        let url = build_page_info_url(lang, word)?;
        let resp = self
            .client
            .get(url)
            .send()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        if !resp.status().is_success() {
            return Err(LookupError::Other(format!(
                "Wiktionary HTTP error: {}",
                resp.status()
            )));
        }

        let v: Value = resp
            .json()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        if v.get("error").is_some() {
            let env: MwErrorEnvelope =
                serde_json::from_value(v).map_err(|err| LookupError::Other(err.to_string()))?;
            if env.error.code == "missingtitle" {
                return Err(LookupError::NotFound {
                    lang: lang.to_string(),
                    word: word.to_string(),
                });
            }
            return Err(LookupError::Api {
                code: env.error.code,
                info: env.error.info,
            });
        }

        let parsed: MwInfoResponse =
            serde_json::from_value(v).map_err(|err| LookupError::Other(err.to_string()))?;
        let page = parsed.query.pages.first().ok_or_else(|| {
            LookupError::Other("Wiktionary response missing page info".to_string())
        })?;
        if page.missing.unwrap_or(false) || page.pageid == Some(-1) {
            return Err(LookupError::NotFound {
                lang: lang.to_string(),
                word: word.to_string(),
            });
        }
        Ok(page.length.unwrap_or(0))
    }

    fn fetch_section1_wikitext(&self, lang: &str, word: &str) -> Result<String, LookupError> {
        let url = build_parse_wikitext_url(lang, word)?;
        let resp = self
            .client
            .get(url)
            .send()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        if !resp.status().is_success() {
            return Err(LookupError::Other(format!(
                "Wiktionary HTTP error: {}",
                resp.status()
            )));
        }

        let v: Value = resp
            .json()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        if v.get("error").is_some() {
            let env: MwErrorEnvelope =
                serde_json::from_value(v).map_err(|err| LookupError::Other(err.to_string()))?;
            if env.error.code == "missingtitle" {
                return Err(LookupError::NotFound {
                    lang: lang.to_string(),
                    word: word.to_string(),
                });
            }
            return Err(LookupError::Api {
                code: env.error.code,
                info: env.error.info,
            });
        }

        let parsed: MwParseResponse =
            serde_json::from_value(v).map_err(|err| LookupError::Other(err.to_string()))?;
        Ok(parsed.parse.wikitext)
    }

    fn fetch_section1_wikitext_with_timeout(
        &self,
        lang: &str,
        word: &str,
        timeout: Duration,
    ) -> Result<String, LookupError> {
        let url = build_parse_wikitext_url(lang, word)?;
        let client = Client::builder()
            .user_agent("Novapad/0.5 (Wiktionary dictionary)")
            .timeout(timeout)
            .build()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        let resp = client
            .get(url)
            .send()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        if !resp.status().is_success() {
            return Err(LookupError::Other(format!(
                "Wiktionary HTTP error: {}",
                resp.status()
            )));
        }

        let v: Value = resp
            .json()
            .map_err(|err| LookupError::Other(err.to_string()))?;
        if v.get("error").is_some() {
            let env: MwErrorEnvelope =
                serde_json::from_value(v).map_err(|err| LookupError::Other(err.to_string()))?;
            if env.error.code == "missingtitle" {
                return Err(LookupError::NotFound {
                    lang: lang.to_string(),
                    word: word.to_string(),
                });
            }
            return Err(LookupError::Api {
                code: env.error.code,
                info: env.error.info,
            });
        }

        let parsed: MwParseResponse =
            serde_json::from_value(v).map_err(|err| LookupError::Other(err.to_string()))?;
        Ok(parsed.parse.wikitext)
    }

    fn dictionary_lookup_with_length(
        &self,
        dictionary_lang: &str,
        word: &str,
    ) -> Result<(LookupOutput, usize), LookupError> {
        let length = self.fetch_page_length(dictionary_lang, word)?;
        let wikitext = if length > MAX_PAGE_LENGTH {
            self.fetch_section1_wikitext_with_timeout(
                dictionary_lang,
                word,
                Duration::from_secs(8),
            )?
        } else {
            self.fetch_section1_wikitext(dictionary_lang, word)?
        };
        let (max_defs, max_lines) = if length > MAX_PAGE_LENGTH {
            (LARGE_PAGE_MAX_DEFS, LARGE_PAGE_MAX_LINES)
        } else {
            (usize::MAX, usize::MAX)
        };
        let defs = extract_definitions_with_subpoints(&wikitext, max_defs, max_lines);
        let syns = if length > MAX_PAGE_LENGTH {
            Vec::new()
        } else {
            extract_synonyms(&wikitext, usize::MAX)
        };

        if defs.is_empty() {
            return Err(LookupError::NotFound {
                lang: dictionary_lang.to_string(),
                word: word.to_string(),
            });
        }

        let output = LookupOutput {
            lang: dictionary_lang.to_string(),
            word: word.to_string(),
            definitions: defs,
            synonyms: syns,
        };
        Ok((output, length))
    }

    pub fn dictionary_lookup(
        &self,
        dictionary_lang: &str,
        word: &str,
    ) -> Result<LookupOutput, LookupError> {
        self.dictionary_lookup_with_length(dictionary_lang, word)
            .map(|(output, _)| output)
    }

    pub fn translate_word(
        &self,
        target_lang: &str,
        word: &str,
    ) -> Result<LookupOutput, LookupError> {
        let length = self.fetch_page_length(target_lang, word)?;
        let wikitext = if length > MAX_PAGE_LENGTH {
            self.fetch_section1_wikitext_with_timeout(target_lang, word, Duration::from_secs(8))?
        } else {
            self.fetch_section1_wikitext(target_lang, word)?
        };
        let (max_defs, max_lines) = if length > MAX_PAGE_LENGTH {
            (LARGE_PAGE_MAX_DEFS, LARGE_PAGE_MAX_LINES)
        } else {
            (usize::MAX, usize::MAX)
        };
        let defs = extract_definitions_with_subpoints(&wikitext, max_defs, max_lines);
        if defs.is_empty() {
            return Err(LookupError::NotFound {
                lang: target_lang.to_string(),
                word: word.to_string(),
            });
        }
        Ok(LookupOutput {
            lang: target_lang.to_string(),
            word: word.to_string(),
            definitions: defs,
            synonyms: Vec::new(),
        })
    }

    pub fn dictionary_and_translation(
        &self,
        dictionary_lang: &str,
        target_lang: Option<&str>,
        word: &str,
    ) -> Result<DictionaryAndTranslation, LookupError> {
        let dict = self.dictionary_lookup(dictionary_lang, word)?;
        let translation = match target_lang {
            None => None,
            Some(t) if t.eq_ignore_ascii_case(dictionary_lang) => None,
            Some(t) => match self.translate_word(t, word) {
                Ok(x) => Some(x),
                Err(LookupError::NotFound { .. }) => None,
                Err(err) => return Err(err),
            },
        };
        Ok(DictionaryAndTranslation {
            dictionary: dict,
            translation,
        })
    }

    pub fn dictionary_and_translation_with_meta(
        &self,
        dictionary_lang: &str,
        target_lang: Option<&str>,
        word: &str,
    ) -> Result<(DictionaryAndTranslation, bool), LookupError> {
        let (dict, length) = self.dictionary_lookup_with_length(dictionary_lang, word)?;
        let translation = match target_lang {
            None => None,
            Some(t) if t.eq_ignore_ascii_case(dictionary_lang) => None,
            Some(t) => match self.translate_word(t, word) {
                Ok(x) => Some(x),
                Err(LookupError::NotFound { .. }) => None,
                Err(err) => return Err(err),
            },
        };
        Ok((
            DictionaryAndTranslation {
                dictionary: dict,
                translation,
            },
            length > MAX_PAGE_LENGTH,
        ))
    }
}

fn language_to_code(language: Language) -> &'static str {
    match language {
        Language::Italian => "it",
        Language::English => "en",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::Vietnamese => "vi",
    }
}

fn translation_target(language: Language, preference: &str) -> Option<String> {
    let pref = preference.trim().to_ascii_lowercase();
    if pref.is_empty() || pref == "auto" {
        return match language {
            Language::English => None,
            _ => Some("en".to_string()),
        };
    }
    if pref == "none" {
        return None;
    }
    let code = match pref.as_str() {
        "it" => "it",
        "en" => "en",
        "es" => "es",
        "pt" => "pt",
        "vi" => "vi",
        _ => {
            return match language {
                Language::English => None,
                _ => Some("en".to_string()),
            };
        }
    };
    let dict_lang = language_to_code(language);
    if code.eq_ignore_ascii_case(dict_lang) {
        None
    } else {
        Some(code.to_string())
    }
}

pub fn lookup_for_language(
    word: &str,
    language: Language,
    translation_preference: &str,
) -> Result<DictionaryAndTranslation, LookupError> {
    let trimmed = word.trim();
    if trimmed.is_empty() {
        return Err(LookupError::Other("Empty word".to_string()));
    }
    let svc = WiktionaryService::new()?;
    let dict_lang = language_to_code(language);
    let target_lang = translation_target(language, translation_preference);
    svc.dictionary_and_translation(dict_lang, target_lang.as_deref(), trimmed)
}

pub fn lookup_for_language_with_meta(
    word: &str,
    language: Language,
    translation_preference: &str,
) -> Result<(DictionaryAndTranslation, bool), LookupError> {
    let trimmed = word.trim();
    if trimmed.is_empty() {
        return Err(LookupError::Other("Empty word".to_string()));
    }
    let svc = WiktionaryService::new()?;
    let dict_lang = language_to_code(language);
    let target_lang = translation_target(language, translation_preference);
    svc.dictionary_and_translation_with_meta(dict_lang, target_lang.as_deref(), trimmed)
}

fn push_definitions_menu(lines: &mut Vec<String>, definitions: &[String]) {
    for def in definitions {
        if def.starts_with("- ") {
            let trimmed = def.trim_start_matches("- ").trim();
            lines.push(trimmed.to_string());
        } else {
            lines.push(def.to_string());
        }
    }
}

pub fn format_output_text(language: Language, entry: &DictionaryAndTranslation) -> String {
    let mut out = String::new();
    let title = i18n::tr_f(
        language,
        "dictionary.word_label",
        &[("word", &entry.dictionary.word)],
    );
    out.push_str(&title);
    out.push_str("\n\n");

    out.push_str(&i18n::tr(language, "dictionary.definitions"));
    out.push_str(":\n");
    for line in format_definition_lines(&entry.dictionary.definitions) {
        out.push_str(&line);
        out.push('\n');
    }

    if !entry.dictionary.synonyms.is_empty() {
        out.push('\n');
        out.push_str(&i18n::tr(language, "dictionary.synonyms"));
        out.push_str(":\n");
        for syn in &entry.dictionary.synonyms {
            out.push_str(syn);
            out.push('\n');
        }
    }

    if let Some(translation) = &entry.translation {
        out.push('\n');
        let label = i18n::tr_f(
            language,
            "dictionary.translation_label",
            &[("lang", &translation.lang)],
        );
        out.push_str(&label);
        out.push('\n');
        for line in format_definition_lines(&translation.definitions) {
            out.push_str(&line);
            out.push('\n');
        }
    }

    out
}

pub fn format_menu_lines(language: Language, entry: &DictionaryAndTranslation) -> Vec<String> {
    let mut lines = Vec::new();
    lines.push(i18n::tr_f(
        language,
        "dictionary.word_label",
        &[("word", &entry.dictionary.word)],
    ));
    lines.push(i18n::tr(language, "dictionary.definitions"));
    push_definitions_menu(&mut lines, &entry.dictionary.definitions);

    if !entry.dictionary.synonyms.is_empty() {
        lines.push(i18n::tr(language, "dictionary.synonyms"));
        for syn in &entry.dictionary.synonyms {
            lines.push(syn.clone());
        }
    }

    if let Some(translation) = &entry.translation {
        let label = i18n::tr_f(
            language,
            "dictionary.translation_label",
            &[("lang", &translation.lang)],
        );
        lines.push(label);
        push_definitions_menu(&mut lines, &translation.definitions);
    }

    lines
        .into_iter()
        .map(|line| line.trim().to_string())
        .filter(|line| !line.is_empty())
        .collect()
}

fn format_definition_lines(definitions: &[String]) -> Vec<String> {
    let mut lines = Vec::new();
    let mut main_index = 1;
    for def in definitions {
        if def.starts_with("- ") {
            let trimmed = def.trim_start_matches("- ").trim();
            lines.push(trimmed.to_string());
        } else {
            lines.push(format!("{main_index}. {def}"));
            main_index += 1;
        }
    }
    lines
}
