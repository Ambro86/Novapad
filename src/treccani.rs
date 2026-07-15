use reqwest::Url;
use reqwest::blocking::Client;
use scraper::{ElementRef, Html, Selector};
use serde_json::Value;
use std::collections::HashSet;
use std::fmt;
use std::time::Duration;

const TRECCANI_BASE_URL: &str = "https://www.treccani.it";
const TRECCANI_SEARCH_BASE: &str = "https://www.treccani.it/enciclopedia/ricerca/";

#[derive(Debug, Clone)]
pub struct SearchResult {
    pub title: String,
    pub url: String,
    pub description: String,
}

#[derive(Debug, Clone)]
pub struct ExtractResult {
    pub extract: String,
    pub url: String,
    pub sections: Vec<ArticleSection>,
}

#[derive(Debug, Clone)]
pub struct ArticleSection {
    pub title: String,
    pub level: usize,
    pub text: String,
}

#[derive(Debug)]
pub enum TreccaniError {
    NotFound,
    InvalidUrl,
    Other(String),
}

impl fmt::Display for TreccaniError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TreccaniError::NotFound => write!(f, "Voce Treccani non trovata"),
            TreccaniError::InvalidUrl => write!(f, "Indirizzo Treccani non valido"),
            TreccaniError::Other(message) => write!(f, "{message}"),
        }
    }
}

impl std::error::Error for TreccaniError {}

fn http_client() -> Result<Client, TreccaniError> {
    Client::builder()
        .timeout(Duration::from_secs(20))
        .user_agent("Sonarpad/0.8.1 (Treccani encyclopedia reader)")
        .build()
        .map_err(|err| TreccaniError::Other(err.to_string()))
}

fn build_search_url(query: &str) -> Result<Url, TreccaniError> {
    let mut url =
        Url::parse(TRECCANI_SEARCH_BASE).map_err(|err| TreccaniError::Other(err.to_string()))?;
    url.path_segments_mut()
        .map_err(|_| TreccaniError::InvalidUrl)?
        .push(query)
        .push("");
    Ok(url)
}

fn normalize_result_url(raw_url: &str) -> Option<String> {
    let base = Url::parse(TRECCANI_BASE_URL).ok()?;
    let url = base.join(raw_url.trim()).ok()?;
    let host = url.host_str()?.to_ascii_lowercase();
    if host != "treccani.it" && !host.ends_with(".treccani.it") {
        return None;
    }
    if !matches!(url.scheme(), "http" | "https") {
        return None;
    }
    let path = url.path().to_ascii_lowercase();
    if !path.starts_with("/enciclopedia/") || path.contains("/enciclopedia/ricerca/") {
        return None;
    }
    Some(url.to_string())
}

fn find_matches_array(value: &Value) -> Option<&Vec<Value>> {
    match value {
        Value::Object(map) => {
            if let Some(Value::Array(matches)) = map.get("matches") {
                return Some(matches);
            }
            map.values().find_map(find_matches_array)
        }
        Value::Array(items) => items.iter().find_map(find_matches_array),
        _ => None,
    }
}

fn parse_search_results(html: &str, limit: usize) -> Vec<SearchResult> {
    let document = Html::parse_document(html);
    let selector = match Selector::parse("script[type=\"application/json\"]") {
        Ok(selector) => selector,
        Err(_) => return Vec::new(),
    };

    let mut results = Vec::new();
    let mut seen_urls = HashSet::new();

    for script in document.select(&selector) {
        let json_text = script.text().collect::<String>();
        let Ok(value) = serde_json::from_str::<Value>(json_text.trim()) else {
            continue;
        };
        let Some(matches) = find_matches_array(&value) else {
            continue;
        };

        for item in matches {
            let Some(object) = item.as_object() else {
                continue;
            };
            let raw_url = object.get("url").and_then(Value::as_str).unwrap_or("");
            let Some(url) = normalize_result_url(raw_url) else {
                continue;
            };
            if !seen_urls.insert(url.clone()) {
                continue;
            }

            let title = object
                .get("title")
                .and_then(Value::as_str)
                .unwrap_or("")
                .trim();
            if title.is_empty() {
                continue;
            }
            let section = object
                .get("section")
                .and_then(Value::as_str)
                .unwrap_or("")
                .trim()
                .to_ascii_lowercase();
            if !section.is_empty() && section != "enciclopedia" {
                continue;
            }
            let description = object
                .get("description")
                .and_then(Value::as_str)
                .unwrap_or("")
                .trim()
                .to_string();
            results.push(SearchResult {
                title: title.to_string(),
                url,
                description,
            });
            if results.len() >= limit {
                return results;
            }
        }
    }

    if results.len() < limit {
        let link_selector = match Selector::parse("a[href]") {
            Ok(selector) => selector,
            Err(_) => return results,
        };
        for link in document.select(&link_selector) {
            let raw_url = link.value().attr("href").unwrap_or("");
            let Some(url) = normalize_result_url(raw_url) else {
                continue;
            };
            if !seen_urls.insert(url.clone()) {
                continue;
            }
            let title = normalize_text(&link.text().collect::<String>());
            let generic = title.to_ascii_lowercase();
            if title.is_empty()
                || matches!(
                    generic.as_str(),
                    "leggi" | "scopri" | "vai" | "approfondisci"
                )
            {
                continue;
            }
            results.push(SearchResult {
                title,
                url,
                description: String::new(),
            });
            if results.len() >= limit {
                break;
            }
        }
    }

    results
}

fn normalize_text(text: &str) -> String {
    text.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn element_visible_text(element: &ElementRef<'_>) -> String {
    let mut output = String::new();
    for descendant in element.descendants() {
        let Some(text) = descendant.value().as_text() else {
            continue;
        };
        let is_hidden = descendant
            .ancestors()
            .filter_map(ElementRef::wrap)
            .any(|ancestor| matches!(ancestor.value().name(), "style" | "script"));
        if !is_hidden {
            output.push_str(text);
        }
    }
    normalize_text(&output)
}

fn normalize_article_text(text: &str) -> String {
    let mut output = String::new();
    let mut previous_blank = false;
    for line in text.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            if !previous_blank && !output.is_empty() {
                output.push('\n');
            }
            previous_blank = true;
            continue;
        }
        if !output.is_empty() && !output.ends_with('\n') {
            output.push('\n');
        }
        output.push_str(trimmed);
        previous_blank = false;
    }
    output.trim().to_string()
}

fn heading_marker(level: usize, text: &str) -> String {
    let level = level.clamp(2, 6);
    let marker = "=".repeat(level);
    format!("{marker} {text} {marker}")
}

fn element_has_skipped_ancestor(element: &ElementRef<'_>) -> bool {
    for ancestor in element.ancestors().filter_map(ElementRef::wrap) {
        let name = ancestor.value().name();
        if matches!(name, "header" | "nav" | "aside" | "footer" | "form") {
            return true;
        }
        let href = ancestor
            .value()
            .attr("href")
            .unwrap_or("")
            .to_ascii_lowercase();
        if href.starts_with("/enciclopedia/tag/") {
            return true;
        }
        let id = ancestor.value().id().unwrap_or("").to_ascii_lowercase();
        let classes = ancestor
            .value()
            .classes()
            .collect::<Vec<_>>()
            .join(" ")
            .to_ascii_lowercase();
        let marker = format!("{id} {classes}");
        if [
            "breadcrumb",
            "related",
            "correlat",
            "social",
            "share",
            "subscription",
            "advert",
            "banner",
            "modal",
            "cookie",
            "newsletter",
        ]
        .iter()
        .any(|needle| marker.contains(needle))
        {
            return true;
        }
    }
    false
}

fn is_boilerplate(text: &str) -> bool {
    let normalized = text.trim().to_ascii_lowercase();
    normalized.is_empty()
        || normalized == "indice"
        || normalized == "categorie"
        || normalized == "tag"
        || normalized == "dal vocabolario"
        || normalized == "lemmi correlati"
        || normalized == "scarica"
        || normalized == "salta al contenuto"
        || normalized.starts_with("vuoi navigare il sito treccani")
        || normalized.starts_with("scarica l'app")
}

fn is_copyright(text: &str) -> bool {
    let normalized = text.trim().to_ascii_lowercase();
    normalized.starts_with('©')
        || normalized.contains("istituto della enciclopedia italiana fondata da giovanni treccani")
        || normalized.contains("riproduzione riservata")
}

fn heading_title(line: &str) -> Option<(usize, String)> {
    let trimmed = line.trim();
    for level in 2..=6 {
        let marker = "=".repeat(level);
        let prefix = format!("{marker} ");
        let suffix = format!(" {marker}");
        let Some(body) = trimmed
            .strip_prefix(&prefix)
            .and_then(|value| value.strip_suffix(&suffix))
        else {
            continue;
        };
        if body.contains("==") {
            return None;
        }
        let title = body.trim();
        if !title.is_empty() {
            return Some((level, title.to_string()));
        }
    }
    None
}

fn legacy_smallcaps_heading(element: &ElementRef<'_>, text: &str) -> Option<String> {
    if element.value().name() != "p" {
        return None;
    }
    let selector = Selector::parse("span.tc-smallcaps").ok()?;
    let labels = element
        .select(&selector)
        .map(|span| element_visible_text(&span))
        .filter(|label| !label.is_empty())
        .collect::<Vec<_>>();
    if labels.len() != 1 {
        return None;
    }
    let label = &labels[0];
    if text.eq_ignore_ascii_case("bibliografia") && label.eq_ignore_ascii_case("bibliografia") {
        return Some("Bibliografia".to_string());
    }
    let (number, title) = text.split_once(". ")?;
    if number.is_empty()
        || number.len() > 3
        || !number.chars().all(|character| character.is_ascii_digit())
        || !title.eq_ignore_ascii_case(label)
    {
        return None;
    }
    Some(format!("{number}. {title}"))
}

fn extract_article_sections(text: &str) -> Vec<ArticleSection> {
    let lines = text.lines().collect::<Vec<_>>();
    let headings = lines
        .iter()
        .enumerate()
        .filter_map(|(index, line)| heading_title(line).map(|(level, title)| (index, level, title)))
        .collect::<Vec<_>>();
    let mut sections = Vec::new();

    for (position, (start, level, title)) in headings.iter().enumerate() {
        let end = headings
            .iter()
            .skip(position + 1)
            .find(|(_, next_level, _)| next_level <= level)
            .map(|(index, _, _)| *index)
            .unwrap_or(lines.len());
        let section_text = lines[*start..end].join("\n").trim().to_string();
        if !section_text.is_empty() {
            sections.push(ArticleSection {
                title: title.clone(),
                level: *level,
                text: section_text,
            });
        }
    }

    sections
}

fn extract_blocks_from_elements<'a, I>(elements: I, title: &str) -> Vec<String>
where
    I: IntoIterator<Item = ElementRef<'a>>,
{
    let elements = elements.into_iter().collect::<Vec<_>>();
    let contains_title = elements.iter().any(|element| {
        element.value().name() == "h1" && element_visible_text(element).eq_ignore_ascii_case(title)
    });

    let mut blocks = Vec::new();
    let mut found_title = !contains_title;
    let mut skip_index_entries = false;
    let mut skip_auxiliary_entries = false;
    let mut seen_blocks = HashSet::new();

    blocks.push(title.to_string());
    seen_blocks.insert(title.to_ascii_lowercase());

    for element in elements {
        if element_has_skipped_ancestor(&element) {
            continue;
        }
        let tag = element.value().name();
        let text = element_visible_text(&element);
        if text.is_empty() {
            continue;
        }
        if tag == "h1" {
            if text.eq_ignore_ascii_case(title) {
                found_title = true;
            }
            continue;
        }
        if !found_title {
            continue;
        }
        if is_copyright(&text) {
            break;
        }
        if text.eq_ignore_ascii_case("Indice") {
            skip_index_entries = true;
            continue;
        }
        if skip_index_entries && tag == "li" {
            continue;
        }
        if skip_index_entries && tag != "li" {
            skip_index_entries = false;
        }

        let normalized = text.to_ascii_lowercase();
        if matches!(
            normalized.as_str(),
            "categorie" | "tag" | "dal vocabolario" | "lemmi correlati"
        ) {
            skip_auxiliary_entries = true;
            continue;
        }
        if skip_auxiliary_entries {
            let starts_article_body = matches!(tag, "h2" | "h3" | "h4" | "h5" | "h6")
                || (tag == "p"
                    && (text.len() >= 80
                        || text.ends_with('.')
                        || text.ends_with('!')
                        || text.ends_with('?')));
            if starts_article_body {
                skip_auxiliary_entries = false;
            } else {
                continue;
            }
        }
        if is_boilerplate(&text) {
            continue;
        }

        let legacy_heading = legacy_smallcaps_heading(&element, &text);
        let is_article_metadata_heading = matches!(tag, "h5" | "h6")
            && (normalized.starts_with("enciclopedia ") || normalized.starts_with("di "));
        let block = if let Some(title) = legacy_heading {
            heading_marker(2, &title)
        } else if normalized == "enciclopedia on line" || is_article_metadata_heading {
            text
        } else {
            match tag {
                "h2" => heading_marker(2, &text),
                "h3" => heading_marker(3, &text),
                "h4" => heading_marker(4, &text),
                "h5" => heading_marker(5, &text),
                "h6" => heading_marker(6, &text),
                "li" => format!("- {text}"),
                _ => text,
            }
        };
        let dedupe_key = block.to_ascii_lowercase();
        if seen_blocks.insert(dedupe_key) {
            blocks.push(block);
        }
    }

    blocks
}

fn article_blocks_score(blocks: &[String], title: &str) -> usize {
    blocks
        .iter()
        .filter(|block| !block.eq_ignore_ascii_case(title))
        .map(|block| {
            if block.starts_with("- ") || heading_title(block).is_some() {
                0
            } else {
                block.len()
            }
        })
        .sum()
}

fn parse_article_html(html: &str) -> Result<(String, String, Vec<ArticleSection>), TreccaniError> {
    let document = Html::parse_document(html);
    let root_selector = Selector::parse("article, main, [role=\"main\"]")
        .map_err(|err| TreccaniError::Other(err.to_string()))?;
    let content_selector = Selector::parse("h1, h2, h3, h4, h5, h6, p, li")
        .map_err(|err| TreccaniError::Other(err.to_string()))?;
    let h1_selector = Selector::parse("h1").map_err(|err| TreccaniError::Other(err.to_string()))?;

    let title = document
        .select(&h1_selector)
        .map(|element| element_visible_text(&element))
        .find(|value| !value.is_empty())
        .ok_or(TreccaniError::NotFound)?;

    // Treccani can render the title/index and the actual article body in different
    // <main>/<article> containers. Select the container with the largest amount of
    // substantial prose instead of assuming that the first <main> is the article.
    let mut best_blocks = Vec::new();
    let mut best_score = 0usize;
    for root in document.select(&root_selector) {
        let blocks = extract_blocks_from_elements(root.select(&content_selector), &title);
        let score = article_blocks_score(&blocks, &title);
        if score > best_score {
            best_score = score;
            best_blocks = blocks;
        }
    }

    // Also inspect the complete document. Some Treccani layouts keep the title/index
    // in one semantic container and place the article body in a sibling <div>. Prefer
    // the complete document only when it contains substantially more prose, otherwise
    // keep the cleaner semantic-container extraction.
    let document_blocks = extract_blocks_from_elements(document.select(&content_selector), &title);
    let document_score = article_blocks_score(&document_blocks, &title);
    if best_score == 0 || document_score > best_score.saturating_mul(3) / 2 {
        best_blocks = document_blocks;
    }

    let extract = normalize_article_text(&best_blocks.join("\n\n"));
    if article_blocks_score(&best_blocks, &title) == 0 {
        return Err(TreccaniError::NotFound);
    }
    let sections = extract_article_sections(&extract);
    Ok((title, extract, sections))
}

pub fn search_articles(query: &str, limit: usize) -> Result<Vec<SearchResult>, TreccaniError> {
    let trimmed = query.trim();
    if trimmed.is_empty() {
        return Ok(Vec::new());
    }
    let url = build_search_url(trimmed)?;
    let response = http_client()?
        .get(url)
        .send()
        .map_err(|err| TreccaniError::Other(err.to_string()))?
        .error_for_status()
        .map_err(|err| TreccaniError::Other(err.to_string()))?;
    let html = response
        .text()
        .map_err(|err| TreccaniError::Other(err.to_string()))?;
    Ok(parse_search_results(&html, limit))
}

pub fn fetch_extract(article_url: &str) -> Result<ExtractResult, TreccaniError> {
    let url = normalize_result_url(article_url).ok_or(TreccaniError::InvalidUrl)?;
    let response = http_client()?
        .get(&url)
        .send()
        .map_err(|err| TreccaniError::Other(err.to_string()))?
        .error_for_status()
        .map_err(|err| TreccaniError::Other(err.to_string()))?;
    let html = response
        .text()
        .map_err(|err| TreccaniError::Other(err.to_string()))?;
    let (_, extract, sections) = parse_article_html(&html)?;
    Ok(ExtractResult {
        extract,
        url,
        sections,
    })
}

#[cfg(test)]
mod tests {
    use super::{
        fetch_extract, normalize_result_url, parse_article_html, parse_search_results,
        search_articles,
    };

    #[test]
    fn treccani_search_reads_only_encyclopedia_results() {
        let html = r#"
        <html><body><script type="application/json">
        {"props":{"pageProps":{"data":{"matches":[
          {"title":"Enciclopedia","url":"/enciclopedia/enciclopedia/","section":"enciclopedia","description":"Enciclopedia on line"},
          {"title":"Enciclopedia","url":"/vocabolario/enciclopedia/","section":"vocabolario","description":"Vocabolario"}
        ]}}}}
        </script></body></html>
        "#;
        let results = parse_search_results(html, 20);
        assert_eq!(results.len(), 1);
        assert_eq!(results[0].title, "Enciclopedia");
        assert!(results[0].url.contains("/enciclopedia/enciclopedia/"));
    }

    #[test]
    fn treccani_article_parser_keeps_sections_and_skips_index() {
        let html = r#"
        <html><body><main>
          <h1>Enciclopedia</h1>
          <p>Enciclopedia on line</p>
          <div><p>Indice</p><ul><li>Storia</li><li>Opere</li></ul></div>
          <p>Introduzione dell'articolo.</p>
          <h2>Storia</h2><p>Testo della storia.</p>
          <h3>Origini</h3><p>Testo delle origini.</p>
          <h2>Opere</h2><p>Testo delle opere.</p>
          <footer><p>© Istituto della Enciclopedia Italiana - Riproduzione riservata</p></footer>
        </main></body></html>
        "#;
        let (title, extract, sections) = parse_article_html(html).expect("valid article");
        assert_eq!(title, "Enciclopedia");
        assert!(extract.contains("== Storia =="));
        assert!(extract.contains("=== Origini ==="));
        assert!(!extract.contains("- Storia"));
        assert_eq!(sections.len(), 3);
        assert!(sections[0].text.contains("=== Origini ==="));
    }

    #[test]
    fn treccani_article_parser_supports_short_entries_without_sections() {
        let html = r#"
        <html><body><main><h1>Query</h1><p>Enciclopedia della Matematica</p>
        <p>Query in informatica, istruzione che permette l'accesso ai dati.</p>
        <p>Un secondo paragrafo.</p></main></body></html>
        "#;
        let (_, extract, sections) = parse_article_html(html).expect("valid article");
        assert!(extract.contains("Query in informatica"));
        assert!(sections.is_empty());
    }

    #[test]
    fn treccani_article_parser_selects_body_from_a_later_semantic_container() {
        let html = r#"
        <html><body>
          <main>
            <h1>Chimica</h1>
            <h5>Enciclopedia on line</h5>
            <p>Indice</p>
            <ul><li>1 Storia</li><li>2 Le origini</li></ul>
            <p>TAG</p>
            <p>Risonanza magnetica nucleare</p>
          </main>
          <article>
            <div class="paywall term-paragraph0">
            <p>Scienza che studia le proprietà, la composizione e il modo di reagire delle sostanze naturali e artificiali.</p>
            </div>
            <div class="paywall term-paragraph1">
            <h2>Storia</h2>
            <p>La nascita della chimica si fa risalire alla seconda metà del diciottesimo secolo, quando si svilupparono le sue basi sperimentali.</p>
            </div>
            <div class="paywall term-paragraph2">
            <h2>Le origini</h2>
            <p>Le prime conoscenze chimiche furono legate allo sviluppo delle capacità tecniche dei popoli antichi.</p>
            </div>
          </article>
        </body></html>
        "#;
        let (_, extract, sections) = parse_article_html(html).expect("valid split article");
        assert!(extract.contains("Scienza che studia le proprietà"));
        assert!(extract.contains("La nascita della chimica"));
        assert!(!extract.contains("Risonanza magnetica nucleare"));
        assert_eq!(sections.len(), 2);
    }

    #[test]
    fn treccani_article_parser_falls_back_to_document_for_body_in_plain_div() {
        let html = r#"
        <html><body>
          <main>
            <h1>Chimica</h1>
            <p>Enciclopedia on line</p>
            <p>Indice</p>
            <ul><li>1 Storia</li><li>2 Le origini</li></ul>
            <p>TAG</p>
            <ul><li>Risonanza magnetica nucleare</li><li>Rivoluzione industriale</li></ul>
            <p>Lemmi correlati</p>
          </main>
          <div class="paywall term-paragraph0">
            <p>Scienza che studia le proprietà, la composizione, l’identificazione, la preparazione e il modo di reagire delle sostanze naturali e artificiali.</p>
          </div>
          <div class="paywall term-paragraph1">
            <h2>Storia</h2>
            <p>La nascita della chimica si fa in genere risalire alla seconda metà del diciottesimo secolo, quando si svilupparono le basi sperimentali e teoriche della scienza moderna.</p>
          </div>
          <div class="paywall term-paragraph2">
            <h2>Le origini</h2>
            <p>Le prime conoscenze chimiche nacquero insieme alle capacità tecniche dei popoli antichi e alla lavorazione dei metalli, del vetro e dei coloranti.</p>
          </div>
        </body></html>
        "#;
        let (_, extract, sections) = parse_article_html(html).expect("valid split article");
        assert!(extract.contains("Scienza che studia le proprietà"));
        assert!(extract.contains("La nascita della chimica"));
        assert!(!extract.contains("Risonanza magnetica nucleare"));
        assert!(!extract.contains("Rivoluzione industriale"));
        assert_eq!(sections.len(), 2);
    }

    #[test]
    fn treccani_article_parser_recognizes_legacy_smallcaps_index() {
        let html = r#"
        <html><body>
          <h1>Autorita</h1>
          <h6>di Augusto Del Noce</h6>
          <h5>Enciclopedia del Novecento (1975)</h5>
          <div class="paywall term-paragraph2"><p>
            <style>.css-author{font-weight:400}</style>
            di <em>Augusto Del Noce</em>
          </p></div>
          <div class="paywall term-paragraph3"><p><span class="tc-smallcaps">sommario</span>:
            1. Eclissi dell'idea di autorità. 2. Autorità e potere. □ Bibliografia.
          </p></div>
          <div class="paywall term-paragraph4"><p>1. <span class="tc-smallcaps">Eclissi dell'idea di autorità</span></p></div>
          <div class="paywall term-paragraph5"><p>Primo paragrafo sostanziale dedicato all'eclissi dell'autorità nel mondo contemporaneo e alle sue conseguenze.</p></div>
          <div class="paywall term-paragraph6"><p>2. <span class="tc-smallcaps">Autorità e potere</span></p></div>
          <div class="paywall term-paragraph7"><p>Secondo paragrafo sostanziale che distingue attentamente il concetto di autorità da quello differente di potere.</p></div>
          <div class="paywall term-paragraph8"><p><span class="tc-smallcaps">bibliografia</span></p></div>
          <div class="paywall term-paragraph9"><p>AA. VV., Coscienza, legge, autorità, Brescia 1970.</p></div>
        </body></html>
        "#;
        let (_, extract, sections) = parse_article_html(html).expect("valid legacy article");
        assert!(!extract.contains(".css-author"));
        assert!(extract.contains("== 1. Eclissi dell'idea di autorità =="));
        assert!(extract.contains("== 2. Autorità e potere =="));
        assert!(extract.contains("== Bibliografia =="));
        assert_eq!(
            sections
                .iter()
                .map(|section| section.title.as_str())
                .collect::<Vec<_>>(),
            [
                "1. Eclissi dell'idea di autorità",
                "2. Autorità e potere",
                "Bibliografia"
            ]
        );
    }

    #[test]
    fn treccani_search_falls_back_to_article_links() {
        let html = r#"
        <html><body><main>
          <a href="/enciclopedia/filosofia/">Filosofia</a>
          <a href="/vocabolario/filosofia/">Filosofia nel vocabolario</a>
        </main></body></html>
        "#;
        let results = parse_search_results(html, 20);
        assert_eq!(results.len(), 1);
        assert_eq!(results[0].title, "Filosofia");
    }

    #[test]
    fn treccani_url_validation_rejects_external_hosts_and_search_pages() {
        assert!(normalize_result_url("https://example.com/enciclopedia/test/").is_none());
        assert!(normalize_result_url("/enciclopedia/ricerca/test/").is_none());
        assert!(normalize_result_url("/enciclopedia/test/").is_some());
    }

    #[test]
    #[ignore = "requires the live Treccani website"]
    fn treccani_live_del_noce_reads_the_complete_entry() {
        let article = search_articles("del noce", 20)
            .expect("live Del Noce search")
            .into_iter()
            .find(|item| item.url.contains("/enciclopedia/augusto-del-noce"))
            .expect("Del Noce in live search results");
        let result = fetch_extract(&article.url).expect("live Del Noce entry");
        assert!(result.extract.contains("Filosofo italiano"));
        assert!(result.extract.contains("L'epoca della secolarizzazione"));
        assert!(
            !result
                .extract
                .lines()
                .any(|line| line == "Riforma cattolica")
        );
        assert!(result.sections.is_empty());
        eprintln!(
            "Del Noce: {} characters, {} sections",
            result.extract.chars().count(),
            result.sections.len()
        );
    }

    #[test]
    #[ignore = "requires the live Treccani website"]
    fn treccani_live_chimica_reads_body_and_sections() {
        let article = search_articles("chimica", 20)
            .expect("live Chimica search")
            .into_iter()
            .find(|item| item.url.contains("/enciclopedia/chimica"))
            .expect("Chimica in live search results");
        let result = fetch_extract(&article.url).expect("live Chimica entry");
        assert!(result.extract.contains("Scienza che studia le proprietà"));
        assert!(result.extract.contains("== Storia =="));
        assert!(result.extract.contains("== C. organica =="));
        assert!(
            !result
                .extract
                .lines()
                .any(|line| line == "Risonanza magnetica nucleare")
        );
        assert_eq!(result.sections.len(), 11);
        eprintln!(
            "Chimica: {} characters, {} sections",
            result.extract.chars().count(),
            result.sections.len()
        );
    }

    #[test]
    #[ignore = "requires the live Treccani website"]
    fn treccani_live_autorita_del_noce_reads_legacy_index() {
        let results = search_articles("del noce", 20).expect("live Del Noce search");
        let (position, article) = results
            .into_iter()
            .enumerate()
            .find(|(_, item)| {
                let url = item.url.to_ascii_lowercase();
                url.contains("autorita_") && url.contains("enciclopedia-del-novecento")
            })
            .expect("Del Noce Autorita entry in live search results");
        let result = fetch_extract(&article.url).expect("live Autorita entry");
        let expected_titles = [
            "1. Eclissi dell'idea di autorità e crisi del mondo contemporaneo",
            "2. Autorità e potere",
            "3. Autorità e rivoluzione",
            "4. L'Occidente e il tramonto dell'autorità",
            "5. I totalitarismi e la negazione dell'autorità",
            "6. Lo spirito borghese e l'autorità",
            "7. L'idea di autorità nell'età liberale e nel primo dopoguerra",
            "8. Conclusioni",
            "Bibliografia",
        ];
        assert_eq!(
            result
                .sections
                .iter()
                .map(|section| section.title.as_str())
                .collect::<Vec<_>>(),
            expected_titles
        );
        assert!(!result.extract.contains(".css-"));
        assert!(
            result
                .extract
                .contains("== 1. Eclissi dell'idea di autorità")
        );
        assert!(result.extract.contains("== Bibliografia =="));
        eprintln!(
            "Autorita by Del Noce: search position {}, {} characters, {} sections",
            position + 1,
            result.extract.chars().count(),
            result.sections.len()
        );
    }

    #[test]
    #[ignore = "requires 40 live searches and article requests to Treccani"]
    fn treccani_live_bulk_40_entries_have_real_article_bodies() {
        let queries = [
            "chimica",
            "fisica",
            "matematica",
            "biologia",
            "astronomia",
            "geologia",
            "medicina",
            "informatica",
            "filosofia",
            "logica",
            "etica",
            "diritto",
            "economia",
            "sociologia",
            "psicologia",
            "linguistica",
            "letteratura",
            "poesia",
            "musica",
            "pittura",
            "scultura",
            "architettura",
            "fotografia",
            "cinema",
            "Roma",
            "Milano",
            "Napoli",
            "Sicilia",
            "Europa",
            "Rinascimento",
            "Illuminismo",
            "Risorgimento",
            "Dante Alighieri",
            "Leonardo da Vinci",
            "Galileo Galilei",
            "Alessandro Manzoni",
            "Maria Montessori",
            "Augusto Del Noce",
            "Enrico Fermi",
            "Rita Levi Montalcini",
        ];
        let mut seen_urls = std::collections::HashSet::new();
        let mut failures = Vec::new();

        for query in queries {
            let search_results = match search_articles(query, 20) {
                Ok(results) => results,
                Err(error) => {
                    failures.push(format!("{query}: search failed: {error}"));
                    continue;
                }
            };
            let Some(article) = search_results
                .into_iter()
                .find(|item| !seen_urls.contains(&item.url))
            else {
                failures.push(format!("{query}: no distinct encyclopedia result"));
                continue;
            };
            seen_urls.insert(article.url.clone());

            let result = match fetch_extract(&article.url) {
                Ok(result) => result,
                Err(error) => {
                    failures.push(format!("{query}: article fetch failed: {error}"));
                    continue;
                }
            };
            let character_count = result.extract.chars().count();
            let has_substantial_prose = result.extract.lines().any(|line| {
                let line = line.trim();
                line.chars().count() >= 80 && !line.starts_with('=') && !line.starts_with("- ")
            });
            let has_metadata_noise = result.extract.lines().any(|line| {
                matches!(
                    line.trim().to_ascii_lowercase().as_str(),
                    "tag" | "categorie" | "indice" | "dal vocabolario" | "lemmi correlati"
                )
            });
            let sections_are_consistent = result.sections.iter().all(|section| {
                let marker = "=".repeat(section.level.clamp(2, 6));
                let heading = format!("{marker} {} {marker}", section.title);
                result.extract.contains(&heading) && section.text.starts_with(&heading)
            });

            if character_count < 100 {
                failures.push(format!("{query}: only {character_count} characters"));
            }
            if !has_substantial_prose {
                failures.push(format!("{query}: no substantial prose paragraph"));
            }
            if has_metadata_noise {
                failures.push(format!("{query}: metadata/index noise in article body"));
            }
            if !sections_are_consistent {
                failures.push(format!("{query}: inconsistent section extraction"));
            }

            eprintln!(
                "{query} | {} | {character_count} characters | {} sections | {}",
                article.title,
                result.sections.len(),
                article.url
            );
        }

        assert_eq!(
            seen_urls.len(),
            40,
            "not all searches produced distinct entries"
        );
        assert!(
            failures.is_empty(),
            "bulk Treccani failures:\n{}",
            failures.join("\n")
        );
    }
}
