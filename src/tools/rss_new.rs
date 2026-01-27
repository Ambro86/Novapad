pub async fn fetch_article_text(
    url: &str,
    fallback_title: &str,
    fallback_description: &str,
) -> Result<String, String> {
    let start_total = Instant::now();
    let url_str = normalize_url(url);
    if url_str.is_empty() {
        return Err("Empty URL".to_string());
    }

    log_debug(&format!("rss_article_fetch starting via curl-impersonate url=\"{}\"", url_str));
    
    let url_for_curl = url_str.clone();
    let bytes_res = tokio::task::spawn_blocking(move || {
        fetch_article_curl(&url_for_curl)
    }).await.map_err(|e| e.to_string())?;

    let html = match bytes_res {
        Ok(bytes) => String::from_utf8_lossy(&bytes).to_string(),
        Err(err) => {
            log_debug(&format!("rss_article_fetch curl_failed url=\"{}\" error=\"{}\"", url_str, err));
            return Err(err);
        }
    };

    let article = reader::reader_mode_extract(&html).unwrap_or(reader::ArticleContent {
        title: fallback_title.to_string(),
        content: fallback_description.to_string(),
    });
    log_debug(&format!(
        "rss_article_fetch_done ms={} url=\"{}\"",
        start_total.elapsed().as_millis(),
        url_str
    ));
    Ok(format!("{}\n\n{}", article.title, article.content))
}
