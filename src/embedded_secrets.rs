//! Chiavi API embedded a compile time.
//! I valori vengono impostati durante il CI build tramite GitHub Secrets.
//! Se non sono impostati, restano vuoti e l'utente dovrà inserirli manualmente.

/// API Key di default per Podcast Index (impostata durante CI build)
pub fn default_podcast_index_api_key() -> &'static str {
    option_env!("PODCAST_INDEX_API_KEY").unwrap_or("")
}

/// API Secret di default per Podcast Index (impostata durante CI build)
pub fn default_podcast_index_api_secret() -> &'static str {
    option_env!("PODCAST_INDEX_API_SECRET").unwrap_or("")
}

/// Controlla se ci sono chiavi di default embedded
pub fn has_default_podcast_index_keys() -> bool {
    !default_podcast_index_api_key().is_empty() && !default_podcast_index_api_secret().is_empty()
}
