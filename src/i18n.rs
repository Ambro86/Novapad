use crate::settings::Language;
use std::collections::HashMap;
use std::sync::OnceLock;

const EN_JSON: &str = include_str!("../i18n/en.json");
const IT_JSON: &str = include_str!("../i18n/it.json");
const ES_JSON: &str = include_str!("../i18n/es.json");
const PT_JSON: &str = include_str!("../i18n/pt.json");
const VI_JSON: &str = include_str!("../i18n/vi.json");

fn load_map(raw: &str) -> HashMap<String, String> {
    let mut map: HashMap<String, String> = serde_json::from_str(raw).unwrap_or_default();
    for value in map.values_mut() {
        if value.contains("\\n") {
            *value = value.replace("\\n", "\n");
        }
    }
    map
}

fn map_for_language(language: Language) -> &'static HashMap<String, String> {
    static EN: OnceLock<HashMap<String, String>> = OnceLock::new();
    static IT: OnceLock<HashMap<String, String>> = OnceLock::new();
    static ES: OnceLock<HashMap<String, String>> = OnceLock::new();
    static PT: OnceLock<HashMap<String, String>> = OnceLock::new();
    static VI: OnceLock<HashMap<String, String>> = OnceLock::new();
    match language {
        Language::Italian => IT.get_or_init(|| load_map(IT_JSON)),
        Language::Spanish => ES.get_or_init(|| load_map(ES_JSON)),
        Language::Portuguese => PT.get_or_init(|| load_map(PT_JSON)),
        Language::Vietnamese => VI.get_or_init(|| load_map(VI_JSON)),
        Language::English => EN.get_or_init(|| load_map(EN_JSON)),
    }
}

pub fn tr(language: Language, key: &str) -> String {
    map_for_language(language)
        .get(key)
        .cloned()
        .unwrap_or_else(|| key.to_string())
}

pub fn tr_f(language: Language, key: &str, args: &[(&str, &str)]) -> String {
    let mut out = tr(language, key);
    for (name, value) in args {
        let token = format!("{{{name}}}");
        out = out.replace(&token, value);
    }
    out
}
