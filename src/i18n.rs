use crate::settings::Language;
use std::collections::HashMap;
use std::sync::OnceLock;

const EN_JSON: &str = include_str!("../i18n/en.json");
const IT_JSON: &str = include_str!("../i18n/it.json");
const ES_JSON: &str = include_str!("../i18n/es.json");
const PT_JSON: &str = include_str!("../i18n/pt.json");
const SV_JSON: &str = include_str!("../i18n/sv.json");
const VI_JSON: &str = include_str!("../i18n/vi.json");
const CS_JSON: &str = include_str!("../i18n/cs.json");
const PL_JSON: &str = include_str!("../i18n/pl.json");

fn load_map(raw: &str) -> HashMap<String, String> {
    let mut map: HashMap<String, String> = serde_json::from_str(raw).unwrap_or_default();
    for value in map.values_mut() {
        if value.contains("\\n") {
            *value = value.replace("\\n", "\n");
        }
    }
    map
}

fn load_sv_map() -> HashMap<String, String> {
    if let Ok(exe_path) = std::env::current_exe() {
        if let Some(dir) = exe_path.parent() {
            let override_path = dir.join("sv.json");
            if override_path.exists() {
                match std::fs::read_to_string(&override_path) {
                    Ok(content) => {
                        match serde_json::from_str::<HashMap<String, String>>(&content) {
                            Ok(mut map) => {
                                for value in map.values_mut() {
                                    if value.contains("\\n") {
                                        *value = value.replace("\\n", "\n");
                                    }
                                }
                                return map;
                            }
                            Err(err) => {
                                crate::log_debug(&format!(
                                    "Failed to parse Swedish translation override: {err}"
                                ));
                            }
                        }
                    }
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Failed to read Swedish translation override: {err}"
                        ));
                    }
                }
            }
        }
    } else {
        crate::log_debug("Failed to resolve executable path for Swedish translation override.");
    }
    load_map(SV_JSON)
}

fn load_cs_map() -> HashMap<String, String> {
    if let Ok(exe_path) = std::env::current_exe() {
        if let Some(dir) = exe_path.parent() {
            let override_path = dir.join("cs.json");
            if override_path.exists() {
                match std::fs::read_to_string(&override_path) {
                    Ok(content) => {
                        match serde_json::from_str::<HashMap<String, String>>(&content) {
                            Ok(mut map) => {
                                for value in map.values_mut() {
                                    if value.contains("\\n") {
                                        *value = value.replace("\\n", "\n");
                                    }
                                }
                                return map;
                            }
                            Err(err) => {
                                crate::log_debug(&format!(
                                    "Failed to parse Czech translation override: {err}"
                                ));
                            }
                        }
                    }
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Failed to read Czech translation override: {err}"
                        ));
                    }
                }
            }
        }
    } else {
        crate::log_debug("Failed to resolve executable path for Czech translation override.");
    }
    load_map(CS_JSON)
}

fn load_pl_map() -> HashMap<String, String> {
    if let Ok(exe_path) = std::env::current_exe() {
        if let Some(dir) = exe_path.parent() {
            let override_path = dir.join("pl.json");
            if override_path.exists() {
                match std::fs::read_to_string(&override_path) {
                    Ok(content) => {
                        match serde_json::from_str::<HashMap<String, String>>(&content) {
                            Ok(mut map) => {
                                for value in map.values_mut() {
                                    if value.contains("\\n") {
                                        *value = value.replace("\\n", "\n");
                                    }
                                }
                                return map;
                            }
                            Err(err) => {
                                crate::log_debug(&format!(
                                    "Failed to parse Polish translation override: {err}"
                                ));
                            }
                        }
                    }
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Failed to read Polish translation override: {err}"
                        ));
                    }
                }
            }
        }
    } else {
        crate::log_debug("Failed to resolve executable path for Polish translation override.");
    }
    load_map(PL_JSON)
}

fn map_for_language(language: Language) -> &'static HashMap<String, String> {
    static EN: OnceLock<HashMap<String, String>> = OnceLock::new();
    static IT: OnceLock<HashMap<String, String>> = OnceLock::new();
    static ES: OnceLock<HashMap<String, String>> = OnceLock::new();
    static PT: OnceLock<HashMap<String, String>> = OnceLock::new();
    static SV: OnceLock<HashMap<String, String>> = OnceLock::new();
    static VI: OnceLock<HashMap<String, String>> = OnceLock::new();
    static CS: OnceLock<HashMap<String, String>> = OnceLock::new();
    static PL: OnceLock<HashMap<String, String>> = OnceLock::new();
    match language {
        Language::Italian => IT.get_or_init(|| load_map(IT_JSON)),
        Language::Spanish => ES.get_or_init(|| load_map(ES_JSON)),
        Language::Portuguese => PT.get_or_init(|| load_map(PT_JSON)),
        Language::Swedish => SV.get_or_init(load_sv_map),
        Language::Vietnamese => VI.get_or_init(|| load_map(VI_JSON)),
        Language::Czech => CS.get_or_init(load_cs_map),
        Language::Polish => PL.get_or_init(load_pl_map),
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
