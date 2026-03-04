use std::collections::BTreeSet;
use std::fs;
use std::path::Path;

use serde_json::Value;

fn load_keys(path: &Path) -> BTreeSet<String> {
    let raw = fs::read_to_string(path)
        .unwrap_or_else(|err| panic!("Failed to read {}: {err}", path.display()));
    let json: Value = serde_json::from_str(&raw)
        .unwrap_or_else(|err| panic!("Invalid JSON in {}: {err}", path.display()));
    let obj = json
        .as_object()
        .unwrap_or_else(|| panic!("Expected a JSON object at top-level in {}", path.display()));
    obj.keys()
        .filter(|key| !key.starts_with("excluded_from_testing."))
        .cloned()
        .collect()
}

#[test]
fn i18n_files_have_matching_keys() {
    let i18n_dir = Path::new("i18n");
    let en_path = i18n_dir.join("en.json");
    assert!(en_path.exists(), "Missing {}", en_path.display());

    let en_keys = load_keys(&en_path);
    assert!(!en_keys.is_empty(), "en.json has no keys");

    let entries = fs::read_dir(i18n_dir)
        .unwrap_or_else(|err| panic!("Failed to read {}: {err}", i18n_dir.display()));

    let mut checked = 0usize;
    for entry in entries {
        let entry = entry.unwrap_or_else(|err| panic!("Failed dir entry: {err}"));
        let path = entry.path();
        if path.extension().and_then(|ext| ext.to_str()) != Some("json") {
            continue;
        }
        if path.file_name().and_then(|name| name.to_str()) == Some("en.json") {
            continue;
        }

        let keys = load_keys(&path);
        let missing: BTreeSet<_> = en_keys.difference(&keys).cloned().collect();
        let extra: BTreeSet<_> = keys.difference(&en_keys).cloned().collect();
        assert!(
            missing.is_empty() && extra.is_empty(),
            "{} key mismatch.\nMissing: {:?}\nExtra: {:?}",
            path.display(),
            missing,
            extra
        );
        checked += 1;
    }

    assert!(
        checked > 0,
        "No translation files found in {}",
        i18n_dir.display()
    );
}
