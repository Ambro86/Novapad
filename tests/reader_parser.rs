use std::fs;
use std::path::Path;

mod settings {
    #[derive(Clone, Copy)]
    #[allow(dead_code)] // reader.rs matches on all variants; test only needs one.
    pub enum Language {
        English,
        Italian,
        French,
        Spanish,
        Portuguese,
        Swedish,
        Czech,
        Polish,
        Vietnamese,
    }
}

#[path = "../src/tools/reader.rs"]
mod reader;

struct Expectation {
    file: String,
    min_chars: usize,
    min_words: usize,
    require_title: bool,
}

fn load_expectations() -> Vec<Expectation> {
    let data =
        fs::read_to_string("tests/fixtures/article_expectations.csv").expect("read expectations");
    let mut out = Vec::new();
    for (i, line) in data.lines().enumerate() {
        if i == 0 {
            continue;
        }
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        let parts: Vec<_> = line.split(',').map(|s| s.trim()).collect();
        if parts.len() < 4 {
            continue;
        }
        let file = parts[0].to_string();
        let min_chars = parts[1].parse::<usize>().expect("min_chars");
        let min_words = parts[2].parse::<usize>().expect("min_words");
        let require_title = !parts[3].is_empty();
        out.push(Expectation {
            file,
            min_chars,
            min_words,
            require_title,
        });
    }
    out
}

#[test]
fn reader_fixtures_meet_minimums() {
    let expectations = load_expectations();
    assert!(!expectations.is_empty(), "no expectations loaded");

    for exp in expectations {
        let path = Path::new(&exp.file);
        let html = fs::read_to_string(path).expect("read fixture html");
        let article = reader::reader_mode_extract(&html, settings::Language::Italian)
            .unwrap_or_else(|| panic!("no article extracted from {}", exp.file));

        let title = article.title.trim();
        if exp.require_title {
            assert!(!title.is_empty(), "empty title for {}", exp.file);
            assert!(title != "No Title", "missing title for {}", exp.file);
        }

        let content = article.content.trim();
        let chars = content.chars().count();
        let words = content.split_whitespace().count();
        assert!(
            chars >= exp.min_chars,
            "{} chars too short (got {}, expected >= {})",
            exp.file,
            chars,
            exp.min_chars
        );
        assert!(
            words >= exp.min_words,
            "{} words too short (got {}, expected >= {})",
            exp.file,
            words,
            exp.min_words
        );
    }
}
