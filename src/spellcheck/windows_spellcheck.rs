use crate::accessibility::{from_wide, to_wide};
use crate::log_debug;
use windows::Win32::Foundation::S_FALSE;
use windows::Win32::Globalization::{
    IEnumSpellingError, ISpellChecker, ISpellCheckerFactory, ISpellingError, SpellCheckerFactory,
};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoTaskMemFree,
    IEnumString,
};
use windows::core::PCWSTR;

#[derive(Default)]
pub struct WindowsSpellChecker {
    factory: Option<ISpellCheckerFactory>,
    checker: Option<ISpellChecker>,
    language: Option<String>,
}

impl WindowsSpellChecker {
    pub fn supported_languages(&mut self) -> Vec<String> {
        if !self.ensure_com() {
            return Vec::new();
        }
        let factory = match self.factory() {
            Some(factory) => factory,
            None => return Vec::new(),
        };
        unsafe { factory.SupportedLanguages() }
            .ok()
            .map(enum_string_to_vec)
            .unwrap_or_default()
    }

    pub fn is_language_supported(&mut self, language: &str) -> bool {
        if !self.ensure_com() {
            return false;
        }
        let factory = match self.factory() {
            Some(factory) => factory,
            None => return false,
        };
        let wide = to_wide(language);
        unsafe { factory.IsSupported(PCWSTR(wide.as_ptr())) }
            .ok()
            .map(|value| value.as_bool())
            .unwrap_or(false)
    }

    pub fn set_language(&mut self, language: &str) -> bool {
        if self.language.as_deref() == Some(language) && self.checker.is_some() {
            return true;
        }
        if !self.ensure_com() {
            return false;
        }
        let factory = match self.factory() {
            Some(factory) => factory,
            None => return false,
        };
        let wide = to_wide(language);
        let checker = unsafe { factory.CreateSpellChecker(PCWSTR(wide.as_ptr())) };
        match checker {
            Ok(checker) => {
                self.checker = Some(checker);
                self.language = Some(language.to_string());
                true
            }
            Err(err) => {
                log_debug(&format!(
                    "Spellcheck: failed to set language {language}: {err:?}"
                ));
                false
            }
        }
    }

    pub fn check_text(&mut self, text: &str) -> Vec<Misspelling> {
        let Some(checker) = self.checker.as_ref() else {
            return Vec::new();
        };
        let wide = to_wide(text);
        let enum_errors = unsafe { checker.Check(PCWSTR(wide.as_ptr())) };
        let enum_errors = match enum_errors {
            Ok(enum_errors) => enum_errors,
            Err(err) => {
                log_debug(&format!("Spellcheck: Check failed: {err:?}"));
                return Vec::new();
            }
        };
        collect_misspellings(text, &enum_errors)
    }

    pub fn suggestions(&mut self, word: &str) -> Vec<String> {
        let Some(checker) = self.checker.as_ref() else {
            return Vec::new();
        };
        let wide = to_wide(word);
        unsafe { checker.Suggest(PCWSTR(wide.as_ptr())) }
            .ok()
            .map(enum_string_to_vec)
            .unwrap_or_default()
    }

    pub fn add_to_dictionary(&mut self, word: &str) -> bool {
        let Some(checker) = self.checker.as_ref() else {
            return false;
        };
        let wide = to_wide(word);
        unsafe { checker.Add(PCWSTR(wide.as_ptr())) }.is_ok()
    }

    pub fn ignore_once(&mut self, word: &str) -> bool {
        let Some(checker) = self.checker.as_ref() else {
            return false;
        };
        let wide = to_wide(word);
        unsafe { checker.Ignore(PCWSTR(wide.as_ptr())) }.is_ok()
    }

    fn ensure_com(&self) -> bool {
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_ok() || hr.0 == 0x80010106_u32 as i32 {
                true
            } else {
                log_debug(&format!("Spellcheck: CoInitializeEx failed: {hr:?}"));
                false
            }
        }
    }

    fn factory(&mut self) -> Option<ISpellCheckerFactory> {
        if self.factory.is_none() {
            let factory = unsafe { CoCreateInstance(&SpellCheckerFactory, None, CLSCTX_ALL) };
            match factory {
                Ok(factory) => {
                    self.factory = Some(factory);
                }
                Err(err) => {
                    log_debug(&format!("Spellcheck: factory creation failed: {err:?}"));
                }
            }
        }
        self.factory.clone()
    }
}

#[derive(Clone, Debug)]
pub struct Misspelling {
    pub start: usize,
    pub end: usize,
}

fn collect_misspellings(text: &str, enum_errors: &IEnumSpellingError) -> Vec<Misspelling> {
    let mut out = Vec::new();
    loop {
        let mut error: Option<ISpellingError> = None;
        let hr = unsafe { enum_errors.Next(&mut error) };
        if hr == S_FALSE || error.is_none() {
            break;
        }
        let Some(error) = error else {
            break;
        };
        let start = unsafe { error.StartIndex() }.ok();
        let length = unsafe { error.Length() }.ok();
        let (Some(start), Some(length)) = (start, length) else {
            continue;
        };
        let start_utf8 = utf16_offset_to_utf8_byte_offset(text, start);
        let end_utf8 = utf16_offset_to_utf8_byte_offset(text, start.saturating_add(length));
        if start_utf8 < end_utf8 && end_utf8 <= text.len() {
            out.push(Misspelling {
                start: start_utf8,
                end: end_utf8,
            });
        }
    }
    out
}

fn enum_string_to_vec(enum_string: IEnumString) -> Vec<String> {
    let mut out = Vec::new();
    loop {
        let mut fetched = 0u32;
        let mut values = [windows::core::PWSTR::null()];
        let hr = unsafe { enum_string.Next(&mut values, Some(&mut fetched)) };
        if hr == S_FALSE || fetched == 0 {
            break;
        }
        let ptr = values[0];
        if !ptr.is_null() {
            let text = unsafe { from_wide(ptr.0) };
            if !text.is_empty() {
                out.push(text);
            }
            unsafe {
                CoTaskMemFree(Some(ptr.0 as *const _));
            }
        }
    }
    out
}

pub fn utf16_offset_to_utf8_byte_offset(text: &str, utf16_units: u32) -> usize {
    let mut utf16_count = 0usize;
    for (byte_idx, ch) in text.char_indices() {
        let units = ch.len_utf16();
        let next = utf16_count + units;
        if utf16_units as usize <= next {
            if utf16_units as usize == next {
                return byte_idx + ch.len_utf8();
            }
            return byte_idx;
        }
        utf16_count = next;
    }
    text.len()
}

#[cfg(test)]
mod tests {
    use super::utf16_offset_to_utf8_byte_offset;

    #[test]
    fn utf16_offset_ascii() {
        let text = "hello";
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 0), 0);
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 1), 1);
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 5), 5);
    }

    #[test]
    fn utf16_offset_accents() {
        let text = "citta";
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 5), 5);
        let text = "città";
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 4), 4);
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 5), 6);
    }

    #[test]
    fn utf16_offset_emoji() {
        let text = "a😀b";
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 1), 1);
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 3), 5);
        let text = "😀";
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 0), 0);
        assert_eq!(utf16_offset_to_utf8_byte_offset(text, 2), 4);
    }
}
