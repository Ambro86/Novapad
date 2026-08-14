import pathlib
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[2]
SETTINGS = (ROOT / "src" / "settings.rs").read_text(encoding="utf-8")
OPTIONS = (ROOT / "src" / "app_windows" / "options_window.rs").read_text(encoding="utf-8")
MAIN = (ROOT / "src" / "main.rs").read_text(encoding="utf-8")
TRANSLATOR = (ROOT / "src" / "translator.rs").read_text(encoding="utf-8")
AUDIO_DESCRIPTION_WINDOW = (
    ROOT / "src" / "app_windows" / "audio_description_window.rs"
).read_text(encoding="utf-8")


class WindowsGeminiLatestModelsOptionTests(unittest.TestCase):
    def test_ai_default_uses_flash_latest_without_changing_audio_description_default(self):
        self.assertIn(
            'pub const DEFAULT_GEMINI_MODEL: &str = "gemini-flash-latest";', SETTINGS
        )
        self.assertIn(
            'pub const GEMINI_FLASH_LITE_LATEST_MODEL: &str = "gemini-flash-lite-latest";',
            SETTINGS,
        )
        self.assertIn(
            'pub const DEFAULT_AUDIO_DESCRIPTION_GEMINI_MODEL: &str = "gemini-3.5-flash-lite";',
            SETTINGS,
        )
        self.assertIn('const LEGACY_DEFAULT_GEMINI_MODEL: &str = "gemini-3.5-flash";', SETTINGS)

    def test_options_start_with_two_friendly_latest_aliases(self):
        self.assertIn('const GEMINI_FLASH_LATEST_LABEL: &str = "Gemini Flash (latest)";', OPTIONS)
        self.assertIn(
            'const GEMINI_FLASH_LITE_LATEST_LABEL: &str = "Gemini Flash-Lite (latest)";', OPTIONS
        )
        self.assertIn(
            'for model in [DEFAULT_GEMINI_MODEL, GEMINI_FLASH_LITE_LATEST_MODEL]', OPTIONS
        )
        self.assertIn('gemini_model_id_from_combo_text', OPTIONS)

    def test_refresh_still_fetches_and_loads_full_api_model_list(self):
        self.assertIn('pub(crate) fn fetch_gemini_models_for_key', OPTIONS)
        self.assertIn('for model in list', OPTIONS)
        self.assertIn('models.push(clean_name);', OPTIONS)
        self.assertIn('for (index, model) in models.iter().enumerate()', OPTIONS)
        self.assertIn('OPTIONS_ID_GEMINI_REFRESH_MODELS =>', OPTIONS)
        self.assertIn('refresh_gemini_models(hwnd);', OPTIONS)

    def test_latest_aliases_keep_gemini3_fallback_behavior(self):
        self.assertIn('fn gemini_model_uses_compat_fallback', MAIN)
        self.assertIn('model == crate::settings::DEFAULT_GEMINI_MODEL', MAIN)
        self.assertIn('model == crate::settings::GEMINI_FLASH_LITE_LATEST_MODEL', MAIN)

    def test_generate_content_omits_deprecated_sampling_parameters(self):
        generate_start = TRANSLATOR.index('async fn generate_text')
        generate_end = TRANSLATOR.index('pub async fn summarize_same_language', generate_start)
        generate = TRANSLATOR[generate_start:generate_end]
        self.assertNotIn('"temperature"', generate)
        self.assertNotIn('"top_p"', generate)
        self.assertNotIn('"top_k"', generate)

    def test_audio_description_window_not_switched_to_ai_latest_aliases(self):
        self.assertIn('DEFAULT_AUDIO_DESCRIPTION_GEMINI_MODEL', AUDIO_DESCRIPTION_WINDOW)
        self.assertNotIn('GEMINI_FLASH_LITE_LATEST_MODEL', AUDIO_DESCRIPTION_WINDOW)
        self.assertNotIn('GEMINI_FLASH_LATEST_LABEL', AUDIO_DESCRIPTION_WINDOW)


if __name__ == "__main__":
    unittest.main()
