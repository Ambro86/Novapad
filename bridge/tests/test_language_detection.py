import json
import unittest
from unittest import mock

from audio_describer.core import audio_describer, language_detector


class _Response:
    def __init__(self, payload):
        self._body = json.dumps(payload).encode("utf-8")

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


class LanguageDetectorTests(unittest.TestCase):
    def test_google_source_language_and_confidence_are_returned(self):
        payload = [
            "Questa è una descrizione audio inglese sufficientemente lunga.",
            [["Questa è una descrizione audio inglese sufficientemente lunga.",
              "This is a sufficiently long English audio description."]],
            None,
            None,
            [["en"], None, [0.98], ["en"]],
            "en",
        ]
        opener = mock.Mock(return_value=_Response(payload))
        result = language_detector.detect_language(
            "This is a sufficiently long English audio description.",
            target_language="it",
            opener=opener,
        )
        self.assertEqual(result.language, "en")
        self.assertEqual(result.confidence, 0.98)
        request = opener.call_args.args[0]
        self.assertIn("translate-pa.googleapis.com/v1/translate", request.full_url)
        self.assertIn("query.source_language=auto", request.full_url)
        self.assertIn("query.target_language=it", request.full_url)
        self.assertEqual(request.get_header("Content-type"), "application/json+protobuf")

    def test_short_text_is_not_detected(self):
        opener = mock.Mock()
        self.assertIsNone(language_detector.detect_language("A car.", opener=opener))
        opener.assert_not_called()

    def test_short_english_description_is_still_checked(self):
        opener = mock.Mock(return_value=_Response([
            "I nani osservano con gioia",
            [["I nani osservano con gioia", "Dwarves observe joyously"]],
            None,
            None,
            [["en"], None, [0.99], ["en"]],
            "en",
        ]))
        result = language_detector.detect_language(
            "Dwarves observe joyously", target_language="it", opener=opener
        )
        self.assertEqual(result.language, "en")
        opener.assert_called_once()

    def test_language_metadata_fallback_is_supported(self):
        opener = mock.Mock(return_value=_Response([
            "hello friend how are you",
            [["hello friend how are you", "ciao amico come stai"]],
            None,
            None,
            [["it"], None, [1], ["it"]],
            None,
        ]))
        result = language_detector.detect_language(
            "ciao amico come stai", target_language="en", opener=opener
        )
        self.assertEqual(result.language, "it")
        self.assertEqual(result.confidence, 1.0)

    def test_locale_variants_match(self):
        self.assertTrue(language_detector.languages_match("pt-BR", "pt"))
        self.assertTrue(language_detector.languages_match("zh-CN", "zh-TW"))


class DescriptionLanguageCorrectionTests(unittest.TestCase):
    def test_one_english_entry_is_not_hidden_by_italian_entries(self):
        descriptions = [
            (1.0, 2.0, "Amici festosi accorrono"),
            (4.0, 5.0, "Dwarves observe joyously"),
            (7.0, 8.0, "Amici sorridono calorosamente"),
        ]
        corrected_json = json.dumps({
            "descriptions": [
                {"index": 0, "text": "I nani osservano con gioia"},
            ]
        })

        def detect(text, _target):
            language = "en" if text.startswith("Dwarves") else "it"
            return language_detector.LanguageDetection(language, 1.0)

        with mock.patch.object(
            audio_describer.config_model, "get_setting", return_value="it"
        ), mock.patch.object(
            language_detector, "detect_language", side_effect=detect
        ) as detect_mock, mock.patch.object(
            audio_describer.gemini, "build_generation_config", return_value=object()
        ), mock.patch.object(
            audio_describer.gemini,
            "generate_content_with_retry",
            return_value=mock.Mock(),
        ) as generate, mock.patch.object(
            audio_describer.gemini, "save_raw_ai_output"
        ), mock.patch.object(
            audio_describer.gemini, "log_token_usage", return_value=None
        ), mock.patch.object(
            audio_describer.gemini,
            "process_gemini_response",
            return_value=(corrected_json, True),
        ):
            result, _usage = audio_describer._correct_description_language(
                mock.Mock(), "gemini-3.5-flash-lite", descriptions
            )

        self.assertEqual(detect_mock.call_count, 3)
        self.assertEqual(result[0], descriptions[0])
        self.assertEqual(result[1], (4.0, 5.0, "I nani osservano con gioia"))
        self.assertEqual(result[2], descriptions[2])
        sent_prompt = generate.call_args.kwargs["contents"][0]
        self.assertIn("Dwarves observe joyously", sent_prompt)
        self.assertNotIn("Amici festosi accorrono", sent_prompt)

    def test_wrong_language_triggers_one_index_preserving_retry(self):
        descriptions = [
            (1.0, 2.0, "A man enters the room."),
            (4.0, 5.0, "He closes the door."),
        ]
        response = mock.Mock()
        corrected_json = json.dumps({
            "descriptions": [
                {"index": 0, "text": "Un uomo entra nella stanza."},
                {"index": 1, "text": "Chiude la porta."},
            ]
        })
        with mock.patch.object(
            audio_describer.config_model, "get_setting", return_value="it"
        ), mock.patch.object(
            language_detector,
            "detect_language",
            return_value=language_detector.LanguageDetection("en", 1.0),
        ), mock.patch.object(
            audio_describer.gemini, "build_generation_config", return_value=object()
        ), mock.patch.object(
            audio_describer.gemini, "generate_content_with_retry", return_value=response
        ) as generate, mock.patch.object(
            audio_describer.gemini, "save_raw_ai_output"
        ), mock.patch.object(
            audio_describer.gemini, "log_token_usage", return_value={"total_tokens": 12}
        ), mock.patch.object(
            audio_describer.gemini,
            "process_gemini_response",
            return_value=(corrected_json, True),
        ):
            result, usage = audio_describer._correct_description_language(
                mock.Mock(), "gemini-2.5-flash-lite", descriptions
            )

        self.assertEqual([(a, b) for a, b, _ in result], [(1.0, 2.0), (4.0, 5.0)])
        self.assertEqual([text for _, _, text in result], [
            "Un uomo entra nella stanza.", "Chiude la porta."
        ])
        self.assertEqual(usage, {"total_tokens": 12})
        self.assertEqual(generate.call_count, 1)

    def test_non_lite_models_are_checked_too(self):
        descriptions = [(0.0, 2.0, "An Italian man enters the room.")]
        with mock.patch.object(
            audio_describer.config_model, "get_setting", return_value="en"
        ), mock.patch.object(
            language_detector,
            "detect_language",
            return_value=language_detector.LanguageDetection("en", 1.0),
        ) as detect:
            result, usage = audio_describer._correct_description_language(
                mock.Mock(), "gemini-2.5-pro", descriptions
            )
        self.assertEqual(result, descriptions)
        self.assertIsNone(usage)
        detect.assert_called_once()


class PromptLanguageCoverageTests(unittest.TestCase):
    def test_all_sonarpad_languages_have_explicit_prompt_names(self):
        expected = {
            "it": "Italian", "en": "English", "de": "German",
            "es": "Spanish", "pt": "Portuguese",
            "pt-BR": "Brazilian Portuguese", "sv": "Swedish",
            "vi": "Vietnamese", "cs": "Czech", "pl": "Polish",
            "fr": "French", "sr": "Serbian", "uk": "Ukrainian",
            "lt": "Lithuanian", "ru": "Russian", "zh": "Chinese",
            "hi": "Hindi",
        }
        for code, expected_name in expected.items():
            with self.subTest(code=code), mock.patch.object(
                audio_describer.config_model,
                "get_setting",
                side_effect=lambda key, code=code: (
                    code if key == "application_language" else False
                ),
            ):
                _detector_code, name, _examples = (
                    audio_describer._target_language_details()
                )
                self.assertEqual(name, expected_name)

    def test_prompt_requires_language_for_descriptions_and_glossary(self):
        settings = {
            "application_language": "de",
            "enable_character_glossary": True,
            "gemini_description_verbosity": "standard",
        }
        with mock.patch.object(
            audio_describer.config_model,
            "get_setting",
            side_effect=lambda key: settings.get(key),
        ):
            system_prompt, user_prompt = audio_describer._build_unified_prompts(
                "", "gemini-test"
            )
        self.assertIn("written entirely in German", system_prompt)
        self.assertIn("character_glossary[].description", user_prompt)
        self.assertIn("Ein Auto rast die Straße entlang.", system_prompt)
        self.assertNotIn("Target Language for `description_text`: English", user_prompt)


class GlossaryLanguageCorrectionTests(unittest.TestCase):
    def test_glossary_description_is_corrected_without_changing_id_or_name(self):
        glossary = [{
            "id": "woman_red_coat",
            "description": "A woman wearing a red coat.",
            "name": "Anna",
        }]
        corrected_json = json.dumps({
            "entries": [{
                "index": 0,
                "description": "Una donna con un cappotto rosso.",
            }]
        })
        settings = {
            "application_language": "it",
            "enable_character_glossary": True,
        }
        with mock.patch.object(
            audio_describer.config_model,
            "get_setting",
            side_effect=lambda key: settings.get(key),
        ), mock.patch.object(
            language_detector,
            "detect_language",
            return_value=language_detector.LanguageDetection("en", 1.0),
        ), mock.patch.object(
            audio_describer.gemini, "build_generation_config", return_value=object()
        ), mock.patch.object(
            audio_describer.gemini, "generate_content_with_retry", return_value=mock.Mock()
        ), mock.patch.object(
            audio_describer.gemini, "save_raw_ai_output"
        ), mock.patch.object(
            audio_describer.gemini, "log_token_usage", return_value={"total_tokens": 5}
        ), mock.patch.object(
            audio_describer.gemini,
            "process_gemini_response",
            return_value=(corrected_json, True),
        ):
            corrected, usage = audio_describer._correct_glossary_language(
                mock.Mock(), "gemini-test", glossary
            )
        self.assertEqual(corrected[0]["id"], "woman_red_coat")
        self.assertEqual(corrected[0]["name"], "Anna")
        self.assertEqual(
            corrected[0]["description"],
            "Una donna con un cappotto rosso.",
        )
        self.assertEqual(usage, {"total_tokens": 5})


if __name__ == "__main__":
    unittest.main()
