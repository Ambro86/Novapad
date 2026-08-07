import unittest
from types import SimpleNamespace
from unittest import mock

from audio_describer.core import gemini_helpers


class _Models:
    def __init__(self, model=None, error=None):
        self.model = model
        self.error = error
        self.requested = None

    def get(self, *, model):
        self.requested = model
        if self.error:
            raise self.error
        return self.model

    def list(self):
        return self.model


class GeminiModelTests(unittest.TestCase):
    def setUp(self):
        gemini_helpers._VALIDATED_GENERATE_MODELS.clear()
        self.lazy_import = mock.patch.object(gemini_helpers, "_lazy_import_gemini_sdk")
        self.lazy_import.start()

    def tearDown(self):
        self.lazy_import.stop()

    def test_migrates_31_pro_to_preview_endpoint(self):
        self.assertEqual(
            gemini_helpers.normalize_model_id("gemini-3.1-pro"),
            "gemini-3.1-pro-preview",
        )

    def test_client_uses_eight_minute_request_timeout(self):
        fake_genai = mock.Mock()
        fake_types = mock.Mock()
        timeout_options = object()
        fake_types.HttpOptions.return_value = timeout_options

        with mock.patch.object(gemini_helpers, "genai", fake_genai), \
             mock.patch.object(gemini_helpers, "types", fake_types):
            gemini_helpers._create_gemini_client("test-key")

        fake_types.HttpOptions.assert_called_once_with(
            timeout=8 * 60 * 1000
        )
        fake_genai.Client.assert_called_once_with(
            api_key="test-key", http_options=timeout_options
        )

    def test_accepts_model_supporting_generate_content(self):
        models = _Models(SimpleNamespace(supported_actions=["generateContent"]))
        client = SimpleNamespace(models=models)

        result = gemini_helpers.validate_model_for_generate_content(
            "models/gemini-3.6-flash", client=client
        )

        self.assertEqual(result, "gemini-3.6-flash")
        self.assertEqual(models.requested, "gemini-3.6-flash")

    def test_rejects_model_without_generate_content(self):
        models = _Models(SimpleNamespace(supported_actions=["bidiGenerateContent"]))
        client = SimpleNamespace(models=models)

        with self.assertRaises(gemini_helpers.GeminiAPIError):
            gemini_helpers.validate_model_for_generate_content(
                "gemini-live-only", client=client
            )

    def test_retries_transient_model_verification(self):
        model_info = SimpleNamespace(supported_actions=["generateContent"])
        models = mock.Mock()
        models.get.side_effect = [TimeoutError("temporary timeout"), model_info]
        client = SimpleNamespace(models=models)
        status_messages = []

        with mock.patch("audio_describer.core.gemini_helpers.time.sleep") as sleep:
            result = gemini_helpers.validate_model_for_generate_content(
                "gemini-3-flash-preview",
                client=client,
                status_callback=status_messages.append,
            )

        self.assertEqual(result, "gemini-3-flash-preview")
        self.assertEqual(models.get.call_count, 2)
        sleep.assert_called_once_with(gemini_helpers.RETRY_DELAY_SEC)
        self.assertTrue(any("attempt 2" in message for message in status_messages))

    def test_lists_current_general_generate_models_from_api(self):
        models = _Models([
            SimpleNamespace(
                name="models/gemini-3.6-flash",
                supported_actions=["generateContent"],
            ),
            SimpleNamespace(
                name="models/gemini-future-image",
                supported_actions=["generateContent"],
            ),
            SimpleNamespace(
                name="models/gemini-live-only",
                supported_actions=["bidiGenerateContent"],
            ),
            SimpleNamespace(
                name="models/text-embedding-new",
                supported_actions=["embedContent"],
            ),
        ])

        result = gemini_helpers.list_generate_content_models(
            client=SimpleNamespace(models=models)
        )

        self.assertEqual(result, ["gemini-3.6-flash"])


if __name__ == "__main__":
    unittest.main()
