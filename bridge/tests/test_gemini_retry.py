import unittest
from types import SimpleNamespace
from unittest import mock

from audio_describer.core import gemini_helpers


class GeminiRetryTests(unittest.TestCase):
    def setUp(self):
        self.lazy_import = mock.patch.object(gemini_helpers, "_lazy_import_gemini_sdk")
        self.lazy_import.start()

    def tearDown(self):
        self.lazy_import.stop()
        gemini_helpers.set_quota_decision_handler(None)

    def test_transient_failure_retries_beyond_thirty_with_fixed_delay(self):
        failures = [TimeoutError("temporary timeout") for _ in range(35)]
        operation = mock.Mock(side_effect=failures + ["ok"])

        with mock.patch.object(
            gemini_helpers, "is_retryable_transient_error", return_value=True
        ), mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            result = gemini_helpers.run_with_retry(operation)

        self.assertEqual(result, "ok")
        self.assertEqual(operation.call_count, 36)
        self.assertEqual(sleep_mock.call_count, 35)
        self.assertTrue(
            all(call.args == (5,) for call in sleep_mock.call_args_list)
        )

    def test_non_transient_failure_still_stops_immediately(self):
        operation = mock.Mock(side_effect=ValueError("invalid request"))

        with mock.patch.object(
            gemini_helpers, "is_retryable_transient_error", return_value=False
        ), mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            with self.assertRaises(ValueError):
                gemini_helpers.run_with_retry(operation)

        self.assertEqual(operation.call_count, 1)
        sleep_mock.assert_not_called()

    def test_prepaid_credits_depleted_stops_generic_retry_immediately(self):
        error = RuntimeError(
            "429 RESOURCE_EXHAUSTED: Your prepayment credits are depleted. "
            "Please go to AI Studio to manage your project and billing."
        )
        operation = mock.Mock(side_effect=error)

        with mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            with self.assertRaises(RuntimeError) as raised:
                gemini_helpers.run_with_retry(
                    operation, operation_label="video upload"
                )

        self.assertIs(raised.exception, error)
        self.assertEqual(operation.call_count, 1)
        sleep_mock.assert_not_called()

    def test_prepaid_credits_depleted_does_not_offer_model_switch(self):
        error = RuntimeError(
            "429 RESOURCE_EXHAUSTED: Your prepayment credits are depleted."
        )
        client = mock.Mock()
        client.models.generate_content.side_effect = error
        decision_handler = mock.Mock(return_value="gemini-new")
        gemini_helpers.set_quota_decision_handler(decision_handler)

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ), mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            with self.assertRaises(RuntimeError) as raised:
                gemini_helpers.generate_content_with_retry(
                    client, "gemini-old", [], object()
                )

        self.assertIs(raised.exception, error)
        self.assertEqual(client.models.generate_content.call_count, 1)
        decision_handler.assert_not_called()
        client.models.get.assert_not_called()
        sleep_mock.assert_not_called()

    def test_prepaid_credits_detection_follows_exception_chain(self):
        inner = RuntimeError(
            "429 RESOURCE_EXHAUSTED: Your prepayment credits are depleted."
        )
        outer = RuntimeError("upload wrapper failed")
        outer.__cause__ = inner

        self.assertTrue(
            gemini_helpers.is_prepaid_credits_depleted_error(outer)
        )

    def test_gemini_504_deadline_exceeded_retries_same_request(self):
        deadline = RuntimeError(
            "504 DEADLINE_EXCEEDED: Deadline expired before operation could complete."
        )
        client = mock.Mock()
        client.models.generate_content.side_effect = [deadline, "ok"]

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ), mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            result = gemini_helpers.generate_content_with_retry(
                client, "gemini-test", [], object()
            )

        self.assertEqual(result, "ok")
        self.assertEqual(client.models.generate_content.call_count, 2)
        sleep_mock.assert_called_once_with(gemini_helpers.RETRY_DELAY_SEC)

    def test_every_structured_http_5xx_is_retryable(self):
        for code in (500, 501, 502, 503, 504, 507, 599):
            error = RuntimeError("server failure")
            error.code = code
            with self.subTest(code=code):
                self.assertTrue(
                    gemini_helpers.is_retryable_transient_error(error)
                )

    def test_http_5xx_in_unstructured_message_is_retryable(self):
        self.assertTrue(
            gemini_helpers.is_retryable_transient_error(
                RuntimeError("502 Bad Gateway")
            )
        )

    def test_permanent_http_4xx_is_not_made_retryable(self):
        for code in (400, 401, 403, 404, 413, 422):
            error = RuntimeError("permanent client failure")
            error.code = code
            with self.subTest(code=code):
                self.assertFalse(
                    gemini_helpers.is_retryable_transient_error(error)
                )

    def test_prohibited_content_without_candidate_retries_same_chunk(self):
        blocked = SimpleNamespace(
            prompt_feedback=SimpleNamespace(
                block_reason=SimpleNamespace(name="PROHIBITED_CONTENT"),
                block_reason_message=None,
            ),
            candidates=[],
        )
        success = SimpleNamespace(prompt_feedback=None, candidates=[])
        client = mock.Mock()
        client.models.generate_content.side_effect = [blocked, success]

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ), mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            result = gemini_helpers.generate_content_with_retry(
                client,
                "gemini-test",
                [],
                object(),
                prohibited_content_max_attempts=2,
            )

        self.assertIs(result, success)
        self.assertEqual(client.models.generate_content.call_count, 2)
        sleep_mock.assert_called_once_with(gemini_helpers.RETRY_DELAY_SEC)

    def test_prohibited_content_hands_back_after_configured_two_attempts(self):
        blocked = SimpleNamespace(
            prompt_feedback=SimpleNamespace(
                block_reason=SimpleNamespace(name="PROHIBITED_CONTENT"),
                block_reason_message=None,
            ),
            candidates=[],
        )
        client = mock.Mock()
        client.models.generate_content.return_value = blocked

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ), mock.patch.object(gemini_helpers.time, "sleep") as sleep_mock:
            with self.assertRaises(gemini_helpers.ContentBlockedError):
                gemini_helpers.generate_content_with_retry(
                    client,
                    "gemini-test",
                    [],
                    object(),
                    prohibited_content_max_attempts=2,
                )

        self.assertEqual(client.models.generate_content.call_count, 2)
        sleep_mock.assert_called_once_with(gemini_helpers.RETRY_DELAY_SEC)

    def test_quota_can_switch_model_without_losing_current_request(self):
        quota_error = RuntimeError("429 RESOURCE_EXHAUSTED: quota exceeded")
        generated_models = []

        def generate(*, model, contents, config):
            generated_models.append(model)
            if model == "gemini-old":
                raise quota_error
            return "ok"

        client = mock.Mock()
        client.models.generate_content.side_effect = generate
        client.models.get.return_value = mock.Mock(
            supported_actions=["generateContent"]
        )
        gemini_helpers.set_quota_decision_handler(
            lambda current_model, error: "gemini-new"
        )

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ):
            result = gemini_helpers.generate_content_with_retry(
                client, "gemini-old", [], object()
            )

        self.assertEqual(result, "ok")
        self.assertEqual(generated_models, ["gemini-old", "gemini-new"])


    def test_quota_switch_rejects_same_model_in_models_prefix_form(self):
        client = mock.Mock()
        client.models.generate_content.side_effect = RuntimeError(
            "429 RESOURCE_EXHAUSTED: quota exceeded"
        )
        client.models.get.return_value = mock.Mock(
            supported_actions=["generateContent"]
        )
        gemini_helpers.set_quota_decision_handler(
            lambda current_model, error: "models/gemini-old"
        )

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ), self.assertRaises(gemini_helpers.GeminiRetryCancelledError):
            gemini_helpers.generate_content_with_retry(
                client, "gemini-old", [], object()
            )

        self.assertEqual(client.models.generate_content.call_count, 1)

    def test_quota_dialog_cancel_stops_immediately(self):
        client = mock.Mock()
        client.models.generate_content.side_effect = RuntimeError(
            "429 RESOURCE_EXHAUSTED: quota exceeded"
        )
        gemini_helpers.set_quota_decision_handler(
            lambda current_model, error: False
        )

        with mock.patch.object(
            gemini_helpers.config_model, "get_setting", return_value=""
        ), self.assertRaises(gemini_helpers.GeminiRetryCancelledError):
            gemini_helpers.generate_content_with_retry(
                client, "gemini-old", [], object()
            )

        self.assertEqual(client.models.generate_content.call_count, 1)


if __name__ == "__main__":
    unittest.main()
