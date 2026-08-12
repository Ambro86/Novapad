from __future__ import annotations

import logging
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import audio_description_bridge as bridge
from audio_describer.utils import logger as bridge_logger


class AudioDescriptionLoggingTests(unittest.TestCase):
    def test_reset_log_file_truncates_the_existing_file_handler_in_place(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            log_path = Path(temp_dir) / "sonarpad_audio_description_bridge.log"
            log_path.write_text("old job\nold details\n", encoding="utf-8")
            handler = logging.FileHandler(log_path, encoding="utf-8")
            try:
                with mock.patch.object(bridge_logger, "_LOG_PATH", str(log_path)), mock.patch.object(
                    bridge_logger.app_logger, "handlers", [handler]
                ):
                    self.assertTrue(bridge_logger.reset_log_file())
                    handler.stream.write("new job\n")
                    handler.flush()

                self.assertEqual(log_path.read_text(encoding="utf-8"), "new job\n")
            finally:
                handler.close()

    def test_bridge_clears_log_once_when_a_real_request_job_starts(self):
        with mock.patch.object(
            bridge.sys, "argv", ["audio_description_bridge.py", "--request", "job.json"]
        ), mock.patch.object(bridge, "_read_request", return_value={"job": True}), mock.patch.object(
            bridge, "reset_log_file", return_value=True
        ) as reset_log, mock.patch.object(
            bridge, "run", return_value={"ok": True}
        ), mock.patch.object(
            bridge, "_emit"
        ), mock.patch.object(
            bridge.app_logger, "info"
        ):
            self.assertEqual(bridge.main(), 0)

        reset_log.assert_called_once_with()

    def test_self_test_does_not_clear_the_previous_job_log(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            model_path = Path(temp_dir) / "model.onnx"
            model_path.write_bytes(b"model")
            with mock.patch.object(
                bridge.sys, "argv", ["audio_description_bridge.py", "--self-test"]
            ), mock.patch.object(
                bridge.speech_detector,
                "get_bundled_model_path",
                return_value=str(model_path),
            ), mock.patch.object(bridge, "reset_log_file") as reset_log, mock.patch.object(
                bridge, "_emit"
            ):
                self.assertEqual(bridge.main(), 0)

            reset_log.assert_not_called()


if __name__ == "__main__":
    unittest.main()
