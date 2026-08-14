from __future__ import annotations

import logging
import unittest
from pathlib import Path

import audio_description_bridge as bridge
from audio_describer.utils import logger as bridge_logger


class AudioDescriptionLoggingTests(unittest.TestCase):
    def test_worker_logger_uses_only_stderr_no_file_handler(self):
        self.assertTrue(bridge_logger.app_logger.handlers)
        self.assertFalse(
            any(isinstance(handler, logging.FileHandler) for handler in bridge_logger.app_logger.handlers)
        )

    def test_worker_has_no_separate_audio_description_log_path(self):
        logger_source = Path(bridge_logger.__file__).read_text(encoding="utf-8")
        bridge_source = Path(bridge.__file__).read_text(encoding="utf-8")
        self.assertNotIn("sonarpad_audio_description_bridge.log", logger_source)
        self.assertNotIn("reset_log_file", bridge_source)
        self.assertNotIn("get_log_file_path", bridge_source)

    def test_rust_host_streams_worker_stderr_into_main_sonarpad_log(self):
        root = Path(__file__).resolve().parents[2]
        rust_source = (root / "src" / "tools" / "audio_description_bridge.rs").read_text(
            encoding="utf-8"
        )
        self.assertIn('crate::log_debug(&format!("audio_description.worker {line}"))', rust_source)
        self.assertIn("STDERR_TAIL_LINES", rust_source)
        self.assertNotIn("Audio description worker stderr: {}", rust_source)


if __name__ == "__main__":
    unittest.main()
