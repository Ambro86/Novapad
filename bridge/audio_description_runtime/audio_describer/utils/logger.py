"""Logging for the headless worker; stdout is reserved for its JSON protocol.

The worker writes only to stderr. The Sonarpad Windows host captures stderr
line by line and merges it into the application's main log.txt, so there is no
separate audio-description log file.
"""
from __future__ import annotations

import logging
import sys

app_logger = logging.getLogger("sonarpad.audio_description")
if not app_logger.handlers:
    app_logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(levelname)s %(name)s:%(lineno)d %(message)s")
    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setFormatter(formatter)
    app_logger.addHandler(stderr_handler)
    app_logger.propagate = False


def update_log_level() -> None:
    from audio_describer.models import config_model

    level = str(config_model.get_setting("logging_level") or "INFO").upper()
    app_logger.setLevel(getattr(logging, level, logging.INFO))
