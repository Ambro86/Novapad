"""Logging for the headless worker; stdout is reserved for its JSON protocol."""
from __future__ import annotations

import logging
import os
import sys
import tempfile

_LOG_PATH = os.path.join(tempfile.gettempdir(), "sonarpad_audio_description_bridge.log")
app_logger = logging.getLogger("sonarpad.audio_description")
if not app_logger.handlers:
    app_logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s:%(lineno)d %(message)s")
    try:
        file_handler = logging.FileHandler(_LOG_PATH, encoding="utf-8")
        file_handler.setFormatter(formatter)
        app_logger.addHandler(file_handler)
    except OSError:
        pass
    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setFormatter(formatter)
    app_logger.addHandler(stderr_handler)
    app_logger.propagate = False


def get_log_file_path() -> str:
    return _LOG_PATH


def reset_log_file() -> bool:
    """Clear the bridge log in-place before a new audio-description job.

    The FileHandler is created when the worker module is imported, so truncate
    its already-open stream instead of replacing the file underneath it. A
    logging maintenance failure must never prevent an audio-description job.
    """
    expected_path = os.path.abspath(_LOG_PATH)
    for handler in app_logger.handlers:
        if not isinstance(handler, logging.FileHandler):
            continue
        if os.path.abspath(handler.baseFilename) != expected_path:
            continue
        handler.acquire()
        try:
            stream = handler.stream
            if stream is None:
                return False
            handler.flush()
            stream.seek(0)
            stream.truncate(0)
            stream.seek(0, os.SEEK_END)
            return True
        except (OSError, ValueError):
            return False
        finally:
            handler.release()

    try:
        with open(_LOG_PATH, "w", encoding="utf-8"):
            pass
        return True
    except OSError:
        return False


def update_log_level() -> None:
    from audio_describer.models import config_model

    level = str(config_model.get_setting("logging_level") or "INFO").upper()
    app_logger.setLevel(getattr(logging, level, logging.INFO))
