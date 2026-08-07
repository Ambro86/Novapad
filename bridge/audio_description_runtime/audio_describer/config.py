"""Minimal Omni-compatible configuration for the Sonarpad audio-description worker."""
from __future__ import annotations

import os
import sys

APP_NAME_BASE_UNTRANSLATED = "Sonarpad Audio Description"
APP_NAME_TRANSLATABLE = APP_NAME_BASE_UNTRANSLATED
APP_VERSION = "1.0.0"
GEMINI_API_KEY = None
GEMINI_MODEL_NAME = "gemini-3.5-flash-lite"
VERBOSITY_SHORT = "short"
VERBOSITY_STANDARD = "standard"
VERBOSITY_DETAILED = "detailed"
DEFAULT_VERBOSITY = VERBOSITY_DETAILED
LOG_FILE = "sonarpad_audio_description_bridge.log"
LOG_LEVEL = "INFO"
SUPPORTED_VIDEO_FORMATS = "*.mp4;*.avi;*.mkv;*.webm;*.mov;*.flv;*.wmv;*.mpeg;*.mpg;*.3gp;*.ogv;*.ts;*.m4v;*.divx"
DEFAULT_YOUTUBE_RESOLUTION = "360p"


def get_app_root() -> str:
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_locale_dir() -> str:
    return os.path.join(get_app_root(), "locale")


def get_app_data_dir() -> str:
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else get_app_root()
TEMP_DIR_NAME = "sonarpad_audio_description_temp"
