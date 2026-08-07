"""Pytest bootstrap for Sonarpad's embedded audio-description runtime."""
from __future__ import annotations

import sys
from pathlib import Path

BRIDGE_DIR = Path(__file__).resolve().parents[1]
RUNTIME_DIR = BRIDGE_DIR / "audio_description_runtime"
for path in (BRIDGE_DIR, RUNTIME_DIR):
    value = str(path)
    if value not in sys.path:
        sys.path.insert(0, value)
