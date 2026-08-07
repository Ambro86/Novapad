"""Run the portable Python portion of Sonarpad's audio-description tests."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

BRIDGE_DIR = Path(__file__).resolve().parent
RUNTIME_DIR = BRIDGE_DIR / "audio_description_runtime"
TESTS_DIR = BRIDGE_DIR / "tests"
for path in (BRIDGE_DIR, RUNTIME_DIR):
    value = str(path)
    if value not in sys.path:
        sys.path.insert(0, value)


def main() -> int:
    suite = unittest.defaultTestLoader.discover(str(TESTS_DIR), pattern="test_*.py")
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    print(f"Sonarpad audio-description Python tests executed: {result.testsRun}")
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
