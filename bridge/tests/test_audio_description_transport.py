from __future__ import annotations

import tempfile
import unittest
from unittest import mock

from audio_describer.core import audio_describer


class AudioDescriptionTransportTests(unittest.TestCase):
    def test_22_5_mb_chunk_stays_inline_instead_of_files_api(self):
        with tempfile.NamedTemporaryFile(suffix=".mkv") as tmp:
            tmp.truncate(int(22.5 * 1024 * 1024))
            inline_part = object()
            with mock.patch.object(
                audio_describer, "_build_inline_video_part", return_value=inline_part
            ) as inline, mock.patch.object(
                audio_describer, "_upload_and_wait_for_active"
            ) as upload:
                part, uploaded = audio_describer._prepare_video_for_gemini(
                    mock.Mock(), tmp.name, trusted_prepared_video=True
                )

        self.assertIs(part, inline_part)
        self.assertIsNone(uploaded)
        inline.assert_called_once()
        upload.assert_not_called()

    def test_permission_denied_is_recognized_for_inline_fallback(self):
        error = RuntimeError(
            "403 PERMISSION_DENIED. The caller does not have permission"
        )
        self.assertTrue(audio_describer._is_permission_denied_error(error))
        self.assertFalse(
            audio_describer._is_permission_denied_error(
                RuntimeError("429 RESOURCE_EXHAUSTED")
            )
        )

    def test_structured_403_is_recognized_through_exception_chain(self):
        inner = RuntimeError("permission failure")
        inner.code = 403
        outer = RuntimeError("Gemini request failed")
        outer.__cause__ = inner
        self.assertTrue(audio_describer._is_permission_denied_error(outer))

    def test_inline_limit_covers_observed_blocking_chunk_with_headroom(self):
        self.assertGreater(audio_describer._INLINE_VIDEO_MAX_BYTES, 23 * 1024 * 1024)
        self.assertLess(audio_describer._INLINE_VIDEO_MAX_BYTES, 100 * 1024 * 1024)


if __name__ == "__main__":
    unittest.main()
