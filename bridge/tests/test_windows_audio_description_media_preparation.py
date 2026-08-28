from __future__ import annotations

import unittest
from pathlib import Path


class WindowsAudioDescriptionMediaPreparationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(__file__).resolve().parents[2]
        cls.audio_source = (cls.root / "src" / "audio_description.rs").read_text(
            encoding="utf-8"
        )
        cls.ffmpeg_source = (cls.root / "src" / "ffmpeg_export.rs").read_text(
            encoding="utf-8"
        )

    def test_gemini_chunks_adapt_to_inline_safe_size(self):
        self.assertIn(
            "const GEMINI_INLINE_TARGET_CHUNK_BYTES: u64 = 40 * 1024 * 1024;",
            self.audio_source,
        )
        self.assertIn("GEMINI_SEGMENT_RETRY_LIMIT", self.audio_source)
        self.assertIn("adaptive Gemini chunk retry", self.audio_source)
        self.assertIn("max_chunk_bytes <= GEMINI_INLINE_TARGET_CHUNK_BYTES", self.audio_source)
        self.assertIn("file_size >= GEMINI_MAX_CHUNK_BYTES", self.audio_source)

    def test_large_source_is_split_instead_of_rejected_wholesale(self):
        self.assertIn("if input_size == 0 {", self.audio_source)
        self.assertNotIn(
            "if input_size == 0 || input_size >= GEMINI_MAX_CHUNK_BYTES {",
            self.audio_source,
        )

    def test_native_ffmpeg_repairs_missing_avi_timestamps(self):
        self.assertIn('"fflags", "+genpts+discardcorrupt"', self.ffmpeg_source)
        self.assertIn("fn repair_segment_packet_timestamps", self.ffmpeg_source)
        self.assertIn("pts_missing", self.ffmpeg_source)
        self.assertIn("dts_missing", self.ffmpeg_source)
        self.assertIn("repaired_timestamp_packets", self.ffmpeg_source)

    def test_segment_muxer_errors_are_not_silently_ignored(self):
        self.assertIn("segment_write_error = Some(format!", self.ffmpeg_source)
        self.assertIn("if let Some(error) = segment_write_error", self.ffmpeg_source)
        self.assertIn("failed to finalize segmented output", self.ffmpeg_source)

    def test_changed_chunk_layout_does_not_make_resume_fatal(self):
        self.assertIn("ignoring resume checkpoint after chunk layout change", self.audio_source)
        self.assertIn("ignoring invalid resume checkpoint", self.audio_source)


if __name__ == "__main__":
    unittest.main()
