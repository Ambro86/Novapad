from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
YOUTUBE = ROOT / "src" / "app_windows" / "youtube_transcript_window.rs"
PROGRESS = ROOT / "src" / "app_windows" / "podcast_save_window.rs"


class YoutubeDownloadFocusAccessibilityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.youtube = YOUTUBE.read_text(encoding="utf-8")
        cls.progress = PROGRESS.read_text(encoding="utf-8")

    def test_stream_progress_respects_external_foreground(self):
        body = self.youtube.split("fn keep_stream_progress_focus", 1)[1].split(
            "fn get_window_text_for_log", 1
        )[0]
        self.assertIn("GetWindowThreadProcessId", body)
        self.assertIn("foreground_pid != std::process::id()", body)
        self.assertIn("respecting external foreground", body)
        self.assertIn("focus_primary_control(dialog)", body)

    def test_cancellable_stream_progress_has_accessible_status(self):
        body = self.youtube.split("fn open_progress_dialog", 1)[1].split(
            "fn restore_stream_parent_focus", 1
        )[0]
        self.assertIn("open_with_labels_and_status_field", body)
        self.assertIn("enable_accessible_progress_context", body)
        self.assertIn("focus_primary_control", body)
        self.assertIn("activate_stream_progress_window", body)
        self.assertNotIn("pin_stream_modal_window(dialog)", body)

    def test_progress_status_reports_item_and_percent_on_reactivation(self):
        self.assertIn("accessible_progress_context: bool", self.progress)
        self.assertIn("refresh_accessible_status_field(state)", self.progress)
        self.assertIn("WM_ACTIVATE =>", self.progress)
        self.assertIn('format!("{} {}%", state.status_text, state.current_pct)', self.progress)
        self.assertIn('format!("{} - {}", state.labels.title, state.status_text)', self.progress)
        self.assertIn("respecting external foreground on progress close", self.progress)
        self.assertIn("stream_audio.playlist_download_progress", self.youtube)
        self.assertIn("report_progress_status(progress, &status);", self.youtube)
        self.assertIn("status_with_title", self.youtube)


if __name__ == "__main__":
    unittest.main()
