from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
WATCHDOG = ROOT / "src" / "watchdog.rs"
PROGRESS = ROOT / "src" / "app_windows" / "podcast_save_window.rs"
YOUTUBE = ROOT / "src" / "app_windows" / "youtube_transcript_window.rs"


class WatchdogProgressHeartbeatTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.watchdog = WATCHDOG.read_text(encoding="utf-8")
        cls.progress = PROGRESS.read_text(encoding="utf-8")
        cls.youtube = YOUTUBE.read_text(encoding="utf-8")

    def test_global_heartbeat_hook_is_available(self):
        self.assertIn("pub fn heartbeat()", self.watchdog)
        body = self.watchdog.split("pub fn heartbeat()", 1)[1].split(
            "/// Stato condiviso del watchdog", 1
        )[0]
        self.assertIn("GLOBAL_WATCHDOG.get()", body)
        self.assertIn("state.heartbeat();", body)

    def test_progress_window_proves_ui_responsiveness_to_watchdog(self):
        progress_body = self.progress.split("WM_PODCAST_SAVE_PROGRESS =>", 1)[1].split(
            "WM_TIMER =>", 1
        )[0]
        timer_body = self.progress.split("WM_TIMER =>", 1)[1].split(
            "WM_PODCAST_SAVE_DONE =>", 1
        )[0]
        self.assertIn("crate::watchdog::heartbeat();", progress_body)
        self.assertIn("SAVE_PROGRESS_TIMER_ID", timer_body)
        self.assertIn("crate::watchdog::heartbeat();", timer_body)

    def test_nested_youtube_message_pump_updates_watchdog(self):
        body = self.youtube.split("fn pump_messages_detect_stream_cancel", 1)[1].split(
            "fn lower_current_stream_worker_priority", 1
        )[0]
        self.assertIn("crate::watchdog::heartbeat();", body)
        self.assertIn("PeekMessageW", body)
        self.assertIn("DispatchMessageW", body)


if __name__ == "__main__":
    unittest.main()
