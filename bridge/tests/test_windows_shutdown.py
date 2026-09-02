from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
MAIN = (ROOT / "src" / "main.rs").read_text(encoding="utf-8")
EDITOR = (ROOT / "src" / "editor_manager.rs").read_text(encoding="utf-8")


class WindowsShutdownTests(unittest.TestCase):
    def test_other_main_window_count_is_limited_to_current_process(self):
        self.assertIn("GetCurrentProcessId", MAIN)
        self.assertIn("GetWindowThreadProcessId(hwnd, Some(&mut window_pid))", MAIN)
        self.assertIn("window_pid != GetCurrentProcessId()", MAIN)

    def test_last_same_process_window_posts_quit(self):
        self.assertIn("let has_other = has_other_main_windows(hwnd);", MAIN)
        self.assertIn("if !has_other", MAIN)
        self.assertIn("PostQuitMessage(0);", MAIN)

    def test_shutdown_has_stage_logging(self):
        self.assertIn("Application shutdown: try_close_app start", EDITOR)
        self.assertIn("Application shutdown: destroying main window", EDITOR)
        self.assertIn("Application shutdown: main-window destroy returned", EDITOR)
        self.assertIn("Application shutdown: message loop exited", MAIN)


if __name__ == "__main__":
    unittest.main()
