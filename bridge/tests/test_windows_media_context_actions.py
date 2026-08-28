from __future__ import annotations

import re
import unittest
from pathlib import Path


class WindowsMediaContextActionsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(__file__).resolve().parents[2]
        cls.selector = (cls.root / "src" / "app_windows" / "interpreter_select_window.rs").read_text(
            encoding="utf-8"
        )
        cls.youtube = (cls.root / "src" / "app_windows" / "youtube_transcript_window.rs").read_text(
            encoding="utf-8"
        )
        cls.raiplay = (cls.root / "src" / "app_windows" / "raiplay_window.rs").read_text(
            encoding="utf-8"
        )
        cls.rai_ad = (cls.root / "src" / "app_windows" / "rai_audiodescrizioni_window.rs").read_text(
            encoding="utf-8"
        )
        cls.main = (cls.root / "src" / "main.rs").read_text(encoding="utf-8")

    def test_context_actions_support_nested_submenus(self):
        self.assertIn("pub children: Vec<InterpreterContextAction>", self.selector)
        self.assertIn("fn append_context_actions_to_menu", self.selector)
        self.assertIn("MF_POPUP", self.selector)
        self.assertIn("append_recursive", self.selector)
        self.assertIn("append_context_actions_to_menu(\n                    menu,", self.selector)
        self.assertIn("append_context_actions_to_menu(\n            menu,", self.youtube)

    def test_youtube_video_context_reuses_existing_media_actions(self):
        block = self.youtube[
            self.youtube.index("fn youtube_video_context_actions(") : self.youtube.index(
                "fn choose_stream_collection_entry_page(",
                self.youtube.index("fn youtube_video_context_actions("),
            )
        ]
        self.assertIn('"playback.download_episode"', block)
        self.assertIn('"playback.transcribe_current"', block)
        self.assertIn('"menu.create_audio_description"', block)
        for format_name in ["Mp4", "Mp3", "M4a", "Opus", "Ogg", "Wav", "Flac"]:
            self.assertIn(f"StreamOutputFormat::{format_name}", block)
        self.assertIn("download_active_streaming_audio_media(parent", block)
        self.assertIn("download_active_streaming_audio_media_for_transcription", block)
        self.assertIn("download_active_youtube_for_audio_description", block)
        self.assertIn("!is_youtube_collection_url(&entry.url)", block)
        self.assertIn("extract_video_id(&entry.url).is_some()", block)

    def test_youtube_selection_carries_download_options_into_context_menu(self):
        self.assertIn("download_options: PlaylistDownloadOptions", self.youtube)
        self.assertGreaterEqual(
            self.youtube.count("download_options: playlist_download_options"), 2
        )

    def test_raiplay_context_actions_hide_media_operations_for_live_streams(self):
        self.assertIn("fn context_target_is_vod", self.raiplay)
        self.assertIn("PlaybackTarget::DirectStream { is_live, .. } => !*is_live", self.raiplay)
        self.assertIn('"playback.download_episode"', self.raiplay)
        self.assertIn('"playback.transcribe_current"', self.raiplay)
        self.assertIn('"menu.create_audio_description"', self.raiplay)
        self.assertIn("crate::save_raiplay_context_media", self.raiplay)
        self.assertIn("crate::start_whisper_transcription_for_remote_media", self.raiplay)
        self.assertIn("crate::create_audio_description_from_raiplay_context", self.raiplay)

    def test_raiplay_save_and_audio_description_reuse_existing_exporter(self):
        start = self.main.index("fn run_raiplay_context_media_action(")
        end = self.main.index("fn download_podcast_episode_with_progress(", start)
        block = self.main[start:end]
        self.assertIn("download_podcast_episode_with_progress(PodcastProgressDownloadRequest", block)
        self.assertIn("rai_origin: RaiAudioOrigin::RaiPlay", block)
        self.assertIn("open_audio_description", block)

    def test_rai_audio_descriptions_offer_save_and_transcribe_but_not_create_ad(self):
        start = self.rai_ad.index("fn catalog_context_actions(")
        end = self.rai_ad.index("fn open_recent_catalog(", start)
        block = self.rai_ad[start:end]
        self.assertIn('"playback.download_episode"', block)
        self.assertIn('"playback.transcribe_current"', block)
        self.assertNotIn('"menu.create_audio_description"', block)
        self.assertIn("crate::download_podcast_episode(", block)
        self.assertIn("crate::start_whisper_transcription_for_remote_media", block)


    def test_media_save_returns_focus_to_originating_context_list(self):
        self.assertIn("ACTIVE_INTERPRETER_SELECT_WINDOWS", self.selector)
        self.assertIn("register_active_interpreter_select(hwnd, parent);", self.selector)
        self.assertIn("unregister_active_interpreter_select(hwnd);", self.selector)
        self.assertIn("pub fn restore_active_interpreter_select_for_parent", self.selector)

        self.assertIn("ACTIVE_MULTILINE_SELECT_WINDOWS", self.youtube)
        self.assertIn("register_active_multiline_select(hwnd, parent);", self.youtube)
        self.assertIn("unregister_active_multiline_select(hwnd);", self.youtube)
        self.assertIn("pub(crate) fn restore_active_media_context_list", self.youtube)

        youtube_context_start = self.youtube.index("fn youtube_video_context_actions(")
        youtube_context_end = self.youtube.index(
            "fn choose_stream_collection_entry_page(", youtube_context_start
        )
        youtube_context = self.youtube[youtube_context_start:youtube_context_end]
        self.assertIn("restore_active_media_context_list(parent);", youtube_context)

        self.assertIn(
            "youtube_transcript_window::restore_active_media_context_list(hwnd)",
            self.main,
        )
        self.assertIn(
            "youtube_transcript_window::has_active_media_context_list(hwnd)",
            self.main,
        )
        self.assertIn(
            "youtube_transcript_window::restore_active_media_context_list(parent);",
            self.rai_ad,
        )

    def test_media_save_progress_does_not_bounce_focus_to_editor_when_list_is_open(self):
        start = self.youtube.index("pub(crate) fn download_active_streaming_audio_media(")
        end = self.youtube.index(
            "pub(crate) fn download_active_streaming_video_for_audio_description(", start
        )
        block = self.youtube[start:end]
        self.assertGreaterEqual(block.count("has_active_media_context_list(parent)"), 2)
        self.assertGreaterEqual(
            block.count("suppress_parent_restore_on_close(progress)"), 2
        )

        progress_start = self.main.index("fn download_podcast_episode_with_progress(")
        progress_end = self.main.index("fn post_podcast_episode_save_result(", progress_start)
        progress_block = self.main[progress_start:progress_end]
        self.assertIn("!open_audio_description", progress_block)
        self.assertIn("has_active_media_context_list(hwnd)", progress_block)
        self.assertIn("suppress_parent_restore_on_close(progress_dialog)", progress_block)

    def test_raiplay_context_save_keeps_progress_dialog_in_front_until_completion(self):
        start = self.raiplay.index("let save_context_action")
        end = self.raiplay.index("let transcribe_context_action", start)
        block = self.raiplay[start:end]
        self.assertIn("crate::save_raiplay_context_media", block)
        self.assertNotIn("restore_active_media_context_list(parent);", block)

        result_start = self.main.index("WM_PODCAST_EPISODE_SAVE_RESULT =>")
        result_end = self.main.index("WM_EDITOR_TRANSLATION_DONE =>", result_start)
        result_block = self.main[result_start:result_end]
        self.assertIn("close_podcast_save_progress_window(hwnd);", result_block)
        self.assertIn("restore_active_media_context_list(hwnd)", result_block)


    def test_context_lists_cannot_steal_focus_from_active_save_progress(self):
        self.assertIn("pub(crate) fn media_action_progress_is_active", self.main)

        tracked_start = self.main.index("fn tracked_progress_modal_while_main_disabled(")
        tracked_end = self.main.index("pub(crate) fn media_action_progress_is_active", tracked_start)
        tracked_block = self.main[tracked_start:tracked_end]
        self.assertIn("state.podcast_save_window", tracked_block)
        self.assertIn("state.transcription_progress_window", tracked_block)

        modal_start = self.main.index("fn reactivate_modal_child_while_main_disabled(")
        modal_end = self.main.index("pub(crate) fn recover_main_window_after_audio_description", modal_start)
        modal_block = self.main[modal_start:modal_end]
        self.assertLess(
            modal_block.index("tracked_progress_modal_while_main_disabled(hwnd)"),
            modal_block.index("GetLastActivePopup(hwnd)"),
        )

        context_start = self.youtube.index("fn show_youtube_comments_context_menu(")
        context_end = self.youtube.index('unsafe extern "system" fn youtube_comments_view_wndproc', context_start)
        context_block = self.youtube[context_start:context_end]
        self.assertIn("media_action_progress_is_active(state.parent)", context_block)
        self.assertIn("preserving active progress modal", context_block)

        proxy_start = self.youtube.index("fn sync_youtube_comments_accessibility_proxy_selection(")
        proxy_end = self.youtube.index("fn selected_youtube_comment_id", proxy_start)
        proxy_block = self.youtube[proxy_start:proxy_end]
        self.assertIn("media_action_progress_is_active(state.parent)", proxy_block)

        self.assertIn(
            "interpreter_select restore focus suppressed for active progress modal",
            self.selector,
        )

    def test_all_rust_context_action_initializers_have_children_field(self):
        rust_sources = "\n".join(
            path.read_text(encoding="utf-8", errors="ignore")
            for path in (self.root / "src").rglob("*.rs")
        )
        starts = [m.start() for m in re.finditer(r"InterpreterContextAction\s*\{", rust_sources)]
        self.assertTrue(starts)
        for start in starts:
            tail = rust_sources[start : start + 1800]
            # Every initializer reaches its first handler/children section well inside this window.
            self.assertIn("children:", tail, msg=tail[:500])


if __name__ == "__main__":
    unittest.main()
