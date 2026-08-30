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
        cls.la7 = (cls.root / "src" / "app_windows" / "la7_play_window.rs").read_text(
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
        self.assertRegex(
            self.selector,
            r"append_context_actions_to_menu\(\s*menu,",
        )
        self.assertRegex(
            self.youtube,
            r"append_context_actions_to_menu\(\s*menu,",
        )

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
        self.assertIn("crate::start_whisper_transcription_for_raiplay_context", self.raiplay)
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
        self.assertIn("crate::start_whisper_transcription_for_remote_media_context", block)


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

        self.assertRegex(
            self.main,
            r"youtube_transcript_window::restore_active_media_context_list\(\s*hwnd\s*\)",
        )
        self.assertIn(
            "youtube_transcript_window::has_active_media_context_list(hwnd)",
            self.main,
        )
        self.assertRegex(
            self.rai_ad,
            r"youtube_transcript_window::restore_active_media_context_list\(\s*parent\s*,?\s*\);",
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


    def test_successful_transcription_dismisses_originating_media_lists_before_editor(self):
        self.assertIn(
            "pub fn close_active_interpreter_selects_for_parent",
            self.selector,
        )
        self.assertIn(
            "pub(crate) fn dismiss_active_media_context_lists_for_editor",
            self.youtube,
        )

        start = self.main.index("fn apply_whisper_transcription_result(")
        end = self.main.index("fn auto_save_whisper_transcription(", start)
        block = self.main[start:end]
        self.assertIn("let transcription_succeeded = !result.cancelled", block)
        self.assertIn("let dismissed_media_context = transcription_succeeded", block)
        self.assertIn(
            "youtube_transcript_window::dismiss_active_media_context_lists_for_editor(",
            block,
        )
        self.assertLess(
            block.index("dismiss_active_media_context_lists_for_editor("),
            block.index("close_whisper_progress_window(hwnd);"),
        )
        self.assertIn("if dismissed_media_context", block)
        self.assertLess(
            block.index("crate::enable_window_safe(hwnd, true);"),
            block.index("editor_manager::new_document(hwnd);"),
        )

        dismiss_start = self.youtube.index(
            "pub(crate) fn dismiss_active_media_context_lists_for_editor"
        )
        dismiss_end = self.youtube.index("const EVENT_OBJECT_FOCUS", dismiss_start)
        dismiss_block = self.youtube[dismiss_start:dismiss_end]
        self.assertIn("close_active_interpreter_selects_for_parent", dismiss_block)
        self.assertIn("destroy_window_safe(dialog)", dismiss_block)
        self.assertIn("enable_window_safe(parent, true)", dismiss_block)

    def test_context_transcription_is_deferred_until_selector_closes(self):
        self.assertIn(
            "request_close_active_interpreter_selects_for_parent",
            self.selector,
        )
        self.assertIn("WM_START_CONTEXT_WHISPER_TRANSCRIPTION", self.main)
        self.assertIn("start_whisper_transcription_for_media_context", self.main)
        self.assertIn("start_whisper_transcription_for_remote_media_context", self.main)

        youtube_context_start = self.youtube.index("fn youtube_video_context_actions(")
        youtube_context_end = self.youtube.index(
            "fn choose_stream_collection_entry_page(", youtube_context_start
        )
        youtube_context = self.youtube[youtube_context_start:youtube_context_end]
        self.assertIn("start_whisper_transcription_for_media_context", youtube_context)
        self.assertNotIn(
            "start_whisper_transcription_for_media(parent, path, media_title)",
            youtube_context,
        )

        raiplay_start = self.raiplay.index("let transcribe_context_action")
        raiplay_end = self.raiplay.index("let ad_items_for_enabled", raiplay_start)
        self.assertIn(
            "start_whisper_transcription_for_raiplay_context",
            self.raiplay[raiplay_start:raiplay_end],
        )

        rai_start = self.rai_ad.index("let transcribe_action")
        rai_end = self.rai_ad.index("let copy_items", rai_start)
        self.assertIn(
            "start_whisper_transcription_for_remote_media_context",
            self.rai_ad[rai_start:rai_end],
        )

    def test_youtube_context_transcription_unwinds_without_empty_editor_focus(self):
        self.assertIn("YOUTUBE_CONTEXT_TRANSCRIPTION_UNWIND_PARENT", self.youtube)
        self.assertIn("YOUTUBE_CONTEXT_TRANSCRIPTION_ACTIVE_PARENT", self.youtube)
        self.assertIn("pub(crate) fn mark_context_transcription_started", self.youtube)
        self.assertIn(
            "pub(crate) fn context_transcription_keeps_progress_foreground",
            self.youtube,
        )
        self.assertIn("pub(crate) fn finish_context_transcription", self.youtube)
        self.assertIn("fn take_context_transcription_unwind", self.youtube)

        youtube_context_start = self.youtube.index("fn youtube_video_context_actions(")
        youtube_context_end = self.youtube.index(
            "fn choose_stream_collection_entry_page(", youtube_context_start
        )
        youtube_context = self.youtube[youtube_context_start:youtube_context_end]
        self.assertIn("mark_context_transcription_started(parent);", youtube_context)
        self.assertLess(
            youtube_context.index("mark_context_transcription_started(parent);"),
            youtube_context.index("start_whisper_transcription_for_media_context"),
        )

        resolved_start = self.youtube.index("let resolved = match resolve_stream_input_url(")
        resolved_end = self.youtube.index("url = resolved.url;", resolved_start)
        resolved_block = self.youtube[resolved_start:resolved_end]
        self.assertIn("if take_context_transcription_unwind(parent)", resolved_block)
        self.assertLess(
            resolved_block.index("take_context_transcription_unwind(parent)"),
            resolved_block.index("dialog_data.previous_input.clone()"),
        )
        self.assertLess(
            resolved_block.index("take_context_transcription_unwind(parent)"),
            resolved_block.index("post_focus_editor(parent);"),
        )

        self.assertIn("fn restore_context_transcription_progress_focus", self.main)
        self.assertIn(
            "context_transcription_keeps_progress_foreground",
            self.main,
        )
        self.assertIn(
            "focus timer {} suppressed because YouTube context transcription is active",
            self.main,
        )
        apply_start = self.main.index("fn apply_whisper_transcription_result(")
        apply_end = self.main.index("fn auto_save_whisper_transcription(", apply_start)
        apply_block = self.main[apply_start:apply_end]
        self.assertIn(
            "youtube_transcript_window::finish_context_transcription(hwnd)",
            apply_block,
        )

    def test_youtube_context_audio_description_closes_selector_before_opening_window(self):
        self.assertIn("YOUTUBE_CONTEXT_AUDIO_DESCRIPTION_UNWIND_PARENT", self.youtube)
        self.assertIn("YOUTUBE_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT", self.youtube)
        self.assertIn("mark_context_audio_description_started", self.youtube)
        self.assertIn("take_context_audio_description_unwind", self.youtube)
        self.assertIn("download_active_youtube_for_audio_description_context", self.youtube)

        youtube_context_start = self.youtube.index("fn youtube_video_context_actions(")
        youtube_context_end = self.youtube.index(
            "fn choose_stream_collection_entry_page(", youtube_context_start
        )
        youtube_context = self.youtube[youtube_context_start:youtube_context_end]
        self.assertIn(
            "download_active_youtube_for_audio_description_context", youtube_context
        )

        download_start = self.youtube.index(
            "fn download_stream_context_for_audio_description("
        )
        download_end = self.youtube.index(
            "pub(crate) fn download_active_youtube_for_audio_description(", download_start
        )
        download_block = self.youtube[download_start:download_end]
        self.assertIn("youtube_context_menu: bool", download_block)
        self.assertGreaterEqual(
            download_block.count("suppress_parent_restore_on_close(progress)"), 2
        )
        self.assertIn("crate::queue_youtube_context_audio_description", download_block)
        self.assertIn(
            "audio_description_window::open_with_input(parent, target_path)", download_block
        )

        resolved_start = self.youtube.index("let resolved = match resolve_stream_input_url(")
        resolved_end = self.youtube.index("url = resolved.url;", resolved_start)
        resolved_block = self.youtube[resolved_start:resolved_end]
        self.assertIn("take_context_audio_description_unwind(parent)", resolved_block)
        self.assertLess(
            resolved_block.index("take_context_audio_description_unwind(parent)"),
            resolved_block.index("post_focus_editor(parent);"),
        )

        self.assertIn("WM_START_YOUTUBE_CONTEXT_AUDIO_DESCRIPTION", self.main)
        self.assertIn("queue_youtube_context_audio_description", self.main)
        self.assertIn("restore_youtube_context_audio_description_focus", self.main)
        self.assertIn(
            "focus timer {} suppressed because YouTube context audio description is opening",
            self.main,
        )

        shortcut_start = self.main.index("fn open_audio_description_from_current_context(")
        shortcut_end = self.main.index(
            "fn restore_context_transcription_progress_focus", shortcut_start
        )
        shortcut_block = self.main[shortcut_start:shortcut_end]
        self.assertIn("download_active_youtube_for_audio_description(", shortcut_block)
        self.assertNotIn(
            "download_active_youtube_for_audio_description_context", shortcut_block
        )

    def test_raiplay_context_audio_description_closes_browser_and_protects_focus(self):
        self.assertIn("RAIPLAY_CONTEXT_AUDIO_DESCRIPTION_PENDING_PARENT", self.raiplay)
        self.assertIn("RAIPLAY_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT", self.raiplay)
        self.assertIn("RAIPLAY_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT", self.raiplay)
        self.assertIn("mark_context_audio_description_started", self.raiplay)
        self.assertIn("take_context_audio_description_browser_exit", self.raiplay)
        cancelled_start = self.raiplay.index("MultilineSelectionResult::Cancelled =>")
        cancelled_end = self.raiplay.index("if let Some((previous_page_path", cancelled_start)
        cancelled_block = self.raiplay[cancelled_start:cancelled_end]
        self.assertIn("take_context_audio_description_browser_exit(parent)", cancelled_block)

        download_start = self.main.index("fn download_podcast_episode_with_progress(")
        download_end = self.main.index("fn post_podcast_episode_save_result", download_start)
        download_block = self.main[download_start:download_end]
        self.assertIn("raiplay_context_audio_description", download_block)
        self.assertIn("mark_context_audio_description_started(hwnd)", download_block)
        self.assertIn("suppress_parent_restore_on_close(progress_dialog)", download_block)

        self.assertIn("WM_START_RAIPLAY_CONTEXT_AUDIO_DESCRIPTION", self.main)
        self.assertIn("queue_raiplay_context_audio_description", self.main)
        self.assertIn("restore_raiplay_context_audio_description_focus", self.main)
        self.assertIn(
            "focus timer {} suppressed because RaiPlay context audio description is opening",
            self.main,
        )
        save_result_start = self.main.index("WM_PODCAST_EPISODE_SAVE_RESULT =>")
        save_result_end = self.main.index("WM_EDITOR_TRANSLATION_DONE =>", save_result_start)
        save_result_block = self.main[save_result_start:save_result_end]
        self.assertIn("finish_context_audio_description_download", save_result_block)
        self.assertIn("dismiss_active_media_context_lists_for_editor", save_result_block)
        self.assertIn("queue_raiplay_context_audio_description", save_result_block)

    def test_la7_context_actions_are_available_only_for_resolvable_vod_media(self):
        self.assertIn("type La7ContextTargetCache", self.la7)
        self.assertIn("fn cached_context_vod_url", self.la7)
        self.assertIn("item.kind != ItemKind::Media", self.la7)
        self.assertIn("la7_play::resolve_vod(&item.target)", self.la7)
        self.assertIn('"playback.download_episode"', self.la7)
        self.assertIn('"playback.transcribe_current"', self.la7)
        self.assertIn('"menu.create_audio_description"', self.la7)
        self.assertIn("crate::save_la7_context_media", self.la7)
        self.assertIn("crate::start_whisper_transcription_for_la7_context", self.la7)
        self.assertIn("crate::create_audio_description_from_la7_context", self.la7)

    def test_la7_context_audio_description_closes_browser_and_protects_focus(self):
        self.assertIn("LA7_CONTEXT_AUDIO_DESCRIPTION_PENDING_PARENT", self.la7)
        self.assertIn("LA7_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT", self.la7)
        self.assertIn("LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT", self.la7)
        self.assertIn("take_context_audio_description_browser_exit", self.la7)
        cancel_start = self.la7.index("MultilineSelectionResult::Cancelled =>")
        cancel_end = self.la7.index("if let Some((src, id))", cancel_start)
        cancel_block = self.la7[cancel_start:cancel_end]
        self.assertIn("take_context_audio_description_browser_exit(parent)", cancel_block)

        self.assertIn("la7_context_audio_description: bool", self.main)
        self.assertIn("mark_context_audio_description_started(hwnd)", self.main)
        self.assertIn("WM_START_LA7_CONTEXT_AUDIO_DESCRIPTION", self.main)
        self.assertIn("queue_la7_context_audio_description", self.main)
        self.assertIn("restore_la7_context_audio_description_focus", self.main)
        self.assertIn(
            "focus timer {} suppressed because La7 Play context audio description is opening",
            self.main,
        )

    def test_la7_context_transcription_closes_browser_and_keeps_whisper_in_front(self):
        self.assertIn("LA7_CONTEXT_TRANSCRIPTION_PENDING_PARENT", self.la7)
        self.assertIn("LA7_CONTEXT_TRANSCRIPTION_EXIT_PARENT", self.la7)
        self.assertIn("LA7_CONTEXT_TRANSCRIPTION_FOCUS_PARENT", self.la7)
        self.assertIn("context_transcription_keeps_progress_foreground", self.la7)
        self.assertIn("take_context_transcription_browser_exit", self.la7)
        self.assertIn("close_la7_on_success: bool", self.main)
        self.assertIn("start_whisper_transcription_for_la7_context", self.main)
        self.assertIn(
            "app_windows::la7_play_window::mark_context_transcription_started(hwnd)",
            self.main,
        )
        self.assertIn(
            "app_windows::la7_play_window::finish_context_transcription(hwnd, transcription_succeeded)",
            self.main,
        )

    def test_la7_context_save_and_ad_reuse_existing_la7_exporter(self):
        start = self.main.index("fn run_la7_context_media_action(")
        end = self.main.index("fn download_podcast_episode_with_progress(", start)
        block = self.main[start:end]
        self.assertIn("download_podcast_episode_with_progress(PodcastProgressDownloadRequest", block)
        self.assertIn("rai_origin: RaiAudioOrigin::La7Play", block)
        self.assertIn("la7_context_audio_description: open_audio_description", block)

    def test_transcription_documents_confirm_before_clean_close(self):
        editor = (self.root / "src" / "editor_manager.rs").read_text(encoding="utf-8")
        self.assertIn("pub confirm_close_when_clean: bool", editor)
        self.assertIn('crate::i18n::tr(language, "whisper.close_confirm")', editor)
        self.assertIn("doc.confirm_close_when_clean = true;", self.main)
        for path in sorted((self.root / "i18n").glob("*.json")):
            values = __import__("json").loads(path.read_text(encoding="utf-8"))
            self.assertIn("whisper.close_confirm", values, path.name)

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

    def test_successful_raiplay_context_transcription_exits_browser_history(self):
        self.assertIn("RAIPLAY_CONTEXT_TRANSCRIPTION_PENDING_PARENT", self.raiplay)
        self.assertIn("RAIPLAY_CONTEXT_TRANSCRIPTION_EXIT_PARENT", self.raiplay)
        self.assertIn("pub(crate) fn mark_context_transcription_started", self.raiplay)
        self.assertIn("pub(crate) fn finish_context_transcription", self.raiplay)
        self.assertIn("fn take_context_transcription_browser_exit", self.raiplay)

        cancel_start = self.raiplay.index("MultilineSelectionResult::Cancelled =>")
        cancel_end = self.raiplay.index("let Some(selected_item)", cancel_start)
        cancel_block = self.raiplay[cancel_start:cancel_end]
        self.assertIn("take_context_transcription_browser_exit(parent)", cancel_block)
        self.assertLess(
            cancel_block.index("take_context_transcription_browser_exit(parent)"),
            cancel_block.index("history.pop()"),
        )

        self.assertIn("close_raiplay_on_success: bool", self.main)
        self.assertIn("start_whisper_transcription_for_raiplay_context", self.main)
        self.assertIn(
            "app_windows::raiplay_window::mark_context_transcription_started(hwnd)",
            self.main,
        )

        apply_start = self.main.index("fn apply_whisper_transcription_result(")
        apply_end = self.main.index("fn auto_save_whisper_transcription(", apply_start)
        apply_block = self.main[apply_start:apply_end]
        self.assertIn(
            "app_windows::raiplay_window::finish_context_transcription(hwnd, transcription_succeeded)",
            apply_block,
        )
        self.assertLess(
            apply_block.index("finish_context_transcription(hwnd, transcription_succeeded)"),
            apply_block.index("dismiss_active_media_context_lists_for_editor("),
        )


if __name__ == "__main__":
    unittest.main()
