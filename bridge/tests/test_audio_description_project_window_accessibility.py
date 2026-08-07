import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
PROJECT = (
    ROOT / "src" / "app_windows" / "audio_description_project_window.rs"
).read_text(encoding="utf-8")
WINDOW = (
    ROOT / "src" / "app_windows" / "audio_description_window.rs"
).read_text(encoding="utf-8")
MAIN = (ROOT / "src" / "main.rs").read_text(encoding="utf-8")
AUDIO = (ROOT / "src" / "audio_description.rs").read_text(encoding="utf-8")


class AudioDescriptionProjectWindowAccessibilityTests(unittest.TestCase):
    def test_project_opens_and_reopens_on_inserted_descriptions_list(self):
        self.assertIn("fn focus_descriptions_list", PROJECT)
        self.assertIn("SetFocus((*pointer).list);", PROJECT)
        self.assertIn("focus_descriptions_list(existing);", PROJECT)
        self.assertIn("SetFocus(list);", PROJECT)

    def test_tab_and_shift_tab_leave_multiline_description_editor(self):
        navigation = PROJECT[
            PROJECT.index("pub fn handle_navigation"):
            PROJECT.index("fn get_text")
        ]
        self.assertIn("key == VK_TAB.0 as u32", navigation)
        self.assertIn("VK_SHIFT", navigation)
        self.assertIn("get_next_dlg_tab_item_safe", navigation)
        self.assertIn("set_focus_safe(next)", navigation)
        self.assertIn("focus_descriptions_list(hwnd)", navigation)

    def test_navigation_only_handles_messages_from_project_window(self):
        navigation = PROJECT[
            PROJECT.index("pub fn handle_navigation"):
            PROJECT.index("fn get_text")
        ]
        self.assertIn("msg.hwnd != hwnd", navigation)
        self.assertIn("!crate::is_child_safe(hwnd, msg.hwnd)", navigation)

    def test_project_is_registered_before_main_focus_timers_can_steal_focus(self):
        create = PROJECT[PROJECT.index("WM_CREATE =>"):PROJECT.index("WM_COMMAND =>")]
        register = create.index("state.audio_description_project_window = hwnd")
        focus = create.index("SetFocus(list);")
        self.assertLess(register, focus)
        self.assertIn("state.audio_description_project_window.0 == 0", MAIN)
        self.assertIn("!IsWindowVisible(state.audio_description_project_window).as_bool()", MAIN)
        self.assertIn("audio_description_window::blocks_parent_focus", MAIN)

    def test_explorer_open_uses_separate_main_window_while_creation_window_is_active(self):
        self.assertIn("COPYDATA_RESULT_OPEN_NEW_WINDOW", MAIN)
        self.assertIn("fn should_route_external_file_open_to_new_window", MAIN)
        self.assertIn(
            "audio_description_window::blocks_parent_focus(hwnd, audio_description_window)",
            MAIN,
        )
        self.assertIn("fn spawn_new_window_with_paths(paths: &[PathBuf])", MAIN)
        self.assertIn('command.arg("--new-window")', MAIN)
        self.assertIn("command.args(paths)", MAIN)
        self.assertIn("if spawn_new_window_with_paths(&paths)", MAIN)
        self.assertIn(
            "WM_COPYDATA open spawned a separate Sonarpad window",
            MAIN,
        )
        self.assertIn('let force_new_window = args.iter().any(|arg| arg == "--new-window")', MAIN)
        self.assertIn("if !force_new_window", MAIN)
        self.assertIn(
            "Existing Sonarpad instance spawned a separate window for Explorer file(s); sender will exit",
            MAIN,
        )

    def test_creation_window_survives_show_desktop_focus_round_trip(self):
        self.assertIn("pub(crate) fn restore_on_parent_activation", WINDOW)
        self.assertIn("if foreground != parent", WINDOW)
        self.assertIn("ShowWindow(window, SW_SHOW);", WINDOW)
        self.assertIn("SetForegroundWindow(window);", WINDOW)
        self.assertGreaterEqual(
            MAIN.count("audio_description_window::restore_on_parent_activation(hwnd)"),
            3,
        )
        self.assertGreaterEqual(
            MAIN.count("audio_description_window::blocks_parent_focus("),
            2,
        )

    def test_project_is_owned_by_creation_window_and_returns_focus_on_close(self):
        self.assertIn("open_from_dialog(parent: HWND, owner: HWND)", PROJECT)
        self.assertIn("hwndOwner: owner", PROJECT)
        self.assertIn("parent, owner, project_path.to_path_buf(), project", PROJECT)
        self.assertIn("SetForegroundWindow(owner);", PROJECT)
        self.assertIn("SetFocus(owner);", PROJECT)
        call = WINDOW[
            WINDOW.index("audio_description_project_window::open_from_dialog"):
            WINDOW.index("audio_description_project_window::open_from_dialog") + 220
        ]
        self.assertIn("state.parent", call)
        self.assertIn("hwnd", call)

    def test_space_enter_and_context_menu_choose_exported_or_modified_preview(self):
        navigation = PROJECT[
            PROJECT.index("pub fn handle_navigation"):
            PROJECT.index("fn get_text")
        ]
        self.assertIn("VK_SPACE", navigation)
        self.assertIn("VK_RETURN", navigation)
        preview = PROJECT[
            PROJECT.index("fn stop_preview"):
            PROJECT.index("fn delete_selected_description")
        ]
        self.assertIn("fn play_exported_description", preview)
        self.assertIn("state.project.output_mp3_path", preview)
        self.assertIn("description.output_start_sec", preview)
        self.assertIn(".output_end_sec", preview)
        self.assertIn("if draft != description.text || description.rendered_text != description.text", preview)
        self.assertIn("start_draft_preview(hwnd, state, index, draft)", preview)
        self.assertIn("synthesize_audio_description_project_preview", PROJECT)
        self.assertIn("WM_PROJECT_DRAFT_PREVIEW_DONE", PROJECT)
        self.assertIn("audio_description.project.play_description", PROJECT)
        self.assertIn("ID_CONTEXT_PLAY", PROJECT)

    def test_modified_preview_overlays_ducked_original_audio_without_saving(self):
        preview = PROJECT[
            PROJECT.index("fn play_modified_preview"):
            PROJECT.index("fn play_selected_description")
        ]
        self.assertIn("BassOutput::start(preview_audio.path()", preview)
        self.assertIn("BassOutput::start_with_ffmpeg_at", preview)
        self.assertIn("source_start", preview)
        self.assertNotIn("source.seek_to_seconds(source_start)", preview)
        self.assertIn("state.project.source_path", preview)
        self.assertIn("state.project.ducking_db", preview)
        self.assertIn("description.source_start_sec", preview)
        self.assertIn("description.extended_pause", preview)
        self.assertNotIn("save_audio_description_project", preview)
        synthesis = AUDIO[
            AUDIO.index("pub fn synthesize_audio_description_project_preview"):
            AUDIO.index("pub fn apply_audio_description_project_edit")
        ]
        self.assertIn("synthesize_description(", synthesis)
        self.assertIn("modified_description_preview.wav", synthesis)
        self.assertNotIn("save_audio_description_project", synthesis)



    def test_modified_preview_uses_precise_ffmpeg_start_without_bass_reseek(self):
        bass_output = (ROOT / "src" / "bass_output.rs").read_text(encoding="utf-8")
        bass_stream = (ROOT / "src" / "bass_ffmpeg_stream.rs").read_text(encoding="utf-8")
        ffmpeg_source = (ROOT / "src" / "ffmpeg_source.rs").read_text(encoding="utf-8")
        self.assertIn("pub fn start_with_ffmpeg_at", bass_output)
        self.assertIn("start_seconds: f64", bass_stream)
        self.assertIn("FfmpegSource::try_new_at", bass_stream)
        self.assertIn("pub fn try_new_at", ffmpeg_source)
        self.assertIn("Duration::from_secs_f64(start_seconds)", ffmpeg_source)

    def test_applied_but_not_exported_text_still_uses_synthesized_preview(self):
        self.assertIn("pub rendered_text: String", AUDIO)
        loader = AUDIO[
            AUDIO.index("pub fn load_audio_description_project"):
            AUDIO.index("pub fn save_audio_description_project")
        ]
        self.assertIn("description.rendered_text.is_empty()", loader)
        apply_fn = AUDIO[
            AUDIO.index("pub fn apply_audio_description_project_edit"):
            AUDIO.index("pub fn delete_audio_description_project_description")
        ]
        self.assertIn("description.text = normalized_text.to_string()", apply_fn)
        self.assertNotIn("description.rendered_text =", apply_fn)
        builder = AUDIO[
            AUDIO.index("fn build_audio_description_project"):
            AUDIO.index("pub fn load_audio_description_project")
        ]
        self.assertIn("rendered_text: description.text.clone()", builder)

    def test_context_menu_deletes_and_saves_selected_description(self):
        self.assertIn("WM_CONTEXTMENU", PROJECT)
        self.assertIn("ID_CONTEXT_DELETE", PROJECT)
        self.assertIn("delete_audio_description_project_description", PROJECT)
        delete = AUDIO[
            AUDIO.index("pub fn delete_audio_description_project_description"):
            AUDIO.index("fn validate_job")
        ]
        self.assertIn("updated.descriptions.remove(index)", delete)
        self.assertIn("save_audio_description_project(project_path, &updated)?", delete)

    def test_apply_synthesizes_checks_duration_then_saves_immediately(self):
        apply_fn = AUDIO[
            AUDIO.index("pub fn apply_audio_description_project_edit"):
            AUDIO.index("pub fn delete_audio_description_project_description")
        ]
        synthesis = apply_fn.index("synthesize_description(")
        duration_check = apply_fn.index("validate_audio_description_project_edit_duration(")
        save = apply_fn.index("save_audio_description_project(project_path, &updated)")
        self.assertLess(synthesis, duration_check)
        self.assertLess(duration_check, save)
        self.assertIn("AudioDescriptionProjectEditError::TooLong", AUDIO)
        self.assertIn("start_apply(hwnd, state)", PROJECT)
        self.assertIn("WM_PROJECT_APPLY_DONE", PROJECT)

    def test_apply_error_returns_to_project_description_list(self):
        apply_done = PROJECT[
            PROJECT.index("WM_PROJECT_APPLY_DONE =>"):
            PROJECT.index("WM_PROJECT_PROGRESS =>")
        ]
        too_long = apply_done[
            apply_done.index("AudioDescriptionProjectEditError::TooLong"):
            apply_done.index("AudioDescriptionProjectEditError::Other")
        ]
        self.assertIn("show_project_error(hwnd, state.language, &message)", too_long)
        self.assertNotIn("show_error(hwnd", too_long)
        helper = PROJECT[
            PROJECT.index("fn show_project_error"):
            PROJECT.index("pub fn handle_navigation")
        ]
        self.assertIn("MessageBoxW(", helper)
        self.assertIn("crate::watchdog::enter_modal_dialog()", helper)
        self.assertIn("crate::watchdog::exit_modal_dialog()", helper)
        self.assertIn("crate::is_window_handle_valid(hwnd)", helper)
        self.assertIn("SetForegroundWindow(hwnd)", helper)
        self.assertIn("focus_descriptions_list(hwnd)", helper)
        self.assertNotIn("show_error(", helper)
        self.assertNotIn("show_blocking_modal_message_box", helper)

    def test_export_result_messages_return_to_project_description_list(self):
        done = PROJECT[
            PROJECT.index("WM_PROJECT_DONE =>"):
            PROJECT.index("WM_SETFOCUS =>")
        ]
        self.assertIn(
            "show_project_info(hwnd, state.language, &labels(state.language).success)",
            done,
        )
        self.assertNotIn("show_info(state.parent", done)
        self.assertIn("show_project_error(hwnd, state.language, &error)", done)
        helper = PROJECT[
            PROJECT.index("fn show_project_info"):
            PROJECT.index("pub fn handle_navigation")
        ]
        self.assertIn("MessageBoxW(", helper)
        self.assertIn("MB_ICONINFORMATION", helper)
        self.assertIn("crate::watchdog::enter_modal_dialog()", helper)
        self.assertIn("crate::watchdog::exit_modal_dialog()", helper)
        self.assertIn("crate::is_window_handle_valid(hwnd)", helper)
        self.assertIn("SetForegroundWindow(hwnd)", helper)
        self.assertIn("focus_descriptions_list(hwnd)", helper)
        self.assertNotIn("show_info(", helper)

    def test_selection_change_does_not_silently_save_unapplied_text(self):
        command = PROJECT[PROJECT.index("WM_COMMAND =>"):PROJECT.index("WM_CONTEXTMENU =>")]
        selection = command[
            command.index("ID_LIST if notification == LBN_SELCHANGE"):
            command.index("ID_APPLY if !state.running")
        ]
        self.assertNotIn("description.text = text", selection)
        self.assertIn("set_text(state.edit, &description.text)", selection)
        self.assertIn("apply_before_export", PROJECT)

    def test_output_player_return_focuses_close_not_start(self):
        returned = WINDOW[
            WINDOW.index("WM_AD_PLAYER_RETURN =>"):
            WINDOW.index("WM_AD_SET_INPUT =>")
        ]
        self.assertIn("return_to_editor_after_player = true", returned)
        self.assertIn("SetFocus((*pointer).close_button);", returned)
        self.assertNotIn("SetFocus((*pointer).start_button);", returned)

    def test_audio_description_windows_block_player_keyboard_shortcuts(self):
        keyboard = MAIN[
            MAIN.index("// Audiobook keyboard controls (ONLY if no secondary window is open)"):
            MAIN.index("// Exclude voice panel controls from player keyboard handling")
        ]
        self.assertIn("audio_description_window::blocks_parent_focus", keyboard)
        self.assertIn("state.audio_description_project_window", keyboard)
        self.assertIn("IsWindowVisible(state.audio_description_project_window)", keyboard)

    def test_closing_after_output_preview_returns_to_normal_editor(self):
        destroyed = WINDOW[WINDOW.index("WM_DESTROY =>"):WINDOW.index("WM_NCDESTROY =>")]
        self.assertIn("return_to_editor_after_player", destroyed)
        self.assertIn("source_player_path", destroyed)
        self.assertIn("finish_audio_description_after_output_preview", destroyed)

        cleanup = MAIN[
            MAIN.index("pub(crate) fn finish_audio_description_after_output_preview"):
            MAIN.index("fn pause_active_playback_for_audio_description")
        ]
        self.assertIn("stop_managed_mpv_playback(hwnd)", cleanup)
        self.assertIn("matches!(doc.format, FileFormat::Audiobook)", cleanup)
        self.assertIn("editor_manager::close_document_at(hwnd, index)", cleanup)
        self.assertIn("clear_active_podcast_chapters(hwnd)", cleanup)
        self.assertIn("clear_active_youtube_return_context(hwnd)", cleanup)
        self.assertIn("enable_window_safe(hwnd, true)", cleanup)
        self.assertIn("WM_CANCELMODE", cleanup)
        self.assertIn("state.alt_menu_suppressed = false", cleanup)
        self.assertIn("state.alt_menu_used_with_key = false", cleanup)
        self.assertIn("DrawMenuBar(hwnd)", cleanup)
        self.assertIn("WM_FOCUS_EDITOR", cleanup)
        self.assertIn("post_message_w_safe", cleanup)
        self.assertNotIn("focus_editor(hwnd);", cleanup)

    def test_output_preview_focus_is_deferred_until_audio_description_window_is_destroyed(self):
        destroyed = WINDOW[WINDOW.index("WM_DESTROY =>"):WINDOW.index("WM_NCDESTROY =>")]
        self.assertIn("audio_description_window = HWND(0)", destroyed)
        self.assertIn("finish_audio_description_after_output_preview", destroyed)

        cleanup = MAIN[
            MAIN.index("pub(crate) fn finish_audio_description_after_output_preview"):
            MAIN.index("fn pause_active_playback_for_audio_description")
        ]
        self.assertIn("scheduling deferred editor focus after preview cleanup", cleanup)
        self.assertIn("post_message_w_safe", cleanup)
        self.assertNotIn("bring_window_to_foreground(hwnd);", cleanup)


if __name__ == "__main__":
    unittest.main()
