import sys
import re
import os

def main():
    files_to_check = [
        "src/main.rs", "src/search.rs", "src/tts_engine.rs", "src/editor_manager.rs",
        "src/audio_player.rs", "src/mf_encoder.rs", "src/sapi5_engine.rs",
        "src/app_windows/youtube_transcript_window.rs", "src/app_windows/prompt_window.rs",
        "src/app_windows/podcast_window.rs", "src/app_windows/podcast_save_window.rs",
        "src/app_windows/podcasts_window.rs", "src/app_windows/options_window.rs",
        "src/app_windows/marker_select_window.rs", "src/app_windows/help_window.rs",
        "src/app_windows/find_in_files_window.rs", "src/app_windows/dictionary_window.rs",
        "src/app_windows/bookmarks_window.rs", "src/app_windows/batch_audiobooks_window.rs",
        "src/app_windows/audiobook_window.rs", "src/app_windows/rss_window.rs",
        "src/app_windows/wiktionary_window.rs"
    ]

    for file_path in files_to_check:
        if not os.path.exists(file_path):
            continue
            
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            
        changed = False
        for i in range(len(lines)):
            line = lines[i]
            
            # 1. Fix unused variable e -> _e
            if "if let Err(e) =" in line:
                lines[i] = line.replace("if let Err(e) =", "if let Err(_e) =")
                changed = True
            
            # 2. Fix unreachable match arms by adding missing 'if'
            # Look for with_ calls that start a line and end with ).is_none() {
            stripped = line.lstrip()
            if stripped.startswith("with_") and ").is_none() {" in line:
                indent = line[:line.find("with_")]
                lines[i] = indent + "if " + stripped
                changed = True

        if changed:
            with open(file_path, "w", encoding="utf-8") as f:
                f.writelines(lines)

if __name__ == "__main__":
    main()
