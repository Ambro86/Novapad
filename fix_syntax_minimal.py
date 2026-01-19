import os
import re

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
            
            if "if let Err(e) = " in line:
                stripped = line.strip()
                if stripped.endswith(");"):
                    indent = line[:line.find("if let Err")]
                    call = stripped[:-1]
                    lines[i] = indent + call + " { crate::log_debug(&format!(\"Error: {:?}\", e)); }" + chr(10)
                    changed = True

            if "if MessageBoxW(" in line:
                stripped = line.strip()
                if stripped.endswith(");"):
                    indent = line[:line.find("if MessageBoxW")]
                    call = stripped[3:-1]
                    lines[i] = indent + call + ";" + chr(10)
                    changed = True

        if changed:
            with open(file_path, "w", encoding="utf-8") as f:
                f.writelines(lines)

if __name__ == "__main__":
    main()
