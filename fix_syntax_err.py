import os
import re

def fix_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    original_content = content
    lines = content.splitlines(keepends=True)
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # 1. Fix if let Err(e) = ...; -> if let Err(e) = ... { }
        # Look for the start of if let Err
        if "if let Err(e) = " in line:
            stripped = line.strip()
            # If it's single line and ends with ;
            if stripped.endswith(");"):
                indent = line[:line.find("if let Err")]
                code = stripped[:-1]
                line = indent + code + " { } " + chr(10)
            else:
                # Multi-line case: look for ); alone on a line
                for j in range(i + 1, min(i + 10, len(lines))):
                    if lines[j].strip() == ");":
                        # Wrap the whole thing
                        # But wait, it's safer to just fix the semicolon line
                        lines[j] = lines[j].replace(");", ") { }")
                        break

        new_lines.append(line)
        i += 1

    content = "".join(new_lines)
    if content != original_content:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    return False

files_to_fix = [
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

for f in files_to_fix:
    if os.path.exists(f):
        if fix_file(f):
            print("Fixed " + f)
        else:
            print("No changes for " + f)
