import sys
import re
import os

def main():
    # Targets all with_* functions that commonly cause this
    with_funcs = [
        "with_state", "with_podcast_state", "with_save_state", "with_import_state",
        "with_marker_state", "with_help_state", "with_find_state", "with_options_state",
        "with_batch_state", "with_progress_state", "with_prompt_state"
    ]

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
            stripped = lines[i].lstrip()
            # If it starts with a with_ function and NOT with 'if ', 'let ', 'return ', etc.
            starts_with_with = any(stripped.startswith(f + "(") for f in with_funcs)
            if starts_with_with and not (stripped.startswith("if ") or stripped.startswith("let ") or stripped.startswith("return ")):
                # Look ahead for .is_none() {
                is_none_found = False
                for j in range(i, min(i + 50, len(lines))):
                    if ").is_none() {" in lines[j] or "}.is_none() {" in lines[j]:
                        is_none_found = True
                        break
                    if "fn " in lines[j] and lines[j].strip().startswith("fn "):
                        break
                
                if is_none_found:
                    indent = lines[i][:lines[i].find(stripped[0:5])] # find approximate indent
                    # Precise indent
                    for f_name in with_funcs:
                        if stripped.startswith(f_name):
                            indent = lines[i][:lines[i].find(f_name)]
                            break
                    
                    lines[i] = indent + "if " + stripped
                    changed = True
                    print(f"Fixed {file_path} at line {i+1}")

        if changed:
            with open(file_path, "w", encoding="utf-8") as f:
                f.writelines(lines)

if __name__ == "__main__":
    main()