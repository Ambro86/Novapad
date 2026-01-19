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
        
        # 1. Add missing 'if' before with_state calls that end with .is_none() {
        # Check if line starts with with_ (ignoring whitespace) and has matching ).is_none() {
        stripped = line.lstrip()
        if (stripped.startswith("with_") or stripped.startswith("with_podcast_state") or \
            stripped.startswith("with_save_state") or stripped.startswith("with_import_state") or \
            stripped.startswith("with_marker_state") or stripped.startswith("with_help_state") or \
            stripped.startswith("with_find_state") or stripped.startswith("with_options_state") or \
            stripped.startswith("with_batch_state") or stripped.startswith("with_progress_state") or \
            stripped.startswith("with_prompt_state")) and "(| " in stripped:
            
            indent = line[:line.find("with_")]
            # Look ahead for ).is_none() {
            found = False
            for j in range(i, min(i + 100, len(lines))):
                if ").is_none() {" in lines[j] or "}.is_none() {" in lines[j]:
                    found = True
                    break
                if "fn " in lines[j] and lines[j].strip().startswith("fn "):
                    break
            
            if found:
                line = indent + "if " + stripped

        # 2. Fix if let Err(e) = ...; -> if let Err(e) = ... { log }
        if "if let Err(e) = " in line and line.strip().endswith(");"):
            indent = line[:line.find("if let Err")]
            code = line.strip()[:-1]
            line = indent + code + " { crate::log_debug(&format!(\"Error: {:?}\", e)); }" + chr(10)
        
        # 3. Special case for MessageBoxW which was also broken
        if "if MessageBoxW(" in line and line.strip().endswith(");"):
             indent = line[:line.find("if MessageBoxW")]
             code = line.strip()[3:-1] # remove 'if ' and ';'
             line = indent + code + ";" + chr(10)

        new_lines.append(line)
        i += 1

    content = "".join(new_lines)
    
    # Final cleanup for specific main.rs version mismatch
    if "src/main.rs" in file_path:
        content = content.replace("if last_seen != current_version", "if last_seen != &current_version")

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
