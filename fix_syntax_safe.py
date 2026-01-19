import os
import re

def fix_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    original_content = content
    
    # Target: lines starting with whitespace and with_state (or other with_*)
    # that eventually (potentially many lines later) end with .is_none() {
    # and don't already have 'if ' at the start.
    
    lines = content.splitlines(keepends=True)
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.lstrip()
        
        # Look for the start of a with_* block
        if (stripped.startswith("with_state") or stripped.startswith("with_podcast_state") or \
            stripped.startswith("with_save_state") or stripped.startswith("with_import_state") or \
            stripped.startswith("with_marker_state") or stripped.startswith("with_help_state") or \
            stripped.startswith("with_find_state") or stripped.startswith("with_options_state") or \
            stripped.startswith("with_batch_state") or stripped.startswith("with_progress_state") or \
            stripped.startswith("with_prompt_state")) and "(|" in stripped:
            
            indent = line[:line.find("with_")]
            
            # Find the balanced closing brace and see if it's followed by .is_none() {
            # This is a bit complex, let's use a simpler heuristic first: 
            # if we see .is_none() { before the next function definition or major block
            
            found_is_none = False
            for j in range(i, min(i + 100, len(lines))):
                if ").is_none() {" in lines[j] or "}.is_none() {" in lines[j]:
                    found_is_none = True
                    break
                if "fn " in lines[j] and lines[j].strip().startswith("fn "): # stop at next function
                    break
            
            if found_is_none:
                line = indent + "if " + stripped

        # Also fix the if let Err(e) = ...; case which might be multi-line
        if "if let Err(e) = " in line:
            # If it ends with ); it's definitely the bug
            if line.strip().endswith(");"):
                indent = line[:line.find("if let Err")]
                code = line.strip()[:-1]
                line = indent + code + " { crate::log_debug(&format!(\"Error: {{:?}}\", e)); }" + chr(10)
            else:
                # Check next few lines for the closing );
                for j in range(i + 1, min(i + 10, len(lines))):
                    if lines[j].strip() == ");":
                        # We found it! We need to wrap it.
                        # This is getting complex, let's do it manually if it fails.
                        pass

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
