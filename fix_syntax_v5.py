import os
import re

def fix_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    original_content = content
    
    # 1. Fix if let Err(_e) = ... { ... e ... } -> use _e everywhere inside
    # We find the range of the block and replace e with _e
    def fix_err_blocks(match):
        block = match.group(0)
        # Only replace ' e ' or ' e,' or ' e)' or ' e.' to avoid partial words
        block = re.sub(r'(\W)e(\W)', r'\1_e\2', block)
        return block

    content = re.sub(r'if let Err\(_e\) = .*? \{.*?\}', fix_err_blocks, content, flags=re.DOTALL)

    # 2. Fix the specific cases of missing braces or extra semicolons
    # This is hard to do globally, but we can fix the common patterns
    
    # Add missing 'if let Some' or 'if' where it was missed
    # Pattern: .flatten() 
 { ... } -> .flatten(); 
 if let Some(x) = ... { 
    # This is too specific to main.rs:160, let's look for others
    
    lines = content.splitlines(keepends=True)
    new_lines = []
    with_funcs = ["with_state", "with_podcast_state", "with_save_state", "with_import_state", "with_marker_state", "with_help_state", "with_find_state", "with_options_state", "with_batch_state", "with_progress_state", "with_prompt_state"]
    
    for i in range(len(lines)):
        line = lines[i]
        stripped = line.lstrip()
        
        # Add missing 'if' before with_* calls ending in .is_none()
        is_with_start = any(stripped.startswith(f + "(") for f in with_funcs)
        if is_with_start and not any(stripped.startswith(prefix) for prefix in ["if ", "let ", "return ", "match ", "unsafe ", "pub ", "fn ", "Some(", "None", "Ok(", "Err("]):
            found_is_none = False
            for j in range(i, min(i + 50, len(lines))):
                if ").is_none() {" in lines[j] or "}.is_none() {" in lines[j]:
                    found_is_none = True
                    break
            if found_is_none:
                indent = line[:line.find(stripped[:5])]
                for f_name in with_funcs:
                    if stripped.startswith(f_name):
                        indent = line[:line.find(f_name)]
                        break
                line = indent + "if " + stripped

        new_lines.append(line)
    content = "".join(new_lines)

    if content != original_content:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    return False

files = [
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

for f in files:
    if os.path.exists(f):
        if fix_file(f):
            print("Fixed " + f)
