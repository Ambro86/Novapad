import os
import re

def fix_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    original_content = content
    
    # Heuristic: Find any with_* call that is preceded by whitespace/start of line
    # and followed by ).is_none() { or }.is_none() {
    
    with_funcs = [
        "with_state", "with_podcast_state", "with_save_state", "with_import_state",
        "with_marker_state", "with_help_state", "with_find_state", "with_options_state",
        "with_batch_state", "with_progress_state", "with_prompt_state"
    ]
    
    for func in with_funcs:
        # Pattern: look for func( but NOT preceded by "if ", "let ", "return ", etc.
        # We use a negative lookbehind for common starters.
        # The [ \t]* matches leading whitespace.
        # The [^a-zA-Z0-0_] ensures we are at the start of a word.
        
        regex = r'(?m)^([ \t]*)(?<!if )(with_[a-z_]+\(.*\)(?:\.is_none\(\) \{|\.unwrap_or\(.*\)))'
        # This is hard because of multiline.
        
        # Let's use a line-by-line approach with state
        lines = content.splitlines(keepends=True)
        new_lines = []
        for i in range(len(lines)):
            line = lines[i]
            stripped = line.lstrip()
            
            is_with_start = any(stripped.startswith(f + "(") for f in with_funcs)
            if is_with_start and not (stripped.startswith("if ") or stripped.startswith("let ") or \
                                     stripped.startswith("return ") or stripped.startswith("match ") or \
                                     stripped.startswith("unsafe ")):
                
                # Look ahead for is_none or unwrap or just semicolon
                found_ending = False
                for j in range(i, min(i + 100, len(lines))):
                    if ").is_none() {" in lines[j] or "}.is_none() {" in lines[j]:
                        found_ending = True
                        break
                    if "fn " in lines[j] and lines[j].strip().startswith("fn "):
                        break
                
                if found_ending:
                    indent = line[:line.find(stripped[:5])] # rough indent
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
