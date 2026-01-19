import os
import re

def fix_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    original_content = content
    
    # 1. Fix argument never used errors by using _e instead of e
    # Pattern: if let Err(_e) = ... { ... e ... }
    # We should replace e with _e inside the block if it was changed by previous script
    content = re.sub(r'if let Err\(_e\) = (.*) \{ (.*)e(.*) \}', r'if let Err(_e) = \1 { \2_e\3 }', content)

    # 2. Fix the specific cases of }; inside unsafe/if blocks
    # This is better done with a line-by-line approach
    lines = content.splitlines(keepends=True)
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Remove suspicious isolated }; that likely break blocks
        if line.strip() == "};" and i > 0:
            prev = lines[i-1].strip()
            if ".unwrap_or" in prev:
                # This might be correct closing of a let assignment, 
                # but if the NEXT line is an 'if', it might need to be just '}'
                if i + 1 < len(lines) and lines[i+1].lstrip().startswith("if "):
                    line = line.replace("};", "}")
        
        new_lines.append(line)
        i += 1
    
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
