import os
import re

def fix_syntax_errors(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # 1. Fix missing 'if' before with_state(...).is_none()
        # Look for lines starting with with_ and containing .is_none() {
        # but NOT already having 'if ' at the start (ignoring leading whitespace)
        match_with = re.match(r'^(\s*)(with_[a-z_]+\(.*\)\.is_none\(\) \{)$', line)
        if match_with:
            indent = match_with.group(1)
            rest = match_with.group(2)
            line = indent + "if " + rest + "\n"
        
        # Multiline variant of the above
        if i + 1 < len(lines) and "}).is_none() {" in lines[i+1]:
             match_start = re.match(r'^(\s*)(with_[a-z_]+\(.*)$', line)
             if match_start:
                 indent = match_start.group(1)
                 rest = match_start.group(2)
                 line = indent + "if " + rest + "\n"

        # 2. Fix if let Err(e) = ...; -> wrap in braces
        # Single line: if let Err(e) = some_func();
        match_err = re.match(r'^(\s*)if let Err\(e\) = ([^;\{]+);$', line)
        if match_err:
            indent = match_err.group(1)
            call = match_err.group(2)
            line = f'{indent}if let Err(e) = {call} {{ crate::log_debug(&format!("Error: {{:?}}", e)); }}\n'

        new_lines.append(line)
        i += 1

    with open(file_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)

# Apply to main files
files = ["src/main.rs", "src/mf_encoder.rs", "src/sapi5_engine.rs"]
for f in files:
    if os.path.exists(f):
        fix_syntax_errors(f)
