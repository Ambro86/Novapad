with open("src/main.rs", "r", encoding="utf-8") as f:
    content = f.read()

stack = []
for i, char in enumerate(content):
    if char == "{":
        stack.append(i)
    elif char == "}":
        if not stack:
            print(f"Extra closing brace at index {i}")
        else:
            stack.pop()

if stack:
    print(f"Total unclosed braces: {len(stack)}")
    for start_idx in stack:
        line_num = content.count('\n', 0, start_idx) + 1
        print(f"Unclosed brace at line {line_num} (index {start_idx})")
        # Context
        start = max(0, start_idx - 20)
        end = min(len(content), start_idx + 50)
        print(f"  Context: {content[start:end].replace('\n', '\\n')}")
else:
    print("All braces are balanced.")