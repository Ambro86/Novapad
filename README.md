# Sonarpad

[Read it in Italian 🇮🇹](README_IT.md)

**Download the latest release:**


- [Portable (EXE)](https://github.com/Ambro86/Sonarpad/releases/latest/download/sonarpad.exe)
- [Installer (Setup)](https://github.com/Ambro86/Sonarpad/releases/latest/download/sonarpad_x64-setup.exe)
- [Installer (MSI)](https://github.com/Ambro86/Sonarpad/releases/latest/download/sonarpad_x64_en-US.msi)
- [Mac Apple Silicon (DMG)](https://github.com/Ambro86/Sonarpad-Mac/releases/latest/download/Sonarpad-macOS-AppleSilicon.dmg)
- [Mac Intel (DMG)](https://github.com/Ambro86/Sonarpad-Mac/releases/latest/download/Sonarpad-macOS-Intel.dmg)
- [Mac project](https://github.com/Ambro86/Sonarpad-Mac/)
**Sonarpad** is a modern, feature-rich Notepad alternative for Windows, built with Rust.
It extends traditional text editing with multi-format document support,
advanced accessibility features, and Text-to-Speech (TTS) capabilities.

It also includes an integrated **media/audiobook player**, a **bookmark system for both text and audio**,
the ability to **create audiobooks directly from text using Microsoft voices (Edge Neural) and SAPI5**,
**audio transcription with Whisper/faster-whisper**, and **podcast recording from microphone and system audio**.

> ⚠️ **License Notice**
> This project is **source-available but NOT open source**.
> Commercial use, redistribution, and derivative works are strictly prohibited.

---

## Features

- **Native Windows UI**
  Built directly on the Windows API for maximum performance and accessibility.
- **Multi-Format Support**
  - Plain text files
  - PDF documents (text extraction)
  - Microsoft Word (DOCX)
  - Spreadsheets (Excel / ODS via `calamine`)
  - EPUB e-books
- **Text-to-Speech (TTS) & Audiobook Creation**
  - Read documents aloud using Microsoft voices (Edge Neural) and SAPI5 (including OneCore)
  - Create MP3 audiobooks directly from text
  - Split audiobooks by fixed parts or by marker text (case sensitive, line start). Example: with "Chapter" it creates one part per chapter; it includes author/introduction in the first part up to the first Chapter. Other options: 2, 4, 6, 8 parts
  - Supports both Microsoft voices and SAPI5/OneCore for playback and audiobook saving
  - Add voices to favorites and switch quickly during reading
  - Dictionary with user-defined word replacements applied during reading and audiobook creation
- **Media and Streaming Playback**
  - Open and play local media files (including MP3 audiobooks)
  - Open and play direct streaming URLs
  - Seek forward/backward using arrow keys
  - Play/Pause with the Space bar
  - Volume up/down using arrow keys
- **Audio-description creation**
  - Dialogue analysis with Pyannote and description generation with Gemini
  - Narration through Sonarpad's existing Edge, Google or SAPI voices
  - Ducking, extended pauses and MP3 export through the bundled Rust FFmpeg libraries
  - Optional editable project containing only descriptions actually present in the MP3
  - Text editing and re-export without repeating Gemini analysis
- **Audio Transcription (Whisper / faster-whisper)**
  - Transcribe the currently playing audio/media
  - Optional timestamped transcript output
  - Optional CUDA acceleration (NVIDIA GPUs) with runtime package download
  - Automatic bridge/model download on first use
- **Bookmarks**
  - Create and manage bookmarks for both text files and MP3 playback
  - Quickly jump to saved positions in documents or audio
- **Podcast Recording**
  - Record from microphone and/or system audio (`Voice and audio` menu, Ctrl+Shift+R)
- **Voice and Audio Menu**
  - Start reading (F5)
  - Pause reading (F4)
  - Stop reading (F6)
  - Record audiobook (Ctrl+R)
  - Batch audiobooks (Ctrl+Shift+B)
  - Record podcast (Ctrl+Shift+R)
  - Convert Audio (Ctrl+Shift+A)
  - Transcribe current audio/media
- **Accessibility-Focused**
  Designed to work correctly with screen readers such as NVDA and JAWS.
- **Accessible Terminal**
  - Dedicated terminal window with stable output for screen readers
  - Shortcuts: Ctrl+Shift+P, Alt+I (input), Alt+O (output)
  - Options: auto-scroll, strip ANSI, beep on idle, prevent sleep; choose cmd/PowerShell/Codex CLI
- **RSS / Articles Reader**
  - Browse and import articles from RSS feeds
- **Readable UI Options**
  - Text color and text size controls for better readability, including light/dark colors and larger sizes
- **Editing tools**
  - Find in Files, Strip Markdown Tags, Normalize Whitespace, Hard Line Break
- **Voice Tuning Options**
  - Set pitch, speed, and volume for voices (Microsoft and SAPI5), applied to reading and audiobook creation
- **Modern Tech Stack**
  Written in Rust for safety, performance, and reliability.
- **YouTube Transcript Import**
  - Import YouTube captions with language selection and optional timestamps.
- **Localization**
  - English, Italian, Spanish, Portuguese, Swedish.

---

## Build Instructions

Ensure you have the Rust toolchain installed.
Formatting is enforced with `cargo fmt --check`.
Clippy runs in CI as advisory (Windows/TTS glue still emits warnings).
Core text logic is the target for stricter linting/tests.

```bash
git clone https://github.com/Ambro86/Sonarpad.git
cd Sonarpad
cargo build --release
```

Run the application:

```bash
cargo run --release
```

---

## Legal & Licensing

This repository is published for **transparency, evaluation, and personal use only**.

### You MAY:
- View and study the source code
- Build and run the software for personal or evaluation purposes

### You MAY NOT:
- Use the software for commercial purposes
- Redistribute the source code or binaries
- Fork this repository for redistribution
- Include this software in other products or projects
- Create or distribute derivative works without written permission

Text-to-Speech features may rely on Microsoft voices and are subject to
Microsoft Terms of Service.
**Commercial usage is explicitly prohibited.**

Refer to the `LICENSE` file for full terms.

## Third-Party Binaries and Notices

This project includes third-party binaries (for example BASS-related DLLs and optional CUDA runtime files) and may download additional third-party assets at runtime.

See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for details and upstream references.

If you redistribute binaries, you are responsible for compliance with all third-party licenses and redistribution terms.

---

## Author

**Ambrogio Riili**
