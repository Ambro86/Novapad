# Changelog

Version 0.5.9 - 2025-01-13
New features
• Added RSS reordering from the context menu (up/down/to position) with invalid-position checks.
• Added an article context menu with open original site and share via WhatsApp, Facebook, and X.
• Added Esc shortcut to return from imported articles to the RSS list.
• Added podcast mode: search, subscribe, listen; reorder subscriptions; Esc stops playback and returns to the list; Enter on an episode starts playback.
• Added playback speed control for podcasts and MP3 files.
• Added Ctrl+T to jump to a specific time.
• Added a voice preview button after the volume combo.
• Added regex find and replace (Notepad++ style).
• Added RSS import from OPML and TXT files.
• Added an option to enable "Open with Novapad" in File Explorer, including portable builds.
Improvements
• Improved voice speed/pitch/volume selection, respecting TTS max limits.
• Various RSS improvements to download all articles without moving NVDA focus during updates.
• Improved audio playback with a dedicated menu, Ctrl+I time announce, and volume up to 300%.
• Added missing shortcuts for some functions.
• Reorganized the Edit menu with a text cleanup submenu.
• Reorganized Options into tabs, with Ctrl+Tab and Ctrl+Shift+Tab navigation.
• RSS reader now downloads full article content, matching the browser view.
Fixes
• Fixed Markdown cleanup removing numbers at the start of lines.
• Fixed AltGr+Z triggering undo.
• Fixed audiobook recording cancellation so it stops quickly.
Localization
• Added Vietnamese translation (thanks to Anh Đức Nguyễn).

Version 0.5.8 - 2026-01-10
New features
• Added volume control for the microphone and system audio when recording podcasts.
• Added a new feature to import articles from websites or RSS feeds, including the most important feeds for each language.
• Added a function to remove all bookmarks for the current file.
• Added a function to remove duplicate lines and duplicate consecutive lines.
• Added a function to close all tabs or windows except the current one.
• Added a Donations entry in the Help menu for all languages.
Improvements
• Improved the accessible terminal to prevent some crashes.
• Improved and fixed access keys and keyboard shortcuts across the app.
• Fixed an issue where closing the audio playback window did not stop playback.
• Added confirmation dialogs for important actions (e.g., remove duplicate lines, remove end-of-line hyphens, remove all bookmarks in the current file). No dialog is shown when the action does not apply.
• Added the ability to delete RSS feeds/sites from the library by selecting them and pressing Delete.
• Added a context menu in the RSS window to edit or delete RSS feeds/sites.
• Removed the setting to move settings to the current folder; the app now handles this automatically based on location (if the exe folder is named "novapad portable" or the exe is on a removable drive, settings go to the exe folder in `config`, otherwise `%APPDATA%\\Novapad`, with fallback to the exe `config` if the preferred folder is not writable).

Version 0.5.7 - 2026-01-05
New features
• Added Batch Audiobooks feature to convert multiple files/folders at once.
• Added support for Markdown files (.md).
• Added file encoding selection when opening text files.
• Added option in the accessible terminal to announce new lines with NVDA.
Improvements
• Audiobook recording now saves natively to MP3 when selected.
• User can now choose the position of the "unsaved changes" asterisk (*) in the window title.
• Improved the update system robustness across different scenarios.
• Added "Remove Hyphens" in Edit menu to fix OCR line-endings.

Version 0.5.6 - 2026-01-04
Fixes
  Improved Find in Files so pressing Enter opens the file exactly at the selected snippet.
Improvements
  Added PPT/PPTX support (open as text).
  Opening non-text formats now saves as .txt to avoid formatting corruption (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Added podcast recording from microphone and system audio (File menu, Ctrl+Shift+R).

Version 0.5.5 – 2026-01-03
New features
• Added an accessible terminal optimized for large output and screen readers (Ctrl+Shift+P).
• Added a setting to save user settings in the current folder (portable mode).
Fixes
• Improved Find in Files snippets so the preview stays aligned with the match.

Version 0.5.4 – 2026-01-03
Improvements
• Fixed Normalize Whitespace (Ctrl+Shift+Enter).
• Added HTML/HTM support (open as text).

Version 0.5.3 – 2026-01-02
New features
• Added Find in Files.
• Added new text tools: Normalize Whitespace, Hard Line Break, and Strip Markdown.
• Added Text Statistics (Alt+Y).
• Added new list commands in the Edit menu:
• Order Items (Alt+Shift+O)
• Keep Unique Items (Alt+Shift+K)
• Reverse Items (Alt+Shift+Z)
• Added Quote / Unquote Lines (Ctrl+Q / Ctrl+Shift+Q).
Localization
• Added Spanish localization.
• Added Portuguese localization.
Improvements
• When an EPUB file is open, Save now automatically switches to Save As and exports the content as a .txt file to prevent EPUB corruption.

## 0.5.2 - 2026-01-01
- Added a changelog.
- Added open-with-Novapad options and file associations for supported files during installation.
- Improved message localization (errors, dialogs, audiobook export).
- Added part selection when using "Split audiobook based on text", with a "Require the marker at line start" option.
- Added YouTube transcript import with language selection, timestamp option, and improved focus handling.

## 0.5.1 - 2025-12-31
- Automatic updates with confirmation, improved error handling and notifications.
- Audiobook export improvements (text-based split, SAPI5/Media Foundation, advanced controls).
- TTS improvements (pause/resume, replacement dictionary, favorites).
- View menu and voice/favorites panels, text color and size.
- Default language from system locale and localization improvements.
- CI and Windows packaging (artifacts, MSI/NSIS, cache).

## 0.5.0 - 2025-12-27
- Modular refactor (editor, file handler, menu, search).
- Windows build/packaging workflow and README/license updates.
- Fix TAB navigation in the Help window.

## 0.5 - 2025-12-27
- Preliminary version bump.

## 0.1.0 - 2025-12-25
- Initial release: project structure and README.
