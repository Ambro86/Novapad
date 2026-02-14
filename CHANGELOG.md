# Changelog

Version 0.6.7 – 2026-02-13
Improvements
• Added a new View > Read-only mode to lock editor text from accidental edits while keeping documents fully readable and navigable.
• Added a new View menu toggle for Word wrap, so line wrapping can be switched quickly without opening Options.
• Added new Edit > Text actions for indent/outdent, with shortcuts Ctrl+Alt+Right (indent) and Ctrl+Alt+Left (outdent), because when “Show voices in editor” is enabled the Tab key is reserved for voice-panel navigation.
• Added localized date/time in RSS articles and podcast episodes, with formatting adapted to the current interface language.
• Added a new RSS context-menu action to share the selected article by email.
• Added granular delete-confirmation options for RSS/Podcast in Options > RSS and podcast: RSS (feed/article/both/none) and Podcasts (podcast/episode/both/none).
• Added configurable quick RSS copy with Ctrl+C (Options > RSS and podcast): copy title, URL, article content, or all combined.
• Pressing Ctrl+A now announces completion for clearer screen-reader feedback.
• Improved replace feedback messages with proper singular/plural forms (e.g. “1 replacement made” vs “2 replacements made”).
• Added dictionary lookup language selection in the dictionary window, with default Auto (interface language) and optional manual override.
Bug fixes
• Fixed Google News RSS item ordering: articles are now shown in descending publish date (newest first) when dates are available.
• Fixed NVDA label association in the dictionary window: search field and language combo now announce the correct labels.
• Fixed RSS/Podcast Properties window keyboard handling: Tab/Shift+Tab now reaches the OK button, Enter activates OK, Esc closes safely, and focus correctly returns to the RSS/Podcast list.
• Fixed RSS/Podcast undo history: Ctrl+Z now supports multi-level undo for removals (articles/episodes and sources), not just the last action.
• Improved RSS/Podcast removal feedback with explicit status announcements (RSS removed, RSS article removed, podcast episode removed).
• Improved RSS/Podcast focus behavior after delete/undo: RSS now reliably focuses the first feed when needed and avoids repeated screen-reader announcements during delayed reselection.

Version 0.6.6 – 2026-02-13
Improvements
• Added "Auto format for TTS" in the Edit menu to quickly prepare text for speech (removes markdown/quotes and reflows wrapped lines).
• Improved voice-tag insertion: when text is selected, tags are now applied correctly to both single-line and multi-line selections.
• Added a default audiobook save folder option in Audio settings (default: Documents\\Sonarpad Audiobooks).
• In the audiobook save dialog, when split mode is enabled, added a new default-on option to create a dedicated subfolder for split parts (for cleaner output organization).
• Audiobook export now saves MP3 in stereo with user-selected bitrate for Edge, SAPI5, and SAPI4 voices.
• Added support for 32-bit SAPI5 voices via bridge, so voices available only in 32-bit engines can also be used in Sonarpad.
• Reorganized voice features into a dedicated "Voice and audio" menu and added/clarified "Convert Audio", useful for converting any supported media file to MP3, AAC, OGG, Opus, FLAC, WAV, and AIFF.
• Added removal of individual RSS articles and podcast episodes (Delete key + context menu with confirmation), without removing the entire RSS/podcast source, plus undo for the last removal (single article/episode or entire RSS/podcast source).
• Added RSS feed export to OPML in the RSS window, so current RSS sources can be saved and re-imported easily.
• Added "Search RSS by keyword" in the RSS window: entering a keyword now generates a Google News RSS URL automatically and opens the add-source dialog prefilled, so keyword feeds can be created in one step.
• Added Serbian translation thanks to Mila Kuran.
• Added Ukrainian translation thanks to Ivan Shtefuriak.
• Added multi-file media opening: selecting/opening multiple media files now builds a playback queue instead of replacing the current file.
• Added variable seek shortcuts during playback: with a 1-minute base skip, Left/Right seeks 60s, Shift+Left/Right seeks 20s, and Ctrl+Left/Right seeks 3 minutes.
• Added previous/next track navigation shortcuts in the player: Ctrl+PageUp and Ctrl+PageDown.
• Added "Reset volume" and grouped reset actions into a dedicated "Reset" submenu in Playback, alongside "Reset speed" and "Reset pitch".
• Installer improvements: setup.exe now lets users choose between associating all supported file types or selecting extensions manually; MSI now exposes per-extension file-association choices in the feature tree (default remains all enabled).
• Added a new "Window" menu with "Open documents..." to quickly switch to any currently open file.
• Updated View > Font: replaced the old chooser with a quick submenu of common fonts (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia) while preserving the current text size.
• Improved RSS/Podcast announcements with a dual status model: source nodes announce "new items" when a feed/podcast has updates, while individual RSS articles and podcast episodes announce "unread"/"unplayed"; this behavior can be disabled from Options.
Bug fixes
• Fixed EPUB text extraction for books containing inline HTML comments (<!-- ... -->): chapter text is now parsed correctly instead of being partially or fully skipped.
• Fixed Spanish Wiktionary lookups and dictionary cache handling: Spanish entries like "agua" now load correctly, and old "Word not found" cache entries are no longer reused.
• Fixed RSS article import character encoding for some Spanish sources (e.g., El Mundo): accented letters and "ñ" are now preserved correctly in the temporary editor.
• Fixed ANSI text decoding for Central European files (e.g., Czech/Polish): Sonarpad now better distinguishes UTF-8 vs ANSI and chooses the correct code page (including Windows-1250) to prevent garbled diacritics.
• Fixed RSS source persistence for feeds with URL query parameters (e.g., `rss.aspx?c=...`): these feeds are now saved and restored correctly after restarting Sonarpad.
• Fixed opening Google Drive pointer files (`.gdoc`, `.gsheet`, `.gslides`) from Explorer context menu: when direct read fails with “Incorrect function (os error 1)”, Sonarpad now falls back to shell-open so the document still opens correctly.
• Fixed legacy Excel 2010 `.xls` reading: old binary Excel files are now detected and decoded correctly instead of showing garbled text (e.g. `ÐÏ_à¡±...`).
• Fixed spellcheck announcement flow: misspellings are now announced again when reviewing text later, and the same mistake is reported again if it is deleted and retyped.
• Fixed line-based text actions (e.g. Ctrl+Q / Ctrl+Shift+Q, sort/reverse/unique/join lines): selecting a single line with Shift+Down no longer merges or truncates adjacent lines.
• Fixed multi-line behavior for line-based text actions (Ctrl+Q / Ctrl+Shift+Q and related line tools): RichEdit selections using CR-only separators are now normalized correctly, so all selected lines are processed without cutting first characters.
• Extended TTS input normalization for visible whitespace symbols (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424) to prevent repeated paragraph playback with multilingual voices.
• Refined Edge TTS text sanitization with a single validation pipeline: weird/invisible spaces are normalized, long punctuation runs (like "...", "!!!", "???") are compacted, and punctuation-only chunks are skipped to prevent playback loops.
• Fixed playback time announcement (Ctrl+I) for MP3/podcast streams: current time is now clamped to track duration, and playback is auto-stopped if position runs past the end.
• Improved installer localization coverage: setup.exe now includes additional installer languages (Czech, Polish, French, Serbian), while MSI is kept as a single en-US package to avoid release confusion.
• Fixed uninstall cleanup for context menu entries: "Open with Sonarpad" is now removed reliably, including legacy registry scenarios.
• Fixed SAPI5 pause/resume reliability: F4 pause now works correctly and resume returns to the expected position instead of restarting from the beginning.
• Fixed pause + seek + resume flow for media playback: after pausing and seeking with Left/Right, pressing Space now reliably resumes from the current position instead of stopping or restarting from the beginning.

Version 0.6.5 – 2026-02-07
Improvements
• Spanish translation improved thanks to Arturo Fernandez Rivas.
• Added an option to split EPUB audiobooks by chapters.
• RSS imports now use a dedicated temporary tab (localized title); Save As converts it to a normal document.
• Screen reader messages are now also sent to JAWS when available.
Bug fixes
• Reading from the caret (F5) now starts exactly at the cursor. Previously it could start a couple of lines above because the caret offset did not match CRLF/UTF-16 positions.
• Fixed a redraw issue where typing over a selection could make earlier text temporarily disappear until the selection moved.
• Fixed EPUB chapter parsing so cover or image-only pages no longer produce spoken CSS (e.g., "padding") or "Sconosciuto" titles.
• Fixed audiobook time-splitting from EPUBs with Edge TTS failing on empty/oversized chunks ("Edge audio not sent").
• RSS articles now decode HTML entities (e.g., &quot;, &amp;, &lt;, &gt;).
• Save/Save As now suggests the existing filename when saving non-overwritable formats (e.g., EPUB) instead of the first line.
• Fixed a bug where podcasts with new episodes were not announced as unplayed, and renamed "Unheard" to "Unplayed" for a more professional label.

Version 0.6.4 – 2026-02-05
Improvements
• The program has been renamed to Sonarpad to emphasize sound and audio as the key focus.
• Added audio track selection in the Playback menu for media files with multiple audio tracks (e.g., MKV files with multiple languages).
• Podcasts now clearly indicate unheard episodes with an "Unheard" prefix before the name.
• New tag-based voice switching in text. Examples:
  - Microsoft voices (Edge): <voice edge it-IT-IsabellaNeural>Hello</voice>
  - SAPI5 voices: <voice sapi5 Microsoft Helena Desktop>Hello</voice>
  - SAPI4 voices: <voice sapi4 #1>Hello</voice>
  - With speed/pitch/volume: <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Hello</voice>
• Enriched podcast categories.
• Improved PDF reading with automatic fallback to PDFium.
• Improved the article parser for cases where content was not read in full.
• Added pitch reset to the Playback menu.
• Added a context menu option to create an audiobook from the selected text.
• Added audiobook splitting by duration, with the ability to choose the first file name.
• Localized the author label in article reading (e.g., "by", "di", "par").
• Added indentation options (tabs/spaces with width) and Tab/Shift+Tab indent/outdent on selected lines.
• Fixed Markdown cleanup to handle '*' list bullets when bullet preservation is disabled.
• Added an option to use the legacy "Novapad" name in the window title and Start Menu shortcuts.
Bug fixes
• Fixed a bug where SAPI4 audiobooks could be created differently than expected.
• Fixed a bug where seeking past the end of a media file restarted playback from the beginning.
• Find in Files window: pressing Enter on a result now opens at the correct snippet position, and Esc returns to results.
• Options window: improved visual layout across General, Voice, Editor, and Audio tabs to prevent missing or clipped controls.
• Fixed a bookmark issue when changing playback speed.
• Fixed Podcast Index categories not displaying correctly.
• Fixed apostrophes breaking reading by removing separate dialogue reading; voice tags are used instead.

Version 0.6.3 – 2026-01-30
Improvements
• Improved microphone detection.
• Added instant playback support for all formats.
Bug fixes
• Fixed crash in podcast categories window.

Version 0.6.2 – 2026-01-30
New features
• Added file execution support (Shift+F5). Users can select an interpreter (e.g., python) in Options, search for it on the computer, and pressing Shift+F5 runs the current script. HTML files open in the browser.
• Added support for Google Docs pointer files (.gdoc, .gsheet, .gslides), which automatically open in the default browser.
• Added support for M4B audiobook format (Apple/AAC).
• Added "Show episodes" option in podcast search results context menu to browse and play episodes without subscribing.
• Added "Go to Line" feature (Edit menu or Ctrl+J) to quickly jump to a specific line number.
• Added context menu options to order RSS feeds and podcasts (alphabetically or by date).
• Added Vietnamese default RSS feeds.
• Added a microphone test box in the recording dialog to check levels before starting.
• Added "Show description" for podcast episodes in the context menu.
• Added support for extended audio/video formats via FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Added synchronized subtitle reading support (srt, vtt, ass, sub, sbv, lrc, smi) with NVDA or selected voice. The program searches for a subtitle file with the same name as the media file. Added "Import subtitles" and "Remove subtitles" options in the Playback menu for files with different names.
• Added file associations for all new supported audio/video formats in the "Open with Sonarpad" context menu.
• Added pitch adjustment setting for any file.
• Added option in General settings to enable or disable anonymous error reports. Added a menu item in Help to create a diagnostic ZIP file.
• Added option to use a different voice for dialogues, both for live reading and audiobook creation.
• Added podcast categories browser to explore podcasts by category (business, art, sport, etc.).
Improvements
• Opening an audio/video file from Explorer now opens the player view directly instead of the text editor.
• Removed the OCR prompt for inaccessible PDFs; OCR is now performed automatically to improve speed and user experience.
• Improved Accessible Terminal: NVDA reading now remembers the last read line for better continuity.
• SAPI 4: Audiobook creation is now fully parallelized and nearly instant. Added a prompt to choose the number of concurrent processes.
• SAPI 4: Eliminated the WAV-to-MP3 bottleneck by converting chunks in parallel during synthesis.
• SAPI 4: Improved error handling and automatic cleanup of temporary files.
• Find dialog: Renamed "Regex" to "Regular expression" for clarity and added missing translations for search options.
• M4B Audiobooks: Better output handling; splitting by parts/markers now produces a single M4B file with proper metadata chapters including title and author.
• Player: Fixed bookmark and time announcement precision when playback speed is not 1.0x.
• Restored Ctrl+Tab and Ctrl+Shift+Tab navigation in Options.
• Added an option in the Playback menu to instantly reset speed to Normal (1.0x).
• Updated all dependencies to the latest versions for better performance and stability.
• Integrated FFmpeg with dynamic DLL loading to ensure compatibility without blocking startup.
• Updated podcast download filters to include new audio/video formats.
• Prevented Ctrl+S from saving audio/video files to avoid corruption.
• Improved YouTube transcript import making it more robust and resilient.
• Improved audiobook part splitting robustness, ensuring no text is lost.
• Installer is now fully multilingual, supporting Italian, English, Spanish, Portuguese, Swedish, and Vietnamese based on the user's system language. English is the default for unsupported systems.
• Podcast categories: pressing Enter on a category now confirms the selection (equivalent to OK button).
• Improved hang detection system to avoid false positives when modal dialogs (error messages, "text not found") are open.
Fixes
• Fixed a bug where the changelog did not open on startup.
• Fixed a bug where the OCR prompt did not appear for inaccessible PDFs opened from Explorer.
• Fixed a startup bug that could cause loss of focus or window closure immediately after opening.
• Fixed a critical bug in regex search that prevented finding text, including issues with "Wrap around" search and "Dot matches newline" option with Windows line endings.
Localization
• Added Polish translation.
• Added French translation.
• Added Czech translation (thanks to Radek Žalud and Jiri Holzinger).

Version 0.6.1 – 2026-01-20
Fixes
• Fixed a bug where enabling “Show voices in the editor” caused podcast playback to stop.
• Fixed an issue where some podcasts could not be added via URL because the URL was being truncated.
• Fixed a bug where normal URLs could no longer be added in the RSS feed feature.
• Fixed an issue where the Wikipedia language option was shown multiple times across different settings tabs.
• Removed the creation of debug files that were incorrectly generated even in release mode.
Improvements
• Improved support for Microsoft voices, which now use a dedicated playback method with a different user agent.
• Added support for MP4 files.

Version 0.6.0 – 2025-01-20
New features
• Added spell checker. From the context menu, users can check whether the current word is correct and, if not, get spelling suggestions.
• Added podcast import and export via OPML files.
• Added Podcast Index search support in addition to iTunes. Users can enter their free API key and secret (generated using only an email address).
• Added support for SAPI4 voices, both for real-time reading and audiobook creation.
• Added automatic OCR fallback for non-accessible PDFs: when no extractable text is found, the document is recognized via OCR.
• Added dictionary support using Wiktionary. Pressing the Applications key shows definitions, and when available, synonyms and translations into other languages.
• Added Wikipedia article import with search, result selection, and direct import into the editor.
• Added Shift+Enter shortcut in the RSS module to open an article directly in the original website.
Improvements
• Microphone selection is now always respected by the application.
• In the podcast window, pressing Enter on an episode now immediately announces “loading” via NVDA to confirm the action.
• In podcast search results, pressing Enter now subscribes to the selected podcast.
• Fixed and improved labels for Ctrl+Shift+O and Podcast Ctrl+Shift+P shortcuts.
• Playback speed and volume are now saved in settings and persist across all audio files.
• Added a dedicated cache folder for podcast episodes. Users can keep episodes via “Keep podcast” in the Playback menu. The cache is automatically cleaned when exceeding the user-defined size (Options → Audio).
• Improved RSS article fetching significantly using libcurl impersonation with Chrome and iPhone profiles, ensuring compatibility with ~99% of sites.
• Added read / unread state for RSS articles, with clear indication in the RSS list.
• Replace All now reports the number of replacements performed.
• Added a Delete Podcast button when navigating the podcast library using Tab.
Fixes
• Removed the redundant “pending update” entry from the Help menu (updates are already handled automatically).
• Fixed a bug where pressing Ctrl+S on an opened MP3 file would save and corrupt the file.
• Fixed a UI issue where “Batch Audiobooks” was shown as “(B)… Ctrl+Shift+B” (removed redundant label).
• Fixed smart quotes: when enabled, normal quotes are now correctly replaced with smart quotes.
• Fixed a bug where using “Go to bookmark” reset the playback speed to 1.0.
• Fixed an issue where already-downloaded podcast episodes were re-downloaded instead of using the cached version.
Keyboard shortcuts
• F1 now opens the Help guide.
• F2 now checks for updates.
• F7 / F8 now jump to the previous or next spelling error.
• F9 / F10 now quickly switch between favorite voices.
Developer improvements
• Errors are no longer silently dropped: all let _ = patterns have been removed, and errors are now explicitly handled (propagated, logged, or handled with fallbacks as appropriate).
• The project now fails to compile if there are warnings: both cargo check and cargo clippy must pass cleanly, with lints tightened and allow removed where possible.
• Custom implementations such as strlen / wcslen-style helpers have been removed. String and UTF-16 buffer lengths are now derived from Rust-owned data instead of scanning memory.
• DLL handling has been cleaned up and consolidated around libloading, avoiding custom loader logic and PE parsing.
• Hand-rolled byte parsing helpers were removed; all byte parsing now uses standard from_le_bytes / from_be_bytes on checked slices.
These changes reduce unnecessary unsafe usage, eliminate potential undefined behavior, and make the codebase more idiomatic, robust, and maintainable.

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
• Added an option to enable "Open with Sonarpad" in File Explorer, including portable builds.
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
• Removed the setting to move settings to the current folder; the app now handles this automatically based on location (if the exe folder is named "sonarpad portable" or the exe is on a removable drive, settings go to the exe folder in `config`, otherwise `%APPDATA%\\Sonarpad`, with fallback to the exe `config` if the preferred folder is not writable).

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
- Added open-with-Sonarpad options and file associations for supported files during installation.
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






