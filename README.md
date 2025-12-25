[Leggilo in Italiano 🇮🇹](README_IT.md)

# Novapad

**Novapad** is a modern, feature-rich Notepad alternative for Windows, built with Rust. It extends traditional text editing with support for various document formats, Text-to-Speech (TTS) capabilities, and more.

## Features

- **Native Windows UI:** Built using the Windows API for a lightweight and native feel.
- **Multi-Format Support:**
    - Read and write plain text files.
    - View and extract text from **PDF** documents.
    - Support for **Microsoft Word (DOCX)** files.
    - Support for **Spreadsheets** (Excel/ODS via `calamine`).
    - Support for **EPUB** e-books.
- **Text-to-Speech (TTS):** Integrated audio playback features.
- **Modern Tech Stack:** Powered by Rust for performance and safety.

## Installation / Build

This project is built with Rust. Ensure you have the Rust toolchain installed.

1.  Clone the repository:
    ```bash
    git clone https://github.com/Ambro86/Novapad.git
    cd Novapad
    ```

2.  Build the project:
    ```bash
    cargo build --release
    ```

3.  Run the application:
    ```bash
    cargo run --release
    ```

## Dependencies

Novapad relies on several powerful Rust crates, including:
- `windows-rs`: For native Windows API integration.
- `printpdf` & `pdf-extract`: For PDF handling.
- `docx-rs`: For Word document support.
- `rodio`: For audio playback.
- `tokio`: For asynchronous operations.

## License

This project is licensed under the [Creative Commons License](https://creativecommons.org/licenses/by/4.0/).

## Author

Ambrogio Riili
