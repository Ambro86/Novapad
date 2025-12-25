[Read in English 🇺🇸](README.md)

# Novapad

**Novapad** è una moderna alternativa al Blocco Note per Windows ricca di funzionalità, costruita con Rust. Estende l'editing di testo tradizionale con il supporto per vari formati di documenti, funzionalità Text-to-Speech (TTS) e altro ancora.

## Funzionalità

- **Interfaccia Nativa Windows:** Costruita utilizzando le API di Windows per un aspetto leggero e nativo.
- **Supporto Multi-Formato:**
    - Lettura e scrittura di file di testo semplice.
    - Visualizzazione ed estrazione testo da documenti **PDF**.
    - Supporto per file **Microsoft Word (DOCX)**.
    - Supporto per **Fogli di calcolo** (Excel/ODS tramite `calamine`).
    - Supporto per e-book **EPUB**.
- **Text-to-Speech (TTS):** Funzionalità di riproduzione audio integrate.
- **Stack Tecnologico Moderno:** Basato su Rust per prestazioni e sicurezza.

## Installazione / Compilazione

Questo progetto è costruito con Rust. Assicurati di avere installato la toolchain Rust.

1.  Clona il repository:
    ```bash
    git clone https://github.com/Ambro86/Novapad.git
    cd Novapad
    ```

2.  Compila il progetto:
    ```bash
    cargo build --release
    ```

3.  Esegui l'applicazione:
    ```bash
    cargo run --release
    ```

## Dipendenze

Novapad si affida a diverse potenti librerie Rust (crates), tra cui:
- `windows-rs`: Per l'integrazione nativa con le API di Windows.
- `printpdf` & `pdf-extract`: Per la gestione dei PDF.
- `docx-rs`: Per il supporto ai documenti Word.
- `rodio`: Per la riproduzione audio.
- `tokio`: Per le operazioni asincrone.

## Licenza

Questo progetto è concesso in licenza sotto la [Licenza Creative Commons](https://creativecommons.org/licenses/by/4.0/).

## Autore

Ambrogio Riili
