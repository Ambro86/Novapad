# Novapad

[Leggilo in Inglese 🇬🇧](README.md)

**Scarica l'ultima versione:**


- [Portable (EXE)](https://github.com/Ambro86/Novapad/releases/latest/download/novapad.exe)
- [Installer (Setup)](https://github.com/Ambro86/Novapad/releases/latest/download/novapad_x64-setup.exe)
- [Installer (MSI)](https://github.com/Ambro86/Novapad/releases/latest/download/novapad_x64_en-US.msi)
**Novapad** è un Notepad moderno e avanzato per Windows, sviluppato in Rust.
Estende il classico editor di testo con il supporto a più formati di documento,
funzionalità avanzate di accessibilità e capacità di Text-to-Speech (TTS).

Include inoltre un **player MP3 per audiolibri**, un **sistema di segnalibri per testo e audio**
e la possibilit… di **creare audiolibri direttamente dal testo utilizzando le voci Microsoft (Edge Neural) e SAPI5**.

> ⚠️ **Avviso di licenza**
> Questo progetto è **source-available ma NON open source**.
> L’uso commerciale, la redistribuzione e la creazione di opere derivate
> sono espressamente vietati.

---

## Funzionalità

- **Interfaccia nativa Windows**
  Costruita direttamente sulle Windows API per garantire prestazioni elevate
  e piena integrazione con le tecnologie di accessibilità.
- **Supporto multi-formato**
  - File di testo semplice
  - Documenti PDF (estrazione del testo)
  - Documenti Microsoft Word (DOCX)
  - Fogli di calcolo (Excel / ODS tramite `calamine`)
  - E-book EPUB
- **Text-to-Speech (TTS) e creazione di audiolibri**
  - Lettura vocale dei documenti tramite le voci Microsoft (Edge Neural) e SAPI5 (incluse OneCore)
  - Creazione di audiolibri in formato MP3 direttamente dal testo
  - Divisione audiolibri in parti fisse o in base a testo (case sensitive, inizio riga). Esempio: con "Capitolo" crea una parte per ogni capitolo; include autore e introduzione nella prima parte fino al primo Capitolo. Altre opzioni: 2, 4, 6, 8 parti
  - Supporto voci Microsoft e SAPI5/OneCore per lettura e salvataggio audiolibri
  - Aggiunta di voci ai preferiti e cambio rapido durante la lettura
  - Dizionario con sostituzioni personalizzate applicate alla lettura e agli audiolibri
- **Player MP3 (audiolibri)**
  - Apertura e riproduzione di file MP3
  - Avanzamento e riavvolgimento con i tasti freccia
  - Play/Pausa con la barra spaziatrice
  - Volume su/giù con i tasti freccia
- **Segnalibri**
  - Creazione e gestione di segnalibri sia per file di testo sia per la riproduzione MP3
  - Salto rapido alle posizioni salvate nei documenti o nell’audio
- **Accessibilità**
  Progettato per funzionare correttamente con screen reader
  come NVDA e JAWS.
- **Prompt accessibile**
  - Finestra prompt dedicata con output stabile per gli screen reader
  - Scorciatoie: Ctrl+Shift+P, Alt+I (input), Alt+O (output)
  - Opzioni: scorrimento automatico, rimuovi ANSI, beep dopo inattivita', evita sospensione; scelta tra cmd/PowerShell/Codex CLI
- **Opzioni di leggibilita'**
  - Controlli per colore e dimensione del testo per una migliore lettura, con colori chiari/scuri e dimensioni grandi
- **Strumenti di modifica**
  - Trova nei file, Pulisci testo Markdown, Normalizza spazi bianchi, Riformatta righe
- **Regolazione voci**
  - Impostazioni di tono, velocita' e volume per le voci (Microsoft e SAPI5), valide per lettura e creazione audiolibri
- **Importazione trascrizioni YouTube**
  - Importa le trascrizioni con selezione lingua e timestamp opzionali.
- **Localizzazione**
  - Italiano, inglese, spagnolo, portoghese.
- **Tecnologia moderna**
  Scritto in Rust per garantire sicurezza, affidabilità e ottime prestazioni.

---

## Compilazione e utilizzo

Assicurati di avere installato il toolchain Rust.
La formattazione del codice e' gestita con `cargo fmt`.

Clona il repository:

```bash
git clone https://github.com/Ambro86/Novapad.git
cd Novapad
```

Compila il progetto:

```bash
cargo build --release
```

Avvia l’applicazione:

```bash
cargo run --release
```

---

## Aspetti legali e licenza

Questo repository è pubblicato **esclusivamente per scopi di trasparenza,
studio, valutazione e uso personale**.

### È consentito:
- Visualizzare e studiare il codice sorgente
- Compilare ed eseguire il software per uso personale o di test

### NON è consentito:
- Utilizzare il software per scopi commerciali
- Redistribuire il codice sorgente o i binari
- Effettuare fork del repository per la distribuzione
- Integrare Novapad in altri progetti o prodotti
- Creare e distribuire opere derivate senza autorizzazione scritta

Le funzionalità di Text-to-Speech possono utilizzare voci Microsoft
e sono soggette ai termini di servizio Microsoft.
**L’uso commerciale è espressamente vietato.**

Per i dettagli completi fare riferimento al file `LICENSE`.

---

## Autore

**Ambrogio Riili**
