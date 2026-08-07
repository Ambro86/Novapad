# Verifica v7 – suite di test audiodescrizione

## Composizione

La suite contiene **90 test dedicati**:

1. 79 test Python in `bridge/tests`.
2. 11 test Rust con nomi che iniziano per `omni_port_`.

Gli 84 casi della suite originale di Omni Describer sono stati mantenuti o adattati alle responsabilità reali di Sonarpad. Sono stati aggiunti 6 test per il protocollo headless specifico di Sonarpad.

## Test Python eseguiti

Comandi eseguiti nell'ambiente disponibile:

```text
python bridge/run_audio_description_tests.py
python -m pytest bridge/tests -q
```

Risultati:

```text
Ran 79 tests ... OK
79 passed, 17 subtests passed
```

Sono stati verificati senza chiamate reali alla rete:

1. Timestamp relativi e assoluti, ordinamento e correzione degli offset.
2. Chunk Gemini da 180 secondi e delega della preparazione multimediale a Rust.
3. Retry per timeout, rete, HTTP 5xx, contenuti bloccati e file Gemini in elaborazione.
4. Quota esaurita con cambio modello, attesa o interruzione senza perdere il chunk corrente.
5. Validazione e recupero dell'elenco dei modelli Gemini.
6. Controllo e correzione selettiva della lingua delle descrizioni.
7. Prompt intensivi, copertura degli slot e continuità dei personaggi.
8. Attivazione e disattivazione effettiva del glossario dei personaggi.
9. Modello Pyannote ONNX incluso, checksum, cache della sessione e post-elaborazione NumPy.
10. Intervalli protetti, spostamento massimo di 5 secondi, descrizioni escluse e pause estese.
11. Protocollo `QUOTA` e validazione dei chunk preparati dalle librerie FFmpeg Rust.

## Test Rust aggiunti

Gli 11 test Rust verificano:

1. Rimozione del silenzio finale Edge senza tagliare il parlato.
2. Conservazione di una coda silenziosa breve.
3. Missaggio con durata invariata e modifica soltanto degli intervalli previsti.
4. Inserimento di una pausa estesa e allungamento della timeline.
5. Round-trip del progetto JSON con i tempi realmente inseriti nell'MP3.
6. Compatibilità dei progetti precedenti senza il campo dei personaggi.
7. Retry Edge in assenza di audio senza limite artificiale.
8. Arresto del retry infinito per parametri voce non validi.
9. Retry per audio vuoto o errore di decodifica.
10. Parallelismo Edge limitato a otto operazioni.
11. Esclusione dei motori SAPI dai batch paralleli Edge.

## Automazione

La workflow `.github/workflows/ci.yml` ora:

1. configura Python 3.14;
2. installa `bridge/requirements-test.txt`;
3. esegue i 79 test Python;
4. esegue successivamente `cargo test --all`, che comprende gli 11 test Rust.

Lo script locale `run_audio_description_tests.ps1` consente di eseguire entrambe le parti con un solo comando.

## Limite della verifica locale

In questo ambiente non sono installati `cargo`, `rustc`, Win32 e le librerie FFmpeg Windows necessarie alla compilazione completa di Sonarpad. I 79 test Python sono stati realmente eseguiti e superati. Gli 11 test Rust sono stati controllati staticamente, ma devono essere compilati ed eseguiti su Windows con:

```text
cargo test omni_port_ -- --nocapture
```

La CI Windows è configurata per eseguirli automaticamente attraverso il normale `cargo test --all`.
