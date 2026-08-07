# Sonarpad Audio Description Bridge

Worker Python headless per l'analisi Pyannote ONNX e la generazione Gemini del modulo **Crea audiodescrizione**.

Il worker contiene esclusivamente Pyannote ONNX e la logica Gemini necessaria a produrre testi e timestamp. Non include `accessible-output2`, SAPI5, `pyttsx3`, Edge TTS, Google TTS, VLC, riproduzione, ducking o esportazione. Queste funzioni appartengono al codice Rust di Sonarpad.

Protocollo stdout:

```text
STATUS:{json}
PROGRESS:0-100
QUOTA:{"model":"...","error":"..."}
RESULT:{json}
```

Costruzione e pubblicazione sono descritte in `../AUDIO_DESCRIPTION_INTEGRATION.md`.

Il worker non include e non avvia `ffmpeg.exe` o `ffprobe.exe`. Riceve da Sonarpad un WAV mono 16 kHz per Pyannote e i chunk video già creati dalle librerie FFmpeg Rust del programma.

Il worker non sceglie alcuna voce e non crea file vocali. Il risultato JSON contiene soltanto intervalli protetti, descrizioni testuali e timestamp; Sonarpad sintetizza ogni descrizione con il motore selezionato dall’utente e crea direttamente l’MP3 finale.

Quando emette `QUOTA`, il worker non perde il lavoro già completato e attende una riga JSON su stdin: `{"action":"switch","model":"..."}`, `{"action":"wait"}` oppure `{"action":"stop"}`. Il modello scelto viene applicato alla richiesta corrente e ai chunk successivi.
La richiesta contiene anche `recognize_characters`. Quando è `true`, il worker attiva glossario e continuità nominale; quando è `false`, impone un glossario vuoto e non identifica i personaggi per nome.

## Test derivati da Omni Describer

Gli 89 test Python sono in `bridge/tests`: 73 portano la copertura headless di Omni e 16 verificano protocollo, lingue, glossario, localizzazione, stati strutturati e script di costruzione specifici di Sonarpad. Insieme ai 17 test Rust formano una suite di 106 test. Vedere `AUDIO_DESCRIPTION_TESTS.md` e `run_audio_description_tests.ps1` nella cartella principale.


## Catalogo persistente dei personaggi (v22)

La richiesta può includere `initial_character_glossary`, una lista opzionale di oggetti con `id`, `name` e `description`. Il worker la normalizza, la usa come continuità stabilita per la serie e restituisce il glossario aggiornato in `character_glossary`. Il salvataggio su disco resta responsabilità del backend Rust di Sonarpad e avviene soltanto dopo l'esportazione riuscita dell'MP3. Il self-test espone `persistent_character_catalog_seed: true`.

## Esclusione dei moduli di test Google

Il file `audio_description_bridge.spec` raccoglie separatamente dati, librerie e moduli runtime di `google.genai`. I pacchetti `google.genai.tests`, `google.genai._test_api_client`, `pytest` e `_pytest` sono esclusi perché non servono al worker e causavano un avviso PyInstaller quando `pytest` non era installato.
