# Test dell'audiodescrizione portati da Omni Describer

Sonarpad v30 documenta **154 test dedicati all'audiodescrizione**:

1. **125 test Python**, eseguiti senza rete e senza chiave Gemini.
2. **29 test Rust** con prefisso `omni_port_` o dedicati alla localizzazione e all'integrazione Win32.

## Copertura principale

1. Chunk Gemini da 180 secondi, timestamp assoluti e relativi, retry, quota e contenuti bloccati.
2. Pyannote ONNX incluso, intervalli di dialogo, slot disponibili, spostamenti e pause estese.
3. Correzione linguistica delle descrizioni e del glossario in tutte le 17 lingue di Sonarpad.
4. TTS Edge, Google e SAPI, misura della durata reale, pulizia del silenzio finale e scheduling.
5. Esportazione MP3 con librerie FFmpeg Rust, ducking, pause estese e progetto JSON transazionale.
6. Apertura contestuale da YouTube, video locali, streaming yt-dlp, RaiPlay e La7.
7. Preferenze persistenti del modulo e cataloghi dei personaggi riutilizzabili tra episodi.
8. Focus, Tab, riapertura e isolamento della finestra di modifica progetto.

## Test specifici aggiunti nella v24

1. **Spazio** e **Invio** sulla lista richiamano l'anteprima della descrizione selezionata.
2. L'anteprima usa l'MP3 audiodescritto esistente e i tempi `output_start_sec`/`output_end_sec`.
3. L'anteprima usa BASS in modo invisibile e non richiama la sintesi TTS né il player principale.
4. Il menu contestuale contiene **Riproduci descrizione** ed **Elimina descrizione**.
5. L'eliminazione rimuove la descrizione selezionata e salva immediatamente il JSON.
6. L'unica descrizione rimasta non può essere eliminata.
7. **Applica modifica** sintetizza prima il testo candidato, misura la durata e salva soltanto dopo il controllo.
8. Una frase troppo lunga restituisce `TooLong` e lascia invariato il progetto precedente.
9. Per le pause estese il controllo non impone un limite fisso.
10. Il passaggio a un'altra descrizione non salva silenziosamente il testo non applicato.
11. L'esportazione viene bloccata se esiste una modifica non ancora applicata.

## Esecuzione su Windows

Dalla cartella principale:

```powershell
.\run_audio_description_tests.ps1 `
  -Python py `
  -PythonVersion 3.14 `
  -InstallPythonTestDependencies
```

Il comando esegue i 125 test Python e poi:

```powershell
cargo test omni_port_ -- --nocapture
```

Per i soli test Python:

```powershell
py -3.14 .\bridge\run_audio_description_tests.py
```

I test non richiedono film reali, `ffmpeg.exe`, `ffprobe.exe`, wxPython, VLC o dispositivi audio. I test Python sono stati eseguiti nel pacchetto v27; i test Rust devono essere confermati nell'ambiente Windows del progetto.

## Verifica compilazione v25

La v25 non aggiunge nuovi casi funzionali. Corregge la compilazione derivando `Debug` per `Language` e `TtsEngine`. La verifica richiesta su Windows è `cargo clippy -- -D warnings`; il worker e i 119 test Python della v24 non cambiano.


## Verifica Clippy v26

La v26 non aggiunge nuovi casi funzionali. Corregge due lint Rust 1.97 in `src/audio_description.rs`: `manual_pattern_char_comparison` e `unnecessary_sort_by`. I 119 test Python e i 29 test Rust restano invariati; la verifica finale richiesta su Windows è `cargo clippy -- -D warnings`.


## Test specifico aggiunto nella v27

1. L'errore di descrizione troppo lunga usa la finestra progetto come proprietaria del messaggio.
2. Dopo **OK**, la finestra progetto viene riportata in primo piano.
3. Il focus torna alla lista delle descrizioni e non all'editor principale.
4. Il medesimo recupero viene verificato anche per gli altri errori di **Applica modifica**.

La suite Python v27 contiene **120 test**; i 29 test Rust restano invariati, per un totale documentato di **149 test**.

## Regressione access violation corretta nella v28

Il test del ritorno del focus dopo un errore di **Applica modifica** verifica ora anche che:

1. il ramo `TooLong` richiami `show_project_error` e non la funzione globale `show_error`;
2. il messaggio venga creato direttamente con `MessageBoxW`;
3. il watchdog venga sospeso e riattivato correttamente durante il messaggio;
4. prima del recupero del focus venga controllata la validità della finestra;
5. il helper locale non usi `show_blocking_modal_message_box` e quindi non interpreti lo stato Win32 della finestra progetto come `AppState`.

La suite v28 mantiene **120 test Python** e **29 test Rust**, per un totale documentato di **149 test**.


## Test specifico aggiunto nella v29

1. Il messaggio di esportazione riuscita usa `show_project_info(hwnd, ...)` e non `show_info(state.parent, ...)`.
2. Il messaggio informativo è creato direttamente con `MessageBoxW` e `MB_ICONINFORMATION`.
3. Dopo **OK** vengono controllate validità e visibilità della finestra progetto.
4. La finestra progetto viene riportata in primo piano e il focus torna alla lista delle descrizioni.
5. Gli errori conclusivi dell'esportazione usano lo stesso percorso locale sicuro.

La suite v29 contiene **121 test Python** e **29 test Rust**, per un totale documentato di **150 test**.

## Test specifici aggiunti nella v30

1. Spazio, Invio e menu contestuale scelgono l’MP3 esistente soltanto quando il testo coincide con quello realmente esportato.
2. Un testo non applicato viene sintetizzato come anteprima temporanea senza salvare il progetto.
3. Un testo applicato ma non ancora riesportato viene riconosciuto tramite `rendered_text` e continua a usare l’anteprima sintetizzata, anche dopo la riapertura del JSON.
4. L’anteprima normale apre l’audio originale da `source_start_sec`, applica il ducking del progetto e sovrappone la nuova voce.
5. Le pause estese riproducono soltanto la nuova voce.
6. Le anteprime superate vengono annullate e i file temporanei vengono rimossi.
7. Il file `.spec` filtra i pacchetti di test Google prima della raccolta e non usa più `collect_all("google.genai")`.
8. `pytest`, `_pytest`, `google.genai.tests` e `google.genai._test_api_client` sono esclusi dall’analisi PyInstaller.

La suite v30 contiene **125 test Python**. I **29 test Rust** restano invariati, per un totale documentato di **154 test**. I test Python sono stati eseguiti nel pacchetto v30; Cargo, Clippy e la costruzione PyInstaller devono essere confermati su Windows.


## Test specifico aggiunto nella v31

1. L'anteprima modificata usa `BassOutput::start_with_ffmpeg_at` con `source_start_sec` completo.
2. Non viene più richiamato `seek_to_seconds(source_start)` dopo l'apertura del flusso FFmpeg.
3. Il backend FFmpeg espone `FfmpegSource::try_new_at` e conserva i secondi frazionari tramite `Duration::from_secs_f64`.
4. Le API precedenti in secondi interi restano disponibili come wrapper compatibili.

La suite v31 contiene **126 test Python**. I **29 test Rust** restano invariati, per un totale documentato di **155 test**. I test Python sono stati eseguiti nel pacchetto v31; Cargo e Clippy devono essere confermati su Windows.
