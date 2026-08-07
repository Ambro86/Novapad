# Verifica integrazione audiodescrizione v6

## Obiettivo della v6

La v6 aggiunge la gestione interattiva delle quote Gemini senza riavviare l'analisi e senza perdere i chunk già completati. Aggiunge inoltre la scelta accessibile per attivare o disattivare il riconoscimento nominale dei personaggi e il glossario.

## Comportamento verificato

1. Il worker distingue gli errori transitori generici dalle risposte di quota `429`/`RESOURCE_EXHAUSTED` che contengono riferimenti a quota o rate limit.
2. Timeout, rete, 5xx, sovraccarico e altri errori temporanei continuano a essere ritentati ogni 5 secondi.
3. Alla prima quota esaurita per un modello nella richiesta corrente, il worker emette `QUOTA:{...}` e si blocca su stdin senza terminare il processo.
4. Sonarpad mostra una domanda accessibile con tre scelte: provare un altro modello, continuare ad attendere o interrompere.
5. Se viene scelto un altro modello, Sonarpad propone un identificatore alternativo dall'elenco già caricato e permette di modificarlo mediante una finestra di testo accessibile.
6. La risposta Rust viene inviata al worker come JSON line-oriented: `switch`, `wait` oppure `stop`.
7. Con `switch`, il worker valida il modello, ripete soltanto la richiesta Gemini corrente e aggiorna l'override in memoria per tutti i chunk successivi.
8. Con `wait`, il modello non cambia e riprendono i retry automatici.
9. Con `stop`, il worker restituisce `cancelled=true`; Sonarpad non prosegue con TTS, ducking o esportazione.
10. Il modello effettivamente usato al termine viene restituito dal worker e salvato nel progetto JSON, invece di registrare necessariamente quello selezionato all'avvio.
11. La preferenza `audio_description_gemini_model` viene aggiornata immediatamente quando l'utente sceglie un modello alternativo.
12. La cache del worker passa a `audio_description_bridge_v6.exe`, impedendo il riuso del vecchio eseguibile privo del protocollo bidirezionale.

## Riconoscimento dei personaggi

1. La casella **Prova a riconoscere i personaggi e usa i loro nomi** è selezionata per impostazione predefinita.
2. Il valore viene trasmesso da `audio_description_window.rs` a `AudioDescriptionJob`, quindi a `AudioDescriptionBridgeRequest` e infine a `enable_character_glossary` nel worker.
3. Con la casella attiva, Gemini crea il glossario, riusa i nomi già riconosciuti nei chunk successivi e applica la pulizia delle ripetizioni nominali.
4. Con la casella disattivata, il prompt impone `character_glossary: []`, vieta di identificare o nominare i personaggi, non trasferisce identità tra chunk e salta la pulizia basata sui nomi.
5. La scelta effettiva viene salvata nel progetto JSON. I progetti v4/v5 privi del nuovo campo vengono letti come `true` grazie al default Serde compatibile.
6. La finestra inizializza la casella con `BST_CHECKED`; la scelta non viene imposta globalmente alle altre funzioni Gemini.

## Verifiche automatiche eseguite

1. `python3 -m py_compile` sul worker e su `gemini_helpers.py`: superato.
2. `audio_description_bridge.py --self-test`: superato; il risultato include `interactive_quota_decisions=true` e `optional_character_glossary=true`.
3. Test diretto del gestore quota con stdin simulato: `switch` restituisce il nuovo modello, `wait` restituisce attesa e `stop` restituisce annullamento.
4. Test simulato di `generate_content_with_retry`: quota sul modello A, scelta del modello B, validazione di B, ripetizione della sola richiesta corrente e persistenza di B per le richieste successive: superato.
5. Test diretto dei prompt: con riconoscimento attivo sono presenti schema e continuità del glossario; con riconoscimento disattivo il glossario è imposto vuoto e i nomi sono vietati: superato.
6. Parsing JSON di tutti i 17 file di traduzione: superato.
7. Verifica della presenza delle quattro nuove chiavi i18n, inclusa `audio_description.recognize_characters`, in tutte le lingue: superata.
8. Verifica statica che ogni costruzione di `AudioDescriptionCallbacks` inizializzi il nuovo campo `quota`: superata.
9. Verifica statica del protocollo Rust: stdin del worker è `piped`, gli eventi `QUOTA` vengono deserializzati e le decisioni vengono scritte e flushate sul processo figlio.
10. Verifica lessicale del bilanciamento di parentesi e delimitatori nei quattro file Rust modificati: superata.
11. Verifica che il worker non importi né esegua `ffmpeg.exe`, `ffprobe.exe`, TTS, SAPI di Omni, VLC o `accessible-output2`: superata.
12. Applicazione della patch v5→v6 a una copia pulita della v5 e confronto ricorsivo con il sorgente v6: nessuna differenza.

## Limite della verifica

In questo ambiente non sono presenti `cargo`, `rustc`, Windows, le DLL FFmpeg di Sonarpad e i backend TTS Windows. Non è stato possibile eseguire `cargo check` o simulare una finestra Win32 reale. Prima della pubblicazione occorre compilare su Windows e provare almeno questi casi: errore temporaneo 503, quota sul modello iniziale con cambio modello, scelta “continua ad attendere”, scelta “interrompi” e salvataggio del progetto con il modello effettivamente usato.
