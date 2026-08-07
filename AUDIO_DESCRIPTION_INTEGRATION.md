# Crea audiodescrizione con IA: architettura e distribuzione

### Scorciatoie nel player

`Ctrl+I` annuncia esclusivamente il tempo corrente e totale. Il comando del player richiede che Shift e Alt non siano premuti, quindi `Ctrl+Shift+I` non viene intercettato come Annuncia tempo e raggiunge sempre **Crea audiodescrizione con IA**.

## Obiettivo

Il nuovo comando **Strumenti > Multimedia > Crea audiodescrizione con IA...** integra in Sonarpad il flusso di Omni Describer senza duplicare i sistemi di sintesi e di esportazione già presenti nel programma.

La pipeline è divisa deliberatamente in due componenti indipendenti:

1. **Worker Python scaricabile**: analisi del WAV già decodificato da Sonarpad, rilevamento delle zone con parlato tramite il modello Pyannote Segmentation ONNX e generazione delle descrizioni sui segmenti video forniti dall’host.
2. **Host Rust di Sonarpad**: sintesi con il motore e la voce scelti dall'utente, controllo della durata reale, collocazione definitiva, ducking e codifica MP3 tramite le librerie FFmpeg dinamiche già incluse in Sonarpad.

Pyannote è l'unico componente che stabilisce gli intervalli di parlato protetti. FFmpeg non viene usato come rilevatore di silenzio.

## Impostazioni predefinite

La finestra usa questi valori iniziali:

1. Lingua dell'audiodescrizione uguale alla lingua dell'interfaccia di Sonarpad.
2. Brevità **Dettagliata**.
3. Analisi Gemini intensiva.
4. Segmenti Gemini con obiettivo di 180 secondi, non esposti come opzione; un file già entro 180 secondi viene passato direttamente senza crearne una copia.
5. Soglia intensiva minima di 3 secondi.
6. Protezione dei dialoghi attiva.
7. Casella **Prova a riconoscere i personaggi e usa i loro nomi** attiva. Se viene disattivata, Gemini restituisce un glossario vuoto, non mantiene identità nominali tra i chunk e usa riferimenti generici.
8. Verifica temporale dei chunk attiva.
9. Filtri di sicurezza Gemini impostati a `BLOCK_NONE`/`OFF` quando supportato dall'SDK; restano applicabili i blocchi non disattivabili del servizio.
10. Temperatura Gemini 0,3.
11. Casella **Attiva pause estese: interrompi il film in casi eccezionali...** attiva; l'utente può disattivarla per impedire qualsiasi interruzione della timeline originale.
12. Chiave API Gemini precaricata dalla stessa impostazione protetta usata da **AI e trascrizione**; una modifica effettuata nella finestra viene salvata nelle impostazioni di Sonarpad quando si avvia il lavoro o si aggiorna l'elenco dei modelli.
13. Pulsante **Ottieni chiave API...** che apre Google AI Studio.
14. Modello dell'audiodescrizione predefinito `gemini-3.5-flash-lite`, modificabile mediante una casella combinata e indipendente dal modello Gemini generale di Sonarpad.
15. Pulsante **Aggiorna elenco modelli** che interroga l'API Gemini con la chiave inserita e mostra i modelli compatibili con `generateContent`.
16. Motore TTS **Microsoft Edge**.
17. Per l'italiano, preferenza per `it-IT-IsabellaNeural` quando disponibile.
18. Ducking a -15 dB con dissolvenze di 150 ms.
19. Cartella di uscita predefinita `Documenti\Sonarpad\Audiodescriptions`, modificabile dalla scheda **Audio** delle impostazioni scegliendo **Audiodescrizioni** nella casella combinata delle cartelle predefinite.

La preferenza del modello viene salvata in `audio_description_gemini_model`; il worker riceve sempre il valore scelto e usa `gemini-3.5-flash-lite` anche come fallback se il campo ricevuto fosse vuoto. La chiave API resta condivisa con il resto delle funzioni Gemini di Sonarpad, mentre la scelta del modello non modifica traduzione, riassunto o trascrizione.

## Cartella di salvataggio

Quando viene selezionato un film, Sonarpad propone automaticamente un nome come `NomeFilm_audiodescritto.mp3` nella cartella configurata per le audiodescrizioni. Il valore iniziale è `Documenti\Sonarpad\Audiodescriptions`.

Se la casella **Salva anche il progetto per modifiche future** è attiva, il file `NomeFilm_audiodescritto.sonarpad-ad.json` viene salvato accanto all’MP3 soltanto dopo il completamento riuscito dell’esportazione. Cambiando la cartella **Audiodescrizioni** nella scheda **Audio** delle impostazioni, cambia quindi sia la destinazione proposta per l’MP3 sia quella del progetto. La finestra **Modifica progetto audiodescrizione** si apre inizialmente nella stessa cartella.

L’utente può comunque scegliere manualmente un’altra destinazione con il pulsante **Sfoglia**; in quel caso il progetto resta accanto all’MP3 scelto.

## Creazione diretta dal video in riproduzione

Quando il player interno sta riproducendo un video YouTube oppure un file video locale, il menu **Riproduci** mostra **Crea audiodescrizione con IA...** (`Ctrl+Shift+I`). La stessa scorciatoia è mostrata anche in **Strumenti > Multimedia**. Il comando è assente per file audio locali, radio, podcast, flussi HLS e dirette TV.

1. **Video YouTube**: Sonarpad riutilizza l'URL del video già aperto e, quando disponibile, anche il contesto yt-dlp già registrato; per gli altri ingressi YouTube ricostruisce automaticamente il contesto usando yt-dlp.
2. yt-dlp scarica un singolo flusso già combinato con immagini e audio, senza richiedere un `ffmpeg.exe` esterno per il merge.
3. Il file YouTube viene salvato automaticamente nella cartella `media_save_folder` configurata nelle impostazioni. Se il valore è vuoto viene usata la cartella predefinita `Documenti\Sonarpad\Media`.
4. Non viene mostrata la finestra **Salva con nome**. In caso di nome già esistente, Sonarpad crea `Titolo (2).webm`, `Titolo (3).webm` e così via, senza sovrascrivere.
5. **Video locale**: non viene eseguito alcun download. Sonarpad apre direttamente **Crea audiodescrizione con IA** con il percorso del file locale già inserito nel campo sorgente.
6. In entrambi i casi il percorso MP3 viene proposto automaticamente nella cartella Audiodescrizioni.
7. La voce del menu Riproduci riconosce soltanto estensioni video supportate, tra cui MP4, MKV, AVI, MOV, M4V e WebM; non compare per MP3, M4A, WAV o altri file solo audio.
8. Premendo **Ctrl+Shift+I** durante la riproduzione, Sonarpad usa automaticamente il contesto attivo: file locale, YouTube oppure finestra vuota quando non è in riproduzione un video.
9. Il video YouTube può continuare a essere riprodotto durante il download; il normale pannello di avanzamento yt-dlp rimane annullabile.

## Flusso operativo

1. Sonarpad legge durata e tracce direttamente tramite le DLL FFmpeg già caricate dal codice Rust.
2. Sonarpad decodifica il soundtrack in un WAV PCM mono 16 kHz; per i film oltre 180 secondi crea i segmenti video tramite il backend FFmpeg Rust, mentre per i file più brevi usa direttamente il media originale già verificato.
3. Il worker riceve il WAV e i percorsi dei chunk già preparati; non cerca e non avvia `ffmpeg.exe` o `ffprobe.exe`.
4. Pyannote ONNX analizza il WAV e restituisce gli intervalli con parlato, che diventano zone protette.
5. Gemini analizza i chunk fisici con obiettivo di 180 secondi e modalità intensiva. Se il riconoscimento dei personaggi è attivo, crea il glossario e mantiene la continuità nominale tra segmenti; se è disattivo, non crea il glossario e usa riferimenti generici.
6. Il worker restituisce a Rust soltanto timestamp, testi, intervalli Pyannote e metadati di analisi. Non sintetizza e non esporta audio.
7. Sonarpad sintetizza ogni testo con il motore e la voce selezionati, riutilizzando `tts_engine.rs`.
8. Per Edge viene applicata in PCM la stessa regola di pulizia della coda usata da Omni: soglia massima tra -55 dBFS e media meno 35 dB, finestra 60 ms, passo 5 ms, conservazione di 30 ms e rimozione soltanto da 60 ms in su.
9. La durata reale di ogni file TTS viene ricontrollata contro gli intervalli liberi da parlato prodotti da Pyannote.
10. Se una descrizione non entra, viene spostata al massimo di 5 secondi. Se ancora non entra, viene inserita come pausa estesa soltanto quando la casella relativa è attiva; altrimenti viene scartata.
11. `ffmpeg_export.rs` decodifica l'audio originale, applica il ducking, inserisce le pause estese e crea un WAV temporaneo PCM.
12. Il WAV viene codificato in MP3 tramite `convert_audio_file`, quindi tramite le DLL FFmpeg già caricate dinamicamente da Sonarpad. Non viene avviato `ffmpeg.exe` per il montaggio o l'esportazione finale.


## Lingua dei prompt, controllo e localizzazione

La lingua scelta nella finestra viene passata al worker come codice BCP 47 e governa l’intero contenuto naturale restituito da Gemini. La tabella interna contiene una voce esplicita per tutte le 17 lingue di Sonarpad: ceco, cinese, francese, hindi, inglese, italiano, lituano, polacco, portoghese brasiliano, portoghese europeo, russo, serbo, spagnolo, svedese, tedesco, ucraino e vietnamita. `pt-BR` viene mantenuto distinto da `pt` nel prompt, mentre il rilevatore può usare il codice base quando necessario.

Il prompt impone che siano nella lingua selezionata sia `description_text` sia la descrizione fisica contenuta in `character_glossary[].description`. I nomi propri, gli identificatori JSON e i timestamp non vengono tradotti. Anche l’esempio JSON mostrato a Gemini è scritto nella lingua scelta, per ridurre il rischio che il modello torni all’inglese.

Dopo ogni risposta, il controllo linguistico derivato da Omni verifica separatamente le descrizioni. Se il glossario è attivo, verifica separatamente anche ogni descrizione fisica dei personaggi. Solo gli elementi rilevati nella lingua sbagliata vengono inviati a una richiesta correttiva; identificatori, nomi, ordine e timestamp restano invariati. Se il rilevamento non è affidabile o la correzione fallisce, il testo originale viene conservato invece di interrompere l’intero lavoro.

L’interfaccia comprende 86 chiavi dedicate al modulo in ciascuno dei 17 file linguistici. Il worker non espone più direttamente frasi inglesi di avanzamento: invia identificatori stabili come `pyannote_done`, `gemini_chunk` o `language_correction`, e Sonarpad li traduce nella lingua dell’interfaccia preservando valori dinamici come numero del chunk e conteggio degli intervalli.

## Retry Gemini e cambio modello su quota esaurita

Il worker conserva i retry automatici illimitati ogni 5 secondi per timeout, errori di rete, risposte 5xx, sovraccarico temporaneo, `PROHIBITED_CONTENT` transitorio e fallimenti temporanei di elaborazione dei file Gemini. Una richiesta bloccata ha inoltre un timeout di 8 minuti, dopo il quale viene riprovata.

Una vera risposta di quota esaurita (`429`/`RESOURCE_EXHAUSTED` accompagnata da un riferimento a quota o rate limit) segue invece un protocollo interattivo: il worker emette un evento `QUOTA` e rimane in attesa sullo stesso chunk. Sonarpad presenta tre possibilità:

1. **Sì – prova un altro modello**: viene proposto un modello alternativo tra quelli già caricati; l'utente può modificare l'identificatore. Il worker valida il nuovo modello e ripete soltanto la richiesta corrente.
2. **No – continua ad attendere**: il worker mantiene il modello corrente e continua i retry automatici.
3. **Annulla – interrompi**: il worker termina come operazione annullata, senza passare alla sintesi o creare un MP3 parziale.

Il processo Python non viene riavviato durante la scelta: intervalli Pyannote, l’eventuale glossario, descrizioni dei chunk precedenti e file caricati restano in memoria. Dopo un cambio riuscito, il nuovo modello viene usato anche dai chunk successivi, viene salvato nella preferenza dell'audiodescrizione di Sonarpad e, se richiesto, nel JSON del progetto finale.

Il canale è bidirezionale e line-oriented: stdout trasporta `STATUS`, `PROGRESS`, `QUOTA` e `RESULT`; stdin riceve una singola decisione JSON (`switch`, `wait` o `stop`) per ogni evento di quota.

## Perché il worker Pyannote è separato da Faster Whisper

Il nuovo worker usa lo stesso modello di distribuzione di Faster Whisper, ma non lo stesso processo o eseguibile. In questo modo:

1. Aggiornamenti di Pyannote/ONNX Runtime non interferiscono con Whisper.
2. Un errore dell'audiodescrizione non interrompe dettatura o trascrizione.
3. I pacchetti e i modelli non necessari non vengono caricati in memoria insieme.
4. Sonarpad può aggiornare separatamente `audio_description_bridge_v1.exe`.

## Costruzione del worker su Windows

Dalla cartella `bridge`:

```powershell
.\build_audio_description_bridge.ps1 `
  -Python py `
  -PythonVersion 3.14
```

Non servono `ffmpeg.exe` o `ffprobe.exe`, né durante la costruzione né sul computer dell'utente. Il worker contiene soltanto Pyannote ONNX, ONNX Runtime e il client Gemini. Tutta la preparazione multimediale e l'esportazione vengono eseguite da Sonarpad attraverso le DLL FFmpeg già incluse.

L'output è:

```text
dll\audio_description_bridge.exe
```

Prima della pubblicazione eseguire:

```powershell
.\dll\audio_description_bridge.exe --self-test
```

Il risultato deve indicare:

```json
{"ok":true,"chunk_duration_sec":180,"exports_audio":false,"uses_external_ffmpeg":false,"expects_host_prepared_media":true,"contains_tts_or_playback":false,"interactive_quota_decisions":true,"optional_character_glossary":true}
```

## Pubblicazione per il download al primo utilizzo

Caricare `audio_description_bridge.exe` nella release `0.7` oppure `v0.7` del repository `Ambro86/Sonarpad-Tools`. Sonarpad prova entrambi gli indirizzi e salva il file in:

```text
%APPDATA%\Sonarpad\tools\audio_description_bridge_v1.exe
```

Il file scaricato deve essere un PE Windows valido e misurare almeno 5 MB. In build di debug Sonarpad preferisce `dll\audio_description_bridge.exe`.

## File principali modificati

1. `src/app_windows/audio_description_window.rs`: interfaccia accessibile, pause estese, riconoscimento personaggi, chiave/modello Gemini e selezione motore/voce.
2. `src/audio_description.rs`: orchestrazione Pyannote/Gemini, TTS, controllo durata e scheduling.
3. `src/tools/audio_description_bridge.rs`: download ed esecuzione del worker.
4. `src/ffmpeg_export.rs`: mix, ducking, pause estese e MP3 con FFmpeg Rust.
5. `bridge/audio_description_bridge.py`: worker headless derivato dal flusso Omni.
6. `bridge/audio_description_runtime/`: nucleo minimo Omni, modello Pyannote ONNX e licenze.
7. `src/settings.rs`: chiave Gemini condivisa e preferenza separata `audio_description_gemini_model`, con default `gemini-3.5-flash-lite`.
8. `src/app_windows/options_window.rs`: funzioni condivise per aprire Google AI Studio e recuperare l'elenco dei modelli Gemini.

## Separazione completa della sintesi vocale

Il worker non incorpora né importa i componenti di riproduzione e sintesi di Omni Describer. In particolare non sono presenti `accessible-output2`, `sapi32.py`, SAPI5, `pyttsx3`, `tts_engine.py`, `sound_player.py`, Edge TTS, VLC o moduli UI/player.

Il worker termina restituendo a Sonarpad testi, timestamp e intervalli Pyannote. Sonarpad usa quindi il motore TTS scelto nella propria interfaccia per generare i WAV temporanei, controlla la loro durata reale, elimina il silenzio finale Edge quando necessario e produce l’MP3 mediante il backend FFmpeg Rust. Il ducking predefinito di -15 dB è definito nel codice Rust e non arriva più dal worker.

La presenza di `src/sapi5_engine.rs` nel sorgente generale non deriva da Omni Describer: è il backend SAPI già utilizzato dalle altre funzioni di Sonarpad. Il nuovo worker non lo importa e non dipende da esso; potrà essere usato soltanto se l’utente lo seleziona esplicitamente nella combinazione dei motori di Sonarpad.

## Progetto modificabile dopo l'esportazione

La finestra include la casella **Salva anche il progetto per modifiche future**, disattivata per impostazione predefinita. Quando è attiva, il comportamento è deliberatamente successivo all'esportazione:

1. Gemini produce le descrizioni candidate.
2. Sonarpad sintetizza realmente ogni testo con il motore scelto e misura l'audio ottenuto dopo l'eventuale pulizia Edge.
3. Lo scheduler decide quali descrizioni entrano nei silenzi Pyannote, quali richiedono una pausa estesa e quali devono essere escluse.
4. Sonarpad crea e valida l'MP3 completo.
5. Soltanto dopo l'esportazione riuscita viene costruito il file `nomefilm.sonarpad-ad.json`.

L'elenco principale `descriptions` del JSON contiene quindi esclusivamente le descrizioni realmente presenti nell'ultimo MP3 esportato. Per ogni elemento vengono salvati il testo pronunciato, il tempo proposto da Gemini, la posizione effettiva nella timeline originale, la posizione iniziale e finale nella timeline dell'MP3, la durata TTS reale, l'eventuale pausa estesa e l'intervallo di ducking. Le descrizioni non inserite sono conservate separatamente in `excluded_descriptions` con il motivo dell'esclusione.

I timestamp dell'MP3 tengono conto delle pause estese già inserite: una descrizione successiva viene quindi salvata nella posizione in cui si ascolta davvero nel file finale, non soltanto nella posizione della scena originale.

### Coerenza tra MP3 e JSON

Quando viene richiesto il progetto, MP3 e JSON vengono prima creati come file temporanei. I file definitivi vengono sostituiti insieme soltanto quando entrambi sono validi. Se la scrittura del JSON o la sostituzione finale fallisce, Sonarpad ripristina la coppia precedente. In questo modo un progetto non può descrivere un MP3 diverso da quello associato.

### Modifica successiva

Il pulsante **Modifica progetto audiodescrizione...** apre un progetto già esportato e mostra soltanto le descrizioni presenti nell'ultimo MP3. L'utente può modificare il testo e premere **Applica modifica**. Sonarpad sintetizza soltanto la frase candidata con la voce del progetto, misura la durata reale e aggiorna immediatamente il JSON esclusivamente se la nuova voce entra nello spazio disponibile. Se è troppo lunga, il testo precedente resta invariato.

Con **Esporta nuovamente MP3**, Sonarpad:

1. sintetizza di nuovo i testi del progetto;
2. misura le nuove durate reali;
3. ripete il posizionamento sugli intervalli Pyannote già salvati;
4. ricrea ducking, pause estese e MP3 tramite le librerie FFmpeg Rust;
5. aggiorna il JSON soltanto dopo il successo, con le nuove posizioni effettive.

Se una descrizione modificata non entra più, non resta falsamente nell'elenco principale: viene spostata tra le descrizioni escluse. Il worker Gemini/Pyannote non viene richiamato durante questa riesportazione e non vengono consumati nuovamente token Gemini.

## Suite di test derivata da Omni Describer

La versione v30 documenta 154 test dedicati all'audiodescrizione: 125 test Python e 29 test Rust. Oltre agli 84 casi derivati da Omni Describer, la suite verifica protocollo quota, cartella di uscita, tutte le 17 lingue dei prompt, correzione linguistica del glossario, localizzazione completa e stati dinamici tradotti. La suite non usa la rete e non richiede una chiave Gemini, film reali, `ffmpeg.exe`, `ffprobe.exe`, wxPython o dispositivi audio. Le istruzioni complete sono in `AUDIO_DESCRIPTION_TESTS.md`.
## Compatibilità Clippy v10

La v10 corregge tutte le 73 diagnostiche riportate da `cargo clippy -- -D warnings` dopo la v9. Le correzioni non usano attributi `allow` e non cambiano il protocollo del worker, Pyannote, Gemini, TTS, ducking, cartelle di output o formato del progetto JSON.


## Interfaccia di elaborazione e ascolto immediato

Durante la generazione Sonarpad nasconde tutti i campi di configurazione della finestra. Restano esposti soltanto la barra di avanzamento, il messaggio corrente e il pulsante **Annulla**, così NVDA non incontra più file, modello, voce e altre opzioni mentre il lavoro è in corso. Al termine i controlli vengono ripristinati.

Dopo il messaggio di completamento, premendo **OK** l'MP3 appena creato viene aperto direttamente nel player audio interno di Sonarpad. La finestra **Crea audiodescrizione** viene temporaneamente nascosta. Premendo **Esc** nel player, Sonarpad arresta la riproduzione, chiude la scheda audio e riporta il focus alla finestra dell'audiodescrizione, sul pulsante di creazione.

## Stati del worker completamente localizzati

Il worker non inoltra più all'interfaccia le frasi inglesi prodotte internamente da Gemini. Caricamento del video, attesa dell'elaborazione, invio della richiesta, ricezione della risposta, riparazione JSON e retry temporanei sono convertiti in identificatori stabili. Sonarpad traduce questi identificatori in tutte le 17 lingue; anche eventuali nuovi stati con prefisso `gemini` ricadono su un messaggio Gemini localizzato invece di mostrare il testo inglese originale.

## Costruzione con Python Launcher

Lo script `bridge\build_audio_description_bridge.ps1` non usa più `@launcherArgs`. Quando viene scelto `py`, passa esplicitamente il selettore, per esempio `-3.14`, a pip e PyInstaller. In questo modo Windows PowerShell non avvia accidentalmente la console Python interattiva. La stessa correzione è applicata allo script di test del bridge.

## Correzione dei video YouTube salvati in formato automatico

Nelle versioni precedenti il formato **Automatico** usato da yt-dlp selezionava sempre `bestaudio/best`. Un file poteva quindi chiamarsi `.webm`, pur contenendo soltanto una traccia audio. Il modulo audiodescrizione lo rifiutava correttamente perché Gemini non aveva immagini da analizzare.

La v16 distingue il contesto:

1. quando il player di Sonarpad sta mostrando un video, il salvataggio automatico richiede un singolo flusso già multiplexato con video e audio;
2. non viene richiesto `ffmpeg.exe` a yt-dlp per unire due flussi separati;
3. trascrizione e streaming solo audio continuano a scaricare `bestaudio/best`;
4. la verifica FFmpeg accetta anche video WebM/Matroska validi che espongono larghezza e altezza soltanto all'apertura del decoder, purché la traccia non sia una semplice copertina allegata.

I vecchi WebM realmente audio-only non possono essere trasformati in video: devono essere scaricati nuovamente dopo la correzione oppure salvati esplicitamente come MP4.

## Preferenze del modulo ricordate

La finestra **Crea audiodescrizione con IA** conserva preferenze proprie, separate dalle impostazioni TTS generali di Sonarpad. Ogni modifica viene salvata immediatamente e ripristinata alla successiva apertura. Vengono ricordati:

1. lingua dell'audiodescrizione;
2. livello di dettaglio;
3. motore TTS e voce selezionata;
4. stato della casella per le pause estese;
5. stato della casella per il riconoscimento dei personaggi;
6. stato della casella per il salvataggio del progetto modificabile.

Per una configurazione precedente alla v20, la lingua continua inizialmente a seguire quella dell'interfaccia, Edge resta il motore predefinito, le pause estese e il riconoscimento dei personaggi restano attivi e il salvataggio del progetto resta disattivato. La voce dell'audiodescrizione non modifica la voce generale usata per leggere i documenti. Se la voce salvata non è più disponibile, Sonarpad sceglie una voce compatibile con la lingua selezionata e registra il nuovo valore.


## Avvio contestuale da streaming, RaiPlay e La7

La voce **Riproduci > Crea audiodescrizione con IA…** è ora disponibile anche nei tre contesti richiesti, senza estendere il comportamento ad altri servizi o alle dirette generiche:

1. un contenuto video aperto da **Riproduci audio da streaming**, quando il contesto yt-dlp registrato da Sonarpad indica che il player sta riproducendo video;
2. un contenuto on demand di RaiPlay;
3. un contenuto on demand di La7 Play.

Per lo streaming yt-dlp Sonarpad riusa il medesimo contesto di salvataggio creato all'apertura del link, forza un file con video e audio, lo salva nella cartella Media configurata e apre la finestra dell'audiodescrizione con il sorgente già compilato. Il ramo è separato da quello YouTube, che continua a usare il proprio recupero dedicato.

Per RaiPlay e La7 Sonarpad non introduce un nuovo downloader: richiama l'esportatore MP4 già usato dal comando di salvataggio dei due servizi. In questo percorso non viene chiesta una destinazione manuale; il file riceve un nome non sovrascrivente nella cartella Media e, soltanto dopo un'esportazione riuscita, viene passato alla finestra **Crea audiodescrizione con IA**. Le dirette RaiPlay e La7 restano escluse.


## Editor progetto: anteprima, applicazione verificata ed eliminazione

La v24 rende operative le azioni sulla singola descrizione senza aprire il player principale.

1. Quando il focus è sulla lista delle descrizioni, **Spazio** o **Invio** riproducono direttamente l’intervallo già presente nell’MP3 audiodescritto, usando `output_start_sec` e `output_end_sec` salvati nel progetto. Non viene eseguita alcuna sintesi TTS e non viene aperta una nuova scheda del player.
2. Il menu contestuale della descrizione selezionata contiene **Riproduci descrizione** ed **Elimina descrizione**.
3. L’eliminazione aggiorna e salva immediatamente il progetto JSON. L’MP3 esistente non viene modificato finché l’utente non sceglie **Esporta nuovamente MP3**. Non è consentito eliminare l’unica descrizione rimasta.
4. **Applica modifica** sintetizza soltanto il nuovo testo con motore, voce, velocità, tono e volume del progetto. La durata viene misurata dall’audio realmente generato.
5. Per una descrizione normale, la durata viene confrontata con lo spazio libero effettivo, limitato dalla fine del silenzio Pyannote e dall’inizio della descrizione successiva. Per una pausa estesa non esiste un limite fisso, perché la timeline può essere allungata.
6. Se la nuova voce supera il tempo disponibile, Sonarpad mostra durata ottenuta e durata disponibile, non salva il testo e mantiene integralmente la descrizione precedente.
7. Se la verifica riesce, il testo viene salvato subito nel JSON. L’audio dell’MP3 viene aggiornato soltanto con **Esporta nuovamente MP3**; fino a quel momento l’anteprima riproduce correttamente il tratto dell’MP3 esistente.
8. Se il campo contiene una modifica non ancora applicata, l’esportazione viene bloccata e Sonarpad chiede di applicarla prima.

## Correzione compilazione v25

La v25 aggiunge `Debug` agli enum `Language` e `TtsEngine`, richiesto dalla struttura `AudioDescriptionProject` introdotta nel flusso di modifica. La correzione risolve gli errori Rust `E0277` senza modificare serializzazione, impostazioni o comportamento del modulo.


## Compatibilità Clippy v26

La v26 corregge due diagnostiche Clippy segnalate da Rust 1.97 nel supporto ai cataloghi dei personaggi:

1. `trim_end_matches` usa direttamente il pattern `['.', ' ']` invece di una chiusura con confronto manuale dei caratteri;
2. l'ordinamento alfabetico dei cataloghi usa `sort_by_key` con il nome normalizzato in minuscolo.

Le modifiche non alterano nomi file, contenuto dei cataloghi, ordinamento visibile o comportamento del modulo. Non sono stati usati attributi `allow`.


## Ripristino del focus dopo errore di applicazione v27

Quando **Applica modifica** rifiuta una descrizione perché la sintesi supera il tempo disponibile, il messaggio di errore appartiene ora alla finestra **Modifica progetto audiodescrizione**. Dopo la conferma con **OK**, Sonarpad riporta esplicitamente quella finestra in primo piano e rimette il focus sulla lista delle descrizioni inserite. Lo stesso recupero viene applicato agli altri errori restituiti dal controllo di applicazione, evitando il ritorno accidentale all'editor principale.

## Correzione access violation nel messaggio di errore v28

La v27 passava la finestra **Modifica progetto audiodescrizione** alla funzione globale `show_error`. Quella funzione è progettata esclusivamente per la finestra principale di Sonarpad: registra lo stato modale tramite `with_state` e, dopo la chiusura del messaggio, tenta di ripristinare l'editor principale. Usata con la finestra progetto, interpretava quindi il relativo `GWLP_USERDATA` come se contenesse l'`AppState` principale, causando `STATUS_ACCESS_VIOLATION` subito dopo **OK**.

La v28 introduce `show_project_error`, dedicata alla finestra progetto:

1. usa direttamente `MessageBoxW` con la finestra progetto come proprietaria;
2. informa il watchdog dell'ingresso e dell'uscita dal messaggio modale;
3. non chiama `with_state`, `show_error` o `show_blocking_modal_message_box` con una maniglia secondaria;
4. dopo **OK** verifica che la finestra progetto esista ancora e sia visibile;
5. solo dopo tali verifiche la riporta in primo piano e rimette il focus sulla lista delle descrizioni.

La verifica di durata, il testo precedente, il progetto JSON e la sintesi TTS non cambiano.


## Ritorno del focus dopo esportazione v29

Al termine di **Esporta nuovamente MP3**, il messaggio **MP3 e progetto aggiornati correttamente** appartiene ora alla finestra **Modifica progetto audiodescrizione**. Dopo **OK**, Sonarpad verifica che la finestra sia ancora valida e visibile, la riporta in primo piano e rimette il focus sulla lista delle descrizioni.

Il ramo di completamento non usa più `show_info(state.parent, ...)`, che restituiva il focus alla finestra principale e quindi all'editor. Anche gli errori conclusivi dell'esportazione usano il messaggio locale sicuro della finestra progetto. Sintesi, progetto JSON, MP3 e worker non cambiano.

## Anteprima del testo modificato e worker pulito v30

La v30 distingue lo stato realmente contenuto nell’MP3 dal testo presente nell’editor del progetto.

1. Se il testo non è stato modificato e coincide con quello già esportato, **Spazio**, **Invio** e **Riproduci descrizione** continuano a leggere esattamente l’intervallo `output_start_sec`–`output_end_sec` dell’MP3 audiodescritto esistente.
2. Se il campo contiene un testo nuovo, Sonarpad lo sintetizza in una cartella temporanea senza salvare il JSON e senza modificare l’MP3.
3. Per una descrizione normale, la nuova voce viene riprodotta insieme all’audio originale a partire da `source_start_sec`; l’audio originale usa lo stesso valore di ducking del progetto, così l’anteprima è vicina al mix finale.
4. Per una descrizione con pausa estesa viene riprodotta soltanto la voce, perché nel file finale la timeline originale viene sospesa e una sovrapposizione farebbe ascoltare un risultato non rappresentativo.
5. Se **Applica modifica** è già stato premuto ma l’MP3 non è ancora stato riesportato, il campo `rendered_text` conserva quale testo è realmente presente nell’MP3. L’anteprima continua quindi a sintetizzare il testo nuovo anche dopo la riapertura del progetto. I vecchi progetti privi del campo vengono caricati assumendo che il testo corrente sia quello già esportato.
6. Una seconda anteprima annulla la sintesi precedente. I file WAV temporanei vengono eliminati al termine o quando l’anteprima viene scartata.
7. Se l’audio originale non può essere aperto o posizionato, Sonarpad riproduce comunque la nuova voce da sola invece di fallire completamente.

Il file `bridge/audio_description_bridge.spec` non usa più `collect_all("google.genai")`, che tentava di importare anche `google.genai.tests` e produceva l’avviso relativo a `pytest`. Dati, librerie e moduli runtime vengono raccolti separatamente; `google.genai.tests`, `google.genai._test_api_client`, `pytest` e `_pytest` vengono filtrati ed esclusi. Per applicare questa correzione al worker è necessario ricostruire `dll/audio_description_bridge.exe`.


## Correzione sovrapposizione anteprima modificata v31

La v30 apriva correttamente il flusso originale con FFmpeg al tempo della descrizione, ma tentava subito dopo un secondo posizionamento tramite `BASS_ChannelSetPosition`. I flussi BASS alimentati dal callback FFmpeg non supportano quel seek e restituivano l'errore BASS 27; il codice interpretava il risultato come apertura fallita e riproduceva soltanto la voce modificata.

La v31 introduce un avvio FFmpeg preciso in secondi frazionari:

1. `FfmpegSource::try_new_at` accetta un tempo `f64` e usa `Duration::from_secs_f64`;
2. `FfmpegBassStream::new` mantiene il tempo preciso senza arrotondarlo al secondo intero;
3. `BassOutput::start_with_ffmpeg_at` apre direttamente il flusso originale nella posizione richiesta;
4. l'anteprima modificata non chiama più `seek_to_seconds` sul canale BASS personalizzato;
5. audio originale abbassato e voce temporanea vengono avviati insieme, mentre il ripiego alla sola voce resta disponibile soltanto se l'apertura FFmpeg fallisce realmente.

Le API esistenti basate su secondi interi restano compatibili e delegano alla nuova variante precisa. Il progetto JSON, il worker Python e il file `.spec` non cambiano rispetto alla v30.
