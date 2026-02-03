# Changelog

Versione 0.6.4 – 2026-01-30
Miglioramenti
• Il programma e' stato rinominato in Sonarpad per dare maggiore enfasi a suono e audio, che sono la chiave di questo programma.
• Aggiunta la selezione delle tracce audio nel menu Riproduzione per i file multimediali con più tracce audio (es. MKV con più lingue).
• I podcast ora indicano chiaramente quelli non ascoltati con il prefisso "Non ascoltato" prima del nome.
• Nuovo sistema di tag per cambiare voce nel testo. Esempi:
  - Voci Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Ciao</voice>
  - Voci SAPI5: <voice sapi5 Microsoft Helena Desktop>Ciao</voice>
  - Voci SAPI4: <voice sapi4 #1>Ciao</voice>
• Arricchite le categorie dei podcast.
• Migliorata la lettura dei PDF grazie al fallback automatico su PDFium.
• Migliorato il parser degli articoli che in alcuni casi non venivano letti in modo integrale.
• Aggiunto il reset del pitch nel menu Riproduci.
• Aggiunta un'opzione nel menu contestuale per creare un audiolibro dalla selezione.
• Aggiunta un'opzione per usare il nome legacy "Novapad" nel titolo della finestra e nei collegamenti del menu Avvio.
Correzioni di bug
• Corretto un bug per cui gli audiolibri con SAPI4 potevano essere creati in modo diverso da quanto previsto.
• Corretto un bug per cui, andando oltre la fine con il seek, la riproduzione ripartiva dall'inizio.

Versione 0.6.3 – 2026-01-30
Miglioramenti
• Migliorata la rilevazione del microfono.
• Aggiunta la riproduzione istantanea per tutti i formati.
Correzioni
• Corretto il crash nella finestra delle categorie podcast.

Versione 0.6.2 – 2026-01-30
Nuove funzionalità
• Aggiunta l'esecuzione dei file (Shift+F5). È possibile scegliere l'interprete (es. python) nelle Opzioni, cercarlo nel computer, e premendo Shift+F5 viene eseguito lo script corrente. I file HTML si aprono nel browser.
• Aggiunto il supporto per i file puntatori di Google Docs (.gdoc, .gsheet, .gslides), che si aprono automaticamente nel browser predefinito.
• Aggiunto il supporto per il formato audiolibro M4B (Apple/AAC).
• Aggiunta l'opzione "Mostra episodi" nel menu contestuale dei risultati di ricerca podcast per sfogliare e riprodurre episodi senza iscriversi.
• Aggiunta la funzione "Vai alla riga" (menu Modifica o Ctrl+J) per saltare rapidamente a un numero di riga specifico.
• Aggiunte opzioni nel menu contestuale per ordinare feed RSS e podcast (alfabeticamente o per data).
• Aggiunti feed RSS predefiniti in vietnamita.
• Aggiunta una casella di test microfono nella finestra di registrazione per verificare i livelli prima di iniziare.
• Aggiunta "Mostra descrizione" per gli episodi podcast nel menu contestuale.
• Aggiunto il supporto per formati audio/video estesi tramite FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Aggiunta la lettura sincronizzata dei sottotitoli (srt, vtt, ass, sub, sbv, lrc, smi) con NVDA o voce selezionata. Il programma cerca un file sottotitoli con lo stesso nome del file multimediale. Aggiunte le opzioni "Importa sottotitoli" e "Rimuovi sottotitoli" nel menu Riproduzione per file con nomi diversi.
• Aggiunte le associazioni file per tutti i nuovi formati audio/video supportati nel menu contestuale "Apri con Sonarpad".
• Aggiunta l'impostazione per regolare il pitch di qualsiasi file.
• Aggiunta nelle impostazioni Generali la casella per attivare o disattivare le segnalazioni di errore anonime. Aggiunta una voce nel menu Aiuto per creare un file ZIP diagnostico.
• Aggiunta l'opzione per usare una voce diversa per i dialoghi, sia per la lettura dal vivo che per la creazione di audiolibri.
• Aggiunto il browser delle categorie podcast per esplorare i podcast per categoria (business, arte, sport, ecc.).
Miglioramenti
• L'apertura di un file audio/video da Esplora risorse ora apre direttamente la vista player invece dell'editor di testo.
• Rimossa la richiesta OCR per i PDF non accessibili; l'OCR viene ora eseguito automaticamente per migliorare velocità ed esperienza utente.
• Migliorato il Terminale Accessibile: la lettura NVDA ora ricorda l'ultima riga letta per una migliore continuità.
• SAPI 4: La creazione di audiolibri è ora completamente parallelizzata e quasi istantanea. Aggiunta una richiesta per scegliere il numero di processi contemporanei.
• SAPI 4: Eliminato il collo di bottiglia WAV-MP3 convertendo i frammenti in parallelo durante la sintesi.
• SAPI 4: Migliorata la gestione degli errori e la pulizia automatica dei file temporanei.
• Finestra Trova: Rinominato "Regex" in "Espressione regolare" per chiarezza e aggiunte le traduzioni mancanti per le opzioni di ricerca.
• Audiolibri M4B: Migliore gestione dell'output; la divisione per parti/marcatori ora produce un singolo file M4B con metadati dei capitoli inclusi titolo e autore.
• Player: Corretta la precisione dei segnalibri e degli annunci del tempo quando la velocità di riproduzione non è 1.0x.
• Ripristinata la navigazione Ctrl+Tab e Ctrl+Shift+Tab nelle Opzioni.
• Aggiunta un'opzione nel menu Riproduzione per ripristinare istantaneamente la velocità Normale (1.0x).
• Aggiornate tutte le dipendenze alle ultime versioni per migliori prestazioni e stabilità.
• Integrato FFmpeg con caricamento dinamico delle DLL per garantire compatibilità senza bloccare l'avvio.
• Aggiornati i filtri di download podcast per includere i nuovi formati audio/video.
• Impedito a Ctrl+S di salvare file audio/video per evitare corruzione.
• Migliorata l'importazione delle trascrizioni YouTube rendendola più robusta e resiliente.
• Migliorata la robustezza della divisione in parti degli audiolibri, assicurando che nessun testo venga perso.
• L'installer è ora completamente multilingua, supportando Italiano, Inglese, Spagnolo, Portoghese, Svedese e Vietnamita in base alla lingua del sistema dell'utente. L'inglese è la lingua predefinita per i sistemi non supportati.
• Categorie podcast: premendo Invio su una categoria ora si conferma la selezione (equivalente al pulsante OK).
• Migliorato il sistema di rilevamento blocchi per evitare falsi positivi quando sono aperti dialoghi modali (messaggi di errore, "testo non trovato").
Correzioni
• Corretto un bug per cui il changelog non si apriva all'avvio.
• Corretto un bug per cui la richiesta OCR non appariva per i PDF non accessibili aperti da Esplora risorse.
• Corretto un bug all'avvio che poteva causare perdita di focus o chiusura delle finestre subito dopo l'apertura.
• Corretto un bug critico nella ricerca regex che impediva di trovare il testo, inclusi problemi con la "Ricerca circolare" e l'opzione "Il punto equivale a nuova riga" con le terminazioni di riga Windows.
Localizzazione
• Aggiunta la traduzione in polacco.
• Aggiunta la traduzione in francese.
• Aggiunta la traduzione in ceco (grazie a Radek Žalud e Jiri Holzinger).

Versione 0.6.1 – 2026-01-20
Correzioni
• Corretto un bug per cui, attivando “Visualizza le voci nell’editor” e riproducendo un podcast, la riproduzione veniva interrotta.
• Corretto un problema per cui alcuni podcast non potevano essere aggiunti tramite URL perché l’indirizzo veniva troncato.
• Corretto un bug per cui non era più possibile aggiungere URL normali nella funzione dei feed RSS.
• Corretto un problema per cui la lingua di Wikipedia veniva mostrata in più schede delle opzioni.
• Rimossa la creazione di alcuni file di debug che venivano generati anche in modalità release.
Miglioramenti
• Migliorato il supporto per le voci Microsoft, che ora vengono riprodotte utilizzando una modalità dedicata con un diverso user agent.
• Aggiunto il supporto per i file MP4.

Versione 0.6.0 – 2025-01-20
Nuove funzionalità
• Aggiunto il correttore ortografico. Dal menu contestuale è possibile verificare se la parola corrente è corretta e, in caso contrario, ottenere suggerimenti.
• Aggiunta l’importazione ed esportazione dei podcast tramite file OPML.
• Aggiunto il supporto alla ricerca Podcast Index oltre a iTunes. L’utente può inserire la propria API key e API secret gratuiti (generabili inserendo solo la propria email).
• Aggiunto il supporto alle voci SAPI4, sia per la lettura in tempo reale sia per la creazione di audiolibri
• Aggiunto il fallback automatico OCR per i PDF non accessibili: quando non viene trovato testo estraibile, il documento viene riconosciuto tramite OCR..
• Aggiunto il supporto al dizionario tramite Wiktionary. Premendo il tasto Applicazioni vengono mostrate le definizioni e, quando disponibili, anche sinonimi e traduzioni in altre lingue.
• Aggiunta l’importazione degli articoli da Wikipedia con ricerca, selezione dei risultati e importazione diretta nell’editor.
• Aggiunta la scorciatoia Shift+Invio nel modulo RSS per aprire un articolo direttamente nel sito web originale.
Miglioramenti
• La selezione del microfono ora viene sempre rispettata dall’applicazione.
• Nella finestra dei podcast, premendo Invio su un episodio NVDA annuncia immediatamente “caricamento”, dando subito conferma dell’operazione.
• Nei risultati di ricerca dei podcast, premendo Invio ora ci si sottoscrive al podcast selezionato.
• Corrette e migliorate le etichette delle scorciatoie Ctrl+Shift+O e Podcast Ctrl+Shift+P.
• La velocità di riproduzione e il volume ora vengono salvati nelle impostazioni e persistono per tutti i file audio.
• Aggiunta una cartella cache dedicata per gli episodi dei podcast. L’utente può conservare gli episodi tramite “Conserva podcast” nel menu Riproduci. La cache viene svuotata automaticamente quando supera la dimensione impostata dall’utente (Opzioni → Audio).
• Migliorato in modo significativo il recupero degli articoli RSS usando libcurl con impersonazione Chrome e iPhone, garantendo la compatibilità con circa il 99% dei siti.
• Aggiunto lo stato letto / non letto per gli articoli RSS, con indicazione chiara nella lista RSS.
• La funzione Sostituisci tutto ora mostra anche il numero di sostituzioni effettuate.
• Aggiunto il pulsante Elimina podcast quando si naviga la libreria dei podcast tramite Tab.
Correzioni
• Rimossa la voce ridondante “pending update” dal menu Aiuto (gli aggiornamenti sono già gestiti automaticamente).
• Corretto un bug per cui, aprendo un file MP3 e premendo Ctrl+S, il file veniva salvato e quindi corrotto.
• Corretto un problema nell’interfaccia in cui “Batch Audiobooks” veniva mostrato come “(B)… Ctrl+Shift+B” (rimossa l’etichetta ridondante).
• Corretto il funzionamento delle virgolette smart: quando abilitate, le virgolette normali vengono ora sostituite correttamente con quelle tipografiche.
• Corretto un bug per cui, usando “Vai al segnalibro”, la velocità di riproduzione veniva ripristinata a 1.0.
• Corretto un problema per cui gli episodi dei podcast già scaricati venivano riscaricati invece di usare la versione in cache.
Scorciatoie da tastiera
• F1 ora apre la guida.
• F2 ora controlla la presenza di aggiornamenti.
• F7 / F8 ora permettono di spostarsi all’errore ortografico precedente o successivo.
• F9 / F10 ora permettono di passare rapidamente tra le voci salvate nei preferiti.
Miglioramenti per sviluppatori
• Gli errori non vengono più ignorati silenziosamente: tutti i pattern let _ = sono stati rimossi e gli errori ora vengono gestiti esplicitamente (propagati, loggati o gestiti con fallback appropriati).
• Il progetto ora non compila in presenza di warning: sia cargo check sia cargo clippy devono completarsi senza avvisi, con lint più restrittivi e rimozione degli allow dove possibile.
• Rimosse le implementazioni personalizzate in stile strlen / wcslen. Le lunghezze delle stringhe e dei buffer UTF-16 ora derivano dai dati gestiti da Rust, senza scansioni manuali della memoria.
• La gestione delle DLL è stata ripulita e centralizzata attorno a libloading, evitando logiche di caricamento personalizzate e parsing PE.
• Rimossi gli helper artigianali per il parsing dei byte: ora tutto il parsing utilizza from_le_bytes / from_be_bytes su slice verificate.
Queste modifiche riducono l’uso superfluo di unsafe, eliminano potenziali comportamenti indefiniti e rendono il codice più idiomatico, robusto e manutenibile.

Versione 0.5.9 - 2025-01-13
Nuove funzionalita
• Aggiunta la possibilita di riordinare gli RSS dal menu contestuale (su/giu/posizione) con controlli per posizioni non valide.
• Aggiunto il menu contestuale anche per gli articoli, con apertura del sito originale e condivisione via WhatsApp, Facebook e X.
• Aggiunta la scorciatoia Esc per tornare rapidamente dagli articoli importati all'elenco RSS.
• Aggiunta la modalita podcast: ricerca, iscrizione e ascolto; riordinamento delle sottoscrizioni; Esc per fermare la riproduzione e tornare all'elenco; Invio su un episodio avvia la riproduzione.
• Aggiunta la regolazione della velocita di riproduzione per podcast e file MP3.
• Aggiunto Ctrl+T per andare a un tempo specifico.
• Aggiunto un pulsante di anteprima voci dopo la casella volume.
• Aggiunta la funzione regex per Trova e Sostituisci, stile Notepad++.
• Aggiunta l'importazione RSS da file OPML e TXT.
• Aggiunta nelle Opzioni la casella per abilitare "Apri con Sonarpad" in Esplora risorse, anche in versione portable.
• Aggiunto supporto OCR per PDF scansionati (richiede Windows 10/11): se un PDF non contiene testo, viene proposto il riconoscimento automatico.
Miglioramenti
• Migliorata la selezione di velocita, tono e volume delle voci, rispettando i limiti massimi del TTS.
• Vari miglioramenti alla modalita RSS per scaricare tutti gli articoli senza spostare il focus di NVDA durante gli aggiornamenti.
• Migliorata la riproduzione audio con un menu dedicato, annuncio tempo con Ctrl+I e volume fino al 300%.
• Aggiunte scorciatoie mancanti per alcune funzioni.
• Riordinato il menu Modifica con un sottomenu per le funzioni di pulizia testo.
• Riordinate le Opzioni in schede, con Ctrl+Tab e Ctrl+Shift+Tab per spostarsi tra le schede.
• Risolti i problemi di lettura degli articoli: il lettore RSS ora legge integralmente gli articoli come da browser.
Fix
• Corretto un problema per cui la pulizia Markdown eliminava i numeri a inizio riga.
• Corretto il problema AltGr+Z che attivava Undo.
• Corretto un problema per cui la registrazione di un audiolibro non si poteva interrompere rapidamente.
Localizzazione
• Aggiunta la traduzione vietnamita (grazie a Anh Đức Nguyễn).

Versione 0.5.8 - 2026-01-10
Nuove funzionalita
• Aggiunto il controllo volume per microfono e audio di sistema durante la registrazione podcast.
• Aggiunta una nuova funzione per importare articoli da siti web o feed RSS, includendo per ogni lingua i feed piu importanti.
• Aggiunta la funzione per rimuovere tutti i segnalibri del file corrente.
• Aggiunta la funzione per rimuovere le linee duplicate e le linee duplicate consecutive.
• Aggiunta la funzione per chiudere tutti i tab o le finestre tranne quella corrente.
• Inserita la voce Donazioni nel menu Aiuto per tutte le lingue.
Miglioramenti
• Migliorato il terminale accessibile evitando alcuni crash.
• Migliorati e sistemati access key e scorciatoie da tastiera del programma.
• Corretto un problema per cui chiudendo la finestra di riproduzione audio la riproduzione non si fermava.
• Aggiunte finestre di conferma per azioni importanti (es. rimozione linee duplicate, rimozione trattini a fine riga, rimozione di tutti i segnalibri del file corrente). Nessuna conferma se l'azione non si applica.
• Aggiunta la possibilita di eliminare feed/siti RSS dalla libreria selezionandoli e premendo Canc.
• Aggiunto un menu contestuale nella finestra RSS per modificare o eliminare feed/siti RSS.
• Rimossa la casella per spostare le impostazioni nella cartella corrente: ora il programma lo gestisce automaticamente (se la cartella dell'exe si chiama "sonarpad portable" o l'exe e su un drive rimovibile, salva nella cartella dell'exe in `config`, altrimenti in `%APPDATA%\\Sonarpad`, con fallback a `config` se la cartella preferita non e scrivibile).

Versione 0.5.7 - 2026-01-05
Nuove funzionalita
• Aggiunta l'opzione per registrare audiolibri in batch (conversione multipla di file e cartelle).
• Aggiunto il supporto per i file Markdown (.md).
• Aggiunta la scelta della codifica (encoding) all'apertura dei file di testo.
• Aggiunta l'opzione nel terminale per annunciare con NVDA le nuove righe in arrivo.
Miglioramenti
• Il salvataggio delle registrazioni (audiolibri) avviene ora in MP3 nativo quando selezionato.
• L'utente può scegliere dove inserire l'asterisco * che indica le modifiche non salvate (titolo finestra).
• Migliorato il sistema di aggiornamento per renderlo più robusto in diversi scenari.
• Aggiunta nel menu Modifica la funzione per rimuovere i trattini a fine riga (utile per testi OCR).

Versione 0.5.6 - 2026-01-04
Fix
  Migliorata Trova nei file: premendo Invio apre il file esattamente alla posizione dello snippet selezionato.
Miglioramenti
  Aggiunto supporto PPT/PPTX.
  Per i formati non testuali, Salva ora propone sempre .txt per evitare di rovinare la formattazione (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Aggiunta registrazione podcast da microfono e audio di sistema (menu File, Ctrl+Shift+R).

Versione 0.5.5 - 2026-01-03
Nuove funzionalita
• Aggiunto un terminale accessibile ottimizzato per programmi che inviano molto output agli screen reader (Ctrl+Shift+P).
• Aggiunta l'opzione per salvare le impostazioni utente nella cartella corrente (modalita' portable).
Fix
• Migliorati gli snippet di Trova nei file per mantenere l'anteprima allineata alla corrispondenza.

Versione 0.5.4 – 2026-01-03
Miglioramenti
• Fix alla funzione Normalizza spazi bianchi (Ctrl+Shift+Invio).
• Aggiunto supporto HTML/HTM (apertura come testo).

Versione 0.5.3 – 2026-01-02
Nuove funzionalita
• Aggiunto Trova nei file.
• Aggiunti nuovi strumenti di testo: Normalizza spazi bianchi, Riformatta righe e Pulisci testo Markdown.
• Aggiunte Statistiche testo (Alt+Y).
• Aggiunti nuovi comandi lista nel menu Modifica:
• Ordina righe (Alt+Shift+O)
• Rimuovi duplicati (Alt+Shift+K)
• Inverti righe (Alt+Shift+Z)
• Aggiunti Commenta / Decommenta righe (Ctrl+Q / Ctrl+Shift+Q).
Localizzazione
• Aggiunta la lingua spagnola.
• Aggiunta la lingua portoghese.
Miglioramenti
• Quando un file EPUB e' aperto, Salva passa automaticamente a Salva con nome ed esporta il contenuto come .txt per evitare corruzione dell'EPUB.

## 0.5.2 - 2026-01-01

* Aggiunto il changelog.
* Aggiunte le opzioni "Apri con Sonarpad" e le associazioni per i file supportati durante l'installazione.
* Migliorata la localizzazione dei messaggi (errori, dialoghi, esportazione audiolibro).
* Aggiunta la selezione delle parti quando si usa "Dividi l'audiolibro in base al testo", con opzione "Il testo deve iniziare a capo".
* Aggiunta l'importazione trascrizioni da YouTube con selezione lingua, opzione timestamp e gestione focus.

## 0.5.1 - 2025-12-31

* Aggiornamento automatico con conferma, gestione errori e notifiche migliorate.
* Esportazione audiolibro migliorata (split per testo, SAPI5/Media Foundation, controlli avanzati).
* Miglioramenti TTS (pausa/riprendi, dizionario sostituzioni, preferiti).
* Menu Vista e pannelli voci/favoriti, colore e dimensione testo.
* Lingua predefinita dal sistema e miglioramenti localizzazione.
* CI e packaging Windows (artefatti, MSI/NSIS, cache).

## 0.5.0 - 2025-12-27

* Refactor modulare (editor, file handler, menu, ricerca).
* Workflow di build/packaging Windows e aggiornamenti README/licenza.
* Fix navigazione TAB in finestra Guida.

## 0.5 - 2025-12-27

* Aggiornamento numero versione preliminare.

## 0.1.0 - 2025-12-25

* Prima versione: struttura progetto e README iniziale.
