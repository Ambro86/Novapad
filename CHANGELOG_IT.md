# Changelog

Versione 0.6.8 – 2026-03-04
Correzioni
• Corretto un bug per cui l’elenco dei file aperti veniva mostrato nel menu Aiuto invece che nel menu Finestra.
• Corretto un caso limite nello streaming in cui la riproduzione poteva partire ma la finestra “Download streaming” restava aperta quando il file scaricato era già nel formato di destinazione.
• Aggiunto il supporto ai capitoli podcast embedded nei file multimediali locali (es. metadati capitoli MP3): quando feed/URL non forniscono capitoli, Sonarpad li legge dal file scaricato in background, così la riproduzione parte subito e i capitoli vengono applicati appena disponibili.
• Migliorato il rilevamento codifica per file `.txt` giapponesi: aggiunto fallback sicuro Shift_JIS/CP932 nei casi di mojibake, preservando il comportamento esistente su UTF/diacritici/cinese.
• Refactor interno sulla sicurezza: conversione a implementazioni safe dove possibile e riduzione drastica delle righe di codice unsafe.
Versione 0.6.7 – 2026-03-02
Miglioramenti
• Aggiornata la traduzione polacca grazie a DJ Graco.
• Aggiunta la traduzione lituana.
• Aggiunta la traduzione cinese.
• D’ora in poi, build beta frequenti saranno pubblicate nella sezione Releases del progetto, così gli utenti potranno testare le nuove modifiche prima della prossima versione stabile.
• Aggiunta la scorciatoia `Ctrl+.` per inserire il carattere di ellissi (…).
• Migliorato il supporto ai capitoli podcast: la navigazione capitoli è ora più affidabile anche negli episodi diretti/streaming in cui i capitoli non sono incorporati nel file MP3, usando quando disponibili i metadati capitolo dal feed/URL come fallback. Aggiunte le scorciatoie `Ctrl+Alt+Pagina su` (capitolo precedente) e `Ctrl+Alt+Pagina giù` (capitolo successivo).
• Riorganizzate le cartelle di output in `Documenti\\Sonarpad`: i file ora vengono salvati nelle sottocartelle dedicate `audiobooks`, `documents`, `recordings` e `media`, con migrazione automatica dai percorsi legacy.
• Migliorato il supporto per file di testo molto grandi (anche 60 MB): apertura e navigazione riga per riga più fluide, in particolare con gli screen reader.
• Aggiornate le guide per tutte le lingue e aggiornate le risorse di localizzazione dell'app, incluse testo donazioni e traduzioni setup NSIS (nuove stringhe installer in cinese semplificato e lituano, più completamento della traduzione ucraina del setup).
• Aggiunto il supporto proxy di rete globale (HTTP/HTTPS e SOCKS5/SOCKS5H) per le funzioni online, con validazione al salvataggio Opzioni: i proxy non validi vengono segnalati e rimossi automaticamente.
• Aggiunta una nuova funzione in Strumenti: "Riproduci audio da streaming...", che permette di inserire un URL (YouTube o link media diretto), scegliere il formato di output e il profilo qualità/bitrate (inclusa qualità/bitrate originale per MP3 e MP4) e avviare la riproduzione nell’audio player di Sonarpad.
• Aggiunto il supporto al tasto multimediale Play/Pausa di sistema (cuffie/tastiera): ora controlla sia la riproduzione media sia la pausa/ripresa della lettura testo (con priorità al player media quando entrambi sono attivi).
• Aggiunta nel menu File > File recenti la nuova voce "Svuota file recenti" per cancellare rapidamente l’elenco dei documenti recenti.
• Ampliate le opzioni di bitrate nella conversione audio e nella registrazione podcast: aggiunti valori più bassi (64/96 kbps) ed esteso MP3 fino a 320 kbps, con validazione e gestione encoder allineate.
• Estese le opzioni di divisione audiolibro in base al tempo fino a 60 minuti.
• Migliorata la divisione audiolibro in parti: ora il numero di parti è inseribile manualmente, con validazione da 1 a 100.
• Aggiunta la nuova modalità Visualizza > Sola lettura per bloccare modifiche accidentali nel testo mantenendo piena lettura e navigazione dei documenti.
• Aggiunta una barra di progresso accessibile durante gli aggiornamenti del programma, così i lettori di schermo possono seguire in tempo reale l’avanzamento del download.
• Aggiunta una nuova barra di stato discreta nella finestra principale con conteggio caratteri, parole e posizione riga/colonna (esempio: "Caratteri (con spazi): 11. | Parole: 2. | Ln 1, Col 12"), senza interferire con il focus di NVDA.
• Aggiunta nel menu Visualizza la nuova voce A capo automatico, per attivare/disattivare rapidamente il wrapping delle righe senza aprire Opzioni.
• Aggiunte nel menu Modifica > Testo le nuove azioni per aumentare/ridurre il rientro, con scorciatoie Ctrl+Shift+. (indent) e Ctrl+Shift+, (de-indent), perché quando “Mostra voci nell’editor” è attivo il tasto Tab è riservato alla navigazione del pannello voci.
• Aggiunta la visualizzazione localizzata di data e ora per articoli RSS ed episodi podcast, con formato adattato alla lingua dell'interfaccia.
• Aggiunta nel menu contestuale RSS una nuova voce per condividere via email l'articolo selezionato.
• Aggiunte opzioni granulari di conferma eliminazione in Opzioni > RSS e podcast: per RSS (feed/articolo/entrambi/nessuno) e per Podcast (podcast/episodio/entrambi/nessuno).
• Aggiunta la copia rapida RSS configurabile con Ctrl+C (Opzioni > RSS e podcast): copia titolo, URL, contenuto articolo oppure tutto insieme.
• Unificato il flusso RSS: “Aggiungi Fonte” ora accetta sia URL feed sia parole chiave (con generazione automatica del feed Google News), senza necessità di una ricerca separata.
• Premendo Ctrl+A ora viene annunciato il completamento dell'azione per un feedback più chiaro con gli screen reader.
• Aggiunta la scorciatoia Shift+F3 per "Trova precedente" nel menu Modifica, in aggiunta a F3 "Trova successivo".
• Migliorato il messaggio di conferma delle sostituzioni con gestione corretta di singolare/plurale (es. “1 sostituzione” vs “2 sostituzioni”).
• Aggiunta nella finestra Dizionario la selezione della lingua di ricerca, con predefinito Auto (lingua interfaccia) e possibilità di override manuale.
• Aggiunta una nuova scheda Scorciatoie nelle Opzioni per personalizzare i tasti rapidi, con rilevamento dei conflitti e avviso quando una combinazione è già assegnata a un'altra azione.
• Aggiunto il supporto iniziale ai parametri da riga di comando: `-h`/`--help` mostrano la guida rapida e `--version` mostra la versione del programma.
• Resa più chiara la regolazione manuale di velocità e tono: i campi ora usano una scala centrata su 100, dove 100 corrisponde al valore normale.
• Migliorata la selezione delle voci Microsoft sia in Opzioni > Voce sia nel pannello voci dell’editor: aggiunta una casella combinata lingua localizzata per filtrare le voci per lingua, mantenendo la modalità “solo voci multilingua” come elenco unico non diviso per lingua (con combo lingua nascosta quando attiva).
• Aggiunta la configurazione della voce per i dialoghi in Opzioni > Voce con navigazione completa via Tab, usando lo stesso modello voci dell’interfaccia principale (sistema, filtro lingua Edge, voce e velocità/tono/volume con etichette); aggiunta anche la seconda voce dialoghi opzionale con gli stessi controlli (sistema, filtro lingua Edge, voce, velocità/tono/volume) per alternare i dialoghi; le regole dialoghi vengono salvate in configurazione `.ini`, senza modificare il testo del documento.
• Migliorata l’etichetta di Annulla: la voce Modifica > Annulla ora mostra l’azione che verrà annullata (ad esempio modifica testo, commenta/decommenta righe o inserimento tag voce), restando non disponibile quando non esiste nulla da annullare.
Correzioni di bug
• Corretto il supporto apertura RTF: i file `.rtf` ora vengono estratti e mostrati come testo leggibile, non più come markup RTF grezzo (es. `{\\rtf1...}`).
• Corretta l'apertura dei file di testo cinesi in codifica GB18030/GBK: Sonarpad ora li rileva e decodifica correttamente, evitando testo illeggibile (mojibake).
• Migliorata la creazione degli audiolibri M4B con metadata capitoli e marker capitolo; risolto il problema "chipmunk" (voce troppo veloce/acuta) nei file M4B generati.
• Corretta l'interfaccia bitrate nella finestra di salvataggio audiolibro: rimossi i testi hardcoded in italiano e aggiunta l'opzione 64 kbps tra i bitrate selezionabili.
• Corretto "Salva tutto" (Ctrl+Shift+S): ora tutti i documenti aperti modificati vengono rilevati in modo affidabile (inclusi tab nuovi/non salvati) e il salvataggio procede correttamente su ciascun file, aprendo "Salva con nome" quando necessario.
• Corretto l'ordinamento degli articoli RSS di Google News: quando la data è disponibile, gli articoli vengono ora mostrati dal più recente al meno recente.
• Corretta l'associazione etichette NVDA nella finestra Dizionario: campo ricerca e combobox lingua ora annunciano l'etichetta giusta.
• Corretta la gestione tastiera nella finestra Proprietà di RSS/Podcast: Tab/Shift+Tab raggiungono il pulsante OK, Invio attiva OK, Esc chiude in modo sicuro e il focus torna correttamente all'elenco RSS/Podcast.
• Corretto lo storico annullamento in RSS/Podcast: Ctrl+Z ora supporta annullamento multi-livello per rimozioni (articoli/episodi e fonti), non solo l'ultima azione.
• Migliorati gli annunci di rimozione in RSS/Podcast con messaggi espliciti (RSS rimosso, articolo RSS rimosso, episodio podcast rimosso).
• Migliorata la gestione del focus dopo elimina/annulla in RSS/Podcast: negli RSS viene selezionato in modo affidabile il primo feed quando necessario e sono state ridotte le ripetizioni degli annunci screen reader durante la riselezione ritardata.

Versione 0.6.6 – 2026-02-13
Miglioramenti
• Aggiunta "Formattazione automatica per TTS" nel menu Modifica per preparare rapidamente il testo alla lettura vocale (rimuove markdown/virgolette e ricompone le righe spezzate).
• Migliorato l'inserimento dei tag voce: ora, se è presente una selezione, i tag vengono applicati correttamente sia a una singola riga sia a più righe selezionate.
• Aggiunta un'opzione nelle impostazioni Audio per scegliere la cartella predefinita di salvataggio audiolibri (predefinita: Documenti\\Sonarpad Audiobooks).
• Nella finestra di salvataggio audiolibro, quando è attiva la divisione in parti, è stata aggiunta una nuova opzione (attiva di default) per creare una sottocartella dedicata alle parti generate.
• L'export audiolibri ora salva gli MP3 in stereo con bitrate scelto dall'utente per voci Edge, SAPI5 e SAPI4.
• Aggiunto supporto alle voci SAPI5 a 32 bit tramite bridge, così possono essere usate anche le voci disponibili solo nei motori a 32 bit.
• Riorganizzate le funzioni vocali in un menu dedicato "Voce e audio" e aggiunta/esplicitata la voce "Converti audio", utile per convertire qualunque file multimediale supportato in MP3, AAC, OGG, Opus, FLAC, WAV e AIFF.
• Aggiunta la rimozione dei singoli articoli RSS e dei singoli episodi podcast (tasto Canc + menu contestuale con conferma), senza eliminare l'intera fonte RSS/podcast, con annullamento dell'ultima rimozione (singolo articolo/episodio oppure intero podcast/feed RSS).
• Aggiunto l'export dei feed RSS in OPML nella finestra RSS, così le fonti correnti possono essere salvate e reimportate facilmente.
• Aggiunta la funzione "Cerca RSS per parola chiave" nella finestra RSS: inserendo una parola chiave viene generato automaticamente l'URL RSS di Google News e si apre la finestra di aggiunta fonte già precompilata, così i feed tematici si creano in un solo passaggio.
• Aggiunta la traduzione serba grazie a Mila Kuran.
• Aggiunta la traduzione ucraina grazie a Ivan Shtefuriak.
• Aggiunta l'apertura multipla dei file media: aprendo più file insieme viene creata una coda di riproduzione invece di sostituire il file corrente.
• Aggiunte scorciatoie di seek variabile durante la riproduzione: con base di 1 minuto, Freccia sinistra/destra sposta di 60s, Shift+Freccia sinistra/destra di 20s e Ctrl+Freccia sinistra/destra di 3 minuti.
• Aggiunte le scorciatoie per brano precedente/successivo nel player: Ctrl+Pagina su e Ctrl+Pagina giù.
• Aggiunta la voce "Reset volume" e raggruppate le azioni di ripristino in un sottomenu dedicato "Reset" in Riproduci, insieme a "Reset speed" e "Reset pitch".
• Migliorato l'installer: setup.exe ora permette di scegliere tra associare tutti i tipi file supportati oppure selezionare manualmente le singole estensioni; anche MSI ora espone la scelta per estensione nell'albero funzionalità (default invariato: tutte attive).
• Aggiunto il nuovo menu "Finestra" con la voce "Documenti aperti..." per passare rapidamente a uno dei file attualmente aperti.
• Aggiornata la voce Visualizza > Carattere: al posto del selettore completo ora c'è un sottomenu rapido con font comuni (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia), mantenendo la dimensione testo già impostata.
• Migliorata la lettura di RSS e podcast con due annunci distinti: i nodi sorgente annunciano "nuovi elementi" quando il feed/podcast ha aggiornamenti, mentre i singoli articoli RSS e i singoli episodi podcast annunciano "non letto"/"non riprodotto"; il comportamento è disattivabile dalle Opzioni.
Correzioni di bug
• Corretto il parsing del testo EPUB per i libri che contengono commenti HTML inline (<!-- ... -->): il testo dei capitoli ora viene estratto correttamente invece di essere saltato in parte o del tutto.
• Corretto il dizionario Wiktionary in spagnolo e la gestione cache del dizionario: parole come "agua" ora vengono trovate correttamente e le vecchie cache "parola non trovata" non vengono più riutilizzate.
• Corretto l'encoding nell'import degli articoli RSS per alcune fonti spagnole (es. El Mundo): accenti e "ñ" ora vengono mantenuti correttamente nell'editor temporaneo.
• Corretta la decodifica ANSI dei file in lingue centro-europee (es. ceco/polacco): Sonarpad ora distingue meglio UTF-8 e ANSI e seleziona la code page corretta (inclusa Windows-1250), evitando diacritici corrotti.
• Corretta la persistenza delle fonti RSS con parametri nella URL (es. `rss.aspx?c=...`): questi feed ora vengono salvati e ripristinati correttamente dopo il riavvio di Sonarpad.
• Corretta l'apertura dei file puntatore Google Drive (`.gdoc`, `.gsheet`, `.gslides`) dal menu contestuale di Esplora file: se la lettura diretta fallisce con “Incorrect function (os error 1)”, Sonarpad ora usa un fallback shell-open e il documento si apre correttamente.
• Corretta la lettura dei file Excel legacy `.xls` (Excel 2010): ora i file binari vecchi vengono riconosciuti/decodificati correttamente invece di mostrare testo corrotto (es. `ÐÏ_à¡±...`).
• Corretto il flusso di annuncio del correttore ortografico: gli errori vengono ora riannunciati quando si rilegge il testo, e lo stesso errore viene segnalato di nuovo se viene cancellato e riscritto.
• Corrette le operazioni testuali a livello riga (es. Ctrl+Q / Ctrl+Shift+Q, ordina/inverti/righe uniche/unisci): selezionando una sola riga con Maiusc+Freccia giù non vengono più unite o troncate le righe adiacenti.
• Corretta la gestione delle selezioni multilinea nelle operazioni testuali a riga (Ctrl+Q / Ctrl+Shift+Q e strumenti correlati): quando RichEdit fornisce separatori di riga solo CR, il testo viene normalizzato correttamente e vengono elaborate tutte le righe selezionate senza tagli di caratteri.
• Estesa la normalizzazione input TTS per simboli visibili di spazi/tab/newline (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), che con voci multilingua potevano causare ripetizioni dei paragrafi.
• Raffinata la sanitizzazione del testo Edge TTS con una pipeline unica di validazione: normalizzazione di spazi strani/invisibili, compattazione delle sequenze lunghe di punteggiatura (come "...", "!!!", "???") e salto dei chunk composti solo da punteggiatura per evitare loop di riproduzione.
• Corretto l'annuncio del tempo di riproduzione (Ctrl+I) per stream MP3/podcast: il tempo corrente ora viene limitato alla durata della traccia e la riproduzione viene fermata automaticamente se la posizione supera la fine.
• Migliorata la copertura di localizzazione dell'installer: setup.exe ora include anche ceco, polacco, francese e serbo, mentre l'MSI resta un unico pacchetto en-US per evitare confusione nelle release.
• Corretta la pulizia in disinstallazione delle voci del menu contestuale: "Apri con Sonarpad" ora viene rimosso in modo affidabile, anche in scenari legacy del registro.
• Corretta l'affidabilità di pausa/riprendi con SAPI5: la pausa con F4 ora funziona correttamente e la ripresa torna al punto previsto invece di ripartire dall'inizio.
• Corretto il flusso pausa + seek + riprendi nella riproduzione media: dopo pausa e spostamento con Freccia sinistra/destra, premendo Spazio la riproduzione riprende in modo affidabile dal punto corrente invece di fermarsi o ripartire dall'inizio.

Versione 0.6.5 – 2026-02-07
Miglioramenti
• Traduzione spagnola migliorata grazie ad Arturo Fernandez Rivas.
• Aggiornati i feed predefiniti: Affaritaliani, HuffPost Italia, La Gazzetta dello Sport. Rimosso Wired Italia.
• Aggiunta un'opzione per dividere gli audiolibri EPUB per capitoli.
• Ora la finestra per registrare i podcast è indipendente, in modo che possiate fare delle registrazioni e allo stesso tempo usare il programma Sonarpad!
• Gli articoli RSS ora usano una scheda temporanea dedicata (titolo localizzato); con Salva con nome diventa un documento normale.
• I messaggi dello screen reader ora vengono inviati anche a JAWS quando disponibile.
Correzioni di bug
• La lettura da cursore (F5) ora parte esattamente dal punto del cursore. Prima poteva partire alcune righe sopra perché l'offset del cursore non coincideva con le posizioni CRLF/UTF-16.
• Corretto un problema di redraw: digitando su una selezione il testo precedente poteva sparire finché non si spostava la selezione.
• Corretto il parsing dei capitoli EPUB: le pagine di copertina o solo immagini non generano più letture di CSS (es. "padding") o titoli "Sconosciuto".
• Corretto il problema degli audiolibri da EPUB con divisione per tempo: Edge TTS poteva fallire su chunk vuoti o troppo lunghi ("Edge audio not sent").
• Gli articoli RSS ora decodificano le entità HTML (es. &quot;, &amp;, &lt;, &gt;).
• Salva/Salva con nome ora propone il nome del file esistente quando si salvano formati non sovrascrivibili (es. EPUB), invece della prima riga.
• Risolto un problema per cui i podcast con nuovi episodi non venivano annunciati come non riprodotti, e rinominato "non ascoltato" in "non riprodotto" perché più professionale.

Versione 0.6.4 – 2026-02-05
Miglioramenti
• Il programma e' stato rinominato in Sonarpad per dare maggiore enfasi a suono e audio, che sono la chiave di questo programma.
• Aggiunta la selezione delle tracce audio nel menu Riproduzione per i file multimediali con più tracce audio (es. MKV con più lingue).
• I podcast ora indicano chiaramente quelli non ascoltati con il prefisso "Non ascoltato" prima del nome.
• Nuovo sistema di tag per cambiare voce nel testo. Esempi:
  - Voci Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Ciao</voice>
  - Voci SAPI5: <voice sapi5 Microsoft Helena Desktop>Ciao</voice>
  - Voci SAPI4: <voice sapi4 #1>Ciao</voice>
  - Con velocita/tono/volume: <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Ciao</voice>
• Arricchite le categorie dei podcast.
• Migliorata la lettura dei PDF grazie al fallback automatico su PDFium.
• Migliorato il parser degli articoli che in alcuni casi non venivano letti in modo integrale.
• Aggiunto il reset del pitch nel menu Riproduci.
• Aggiunta un'opzione nel menu contestuale per creare un audiolibro dalla selezione.
• Aggiunta la divisione degli audiolibri in base alla durata, con la possibilita di scegliere il nome del primo file.
• Localizzata la voce che indica l'autore nella lettura degli articoli (es. "di", "by", "par").
• Aggiunte opzioni di indentazione (tab/spazi con larghezza) e Tab/Shift+Tab per indentare/deindentare le righe selezionate.
• Corretto il ripulisci Markdown: ora gestisce anche i bullet '*' quando non si mantengono le liste.
• Aggiunta un'opzione per usare il nome legacy "Novapad" nel titolo della finestra e nei collegamenti del menu Avvio.
Correzioni di bug
• Corretto un bug per cui gli audiolibri con SAPI4 potevano essere creati in modo diverso da quanto previsto.
• Corretto un bug per cui, andando oltre la fine con il seek, la riproduzione ripartiva dall'inizio.
• Finestra Trova nei file: premendo Invio su un risultato ora apre alla posizione corretta dello snippet e Esc torna ai risultati.
• Finestra Opzioni: sistemato il layout visivo delle schede Generale, Voce, Editor e Audio per evitare controlli mancanti o tagliati.
• Corretto un problema dei segnalibri quando si cambiava la velocità di riproduzione.
• Corretto un problema con Podcast Index e le categorie che non si visualizzavano correttamente.
• Corretto il problema dell'apostrofo che spezzava la lettura: ora non esiste più una lettura separata per i dialoghi, si usano i tag voce.

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







