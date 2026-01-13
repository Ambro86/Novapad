# Changelog

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
• Aggiunta nelle Opzioni la casella per abilitare "Apri con Novapad" in Esplora risorse, anche in versione portable.
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
• Rimossa la casella per spostare le impostazioni nella cartella corrente: ora il programma lo gestisce automaticamente (se la cartella dell'exe si chiama "novapad portable" o l'exe e su un drive rimovibile, salva nella cartella dell'exe in `config`, altrimenti in `%APPDATA%\\Novapad`, con fallback a `config` se la cartella preferita non e scrivibile).

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
* Aggiunte le opzioni "Apri con Novapad" e le associazioni per i file supportati durante l'installazione.
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
