# Changelog

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
 Migliorata Trova nei file: premendo Invio apre il file esattamente alla posizione dello snippet selezionato.
Miglioramenti
 Aggiunto supporto PPT/PPTX.
 Per i formati non testuali, Salva ora propone sempre .txt per evitare di rovinare la formattazione (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
 Aggiunta registrazione podcast da microfono e audio di sistema (menu File, Ctrl+Shift+R).

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
