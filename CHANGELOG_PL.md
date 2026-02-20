# Dziennik zmian

Wersja 0.6.7 – 2026-02-13
Ulepszenia
• Dodano globalną obsługę proxy sieciowego (HTTP/HTTPS oraz SOCKS5/SOCKS5H) dla funkcji online, z walidacją przy zapisie opcji: nieprawidłowe proxy jest sygnalizowane i usuwane automatycznie.
• Rozszerzono opcje bitrate w konwersji audio i ustawieniach nagrywania podcastu: dodano niższe wartości (64/96 kbps) oraz zwiększono MP3 do 320 kbps, z ujednoliconą walidacją i obsługą enkodera.
• Dodano nowy tryb Widok > Tylko do odczytu, aby blokować przypadkowe edycje tekstu przy zachowaniu pełnego odczytu i nawigacji po dokumentach.
• Dodano dostępny pasek postępu podczas aktualizacji programu, aby czytniki ekranu mogły na bieżąco śledzić postęp pobierania.
• Dodano nowy, dyskretny pasek stanu w głównym oknie z liczbą znaków, słów oraz pozycją wiersz/kolumna (np. "Znaki (ze spacjami): 11. | Słowa: 2. | Ln 1, Col 12"), bez zakłócania fokusu NVDA.
• Dodano nową opcję w menu Widok dla zawijania wierszy, aby można było szybko włączać/wyłączać zawijanie bez otwierania Opcji.
• Dodano w Edycja > Tekst nowe akcje zwiększania/zmniejszania wcięcia ze skrótami Ctrl+Shift+. (wcięcie) i Ctrl+Shift+, (usuń wcięcie), ponieważ gdy włączone jest „Pokaż głosy w edytorze”, klawisz Tab jest zarezerwowany do nawigacji w panelu głosów.
• Dodano lokalizowane wyświetlanie daty i godziny w artykułach RSS oraz odcinkach podcastów, z formatem dopasowanym do języka interfejsu.
• Dodano nową akcję w menu kontekstowym RSS do udostępniania wybranego artykułu przez e-mail.
• Dodano szczegółowe opcje potwierdzania usuwania w Opcje > RSS i podcasty: dla RSS (kanał/artykuł/oba/brak) i dla podcastów (podcast/odcinek/oba/brak).
• Dodano konfigurowalne szybkie kopiowanie RSS skrótem Ctrl+C (Opcje > RSS i podcasty): kopiowanie tytułu, URL, treści artykułu albo wszystkiego razem.
• Ujednolicono przepływ RSS: „Dodaj źródło” obsługuje teraz zarówno adresy URL feedów, jak i słowa kluczowe (z automatycznym generowaniem feedu Google News), bez osobnej funkcji wyszukiwania.
• Po naciśnięciu Ctrl+A program ogłasza teraz zakończenie akcji, co daje czytelniejszą informację zwrotną dla czytników ekranu.
• Dodano Shift+F3 dla „Znajdź poprzedni” w menu Edycja, jako uzupełnienie F3 „Znajdź następny”.
• Ulepszono komunikat po zamianie, dodając poprawne formy liczby pojedynczej i mnogiej (np. „1 zamianę” vs „2 zamiany”).
• Dodano w oknie słownika wybór języka wyszukiwania, z domyślną opcją Auto (język interfejsu) oraz możliwością ręcznego wyboru.
• Dodano nową kartę Skróty w Opcjach do personalizacji skrótów klawiaturowych, z wykrywaniem konfliktów i ostrzeżeniem, gdy skrót jest już przypisany do innej akcji.
• Uproszczono ręczną regulację prędkości i tonu: pola ręczne używają teraz skali wyśrodkowanej na 100, gdzie 100 oznacza wartość normalną.
• Ulepszono wybór głosów Microsoft zarówno w Opcje > Głos, jak i w panelu głosów edytora: dodano zlokalizowaną listę języków do filtrowania głosów po języku, a tryb „tylko głosy wielojęzyczne” pozostał pojedynczą listą bez podziału na języki (lista języków jest wtedy ukrywana).
• Ulepszono etykietę Cofnij: pozycja Edycja > Cofnij pokazuje teraz, co zostanie cofnięte (np. edycja tekstu, cytowanie/odcytowanie linii lub wstawienie tagu głosu), pozostając niedostępna, gdy nie ma czego cofać.
Poprawki błędów
• Naprawiono interfejs bitrate w oknie zapisu audiobooka: usunięto teksty zakodowane na stałe po włosku i dodano 64 kbps do listy wybieralnych bitrate.
• Naprawiono „Zapisz wszystko” (Ctrl+Shift+S): wszystkie otwarte zmodyfikowane dokumenty są teraz wykrywane niezawodnie (także nowe/niezapisane karty), a zapis wszystkich działa poprawnie dla każdego pliku, otwierając „Zapisz jako”, gdy to potrzebne.
• Poprawiono kolejność artykułów RSS z Google News: gdy data publikacji jest dostępna, artykuły są teraz wyświetlane od najnowszego do najstarszego.
• Poprawiono powiązanie etykiet NVDA w oknie słownika: pole wyszukiwania i lista języka ogłaszają teraz właściwe etykiety.
• Poprawiono obsługę klawiatury w oknie Właściwości RSS/Podcast: Tab/Shift+Tab przechodzą teraz do przycisku OK, Enter aktywuje OK, Esc bezpiecznie zamyka okno, a fokus poprawnie wraca do listy RSS/Podcast.
• Poprawiono historię cofania w RSS/Podcast: Ctrl+Z obsługuje teraz wielopoziomowe cofanie usunięć (artykułów/odcinków i źródeł), a nie tylko ostatniej akcji.
• Ulepszono komunikaty usuwania w RSS/Podcast dzięki jawnym ogłoszeniom (usunięto RSS, usunięto artykuł RSS, usunięto odcinek podcastu).
• Ulepszono zachowanie fokusu po usunięciu/cofnięciu w RSS/Podcast: w RSS w razie potrzeby niezawodnie wybierany jest pierwszy kanał, a powtarzanie komunikatów czytnika ekranu podczas opóźnionej ponownej selekcji zostało ograniczone.

Wersja 0.6.6 – 2026-02-13
Ulepszenia
• Dodano „Automatyczne formatowanie dla TTS” w menu Edycja, aby szybko przygotować tekst do odczytu (usuwa markdown/cudzysłowy i scala połamane linie).
• Ulepszono wstawianie tagów głosu: gdy tekst jest zaznaczony, tagi są teraz poprawnie nakładane zarówno na pojedynczą linię, jak i na zaznaczenie wielowierszowe.
• Dodano opcję w ustawieniach Audio, aby wybrać domyślny folder zapisu audiobooków (domyślnie: Dokumenty\\Sonarpad Audiobooks).
• W oknie zapisu audiobooka, gdy aktywny jest podział na części, dodano nową opcję (włączoną domyślnie) tworzenia dedykowanego podfolderu dla wygenerowanych części.
• Eksport audiobooków zapisuje teraz MP3 w stereo z bitrate wybranym przez użytkownika dla głosów Edge, SAPI5 i SAPI4.
• Dodano obsługę głosów SAPI5 32-bit przez bridge, aby można było używać także głosów dostępnych tylko w silnikach 32-bit.
• Przeniesiono funkcje głosowe do dedykowanego menu „Głos i audio” oraz dodano/doprecyzowano opcję „Konwertuj audio”, która służy do konwersji dowolnego obsługiwanego pliku multimedialnego do MP3, AAC, OGG, Opus, FLAC, WAV i AIFF.
• Dodano usuwanie pojedynczych artykułów RSS i pojedynczych odcinków podcastów (klawisz Delete + menu kontekstowe z potwierdzeniem), bez usuwania całego źródła RSS/podcastu, wraz z cofnięciem ostatniego usunięcia (pojedynczy artykuł/odcinek lub całe źródło RSS/podcastu).
• Dodano eksport źródeł RSS do OPML w oknie RSS, aby łatwo zapisać i ponownie zaimportować aktualne źródła.
• Dodano funkcję „Wyszukaj RSS po słowie kluczowym” w oknie RSS: po wpisaniu słowa kluczowego Sonarpad automatycznie generuje adres RSS Google News i otwiera okno dodawania źródła z już uzupełnionymi polami, dzięki czemu feed tematyczny można utworzyć jednym krokiem.
• Dodano tłumaczenie serbskie dzięki Mila Kuran.
• Dodano tłumaczenie ukraińskie dzięki Ivan Shtefuriak.
• Dodano otwieranie wielu plików multimedialnych naraz: po otwarciu kilku plików tworzona jest kolejka odtwarzania zamiast zastępowania bieżącego pliku.
• Dodano skróty zmiennego przewijania podczas odtwarzania: przy bazie 1 minuty strzałka lewo/prawo przewija o 60 s, Shift+strzałka lewo/prawo o 20 s, a Ctrl+strzałka lewo/prawo o 3 minuty.
• Dodano skróty do poprzedniego/następnego utworu w odtwarzaczu: Ctrl+PageUp i Ctrl+PageDown.
• Dodano opcję „Reset głośności” i zgrupowano akcje resetu w dedykowanym podmenu „Reset” w Odtwarzaniu, razem z „Reset prędkości” i „Reset tonu”.
• Ulepszenia instalatora: setup.exe pozwala teraz wybrać między skojarzeniem wszystkich obsługiwanych typów plików a ręcznym wyborem rozszerzeń; MSI również udostępnia wybór per rozszerzenie w drzewie funkcji (domyślnie bez zmian: wszystko włączone).
• Dodano nowe menu „Okno” z opcją „Otwarte dokumenty...”, aby szybko przełączać się między aktualnie otwartymi plikami.
• Zaktualizowano opcję Widok > Czcionka: pełny selektor został zastąpiony szybkim podmenu z popularnymi czcionkami (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia), z zachowaniem aktualnego rozmiaru tekstu.
• Ulepszono odczyt RSS i podcastów dzięki dwóm osobnym komunikatom: węzły źródeł ogłaszają „nowe elementy”, gdy feed/podcast ma nowości, a pojedyncze artykuły RSS i odcinki podcastów ogłaszają „nieprzeczytane”/„nieodtworzone”; to zachowanie można wyłączyć w Opcjach.
Poprawki błędów
• Naprawiono wyodrębnianie tekstu EPUB dla książek zawierających komentarze HTML inline (<!-- ... -->): tekst rozdziałów jest teraz poprawnie parsowany zamiast być częściowo lub całkowicie pomijany.
• Naprawiono słownik Wiktionary dla języka hiszpańskiego i obsługę cache słownika: słowa takie jak „agua” są teraz poprawnie znajdowane, a stare wpisy „nie znaleziono słowa” nie są już ponownie używane.
• Naprawiono kodowanie przy imporcie artykułów RSS dla niektórych hiszpańskich źródeł (np. El Mundo): akcenty i „ñ” są teraz poprawnie zachowywane w tymczasowym edytorze.
• Naprawiono dekodowanie ANSI dla plików środkowoeuropejskich (np. czeski/polski): Sonarpad lepiej rozróżnia teraz UTF-8 i ANSI oraz wybiera właściwą stronę kodową (w tym Windows-1250), dzięki czemu znaki diakrytyczne nie są zniekształcone.
• Naprawiono zapisywanie źródeł RSS z parametrami w URL (np. `rss.aspx?c=...`): takie feedy są teraz poprawnie zapisywane i przywracane po ponownym uruchomieniu Sonarpad.
• Naprawiono otwieranie plików wskaźnikowych Google Drive (`.gdoc`, `.gsheet`, `.gslides`) z menu kontekstowego Eksploratora: gdy bezpośredni odczyt kończył się błędem „Incorrect function (os error 1)”, Sonarpad używa teraz fallbacku shell-open i dokument otwiera się poprawnie.
• Naprawiono odczyt starszych plików Excel `.xls` (Excel 2010): stare pliki binarne są teraz poprawnie wykrywane i dekodowane zamiast pokazywać uszkodzony tekst (np. `ÐÏ_à¡±...`).
• Naprawiono sposób ogłaszania błędów pisowni: błędy są teraz ponownie odczytywane podczas późniejszego przeglądania tekstu, a ten sam błąd jest znów zgłaszany po usunięciu i ponownym wpisaniu.
• Naprawiono operacje tekstowe na liniach (np. Ctrl+Q / Ctrl+Shift+Q, sortowanie/odwracanie/unikalne/łączenie linii): po zaznaczeniu jednej linii przez Shift+Strzałka w dół sąsiednie linie nie są już łączone ani obcinane.
• Naprawiono działanie wielowierszowe operacji liniowych (Ctrl+Q / Ctrl+Shift+Q i narzędzia pokrewne): gdy RichEdit zwraca separatory linii jako samo CR, są one teraz poprawnie normalizowane i przetwarzane są wszystkie zaznaczone linie bez ucinania pierwszego znaku.
• Rozszerzono normalizację wejścia TTS o widoczne symbole spacji/tabulatora/nowej linii (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), które przy głosach wielojęzycznych mogły powodować powtarzanie akapitów.
• Dopracowano sanitizację tekstu dla Edge TTS przez jedną, wspólną ścieżkę walidacji: normalizacja nietypowych/niewidocznych spacji, skracanie długich sekwencji interpunkcji (np. "...", "!!!", "???") oraz pomijanie fragmentów złożonych wyłącznie z interpunkcji, aby uniknąć pętli odtwarzania.
• Naprawiono komunikat czasu odtwarzania (Ctrl+I) dla strumieni MP3/podcast: bieżący czas jest teraz ograniczany do długości ścieżki, a odtwarzanie zatrzymuje się automatycznie, gdy pozycja wyjdzie poza koniec.
• Ulepszono pokrycie lokalizacji instalatora: setup.exe zawiera teraz także czeski, polski, francuski i serbski, a MSI pozostaje pojedynczym pakietem en-US, aby uniknąć zamieszania w wydaniach.
• Naprawiono czyszczenie podczas deinstalacji wpisów menu kontekstowego: „Otwórz za pomocą Sonarpad” jest teraz usuwane niezawodnie, także w starszych scenariuszach rejestru.
• Naprawiono niezawodność pauzy/wznawiania w SAPI5: pauza pod F4 działa teraz poprawnie, a wznowienie wraca do oczekiwanego miejsca zamiast startować od początku.
• Naprawiono przepływ pauza + przewijanie + wznowienie w odtwarzaniu multimediów: po pauzie i przewinięciu Strzałka lewo/prawo naciśnięcie Spacji niezawodnie wznawia od bieżącej pozycji zamiast zatrzymywać odtwarzanie lub uruchamiać je od początku.

Wersja 0.6.5 – 2026-02-07
Ulepszenia
• Poprawiona wersja hiszpańska dzięki Arturo Fernandez Rivas.
• Dodano opcję dzielenia audiobooków EPUB na rozdziały.
• Importy RSS używają teraz dedykowanej tymczasowej karty (zlokalizowany tytuł); Zapisz jako zamienia ją w zwykły dokument.
• Komunikaty czytnika ekranu są teraz również wysyłane do JAWS, gdy jest dostępny.
Poprawki błędów
• Odczyt od kursora (F5) teraz zaczyna się dokładnie w miejscu kursora. Wcześniej mógł startować kilka linii wyżej, bo offset kursora nie odpowiadał pozycjom CRLF/UTF-16.
• Naprawiono problem z odświeżaniem: przy pisaniu na zaznaczeniu wcześniejszy tekst mógł znikać do czasu przesunięcia zaznaczenia.
• Poprawiono parser rozdziałów EPUB: strony okładki lub tylko z obrazami nie generują już odczytu CSS (np. "padding") ani tytułów „Sconosciuto”.
• Naprawiono błąd przy dzieleniu audiobooków z EPUB według czasu: Edge TTS mógł się wyłożyć na pustych lub zbyt długich fragmentach ("Edge audio not sent").
• Artykuły RSS teraz dekodują encje HTML (np. &quot;, &amp;, &lt;, &gt;).
• Zapis/Zapisz jako teraz podpowiada nazwę istniejącego pliku przy zapisie formatów nienadpisywalnych (np. EPUB), zamiast pierwszej linii.
• Naprawiono problem, przez który podcasty z nowymi odcinkami nie były oznaczane jako nieodtworzone, oraz zmieniono „Nieodsłuchane” na „Nieodtworzone” jako bardziej profesjonalne.

Wersja 0.6.4 – 2026-02-05
Ulepszenia
• Program został przemianowany na Sonarpad, aby podkreślić dźwięk i audio jako klucz tego programu.
• Dodano wybór ścieżek audio w menu Odtwarzanie dla plików multimedialnych z wieloma ścieżkami audio (np. MKV z wieloma językami).
• Podcasty teraz wyraźnie wskazują te nieprzesłuchane z prefiksem „Nieprzesłuchane" przed nazwą.
• Nowy system tagów do zmiany głosu w tekście. Przykłady:
  - Głosy Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Cześć</voice>
  - Głosy SAPI5: <voice sapi5 Microsoft Helena Desktop>Cześć</voice>
  - Głosy SAPI4: <voice sapi4 #1>Cześć</voice>
  - Z prędkością/tonem/głośnością: <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Cześć</voice>
• Rozszerzono kategorie podcastów.
• Dodano opcję w menu kontekstowym tworzenia audiobooka z zaznaczenia.
• Dodano podział audiobooka według długości, z możliwością wyboru nazwy pierwszego pliku.
• Zlokalizowano etykietę autora w odczycie artykułów (np. "by", "di", "par").
• Dodano opcje wcięć (tabulatory/spacje z szerokością) oraz Tab/Shift+Tab do wcinania/odwcinania zaznaczonych linii.
• Poprawiono czyszczenie Markdown: obsługa wypunktowań „*” gdy zachowanie list jest wyłączone.
Poprawki błędów
• Naprawiono błąd, przez który audiobooki SAPI4 mogły być tworzone inaczej niż oczekiwano.
• Okno Znajdź w plikach: naciśnięcie Enter na wyniku otwiera teraz w prawidłowej pozycji fragmentu, a Esc wraca do wyników.
• Okno Opcje: poprawiono układ wizualny kart Ogólne, Głos, Edytor i Audio, aby uniknąć brakujących lub uciętych kontrolek.
• Naprawiono problem z zakładkami podczas zmiany prędkości odtwarzania.
• Naprawiono problem z Podcast Index i kategoriami, które nie wyświetlały się poprawnie.
• Naprawiono problem z apostrofem rozbijającym czytanie: nie ma już osobnego czytania dialogów, używane są tagi głosu.

Wersja 0.6.3 – 30.01.2026
Ulepszenia
• Ulepszone wykrywanie mikrofonu.
• Dodano natychmiastowe odtwarzanie dla wszystkich formatów.
Poprawki błędów
• Naprawiono awarię w oknie kategorii podcastów.

Wersja 0.6.2 – 30.01.2026
Nowe funkcje
• Dodano obsługę wykonywania plików (Shift+F5). Użytkownicy mogą wybrać interpreter (np. python) w Opcjach, wyszukać go na komputerze, a naciśnięcie Shift+F5 uruchamia bieżący skrypt. Pliki HTML otwierają się w przeglądarce.
• Dodano obsługę plików wskaźników Google Docs (.gdoc, .gsheet, .gslides), które automatycznie otwierają się w domyślnej przeglądarce.
• Dodano obsługę formatu audiobooków M4B (Apple/AAC).
• Dodano opcję "Pokaż odcinki" w menu kontekstowym wyników wyszukiwania podcastów, aby przeglądać i odtwarzać odcinki bez subskrypcji.
• Dodano funkcję "Idź do linii" (menu Edycja lub Ctrl+J), aby szybko przejść do konkretnego numeru linii.
• Dodano opcje menu kontekstowego do porządkowania kanałów RSS i podcastów (alfabetycznie lub według daty).
• Dodano domyślne wietnamskie kanały RSS.
• Dodano pole testowe mikrofonu w oknie nagrywania, aby sprawdzić poziomy przed rozpoczęciem.
• Dodano "Pokaż opis" dla odcinków podcastów w menu kontekstowym.
• Dodano obsługę rozszerzonych formatów audio/wideo przez FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Dodano obsługę zsynchronizowanego odczytu napisów (srt, vtt, ass, sub, sbv, lrc, smi) za pomocą NVDA lub wybranego głosu. Program szuka pliku napisów o tej samej nazwie co plik multimedialny. Dodano opcje "Importuj napisy" i "Usuń napisy" w menu Odtwarzanie dla plików o różnych nazwach.
• Dodano skojarzenia plików dla wszystkich nowych obsługiwanych formatów audio/wideo w menu kontekstowym "Otwórz w Sonarpad".
• Dodano ustawienie regulacji wysokości dźwięku dla dowolnego pliku.
• Dodano opcję w Ustawieniach ogólnych, aby włączyć lub wyłączyć anonimowe raporty o błędach. Dodano wpis w menu Pomoc, aby utworzyć diagnostyczny plik ZIP.
• Dodano opcję używania innego głosu dla dialogów, zarówno do czytania na żywo, jak i tworzenia audiobooków.
• Dodano przeglądarkę kategorii podcastów do eksplorowania podcastów według kategorii (biznes, sztuka, sport itp.).
Ulepszenia
• Otwarcie pliku audio/wideo z Eksploratora otwiera teraz bezpośrednio widok odtwarzacza zamiast edytora tekstu.
• Usunięto monit o OCR dla niedostępnych plików PDF; OCR jest teraz wykonywany automatycznie, aby poprawić szybkość i wrażenia użytkownika.
• Ulepszono Dostępny Terminal: odczyt NVDA pamięta teraz ostatnio przeczytaną linię dla lepszej ciągłości.
• SAPI 4: Tworzenie audiobooków jest teraz w pełni zrównoleglone i prawie natychmiastowe. Dodano monit o wybór liczby współbieżnych procesów.
• SAPI 4: Wyeliminowano wąskie gardło konwersji WAV do MP3 poprzez równoległe konwertowanie fragmentów podczas syntezy.
• SAPI 4: Ulepszono obsługę błędów i automatyczne czyszczenie plików tymczasowych.
• Okno Znajdź: Zmieniono nazwę "Regex" na "Wyrażenie regularne" dla jasności i dodano brakujące tłumaczenia dla opcji wyszukiwania.
• Audiobooki M4B: Lepsza obsługa wyjścia; podział na części/znaczniki tworzy teraz pojedynczy plik M4B z metadanymi rozdziałów, w tym tytułem i autorem.
• Odtwarzacz: Naprawiono precyzję zakładek i ogłaszania czasu, gdy prędkość odtwarzania nie wynosi 1.0x.
• Przywrócono nawigację Ctrl+Tab i Ctrl+Shift+Tab w Opcjach.
• Dodano opcję w menu Odtwarzanie, aby natychmiast zresetować prędkość do Normalnej (1.0x).
• Zaktualizowano wszystkie zależności do najnowszych wersji dla lepszej wydajności i stabilności.
• Zintegrowano FFmpeg z dynamicznym ładowaniem DLL, aby zapewnić kompatybilność bez blokowania uruchamiania.
• Zaktualizowano filtry pobierania podcastów, aby uwzględnić nowe formaty audio/wideo.
• Zablokowano zapisywanie plików audio/wideo przez Ctrl+S, aby uniknąć uszkodzenia.
• Ulepszono import transkrypcji YouTube, czyniąc go bardziej odpornym i niezawodnym.
• Ulepszono odporność dzielenia audiobooków na części, zapewniając, że żaden tekst nie zostanie utracony.
• Instalator jest teraz w pełni wielojęzyczny, obsługując włoski, angielski, hiszpański, portugalski, szwedzki i wietnamski w oparciu o język systemu użytkownika. Angielski jest domyślny dla nieobsługiwanych systemów.
• Kategorie podcastów: naciśnięcie Enter na kategorii potwierdza teraz wybór (odpowiednik przycisku OK).
• Ulepszono system wykrywania zawieszania, aby uniknąć fałszywych alarmów, gdy otwarte są modalne okna dialogowe (komunikaty o błędach, "nie znaleziono tekstu").
Poprawki
• Naprawiono błąd, przez który dziennik zmian nie otwierał się przy starcie.
• Naprawiono błąd, przez który monit o OCR nie pojawiał się dla niedostępnych plików PDF otwieranych z Eksploratora.
• Naprawiono błąd przy starcie, który mógł powodować utratę fokusu lub zamknięcie okna natychmiast po otwarciu.
• Naprawiono krytyczny błąd w wyszukiwaniu regex, który uniemożliwiał znalezienie tekstu, w tym problemy z "Wyszukiwaniem cyklicznym" i opcją "Kropka odpowiada nowej linii" z zakończeniami linii Windows.
Lokalizacja
• Dodano tłumaczenie na język polski.
• Dodano tłumaczenie na język francuski.
• Dodano tłumaczenie na język czeski (dzięki Radkowi Žaludowi i Jiri Holzingerowi).

Wersja 0.6.1 – 20.01.2026
Poprawki
• Naprawiono błąd, przez który włączenie "Pokaż głosy w edytorze" powodowało zatrzymanie odtwarzania podcastu.
• Naprawiono problem, przez który niektórych podcastów nie można było dodać przez URL, ponieważ URL był ucinany.
• Naprawiono błąd, przez który normalne adresy URL nie mogły być już dodawane w funkcji kanału RSS.
• Naprawiono problem, przez który opcja języka Wikipedii była wyświetlana wielokrotnie w różnych zakładkach ustawień.
• Usunięto tworzenie plików debugowania, które były nieprawidłowo generowane nawet w trybie release.
Ulepszenia
• Ulepszono obsługę głosów Microsoft, które teraz używają dedykowanej metody odtwarzania z innym user agentem.
• Dodano obsługę plików MP4.

Wersja 0.6.0 – 20.01.2025
Nowe funkcje
• Dodano sprawdzanie pisowni. Z menu kontekstowego użytkownicy mogą sprawdzić, czy bieżące słowo jest poprawne, a jeśli nie, uzyskać sugestie pisowni.
• Dodano import i eksport podcastów przez pliki OPML.
• Dodano obsługę wyszukiwania Podcast Index oprócz iTunes. Użytkownicy mogą wprowadzić swój darmowy klucz API i sekret (generowane tylko przy użyciu adresu e-mail).
• Dodano obsługę głosów SAPI4, zarówno do czytania w czasie rzeczywistym, jak i tworzenia audiobooków.
• Dodano automatyczny fallback OCR dla niedostępnych plików PDF: gdy nie znaleziono tekstu do wyodrębnienia, dokument jest rozpoznawany przez OCR.
• Dodano obsługę słownika przy użyciu Wiktionary. Naciśnięcie klawisza Aplikacji pokazuje definicje, a gdy dostępne, synonimy i tłumaczenia na inne języki.
• Dodano import artykułów z Wikipedii z wyszukiwaniem, wyborem wyników i bezpośrednim importem do edytora.
• Dodano skrót Shift+Enter w module RSS, aby otworzyć artykuł bezpośrednio na oryginalnej stronie internetowej.
Ulepszenia
• Wybór mikrofonu jest teraz zawsze respektowany przez aplikację.
• W oknie podcastów naciśnięcie Enter na odcinku teraz natychmiast ogłasza "ładowanie" przez NVDA, aby potwierdzić akcję.
• W wynikach wyszukiwania podcastów naciśnięcie Enter teraz subskrybuje wybrany podcast.
• Poprawiono i ulepszono etykiety dla skrótów Ctrl+Shift+O i Podcast Ctrl+Shift+P.
• Prędkość i głośność odtwarzania są teraz zapisywane w ustawieniach i zachowywane dla wszystkich plików audio.
• Dodano dedykowany folder pamięci podręcznej dla odcinków podcastów. Użytkownicy mogą zachować odcinki przez "Zachowaj podcast" w menu Odtwarzanie. Pamięć podręczna jest automatycznie czyszczona po przekroczeniu rozmiaru zdefiniowanego przez użytkownika (Opcje → Audio).
• Znacznie ulepszono pobieranie artykułów RSS przy użyciu podszywania się pod libcurl z profilami Chrome i iPhone, zapewniając kompatybilność z ~99% stron.
• Dodano stan przeczytany / nieprzeczytany dla artykułów RSS, z wyraźnym wskazaniem na liście RSS.
• Zamień wszystko teraz raportuje liczbę wykonanych zamian.
• Dodano przycisk Usuń podcast podczas nawigacji po bibliotece podcastów za pomocą Tab.
Poprawki
• Usunięto zbędny wpis "oczekująca aktualizacja" z menu Pomoc (aktualizacje są już obsługiwane automatycznie).
• Naprawiono błąd, przez który naciśnięcie Ctrl+S na otwartym pliku MP3 zapisywało i uszkadzało plik.
• Naprawiono problem z interfejsem użytkownika, gdzie "Audiobooki wsadowo" były wyświetlane jako "(B)… Ctrl+Shift+B" (usunięto zbędną etykietę).
• Naprawiono inteligentne cudzysłowy: po włączeniu normalne cudzysłowy są teraz poprawnie zastępowane inteligentnymi cudzysłowami.
• Naprawiono błąd, przez który użycie "Idź do zakładki" resetowało prędkość odtwarzania do 1.0.
• Naprawiono problem, przez który już pobrane odcinki podcastów były pobierane ponownie zamiast używać wersji z pamięci podręcznej.
Skróty klawiszowe
• F1 otwiera teraz Przewodnik pomocy.
• F2 sprawdza teraz dostępność aktualizacji.
• F7 / F8 skaczą teraz do poprzedniego lub następnego błędu pisowni.
• F9 / F10 szybko przełączają między ulubionymi głosami.
Ulepszenia deweloperskie
• Błędy nie są już milcząco pomijane: usunięto wszystkie wzorce let _ =, a błędy są teraz jawnie obsługiwane (propagowane, logowane lub obsługiwane z fallbackami w stosownych przypadkach).
• Projekt teraz nie kompiluje się, jeśli są ostrzeżenia: zarówno cargo check, jak i cargo clippy muszą przejść czysto, z zaostrzonymi lintami i usuniętymi allow tam, gdzie to możliwe.
• Usunięto niestandardowe implementacje, takie jak pomocniki w stylu strlen / wcslen. Długości ciągów i buforów UTF-16 są teraz wyznaczane z danych posiadanych przez Rust zamiast skanowania pamięci.
• Obsługa DLL została uporządkowana i skonsolidowana wokół libloading, unikając niestandardowej logiki ładowania i parsowania PE.
• Usunięto ręcznie pisane pomocniki parsowania bajtów; całe parsowanie bajtów używa teraz standardowych from_le_bytes / from_be_bytes na sprawdzonych wycinkach.
Te zmiany zmniejszają niepotrzebne użycie unsafe, eliminują potencjalne niezdefiniowane zachowanie i sprawiają, że baza kodu jest bardziej idiomatyczna, solidna i łatwa w utrzymaniu.

Wersja 0.5.9 - 13.01.2025
Nowe funkcje
• Dodano zmianę kolejności RSS z menu kontekstowego (góra/dół/do pozycji) ze sprawdzaniem nieprawidłowej pozycji.
• Dodano menu kontekstowe artykułu z otwieraniem oryginalnej strony i udostępnianiem przez WhatsApp, Facebook i X.
• Dodano skrót Esc, aby powrócić z importowanych artykułów do listy RSS.
• Dodano tryb podcastów: wyszukiwanie, subskrypcja, słuchanie; zmiana kolejności subskrypcji; Esc zatrzymuje odtwarzanie i wraca do listy; Enter na odcinku rozpoczyna odtwarzanie.
• Dodano kontrolę prędkości odtwarzania dla podcastów i plików MP3.
• Dodano Ctrl+T, aby przejść do określonego czasu.
• Dodano przycisk podglądu głosu za polem wyboru głośności.
• Dodano wyszukiwanie i zamianę za pomocą wyrażeń regularnych (styl Notepad++).
• Dodano import RSS z plików OPML i TXT.
• Dodano opcję włączenia "Otwórz w Sonarpad" w Eksploratorze plików, w tym w wersjach portable.
Ulepszenia
• Ulepszono wybór prędkości/tonu/głośności głosu, respektując maksymalne limity TTS.
• Różne ulepszenia RSS, aby pobierać wszystkie artykuły bez przesuwania fokusu NVDA podczas aktualizacji.
• Ulepszono odtwarzanie dźwięku dzięki dedykowanemu menu, ogłaszaniu czasu Ctrl+I i głośności do 300%.
• Dodano brakujące skróty dla niektórych funkcji.
• Zreorganizowano menu Edycja z podmenu czyszczenia tekstu.
• Zreorganizowano Opcje w zakładki, z nawigacją Ctrl+Tab i Ctrl+Shift+Tab.
• Czytnik RSS pobiera teraz pełną treść artykułu, pasującą do widoku przeglądarki.
Poprawki
• Naprawiono usuwanie cyfr na początku linii podczas czyszczenia Markdown.
• Naprawiono wyzwalanie cofania przez AltGr+Z.
• Naprawiono anulowanie nagrywania audiobooka, aby zatrzymywało się szybko.
Lokalizacja
• Dodano tłumaczenie wietnamskie (dzięki Anh Đức Nguyễn).

Wersja 0.5.8 - 10.01.2026
Nowe funkcje
• Dodano regulację głośności dla mikrofonu i dźwięku systemowego podczas nagrywania podcastów.
• Dodano nową funkcję importu artykułów ze stron internetowych lub kanałów RSS, w tym najważniejsze kanały dla każdego języka.
• Dodano funkcję usuwania wszystkich zakładek dla bieżącego pliku.
• Dodano funkcję usuwania zduplikowanych linii i zduplikowanych kolejnych linii.
• Dodano funkcję zamykania wszystkich kart lub okien z wyjątkiem bieżącego.
• Dodano wpis Darowizny w menu Pomoc dla wszystkich języków.
Ulepszenia
• Ulepszono dostępny terminal, aby zapobiec niektórym awariom.
• Ulepszono i naprawiono klawisze dostępu i skróty klawiaturowe w całej aplikacji.
• Naprawiono problem, przez który zamknięcie okna odtwarzania audio nie zatrzymywało odtwarzania.
• Dodano okna dialogowe potwierdzenia dla ważnych działań (np. usuń zduplikowane linie, usuń łączniki na końcu linii, usuń wszystkie zakładki w bieżącym pliku). Okno dialogowe nie jest wyświetlane, gdy akcja nie ma zastosowania.
• Dodano możliwość usuwania kanałów/stron RSS z biblioteki poprzez wybranie ich i naciśnięcie Delete.
• Dodano menu kontekstowe w oknie RSS do edycji lub usuwania kanałów/stron RSS.
• Usunięto ustawienie przenoszenia ustawień do bieżącego folderu; aplikacja teraz obsługuje to automatycznie na podstawie lokalizacji (jeśli folder exe nazywa się "sonarpad portable" lub exe znajduje się na dysku wymiennym, ustawienia trafiają do folderu exe w `config`, w przeciwnym razie `%APPDATA%\Sonarpad`, z fallbackiem do exe `config`, jeśli preferowany folder nie jest zapisywalny).

Version 0.5.7 - 2026-01-05
New features
• Added Batch Audiobooks feature to convert multiple files/folders at once.
• Added support for Markdown files (.md).
• Added file encoding selection when opening text files.
• Added option in the accessible terminal to announce new lines with NVDA.
Improvements
• Audiobook recording now saves natively to MP3 when selected.
• User can now choose the position of the "unsaved changes" asterisk (*) in the window title.
• Improved the update system robustness across different scenarios.
• Added "Remove Hyphens" in Edit menu to fix OCR line-endings.

Version 0.5.6 - 2026-01-04
Fixes
  Improved Find in Files so pressing Enter opens the file exactly at the selected snippet.
Improvements
  Added PPT/PPTX support (open as text).
  Opening non-text formats now saves as .txt to avoid formatting corruption (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Added podcast recording from microphone and system audio (File menu, Ctrl+Shift+R).

Version 0.5.5 – 2026-01-03
New features
• Added an accessible terminal optimized for large output and screen readers (Ctrl+Shift+P).
• Added a setting to save user settings in the current folder (portable mode).
Fixes
• Improved Find in Files snippets so the preview stays aligned with the match.

Version 0.5.4 – 2026-01-03
Improvements
• Fixed Normalize Whitespace (Ctrl+Shift+Enter).
• Added HTML/HTM support (open as text).

Version 0.5.3 – 2026-01-02
New features
• Added Find in Files.
• Added new text tools: Normalize Whitespace, Hard Line Break, and Strip Markdown.
• Added Text Statistics (Alt+Y).
• Added new list commands in the Edit menu:
• Order Items (Alt+Shift+O)
• Keep Unique Items (Alt+Shift+K)
• Reverse Items (Alt+Shift+Z)
• Added Quote / Unquote Lines (Ctrl+Q / Ctrl+Shift+Q).
Localization
• Added Spanish localization.
• Added Portuguese localization.
Improvements
• When an EPUB file is open, Save now automatically switches to Save As and exports the content as a .txt file to prevent EPUB corruption.

## 0.5.2 - 2026-01-01
- Added a changelog.
- Added open-with-Sonarpad options and file associations for supported files during installation.
- Improved message localization (errors, dialogs, audiobook export).
- Added part selection when using "Split audiobook based on text", with a "Require the marker at line start" option.
- Added YouTube transcript import with language selection, timestamp option, and improved focus handling.

## 0.5.1 - 2025-12-31
- Automatic updates with confirmation, improved error handling and notifications.
- Audiobook export improvements (text-based split, SAPI5/Media Foundation, advanced controls).
- TTS improvements (pause/resume, replacement dictionary, favorites).
- View menu and voice/favorites panels, text color and size.
- Default language from system locale and localization improvements.
- CI and Windows packaging (artifacts, MSI/NSIS, cache).

## 0.5.0 - 2025-12-27
- Modular refactor (editor, file handler, menu, search).
- Windows build/packaging workflow and README/license updates.
- Fix TAB navigation in the Help window.

## 0.5 - 2025-12-27
- Preliminary version bump.

## 0.1.0 - 2025-12-25
- Initial release: project structure and README.






