# Přehled změn

Verze 0.7.0 – 2026-04-12

Co je nového
• Přidána podpora přehrávače mpv pro streamované přehrávání. Videa z YouTube a podporovaných webů se nyní přehrávají okamžitě; pokud si je uživatel chce uložit, stáhnou se jako dříve. Při přepisu streamovaného obsahu se nejprve stáhne a poté přepíše. Přehrávač mpv se také používá pro otevírání lokálních videí a práci s titulky, což zajišťuje lepší kompatibilitu s mnoha formáty.
• Vylepšeno nahrávání podcastů systémového zvuku: nyní si můžete zvolit, zda chcete nahrávat veškerý systémový zvuk, jednu aplikaci nebo více aplikací současně. Tato volba je integrována do běžného nahrávání, takže mikrofon lze stále samostatně zapnout nebo vypnout.
• Přidán jazyk hindština. Rozhraní přeloženo, přidány RSS, seznam změn a příručka Sonarpad.
• Přidána možnost v kartě Editor, která při použití šipek nahoru a dolů vždy přesune kurzor na začátek řádku.
• Přidána možnost v nabídce "Převést audio" pro převod zvuku do formátu M4B.

Opravy
• Opravena klávesa `F10`, takže při čtení textu znovu přepíná na další oblíbený hlas.
• Když probíhá nahrávání podcastu, zavření jiného dokumentu už nezavře také aktivní nahrávání.
• V komentářích YouTube otevřených z „Přehrát streamované audio...“ nyní Sonarpad nejprve načte pouze prvních 50 komentářů nejvyšší úrovně, vždy včetně všech odpovědí k těmto komentářům, a na konci přidá položku pro načtení všech komentářů podle potřeby.
• Záložky se nyní zobrazují a zpracovávají podle své pozice jak v textových dokumentech, tak v multimediálních souborech, místo aby sledovaly pořadí vytvoření. Pokud už záložka na stejné pozici existuje, znovu se nepřidá.

Verze 0.6.9 – 2026-04-08

Opravy
• Vylepšeno používání funkce Najít v souborech: při otevření Procházet složku se zaměření nyní přesune přímo na seznam složek; otevření výsledku klávesou Enter už nenarušuje klávesové příkazy; stisknutí Esc vrátí zaměření na dříve vybraný výsledek; a při návratu pomocí Alt+Tab se zaměření přesune buď na pole hledání, nebo na seznam výsledků, pokud jsou výsledky otevřené.
• Klávesa F5 vždy spouštěla čtení od začátku. To bylo nyní opraveno a čtení začíná od aktuální pozice kurzoru, přičemž `Shift+F5` a `Ctrl+F5` zůstávají zachovány pro navigaci na předchozí a další větu.
• Po použití funkce Přejít na řádek mohlo stisknutí Esc přesunout zaměření mimo Sonarpad. Nyní správně vrací zaměření do editoru.
• Možnost „Zalamovat řádky“ se nyní použije okamžitě i na již otevřené dokumenty, místo aby se projevila až po znovuotevření souboru.

Verze 0.6.8 – 2026-04-07

Co je nového
• Přidána nová položka do nabídky Přehrávání, která umožňuje přepisovat jakýkoli audio nebo video soubor pomocí Whisper. V Možnostech je nyní dostupná nová sekce „AI a přepis“, kde lze vybrat model, zapnout volitelnou podporu CUDA pro grafické karty NVIDIA, zachovat původní jazyk a zapnout nebo vypnout časové značky.
• Přidána nová akce „Přepsat aktuální složku“ v nabídce Přehrávání, která zpracuje všechny podporované audio soubory ve složce právě otevřeného média do jednoho společného dokumentu, s vlastním ukazatelem průběhu, stavem aktuálního souboru a podporou zrušení. Lze ji také spustit pomocí Alt+Shift+C.
• Přidáno offline hlasové diktování, které využívá stejný postup jako přepis audia. Ve výchozím nastavení stiskněte Ctrl+Shift+Space pro spuštění diktování a stejnou zkratku znovu pro zastavení; tuto zkratku lze změnit v Možnostech. Od druhého použití je diktování rychlejší, protože modul zůstává připravený v paměti; toto přednačtení a opětovné použití se automaticky vypnou na počítačích s méně než 4 GB RAM.
• Přidána nová možnost Editoru, ve výchozím nastavení vypnutá, která umožňuje, aby Esc zavřelo okno editoru.
• Vyhledávání podcastů nyní ve výchozím nastavení používá iTunes + Spreaker, s filtrováním duplicit, pokud je stejný podcast nalezen na obou platformách.
• Vylepšeno procházení a vyhledávání Apple podcastů: vyhledávání podcastů, procházení kategorií a nejlepší podcasty podle kategorií nyní používají vybranou zemi adresáře podcastů. V Možnosti > RSS a podcasty můžete ponechat Automaticky pro použití systémové země, nebo ručně zvolit jinou zemi.
• Zvýšen limit výsledků pro kategorie Apple podcastů. Při prvním otevření se stále načte prvních 50 výsledků jako dříve; pokud zvolíte Načíst další výsledky, Sonarpad načte až 200 výsledků celkem (limit Apple) a umožní procházet další stránky při zachování plynulého ovládání.
• Sonarpad je nyní dostupný také na Macu s omezenou sadou funkcí. Odkaz na projekt: https://github.com/Ambro86/Sonarpad-Mac

Vylepšení
• Přidáno více než 50 volitelných zemí pro adresář podcastů, takže uživatelé mohou vybírat z mnohem širší nabídky národních katalogů.
• „Přehrát streamované audio...“ nyní umí také vyhledávat na YouTube podle libovolného textového dotazu nebo přijmout odkaz na YouTube kanál či playlist a zobrazit jeho výsledky.
• Vylepšeno zobrazení výsledků v „Přehrát streamované audio...“: položky YouTube nyní přehledněji zobrazují název, délku, kanál a počet zhlédnutí.
• „Přehrát streamované audio...“ nyní podporuje také komentáře YouTube: lze je otevřít z kontextové nabídky, číst odpovědi a rozbalovat vlákna komentářů pomocí klávesy Šipka vpravo.
• Přidány oblíbené položky YouTube pro kanály a playlisty v „Přehrát streamované audio...“: lze je přidat z výsledků přes kontextovou nabídku, otevřít přímo ze seznamu Oblíbené, který je dostupný klávesou Tab hned za polem URL/dotazu YouTube, a později je odstranit z téhož seznamu pomocí kontextové nabídky. Ve výsledcích hledání YouTube je kontextová nabídka dostupná pouze pro kanály a playlisty.
• „Přehrát streamované audio...“ nyní může vyžadovat přihlašovací údaje, když streamovací web vyžaduje přihlášení. Uživatelé je mohou zadat, uložit pro daný web a později spravovat uložené přihlašovací údaje v Možnosti > Audio.
• Vylepšena práce se zaměřením během „Přehrát streamované audio...“, takže okno průběhu zůstává stabilnější během stahování a převodu.
• Přidány dvě nové akce pro navigaci při čtení v nabídce Hlas a zvuk: Předchozí věta a Další věta, s nastavitelnými zkratkami pro skoky při čtení textu.
• Výchozí zkratka pro Spustit soubor interpretem je nyní Ctrl+Shift+F5, takže Shift+F5 lze ve výchozím nastavení použít pro akci Předchozí věta.
• Přidány hlasové profily v Možnosti > Hlas: profily lze přidávat, používat a mazat.
• Rozšířeny možnosti intervalu přeskočení médií v Možnosti > Audio o další hodnoty od 1 sekundy až do 2 hodin.
• Přidán ruský překlad díky Dmitriyovi.
• Přidána nová možnost v Možnosti > Audio pro výběr formátu pojmenování částí audioknihy: Název + číslo, Pouze číslo nebo Číslo + název.
• Přidány oblíbené články RSS: z kontextové nabídky článku lze položky přidat do zvláštního kanálu Oblíbené.
• RSS kanál Oblíbené lze smazat a při přidání nového článku do oblíbených se automaticky znovu vytvoří.
• Přidány klávesové zkratky RSS pro přesun kanálů nahoru/dolů: Ctrl+Shift+Šipka nahoru a Ctrl+Shift+Šipka dolů.
• Vylepšeno okno RSS o vestavěný náhled článku, takže text článku lze zkontrolovat přímo tam a rychle k němu přejít pomocí Tab ještě před otevřením celého článku v editoru.
• Přidána výslovná položka RSS „Načíst další zprávy“ na konci kanálů, pokud jsou dostupné další položky; stisknutí Enter načte další dávku a přesune zaměření na první nově načtený článek.
• Ve slovníku hlasů je nyní při přidávání nebo úpravě náhrady k dispozici zaškrtávací pole „Rozlišovat velikost písmen“, takže každá náhrada může buď respektovat, nebo ignorovat velikost písmen.

Opravy
• „Přehrát streamované audio...“ nyní respektuje limit mezipaměti podcastů již nastavený v Možnostech a stejný limit se nyní vztahuje také na přehrávání audiopopisů.
• Opraven import z Wikipedie, takže bloky citací přítomné na stránkách se nyní importují správně.
• Vylepšen parser webových stránek pro stránky WordPress, kde mohly být vynechány položky seznamu a některé nadpisy sekcí.
• „Přejít na řádek“ nyní předvyplní pole aktuálním číslem řádku.
• Opraven export OPML pro podcasty a RSS, takže exportované soubory jsou nyní přijímány iTunes.
• Přidány lokalizované potvrzovací zprávy pro správný import a export OPML RSS kanálů a podcastů.
• Opraven problém, kdy v „Přehrát streamované audio...“ zadání hledaného textu a výběr YouTube kanálu z výsledků mohl způsobit, že program vypadal jako zaseknutý, místo aby otevřel videa daného kanálu.
• Opraven problém, kdy se seznam otevřených dokumentů zobrazoval v nabídce Nápověda místo v nabídce Okno.
• Opraven okrajový problém streamování, kdy se přehrávání mohlo spustit, ale dialog „Stahování streamu“ zůstal otevřený, když stažený soubor již odpovídal cílovému formátu.
• Opraveno chování převodu MP3 streamů: pokud je stream již ve formátu MP3 a uživatel zvolí konkrétní bitrate MP3 (například 128 kbps), Sonarpad nyní znovu zakóduje na vybraný bitrate místo přeskočení převodu.
• Opraveny dokumenty přepisu médií, takže jejich zavření nyní vyžaduje potvrzení uložení a navrhovaný název souboru správně znovu používá název přepsaného mediálního souboru místo prvního řádku textu.
• Opravena zkratka Alt+Shift+L: nyní správně otevírá seznam kapitol během přehrávání.
• Opravena zkratka Alt+Shift+T: nyní správně spouští „Přepsat aktuální audio“ místo otevření nabídky Nástroje.
• Opraveno zastavení přehrávání v nabídce Přehrávání: stisknutí . se nyní chová jako Zastavit a zastaví pouze aktuální stopu místo toho, aby zároveň ukončilo přehrávač/epizodu.
• Opravena položka uložení v nabídce Přehrávání pro média otevřená z Nedávných souborů: pokud soubor pochází z místní cache Sonarpad, lokalizovaná akce uložení se nyní správně zobrazuje i tam.
• Když přepis začne ve chvíli, kdy se již přehrává audio, Sonarpad nyní toto audio automaticky pozastaví před zahájením přepisu.
• Opraven problém, kdy import článku z Wikipedie mohl uspět, aniž by se text článku zobrazil na obrazovce.
• Přidána podpora vložených kapitol podcastů z místních mediálních souborů (např. metadata kapitol MP3): když nejsou k dispozici kapitoly z feedu/URL, Sonarpad nyní načte kapitoly ze staženého souboru na pozadí, takže přehrávání začne okamžitě a data kapitol se použijí, jakmile budou připravena.
• Opraveno načítání kapitol u stažených epizod podcastů otevřených jako běžné místní mediální soubory: vložené kapitoly jsou nyní dostupné i zde, nejen když přehrávání začíná z okna Podcasty.
• Opravena finalizace MP3 audioknih pro SAPI4 a SAPI5: konečný výstup je nyní správně dokončen, aby se předešlo neúplným nebo křehkým souborům po dlouhých exportech.
• Přidán výslovný ukazatel průběhu finalizace pro všechny režimy vytváření audioknih: po fázi vytváření nyní Sonarpad oznamuje a zobrazuje zvláštní fázi finalizace s viditelným průběhem.
• Opraveno ladění hlasů dialogů: nastavení rychlosti/výšky/hlasitosti se nyní správně používá pro první i druhý hlas dialogu během syntézy.
• Vylepšena detekce kódování textu pro japonské soubory .txt: přidána bezpečná záložní volba Shift_JIS/CP932 pro případy zkomoleného textu, při zachování stávajícího chování pro UTF/diakritiku/čínštinu.
• Interní bezpečnostní refaktor: funkce byly tam, kde to bylo možné, převedeny na bezpečné implementace a počet řádků s nebezpečným kódem byl výrazně snížen.

Verze 0.6.7 – 2026-03-02

Vylepšení
• Program nyní dokáže hromadně zpracovat funkci Nahradit vše i u velkých souborů s velmi vysokým počtem nahrazení.
• Aktualizován polský překlad díky DJ Graco.
• Přidán litevský překlad.
• Přidán čínský překlad.
• Od této chvíle budou v sekci vydání projektu pravidelně zveřejňovány časté beta verze, aby uživatelé mohli testovat nové změny před další stabilní verzí.
• Přidána zkratka Ctrl+tečka pro vložení znaku výpustky (…).
• Vylepšena podpora kapitol podcastů: navigace mezi kapitolami nyní funguje spolehlivěji, včetně přímých/streamovaných epizod, kde kapitoly nejsou vložené v souboru MP3, díky použití záložních metadat kapitol z feedu/URL, pokud jsou k dispozici. Přidány zkratky pro navigaci mezi kapitolami Ctrl+Alt+PageUp (předchozí kapitola) a Ctrl+Alt+PageDown (další kapitola).
• Přeskupeny výstupní složky Sonarpad do Documents\Sonarpad: soubory jsou nyní ukládány do vyhrazených podsložek audiobooks, documents, recordings a media, s automatickou migrací ze starších cest.
• Vylepšena podpora velmi velkých textových souborů (včetně 60 MB): plynulejší otevírání a navigace po řádcích, zejména se čtečkami obrazovky.
• Aktualizovány příručky pro všechny jazyky a obnoveny lokalizační zdroje napříč aplikací, včetně textů pro dary a překladů instalátoru NSIS (nové řetězce instalátoru pro zjednodušenou čínštinu a litevštinu, plus dokončený ukrajinský překlad instalátoru).
• Přidána globální podpora síťové proxy (HTTP/HTTPS a SOCKS5/SOCKS5H) pro online funkce, s ověřením proxy při ukládání v Možnostech: neplatná proxy jsou oznámena a automaticky odstraněna.
• Přidána nová akce v nabídce Nástroje: „Přehrát streamované audio...“, která umožňuje vložit URL (YouTube nebo přímý odkaz na médium), zvolit výstupní formát a kvalitu a přehrát jej přímo v audio přehrávači Sonarpad.
• Přidána podpora systémové klávesy Přehrát/Pozastavit média (sluchátka/klávesnice): nyní ovládá jak přehrávání médií, tak pozastavení/obnovení čtení textu (s prioritou přehrávání médií, pokud jsou aktivní obě).
• Přidána nová položka v Soubor > Nedávné soubory: „Vymazat nedávné soubory“ pro rychlé smazání seznamu posledních dokumentů.
• Rozšířeny možnosti bitrate v Převést audio a v nastavení nahrávání podcastů: přidány nižší hodnoty (64/96 kbps) a MP3 rozšířeno až na 320 kbps, s odpovídající validací a zpracováním v enkodéru.
• Rozšířeny možnosti rozdělení audioknih podle času až na 60 minut.
• Vylepšeno rozdělení audioknih podle částí: uživatelé nyní mohou ručně zadat počet částí, s validací od 1 do 100.
• V nabídce Zobrazení byl přidán nový režim „Jen pro čtení“, který uzamkne editor proti nechtěným úpravám a přitom ponechá dokumenty plně čitelné a procházetelné.
• Během aktualizací programu přidán přístupný ukazatel průběhu, aby čtečky obrazovky mohly v reálném čase sledovat průběh stahování.
• Přidán nový tichý stavový řádek v hlavním okně zobrazující znaky, slova a řádek/sloupec (například: „Znaky (včetně mezer): 11. | Slova: 2. | Řádek 1, sloupec 12“) bez narušení zaměření NVDA.
• Přidána nová položka „Zalamovat řádky“ v nabídce Zobrazení, takže zalamování lze rychle měnit bez otevírání Možností.
• Přidány nové akce v nabídce Úpravy > Text: „Zvětšit odsazení řádku/bloku“ a „Zmenšit odsazení řádku/bloku“ se zkratkami Ctrl+Shift+. (odsadit) a Ctrl+Shift+, (zmenšit odsazení), protože když je zapnuto „Zobrazit hlasy v editoru“, klávesa Tab je vyhrazena pro navigaci v panelu hlasů.
• Přidáno lokalizované datum/čas v RSS článcích a epizodách podcastů, s formátováním přizpůsobeným aktuálnímu jazyku rozhraní.
• Přidána nová akce v kontextovém menu RSS pro sdílení vybraného článku e-mailem.
• Přidány podrobné možnosti potvrzení mazání pro RSS a podcasty v Možnosti > RSS a podcasty: RSS (feed/článek/oboje/žádné) a Podcasty (podcast/epizoda/oboje/žádné).
• Přidáno nastavitelné rychlé kopírování RSS pomocí Ctrl+C (Možnosti > RSS a podcasty): kopírovat titulek, URL, obsah článku nebo vše dohromady.
• Sjednoceno vytváření RSS zdrojů: „Přidat zdroj“ nyní přijímá jak přímé URL feedu, tak zadání klíčového slova (automaticky generuje Google News RSS), čímž nahrazuje potřebu samostatné akce pro vyhledávání podle klíčového slova.
• Stisknutí Ctrl+A nyní oznámí dokončení pro jasnější zpětnou vazbu čtečkám obrazovky.
• Přidána klávesa Shift+F3 pro „Najít předchozí“ v menu Úpravy, jako doplněk k F3 „Najít další“.
• Vylepšeny zprávy zpětné vazby při nahrazování se správnými tvary jednotného a množného čísla (např. „Provedena 1 náhrada“ vs „Provedeny 2 náhrady“).
• Přidán výběr jazyka pro slovník ve slovníkovém okně, s výchozí volbou Auto (jazyk rozhraní) a volitelným ručním nastavením.
• Přidána nová karta Zkratky v Možnostech pro přizpůsobení klávesových zkratek s detekcí konfliktů, která upozorní, pokud je zkratka již přiřazena jiné akci.
• Přidána počáteční podpora přepínačů příkazové řádky: -h/--help nyní zobrazují informace o použití a --version vypíše verzi programu.
• Vylepšena srozumitelnost ručního nastavování rychlosti a výšky hlasu: ruční pole nyní používají stupnici se středem 100, kde 100 odpovídá normální hodnotě.
• Vylepšen výběr hlasů Microsoft v Možnosti > Hlas i v panelu Hlas v editoru: přidán lokalizovaný jazykový seznam pro filtrování hlasů podle jazyka, přičemž režim pouze vícejazyčných hlasů zůstává jako jediný neseskupený seznam hlasů (jazykový seznam je při jeho zapnutí skryt).
• Přidána konfigurace hlasu dialogů v Možnosti > Hlas s plnou navigací pomocí Tab, používající stejný systém TTS jako hlavní rozhraní (systém TTS, jazyk hlasu, hlas a ruční ladění hlasu); přidán volitelný druhý hlas dialogů se stejnými ovládacími prvky pro střídající se dialogy; pravidla pro hlasy dialogů jsou ukládána do konfiguračního .ini, takže text dokumentu není upravován.
• Vylepšeno označení Zpět: položka Úpravy > Zpět nyní zobrazuje, jaká akce bude vrácena zpět (například úpravy textu, citovat/odcitovat řádky nebo vložení hlasového tagu), přičemž zůstává zakázaná, když není co vracet.

Opravy
• Opraveno otevírání souborů RTF: dokumenty .rtf jsou nyní parsovány a zobrazovány jako obyčejný čitelný text namísto surového RTF zápisu (např. {\\rtf1...}).
• Opraveno otevírání čínských textových souborů kódovaných v GB18030/GBK: Sonarpad nyní tyto soubory správně rozpozná a dekóduje, čímž se zabrání zkomolenému výstupu.
• Vylepšeno vytváření audioknih M4B s metadaty kapitol a značkami kapitol; opraven problém „chipmunk“ přehrávání (vysoká výška/rychlost) u vygenerovaných souborů M4B.
• Opraveno uživatelské rozhraní bitrate v dialogu ukládání audioknih: odstraněny natvrdo vložené italské popisky a přidána možnost 64 kbps mezi volitelné bitrate.
• Opraveno Uložit vše (Ctrl+Shift+S): všechny otevřené upravené dokumenty jsou nyní spolehlivě rozpoznány (včetně neuložených/nových karet) a Uložit vše správně uloží každý dokument nebo otevře Uložit jako, pokud je to potřeba.
• Opraveno řazení položek Google News RSS: články jsou nyní zobrazovány podle data publikace sestupně (nejnovější první), pokud jsou data dostupná.
• Opravena asociace popisků pro NVDA ve slovníkovém okně: pole hledání a jazykový seznam nyní oznamují správné popisky.
• Opraveno ovládání klávesnice v okně Vlastnosti RSS/Podcast: Tab/Shift+Tab nyní dosáhne na tlačítko OK, Enter aktivuje OK, Esc bezpečně zavře okno a zaměření se správně vrací do seznamu RSS/Podcast.
• Opravena historie vrácení změn RSS/Podcast: Ctrl+Z nyní podporuje víceúrovňové vracení odstranění (článků/epizod i zdrojů), nejen poslední akci.
• Vylepšena zpětná vazba při odstraňování RSS/Podcastů pomocí výslovných stavových oznámení (RSS odstraněno, RSS článek odstraněn, epizoda podcastu odstraněna).
• Vylepšeno chování fokusu RSS/Podcast po smazání/vrácení: RSS nyní spolehlivě zaostří první feed, když je to potřeba, a vyhýbá se opakovaným oznámením čtečky obrazovky při zpožděném znovuvýběru.

Verze 0.6.6 – 2026-02-13

Vylepšení
• Přidána možnost „Automatické formátování pro TTS“ v menu Úpravy pro rychlou přípravu textu pro řeč (odstraní markdown/uvozovky a znovu spojí zalomené řádky).
• Vylepšeno vkládání hlasových tagů: pokud je text vybrán, tagy se nyní správně použijí jak na výběr v jednom řádku, tak na víceřádkový výběr.
• Přidána možnost výchozí složky pro ukládání audioknih v nastavení Audio (výchozí: Documents\Sonarpad Audiobooks).
• V dialogu ukládání audioknih při zapnutém rozdělování přidána nová výchozí možnost pro vytvoření samostatné podsložky pro rozdělené části (pro přehlednější organizaci výstupu).
• Export audioknih nyní ukládá MP3 ve stereu s uživatelem zvoleným bitrate pro hlasy Edge, SAPI5 a SAPI4.
• Přidána podpora 32bitových hlasů SAPI5 přes bridge, takže hlasy dostupné pouze v 32bitových enginech lze také použít v Sonarpad.
• Funkce hlasu byly přesunuty do samostatné nabídky „Hlas a zvuk“ a byla přidána/upřesněna funkce „Převést audio...“, užitečná pro převod jakéhokoli podporovaného mediálního souboru do MP3, AAC (M4A), OGG (Vorbis), Opus, FLAC, WAV a AIFF.
• Přidáno odstraňování jednotlivých RSS článků a epizod podcastů (klávesa Delete + kontextové menu s potvrzením), aniž by byl odstraněn celý RSS/podcastový zdroj, plus vrácení posledního odstranění (jediný článek/epizoda nebo celý RSS/podcastový zdroj).
• Přidán export RSS zdrojů do OPML v RSS okně, takže aktuální RSS zdroje lze snadno uložit a znovu importovat.
• Přidána funkce „Hledat RSS podle klíčového slova“ v RSS okně: zadání klíčového slova nyní automaticky vygeneruje URL Google News RSS a otevře dialog přidání zdroje s předvyplněnými údaji, takže lze feedy podle klíčových slov vytvářet v jednom kroku.
• Přidán srbský překlad díky Mila Kuran.
• Přidán ukrajinský překlad díky Ivan Shtefuriak.
• Přidáno otevírání více mediálních souborů najednou: výběr/otevření více mediálních souborů nyní vytvoří frontu přehrávání místo nahrazení aktuálního souboru.
• Přidány zkratky pro proměnlivé posouvání během přehrávání: se základním skokem 1 minuta posouvají Left/Right o 60 s, Shift+Left/Right o 20 s a Ctrl+Left/Right o 3 minuty.
• Přidány zkratky pro předchozí/další stopu v přehrávači: Ctrl+PageUp a Ctrl+PageDown.
• Přidána funkce „Normální hlasitost (100%)“ a seskupeny obnovovací akce do samostatného podmenu „Reset“ v Přehrávání vedle „Normální rychlost (1x)“ a „Normální výška (0)“.
• Vylepšení instalátoru: setup.exe nyní umožňuje uživatelům zvolit mezi přiřazením všech podporovaných typů souborů nebo ručním výběrem přípon; MSI nyní nabízí volby asociací souborů po jednotlivých příponách ve stromu funkcí (výchozí zůstává vše povoleno).
• Přidána nová nabídka „Okno“ s položkou „Otevřít dokumenty...“ pro rychlé přepnutí na libovolný aktuálně otevřený soubor.
• Aktualizováno Zobrazení > Písmo: starý výběr byl nahrazen rychlou podnabídkou běžných písem (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia) při zachování aktuální velikosti textu.
• Vylepšena oznámení RSS/Podcastů pomocí duálního modelu stavu: uzly zdrojů oznamují „nové položky“, když má zdroj nebo podcast aktualizace, zatímco jednotlivé RSS články a epizody podcastů oznamují „nepřečteno“ / „nepřehráno“; toto chování lze vypnout v Možnostech.

Opravy
• Opravena extrakce textu z EPUB pro knihy obsahující vložené HTML komentáře (`<!-- ... -->`): text kapitol je nyní parsován správně místo částečného nebo úplného přeskočení.
• Opraveno vyhledávání ve španělském Wiktionary a zpracování cache slovníku: španělské položky jako „agua“ se nyní načítají správně a staré záznamy cache „Slovo nenalezeno“ se již znovu nepoužívají.
• Opraveno kódování znaků při importu RSS článků z některých španělských zdrojů (např. El Mundo): písmena s diakritikou a „ñ“ jsou nyní správně zachována v dočasném editoru.
• Opraveno dekódování ANSI textu pro středoevropské soubory (např. čeština/polština): Sonarpad nyní lépe rozlišuje UTF-8 vs ANSI a vybírá správnou kódovou stránku (včetně Windows-1250), aby se zabránilo poškození diakritiky.
• Opravena perzistence RSS zdrojů pro feedy s parametry dotazu v URL (např. rss.aspx?c=...): tyto feedy se nyní po restartu Sonarpad správně ukládají a obnovují.
• Opraveno otevírání souborů ukazatelů Google Drive (.gdoc, .gsheet, .gslides) z kontextového menu Průzkumníka: když přímé čtení selže s chybou „Incorrect function (os error 1)“, Sonarpad nyní použije shell-open, aby se dokument přesto správně otevřel.
• Opraveno čtení starých souborů Excel 2010 .xls: staré binární excelové soubory jsou nyní správně rozpoznány a dekódovány místo zobrazení zkomoleného textu (např. ÐÏ_à¡±...).
• Opraven průběh oznamování kontroly pravopisu: chybně napsaná slova jsou nyní znovu oznamována při pozdější kontrole textu a stejná chyba je znovu nahlášena, pokud je smazána a znovu napsána.
• Opraveny textové akce založené na řádcích (např. Ctrl+Q / Ctrl+Shift+Q, řazení/obrácení/jedinečné/sloučení řádků): výběr jednoho řádku pomocí Shift+Down již neslučuje ani nezkracuje sousední řádky.
• Opraveno víceřádkové chování pro textové akce založené na řádcích (Ctrl+Q / Ctrl+Shift+Q a související nástroje): výběry RichEdit používající oddělovače pouze CR jsou nyní správně normalizovány, takže všechny vybrané řádky jsou zpracovány bez oříznutí prvních znaků.
• Rozšířena normalizace vstupu TTS pro viditelné symboly bílých znaků (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), aby se zabránilo opakovanému přehrávání odstavců u vícejazyčných hlasů.
• Upřesněna sanitizace textu Edge TTS pomocí jediného validačního řetězce: zvláštní/neviditelné mezery jsou normalizovány, dlouhé sekvence interpunkce (např. "...", "!!!", "???") jsou zkráceny a úseky obsahující pouze interpunkci jsou přeskočeny, aby se zabránilo smyčkám přehrávání.
• Opraveno oznamování času přehrávání (Ctrl+I) pro streamy MP3/podcastů: aktuální čas je nyní omezen délkou stopy a přehrávání se automaticky zastaví, pokud pozice překročí konec.
• Vylepšeno pokrytí lokalizace instalátoru: setup.exe nyní obsahuje další jazyky instalátoru (čeština, polština, francouzština, srbština), zatímco MSI zůstává jako jediný balíček en-US, aby se předešlo zmatku při vydání.
• Opraveno vyčištění při odinstalaci pro položky kontextového menu: „Otevřít v Sonarpadu“ je nyní spolehlivě odstraněno, včetně starších scénářů registru.
• Opravena spolehlivost pozastavení/obnovení u SAPI5: F4 nyní správně pozastaví a obnovení pokračuje z očekávané pozice místo restartu od začátku.
• Opraven průběh pozastavení + posun + obnovení pro přehrávání médií: po pozastavení a posunu pomocí Left/Right nyní stisknutí Space spolehlivě pokračuje z aktuální pozice místo zastavení nebo restartu od začátku.

Verze 0.6.5 – 2026-02-07

Vylepšení
• Vylepšen španělský překlad díky Arturo Fernandez Rivas.
• Přidána možnost rozdělit EPUB audioknihy podle kapitol.
• Importy RSS nyní používají vyhrazenou dočasnou kartu (lokalizovaný název); Uložit jako ji převede na běžný dokument.
• Zprávy pro čtečky obrazovky jsou nyní při dostupnosti posílány také do JAWS.

Opravy
• Čtení od kurzoru (F5) nyní začíná přesně na pozici kurzoru. Dříve mohlo začínat o několik řádků výše, protože posun kurzoru neodpovídal pozicím CRLF/UTF-16.
• Opraven problém s překreslováním, kdy psaní přes výběr mohlo způsobit dočasné zmizení dřívějšího textu, dokud se výběr neposunul.
• Opraveno parsování kapitol EPUB, takže stránky pouze s obálkou nebo obrázkem již nevedou k předčítání CSS (např. „padding“) nebo názvům „Neznámý“.
• Opraveno rozdělení audioknih podle času z EPUB s Edge TTS, které selhávalo na prázdných/příliš velkých úsecích („Edge audio not sent“).
• RSS články nyní dekódují HTML entity (např. `&quot;`, `&amp;`, `&lt;`, `&gt;`).
• Uložit/Uložit jako nyní navrhuje existující název souboru při ukládání nepřepisovatelných formátů (např. EPUB) místo prvního řádku.
• Opraven problém, kdy podcasty s novými epizodami nebyly oznamovány jako nepřehrané, a „Unheard“ bylo přejmenováno na „Unplayed“ pro profesionálnější označení.

Verze 0.6.4 – 2026-02-05

Vylepšení
• Program byl přejmenován na Sonarpad, aby zdůraznil zvuk a audio jako hlavní zaměření.
• Přidán výběr zvukové stopy v menu Přehrávání pro mediální soubory s více zvukovými stopami (např. MKV soubory s více jazyky).
• Podcasty nyní jasně označují nepřehrané epizody předponou „Nepřehráno“ před názvem.
• Nové přepínání hlasů v textu pomocí tagů. Příklady:
  - Hlasy Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Hello</voice>
  - Hlasy SAPI5: <voice sapi5 Microsoft Helena Desktop>Hello</voice>
  - Hlasy SAPI4: <voice sapi4 #1>Hello</voice>
  - Se změnou rychlosti/výšky/hlasitosti: <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Hello</voice>
• Rozšířené kategorie podcastů.
• Vylepšené čtení PDF s automatickým přepnutím na PDFium.
• Vylepšen parser článků pro případy, kdy se obsah nenačetl celý.
• Přidáno resetování výšky hlasu (pitch) v menu Přehrávání.
• Přidána možnost v kontextovém menu „Vytvořit audioknihu z výběru“.
• Přidáno rozdělení audioknihy podle délky, s možností zvolit název prvního souboru.
• Lokalizován štítek autora při čtení článků (např. „by“, „di“, „par“).
• Přidány možnosti odsazení (tabulátory/mezery s nastavením šířky) a odsazení/odsazení zpět pomocí Tab/Shift+Tab na vybraných řádcích.
• Opraveno čištění Markdownu pro správné zpracování odrážek „*“, když je zachování odrážek vypnuto.
• Přidána možnost používat starý název „Novapad“ v titulku okna a ve zkratkách nabídky Start.

Opravy
• Opraven problém, kdy audioknihy SAPI4 byly vytvářeny jinak, než se očekávalo.
• Opraven problém, kdy posun za konec mediálního souboru znovu spustil přehrávání od začátku.
• Okno Najít v souborech: stisknutí Enter na výsledku nyní otevře správnou pozici a Esc vrátí zpět na výsledky.
• Okno Možnosti: vylepšené rozložení na kartách Obecné, Hlas, Editor a Audio, aby se zabránilo chybějícím nebo oříznutým prvkům.
• Opraven problém se záložkami při změně rychlosti přehrávání.
• Opraveno zobrazování kategorií Podcast Index.
• Opraven problém s apostrofy, které narušovaly čtení – odstraněno oddělené čtení dialogů, místo toho se používají voice tagy.

Verze 0.6.3 – 2026-01-30

Vylepšení
• Vylepšena detekce mikrofonu.
• Přidána podpora okamžitého přehrávání pro všechny formáty.

Opravy
• Opraven pád aplikace v okně kategorií podcastů.

Verze 0.6.2 – 2026-01-30

Nové funkce
• Přidána podpora spouštění souborů (Shift+F5). Uživatelé mohou v Možnostech zvolit interpret (např. python), vyhledat ho v počítači a stisknutím Shift+F5 spustit aktuální skript. HTML soubory se otevírají v prohlížeči.
• Přidána podpora odkazových souborů Google Docs (.gdoc, .gsheet, .gslides), které se automaticky otevřou ve výchozím prohlížeči.
• Přidána podpora formátu audioknih M4B (Apple/AAC).
• Přidána možnost „Zobrazit epizody“ v kontextovém menu výsledků vyhledávání podcastů pro procházení a přehrávání epizod bez odběru.
• Přidána funkce „Přejít na řádek“ (menu Úpravy nebo Ctrl+J).
• Přidány možnosti v kontextovém menu pro řazení RSS a podcastů (abecedně nebo podle data).
• Přidány výchozí RSS kanály pro vietnamštinu.
• Přidáno testovací pole mikrofonu v dialogu nahrávání.
• Přidána možnost „Zobrazit popis“ epizod podcastů v kontextovém menu.
• Přidána podpora rozšířených audio/video formátů přes FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Přidána podpora synchronizovaného čtení titulků (srt, vtt, ass, sub, sbv, lrc, smi) pomocí NVDA nebo zvoleného hlasu. Program hledá soubor titulků se stejným názvem jako mediální soubor. Přidány možnosti „Přidat titulky...“ a „Odebrat načtené titulky“ v menu Přehrávání.
• Přidány asociace souborů pro všechny nové podporované formáty v menu „Otevřít v Sonarpadu“.
• Přidáno nastavení výšky hlasu (pitch) pro jakýkoli soubor.
• Přidána možnost v obecném nastavení zapnout nebo vypnout anonymní hlášení chyb. Přidána položka v menu Nápověda pro vytvoření diagnostického ZIP souboru.
• Přidána možnost použít jiný hlas pro dialogy, jak při živém čtení, tak při tvorbě audioknih.
• Přidán prohlížeč kategorií podcastů.

Vylepšení
• Otevření audio/video souboru z Průzkumníka nyní otevře přímo přehrávač místo textového editoru.
• Odstraněn dotaz OCR pro nepřístupné PDF – OCR se nyní provádí automaticky.
• Vylepšen Přístupný terminál – NVDA si pamatuje poslední přečtený řádek.
• SAPI4: tvorba audioknih je nyní paralelní a téměř okamžitá.
• SAPI4: odstraněno úzké hrdlo převodu WAV→MP3 díky paralelnímu zpracování.
• SAPI4: vylepšené zpracování chyb a čištění dočasných souborů.
• V dialogu hledání bylo „Regex“ přejmenováno na „Regular expression“.
• M4B audioknihy: lepší práce s výstupem a kapitolami.
• Přehrávač: opraveny záložky a čas při jiné rychlosti než 1.0x.
• Obnovena navigace Ctrl+Tab a Ctrl+Shift+Tab v Možnostech.
• Přidána možnost rychlého resetu rychlosti na 1.0x.
• Aktualizovány všechny závislosti.
• Integrovaný FFmpeg s dynamickým načítáním DLL.
• Aktualizovány filtry stahování podcastů.
• Zabráněno ukládání audio/video souborů pomocí Ctrl+S.
• Vylepšen import YouTube transkriptů.
• Vylepšeno dělení audioknih bez ztráty textu.
• Instalátor je nyní vícejazyčný.
• Kategorie podcastů: Enter nyní potvrzuje výběr.
• Vylepšen systém detekce zamrznutí.

Opravy
• Opraven problém, kdy se changelog neotevřel při spuštění.
• Opraven problém s OCR při otevření PDF z Průzkumníka.
• Opraven problém při startu způsobující ztrátu zaměření nebo zavření okna.
• Opraven kritický problém v regex hledání (Wrap around, Dot matches newline).

Lokalizace
• Přidán polský překlad.
• Přidán francouzský překlad.
• Přidán český překlad (díky Radek Žalud a Jiri Holzinger).

Verze 0.6.1 – 2026-01-20

Opravy
• Opraven problém, kdy zapnutí „Zobrazit hlasy v editoru“ způsobovalo zastavení přehrávání podcastu.
• Opraven problém, kdy některé podcasty nešlo přidat pomocí URL, protože URL byla zkrácena.
• Opraven problém, kdy běžné URL již nešlo přidat ve funkci RSS kanálů.
• Opraven problém, kdy se možnost jazyka Wikipedie zobrazovala vícekrát v různých kartách nastavení.
• Odstraněno vytváření ladicích souborů, které se chybně generovaly i v produkčním režimu.

Vylepšení
• Vylepšená podpora hlasů Microsoft, které nyní používají speciální metodu přehrávání s odlišným user agentem.
• Přidána podpora souborů MP4.

Verze 0.6.0 – 2026-01-20

Nové funkce
• Přidána kontrola pravopisu. V kontextové nabídce mohou uživatelé zkontrolovat, zda je aktuální slovo správné, a pokud ne, získat návrhy oprav.
• Přidán import a export podcastů pomocí souborů OPML.
• Přidána podpora vyhledávání Podcast Index vedle iTunes. Uživatelé mohou zadat svůj bezplatný API klíč a tajný klíč (generovaný pouze pomocí e-mailu).
• Přidána podpora hlasů SAPI4 pro čtení v reálném čase i tvorbu audioknih.
• Byla přidána automatická podpora OCR pro nepřístupné PDF: pokud není nalezen extrahovatelný text, dokument je rozpoznán pomocí OCR.
• Přidána podpora slovníku pomocí Wiktionary. Stisknutím klávesy Applications se zobrazí definice a pokud jsou dostupné, také synonyma a překlady do jiných jazyků.
• Přidán import článků z Wikipedie s vyhledáváním, výběrem výsledků a přímým importem do editoru.
• Přidána zkratka Shift+Enter v RSS modulu pro otevření článku přímo na původní webové stránce.

Vylepšení
• Výběr mikrofonu je nyní vždy respektován aplikací.
• V okně podcastů nyní stisknutí Enter na epizodě okamžitě oznámí „načítání“ přes NVDA pro potvrzení akce.
• Ve výsledcích vyhledávání podcastů nyní Enter odebírá vybraný podcast.
• Opraveny a vylepšeny popisky pro zkratky Ctrl+Shift+O a Ctrl+Shift+P (Podcast).
• Rychlost přehrávání a hlasitost jsou nyní ukládány v nastavení a platí pro všechny audio soubory.
• Přidána speciální složka cache pro epizody podcastů. Uživatelé mohou epizody uchovat pomocí „Zachovat podcast“ v menu přehrávání. Cache se automaticky čistí při překročení uživatelem definované velikosti (Možnosti → Audio).
• Výrazně vylepšeno načítání RSS článků pomocí libcurl impersonation s profily Chrome a iPhone, což zajišťuje kompatibilitu s ~99 % webů.
• Přidán stav přečteno / nepřečteno pro RSS články s jasným označením v seznamu.
• Funkce Nahradit vše nyní hlásí počet provedených nahrazení.
• Přidáno tlačítko Smazat podcast při navigaci v knihovně podcastů pomocí Tab.

Opravy
• Odstraněna redundantní položka „Čekající aktualizace“ z menu Nápověda (aktualizace jsou již řešeny automaticky).
• Opraven problém, kdy stisknutí Ctrl+S na otevřeném MP3 souboru způsobilo jeho poškození.
• Opraven problém UI, kde „Hromadné audioknihy“ bylo zobrazeno jako „(B)… Ctrl+Shift+B“.
• Opraveny chytré uvozovky: při zapnutí se nyní správně nahrazují běžné uvozovky.
• Opraven problém, kdy „Přejít na záložku“ resetovalo rychlost přehrávání na 1.0.
• Opraven problém, kdy již stažené epizody podcastů byly znovu stahovány místo použití cache.

Klávesové zkratky
• F1 nyní otevře nápovědu.
• F2 nyní zkontroluje aktualizace.
• F7 / F8 nyní přechází na předchozí / další pravopisnou chybu.
• F9 / F10 nyní rychle přepíná mezi oblíbenými hlasy.

Vylepšení pro vývojáře
• Chyby již nejsou tiše ignorovány: všechny vzory let _ = byly odstraněny a chyby jsou nyní explicitně řešeny.
• Projekt nyní selže při kompilaci, pokud existují varování.
• Odstraněny vlastní implementace strlen / wcslen.
• Zpracování DLL bylo sjednoceno s využitím knihovny libloading.
• Odstraněno ruční parsování bajtů, nyní se používají standardní metody.
Tyto změny zvyšují robustnost, bezpečnost a udržovatelnost kódu.

Verze 0.5.9 - 2026-01-13

Nové funkce
• Přidáno řazení RSS v kontextové nabídce (nahoru/dolů/na pozici).
• Přidána nabídka článku: otevřít v prohlížeči a sdílet přes WhatsApp, Facebook a Twitter/X.
• Přidána zkratka Esc pro návrat z článku do RSS seznamu.
• Přidán režim podcastů: vyhledávání, odběr, poslech.
• Přidána kontrola rychlosti přehrávání.
• Přidáno Ctrl+T pro skok na čas.
• Přidáno tlačítko náhledu hlasu.
• Přidáno hledání a nahrazování pomocí regexu.
• Přidán import RSS z OPML a TXT.
• Přidána možnost „Otevřít v Sonarpadu“ do kontextové nabídky.

Vylepšení
• Vylepšeno ovládání hlasu.
• Vylepšeno RSS bez změny zaměření NVDA.
• Vylepšeno audio přehrávání.
• Přidány chybějící zkratky.
• Reorganizováno menu Úpravy.
• Reorganizovány Možnosti do karet.
• RSS nyní načítá celý obsah článku.

Opravy
• Opraveno odstraňování čísel v Markdownu.
• Opraven AltGr+Z (undo).
• Opraveno zrušení nahrávání audioknihy.

Lokalizace
• Přidán vietnamský překlad.

Verze 0.5.8 - 2026-01-10

Nové funkce
• Přidáno ovládání hlasitosti mikrofonu a systému při nahrávání podcastů.
• Přidán import článků z webů a RSS.
• Přidáno odstranění všech záložek.
• Přidáno odstranění duplicitních řádků.
• Přidáno zavření všech oken kromě aktuálního.
• Přidána položka „Přispět na vývoj programu“ v nabídce Nápověda.

Vylepšení
• Vylepšen přístupný terminál.
• Opraveny zkratky.
• Opraveno přehrávání po zavření okna.
• Přidána potvrzení akcí.
• Přidáno mazání RSS pomocí Delete.
• Přidáno menu pro úpravu RSS.
• Odstraněno ruční nastavení složky konfigurace (nyní automatické).

Verze 0.5.7 - 2026-01-05

Nové funkce
• Přidána funkce Hromadné audioknihy.
• Přidána podpora Markdown (.md).
• Přidán výběr kódování souboru.
• Přidáno oznamování nových řádků NVDA.

Vylepšení
• Audioknihy se ukládají přímo do MP3.
• Nastavitelná pozice hvězdičky změn.
• Vylepšen systém aktualizací.
• Přidáno odstranění spojovníků.

Verze 0.5.6 - 2026-01-04

Opravy
• Vylepšeno Najít v souborech.

Vylepšení
• Přidána podpora PPT/PPTX.
• Netextové formáty se ukládají jako .txt.
• Přidáno nahrávání podcastů.

Verze 0.5.5 – 2026-01-03

Nové funkce
• Přidán přístupný terminál.
• Přidán přenosný režim.

Opravy
• Vylepšeno Najít v souborech.

Verze 0.5.4 – 2026-01-03

Vylepšení
• Opravena funkce „Normalizovat mezery“.
• Přidána podpora souborů HTML.

Verze 0.5.3 – 2026-01-02

Nové funkce
• Přidáno Najít v souborech.
• Přidány nástroje pro text.
• Přidána statistika textu.
• Přidány příkazy pro seznamy.
• Přidány funkce „Citovat řádky“ a „Zrušit citaci řádků“.

Lokalizace
• Přidána španělština.
• Přidána portugalština.

Vylepšení
• EPUB se ukládá jako .txt.

Verze 0.5.2 - 2026-01-01
• Přidán seznam změn.
• Přidány asociace souborů.
• Vylepšena lokalizace.
• Přidáno dělení audioknih.
• Přidán import YouTube transkriptů.

Verze 0.5.1 - 2025-12-31
• Automatické aktualizace.
• Vylepšení audioknih.
• Vylepšení TTS.
• Menu Zobrazení a panely.
• Lokalizace.
• CI a balíčkování.

Verze 0.5.0 - 2025-12-27
• Modulární refaktor.
• Workflow pro sestavení Windows verze.
• Oprava TAB v nápovědě.

Verze 0.5 - 2025-12-27
• Předběžné zvýšení verze.

Verze 0.1.0 - 2025-12-25
• První vydání: struktura projektu a README.
