# Journal des modifications

Version 0.6.8 – 2026-03-31

Nouveautés
• Ajout d’un nouvel élément dans le menu Lecture permettant de transcrire n’importe quel fichier audio ou vidéo avec Whisper. Une nouvelle section « IA et transcription » est désormais disponible dans Options, où vous pouvez choisir le modèle, activer la prise en charge optionnelle de CUDA pour les cartes graphiques NVIDIA, conserver la langue d’origine et activer ou désactiver les horodatages.
• Ajout dans le menu Lecture de la nouvelle action `Transcrire le dossier actuel`, qui transcrit tous les fichiers audio pris en charge du dossier du média ouvert et les regroupe dans un document unique, avec une fenêtre de progression dédiée, l’indication du fichier en cours et la possibilité d’annuler. Elle peut aussi être lancée avec `Alt+Shift+C`.
• Ajout de la dictée vocale hors ligne, avec le même fonctionnement que la transcription audio. Par défaut, appuyez sur `Ctrl+Shift+Espace` pour démarrer la dictée puis appuyez à nouveau sur le même raccourci pour l’arrêter ; ce raccourci peut être personnalisé dans les Options. À partir de la deuxième activation, la dictée est plus rapide car le moteur reste prêt en mémoire ; sur les PC disposant de moins de 4 Go de RAM, ce préchargement et cette réutilisation sont désactivés automatiquement.
• La recherche de podcasts utilise désormais `iTunes + Spreaker` par défaut, avec filtrage des doublons lorsque le même podcast est présent sur les deux plateformes.
• Amélioration de la recherche et de l’exploration des podcasts Apple : la recherche de podcasts, la navigation par catégorie et les top podcasts par catégorie utilisent désormais le pays sélectionné pour le répertoire de podcasts. Dans Options > RSS / Podcast, vous pouvez laisser `Automatique` pour utiliser le pays du système ou choisir manuellement un autre pays.
• Sonarpad est désormais disponible aussi sur Mac, même avec un ensemble de fonctions partiel. Lien du projet : https://github.com/Ambro86/Sonarpad-Mac

Améliorations
• Plus de 50 pays sélectionnables ont été ajoutés pour le répertoire de podcasts, afin de pouvoir choisir parmi un éventail bien plus large de catalogues nationaux.
• « Lire l'audio en streaming... » permet désormais aussi d'effectuer une recherche YouTube à partir de n'importe quel texte, ou de coller le lien d'une chaîne ou d'une playlist YouTube pour afficher ses résultats.
• L’affichage des résultats dans « Lire l'audio en streaming... » a été amélioré : les entrées YouTube incluent maintenant le titre, la durée, la chaîne et le nombre de vues dans un format plus clair.
• « Lire l'audio en streaming... » prend désormais aussi en charge les commentaires YouTube : ils peuvent être ouverts depuis le menu contextuel, les réponses peuvent être lues et les fils de commentaires peuvent être développés avec la Flèche droite.
• Ajout des favoris YouTube pour les chaînes et les playlists dans « Lire l'audio en streaming... » : ils peuvent être ajoutés depuis les résultats via le menu contextuel, ouverts directement depuis la liste Favoris accessible avec Tab juste après le champ URL/requête YouTube, puis supprimés depuis cette même liste avec le menu contextuel. Dans les résultats de recherche YouTube, le menu contextuel n’est disponible que pour les chaînes et les playlists.
• « Lire l'audio en streaming... » peut désormais demander des identifiants lorsqu’un site exige une connexion. L’utilisateur peut les saisir, les enregistrer pour ce site et gérer ensuite les identifiants enregistrés dans Options > Audio.
• Amélioration du focus pendant « Lire l'audio en streaming... », afin que la fenêtre de progression reste plus stable pendant le téléchargement et la conversion.
• Ajout de deux nouvelles actions de lecture dans le menu Voix : `Phrase précédente` et `Phrase suivante`, avec des raccourcis configurables pour se déplacer pendant la lecture du texte.
• Le raccourci par défaut de `Exécuter le fichier avec l’interpréteur` est maintenant `Ctrl+Shift+F5`, afin que `Shift+F5` puisse être utilisé par défaut pour `Phrase précédente`.
• Ajout de la gestion des profils vocaux dans Options > Voix : il est possible d'ajouter, de renommer et de supprimer des profils.
• Extension dans Options > Audio des choix pour l’intervalle de retour arrière pendant la lecture, avec de nouvelles valeurs allant de 1 seconde à 2 heures.
• Ajout de la traduction russe grâce à Dmitriy.
• Ajout dans Options > Audio d’un nouveau choix pour le format du nom des parties du livre audio : `Titre + numéro`, `Numéro uniquement` ou `Numéro + titre`.
• Ajout dans le menu contextuel des articles RSS de l'action pour ajouter un article aux favoris.
• La source RSS "Favoris" peut être supprimée et est recréée automatiquement lors du prochain ajout d'un article aux favoris.
• Ajout de raccourcis clavier RSS pour déplacer les sources vers le haut/le bas : `Ctrl+Shift+Flèche haut` et `Ctrl+Shift+Flèche bas`.
• Amélioration de la fenêtre RSS avec un aperçu d'article intégré, afin de consulter directement le texte dans la fenêtre et d'y accéder rapidement avec Tab avant d'ouvrir l'article complet dans l'éditeur.
• Ajout dans RSS d’une entrée explicite « Charger plus d’actualités » à la fin des sources lorsque d’autres éléments sont disponibles ; en appuyant sur Entrée, le bloc suivant est chargé et le focus se déplace vers le premier nouvel article.
• Dans le dictionnaire vocal, lors de l’ajout ou de la modification d’un remplacement, une case « Respecter la casse » permet désormais de choisir si chaque substitution doit respecter ou ignorer les majuscules/minuscules.
Correctifs
• « Lire l'audio en streaming... » respecte désormais la limite de cache des podcasts déjà définie dans les Options, et cette même limite s'applique aussi à la lecture des audiodescriptions.
• Correction de l’importation depuis Wikipédia, qui sur certaines pages n’importait pas correctement les citations présentes dans le texte.
• Amélioration de l’analyseur de pages web : sur certaines pages WordPress, les éléments de liste et certains titres de section n’étaient pas inclus.
• Désormais, lorsque l’on utilise « Aller à la ligne », le champ est prérempli avec la ligne actuelle.
• Correction de l’export OPML des podcasts et des flux RSS : les fichiers générés sont désormais acceptés par iTunes.
• Correction de la transcription des fichiers multimédias : désormais, lorsque l’on ferme avec Alt+F4 le document généré, Sonarpad demande s’il faut l’enregistrer et propose le bon nom en se basant sur le nom du fichier transcrit au lieu de la première ligne du texte.
• Ajout de messages de confirmation localisés pour l’importation et l’exportation OPML corrects des flux RSS et des podcasts.
• Correction d’un problème où, dans « Lire l'audio en streaming... », après avoir saisi une recherche textuelle et sélectionné une chaîne YouTube dans les résultats, le programme pouvait sembler bloqué au lieu d’ouvrir les vidéos de la chaîne.
• Correction d’un bug où la liste des fichiers ouverts s’affichait dans le menu Aide au lieu du menu Fenêtre.
• Correction d’un cas limite de streaming où la lecture pouvait démarrer mais la fenêtre « Téléchargement du flux » restait ouverte lorsque le fichier téléchargé correspondait déjà au format cible.
• Correction du comportement de conversion en streaming MP3 : lorsque le flux est déjà en MP3 et que l’utilisateur choisit un bitrate MP3 explicite (par exemple 128 kbps), Sonarpad réencode désormais au bitrate sélectionné au lieu d’ignorer la conversion.
• Correction du raccourci `Alt+Shift+L` : il ouvre désormais correctement la liste des chapitres pendant la lecture.
• Correction du raccourci `Alt+Shift+T` : il lance désormais correctement « Transcrire l’audio en cours » au lieu d’ouvrir le menu Outils.
• Si un audio est déjà en cours de lecture, Sonarpad le met désormais automatiquement en pause avant de démarrer la transcription.
• Correction d’un problème où l’import d’un article depuis Wikipédia pouvait réussir sans afficher le texte de l’article à l’écran.
• Ajout de la prise en charge des chapitres de podcast intégrés dans les fichiers multimédias locaux (par ex. métadonnées de chapitres MP3) : lorsque le flux/URL ne fournit pas de chapitres, Sonarpad les charge désormais depuis le fichier téléchargé en arrière-plan, ce qui permet de démarrer la lecture immédiatement puis d’appliquer les chapitres dès qu’ils sont disponibles.
• Correction du chargement des chapitres pour les épisodes de podcast téléchargés et ouverts comme de simples fichiers multimédias locaux : les chapitres intégrés sont désormais disponibles aussi dans ce cas, et pas seulement quand la lecture démarre depuis la fenêtre Podcasts.
• Correction de la finalisation des livres audio MP3 avec SAPI4 et SAPI5 : le fichier final est désormais finalisé correctement, ce qui évite les fichiers incomplets ou fragiles après de longues exportations.
• Ajout d’une barre de progression explicite pour la phase de finalisation dans tous les modes de création de livres audio : après la phase de création, Sonarpad annonce et affiche la finalisation avec une progression visible.
• Correction d’un bug des voix de dialogue : les paramètres vitesse/tonalité/volume de la première et de la seconde voix de dialogue sont désormais correctement appliqués pendant la synthèse.
• Amélioration de la détection d’encodage pour les fichiers japonais `.txt` : ajout d’un fallback Shift_JIS/CP932 sûr en cas de mojibake, tout en préservant le comportement existant pour UTF/diacritiques/chinois.
• Refactorisation interne de sûreté : conversion vers des implémentations safe lorsque possible et réduction drastique des lignes de code unsafe.
Version 0.6.7 – 2026-03-02
Améliorations
• Ora il programma riesce a gestire Sostituisci tutto in modo massivo su file grandi con un gran numero di sostituzioni.
• Mise à jour de la traduction polonaise grâce à DJ Graco.
• Ajout de la traduction lituanienne.
• Ajout de la traduction chinoise.
• À partir de maintenant, des versions bêta fréquentes seront publiées dans la section Releases du projet, afin que les utilisateurs puissent tester les nouvelles modifications avant la prochaine version stable.
• Ajout du raccourci `Ctrl+.` pour insérer le caractère de points de suspension (…).
• Amélioration de la prise en charge des chapitres de podcast : la navigation entre chapitres est désormais plus fiable, y compris pour les épisodes directs/streaming où les chapitres ne sont pas intégrés au fichier MP3, grâce à un fallback sur les métadonnées de chapitres du flux/URL lorsqu’elles sont disponibles. Ajout des raccourcis `Ctrl+Alt+PageUp` (chapitre précédent) et `Ctrl+Alt+PageDown` (chapitre suivant).
• Réorganisation des dossiers de sortie dans `Documents\\Sonarpad` : les fichiers sont désormais enregistrés dans des sous-dossiers dédiés `audiobooks`, `documents`, `recordings` et `media`, avec migration automatique depuis les anciens chemins.
• Amélioration de la prise en charge des fichiers texte très volumineux (jusqu’à 60 Mo) : ouverture et navigation ligne par ligne plus fluides, en particulier avec les lecteurs d’écran.
• Guides mis à jour pour toutes les langues et ressources de localisation actualisées dans toute l’application, y compris les textes de dons et les traductions de l’installateur NSIS (nouvelles chaînes en chinois simplifié et lituanien, ainsi que finalisation de la traduction ukrainienne du setup).
• Ajout de la prise en charge globale du proxy réseau (HTTP/HTTPS et SOCKS5/SOCKS5H) pour les fonctions en ligne, avec validation à l'enregistrement des options : les proxys invalides sont signalés puis supprimés automatiquement.
• Ajout d'une nouvelle fonction dans Outils : « Lire l'audio en streaming... », permettant de coller une URL (YouTube ou lien média direct), de choisir le format de sortie et le profil qualité/débit binaire (y compris qualité/débit d'origine pour MP3 et MP4), puis de lancer la lecture dans le lecteur audio de Sonarpad.
• Ajout de la prise en charge de la touche multimédia système Lecture/Pause (casque/clavier) : elle contrôle désormais à la fois la lecture multimédia et la pause/reprise de la lecture de texte (priorité au lecteur multimédia lorsque les deux sont actifs).
• Ajout d'une nouvelle entrée dans Fichier > Fichiers récents : « Vider les fichiers récents » pour effacer rapidement la liste des documents récents.
• Extension des options de débit binaire dans Convertir l’audio et dans les paramètres d’enregistrement de podcast : ajout de valeurs plus basses (64/96 kbps) et extension du MP3 jusqu’à 320 kbps, avec validation et gestion de l’encodeur harmonisées.
• Extension des options de découpage de livre audio par durée jusqu’à 60 minutes.
• Amélioration du découpage en parties : il est désormais possible de saisir manuellement le nombre de parties, avec validation de 1 à 100.
• Ajout du nouveau mode Affichage > Lecture seule pour éviter les modifications accidentelles tout en conservant une lecture et une navigation complètes des documents.
• Ajout d’une barre de progression accessible pendant les mises à jour du programme, afin que les lecteurs d’écran puissent suivre en temps réel l’avancement du téléchargement.
• Ajout d’une nouvelle barre d’état discrète dans la fenêtre principale avec le nombre de caractères, de mots et la position ligne/colonne (par exemple : "Caractères (avec espaces) : 11. | Mots : 2. | Ln 1, Col 12"), sans perturber le focus NVDA.
• Ajout d’une nouvelle option dans le menu Affichage pour le retour à la ligne, afin d’activer ou désactiver rapidement l’habillage sans ouvrir les Options.
• Ajout, dans Édition > Texte, de nouvelles actions pour augmenter/réduire le retrait, avec les raccourcis Ctrl+Shift+. (indenter) et Ctrl+Shift+, (désindenter), car lorsque « Afficher les voix dans l’éditeur » est activé, la touche Tab est réservée à la navigation du panneau des voix.
• Ajout de l'affichage localisé de la date et de l'heure dans les articles RSS et les épisodes de podcast, avec un format adapté à la langue de l'interface.
• Ajout d'une nouvelle action dans le menu contextuel RSS pour partager l'article sélectionné par e-mail.
• Ajout d'options granulaires de confirmation de suppression dans Options > RSS et podcast : pour RSS (flux/article/les deux/aucun) et pour Podcasts (podcast/épisode/les deux/aucun).
• Ajout d'une copie rapide RSS configurable avec Ctrl+C (Options > RSS et podcast) : copier le titre, l'URL, le contenu de l'article ou l'ensemble.
• Unification du flux RSS : « Ajouter une source » accepte désormais à la fois les URL de flux et les mots-clés (avec génération automatique d'un flux Google News), sans recherche séparée.
• Un appui sur Ctrl+A annonce désormais la fin de l'action pour un retour plus clair avec les lecteurs d'écran.
• Ajout de Shift+F3 pour "Rechercher le précédent" dans le menu Édition, en complément de F3 "Rechercher le suivant".
• Amélioration du message de remplacement avec une gestion correcte du singulier/pluriel (par ex. « 1 remplacement effectué » vs « 2 remplacements effectués »).
• Ajout dans la fenêtre du dictionnaire d'une sélection de langue de recherche, avec Auto (langue de l'interface) par défaut et possibilité de choix manuel.
• Ajout d'un nouvel onglet Raccourcis dans les Options pour personnaliser les combinaisons de touches, avec détection des conflits et avertissement lorsqu'un raccourci est déjà attribué à une autre action.
• Ajout d’un support initial des paramètres en ligne de commande : `-h`/`--help` affichent l’aide rapide et `--version` affiche la version du programme.
• Réglage manuel de la vitesse et de la tonalité rendu plus clair : les champs manuels utilisent désormais une échelle centrée sur 100, où 100 correspond à la valeur normale.
• Amélioration de la sélection des voix Microsoft dans Options > Voix et dans le panneau des voix de l’éditeur : ajout d’une liste de langue localisée pour filtrer les voix par langue, tout en conservant le mode « voix multilingues uniquement » comme une liste unique non séparée par langue (liste de langue masquée lorsqu’il est activé).
• Ajout de la configuration de la voix pour les dialogues dans Options > Voix avec navigation complète au clavier (Tab), en réutilisant le même modèle de voix que l’interface principale (moteur, filtre de langue Edge, voix et vitesse/tonalité/volume avec libellés) ; ajout également d’une deuxième voix de dialogue optionnelle avec les mêmes contrôles (moteur, filtre de langue Edge, voix, vitesse/tonalité/volume) pour alterner les dialogues ; les règles de dialogue sont enregistrées dans la configuration `.ini`, sans modifier le texte du document.
• Amélioration de l'étiquette Annuler : l'entrée Édition > Annuler affiche désormais l'action qui sera annulée (par exemple édition de texte, commenter/décommenter des lignes ou insertion de balise de voix), tout en restant indisponible lorsqu'il n'y a rien à annuler.
Améliorations
• Ajout de la gestion des profils vocaux dans Options > Voix : il est possible d'ajouter, de renommer et de supprimer des profils.
• Ajout dans le menu contextuel des articles RSS de l'action pour ajouter un article aux favoris.
• La source RSS "Favoris" peut être supprimée et est recréée automatiquement lors du prochain ajout d'un article aux favoris.
• Ajout de raccourcis clavier RSS pour déplacer les sources vers le haut/le bas : `Ctrl+Shift+Flèche haut` et `Ctrl+Shift+Flèche bas`.
Corrections de bugs
• Correction de l'ouverture des fichiers RTF : les fichiers `.rtf` sont désormais extraits et affichés en texte lisible, au lieu du balisage RTF brut (ex. `{\\rtf1...}`).
• Correction de l'ouverture des fichiers texte chinois encodés en GB18030/GBK : Sonarpad les détecte et les décode correctement, évitant le texte illisible (mojibake).
• Amélioration de la création des livres audio M4B avec métadonnées de chapitres et marqueurs de chapitre ; correction du problème « chipmunk » (voix trop aiguë/rapide) dans les fichiers M4B générés.
• Correction de l’interface bitrate dans la fenêtre d’enregistrement de livre audio : suppression des libellés italiens codés en dur et ajout de 64 kbps dans les débits sélectionnables.
• Correction de « Enregistrer tout » (Ctrl+Shift+S) : tous les documents ouverts modifiés sont désormais détectés de façon fiable (y compris les onglets nouveaux/non enregistrés) et l’enregistrement global fonctionne correctement pour chacun, avec « Enregistrer sous » lorsque nécessaire.
• Correction de l'ordre des articles RSS Google News : lorsque la date est disponible, les articles sont désormais affichés du plus récent au plus ancien.
• Correction de l'association des étiquettes NVDA dans la fenêtre du dictionnaire : le champ de recherche et la liste de langue annoncent désormais la bonne étiquette.
• Correction de la navigation clavier dans la fenêtre Propriétés RSS/Podcast : Tab/Maj+Tab atteignent désormais le bouton OK, Entrée active OK, Échap ferme la fenêtre en toute sécurité et le focus revient correctement à la liste RSS/Podcast.
• Correction de l'historique d'annulation RSS/Podcast : Ctrl+Z prend désormais en charge une annulation multi-niveaux pour les suppressions (articles/épisodes et sources), et plus seulement la dernière action.
• Amélioration des annonces de suppression RSS/Podcast avec des messages explicites (flux RSS supprimé, article RSS supprimé, épisode de podcast supprimé).
• Amélioration du focus après suppression/annulation dans RSS/Podcast : en RSS, le premier flux est sélectionné de manière fiable si nécessaire, et les répétitions d'annonces du lecteur d'écran ont été réduites pendant la resélection différée.

Version 0.6.6 – 2026-02-13
Améliorations
• Ajout de « Formatage automatique pour TTS » dans le menu Édition pour préparer rapidement le texte à la lecture vocale (suppression markdown/guillemets et recomposition des lignes coupées).
• Amélioration de l’insertion des balises de voix : lorsqu’un texte est sélectionné, les balises sont désormais appliquées correctement aussi bien sur une seule ligne que sur une sélection multiligne.
• Ajout d’une option dans les paramètres Audio pour choisir le dossier par défaut d’enregistrement des livres audio (par défaut : Documents\\Sonarpad Audiobooks).
• Dans la fenêtre d’enregistrement du livre audio, lorsque le découpage est actif, ajout d’une nouvelle option (activée par défaut) pour créer un sous-dossier dédié aux parties générées.
• L’export des livres audio enregistre désormais les MP3 en stéréo avec un bitrate choisi par l’utilisateur pour les voix Edge, SAPI5 et SAPI4.
• Ajout de la prise en charge des voix SAPI5 32 bits via bridge, afin d’utiliser aussi les voix disponibles uniquement dans les moteurs 32 bits.
• Réorganisation des fonctions vocales dans un menu dédié « Voix et audio » et ajout/clarification de l’option « Convertir l’audio », utile pour convertir tout média pris en charge en MP3, AAC, OGG, Opus, FLAC, WAV et AIFF.
• Ajout de la suppression d’articles RSS individuels et d’épisodes de podcast individuels (touche Suppr + menu contextuel avec confirmation), sans supprimer toute la source RSS/podcast, avec annulation de la dernière suppression (article/épisode individuel ou source RSS/podcast complète).
• Ajout de l'export des flux RSS en OPML dans la fenêtre RSS, afin de sauvegarder et réimporter facilement les sources actuelles.
• Ajout de la fonction « Rechercher un flux RSS par mot-clé » dans la fenêtre RSS : en saisissant un mot-clé, l'URL RSS Google News est générée automatiquement et la fenêtre d'ajout de source s'ouvre préremplie, afin de créer un flux thématique en une seule étape.
• Ajout de la traduction serbe grâce à Mila Kuran.
• Ajout de la traduction ukrainienne grâce à Ivan Shtefuriak.
• Ajout de l'ouverture multiple de fichiers média : en ouvrant plusieurs fichiers à la fois, une file de lecture est créée au lieu de remplacer le fichier en cours.
• Ajout de raccourcis de déplacement variable pendant la lecture : avec une base de 1 minute, Gauche/Droite avance-recule de 60 s, Maj+Gauche/Droite de 20 s et Ctrl+Gauche/Droite de 3 minutes.
• Ajout des raccourcis piste précédente/suivante dans le lecteur : Ctrl+PageUp et Ctrl+PageDown.
• Ajout de l'option « Réinitialiser le volume » et regroupement des actions de réinitialisation dans un sous-menu dédié « Réinitialiser » dans Lecture, avec « Réinitialiser la vitesse » et « Réinitialiser la tonalité ».
• Amélioration de l'installateur : setup.exe permet désormais de choisir entre associer tous les types de fichiers pris en charge ou sélectionner manuellement les extensions ; le MSI propose aussi une sélection extension par extension dans l'arborescence des fonctionnalités (valeur par défaut inchangée : tout activé).
• Ajout du nouveau menu « Fenêtre » avec l'option « Documents ouverts... » pour basculer rapidement vers n'importe quel fichier actuellement ouvert.
• Mise à jour de l'option Affichage > Police : le sélecteur complet a été remplacé par un sous-menu rapide de polices courantes (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia), tout en conservant la taille de texte actuelle.
• Amélioration de la lecture RSS/podcasts avec deux annonces distinctes : les nœuds de source annoncent « nouveaux éléments » lorsqu’un flux/podcast a des nouveautés, tandis que les articles RSS et épisodes de podcast individuels annoncent « non lu »/« non joué » ; ce comportement peut être désactivé dans les Options.
Corrections de bugs
• Correction de l’extraction de texte EPUB pour les livres contenant des commentaires HTML inline (<!-- ... -->) : le texte des chapitres est désormais correctement analysé au lieu d’être partiellement ou totalement ignoré.
• Correction du dictionnaire Wiktionary en espagnol et de la gestion du cache : des mots comme « agua » sont maintenant trouvés correctement et les anciennes entrées « mot introuvable » ne sont plus réutilisées.
• Correction de l’encodage lors de l’import d’articles RSS pour certaines sources espagnoles (ex. El Mundo) : les accents et le « ñ » sont désormais correctement conservés dans l’éditeur temporaire.
• Correction du décodage ANSI des fichiers d’Europe centrale (ex. tchèque/polonais) : Sonarpad distingue désormais mieux UTF-8 et ANSI et choisit la bonne page de codes (y compris Windows-1250), évitant les diacritiques corrompus.
• Correction de la persistance des sources RSS avec paramètres d’URL (ex. `rss.aspx?c=...`) : ces flux sont maintenant correctement sauvegardés et restaurés après redémarrage de Sonarpad.
• Correction de l’ouverture des fichiers pointeurs Google Drive (`.gdoc`, `.gsheet`, `.gslides`) depuis le menu contextuel de l’Explorateur : si la lecture directe échoue avec « Incorrect function (os error 1) », Sonarpad utilise désormais un fallback shell-open et le document s’ouvre correctement.
• Correction de la lecture des fichiers Excel legacy `.xls` (Excel 2010) : les anciens fichiers binaires sont maintenant détectés/décodés correctement au lieu d’afficher du texte corrompu (ex. `ÐÏ_à¡±...`).
• Correction du flux d’annonce du correcteur orthographique : les fautes sont désormais réannoncées lors d’une relecture ultérieure du texte, et la même faute est de nouveau signalée si elle est supprimée puis retapée.
• Correction des opérations de texte par ligne (ex. Ctrl+Q / Ctrl+Shift+Q, trier/inverser/lignes uniques/fusionner les lignes) : en sélectionnant une seule ligne avec Maj+Flèche bas, les lignes adjacentes ne sont plus fusionnées ni tronquées.
• Correction du comportement multilignes des opérations de texte par ligne (Ctrl+Q / Ctrl+Shift+Q et outils associés) : lorsque RichEdit fournit des séparateurs de ligne en CR seul, ils sont désormais normalisés correctement et toutes les lignes sélectionnées sont traitées sans couper le premier caractère.
• Extension de la normalisation d’entrée TTS pour les symboles visibles d’espace/tabulation/saut de ligne (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), qui pouvaient provoquer des répétitions de paragraphes avec les voix multilingues.
• Affinement de la sanitisation du texte Edge TTS avec une pipeline unique de validation : normalisation des espaces étranges/invisibles, compactage des longues séquences de ponctuation (comme "...", "!!!", "???") et suppression des segments composés uniquement de ponctuation pour éviter les boucles de lecture.
• Correction de l’annonce du temps de lecture (Ctrl+I) pour les flux MP3/podcast : le temps courant est désormais borné à la durée de la piste, et la lecture est arrêtée automatiquement si la position dépasse la fin.
• Amélioration de la couverture de localisation de l’installateur : setup.exe inclut désormais aussi le tchèque, le polonais, le français et le serbe, tandis que le MSI reste un paquet unique en-US pour éviter la confusion en release.
• Correction du nettoyage à la désinstallation des entrées du menu contextuel : « Ouvrir avec Sonarpad » est maintenant supprimé de façon fiable, y compris dans des scénarios de registre legacy.
• Correction de la fiabilité pause/reprise en SAPI5 : la pause avec F4 fonctionne désormais correctement et la reprise revient au point attendu au lieu de redémarrer depuis le début.
• Correction du flux pause + recherche + reprise en lecture média : après une pause puis un déplacement avec Gauche/Droite, la touche Espace reprend désormais de manière fiable à la position courante au lieu de s'arrêter ou de repartir du début.

Version 0.6.5 – 2026-02-07
Améliorations
• Traduction espagnole améliorée grâce à Arturo Fernandez Rivas.
• Les imports RSS utilisent désormais un onglet temporaire dédié (titre localisé) ; Enregistrer sous le convertit en document normal.
• Les messages du lecteur d’écran sont désormais également envoyés à JAWS lorsqu’il est disponible.
Corrections de bugs
• La lecture depuis le curseur (F5) démarre exactement au niveau du curseur. Avant, elle pouvait commencer quelques lignes au-dessus car l’offset du curseur ne correspondait pas aux positions CRLF/UTF-16.
• Correction d’un problème de redessin : en tapant sur une sélection, le texte précédent pouvait disparaître jusqu’au déplacement de la sélection.
• Correction du parsing des chapitres EPUB : les pages de couverture ou uniquement images ne génèrent plus de lecture de CSS (ex. « padding ») ni de titres « Sconosciuto ».
• Correction d’un échec lors du découpage par durée des EPUB : Edge TTS pouvait échouer avec des blocs vides ou trop longs ("Edge audio not sent").
• La fenêtre d’enregistrement de podcast est maintenant indépendante : vous pouvez utiliser l’éditeur pendant l’enregistrement.
• Les articles RSS décodent désormais les entités HTML (par ex. &quot;, &amp;, &lt;, &gt;).
• Enregistrer/Enregistrer sous propose désormais le nom du fichier existant lors de l’enregistrement de formats non réécrivables (ex. EPUB), au lieu de la première ligne.
• Correction d’un problème où les podcasts avec de nouveaux épisodes n’étaient pas annoncés comme non joués, et renommage de « Non écouté » en « Non joué » pour un libellé plus professionnel.

Version 0.6.4 – 2026-02-05
Améliorations
• Le programme a été renommé en Sonarpad pour mettre davantage l'accent sur le son et l'audio, qui sont la clé de ce programme.
• Ajout de la sélection des pistes audio dans le menu Lecture pour les fichiers multimédias avec plusieurs pistes audio (ex. MKV avec plusieurs langues).
• Les podcasts indiquent maintenant clairement ceux non écoutés avec le préfixe « Non écouté » avant le nom.
• Nouveau système de balises pour changer la voix dans le texte. Exemples :
  - Voix Microsoft (Edge) : <voice edge it-IT-IsabellaNeural>Bonjour</voice>
  - Voix SAPI5 : <voice sapi5 Microsoft Helena Desktop>Bonjour</voice>
  - Voix SAPI4 : <voice sapi4 #1>Bonjour</voice>
  - Avec vitesse/tonalité/volume : <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Bonjour</voice>
• Catégories de podcasts enrichies.
• Ajout d’une option dans le menu contextuel pour créer un livre audio à partir de la sélection.
• Ajout du découpage des livres audio par durée, avec la possibilité de choisir le nom du premier fichier.
• Libellé de l’auteur localisé dans la lecture des articles (ex. « par », « by », « di »).
• Ajout d’options d’indentation (tabulations/espaces avec largeur) et de Tab/Maj+Tab pour indenter/désindenter les lignes sélectionnées.
• Correction du nettoyage Markdown : gestion des puces « * » lorsque la conservation des listes est désactivée.
Corrections de bugs
• Corrigé un bug où les livres audio SAPI4 pouvaient être créés différemment de ce qui était attendu.
• Fenêtre Rechercher dans les fichiers : Entrée sur un résultat ouvre maintenant à la position correcte de l’extrait et Échap retourne aux résultats.
• Fenêtre Options : ajustement du layout visuel des onglets Général, Voix, Éditeur et Audio pour éviter des contrôles manquants ou coupés.
• Correction d’un problème de signets lors du changement de vitesse de lecture.
• Correction d’un problème avec Podcast Index et les catégories qui ne s’affichaient pas correctement.
• Correction du problème de l’apostrophe qui coupait la lecture : plus de lecture séparée pour les dialogues, utilisation des balises de voix.

Version 0.6.3 – 2026-01-30
Améliorations
• Amélioration de la détection du microphone.
• Ajout de la lecture instantanée pour tous les formats.
Corrections de bugs
• Correction du plantage dans la fenêtre des catégories de podcasts.

Version 0.6.2 – 2026-01-30
Nouvelles fonctionnalités
• Ajout de la prise en charge de l'exécution de fichiers (Shift+F5). Les utilisateurs peuvent sélectionner un interpréteur (par exemple, python) dans les Options, le rechercher sur l'ordinateur, et appuyer sur Shift+F5 exécute le script actuel. Les fichiers HTML s'ouvrent dans le navigateur.
• Ajout de la prise en charge des fichiers pointeurs Google Docs (.gdoc, .gsheet, .gslides), qui s'ouvrent automatiquement dans le navigateur par défaut.
• Ajout de la prise en charge du format de livre audio M4B (Apple/AAC).
• Ajout de l'option "Afficher les épisodes" dans le menu contextuel des résultats de recherche de podcasts pour parcourir et lire des épisodes sans s'abonner.
• Ajout de la fonctionnalité "Aller à la ligne" (menu Édition ou Ctrl+J) pour accéder rapidement à un numéro de ligne spécifique.
• Ajout d'options de menu contextuel pour ordonner les flux RSS et les podcasts (alphabétiquement ou par date).
• Ajout de flux RSS vietnamiens par défaut.
• Ajout d'une case de test du microphone dans la boîte de dialogue d'enregistrement pour vérifier les niveaux avant de commencer.
• Ajout de "Afficher la description" pour les épisodes de podcast dans le menu contextuel.
• Ajout de la prise en charge des formats audio/vidéo étendus via FFmpeg : mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Ajout de la prise en charge de la lecture synchronisée des sous-titres (srt, vtt, ass, sub, sbv, lrc, smi) avec NVDA ou la voix sélectionnée. Le programme recherche un fichier de sous-titres portant le même nom que le fichier multimédia. Ajout des options "Importer des sous-titres" et "Supprimer les sous-titres" dans le menu Lecture pour les fichiers aux noms différents.
• Ajout d'associations de fichiers pour tous les nouveaux formats audio/vidéo pris en charge dans le menu contextuel "Ouvrir avec Sonarpad".
• Ajout d'un paramètre de réglage de la hauteur tonale pour n'importe quel fichier.
• Ajout d'une option dans les Paramètres généraux pour activer ou désactiver les rapports d'erreurs anonymes. Ajout d'une entrée dans le menu Aide pour créer un fichier ZIP de diagnostic.
• Ajout d'une option pour utiliser une voix différente pour les dialogues, à la fois pour la lecture en direct et la création de livres audio.
• Ajout d'un navigateur de catégories de podcasts pour explorer les podcasts par catégorie (affaires, art, sport, etc.).
Améliorations
• L'ouverture d'un fichier audio/vidéo depuis l'Explorateur ouvre désormais directement la vue du lecteur au lieu de l'éditeur de texte.
• Suppression de l'invite OCR pour les PDF inaccessibles ; l'OCR est désormais effectué automatiquement pour améliorer la vitesse et l'expérience utilisateur.
• Amélioration du terminal accessible : la lecture NVDA mémorise désormais la dernière ligne lue pour une meilleure continuité.
• SAPI 4 : La création de livres audio est désormais entièrement parallélisée et presque instantanée. Ajout d'une invite pour choisir le nombre de processus simultanés.
• SAPI 4 : Élimination du goulot d'étranglement de la conversion WAV vers MP3 en convertissant les morceaux en parallèle pendant la synthèse.
• SAPI 4 : Amélioration de la gestion des erreurs et du nettoyage automatique des fichiers temporaires.
• Boîte de dialogue Rechercher : Renommage de "Regex" en "Expression régulière" pour plus de clarté et ajout des traductions manquantes pour les options de recherche.
• Livres audio M4B : Meilleure gestion de la sortie ; la division par parties/marqueurs produit désormais un seul fichier M4B avec des métadonnées de chapitres incluant le titre et l'auteur.
• Lecteur : Correction de la précision des signets et de l'annonce du temps lorsque la vitesse de lecture n'est pas de 1.0x.
• Restauration de la navigation Ctrl+Tab et Ctrl+Maj+Tab dans les Options.
• Ajout d'une option dans le menu Lecture pour réinitialiser instantanément la vitesse à la normale (1.0x).
• Mise à jour de toutes les dépendances vers les dernières versions pour de meilleures performances et stabilité.
• Intégration de FFmpeg avec chargement dynamique de DLL pour assurer la compatibilité sans bloquer le démarrage.
• Mise à jour des filtres de téléchargement de podcasts pour inclure les nouveaux formats audio/vidéo.
• Empêchement de Ctrl+S d'enregistrer les fichiers audio/vidéo pour éviter la corruption.
• Amélioration de l'importation des transcriptions YouTube, la rendant plus robuste et résiliente.
• Amélioration de la robustesse de la division des livres audio en parties, garantissant qu'aucun texte n'est perdu.
• L'installateur est désormais entièrement multilingue, prenant en charge l'italien, l'anglais, l'espagnol, le portugais, le suédois et le vietnamien en fonction de la langue du système de l'utilisateur. L'anglais est la valeur par défaut pour les systèmes non pris en charge.
• Catégories de podcasts : appuyer sur Entrée sur une catégorie confirme désormais la sélection (équivalent au bouton OK).
• Amélioration du système de détection des blocages pour éviter les faux positifs lorsque des boîtes de dialogue modales sont ouvertes (messages d'erreur, "texte non trouvé").
Corrections
• Correction d'un bug où le journal des modifications ne s'ouvrait pas au démarrage.
• Correction d'un bug où l'invite OCR n'apparaissait pas pour les PDF inaccessibles ouverts depuis l'Explorateur.
• Correction d'un bug au démarrage pouvant entraîner une perte de focus ou la fermeture de la fenêtre immédiatement après l'ouverture.
• Correction d'un bug critique dans la recherche par expression régulière empêchant de trouver du texte, y compris des problèmes avec la "Recherche circulaire" et l'option "Le point équivaut à une nouvelle ligne" avec les fins de ligne Windows.
Localisation
• Ajout de la traduction en polonais.
• Ajout de la traduction en français.
• Ajout de la traduction en tchèque (merci à Radek Žalud et Jiri Holzinger).

Version 0.6.1 – 2026-01-20
Corrections
• Correction d'un bug où l'activation de "Afficher les voix dans l'éditeur" provoquait l'arrêt de la lecture du podcast.
• Correction d'un problème où certains podcasts ne pouvaient pas être ajoutés via URL car l'URL était tronquée.
• Correction d'un bug où les URL normales ne pouvaient plus être ajoutées dans la fonctionnalité de flux RSS.
• Correction d'un problème où l'option de langue de Wikipédia était affichée plusieurs fois dans différents onglets de paramètres.
• Suppression de la création de fichiers de débogage qui étaient générés incorrectement même en mode release.
Améliorations
• Amélioration de la prise en charge des voix Microsoft, qui utilisent désormais une méthode de lecture dédiée avec un agent utilisateur différent.
• Ajout de la prise en charge des fichiers MP4.

Version 0.6.0 – 2025-01-20
Nouvelles fonctionnalités
• Ajout du correcteur orthographique. Depuis le menu contextuel, les utilisateurs peuvent vérifier si le mot actuel est correct et, sinon, obtenir des suggestions d'orthographe.
• Ajout de l'importation et de l'exportation de podcasts via des fichiers OPML.
• Ajout de la prise en charge de la recherche Podcast Index en plus d'iTunes. Les utilisateurs peuvent saisir leur clé API et leur secret gratuits (générés uniquement à l'aide d'une adresse e-mail).
• Ajout de la prise en charge des voix SAPI4, tant pour la lecture en temps réel que pour la création de livres audio.
• Ajout du repli automatique OCR pour les PDF non accessibles : lorsqu'aucun texte extractible n'est trouvé, le document est reconnu via OCR.
• Ajout de la prise en charge du dictionnaire utilisant le Wiktionnaire. Appuyer sur la touche Applications affiche les définitions et, lorsqu'ils sont disponibles, les synonymes et les traductions dans d'autres langues.
• Ajout de l'importation d'articles Wikipédia avec recherche, sélection de résultats et importation directe dans l'éditeur.
• Ajout du raccourci Maj+Entrée dans le module RSS pour ouvrir un article directement sur le site web d'origine.
Améliorations
• La sélection du microphone est désormais toujours respectée par l'application.
• Dans la fenêtre des podcasts, appuyer sur Entrée sur un épisode annonce désormais immédiatement "chargement" via NVDA pour confirmer l'action.
• Dans les résultats de recherche de podcasts, appuyer sur Entrée s'abonne désormais au podcast sélectionné.
• Correction et amélioration des étiquettes pour les raccourcis Ctrl+Maj+O et Podcast Ctrl+Maj+P.
• La vitesse et le volume de lecture sont désormais enregistrés dans les paramètres et persistent pour tous les fichiers audio.
• Ajout d'un dossier cache dédié pour les épisodes de podcast. Les utilisateurs peuvent conserver les épisodes via "Garder le podcast" dans le menu Lecture. Le cache est automatiquement nettoyé lorsqu'il dépasse la taille définie par l'utilisateur (Options → Audio).
• Amélioration significative de la récupération des articles RSS en utilisant l'emprunt d'identité libcurl avec des profils Chrome et iPhone, assurant une compatibilité avec ~99% des sites.
• Ajout de l'état lu / non lu pour les articles RSS, avec une indication claire dans la liste RSS.
• Tout remplacer signale désormais le nombre de remplacements effectués.
• Ajout d'un bouton Supprimer le podcast lors de la navigation dans la bibliothèque de podcasts à l'aide de Tab.
Corrections
• Suppression de l'entrée redondante "mise à jour en attente" du menu Aide (les mises à jour sont déjà gérées automatiquement).
• Correction d'un bug où appuyer sur Ctrl+S sur un fichier MP3 ouvert enregistrait et corrompait le fichier.
• Correction d'un problème d'interface utilisateur où "Audiolivres par lots" était affiché comme "(B)… Ctrl+Maj+B" (suppression de l'étiquette redondante).
• Correction des guillemets intelligents : lorsqu'ils sont activés, les guillemets normaux sont désormais correctement remplacés par des guillemets intelligents.
• Correction d'un bug où l'utilisation de "Aller au signet" réinitialisait la vitesse de lecture à 1.0.
• Correction d'un problème où les épisodes de podcast déjà téléchargés étaient retéléchargés au lieu d'utiliser la version en cache.
Raccourcis clavier
• F1 ouvre désormais le Guide d'aide.
• F2 vérifie désormais les mises à jour.
• F7 / F8 sautent désormais à l'erreur d'orthographe précédente ou suivante.
• F9 / F10 basculent désormais rapidement entre les voix favorites.
Améliorations développeur
• Les erreurs ne sont plus ignorées silencieusement : tous les modèles let _ = ont été supprimés, et les erreurs sont désormais gérées explicitement.
• Le projet ne compile plus s'il y a des avertissements.
• Les implémentations personnalisées telles que les aides de style strlen / wcslen ont été supprimées.
• La gestion des DLL a été nettoyée et consolidée autour de libloading.
• Les aides d'analyse d'octets manuelles ont été supprimées au profit des méthodes standard.

Version 0.5.9 - 2025-01-13
Nouvelles fonctionnalités
• Ajout de la réorganisation RSS depuis le menu contextuel (monter/descendre/vers la position) avec vérification de position invalide.
• Ajout d'un menu contextuel d'article avec ouverture du site d'origine et partage via WhatsApp, Facebook et X.
• Ajout du raccourci Échap pour revenir des articles importés à la liste RSS.
• Ajout du mode podcast : recherche, abonnement, écoute ; réorganisation des abonnements ; Échap arrête la lecture et revient à la liste ; Entrée sur un épisode démarre la lecture.
• Ajout du contrôle de la vitesse de lecture pour les podcasts et les fichiers MP3.
• Ajout de Ctrl+T pour sauter à un temps spécifique.
• Ajout d'un bouton d'aperçu vocal après le combo de volume.
• Ajout de la recherche et du remplacement par regex (style Notepad++).
• Ajout de l'importation RSS depuis des fichiers OPML et TXT.
• Ajout d'une option pour activer "Ouvrir avec Sonarpad" dans l'Explorateur de fichiers, y compris pour les versions portables.
Améliorations
• Amélioration de la sélection de la vitesse/hauteur/volume de la voix, respectant les limites maximales TTS.
• Diverses améliorations RSS pour télécharger tous les articles sans déplacer le focus NVDA pendant les mises à jour.
• Amélioration de la lecture audio avec un menu dédié, annonce du temps Ctrl+I, et volume jusqu'à 300%.
• Ajout de raccourcis manquants pour certaines fonctions.
• Réorganisation du menu Édition avec un sous-menu de nettoyage de texte.
• Réorganisation des Options en onglets, avec navigation Ctrl+Tab et Ctrl+Maj+Tab.
• Le lecteur RSS télécharge désormais le contenu complet de l'article, correspondant à la vue du navigateur.
Corrections
• Correction du nettoyage Markdown supprimant les numéros au début des lignes.
• Correction de AltGr+Z déclenchant l'annulation.
• Correction de l'annulation de l'enregistrement de livre audio pour qu'il s'arrête rapidement.
Localisation
• Ajout de la traduction vietnamienne (merci à Anh Đức Nguyễn).

Version 0.5.8 - 2026-01-10
Nouvelles fonctionnalités
• Ajout du contrôle du volume pour le microphone et l'audio système lors de l'enregistrement de podcasts.
• Ajout d'une nouvelle fonctionnalité pour importer des articles depuis des sites web ou des flux RSS, y compris les flux les plus importants pour chaque langue.
• Ajout d'une fonction pour supprimer tous les signets du fichier actuel.
• Ajout d'une fonction pour supprimer les lignes dupliquées et les lignes consécutives dupliquées.
• Ajout d'une fonction pour fermer tous les onglets ou fenêtres sauf l'actuel.
• Ajout d'une entrée Dons dans le menu Aide pour toutes les langues.
Améliorations
• Amélioration du terminal accessible pour éviter certains plantages.
• Amélioration et correction des touches d'accès et des raccourcis clavier dans toute l'application.
• Correction d'un problème où la fermeture de la fenêtre de lecture audio n'arrêtait pas la lecture.
• Ajout de boîtes de dialogue de confirmation pour les actions importantes (ex: supprimer les lignes dupliquées, supprimer les traits d'union de fin de ligne, supprimer tous les signets).
• Ajout de la possibilité de supprimer des flux/sites RSS de la bibliothèque en les sélectionnant et en appuyant sur Suppr.
• Ajout d'un menu contextuel dans la fenêtre RSS pour modifier ou supprimer des flux/sites RSS.
• Suppression du paramètre pour déplacer les paramètres vers le dossier actuel ; l'application gère désormais cela automatiquement en fonction de l'emplacement.

Version 0.5.7 - 2026-01-05
Nouvelles fonctionnalités
• Ajout de la fonctionnalité Audiolivres par lots pour convertir plusieurs fichiers/dossiers à la fois.
• Ajout de la prise en charge des fichiers Markdown (.md).
• Ajout de la sélection de l'encodage de fichier lors de l'ouverture de fichiers texte.
• Ajout d'une option dans le terminal accessible pour annoncer les nouvelles lignes avec NVDA.
Améliorations
• L'enregistrement de livre audio sauvegarde désormais nativement en MP3 lorsqu'il est sélectionné.
• L'utilisateur peut désormais choisir la position de l'astérisque (*) "modifications non enregistrées" dans le titre de la fenêtre.
• Amélioration de la robustesse du système de mise à jour.
• Ajout de "Supprimer les traits d'union" dans le menu Édition pour corriger les fins de ligne OCR.

Version 0.5.6 - 2026-01-04
Corrections
  Amélioration de la recherche dans les fichiers pour que l'appui sur Entrée ouvre le fichier exactement à l'extrait sélectionné.
Améliorations
  Ajout de la prise en charge PPT/PPTX (ouvrir comme texte).
  L'ouverture de formats non textuels enregistre désormais en .txt pour éviter la corruption de formatage (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Ajout de l'enregistrement de podcast à partir du microphone et de l'audio système (menu Fichier, Ctrl+Maj+R).

Version 0.5.5 – 2026-01-03
Nouvelles fonctionnalités
• Ajout d'un terminal accessible optimisé pour les grandes sorties et les lecteurs d'écran (Ctrl+Maj+P).
• Ajout d'un paramètre pour enregistrer les paramètres utilisateur dans le dossier actuel (mode portable).
Corrections
• Amélioration des extraits de recherche dans les fichiers pour que l'aperçu reste aligné avec la correspondance.

Version 0.5.4 – 2026-01-03
Améliorations
• Correction de la normalisation des espaces (Ctrl+Maj+Entrée).
• Ajout de la prise en charge HTML/HTM (ouvrir comme texte).

Version 0.5.3 – 2026-01-02
Nouvelles fonctionnalités
• Ajout de la recherche dans les fichiers.
• Ajout de nouveaux outils de texte : Normaliser les espaces, Saut de ligne dur et Supprimer Markdown.
• Ajout des statistiques de texte (Alt+Y).
• Ajout de nouvelles commandes de liste dans le menu Édition :
• Ordonner les éléments (Alt+Maj+O)
• Garder les éléments uniques (Alt+Maj+K)
• Inverser les éléments (Alt+Maj+Z)
• Ajout de Citer / Retirer la citation des lignes (Ctrl+Q / Ctrl+Maj+Q).
Localisation
• Ajout de la localisation espagnole.
• Ajout de la localisation portugaise.
Améliorations
• Lorsqu'un fichier EPUB est ouvert, Enregistrer bascule désormais automatiquement vers Enregistrer sous et exporte le contenu en fichier .txt pour éviter la corruption de l'EPUB.

## 0.5.2 - 2026-01-01
- Ajout d'un journal des modifications.
- Ajout des options ouvrir avec Sonarpad et des associations de fichiers lors de l'installation.
- Amélioration de la localisation des messages.
- Ajout de la sélection de partie lors de l'utilisation de "Diviser le livre audio par texte".
- Ajout de l'importation de transcription YouTube.

## 0.5.1 - 2025-12-31
- Mises à jour automatiques avec confirmation.
- Améliorations de l'exportation de livres audio.
- Améliorations TTS.
- Menu Affichage et panneaux voix/favoris.
- Langue par défaut du système et améliorations de la localisation.
- CI et empaquetage Windows.

## 0.5.0 - 2025-12-27
- Refactorisation modulaire.
- Flux de travail de construction/empaquetage Windows.
- Correction de la navigation par TAB dans la fenêtre d'aide.

## 0.5 - 2025-12-27
- Changement de version préliminaire.

## 0.1.0 - 2025-12-25
- Version initiale.












