# Changelog

Versión 0.6.7 – 2026-02-13
Mejoras
• Ampliadas las opciones de bitrate en la conversión de audio y en la grabación de podcast: añadidos valores más bajos (64/96 kbps) y ampliado MP3 hasta 320 kbps, con validación y manejo del codificador alineados.
• Añadido el nuevo modo Ver > Solo lectura para bloquear ediciones accidentales del texto manteniendo la lectura y navegación completas de los documentos.
• Añadida una barra de progreso accesible durante las actualizaciones del programa, para que los lectores de pantalla puedan seguir en tiempo real el avance de la descarga.
• Añadida una nueva barra de estado discreta en la ventana principal con recuento de caracteres, palabras y posición línea/columna (por ejemplo: "Caracteres (con espacios): 11. | Palabras: 2. | Ln 1, Col 12"), sin interferir con el foco de NVDA.
• Añadida una nueva opción en el menú Ver para Ajuste de línea, que permite activar o desactivar rápidamente el ajuste sin abrir Opciones.
• Añadidas en Editar > Texto nuevas acciones para aumentar/reducir sangría, con atajos Ctrl+Shift+. (indentar) y Ctrl+Shift+, (desindentar), porque cuando “Mostrar voces en el editor” está activo la tecla Tab queda reservada para navegar el panel de voces.
• Añadida la visualización localizada de fecha y hora en artículos RSS y episodios de podcast, con formato adaptado al idioma de la interfaz.
• Añadida en el menú contextual RSS una nueva acción para compartir por correo electrónico el artículo seleccionado.
• Añadidas opciones granulares de confirmación de borrado en Opciones > RSS y podcast: para RSS (feed/artículo/ambos/ninguno) y para Podcasts (podcast/episodio/ambos/ninguno).
• Añadida copia rápida RSS configurable con Ctrl+C (Opciones > RSS y podcast): copiar titular, URL, contenido del artículo o todo junto.
• Unificado el flujo de RSS: “Añadir fuente” ahora acepta tanto URL de feed como palabras clave (generando automáticamente el feed de Google News), sin necesidad de una búsqueda separada.
• Al pulsar Ctrl+A ahora se anuncia la finalización de la acción para ofrecer una respuesta más clara en lectores de pantalla.
• Añadido Shift+F3 para "Buscar anterior" en el menú Editar, junto a F3 "Buscar siguiente".
• Mejorado el mensaje de reemplazo con singular/plural correcto (por ejemplo, “1 reemplazo realizado” frente a “2 reemplazos realizados”).
• Añadida en la ventana del diccionario la selección de idioma de búsqueda, con valor predeterminado Auto (idioma de la interfaz) y opción de ajuste manual.
• Añadida una nueva pestaña de Atajos en Opciones para personalizar combinaciones de teclas, con detección de conflictos y aviso cuando un atajo ya está asignado a otra acción.
• Mejorada la claridad del ajuste manual de velocidad y tono: los campos manuales ahora usan una escala centrada en 100, donde 100 corresponde al valor normal.
• Mejorada la selección de voces Microsoft tanto en Opciones > Voz como en el panel de voces del editor: se añadió un combobox de idioma localizado para filtrar voces por idioma, manteniendo el modo “solo voces multilingües” como una lista única sin división por idioma (con el combobox de idioma oculto cuando está activo).
• Mejorada la etiqueta de Deshacer: la opción Editar > Deshacer ahora muestra qué acción se va a deshacer (por ejemplo, edición de texto, comentar/descomentar líneas o inserción de etiqueta de voz), manteniéndose no disponible cuando no hay nada que deshacer.
Correcciones de errores
• Corregida la interfaz de bitrate en la ventana de guardado de audiolibros: se eliminaron textos hardcoded en italiano y se añadió 64 kbps entre los bitrates seleccionables.
• Corregido "Guardar todo" (Ctrl+Shift+S): ahora se detectan de forma fiable todos los documentos abiertos modificados (incluidas pestañas nuevas/sin guardar) y Guardar todo guarda cada uno correctamente, abriendo "Guardar como" cuando hace falta.
• Corregido el orden de artículos RSS de Google News: cuando hay fecha disponible, los artículos ahora se muestran del más reciente al más antiguo.
• Corregida la asociación de etiquetas en NVDA en la ventana del diccionario: el campo de búsqueda y el combo de idioma ahora anuncian la etiqueta correcta.
• Corregida la navegación de teclado en la ventana Propiedades de RSS/Podcast: Tab/Shift+Tab ahora llegan al botón Aceptar, Enter activa Aceptar, Esc cierra de forma segura y el foco vuelve correctamente a la lista RSS/Podcast.
• Corregido el historial de deshacer en RSS/Podcast: Ctrl+Z ahora admite deshacer multinivel para eliminaciones (artículos/episodios y fuentes), no solo la última acción.
• Mejorados los avisos de eliminación en RSS/Podcast con mensajes explícitos (RSS eliminado, artículo RSS eliminado, episodio de podcast eliminado).
• Mejorado el foco tras eliminar/deshacer en RSS/Podcast: en RSS se selecciona de forma fiable el primer feed cuando hace falta y se reducen repeticiones de anuncios del lector de pantalla durante la reselección diferida.

Versión 0.6.6 – 2026-02-13
Mejoras
• Añadida "Formateo automático para TTS" en el menú Editar para preparar rápidamente el texto para voz (elimina markdown/comillas y recompone líneas partidas).
• Mejorada la inserción de etiquetas de voz: cuando hay texto seleccionado, ahora las etiquetas se aplican correctamente tanto a una sola línea como a selecciones multilínea.
• Añadida una opción en Configuración de audio para elegir la carpeta predeterminada de guardado de audiolibros (predeterminada: Documentos\\Sonarpad Audiobooks).
• En el cuadro de guardado de audiolibro, cuando está activa la división en partes, se añadió una nueva opción (activada por defecto) para crear una subcarpeta dedicada a las partes generadas.
• La exportación de audiolibros ahora guarda MP3 en estéreo con bitrate elegido por el usuario para voces Edge, SAPI5 y SAPI4.
• Añadido soporte para voces SAPI5 de 32 bits mediante bridge, para usar también voces disponibles solo en motores de 32 bits.
• Reorganizadas las funciones de voz en un menú dedicado "Voz y audio" y añadida/aclarada la opción "Convertir audio", útil para convertir cualquier archivo multimedia compatible a MP3, AAC, OGG, Opus, FLAC, WAV y AIFF.
• Añadida la eliminación de artículos RSS individuales y episodios de podcast individuales (tecla Supr + menú contextual con confirmación), sin eliminar toda la fuente RSS/podcast, con deshacer de la última eliminación (artículo/episodio individual o fuente RSS/podcast completa).
• Añadida la exportación de feeds RSS a OPML en la ventana RSS, para guardar y reimportar fácilmente las fuentes actuales.
• Añadida la función "Buscar RSS por palabra clave" en la ventana RSS: al introducir una palabra clave se genera automáticamente la URL RSS de Google News y se abre el diálogo de añadir fuente ya prellenado, para crear feeds temáticos en un solo paso.
• Añadida la traducción serbia gracias a Mila Kuran.
• Añadida la traducción ucraniana gracias a Ivan Shtefuriak.
• Añadida la apertura múltiple de archivos multimedia: al abrir varios archivos juntos se crea una cola de reproducción en lugar de sustituir el archivo actual.
• Añadidos atajos de salto variable durante la reproducción: con base de 1 minuto, Izquierda/Derecha salta 60s, Shift+Izquierda/Derecha salta 20s y Ctrl+Izquierda/Derecha salta 3 minutos.
• Añadidos atajos para pista anterior/siguiente en el reproductor: Ctrl+PageUp y Ctrl+PageDown.
• Añadida la opción "Restablecer volumen" y agrupadas las acciones de reinicio en un submenú dedicado "Restablecer" en Reproducción, junto con "Restablecer velocidad" y "Restablecer tono".
• Mejoras del instalador: setup.exe ahora permite elegir entre asociar todos los tipos de archivo compatibles o seleccionar manualmente las extensiones; el MSI también ofrece selección por extensión en el árbol de características (el valor predeterminado se mantiene: todas activadas).
• Añadido el nuevo menú "Ventana" con la opción "Documentos abiertos..." para cambiar rápidamente a cualquiera de los archivos abiertos.
• Actualizada la opción Ver > Fuente: se reemplazó el selector completo por un submenú rápido con fuentes comunes (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia), manteniendo el tamaño de texto actual.
• Mejorada la lectura de RSS y podcasts con dos avisos diferenciados: los nodos de fuente anuncian "nuevos elementos" cuando un feed/podcast tiene novedades, mientras que los artículos RSS y episodios de podcast individuales anuncian "no leído"/"no reproducido"; este comportamiento se puede desactivar desde Opciones.
Correcciones de errores
• Corregida la extracción de texto EPUB para libros con comentarios HTML inline (<!-- ... -->): ahora el texto de los capítulos se analiza correctamente en lugar de omitirse parcial o totalmente.
• Corregido el diccionario Wiktionary en español y la caché del diccionario: palabras como "agua" ahora se encuentran correctamente y las entradas antiguas de "Palabra no encontrada" ya no se reutilizan.
• Corregida la codificación al importar artículos RSS en algunas fuentes españolas (p. ej., El Mundo): los acentos y la "ñ" ahora se muestran correctamente en el editor temporal.
• Corregida la decodificación ANSI de archivos centroeuropeos (p. ej., checo/polaco): Sonarpad ahora distingue mejor entre UTF-8 y ANSI y elige la página de códigos correcta (incluida Windows-1250), evitando diacríticos corruptos.
• Corregida la persistencia de fuentes RSS con parámetros en la URL (p. ej., `rss.aspx?c=...`): estos feeds ahora se guardan y restauran correctamente tras reiniciar Sonarpad.
• Corregida la apertura de archivos puntero de Google Drive (`.gdoc`, `.gsheet`, `.gslides`) desde el menú contextual del Explorador: si la lectura directa falla con “Incorrect function (os error 1)”, Sonarpad ahora usa un fallback por shell-open y el documento se abre correctamente.
• Corregida la lectura de archivos Excel legacy `.xls` (Excel 2010): los archivos binarios antiguos ahora se detectan y decodifican correctamente en lugar de mostrar texto corrupto (p. ej. `ÐÏ_à¡±...`).
• Corregido el flujo de anuncios del corrector ortográfico: los errores vuelven a anunciarse al revisar el texto más tarde, y el mismo error se informa de nuevo si se borra y se vuelve a escribir.
• Corregidas las operaciones de texto por líneas (p. ej. Ctrl+Q / Ctrl+Shift+Q, ordenar/invertir/únicas/unir líneas): al seleccionar una sola línea con Mayús+Flecha abajo ya no se unen ni se truncan las líneas adyacentes.
• Corregido el comportamiento multinea en operaciones de texto por líneas (Ctrl+Q / Ctrl+Shift+Q y herramientas relacionadas): cuando RichEdit entrega separadores de línea solo CR, ahora se normalizan correctamente y se procesan todas las líneas seleccionadas sin cortar el primer carácter.
• Ampliada la normalización de entrada TTS para símbolos visibles de espacio/tab/nueva línea (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), que con voces multilingües podían causar repetición de párrafos.
• Refinada la sanitización del texto para Edge TTS con una única tubería de validación: normalización de espacios extraños/invisibles, compactación de secuencias largas de puntuación (como "...", "!!!", "???") y omisión de fragmentos formados solo por puntuación para evitar bucles de reproducción.
• Corregido el anuncio del tiempo de reproducción (Ctrl+I) para streams MP3/podcast: el tiempo actual ahora se limita a la duración de la pista y la reproducción se detiene automáticamente si la posición supera el final.
• Mejorada la cobertura de localización del instalador: setup.exe ahora incluye también checo, polaco, francés y serbio, mientras que el MSI se mantiene como un único paquete en-US para evitar confusión en las releases.
• Corregida la limpieza en desinstalación de entradas del menú contextual: "Abrir con Sonarpad" ahora se elimina de forma fiable, también en escenarios de registro heredados.
• Corregida la fiabilidad de pausa/reanudar en SAPI5: la pausa con F4 ahora funciona correctamente y al reanudar vuelve al punto esperado en lugar de reiniciar desde el principio.
• Corregido el flujo pausa + salto + reanudar en la reproducción multimedia: tras pausar y mover con Izquierda/Derecha, al pulsar Espacio ahora reanuda de forma fiable desde la posición actual en lugar de detenerse o reiniciar desde el inicio.

Version 0.6.5 – 2026-02-07
Mejoras
• Traducción al español mejorada gracias a Arturo Fernandez Rivas.
• Se agregó una opción para dividir audiolibros EPUB por capítulos.
• Las importaciones RSS ahora usan una pestaña temporal dedicada (título localizado); Guardar como la convierte en un documento normal.
• Los mensajes del lector de pantalla ahora también se envían a JAWS cuando está disponible.
Correcciones de errores
• La lectura desde el cursor (F5) ahora empieza exactamente en el cursor. Antes podía comenzar un par de líneas arriba porque el desplazamiento del cursor no coincidía con las posiciones CRLF/UTF-16.
• Corregido un problema de redibujado: al escribir sobre una selección, el texto anterior podía desaparecer hasta mover la selección.
• Corregido el parseo de capítulos EPUB: las páginas de portada o solo con imágenes ya no generan lectura de CSS (p. ej., "padding") ni títulos "Sconosciuto".
• Corregido el fallo al dividir por tiempo audiolibros desde EPUB: Edge TTS podía fallar con chunks vacíos o demasiado largos ("Edge audio not sent").
• Se decodifican las entidades HTML en los artículos RSS (por ejemplo &quot;, &amp;, &lt;, &gt;).
• Guardar/Guardar como ahora propone el nombre del archivo existente al guardar formatos no sobrescribibles (p. ej., EPUB), en lugar de la primera línea.
• Se corrigió un problema por el que los podcasts con nuevos episodios no se anunciaban como no reproducidos, y se renombró "No escuchado" a "No reproducido" por ser más profesional.

Version 0.6.4 – 2026-02-05
Mejoras
• El programa se ha renombrado a Sonarpad para dar mayor enfasis al sonido y al audio, que son la clave de este programa.
• Añadida la selección de pistas de audio en el menú Reproducción para archivos multimedia con múltiples pistas de audio (ej. MKV con varios idiomas).
• Los podcasts ahora indican claramente los no escuchados con el prefijo "No escuchado" antes del nombre.
• Nuevo sistema de etiquetas para cambiar la voz en el texto. Ejemplos:
  - Voces Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Hola</voice>
  - Voces SAPI5: <voice sapi5 Microsoft Helena Desktop>Hola</voice>
  - Voces SAPI4: <voice sapi4 #1>Hola</voice>
  - Con velocidad/tono/volumen: <voice edge it-IT-ElsaNeural speed=-20 pitch=-5 volume=-10>Hola</voice>
• Enriquecidas las categorias de podcasts.
• Añadida una opción en el menú contextual para crear un audiolibro desde la selección.
• Añadida la división de audiolibros por duración, con la posibilidad de elegir el nombre del primer archivo.
• Localizada la etiqueta del autor en la lectura de artículos (ej. "por", "by", "di").
• Añadidas opciones de indentación (tabulaciones/espacios con anchura) y Tab/Mayús+Tab para indentar/desindentar líneas seleccionadas.
• Corregida la limpieza de Markdown: ahora gestiona los bullets '*' cuando no se conservan las listas.
Correcciones de errores
• Corregido un error por el cual los audiolibros con SAPI4 podian crearse de forma diferente a lo esperado.
• Ventana Buscar en archivos: al pulsar Intro en un resultado ahora abre en la posición correcta del fragmento y Esc vuelve a los resultados.
• Ventana Opciones: se ajustó el diseño visual de las pestañas General, Voz, Editor y Audio para evitar controles faltantes o recortados.
• Corregido un problema de marcadores al cambiar la velocidad de reproducción.
• Corregido un problema con Podcast Index y las categorías que no se mostraban correctamente.
• Corregido el problema del apóstrofo que cortaba la lectura: ya no hay lectura separada para diálogos, se usan las etiquetas de voz.

Versión 0.6.3 – 2026-01-30
Mejoras
• Mejorada la detección del micrófono.
• Añadida reproducción instantánea para todos los formatos.
Correcciones
• Corregido el fallo en la ventana de categorías de podcasts.

Versión 0.6.2 – 2026-01-30
Nuevas funcionalidades
• Añadida la ejecución de archivos (Shift+F5). Los usuarios pueden seleccionar un intérprete (ej. python) en Opciones, buscarlo en el ordenador, y presionando Shift+F5 se ejecuta el script actual. Los archivos HTML se abren en el navegador.
• Añadido soporte para archivos de puntero de Google Docs (.gdoc, .gsheet, .gslides), que se abren automáticamente en el navegador predeterminado.
• Añadido soporte para el formato de audiolibro M4B (Apple/AAC).
• Añadida la opción "Mostrar episodios" en el menú contextual de resultados de búsqueda de podcasts para explorar y reproducir episodios sin suscribirse.
• Añadida la función "Ir a línea" (menú Editar o Ctrl+J) para saltar rápidamente a un número de línea específico.
• Añadidas opciones en el menú contextual para ordenar feeds RSS y podcasts (alfabéticamente o por fecha).
• Añadidos feeds RSS predeterminados en vietnamita.
• Añadida una casilla de prueba de micrófono en el diálogo de grabación para verificar los niveles antes de comenzar.
• Añadida "Mostrar descripción" para episodios de podcast en el menú contextual.
• Añadido soporte para formatos de audio/video extendidos vía FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Añadida lectura sincronizada de subtítulos (srt, vtt, ass, sub, sbv, lrc, smi) con NVDA o voz seleccionada. El programa busca un archivo de subtítulos con el mismo nombre que el archivo multimedia. Añadidas las opciones "Importar subtítulos" y "Eliminar subtítulos" en el menú Reproducción para archivos con nombres diferentes.
• Añadidas asociaciones de archivos para todos los nuevos formatos de audio/video soportados en el menú contextual "Abrir con Sonarpad".
• Añadida configuración para ajustar el pitch de cualquier archivo.
• Añadida opción en Configuración General para activar o desactivar los informes de errores anónimos. Añadida una entrada en el menú Ayuda para crear un archivo ZIP de diagnóstico.
• Añadida opción para usar una voz diferente para los diálogos, tanto para la lectura en vivo como para la creación de audiolibros.
• Añadido el explorador de categorías de podcasts para explorar podcasts por categoría (negocios, arte, deportes, etc.).
Mejoras
• Abrir un archivo de audio/video desde el Explorador ahora abre directamente la vista del reproductor en lugar del editor de texto.
• Eliminada la solicitud de OCR para PDFs no accesibles; el OCR ahora se realiza automáticamente para mejorar la velocidad y experiencia del usuario.
• Mejorada la Terminal Accesible: la lectura NVDA ahora recuerda la última línea leída para mejor continuidad.
• SAPI 4: La creación de audiolibros ahora está completamente paralelizada y es casi instantánea. Añadida una solicitud para elegir el número de procesos concurrentes.
• SAPI 4: Eliminado el cuello de botella WAV-MP3 convirtiendo fragmentos en paralelo durante la síntesis.
• SAPI 4: Mejorado el manejo de errores y limpieza automática de archivos temporales.
• Diálogo Buscar: Renombrado "Regex" a "Expresión regular" para mayor claridad y añadidas las traducciones faltantes para las opciones de búsqueda.
• Audiolibros M4B: Mejor manejo de salida; dividir por partes/marcadores ahora produce un único archivo M4B con metadatos de capítulos incluyendo título y autor.
• Reproductor: Corregida la precisión de marcadores y anuncios de tiempo cuando la velocidad de reproducción no es 1.0x.
• Restaurada la navegación Ctrl+Tab y Ctrl+Shift+Tab en Opciones.
• Añadida una opción en el menú Reproducción para restablecer instantáneamente la velocidad Normal (1.0x).
• Actualizadas todas las dependencias a las últimas versiones para mejor rendimiento y estabilidad.
• Integrado FFmpeg con carga dinámica de DLL para garantizar compatibilidad sin bloquear el inicio.
• Actualizados los filtros de descarga de podcasts para incluir los nuevos formatos de audio/video.
• Impedido que Ctrl+S guarde archivos de audio/video para evitar corrupción.
• Mejorada la importación de transcripciones de YouTube haciéndola más robusta y resiliente.
• Mejorada la robustez de la división en partes de audiolibros, asegurando que no se pierda texto.
• El instalador ahora es completamente multilingüe, soportando Italiano, Inglés, Español, Portugués, Sueco y Vietnamita según el idioma del sistema del usuario. El inglés es el predeterminado para sistemas no soportados.
• Categorías de podcasts: presionar Enter en una categoría ahora confirma la selección (equivalente al botón OK).
• Mejorado el sistema de detección de bloqueos para evitar falsos positivos cuando hay diálogos modales abiertos (mensajes de error, "texto no encontrado").
Correcciones
• Corregido un error donde el changelog no se abría al inicio.
• Corregido un error donde la solicitud de OCR no aparecía para PDFs no accesibles abiertos desde el Explorador.
• Corregido un error de inicio que podía causar pérdida de foco o cierre de ventanas inmediatamente después de abrir.
• Corregido un error crítico en la búsqueda regex que impedía encontrar texto, incluyendo problemas con "Búsqueda circular" y la opción "El punto equivale a nueva línea" con terminaciones de línea de Windows.
Localización
• Añadida la traducción al polaco.
• Añadida la traducción al francés.
• Añadida la traducción al checo (gracias a Radek Žalud y Jiri Holzinger).

Versión 0.6.1 – 2026-01-20
Correcciones
• Corregido un error por el cual, al activar “Mostrar voces en el editor” y reproducir un podcast, la reproducción se detenía.
• Corregido un problema por el cual algunos podcasts no podían añadirse mediante URL porque la dirección se truncaba.
• Corregido un error por el cual ya no era posible añadir URLs normales en la función de feeds RSS.
• Corregido un problema por el cual el idioma de Wikipedia se mostraba en varias pestañas de configuración.
• Eliminada la creación de algunos archivos de depuración que se generaban incluso en modo release.
Mejoras
• Mejorado el soporte para las voces de Microsoft, que ahora se reproducen mediante un método dedicado con un user agent diferente.
• Añadido soporte para archivos MP4.

Versión 0.6.0 – 2025-01-20
Nuevas funciones
• Añadido el corrector ortográfico. Desde el menú contextual es posible comprobar si la palabra actual es correcta y, en caso contrario, obtener sugerencias.
• Añadida la importación y exportación de podcasts mediante archivos OPML.
• Añadido soporte para la búsqueda en Podcast Index además de iTunes. El usuario puede introducir su API key y API secret gratuitos (generados usando solo su correo electrónico).
• Añadido soporte para voces SAPI4, tanto para la lectura en tiempo real como para la creación de audiolibros.
• Añadido un fallback automático de OCR para PDFs no accesibles: cuando no se encuentra texto extraíble, el documento se reconoce mediante OCR.
• Añadido soporte de diccionario mediante Wiktionary. Al pulsar la tecla Aplicaciones se muestran las definiciones y, cuando están disponibles, también sinónimos y traducciones a otros idiomas.
• Añadida la importación de artículos desde Wikipedia con búsqueda, selección de resultados e importación directa en el editor.
• Añadido el atajo Shift+Enter en el módulo RSS para abrir un artículo directamente en el sitio web original.
Mejoras
• La selección del micrófono ahora siempre es respetada por la aplicación.
• En la ventana de podcasts, al pulsar Enter sobre un episodio, NVDA anuncia inmediatamente “cargando”, proporcionando confirmación inmediata de la acción.
• En los resultados de búsqueda de podcasts, al pulsar Enter ahora se realiza la suscripción al podcast seleccionado.
• Corregidas y mejoradas las etiquetas de los atajos Ctrl+Shift+O y Podcast Ctrl+Shift+P.
• La velocidad de reproducción y el volumen ahora se guardan en la configuración y se mantienen para todos los archivos de audio.
• Añadida una carpeta de caché dedicada para los episodios de podcasts. El usuario puede conservar los episodios mediante “Conservar podcast” en el menú Reproducir. La caché se limpia automáticamente cuando supera el tamaño configurado por el usuario (Opciones → Audio).
• Mejorada de forma significativa la obtención de artículos RSS utilizando libcurl con impersonación de Chrome e iPhone, garantizando compatibilidad con aproximadamente el 99 % de los sitios.
• Añadido el estado leído / no leído para los artículos RSS, con indicación clara en la lista RSS.
• La función Reemplazar todo ahora muestra también el número de reemplazos realizados.
• Añadido el botón Eliminar podcast al navegar por la biblioteca de podcasts mediante Tab.
Correcciones
• Eliminada la entrada redundante “pending update” del menú Ayuda (las actualizaciones ya se gestionan automáticamente).
• Corregido un error por el cual, al abrir un archivo MP3 y pulsar Ctrl+S, el archivo se guardaba y quedaba corrupto.
• Corregido un problema de interfaz donde “Batch Audiobooks” se mostraba como “(B)… Ctrl+Shift+B” (se eliminó la etiqueta redundante).
• Corregido el funcionamiento de las comillas inteligentes: cuando están habilitadas, las comillas normales ahora se sustituyen correctamente por comillas tipográficas.
• Corregido un error por el cual, al usar “Ir al marcador”, la velocidad de reproducción se restablecía a 1.0.
• Corregido un problema por el cual los episodios de podcasts ya descargados se volvían a descargar en lugar de usar la versión en caché.
Atajos de teclado
• F1 ahora abre la guía.
• F2 ahora comprueba si hay actualizaciones.
• F7 / F8 ahora permiten desplazarse al error ortográfico anterior o siguiente.
• F9 / F10 ahora permiten cambiar rápidamente entre las voces guardadas en favoritos.
Mejoras para desarrolladores
• Los errores ya no se ignoran silenciosamente: se han eliminado todos los patrones let _ = y los errores ahora se gestionan explícitamente (propagados, registrados o tratados con mecanismos de respaldo adecuados).
• El proyecto ahora no compila si hay advertencias: tanto cargo check como cargo clippy deben completarse sin avisos, con lints más estrictos y eliminación de allow donde sea posible.
• Eliminadas las implementaciones personalizadas de tipo strlen / wcslen. Las longitudes de cadenas y buffers UTF-16 ahora se derivan de datos gestionados por Rust, sin escanear memoria manualmente.
• La gestión de DLL se ha limpiado y centralizado en torno a libloading, evitando lógica de carga personalizada y análisis PE.
• Eliminados los helpers manuales para el parsing de bytes: ahora todo el parsing utiliza from_le_bytes / from_be_bytes sobre slices verificadas.
Estos cambios reducen el uso innecesario de unsafe, eliminan posibles comportamientos indefinidos y hacen que el código sea más idiomático, robusto y mantenible.

Version 0.5.9 - 2025-01-13
Nuevas funciones
• Aniadida la posibilidad de reordenar RSS desde el menu contextual (arriba/abajo/a posicion) con controles para posiciones no validas.
• Aniadido un menu contextual para los articulos con abrir sitio original y compartir por WhatsApp, Facebook y X.
• Aniadido el atajo Esc para volver desde articulos importados a la lista RSS.
• Aniadido el modo podcast: buscar, suscribirse y escuchar; reordenar suscripciones; Esc detiene la reproduccion y vuelve a la lista; Enter en un episodio inicia la reproduccion.
• Aniadido el control de velocidad de reproduccion para podcasts y archivos MP3.
• Aniadido Ctrl+T para ir a un tiempo especifico.
• Aniadido un boton de vista previa de voz despues del combo de volumen.
• Aniadida la funcion de regex para Buscar y Reemplazar, estilo Notepad++.
• Aniadida la importacion de RSS desde archivos OPML y TXT.
• Aniadida la casilla en Opciones para habilitar "Abrir con Sonarpad" en el Explorador de archivos, tambien en version portable.
Mejoras
• Mejorada la seleccion de velocidad, tono y volumen de las voces, respetando los limites maximos del TTS.
• Varias mejoras de RSS para descargar todos los articulos sin mover el foco de NVDA durante las actualizaciones.
• Mejorada la reproduccion de audio con un menu dedicado, anuncio del tiempo con Ctrl+I y volumen hasta el 300%.
• Aniadidos atajos faltantes para algunas funciones.
• Reorganizado el menu Editar con un submenu para las funciones de limpieza de texto.
• Reorganizadas las Opciones en pestanas, con Ctrl+Tab y Ctrl+Shift+Tab para moverse entre ellas.
• Resueltos los problemas de lectura de articulos: el lector RSS ahora muestra los articulos completos como en el navegador.
Correcciones
• Corregido un problema por el que la limpieza de Markdown eliminaba numeros al inicio de linea.
• Corregido AltGr+Z que activaba Undo.
• Corregido un problema por el que al grabar un audiolibro no se podia detener rapidamente.
Localizacion
• Aniadida la traduccion vietnamita (gracias a Anh Duc Nguyen).

Version 0.5.8 - 2026-01-10
Nuevas funciones
• Aniadido control de volumen para microfono y audio del sistema al grabar podcasts.
• Aniadida una nueva funcion para importar articulos desde sitios web o feeds RSS, incluyendo los feeds mas importantes para cada idioma.
• Aniadida una funcion para eliminar todos los marcadores del archivo actual.
• Aniadida la funcion para eliminar lineas duplicadas y lineas duplicadas consecutivas.
• Aniadida la funcion para cerrar todas las pestanas o ventanas excepto la actual.
• Aniadida la entrada Donaciones en el menu Ayuda para todos los idiomas.
Mejoras
• Mejorado el terminal accesible para evitar algunos bloqueos.
• Mejoradas y corregidas las access key y los atajos de teclado del programa.
• Corregido un problema por el que al cerrar la ventana de reproduccion de audio la reproduccion no se detenia.
• Aniadidas ventanas de confirmacion para acciones importantes (p. ej., eliminar lineas duplicadas, eliminar guiones al final de linea, eliminar todos los marcadores del archivo actual). No se muestra confirmacion si la accion no se aplica.
• Aniadida la posibilidad de eliminar feeds/sitios RSS de la biblioteca seleccionandolos y pulsando Supr.
• Aniadido un menu contextual en la ventana RSS para modificar o eliminar feeds/sitios RSS.
• Eliminada la casilla para mover la configuracion a la carpeta actual; ahora el programa lo gestiona automaticamente (si la carpeta del exe se llama "sonarpad portable" o el exe esta en una unidad extraible, guarda en la carpeta del exe en `config`, si no en `%APPDATA%\\Sonarpad`, con fallback a `config` si la carpeta preferida no es escribible).

Version 0.5.7 - 2026-01-05
Nuevas funciones
• Aniadida opcion para grabar audiolibros en lote (conversion multiple de archivos y carpetas).
• Aniadido soporte para archivos Markdown (.md).
• Aniadida eleccion de codificacion al abrir archivos de texto.
• Aniadida opcion en el terminal para anunciar nuevas lineas con NVDA.
Mejoras
• La grabacion de audiolibros se guarda ahora en MP3 nativo cuando se selecciona.
• El usuario puede elegir donde colocar el asterisco * que indica cambios no guardados.
• Mejorado el sistema de actualizacion para ser mas robusto en diferentes escenarios.
• Aniadida en el menu Editar la funcion para eliminar guiones al final de linea (util para textos OCR).

Version 0.5.6 - 2026-01-04
Correcciones
  Mejorado Buscar en archivos: al pulsar Enter abre el archivo exactamente en el fragmento seleccionado.
Mejoras
  Soporte PPT/PPTX.
  Para formatos no textuales, Guardar ahora propone siempre .txt para no romper el formato (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Grabacion de podcast desde microfono y/o audio del sistema (menu Archivo, Ctrl+Shift+R).

Version 0.5.5 - 2026-01-03
Nuevas funciones
• Aniadido un terminal accesible optimizado para mucho output y lectores de pantalla (Ctrl+Shift+P).
• Aniadida la opcion de guardar la configuracion en la carpeta actual (modo portable).
Correcciones
• Mejorados los fragmentos de Buscar en archivos para que la vista previa quede alineada con la coincidencia.

Version 0.5.4 – 2026-01-03
Mejoras
• Correccion de Normalizar espacios en blanco (Ctrl+Shift+Enter).
• Soporte HTML/HTM (abrir como texto).

Version 0.5.3 – 2026-01-02
Nuevas funciones
• Se agrego Buscar en archivos.
• Se agregaron nuevas herramientas de texto: Normalizar espacios en blanco, Salto de linea duro y Quitar Markdown.
• Se agrego Estadisticas de texto (Alt+Y).
• Se agregaron nuevos comandos de lista en el menu Editar:
• Ordenar lineas (Alt+Shift+O)
• Eliminar duplicados (Alt+Shift+K)
• Invertir lineas (Alt+Shift+Z)
• Se agregaron Comentar / Descomentar lineas (Ctrl+Q / Ctrl+Shift+Q).
Localizacion
• Se agrego la localizacion en espanol.
• Se agrego la localizacion en portugues.
Mejoras
• Cuando un archivo EPUB esta abierto, Guardar cambia automaticamente a Guardar como y exporta el contenido como .txt para evitar la corrupcion del EPUB.

## 0.5.2 - 2026-01-01
- Se agrego un changelog.
- Se agregaron opciones "Abrir con Sonarpad" y asociaciones de archivos compatibles durante la instalacion.
- Se mejoro la localizacion de mensajes (errores, dialogos, exportacion de audiolibro).
- Se agrego la seleccion de partes al usar "Dividir audiolibro por texto", con la opcion "Requerir el marcador al inicio de la linea".
- Se agrego la importacion de transcripciones de YouTube con seleccion de idioma, opcion de marca de tiempo y mejoras de foco.

## 0.5.1 - 2025-12-31
- Actualizaciones automaticas con confirmacion, manejo de errores y notificaciones mejoradas.
- Mejoras en exportacion de audiolibros (division por texto, SAPI5/Media Foundation, controles avanzados).
- Mejoras en TTS (pausa/reanudar, diccionario de reemplazos, favoritos).
- Menu Ver y paneles de voces/favoritos, color y tamano del texto.
- Idioma predeterminado del sistema y mejoras de localizacion.
- CI y empaquetado Windows (artefactos, MSI/NSIS, cache).

## 0.5.0 - 2025-12-27
- Refactor modular (editor, manejo de archivos, menu, busqueda).
- Workflow de compilacion/empaquetado Windows y actualizaciones de README/licencia.
- Arreglo de navegacion TAB en la ventana de Ayuda.

## 0.5 - 2025-12-27
- Actualizacion preliminar del numero de version.

## 0.1.0 - 2025-12-25
- Version inicial: estructura del proyecto y README.






