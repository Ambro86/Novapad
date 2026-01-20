# Changelog

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
• Aniadida la casilla en Opciones para habilitar "Abrir con Novapad" en el Explorador de archivos, tambien en version portable.
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
• Eliminada la casilla para mover la configuracion a la carpeta actual; ahora el programa lo gestiona automaticamente (si la carpeta del exe se llama "novapad portable" o el exe esta en una unidad extraible, guarda en la carpeta del exe en `config`, si no en `%APPDATA%\\Novapad`, con fallback a `config` si la carpeta preferida no es escribible).

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
- Se agregaron opciones "Abrir con Novapad" y asociaciones de archivos compatibles durante la instalacion.
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
