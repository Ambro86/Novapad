# Changelog
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
 Mejorado Buscar en archivos: al pulsar Enter abre el archivo exactamente en el fragmento seleccionado.
Mejoras
 Soporte PPT/PPTX.
 Para formatos no textuales, Guardar ahora propone siempre .txt para no romper el formato (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
 Grabacion de podcast desde microfono y/o audio del sistema (menu Archivo, Ctrl+Shift+R).

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
