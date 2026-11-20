# Excel de once columnas

Dos hojas: Resumen y DETALLE. DETALLE conserva las once columnas, encabezados en fila 1, autofiltro y una fila por persona/tarea identificable. Cada documento aporta al menos una fila, incluso pendiente. Los comentarios del enlace conservan identificador y citas para revisión.

| Orden | Encabezado | Contenido |
|---|---|---|
| 1 | NOMBRE DETECTADO | Persona tal como aparece en la evidencia. Su rol y cita figuran en el comentario del enlace. |
| 2 | NOMBRE BONIFICADO | Nombre completo trazado al listado de personal. |
| 3 | CÓDIGO DETECTADO | Etiqueta literal de proyecto, incluidos prefijos y sufijos observados. |
| 4 | CÓDIGO BONIFICADO | Código canónico del listado de proyectos. Si hay varios, conserva la posición de cada literal. |
| 5 | FECHA | Fecha acreditada. |
| 6 | TÍTULO TAREA | Título observado. |
| 7 | RESUMEN EVIDENCIA | Qué acredita el documento sobre la actividad. |
| 8 | UBICACIÓN | Página, celda o región. |
| 9 | ARCHIVO EVIDENCIA | Enlace relativo y comentario con citas, rol y origen del cruce. |
| 10 | REVISIÓN | Incidencias y límites del contraste. |
| 11 | DESCRIPCIÓN TAREA | Acción y objeto concretos de esta actividad. |

Los datos textuales se escriben como texto, incluidos los que empiezan por signos de fórmula. Las fechas desconocidas permanecen vacías. Los nombres y códigos conservan sus caracteres. El libro y la carpeta evidencias deben mantenerse juntos para utilizar los enlaces relativos. El ZIP incorpora únicamente esos elementos.

Recuentos requeridos antes de entregar, documentos = copias, todos los documentos presentes en filas y filas de detalle = suma de combinaciones persona/tarea por documento, y filas de Resumen = documentos. Un archivo idéntico a otro se conserva y se comunica. Cada envío se contrasta por separado.

## Resumen documental (versión 1.3.0)

La primera hoja, Resumen, tiene exactamente nueve columnas: NOMBRE PRINCIPAL DE LA EVIDENCIA, OTROS NOMBRES, CÓDIGO DE PROYECTO, CÓDIGO DETECTADO, FECHA, TÍTULO, RESUMEN, NOMBRE DETECTADO, DESCRIPCIÓN TAREA. Una fila por identidad del manifiesto: siete PNG producen siete filas, aunque sean idénticos o tengan varias tareas. Un PDF de varias páginas cuenta como un documento.

Nombre principal es la persona protagonista de la evidencia, respaldada por su contexto; otros nombres recoge las restantes personas únicas. No elijas protagonista por orden de aparición ni por pertenecer al listado bonificado. Si hay varias personas sin principal inequívoco, deja principal vacío y conserva todas en otros nombres. Código de proyecto conserva códigos literales; código detectado contiene únicamente coincidencias canónicas con el listado de proyectos opcional. Sin listado, queda vacío. Si hay varios códigos, consérvalos en la misma celda. Fecha contiene una fecha única; con varias fechas distintas queda vacía y el comentario remite al detalle. Título y resumen sintetizan el documento completo sin multiplicar filas. El título enlaza a la copia; su comentario conserva archivo original, ID, citas e incidencias. Si no consta título, usa la etiqueta «Sin título identificado».

DETALLE conserva las once columnas. El cruce de NOMBRE BONIFICADO se realiza para la persona de cada fila, con coincidencias exactas, abreviadas o por alias cuando la identificación es inequívoca. El comentario del enlace conserva nombre canónico, origen del listado y citas del documento; permite trazar Resumen → DETALLE → copia original. No completes el campo con otros compañeros de la misma tarea. Sin listado, el campo queda vacío y REVISIÓN indica contraste pendiente.

En Resumen, NOMBRE DETECTADO reúne los nombres canónicos contrastados de todas las personas de la evidencia, una vez por nombre. Para DESCRIPCIÓN TAREA, redacta en resumen_documento.descripcion_tarea una frase breve de acción y objeto con valor, ubicación y cita. Por ejemplo, «Desarrollo del modelo de la cabra» o «Compositing del plano 024». Si hay varias tareas, sintetiza el trabajo acreditado. Si no se puede identificar, deja el valor vacío y anota el motivo en incidencias_documento.
