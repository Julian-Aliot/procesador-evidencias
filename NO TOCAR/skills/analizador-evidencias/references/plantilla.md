# Excel de once columnas

Dos hojas: Resumen y Evidencias. Evidencias conserva las once columnas, encabezados en fila 1, autofiltro y una fila por persona/tarea identificable. Cada documento aporta al menos una fila, incluso pendiente. Los comentarios del enlace conservan identificador y citas para revisión.

| Orden | Encabezado | Contenido |
|---|---|---|
| 1 | PERSONAS | Nombres literales y rol observado, una persona por fila y su rol; la misma tarea puede aparecer para varias personas con respaldo documental. |
| 2 | CÓDIGO PROYECTO | Códigos literales, conservando ceros y puntuación. |
| 3 | NOMBRE PROYECTO | Denominación presente en la evidencia. |
| 4 | FECHA | Fecha real de Excel con presentación dd/mm/aaaa. |
| 5 | TÍTULO TAREA | Título literal cuando consta. |
| 6 | RESUMEN EVIDENCIA | Explicación breve en español de lo que acredita el documento. |
| 7 | UBICACIÓN | Página, hoja/celda, párrafo o región. |
| 8 | ARCHIVO EVIDENCIA | Enlace relativo a la copia en evidencias/. |
| 9 | REVISIÓN | Ausencias, ambigüedades y límites del contraste. |
| 10 | PERSONAL BONIFICADO DETECTADO | Nombres canónicos confirmados contra el listado aportado. |
| 11 | CÓDIGO DE PROYECTO DETECTADO | Códigos canónicos confirmados contra el listado aportado. |

Los datos textuales se escriben como texto, incluidos los que empiezan por signos de fórmula. Las fechas desconocidas permanecen vacías. Los nombres y códigos conservan sus caracteres. El libro y la carpeta evidencias deben mantenerse juntos para utilizar los enlaces relativos. El ZIP incorpora únicamente esos elementos.

Recuentos requeridos antes de entregar, documentos = copias, todos los documentos presentes en filas y filas de detalle = suma de combinaciones persona/tarea por documento, y filas de Resumen = documentos. Un archivo idéntico a otro se conserva y se comunica. Cada envío se contrasta por separado.

## Resumen documental (versión 1.1.0)

La primera hoja, Resumen, tiene exactamente siete columnas: NOMBRE PRINCIPAL DE LA EVIDENCIA, OTROS NOMBRES, CÓDIGO DE PROYECTO, CÓDIGO DETECTADO, FECHA, TÍTULO, RESUMEN. Una fila por identidad del manifiesto: siete PNG producen siete filas, aunque sean idénticos o tengan varias tareas. Un PDF de varias páginas cuenta como un documento.

Nombre principal es la persona protagonista de la evidencia, respaldada por su contexto; otros nombres recoge las restantes personas únicas. No elijas protagonista por orden de aparición ni por pertenecer al listado bonificado. Si hay varias personas sin principal inequívoco, deja principal vacío y conserva todas en otros nombres. Código de proyecto conserva códigos literales; código detectado contiene únicamente coincidencias canónicas con el listado de proyectos opcional. Sin listado, queda vacío. Si hay varios códigos, consérvalos en la misma celda. Fecha contiene una fecha única; con varias fechas distintas queda vacía y el comentario remite al detalle. Título y resumen sintetizan el documento completo sin multiplicar filas. El título enlaza a la copia; su comentario conserva archivo original, ID, citas e incidencias. Si no consta título, usa la etiqueta «Sin título identificado».

Evidencias conserva las once columnas. El cruce de PERSONAL BONIFICADO DETECTADO se realiza para la persona de cada fila, con coincidencias exactas normalizadas o alias autorizados únicos. El comentario del enlace conserva nombre canónico, origen del listado y citas del documento; permite trazar Resumen → Evidencias → copia original. No completes el campo con otros compañeros de la misma tarea. Sin listado, el campo queda vacío y REVISIÓN indica contraste pendiente.
