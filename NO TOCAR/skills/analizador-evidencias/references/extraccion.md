# Contrato de extracción

El preparador crea un JSON con version, envio_id y documentos. Conserva los campos id y sha256 de cada documento. `lectura_completa` se establece en true después de revisar todo el contenido relevante, incluidas imágenes y hojas ocultas. Los documentos ilegibles conservan false, una incidencia y una actividad con datos pendientes.

Cada documento contiene `incidencias_documento` como lista de textos y `actividades` con al menos un elemento. Divide en varias actividades cuando el documento respalde tareas, fechas o relaciones diferentes. Evita multiplicar combinaciones de personas y proyectos sin respaldo. Los listados de referencia orientan las coincidencias, mientras los datos extraídos proceden de la evidencia.

Cada actividad sigue esta estructura. Los valores de demostración ilustran el esquema y se sustituyen por los datos del documento.

```json
{
  "personas": [
    {"valor": "Nombre literal", "rol": "participante", "ubicacion": "Página 1, sección Asistentes", "cita": "Asistentes: Nombre literal"}
  ],
  "codigos_proyecto": [
    {"valor": "PR-001", "ubicacion": "Página 1, encabezado", "cita": "Proyecto PR-001"}
  ],
  "nombre_proyecto": {"valor": null, "ubicacion": null, "cita": null},
  "fecha": {"valor": "2026-11-19", "ubicacion": "Página 1, Fecha de actividad", "cita": "Fecha de actividad: 19/11/2026"},
  "tipo_fecha": "realizacion",
  "titulo_tarea": {"valor": "Título literal", "ubicacion": "Página 1, título", "cita": "Título literal"},
  "resumen": {"valor": "Descripción breve de la evidencia.", "ubicacion": "Página 1, cuerpo", "cita": "Fragmento que respalda el resumen"},
  "motivos_revision": []
}
```

`valor` es texto literal salvo el resumen, que se redacta en español apoyado en una cita. `ubicacion` y `cita` son obligatorias para cada valor conocido. En hojas de cálculo usa hoja y celda o rango; en texto, líneas; en imágenes, imagen y región; en documentos, página o párrafo. Guarda citas breves y suficientes para comprobar la relación.

Los roles admitidos son participante, autor, responsable, destinatario, mencionado y no_determinado. La última opción genera revisión. Un nombre en una firma o en una copia de correo conserva ese contexto. Un código en un pie de página o catálogo general requiere revisión antes de vincularlo a la tarea.

Las fechas usan AAAA-MM-DD. Su tipo es realizacion, participacion, documento, envio u otra. Los tipos documento, envio y otra se explican en REVISIÓN. Las fechas ambiguas o periodos sin día concreto se conservan en citas y resumen, dejando fecha vacía y un motivo. Para varias fechas de realización distintas, crea actividades distintas cuando se puedan atribuir.

Conserva completos los códigos, las nomenclaturas y el título observado. Los nombres abreviados se mantienen tal como aparecen. Una captura puede carecer de título, proyecto o fecha; registra esa ausencia. Los documentos corruptos y formatos sin lector conservan una copia y una fila pendiente.

El control automático comprueba estructura, tipos, procedencia declarada e integridad. El agente verifica que las citas realmente figuren y apoyen los valores. La fila expresa la evidencia encontrada y sus límites.

## Resumen por documento

Cada documento puede incluir `resumen_documento`, objeto con `nombre_principal`, `titulo` y `resumen`, cada uno con el contrato `{valor, ubicacion, cita}` (null para ausencias). Cumplimenta título y resumen con una síntesis respaldada del documento completo. Nombre principal debe coincidir literalmente con una persona de sus actividades y tener respaldo contextual. Para varios nombres sin protagonista inequívoco, usa null. Sin este objeto, el script agrupa valores únicos de las actividades; selecciona nombre principal automáticamente solo cuando hay una única persona. No inventes una persona para completar el resumen.

En actividades, vincula únicamente personas realmente relacionadas con la tarea y conserva sus roles. El generador desdobla cada actividad en una fila por persona (o una fila pendiente si no hay personas). Las coincidencias se contrastan y trazan individualmente. No crees duplicados artificiales ni relaciones por mera proximidad textual.

En Resumen, NOMBRE DETECTADO reúne los nombres canónicos contrastados de todas las personas de la evidencia, una vez por nombre. Para DESCRIPCIÓN TAREA, redacta en resumen_documento.descripcion_tarea una frase breve de acción y objeto con valor, ubicación y cita. Por ejemplo, «Desarrollo del modelo de la cabra» o «Compositing del plano 024». Si hay varias tareas, sintetiza el trabajo acreditado. Si no se puede identificar, deja el valor vacío y anota el motivo en incidencias_documento.

En capturas, transcribe códigos junto con sus prefijos o sufijos visibles en codigos_proyecto, con región y cita. Conserva cada nombre literal. Si es dudosa la correspondencia de identidad, añade identidad_dudosa: true a esa persona.
