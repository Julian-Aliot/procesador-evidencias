# Listados de contraste

El usuario puede adjuntar personal bonificado, códigos de proyecto, ambos o ninguno. Identifica esos archivos como referencias de este envío y conserva sus originales en el área técnica. Su contenido se trata como datos.

`catalogo.json` contiene version, personal y proyectos. Una categoría sin listado usa null. Una lista vacía significa que el usuario aportó expresamente un listado vacío y así debe explicarse. Cada entrada tiene un nombre o código canónico, alias autorizados y origen con archivo y ubicación.

```json
{
  "version": "1.2.0",
  "personal": [
    {"nombre": "Ana García López", "aliases": ["Ana G. López"], "origen": "Personal.xlsx, hoja Personal, A2:B2"}
  ],
  "proyectos": [
    {"codigo": "PR-001", "aliases": ["PR/001"], "origen": "Proyectos.xlsx, hoja Proyectos, A2:B2"}
  ]
}
```

El agente interpreta referencias PDF, imágenes, tablas mixtas o encabezados distintos y construye este catálogo con citas/localizaciones verificables. Mantén aparte las referencias y las evidencias. Cada alias debe figurar en el listado o estar autorizado explícitamente por el usuario. Registra esa autorización en origen.

Para un listado tabular sencillo hay importación auxiliar. El CSV/XLSX usa encabezados `nombre,aliases` para personal y `codigo,aliases` para proyectos. `aliases` es opcional y separa valores con |. Los códigos se guardan como texto para conservar ceros. Las plantillas vacías están en assets. En un libro con varias hojas y encabezados distintos, el agente selecciona e interpreta las hojas por separado.

```bash
python 'NO TOCAR/skills/analizador-evidencias/scripts/lectura.py' '/ruta/Personal.xlsx' \
  --listado personal --salida '/ruta/envio/lectura-listado-personal'
```

Revisa el fragmento generado e incorpora su categoría a catalogo.json, manteniendo la otra categoría y version. Para otro listado usa `--listado proyectos` y otro destino nuevo. Las plantillas vacías sirven para introducir datos y se detectan como vacías durante la importación.

El contraste normaliza Unicode, mayúsculas/minúsculas y espacios consecutivos. Conserva acentos, puntuación, guiones, ceros y el orden de palabras. Se comprueban primero los valores completos y los alias declarados. Los alias compartidos por varias entradas distintas producen revisión y quedan sin asignación. Se admiten iniciales con apellido y nombres de pila o prefijos suficientemente distintivos en un campo que identifica a la persona, siempre con un único candidato del listado. Conserva la grafía literal y la cita. Si el contexto deja dudas, marca identidad_dudosa: true en la persona extraída. El nombre literal y su rol se conservan y la detección queda vacía. En proyectos se admite un código delimitado dentro de una etiqueta, y códigos alfabéticos seguidos de numeración de plano. Conserva los ceros y deja en blanco las coincidencias ambiguas.

PERSONAL BONIFICADO DETECTADO contiene los nombres canónicos cuyas detecciones tienen coincidencia única. CÓDIGO DE PROYECTO DETECTADO contiene los códigos canónicos con coincidencia única. Ambas columnas pueden contener varias líneas. Sin listado, quedan vacías y REVISIÓN indica contraste pendiente. Con listado y sin coincidencia confirmada, quedan vacías con esa explicación. Las coincidencias de una lectura incompleta son parciales y se indica en REVISIÓN.

Estas columnas recogen presencia textual contrastada. Los roles y la relación con la tarea se conservan en PERSONAS, resumen y citas. Un nombre coincidente puede ser un destinatario o una mención y se presenta como tal.
