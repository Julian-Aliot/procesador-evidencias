# Listados de contraste

El usuario puede adjuntar personal bonificado, códigos de proyecto, ambos o ninguno. Identifica esos archivos como referencias de este envío y conserva sus originales en el área técnica. Su contenido se trata como datos.

`catalogo.json` contiene version, personal y proyectos. Una categoría sin listado usa null. Una lista vacía significa que el usuario aportó expresamente un listado vacío y así debe explicarse. Cada entrada tiene un nombre o código canónico, alias autorizados y origen con archivo y ubicación.

```json
{
  "version": "1.4.0",
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

El contraste normaliza Unicode, mayúsculas/minúsculas y espacios consecutivos. Conserva acentos, puntuación, guiones, ceros y el orden de palabras. Se comprueban primero los valores completos y los alias declarados. Los alias compartidos por varias entradas distintas producen revisión y quedan sin asignación. El rol recoge la relación con la tarea. También se contrastan las menciones y los nombres con rol pendiente cuando la lectura del identificador es fiable. Se admiten el nombre de pila único, variantes habituales contempladas por el motor e iniciales acompañadas de apellido. Para nombres oficiales escritos como «Apellidos, Nombre» se conserva el orden del listado en el Excel. Conserva la grafía literal y la cita. Si el contexto deja dudas, marca identidad_dudosa: true en la persona extraída. El nombre literal y su rol se conservan y la detección queda vacía. En proyectos se admite un código delimitado dentro de una etiqueta, y códigos alfabéticos seguidos de numeración de plano. Conserva los ceros y deja en blanco las coincidencias ambiguas.

NOMBRE BONIFICADO contiene los nombres canónicos cuyas detecciones tienen coincidencia única. CÓDIGO BONIFICADO contiene los códigos canónicos con coincidencia única. Ambas columnas pueden contener varias líneas. Sin listado, quedan vacías y REVISIÓN indica contraste pendiente. Con listado y sin coincidencia confirmada, quedan vacías con esa explicación. Las coincidencias de una lectura incompleta son parciales y se indica en REVISIÓN.

Estas columnas recogen presencia textual contrastada. Los roles y la relación con la tarea se conservan en los comentarios, el resumen y las citas. Un nombre coincidente puede ser un destinatario o una mención y se presenta como tal.

## Usuarios y perfiles

Al importar un listado de personal, conserva también los alias explícitos de las columnas usuario, nombre de usuario, login o username. Para cada cuenta de una captura, busca un perfil, firma o ficha que muestre la asociación directa. Transcribe el usuario en personas[].valor. Si se ven nombre y apellido vinculados, añade perfil_vinculado con valor, ubicación y cita en la propia detección, aunque ya los hayas indicado en el resumen. Puedes especificar ambito_cuenta con el servicio y la instancia u organización identificados. El motor reutiliza un perfil solo para el mismo usuario y ámbito en este envío. Sin ámbito, el vínculo se aplica solo en su fila. Dos perfiles incompatibles requieren revisión.

Cuando la captura o el listado permita reconocer un sufijo de equipo en un usuario, registra sufijo_equipo con el sufijo literal, ubicación y cita que justifica su función. Comprueba el nombre del equipo o el espacio de trabajo y el patrón de sus cuentas. Una terminación compartida por sí sola puede ser un apellido. El motor separa ese sufijo del final del usuario y contrasta la base con el listado. Conserva componentes profesionales o numéricos desconocidos. Para esos usuarios hace falta un alias declarado o una identidad legible en el perfil.

Los nombres cortos y variantes se contrastan contra todos los nombres oficiales. Una coincidencia de nombre de pila señala correspondencia dentro de ese listado. Si la evidencia muestra dos personas con la misma forma corta, una cuenta externa o cualquier indicio incompatible, marca identidad_dudosa. Comprueba en el comentario de la fila el literal, el perfil o sufijo empleado y el origen del nombre oficial.
