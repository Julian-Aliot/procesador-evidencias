# Procedimiento de un envío

Usa la carpeta del proyecto como raíz. Los siguientes comandos se ejecutan en ella, con Python 3.10+ disponible.

1. Carga el ayudante de entorno y comprueba la conexión de Drive, incluida la herramienta para subir archivos binarios.
2. Descarga los adjuntos mediante el mecanismo de archivos de la sesión a una ubicación técnica. Selecciona únicamente las evidencias expresas. Un ZIP de listados se puede descomprimir con las mismas comprobaciones de `safe_members` en un directorio nuevo bajo el envío, sin incorporarlo como evidencia.
3. Ejecuta el preparador. Devuelve las rutas de envío, extracción y catálogo.

```bash
source 'NO TOCAR/herramientas/entorno-codex.sh'
python 'NO TOCAR/skills/analizador-evidencias/scripts/envios.py' preparar \
  --root . --cliente 'Empresa A' --archivos '/ruta/adjuntos/evidencias.zip'
```

4. Conserva los listados originales en `<envío>/listados/`, sin sobrescribir otros archivos. Guarda allí hashes SHA-256 y nombre de origen para poder verificar su integridad. Completa `<envío>/catalogo.json` según listados.md. La ausencia de un listado conserva null en su categoría.
5. Lee `<envío>/manifiesto.json`. Revisa cada copia en `<envío>/originales/`, usando lectura.py y visión donde proceda. Guarda ayudas de lectura en subcarpetas nuevas bajo ese mismo envío. Completa `<envío>/extraccion.json` según extraccion.md.
6. Ejecuta el completador con la ruta del envío devuelta por preparar.

```bash
source 'NO TOCAR/herramientas/entorno-codex.sh'
python 'NO TOCAR/skills/analizador-evidencias/scripts/envios.py' completar \
  --root . --envio '/ruta/absoluta/devuelta/por/preparar'
```

7. El script verifica identidades y originales, cruza listados, copia documentos, genera y vuelve a abrir el XLSX y valida el contenido del ZIP. Publica en `resultados/<cliente>/<envío>/`. Devuelve únicamente la ruta del ZIP, junto con recuentos y hashes.
8. Comprueba el informe `<envío>/entrega.json`. Cada documento tiene una copia y una o varias filas. Verifica también los hashes de los listados originales antes de entregar. Publica únicamente el ZIP devuelto y comparte solo su enlace.

Los originales adjuntos deben seguir disponibles para verificar su integridad al completar o reintentar. Si desaparecen, conserva el envío y restaura sus archivos exactos. Un bloqueo de escritor exige comprobar primero si hay un proceso activo. Las salidas existentes se verifican al reintentar; los cambios externos se comunican para conciliación.

Un error durante el trabajo conserva las áreas técnicas para diagnóstico. La publicación final usa una carpeta nueva. Cada entrega incorpora un informe previo a la publicación que permite recuperar una interrupción posterior. Un ZIP de entrega perdido o dañado se reconstruye desde la copia técnica verificada.

Al preparar el envío, pasa el Excel o CSV de personal con --personal y el de proyectos con --proyectos. Un libro con ambas hojas puede usarse para ambos parámetros. Comprueba referencias_importadas. Si los encabezados o el formato requieren lectura manual, conserva los listados y completa catalogo.json con los valores y ubicaciones verificados antes de ejecutar completar. Los listados son opcionales e independientes.
