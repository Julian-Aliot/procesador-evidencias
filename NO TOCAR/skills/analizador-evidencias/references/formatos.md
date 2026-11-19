# Lectura de documentos

`scripts/lectura.py` genera apoyos de lectura con ubicaciones y extrae imágenes Office a una carpeta técnica nueva. Su salida siempre conserva lectura_completa=false para que el agente la confirme tras revisar contenido y contexto.

```bash
python 'NO TOCAR/skills/analizador-evidencias/scripts/lectura.py' '/ruta/documento.xlsx' \
  --salida '/ruta/envio/lectura-D00001'
```

PDF digital usa pdftotext si está instalado. Un escaneo requiere visión y renderizado local, por ejemplo pdftoppm, cuando esté disponible. Revisa todas las páginas, aunque algunas ya tengan texto. Las imágenes se inspeccionan con visión de la sesión. TIFF multipágina requiere revisar todos los fotogramas con un lector local.

XLSX y XLSM usan openpyxl con fórmulas literales, enlaces externos sin resolver y macros sin ejecutar. Revisa hojas ocultas, celdas combinadas, comentarios, imágenes incrustadas y gráficos cuando contengan evidencias. El helper extrae imágenes Office y enumera celdas. La conexión entre una imagen y su hoja requiere revisar los anclajes o el libro visualmente.

DOCX, PPTX, ODT y ODS admiten apoyo de texto XML e imágenes incrustadas. Revisa tablas, notas y ubicación visual para comprender relaciones. El XML conserva ubicaciones técnicas cuando no haya paginación estable. TXT, Markdown, CSV y EML se leen como texto con líneas; el correo puede necesitar revisión MIME y adjuntos específicos.

XLS y DOC antiguos necesitan un lector local compatible. Si LibreOffice está disponible, puede convertir una copia a un área técnica para lectura, manteniendo el original intacto. Comprueba primero su disponibilidad y los permisos del entorno. Ante falta de lector, formato cifrado, documento corrupto, gráfico no interpretable o límite de tamaño, conserva copia y fila y explica qué queda pendiente.

Los ZIP se descomprimen una sola capa. Los archivos ZIP anidados permanecen como documentos pendientes hasta que el usuario seleccione expresamente su contenido y se vuelva a validar su extracción. Los límites predeterminados son configurables en config.json. El preparador admite hasta 10000 documentos, 512 MiB por archivo y 2 GiB por envío, con control de expansión. El lector auxiliar usa límites más reducidos para formatos estructurados.
