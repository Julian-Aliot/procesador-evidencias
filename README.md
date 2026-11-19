# Analizador de evidencias

Versión 1.0.0. Proyecto independiente para usar en su propio chat con Codex. Analiza evidencias de tareas, trabajos o participación en proyectos y entrega un Excel de once columnas y un ZIP de documentos en Google Drive privado.

## Uso sencillo

Abre esta carpeta como raíz de proyecto. Escribe «Lee AGENTS.md y sigue la Skill de este proyecto». Después adjunta evidencias y pide «Analiza estas evidencias para Empresa A». Puedes añadir listados de personal bonificado y códigos de proyecto indicando que son referencias para contrastar.

Cada envío es independiente y conserva los anteriores. Una tarea genera una fila. Un documento con varias tareas justificadas genera varias filas que apuntan a su copia. Los datos desconocidos quedan pendientes de revisión. Una persona mencionada conserva su rol para poder interpretar la evidencia.

Las columnas PERSONAL BONIFICADO DETECTADO y CÓDIGO DE PROYECTO DETECTADO se calculan a partir de coincidencias únicas contra los listados aportados. Los alias requieren autorización o presencia explícita en el listado. Sin listado, ambas columnas permanecen disponibles y se señala contraste pendiente.

## Archivos

`entrada/` puede contener documentos seleccionados expresamente por el usuario. Los resultados se publican en `resultados/<cliente>/<envío>/`. Las ayudas y los controles viven bajo `NO TOCAR/`. El ZIP contiene el Excel y `evidencias/`, con enlaces relativos. Descarga y extrae el ZIP para usar esos enlaces.

## Dependencias y lectura

Se necesitan Python 3.10+ y openpyxl 3.1. El ayudante `NO TOCAR/herramientas/entorno-codex.sh` comprueba dependencias y lectores disponibles. La instalación se explica y autoriza antes de ejecutarla. El entorno de Codex debe proporcionar visión para imágenes y escaneos, lectores PDF locales para extraer/renderizar y una conexión Drive con subida binaria para entregar.

Los scripts de lectura apoyan PDF, imágenes mediante visión del agente, XLSX/XLSM, texto y formatos Office/OpenDocument. XLS y DOC necesitan un lector compatible disponible. Los documentos que no se puedan leer se conservan y se explican como pendientes. El agente interpreta los documentos con evidencias de origen; los scripts realizan los controles y generan el Excel localmente con openpyxl.

## Publicación en tu nuevo repositorio

Descomprime el paquete y sube el contenido de esta carpeta a un repositorio propio. Incluye los archivos ocultos `.gitignore` y `.gitattributes`. Las carpetas de datos están excluidas de Git. Publica antes de introducir documentos personales. Conserva `.gitkeep` en las carpetas vacías.

Tras publicarlo puedes configurar la URL real en `NO TOCAR/config/config.json`, campo repositorio_url. También se reconoce el remoto origin de un clon del nuevo repositorio. Los envíos posteriores comprobarán la rama main según la referencia versiones.md.

## Arranque de otro chat

Puedes pegar este texto y sustituir la URL cuando exista.

> Descarga la rama main de MI_REPOSITORIO en una carpeta nueva y ábrela como raíz del proyecto. Lee AGENTS.md, NO TOCAR/AGENTS.md y NO TOCAR/skills/analizador-evidencias/SKILL.md con sus referencias. Comprueba las dependencias y la subida a mi Drive conectado. Cada envío es independiente. Analiza únicamente los adjuntos que seleccione, conserva anteriores y entrega Excel y ZIP privados. Espera a que adjunte evidencias e indique el cliente. Si aporto listados de personal bonificado o códigos de proyecto, úsalos para las dos columnas de detección.

## Verificación del programa

Las pruebas locales usan datos sintéticos dentro de directorios temporales. Comprueban separación de envíos, enlaces, integridad, listados, protección de fórmulas y seguridad ZIP.

```bash
python -m unittest discover -s 'NO TOCAR/tests' -v
```

Este paquete contiene programa, documentación y plantillas vacías. Las entregas de usuarios contienen sus propios resultados y se guardan privadas en Drive.
