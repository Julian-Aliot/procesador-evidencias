#!/usr/bin/env bash
# Cargar desde la raíz del proyecto. No instala dependencias.
python - <<'PY'
import importlib.util
import shutil
import sys
if sys.version_info < (3, 10):
    raise SystemExit('Se necesita Python 3.10 o posterior.')
if importlib.util.find_spec('openpyxl') is None:
    raise SystemExit('Falta openpyxl. Explica su necesidad y solicita autorización antes de instalarlo.')
import openpyxl
print(f'Python {sys.version.split()[0]}; openpyxl {openpyxl.__version__}. Excel disponible.')
for program in ['pdftotext', 'pdftoppm', 'pdfinfo', 'libreoffice']:
    print(program + ' ' + (shutil.which(program) or 'no disponible'))
print('Comprueba también visión en las herramientas de la sesión para imágenes y escaneos.')
PY
