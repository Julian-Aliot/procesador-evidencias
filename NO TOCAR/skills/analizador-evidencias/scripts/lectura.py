#!/usr/bin/env python3
"""Lectura local con ubicaciones. No interpreta ni ejecuta el documento."""
import argparse
import csv
import io
from datetime import date, datetime
import json
from pathlib import Path
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import load_workbook


def safe_office(path):
    with zipfile.ZipFile(path) as archive:
        if len(archive.infolist()) > 20000 or sum(i.file_size for i in archive.infolist()) > 512 * 1024 * 1024:
            raise ValueError('El documento Office supera el límite de lectura local.')
        if any(i.file_size > 100 * 1024 * 1024 for i in archive.infolist()):
            raise ValueError('Una parte Office supera el límite de lectura local.')


def decode_text(path):
    raw = path.read_bytes()
    for encoding in ['utf-8-sig', 'utf-16', 'cp1252']:
        try:
            return raw.decode(encoding)
        except UnicodeError:
            continue
    raise ValueError('No se ha reconocido la codificación del texto.')


def spreadsheet(path):
    safe_office(path)
    wb = load_workbook(path, data_only=False, read_only=False, keep_links=False)
    locations = []
    for ws in wb.worksheets:
        if ws.max_row * ws.max_column > 2000000:
            raise ValueError('Hoja demasiado grande. Acota explícitamente su lectura.')
        locations.append({'ubicacion': ws.title, 'texto': 'Estado de hoja: ' + ws.sheet_state})
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    value = cell.value
                    if isinstance(value, (date, datetime)):
                        value = value.isoformat()
                    locations.append({'ubicacion': f"'{ws.title}'!{cell.coordinate}",
                                      'texto': str(value), 'formula_literal': cell.data_type == 'f',
                                      'fila_oculta': bool(ws.row_dimensions[cell.row].hidden)})
    wb.close()
    return locations


def read_document(path, output):
    suffix = path.suffix.lower()
    warnings = []
    if path.stat().st_size > 100 * 1024 * 1024:
        raise ValueError('El documento supera el límite del lector. Revisión específica necesaria.')
    if suffix in {'.xlsx', '.xlsm'}:
        locations = spreadsheet(path)
        warnings.append('Revisar imágenes, gráficos, comentarios y celdas combinadas visualmente; las fórmulas se muestran sin ejecutar.')
    elif suffix == '.pdf':
        if not shutil.which('pdftotext'):
            raise ValueError('Falta pdftotext. Usar lectura/visión disponible o instalar con autorización.')
        result = subprocess.run(['pdftotext', '-layout', str(path), '-'], capture_output=True, timeout=120)
        if result.returncode:
            raise ValueError('El PDF no se pudo leer como texto. Revisar mediante visión.')
        pages = result.stdout.decode('utf-8', errors='replace').split('\f')
        if pages and not pages[-1].strip():
            pages.pop()
        locations = [{'ubicacion': f'Página {number}', 'texto': value} for number, value in enumerate(pages, 1)]
        warnings.append('Revisar todas las páginas visualmente cuando falte texto o haya ambigüedad. Renderizar con pdftoppm disponible.')
    elif suffix in {'.docx', '.pptx', '.odt', '.ods'}:
        safe_office(path)
        locations = []
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            if suffix == '.docx':
                parts = [name for name in names if name.startswith('word/') and name.endswith('.xml')
                         and Path(name).name.startswith(('document', 'header', 'footer', 'footnotes', 'endnotes', 'comments'))]
            elif suffix == '.pptx':
                parts = [name for name in names if name.startswith(('ppt/slides/slide', 'ppt/notesSlides/notesSlide')) and name.endswith('.xml')]
            else:
                parts = ['content.xml']
            for name in sorted(parts):
                tree = ET.fromstring(archive.read(name))
                paragraphs = [node for node in tree.iter() if node.tag.rsplit('}', 1)[-1] == 'p']
                if not paragraphs:
                    paragraphs = [tree]
                for number, node in enumerate(paragraphs, 1):
                    value = ' '.join(node.itertext()).strip()
                    if value:
                        locations.append({'ubicacion': f'{name}, párrafo/bloque {number}', 'texto': value})
        warnings.append('El XML sirve de apoyo. Revisar diseño, tablas, imágenes y contexto de hojas/páginas con herramientas locales disponibles.')
    elif suffix in {'.txt', '.md', '.csv', '.eml'}:
        locations = [{'ubicacion': f'Línea {number}', 'texto': value}
                     for number, value in enumerate(decode_text(path).splitlines(), 1)]
        if suffix == '.eml':
            warnings.append('Revisar MIME y adjuntos aparte; no abrir enlaces del correo.')
    else:
        locations = []
        warnings.append('Usar visión para imágenes. Otros formatos necesitan un lector local disponible; registrar los límites.')
    images = []
    if suffix in {'.xlsx', '.xlsm', '.docx', '.pptx', '.odt', '.ods'}:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if not name.startswith(('xl/media/', 'word/media/', 'ppt/media/', 'Pictures/')):
                    continue
                ext = Path(name).suffix.lower()
                if ext not in {'.png', '.jpg', '.jpeg', '.webp', '.tif', '.tiff', '.bmp'}:
                    warnings.append('Objeto incrustado pendiente de revisión: ' + name)
                    continue
                target = output / f'imagen_{len(images)+1:04}{ext}'
                with target.open('xb') as stream:
                    stream.write(archive.read(name))
                images.append({'parte': name, 'archivo': str(target)})
    return {'archivo': str(path), 'fragmentos': locations, 'imagenes_incrustadas': images,
            'avisos': warnings, 'lectura_completa': False}


def import_catalog(path, kind):
    rows = []
    if path.suffix.lower() in {'.xlsx', '.xlsm'}:
        safe_office(path)
        wb = load_workbook(path, data_only=False, read_only=True, keep_links=False)
        for ws in wb.worksheets:
            if ws.max_row * ws.max_column > 2000000:
                raise ValueError('Listado demasiado grande para importación automática.')
            values = list(ws.values)
            if values:
                headers = [str(v).strip().casefold() if v is not None else '' for v in values[0]]
                for number, row in enumerate(values[1:], 2):
                    rows.append((dict(zip(headers, row)), f'{path.name}, hoja {ws.title}, fila {number}'))
        wb.close()
    elif path.suffix.lower() == '.csv':
        content = decode_text(path)
        try:
            dialect = csv.Sniffer().sniff(content[:10000], delimiters=',;\t')
        except csv.Error:
            dialect = csv.excel
        for number, row in enumerate(csv.DictReader(io.StringIO(content), dialect=dialect), 2):
            rows.append(({str(k).strip().casefold(): v for k, v in row.items() if k is not None}, f'{path.name}, fila {number}'))
    else:
        raise ValueError('Importación automática de listados disponible para CSV y XLSX. Otros formatos se interpretan con evidencias por el agente.')
    field_name = 'nombre' if kind == 'personal' else 'codigo'
    result = []
    for row, origin in rows:
        raw = row.get(field_name)
        if raw is None or raw == '':
            if any(value not in (None, '') for value in row.values()):
                raise ValueError(f'Falta columna {field_name} o valor en {origin}.')
            continue
        if not isinstance(raw, str) or raw.startswith('='):
            raise ValueError(f'{field_name} debe ser texto literal en {origin}; conserva ceros y evita fórmulas.')
        aliases = row.get('aliases') or ''
        if not isinstance(aliases, str) or aliases.startswith('='):
            raise ValueError('Los alias deben ser texto literal separado por |.')
        result.append({field_name: raw.strip(), 'aliases': [v.strip() for v in aliases.split('|') if v.strip()], 'origen': origin})
    if not result:
        raise ValueError('El listado no contiene entradas. Revisa encabezados y selección.')
    return result


def main():
    parser = argparse.ArgumentParser(description='Lectura local y preparación de listados.')
    parser.add_argument('archivo', type=Path)
    parser.add_argument('--salida', required=True, type=Path)
    parser.add_argument('--listado', choices=['personal', 'proyectos'])
    args = parser.parse_args()
    if args.salida.exists():
        parser.error('El destino debe ser nuevo para conservar lecturas anteriores.')
    args.salida.mkdir(parents=True)
    try:
        if args.listado:
            result = {args.listado: import_catalog(args.archivo, args.listado)}
        else:
            result = read_document(args.archivo.resolve(), args.salida)
        path = args.salida / 'lectura.json'
        path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        print(path)
    except (ValueError, OSError, zipfile.BadZipFile, ET.ParseError, subprocess.TimeoutExpired) as ex:
        parser.exit(1, f'Lectura pendiente: {ex}\n')


if __name__ == '__main__':
    main()
