#!/usr/bin/env python3
"""Operaciones locales. La interpretación de los documentos corresponde al agente."""
import argparse
from contextlib import contextmanager
from datetime import date, datetime, timezone
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import unicodedata
from urllib.parse import quote, unquote
import uuid
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

VERSION = '1.2.0'
HEADERS = ['PERSONAS', 'CÓDIGO PROYECTO', 'NOMBRE PROYECTO', 'FECHA',
           'TÍTULO TAREA', 'RESUMEN EVIDENCIA', 'UBICACIÓN', 'ARCHIVO EVIDENCIA',
           'REVISIÓN', 'PERSONAL BONIFICADO DETECTADO', 'CÓDIGO DE PROYECTO DETECTADO']
SUMMARY_HEADERS = ['NOMBRE PRINCIPAL DE LA EVIDENCIA', 'OTROS NOMBRES', 'CÓDIGO DE PROYECTO',
                   'CÓDIGO DETECTADO', 'FECHA', 'TÍTULO', 'RESUMEN', 'NOMBRE DETECTADO', 'DESCRIPCIÓN TAREA']
SUPPORTED = {'.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.webp', '.bmp',
             '.xlsx', '.xlsm', '.xls', '.ods', '.csv', '.docx', '.doc', '.odt',
             '.pptx', '.txt', '.md', '.eml'}
SCALARS = ['nombre_proyecto', 'fecha', 'titulo_tarea', 'resumen']
ROLES = {'participante', 'autor', 'responsable', 'destinatario', 'mencionado', 'no_determinado'}


def read_json(path):
    with Path(path).open(encoding='utf-8') as stream:
        return json.load(stream)


def write_json(path, value):
    path = Path(path)
    temp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    with temp.open('x', encoding='utf-8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.write('\n')
    os.replace(temp, path)


def digest(path):
    result = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            result.update(block)
    return result.hexdigest()


def simple_name(value):
    if not isinstance(value, str) or not value.strip() or value.strip() in {'.', '..'}:
        raise ValueError('Indica un nombre de cliente sencillo.')
    if any(c in value for c in '/\\\x00') or re.search(r'[<>:"|?*\x00-\x1f]', value):
        raise ValueError('El cliente debe ser un nombre, sin separadores ni caracteres de ruta.')
    return value.strip()


def slug(value):
    value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode()
    return re.sub(r'[^a-zA-Z0-9._-]+', '_', value).strip('._')[:70] or 'documento'


def normalized(value):
    """Conserva orden de palabras, ceros, guiones y puntuación."""
    return ' '.join(unicodedata.normalize('NFKC', value).casefold().split())


def config(root):
    root = Path(root).resolve()
    value = read_json(root / 'NO TOCAR/config/config.json')
    if value.get('programa') != 'analizador-evidencias':
        raise ValueError('La raíz no pertenece a este analizador de evidencias.')
    return root, value


def contained(root, relative):
    path = (root / relative).resolve()
    if not path.is_relative_to(root):
        raise ValueError('Ruta fuera del envío.')
    return path


@contextmanager
def lock(envio):
    path = envio / 'escritor.lock'
    try:
        descriptor = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError as ex:
        raise ValueError('Envío bloqueado. Comprueba el proceso escritor antes de retirar su bloqueo.') from ex
    try:
        os.write(descriptor, str(os.getpid()).encode())
        os.close(descriptor)
        yield
    finally:
        path.unlink(missing_ok=True)


def selected_files(paths):
    found = []
    for raw in paths:
        source = Path(raw).absolute()
        if source.is_symlink():
            raise ValueError(f'Enlace simbólico rechazado: {source.name}')
        if source.is_dir():
            members = sorted(source.rglob('*'))
            if any(p.is_symlink() for p in members):
                raise ValueError('La selección contiene enlaces simbólicos.')
            for path in members:
                if path.is_file():
                    found.append((path, source.name + '/' + path.relative_to(source).as_posix()))
                elif not path.is_dir():
                    raise ValueError('La selección contiene un archivo especial.')
        elif source.is_file():
            found.append((source, source.name))
        else:
            raise ValueError(f'Archivo inaccesible: {source.name}')
    resolved = [p.resolve() for p, _ in found]
    if len(set(resolved)) != len(resolved):
        raise ValueError('La selección solapa archivos. Selecciona cada ruta una vez.')
    return found


def safe_members(archive, limits):
    members = archive.infolist()
    if len(members) > limits['max_archivos']:
        raise ValueError('El ZIP supera el límite de entradas.')
    seen = set()
    total = 0
    for item in members:
        name = item.filename
        path = PurePosixPath(name)
        mode = item.external_attr >> 16
        if (not name or '\\' in name or '\x00' in name or ':' in name or path.is_absolute()
                or '..' in path.parts or name.startswith('/')
                or stat.S_ISLNK(mode) or stat.S_IFMT(mode) not in (0, stat.S_IFREG, stat.S_IFDIR)):
            raise ValueError('ZIP con ruta o tipo de archivo inseguro.')
        key = normalized(path.as_posix().rstrip('/'))
        if key in seen:
            raise ValueError('ZIP con nombres repetidos o ambiguos.')
        seen.add(key)
        total += item.file_size
        if item.flag_bits & 1:
            raise ValueError('El ZIP está cifrado.')
        if (item.file_size > limits['max_bytes_archivo'] or total > limits['max_bytes_envio']
                or item.file_size > max(1, item.compress_size) * limits['max_ratio_zip']):
            raise ValueError('El ZIP supera los límites de descompresión configurados.')
    return [item for item in members if not item.is_dir()]


def field():
    return {'valor': None, 'ubicacion': None, 'cita': None}


def empty_activity(reason):
    return {'personas': [], 'codigos_proyecto': [],
            **{key: field() for key in SCALARS}, 'tipo_fecha': None,
            'motivos_revision': [reason]}


def prepare(root, cliente, archivos):
    root, cfg = config(root)
    cliente = simple_name(cliente)
    sources = selected_files(archivos)
    if not sources:
        raise ValueError('La selección no contiene archivos.')
    identifier = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ') + '_' + uuid.uuid4().hex[:10]
    envio = root / 'NO TOCAR/trabajo/envios' / identifier
    envio.mkdir(parents=True, exist_ok=False)
    (envio / 'originales').mkdir()
    documents, source_info = [], []
    size = 0

    def snapshot(stream, label, suffix):
        nonlocal size
        if len(documents) >= cfg['limites']['max_archivos']:
            raise ValueError('El envío supera el límite de documentos.')
        doc_id = f'D{len(documents) + 1:05}'
        name = f'{doc_id}_{slug(Path(label).stem)}{suffix.lower()}'
        destination = envio / 'originales' / name
        count = 0
        with destination.open('xb') as output:
            for block in iter(lambda: stream.read(1024 * 1024), b''):
                count += len(block)
                size += len(block)
                if count > cfg['limites']['max_bytes_archivo'] or size > cfg['limites']['max_bytes_envio']:
                    raise ValueError('El envío supera los límites de tamaño.')
                output.write(block)
        documents.append({'id': doc_id, 'nombre_original': label,
                          'ruta': destination.relative_to(envio).as_posix(),
                          'sha256': digest(destination), 'bytes': count,
                          'formato_previsto': suffix.lower() in SUPPORTED})

    for source, label in sources:
        before = digest(source)
        source_info.append({'ruta': str(source), 'sha256': before})
        if source.suffix.lower() == '.zip':
            with zipfile.ZipFile(source) as archive:
                for item in safe_members(archive, cfg['limites']):
                    with archive.open(item) as stream:
                        snapshot(stream, label + ' / ' + item.filename, Path(item.filename).suffix)
        else:
            with source.open('rb') as stream:
                snapshot(stream, label, source.suffix)
        if digest(source) != before:
            raise ValueError('Un original ha cambiado durante la preparación.')
    if not documents:
        raise ValueError('El ZIP no contiene documentos.')
    manifest = {'version': VERSION, 'envio_id': identifier, 'cliente': cliente,
                'fuentes': source_info, 'documentos': documents}
    write_json(envio / 'manifiesto.json', manifest)
    write_json(envio / 'extraccion.json', {
        'version': VERSION, 'envio_id': identifier,
        'documentos': [{'id': d['id'], 'sha256': d['sha256'], 'lectura_completa': False,
                        'incidencias_documento': [],
                        'resumen_documento': {'descripcion_tarea': field()},
                        'actividades': [empty_activity('Pendiente de lectura por el agente.')]} for d in documents]})
    write_json(envio / 'catalogo.json', {'version': VERSION, 'personal': None, 'proyectos': None})
    write_json(envio / 'estado.json', {'version': VERSION, 'envio_id': identifier,
                                    'manifiesto_sha256': digest(envio / 'manifiesto.json')})
    return {'envio': str(envio), 'envio_id': identifier, 'documentos': len(documents),
            'extraccion': str(envio / 'extraccion.json'), 'catalogo': str(envio / 'catalogo.json')}


def text(value, label):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f'{label} debe contener texto.')
    if len(value) > 30000:
        raise ValueError(f'{label} supera el tamaño admitido por Excel. Divide la actividad justificadamente.')
    return value.strip()


def text_list(value, label):
    if not isinstance(value, list):
        raise ValueError(f'{label} debe ser una lista.')
    return [text(v, label) for v in value]


def validate_field(value, label):
    if not isinstance(value, dict) or 'valor' not in value:
        raise ValueError(f'Falta el objeto con valor, ubicación y cita para {label}.')
    if value['valor'] is None:
        return None
    result = text(value['valor'], label)
    text(value.get('ubicacion'), f'Ubicación de {label}')
    text(value.get('cita'), f'Cita de {label}')
    return result


def catalog_indexes(catalog):
    indexes = {}
    for group, key in [('personal', 'nombre'), ('proyectos', 'codigo')]:
        entries = catalog.get(group)
        if entries is None:
            indexes[group] = None
            continue
        if not isinstance(entries, list):
            raise ValueError('El catálogo debe contener listas o null.')
        index = {}
        for entry in entries:
            canonical = text(entry.get(key), f'Catálogo {key}')
            text(entry.get('origen'), 'Origen del listado y su ubicación')
            aliases = text_list(entry.get('aliases', []), 'Alias autorizados')
            for alias in [canonical] + aliases:
                index.setdefault(normalized(alias), set()).add(canonical)
        indexes[group] = index
    return indexes


def name_parts(value):
    plain = ''.join(c for c in unicodedata.normalize('NFD', normalized(value))
                    if not unicodedata.combining(c))
    return re.findall(r"[^\W\d_]+", plain)


def short_name(observed, full, identified):
    if any(char.isdigit() for char in observed) or not all(c.isalpha() or c in " .-’'" for c in observed):
        return False
    words, target = name_parts(observed), name_parts(full)
    if not words or not target or len(words) > len(target):
        return False
    if words == target:
        return True
    if len(words) == 1:
        if not identified:
            return False
        return (len(words[0]) >= 3 and words[0] == target[0]) or (
            len(words[0]) >= 4 and target[0].startswith(words[0]))
    if not any(len(word) >= 3 and word == ref for word, ref in zip(words, target)):
        return False
    return all(word == ref or (ref.startswith(word) and
               (len(word) == 1 or (i == 0 and len(word) >= 3)))
               for i, (word, ref) in enumerate(zip(words, target)))


def embedded_code(value, code):
    # Los delimitadores incluyen guiones bajos de nombres de planos y archivos.
    if code.isdecimal():
        return re.search(r'(?:proyecto|project|código|codigo)\s*[:#_-]?\s*' + re.escape(code)
                         + r'(?![^\W_])', value) is not None
    if not code or len(code) < 2:
        return False
    after = r'(?![^\W_])'
    if code.isalpha() and len(code) >= 3:
        after = r'(?:[0-9]+)?(?![^\W_])'
    return re.search(r'(?<![^\W_])' + re.escape(code) + after, value) is not None


def matches(values, index, label, reasons, identified=False):
    if index is None:
        reasons.append(f'Listado de {label} no aportado; contraste pendiente.')
        return ''
    results = []
    for value in values:
        candidates = index.get(normalized(value), set())
        if not candidates:
            candidates = set()
            for alias, owners in index.items():
                if label == 'proyectos':
                    hit = embedded_code(normalized(value), alias)
                else:
                    hit = short_name(value, alias, identified)
                if hit:
                    candidates.update(owners)
        if len(candidates) > 1:
            reasons.append(f'Coincidencia ambigua en {label}: {value}.')
        elif len(candidates) == 1:
            canonical = next(iter(candidates))
            if canonical not in results:
                results.append(canonical)
    if not results:
        reasons.append(f'Sin coincidencias confirmadas en el listado de {label}.')
    return '\n'.join(results)


def build_rows(manifest, extraction, catalog):
    if extraction.get('envio_id') != manifest['envio_id']:
        raise ValueError('La extracción pertenece a otro envío.')
    documents = extraction.get('documentos')
    if not isinstance(documents, list) or len(documents) != len(manifest['documentos']):
        raise ValueError('Debe existir una extracción por documento.')
    indexed = {d['id']: d for d in documents}
    if len(indexed) != len(documents):
        raise ValueError('Identidades de extracción repetidas.')
    indexes = catalog_indexes(catalog)
    rows, counts = [], {}
    seen_hashes = {}
    for doc in manifest['documentos']:
        if doc['id'] not in indexed:
            raise ValueError('Falta un documento en la extracción.')
        item = indexed[doc['id']]
        if item.get('sha256') != doc['sha256']:
            raise ValueError('Se ha cambiado la identidad de un documento.')
        if type(item.get('lectura_completa')) is not bool:
            raise ValueError('lectura_completa debe ser true o false.')
        doc_reasons = text_list(item.get('incidencias_documento'), 'Incidencias de documento')
        if not item['lectura_completa']:
            doc_reasons.append('Lectura incompleta; revisión del documento pendiente.')
        if not doc['formato_previsto']:
            doc_reasons.append('Formato sin lector previsto; capacidad de lectura por verificar.')
        if doc['sha256'] in seen_hashes:
            doc_reasons.append('Archivo idéntico a ' + seen_hashes[doc['sha256']] + '; se conserva su copia.')
        seen_hashes.setdefault(doc['sha256'], doc['id'])
        activities = item.get('actividades')
        if not isinstance(activities, list) or not activities:
            raise ValueError('Cada documento debe tener al menos una actividad, incluso pendiente.')
        counts[doc['id']] = 0
        for number, activity in enumerate(activities, 1):
            reasons = doc_reasons + text_list(activity.get('motivos_revision'), 'Motivos de revisión')
            locations, citations, names, name_roles, codes = [], [], [], [], []
            for group in ['personas', 'codigos_proyecto']:
                values = activity.get(group)
                if not isinstance(values, list):
                    raise ValueError(group + ' debe ser una lista.')
                for element in values:
                    value = validate_field(element, group)
                    if value is None:
                        raise ValueError('Las listas contienen solo detecciones con valor y evidencia.')
                    locations.append(element['ubicacion'])
                    citations.append(group + ': ' + element['cita'])
                    if group == 'personas':
                        role = element.get('rol')
                        if role not in ROLES:
                            raise ValueError('Asigna un rol válido a cada persona detectada.')
                        names.append(value)
                        name_roles.append(value + ' (' + role.replace('_', ' ') + ')')
                        if role == 'no_determinado':
                            reasons.append('Papel de la persona sin determinar: ' + value)
                    else:
                        codes.append(value)
            scalars = {}
            for key in SCALARS:
                scalars[key] = validate_field(activity.get(key), key)
                if scalars[key] is not None:
                    locations.append(activity[key]['ubicacion'])
                    citations.append(key + ': ' + activity[key]['cita'])
            if not names:
                reasons.append('Personas no identificadas en esta actividad.')
            if not codes:
                reasons.append('Código de proyecto no identificado en esta actividad.')
            for key in ['fecha', 'titulo_tarea', 'resumen']:
                if scalars[key] is None:
                    reasons.append(key.replace('_', ' ').capitalize() + ' pendiente.')
            excel_date = None
            if scalars['fecha'] is not None:
                if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', scalars['fecha']):
                    raise ValueError('Las fechas deben ser ISO AAAA-MM-DD; las ambiguas quedan null.')
                excel_date = date.fromisoformat(scalars['fecha'])
                kind = activity.get('tipo_fecha')
                if kind not in {'realizacion', 'participacion', 'documento', 'envio', 'otra'}:
                    raise ValueError('Debe indicarse el tipo de fecha extraída.')
                if kind not in {'realizacion', 'participacion'}:
                    reasons.append('La fecha corresponde a ' + kind + '; la realización queda pendiente.')
            detected_codes = matches(codes, indexes['proyectos'], 'proyectos', reasons)
            if not item['lectura_completa']:
                reasons.append('Las detecciones reflejan únicamente el contenido legible revisado.')
            values = ['\n'.join(name_roles), '\n'.join(codes), scalars['nombre_proyecto'], excel_date,
                      scalars['titulo_tarea'], scalars['resumen'], '\n'.join(dict.fromkeys(locations)),
                      'Abrir evidencia', '\n'.join(dict.fromkeys(reasons)), '', detected_codes]
            if any(isinstance(v, str) and len(v) > 32767 for v in values):
                raise ValueError('Una celda supera 32767 caracteres; divide actividades con respaldo documental.')
            # Una fila por persona/tarea; no atribuir el resto de personas a esa fila.
            for person in activity['personas'] or [None]:
                person_values = list(values)
                person_reasons = list(reasons)
                if person is not None:
                    person_values[0] = person['valor'] + ' (' + person['rol'].replace('_', ' ') + ')'
                    if person.get('identidad_dudosa') is True:
                        person_reasons.append('Identidad dudosa según la evidencia: ' + person['valor'])
                    else:
                        person_values[9] = matches([person['valor']], indexes['personal'],
                                               'personal bonificado', person_reasons,
                                               identified=person['rol'] in {'autor', 'participante', 'responsable', 'destinatario'})
                else:
                    matches([], indexes['personal'], 'personal bonificado', person_reasons)
                person_values[8] = '\n'.join(dict.fromkeys(person_reasons))
                trace = []
                for entry in catalog.get('personal') or []:
                    if entry['nombre'] in person_values[9].split('\n'):
                        trace.append('Personal bonificado: ' + entry['nombre'] + ' | listado: ' + entry['origen'])
                rows.append({'documento_id': doc['id'], 'actividad': number, 'values': person_values,
                             'persona': person['valor'] if person else None,
                             'citas': '\n'.join(citations + trace)[:30000],
                             'copia': Path(doc['ruta']).name})
                counts[doc['id']] += 1
    return rows, counts


def build_summary_rows(manifest, extraction, rows):
    """Agrupa por identidad documental, nunca por hash o nombre de archivo."""
    summaries = []
    indexed = {d['id']: d for d in extraction['documentos']}
    for doc in manifest['documentos']:
        detail = [r for r in rows if r['documento_id'] == doc['id']]
        def unique(values):
            return list(dict.fromkeys(v for v in values if v is not None and v != ''))
        names = unique(r.get('persona') for r in detail)
        item = indexed[doc['id']]
        meta = item.get('resumen_documento', {})
        if not isinstance(meta, dict):
            raise ValueError('resumen_documento debe ser un objeto.')
        principal = validate_field(meta['nombre_principal'], 'nombre principal') if 'nombre_principal' in meta else (names[0] if len(names) == 1 else None)
        if principal is not None and principal not in names:
            raise ValueError('El nombre principal debe figurar entre las personas detectadas del documento.')
        def joined(column):
            return '\n'.join(unique(v for r in detail for v in (r['values'][column] or '').split('\n')))
        dates = unique(r['values'][3] for r in detail)
        title = validate_field(meta['titulo'], 'título del documento') if 'titulo' in meta else joined(4)
        summary = validate_field(meta['resumen'], 'resumen del documento') if 'resumen' in meta else joined(5)
        task = validate_field(meta['descripcion_tarea'], 'descripción tarea') if 'descripcion_tarea' in meta else None
        if extraction.get('version') == VERSION and summary and not task and item['lectura_completa'] and not item['incidencias_documento']:
            raise ValueError('Completa descripcion_tarea con la acción y su objeto o registra por qué no se puede identificar.')
        values = [principal, '\n'.join(n for n in names if n != principal), joined(1), joined(10),
                  dates[0] if len(dates) == 1 else None, title or 'Sin título identificado', summary, joined(9), task]
        if any(isinstance(v, str) and len(v) > 32767 for v in values):
            raise ValueError('Resumen documental demasiado largo; redacta resumen_documento de forma concisa.')
        notes = [doc['id'], doc['nombre_original'], 'Detalle en hoja Evidencias.']
        if len(dates) > 1:
            notes.append('Varias fechas; consultar el detalle: ' + ', '.join(str(d) for d in dates))
        notes.extend(unique(r['values'][8] for r in detail))
        notes.extend(r['citas'] for r in detail)
        for value in meta.values():
            if isinstance(value, dict) and value.get('valor') is not None:
                notes.append(str(value.get('ubicacion')) + ': ' + str(value.get('cita')))
        summaries.append({'documento_id': doc['id'], 'values': values,
                          'copia': Path(doc['ruta']).name, 'citas': '\n'.join(notes)[:30000]})
    return summaries


def make_excel(path, rows, summaries):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Evidencias'
    ws.append(HEADERS)
    widths = [34, 25, 32, 14, 45, 55, 35, 23, 55, 35, 30]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, index).column_letter].width = width
    for row_number, item in enumerate(rows, 2):
        for column, value in enumerate(item['values'], 1):
            cell = ws.cell(row_number, column, value)
            if isinstance(value, str):
                cell.data_type = 's'  # Incluye textos que comienzan por =, +, - o @.
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row_number, 4).number_format = 'dd/mm/yyyy'
        link = ws.cell(row_number, 8)
        link.hyperlink = 'evidencias/' + quote(item['copia'], safe='')
        link.font = Font(color='0563C1', underline='single')
        link.comment = Comment(item['documento_id'] + '\n' + item['citas'], 'Evidencias')
        ws.row_dimensions[row_number].height = 110
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.font = Font(color='FFFFFF', bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = 'A2'
    table = Table(displayName='TablaEvidencias', ref=f'A1:K{len(rows) + 1}')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(table)
    summary = wb.create_sheet('Resumen', 0)
    summary.append(SUMMARY_HEADERS)
    for col, width in enumerate([34, 42, 28, 28, 16, 50, 80, 38, 55], 1):
        summary.column_dimensions[summary.cell(1, col).column_letter].width = width
    for number, item in enumerate(summaries, 2):
        for col, value in enumerate(item['values'], 1):
            cell = summary.cell(number, col, value)
            if isinstance(value, str):
                cell.data_type = 's'
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        summary.cell(number, 5).number_format = 'dd/mm/yyyy'
        link = summary.cell(number, 6)
        link.hyperlink = 'evidencias/' + quote(item['copia'], safe='')
        link.font = Font(color='0563C1', underline='single')
        link.comment = Comment(item['citas'], 'Evidencias')
        summary.row_dimensions[number].height = 110
    for cell in summary[1]:
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.font = Font(color='FFFFFF', bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    summary.row_dimensions[1].height = 36
    summary.freeze_panes = 'A2'
    table = Table(displayName='TablaResumen', ref=f'A1:I{len(summaries) + 1}')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    summary.add_table(table)
    wb.active = 0
    wb.save(path)


def verify_excel(path, rows, summaries):
    wb = load_workbook(path)
    if wb.sheetnames != ['Resumen', 'Evidencias']:
        raise ValueError('El libro debe tener las hojas Resumen y Evidencias.')
    ws = wb['Evidencias']
    if [c.value for c in ws[1]] != HEADERS or ws.max_row != len(rows) + 1 or ws.max_column != 11:
        raise ValueError('Columnas o recuentos de Excel incorrectos.')
    for number, item in enumerate(rows, 2):
        for column, expected in enumerate(item['values'], 1):
            cell = ws.cell(number, column)
            actual = cell.value
            if isinstance(actual, datetime):
                actual = actual.date()
            if actual != (expected if expected != '' else None):
                raise ValueError('El Excel no conserva los valores extraídos.')
            if isinstance(expected, str) and expected and cell.data_type != 's':
                raise ValueError('Un texto se ha convertido en fórmula o número.')
        link = ws.cell(number, 8).hyperlink
        if link is None or link.target != 'evidencias/' + quote(item['copia'], safe=''):
            raise ValueError('Enlace relativo incorrecto.')
        if not (path.parent / unquote(link.target)).is_file():
            raise ValueError('El enlace apunta a una copia ausente.')
    summary = wb['Resumen']
    if ([c.value for c in summary[1]] != SUMMARY_HEADERS or summary.max_row != len(summaries) + 1
            or summary.max_column != len(SUMMARY_HEADERS)):
        raise ValueError('Resumen debe contener exactamente una fila por documento y nueve columnas.')
    for number, item in enumerate(summaries, 2):
        for col, expected in enumerate(item['values'], 1):
            cell = summary.cell(number, col)
            actual = cell.value.date() if isinstance(cell.value, datetime) else cell.value
            if actual != (expected if expected != '' else None):
                raise ValueError('El resumen no conserva sus valores.')
            if isinstance(expected, str) and expected and cell.data_type != 's':
                raise ValueError('Texto de resumen convertido en fórmula o número.')
        link = summary.cell(number, 6).hyperlink
        if link is None or link.target != 'evidencias/' + quote(item['copia'], safe=''):
            raise ValueError('Enlace del resumen incorrecto.')
        if not (path.parent / unquote(link.target)).is_file():
            raise ValueError('Copia del resumen ausente.')
    wb.close()


def verify_originals(envio, manifest):
    for doc in manifest['documentos']:
        path = contained(envio, doc['ruta'])
        if digest(path) != doc['sha256']:
            raise ValueError('La copia de un original ha cambiado: ' + doc['id'])
    for source in manifest['fuentes']:
        path = Path(source['ruta'])
        if not path.is_file() or digest(path) != source['sha256']:
            raise ValueError('Un adjunto original ha cambiado o falta; conserva/restaura el original para continuar.')


def finish(root, envio_path):
    root, _ = config(root)
    allowed = root / 'NO TOCAR/trabajo/envios'
    envio = Path(envio_path).resolve()
    if envio.parent != allowed or not envio.is_dir():
        raise ValueError('Selecciona un envío aislado de este proyecto.')
    with lock(envio):
        state = read_json(envio / 'estado.json')
        if digest(envio / 'manifiesto.json') != state['manifiesto_sha256']:
            raise ValueError('El manifiesto ha sido modificado.')
        manifest = read_json(envio / 'manifiesto.json')
        if manifest['envio_id'] != envio.name or state['envio_id'] != envio.name:
            raise ValueError('Identidad de envío incorrecta.')
        verify_originals(envio, manifest)
        extraction = read_json(envio / 'extraccion.json')
        inputs_hash = [digest(envio / name) for name in ['extraccion.json', 'catalogo.json']]
        destination = root / 'resultados' / simple_name(manifest['cliente']) / envio.name
        archive = envio / 'entrega.zip'
        public_zip = destination / 'entrega.zip'
        receipt = envio / 'entrega.json'
        if receipt.exists():
            report = read_json(receipt)
            if report['datos_sha256'] != inputs_hash:
                raise ValueError('La entrega ya existe con otros datos. Conserva la entrega y crea otro envío.')
            if 'excel' in report:  # Una entrega anterior conserva su formato y sus archivos.
                if not archive.is_file() or digest(archive) != report['zip_sha256']:
                    for relative, expected in report['archivos_sha256'].items():
                        if digest(contained(destination, relative)) != expected:
                            raise ValueError('Una salida anterior ha cambiado.')
                    create_zip(archive, destination)
                    report['zip_sha256'] = digest(archive)
                    write_json(receipt, report)
                return {k: v for k, v in report.items() if k not in {'excel', 'archivos_sha256'}}
            return publish_zip(archive, public_zip, report)
        if destination.exists():
            raise ValueError('Carpeta de resultados existente sin informe de entrega.')
        rows, counts = build_rows(manifest, extraction, read_json(envio / 'catalogo.json'))
        summaries = build_summary_rows(manifest, extraction, rows)
        stage = envio / ('preparacion_' + uuid.uuid4().hex)
        (stage / 'evidencias').mkdir(parents=True)
        for doc in manifest['documentos']:
            target = stage / 'evidencias' / Path(doc['ruta']).name
            shutil.copyfile(contained(envio, doc['ruta']), target)
            if digest(target) != doc['sha256']:
                raise ValueError('Una copia no coincide con su original.')
        excel_name = 'Evidencias_' + slug(manifest['cliente']) + '.xlsx'
        make_excel(stage / excel_name, rows, summaries)
        verify_excel(stage / excel_name, rows, summaries)
        if len(list((stage / 'evidencias').iterdir())) != len(manifest['documentos']):
            raise ValueError('No coincide el número de documentos y copias.')
        verify_originals(envio, manifest)
        create_zip(archive, stage)
        report = {'version': VERSION, 'envio_id': envio.name, 'cliente': manifest['cliente'],
                  'documentos': len(manifest['documentos']), 'copias': len(manifest['documentos']),
                  'filas': len(rows), 'filas_resumen': len(summaries), 'entregables': ['zip'], 'filas_por_documento': counts,
                  'filas_con_revision': sum(bool(r['values'][8]) for r in rows),
                  'zip': str(public_zip),
                  'datos_sha256': inputs_hash, 'zip_sha256': digest(archive)}
        # El informe precede a la publicación. Permite recuperar una interrupción posterior al rename.
        write_json(envio / 'entrega.json', report)
        shutil.rmtree(stage)
        return publish_zip(archive, public_zip, report)


def publish_zip(archive, public_zip, report):
    public_zip.parent.mkdir(parents=True, exist_ok=True)
    if public_zip.exists():
        if digest(public_zip) != report['zip_sha256']:
            raise ValueError('Una salida anterior ha cambiado.')
        return report
    if digest(archive) != report['zip_sha256']:
        raise ValueError('El ZIP técnico ha cambiado.')
    temporary = public_zip.with_name('publicacion_' + uuid.uuid4().hex + '.tmp')
    try:
        shutil.copyfile(archive, temporary)
        os.replace(temporary, public_zip)
    finally:
        temporary.unlink(missing_ok=True)
    return report


def create_zip(archive, folder):
    temporary = archive.with_name('entrega_' + uuid.uuid4().hex + '.zip')
    files = sorted(p for p in folder.rglob('*') if p.is_file())
    with zipfile.ZipFile(temporary, 'x', zipfile.ZIP_DEFLATED) as output:
        for path in files:
            output.write(path, path.relative_to(folder).as_posix())
    with zipfile.ZipFile(temporary) as output:
        if output.testzip() is not None or set(output.namelist()) != {p.relative_to(folder).as_posix() for p in files}:
            raise ValueError('ZIP inválido.')
        for path in files:
            with output.open(path.relative_to(folder).as_posix()) as stream:
                checksum = hashlib.sha256()
                for block in iter(lambda: stream.read(1024 * 1024), b''):
                    checksum.update(block)
                if checksum.hexdigest() != digest(path):
                    raise ValueError('El ZIP no conserva los archivos.')
    os.replace(temporary, archive)


def main():
    parser = argparse.ArgumentParser(description='Analizador de evidencias: envíos independientes.')
    parser.add_argument('accion', choices=['preparar', 'completar'])
    parser.add_argument('--root', required=True)
    parser.add_argument('--cliente')
    parser.add_argument('--archivos', nargs='+')
    parser.add_argument('--envio')
    args = parser.parse_args()
    try:
        if args.accion == 'preparar':
            if not args.cliente or not args.archivos:
                parser.error('preparar requiere --cliente y --archivos')
            result = prepare(args.root, args.cliente, args.archivos)
        else:
            if not args.envio:
                parser.error('completar requiere --envio')
            result = finish(args.root, args.envio)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except (ValueError, OSError, KeyError, TypeError, zipfile.BadZipFile, RuntimeError) as ex:
        parser.exit(1, f'No se ha completado la operación: {ex}\n')


if __name__ == '__main__':
    main()
