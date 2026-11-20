import importlib.util
import io
import json
from pathlib import Path
import stat
import subprocess
import sys
import tempfile
import unittest
import zipfile

from openpyxl import Workbook, load_workbook

PROJECT = Path(__file__).resolve().parents[2]
SCRIPT = PROJECT / 'NO TOCAR/skills/analizador-evidencias/scripts/envios.py'
SPEC = importlib.util.spec_from_file_location('envios', SCRIPT)
engine = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(engine)


class EnvioTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.root = self.base / 'proyecto'
        cfg = self.root / 'NO TOCAR/config'
        cfg.mkdir(parents=True)
        engine.write_json(cfg / 'config.json', engine.read_json(PROJECT / 'NO TOCAR/config/config.json'))
        self.source = self.base / 'evidencia.txt'
        self.source.write_text('Ana García\nPR-001\nTarea A\n19/11/2026', encoding='utf-8')

    def prepare(self):
        return Path(engine.prepare(self.root, 'Cliente', [self.source])['envio'])

    def workbook(self, report):
        with zipfile.ZipFile(report['zip']) as archive:
            [book] = [name for name in archive.namelist() if '/' not in name and name.endswith('.xlsx')]
            return load_workbook(io.BytesIO(archive.read(book)))

    def fill(self, envio, twice=False):
        data = engine.read_json(envio / 'extraccion.json')
        def evidence(value):
            return {'valor': value, 'ubicacion': 'Línea 1', 'cita': value}
        activity = engine.empty_activity('')
        activity['motivos_revision'] = []
        activity['personas'] = [dict(evidence('Ana García'), rol='participante')]
        activity['codigos_proyecto'] = [evidence('PR-001')]
        activity['fecha'] = evidence('2026-11-19')
        activity['tipo_fecha'] = 'realizacion'
        activity['titulo_tarea'] = evidence('=Tarea literal')
        activity['resumen'] = evidence('Actividad revisada')
        data['documentos'][0]['resumen_documento']['descripcion_tarea'] = evidence('Tarea de revisión del modelo')
        data['documentos'][0]['lectura_completa'] = True
        data['documentos'][0]['actividades'] = [activity, activity] if twice else [activity]
        engine.write_json(envio / 'extraccion.json', data)
        catalog = {'personal': [{'nombre': 'Ana García', 'aliases': [], 'origen': 'Personal, A2'}],
                   'proyectos': [{'codigo': 'PR-001', 'aliases': [], 'origen': 'Proyectos, A2'}]}
        engine.write_json(envio / 'catalogo.json', catalog)

    def test_end_to_end_multiple_activities_and_retry(self):
        envio = self.prepare()
        self.fill(envio, twice=True)
        report = engine.finish(self.root, envio)
        self.assertEqual((report['documentos'], report['copias'], report['filas']), (1, 1, 2))
        wb = self.workbook(report)
        self.addCleanup(wb.close)
        self.assertEqual(wb.sheetnames, ['Resumen', 'DETALLE'])
        ws = wb['DETALLE']
        self.assertEqual([c.value for c in ws[1]], engine.HEADERS)
        self.assertEqual(ws['F2'].value, '=Tarea literal')
        self.assertEqual(ws['F2'].data_type, 's')
        self.assertTrue(ws['E2'].is_date)
        self.assertEqual(ws['B2'].value, 'Ana García')
        self.assertEqual(ws['D2'].value, 'PR-001')
        self.assertEqual(ws['I2'].hyperlink.target, ws['I3'].hyperlink.target)
        with zipfile.ZipFile(report['zip']) as z:
            self.assertEqual(len(z.namelist()), 2)
            self.assertIsNone(z.testzip())
        self.assertEqual(engine.finish(self.root, envio), report)
        Path(report['zip']).unlink()
        repaired = engine.finish(self.root, envio)
        self.assertTrue(Path(repaired['zip']).is_file())

    def test_seven_documents_summary_and_person_task_trace(self):
        sources = []
        for n in range(7):
            source = self.base / f'image{n}.png'
            source.write_bytes(b'identical image bytes')
            sources.append(source)
        envio = Path(engine.prepare(self.root, 'Cliente', sources)['envio'])
        self.fill(envio, twice=True)
        data = engine.read_json(envio / 'extraccion.json')
        activity = data['documentos'][0]['actividades'][0]
        activity['personas'].append({'valor': 'Luis Pérez', 'rol': 'autor',
                                    'ubicacion': 'Firma', 'cita': 'Luis Pérez'})
        for document in data['documentos']:
            document['resumen_documento'] = data['documentos'][0]['resumen_documento']
            document['actividades'] = [activity, activity]
            document['lectura_completa'] = True
        engine.write_json(envio / 'extraccion.json', data)
        report = engine.finish(self.root, envio)
        self.assertEqual(report['filas_resumen'], 7)
        self.assertEqual(report['filas'], 28)
        self.assertEqual(report['entregables'], ['zip'])
        wb = self.workbook(report)
        self.addCleanup(wb.close)
        summary, detail = wb['Resumen'], wb['DETALLE']
        self.assertNotIn('excel', report)
        self.assertEqual(list(Path(report['zip']).parent.iterdir()), [Path(report['zip'])])
        self.assertEqual(summary['H2'].value, 'Ana García')
        self.assertEqual(summary['I2'].value, 'Tarea de revisión del modelo')
        self.assertEqual(summary.max_row, 8)
        self.assertEqual(summary.max_column, 9)
        self.assertIsNone(summary['A2'].value)
        self.assertEqual(summary['B2'].value, 'Ana García\nLuis Pérez')
        self.assertEqual(summary['D2'].value, 'PR-001')
        self.assertEqual(detail['B2'].value, 'Ana García')
        self.assertIsNone(detail['B3'].value)
        self.assertIn('Personal, A2', detail['I2'].comment.text)
        self.assertEqual(summary['F2'].data_type, 's')
        self.assertEqual(summary['F2'].value, '=Tarea literal')

    def test_summary_metadata_missing_catalog_and_multiple_dates(self):
        envio = self.prepare()
        self.fill(envio, twice=True)
        data = engine.read_json(envio / 'extraccion.json')
        doc = data['documentos'][0]
        doc['actividades'][1] = json.loads(json.dumps(doc['actividades'][1]))
        doc['actividades'][1]['fecha']['valor'] = '2026-11-20'
        doc['resumen_documento'] = {
            key: {'valor': value, 'ubicacion': 'Documento', 'cita': value}
            for key, value in [('nombre_principal', 'Ana García'),
                               ('titulo', 'Título global'), ('resumen', 'Resumen global'), ('descripcion_tarea', 'Tarea de composición del plano 21')]}
        engine.write_json(envio / 'extraccion.json', data)
        engine.write_json(envio / 'catalogo.json', {'personal': None, 'proyectos': None})
        report = engine.finish(self.root, envio)
        wb = self.workbook(report)
        self.addCleanup(wb.close)
        self.assertIsNone(wb['Resumen']['E2'].value)
        self.assertIsNone(wb['Resumen']['D2'].value)
        self.assertEqual(wb['Resumen']['F2'].value, 'Título global')
        self.assertIn('Varias fechas', wb['Resumen']['F2'].comment.text)
        self.assertIsNone(wb['DETALLE']['B2'].value)

    def test_same_client_same_files_produces_separate_shipments(self):
        first, second = self.prepare(), self.prepare()
        self.assertNotEqual(first, second)
        for envio in [first, second]:
            self.fill(envio)
        one = engine.finish(self.root, first)
        original_hash = engine.digest(one['zip'])
        two = engine.finish(self.root, second)
        self.assertNotEqual(one['zip'], two['zip'])
        self.assertEqual(engine.digest(one['zip']), original_hash)
        self.assertEqual(two['documentos'], 1)

    def test_catalog_matching_full_names_codes_and_ambiguous_aliases(self):
        indexes = engine.catalog_indexes({'personal': [
            {'nombre': 'Ana García', 'aliases': ['A. García'], 'origen': 'A2'},
            {'nombre': 'Alba García', 'aliases': ['A. García'], 'origen': 'A3'}], 'proyectos': None})
        reasons = []
        self.assertEqual(engine.matches([' ANA   GARCÍA '], indexes['personal'], 'personal', reasons), 'Ana García')
        self.assertEqual(engine.matches(['Ana'], indexes['personal'], 'personal', reasons), '')
        self.assertEqual(engine.matches(['A. García'], indexes['personal'], 'personal', reasons), '')
        self.assertTrue(any('ambigua' in v for v in reasons))
        self.assertNotEqual(engine.normalized('PR-001'), engine.normalized('PR-1'))
        self.assertNotEqual(engine.normalized('PR-001'), engine.normalized('PR001'))

    def test_names_and_decorated_codes_in_delivered_zip(self):
        envio = self.prepare()
        self.fill(envio)
        data = engine.read_json(envio / 'extraccion.json')
        activity = data['documentos'][0]['actividades'][0]
        def run(name, code, people):
            activity['personas'][0].update(valor=name, cita='Responsable: ' + name)
            activity['codigos_proyecto'][0].update(valor=code, cita=code)
            # Cada prueba tiene su propio envío.
            current = self.prepare()
            self.fill(current)
            data['envio_id'] = current.name
            engine.write_json(current / 'extraccion.json', data)
            cat = engine.read_json(current / 'catalogo.json')
            cat['personal'] = [{'nombre': n, 'origen': 'Personal.xlsx, A2', 'aliases': []} for n in people]
            engine.write_json(current / 'catalogo.json', cat)
            wb = self.workbook(engine.finish(self.root, current))
            self.addCleanup(wb.close)
            return [wb['Resumen'][c].value for c in ['H2', 'D2']] + [wb['DETALLE']['B2'].value]
        self.assertEqual(run('A. García', 'plano_PR-001_004', ['Ana García']),
                         ['Ana García', 'PR-001', 'Ana García'])
        self.assertEqual(run('Fran', 'PR-0012', ['Francisco López']), ['Francisco López', None, 'Francisco López'])
        self.assertEqual(run('Ana', 'PR-001', ['Ana García', 'Ana Pérez']), [None, 'PR-001', None])
        self.assertTrue(engine.embedded_code('seq_atlas042_comp', 'atlas'))
        self.assertFalse(engine.embedded_code('atlasmar042', 'atlas'))

    def test_cli_imports_reference_workbook_and_publishes_literal_and_canonical_columns(self):
        path = self.base / 'Listados.xlsx'
        wb = Workbook()
        wb.active.title = 'Personal'
        wb.active.append(['Equipo 2026'])
        wb.active.append(['Nombre', 'Apellidos'])
        wb.active.append(['Ana', 'García López'])
        projects = wb.create_sheet('Proyectos')
        projects.append(['Código de proyecto'])
        projects.append(['PR-001'])
        wb.save(path)
        for people, codes in [(True, True), (False, False), (True, False), (False, True)]:
            with self.subTest(people=people, codes=codes):
                cmd = [sys.executable, str(SCRIPT), 'preparar', '--root', str(self.root),
                       '--cliente', 'Cliente', '--archivos', str(self.source)]
                if people:
                    cmd += ['--personal', str(path)]
                if codes:
                    cmd += ['--proyectos', str(path)]
                envio = Path(json.loads(subprocess.check_output(cmd))['envio'])
                extracted = engine.read_json(envio / 'extraccion.json')
                doc = extracted['documentos'][0]
                activity = doc['actividades'][0]
                def observed(value):
                    return {'valor': value, 'ubicacion': 'Captura, panel de tarea', 'cita': value}
                activity['personas'] = [dict(observed('A. García'), rol='autor')]
                activity['codigos_proyecto'] = [observed('shot_PR-001_042')]
                activity['descripcion_tarea'] = observed('Compositing del plano 042')
                activity['motivos_revision'] = []
                doc['lectura_completa'] = True
                engine.write_json(envio / 'extraccion.json', extracted)
                report = json.loads(subprocess.check_output(
                    [sys.executable, str(SCRIPT), 'completar', '--root', str(self.root), '--envio', str(envio)]))
                book = self.workbook(report)
                self.addCleanup(book.close)
                detail = book['DETALLE']
                self.assertEqual([c.value for c in detail[1][:4]],
                                 ['NOMBRE DETECTADO', 'NOMBRE BONIFICADO', 'CÓDIGO DETECTADO', 'CÓDIGO BONIFICADO'])
                self.assertEqual([c.value for c in detail[2][:4]],
                                 ['A. García', 'Ana García López' if people else None,
                                  'shot_PR-001_042', 'PR-001' if codes else None])
                self.assertEqual(detail['K2'].value, 'Compositing del plano 042')
                self.assertEqual(book['Resumen']['I2'].value, detail['K2'].value)
                self.assertEqual(book['Resumen']['H2'].value, detail['B2'].value)
                self.assertEqual(book['Resumen']['D2'].value, detail['D2'].value)

    def test_no_catalog_and_unreadable_file_retained(self):
        envio = self.prepare()
        report = engine.finish(self.root, envio)
        self.assertEqual(report['filas'], 1)
        self.assertEqual(report['filas_con_revision'], 1)
        wb = self.workbook(report)
        self.addCleanup(wb.close)
        self.assertIn('Listado de personal bonificado no aportado', wb['DETALLE']['J2'].value)
        self.assertIn('Lectura incompleta', wb['DETALLE']['J2'].value)
        self.assertIsNone(wb['DETALLE']['B2'].value)

    def test_tampered_original_or_extraction_is_blocked(self):
        envio = self.prepare()
        self.fill(envio)
        self.source.write_text('modificado', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'original ha cambiado'):
            engine.finish(self.root, envio)

    def test_output_tampering_and_changed_completed_data_blocked(self):
        envio = self.prepare()
        self.fill(envio)
        report = engine.finish(self.root, envio)
        original_extraction = (envio / 'extraccion.json').read_bytes()
        data = engine.read_json(envio / 'extraccion.json')
        data['documentos'][0]['incidencias_documento'] = ['Cambio']
        engine.write_json(envio / 'extraccion.json', data)
        with self.assertRaisesRegex(ValueError, 'otros datos'):
            engine.finish(self.root, envio)
        (envio / 'extraccion.json').write_bytes(original_extraction)
        Path(report['zip']).write_bytes(b'alterado')
        with self.assertRaisesRegex(ValueError, 'salida anterior ha cambiado'):
            engine.finish(self.root, envio)

    def test_secure_zip_paths_symlinks_and_duplicate_names(self):
        limits = engine.read_json(PROJECT / 'NO TOCAR/config/config.json')['limites']
        for names in [['../escape.txt'], ['/absoluta'], ['C:/ruta.txt'], ['a\\b.txt'], ['A.txt', 'a.txt']]:
            memory = io.BytesIO()
            with zipfile.ZipFile(memory, 'w') as archive:
                for name in names:
                    archive.writestr(name, 'dato')
            with zipfile.ZipFile(memory) as archive:
                with self.assertRaises(ValueError):
                    engine.safe_members(archive, limits)
        memory = io.BytesIO()
        with zipfile.ZipFile(memory, 'w') as archive:
            info = zipfile.ZipInfo('enlace')
            info.external_attr = (stat.S_IFLNK | 0o777) << 16
            archive.writestr(info, '/etc/passwd')
        with zipfile.ZipFile(memory) as archive:
            with self.assertRaises(ValueError):
                engine.safe_members(archive, limits)

    def test_valid_zip_keeps_identical_documents_and_unknown_format(self):
        archive = self.base / 'documentos.zip'
        with zipfile.ZipFile(archive, 'w') as z:
            z.writestr('carpeta/uno.txt', 'idéntico')
            z.writestr('carpeta/dos.txt', 'idéntico')
            z.writestr('corrupto.unknown', b'\x00\x01')
        envio = Path(engine.prepare(self.root, 'Cliente', [archive])['envio'])
        report = engine.finish(self.root, envio)
        self.assertEqual((report['documentos'], report['copias'], report['filas']), (3, 3, 3))
        wb = self.workbook(report)
        self.addCleanup(wb.close)
        notes = '\n'.join(str(wb['DETALLE'].cell(r, 10).value) for r in range(2, 5))
        self.assertIn('Archivo idéntico', notes)
        self.assertIn('Formato sin lector', notes)


if __name__ == '__main__':
    unittest.main()
