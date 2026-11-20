import importlib.util
import io
import json
from pathlib import Path
import stat
import tempfile
import unittest
import zipfile

from openpyxl import load_workbook

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
        wb = load_workbook(report['excel'])
        self.addCleanup(wb.close)
        self.assertEqual(wb.sheetnames, ['Resumen', 'Evidencias'])
        ws = wb['Evidencias']
        self.assertEqual([c.value for c in ws[1]], engine.HEADERS)
        self.assertEqual(ws['E2'].value, '=Tarea literal')
        self.assertEqual(ws['E2'].data_type, 's')
        self.assertTrue(ws['D2'].is_date)
        self.assertEqual(ws['J2'].value, 'Ana García')
        self.assertEqual(ws['K2'].value, 'PR-001')
        self.assertEqual(ws['H2'].hyperlink.target, ws['H3'].hyperlink.target)
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
            document['actividades'] = [activity, activity]
            document['lectura_completa'] = True
        engine.write_json(envio / 'extraccion.json', data)
        report = engine.finish(self.root, envio)
        self.assertEqual(report['filas_resumen'], 7)
        self.assertEqual(report['filas'], 28)
        self.assertEqual(report['entregables'], ['zip'])
        with zipfile.ZipFile(report['zip']) as archive:
            wb = load_workbook(io.BytesIO(archive.read(Path(report['excel']).name)))
        self.addCleanup(wb.close)
        summary, detail = wb['Resumen'], wb['Evidencias']
        self.assertEqual(summary.max_row, 8)
        self.assertEqual(summary.max_column, 7)
        self.assertIsNone(summary['A2'].value)
        self.assertEqual(summary['B2'].value, 'Ana García\nLuis Pérez')
        self.assertEqual(summary['D2'].value, 'PR-001')
        self.assertEqual(detail['J2'].value, 'Ana García')
        self.assertIsNone(detail['J3'].value)
        self.assertIn('Personal, A2', detail['H2'].comment.text)
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
                               ('titulo', 'Título global'), ('resumen', 'Resumen global')]}
        engine.write_json(envio / 'extraccion.json', data)
        engine.write_json(envio / 'catalogo.json', {'personal': None, 'proyectos': None})
        report = engine.finish(self.root, envio)
        wb = load_workbook(report['excel'])
        self.addCleanup(wb.close)
        self.assertIsNone(wb['Resumen']['E2'].value)
        self.assertIsNone(wb['Resumen']['D2'].value)
        self.assertEqual(wb['Resumen']['F2'].value, 'Título global')
        self.assertIn('Varias fechas', wb['Resumen']['F2'].comment.text)
        self.assertIsNone(wb['Evidencias']['J2'].value)

    def test_same_client_same_files_produces_separate_shipments(self):
        first, second = self.prepare(), self.prepare()
        self.assertNotEqual(first, second)
        for envio in [first, second]:
            self.fill(envio)
        one = engine.finish(self.root, first)
        original_hash = engine.digest(one['excel'])
        two = engine.finish(self.root, second)
        self.assertNotEqual(one['excel'], two['excel'])
        self.assertEqual(engine.digest(one['excel']), original_hash)
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

    def test_no_catalog_and_unreadable_file_retained(self):
        envio = self.prepare()
        report = engine.finish(self.root, envio)
        self.assertEqual(report['filas'], 1)
        self.assertEqual(report['filas_con_revision'], 1)
        wb = load_workbook(report['excel'])
        self.addCleanup(wb.close)
        self.assertIn('Listado de personal bonificado no aportado', wb['Evidencias']['I2'].value)
        self.assertIn('Lectura incompleta', wb['Evidencias']['I2'].value)
        self.assertIsNone(wb['Evidencias']['J2'].value)

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
        Path(report['excel']).write_bytes(b'alterado')
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
        wb = load_workbook(report['excel'])
        self.addCleanup(wb.close)
        notes = '\n'.join(str(wb['Evidencias'].cell(r, 9).value) for r in range(2, 5))
        self.assertIn('Archivo idéntico', notes)
        self.assertIn('Formato sin lector', notes)


if __name__ == '__main__':
    unittest.main()
