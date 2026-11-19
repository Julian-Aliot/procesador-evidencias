import importlib.util
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

PROJECT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location('lectura', PROJECT / 'NO TOCAR/skills/analizador-evidencias/scripts/lectura.py')
reader = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(reader)


class ReaderTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)

    def test_catalog_preserves_text_codes_and_aliases(self):
        path = self.base / 'Proyectos.xlsx'
        wb = Workbook()
        wb.active.append(['codigo', 'aliases'])
        wb.active.append(['0001-PR', 'PR/0001|P-0001'])
        wb.save(path)
        rows = reader.import_catalog(path, 'proyectos')
        self.assertEqual(rows[0]['codigo'], '0001-PR')
        self.assertEqual(rows[0]['aliases'], ['PR/0001', 'P-0001'])
        self.assertIn('fila 2', rows[0]['origen'])

    def test_numeric_code_requires_review_instead_of_losing_zeros(self):
        path = self.base / 'Proyectos.xlsx'
        wb = Workbook()
        wb.active.append(['codigo'])
        wb.active.append([1])
        wb.save(path)
        with self.assertRaisesRegex(ValueError, 'texto literal'):
            reader.import_catalog(path, 'proyectos')

    def test_read_hidden_sheets_and_formula_without_execution(self):
        path = self.base / 'Evidencia.xlsx'
        wb = Workbook()
        wb.active['A1'] = '=1+2'
        wb.create_sheet('Oculta')['B3'] = 'Dato oculto'
        wb['Oculta'].sheet_state = 'hidden'
        wb.save(path)
        fragments = reader.spreadsheet(path)
        self.assertTrue(any(f.get('formula_literal') and f['texto'] == '=1+2' for f in fragments))
        self.assertTrue(any(f['ubicacion'] == "'Oculta'!B3" for f in fragments))


if __name__ == '__main__':
    unittest.main()
