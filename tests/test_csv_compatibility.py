import csv
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from main import process_csv
from openpyxl import load_workbook


class CsvCompatibilityTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.source = Path(self.directory.name) / "input.csv"

    def report(self, text):
        self.source.write_text(text, encoding="utf-8", newline="")
        output = process_csv(str(self.source), self.directory.name)
        workbook = load_workbook(output)
        self.addCleanup(workbook.close)
        return workbook.active

    def test_bom_blank_lines_quoted_commas_and_multiline_text(self):
        sheet = self.report('\ufeff\r\n \r\nid,texto\r\n001,"João, ""Silva""\r\nsegunda linha"\r\n')
        self.assertEqual(sheet["A1"].value, "id")
        self.assertEqual(sheet["A2"].value, "001")
        self.assertEqual(sheet["B2"].value, 'João, "Silva"\r\nsegunda linha')
        self.assertIn("Total rows: 1", sheet["A3"].value)

    def test_duplicate_and_unnamed_headers_do_not_overwrite_columns(self):
        sheet = self.report(',Unnamed: 0,a,a,a.1\nx,y,1,2,3\n')
        self.assertEqual(
            [cell.value for cell in sheet[1]],
            ["Unnamed: 0.1", "Unnamed: 0", "a", "a.2", "a.1"],
        )
        self.assertEqual([cell.value for cell in sheet[2]], ["x", "y", "1", "2", "3"])

    def test_missing_trailing_fields_use_selected_null_treatment(self):
        sheet = self.report('id,a,b\n01,x,y\n02,z\n')
        self.assertEqual(sheet["B3"].value, "z")
        self.assertEqual(sheet["C3"].value, "N/A")

    def test_implicit_index_fields_are_not_exported(self):
        sheet = self.report('id,valor\nignored,001,x\nignored2,002,y\n')
        self.assertEqual([cell.value for cell in sheet[2]], ["001", "x"])
        self.assertEqual([cell.value for cell in sheet[3]], ["002", "y"])

    def test_image_detection_keeps_majority_rule(self):
        sheet = self.report('links,id\na.jpg,001\ntext,002\n')
        self.assertEqual([cell.value for cell in sheet[1]], ["links", "id"])
        sheet = self.report('links,id\na.jpg,001\ntext,002\nb.png?size=1,003\n')
        self.assertEqual([cell.value for cell in sheet[1]], ["id", "links"])
        self.assertEqual(sheet["B4"].value, "b.png?size=1")

    def test_large_fields_are_accepted_and_excel_width_stays_capped(self):
        sheet = self.report('id,texto\n001,' + 'x' * 200000 + '\n')
        self.assertEqual(sheet.column_dimensions["B"].width, 50)
        self.assertEqual(sheet["B2"].value, 'x' * 32767)

    def test_header_only_file_produces_a_report(self):
        sheet = self.report('id,valor\n')
        self.assertIn("Total rows: 0", sheet["A2"].value)
        self.assertEqual(sheet.column_dimensions["B"].width, 9)

    def test_empty_and_malformed_inputs_fail_without_an_output_file(self):
        for text in ('', '\n \n', 'id,a\n1,x\n2,y,z\n', 'id,a\n1,"unclosed\n'):
            with self.subTest(text=text):
                self.source.write_text(text, encoding="utf-8")
                with self.assertRaises((ValueError, csv.Error)):
                    process_csv(str(self.source), self.directory.name)
                self.assertFalse(self.source.with_suffix('.xlsx').exists())

    def test_processing_works_without_pandas_numpy_or_pillow(self):
        self.source.write_text('id,valor\n001,teste\n', encoding="utf-8")
        subprocess.run(
            [sys.executable, "-c", '''
import sys
from pathlib import Path

class BlockOptionalLibraries:
    def find_spec(self, fullname, path=None, target=None):
        if fullname.split('.')[0] in {'pandas', 'numpy', 'PIL'}:
            raise ImportError(fullname)

sys.meta_path.insert(0, BlockOptionalLibraries())
sys.path.insert(0, 'src')
from main import process_csv
from openpyxl import load_workbook
source = Path(sys.argv[1])
output = process_csv(str(source), str(source.parent))
workbook = load_workbook(output)
assert workbook.active['A2'].value == '001'
assert workbook.active['B2'].value == 'teste'
workbook.close()
''', str(self.source)],
            cwd=ROOT,
            check=True,
        )


if __name__ == "__main__":
    unittest.main()
