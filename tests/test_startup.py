import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import app
from main import process_csv
from openpyxl import load_workbook


class StartupTests(unittest.TestCase):
    def test_gui_import_does_not_load_spreadsheet_engine(self):
        subprocess.run(
            [sys.executable, "-c", (
                "import sys; sys.path.insert(0, 'src'); import app; "
                "assert not {'main', 'pandas', 'numpy', 'openpyxl'} & sys.modules.keys()"
            )],
            cwd=ROOT,
            check=True,
        )

    def test_worker_generates_report_with_delayed_import(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.csv"
            source.write_text('id,nome\n1,João\n', encoding="utf-8")
            button, status = Mock(), Mock()
            with patch.object(app.messagebox, "showinfo") as success, patch.object(
                app.messagebox, "showerror"
            ) as error:
                app.run_processing(str(source), directory, False, "00AD32", "na", button, status)
            error.assert_not_called()
            success.assert_called_once()
            workbook = load_workbook(Path(directory) / "input.xlsx")
            self.assertEqual(workbook.active["B2"].value, "João")
            workbook.close()
            self.assertEqual(button.config.call_args.kwargs["state"], app.tk.NORMAL)

    def test_report_options_and_image_order_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.csv"
            source.write_text(
                'foto,id,nome,valor\n' + 'https://example.com/a.jpg,001,"João, Silva",\n' * 12,
                encoding="utf-8",
            )
            for treatment, expected in (("na", "N/A"), ("vazio", None), ("zero", 0)):
                with self.subTest(treatment=treatment):
                    output = process_csv(str(source), directory, True, "123456", treatment)
                    workbook = load_workbook(output)
                    sheet = workbook.active
                    self.assertEqual([cell.value for cell in sheet[1]], ["id", "nome", "valor", "foto"])
                    self.assertEqual(sheet["A2"].value, "001")
                    self.assertEqual(sheet["B2"].value, "João, Silva")
                    self.assertEqual(sheet["C2"].value, expected)
                    self.assertEqual(sheet["A1"].fill.fgColor.rgb, "00123456")
                    self.assertTrue(sheet.protection.sheet)
                    self.assertEqual(sheet.freeze_panes, "A2")
                    self.assertIn("Total rows: 12", sheet["A14"].value)
                    self.assertIn("valor (12)", sheet["A15"].value)
                    workbook.close()


if __name__ == "__main__":
    unittest.main()
