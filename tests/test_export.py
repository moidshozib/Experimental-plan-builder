import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from main import export_plan, import_plan


class ExportTests(unittest.TestCase):
    def test_export_and_import_round_trip(self):
        columns = ["Run", "Sample", "Temperature"]
        rows = [["1", "A", "25 C"], ["2", "B", "30 C"]]
        details = {"plan_name": "Demo", "researcher": "Tester"}
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "plan.xlsx"
            export_plan(path, columns, rows, details)
            workbook = load_workbook(path, data_only=True)
            self.assertEqual(workbook.sheetnames, ["Experimental Plan", "Plan Details"])
            self.assertEqual(workbook["Experimental Plan"].max_row, 3)
            self.assertEqual(workbook["Experimental Plan"]["B2"].value, "A")
            workbook.close()
            imported_columns, imported_rows = import_plan(path)
            self.assertEqual(imported_columns, columns)
            self.assertEqual(imported_rows, rows)


if __name__ == "__main__":
    unittest.main()
