import importlib
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = REPO_ROOT / "skills" / "csv-doctor" / "scripts"
LOADER_PATH = SCRIPTS_DIR / "loader.py"
DETECTOR_PATH = SCRIPTS_DIR / "column_detector.py"
HEAL_PATH = SCRIPTS_DIR / "heal.py"

sys.path.insert(0, str(SCRIPTS_DIR))


def load_module(path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


class DataShapeEdgeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.loader = load_module(LOADER_PATH, "sheet_doctor_loader_shape_tests")
        cls.detector = load_module(DETECTOR_PATH, "sheet_doctor_detector_shape_tests")
        cls.heal = load_module(HEAL_PATH, "sheet_doctor_heal_shape_tests")
        cls.normalization = importlib.import_module("heal_modules.normalization")

    def test_one_column_file_loads_and_profiles(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "single.csv"
            path.write_text("notes\nhello world\n", encoding="utf-8")
            loaded = self.loader.load_file(path)

        self.assertEqual(loaded["dataframe"].shape, (1, 1))
        analysis = self.detector.analyse_dataframe(loaded["dataframe"])
        self.assertEqual(analysis["summary"]["total_columns"], 1)
        self.assertEqual(analysis["columns"]["notes"]["detected_type"], "free text")

    def test_one_row_file_heals_cleanly(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "one_row.csv"
            path.write_text(
                "Employee Name,Department,Date,Amount,Currency,Category,Status,Notes\n"
                "Ada Lovelace,Finance,2023-01-15,$10.00,USD,Travel,Approved,Taxi\n",
                encoding="utf-8",
            )
            result = self.heal.execute_healing(path)

        self.assertEqual(result["mode"], "schema-specific")
        self.assertEqual(len(result["clean_data"]), 1)
        self.assertEqual(len(result["quarantine"]), 0)

    def test_duplicate_headers_are_deduped_in_generic_mode(self):
        rows = [
            ["Name", "Name", "Amount"],
            ["Ada", "Ada Lovelace", "10"],
        ]
        clean, quarantine, _, headers, _ = self.heal.process_generic(rows, ",")

        self.assertEqual(headers, ["Name", "Name_2", "Amount"])
        self.assertEqual(len(clean), 1)
        self.assertEqual(len(quarantine), 0)

    def test_500_plus_columns_load(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "wide.csv"
            headers = [f"col_{i}" for i in range(501)]
            values = [str(i) for i in range(501)]
            path.write_text(",".join(headers) + "\n" + ",".join(values) + "\n", encoding="utf-8")
            loaded = self.loader.load_file(path)

        self.assertEqual(loaded["dataframe"].shape, (1, 501))

    def test_very_long_text_and_unicode_values_survive_analysis(self):
        long_text = "x" * 12000
        emoji = "Ready ✅"
        rtl = "مرحبا بالعالم"
        cjk = "こんにちは世界"
        df = pd.DataFrame({
            "notes": [long_text, long_text[::-1]],
            "emoji": [emoji, emoji],
            "rtl": [rtl, rtl],
            "cjk": [cjk, cjk],
        })

        analysis = self.detector.analyse_dataframe(df)["columns"]

        self.assertEqual(analysis["notes"]["detected_type"], "free text")
        self.assertTrue(any(len(value) >= 12000 for value in analysis["notes"]["sample_values"]))
        self.assertIn(emoji, analysis["emoji"]["sample_values"])
        self.assertIn(rtl, analysis["rtl"]["sample_values"])
        self.assertIn(cjk, analysis["cjk"]["sample_values"])

    def test_numbers_stored_as_text_are_profiled_as_numbers(self):
        df = pd.DataFrame({"Column1": ["10", "20", "30"]})
        column = self.detector.analyse_dataframe(df)["columns"]["Column1"]

        self.assertEqual(column["detected_type"], "plain number")
        self.assertEqual(column["min_value"], 10.0)
        self.assertEqual(column["max_value"], 30.0)

    def test_dates_stored_as_numbers_normalise_from_excel_serial(self):
        value, changed, _ = self.normalization.normalise_date("44910")

        self.assertTrue(changed)
        self.assertRegex(value, r"^\d{4}-\d{2}-\d{2}$")

    def test_negative_amounts_normalise_from_accounting_format(self):
        value, changed, _ = self.normalization.normalise_amount("(500)")

        self.assertTrue(changed)
        self.assertEqual(value, "-500.00")

    def test_all_null_and_all_identical_columns_are_flagged(self):
        df = pd.DataFrame(
            {
                "empty_col": [None, None, None],
                "same_col": ["Approved", "Approved", "Approved"],
            }
        )
        analysis = self.detector.analyse_dataframe(df)["columns"]

        self.assertEqual(analysis["empty_col"]["null_count"], 3)
        self.assertEqual(analysis["empty_col"]["detected_type"], "unknown")
        self.assertIn("Values suspiciously all the same", analysis["same_col"]["suspected_issues"])


if __name__ == "__main__":
    unittest.main()
