import importlib.util
import io
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

from openpyxl import Workbook


_XLRD_AVAILABLE = importlib.util.find_spec("xlrd") is not None
_ODFPY_AVAILABLE = importlib.util.find_spec("odf") is not None


REPO_ROOT = Path(__file__).resolve().parents[1]
LOADER_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "loader.py"
FIXTURE_DIR = REPO_ROOT / "tests" / "fixtures" / "loader"


def load_loader_module():
    spec = importlib.util.spec_from_file_location("sheet_doctor_loader", LOADER_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class LoaderLocalBehaviorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.loader = load_loader_module()

    def test_plain_txt_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "notes.txt"
            path.write_text("This is a text file.\n", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "does not appear to contain delimited/tabular data"):
                self.loader.load_file(path)

    def test_empty_text_file_raises_clear_error(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "empty.csv"
            path.write_text("", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "File is empty"):
                self.loader.load_file(path)

    def test_missing_xlrd_raises_clear_importerror(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "legacy.xls"
            path.write_bytes(b"not-a-real-xls")

            original_import = __import__

            def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
                if name == "xlrd":
                    raise ImportError("simulated missing xlrd")
                return original_import(name, globals, locals, fromlist, level)

            with mock.patch("builtins.__import__", side_effect=fake_import):
                with self.assertRaisesRegex(ImportError, r"\.xls files require xlrd"):
                    self.loader.load_file(path)

    def test_missing_odfpy_raises_clear_importerror(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sheet.ods"
            path.write_bytes(b"not-a-real-ods")

            original_import = __import__

            def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
                if name == "odf":
                    raise ImportError("simulated missing odf")
                return original_import(name, globals, locals, fromlist, level)

            with mock.patch("builtins.__import__", side_effect=fake_import):
                with self.assertRaisesRegex(ImportError, r"\.ods files require odfpy"):
                    self.loader.load_file(path)

    def test_large_text_inputs_emit_degraded_mode_warning(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "large.csv"
            path.write_text("name,value\nAda,10\nGrace,11\n", encoding="utf-8")

            with mock.patch.object(self.loader, "LARGE_FILE_WARNING_BYTES", 1), \
                 mock.patch.object(self.loader, "LARGE_FILE_DEGRADED_BYTES", 1), \
                 mock.patch.object(self.loader, "LARGE_FILE_HARD_LIMIT_BYTES", 10_000), \
                 mock.patch.object(self.loader, "LARGE_ROW_WARNING_COUNT", 1_000), \
                 mock.patch.object(self.loader, "LARGE_ROW_DEGRADED_COUNT", 2_000), \
                 mock.patch.object(self.loader, "LARGE_ROW_HARD_LIMIT_COUNT", 10_000):
                result = self.loader.load_file(path)

            self.assertTrue(result["degraded_mode"]["active"])
            self.assertTrue(any("Degraded mode active" in warning for warning in result["warnings"]))
            self.assertTrue(any("Large file size" in reason for reason in result["degraded_mode"]["reasons"]))

    def test_hard_limit_rejects_oversized_input(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "tiny.csv"
            path.write_text("a,b\n1,2\n", encoding="utf-8")

            with mock.patch.object(self.loader, "LARGE_FILE_HARD_LIMIT_BYTES", 1):
                with self.assertRaisesRegex(ValueError, "too large for safe in-memory processing"):
                    self.loader.load_file(path)

    def test_multisheet_xlsx_requires_explicit_selection_noninteractive(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "multi.xlsx"
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "Visible"
            ws1.append(["name", "score"])
            ws1.append(["Ada", 10])
            ws2 = wb.create_sheet("Backup")
            ws2.append(["name", "score"])
            ws2.append(["Grace", 11])
            wb.save(path)

            fake_stdin = mock.Mock()
            fake_stdin.isatty.return_value = False

            with mock.patch.object(self.loader.sys, "stdin", fake_stdin):
                with self.assertRaisesRegex(ValueError, "Multiple sheets found"):
                    self.loader.load_file(path)

    def test_multisheet_xlsx_loads_when_sheet_name_is_explicit(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "multi.xlsx"
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "Visible"
            ws1.append(["name", "score"])
            ws1.append(["Ada", 10])
            ws2 = wb.create_sheet("Backup")
            ws2.append(["name", "score"])
            ws2.append(["Grace", 11])
            wb.save(path)

            fake_stdin = mock.Mock()
            fake_stdin.isatty.return_value = False

            with mock.patch.object(self.loader.sys, "stdin", fake_stdin):
                result = self.loader.load_file(path, sheet_name="Backup")

            self.assertEqual(result["sheet_name"], "Backup")
            self.assertEqual(result["sheet_names"], ["Visible", "Backup"])
            self.assertEqual(result["dataframe"].shape, (1, 2))
            self.assertEqual(result["dataframe"].iloc[0].to_dict(), {"name": "Grace", "score": "11"})

    def test_multisheet_xlsx_can_consolidate_explicitly(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "multi.xlsx"
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "Q1"
            ws1.append(["name", "score"])
            ws1.append(["Ada", 10])
            ws2 = wb.create_sheet("Q2")
            ws2.append(["name", "score"])
            ws2.append(["Grace", 11])
            wb.save(path)

            fake_stdin = mock.Mock()
            fake_stdin.isatty.return_value = False

            with mock.patch.object(self.loader.sys, "stdin", fake_stdin):
                result = self.loader.load_file(path, consolidate_sheets=True)

            self.assertEqual(result["sheet_name"], "[all 2 sheets]")
            self.assertEqual(result["sheet_names"], ["Q1", "Q2"])
            self.assertEqual(result["dataframe"].shape, (2, 2))
            self.assertIn("Consolidated 2 sheets into one table.", result["warnings"])


class LoaderFixtureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.loader = load_loader_module()
        cls.csv_path = FIXTURE_DIR / "semicolon_people.csv"
        cls.tsv_path = FIXTURE_DIR / "sample.tsv"
        cls.xlsx_path = FIXTURE_DIR / "multisheet.xlsx"
        cls.xlsm_path = FIXTURE_DIR / "multisheet.xlsm"
        cls.ods_path = FIXTURE_DIR / "multisheet.ods"
        cls.json_path = FIXTURE_DIR / "records.json"
        cls.json_nested_path = FIXTURE_DIR / "nested_records.json"
        cls.jsonl_path = FIXTURE_DIR / "records.jsonl"
        cls.corrupt_xls_path = FIXTURE_DIR / "corrupt.xls"
        cls.corrupt_xlsx_path = FIXTURE_DIR / "corrupt.xlsx"
        cls.encrypted_xlsx_path = FIXTURE_DIR / "encrypted.xlsx"

    def test_fixture_csv_loads_with_semicolon_detection(self):
        result = self.loader.load_file(self.csv_path)

        self.assertEqual(result["detected_format"], "csv")
        self.assertEqual(result["delimiter"], ";")
        self.assertEqual(result["dataframe"].shape, (2, 3))

    def test_fixture_tsv_loads(self):
        result = self.loader.load_file(self.tsv_path)

        self.assertEqual(result["detected_format"], "tsv")
        self.assertEqual(result["delimiter"], "\t")
        self.assertEqual(result["dataframe"].shape, (2, 3))

    def test_fixture_xlsx_requires_sheet_name_noninteractive(self):
        with self.assertRaisesRegex(ValueError, "Available sheets"):
            self.loader.load_file(self.xlsx_path)

    def test_fixture_xlsx_loads_selected_sheet(self):
        result = self.loader.load_file(self.xlsx_path, sheet_name="Visible")

        self.assertEqual(result["sheet_name"], "Visible")
        self.assertEqual(result["sheet_names"], ["Visible", "Hidden", "VeryHidden"])
        self.assertEqual(result["dataframe"].shape, (4, 2))

    def test_fixture_xlsm_loads_selected_sheet(self):
        result = self.loader.load_file(self.xlsm_path, sheet_name="Visible")

        self.assertEqual(result["detected_format"], "xlsm")
        self.assertEqual(result["sheet_name"], "Visible")
        self.assertEqual(result["dataframe"].shape, (4, 2))

    def test_fixture_ods_loads_selected_sheet(self):
        if not _ODFPY_AVAILABLE:
            self.skipTest("odfpy not installed — run: pip install odfpy")

        result = self.loader.load_file(self.ods_path, sheet_name="Visible")

        self.assertEqual(result["detected_format"], "ods")
        self.assertEqual(result["sheet_name"], "Visible")
        self.assertEqual(result["dataframe"].shape, (2, 2))

    def test_fixture_json_loads(self):
        result = self.loader.load_file(self.json_path)

        self.assertEqual(result["detected_format"], "json")
        self.assertEqual(result["dataframe"].shape, (2, 4))

    def test_fixture_nested_json_flattens(self):
        result = self.loader.load_file(self.json_nested_path)

        self.assertEqual(result["detected_format"], "json")
        self.assertEqual(result["dataframe"].shape, (2, 2))
        self.assertIn("Nested JSON: used array at top-level key 'rows'", result["warnings"])

    def test_fixture_jsonl_loads(self):
        result = self.loader.load_file(self.jsonl_path)

        self.assertEqual(result["detected_format"], "jsonl")
        self.assertEqual(result["dataframe"].shape, (2, 4))

    def test_fixture_encrypted_xlsx_raises_clear_error(self):
        with self.assertRaisesRegex(ValueError, "Password-protected Excel workbooks are not supported"):
            self.loader.load_file(self.encrypted_xlsx_path)

    def test_fixture_corrupt_xlsx_raises_clear_error(self):
        with self.assertRaisesRegex(ValueError, "Could not open workbook"):
            self.loader.load_file(self.corrupt_xlsx_path)

    def test_fixture_corrupt_xls_raises_clear_error(self):
        if not _XLRD_AVAILABLE:
            self.skipTest("xlrd not installed — run: pip install xlrd")

        with self.assertRaisesRegex(ValueError, "Could not open workbook"):
            self.loader.load_file(self.corrupt_xls_path)

    def test_fixture_corrupt_xls_does_not_leak_parser_noise(self):
        if not _XLRD_AVAILABLE:
            self.skipTest("xlrd not installed — run: pip install xlrd")

        stdout = io.StringIO()
        stderr = io.StringIO()
        with redirect_stdout(stdout), redirect_stderr(stderr):
            with self.assertRaisesRegex(ValueError, "Could not open workbook"):
                self.loader.load_file(self.corrupt_xls_path)

        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")


if __name__ == "__main__":
    unittest.main()
