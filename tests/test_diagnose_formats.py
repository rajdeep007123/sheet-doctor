import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook


REPO_ROOT = Path(__file__).resolve().parents[1]
DIAGNOSE_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "diagnose.py"
REPORTER_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "reporter.py"
SAMPLE_XLSX = REPO_ROOT / "sample-data" / "messy_sample.xlsx"

_ODFPY_AVAILABLE = importlib.util.find_spec("odf") is not None


def load_module(path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


class DiagnoseFormatCoverageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.diagnose = load_module(DIAGNOSE_PATH, "sheet_doctor_diagnose_format_tests")
        cls.reporter = load_module(REPORTER_PATH, "sheet_doctor_reporter_format_tests")

    def test_diagnose_supports_multisheet_xlsx_with_explicit_sheet(self):
        report = self.diagnose.build_report(SAMPLE_XLSX, sheet_name="Orders")
        self.assertEqual(report["detected_format"], "xlsx")
        self.assertEqual(report["sheet_name"], "Orders")
        self.assertEqual(report["run_summary"]["metrics"]["sheet_name"], "Orders")

    def test_reporter_supports_multisheet_xlsx_with_explicit_sheet(self):
        report = self.reporter.build_report(SAMPLE_XLSX, sheet_name="Orders")
        self.assertEqual(report["file_overview"]["format"], "xlsx")
        self.assertEqual(report["file_overview"]["sheet_name"], "Orders")
        self.assertEqual(report["source_reports"]["diagnose"]["sheet_name"], "Orders")

    def test_diagnose_supports_json(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "data.json"
            path.write_text(
                json.dumps(
                    [
                        {"name": "Ada Lovelace", "email": "ada@example.com", "joined": "2023-01-15"},
                        {"name": "Grace Hopper", "email": "grace@example.com", "joined": "2023-01-16"},
                    ]
                ),
                encoding="utf-8",
            )
            report = self.diagnose.build_report(path)

        self.assertEqual(report["detected_format"], "json")
        self.assertEqual(report["column_semantics"]["columns"]["email"]["detected_type"], "email address")

    def test_reporter_supports_json(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "data.json"
            path.write_text(
                json.dumps(
                    [
                        {"name": "Ada Lovelace", "amount": "$125.00", "status": "approved"},
                        {"name": "Grace Hopper", "amount": "$130.00", "status": "pending"},
                    ]
                ),
                encoding="utf-8",
            )
            report = self.reporter.build_report(path)

        self.assertEqual(report["file_overview"]["format"], "json")
        self.assertEqual(report["source_reports"]["diagnose"]["detected_format"], "json")

    def test_diagnose_supports_jsonl(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "data.jsonl"
            path.write_text(
                '{"name":"Ada","url":"https://example.com/a"}\n'
                '{"name":"Grace","url":"https://example.com/b"}\n',
                encoding="utf-8",
            )
            report = self.diagnose.build_report(path)

        self.assertEqual(report["detected_format"], "jsonl")
        self.assertEqual(report["column_semantics"]["columns"]["url"]["detected_type"], "URL")

    def test_reporter_supports_jsonl(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "data.jsonl"
            path.write_text(
                '{"name":"Ada","status":"approved"}\n'
                '{"name":"Grace","status":"pending"}\n',
                encoding="utf-8",
            )
            report = self.reporter.build_report(path)

        self.assertEqual(report["file_overview"]["format"], "jsonl")
        self.assertEqual(report["source_reports"]["diagnose"]["detected_format"], "jsonl")

    def test_diagnose_supports_xlsm(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "macro_like.xlsm"
            wb = Workbook()
            ws = wb.active
            ws.title = "Data"
            ws.append(["Employee", "Amount", "Status"])
            ws.append(["Ada", "$125.00", "Approved"])
            wb.save(path)
            report = self.diagnose.build_report(path)

        self.assertEqual(report["detected_format"], "xlsm")
        self.assertEqual(report["sheet_name"], "Data")

    def test_reporter_supports_xlsm(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "macro_like.xlsm"
            wb = Workbook()
            ws = wb.active
            ws.title = "Data"
            ws.append(["Employee", "Amount", "Status"])
            ws.append(["Ada", "$125.00", "Approved"])
            wb.save(path)
            report = self.reporter.build_report(path)

        self.assertEqual(report["file_overview"]["format"], "xlsm")
        self.assertEqual(report["source_reports"]["diagnose"]["detected_format"], "xlsm")

    def test_diagnose_supports_ods_when_dependency_available(self):
        if not _ODFPY_AVAILABLE:
            self.skipTest("odfpy not installed")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.ods"
            pd.DataFrame(
                {
                    "Employee": ["Ada", "Grace"],
                    "Amount": ["125.00", "130.00"],
                    "Status": ["Approved", "Pending"],
                }
            ).to_excel(path, index=False, engine="odf")
            report = self.diagnose.build_report(path)

        self.assertEqual(report["detected_format"], "ods")

    def test_reporter_supports_ods_when_dependency_available(self):
        if not _ODFPY_AVAILABLE:
            self.skipTest("odfpy not installed")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.ods"
            pd.DataFrame(
                {
                    "Employee": ["Ada", "Grace"],
                    "Amount": ["125.00", "130.00"],
                    "Status": ["Approved", "Pending"],
                }
            ).to_excel(path, index=False, engine="odf")
            report = self.reporter.build_report(path)

        self.assertEqual(report["file_overview"]["format"], "ods")


if __name__ == "__main__":
    unittest.main()
