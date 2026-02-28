from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CSV_SCRIPTS = ROOT / "skills" / "csv-doctor" / "scripts"
EXCEL_SCRIPTS = ROOT / "skills" / "excel-doctor" / "scripts"
SAMPLE_CSV = ROOT / "sample-data" / "extreme_mess.csv"
SAMPLE_XLSX = ROOT / "sample-data" / "messy_sample.xlsx"
SCHEMAS_DIR = ROOT / "schemas"

sys.path.insert(0, str(CSV_SCRIPTS))

from diagnose import build_report as build_csv_diagnose_report
from heal import build_structured_summary as build_csv_heal_summary
from heal import execute_healing as execute_csv_healing
from reporter import build_report as build_csv_report


def load_module(module_path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


EXCEL_DIAGNOSE = load_module(EXCEL_SCRIPTS / "diagnose.py", "excel_doctor_diagnose")
EXCEL_HEAL = load_module(EXCEL_SCRIPTS / "heal.py", "excel_doctor_heal")


class ContractTests(unittest.TestCase):
    def test_csv_diagnose_emits_versioned_contract_and_run_summary(self):
        report = build_csv_diagnose_report(SAMPLE_CSV)
        self.assertEqual(report["contract"]["name"], "csv_doctor.diagnose")
        self.assertEqual(report["schema_version"], report["contract"]["version"])
        self.assertIn("tool_version", report)
        self.assertEqual(report["run_summary"]["tool"], "csv-doctor")
        self.assertEqual(report["run_summary"]["script"], "diagnose.py")
        self.assertIn("issues_found", report["run_summary"]["metrics"])

    def test_csv_report_emits_versioned_contract_and_run_summary(self):
        report = build_csv_report(SAMPLE_CSV)
        self.assertEqual(report["contract"]["name"], "csv_doctor.report")
        self.assertEqual(report["schema_version"], report["contract"]["version"])
        self.assertEqual(report["run_summary"]["tool"], "csv-doctor")
        self.assertIn("recoverability_score", report["run_summary"]["metrics"])
        self.assertIn("text_report", report)

    def test_csv_heal_summary_emits_versioned_contract(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "healed.xlsx"
            result = execute_csv_healing(SAMPLE_CSV)
            summary = build_csv_heal_summary(result, input_path=SAMPLE_CSV, output_path=output_path)
        self.assertEqual(summary["contract"]["name"], "csv_doctor.heal_summary")
        self.assertEqual(summary["schema_version"], summary["contract"]["version"])
        self.assertEqual(summary["run_summary"]["script"], "heal.py")
        self.assertIn("clean_rows", summary["rows"])

    def test_csv_heal_summary_persists_confirmed_workbook_plan(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "healed.xlsx"
            result = execute_csv_healing(SAMPLE_CSV)
            summary = build_csv_heal_summary(
                result,
                input_path=SAMPLE_CSV,
                output_path=output_path,
                sheet_name="Transactions",
                consolidate_sheets=False,
                header_row_override=3,
                role_overrides={1: "department", 4: "amount"},
                plan_confirmed=True,
            )
        self.assertEqual(summary["workbook_plan"]["sheet_name"], "Transactions")
        self.assertEqual(summary["workbook_plan"]["header_row_override"], 3)
        self.assertEqual(summary["workbook_plan"]["role_overrides"], {"2": "department", "5": "amount"})
        self.assertTrue(summary["workbook_plan"]["plan_confirmed"])
        self.assertTrue(summary["run_summary"]["metrics"]["plan_confirmed"])

    def test_excel_diagnose_emits_versioned_contract_and_run_summary(self):
        report = EXCEL_DIAGNOSE.build_report(SAMPLE_XLSX)
        self.assertEqual(report["contract"]["name"], "excel_doctor.diagnose")
        self.assertEqual(report["schema_version"], report["contract"]["version"])
        self.assertEqual(report["run_summary"]["tool"], "excel-doctor")
        self.assertIn("verdict", report["run_summary"]["metrics"])

    def test_excel_heal_summary_emits_versioned_contract(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "healed.xlsx"
            changes, stats = EXCEL_HEAL.execute_healing(SAMPLE_XLSX, output_path)
            summary = EXCEL_HEAL.build_structured_summary(
                input_path=SAMPLE_XLSX,
                output_path=output_path,
                changes=changes,
                stats=stats,
            )
        self.assertEqual(summary["contract"]["name"], "excel_doctor.heal_summary")
        self.assertEqual(summary["schema_version"], summary["contract"]["version"])
        self.assertEqual(summary["run_summary"]["script"], "heal.py")
        self.assertIn("changes_logged", summary)

    def test_schema_files_are_valid_json(self):
        schema_files = sorted(SCHEMAS_DIR.glob("*.json"))
        self.assertGreaterEqual(len(schema_files), 5)
        for schema_path in schema_files:
            with self.subTest(schema=schema_path.name):
                payload = json.loads(schema_path.read_text(encoding="utf-8"))
                self.assertIn("$schema", payload)
                self.assertIn("title", payload)


if __name__ == "__main__":
    unittest.main()
