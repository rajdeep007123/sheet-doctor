from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CLI = [sys.executable, "-m", "sheet_doctor.cli"]


def run_cli(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [*CLI, *args],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )


class SheetDoctorCliTests(unittest.TestCase):
    def test_diagnose_csv_outputs_json_report(self):
        proc = run_cli("diagnose", "sample-data/extreme_mess.csv")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        report = json.loads(proc.stdout)
        self.assertEqual(report["file"], "extreme_mess.csv")
        self.assertEqual(report["detected_format"], "csv")

    def test_diagnose_xlsx_defaults_to_excel_backend(self):
        proc = run_cli("diagnose", "sample-data/messy_sample.xlsx")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        report = json.loads(proc.stdout)
        self.assertEqual(report["file_type"], ".xlsx")
        self.assertEqual(report["workbook_mode"], "workbook-native")

    def test_diagnose_sheet_flag_routes_xlsx_to_tabular_backend(self):
        proc = run_cli("diagnose", "sample-data/messy_sample.xlsx", "--sheet", "Orders")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        report = json.loads(proc.stdout)
        self.assertEqual(report["detected_format"], "xlsx")
        self.assertEqual(report["sheet_name"], "Orders")

    def test_heal_csv_writes_output_workbook_and_summary(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "extreme_cli_healed.xlsx"
            summary_path = Path(tmpdir) / "extreme_cli_healed.json"
            proc = run_cli(
                "heal",
                "sample-data/extreme_mess.csv",
                "--output",
                str(output_path),
                "--json-summary",
                str(summary_path),
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertIn("Mode: schema-specific", proc.stdout)
            self.assertTrue(output_path.exists())
            self.assertTrue(summary_path.exists())
            summary = json.loads(summary_path.read_text())
            self.assertEqual(summary["mode"], "schema-specific")

    def test_heal_xlsx_writes_workbook_native_output_and_summary(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "messy_sample_cli_healed.xlsx"
            summary_path = Path(tmpdir) / "messy_sample_cli_healed.json"
            proc = run_cli(
                "heal",
                "sample-data/messy_sample.xlsx",
                str(output_path),
                "--json-summary",
                str(summary_path),
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertIn("Mode: workbook-native", proc.stdout)
            self.assertTrue(output_path.exists())
            summary = json.loads(summary_path.read_text())
            self.assertEqual(summary["mode"], "workbook-native")
            self.assertIn("before_after_issue_summary", summary)

    def test_report_csv_outputs_text_report(self):
        proc = run_cli("report", "sample-data/extreme_mess.csv")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("SECTION 1 — FILE OVERVIEW", proc.stdout)
        self.assertIn("SECTION 2 — HEALTH SCORE", proc.stdout)

    def test_report_xlsx_outputs_workbook_report_in_auto_mode(self):
        proc = run_cli("report", "sample-data/messy_sample.xlsx")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("sheet-doctor workbook report", proc.stdout)
        self.assertIn("Triage:", proc.stdout)

    def test_workbook_mode_is_rejected_for_legacy_xls(self):
        proc = run_cli(
            "diagnose",
            "tests/fixtures/loader/corrupt.xls",
            "--mode",
            "workbook",
        )
        self.assertNotEqual(proc.returncode, 0)
        self.assertIn("Workbook-native mode is not supported for .xls", proc.stderr)

    def test_report_can_write_json_output_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "report.json"
            proc = run_cli(
                "report",
                "sample-data/extreme_mess.csv",
                "--format",
                "json",
                "--output",
                str(output_path),
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertTrue(output_path.exists())
            payload = json.loads(output_path.read_text())
            self.assertEqual(payload["file_overview"]["file"], "extreme_mess.csv")


if __name__ == "__main__":
    unittest.main()
