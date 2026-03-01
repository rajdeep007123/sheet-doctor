from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CLI = [sys.executable, "-m", "sheet_doctor.cli"]
FIXED_STAMP = "20260301T010203Z"


def run_cli(*args: str, env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    merged_env = dict(os.environ)
    merged_env["SHEET_DOCTOR_OUTPUT_STAMP"] = FIXED_STAMP
    if env:
        merged_env.update(env)
    return subprocess.run(
        [*CLI, *args],
        cwd=ROOT,
        capture_output=True,
        text=True,
        env=merged_env,
    )


class SheetDoctorCliTests(unittest.TestCase):
    def test_diagnose_dirty_csv_returns_exit_3_and_writes_default_report(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            proc = run_cli(
                "diagnose",
                "sample-data/extreme_mess.csv",
                "--out",
                tmpdir,
            )
            self.assertEqual(proc.returncode, 3, proc.stderr)
            self.assertIn("Report written:", proc.stderr)
            report_path = Path(tmpdir) / "report.json"
            self.assertTrue(report_path.exists())
            report = json.loads(report_path.read_text())
            self.assertEqual(report["file"], "extreme_mess.csv")

    def test_diagnose_clean_csv_returns_exit_0(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            clean_path = Path(tmpdir) / "clean.csv"
            clean_path.write_text("name,amount\nAlice,10\nBob,20\n", encoding="utf-8")
            proc = run_cli("diagnose", str(clean_path), "--json")
            self.assertEqual(proc.returncode, 0, proc.stderr)
            report = json.loads(proc.stdout)
            self.assertEqual(report["summary"]["issue_count"], 0)

    def test_diagnose_unreadable_input_returns_exit_2(self):
        proc = run_cli("diagnose", "tests/fixtures/loader/corrupt.xlsx")
        self.assertEqual(proc.returncode, 2)
        self.assertIn("Could not read workbook", proc.stderr)

    def test_diagnose_json_stdout_contains_only_json(self):
        proc = run_cli("diagnose", "sample-data/extreme_mess.csv", "--json")
        self.assertEqual(proc.returncode, 3, proc.stderr)
        report = json.loads(proc.stdout)
        self.assertEqual(report["file"], "extreme_mess.csv")
        self.assertEqual(proc.stderr.strip(), "")

    def test_diagnose_default_output_directory_uses_contract_layout(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            proc = run_cli(
                "diagnose",
                "sample-data/extreme_mess.csv",
                env={"PWD": tmpdir},
            )
            self.assertEqual(proc.returncode, 3, proc.stderr)
            output_dir = ROOT / "sheet-doctor-output" / f"extreme_mess-{FIXED_STAMP}"
            try:
                self.assertTrue((output_dir / "report.json").exists())
            finally:
                if output_dir.parent.exists():
                    for child in output_dir.parent.iterdir():
                        if child.is_dir() and child.name == f"extreme_mess-{FIXED_STAMP}":
                            for nested in child.iterdir():
                                nested.unlink()
                            child.rmdir()
                    if not any(output_dir.parent.iterdir()):
                        output_dir.parent.rmdir()

    def test_heal_returns_exit_4_when_quarantine_exists(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            proc = run_cli("heal", "sample-data/extreme_mess.csv", "--out", tmpdir)
            self.assertEqual(proc.returncode, 4, proc.stderr)
            self.assertIn("Quarantine rows:", proc.stderr)
            self.assertTrue((Path(tmpdir) / "heal-summary.json").exists())

    def test_heal_fail_on_quarantine_returns_exit_5(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            proc = run_cli(
                "heal",
                "sample-data/extreme_mess.csv",
                "--out",
                tmpdir,
                "--fail-on-quarantine",
            )
            self.assertEqual(proc.returncode, 5, proc.stderr)

    def test_heal_dry_run_does_not_write_files(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            proc = run_cli(
                "heal",
                "sample-data/extreme_mess.csv",
                "--out",
                tmpdir,
                "--dry-run",
                "--json",
            )
            self.assertEqual(proc.returncode, 4, proc.stderr)
            payload = json.loads(proc.stdout)
            self.assertEqual(payload["mode"], "schema-specific")
            self.assertFalse(any(Path(tmpdir).iterdir()))

    def test_heal_xlsx_defaults_to_workbook_native(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            proc = run_cli(
                "heal",
                "sample-data/messy_sample.xlsx",
                "--out",
                tmpdir,
                "--json",
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            summary = json.loads(proc.stdout)
            self.assertEqual(summary["mode"], "workbook-native")
            self.assertEqual(summary["workbook_triage"]["classification"], "manual_spreadsheet_review_required")

    def test_report_json_stdout_is_machine_readable(self):
        proc = run_cli("report", "sample-data/extreme_mess.csv", "--json")
        self.assertEqual(proc.returncode, 3, proc.stderr)
        payload = json.loads(proc.stdout)
        self.assertIn("run_summary", payload)
        self.assertIn("source_reports", payload)
        self.assertEqual(proc.stderr.strip(), "")

    def test_validate_success_and_failure(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            schema_path = Path(tmpdir) / "schema.json"
            schema_path.write_text(
                json.dumps(
                    {
                        "required_columns": ["Employee Name", "Amount"],
                        "types": {"Amount": "currency/amount"},
                    }
                ),
                encoding="utf-8",
            )
            good = run_cli("validate", "sample-data/extreme_mess.csv", "--schema", str(schema_path), "--json")
            self.assertEqual(good.returncode, 0, good.stderr)
            self.assertTrue(json.loads(good.stdout)["valid"])

            bad_schema = Path(tmpdir) / "bad_schema.json"
            bad_schema.write_text(
                json.dumps(
                    {
                        "required_columns": ["Employee Name", "Missing Column"],
                        "types": {"Amount": "date"},
                    }
                ),
                encoding="utf-8",
            )
            bad = run_cli("validate", "sample-data/extreme_mess.csv", "--schema", str(bad_schema), "--json")
            self.assertEqual(bad.returncode, 5, bad.stderr)
            payload = json.loads(bad.stdout)
            self.assertFalse(payload["valid"])
            self.assertIn("Missing Column", payload["missing_columns"])

    def test_config_init_writes_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "sheet-doctor.yml"
            proc = run_cli("config", "init", "--path", str(config_path))
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertTrue(config_path.exists())
            self.assertIn("max_rows_scan", config_path.read_text())

    def test_explain_outputs_stable_rule_text(self):
        proc = run_cli("explain", "date_mixed_formats")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("Rule: date_mixed_formats", proc.stdout)
        self.assertIn("Auto-fixable:", proc.stdout)

    def test_version_prints_version(self):
        proc = run_cli("version")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertRegex(proc.stdout.strip(), r"^\d+\.\d+\.\d+$")


if __name__ == "__main__":
    unittest.main()
