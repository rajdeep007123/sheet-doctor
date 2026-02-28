import importlib.util
import sys
import unittest
from pathlib import Path


REPO_ROOT = Path("/Users/razzo/Documents/For Codex/sheet-doctor")
REPORTER_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "reporter.py"
LOADER_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "loader.py"
EXTREME_MESS_PATH = REPO_ROOT / "sample-data" / "extreme_mess.csv"


def load_module(path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


class ReporterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.reporter = load_module(REPORTER_PATH, "sheet_doctor_reporter_tests")
        cls.loader = load_module(LOADER_PATH, "sheet_doctor_loader_for_reporter_tests")

    def test_file_overview_uses_raw_and_parsed_row_counts(self):
        report = self.reporter.build_report(EXTREME_MESS_PATH)
        overview = report["file_overview"]

        self.assertEqual(overview["rows"], 51)
        self.assertEqual(overview["parsed_rows"], 46)
        self.assertEqual(overview["malformed_rows"], 8)
        self.assertEqual(overview["dropped_rows"], 5)
        self.assertEqual(overview["columns"], 8)
        self.assertIn("Parsed cleanly: 46 rows", report["text_report"])

    def test_reporter_uses_truthful_auto_fixable_flags(self):
        report = self.reporter.build_report(EXTREME_MESS_PATH)
        issues = report["issues"]["warning"] + report["issues"]["info"] + report["issues"]["critical"]

        def find(issue_id: str, column: str):
            for issue in issues:
                if issue["id"] == issue_id and issue["columns"] == [column]:
                    return issue
            self.fail(f"Could not find {issue_id} for {column}")

        self.assertTrue(find("semantic_inconsistent_capitalisation", "Department")["auto_fixable"])
        self.assertTrue(find("semantic_near_duplicates", "Currency")["auto_fixable"])
        self.assertTrue(find("semantic_near_duplicates", "Amount")["auto_fixable"])
        self.assertTrue(find("semantic_near_duplicates", "Status")["auto_fixable"])
        self.assertFalse(find("semantic_near_duplicates", "Department")["auto_fixable"])

    def test_loader_exposes_row_accounting_for_text_files(self):
        result = self.loader.load_file(EXTREME_MESS_PATH)
        row_accounting = result["row_accounting"]

        self.assertEqual(row_accounting["raw_rows_total"], 52)
        self.assertEqual(row_accounting["raw_data_rows_total"], 51)
        self.assertEqual(row_accounting["parsed_rows_total"], 46)
        self.assertEqual(row_accounting["dropped_rows_total"], 5)
        self.assertEqual(row_accounting["malformed_rows_total"], 8)
        self.assertTrue(result["warnings"])

    def test_reporter_exposes_raw_recoverability_and_post_heal_scores(self):
        report = self.reporter.build_report(EXTREME_MESS_PATH)

        raw_score = report["raw_health_score"]["score"]
        recoverability_score = report["recoverability_score"]["score"]
        post_heal_score = report["post_heal_score"]["score"]

        self.assertEqual(raw_score, 32)
        self.assertGreater(recoverability_score, raw_score)
        self.assertGreater(post_heal_score, recoverability_score)
        self.assertIn("Raw Health Score: 32/100", report["text_report"])
        self.assertIn("Recoverability Score:", report["text_report"])
        self.assertIn("Post-Heal Score:", report["text_report"])

    def test_recommended_actions_use_actual_healing_projection(self):
        report = self.reporter.build_report(EXTREME_MESS_PATH)
        actions = report["recommended_actions"]
        projection = report["healing_projection"]

        self.assertTrue(actions)
        self.assertIn("Run sheet-doctor healing now", actions[0])
        self.assertTrue(any("Quarantine tab" in action for action in actions))
        self.assertTrue(any("needs_review=TRUE" in action for action in actions))
        self.assertEqual(projection["clean_rows"], 40)
        self.assertEqual(projection["quarantine_rows"], 5)
        self.assertEqual(projection["needs_review_rows"], 8)


if __name__ == "__main__":
    unittest.main()
