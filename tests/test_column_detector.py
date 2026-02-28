import importlib.util
import unittest
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
LOADER_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "loader.py"
DETECTOR_PATH = REPO_ROOT / "skills" / "csv-doctor" / "scripts" / "column_detector.py"
EXTREME_MESS_PATH = REPO_ROOT / "sample-data" / "extreme_mess.csv"


def load_module(path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class ColumnDetectorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.loader = load_module(LOADER_PATH, "sheet_doctor_loader_for_detector_tests")
        cls.detector = load_module(DETECTOR_PATH, "sheet_doctor_column_detector")

    def test_extreme_mess_semantics_match_expected_columns(self):
        loaded = self.loader.load_file(EXTREME_MESS_PATH)
        analysis = self.detector.analyse_dataframe(loaded["dataframe"])
        columns = analysis["columns"]

        self.assertEqual(columns["Employee Name"]["detected_type"], "name")
        self.assertIn(
            "Possible PII detected (emails/phones/names)",
            columns["Employee Name"]["suspected_issues"],
        )

        self.assertEqual(columns["Date"]["detected_type"], "date")
        self.assertIn(
            "Mixed date formats detected",
            columns["Date"]["suspected_issues"],
        )
        self.assertEqual(columns["Date"]["min_value"], "2023-01-05")
        self.assertEqual(columns["Date"]["max_value"], "2023-12-28")

        self.assertEqual(columns["Amount"]["detected_type"], "currency/amount")
        self.assertIn(
            "Outliers detected (values outside 3 standard deviations)",
            columns["Amount"]["suspected_issues"],
        )

        self.assertEqual(columns["Currency"]["detected_type"], "currency code")
        self.assertEqual(columns["Notes"]["detected_type"], "free text")
        self.assertEqual(
            analysis["summary"]["detected_types"],
            {
                "categorical": 3,
                "currency code": 1,
                "currency/amount": 1,
                "date": 1,
                "free text": 1,
                "name": 1,
            },
        )

    def test_generic_headers_still_infer_semantic_types(self):
        df = pd.DataFrame(
            {
                "Column1": ["amy@example.com", "bob@example.com", "cara@example.com"],
                "A": ["yes", "no", "yes"],
                "data": ["https://a.example", "https://b.example", "https://c.example"],
                "misc": ["AB-1001", "AB-1002", "AB-1003"],
            }
        )

        analysis = self.detector.analyse_dataframe(df)["columns"]

        self.assertEqual(analysis["Column1"]["detected_type"], "email address")
        self.assertEqual(analysis["A"]["detected_type"], "boolean")
        self.assertEqual(analysis["data"]["detected_type"], "URL")
        self.assertEqual(analysis["misc"]["detected_type"], "ID/code")

    def test_quality_issue_detection_flags_whitespace_and_near_duplicates(self):
        df = pd.DataFrame(
            {
                "Column1": [" Alpha", "alpha", "ALPHA ", "Alpha", None],
            }
        )

        column = self.detector.analyse_dataframe(df)["columns"]["Column1"]

        self.assertEqual(column["detected_type"], "categorical")
        self.assertIn("Inconsistent capitalisation", column["suspected_issues"])
        self.assertIn("Possible duplicates with slight differences", column["suspected_issues"])
        self.assertTrue(
            any(issue.startswith("Trailing/leading whitespace in ") for issue in column["suspected_issues"])
        )

    def test_numeric_and_percentage_ranges_are_computed(self):
        df = pd.DataFrame(
            {
                "Column1": ["10%", "15%", "25%", "50%"],
                "Column2": ["100", "125.5", "90", "110"],
            }
        )

        analysis = self.detector.analyse_dataframe(df)["columns"]

        self.assertEqual(analysis["Column1"]["detected_type"], "percentage")
        self.assertEqual(analysis["Column1"]["min_value"], 10.0)
        self.assertEqual(analysis["Column1"]["max_value"], 50.0)

        self.assertEqual(analysis["Column2"]["detected_type"], "plain number")
        self.assertEqual(analysis["Column2"]["min_value"], 90.0)
        self.assertEqual(analysis["Column2"]["max_value"], 125.5)

    def test_report_builder_returns_loader_metadata(self):
        report = self.detector.build_report(EXTREME_MESS_PATH)

        self.assertEqual(report["file"], str(EXTREME_MESS_PATH))
        self.assertEqual(report["detected_format"], "csv")
        self.assertEqual(report["delimiter"], ",")
        self.assertIn("columns", report)
        self.assertIn("summary", report)


if __name__ == "__main__":
    unittest.main()
