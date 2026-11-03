import io
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

from test_inferred_matching import campaign_pair
from test_streamlit_app import BrowserFile
from test_web_checker import xlsx_upload


class InferredMatchingUITests(unittest.TestCase):
    def setUp(self):
        self.database = None
        self.files = []
        uploader_patch = patch("streamlit.file_uploader", side_effect=self.uploader)
        uploader_patch.start()
        self.addCleanup(uploader_patch.stop)

    def uploader(self, label, *args, **kwargs):
        key = kwargs.get("key", "")
        if key.startswith("database_upload_"):
            return self.database if key.endswith("_0") else None
        return self.files if key.endswith("_0") else []

    def load_profile(self, profile="PARKED"):
        database, upload, headers, rows = campaign_pair(profile)
        self.database = BrowserFile(database)
        self.files = [BrowserFile(upload)]
        self.app = AppTest.from_file(
            str(Path(__file__).resolve().parents[1] / "streamlit_app.py"), default_timeout=20,
        )
        return database, upload

    def run_check(self):
        self.app.run()
        self.assertFalse(self.app.exception)
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)
        return self.app.session_state["check_run"]

    def detection_rows(self):
        return next(frame.value for frame in self.app.dataframe if "Lookup column" in frame.value.columns)

    def test_profiles_display_detected_mapping_evidence_and_downloadable_results(self):
        for profile in ("PARKED", "LEGACY"):
            with self.subTest(profile=profile):
                self.load_profile(profile)
                self.app.run()
                self.assertFalse(self.app.exception)
                detected = self.detection_rows().iloc[0]
                self.assertEqual(detected["Lookup column"], "Unnamed column K" if profile == "PARKED" else "Debtor")
                self.assertEqual(detected["Database lookup column"], "Debtor ID")
                self.assertEqual(detected["Check"], "Detected field mappings")
                self.assertEqual(detected["Status"], "Ready")
                evidence = next(frame.value for frame in self.app.dataframe if "Evidence" in frame.value.columns)
                self.assertEqual(len(evidence), 4)
                phone = evidence.set_index("File column").loc["Phone"]
                self.assertEqual((phone["Database column"], phone["Comparison"]), ("Email", "Phone"))
                self.assertTrue(any(item.label.startswith("Detected field mappings:") for item in self.app.expander))

                self.app.button(key="run_check").click().run()
                self.assertFalse(self.app.exception)
                run = self.app.session_state["check_run"]
                self.assertEqual(run.results[0].count("MATCHED"), 5)
                self.assertEqual(run.results[0].verdict, "MAPPED ROWS MATCH")
                coverage = next(frame.value for frame in self.app.dataframe if "Coverage" in frame.value.columns)
                coverage = coverage.set_index("File column")
                self.assertEqual(coverage.loc["Phone", "Database column"], "Email")
                self.assertEqual(coverage.loc["Phone", "Coverage"], "CHECKED")
                for field in ("FirstName", "CHCode", "OB", "Monthly", "Terms", "Promo", "Level"):
                    self.assertEqual(coverage.loc[field, "Coverage"], "NOT CHECKED")
                self.assertTrue(any("same database record" in caption.value and "Debtor ID" in caption.value
                                    for caption in self.app.caption))
                self.assertEqual(len(self.app.get("download_button")), 1)
                self.assertTrue(run.report.startswith(b"PK"))
                book = load_workbook(io.BytesIO(run.report), data_only=True)
                try:
                    self.assertEqual(book["File 001"].max_row, 6)
                finally:
                    book.close()

    def test_replacing_inferred_upload_with_same_name_clears_results_and_rechecks(self):
        self.load_profile()
        self.assertEqual(self.run_check().results[0].count("MATCHED"), 5)

        def change(headers, rows):
            rows[0][headers.index("Phone")] = rows[1][headers.index("Phone")]

        database, changed_upload, headers, rows = campaign_pair(change=change)
        self.files = [BrowserFile(changed_upload)]
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(len(self.app.get("download_button")), 0)
        self.assertEqual(self.detection_rows().iloc[0]["Database lookup column"], "Debtor ID")
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.count("MATCHED"), 4)
        self.assertEqual(result.count("DIFFERENT"), 1)

    def test_apostrophes_do_not_override_standard_csv_quotes_or_drop_rows(self):
        self.load_profile()

        def change(headers, rows):
            # Most names include literal surrounding apostrophes. Sniffer can
            # mistake them for CSV quoting, splitting the double-quoted comma
            # name into two columns when its guessed quote character is used.
            for number, row in enumerate(rows, 1):
                row[0] = (f"'EXAMPLE CUSTOMER {number:02d}'" if number < len(rows)
                          else "O'BRIEN, EXAMPLE CUSTOMER")

        database, upload, headers, file_rows = campaign_pair(change=change)
        book = load_workbook(io.BytesIO(database.data), data_only=True)
        try:
            database_rows = [list(row) for row in book.active.values]
        finally:
            book.close()
        for database_row, file_row in zip(database_rows[1:], file_rows):
            database_row[0] = file_row[0]
        self.database = BrowserFile(xlsx_upload(database.name, database_rows))
        self.files = [BrowserFile(upload)]
        run = self.run_check()
        result = run.results[0]
        self.assertEqual(result.count("MATCHED"), 5)
        self.assertEqual(len(result.records), 5)
        self.assertEqual(len(result.headers), len(headers))
        self.assertEqual([record[3][0] for record in result.records], [row[0] for row in file_rows])

    def test_replacing_inferred_database_with_same_name_clears_results_and_rechecks(self):
        database, upload = self.load_profile("LEGACY")
        self.assertEqual(self.run_check().results[0].count("MATCHED"), 5)
        book = load_workbook(io.BytesIO(database.data), data_only=True)
        try:
            database_rows = [list(row) for row in book.active.values]
        finally:
            book.close()
        database_rows[1][database_rows[0].index("Email")] = "9179999999"
        self.database = BrowserFile(xlsx_upload(database.name, database_rows))
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(len(self.app.get("download_button")), 0)
        self.assertEqual(self.detection_rows().iloc[0]["Check"], "Detected field mappings")
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.count("MATCHED"), 4)
        self.assertEqual(result.count("DIFFERENT"), 1)


if __name__ == "__main__":
    unittest.main()
