import io
import unittest
from pathlib import Path
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from test_web_checker import csv_upload, database_upload


class BrowserFile(io.BytesIO):
    def __init__(self, upload):
        super().__init__(upload.data)
        self.name = upload.name


class StreamlitAppTests(unittest.TestCase):
    def setUp(self):
        self.database = BrowserFile(database_upload())
        self.files = [BrowserFile(csv_upload("campaign.csv", [
            ["ChCode", "primary_key", "email", "chname", "chname_format", "OB_AT_ENDORSEMENT", "product_type", "endo", "FLOW_STAGE"],
            ["00123", "00123", "alice@example.test", "Alice Smith", "Smith", "1,234.50", "Credit Card", "12OCT ENDO", "NO_REPLY"],
            ["456", "456", "bob@example.test", "Changed Name", "Jones", "500.00", "Personal Loan", "12OCT ENDO", "NO_REPLY"],
        ]))]
        self.uploader_patch = patch("streamlit.file_uploader", side_effect=self.uploader)
        self.uploader_patch.start()
        self.addCleanup(self.uploader_patch.stop)
        self.app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "streamlit_app.py"), default_timeout=20)

    def uploader(self, label, *args, **kwargs):
        key = kwargs.get("key", "")
        if key.startswith("database_upload_"):
            return self.database if key.endswith("_0") else None
        return self.files if key.endswith("_0") else []

    def run_check(self):
        self.app.run()
        self.assertFalse(self.app.exception)
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)

    def test_no_database_displays_upload_guidance(self):
        self.database = None
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertTrue(self.app.button(key="run_check").disabled)
        self.assertTrue(any("Add a database and your files" in caption.value for caption in self.app.caption))
        self.assertFalse(self.app.expander[0].proto.expanded)
        self.assertEqual(len(self.app.get("download_button")), 0)

    def test_upload_check_preview_and_report(self):
        self.run_check()
        run = self.app.session_state["check_run"]
        self.assertEqual(run.results[0].count("MATCHED"), 1)
        self.assertEqual(run.results[0].count("DIFFERENT"), 1)
        self.assertEqual(len(run.results[0].mappings), 7)
        self.assertEqual(run.results[0].ignored_columns, ["endo"])
        self.assertEqual(run.results[0].flow_stage, "NO_REPLY")
        self.assertEqual(self.app.metric[0].value, "1")
        self.assertEqual(len(self.app.metric), 3)
        self.assertTrue(all(not expander.proto.expanded for expander in self.app.expander))
        self.assertEqual(len(self.app.get("download_button")), 1)
        self.assertTrue(run.report.startswith(b"PK"))
        self.app.selectbox(key="result_filter_0").select("DIFFERENT").run()
        self.assertFalse(self.app.exception)
        self.assertIn("check_run", self.app.session_state)

    def test_options_change_removes_previous_report(self):
        self.run_check()
        self.app.checkbox(key="trim_spaces").uncheck().run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(len(self.app.get("download_button")), 0)

    def test_replacing_file_with_same_name_clears_results(self):
        self.run_check()
        self.files = [BrowserFile(csv_upload("campaign.csv", [["CHCODE"], ["00123"]]))]
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)

    def test_clear_removes_uploads_and_results(self):
        self.run_check()
        self.app.button(key="clear_session").click().run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(self.app.session_state["upload_epoch"], 1)

    def test_missing_database_key_is_a_readable_error(self):
        self.database = BrowserFile(csv_upload("db.csv", [["Other"], ["A"]]))
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertIn("CHCODE header", self.app.error[0].value)

    def test_bad_target_is_reported_while_other_files_are_checked(self):
        self.files.append(BrowserFile(csv_upload("missing-key.csv", [["Other"], ["A"]])))
        self.run_check()
        run = self.app.session_state["check_run"]
        self.assertEqual(len(run.results), 2)
        self.assertEqual(run.results[1].verdict, "ERROR")
        self.assertEqual(self.app.metric[2].value, "2")

    def test_full_copy_mode_is_available(self):
        self.app.run()
        self.app.selectbox(key="check_mode").select("Compare complete database copy").run()
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)
        self.assertFalse(self.app.session_state["check_run"].options.mapped_fields)
        self.assertEqual(self.app.session_state["check_run"].results[0].verdict, "DIFFERENT ROWS")

    def test_no_reply_file_does_not_require_a_database_flow_stage_column(self):
        self.files = [BrowserFile(csv_upload("reply.csv", [["CHCODE", "chname", "FLOW_STAGE"], ["00123", "Alice Smith", "NO_REPLY"]]))]
        self.run_check()
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.flow_stage, "NO_REPLY")
        self.assertEqual(result.verdict, "EXACT MATCH")
        self.assertNotIn("FLOW_STAGE", result.ignored_columns)

    def test_file_without_stage_cannot_pass(self):
        self.files = [BrowserFile(csv_upload("missing-stage.csv", [["CHCODE"], ["00123"]]))]
        self.run_check()
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("FLOW_STAGE", result.error)

    def test_headerless_file_checks_column_b_and_format_toggle_clears_results(self):
        self.files = [BrowserFile(csv_upload("no_headers.csv", [
            ["First row", "00123", "Unmapped value"],
            ["Second row", "00999", "Another value"],
            ["Blank key", "", "Still a data row"],
        ]))]
        self.app.run()
        headerless = next(item for item in self.app.checkbox if item.label.startswith("No headers"))
        key = headerless.key
        headerless.check().run()
        self.assertFalse(self.app.exception)
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)
        result = self.app.session_state["check_run"].results[0]
        self.assertFalse(result.error)
        self.assertFalse(result.source.has_headers)
        self.assertEqual([(row[0], row[1]) for row in result.records],
                         [(1, "MATCHED"), (2, "NOT FOUND"), (3, "BLANK KEY")])
        self.assertEqual(result.records[0][3], ("First row", "00123", "Unmapped value"))
        self.assertEqual([item.file_column for item in result.mappings], ["CHCODE"])
        self.assertEqual(result.ignored_columns, ["Column A", "Column C"])
        self.assertFalse(result.flow_stage_required)
        self.assertEqual(len(self.app.get("download_button")), 1)
        self.app.checkbox(key=key).uncheck().run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(len(self.app.get("download_button")), 0)

    def test_headerless_full_copy_is_a_readable_file_error(self):
        self.files = [BrowserFile(csv_upload("no_headers.csv", [["First row", "00123"]]))]
        self.app.run()
        next(item for item in self.app.checkbox if item.label.startswith("No headers")).check().run()
        self.app.selectbox(key="check_mode").select("Compare complete database copy").run()
        self.app.button(key="run_check").click().run()
        self.assertFalse(self.app.exception)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("uploaded-row", result.error)
