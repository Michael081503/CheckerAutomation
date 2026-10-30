import io
import unittest
from pathlib import Path
from unittest.mock import patch

from streamlit.testing.v1 import AppTest
from test_cbs_pl import cbs_pl_account_uploads
from test_sbf_pl_viber import BLAST_HEADERS as PL_HEADERS, pl_database, pl_row
from test_salad_endo import endo_blast, endo_database, regular_database

from test_sbf_salad_viber import BLAST_HEADERS, salad_database, salad_row, salad_viber_uploads
from test_web_checker import cbs_apd_uploads, cbs_pl_uploads, csv_upload, database_upload, datagrid_upload, email_blast_upload, hsbc_mobile_uploads, xlsx_upload


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

    def detection_rows(self):
        return next(item.value for item in self.app.dataframe if "Lookup column" in item.value.columns)

    def test_uploads_detect_layouts_without_manual_options(self):
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertFalse(self.app.checkbox)
        self.assertFalse(self.app.selectbox)
        self.assertFalse(self.app.get("data_editor"))
        self.assertFalse(any(item.label == "Advanced options" for item in self.app.expander))
        summary = self.detection_rows()
        self.assertEqual(summary.iloc[0]["Header row"], "1")
        self.assertEqual(summary.iloc[0]["Lookup column"], "ChCode")
        self.assertEqual(summary.iloc[0]["Status"], "Ready")

    def test_no_database_displays_upload_guidance(self):
        self.database = None
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertTrue(self.app.button(key="run_check").disabled)
        self.assertTrue(any("Add a database and your files" in caption.value for caption in self.app.caption))
        self.assertEqual(len(self.app.get("download_button")), 0)

    def test_cbs_pl_buckets_pass_and_disclose_unused_legacy_identifier(self):
        database, files = cbs_pl_account_uploads()
        self.database = BrowserFile(database)
        self.files = [BrowserFile(upload) for upload in files]
        self.run_check()
        run = self.app.session_state["check_run"]
        self.assertTrue(all(result.count("MATCHED") == 2 and len(result.mappings) == 5 for result in run.results))
        self.assertTrue(all(result.source.column == "primary_key" for result in run.results))
        self.assertTrue(any("indexone_uid: Optional identifier is blank in every row" in caption.value
                            for caption in self.app.caption))
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_upload_check_preview_and_report(self):
        self.run_check()
        run = self.app.session_state["check_run"]
        self.assertEqual(run.results[0].count("MATCHED"), 1)
        self.assertEqual(run.results[0].count("DIFFERENT"), 1)
        self.assertEqual(len(run.results[0].mappings), 7)
        self.assertEqual(run.results[0].ignored_columns, ["endo"])
        self.assertEqual(run.results[0].flow_stage, "NO_REPLY")
        self.assertTrue(run.options.ignore_case)
        self.assertTrue(run.options.trim_spaces)
        self.assertTrue(run.options.mapped_fields)
        self.assertTrue(run.options.require_flow_stage)
        self.assertEqual(self.app.metric[0].value, "1")
        self.assertEqual(len(self.app.metric), 3)
        self.assertTrue(all(not expander.proto.expanded for expander in self.app.expander))
        self.assertEqual(len(self.app.get("download_button")), 1)
        self.assertTrue(run.report.startswith(b"PK"))
        self.app.selectbox(key="result_filter_0").select("DIFFERENT").run()
        self.assertFalse(self.app.exception)
        self.assertIn("check_run", self.app.session_state)

    def test_replacing_file_with_same_name_clears_results_and_redetects(self):
        self.run_check()
        self.files = [BrowserFile(csv_upload("campaign.csv", [["First row", "00123"]]))]
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        summary = self.detection_rows()
        self.assertEqual(summary.iloc[0]["Header row"], "No headers")
        self.app.button(key="run_check").click().run()
        result = self.app.session_state["check_run"].results[0]
        self.assertFalse(result.source.has_headers)
        self.assertEqual(result.records[0][0], 1)

    def test_replacing_database_clears_results_and_reselects_file_key(self):
        _, blast = cbs_pl_uploads()
        self.database = BrowserFile(datagrid_upload())
        self.files = [BrowserFile(blast)]
        self.run_check()
        self.assertEqual(self.app.session_state["check_run"].results[0].source.column, "indexone_uid")
        self.database = BrowserFile(xlsx_upload("DataGrid.xlsx", [
            ["Name", "Old I.C.", "Account No.", "Email"],
            ["Alice Smith", "different-primary", "001234567890", "alice@example.test"],
            ["Bob Jones", "456", "009876543210", "bob@example.test"],
        ]))
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(self.detection_rows().iloc[0]["Lookup column"], "primary_key")
        self.app.button(key="run_check").click().run()
        self.assertEqual(self.app.session_state["check_run"].results[0].source.column, "primary_key")

    def test_clear_removes_uploads_results_and_detection(self):
        self.run_check()
        self.app.button(key="clear_session").click().run()
        self.assertFalse(self.app.exception)
        self.assertNotIn("check_run", self.app.session_state)
        self.assertEqual(self.app.session_state["upload_epoch"], 1)
        self.assertTrue(self.app.button(key="run_check").disabled)
        self.assertFalse(self.app.dataframe)

    def test_missing_database_key_is_a_readable_error(self):
        self.database = BrowserFile(csv_upload("db.csv", [["Other"], ["A"]]))
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertIn("CHCODE", self.app.error[0].value)
        self.assertTrue(self.app.button(key="run_check").disabled)

    def test_datagrid_blast_detects_headers_and_shows_mapped_results(self):
        self.database = BrowserFile(datagrid_upload())
        self.files = [BrowserFile(email_blast_upload(changed=True))]
        self.run_check()
        self.assertFalse(self.app.error)
        run = self.app.session_state["check_run"]
        result = run.results[0]
        self.assertEqual(run.database.column, "Old I.C.")
        self.assertTrue(result.source.has_headers)
        self.assertEqual(result.source.column, "primary_key")
        self.assertEqual(result.file_identifier, "email_blast.csv")
        self.assertFalse(result.flow_stage_required)
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual(len(result.mappings), 5)
        self.assertEqual(result.ignored_columns, ["cc", "bcc"])
        self.assertTrue(any("Old I.C." in caption.value and "row 3" in caption.value for caption in self.app.caption))
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_datagrid_blast_success_uses_mapped_fields_label(self):
        self.database = BrowserFile(datagrid_upload())
        self.files = [BrowserFile(email_blast_upload())]
        self.run_check()
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
        self.assertTrue(any("email_blast.csv" in item.label and "Mapped fields match" in item.label
                            for item in self.app.get("status")))
        self.assertFalse(self.app.warning)

    def test_primary_key_is_selected_from_database_matches(self):
        database, blast = cbs_apd_uploads(total_rows=12, changed_emails=11, changed_secondary_uid=True)
        self.database = BrowserFile(database)
        self.files = [BrowserFile(blast)]
        self.run_check()
        self.assertFalse(self.app.error)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.source.column, "primary_key")
        self.assertEqual(self.detection_rows().iloc[0]["Lookup column"], "primary_key")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.count("DIFFERENT"), 11)
        self.assertEqual(len(result.mappings), 6)
        self.assertEqual(result.ignored_columns, ["cc", "bcc"])
        frames = [item.value for item in self.app.dataframe]
        issues = next(frame for frame in frames if "Field" in frame.columns)
        self.assertEqual(len(issues), 12)
        self.assertEqual(issues.iloc[0]["Your file"], "changed1@example.test")
        coverage = next(frame for frame in frames if "Coverage" in frame.columns).set_index("File column")
        self.assertEqual(coverage.loc["campaignbucket", "Database column"], "Account Type")
        self.assertEqual(sum(coverage["Coverage"] == "CHECKED"), 6)
        self.assertTrue(all(coverage.loc[column, "Coverage"] == "NOT CHECKED" for column in ("cc", "bcc")))

    def test_uid_is_selected_when_it_matches_database_better_than_primary(self):
        database, blast = cbs_pl_uploads()
        self.database = BrowserFile(database)
        self.files = [BrowserFile(blast)]
        self.run_check()
        result = self.app.session_state["check_run"].results[0]
        self.assertFalse(result.error)
        self.assertEqual(result.source.column, "indexone_uid")
        self.assertEqual(self.detection_rows().iloc[0]["Lookup column"], "indexone_uid")
        self.assertEqual(result.count("NOT FOUND"), 0)
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertTrue(result.differences[0][2].startswith("primary_key"))

    def test_same_headers_select_different_keys_per_file_from_data(self):
        database, uid_blast = cbs_pl_uploads()
        primary_blast = csv_upload("SBF PL.csv", [
            ["email", "cc", "bcc", "indexone_uid", "primary_key", "chname", "accountnumber", "campaignbucket"],
            ["alice@example.test", "", "", "different-secondary", "00123", "Alice Smith", "001234567890", "PL"],
            ["bob@example.test", "", "", "456", "456", "Bob Jones", "009876543210", "PL"],
        ])
        self.database = BrowserFile(database)
        self.files = [BrowserFile(uid_blast), BrowserFile(primary_blast)]
        self.run_check()
        first, second = self.app.session_state["check_run"].results
        self.assertEqual(first.source.column, "indexone_uid")
        self.assertEqual(second.source.column, "primary_key")
        self.assertEqual(first.count("NOT FOUND"), 0)
        self.assertEqual(second.count("NOT FOUND"), 0)

    def test_hsbc_mobile_blasts_detect_without_stage_and_show_fields(self):
        database, blasts = hsbc_mobile_uploads()
        self.database = BrowserFile(database)
        self.files = [BrowserFile(blast) for blast in blasts]
        self.run_check()
        self.assertFalse(self.app.error)
        self.assertFalse(self.app.warning)
        run = self.app.session_state["check_run"]
        self.assertEqual(len(run.results), 2)
        for result, blast in zip(run.results, blasts):
            with self.subTest(file=blast.name):
                self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
                self.assertEqual(result.file_identifier, blast.name)
                self.assertFalse(result.flow_stage_required)
                self.assertEqual(result.count("MATCHED"), 2)
                self.assertEqual(len(result.mappings), 5)
                phone = next(item for item in result.mappings if item.file_column == "formatted mobile")
                self.assertEqual((phone.database_column, phone.comparison), ("Mobile phone", "Phone"))
        coverage_frames = [item.value for item in self.app.dataframe if "Coverage" in item.value.columns]
        self.assertEqual(len(coverage_frames), 2)
        self.assertTrue(all(all(frame["Coverage"] == "CHECKED") for frame in coverage_frames))
        self.assertTrue(all(result.source.column == "ChCode" for result in run.results))
        self.assertTrue(any("Phone numbers are compared with the database record for the same CHCODE" in caption.value
                            for caption in self.app.caption))
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_phone_upload_without_stage_shows_wrong_account_phone_and_blank_numbers(self):
        self.database = BrowserFile(xlsx_upload("phones.xlsx", [
            ["CHCODE", "Mobile phone"], ["00123", "+639171234567"],
            ["456", "+639181234567"],
        ]))
        self.files = [BrowserFile(csv_upload("phone_blast.csv", [
            ["Phone Number", "CHCODE"],
            ["09171234567", "00123"],
            ["09181234567", "00123"],
            ["", "00123"],
            ["09171234567", "unknown"],
            ["09171234567", ""],
        ]))]
        self.run_check()
        self.assertFalse(self.app.error)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.source.column, "CHCODE")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.count("DIFFERENT"), 2)
        self.assertEqual(result.count("NOT FOUND"), 1)
        self.assertEqual(result.count("BLANK KEY"), 1)
        self.assertEqual(self.detection_rows().iloc[0]["Lookup column"], "CHCODE")
        self.assertEqual(self.detection_rows().iloc[0]["Check"], "Phone by CHCODE")
        self.assertTrue(all(difference[1] == 2 for difference in result.differences))
        issue_rows = next(item.value for item in self.app.dataframe if "Issue" in item.value.columns)
        self.assertEqual(list(issue_rows["Issue"]), ["NOT FOUND", "BLANK KEY"])
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_hsbc_last_name_title_is_ignored_while_real_surname_differences_remain(self):
        self.database = BrowserFile(xlsx_upload("HSBC.xlsx", [
            ["CHCODE", "Mobile phone", "Name", "Last Name", "PROD TYPE"],
            ["00123", "09171234567", "Alice Smith", "Mr./Ms. Smith", "CARDS"],
        ]))
        self.files = [BrowserFile(csv_upload("HSBC SMS.csv", [
            ["formatted mobile", "ChCode", "ChName", "LastName", "ProductType"],
            ["09171234567", "00123", "Alice Smith", "Smith", "Credit Card"],
            ["09171234567", "00123", "Alice Smith", "Jones", "Credit Card"],
        ]))]
        self.run_check()
        self.assertFalse(self.app.error)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual((result.count("MATCHED"), result.count("DIFFERENT")), (1, 1))
        self.assertEqual(result.differences[0][2:], ("LastName → Last Name", "Jones", "Mr./Ms. Smith"))
        self.assertTrue(any("Leading Mr./Ms." in caption.value for caption in self.app.caption))
        coverage = next(frame.value for frame in self.app.dataframe if "Coverage" in frame.value.columns)
        self.assertEqual(list(coverage["Coverage"]).count("CHECKED"), 5)

    def test_salad_viber_blasts_automatically_check_phone_and_name_with_datagrid(self):
        database, blasts = salad_viber_uploads(changed_name=True)
        self.database = BrowserFile(database)
        self.files = [BrowserFile(blast) for blast in blasts]
        self.run_check()
        self.assertFalse(self.app.error)
        run = self.app.session_state["check_run"]
        self.assertEqual([result.count("MATCHED") for result in run.results], [2, 2, 1])
        self.assertEqual([result.count("DIFFERENT") for result in run.results], [0, 0, 1])
        self.assertTrue(all(result.database_match_column == "Last Called No." for result in run.results))
        summary = self.detection_rows()
        self.assertEqual(list(summary["Lookup column"]), ["Phone"] * 3)
        self.assertEqual(list(summary["Check"]), ["Viber phone and name"] * 3)
        coverage = [frame.value for frame in self.app.dataframe if "Coverage" in frame.value.columns]
        self.assertEqual(len(coverage), 3)
        self.assertTrue(all(list(frame["Coverage"]).count("CHECKED") == 2 for frame in coverage))
        self.assertTrue(any("FullName must match Name on the same database row" in caption.value
                            for caption in self.app.caption))
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_pl_viber_uses_debtor_id_and_shows_differences_and_field_coverage(self):
        self.database = BrowserFile(pl_database())
        self.files = [BrowserFile(csv_upload("PL.csv", [
            PL_HEADERS, pl_row(), pl_row("1001", "Bob Jones", "09181234567"),
        ]))]
        self.run_check()
        self.assertFalse(self.app.error)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual((result.count("MATCHED"), result.count("DIFFERENT")), (1, 1))
        self.assertEqual(result.database_match_column, "Debtor ID")
        self.assertEqual(self.detection_rows().iloc[0]["Lookup column"], "Debtorid")
        self.assertEqual(self.detection_rows().iloc[0]["Check"], "Viber debtor ID, phone and name")
        coverage = next(frame.value for frame in self.app.dataframe if "Coverage" in frame.value.columns)
        self.assertEqual(list(coverage["Coverage"]).count("CHECKED"), 3)
        self.assertTrue(any("Debtorid is matched to DataGrid Debtor ID" in caption.value
                            for caption in self.app.caption))
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_customer_names_ignore_question_marks_and_show_the_rule(self):
        self.database = BrowserFile(salad_database())
        self.files = [BrowserFile(csv_upload("viber.csv", [BLAST_HEADERS, salad_row("Al?ice Smith??")]))]
        self.run_check()
        self.assertFalse(self.app.error)
        self.assertFalse(self.app.warning)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.differences, [])
        self.assertTrue(any("Question marks are ignored" in caption.value for caption in self.app.caption))

    def test_replacement_glyph_names_match_and_header_rows_are_excluded(self):
        self.database = BrowserFile(pl_database([
            ["Ana Peña", "00123", "000001", "a@example.test", "09171234567", 1001],
        ]))
        self.files = [BrowserFile(csv_upload("viber.csv", [PL_HEADERS, pl_row(name="Ana Pe\ufffda")]))]
        self.run_check()
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual([row[:3] for row in result.records], [(2, "MATCHED", "2")])
        self.assertEqual(result.differences, [])
        self.assertTrue(any("Header rows are excluded" in caption.value for caption in self.app.caption))
        self.assertTrue(any("replacement character" in caption.value for caption in self.app.caption))

    def test_salad_endorsement_email_blast_uses_account_number_and_leaves_staff_unchecked(self):
        self.database = BrowserFile(endo_database())
        self.files = [BrowserFile(endo_blast())]
        self.run_check()
        self.assertFalse(self.app.error)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.count("MATCHED"), 2)
        self.assertEqual(len(result.records), 2)
        self.assertEqual(self.detection_rows().iloc[0]["Lookup column"], "accountnumber")
        self.assertEqual(self.detection_rows().iloc[0]["Check"], "Account number, name and email")
        self.assertTrue(any("staff assignments are not compared" in caption.value for caption in self.app.caption))
        coverage = next(frame.value for frame in self.app.dataframe if "Coverage" in frame.value.columns)
        self.assertEqual(list(coverage["Coverage"]).count("CHECKED"), 3)
        self.assertEqual(len(self.app.get("download_button")), 1)

    def test_regular_pl_endorsement_is_detected_and_checked_in_the_web_app(self):
        self.database = BrowserFile(regular_database())
        upload = endo_blast()
        self.files = [BrowserFile(type(upload)("SBF PL EMAIL BLAST NEW ENDO.csv", upload.data))]
        self.run_check()
        self.assertFalse(self.app.error)
        result = self.app.session_state["check_run"].results[0]
        self.assertEqual(result.count("MATCHED"), 2)
        self.assertEqual(result.source.column, "accountnumber")
        self.assertEqual(result.database_match_column, "caseAccountNo")
        self.assertEqual(self.detection_rows().iloc[0]["Check"], "Account number, name and email")
        self.assertTrue(any("SBF endorsement record (SALAD / PL)" in caption.value for caption in self.app.caption))
        self.assertIn("agentcode", result.ignored_columns)

    def test_headerless_file_preserves_row_one_and_checks_column_b(self):
        self.files = [BrowserFile(csv_upload("no_headers.csv", [
            ["First row", "00123", "Unmapped value"],
            ["Second row", "00999", "Another value"],
            ["Blank key", "", "Still a data row"],
        ]))]
        self.run_check()
        result = self.app.session_state["check_run"].results[0]
        self.assertFalse(result.error)
        self.assertFalse(result.source.has_headers)
        self.assertEqual(self.detection_rows().iloc[0]["Header row"], "No headers")
        self.assertEqual([(row[0], row[1]) for row in result.records],
                         [(1, "MATCHED"), (2, "NOT FOUND"), (3, "BLANK KEY")])
        self.assertEqual(result.records[0][3], ("First row", "00123", "Unmapped value"))
        self.assertEqual([item.file_column for item in result.mappings], ["CHCODE"])
        self.assertEqual(result.ignored_columns, ["Column A", "Column C"])
        self.assertFalse(result.flow_stage_required)

    def test_mixed_header_formats_are_detected_per_file(self):
        self.files.append(BrowserFile(xlsx_upload("headerless.xlsx", [
            ["First row", "00123", "Any value"], ["Second row", "456", "Other value"],
        ])))
        self.run_check()
        first, second = self.app.session_state["check_run"].results
        self.assertTrue(first.source.has_headers)
        self.assertFalse(second.source.has_headers)
        self.assertEqual(second.count("MATCHED"), 2)
        self.assertEqual(second.records[0][0], 1)
        self.assertEqual(list(self.detection_rows()["Header row"]), ["1", "No headers"])

    def test_database_header_after_preamble_is_detected_in_excel_and_csv(self):
        for make_upload, suffix in ((csv_upload, "csv"), (xlsx_upload, "xlsx")):
            with self.subTest(format=suffix):
                self.database = BrowserFile(make_upload(f"database.{suffix}", [
                    ["Account database export"], [],
                    ["Name", " cHcOdE ", "Last Name", "Email address", "OB", "PROD TYPE"],
                    ["Alice Smith", "00123", "Smith", "alice@example.test", 1234.5, "CARDS"],
                    ["Bob Jones", "456", "Jones", "bob@example.test", 500, "LOAN"],
                ]))
                self.run_check()
                self.assertFalse(self.app.error)
                run = self.app.session_state["check_run"]
                self.assertEqual(run.database.column, "cHcOdE")
                self.assertTrue(run.database.has_headers)
                self.assertEqual([record[:3] for record in run.results[0].records], [
                    (2, "MATCHED", "4"), (3, "DIFFERENT", "5"),
                ])

    def test_bad_target_is_reported_while_other_files_are_checked(self):
        self.files.append(BrowserFile(csv_upload("missing-key.csv", [["Other"], ["A"]])))
        self.run_check()
        run = self.app.session_state["check_run"]
        self.assertEqual(len(run.results), 2)
        self.assertEqual(run.results[1].verdict, "ERROR")
        self.assertEqual(self.app.metric[2].value, "2")
        self.assertNotEqual(self.detection_rows().iloc[1]["Status"], "Ready")
        self.assertFalse(run.results[0].error)

    def test_unsupported_header_is_not_treated_as_headerless(self):
        self.files.append(BrowserFile(csv_upload("unsupported.csv", [
            ["email", "accountnumber", "customer_name"],
            ["alice@example.test", "00123", "Alice Smith"],
        ])))
        self.run_check()
        result = self.app.session_state["check_run"].results[1]
        self.assertTrue(result.error)
        self.assertNotEqual(self.detection_rows().iloc[1]["Status"], "Ready")

    def test_chcode_alias_headers_work_without_stage_or_full_blast_layout(self):
        headers = ["CHCODE", "primary_key", " INDEXONE_UID "]
        self.files = [BrowserFile(csv_upload(f"alias-{index}.csv", [
            ["Name", "Memo", header], ["Alice Smith", "Campaign note", "00123"],
        ])) for index, header in enumerate(headers)]
        self.run_check()
        run = self.app.session_state["check_run"]
        self.assertTrue(run.options.allow_missing_flow_stage)
        self.assertTrue(all(row == "Ready" for row in self.detection_rows()["Status"]))
        for result, header in zip(run.results, headers):
            self.assertEqual(result.error, "")
            self.assertEqual(result.source.column, header.strip())
            self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
            self.assertEqual(result.count("MATCHED"), 1)
            self.assertEqual(len(result.mappings), 2)
            self.assertEqual(result.ignored_columns, ["Memo"])
            self.assertEqual(result.file_identifier_column, "Filename")

    def test_preparation_is_cached_across_result_filter_changes(self):
        import web_checker
        with patch("web_checker.prepare_uploaded_check", wraps=web_checker.prepare_uploaded_check) as prepare:
            self.run_check()
            preparation = self.app.session_state["preparation"]
            self.app.selectbox(key="result_filter_0").select("DIFFERENT").run()
            self.assertIs(self.app.session_state["preparation"], preparation)
            self.assertEqual(prepare.call_count, 1)
            self.assertFalse(self.app.exception)
