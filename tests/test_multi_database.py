"""Independent database checks and a combined verdict for each campaign row."""

import io
import unittest
from dataclasses import replace

from openpyxl import load_workbook

from checker import ColumnMapping, _is_datagrid_total, match_column
from test_web_checker import csv_upload, xlsx_upload
from web_checker import Upload, prepare_uploaded_check, run_prepared_check, run_prepared_checks


ML_HEADERS = ["UID Code/Old IC", None, "CH CODE", "caseNo", "caseAccountNo", "caseBaseAccountNo",
              "caseCustomerName", "caseCustomerEmail"]
BLAST_HEADERS = ["email", "cc", "bcc", "indexone_uid", "primary_key", "accountnumber",
                 "financingnumber", "chname", "agentcode", "agentname", "last4"]


def dual_database_uploads(*, ml_changed=False):
    grid = xlsx_upload("DataGrid(1).xlsx", [
        ["No", "Name", "Old I.C.", "Account No.", "Financing/Card No.", "Email", "Balance"],
        [1, "Alice Smith", "00123", "000001", "000011", "alice@example.test", 100],
        [2, "Bob Jones", "00456", "000002", "000022", "bob@example.test", 200],
        [2, None, None, None, None, "#N/A", 300],
    ])
    master = xlsx_upload("SALAD ML.xlsx", [
        ML_HEADERS,
        ["00123", None, "staff-a", "case-a", "000001", "000011", "Alice Smith", "alice@example.test"],
        ["00456", "preserved unnamed value", "staff-b", "case-b", "000002", "000022", "Bob Jones",
         "changed@example.test" if ml_changed else "bob@example.test"],
    ])
    blast = csv_upload("renamed future campaign.csv", [
        BLAST_HEADERS,
        ["alice@example.test", "", "", "00123", "00123", "000001", "000011", "Alice Smith", "agent-a", "Agent A", "0011"],
        ["bob@example.test", "", "", "00456", "00456", "000002", "000022", "Bob Jones", "agent-b", "Agent B", "0022"],
    ])
    return [grid, master], blast


class MultipleDatabaseTests(unittest.TestCase):
    def prepare(self, databases, files):
        checks = [prepare_uploaded_check(database, files) for database in databases]
        for check in checks:
            self.assertEqual(check.database_error, "")
        return checks

    def test_both_database_schemas_and_export_details_are_detected(self):
        databases, blast = dual_database_uploads()
        checks = self.prepare(databases, [blast])
        self.assertEqual([check.database.column for check in checks], ["Old I.C.", "UID Code/Old IC"])
        self.assertIn("Unnamed column B", checks[1].database_info.headers)
        self.assertEqual(match_column(ML_HEADERS), "UID Code/Old IC")
        progress = []
        run = run_prepared_checks(checks, lambda i, total, result: progress.append((i, total)))
        self.assertEqual(progress, [(1, 2), (2, 2)])
        self.assertEqual([result.count("MATCHED") for result in run.results], [2, 2])
        self.assertEqual(run.matched_in_all, 2)
        self.assertEqual([len(result.mappings) for result in run.results], [6, 6])
        self.assertEqual([result.database_source.path.name for result in run.results], [db.name for db in databases])
        self.assertIn(ColumnMapping("financingnumber", "caseBaseAccountNo"), run.results[1].mappings)
        self.assertEqual(run.results[1].ignored_columns, ["cc", "bcc", "agentcode", "agentname", "last4"])

    def test_one_database_difference_fails_combined_row_and_keeps_provenance(self):
        databases, blast = dual_database_uploads(ml_changed=True)
        run = run_prepared_checks(self.prepare(databases, [blast]))
        self.assertEqual([result.count("MATCHED") for result in run.results], [2, 1])
        self.assertEqual(run.matched_in_all, 1)
        self.assertEqual(run.results[1].differences[0][2:],
                         ("email → caseCustomerEmail", "bob@example.test", "changed@example.test"))
        book = load_workbook(io.BytesIO(run.report), data_only=True)
        try:
            summary = list(book["Summary"].values)
            headers = summary[7]
            self.assertEqual([row[headers.index("Database")] for row in summary[8:]], [db.name for db in databases])
            self.assertEqual([row[headers.index("Database match column")] for row in summary[8:]],
                             ["Old I.C.", "UID Code/Old IC"])
            combined = list(book["Combined 001"].values)
            self.assertEqual(combined[1][:6], (2, "MATCHED", "MATCHED", "2", "MATCHED", "2"))
            self.assertEqual(combined[2][:6], (3, "NEEDS REVIEW", "MATCHED", "3", "DIFFERENT", "3"))
            self.assertEqual(book["File 002"]["G2"].value, "00123")
            self.assertEqual(book["File 002"]["G2"].data_type, "s")
        finally:
            book.close()

    def test_wrong_and_missing_identifiers_do_not_fall_back_to_account_number(self):
        databases, _ = dual_database_uploads()
        blast = csv_upload("issues.csv", [
            BLAST_HEADERS,
            ["alice@example.test", "", "", "00123", "00123", "000001", "000011", "Alice Smith", "", "", ""],
            ["alice@example.test", "", "", "00123", "wrong", "000001", "000011", "Alice Smith", "", "", ""],
            ["alice@example.test", "", "", "missing", "missing", "000001", "000011", "Alice Smith", "", "", ""],
            ["alice@example.test", "", "", "", "", "000001", "000011", "Alice Smith", "", "", ""],
        ])
        run = run_prepared_checks(self.prepare(databases, [blast]))
        for result in run.results:
            self.assertEqual(result.source.column, "indexone_uid")
            self.assertEqual([row[1] for row in result.records], ["MATCHED", "DIFFERENT", "NOT FOUND", "BLANK KEY"])

    def test_duplicate_uid_requires_all_fields_on_one_record(self):
        databases, blast = dual_database_uploads()
        databases[1] = xlsx_upload("SALAD ML.xlsx", [
            ML_HEADERS,
            ["00123", None, "staff-a", "a", "000001", "wrong-finance", "Alice Smith", "alice@example.test"],
            ["00123", None, "staff-a", "b", "wrong-account", "000011", "Alice Smith", "alice@example.test"],
            ["00456", None, "staff-b", "c", "000002", "000022", "Bob Jones", "bob@example.test"],
        ])
        run = run_prepared_checks(self.prepare(databases, [blast]))
        self.assertEqual([row[1] for row in run.results[1].records], ["DIFFERENT", "MATCHED"])

    def test_multiple_files_and_broken_file_stay_in_each_database_report(self):
        databases, blast = dual_database_uploads()
        broken = Upload("broken.xlsx", b"not a workbook")
        run = run_prepared_checks(self.prepare(databases, [broken, blast]))
        self.assertEqual([result.verdict for result in run.results],
                         ["ERROR", "MAPPED ROWS MATCH", "ERROR", "MAPPED ROWS MATCH"])
        self.assertEqual(run.matched_in_all, 2)
        book = load_workbook(io.BytesIO(run.report), data_only=True)
        try:
            self.assertEqual([row[1] for row in list(book["Combined 002"].values)[1:]], ["MATCHED", "MATCHED"])
            self.assertEqual(book["Combined 001"]["A2"].value, "ERROR")
        finally:
            book.close()

    def test_one_layout_error_cannot_pass_combined_rows_from_the_other_database(self):
        databases, blast = dual_database_uploads()
        checks = self.prepare(databases, [blast])
        file = checks[1].files[0]
        checks[1] = replace(checks[1], files=(replace(file, source=replace(file.source, detection_error="Unsupported layout")),))
        run = run_prepared_checks(checks)
        self.assertEqual(run.results[1].verdict, "ERROR")
        self.assertEqual(run.matched_in_all, 0)
        book = load_workbook(io.BytesIO(run.report), data_only=True)
        try:
            self.assertEqual(book["Combined 001"]["B2"].value, "NEEDS REVIEW")
            self.assertEqual(book["Combined 001"]["E2"].value, "ERROR")
        finally:
            book.close()

    def test_invalid_database_and_inconsistent_file_groups_are_rejected(self):
        databases, blast = dual_database_uploads()
        checks = self.prepare(databases, [blast])
        bad = prepare_uploaded_check(Upload("bad.xlsx", b"invalid"), [blast])
        with self.assertRaises(ValueError):
            run_prepared_checks([checks[0], bad])
        with self.assertRaises(ValueError):
            run_prepared_checks([checks[0], replace(checks[1], files=())])

    def test_single_master_list_uses_the_same_six_checks(self):
        databases, blast = dual_database_uploads()
        run = run_prepared_check(self.prepare(databases[1:], [blast])[0])
        self.assertEqual(run.results[0].count("MATCHED"), 2)
        self.assertEqual(len(run.results[0].mappings), 6)

    def test_error_contact_footer_requires_reconciled_totals_and_no_customer_data(self):
        headers = ["No", "Name", "Old I.C.", "Account No.", "Email", "Balance"]
        customer = (1, "Alice Smith", "00123", "000001", "alice@example.test", "100.50")
        footer = (1, None, None, None, "#N/A", "100.50")
        self.assertTrue(_is_datagrid_total(headers, [(2, customer), (3, footer)]))
        for changes in ({0: 2}, {1: "Customer"}, {4: "real@example.test"}, {5: "101"}):
            values = list(footer)
            for index, value in changes.items():
                values[index] = value
            self.assertFalse(_is_datagrid_total(headers, [(2, customer), (3, tuple(values))]))


if __name__ == "__main__":
    unittest.main()
