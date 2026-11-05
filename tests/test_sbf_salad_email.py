"""SALAD email detection and same-account checks, using synthetic customer data."""

import io
import unittest

from openpyxl import load_workbook

from checker import ColumnMapping
from test_web_checker import csv_upload, xlsx_upload
from web_checker import prepare_uploaded_check, run_prepared_check


HEADERS = ["email", "cc", "bcc", "indexone_uid", "primary_key", "accountnumber", "chname",
           "agentcode", "agentname", "OB", "CRITERIA", "monthly1", "monthly2", "monthly3",
           "amount1", "amount2", "amount3"]
ROWS = [
    ["alice@example.test", "", "", "000123", "000123", "000987", "Alice Smith",
     "AG01", "First Agent", "1,234.50", "CAMPAIGN A", 1, 2, 3, 100, 200, 300],
    ["bob@example.test", "", "", "000456", "000456", "000654", "Bob Jones",
     "AG02", "Second Agent", "500.00", "CAMPAIGN B", 2, 4, 6, 50, 100, 150],
]


def salad_email_pair(change=None):
    database = xlsx_upload("renamed-database.xlsx", [
        ["DataGrid export"], [],
        ["Name", "Old I.C.", "Account No.", "Email", "OB", "AgentCode", "AgentName"],
        ["Alice Smith", "000123", "000987", "alice@example.test", 1234.5, "AG01", "First Agent"],
        ["Bob Jones", "000456", "000654", "bob@example.test", 500, "AG02", "Second Agent"],
    ])
    rows = [list(row) for row in ROWS]
    if change:
        change(rows)
    return database, csv_upload("renamed-campaign.csv", [HEADERS, *rows])


class SaladEmailTests(unittest.TestCase):
    def check(self, change=None):
        database, upload = salad_email_pair(change)
        prepared = prepare_uploaded_check(database, [upload])
        self.assertEqual(prepared.database_error, "")
        self.assertEqual(prepared.files[0].error, "")
        self.assertEqual(prepared.files[0].source.column, "primary_key")
        run = run_prepared_check(prepared)
        self.assertEqual(run.results[0].error, "")
        return run

    def test_renamed_files_detect_eight_fields_and_preserve_original_rows_in_report(self):
        run = self.check()
        result = run.results[0]
        self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
        self.assertEqual(result.database_match_column, "Old I.C.")
        self.assertEqual(result.count("MATCHED"), 2)
        self.assertEqual(len(result.mappings), 8)
        self.assertIn(ColumnMapping("OB", "OB", "Number"), result.mappings)
        self.assertIn(ColumnMapping("agentcode", "AgentCode"), result.mappings)
        self.assertIn(ColumnMapping("agentname", "AgentName"), result.mappings)
        self.assertEqual(result.ignored_columns, ["cc", "bcc", "CRITERIA", "monthly1", "monthly2",
                                                  "monthly3", "amount1", "amount2", "amount3"])
        book = load_workbook(io.BytesIO(run.report), data_only=True)
        try:
            self.assertEqual(book["File 001"].max_row, 3)
            self.assertEqual(book["File 001"]["H2"].value, "000123")
            mappings = {row[0]: row[1:] for row in book["Mappings 001"].iter_rows(min_row=2, values_only=True)}
            self.assertEqual(mappings["OB"][:3], ("OB", "Number", "CHECKED"))
            self.assertEqual(mappings["monthly1"][2], "NOT CHECKED")
            self.assertIn("no confirmed database mapping", mappings["monthly1"][3])
        finally:
            book.close()

    def test_each_field_must_belong_to_the_identified_account(self):
        for field in ("email", "accountnumber", "chname", "agentcode", "agentname", "OB"):
            with self.subTest(field=field):
                def change(rows):
                    rows[0][HEADERS.index(field)] = rows[1][HEADERS.index(field)]
                result = self.check(change).results[0]
                self.assertEqual(result.count("MATCHED"), 1)
                self.assertEqual(result.count("DIFFERENT"), 1)
                self.assertTrue(any(item[2].split(" ")[0] == field for item in result.differences))

    def test_missing_blank_and_wrong_secondary_identifiers_do_not_switch_accounts(self):
        for key, status in (("unknown", "NOT FOUND"), ("", "BLANK KEY")):
            with self.subTest(key=key):
                def change(rows):
                    for field in ("primary_key", "indexone_uid"):
                        rows[0][HEADERS.index(field)] = key
                result = self.check(change).results[0]
                self.assertEqual(result.count(status), 1)
                self.assertEqual(result.count("MATCHED"), 1)
        def change_secondary(rows):
            rows[0][HEADERS.index("indexone_uid")] = "unknown-secondary"
        result = self.check(change_secondary).results[0]
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertTrue(any(item[2].startswith("indexone_uid") for item in result.differences))


if __name__ == "__main__":
    unittest.main()
