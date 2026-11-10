"""Viber discount batches against the standard DataGrid Balance/contact export."""

import io
import unittest

from openpyxl import load_workbook

from checker import ColumnMapping, Options, is_datagrid_balance_viber_pair
from test_sbf_balance_viber import BLAST_HEADERS, balance_database, balance_row
from test_web_checker import csv_upload, xlsx_upload
from web_checker import UploadedSource, prepare_uploaded_check, run_prepared_check, run_uploaded_check


DATABASE_HEADERS = ["Name", "Old I.C.", "Account No.", "Email", "Last Called No.",
                    "Balance", "Amount Referred", "Debtor ID"]


def discount_database(rows=None, headers=None):
    return xlsx_upload("renamed database.xlsx", [headers or DATABASE_HEADERS, *(rows if rows is not None else [
        ["Alice Smith", "00123", "000001", "alice@example.test", "+639171234567", 1234.5, 9000, 1001],
        ["Bob Jones", "00456", "000002", "bob@example.test", "09181234567", 5000, 8000, 1002],
        ["No Contact", "00789", "000003", "other@example.test", None, 300, 300, 1003],
    ])])


class DiscountViberTests(unittest.TestCase):
    def run_rows(self, rows, database=None):
        prepared = prepare_uploaded_check(database or discount_database(), [
            csv_upload("any future batch.csv", [BLAST_HEADERS, *rows]),
        ])
        self.assertEqual(prepared.database_error, "")
        self.assertEqual(prepared.files[0].error, "")
        return run_prepared_check(prepared)

    def test_renamed_csv_and_excel_preambles_reordered_columns_and_extra_fields(self):
        headers = [*reversed(BLAST_HEADERS), "New field"]
        values = [*reversed(balance_row(level="2")), "unverified"]
        uploads = [make(f"unfamiliar filename.{suffix}", [
            ["Export"], [], [h.upper() for h in headers], values,
        ]) for make, suffix in ((csv_upload, "csv"), (xlsx_upload, "xlsx"))]
        prepared = prepare_uploaded_check(discount_database(), uploads)
        self.assertEqual(prepared.database_error, "")
        for item in prepared.files:
            self.assertEqual(item.error, "")
            self.assertEqual(item.source.column, "PHONE")
            self.assertEqual(item.inspection.header_row, 3)
            self.assertFalse(item.inferred)
            self.assertIn("Last Called No.", item.note)
            self.assertIn("Balance", item.note)
            self.assertIn("Level", item.note)
        for result in run_prepared_check(prepared).results:
            self.assertEqual(result.error, "")
            self.assertEqual(result.count("MATCHED"), 1)
            self.assertEqual(result.database_match_column, "Last Called No.")
            self.assertEqual(result.mappings, [ColumnMapping("OB", "Balance", "Number"),
                ColumnMapping("PHONE", "Last Called No.", "Phone"), ColumnMapping("FULLNAME", "Name")])
            self.assertIn("LEVEL", result.ignored_columns)
            self.assertIn("NEW FIELD", result.ignored_columns)

    def test_wrong_values_unknown_and_blank_phones_and_repeated_rows(self):
        result = self.run_rows([
            balance_row(), balance_row(name="Bob Jones"), balance_row(ob="9000"),
            balance_row(phone="09199999999"), balance_row(phone=""), balance_row(),
        ]).results[0]
        self.assertEqual(result.error, "")
        self.assertEqual([row[1] for row in result.records],
                         ["MATCHED", "DIFFERENT", "DIFFERENT", "NOT FOUND", "BLANK KEY", "MATCHED"])
        self.assertEqual([row[0] for row in result.differences], [3, 4])
        self.assertEqual(result.differences[1][3:], ("9000", 1234.5))

    def test_fixed_phone_and_balance_columns_never_switch_to_make_bad_values_pass(self):
        database = discount_database([
            ["Alice Smith", "00123", "000001", "09171234567", "09181234567", 1234.5, 9000, 1001],
        ])
        result = self.run_rows([balance_row(), balance_row(phone="09181234567", ob="9000")], database).results[0]
        self.assertEqual(result.error, "")
        self.assertEqual([row[1] for row in result.records], ["NOT FOUND", "DIFFERENT"])
        self.assertEqual(result.database_match_column, "Last Called No.")
        self.assertEqual(result.differences[0][3:], ("9000", 1234.5))

    def test_shared_phone_requires_name_and_balance_on_one_database_row(self):
        database = discount_database([
            ["Alice Smith", "00123", "000001", "a@example.test", "09171234567", 1234.5, 9000, 1001],
            ["Bob Jones", "00456", "000002", "b@example.test", "+639171234567", 5000, 8000, 1002],
        ])
        result = self.run_rows([balance_row("Bob Jones", ob="5000"),
                                balance_row("Alice Smith", ob="5000"), balance_row()], database).results[0]
        self.assertEqual([row[1] for row in result.records], ["MATCHED", "DIFFERENT", "MATCHED"])
        self.assertEqual(result.records[0][2], "3")

    def test_missing_required_fields_are_errors_and_not_replaced_by_other_columns(self):
        original = ["Alice Smith", "00123", "000001", "09171234567", "09171234567", 1234.5, 1234.5, 1001]
        for missing in ("Last Called No.", "Balance"):
            headers = [h for h in DATABASE_HEADERS if h != missing]
            values = [value for h, value in zip(DATABASE_HEADERS, original) if h != missing]
            result = self.run_rows([balance_row()], discount_database([values], headers)).results[0]
            self.assertEqual(result.verdict, "ERROR")
            self.assertIn("Balance Viber", result.error)

    def test_report_preserves_phone_and_marks_level_and_campaign_fields_unchecked(self):
        run = self.run_rows([balance_row(ob="9000", level="1")])
        book = load_workbook(io.BytesIO(run.report))
        try:
            summary = dict(zip(next(book["Summary"].iter_rows(min_row=8, max_row=8, values_only=True)),
                               next(book["Summary"].iter_rows(min_row=9, max_row=9, values_only=True))))
            self.assertEqual(summary["Database match column"], "Last Called No.")
            self.assertEqual(book["File 001"]["G2"].value, "09171234567")
            self.assertEqual(book["File 001"]["G2"].data_type, "s")
            self.assertEqual(book["File 001"]["M2"].value, "1")
            coverage = {row[0]: row[1:] for row in book["Mappings 001"].iter_rows(min_row=2, values_only=True)}
            self.assertEqual(coverage["Phone"][:3], ("Last Called No.", "Phone", "CHECKED"))
            self.assertEqual(coverage["OB"][:3], ("Balance", "Number", "CHECKED"))
            self.assertIn("same database row", coverage["OB"][3])
            self.assertEqual(coverage["Level"][2], "NOT CHECKED")
            self.assertIn("no confirmed", coverage["Level"][3])
            for column in ("FirstName", "LastName", "CHCode", "Monthly", "Terms", "Promo"):
                self.assertEqual(coverage[column][2], "NOT CHECKED")
        finally:
            book.close()

    def test_explicit_sources_and_flow_stage_still_use_the_fixed_profile(self):
        database = UploadedSource(discount_database(), "Sheet", "Old I.C.")
        for stage in (None, ""):
            rows = [BLAST_HEADERS, balance_row()] if stage is None else [
                [*BLAST_HEADERS, "FLOW_STAGE"], [*balance_row(), stage]]
            source = UploadedSource(csv_upload("direct.csv", rows), "CSV", "Phone")
            result = run_uploaded_check(database, [source], Options(mapped_fields=True, require_flow_stage=True)).results[0]
            if stage is None:
                self.assertEqual(result.count("MATCHED"), 1)
                self.assertEqual(result.database_match_column, "Last Called No.")
            else:
                self.assertIn("FLOW_STAGE is blank", result.error)

    def test_existing_phone_ob_level_export_retains_four_checks(self):
        result = self.run_rows([balance_row()], balance_database()).results[0]
        self.assertEqual(result.database_match_column, "Phone")
        self.assertEqual(len(result.mappings), 4)
        self.assertNotIn("Level", result.ignored_columns)
        self.assertTrue(is_datagrid_balance_viber_pair(BLAST_HEADERS, DATABASE_HEADERS))
        self.assertFalse(is_datagrid_balance_viber_pair(BLAST_HEADERS, DATABASE_HEADERS, "CHCode"))


if __name__ == "__main__":
    unittest.main()
