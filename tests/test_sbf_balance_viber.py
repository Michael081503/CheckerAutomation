"""Phone-based DataGrid Viber exports, independent of filenames and batches."""

import io
import unittest

from openpyxl import load_workbook

from checker import ColumnMapping, Options, is_datagrid_balance_viber_blast
from test_web_checker import csv_upload, xlsx_upload
from web_checker import UploadedSource, prepare_uploaded_check, run_prepared_check, run_uploaded_check


DATABASE_HEADERS = ["Name", "Old I.C.", "Account No.", "Phone", "OB", "Level", "Last Called No."]
BLAST_HEADERS = ["FullName", "FirstName", "LastName", "Phone", "CHCode", "OB",
                 "Monthly", "Terms", "Promo", "Level"]


def balance_database(rows=None, headers=None):
    return xlsx_upload("renamed database.xlsx", [headers or DATABASE_HEADERS, *(rows if rows is not None else [
        ["Alice Smith", "00123", "000001", "+639171234567", 1234.5, "Level_1", "09181234567"],
        ["Bob Jones", "00456", "000002", "09181234567", 5000, "Level_5", None],
    ])])


def balance_row(name="Alice Smith", phone="09171234567", ob="1,234.50", level="Level_1"):
    return [name, "Staff", "Salutation", phone, "STAFFCODE", ob, "600", "2", "Offer", level]


class BalanceViberTests(unittest.TestCase):
    def run_rows(self, rows, database=None):
        prepared = prepare_uploaded_check(database or balance_database(), [
            csv_upload("future export.csv", [BLAST_HEADERS, *rows]),
        ])
        self.assertEqual(prepared.database_error, "")
        self.assertEqual(prepared.files[0].error, "")
        self.assertEqual(prepared.files[0].source.column, "Phone")
        self.assertFalse(prepared.files[0].inferred)
        return run_prepared_check(prepared)

    def test_csv_and_excel_future_batches_use_phone_without_email_or_debtor_id(self):
        uploads = [make(f"future batch.{suffix}", [
            ["Campaign export"], [], BLAST_HEADERS, balance_row(),
            balance_row("Bob Jones", "639181234567", "5,000.00", "Level_5"),
        ]) for make, suffix in ((csv_upload, "csv"), (xlsx_upload, "xlsx"))]
        prepared = prepare_uploaded_check(balance_database(), uploads)
        self.assertEqual(prepared.database.column, "Old I.C.")
        self.assertTrue(all(not item.error for item in prepared.files))
        self.assertTrue(all(item.inspection.header_row == 3 for item in prepared.files))
        for result in run_prepared_check(prepared).results:
            self.assertEqual(result.error, "")
            self.assertEqual(result.database_match_column, "Phone")
            self.assertEqual(result.count("MATCHED"), 2)
            self.assertEqual(result.mappings, [ColumnMapping("FullName", "Name"),
                ColumnMapping("Phone", "Phone", "Phone"), ColumnMapping("OB", "OB", "Number"),
                ColumnMapping("Level", "Level")])
            self.assertEqual(result.ignored_columns, ["FirstName", "LastName", "CHCode", "Monthly", "Terms", "Promo"])

    def test_wrong_name_balance_level_unknown_and_blank_phones_are_reported(self):
        result = self.run_rows([
            balance_row(), balance_row(name="Bob Jones"), balance_row(ob="999"),
            balance_row(level="Level_5"), balance_row(phone="09199999999"),
            balance_row(phone=""), balance_row(),
        ]).results[0]
        self.assertEqual([row[1] for row in result.records], [
            "MATCHED", "DIFFERENT", "DIFFERENT", "DIFFERENT", "NOT FOUND", "BLANK KEY", "MATCHED"])
        self.assertEqual(result.error, "")
        self.assertEqual(len(result.differences), 3)

    def test_duplicate_phone_fields_must_match_one_record(self):
        database = balance_database([
            ["Alice Smith", "00123", "000001", "09171234567", 1234.5, "Level_1", None],
            ["Bob Jones", "00456", "000002", "+639171234567", 5000, "Level_5", None],
        ])
        result = self.run_rows([
            balance_row("Bob Jones", ob="5,000", level="Level_5"),
            balance_row("Alice Smith", ob="5,000", level="Level_5"),
            balance_row(), balance_row(),
        ], database).results[0]
        self.assertEqual([row[1] for row in result.records], ["MATCHED", "DIFFERENT", "MATCHED", "MATCHED"])
        self.assertEqual(result.records[0][2], "3")

    def test_missing_required_database_fields_cannot_silently_reduce_checks(self):
        for missing in ("Phone", "OB", "Level"):
            headers = [h for h in DATABASE_HEADERS if h != missing]
            row = [v for h, v in zip(DATABASE_HEADERS,
                ["Alice Smith", "00123", "000001", "09171234567", 1234.5, "Level_1", "09171234567"])
                if h != missing]
            result = self.run_rows([balance_row()], balance_database([row], headers)).results[0]
            self.assertEqual(result.verdict, "ERROR")
            self.assertIn("Balance Viber", result.error)

    def test_reordered_case_insensitive_headers_and_new_columns(self):
        headers = [*reversed(BLAST_HEADERS), "New campaign field"]
        upload = csv_upload("a new filename.csv", [
            [h.upper() for h in headers], [*reversed(balance_row()), "unverified"],
        ])
        prepared = prepare_uploaded_check(balance_database(), [upload])
        self.assertEqual(prepared.files[0].error, "")
        result = run_prepared_check(prepared).results[0]
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertIn("NEW CAMPAIGN FIELD", result.ignored_columns)

    def test_debtor_templates_retain_their_identifier_detection(self):
        for extra in ("Debtorid", "Debtor", "", "Unnamed column K"):
            self.assertFalse(is_datagrid_balance_viber_blast([*BLAST_HEADERS, extra]))

    def test_report_preserves_originals_differences_and_unchecked_columns(self):
        run = self.run_rows([balance_row(ob="999")])
        book = load_workbook(io.BytesIO(run.report))
        try:
            self.assertEqual(book["File 001"]["I2"].value, "999")
            self.assertEqual(book["File 001"]["G2"].value, "09171234567")
            self.assertEqual(book["File 001"]["G2"].data_type, "s")
            values = list(book["Mappings 001"].values)
            monthly = next(row for row in values if row[0] == "Monthly")
            self.assertIn("NOT CHECKED", monthly)
            self.assertEqual(run.results[0].differences[0][3:], ("999", 1234.5))
        finally:
            book.close()

    def test_explicit_sources_also_support_new_profile_and_present_stage_is_validated(self):
        database = UploadedSource(balance_database(), "Sheet", "Old I.C.")
        for stage in (None, ""):
            rows = [BLAST_HEADERS, balance_row()] if stage is None else [
                [*BLAST_HEADERS, "FLOW_STAGE"], [*balance_row(), stage]]
            source = UploadedSource(csv_upload("direct.csv", rows), "CSV", "Phone")
            result = run_uploaded_check(database, [source], Options(mapped_fields=True, require_flow_stage=True)).results[0]
            if stage is None:
                self.assertEqual(result.count("MATCHED"), 1)
                self.assertEqual(result.database_match_column, "Phone")
            else:
                self.assertIn("FLOW_STAGE is blank", result.error)


if __name__ == "__main__":
    unittest.main()
