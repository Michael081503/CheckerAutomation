import io
import unittest

from openpyxl import load_workbook

from checker import ColumnMapping, Options
from test_web_checker import csv_upload, database_upload, xlsx_upload
from web_checker import UploadedSource, prepare_uploaded_check, run_prepared_check, run_uploaded_check


def phone_database():
    return xlsx_upload("phones.xlsx", [
        ["CHCODE", "Mobile phone", "Name"],
        ["00123", 9171234567, "Alice Smith"],
        ["456", "+639181234567", "Bob Jones"],
        ["789", "09171234567", "Shared phone"],
        ["999", None, "No phone"],
    ])


class PhoneByCHCODETests(unittest.TestCase):
    def check(self, upload, database=None):
        prepared = prepare_uploaded_check(database or phone_database(), [upload])
        self.assertEqual(prepared.database_error, "")
        self.assertEqual(prepared.files[0].error, "")
        self.assertTrue(prepared.files[0].source.has_headers)
        return run_prepared_check(prepared)

    def test_missing_stage_checks_phone_only_against_the_same_chcode(self):
        upload = csv_upload("blast.csv", [
            ["Campaign", "CHCODE", "formatted mobile"],
            ["Matching account", "00123", "+63 (917) 123-4567"],
            ["Another account's phone", "00123", "09181234567"],
            ["Known phone, unknown account", "unknown-account", "09171234567"],
            ["Repeated number", "", "639171234567"],
            ["Unknown phone", "00123", "09199999999"],
            ["Blank phone", "456", ""],
            ["Shared phone", "789", "09171234567"],
        ])
        run = self.check(upload)
        result = run.results[0]
        self.assertEqual(result.error, "")
        self.assertFalse(result.flow_stage_required)
        self.assertEqual(result.file_identifier, "blast.csv")
        self.assertEqual(result.source.column, "CHCODE")
        self.assertEqual(result.verdict, "ROW ISSUES")
        self.assertEqual([record[:3] for record in result.records], [
            (2, "MATCHED", "2"), (3, "DIFFERENT", "2"), (4, "NOT FOUND", ""),
            (5, "BLANK KEY", ""), (6, "DIFFERENT", "2"), (7, "DIFFERENT", "3"),
            (8, "MATCHED", "4"),
        ])
        self.assertEqual(result.mappings, [ColumnMapping("CHCODE", "CHCODE"),
                                           ColumnMapping("formatted mobile", "Mobile phone", "Phone")])
        self.assertEqual(result.ignored_columns, ["Campaign"])
        self.assertEqual([(row[0], row[1], row[4]) for row in result.differences], [
            (3, 2, 9171234567), (6, 2, 9171234567), (7, 3, "+639181234567"),
        ])
        book = load_workbook(io.BytesIO(run.report))
        try:
            self.assertEqual(book["Summary"]["D9"].value, "CHCODE")
            self.assertEqual(book["Summary"]["S9"].value, "Uploaded rows")
            self.assertEqual(book["File 001"]["F2"].value, "+63 (917) 123-4567")
            self.assertEqual(book["File 001"]["B5"].value, "BLANK KEY")
            coverage = list(book["Mappings 001"].values)[1:]
            self.assertEqual([row[3] for row in coverage], ["NOT CHECKED", "CHECKED", "CHECKED"])
            self.assertIn("same CHCODE", coverage[-1][4])
            self.assertEqual(book["Differences 001"]["B2"].value, 2)
            self.assertEqual(book["Differences 001"]["E2"].value, 9171234567)
        finally:
            book.close()

    def test_phone_aliases_with_chcode_work_with_csv_excel_and_preambles(self):
        for make_upload, suffix in ((csv_upload, "csv"), (xlsx_upload, "xlsx")):
            for header in ("Phone", "Phone Number", "mobile", "Mobile Number", "Mobile phone",
                           "formatted mobile", "CONTACT_NUMBER", "PhoneNumber"):
                with self.subTest(format=suffix, header=header):
                    upload = make_upload(f"phones.{suffix}", [
                        ["Campaign export"], [], [header, "CHCODE"],
                        ["09171234567", "00123"], [639181234567, "456"],
                    ])
                    result = self.check(upload).results[0]
                    self.assertEqual(result.error, "")
                    self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
                    self.assertEqual(result.source.column, "CHCODE")
                    self.assertIn(ColumnMapping(header, "Mobile phone", "Phone"), result.mappings)
                    self.assertEqual([record[0] for record in result.records], [4, 5])
                    self.assertEqual(result.count("MATCHED"), 2)

    def test_missing_database_phone_column_is_a_file_error(self):
        result = self.check(csv_upload("phones.csv", [
            ["phone", "CHCODE"], ["09171234567", "00123"],
        ]), database_upload()).results[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("database has no supported phone column", result.error)
        self.assertEqual(result.records, [])

    def test_literal_phone_in_headerless_data_keeps_chcode_check_and_row_one(self):
        for rows in ([
            ["Phone", "00123", "Original first row"],
            ["Other account", "456", "Keep this row"],
        ], [
            ["First account", "00123", "Keep this row"],
            ["Phone", "456", "Original second row"],
        ]):
            with self.subTest(rows=rows):
                upload = csv_upload("raw.csv", rows)
                prepared = prepare_uploaded_check(phone_database(), [upload])
                self.assertEqual(prepared.files[0].error, "")
                self.assertFalse(prepared.files[0].source.has_headers)
                result = run_prepared_check(prepared).results[0]
                self.assertEqual([record[0] for record in result.records], [1, 2])
                self.assertEqual(result.count("MATCHED"), 2)

    def test_present_stage_uses_account_comparison_and_validates_stage(self):
        for stage, expected_error in (("", "FLOW_STAGE is blank"), ("MAIL", "")):
            with self.subTest(stage=stage):
                result = self.check(csv_upload("staged.csv", [
                    ["CHCODE", "formatted mobile", "FLOW_STAGE"],
                    ["00123", "09181234567", stage],
                ])).results[0]
                self.assertTrue(result.flow_stage_required)
                if expected_error:
                    self.assertIn(expected_error, result.error)
                else:
                    self.assertEqual(result.error, "")
                    self.assertEqual(result.verdict, "ROW ISSUES")
                    self.assertEqual(result.count("DIFFERENT"), 1)

    def test_empty_phone_file_is_an_error(self):
        result = self.check(csv_upload("empty.csv", [["Phone", "CHCODE"]])).results[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("no data rows", result.error)

    def test_full_copy_keeps_stage_requirement_even_with_phone_fallback_enabled(self):
        database = UploadedSource(phone_database(), "Sheet", "CHCODE")
        source = UploadedSource(csv_upload("phones.csv", [
            ["CHCODE", "Mobile phone", "Name"], ["00123", "09171234567", "Alice Smith"],
        ]), "CSV", "CHCODE")
        result = run_uploaded_check(database, [source], Options(
            require_flow_stage=True, allow_missing_flow_stage=True,
        )).results[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("FLOW_STAGE header", result.error)

    def test_phone_only_files_require_chcode_and_remain_errors_in_report(self):
        for make_upload, suffix in ((csv_upload, "csv"), (xlsx_upload, "xlsx")):
            with self.subTest(format=suffix):
                upload = make_upload(f"phone_only.{suffix}", [["Phone"], ["09171234567"]])
                prepared = prepare_uploaded_check(phone_database(), [upload])
                self.assertIn("CHCODE is required", prepared.files[0].error)
                result = run_prepared_check(prepared).results[0]
                self.assertEqual(result.verdict, "ERROR")
                self.assertEqual(result.records, [])

    def test_blank_phone_does_not_match_blank_database_phone(self):
        result = self.check(csv_upload("blank.csv", [
            ["CHCODE", "Phone"], ["999", ""],
        ])).results[0]
        self.assertEqual(result.verdict, "ROW ISSUES")
        self.assertEqual(result.records[0][:3], (2, "DIFFERENT", "5"))
        self.assertEqual(result.differences[0][4], "(phone number required)")

    def test_duplicate_chcodes_match_only_phones_for_that_chcode_and_allow_repeated_uploads(self):
        database = xlsx_upload("duplicates.xlsx", [
            ["CHCODE", "Mobile phone"],
            ["00123", "09171234567"], ["00123", "09181234567"],
            ["456", "09199999999"],
        ])
        result = self.check(csv_upload("repeated.csv", [
            ["primary_key", "Phone Number"],
            ["00123", "+639181234567"], ["00123", "639181234567"],
            ["00123", "09199999999"],
        ]), database).results[0]
        self.assertEqual(result.source.column, "primary_key")
        self.assertEqual([record[:3] for record in result.records], [
            (2, "MATCHED", "3"), (3, "MATCHED", "3"), (4, "DIFFERENT", "2"),
        ])

    def test_other_mapped_fields_are_still_compared(self):
        result = self.check(csv_upload("name.csv", [
            ["CHCODE", "Mobile phone", "Name"], ["00123", "09171234567", "Changed name"],
        ])).results[0]
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual([item[2] for item in result.differences], ["Name"])


if __name__ == "__main__":
    unittest.main()
