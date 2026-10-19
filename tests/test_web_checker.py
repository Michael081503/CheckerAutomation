import csv
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from checker import ColumnMapping, Options
from web_checker import Upload, UploadedSource, clean_filename, configuration_signature, inspect_upload, run_uploaded_check


def csv_upload(name, rows):
    stream = io.StringIO(newline="")
    csv.writer(stream).writerows(rows)
    return Upload(name, stream.getvalue().encode("utf-8-sig"))


def database_upload():
    stream = io.BytesIO()
    book = Workbook()
    sheet = book.active
    sheet.title = "Accounts"
    sheet.append(["CHCODE", "CHCode", "Name", "Last Name", "Email address", "OB", "PROD TYPE"])
    sheet.append(["00123", "00123", "Alice Smith", "Smith", "alice@example.test", 1234.5, "CARDS"])
    sheet.append(["456", "456", "Bob Jones", "Jones", "bob@example.test", 500, "LOAN"])
    book.create_sheet("Other").append(["Other"])
    book.save(stream)
    book.close()
    return Upload("database.xlsx", stream.getvalue())


class WebCheckerTests(unittest.TestCase):
    def database(self):
        return UploadedSource(database_upload(), "Accounts", "CHCODE")

    def test_workbook_upload_inspection_supports_sheets_and_duplicate_headers(self):
        info = inspect_upload(database_upload())
        self.assertEqual(info.sheets, ("Accounts", "Other"))
        self.assertEqual(info.column, "CHCODE")
        self.assertEqual(info.headers[:2], ("CHCODE", "CHCode [2]"))
        self.assertEqual(inspect_upload(database_upload(), "Other").column, "")

    def test_browser_uploads_generate_a_downloadable_report(self):
        first = csv_upload("campaign.csv", [
            ["ChCode", "primary_key", "email", "chname", "chname_format", "OB_AT_ENDORSEMENT", "product_type", "endo"],
            ["00123", "00123", "alice@example.test", "Alice Smith", "Smith", "1,234.50", "Credit Card", "12OCT ENDO"],
        ])
        second = csv_upload("changed.csv", [["CHCODE", "Name"], ["456", "Changed"]])
        progress = []
        run = run_uploaded_check(self.database(), [UploadedSource(first, "CSV", "ChCode"), UploadedSource(second, "CSV", "CHCODE")],
                                 Options(mapped_fields=True), lambda index, total, result: progress.append((index, total)))
        self.assertEqual(run.results[0].verdict, "MAPPED ROWS MATCH")
        self.assertEqual(run.results[1].count("DIFFERENT"), 1)
        self.assertEqual(progress, [(1, 2), (2, 2)])
        self.assertTrue(run.report.startswith(b"PK"))
        self.assertEqual(run.database.path, Path("database.xlsx"))
        book = load_workbook(io.BytesIO(run.report))
        try:
            self.assertEqual(book["Summary"]["B1"].value, "database.xlsx")
            self.assertEqual(book["Summary"]["B9"].value, "campaign.csv")
            self.assertEqual(book["File 001"]["D2"].value, "00123")
            self.assertEqual(book["Mappings 001"]["D9"].value, "NOT CHECKED")
        finally:
            book.close()

    def test_uploads_with_same_filenames_are_kept_separate(self):
        database = UploadedSource(csv_upload("same.csv", [["CHCODE"], ["A"]]), "CSV", "CHCODE")
        first = UploadedSource(csv_upload("same.csv", [["CHCODE"], ["A"]]), "CSV", "CHCODE")
        second = UploadedSource(csv_upload("same.csv", [["CHCODE"], ["X"]]), "CSV", "CHCODE")
        run = run_uploaded_check(database, [first, second], Options(mapped_fields=True))
        self.assertEqual(run.results[0].count("MATCHED"), 1)
        self.assertEqual(run.results[1].count("NOT FOUND"), 1)

    def test_corrupt_target_does_not_stop_valid_uploads(self):
        invalid = UploadedSource(Upload("broken.xlsx", b"not an Excel workbook"), "Sheet1", "CHCODE")
        valid = UploadedSource(csv_upload("valid.csv", [["CHCODE"], ["00123"]]), "CSV", "CHCODE")
        run = run_uploaded_check(self.database(), [invalid, valid], Options(mapped_fields=True))
        self.assertEqual(run.results[0].verdict, "ERROR")
        self.assertEqual(run.results[1].count("MATCHED"), 1)

    def test_custom_field_mappings_survive_browser_adapter(self):
        upload = csv_upload("custom.csv", [["CHCODE", "Recipient"], ["00123", "alice@example.test"]])
        mappings = (ColumnMapping("CHCODE", "CHCODE"), ColumnMapping("Recipient", "Email address"))
        run = run_uploaded_check(self.database(), [UploadedSource(upload, "CSV", "CHCODE", mappings)], Options(mapped_fields=True))
        self.assertEqual(run.results[0].count("MATCHED"), 1)
        self.assertEqual(run.results[0].mappings, list(mappings))

    def test_input_or_option_changes_invalidate_signature(self):
        database = self.database()
        first = UploadedSource(csv_upload("file.csv", [["CHCODE"], ["00123"]]), "CSV", "CHCODE")
        second = UploadedSource(csv_upload("file.csv", [["CHCODE"], ["456"]]), "CSV", "CHCODE")
        signature = configuration_signature(database, [first], Options(mapped_fields=True))
        self.assertNotEqual(signature, configuration_signature(database, [second], Options(mapped_fields=True)))
        self.assertNotEqual(signature, configuration_signature(database, [first], Options(False, True, True)))
        changed_mapping = UploadedSource(first.upload, first.sheet, first.column, (ColumnMapping("CHCODE", "CHCODE"),))
        self.assertNotEqual(signature, configuration_signature(database, [changed_mapping], Options(mapped_fields=True)))

    def test_temporary_uploads_are_removed_after_success_and_failure(self):
        original = tempfile.TemporaryDirectory
        directories = []

        def tracked(*args, **kwargs):
            directory = original(*args, **kwargs)
            directories.append(Path(directory.name))
            return directory

        valid = UploadedSource(csv_upload("valid.csv", [["CHCODE"], ["00123"]]), "CSV", "CHCODE")
        with patch("web_checker.tempfile.TemporaryDirectory", side_effect=tracked):
            run_uploaded_check(self.database(), [valid], Options(mapped_fields=True))
            invalid_database = UploadedSource(csv_upload("invalid.csv", [["Other"], ["A"]]), "CSV", "CHCODE")
            with self.assertRaises(ValueError):
                run_uploaded_check(invalid_database, [valid], Options(mapped_fields=True))
        self.assertEqual(len(directories), 2)
        self.assertTrue(all(not directory.exists() for directory in directories))

    def test_client_paths_cannot_escape_upload_directory(self):
        self.assertEqual(clean_filename("../../database.xlsx"), "database.xlsx")
        self.assertEqual(clean_filename(r"C:\Users\someone\file.csv"), "file.csv")
        with self.assertRaises(ValueError):
            clean_filename("..")

    def test_no_target_files_is_an_error(self):
        with self.assertRaisesRegex(ValueError, "at least one"):
            run_uploaded_check(self.database(), [], Options(mapped_fields=True))
