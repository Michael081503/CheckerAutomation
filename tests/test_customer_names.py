import io
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from checker import ColumnMapping, Options, Source, check_files, compare_value
from test_sbf_salad_viber import BLAST_HEADERS, salad_database, salad_row
from test_web_checker import csv_upload, xlsx_upload
from web_checker import UploadedSource, prepare_uploaded_check, run_prepared_check, run_uploaded_check


class CustomerNameTests(unittest.TestCase):
    def test_question_marks_are_ignored_on_both_sides_with_existing_case_and_space_options(self):
        self.assertEqual(compare_value(" ?Al?ice Smith?? ", "Text", Options(), customer_name=True), "alice smith")
        self.assertEqual(compare_value("?Alice?", "Text", Options(False, False), customer_name=True), "Alice")
        self.assertEqual(compare_value(" Alice? ", "Text", Options(False, False), customer_name=True), " Alice ")
        self.assertNotEqual(compare_value("Alice?", "Text", Options(ignore_case=False), customer_name=True),
                            compare_value("alice", "Text", Options(ignore_case=False), customer_name=True))
        self.assertNotEqual(compare_value("Anne-Marie?", "Text", Options(), customer_name=True),
                            compare_value("AnneMarie", "Text", Options(), customer_name=True))
        self.assertEqual(compare_value(None, "Text", Options(), customer_name=True), "")
        self.assertEqual(compare_value("test?@example.test", "Text", Options()), "test?@example.test")

    def test_all_supported_name_fields_ignore_questions_and_preserve_report_values(self):
        headers = ["CHCODE", "Name", "FullName", "ChName", "First Name", "Last Name", "Middle Name",
                   "chname_format", "Customer Name", "Name"]
        database = xlsx_upload("database.xlsx", [headers, ["00123", *(["Al?ice Smith"] * 9)]])
        upload = csv_upload("names.csv", [headers, ["00123", *(["Alice? Smith??"] * 9)]])
        run = run_prepared_check(prepare_uploaded_check(database, [upload]))
        result = run.results[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(len(result.mappings), 10)
        book = load_workbook(io.BytesIO(run.report))
        try:
            self.assertEqual(book["File 001"]["E2"].value, "Alice? Smith??")
            coverage = list(book["Mappings 001"].values)[1:]
            self.assertTrue(all("question marks are ignored" in row[4] for row in coverage[1:]))
            self.assertNotIn("question marks", coverage[0][4])
        finally:
            book.close()

    def test_viber_name_matches_stay_on_the_same_phone_record(self):
        upload = csv_upload("viber.csv", [
            BLAST_HEADERS, salad_row("Al?ice Smith??"),
            salad_row("?Bob Jones", "09171234567"),
        ])
        result = run_prepared_check(prepare_uploaded_check(salad_database(), [upload])).results[0]
        self.assertEqual([row[:3] for row in result.records], [(2, "MATCHED", "2"), (3, "DIFFERENT", "2")])
        self.assertEqual(result.differences[0][3:], ("?Bob Jones", "Alice Smith"))

    def test_question_marks_in_keys_emails_and_notes_still_produce_issues(self):
        headers = ["CHCODE", "Name", "Email", "Note"]
        database = csv_upload("database.csv", [headers, ["00123", "Alice", "alice@example.test", "Ready"]])
        upload = csv_upload("other_fields.csv", [
            headers,
            ["00123", "Alice?", "alice?@example.test", "Ready?"],
            ["00123?", "Alice", "alice@example.test", "Ready"],
        ])
        result = run_prepared_check(prepare_uploaded_check(database, [upload])).results[0]
        self.assertEqual([row[1] for row in result.records], ["DIFFERENT", "NOT FOUND"])
        self.assertEqual([difference[2] for difference in result.differences], ["Email", "Note"])

    def test_explicit_customer_name_mapping_also_ignores_questions(self):
        database = UploadedSource(csv_upload("database.csv", [
            ["CHCODE", "Name"], ["00123", "Alice?"],
        ]), "CSV", "CHCODE")
        source = UploadedSource(csv_upload("mapped.csv", [
            ["CHCODE", "Recipient"], ["00123", "?Alice"],
        ]), "CSV", "CHCODE", (ColumnMapping("CHCODE", "CHCODE"), ColumnMapping("Recipient", "Name")))
        result = run_uploaded_check(database, [source], Options(mapped_fields=True)).results[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.count("MATCHED"), 1)

    def test_full_copy_ignores_name_questions_and_keeps_duplicate_counts(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            database = root / "database.csv"
            source = root / "source.csv"
            headers = ["CHCODE", "Name", "Note"]
            database.write_bytes(csv_upload(database.name, [headers, ["00123", "Alice?", "Ready"],
                                                          ["00123", "?Alice", "Ready"]]).data)
            source.write_bytes(csv_upload(source.name, [headers, ["00123", "Alice", "Ready"],
                                                       ["00123", "Alice??", "Ready"]]).data)
            db = Source(database, "CSV", "CHCODE")
            target = Source(source, "CSV", "CHCODE")
            result = check_files(db, [target], Options())[0]
            self.assertEqual(result.verdict, "SAME ROWS")
            source.write_bytes(csv_upload(source.name, [headers, ["00123", "Alice", "Ready"],
                                                       ["00123", "Alice??", "Ready"],
                                                       ["00123", "Alice?", "Ready?"]]).data)
            result = check_files(db, [target], Options())[0]
            self.assertEqual(result.count("MATCHED"), 2)
            self.assertEqual(result.count("EXTRA ROW"), 1)

    def test_full_copy_retains_real_name_and_other_field_differences(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            headers = ["CHCODE", "Name", "Note"]
            database = root / "database.csv"
            target = root / "source.csv"
            database.write_bytes(csv_upload(database.name, [headers, ["00123", "Alice?", "Ready"]]).data)
            target.write_bytes(csv_upload(target.name, [headers, ["00123", "Bob?", "Ready?"]]).data)
            result = check_files(Source(database, "CSV", "CHCODE"), [Source(target, "CSV", "CHCODE")], Options())[0]
            self.assertEqual(result.count("DIFFERENT"), 1)
            self.assertEqual([row[2] for row in result.differences], ["Name", "Note"])


if __name__ == "__main__":
    unittest.main()
