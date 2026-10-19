import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from checker import ColumnMapping, Options, Source, check_files, export_report


class FlowStageTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.options = Options(mapped_fields=True, require_flow_stage=True)

    def source(self, name, rows, mappings=None):
        path = self.folder / name
        with path.open("w", newline="", encoding="utf-8-sig") as stream:
            csv.writer(stream).writerows(rows)
        return Source(path, "CSV", "CHCODE", mappings)

    def database(self):
        return self.source("database.csv", [["CHCODE", "Name"], ["A", "Alice"], ["B", "Bob"]])

    def test_no_reply_identifies_file_and_rows_match_database_without_stage_column(self):
        file = self.source("campaign.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", "NO_REPLY"]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.flow_stage, "NO_REPLY")
        self.assertEqual(result.verdict, "EXACT MATCH")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(len(result.mappings), 2)
        self.assertEqual(result.ignored_columns, [])
        self.assertEqual(result.missing_rows, [])

    def test_same_identifier_does_not_hide_changed_data(self):
        file = self.source("campaign.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Changed", "NO_REPLY"]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.flow_stage, "NO_REPLY")
        self.assertEqual(result.verdict, "ROW ISSUES")
        self.assertEqual(result.differences, [(2, 2, "Name", "Changed", "Alice")])

    def test_identifier_does_not_hide_missing_chcode(self):
        file = self.source("campaign.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["X", "Extra", "NO_REPLY"]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.verdict, "ROW ISSUES")
        self.assertEqual(result.count("NOT FOUND"), 1)

    def test_other_identifiers_label_files_using_same_row_comparison(self):
        files = [self.source(f"{stage}.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", stage]])
                 for stage in ["NEW_ENDO", "SHORT_TERM_ARRANGEMENT"]]
        results = check_files(self.database(), files, self.options)
        self.assertEqual([result.flow_stage for result in results], ["NEW_ENDO", "SHORT_TERM_ARRANGEMENT"])
        self.assertTrue(all(result.verdict == "EXACT MATCH" for result in results))

    def test_header_case_and_stage_case_do_not_change_file_identifier(self):
        file = self.source("case.csv", [["CHCODE", "Name", "Flow_Stage"], ["A", "Alice", " no_reply "]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.flow_stage, "NO_REPLY")
        self.assertEqual(result.verdict, "EXACT MATCH")

    def test_missing_stage_header_is_a_file_error(self):
        file = self.source("missing.csv", [["CHCODE", "Name"], ["A", "Alice"]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("FLOW_STAGE header", result.error)

    def test_blank_stage_cannot_pass(self):
        file = self.source("blank.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", ""]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("blank", result.error)

    def test_mixed_stages_cannot_identify_one_file(self):
        file = self.source("mixed.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", "NO_REPLY"], ["B", "Bob", "NEW_ENDO"]])
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("multiple FLOW_STAGE", result.error)

    def test_stage_is_metadata_even_when_database_has_a_stage_column(self):
        database = self.source("stage-db.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", "UNRELATED_DATABASE_LABEL"]])
        file = self.source("campaign.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", "NO_REPLY"]])
        result = check_files(database, [file], self.options)[0]
        self.assertEqual(result.verdict, "EXACT MATCH")
        self.assertFalse(any(item.file_column == "FLOW_STAGE" for item in result.mappings))

    def test_stage_cannot_be_mapped_to_unrelated_database_fields(self):
        mappings = (ColumnMapping("CHCODE", "CHCODE"), ColumnMapping("FLOW_STAGE", "Name"))
        file = self.source("custom.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", "NO_REPLY"]], mappings)
        result = check_files(self.database(), [file], self.options)[0]
        self.assertEqual(result.verdict, "ERROR")
        self.assertIn("file identifier", result.error)

    def test_full_copy_mode_excludes_identifier_from_data_columns(self):
        file = self.source("copy.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["B", "Bob", "NO_REPLY"], ["A", "Alice", "NO_REPLY"]])
        result = check_files(self.database(), [file], Options(require_flow_stage=True))[0]
        self.assertEqual(result.verdict, "SAME ROWS")
        self.assertEqual(result.flow_stage, "NO_REPLY")
        self.assertFalse(result.extra_columns)

    def test_report_shows_file_identifier_separately_from_checked_fields(self):
        database = self.database()
        file = self.source("campaign.csv", [["CHCODE", "Name", "FLOW_STAGE"], ["A", "Alice", "NO_REPLY"]])
        results = check_files(database, [file], self.options)
        report = self.folder / "report.xlsx"
        export_report(report, database, self.options, results)
        book = load_workbook(report)
        try:
            self.assertEqual(book["Summary"]["K9"].value, "EXACT MATCH")
            self.assertEqual(book["Summary"]["W9"].value, "NO_REPLY")
            self.assertEqual(book["Mappings 001"]["D4"].value, "FILE IDENTIFIER")
            self.assertEqual(book["File 001"]["F2"].value, "NO_REPLY")
        finally:
            book.close()
