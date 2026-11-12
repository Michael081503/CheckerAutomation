"""HSBC SMS blasts retain helper columns and validate customer fields by CHCODE."""

import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from checker import ColumnMapping, Options, read_table
from scripts.check_files import main
from test_web_checker import csv_upload, xlsx_upload
from web_checker import UploadedSource, prepare_uploaded_check, run_prepared_check, run_uploaded_check


DATABASE_HEADERS = ["CHCODE", "Mobile phone", "Name", "Last Name", "PROD TYPE", "CHCode"]
BLAST_HEADERS = ["Mobile", "ChCode", "ChName", "LastName", "ProductType", "discount", None, None, None]
ALICE = ["0917 123 4567", "00123", "Alice Smith", "Smith", "Credit Card", 80,
         "cached helper G", "cached helper H", "cached helper I"]


def sms_database(rows=None, headers=None):
    return xlsx_upload("renamed source.xlsx", [headers or DATABASE_HEADERS, *(rows if rows is not None else [
        ["00123", 9171234567, "Alice Smith", "Mr./Ms. Smith", "CARDS", "decoy"],
        ["00456", "+639189876543", "Bob Jones", "Mr. Jones", "LOAN", "00123"],
    ])])


def sms_blast(rows=None, headers=None, name="future batch.xlsx"):
    return xlsx_upload(name, [headers or BLAST_HEADERS, *(rows if rows is not None else [ALICE])])


class HSBCSMSDiscountTests(unittest.TestCase):
    def prepare(self, *files, database=None):
        prepared = prepare_uploaded_check(database or sms_database(), list(files) or [sms_blast()])
        self.assertEqual(prepared.database_error, "")
        for item in prepared.files:
            self.assertEqual(item.error, "")
            self.assertFalse(item.inferred)
            self.assertEqual(item.source.column.casefold(), "chcode")
        return prepared

    def test_auto_checks_five_fields_preserving_unnamed_helpers_and_leading_zeros(self):
        prepared = self.prepare()
        self.assertIn("NOT CHECKED", prepared.files[0].note)
        run = run_prepared_check(prepared)
        result = run.results[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
        self.assertEqual(result.records, [(2, "MATCHED", "2", tuple(ALICE))])
        self.assertEqual(result.mappings, [
            ColumnMapping("Mobile", "Mobile phone", "Phone"), ColumnMapping("ChCode", "CHCODE"),
            ColumnMapping("ChName", "Name"), ColumnMapping("LastName", "Last Name"),
            ColumnMapping("ProductType", "PROD TYPE", "Product"),
        ])
        self.assertEqual(result.ignored_columns, ["discount", "Unnamed column G", "Unnamed column H", "Unnamed column I"])
        self.assertIn("approved calculation", result.unchecked_reasons["discount"])
        book = load_workbook(io.BytesIO(run.report), read_only=True, data_only=True)
        try:
            self.assertEqual(tuple(book["File 001"].values)[1][3:], tuple(ALICE))
            self.assertEqual(book["File 001"]["E2"].data_type, "s")
            coverage = {r[0]: r[1:] for r in book["Mappings 001"].iter_rows(min_row=2, values_only=True)}
            self.assertEqual(coverage["discount"][2], "NOT CHECKED")
            self.assertEqual(coverage["Unnamed column I"][2], "NOT CHECKED")
            self.assertIn("titles are ignored", coverage["LastName"][3])
        finally:
            book.close()

    def test_changed_fields_unknown_blank_and_repeated_keys_are_checked_independently(self):
        rows = [ALICE]
        for index, wrong in [(0, "09189876543"), (2, "Bob Jones"), (3, "Jones"), (4, "Personal Loan")]:
            row = ALICE.copy()
            row[index] = wrong
            rows.append(row)
        rows.extend([[ALICE[0], key, *ALICE[2:]] for key in ("unknown", "", "00123")])
        result = run_prepared_check(self.prepare(sms_blast(rows))).results[0]
        self.assertEqual([r[1] for r in result.records],
                         ["MATCHED", *(["DIFFERENT"] * 4), "NOT FOUND", "BLANK KEY", "MATCHED"])
        self.assertEqual([d[0] for d in result.differences], [3, 4, 5, 6])

    def test_reordered_uppercase_headers_preambles_future_names_and_csv_work(self):
        order = [5, 4, 3, 2, 1, 0]
        headers = [BLAST_HEADERS[i].upper() for i in order] + ["", "Extra campaign field"]
        row = [ALICE[i] for i in order] + ["helper", "unverified"]
        uploads = [make(f"renamed.{suffix}", [["Export"], [], headers, row])
                   for make, suffix in ((csv_upload, "csv"), (xlsx_upload, "xlsx"))]
        prepared = self.prepare(*uploads)
        for item in prepared.files:
            self.assertEqual(item.inspection.header_row, 3)
        for result in run_prepared_check(prepared).results:
            self.assertEqual(result.count("MATCHED"), 1)
            self.assertEqual(len(result.mappings), 5)
            self.assertIn("Unnamed column G", result.ignored_columns)
            self.assertIn("Extra campaign field", result.ignored_columns)

    def test_fixed_database_fields_never_use_decoys_or_map_helper_values(self):
        database = sms_database(headers=[*DATABASE_HEADERS, "Mobile", "discount", "Unnamed column G"], rows=[
            ["00123", "09189876543", "Alice Smith", "Mr./Ms. Smith", "CARDS", "decoy",
             ALICE[0], 80, ALICE[6]],
        ])
        result = run_prepared_check(self.prepare(database=database)).results[0]
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual(result.differences[0][2], "Mobile → Mobile phone")
        self.assertEqual(len(result.mappings), 5)

    def test_duplicate_chcodes_cannot_combine_customer_fields_from_different_rows(self):
        database = sms_database(rows=[
            ["00123", "09171234567", "Alice Smith", "Ms. Smith", "CARDS", "decoy"],
            ["00123", "09189876543", "Bob Jones", "Mr. Jones", "LOAN", "decoy"],
        ])
        bob = ["09189876543", "00123", "Bob Jones", "Jones", "Personal Loan", *ALICE[5:]]
        mixed = [*bob[:3], "Smith", *bob[4:]]
        result = run_prepared_check(self.prepare(sms_blast([bob, mixed, ALICE]), database=database)).results[0]
        self.assertEqual([r[1] for r in result.records], ["MATCHED", "DIFFERENT", "MATCHED"])
        self.assertEqual(result.records[0][2], "3")

    def test_missing_required_database_fields_are_errors(self):
        original = ["00123", ALICE[0], "Alice Smith", "Mr./Ms. Smith", "CARDS", "decoy"]
        for missing in DATABASE_HEADERS[1:5]:
            with self.subTest(missing=missing):
                headers = [h for h in DATABASE_HEADERS if h != missing]
                values = [v for h, v in zip(DATABASE_HEADERS, original) if h != missing]
                result = run_prepared_check(self.prepare(database=sms_database([values], headers))).results[0]
                self.assertIn("HSBC SMS discount checks require", result.error)
                self.assertEqual(result.verdict, "ERROR")

    def test_empty_files_blank_phones_and_present_flow_stage_do_not_pass(self):
        empty = run_prepared_check(self.prepare(sms_blast([]))).results[0]
        self.assertIn("no data rows", empty.error)
        database = sms_database([["00123", None, "Alice Smith", "Ms. Smith", "CARDS", "decoy"]])
        blank = [None, *ALICE[1:]]
        # Explicit checks use the same profile without needing the browser adapter.
        result = run_uploaded_check(UploadedSource(database, "Sheet", "CHCODE"),
                                    [UploadedSource(sms_blast([blank]), "Sheet", "ChCode")],
                                    Options(mapped_fields=True, require_flow_stage=True)).results[0]
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual(result.differences[0][4], "(phone number required)")
        for stage in ("", None):
            result = run_prepared_check(self.prepare(sms_blast([[*ALICE, stage]],
                                                [*BLAST_HEADERS, "FLOW_STAGE"]))).results[0]
            self.assertIn("FLOW_STAGE is blank", result.error)

    def test_unrelated_layouts_still_reject_unnamed_data_columns(self):
        upload = xlsx_upload("other.xlsx", [["CHCODE", "Name", None], ["00123", "Alice Smith", "unknown"]])
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / upload.name
            path.write_bytes(upload.data)
            with self.assertRaisesRegex(ValueError, "unnamed column"):
                read_table(path)

    def test_automatic_cli_checks_both_blasts_and_does_not_change_inputs(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            uploads = [sms_database(), sms_blast(name="credit.xlsx"),
                       sms_blast([["09189876543", *ALICE[1:]]], name="offus.xlsx")]
            paths = [root / upload.name for upload in uploads]
            for path, upload in zip(paths, uploads):
                path.write_bytes(upload.data)
            output = root / "report.xlsx"
            stdout = io.StringIO()
            args = ["check_files.py", "--auto", "--database", str(paths[0]), "--files",
                    *map(str, paths[1:]), "--output", str(output)]
            with patch.object(sys, "argv", args), redirect_stdout(stdout):
                self.assertEqual(main(), 0)
            results = [json.loads(line) for line in stdout.getvalue().splitlines() if line.startswith("{")]
            self.assertEqual([r["matched"] for r in results], [1, 0])
            self.assertEqual([r["changed"] for r in results], [0, 1])
            self.assertTrue(all(r["checked_fields"] == 5 and not r["error"] for r in results))
            self.assertTrue(all(path.read_bytes() == upload.data for path, upload in zip(paths, uploads)))
            self.assertTrue(output.is_file())


if __name__ == "__main__":
    unittest.main()
