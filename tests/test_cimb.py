"""CIMB checks keep rounded account IDs in review while auditing contacts."""

import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from checker import ColumnMapping, Options
from scripts.check_files import main
from test_auto_detection import workbook_upload
from test_web_checker import csv_upload, xlsx_upload
from web_checker import UploadedSource, prepare_uploaded_check, run_prepared_check, run_uploaded_check


DATABASE_HEADERS = ["PLACEMENT", "LOAN ACCOUNT NO.", "CUST_NAME", "EMAIL ADDRESS", "CUST_NO"]
BLAST_HEADERS = ["email", "cc", "bcc", "name", "primary_key", "indexone_uid", "PRODUCT"]
ALICE = ["alice@example.test", "", "", "Alice Smith", "20860123456789", "", "PERSONAL LOAN"]
DATABASE_ROWS = [
    ["PLOAN CO", "20860123456789", "Alice Smith", "alice@example.test", "000001"],
    ["LAZPAY PCO", "20860123456790", "Bob Jones", "bob@example.test", "000002"],
]


def cimb_database(rows=None, *, multiple_sheets=False):
    active = [DATABASE_HEADERS, *(DATABASE_ROWS if rows is None else rows)]
    if multiple_sheets:
        return workbook_upload("future master.xlsx", [
            ("ENDO SUMMARY", [["Summary"]]),
            ("POUT", [DATABASE_HEADERS, ["PLOAN CO", "20860123456789", "Wrong Name", "alice@example.test", "000001"]]),
            ("ACTIVE", active),
            ("WAIVER", [["loan_no", "cust_no", "Full Name"], ["20860123456789", "000001", "Wrong Name"]]),
        ])
    return xlsx_upload("renamed master.xlsx", active)


def cimb_blast(rows=None):
    return csv_upload("future campaign.csv", [BLAST_HEADERS, *(rows if rows is not None else [ALICE])])


class CIMBTests(unittest.TestCase):
    def prepare(self, file=None, database=None):
        prepared = prepare_uploaded_check(database or cimb_database(), [file or cimb_blast()])
        self.assertEqual(prepared.database_error, "")
        self.assertEqual(prepared.files[0].error, "")
        self.assertFalse(prepared.files[0].inferred)
        return prepared

    def test_active_is_selected_without_combining_other_master_sheets(self):
        prepared = self.prepare(database=cimb_database(multiple_sheets=True))
        self.assertEqual(prepared.database.sheet, "ACTIVE")
        self.assertEqual(prepared.database.column, "LOAN ACCOUNT NO.")
        self.assertEqual(prepared.files[0].source.column, "primary_key")
        result = run_prepared_check(prepared).results[0]
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.mappings, [
            ColumnMapping("email", "EMAIL ADDRESS"), ColumnMapping("name", "CUST_NAME"),
            ColumnMapping("primary_key", "LOAN ACCOUNT NO."), ColumnMapping("PRODUCT", "PLACEMENT", "Product"),
        ])
        self.assertEqual(result.ignored_columns, ["cc", "bcc", "indexone_uid"])

    def test_full_accounts_check_other_fields_and_never_switch_identity(self):
        rows = [ALICE]
        for position, wrong in [(0, "bob@example.test"), (3, "Bob Jones"), (6, "LAZPAY"),
                                (4, "unknown"), (4, "")]:
            row = ALICE.copy()
            row[position] = wrong
            rows.append(row)
        result = run_prepared_check(self.prepare(cimb_blast(rows))).results[0]
        self.assertEqual([r[1] for r in result.records],
                         ["MATCHED", "DIFFERENT", "DIFFERENT", "DIFFERENT", "NOT FOUND", "BLANK KEY"])

    def test_scientific_ids_find_email_candidates_but_cannot_pass(self):
        row = [*ALICE[:4], "2.08601E+13", *ALICE[5:]]
        prepared = self.prepare(cimb_blast([row]))
        self.assertEqual(prepared.files[0].source.column, "email")
        self.assertEqual(prepared.files[0].source.database_column, "EMAIL ADDRESS")
        self.assertIn("cannot pass", prepared.files[0].note)
        run = run_prepared_check(prepared)
        result = run.results[0]
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual(result.differences, [(2, 2, "primary_key → LOAN ACCOUNT NO.", "2.08601E+13", ALICE[4])])
        self.assertIn("cannot be recovered", result.validation_notes[0])
        book = load_workbook(io.BytesIO(run.report), read_only=True, data_only=True)
        try:
            self.assertEqual(book["File 001"]["H2"].value, "2.08601E+13")
            self.assertEqual(book["File 001"]["H2"].data_type, "s")
            coverage = {r[0]: r for r in book["Mappings 001"].iter_rows(min_row=2, values_only=True)}
            self.assertIn("never passes", coverage["primary_key"][4])
            self.assertEqual(coverage["indexone_uid"][3], "NOT CHECKED")
            counts = {r[0]: r[2] for r in book["Field issues 001"].iter_rows(min_row=2, values_only=True)}
            self.assertEqual(counts, {"email": 0, "name": 0, "primary_key": 1, "PRODUCT": 0})
            summary = list(book["Summary"].values)
            self.assertIn("cannot be recovered", summary[8][summary[7].index("Check notes")])
        finally:
            book.close()

    def test_shared_emails_cannot_combine_different_account_records(self):
        db_rows = [DATABASE_ROWS[0], ["LAZPAY CO", "20860123456790", "Bob Jones", ALICE[0], "000002"]]
        rows = [[*ALICE[:4], "2.08601E+13", *ALICE[5:]],
                [ALICE[0], "", "", "Bob Jones", ALICE[4], "", "LAZPAY"],
                [ALICE[0], "", "", "Bob Jones", "20860123456790", "", "LAZPAY"]]
        result = run_prepared_check(self.prepare(cimb_blast(rows), cimb_database(db_rows))).results[0]
        self.assertEqual([r[1] for r in result.records], ["DIFFERENT", "DIFFERENT", "MATCHED"])
        self.assertEqual(result.records[2][2], "3")

    def test_email_audit_checks_name_product_unknown_and_blank_email(self):
        rounded = [*ALICE[:4], "2.08601E+13", *ALICE[5:]]
        rows = [rounded, [*rounded[:3], "Wrong Name", *rounded[4:]],
                [*rounded[:6], "LAZPAY"], ["unknown@example.test", *rounded[1:]], ["", *rounded[1:]]]
        result = run_prepared_check(self.prepare(cimb_blast(rows))).results[0]
        self.assertEqual([r[1] for r in result.records], ["DIFFERENT"] * 3 + ["NOT FOUND", "BLANK KEY"])
        self.assertTrue(any(d[0] == 3 and d[2] == "name → CUST_NAME" for d in result.differences))
        self.assertTrue(any(d[0] == 4 and d[2] == "PRODUCT → PLACEMENT" for d in result.differences))

    def test_even_scientific_ids_in_both_files_cannot_pass(self):
        rounded = [*ALICE[:4], "2.08601E+13", *ALICE[5:]]
        db_row = [DATABASE_ROWS[0][0], "2.08601E+13", *DATABASE_ROWS[0][2:]]
        result = run_prepared_check(self.prepare(cimb_blast([rounded]), cimb_database([db_row]))).results[0]
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertIn("full loan account number required", result.differences[0][4])

    def test_reordered_headers_preambles_and_equivalent_excel_are_detected(self):
        order = [6, 4, 3, 1, 0, 5, 2]
        file = xlsx_upload("renamed.xlsx", [["Export"], [], [BLAST_HEADERS[i].upper() for i in order],
                                              [ALICE[i] for i in order]])
        database = xlsx_upload("source.xlsx", [["Export"], [], DATABASE_HEADERS, DATABASE_ROWS[0]])
        result = run_prepared_check(self.prepare(file, database)).results[0]
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.records[0][0], 4)

    def test_missing_master_fields_and_empty_active_are_errors(self):
        for index in range(4):
            headers = [h for i, h in enumerate(DATABASE_HEADERS) if i != index]
            row = [v for i, v in enumerate(DATABASE_ROWS[0]) if i != index]
            database = xlsx_upload("incomplete.xlsx", [headers, row])
            result = run_uploaded_check(UploadedSource(database, "Sheet", headers[0]),
                                        [UploadedSource(cimb_blast(), "CSV", "primary_key")],
                                        Options(mapped_fields=True)).results[0]
            self.assertIn("CIMB email checks require", result.error)
        prepared = prepare_uploaded_check(cimb_database([], multiple_sheets=True), [cimb_blast()])
        self.assertIn("no nonblank values", prepared.database_error)

    def test_cli_uses_same_detection_and_preserves_inputs(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            uploads = [cimb_database(multiple_sheets=True), cimb_blast()]
            paths = [root / upload.name for upload in uploads]
            for path, upload in zip(paths, uploads):
                path.write_bytes(upload.data)
            output = root / "report.xlsx"
            stdout = io.StringIO()
            args = ["check_files.py", "--auto", "--database", str(paths[0]), "--files", str(paths[1]),
                    "--output", str(output)]
            with patch.object(sys, "argv", args), redirect_stdout(stdout):
                self.assertEqual(main(), 0)
            stats = json.loads(stdout.getvalue().splitlines()[0])
            self.assertEqual(stats["matched"], 1)
            self.assertEqual(stats["checked_fields"], 4)
            self.assertTrue(output.is_file())
            self.assertTrue(all(p.read_bytes() == u.data for p, u in zip(paths, uploads)))

    def test_common_campaign_headers_do_not_force_unrelated_databases_to_be_cimb(self):
        database = xlsx_upload("other source.xlsx", [
            ["CHCODE", "Name", "Email address", "PRODUCT"],
            [ALICE[4], ALICE[3], ALICE[0], "PERSONAL LOAN"],
        ])
        file = csv_upload("other campaign.csv", [["email", "name", "primary_key", "PRODUCT"],
                                                  [ALICE[0], ALICE[3], ALICE[4], "PERSONAL LOAN"]])
        result = run_prepared_check(self.prepare(file, database)).results[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertFalse(result.validation_notes)


if __name__ == "__main__":
    unittest.main()
