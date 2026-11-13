"""Exclude blank database lookup IDs without hiding campaign row issues."""

import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from checker import Options, Source, check_files
from scripts.check_files import main
from test_inferred_matching import campaign_pair
from test_sbf_pl_viber import BLAST_HEADERS as PL_HEADERS, pl_database, pl_row
from test_sbf_salad_viber import BLAST_HEADERS as SALAD_HEADERS, salad_row
from test_web_checker import csv_upload, xlsx_upload
from web_checker import prepare_uploaded_check, run_prepared_check, run_prepared_checks


GRID_HEADERS = ["Name", "Old I.C.", "Account No.", "Email"]
EMAIL_HEADERS = ["email", "primary_key", "chname", "accountnumber"]


def blank_key_pair():
    database = xlsx_upload("DataGrid.xlsx", [
        GRID_HEADERS,
        ["Alice Smith", "00123", "000001", "alice@example.test"],
        ["Unkeyed customer", None, "999999", "unkeyed@example.test"],
        ["Bob Jones", "00456", "000002", None],
        ["Empty ID", "", "999998", None],
        ["Whitespace ID", " \t ", "999997", None],
    ])
    blast = csv_upload("blast.csv", [
        EMAIL_HEADERS,
        ["alice@example.test", "00123", "Alice Smith", "000001"],
        ["", "00456", "Bob Jones", "000002"],
        ["unkeyed@example.test", "", "Unkeyed customer", "999999"],
        ["unkeyed@example.test", "missing", "Unkeyed customer", "999999"],
    ])
    return database, blast


class BlankDatabaseKeyTests(unittest.TestCase):
    def prepare(self, database, blast):
        prepared = prepare_uploaded_check(database, [blast])
        self.assertEqual(prepared.database_error, "")
        self.assertEqual(prepared.files[0].error, "")
        return prepared

    def test_datagrid_skips_blank_ids_and_keeps_non_key_blanks_and_campaign_issues(self):
        database, blast = blank_key_pair()
        run = run_prepared_check(self.prepare(database, blast))
        result = run.results[0]
        self.assertEqual(result.error, "")
        self.assertEqual([r[1] for r in result.records], ["MATCHED", "MATCHED", "BLANK KEY", "NOT FOUND"])
        self.assertEqual([r[2] for r in result.records[:2]], ["2", "4"])
        self.assertEqual(result.skipped_database_rows, [3, 5, 6])
        book = load_workbook(io.BytesIO(run.report), read_only=True, data_only=True)
        try:
            self.assertEqual(tuple(book["Skipped DB 001"].values)[1:], (
                (3, "Old I.C.", "Blank lookup value; excluded from comparison"),
                (5, "Old I.C.", "Blank lookup value; excluded from comparison"),
                (6, "Old I.C.", "Blank lookup value; excluded from comparison"),
            ))
            rows = list(book["Summary"].values)
            self.assertEqual(rows[8][rows[7].index("Skipped database rows")], 3)
            self.assertEqual(book["File 001"]["E2"].value, "00123")
        finally:
            book.close()

    def test_inference_accepts_unique_populated_ids_without_aligning_blank_ids(self):
        database = xlsx_upload("generic.xlsx", [
            ["Customer ID", "Name"], ["001", "Alice"], [None, "Unkeyed"],
            ["002", "Bob"], ["003", "Carol"], ["", "Another unkeyed"],
        ])
        blast = csv_upload("generic.csv", [
            ["Client ID", "Name"], ["001", "Alice"], ["002", "Bob"], ["003", "Carol"], ["", "Unkeyed"],
        ])
        prepared = self.prepare(database, blast)
        self.assertTrue(prepared.files[0].inferred)
        result = run_prepared_check(prepared).results[0]
        self.assertEqual([r[1] for r in result.records], ["MATCHED", "MATCHED", "MATCHED", "BLANK KEY"])
        self.assertEqual(result.skipped_database_rows, [3, 6])

    def test_inferred_per_file_key_excludes_blank_ids_on_its_own_column(self):
        database, blast, _, _ = campaign_pair()
        book = load_workbook(io.BytesIO(database.data))
        sheet = book.active
        sheet.append(["Unkeyed debtor", "OLD99999", "OTHER", "OTHER", "OTHER", None, None, None])
        stream = io.BytesIO()
        book.save(stream)
        book.close()
        result = run_prepared_check(self.prepare(type(database)(database.name, stream.getvalue()), blast)).results[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.database_match_column, "Debtor ID")
        self.assertEqual(result.count("MATCHED"), 5)
        self.assertEqual(result.skipped_database_rows, [7])

    def test_debtor_and_phone_profiles_skip_only_blank_effective_lookup_values(self):
        headers = [*GRID_HEADERS, "Last Called No.", "Debtor ID"]
        rows = [
            ["Alice Smith", "00123", "000001", None, "09171234567", 1001],
            ["No debtor or phone", "00456", "000002", None, None, None],
            ["Bob Jones", None, "000003", None, "09181234567", 1002],
        ]
        database = pl_database(rows, headers)
        blasts = [csv_upload("PL.csv", [PL_HEADERS, pl_row(), pl_row("1002", "Bob Jones", "09181234567")]),
                  csv_upload("SALAD.csv", [SALAD_HEADERS, salad_row(), salad_row("Bob Jones", "09181234567")])]
        prepared = prepare_uploaded_check(database, blasts)
        self.assertEqual(prepared.database_error, "")
        for result, key in zip(run_prepared_check(prepared).results, ["Debtor ID", "Last Called No."]):
            self.assertEqual(result.error, "")
            self.assertEqual(result.database_match_column, key)
            self.assertEqual(result.count("MATCHED"), 2)
            self.assertEqual(result.skipped_database_rows, [3])
            self.assertEqual(result.records[1][2], "4")

    def test_all_blank_ids_still_report_that_no_usable_records_remain(self):
        database = xlsx_upload("empty IDs.xlsx", [GRID_HEADERS, ["Unkeyed", None, "123", None]])
        _, blast = blank_key_pair()
        self.assertIn("no nonblank", prepare_uploaded_check(database, [blast]).database_error)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / database.name
            path.write_bytes(database.data)
            with self.assertRaisesRegex(ValueError, "no nonblank"):
                check_files(Source(path, "Sheet", "Old I.C."), [], Options(mapped_fields=True))

    def test_zero_identifier_is_valid_and_full_copy_keeps_missing_valid_rows(self):
        database = csv_upload("database.csv", [["CHCODE", "Name"], [0, "Zero"], ["", "Unkeyed"], ["2", "Two"]])
        target = csv_upload("copy.csv", [["CHCODE", "Name"], [0, "Zero"]])
        with tempfile.TemporaryDirectory() as folder:
            paths = [Path(folder) / item.name for item in (database, target)]
            for path, item in zip(paths, (database, target)):
                path.write_bytes(item.data)
            result = check_files(Source(paths[0], "CSV", "CHCODE"), [Source(paths[1], "CSV", "CHCODE")], Options())[0]
            self.assertEqual(result.count("MATCHED"), 1)
            self.assertEqual(result.skipped_database_rows, [3])
            self.assertEqual([n for n, _ in result.missing_rows], [4])

    def test_each_database_retains_its_own_exclusions_and_combined_verdict(self):
        database, blast = blank_key_pair()
        second = xlsx_upload("second.xlsx", [GRID_HEADERS,
                            ["No ID", None, None, None],
                            ["Alice Smith", "00123", "000001", "alice@example.test"],
                            ["Bob Jones", "00456", "000002", None]])
        run = run_prepared_checks([self.prepare(item, blast) for item in (database, second)])
        self.assertEqual([r.skipped_database_rows for r in run.results], [[3, 5, 6], [2]])
        self.assertEqual(run.matched_in_all, 2)

    def test_auto_and_explicit_cli_skip_blank_database_ids_without_modifying_inputs(self):
        database, blast = blank_key_pair()
        with tempfile.TemporaryDirectory() as folder:
            paths = [Path(folder) / item.name for item in (database, blast)]
            for path, item in zip(paths, (database, blast)):
                path.write_bytes(item.data)
            for flags in ([], ["--auto"]):
                stdout = io.StringIO()
                arguments = ["check_files.py", *flags, "--database", str(paths[0]), "--files", str(paths[1]),
                             "--output", str(Path(folder) / "result.xlsx")]
                with patch.object(sys, "argv", arguments), redirect_stdout(stdout):
                    self.assertEqual(main(), 0)
                result = json.loads(stdout.getvalue().splitlines()[0])
                self.assertEqual(result["skipped_database_rows"], 3)
                self.assertEqual(result["matched"], 2)
            self.assertTrue(all(path.read_bytes() == item.data for path, item in zip(paths, (database, blast))))


if __name__ == "__main__":
    unittest.main()
