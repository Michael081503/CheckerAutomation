"""Automatic command-line checks share the browser detection and report path."""

import csv
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from scripts.check_files import main
from web_checker import run_prepared_check


class AutomaticCommandLineTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.database = self.csv("database.csv", [
            ["CHCODE", "Name", "Mobile phone"],
            ["00123", "Alice Smith", "09171234567"],
            ["00456", "Bob Jones", "09181234567"],
        ])
        self.target = self.csv("template.csv", [
            ["CHCODE", "Name", "Phone"],
            ["00123", "Alice Smith", "+639171234567"],
            ["00456", "Changed Name", "09181234567"],
            ["missing", "Unknown Customer", "09191234567"],
            ["", "Blank Customer", "09201234567"],
        ])
        self.output = self.folder / "reports" / "result.xlsx"

    def csv(self, name, rows):
        path = self.folder / name
        with path.open("w", newline="", encoding="utf-8-sig") as stream:
            csv.writer(stream).writerows(rows)
        return path

    def run_cli(self, *extra, files=None, output=None):
        arguments = ["check_files.py", "--auto", "--database", str(self.database),
                     "--files", *map(str, files or [self.target]),
                     "--output", str(output or self.output), *extra]
        stdout = io.StringIO()
        with patch.object(sys, "argv", arguments), redirect_stdout(stdout):
            exit_code = main()
        results = [json.loads(line) for line in stdout.getvalue().splitlines() if line.startswith("{")]
        return exit_code, results

    def test_auto_uses_filename_without_stage_and_preserves_all_row_issues(self):
        originals = {path: path.read_bytes() for path in [self.database, self.target]}
        exit_code, results = self.run_cli()
        self.assertEqual(exit_code, 0)
        result = results[0]
        self.assertEqual(result["identifier"], "template.csv")
        self.assertEqual(result["FLOW_STAGE"], "")
        self.assertEqual((result["file_key"], result["database_key"]), ("CHCODE", "CHCODE"))
        self.assertEqual((result["rows"], result["matched"], result["changed"],
                          result["not_found"], result["blank_key"]), (4, 1, 1, 1, 1))
        self.assertEqual(result["checked_fields"], 3)
        self.assertEqual(result["error"], "")
        book = load_workbook(self.output, read_only=True, data_only=True)
        try:
            records = list(book["File 001"].values)[1:]
            self.assertEqual([row[:2] for row in records], [
                (2, "MATCHED"), (3, "DIFFERENT"), (4, "NOT FOUND"), (5, "BLANK KEY"),
            ])
            self.assertEqual(records[0][3], "00123")
            self.assertEqual(book["File 001"]["D2"].data_type, "s")
        finally:
            book.close()
        self.assertTrue(all(path.read_bytes() == original for path, original in originals.items()))

    def test_auto_saves_the_adapter_report_without_running_a_second_check(self):
        captured_reports = []

        def capture(*args, **kwargs):
            run = run_prepared_check(*args, **kwargs)
            captured_reports.append(run.report)
            return run

        with patch("web_checker.run_prepared_check", side_effect=capture) as run, \
                patch("scripts.check_files.check_files", side_effect=AssertionError("duplicate check")), \
                patch("scripts.check_files.export_report", side_effect=AssertionError("duplicate report")):
            self.run_cli()
        self.assertEqual(run.call_count, 1)
        self.assertEqual(self.output.read_bytes(), captured_reports[0])

    def test_auto_bad_file_remains_in_report_while_other_files_are_checked(self):
        broken = self.folder / "broken.xlsx"
        broken.write_bytes(b"not an Excel workbook")
        exit_code, results = self.run_cli(files=[broken, self.target])
        self.assertEqual(exit_code, 1)
        self.assertEqual(results[0]["verdict"], "ERROR")
        self.assertTrue(results[0]["error"])
        self.assertEqual(results[1]["rows"], 4)
        self.assertTrue(self.output.is_file())

    def test_auto_rejects_manual_options_before_writing_a_report(self):
        for extra in (("--full-copy",), ("--database-sheet", "Accounts"),
                      ("--database-key", "CHCODE"), ("--file-key", "CHCODE")):
            with self.subTest(option=extra[0]), redirect_stderr(io.StringIO()), \
                    self.assertRaises(SystemExit) as stopped:
                self.run_cli(*extra)
            self.assertEqual(stopped.exception.code, 2)
            self.assertFalse(self.output.exists())

    def test_auto_output_cannot_overwrite_either_input(self):
        for path in (self.database, self.target):
            original = path.read_bytes()
            with self.subTest(path=path.name), redirect_stderr(io.StringIO()), \
                    self.assertRaises(SystemExit) as stopped:
                self.run_cli(output=path)
            self.assertEqual(stopped.exception.code, 2)
            self.assertEqual(path.read_bytes(), original)

    def test_auto_invalid_database_is_a_readable_error_without_output(self):
        self.database.write_text("Unrelated\nNothing\n", encoding="utf-8")
        stderr = io.StringIO()
        with redirect_stderr(stderr), self.assertRaises(SystemExit) as stopped:
            self.run_cli()
        self.assertEqual(stopped.exception.code, 2)
        self.assertIn("Database:", stderr.getvalue())
        self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
