import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from checker import (
    ColumnMapping, Options, Source, check_files, is_cbs_apd_email_blast,
    match_column, read_table, suggest_mappings,
)


class CbsApdTests(unittest.TestCase):
    database_headers = ["Name", "Old I.C.", "Account No.", "Email", "Account Type"]
    campaign_headers = [
        "email", "cc", "bcc", "indexone_uid", "primary_key", "accountnumber", "chname", "campaignbucket",
    ]
    alice = ["Alice Smith", "00123", "0005678", "alice@example.test", "BUCKET 4 91-120"]
    campaign_alice = [
        "alice@example.test", "", "", "00123", "00123", "0005678", "Alice Smith", "BUCKET 4 91-120",
    ]

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.options = Options(mapped_fields=True, require_flow_stage=True)

    def csv(self, name, rows, delimiter=","):
        path = self.folder / name
        with path.open("w", newline="", encoding="utf-8-sig") as stream:
            csv.writer(stream, delimiter=delimiter).writerows(rows)
        return path

    def excel(self, name, rows):
        path = self.folder / name
        book = Workbook()
        for row in rows:
            book.active.append(row)
        book.save(path)
        book.close()
        return path

    def source(self, path):
        table = read_table(path)
        return Source(path, table.sheet, match_column(table.headers))

    def database(self, rows=None, headers=None):
        return self.source(self.excel("reference.xlsx", [
            headers or self.database_headers, *(rows if rows is not None else [self.alice]),
        ]))

    def campaign(self, rows=None, headers=None):
        return self.source(self.csv("upload.csv", [
            headers or self.campaign_headers, *(rows if rows is not None else [self.campaign_alice]),
        ]))

    def test_cbs_layout_selects_primary_key_and_preserves_other_schemas(self):
        self.assertTrue(is_cbs_apd_email_blast(self.campaign_headers))
        self.assertEqual(match_column(self.campaign_headers), "primary_key")
        self.assertEqual(match_column(["ChCode", *self.campaign_headers]), "ChCode")
        msb_headers = [name for name in self.campaign_headers if name != "campaignbucket"]
        self.assertFalse(is_cbs_apd_email_blast(msb_headers))
        self.assertEqual(match_column(msb_headers), "indexone_uid")
        self.assertFalse(is_cbs_apd_email_blast(["primary_key", "campaignbucket"]))
        self.assertEqual(match_column(["primary_key", "campaignbucket"]), "primary_key")

    def test_cbs_checks_six_fields_against_datagrid(self):
        campaign = self.campaign()
        result = check_files(self.database(), [campaign], self.options)[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertFalse(result.flow_stage_required)
        self.assertEqual(result.file_identifier, "upload.csv")
        self.assertEqual(result.ignored_columns, ["cc", "bcc"])
        self.assertEqual(result.mappings, [
            ColumnMapping("email", "Email"),
            ColumnMapping("indexone_uid", "Old I.C."),
            ColumnMapping("primary_key", "Old I.C."),
            ColumnMapping("accountnumber", "Account No."),
            ColumnMapping("chname", "Name"),
            ColumnMapping("campaignbucket", "Account Type"),
        ])

    def test_primary_key_controls_lookup_when_secondary_uid_disagrees(self):
        values = self.campaign_alice.copy()
        values[3] = "00456"
        database = self.database([self.alice, ["Bob Jones", "00456", "999", "bob@example.test", "OTHER"]])
        result = check_files(database, [self.campaign([values])], self.options)[0]
        self.assertEqual(result.source.column, "primary_key")
        self.assertEqual(result.count("DIFFERENT"), 1)
        self.assertEqual(result.records[0][2], "2")
        self.assertEqual(result.differences, [(2, 2, "indexone_uid → Old I.C.", "00456", "00123")])

    def test_primary_key_missing_and_blank_do_not_fall_back_to_secondary_uid(self):
        rows = []
        for primary_key in ["99999", ""]:
            values = self.campaign_alice.copy()
            values[4] = primary_key
            rows.append(values)
        result = check_files(self.database(), [self.campaign(rows)], self.options)[0]
        self.assertEqual([row[1] for row in result.records], ["NOT FOUND", "BLANK KEY"])
        self.assertEqual(result.verdict, "ROW ISSUES")

    def test_email_and_bucket_differences_remain_visible(self):
        for position, value, label, expected in [
            (0, "wrong@example.test", "email → Email", "alice@example.test"),
            (7, "BUCKET 3 61-90", "campaignbucket → Account Type", "BUCKET 4 91-120"),
            (5, "5678", "accountnumber → Account No.", "0005678"),
        ]:
            with self.subTest(field=label):
                values = self.campaign_alice.copy()
                values[position] = value
                result = check_files(self.database(), [self.campaign([values])], self.options)[0]
                self.assertEqual(result.count("DIFFERENT"), 1)
                self.assertEqual(result.differences, [(2, 2, label, value, expected)])

    def test_database_can_contain_more_rows_than_the_campaign(self):
        rows = [self.alice, *[
            [f"Customer {index}", f"K{index:06}", f"A{index:08}", f"customer{index}@example.test", "OTHER"]
            for index in range(1104)
        ]]
        result = check_files(self.database(rows), [self.campaign()], self.options)[0]
        self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.missing_rows, [])

    def test_case_preambles_and_text_keys_are_preserved(self):
        headers = [f" {name.upper()} " for name in self.campaign_headers]
        paths = [
            self.csv("preamble.csv", [["Export"], [], headers, self.campaign_alice], delimiter=";"),
            self.excel("preamble.xlsx", [["Export"], [], headers, self.campaign_alice]),
        ]
        for path in paths:
            with self.subTest(file=path.name):
                table = read_table(path)
                self.assertEqual(table.header_row, 3)
                self.assertEqual(match_column(table.headers), "PRIMARY_KEY")
                self.assertEqual(table.rows[0][1][4:6], ("00123", "0005678"))
                result = check_files(self.database(), [self.source(path)], self.options)[0]
                self.assertEqual(result.error, "")
                self.assertEqual(result.count("MATCHED"), 1)
                self.assertIn(ColumnMapping("CAMPAIGNBUCKET", "Account Type"), result.mappings)

    def test_bucket_mapping_uses_account_type_before_same_named_unrelated_field(self):
        database_headers = [*self.database_headers, "campaignbucket"]
        mappings = suggest_mappings(self.campaign_headers, database_headers, "primary_key", "Old I.C.")
        self.assertIn(ColumnMapping("campaignbucket", "Account Type"), mappings)
        self.assertNotIn(ColumnMapping("campaignbucket", "campaignbucket"), mappings)

    def test_bucket_without_database_counterpart_is_disclosed_as_unchecked(self):
        database = self.database([self.alice[:-1]], headers=self.database_headers[:-1])
        result = check_files(database, [self.campaign()], self.options)[0]
        self.assertEqual(result.error, "")
        self.assertEqual(result.count("MATCHED"), 1)
        self.assertEqual(result.ignored_columns, ["cc", "bcc", "campaignbucket"])

    def test_selected_primary_key_is_required_in_custom_mapping(self):
        campaign = self.campaign()
        campaign = Source(campaign.path, campaign.sheet, campaign.column, (ColumnMapping("indexone_uid", "Old I.C."),))
        result = check_files(self.database(), [campaign], self.options)[0]
        self.assertIn("primary_key-to-Old I.C. mapping is required", result.error)

    def test_cbs_with_chcode_still_requires_flow_stage(self):
        campaign = self.campaign([[*self.campaign_alice, "00123"]], headers=[*self.campaign_headers, "CHCODE"])
        result = check_files(self.database(), [campaign], self.options)[0]
        self.assertEqual(campaign.column, "CHCODE")
        self.assertIn("FLOW_STAGE header", result.error)

    def test_present_flow_stage_is_validated_for_cbs(self):
        campaign = self.campaign([[*self.campaign_alice, ""]], headers=[*self.campaign_headers, "FLOW_STAGE"])
        result = check_files(self.database(), [campaign], self.options)[0]
        self.assertTrue(result.flow_stage_required)
        self.assertIn("FLOW_STAGE is blank", result.error)


if __name__ == "__main__":
    unittest.main()
