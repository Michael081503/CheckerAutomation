import io
import unittest

from openpyxl import load_workbook

from test_web_checker import csv_upload, xlsx_upload
from web_checker import prepare_uploaded_check, run_prepared_check


DATABASE_HEADERS = [
    "Name", "Old I.C.", "Account No.", "Financing/Card No.", "New I.C.",
    "Email", "Last Called No.", "Debtor ID",
]
CAMPAIGN_HEADERS = [
    "FullName", "FirstName", "LastName", "Phone", "CHCode", "OB",
    "Monthly", "Terms", "Promo",
]


def campaign_pair(profile="PARKED", *, count=5, change=None):
    """Model misleading campaign headers without including customer data."""
    database_rows = []
    file_rows = []
    for number in range(1, count + 1):
        name = f"Example Customer {number:02d}"
        account = f"SPL15R{number:05d}"
        phone = f"917{number:07d}"
        debtor = str(9800000 + number)
        database_rows.append([
            name, f"OLD{number:05d}", account, account, account,
            phone, f"918{number:07d}", int(debtor),
        ])
        values = [
            name, f"Staff {number % 4}", account, phone, f"AGENT{number % 4}",
            "", "", "", account[-4:],
        ]
        values.extend(["Level_5", debtor] if profile == "PARKED" else [debtor, "Level_5"])
        file_rows.append(values)
    headers = CAMPAIGN_HEADERS + (["Level", ""] if profile == "PARKED" else ["Debtor", "Level"])
    if change is not None:
        change(headers, file_rows)
    return (
        xlsx_upload(f"{profile}.xlsx", [DATABASE_HEADERS, *database_rows]),
        csv_upload(f"{profile} TEMPLATE.csv", [headers, *file_rows]),
        headers,
        file_rows,
    )


class InferredMatchingTests(unittest.TestCase):
    def prepare_pair(self, profile="PARKED", *, change=None):
        database, upload, headers, rows = campaign_pair(profile, change=change)
        prepared = prepare_uploaded_check(database, [upload])
        self.assertFalse(prepared.database_error)
        self.assertEqual(len(prepared.files), 1)
        item = prepared.files[0]
        self.assertFalse(item.error)
        self.assertTrue(item.inferred)
        self.assertEqual(item.source.database_column, "Debtor ID")
        self.assertEqual(item.source.column, "Unnamed column K" if profile == "PARKED" else "Debtor")
        return prepared, headers, rows

    def test_parked_and_legacy_detect_actual_fields_despite_misleading_headers(self):
        for profile in ("PARKED", "LEGACY"):
            with self.subTest(profile=profile):
                prepared, headers, rows = self.prepare_pair(profile)
                item = prepared.files[0]
                expected = {
                    "FullName": ("Name", "Text"),
                    "LastName": ("Account No.", "Text"),
                    "Phone": ("Email", "Phone"),
                    item.source.column: ("Debtor ID", "Text"),
                }
                self.assertEqual(
                    {mapping.file_column: (mapping.database_column, mapping.comparison)
                     for mapping in item.source.mappings},
                    expected,
                )
                self.assertEqual(len(item.mapping_evidence), 4)
                for evidence in item.mapping_evidence:
                    self.assertTrue({"File column", "Database column", "Comparison", "Evidence"}.issubset(evidence))
                    self.assertTrue(evidence["Evidence"])
                run = run_prepared_check(prepared)
                result = run.results[0]
                self.assertFalse(result.error)
                self.assertEqual(result.database_match_column, "Debtor ID")
                self.assertEqual(result.count("MATCHED"), 5)
                self.assertEqual(result.verdict, "MAPPED ROWS MATCH")
                self.assertEqual(result.ignored_columns, [
                    "FirstName", "CHCode", "OB", "Monthly", "Terms", "Promo", "Level",
                ])

    def test_changes_are_compared_on_the_same_debtor_record(self):
        for profile in ("PARKED", "LEGACY"):
            for field in ("Phone", "FullName", "LastName"):
                with self.subTest(profile=profile, field=field):
                    def change(headers, rows):
                        position = headers.index(field)
                        # This value exists in the database, on a different debtor.
                        rows[0][position] = rows[1][position]

                    prepared, headers, rows = self.prepare_pair(profile, change=change)
                    result = run_prepared_check(prepared).results[0]
                    self.assertFalse(result.error)
                    self.assertEqual(result.count("MATCHED"), 4)
                    self.assertEqual(result.count("DIFFERENT"), 1)
                    self.assertEqual(result.records[0][1], "DIFFERENT")
                    expected_label = f"{field} \u2192 " + {
                        "Phone": "Email", "FullName": "Name", "LastName": "Account No.",
                    }[field]
                    self.assertTrue(any(difference[:3] == (2, 2, expected_label)
                                        for difference in result.differences))

    def test_unknown_and_blank_debtor_keys_do_not_switch_to_account_or_phone(self):
        for profile in ("PARKED", "LEGACY"):
            for key, expected_status in (("999999999", "NOT FOUND"), ("", "BLANK KEY")):
                with self.subTest(profile=profile, status=expected_status):
                    def change(headers, rows):
                        rows[0][10 if profile == "PARKED" else 9] = key

                    prepared, headers, rows = self.prepare_pair(profile, change=change)
                    result = run_prepared_check(prepared).results[0]
                    self.assertFalse(result.error)
                    self.assertEqual(result.records[0][1], expected_status)
                    self.assertEqual(result.count("MATCHED"), 4)
                    self.assertEqual(result.count(expected_status), 1)

    def test_profile_phone_mapping_is_kept_when_all_phone_values_change_or_are_blank(self):
        for profile in ("PARKED", "LEGACY"):
            for replacement in ("other_database_phone", "blank"):
                with self.subTest(profile=profile, replacement=replacement):
                    def change(headers, rows):
                        for number, row in enumerate(rows, 1):
                            row[headers.index("Phone")] = (f"918{number:07d}"
                                                           if replacement == "other_database_phone" else "")

                    prepared, headers, rows = self.prepare_pair(profile, change=change)
                    phone_mapping = next(mapping for mapping in prepared.files[0].source.mappings
                                         if mapping.file_column == "Phone")
                    self.assertEqual((phone_mapping.database_column, phone_mapping.comparison), ("Email", "Phone"))
                    result = run_prepared_check(prepared).results[0]
                    self.assertFalse(result.error)
                    self.assertEqual(result.count("MATCHED"), 0)
                    self.assertEqual(result.count("DIFFERENT"), 5)
                    self.assertEqual(len([difference for difference in result.differences
                                          if difference[2] == "Phone \u2192 Email"]), 5)

    def test_entirely_blank_or_unknown_debtor_column_keeps_the_selected_lookup(self):
        for profile in ("PARKED", "LEGACY"):
            for expected_status in ("BLANK KEY", "NOT FOUND"):
                with self.subTest(profile=profile, status=expected_status):
                    def change(headers, rows):
                        key_position = 10 if profile == "PARKED" else 9
                        for number, row in enumerate(rows, 1):
                            row[key_position] = "" if expected_status == "BLANK KEY" else str(9990000 + number)

                    prepared, headers, rows = self.prepare_pair(profile, change=change)
                    item = prepared.files[0]
                    self.assertEqual(
                        {mapping.file_column: (mapping.database_column, mapping.comparison)
                         for mapping in item.source.mappings},
                        {
                            "FullName": ("Name", "Text"),
                            "LastName": ("Account No.", "Text"),
                            "Phone": ("Email", "Phone"),
                            item.source.column: ("Debtor ID", "Text"),
                        },
                    )
                    result = run_prepared_check(prepared).results[0]
                    self.assertFalse(result.error)
                    self.assertEqual(result.source.column, item.source.column)
                    self.assertEqual(result.database_match_column, "Debtor ID")
                    self.assertEqual(len(result.headers), len(headers))
                    self.assertEqual(len(result.records), 5)
                    self.assertEqual(result.count(expected_status), 5)
                    self.assertEqual(result.count("MATCHED"), 0)
                    self.assertTrue(all(record[1] == expected_status for record in result.records))

    def test_parked_report_preserves_all_original_values_and_the_unnamed_column(self):
        prepared, headers, original_rows = self.prepare_pair()
        run = run_prepared_check(prepared)
        book = load_workbook(io.BytesIO(run.report), data_only=True)
        try:
            exported = list(book["File 001"].values)
            self.assertEqual(exported[0][3:], tuple(headers[:-1] + ["Unnamed column K"]))
            self.assertEqual(len(exported), len(original_rows) + 1)
            for number, (actual, original) in enumerate(zip(exported[1:], original_rows), 2):
                self.assertEqual(actual[0], number)
                self.assertEqual(actual[1], "MATCHED")
                self.assertEqual(actual[3:], tuple(value if value != "" else None for value in original))
            mapping_rows = list(book["Mappings 001"].values)
            mapping_by_name = {row[0]: row for row in mapping_rows[1:]}
            self.assertEqual(mapping_by_name["Unnamed column K"][1:4], ("Debtor ID", "Text", "CHECKED"))
            self.assertEqual(mapping_by_name["Phone"][1:4], ("Email", "Phone", "CHECKED"))
            self.assertEqual(mapping_by_name["FirstName"][3], "NOT CHECKED")
        finally:
            book.close()

    def test_new_customer_id_layout_is_inferred_from_headers_and_values(self):
        database_rows = [
            [f"CUSTOMER{number:03d}", f"Example Customer {number}", f"0917{number:07d}"]
            for number in range(1, 6)
        ]
        file_rows = [[key, name, "+63" + phone[1:]] for key, name, phone in database_rows]
        prepared = prepare_uploaded_check(
            xlsx_upload("new database.xlsx", [["Customer ID", "Customer Name", "Mobile"], *database_rows]),
            [csv_upload("new export.csv", [["Client ID", "Customer Name", "Mobile"], *file_rows])],
        )
        self.assertFalse(prepared.database_error)
        self.assertFalse(prepared.files[0].error)
        self.assertTrue(prepared.files[0].inferred)
        self.assertEqual(prepared.files[0].source.column, "Client ID")
        self.assertEqual(prepared.files[0].source.database_column, "Customer ID")
        result = run_prepared_check(prepared).results[0]
        self.assertFalse(result.error)
        self.assertEqual(result.count("MATCHED"), 5)
        self.assertEqual({mapping.file_column for mapping in result.mappings},
                         {"Client ID", "Customer Name", "Mobile"})

    def test_ambiguous_identifier_overlap_is_rejected(self):
        # Both identifiers contain the entire uploaded set but identify different
        # records. The constant Status cannot corroborate either interpretation.
        keys = [f"KEY{number:03d}" for number in range(1, 6)]
        prepared = prepare_uploaded_check(
            xlsx_upload("ambiguous.xlsx", [
                ["Customer ID", "Account ID", "Status"],
                *[[key, keys[(number + 1) % len(keys)], "ACTIVE"]
                  for number, key in enumerate(keys)],
            ]),
            [csv_upload("ambiguous export.csv", [
                ["External identifier", "Status"], *[[key, "ACTIVE"] for key in keys],
            ])],
        )
        error = prepared.database_error or (prepared.files[0].error if prepared.files else "")
        self.assertRegex(error, "(?i)ambiguous")

    def test_low_cardinality_values_alone_do_not_select_a_lookup(self):
        prepared = prepare_uploaded_check(
            xlsx_upload("customers.xlsx", [
                ["Customer ID", "Status", "Segment"],
                *[[f"CUSTOMER{number:03d}", "ACTIVE", "A"] for number in range(1, 6)],
            ]),
            [csv_upload("status export.csv", [
                ["Status", "Segment"], *[["ACTIVE", "A"] for _ in range(5)],
            ])],
        )
        self.assertTrue(prepared.database_error or (prepared.files and prepared.files[0].error))


if __name__ == "__main__":
    unittest.main()
