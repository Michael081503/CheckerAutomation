"""Guard identifier inference against plausible but unrelated values."""

import unittest
from pathlib import Path

from auto_match import database_key_candidates, infer_layout
from checker import Table


def table(headers, rows):
    return Table(Path("synthetic.csv"), "CSV", headers,
                 [(index, tuple(row)) for index, row in enumerate(rows, 2)])


class IdentifierSafetyTests(unittest.TestCase):
    def test_names_categories_staff_and_sequences_are_not_database_keys(self):
        headers = ["Account Name", "Account Type", "Account Age", "No", "Staff ID", "Customer ID"]
        database = table(headers, [[f"Person {n}", f"Type {n}", n, n, f"EMP{n}", f"C{n:03}"]
                                   for n in range(1, 6)])
        self.assertEqual(database_key_candidates(database), ["Customer ID"])

    def test_sequence_or_name_values_cannot_supply_an_uploaded_identity(self):
        database = table(["Customer ID", "Name"], [[str(n), f"Person {n}"] for n in range(1, 6)])
        for header in ("No", "Row number", "Account Name", "Staff ID", "Account Type"):
            with self.subTest(header=header):
                target = table([header, "Status"], [[str(n), "Open"] for n in range(1, 6)])
                with self.assertRaisesRegex(ValueError, "No reliable identifier"):
                    infer_layout(target, database)

    def test_global_value_overlap_does_not_establish_a_secondary_field_mapping(self):
        database = table(["Customer ID", "External note"], [[f"C{n}", f"NOTE{n}"] for n in range(1, 6)])
        target = table(["Client ID", "Delivery label"],
                       [[f"C{n}", f"NOTE{n % 5 + 1}"] for n in range(1, 6)])
        layout = infer_layout(target, database)
        self.assertEqual([mapping.file_column for mapping in layout.mappings], ["Client ID"])
        self.assertIn("Only the identifier", layout.note)

    def test_sparse_secondary_intersection_is_not_sufficient_mapping_evidence(self):
        database = table(["Customer ID", "External note"],
                         [[f"C{n}", f"NOTE{n}"] for n in range(1, 6)])
        target = table(["Client ID", "Delivery label"],
                       [[f"C{n}", f"NOTE{n}" if n < 4 else ""] for n in range(1, 6)])
        layout = infer_layout(target, database)
        self.assertEqual(len(layout.mappings), 1)


if __name__ == "__main__":
    unittest.main()
