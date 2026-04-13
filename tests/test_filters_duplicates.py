from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from filters_duplicates import mark_duplicate_accounts


class MarkDuplicateAccountsTests(unittest.TestCase):
    def test_mark_duplicate_accounts_groups_email_username_and_name_matches_transitively(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "1",
                    "username": "alpha",
                    "email": "alpha@example.com",
                    "firstname": "Alice",
                    "lastname": "Doe",
                },
                {
                    "id": "2",
                    "username": "bravo",
                    "email": "alpha@example.com",
                    "firstname": "Alicia",
                    "lastname": "Miller",
                },
                {
                    "id": "3",
                    "username": "charlie",
                    "email": "charlie@example.com",
                    "firstname": "Alice",
                    "lastname": "Doe",
                },
                {
                    "id": "",
                    "username": "charlie",
                    "email": "other@example.com",
                    "firstname": "Bob",
                    "lastname": "Brown",
                },
                {
                    "id": "5",
                    "username": "unique",
                    "email": "unique@example.com",
                    "firstname": "Unique",
                    "lastname": "Person",
                },
            ]
        )

        result = mark_duplicate_accounts(df)

        first_group = result.loc[0, "flag_duplicate_group"]
        self.assertTrue(result.loc[0, "flag_duplicate"])
        self.assertEqual(first_group, result.loc[1, "flag_duplicate_group"])
        self.assertEqual(first_group, result.loc[2, "flag_duplicate_group"])
        self.assertIn("duplicate_email", result.loc[0, "flag_duplicate_reason"])
        self.assertIn("duplicate_name", result.loc[0, "flag_duplicate_reason"])

        self.assertTrue(result.loc[3, "flag_duplicate"])
        self.assertEqual(first_group, result.loc[3, "flag_duplicate_group"])
        self.assertIn("duplicate_username", result.loc[3, "flag_duplicate_reason"])
        self.assertFalse(result.loc[3, "flag_duplicate_keep_candidate"])

        self.assertFalse(result.loc[4, "flag_duplicate"])
        self.assertEqual(result.loc[4, "flag_duplicate_group"], "")
        self.assertTrue(result.loc[4, "flag_duplicate_keep_candidate"])


if __name__ == "__main__":
    unittest.main()
