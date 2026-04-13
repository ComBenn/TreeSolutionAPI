from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from filters_employee_list import mark_by_employee_list


class MarkByEmployeeListTests(unittest.TestCase):
    def test_mark_by_employee_list_matches_email_and_name_variants(self) -> None:
        df_users = pd.DataFrame(
            [
                {
                    "id": "1",
                    "email": "john.doe@example.com",
                    "firstname": "John",
                    "lastname": "Doe",
                },
                {
                    "id": "2",
                    "email": "jane.roe@example.com",
                    "firstname": "Jane",
                    "lastname": "Roe",
                },
                {
                    "id": "3",
                    "email": "nomatch@example.com",
                    "firstname": "No",
                    "lastname": "Match",
                },
            ]
        )
        df_employee_list = pd.DataFrame(
            [
                {"email": "john.doe@example.com", "Vorname": "", "Nachname": ""},
                {"email": "", "Vorname": "Jane", "Nachname": "Roe"},
            ]
        )

        result, stats = mark_by_employee_list(df_users, df_employee_list, return_stats=True)

        self.assertTrue(result.loc[0, "flag_employee_list"])
        self.assertTrue(result.loc[1, "flag_employee_list"])
        self.assertFalse(result.loc[2, "flag_employee_list"])
        self.assertIn("match_email", result.loc[0, "flag_employee_list_reason"])
        self.assertIn("match_fullname", result.loc[1, "flag_employee_list_reason"])
        self.assertEqual(stats["employee_entries_total"], 7)
        self.assertEqual(stats["employee_entries_matched"], 7)
        self.assertEqual(stats["employee_entries_unmatched"], 0)

    def test_mark_by_employee_list_raises_for_unsupported_columns(self) -> None:
        df_users = pd.DataFrame([{"email": "a@example.com", "firstname": "A", "lastname": "User"}])
        df_employee_list = pd.DataFrame([{"foo": "bar"}])

        with self.assertRaises(ValueError):
            mark_by_employee_list(df_users, df_employee_list)


if __name__ == "__main__":
    unittest.main()
