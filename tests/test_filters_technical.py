from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from filters_technical import mark_technical_accounts


class MarkTechnicalAccountsTests(unittest.TestCase):
    def test_mark_technical_accounts_detects_exact_token_and_numeric_matches(self) -> None:
        df = pd.DataFrame(
            [
                {"id": "svc-app", "firstname": "Service", "lastname": "Account"},
                {"id": "u1", "firstname": "Max", "lastname": "Admin"},
                {"id": "u2", "firstname": "123", "lastname": "User"},
                {"id": "u3", "firstname": "Jane", "lastname": "Doe"},
            ]
        )

        result = mark_technical_accounts(df, {"svc-app", "admin"})

        self.assertTrue(result.loc[0, "flag_technical_account"])
        self.assertIn("exact_id:svc-app", result.loc[0, "flag_technical_reason"])

        self.assertTrue(result.loc[1, "flag_technical_account"])
        self.assertIn("exact_lastname:admin", result.loc[1, "flag_technical_reason"])

        self.assertTrue(result.loc[2, "flag_technical_account"])
        self.assertIn("numeric_firstname:123", result.loc[2, "flag_technical_reason"])

        self.assertFalse(result.loc[3, "flag_technical_account"])
        self.assertEqual(result.loc[3, "flag_technical_reason"], "")


if __name__ == "__main__":
    unittest.main()
