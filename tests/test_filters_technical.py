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

    def test_mark_technical_accounts_detects_requested_ecobion_and_secretariat_keywords(self) -> None:
        df = pd.DataFrame(
            [
                {"id": "u1", "firstname": "Dataprotection - Ecobion", "lastname": "Team"},
                {"id": "u2", "firstname": "Ecobion", "lastname": "Secrétariat"},
                {"id": "u3", "firstname": "Secretariat", "lastname": "Ecobion"},
                {"id": "u4", "firstname": "Desk", "lastname": "Sekretariat"},
                {"id": "u5", "firstname": "Segretariato", "lastname": "Hub"},
                {"id": "u6", "firstname": "Jane", "lastname": "Doe"},
            ]
        )

        result = mark_technical_accounts(
            df,
            {"dataprotection", "ecobion", "sekretariat", "secrétariat", "secretariat", "segretariato"},
        )

        self.assertTrue(result.loc[0, "flag_technical_account"])
        self.assertIn("token_firstname:dataprotection", result.loc[0, "flag_technical_reason"])

        self.assertTrue(result.loc[1, "flag_technical_account"])
        self.assertIn("exact_firstname:ecobion", result.loc[1, "flag_technical_reason"])
        self.assertIn("exact_lastname:secrétariat", result.loc[1, "flag_technical_reason"])

        self.assertTrue(result.loc[2, "flag_technical_account"])
        self.assertIn("exact_firstname:secretariat", result.loc[2, "flag_technical_reason"])
        self.assertIn("exact_lastname:ecobion", result.loc[2, "flag_technical_reason"])

        self.assertTrue(result.loc[3, "flag_technical_account"])
        self.assertIn("exact_lastname:sekretariat", result.loc[3, "flag_technical_reason"])

        self.assertTrue(result.loc[4, "flag_technical_account"])
        self.assertIn("exact_firstname:segretariato", result.loc[4, "flag_technical_reason"])

        self.assertFalse(result.loc[5, "flag_technical_account"])
        self.assertEqual(result.loc[5, "flag_technical_reason"], "")

    def test_mark_technical_accounts_detects_accounting_firstname(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "u7",
                    "firstname": "Accounting",
                    "lastname": "CHDE",
                }
            ]
        )

        result = mark_technical_accounts(df, {"accounting"})

        self.assertTrue(result.loc[0, "flag_technical_account"])
        self.assertIn("exact_firstname:accounting", result.loc[0, "flag_technical_reason"])

    def test_mark_technical_accounts_detects_username_and_email_localpart_matches(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "u8",
                    "username": "bio.service@bioanalytica.ch",
                    "email": "anna.mueller@example.com",
                    "firstname": "Anna",
                    "lastname": "Mueller",
                },
                {
                    "id": "u9",
                    "username": "john.smith",
                    "email": "noreply@bioanalytica.ch",
                    "firstname": "John",
                    "lastname": "Smith",
                },
                {
                    "id": "u10",
                    "username": "jane.roe",
                    "email": "jane@patholab.ch",
                    "firstname": "Jane",
                    "lastname": "Roe",
                },
            ]
        )

        result = mark_technical_accounts(df, {"service", "noreply", "lab"})

        self.assertTrue(result.loc[0, "flag_technical_account"])
        self.assertIn("token_username:service", result.loc[0, "flag_technical_reason"])

        self.assertTrue(result.loc[1, "flag_technical_account"])
        self.assertIn("exact_email:noreply", result.loc[1, "flag_technical_reason"])

        self.assertFalse(result.loc[2, "flag_technical_account"])
        self.assertEqual(result.loc[2, "flag_technical_reason"], "")

    def test_mark_technical_accounts_detects_recent_function_account_examples(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "u11",
                    "username": "dia.cml@dianalabs.ch",
                    "email": "dia.cml@dianalabs.ch",
                    "firstname": "Dia",
                    "lastname": "cml",
                },
                {
                    "id": "u12",
                    "username": "dia.bacterio@dianalabs.ch",
                    "email": "dia.bacterio@dianalabs.ch",
                    "firstname": "dia",
                    "lastname": "Bacterio",
                },
                {
                    "id": "u13",
                    "username": "frigoalarm@dianalabs.ch",
                    "email": "frigoalarm@dianalabs.ch",
                    "firstname": "dia",
                    "lastname": "frigoalarm",
                },
                {
                    "id": "u14",
                    "username": "testticket@sonicsuisse.ch",
                    "email": "testticket@sonicsuisse.ch",
                    "firstname": "Forward",
                    "lastname": "Servicedesk",
                },
                {
                    "id": "u15",
                    "username": "tbmicrobio@dianalabs.ch",
                    "email": "tbmicrobio@dianalabs.ch",
                    "firstname": "Dianalabs",
                    "lastname": "Microbio",
                },
                {
                    "id": "u16",
                    "username": "dianalabs.champeltest@dianalabs.ch",
                    "email": "dianalabs.champeltest@dianalabs.ch",
                    "firstname": "Dianalabs",
                    "lastname": "champeltest",
                },
                {
                    "id": "u17",
                    "username": "dialabo@dianalabs.ch",
                    "email": "dialabo@dianalabs.ch",
                    "firstname": "Dianalabs",
                    "lastname": "Colline",
                },
            ]
        )

        result = mark_technical_accounts(
            df,
            {"cml", "bacterio", "frigoalarm", "servicedesk", "microbio", "champeltest", "dialabo"},
        )

        self.assertIn("exact_lastname:cml", result.loc[0, "flag_technical_reason"])
        self.assertIn("exact_lastname:bacterio", result.loc[1, "flag_technical_reason"])
        self.assertIn("exact_lastname:frigoalarm", result.loc[2, "flag_technical_reason"])
        self.assertIn("exact_lastname:servicedesk", result.loc[3, "flag_technical_reason"])
        self.assertIn("exact_lastname:microbio", result.loc[4, "flag_technical_reason"])
        self.assertIn("exact_lastname:champeltest", result.loc[5, "flag_technical_reason"])
        self.assertIn("exact_username:dialabo", result.loc[6, "flag_technical_reason"])


if __name__ == "__main__":
    unittest.main()
