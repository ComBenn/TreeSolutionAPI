from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from exporter import build_upload_export


class BuildUploadExportTests(unittest.TestCase):
    def test_build_upload_export_creates_numbered_departments_and_removes_helper_columns(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "1",
                    "username": "jdoe",
                    "email": "jdoe@example.com",
                    "firstname": "John",
                    "lastname": "Doe",
                    "department": "Legacy",
                    "flag_duplicate": True,
                    "__batch_id": "1",
                }
            ]
        )

        result = build_upload_export(df, department_overrides=["IT", "HR"])

        self.assertIn("department1", result.columns)
        self.assertIn("department2", result.columns)
        self.assertNotIn("department", result.columns)
        self.assertNotIn("flag_duplicate", result.columns)
        self.assertNotIn("__batch_id", result.columns)
        self.assertEqual(result.loc[0, "department1"], "IT")
        self.assertEqual(result.loc[0, "department2"], "HR")
        self.assertEqual(result.loc[0, "institution"], "Sonic Suisse SA")
        self.assertEqual(result.loc[0, "auth"], "iomadoidc")

    def test_build_upload_export_preserves_single_department_when_no_override_is_set(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "2",
                    "username": "asmith",
                    "email": "asmith@example.com",
                    "firstname": "Alice",
                    "lastname": "Smith",
                    "department": "Finance",
                }
            ]
        )

        result = build_upload_export(df)

        self.assertIn("department", result.columns)
        self.assertEqual(result.loc[0, "department"], "Finance")
        self.assertNotIn("department1", result.columns)


if __name__ == "__main__":
    unittest.main()
