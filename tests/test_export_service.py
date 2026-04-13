from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from export_service import build_export_df, format_export_log_message


class ExportServiceTests(unittest.TestCase):
    def test_build_export_df_passes_multiple_departments_through(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "id": "1",
                    "username": "user",
                    "email": "user@example.com",
                    "firstname": "User",
                    "lastname": "Example",
                }
            ]
        )

        result = build_export_df(df, ["IT", "HR"])

        self.assertEqual(result.loc[0, "department1"], "IT")
        self.assertEqual(result.loc[0, "department2"], "HR")

    def test_format_export_log_message_handles_with_and_without_departments(self) -> None:
        self.assertIn("Departments: IT, HR", format_export_log_message("out.csv", 3, ["IT", "HR"]))
        self.assertEqual(
            format_export_log_message("out.csv", 3, []),
            "Export geschrieben: out.csv | Zeilen: 3",
        )


if __name__ == "__main__":
    unittest.main()
