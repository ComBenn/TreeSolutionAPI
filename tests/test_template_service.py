from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from template_service import apply_employee_templates, sanitize_employee_templates


class TemplateServiceTests(unittest.TestCase):
    def test_sanitize_employee_templates_normalizes_invalid_values(self) -> None:
        raw = [
            {
                "name": " Team A ",
                "file": " team.csv ",
                "sheet": " Sheet1 ",
                "mode": "INVALID",
                "kind": "INVALID",
                "readonly": 1,
                "internal_ids": ["1", " 1 ", "", "2"],
                "internal_rows": [{"id": 1, "name": "A"}],
            }
        ]

        result = sanitize_employee_templates(raw)

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["name"], "Team A")
        self.assertEqual(result[0]["file"], "team.csv")
        self.assertEqual(result[0]["sheet"], "Sheet1")
        self.assertEqual(result[0]["mode"], "include")
        self.assertEqual(result[0]["kind"], "employee")
        self.assertEqual(result[0]["internal_ids"], ["1", "2"])
        self.assertEqual(result[0]["internal_rows"], [{"id": "1", "name": "A"}])

    def test_apply_employee_templates_lets_include_templates_override_excludes(self) -> None:
        df_base = pd.DataFrame(
            [
                {"id": "1", "username": "one"},
                {"id": "2", "username": "two"},
                {"id": "3", "username": "three"},
            ]
        )
        templates = [
            {
                "name": "Include",
                "mode": "include",
                "kind": "employee",
                "internal_ids": ["1", "2"],
                "internal_rows": [],
            },
            {
                "name": "Exclude",
                "mode": "exclude",
                "kind": "employee",
                "internal_ids": ["2"],
                "internal_rows": [],
            },
        ]

        selected_df, include_count, exclude_count = apply_employee_templates(
            df_base,
            templates,
            [0, 1],
            rebuild_callback=lambda _template: ([], [], 0),
        )

        self.assertEqual(include_count, 1)
        self.assertEqual(exclude_count, 1)
        self.assertEqual(selected_df["id"].tolist(), ["1", "2", "3"])


if __name__ == "__main__":
    unittest.main()
