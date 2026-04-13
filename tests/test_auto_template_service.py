from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from auto_template_service import (
    build_internal_duplicate_template_data,
    build_internal_technical_template_data,
    find_template_index_by_name,
    upsert_auto_template,
)


class AutoTemplateServiceTests(unittest.TestCase):
    def test_find_template_index_by_name_is_case_insensitive(self) -> None:
        templates = [{"name": "Technische Accounts (Auto)"}, {"name": "Team A"}]

        self.assertEqual(find_template_index_by_name(templates, "team a"), 1)
        self.assertIsNone(find_template_index_by_name(templates, "missing"))

    def test_build_internal_technical_template_data_filters_and_removes_helper_columns(self) -> None:
        marked_df = pd.DataFrame(
            [
                {"id": "1", "username": "tech", "flag_technical_account": True, "flag_technical_reason": "exact_id:svc"},
                {"id": "2", "username": "user", "flag_technical_account": False, "flag_technical_reason": ""},
            ]
        )

        ids, rows, hits = build_internal_technical_template_data(marked_df, "id")

        self.assertEqual(ids, ["1"])
        self.assertEqual(hits, 1)
        self.assertEqual(rows, [{"id": "1", "username": "tech"}])

    def test_build_internal_duplicate_template_data_uses_only_excluded_ids(self) -> None:
        marked_df = pd.DataFrame(
            [
                {"id": "1", "flag_duplicate": True, "flag_duplicate_group": "dup-0001"},
                {"id": "2", "flag_duplicate": True, "flag_duplicate_group": "dup-0001"},
                {"id": "3", "flag_duplicate": False, "flag_duplicate_group": ""},
            ]
        )

        ids, rows, hits = build_internal_duplicate_template_data(marked_df, {"2"}, "id")

        self.assertEqual(ids, ["2"])
        self.assertEqual(hits, 1)
        self.assertEqual(rows, [{"id": "2", "flag_duplicate": "True", "flag_duplicate_group": "dup-0001"}])

    def test_upsert_auto_template_inserts_and_updates(self) -> None:
        templates: list[dict] = []

        inserted, payload = upsert_auto_template(
            templates,
            template_name="Duplikate ausgeschlossen (Auto)",
            file_marker="<auto:duplicate_review>",
            kind="duplicate",
            ids=["1"],
            rows=[{"id": "1"}],
            insert_at=0,
        )

        self.assertTrue(inserted)
        self.assertEqual(len(templates), 1)
        self.assertEqual(payload["internal_ids"], ["1"])

        inserted, payload = upsert_auto_template(
            templates,
            template_name="Duplikate ausgeschlossen (Auto)",
            file_marker="<auto:duplicate_review>",
            kind="duplicate",
            ids=["2"],
            rows=[{"id": "2"}],
            insert_at=0,
        )

        self.assertFalse(inserted)
        self.assertEqual(len(templates), 1)
        self.assertEqual(templates[0]["internal_ids"], ["2"])
        self.assertEqual(payload["internal_rows"], [{"id": "2"}])


if __name__ == "__main__":
    unittest.main()
