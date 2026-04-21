from pathlib import Path
import sys
import unittest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from duplicate_dialogs import (
    _extract_department_values,
    _filter_row_records,
    _resolve_initial_excluded_ids,
    _sort_row_records,
)


class DuplicateDialogHelpersTests(unittest.TestCase):
    def setUp(self) -> None:
        self.row_records = [
            {
                "iid": "dup-0",
                "excluded": False,
                "group": "dup-0001",
                "data": {"firstname": "Alice", "lastname": "Doe", "id": "10", "flag_duplicate_group": "dup-0001"},
            },
            {
                "iid": "dup-1",
                "excluded": True,
                "group": "dup-0001",
                "data": {"firstname": "Bob", "lastname": "Brown", "id": "2", "flag_duplicate_group": "dup-0001"},
            },
            {
                "iid": "dup-2",
                "excluded": False,
                "group": "dup-0002",
                "data": {"firstname": "ALINA", "lastname": "Miller", "id": "", "flag_duplicate_group": "dup-0002"},
            },
        ]

    def test_filter_row_records_filters_case_insensitive_by_selected_column(self) -> None:
        result = _filter_row_records(self.row_records, "firstname", "ali")

        self.assertEqual([record["iid"] for record in result], ["dup-0", "dup-2"])

    def test_sort_row_records_sorts_numeric_then_empty_values(self) -> None:
        result = _sort_row_records(self.row_records, "id", False)

        self.assertEqual([record["iid"] for record in result], ["dup-1", "dup-0", "dup-2"])

    def test_sort_row_records_can_sort_descending_text(self) -> None:
        result = _sort_row_records(self.row_records, "lastname", True)

        self.assertEqual([record["iid"] for record in result], ["dup-2", "dup-0", "dup-1"])

    def test_extract_department_values_collects_single_and_numbered_columns(self) -> None:
        result = _extract_department_values(
            {
                "department2": "HR",
                "department": "Finance",
                "department1": "duplicates",
                "Department3": "HR",
                "username": "alice",
            }
        )

        self.assertEqual(result, ["Finance", "duplicates", "HR"])

    def test_resolve_initial_excluded_ids_prefers_reviewed_state_and_defaults_new_duplicates_department(self) -> None:
        row_records = [
            {"id": "10", "department_values": ["duplicates"]},
            {"id": "11", "department_values": ["Sales"]},
            {"id": "12", "department_values": ["duplicates"]},
            {"id": "13", "department_values": ["HR"]},
        ]

        result = _resolve_initial_excluded_ids(
            row_records,
            saved_excluded_ids={"11"},
            reviewed_ids={"11", "12"},
        )

        self.assertEqual(result, {"10", "11"})

    def test_resolve_initial_excluded_ids_keeps_reviewed_rows_without_saved_exclusion_selected(self) -> None:
        row_records = [{"id": "21", "department_values": ["duplicates"]}]

        result = _resolve_initial_excluded_ids(
            row_records,
            saved_excluded_ids=set(),
            reviewed_ids={"21"},
        )

        self.assertEqual(result, set())


if __name__ == "__main__":
    unittest.main()
