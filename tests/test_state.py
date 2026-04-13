from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

import state


class AppStateTests(unittest.TestCase):
    def _build_state(self, temp_dir: str) -> state.AppState:
        temp_path = Path(temp_dir)
        with patch.object(state, "_bundle_base_dir", return_value=temp_path), patch.object(
            state, "_app_runtime_dir", return_value=temp_path
        ):
            return state.AppState()

    def test_save_and_load_batch_export_tracker_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            app_state = self._build_state(temp_dir)
            app_state.batch_exported_ids = {"2", "1"}
            app_state.save_batch_export_tracker()

            reloaded_state = self._build_state(temp_dir)

            self.assertEqual(reloaded_state.batch_exported_ids, {"1", "2"})

    def test_invalid_batch_tracker_falls_back_to_empty_set(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            (temp_path / "batch_export_tracker.json").write_text("{invalid", encoding="utf-8")

            app_state = self._build_state(temp_dir)

            self.assertEqual(app_state.batch_exported_ids, set())
            self.assertIsNotNone(app_state.batch_export_tracker_error)
            self.assertTrue((temp_path / "batch_export_tracker.json.invalid").exists())
            self.assertTrue(app_state.runtime_warnings)

    def test_load_users_and_reset_restore_original_dataframe(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            csv_path = temp_path / "users.csv"
            pd.DataFrame(
                [
                    {"id": "1", "username": "one", "email": "one@example.com", "firstname": "One", "lastname": "User"},
                    {"id": "2", "username": "two", "email": "two@example.com", "firstname": "Two", "lastname": "User"},
                ]
            ).to_csv(csv_path, index=False)

            app_state = self._build_state(temp_dir)
            app_state.users_file = str(csv_path)
            app_state.users_sheet = None

            app_state.load_users()
            app_state.current_df = app_state.current_df.iloc[:1].copy()
            app_state.reset()

            self.assertEqual(len(app_state.original_df), 2)
            self.assertEqual(len(app_state.current_df), 2)

    def test_reset_batch_export_tracker_clears_error_and_persists_empty_tracker(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            (temp_path / "batch_export_tracker.json").write_text("{invalid", encoding="utf-8")

            app_state = self._build_state(temp_dir)
            app_state.reset_batch_export_tracker()

            self.assertIsNone(app_state.batch_export_tracker_error)
            self.assertEqual(app_state.batch_exported_ids, set())
            self.assertIn('"exported_ids": []', (temp_path / "batch_export_tracker.json").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
