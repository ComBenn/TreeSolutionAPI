from pathlib import Path
from types import SimpleNamespace
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"))

from ui_app import TreeSolutionHelperUI


class TreeSolutionHelperUITests(unittest.TestCase):
    def test_load_users_into_state_applies_all_templates_when_requested(self) -> None:
        ui = TreeSolutionHelperUI.__new__(TreeSolutionHelperUI)
        call_order: list[object] = []
        loaded_df = pd.DataFrame([{"id": "1"}, {"id": "2"}])
        state = SimpleNamespace(current_df=None, users_file="users.xlsx")

        def _load_users() -> None:
            call_order.append("load")
            state.current_df = loaded_df.copy()

        state.load_users = _load_users
        ui.state = state
        ui._sync_state_paths = lambda: call_order.append("sync")
        ui._refresh_auto_flags = lambda: call_order.append("refresh")
        ui._apply_all_employee_templates_to_original_users = (
            lambda label: call_order.append(("apply", label))
        )
        logs: list[str] = []
        ui._log = logs.append

        ui._load_users_into_state(
            load_message="Benutzer geladen",
            template_summary_label="Vorlagen automatisch beim Laden angewendet",
        )

        self.assertEqual(
            call_order,
            [
                "sync",
                "load",
                "refresh",
                ("apply", "Vorlagen automatisch beim Laden angewendet"),
            ],
        )
        self.assertEqual(logs, ["Benutzer geladen: 2 aus users.xlsx"])

    def test_show_technical_accounts_table_export_opens_filtered_export_view(self) -> None:
        ui = TreeSolutionHelperUI.__new__(TreeSolutionHelperUI)
        ui._with_errors = lambda fn: fn()
        ui._ensure_original_users_loaded = lambda: pd.DataFrame()

        marked_df = pd.DataFrame(
            [
                {
                    "id": "1",
                    "username": "svc-app",
                    "flag_technical_account": True,
                    "flag_technical_reason": "exact_id:svc-app",
                },
                {
                    "id": "2",
                    "username": "max",
                    "flag_technical_account": False,
                    "flag_technical_reason": "",
                },
            ]
        )
        ui._get_marked_technical_df = lambda: (marked_df, {"svc-app"})
        refresh_calls: list[bool] = []
        save_calls: list[bool] = []
        ui._ensure_technical_template_present = lambda marked_df=None: refresh_calls.append(marked_df is not None)
        ui._refresh_employee_templates_view = lambda: refresh_calls.append(True)
        ui._save_ui_state = lambda: save_calls.append(True)
        logs: list[str] = []
        ui._log = logs.append
        opened: dict[str, object] = {}
        ui.show_current_table = lambda **kwargs: opened.update(kwargs)

        ui.show_technical_accounts_table_export()

        technical_df = opened["df_override"]
        self.assertIsInstance(technical_df, pd.DataFrame)
        self.assertEqual(technical_df.to_dict(orient="records"), [{"id": "1", "username": "svc-app"}])
        self.assertEqual(opened["title_override"], "Technische Accounts (1 Zeilen)")
        self.assertEqual(opened["sync_state"], False)
        self.assertEqual(refresh_calls, [True, True])
        self.assertEqual(save_calls, [True])
        self.assertEqual(logs, ["Technische Accounts anzeigen: Keywords: 1 | Treffer: 1"])


if __name__ == "__main__":
    unittest.main()
