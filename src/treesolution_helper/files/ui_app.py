from __future__ import annotations

import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from config import (
    COL_ID,
)
from state import AppState
from io_utils import load_table, load_keywords_txt
from filters_duplicates import mark_duplicate_accounts
from filters_technical import mark_technical_accounts
from filters_employee_list import mark_by_employee_list
from exporter import export_utf8_csv
from export_service import build_export_df, format_export_log_message
from auto_template_service import (
    build_internal_duplicate_template_data,
    build_internal_technical_template_data,
    find_template_index_by_name,
    upsert_auto_template,
)
from export_dialogs import open_batch_export_window, open_current_table_dialog
from duplicate_dialogs import open_duplicate_review_dialog
from template_service import (
    apply_employee_templates,
    build_internal_template_data,
    normalize_employee_list_sheet,
    sanitize_employee_templates,
)


class TreeSolutionHelperUI:
    """Tkinter-Hauptfenster fuer Import, Vorlagenverwaltung, Review und Export."""
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("TreeSolution Helper")
        self.root.geometry("1400x920")
        self.root.minsize(1200, 820)

        self.state = AppState()

        self.users_file_var = tk.StringVar(value=self.state.users_file)
        self.users_sheet_var = tk.StringVar(value=self.state.users_sheet or "")
        self.keywords_file_var = tk.StringVar(value=self.state.keywords_file)
        self.output_file_var = tk.StringVar(value=self.state.output_file)
        self.export_department_override_var = tk.StringVar(value="")
        self.export_department_override_vars: list[tk.StringVar] = [self.export_department_override_var]
        self.batch_export_count_var = tk.StringVar(value="250")
        self.employee_template_name_var = tk.StringVar(value="")
        self.employee_list_file_var = tk.StringVar(value="")
        self.employee_list_sheet_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Bereit.")
        self.employee_list_templates: list[dict] = []
        self.employee_templates_tree: ttk.Treeview | None = None
        self.technical_template_name = "Technische Accounts (Auto)"
        self.duplicate_template_name = "Duplikate ausgeschlossen (Auto)"
        self.duplicate_excluded_ids: set[str] = set()
        self.duplicate_reviewed_ids: set[str] = set()
        self._ui_state_load_warning_active = False
        self._ui_state_save_warning_active = False

        self._build_ui()
        self._flush_state_runtime_warnings()
        self._refresh_status()
        self._load_ui_state()
        self._ensure_technical_template_present()
        self._ensure_duplicate_template_present()
        self._refresh_employee_templates_view()
        self.root.after(100, self._auto_load_last_users_file)

    def _build_ui(self) -> None:
        top = tk.Frame(self.root, padx=10, pady=10)
        top.pack(fill="x")

        self._build_file_row(top, "Benutzerdatei", self.users_file_var, self._pick_users_file, row=0, on_enter=self.load_users)
        self._build_entry_row(top, "Users Sheet", self.users_sheet_var, row=1, on_enter=self.load_users)
        self._build_file_row(top, "Keyword-Datei", self.keywords_file_var, self._pick_keywords_file, row=2, on_enter=self.load_users)

        actions = tk.LabelFrame(self.root, text="Aktionen", padx=10, pady=10)
        actions.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(actions, text="Benutzer laden", command=self.load_users, width=24).grid(row=0, column=0, padx=4, pady=4, sticky="w")
        tk.Button(actions, text="Auswahl zurücksetzen", command=self.reset_users, width=24).grid(row=0, column=1, padx=4, pady=4, sticky="w")

        technical = tk.LabelFrame(self.root, text="Technische Accounts", padx=10, pady=10)
        technical.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(technical, text="Keyword-Datei öffnen", command=self.show_keywords, width=22).grid(row=0, column=0, padx=4, pady=4, sticky="w")
        tk.Button(
            technical,
            text="Technische Accounts anzeigen und exportieren",
            command=self.show_technical_accounts_table_export,
            width=40,
        ).grid(row=0, column=1, padx=4, pady=4, sticky="w")

        duplicates = tk.LabelFrame(self.root, text="Duplikate", padx=10, pady=10)
        duplicates.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(duplicates, text="Duplikate prüfen", command=self.review_duplicates, width=22).grid(row=0, column=0, padx=4, pady=4, sticky="w")

        employee = tk.LabelFrame(self.root, text="Mitarbeiterliste", padx=10, pady=10)
        employee.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(employee, text="Vorlage erstellen/aktualisieren", command=self.open_employee_template_dialog, width=32).grid(
            row=0, column=0, padx=4, pady=4, sticky="w"
        )
        tk.Button(employee, text="Vorlage entfernen", command=self.remove_employee_template, width=32).grid(
            row=1, column=0, padx=4, pady=4, sticky="w"
        )
        tk.Button(employee, text="Modus umschalten (ein/aus)", command=self.toggle_selected_template_mode, width=32).grid(
            row=0, column=1, padx=4, pady=4, sticky="w"
        )
        tk.Button(employee, text="Alle Vorlagen anwenden", command=self.mark_employee_list, width=32).grid(
            row=1, column=1, padx=4, pady=4, sticky="w"
        )
        tk.Button(employee, text="Vorlagen anzeigen und exportieren", command=self.show_selected_templates_table_export, width=32).grid(
            row=0, column=2, padx=4, pady=4, sticky="w"
        )

        template_table = tk.Frame(employee)
        template_table.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=4, pady=(8, 0))
        employee.grid_columnconfigure(0, weight=0)
        employee.grid_columnconfigure(1, weight=0)
        employee.grid_columnconfigure(2, weight=1)
        employee.grid_rowconfigure(2, weight=1)

        self.employee_templates_tree = ttk.Treeview(
            template_table,
            columns=("name", "mode", "entries"),
            show="headings",
            height=5,
            selectmode="extended",
        )
        self.employee_templates_tree.heading("name", text="Vorlage")
        self.employee_templates_tree.heading("mode", text="Modus")
        self.employee_templates_tree.heading("entries", text="Einträge")
        self.employee_templates_tree.column("name", width=320, minwidth=240, stretch=False)
        self.employee_templates_tree.column("mode", width=120, minwidth=100, stretch=False)
        self.employee_templates_tree.column("entries", width=100, minwidth=80, stretch=False)
        emp_vsb = ttk.Scrollbar(template_table, orient="vertical", command=self.employee_templates_tree.yview)
        self.employee_templates_tree.configure(yscrollcommand=emp_vsb.set)
        self.employee_templates_tree.grid(row=0, column=0, sticky="nsew")
        emp_vsb.grid(row=0, column=1, sticky="ns")
        template_table.grid_rowconfigure(0, weight=1)
        template_table.grid_columnconfigure(0, weight=1)
        self._bind_employee_template_tree_resize(self.employee_templates_tree)
        self._bind_employee_template_context_menu(self.employee_templates_tree)

        export_box = tk.LabelFrame(self.root, text="Export", padx=10, pady=10)
        export_box.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(
            export_box,
            text="Aktuelle Auswahl anzeigen und exportieren",
            command=self.show_current_table,
            width=40,
        ).grid(row=0, column=0, padx=4, pady=4, sticky="w")
        tk.Button(
            export_box,
            text="Aktuelle Auswahl als Batch exportieren",
            command=self.show_batch_export_window,
            width=40,
        ).grid(row=1, column=0, padx=4, pady=4, sticky="w")
        tk.Button(
            export_box,
            text="Batch-Merkliste zurücksetzen",
            command=self.reset_batch_export_tracker,
            width=32,
        ).grid(row=1, column=1, padx=4, pady=4, sticky="w")

        preview = tk.LabelFrame(self.root, text="Vorschau / Log", padx=10, pady=10)
        preview.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        button_bar = tk.Frame(preview)
        button_bar.pack(fill="x")
        tk.Button(button_bar, text="Vorschau aktualisieren", command=self.preview_current).pack(side="left", padx=(0, 6))
        tk.Button(button_bar, text="Log leeren", command=lambda: self.log.delete("1.0", tk.END)).pack(side="left")
        tk.Label(button_bar, textvariable=self.status_var, anchor="w").pack(side="left", padx=12)

        self.log = scrolledtext.ScrolledText(preview, wrap="none", height=20)
        self.log.pack(fill="both", expand=True, pady=(8, 0))

    def _build_file_row(
        self,
        parent: tk.Misc,
        label: str,
        var: tk.StringVar,
        browse_cmd,
        row: int,
        on_enter=None,
    ) -> tk.Entry:
        tk.Label(parent, text=label, width=16, anchor="w").grid(row=row, column=0, padx=4, pady=4, sticky="w")
        entry = tk.Entry(parent, textvariable=var, width=80)
        entry.grid(row=row, column=1, padx=4, pady=4, sticky="we")
        if on_enter is not None:
            entry.bind("<Return>", lambda _e: on_enter())
        tk.Button(parent, text="Durchsuchen", command=browse_cmd, width=14).grid(row=row, column=2, padx=4, pady=4, sticky="w")
        if hasattr(parent, "grid_columnconfigure"):
            parent.grid_columnconfigure(1, weight=1)
        return entry

    def _build_entry_row(self, parent: tk.Misc, label: str, var: tk.StringVar, row: int, on_enter=None) -> tk.Entry:
        tk.Label(parent, text=label, width=16, anchor="w").grid(row=row, column=0, padx=4, pady=4, sticky="w")
        entry = tk.Entry(parent, textvariable=var, width=80)
        entry.grid(row=row, column=1, padx=4, pady=4, sticky="we")
        if on_enter is not None:
            entry.bind("<Return>", lambda _e: on_enter())
        if hasattr(parent, "grid_columnconfigure"):
            parent.grid_columnconfigure(1, weight=1)
        return entry

    def _build_department_override_controls(
        self,
        parent: tk.Misc,
        start_row: int,
        on_enter=None,
    ) -> None:
        rows_container = tk.Frame(parent)
        rows_container.grid(row=start_row, column=0, columnspan=3, sticky="we")
        if hasattr(parent, "grid_columnconfigure"):
            parent.grid_columnconfigure(1, weight=1)

        def _render_rows() -> None:
            for child in rows_container.winfo_children():
                child.destroy()

            for idx, var in enumerate(self.export_department_override_vars):
                label = "Export department" if idx == 0 else ""
                entry = self._build_entry_row(rows_container, label, var, row=idx)
                if on_enter is not None:
                    entry.bind("<Return>", lambda _e: on_enter())

            button_row = len(self.export_department_override_vars)
            tk.Label(rows_container, text="", width=16, anchor="w").grid(row=button_row, column=0, padx=4, pady=(0, 4), sticky="w")
            tk.Button(
                rows_container,
                text="Weiteres Department",
                command=lambda: self._with_errors(_add_department_row),
                width=20,
            ).grid(row=button_row, column=1, padx=4, pady=(0, 4), sticky="w")

        def _add_department_row() -> None:
            self.export_department_override_vars.append(tk.StringVar(value=""))
            _render_rows()

        _render_rows()

    def _log(self, text: str) -> None:
        self.log.insert(tk.END, text.rstrip() + "\n")
        self.log.see(tk.END)

    def _flush_state_runtime_warnings(self) -> None:
        for warning in self.state.consume_runtime_warnings():
            self._log(f"Warnung: {warning}")

    def _refresh_status(self) -> None:
        rows = "(nicht geladen)" if self.state.current_df is None else str(len(self.state.current_df))
        self.status_var.set(f"Aktuelle Zeilen: {rows}")

    def _sync_state_paths(self) -> None:
        self.state.users_file = self.users_file_var.get().strip()
        self.state.users_sheet = self.users_sheet_var.get().strip() or None
        self.state.keywords_file = self.keywords_file_var.get().strip()
        self.state.output_file = self.output_file_var.get().strip()

    def _set_export_department_override_values(self, values: list[str]) -> None:
        normalized = [str(v).strip() for v in values]
        if not normalized:
            normalized = [""]
        self.export_department_override_vars = [tk.StringVar(value=value) for value in normalized]
        self.export_department_override_var = self.export_department_override_vars[0]

    def _get_export_department_override_values(self) -> list[str]:
        values = [var.get().strip() for var in self.export_department_override_vars if var.get().strip()]
        return values

    def _reset_export_department_override_fields(self) -> None:
        first_value = self.export_department_override_vars[0].get().strip() if self.export_department_override_vars else ""
        self._set_export_department_override_values([first_value])

    def _load_ui_state(self) -> None:
        p = self.state.ui_state_file
        if not p.exists():
            return
        try:
            payload = json.loads(p.read_text(encoding="utf-8"))
            self._ui_state_load_warning_active = False
        except Exception as exc:
            if not self._ui_state_load_warning_active:
                self._log(f"Warnung: UI-Status konnte nicht geladen werden: {p.name} ({exc})")
                self._ui_state_load_warning_active = True
            return

        users_file = str(payload.get("users_file", "")).strip()
        users_sheet = payload.get("users_sheet")
        keywords_file = str(payload.get("keywords_file", "")).strip()
        output_file = str(payload.get("output_file", "")).strip()
        export_department = str(payload.get("export_department_override", "")).strip()
        export_departments_raw = payload.get("export_department_overrides", [])
        employee_list_file = str(payload.get("employee_list_file", "")).strip()
        employee_list_sheet = str(payload.get("employee_list_sheet", "")).strip()
        employee_template_name = str(payload.get("employee_template_name", "")).strip()
        templates_raw = payload.get("employee_list_templates", [])
        duplicate_excluded_ids_raw = payload.get("duplicate_excluded_ids", [])
        duplicate_reviewed_ids_raw = payload.get("duplicate_reviewed_ids", [])

        if users_file:
            self.users_file_var.set(users_file)
        if users_sheet is not None:
            self.users_sheet_var.set(str(users_sheet))
        if keywords_file:
            self.keywords_file_var.set(keywords_file)
        if output_file:
            self.output_file_var.set(output_file)
        if employee_list_file:
            self.employee_list_file_var.set(employee_list_file)
        if employee_list_sheet:
            self.employee_list_sheet_var.set(employee_list_sheet)
        if employee_template_name:
            self.employee_template_name_var.set(employee_template_name)
        export_departments = []
        if isinstance(export_departments_raw, list):
            export_departments = [str(x).strip() for x in export_departments_raw if str(x).strip()]
        elif export_department:
            export_departments = [export_department]
        if not export_departments:
            export_departments = [""]
        self._set_export_department_override_values(export_departments)
        self.duplicate_excluded_ids = {
            str(x).strip()
            for x in duplicate_excluded_ids_raw
            if str(x).strip()
        }
        self.duplicate_reviewed_ids = {
            str(x).strip()
            for x in duplicate_reviewed_ids_raw
            if str(x).strip()
        }
        self.employee_list_templates = self._sanitize_employee_templates(templates_raw)
        self._ensure_technical_template_present()
        self._ensure_duplicate_template_present()
        self._refresh_employee_templates_view()
        self._sync_state_paths()

    def _save_ui_state(self) -> None:
        self._sync_state_paths()
        payload = {
            "users_file": self.users_file_var.get().strip(),
            "users_sheet": self.users_sheet_var.get().strip(),
            "keywords_file": self.keywords_file_var.get().strip(),
            "output_file": self.output_file_var.get().strip(),
            "export_department_override": self.export_department_override_var.get().strip(),
            "export_department_overrides": self._get_export_department_override_values(),
            "employee_list_file": self.employee_list_file_var.get().strip(),
            "employee_list_sheet": self.employee_list_sheet_var.get().strip(),
            "employee_template_name": self.employee_template_name_var.get().strip(),
            "duplicate_excluded_ids": sorted(self.duplicate_excluded_ids),
            "duplicate_reviewed_ids": sorted(self.duplicate_reviewed_ids),
            "employee_list_templates": self.employee_list_templates,
        }
        try:
            self.state.ui_state_file.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            self._ui_state_save_warning_active = False
        except Exception as exc:
            if not self._ui_state_save_warning_active:
                self._log(f"Warnung: UI-Status konnte nicht gespeichert werden: {self.state.ui_state_file.name} ({exc})")
                self._ui_state_save_warning_active = True

    def _sanitize_employee_templates(self, templates_raw) -> list[dict]:
        return sanitize_employee_templates(templates_raw)

    def _find_template_index_by_name(self, name: str) -> int | None:
        return find_template_index_by_name(self.employee_list_templates, name)

    def _bind_employee_template_context_menu(self, tree: ttk.Treeview) -> None:
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Modus umschalten", command=lambda: self._with_errors(self.toggle_selected_template_mode))

        def _show_menu(event) -> None:
            row_id = tree.identify_row(event.y)
            if not row_id:
                return
            current_selection = tree.selection()
            if row_id not in current_selection:
                tree.selection_set(row_id)
            tree.focus(row_id)
            menu.tk_popup(event.x_root, event.y_root)
            menu.grab_release()

        tree.bind("<Button-3>", _show_menu)

    def _build_internal_technical_template_data(
        self,
        marked_df: pd.DataFrame | None = None,
    ) -> tuple[list[str], list[dict], int]:
        if self.state.original_df is None and marked_df is None:
            return [], [], 0
        if marked_df is None:
            marked_df, _keywords = self._get_marked_technical_df()
        return build_internal_technical_template_data(marked_df, COL_ID)

    def _ensure_technical_template_present(self, marked_df: pd.DataFrame | None = None) -> None:
        ids, rows, hits = self._build_internal_technical_template_data(marked_df=marked_df)
        inserted, _payload = upsert_auto_template(
            self.employee_list_templates,
            self.technical_template_name,
            "<auto:keywords_technische_accounts>",
            "technical",
            ids,
            rows,
            insert_at=0,
        )
        if inserted:
            self._log(f"Auto-Vorlage bereitgestellt: {self.technical_template_name} | Treffer: {hits}")

    def _build_internal_duplicate_template_data(
        self,
        marked_df: pd.DataFrame | None = None,
    ) -> tuple[list[str], list[dict], int]:
        if self.state.original_df is None and marked_df is None:
            return [], [], 0
        if marked_df is None:
            marked_df = self._get_marked_duplicate_df()
        return build_internal_duplicate_template_data(marked_df, self.duplicate_excluded_ids, COL_ID)

    def _ensure_duplicate_template_present(self, marked_df: pd.DataFrame | None = None) -> None:
        ids, rows, hits = self._build_internal_duplicate_template_data(marked_df=marked_df)
        insert_at = 1 if self._find_template_index_by_name(self.technical_template_name) is not None else 0
        inserted, _payload = upsert_auto_template(
            self.employee_list_templates,
            self.duplicate_template_name,
            "<auto:duplicate_review>",
            "duplicate",
            ids,
            rows,
            insert_at=insert_at,
        )
        if inserted:
            self._log(f"Auto-Vorlage bereitgestellt: {self.duplicate_template_name} | Treffer: {hits}")

    def _refresh_employee_templates_view(self) -> None:
        if self.employee_templates_tree is None:
            return
        tree = self.employee_templates_tree
        tree.delete(*tree.get_children())
        for i, t in enumerate(self.employee_list_templates):
            mode_label = "einschliessen" if t.get("mode") == "include" else "ausschliessen"
            name_label = str(t.get("name", ""))
            if bool(t.get("readonly", False)):
                name_label = f"{name_label} [auto]"
            entries_count = 0
            internal_rows = t.get("internal_rows", [])
            internal_ids = t.get("internal_ids", [])
            if isinstance(internal_rows, list) and internal_rows:
                entries_count = len(internal_rows)
            elif isinstance(internal_ids, list):
                entries_count = len(internal_ids)
            tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    name_label,
                    mode_label,
                    str(entries_count),
                ),
            )

    def _get_selected_template_indices(self) -> list[int]:
        if self.employee_templates_tree is None:
            return []
        selected = self.employee_templates_tree.selection()
        if not selected:
            return []
        indices: list[int] = []
        for iid_obj in selected:
            iid = str(iid_obj)
            if not iid.isdigit():
                continue
            idx = int(iid)
            if 0 <= idx < len(self.employee_list_templates):
                indices.append(idx)
        return sorted(set(indices))

    def _get_selected_template_index(self) -> int | None:
        indices = self._get_selected_template_indices()
        return indices[0] if indices else None

    def _upsert_employee_template(
        self,
        name: str,
        file_path: str,
        sheet: str,
        selected_idx: int | None = None,
    ) -> None:
        name = name.strip()
        file_path = file_path.strip()
        sheet = sheet.strip()
        if not name:
            raise RuntimeError("Bitte Vorlagenname angeben.")
        if not file_path:
            raise RuntimeError("Bitte Mitarbeiterliste auswählen.")
        if not Path(file_path).exists():
            raise RuntimeError(f"Datei nicht gefunden: {file_path}")
        sheet_norm = self._normalize_employee_list_sheet(file_path, sheet) or ""
        internal_ids, internal_rows, matched_rows = self._build_internal_template_data(file_path, sheet_norm)

        idx_by_name = None
        for i, t in enumerate(self.employee_list_templates):
            if str(t.get("name", "")).strip().casefold() == name.casefold():
                idx_by_name = i
                break

        idx_target = selected_idx if selected_idx is not None else idx_by_name
        if idx_target is None:
            self.employee_list_templates.append(
                {
                    "name": name,
                    "file": file_path,
                    "sheet": sheet_norm,
                    "mode": "include",
                    "kind": "employee",
                    "readonly": False,
                    "internal_ids": internal_ids,
                    "internal_rows": internal_rows,
                }
            )
            self._log(
                f"Mitarbeiterlisten-Vorlage gespeichert: {name} (einschliessen) | "
                f"Interne Treffer: {matched_rows}"
            )
        else:
            if idx_target < 0 or idx_target >= len(self.employee_list_templates):
                raise RuntimeError("Ungültige Vorlagenauswahl.")
            existing = self.employee_list_templates[idx_target]
            mode_existing = str(existing.get("mode", "include")).strip().casefold()
            existing["name"] = name
            existing["file"] = file_path
            existing["sheet"] = sheet_norm
            existing["mode"] = mode_existing if mode_existing in ("include", "exclude") else "include"
            existing["kind"] = "employee"
            existing["readonly"] = False
            existing["internal_ids"] = internal_ids
            existing["internal_rows"] = internal_rows
            self._log(f"Mitarbeiterlisten-Vorlage aktualisiert: {name} | Interne Treffer: {matched_rows}")

        self._refresh_employee_templates_view()
        self._save_ui_state()

    def _build_internal_template_data(self, file_path: str, sheet: str | None) -> tuple[list[str], list[dict], int]:
        df_base, _keywords = self._get_marked_technical_df()
        if "flag_technical_account" in df_base.columns:
            df_base = df_base[df_base["flag_technical_account"] != True].copy()
        return build_internal_template_data(df_base, file_path, sheet)

    def open_employee_template_dialog(self) -> None:
        self._ensure_original_users_loaded()
        selected_indices = self._get_selected_template_indices()
        selected_idx = selected_indices[0] if len(selected_indices) == 1 else None
        if selected_idx is not None and bool(self.employee_list_templates[selected_idx].get("readonly", False)):
            selected_idx = None
        initial = self.employee_list_templates[selected_idx] if selected_idx is not None else {}

        dialog = tk.Toplevel(self.root)
        dialog.title("Vorlage erstellen/aktualisieren")
        dialog.geometry("920x220")
        dialog.resizable(False, False)
        self._make_modal(dialog)

        name_var = tk.StringVar(value=str(initial.get("name", "")))
        file_var = tk.StringVar(value=str(initial.get("file", "")))
        sheet_var = tk.StringVar(value=str(initial.get("sheet", "")))

        body = tk.Frame(dialog, padx=10, pady=10)
        body.pack(fill="both", expand=True)

        self._build_entry_row(body, "Vorlagenname", name_var, row=0)
        self._build_file_row(
            body,
            "Liste",
            file_var,
            lambda: self._pick_employee_list_file_for_var(file_var, name_var),
            row=1,
        )
        self._build_entry_row(body, "Liste Sheet", sheet_var, row=2)

        footer = tk.Frame(body)
        footer.grid(row=3, column=0, columnspan=4, sticky="e", pady=(10, 0))

        def _confirm() -> None:
            try:
                self._upsert_employee_template(
                    name=name_var.get(),
                    file_path=file_var.get(),
                    sheet=sheet_var.get(),
                    selected_idx=selected_idx,
                )
            except Exception as e:
                self._log(f"Fehler: {e}")
                messagebox.showerror("Fehler", str(e))
                return
            dialog.destroy()

        tk.Button(footer, text="Bestätigen", command=_confirm, width=14).pack(side="right", padx=(6, 0))
        tk.Button(footer, text="Abbrechen", command=dialog.destroy, width=14).pack(side="right")

    def _pick_employee_list_file_for_var(self, file_var: tk.StringVar, name_var: tk.StringVar | None = None) -> None:
        path = filedialog.askopenfilename(
            title="Mitarbeiterliste auswählen",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"), ("Alle Dateien", "*.*")],
        )
        if path:
            file_var.set(path)
            if name_var is not None and not name_var.get().strip():
                name_var.set(Path(path).stem)

    def _auto_load_last_users_file(self) -> None:
        if self.state.original_df is not None:
            return
        users_path = self.users_file_var.get().strip()
        if not users_path or not Path(users_path).exists():
            return
        try:
            self._load_users_into_state(
                load_message="Benutzer automatisch beim Start geladen",
                template_summary_label="Vorlagen automatisch beim Start angewendet",
            )
            self.preview_current()
        except Exception as e:
            self._log(f"Auto-Load beim Start fehlgeschlagen: {e}")
            self._refresh_status()

    def _with_errors(self, fn) -> None:
        try:
            fn()
        except Exception as e:
            self._log(f"Fehler: {e}")
            messagebox.showerror("Fehler", str(e))
        finally:
            self._refresh_status()
            self._save_ui_state()

    def _make_modal(self, win: tk.Toplevel) -> None:
        # Modal dialog behavior: keep focus here until closed.
        win.transient(self.root)
        win.grab_set()
        win.focus_set()

    @staticmethod
    def _sort_key(value: str) -> tuple[int, float, str]:
        text = (value or "").strip()
        if text == "":
            return (2, 0.0, "")
        candidate = text.replace(",", ".")
        try:
            return (0, float(candidate), "")
        except ValueError:
            return (1, 0.0, text.casefold())

    def _filter_and_sort_df(
        self,
        df_in: pd.DataFrame,
        filters: dict[str, str],
        sort_col: str | None,
        sort_asc: bool,
    ) -> pd.DataFrame:
        filtered_df = df_in
        for col, term in filters.items():
            if col not in filtered_df.columns:
                continue
            term_norm = str(term).strip().casefold()
            if not term_norm:
                continue
            series = filtered_df[col].fillna("").astype(str).str.casefold()
            filtered_df = filtered_df[series.str.contains(term_norm, regex=False)]

        if sort_col and sort_col in filtered_df.columns:
            sort_series = filtered_df[sort_col].fillna("").astype(str)
            sort_keys = sort_series.map(self._sort_key)
            filtered_df = filtered_df.assign(__sort_key=sort_keys).sort_values(
                by="__sort_key",
                ascending=sort_asc,
                kind="mergesort",
            ).drop(columns=["__sort_key"])
        return filtered_df

    @staticmethod
    def _column_from_tree_event(tree: ttk.Treeview, columns: list[str], event) -> str | None:
        region = tree.identify_region(event.x, event.y)
        if region != "heading":
            return None
        col_id = tree.identify_column(event.x)
        if not col_id.startswith("#"):
            return None
        try:
            pos = int(col_id[1:]) - 1
        except ValueError:
            return None
        if pos < 0 or pos >= len(columns):
            return None
        return columns[pos]

    def _copy_treeview_selection(self, tree: ttk.Treeview, columns: list[str]) -> str:
        selected = tree.selection()
        if not selected:
            return "break"

        lines = ["\t".join(columns)]
        for item_id in selected:
            values = tree.item(item_id, "values")
            row = [str(v) for v in values[: len(columns)]]
            if len(row) < len(columns):
                row.extend([""] * (len(columns) - len(row)))
            lines.append("\t".join(row))

        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self._log(f"Tabellenzeilen in Zwischenablage kopiert: {len(selected)}")
        return "break"

    def _bind_treeview_shortcuts(self, tree: ttk.Treeview, columns: list[str]) -> None:
        def _select_all(_event=None) -> str:
            children = tree.get_children()
            if children:
                tree.selection_set(children)
                tree.focus(children[0])
            return "break"

        def _copy(_event=None) -> str:
            return self._copy_treeview_selection(tree, columns)

        tree.bind("<Control-a>", _select_all)
        tree.bind("<Control-A>", _select_all)
        tree.bind("<Control-c>", _copy)
        tree.bind("<Control-C>", _copy)

    @staticmethod
    def _bind_employee_template_tree_resize(tree: ttk.Treeview) -> None:
        fixed_mode_width = 130
        fixed_entries_width = 90
        min_name_width = 240
        max_name_width = 520
        scrollbar_reserve = 28

        def _resize(_event=None) -> None:
            total_width = tree.winfo_width()
            if total_width <= 1:
                return
            available = total_width - fixed_mode_width - fixed_entries_width - scrollbar_reserve
            name_width = max(min_name_width, min(max_name_width, available))
            tree.column("name", width=name_width, stretch=False)
            tree.column("mode", width=fixed_mode_width, stretch=False)
            tree.column("entries", width=fixed_entries_width, stretch=False)

        tree.bind("<Configure>", _resize)

    def _open_contains_filter_dialog(
        self,
        parent: tk.Toplevel,
        col: str,
        source_df: pd.DataFrame,
        filters: dict[str, str],
        refresh_callback,
        columns: list[str],
    ) -> None:
        if not col:
            return
        if source_df is None or source_df.empty or col not in source_df.columns:
            source_df = pd.DataFrame(columns=columns)

        dialog = tk.Toplevel(parent)
        dialog.title(f"Filter: {col}")
        dialog.geometry("460x140")
        dialog.resizable(False, False)
        self._make_modal(dialog)

        body = tk.Frame(dialog, padx=10, pady=10)
        body.pack(fill="both", expand=True)

        tk.Label(body, text=f"Filter für Spalte '{col}' (enthält):", anchor="w").grid(
            row=0, column=0, padx=4, pady=4, sticky="w"
        )
        raw_values = source_df[col].fillna("").astype(str).map(lambda x: x.strip()) if col in source_df.columns else pd.Series(dtype=str)
        unique_values = sorted({v for v in raw_values if v != ""}, key=lambda x: x.casefold())
        value_var = tk.StringVar(value=str(filters.get(col, "")))
        combo = ttk.Combobox(body, textvariable=value_var, values=unique_values[:1000], width=52)
        combo.grid(row=1, column=0, padx=4, pady=4, sticky="we")
        combo.focus_set()

        def _apply() -> None:
            term = value_var.get().strip()
            if term:
                filters[col] = term
            else:
                filters.pop(col, None)
            refresh_callback()
            dialog.destroy()

        def _clear() -> None:
            filters.pop(col, None)
            refresh_callback()
            dialog.destroy()

        buttons = tk.Frame(body)
        buttons.grid(row=2, column=0, pady=(8, 0), sticky="e")
        tk.Button(buttons, text="Filter setzen", command=lambda: self._with_errors(_apply), width=14).pack(side="left", padx=(0, 6))
        tk.Button(buttons, text="Filter löschen", command=lambda: self._with_errors(_clear), width=14).pack(side="left", padx=(0, 6))
        tk.Button(buttons, text="Abbrechen", command=dialog.destroy, width=12).pack(side="left")

        combo.bind("<Return>", lambda _e: self._with_errors(_apply))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())

    def _ask_export_save_path(self, initialfile: str | None = None) -> str:
        suggested_name = (initialfile or "").strip() or "Upload.csv"
        save_path = filedialog.asksaveasfilename(
            title="Export CSV speichern unter",
            initialfile=suggested_name,
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Alle Dateien", "*.*")],
        )
        if save_path:
            self.state.output_file = save_path
            self.output_file_var.set(save_path)
        return save_path

    def _build_export_df_from_source(self, df_source: pd.DataFrame) -> pd.DataFrame:
        department_overrides = self._get_export_department_override_values()
        return build_export_df(df_source, department_overrides)

    def _log_export_result(self, rows: int) -> None:
        department_overrides = self._get_export_department_override_values()
        self._log(format_export_log_message(self.state.output_file, rows, department_overrides))

    def _export_regular_from_df(self, df_source: pd.DataFrame, initialfile: str = "Upload.csv") -> None:
        save_path = self._ask_export_save_path(initialfile=initialfile)
        if not save_path:
            self._log("Export abgebrochen.")
            return
        self._log(f"Exportquelle (aktuelle Auswahl): {len(df_source)} Zeilen")
        export_df = self._build_export_df_from_source(df_source)
        export_utf8_csv(export_df, self.state.output_file)
        self._log_export_result(len(export_df))
        messagebox.showinfo("Export", f"Export geschrieben:\n{self.state.output_file}")

    def _export_next_batch_from_df(self, df_source: pd.DataFrame) -> None:
        self._ensure_batch_export_tracker_ready()
        df = df_source.copy()
        if COL_ID not in df.columns:
            raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der aktuellen Auswahl.")

        try:
            batch_count = int(self.batch_export_count_var.get().strip())
        except ValueError:
            raise RuntimeError("Export Anzahl muss eine ganze Zahl sein.")
        if batch_count <= 0:
            raise RuntimeError("Export Anzahl muss grösser als 0 sein.")

        ids_series = df[COL_ID].fillna("").astype(str).str.strip()
        df = df.assign(__batch_id=ids_series)
        if (df["__batch_id"] == "").any():
            empty_count = int((df["__batch_id"] == "").sum())
            self._log(f"Hinweis: {empty_count} Zeilen ohne ID werden für Batch-Export ignoriert.")
        eligible_df = df[df["__batch_id"] != ""].copy()
        remaining_df = eligible_df[~eligible_df["__batch_id"].isin(self.state.batch_exported_ids)].copy()

        if remaining_df.empty:
            self._log("Keine neuen Einträge für Batch-Export vorhanden (alle bereits gemerkt/exportiert).")
            return

        batch_df = remaining_df.head(batch_count).drop(columns=["__batch_id"])
        actual_rows = len(batch_df)

        save_path = self._ask_export_save_path(initialfile="Batch-Upload.csv")
        if not save_path:
            self._log("Batch-Export abgebrochen.")
            return

        export_df = self._build_export_df_from_source(batch_df)
        export_utf8_csv(export_df, self.state.output_file)

        exported_ids_now = {
            str(v).strip()
            for v in batch_df[COL_ID].fillna("").astype(str)
            if str(v).strip()
        }
        self.state.batch_exported_ids.update(exported_ids_now)
        self.state.save_batch_export_tracker()

        remaining_after = max(0, len(remaining_df) - actual_rows)
        self._log_export_result(len(export_df))
        self._log(
            "Batch-Export (gemerkt): "
            f"{actual_rows} Zeilen exportiert | "
            f"{len(exported_ids_now)} IDs neu gemerkt | "
            f"Noch offen in aktueller Auswahl: {remaining_after}"
        )
        self._log(
            f"Merkliste persistent gespeichert: {self.state.batch_export_tracker_file.name} "
            f"({len(self.state.batch_exported_ids)} IDs gesamt)"
        )
        messagebox.showinfo(
            "Batch-Export",
            f"{actual_rows} Zeilen exportiert.\nGemerkte IDs gesamt: {len(self.state.batch_exported_ids)}",
        )

    def _require_current_df(self) -> pd.DataFrame:
        if self.state.current_df is None:
            raise RuntimeError("Zuerst Benutzerdatei laden.")
        return self.state.current_df

    def _get_marked_technical_df(self) -> tuple[pd.DataFrame, set[str]]:
        self._sync_state_paths()
        if self.state.original_df is None:
            raise RuntimeError("Zuerst Benutzerdatei laden.")
        keywords = load_keywords_txt(self.state.keywords_file)
        marked_df = mark_technical_accounts(self.state.original_df, keywords)
        return marked_df, keywords

    def _get_marked_duplicate_df(self) -> pd.DataFrame:
        technical_df, _keywords = self._get_marked_technical_df()
        return mark_duplicate_accounts(technical_df)

    def _refresh_auto_flags(self) -> tuple[int, int]:
        technical_df, keywords = self._get_marked_technical_df()
        marked_df = mark_duplicate_accounts(technical_df)
        self.state.current_df = marked_df
        technical_hits = int(marked_df["flag_technical_account"].sum()) if "flag_technical_account" in marked_df.columns else 0
        duplicate_hits = int(marked_df["flag_duplicate"].sum()) if "flag_duplicate" in marked_df.columns else 0
        self._ensure_technical_template_present(marked_df=marked_df)
        self._ensure_duplicate_template_present(marked_df=marked_df)
        self._refresh_employee_templates_view()
        self._log(
            f"Auto-Markierungen aktualisiert. Keywords: {len(keywords)} | "
            f"Technische Treffer: {technical_hits} | Duplikat-Treffer: {duplicate_hits}"
        )
        return technical_hits, duplicate_hits

    def _refresh_technical_flags_from_keywords(self) -> int:
        hits, _duplicate_hits = self._refresh_auto_flags()
        return hits

    def _pick_users_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Benutzerdatei auswählen",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.users_file_var.set(path)
            self._save_ui_state()

    def _pick_keywords_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Keyword-Datei auswählen",
            filetypes=[("Textdateien", "*.txt"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.keywords_file_var.set(path)
            self._save_ui_state()

    def _pick_output_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Output CSV speichern",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.output_file_var.set(path)
            self._save_ui_state()

    def _pick_employee_list_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Mitarbeiterliste auswählen",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.employee_list_file_var.set(path)
            if not self.employee_template_name_var.get().strip():
                self.employee_template_name_var.set(Path(path).stem)
            self._save_ui_state()

    def _normalize_employee_list_sheet(self, list_file: str, list_sheet: str | None) -> str | None:
        return normalize_employee_list_sheet(list_file, list_sheet)

    def _ensure_original_users_loaded(self) -> pd.DataFrame:
        if self.state.original_df is None:
            self._sync_state_paths()
            users_file = str(self.state.users_file or "").strip()
            if not users_file or not Path(users_file).exists():
                raise RuntimeError(
                    "Eine Vorlage kann erst erstellt oder angewendet werden, nachdem eine Benutzerdatei geladen wurde."
                )
            self.state.load_users()
            self._log(f"Benutzer automatisch geladen: {len(self.state.current_df)} aus {self.state.users_file}")
            self._refresh_technical_flags_from_keywords()
        if self.state.original_df is None:
            raise RuntimeError("Eine Vorlage kann erst erstellt oder angewendet werden, nachdem eine Benutzerdatei geladen wurde.")
        return self.state.original_df

    def _load_users_into_state(
        self,
        load_message: str,
        template_summary_label: str | None = None,
    ) -> None:
        """Laedt Benutzer neu und wendet optional danach alle Vorlagen auf die Originaldaten an."""
        self._sync_state_paths()
        self.state.load_users()
        self._log(f"{load_message}: {len(self.state.current_df)} aus {self.state.users_file}")
        self._refresh_auto_flags()
        if template_summary_label:
            self._apply_all_employee_templates_to_original_users(template_summary_label)

    def _evaluate_employee_template(
        self,
        template: dict,
        df_base: pd.DataFrame,
        flag_name: str,
    ) -> tuple[pd.DataFrame, dict]:
        list_file = str(template.get("file", "")).strip()
        if not list_file:
            raise RuntimeError(f"Vorlage '{template.get('name', '')}' hat keine Datei.")
        list_sheet = self._normalize_employee_list_sheet(list_file, str(template.get("sheet", "")))
        df_list = load_table(list_file, list_sheet)
        marked_df, stats = mark_by_employee_list(df_base, df_list, flag_name=flag_name, return_stats=True)
        return marked_df, stats

    def save_employee_template(self) -> None:
        self.open_employee_template_dialog()

    def remove_employee_template(self) -> None:
        def _run() -> None:
            selected_indices = self._get_selected_template_indices()
            if not selected_indices:
                raise RuntimeError("Bitte zuerst eine oder mehrere Vorlagen auswählen.")
            readonly_selected = [i for i in selected_indices if bool(self.employee_list_templates[i].get("readonly", False))]
            if readonly_selected:
                raise RuntimeError("Automatische Vorlagen können nicht entfernt werden.")
            names = [str(self.employee_list_templates[i].get("name", "")) for i in selected_indices]
            for idx in sorted(selected_indices, reverse=True):
                del self.employee_list_templates[idx]
            self._refresh_employee_templates_view()
            self._save_ui_state()
            self._log(f"Mitarbeiterlisten-Vorlagen entfernt: {len(names)} | {', '.join(names)}")
        self._with_errors(_run)

    def _set_selected_template_mode(self, mode: str) -> None:
        idx = self._get_selected_template_index()
        if idx is None:
            raise RuntimeError("Bitte zuerst eine Vorlage auswählen.")
        if mode not in ("include", "exclude"):
            raise RuntimeError(f"Ungültiger Modus: {mode}")
        self.employee_list_templates[idx]["mode"] = mode
        self._refresh_employee_templates_view()
        self._save_ui_state()
        mode_label = "einschliessen" if mode == "include" else "ausschliessen"
        self._log(f"Vorlage '{self.employee_list_templates[idx]['name']}' auf '{mode_label}' gesetzt.")
        self._reapply_all_employee_templates()

    def _reapply_all_employee_templates(self) -> None:
        self._apply_all_employee_templates_to_original_users("Aktuelle Auswahl nach Modusänderung aktualisiert")
        self.preview_current()

    def toggle_selected_template_mode(self) -> None:
        def _run() -> None:
            idx = self._get_selected_template_index()
            if idx is None:
                raise RuntimeError("Bitte zuerst eine Vorlage auswählen.")
            if bool(self.employee_list_templates[idx].get("readonly", False)):
                raise RuntimeError("Der Modus automatischer Vorlagen ist fix auf 'ausschliessen'.")
            current_mode = str(self.employee_list_templates[idx].get("mode", "include")).strip().casefold()
            next_mode = "exclude" if current_mode == "include" else "include"
            self._set_selected_template_mode(next_mode)
        self._with_errors(_run)

    def _apply_employee_templates(
        self,
        df_base: pd.DataFrame,
        template_indices: list[int],
        label: str,
    ) -> tuple[pd.DataFrame, int, int]:
        def _rebuild_template(template: dict) -> tuple[list[str], list[dict], int]:
            template_kind = str(template.get("kind", "employee"))
            if template_kind == "technical":
                return self._build_internal_technical_template_data()
            if template_kind == "duplicate":
                return self._build_internal_duplicate_template_data()
            file_path = str(template.get("file", "")).strip()
            sheet = str(template.get("sheet", "")).strip()
            return self._build_internal_template_data(file_path, sheet) if file_path else ([], [], 0)

        selected_df, include_count, exclude_count = apply_employee_templates(
            df_base,
            self.employee_list_templates,
            template_indices,
            rebuild_callback=_rebuild_template,
            log_callback=lambda msg: self._log(f"{label} | {msg}"),
        )
        self._save_ui_state()
        return selected_df, include_count, exclude_count

    def _apply_all_employee_templates_to_original_users(
        self,
        summary_label: str,
    ) -> tuple[pd.DataFrame, int, int, int]:
        """Wendet die komplette Vorlagenliste auf die geladene Benutzerbasis an."""
        df_base = self._ensure_original_users_loaded()
        all_indices = list(range(len(self.employee_list_templates)))
        if not all_indices:
            selected_df = df_base.copy()
            self.state.current_df = selected_df
            self._log(
                f"{summary_label}: Anzahl Vorlagen: 0 | Einschliessen: 0 | Ausschliessen: 0 | "
                f"Verbleibend: {len(selected_df)}"
            )
            return selected_df, 0, 0, 0
        selected_df, include_count, exclude_count = self._apply_employee_templates(
            df_base,
            all_indices,
            label="Alle Vorlagen",
        )
        self.state.current_df = selected_df
        self._log(
            f"{summary_label}: Anzahl Vorlagen: {len(all_indices)} | "
            f"Einschliessen: {include_count} | Ausschliessen: {exclude_count} | "
            f"Verbleibend: {len(selected_df)}"
        )
        return selected_df, len(all_indices), include_count, exclude_count

    def load_selected_employee_template(self) -> None:
        def _run() -> None:
            selected_indices = self._get_selected_template_indices()
            if not selected_indices:
                raise RuntimeError("Bitte eine oder mehrere Vorlagen auswählen (Ctrl/Shift möglich).")
            df_base = self._ensure_original_users_loaded()
            selected_df, include_count, exclude_count = self._apply_employee_templates(
                df_base,
                selected_indices,
                label="Ausgewählte Vorlagen",
            )
            self.state.current_df = selected_df
            self._log(
                f"Ausgewählte Vorlagen geladen. Anzahl Vorlagen: {len(selected_indices)} | "
                f"Einschliessen: {include_count} | Ausschliessen: {exclude_count} | "
                f"Verbleibend: {len(selected_df)}"
            )
            self.preview_current()
        self._with_errors(_run)

    def show_selected_templates_table_export(self) -> None:
        def _run() -> None:
            selected_indices = self._get_selected_template_indices()
            if not selected_indices:
                raise RuntimeError("Bitte eine oder mehrere Vorlagen auswählen (Ctrl/Shift möglich).")
            df_base = self._ensure_original_users_loaded()
            if COL_ID not in df_base.columns:
                raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der Benutzerdatei.")

            selected_ids: set[str] = set()
            for i in selected_indices:
                template = self.employee_list_templates[i]
                ids_in_template = {
                    str(v).strip()
                    for v in template.get("internal_ids", [])
                    if str(v).strip()
                }
                if not ids_in_template:
                    template_kind = str(template.get("kind", "employee"))
                    if template_kind == "technical":
                        rebuilt_ids, rebuilt_rows, _hits = self._build_internal_technical_template_data()
                        ids_in_template = set(rebuilt_ids)
                        template["internal_ids"] = rebuilt_ids
                        template["internal_rows"] = rebuilt_rows
                    elif template_kind == "duplicate":
                        rebuilt_ids, rebuilt_rows, _hits = self._build_internal_duplicate_template_data()
                        ids_in_template = set(rebuilt_ids)
                        template["internal_ids"] = rebuilt_ids
                        template["internal_rows"] = rebuilt_rows
                    else:
                        file_path = str(template.get("file", "")).strip()
                        sheet = str(template.get("sheet", "")).strip()
                        if file_path:
                            rebuilt_ids, rebuilt_rows, _hits = self._build_internal_template_data(file_path, sheet)
                            ids_in_template = set(rebuilt_ids)
                            template["internal_ids"] = rebuilt_ids
                            template["internal_rows"] = rebuilt_rows
                selected_ids.update(ids_in_template)

            id_series = df_base[COL_ID].fillna("").astype(str).str.strip()
            selected_df = df_base[id_series.isin(selected_ids)].copy()
            self._save_ui_state()
            self._log(
                f"Vorlageninhalt anzeigen: {len(selected_indices)} Vorlagen | "
                f"Interne IDs gesamt: {len(selected_ids)} | Treffer in Benutzerdatei: {len(selected_df)}"
            )

            def _delete_rows_from_selected_templates(rows_to_delete: pd.DataFrame) -> None:
                if COL_ID not in rows_to_delete.columns:
                    return
                ids_to_delete = {
                    str(v).strip()
                    for v in rows_to_delete[COL_ID].fillna("").astype(str)
                    if str(v).strip()
                }
                if not ids_to_delete:
                    return
                changed_templates = 0
                removed_ids_total = 0
                for idx in selected_indices:
                    template = self.employee_list_templates[idx]
                    before_ids = {
                        str(v).strip()
                        for v in template.get("internal_ids", [])
                        if str(v).strip()
                    }
                    after_ids = before_ids - ids_to_delete
                    if after_ids != before_ids:
                        changed_templates += 1
                        removed_ids_total += len(before_ids) - len(after_ids)
                        template["internal_ids"] = sorted(after_ids)
                        rows = template.get("internal_rows", [])
                        if isinstance(rows, list):
                            filtered_rows = []
                            for row in rows:
                                if not isinstance(row, dict):
                                    continue
                                row_id = str(row.get(COL_ID, "")).strip()
                                if row_id and row_id in ids_to_delete:
                                    continue
                                filtered_rows.append(row)
                            template["internal_rows"] = filtered_rows
                if changed_templates > 0:
                    self._save_ui_state()
                    self._refresh_employee_templates_view()
                    self._log(
                        f"Vorlageninhalt gelöscht: {len(ids_to_delete)} IDs ausgewählt | "
                        f"{removed_ids_total} IDs aus {changed_templates} Vorlagen entfernt."
                    )

            self.show_current_table(
                df_override=selected_df,
                title_override=(
                    f"Vorlagen-Inhalt ({len(selected_df)} Zeilen | {len(selected_indices)} Vorlagen)"
                ),
                sync_state=False,
                delete_callback=_delete_rows_from_selected_templates,
                delete_confirm_text=(
                    "Sicher löschen?\n"
                    "Der Eintrag wird dauerhaft aus den ausgewählten Vorlagen entfernt."
                ),
            )
        self._with_errors(_run)

    def load_users(self) -> None:
        def _run() -> None:
            self._load_users_into_state(
                load_message="Benutzer geladen",
                template_summary_label="Vorlagen automatisch beim Laden angewendet",
            )
            self.preview_current()
        self._with_errors(_run)

    def reset_users(self) -> None:
        def _run() -> None:
            self.state.reset()
            self._refresh_auto_flags()
            self._log("Arbeitssatz auf Original zurückgesetzt.")
            self.preview_current()
        self._with_errors(_run)

    def mark_technical(self) -> None:
        def _run() -> None:
            self._sync_state_paths()
            df = self._require_current_df()
            keywords = load_keywords_txt(self.state.keywords_file)
            self.state.current_df = mark_technical_accounts(df, keywords)
            hits = int(self.state.current_df["flag_technical_account"].sum())
            self._log(f"Technische Accounts markiert. Keywords: {len(keywords)} | Treffer: {hits}")
            self.preview_current()
        self._with_errors(_run)

    def keep_technical(self) -> None:
        def _run() -> None:
            self._refresh_technical_flags_from_keywords()
            df = self._require_current_df()
            if "flag_technical_account" not in df.columns:
                raise RuntimeError("Zuerst technische Accounts markieren.")
            self.state.current_df = df[df["flag_technical_account"] == True].copy()
            self._log(f"Nur technische Accounts ausgewählt. Verbleibend: {len(self.state.current_df)}")
            self.preview_current()
        self._with_errors(_run)

    def exclude_technical(self) -> None:
        def _run() -> None:
            self._refresh_technical_flags_from_keywords()
            df = self._require_current_df()
            if "flag_technical_account" not in df.columns:
                raise RuntimeError("Zuerst technische Accounts markieren.")
            self.state.current_df = df[df["flag_technical_account"] != True].copy()
            self._log(f"Technische Accounts ausgeschlossen. Verbleibend: {len(self.state.current_df)}")
            self.preview_current()
        self._with_errors(_run)

    def review_duplicates(self) -> None:
        def _run() -> None:
            open_duplicate_review_dialog(self)
        self._with_errors(_run)

    def show_technical_accounts_table_export(self) -> None:
        """Zeigt die aktuell erkannten technischen Accounts im Standard-Exportfenster."""
        def _run() -> None:
            self._ensure_original_users_loaded()
            marked_df, keywords = self._get_marked_technical_df()
            technical_df = marked_df[marked_df["flag_technical_account"] == True].copy()
            technical_df = technical_df.drop(
                columns=["flag_technical_account", "flag_technical_reason"],
                errors="ignore",
            )
            self._ensure_technical_template_present(marked_df=marked_df)
            self._refresh_employee_templates_view()
            self._save_ui_state()
            self._log(
                f"Technische Accounts anzeigen: Keywords: {len(keywords)} | Treffer: {len(technical_df)}"
            )
            self.show_current_table(
                df_override=technical_df,
                title_override=f"Technische Accounts ({len(technical_df)} Zeilen)",
                sync_state=False,
            )
        self._with_errors(_run)

    def mark_employee_list(self) -> None:
        def _run() -> None:
            self._apply_all_employee_templates_to_original_users("Alle Vorlagen angewendet")
            self.preview_current()
        self._with_errors(_run)

    def show_keywords(self) -> None:
        def _run() -> None:
            self._sync_state_paths()
            p = Path(self.state.keywords_file)
            if not p.exists():
                p.write_text("", encoding="utf-8")
            os.startfile(str(p))
            self._log(f"Keyword-Datei zum Bearbeiten geöffnet: {p}")
        self._with_errors(_run)

    def export_csv(self) -> None:
        def _run() -> None:
            self._sync_state_paths()
            df = self._require_current_df()
            self._export_regular_from_df(df)
        self._with_errors(_run)

    def export_next_batch_csv(self) -> None:
        def _run() -> None:
            self._sync_state_paths()
            df = self._require_current_df()
            self._export_next_batch_from_df(df)
        self._with_errors(_run)

    def reset_batch_export_tracker(self) -> None:
        def _run() -> None:
            count_before = len(self.state.batch_exported_ids)
            tracker_error = self.state.batch_export_tracker_error
            if count_before == 0 and not tracker_error:
                self._log("Batch-Merkliste ist bereits leer.")
                return

            prompt = (
                "Alle gemerkten Batch-Export-IDs wirklich löschen?\n"
                "Danach können Einträge erneut über den Batch-Export exportiert werden."
            )
            if tracker_error:
                prompt = (
                    "Die Batch-Merkliste ist beschädigt und wird aktuell nicht verwendet.\n"
                    "Soll sie jetzt zurückgesetzt und neu angelegt werden?\n\n"
                    f"Details: {tracker_error}"
                )
            confirmed = messagebox.askyesno(
                "Batch-Merkliste zurücksetzen",
                prompt,
            )
            if not confirmed:
                self._log("Zurücksetzen der Batch-Merkliste abgebrochen.")
                return

            self.state.reset_batch_export_tracker()
            self._log(
                "Batch-Merkliste zurückgesetzt: "
                f"{count_before} gemerkte IDs entfernt | Datei: {self.state.batch_export_tracker_file.name}"
            )
            messagebox.showinfo(
                "Batch-Merkliste",
                f"Merkliste wurde zurückgesetzt.\nEntfernte IDs: {count_before}",
            )
        self._with_errors(_run)

    def preview_current(self) -> None:
        df = self.state.current_df
        if df is None:
            self._log("Noch keine Benutzerdatei geladen.")
            self._refresh_status()
            return

        self._log(f"Vorschau ({len(df)} Zeilen):")
        if len(df) == 0:
            self._log("(leer)")
        else:
            self._log(df.head(10).to_string(index=False))
        self._refresh_status()

    def show_current_table(
        self,
        df_override: pd.DataFrame | None = None,
        title_override: str | None = None,
        sync_state: bool = True,
        delete_callback=None,
        delete_confirm_text: str | None = None,
    ) -> None:
        open_current_table_dialog(
            self,
            df_override=df_override,
            title_override=title_override,
            sync_state=sync_state,
            delete_callback=delete_callback,
            delete_confirm_text=delete_confirm_text,
        )

    def _get_batch_remaining_df(
        self, df_source: pd.DataFrame
    ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        self._ensure_batch_export_tracker_ready()
        if COL_ID not in df_source.columns:
            raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der aktuellen Auswahl.")
        df = df_source.copy()
        df["__batch_id"] = df[COL_ID].fillna("").astype(str).str.strip()
        eligible_df = df[df["__batch_id"] != ""].copy()
        already_df = eligible_df[eligible_df["__batch_id"].isin(self.state.batch_exported_ids)].copy()
        remaining_df = eligible_df[~eligible_df["__batch_id"].isin(self.state.batch_exported_ids)].copy()
        return df, eligible_df, already_df, remaining_df

    def _ensure_batch_export_tracker_ready(self) -> None:
        if self.state.batch_export_tracker_error:
            raise RuntimeError(
                "Batch-Export ist blockiert, weil die Batch-Merkliste beschädigt ist. "
                f"Bitte zuerst 'Batch-Merkliste zurücksetzen' ausführen. Details: {self.state.batch_export_tracker_error}"
            )

    def show_batch_export_window(self) -> None:
        open_batch_export_window(self)


def run_ui() -> None:
    """Startet die grafische Anwendung."""
    root = tk.Tk()
    TreeSolutionHelperUI(root)
    root.mainloop()

