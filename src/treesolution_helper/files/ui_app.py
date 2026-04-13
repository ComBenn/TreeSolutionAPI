from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from config import (
    COL_ID,
    COL_EMAIL,
    COL_FIRSTNAME,
    COL_LASTNAME,
    COL_USERNAME,
    DEFAULT_USERS_FILE,
    DEFAULT_USERS_SHEET,
    DEFAULT_KEYWORDS_FILE,
    DEFAULT_OUTPUT_FILE,
)
from io_utils import load_table, load_keywords_txt
from filters_duplicates import mark_duplicate_accounts
from filters_technical import mark_technical_accounts
from filters_employee_list import mark_by_employee_list
from exporter import build_upload_export, export_utf8_csv


def _bundle_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _app_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class AppState:
    def __init__(self) -> None:
        self.bundle_dir = _bundle_base_dir()
        self.runtime_dir = _app_runtime_dir()
        self.runtime_dir.mkdir(parents=True, exist_ok=True)

        self.users_file = str(self.runtime_dir / DEFAULT_USERS_FILE)
        self.users_sheet = DEFAULT_USERS_SHEET
        self.keywords_file = str(self.runtime_dir / DEFAULT_KEYWORDS_FILE)
        self.output_file = str(self.runtime_dir / DEFAULT_OUTPUT_FILE)
        self.original_df: pd.DataFrame | None = None
        self.current_df: pd.DataFrame | None = None
        self.batch_export_tracker_file = self.runtime_dir / "batch_export_tracker.json"
        self.ui_state_file = self.runtime_dir / "ui_state.json"
        self.batch_exported_ids: set[str] = set()
        self._seed_runtime_file("README.md")
        self._seed_runtime_file("keywords_technische_accounts.txt")
        self._seed_runtime_file("batch_export_tracker.json", default_text='{"exported_ids": []}\n')
        self._seed_runtime_file("ui_state.json", default_text="{}\n")
        self._load_batch_export_tracker()

    def _seed_runtime_file(self, filename: str, default_text: str | None = None) -> None:
        dst = self.runtime_dir / filename
        if dst.exists():
            return
        src = self.bundle_dir / filename
        try:
            if src.exists():
                shutil.copy2(src, dst)
            elif default_text is not None:
                dst.write_text(default_text, encoding="utf-8")
        except Exception:
            # Best effort only; file may be created later by normal app flow.
            pass

    def load_users(self) -> None:
        self.original_df = load_table(self.users_file, self.users_sheet)
        self.current_df = self.original_df.copy()

    def reset(self) -> None:
        if self.original_df is None:
            raise RuntimeError("Noch keine Benutzerdatei geladen.")
        self.current_df = self.original_df.copy()

    def _load_batch_export_tracker(self) -> None:
        p = self.batch_export_tracker_file
        if not p.exists():
            self.batch_exported_ids = set()
            return
        try:
            payload = json.loads(p.read_text(encoding="utf-8"))
            ids = payload.get("exported_ids", [])
            self.batch_exported_ids = {str(x) for x in ids if str(x).strip()}
        except Exception:
            self.batch_exported_ids = set()

    def save_batch_export_tracker(self) -> None:
        p = self.batch_export_tracker_file
        payload = {"exported_ids": sorted(self.batch_exported_ids)}
        p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


class TreeSolutionHelperUI:
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

        self._build_ui()
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

    def _load_ui_state(self) -> None:
        p = self.state.ui_state_file
        if not p.exists():
            return
        try:
            payload = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
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
            "employee_list_templates": self.employee_list_templates,
        }
        try:
            self.state.ui_state_file.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _sanitize_employee_templates(self, templates_raw) -> list[dict]:
        out: list[dict] = []
        if not isinstance(templates_raw, list):
            return out
        for item in templates_raw:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            file_path = str(item.get("file", "")).strip()
            sheet = str(item.get("sheet", "")).strip()
            mode = str(item.get("mode", "include")).strip().casefold()
            kind = str(item.get("kind", "employee")).strip().casefold()
            readonly = bool(item.get("readonly", False))
            if not name or not file_path:
                continue
            if mode not in ("include", "exclude"):
                mode = "include"
            if kind not in ("employee", "technical", "duplicate"):
                kind = "employee"
            internal_ids_raw = item.get("internal_ids", [])
            internal_rows_raw = item.get("internal_rows", [])
            internal_ids = []
            if isinstance(internal_ids_raw, list):
                internal_ids = [str(x).strip() for x in internal_ids_raw if str(x).strip()]
            internal_rows = []
            if isinstance(internal_rows_raw, list):
                for r in internal_rows_raw:
                    if isinstance(r, dict):
                        internal_rows.append({str(k): str(v) for k, v in r.items()})
            out.append(
                {
                    "name": name,
                    "file": file_path,
                    "sheet": sheet,
                    "mode": mode,
                    "kind": kind,
                    "readonly": readonly,
                    "internal_ids": sorted(set(internal_ids)),
                    "internal_rows": internal_rows,
                }
            )
        return out

    def _find_template_index_by_name(self, name: str) -> int | None:
        name_cf = str(name).strip().casefold()
        for i, t in enumerate(self.employee_list_templates):
            if str(t.get("name", "")).strip().casefold() == name_cf:
                return i
        return None

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
        if COL_ID not in marked_df.columns:
            return [], [], 0
        if "flag_technical_account" not in marked_df.columns:
            return [], [], 0
        matched_df = marked_df[marked_df["flag_technical_account"] == True].copy()
        if matched_df.empty:
            return [], [], 0
        matched_df = matched_df.drop(
            columns=["flag_technical_account", "flag_technical_reason"],
            errors="ignore",
        ).fillna("").astype(str)
        ids = sorted(
            {
                str(v).strip()
                for v in matched_df[COL_ID].fillna("").astype(str)
                if str(v).strip()
            }
        )
        rows = matched_df.to_dict(orient="records")
        return ids, rows, len(matched_df)

    def _ensure_technical_template_present(self, marked_df: pd.DataFrame | None = None) -> None:
        ids, rows, hits = self._build_internal_technical_template_data(marked_df=marked_df)
        idx = self._find_template_index_by_name(self.technical_template_name)
        payload = {
            "name": self.technical_template_name,
            "file": "<auto:keywords_technische_accounts>",
            "sheet": "",
            "mode": "exclude",
            "kind": "technical",
            "readonly": True,
            "internal_ids": ids,
            "internal_rows": rows,
        }
        if idx is None:
            self.employee_list_templates.insert(0, payload)
            self._log(f"Auto-Vorlage bereitgestellt: {self.technical_template_name} | Treffer: {hits}")
        else:
            existing = self.employee_list_templates[idx]
            existing.update(payload)

    def _build_internal_duplicate_template_data(
        self,
        marked_df: pd.DataFrame | None = None,
    ) -> tuple[list[str], list[dict], int]:
        if self.state.original_df is None and marked_df is None:
            return [], [], 0
        if marked_df is None:
            marked_df = self._get_marked_duplicate_df()
        if COL_ID not in marked_df.columns:
            return [], [], 0
        matched_df = marked_df[
            marked_df[COL_ID].fillna("").astype(str).str.strip().isin(self.duplicate_excluded_ids)
        ].copy()
        if matched_df.empty:
            return [], [], 0
        matched_df = matched_df.fillna("").astype(str)
        ids = sorted(
            {
                str(v).strip()
                for v in matched_df[COL_ID].fillna("").astype(str)
                if str(v).strip()
            }
        )
        rows = matched_df.to_dict(orient="records")
        return ids, rows, len(matched_df)

    def _ensure_duplicate_template_present(self, marked_df: pd.DataFrame | None = None) -> None:
        ids, rows, hits = self._build_internal_duplicate_template_data(marked_df=marked_df)
        idx = self._find_template_index_by_name(self.duplicate_template_name)
        payload = {
            "name": self.duplicate_template_name,
            "file": "<auto:duplicate_review>",
            "sheet": "",
            "mode": "exclude",
            "kind": "duplicate",
            "readonly": True,
            "internal_ids": ids,
            "internal_rows": rows,
        }
        if idx is None:
            insert_at = 1 if self._find_template_index_by_name(self.technical_template_name) is not None else 0
            self.employee_list_templates.insert(insert_at, payload)
            self._log(f"Auto-Vorlage bereitgestellt: {self.duplicate_template_name} | Treffer: {hits}")
        else:
            existing = self.employee_list_templates[idx]
            existing.update(payload)

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
        df_base = self._ensure_original_users_loaded()
        if COL_ID not in df_base.columns:
            raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der Benutzerdatei.")
        df_list = load_table(file_path, sheet or None)
        flag_name = "__template_build_match"
        marked_df, _stats = mark_by_employee_list(df_base, df_list, flag_name=flag_name, return_stats=True)
        matched_df = marked_df[marked_df[flag_name] == True].copy()
        if matched_df.empty:
            return [], [], 0
        matched_df = matched_df.drop(columns=[flag_name, f"{flag_name}_reason"], errors="ignore")
        matched_df = matched_df.fillna("").astype(str)
        internal_ids = sorted(
            {
                str(v).strip()
                for v in matched_df[COL_ID].fillna("").astype(str)
                if str(v).strip()
            }
        )
        internal_rows = matched_df.to_dict(orient="records")
        return internal_ids, internal_rows, len(matched_df)

    def open_employee_template_dialog(self) -> None:
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
            self._sync_state_paths()
            self.state.load_users()
            self._log(
                f"Benutzer automatisch beim Start geladen: {len(self.state.current_df)} aus {self.state.users_file}"
            )
            self._refresh_technical_flags_from_keywords()
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
        department_override = self.export_department_override_var.get().strip()
        department_overrides = self._get_export_department_override_values()
        return build_upload_export(
            df_source,
            department_override=department_override or None,
            department_overrides=department_overrides or None,
        )

    def _log_export_result(self, rows: int) -> None:
        department_overrides = self._get_export_department_override_values()
        if department_overrides:
            self._log(
                f"Export geschrieben: {self.state.output_file} | Zeilen: {rows} | "
                f"Departments: {', '.join(department_overrides)}"
            )
        else:
            self._log(f"Export geschrieben: {self.state.output_file} | Zeilen: {rows}")

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
        if self.state.original_df is None:
            raise RuntimeError("Zuerst Benutzerdatei laden.")
        return mark_duplicate_accounts(self.state.original_df)

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
        sheet = (list_sheet or "").strip() or None
        if Path(list_file).suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
            return None
        return sheet

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
        if not template_indices:
            raise RuntimeError("Keine Vorlagen ausgewählt.")

        if COL_ID not in df_base.columns:
            raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der Benutzerdatei.")
        id_series = df_base[COL_ID].fillna("").astype(str).str.strip()
        include_mask = pd.Series(False, index=df_base.index)
        exclude_mask = pd.Series(False, index=df_base.index)
        include_count = 0
        exclude_count = 0

        for i in template_indices:
            template = self.employee_list_templates[i]
            ids_in_template = {
                str(v).strip()
                for v in template.get("internal_ids", [])
                if str(v).strip()
            }
            if not ids_in_template:
                rows = template.get("internal_rows", [])
                for row in rows if isinstance(rows, list) else []:
                    if isinstance(row, dict):
                        v = str(row.get(COL_ID, "")).strip()
                        if v:
                            ids_in_template.add(v)

            # Migration/Fallback: ältere Vorlagen ohne interne Liste neu aufbauen.
            if not ids_in_template:
                template_kind = str(template.get("kind", "employee"))
                if template_kind == "technical":
                    rebuilt_ids, rebuilt_rows, _hits = self._build_internal_technical_template_data()
                elif template_kind == "duplicate":
                    rebuilt_ids, rebuilt_rows, _hits = self._build_internal_duplicate_template_data()
                else:
                    file_path = str(template.get("file", "")).strip()
                    sheet = str(template.get("sheet", "")).strip()
                    rebuilt_ids, rebuilt_rows, _hits = self._build_internal_template_data(file_path, sheet) if file_path else ([], [], 0)
                ids_in_template = set(rebuilt_ids)
                template["internal_ids"] = rebuilt_ids
                template["internal_rows"] = rebuilt_rows

            row_mask = id_series.isin(ids_in_template)
            hits = int(row_mask.sum())
            mode = str(template.get("mode", "include"))
            mode_label = "einschliessen" if mode == "include" else "ausschliessen"
            self._log(
                f"{label} | Vorlage '{template['name']}' ({mode_label}) geprüft über interne Liste. "
                f"Interne IDs: {len(ids_in_template)} | Treffer in Benutzerdatei: {hits}"
            )
            if mode == "include":
                include_mask = include_mask | row_mask
                include_count += 1
            else:
                exclude_mask = exclude_mask | row_mask
                exclude_count += 1

        self._save_ui_state()

        selected_mask = ~exclude_mask
        if include_count > 0:
            selected_mask = selected_mask | include_mask
        selected_df = df_base[selected_mask].copy()
        return selected_df, include_count, exclude_count

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
            self._sync_state_paths()
            self.state.load_users()
            self._log(f"Benutzer geladen: {len(self.state.current_df)} aus {self.state.users_file}")
            self._refresh_auto_flags()
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
            marked_df = self._get_marked_duplicate_df()
            if "flag_duplicate" not in marked_df.columns:
                raise RuntimeError("Duplikate konnten nicht markiert werden.")

            duplicate_df = marked_df[marked_df["flag_duplicate"] == True].copy()
            if duplicate_df.empty:
                messagebox.showinfo("Keine Duplikate", "Es wurden keine Duplikate gefunden.")
                return
            if COL_ID not in duplicate_df.columns:
                raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der Benutzerdatei.")

            duplicate_df[COL_ID] = duplicate_df[COL_ID].fillna("").astype(str).str.strip()
            duplicate_df = duplicate_df.sort_values(
                by=["flag_duplicate_group", COL_LASTNAME, COL_FIRSTNAME, COL_EMAIL, COL_USERNAME],
                kind="stable",
            )
            group_ids = [g for g in duplicate_df["flag_duplicate_group"].dropna().astype(str).unique() if g.strip()]
            if not group_ids:
                messagebox.showinfo("Keine Duplikate", "Es wurden keine Dublettengruppen gefunden.")
                return

            win = tk.Toplevel(self.root)
            win.title(f"Duplikate prüfen ({len(group_ids)} Gruppen)")
            win.geometry("1320x760")
            self._make_modal(win)

            container = tk.Frame(win, padx=8, pady=8)
            container.pack(fill="both", expand=True)

            tk.Label(
                container,
                text=(
                    "Checkbox aktiviert = Account wird ausgeschlossen. "
                    "Pro Duplikat-Gruppe muss mindestens ein Eintrag aktiv bleiben."
                ),
                anchor="w",
                justify="left",
            ).pack(fill="x", pady=(0, 8))

            columns = [
                "__exclude",
                "flag_duplicate_group",
                COL_ID,
                COL_USERNAME,
                COL_EMAIL,
                COL_FIRSTNAME,
                COL_LASTNAME,
                "flag_duplicate_reason",
            ]
            labels = {
                "__exclude": "Ausschließen",
                "flag_duplicate_group": "Gruppe",
                COL_ID: "id",
                COL_USERNAME: "username",
                COL_EMAIL: "email",
                COL_FIRSTNAME: "firstname",
                COL_LASTNAME: "lastname",
                "flag_duplicate_reason": "Grund",
            }

            table_frame = tk.Frame(container)
            table_frame.pack(fill="both", expand=True)
            tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
            vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)

            for col in columns:
                tree.heading(col, text=labels.get(col, col))
                width = 130
                if col == "__exclude":
                    width = 110
                elif col == "flag_duplicate_reason":
                    width = 240
                tree.column(col, width=width, minwidth=90, stretch=True)

            tree.tag_configure("odd", background="#f6f8fb")
            tree.tag_configure("even", background="#ffffff")

            row_state: dict[str, dict] = {}
            for i, (_, row) in enumerate(duplicate_df.fillna("").astype(str).iterrows()):
                row_id = str(row.get(COL_ID, "")).strip()
                if not row_id:
                    continue
                excluded = row_id in self.duplicate_excluded_ids
                iid = f"dup-{i}"
                tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=[
                        "☑" if excluded else "☐",
                        row.get("flag_duplicate_group", ""),
                        row_id,
                        row.get(COL_USERNAME, ""),
                        row.get(COL_EMAIL, ""),
                        row.get(COL_FIRSTNAME, ""),
                        row.get(COL_LASTNAME, ""),
                        row.get("flag_duplicate_reason", ""),
                    ],
                    tags=("odd" if i % 2 else "even",),
                )
                row_state[iid] = {
                    "id": row_id,
                    "group": str(row.get("flag_duplicate_group", "")).strip(),
                    "excluded": excluded,
                }

            def _refresh_row(iid: str) -> None:
                state = row_state[iid]
                current_values = list(tree.item(iid, "values"))
                current_values[0] = "☑" if state["excluded"] else "☐"
                tree.item(iid, values=current_values)

            def _toggle_selected_row(event=None) -> None:
                current_iid = tree.focus()
                if event is not None and not current_iid:
                    current_iid = tree.identify_row(event.y)
                if not current_iid or current_iid not in row_state:
                    return
                row_state[current_iid]["excluded"] = not bool(row_state[current_iid]["excluded"])
                _refresh_row(current_iid)

            def _toggle_checkbox_click(event) -> None:
                iid = tree.identify_row(event.y)
                col = tree.identify_column(event.x)
                if iid and col == "#1":
                    tree.focus(iid)
                    tree.selection_set(iid)
                    _toggle_selected_row()

            tree.bind("<Button-1>", _toggle_checkbox_click, add="+")
            tree.bind("<Double-1>", _toggle_selected_row)
            tree.bind("<space>", _toggle_selected_row)

            footer = tk.Frame(container)
            footer.pack(fill="x", pady=(8, 0))

            def _validate_selection() -> None:
                active_per_group: dict[str, int] = {}
                for state in row_state.values():
                    group = str(state["group"]).strip()
                    if group and not bool(state["excluded"]):
                        active_per_group[group] = active_per_group.get(group, 0) + 1
                invalid_groups = [group for group in group_ids if active_per_group.get(group, 0) == 0]
                if invalid_groups:
                    raise RuntimeError(
                        "Mindestens ein Account pro Gruppe muss aktiv bleiben. "
                        f"Ungültige Gruppen: {', '.join(invalid_groups[:10])}"
                    )

            def _exclude_all_but_first() -> None:
                for group in group_ids:
                    group_iids = [iid for iid, state in row_state.items() if state["group"] == group]
                    group_iids.sort(key=lambda iid: tuple(str(v) for v in tree.item(iid, "values")[2:6]))
                    for pos, iid in enumerate(group_iids):
                        row_state[iid]["excluded"] = pos > 0
                        _refresh_row(iid)

            def _save_selection() -> None:
                _validate_selection()
                self.duplicate_excluded_ids = {
                    state["id"]
                    for state in row_state.values()
                    if bool(state["excluded"]) and str(state["id"]).strip()
                }
                self._ensure_duplicate_template_present(marked_df=marked_df)
                self._save_ui_state()
                self._refresh_employee_templates_view()

                df_base = self._ensure_original_users_loaded()
                all_indices = list(range(len(self.employee_list_templates)))
                selected_df, include_count, exclude_count = self._apply_employee_templates(
                    df_base,
                    all_indices,
                    label="Alle Vorlagen",
                )
                self.state.current_df = selected_df
                self._log(
                    f"Duplikat-Entscheidungen gespeichert. Gruppen: {len(group_ids)} | "
                    f"Ausgeschlossene IDs: {len(self.duplicate_excluded_ids)} | "
                    f"Vorlagen aktiv: {len(all_indices)} | Einschliessen: {include_count} | Ausschliessen: {exclude_count}"
                )
                self.preview_current()
                win.destroy()

            tk.Button(
                footer,
                text="Alle außer erster ausschließen",
                command=lambda: self._with_errors(_exclude_all_but_first),
                width=28,
            ).pack(side="left")
            tk.Button(
                footer,
                text="Speichern",
                command=lambda: self._with_errors(_save_selection),
                width=14,
            ).pack(side="right", padx=(6, 0))
            tk.Button(footer, text="Abbrechen", command=win.destroy, width=14).pack(side="right")
        self._with_errors(_run)

    def mark_employee_list(self) -> None:
        def _run() -> None:
            df_base = self._ensure_original_users_loaded()
            all_indices = list(range(len(self.employee_list_templates)))
            if not all_indices:
                raise RuntimeError("Keine Vorlagen vorhanden.")
            selected_df, include_count, exclude_count = self._apply_employee_templates(
                df_base,
                all_indices,
                label="Alle Vorlagen",
            )
            self.state.current_df = selected_df
            self._log(
                f"Alle Vorlagen angewendet. Anzahl Vorlagen: {len(all_indices)} | "
                f"Einschliessen: {include_count} | Ausschliessen: {exclude_count} | "
                f"Verbleibend: {len(selected_df)}"
            )
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
            if count_before == 0:
                self._log("Batch-Merkliste ist bereits leer.")
                return

            confirmed = messagebox.askyesno(
                "Batch-Merkliste zurücksetzen",
                "Alle gemerkten Batch-Export-IDs wirklich löschen?\n"
                "Danach können Einträge erneut über den Batch-Export exportiert werden.",
            )
            if not confirmed:
                self._log("Zurücksetzen der Batch-Merkliste abgebrochen.")
                return

            self.state.batch_exported_ids.clear()
            self.state.save_batch_export_tracker()
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
        df = df_override if df_override is not None else self.state.current_df
        if df is None:
            messagebox.showinfo("Keine Daten", "Noch keine Benutzerdatei geladen.")
            return
        view_state = {
            "base_df": df.copy(),
            "display_df": df.copy(),
            "sort_col": None,
            "sort_asc": True,
            "filters": {},
            "iid_to_index": {},
        }
        previous_output_csv = self.output_file_var.get().strip()
        self.output_file_var.set("Upload.csv")
        self.state.output_file = "Upload.csv"

        win = tk.Toplevel(self.root)
        win.title(title_override or f"Aktuelle Auswahl ({len(df)} Zeilen)")
        win.geometry("1200x700")
        self._make_modal(win)

        container = tk.Frame(win, padx=8, pady=8)
        container.pack(fill="both", expand=True)

        info_var = tk.StringVar(value=f"{len(df)} Zeilen | {len(df.columns)} Spalten")
        info = tk.Label(container, textvariable=info_var, anchor="w")
        info.pack(fill="x", pady=(0, 6))

        export_controls = tk.LabelFrame(container, text="Export für diese Auswahl", padx=8, pady=8)
        export_controls.pack(fill="x", pady=(0, 8))
        self._build_entry_row(
            export_controls,
            "Output CSV",
            self.output_file_var,
            row=0,
            on_enter=lambda: self._with_errors(lambda: self._export_regular_from_df(view_state["base_df"].copy(), initialfile="Upload.csv")),
        )
        self._build_department_override_controls(
            export_controls,
            start_row=1,
            on_enter=lambda: self._with_errors(lambda: self._export_regular_from_df(view_state["base_df"].copy())),
        )

        table_toolbar = tk.Frame(container)
        table_toolbar.pack(fill="x", pady=(0, 8))
        tk.Label(
            table_toolbar,
            text="Sortieren: Klick auf Spaltenkopf | Filtern: Rechtsklick auf Spaltenkopf (Dropdown)",
            anchor="w",
        ).pack(side="left")

        table_frame = tk.Frame(container)
        table_frame.pack(fill="both", expand=True)

        columns = [str(c) for c in df.columns]
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        for col in columns:
            tree.heading(col, text=col)
            width = min(max(100, len(col) * 10), 260)
            tree.column(col, width=width, minwidth=80, stretch=True)

        tree.tag_configure("odd", background="#f6f8fb")
        tree.tag_configure("even", background="#ffffff")

        def _apply_filters_and_sort() -> None:
            view_state["display_df"] = self._filter_and_sort_df(
                view_state["base_df"],
                view_state["filters"],
                view_state["sort_col"],
                view_state["sort_asc"],
            ).copy()

        def refresh_selection_table() -> None:
            _apply_filters_and_sort()
            display_df = view_state["display_df"]
            tree.delete(*tree.get_children())
            iid_to_index = {}
            for i, (idx, row) in enumerate(display_df.fillna("").astype(str).iterrows()):
                values = [row.get(col, "") for col in columns]
                iid = str(i)
                tree.insert("", "end", iid=iid, values=values, tags=("odd" if i % 2 else "even",))
                iid_to_index[iid] = idx
            view_state["iid_to_index"] = iid_to_index
            if sync_state:
                self.state.current_df = view_state["base_df"].copy()
            filter_info = ""
            if view_state["filters"]:
                parts = [f"{col}='{term}'" for col, term in view_state["filters"].items()]
                filter_info = " | Filter: " + " ; ".join(parts[:3])
                if len(parts) > 3:
                    filter_info += f" (+{len(parts) - 3})"
            info_var.set(
                f"{len(view_state['base_df'])} Zeilen gesamt | "
                f"{len(display_df)} Zeilen angezeigt | "
                f"{len(columns)} Spalten"
                f"{filter_info}"
            )
            self._refresh_status()

        row_menu = tk.Menu(win, tearoff=0)
        header_menu = tk.Menu(win, tearoff=0)
        active_header_col = {"name": None}

        def remove_selected_entries() -> None:
            selected = tree.selection()
            if not selected:
                return
            current_df = view_state["base_df"]
            if current_df.empty:
                return
            row_indices = [
                view_state["iid_to_index"].get(str(iid))
                for iid in selected
                if str(iid) in view_state["iid_to_index"]
            ]
            valid_indices = [idx for idx in row_indices if idx in current_df.index]
            if not valid_indices:
                return
            if delete_confirm_text:
                confirmed = messagebox.askyesno("Sicher löschen", delete_confirm_text)
                if not confirmed:
                    return
            selected_rows = current_df.loc[valid_indices].copy()
            if delete_callback is not None:
                delete_callback(selected_rows)
            view_state["base_df"] = current_df.drop(index=valid_indices).copy()
            if sync_state:
                self.state.current_df = view_state["base_df"].copy()
            self._log(f"Einträge aus Auswahl entfernt: {len(valid_indices)} | Verbleibend: {len(view_state['base_df'])}")
            refresh_selection_table()

        row_menu.add_command(label="Eintrag entfernen", command=lambda: self._with_errors(remove_selected_entries))

        def _open_header_filter_dialog(col: str) -> None:
            self._open_contains_filter_dialog(
                parent=win,
                col=col,
                source_df=view_state["base_df"],
                filters=view_state["filters"],
                refresh_callback=refresh_selection_table,
                columns=columns,
            )

        def _clear_header_filter(col: str) -> None:
            if not col:
                return
            view_state["filters"].pop(col, None)
            refresh_selection_table()

        header_menu.add_command(
            label="Filter setzen...",
            command=lambda: self._with_errors(lambda: _open_header_filter_dialog(str(active_header_col["name"] or ""))),
        )
        header_menu.add_command(
            label="Filter dieser Spalte löschen",
            command=lambda: self._with_errors(lambda: _clear_header_filter(str(active_header_col["name"] or ""))),
        )

        def sort_by_column(col: str) -> None:
            if view_state["sort_col"] == col:
                view_state["sort_asc"] = not view_state["sort_asc"]
            else:
                view_state["sort_col"] = col
                view_state["sort_asc"] = True
            refresh_selection_table()

        def _column_from_event(event) -> str | None:
            return self._column_from_tree_event(tree, columns, event)

        def show_context_menu(event) -> None:
            header_col = _column_from_event(event)
            if header_col:
                active_header_col["name"] = header_col
                header_menu.tk_popup(event.x_root, event.y_root)
                header_menu.grab_release()
                return
            row_id = tree.identify_row(event.y)
            if row_id:
                current_selection = tree.selection()
                if row_id not in current_selection:
                    tree.selection_set(row_id)
                tree.focus(row_id)
            if tree.selection():
                row_menu.tk_popup(event.x_root, event.y_root)
            row_menu.grab_release()

        tree.bind("<Button-3>", show_context_menu)
        self._bind_treeview_shortcuts(tree, columns)

        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self._with_errors(lambda: sort_by_column(c)))

        def clear_all_filters() -> None:
            view_state["filters"].clear()
            refresh_selection_table()

        tk.Button(
            table_toolbar,
            text="Alle Filter löschen",
            command=lambda: self._with_errors(clear_all_filters),
            width=18,
        ).pack(side="right")
        refresh_selection_table()

        def _restore_output_field_on_close() -> None:
            self.output_file_var.set(previous_output_csv)
            win.destroy()

        footer = tk.Frame(container)
        footer.pack(fill="x", pady=(8, 0))
        tk.Button(
            footer,
            text="Exportieren",
            command=lambda: self._with_errors(lambda: self._export_regular_from_df(view_state["base_df"].copy(), initialfile="Upload.csv")),
            width=16,
        ).pack(side="right", padx=(6, 0))
        tk.Button(footer, text="Abbrechen", command=_restore_output_field_on_close, width=16).pack(side="right")

        win.protocol("WM_DELETE_WINDOW", _restore_output_field_on_close)

    def _get_batch_remaining_df(
        self, df_source: pd.DataFrame
    ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        if COL_ID not in df_source.columns:
            raise RuntimeError(f"Spalte '{COL_ID}' fehlt in der aktuellen Auswahl.")
        df = df_source.copy()
        df["__batch_id"] = df[COL_ID].fillna("").astype(str).str.strip()
        eligible_df = df[df["__batch_id"] != ""].copy()
        already_df = eligible_df[eligible_df["__batch_id"].isin(self.state.batch_exported_ids)].copy()
        remaining_df = eligible_df[~eligible_df["__batch_id"].isin(self.state.batch_exported_ids)].copy()
        return df, eligible_df, already_df, remaining_df

    def show_batch_export_window(self) -> None:
        df = self.state.current_df
        if df is None:
            messagebox.showinfo("Keine Daten", "Noch keine Benutzerdatei geladen.")
            return
        _, eligible_df_start, already_df_start, remaining_df_start = self._get_batch_remaining_df(df)
        df_snapshot = remaining_df_start.drop(columns=["__batch_id"], errors="ignore").copy()
        if df_snapshot.empty:
            messagebox.showinfo(
                "Keine neuen IDs",
                "Es sind keine neuen batch-fähigen IDs mehr vorhanden.\n"
                "Alle IDs aus der aktuellen Auswahl wurden bereits exportiert oder haben keine ID.",
            )
            return
        batch_size_var = tk.StringVar(value=str(len(df_snapshot)))
        previous_output_csv = self.output_file_var.get().strip()
        self.output_file_var.set("Batch-Upload.csv")
        self.state.output_file = "Batch-Upload.csv"

        win = tk.Toplevel(self.root)
        win.title(f"Batch-Export ({len(df_snapshot)} Zeilen nach Ausschluss gemerkter IDs)")
        win.geometry("1250x760")
        self._make_modal(win)

        container = tk.Frame(win, padx=8, pady=8)
        container.pack(fill="both", expand=True)

        header = tk.Label(container, text="Aktuelle Auswahl als Batch exportieren", anchor="w")
        header.pack(fill="x")

        hint = tk.Label(
            container,
            text=(
                "Hinweis: Bereits exportierte Einträge werden über die ID dauerhaft gemerkt "
                "(auch nach dem Schliessen des Programms)."
            ),
            anchor="w",
            justify="left",
        )
        hint.pack(fill="x", pady=(4, 8))

        controls = tk.LabelFrame(container, text="Batch-Export", padx=8, pady=8)
        controls.pack(fill="x", pady=(0, 8))
        self._build_entry_row(controls, "Output CSV", self.output_file_var, row=0)
        self._build_department_override_controls(controls, start_row=1, on_enter=lambda: self._with_errors(refresh_view))
        self._build_entry_row(controls, "Batch-Grösse", batch_size_var, row=3)

        stats_var = tk.StringVar(value="")
        details_var = tk.StringVar(value="")
        tracker_var = tk.StringVar(value="")
        tk.Label(controls, textvariable=stats_var, anchor="w").grid(row=4, column=0, columnspan=3, padx=4, pady=(8, 2), sticky="w")
        tk.Label(controls, textvariable=details_var, anchor="w").grid(row=5, column=0, columnspan=3, padx=4, pady=2, sticky="w")
        tk.Label(controls, textvariable=tracker_var, anchor="w").grid(row=6, column=0, columnspan=3, padx=4, pady=2, sticky="w")

        table_toolbar = tk.Frame(container)
        table_toolbar.pack(fill="x", pady=(0, 8))
        tk.Label(
            table_toolbar,
            text="Sortieren: Klick auf Spaltenkopf | Filtern: Rechtsklick auf Spaltenkopf (Dropdown)",
            anchor="w",
        ).pack(side="left")

        table_frame = tk.Frame(container)
        table_frame.pack(fill="both", expand=True)

        batch_view_state = {
            "sort_col": None,
            "sort_asc": True,
            "filters": {},
            "last_source_df": pd.DataFrame(),
        }
        columns = [str(c) for c in df_snapshot.columns]
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        for col in columns:
            tree.heading(col, text=col)
            width = min(max(100, len(col) * 10), 260)
            tree.column(col, width=width, minwidth=80, stretch=True)
        tree.tag_configure("odd", background="#f6f8fb")
        tree.tag_configure("even", background="#ffffff")

        def _apply_filters_and_sort(df_in: pd.DataFrame) -> pd.DataFrame:
            return self._filter_and_sort_df(
                df_in,
                batch_view_state["filters"],
                batch_view_state["sort_col"],
                batch_view_state["sort_asc"],
            )

        def refresh_view() -> None:
            _, eligible_df, already_df, remaining_df = self._get_batch_remaining_df(df_snapshot)
            try:
                batch_size = int(batch_size_var.get().strip())
            except ValueError:
                raise RuntimeError("Batch-Grösse muss eine ganze Zahl sein.")
            if batch_size <= 0:
                raise RuntimeError("Batch-Grösse muss grösser als 0 sein.")

            selected_batch_df = remaining_df.head(batch_size).copy()
            source_df = selected_batch_df.drop(columns=["__batch_id"], errors="ignore")
            batch_view_state["last_source_df"] = source_df.copy()
            display_df = _apply_filters_and_sort(source_df)
            tree.delete(*tree.get_children())
            for i, (_, row) in enumerate(display_df.fillna("").astype(str).iterrows()):
                values = [row.get(col, "") for col in columns]
                tree.insert("", "end", values=values, tags=("odd" if i % 2 else "even",))

            stats_var.set(
                f"Ausgangsauswahl batch-fähig: {len(eligible_df_start)} | "
                f"Bereits exportiert: {len(already_df_start)} | "
                f"Noch nicht exportiert: {len(remaining_df)}"
            )
            filter_info = ""
            if batch_view_state["filters"]:
                parts = [f"{col}='{term}'" for col, term in batch_view_state["filters"].items()]
                filter_info = " | Filter: " + " ; ".join(parts[:3])
                if len(parts) > 3:
                    filter_info += f" (+{len(parts) - 3})"
            details_var.set(
                f"Batch-Auswahl: {len(source_df)} | Davon angezeigt: {len(display_df)} | "
                f"Bereits exportiert (in dieser Auswahl): {len(already_df)} | "
                f"Batch-Merkliste gesamt: {len(self.state.batch_exported_ids)} IDs"
                f"{filter_info}"
            )
            tracker_var.set(f"Merkliste-Datei: {self.state.batch_export_tracker_file.name}")

        active_header_col = {"name": None}
        header_menu = tk.Menu(win, tearoff=0)

        def _open_header_filter_dialog(col: str) -> None:
            self._open_contains_filter_dialog(
                parent=win,
                col=col,
                source_df=batch_view_state["last_source_df"],
                filters=batch_view_state["filters"],
                refresh_callback=refresh_view,
                columns=columns,
            )

        def _clear_header_filter(col: str) -> None:
            if not col:
                return
            batch_view_state["filters"].pop(col, None)
            refresh_view()

        header_menu.add_command(
            label="Filter setzen...",
            command=lambda: self._with_errors(lambda: _open_header_filter_dialog(str(active_header_col["name"] or ""))),
        )
        header_menu.add_command(
            label="Filter dieser Spalte löschen",
            command=lambda: self._with_errors(lambda: _clear_header_filter(str(active_header_col["name"] or ""))),
        )

        def sort_by_column(col: str) -> None:
            if batch_view_state["sort_col"] == col:
                batch_view_state["sort_asc"] = not batch_view_state["sort_asc"]
            else:
                batch_view_state["sort_col"] = col
                batch_view_state["sort_asc"] = True
            refresh_view()

        def _column_from_event(event) -> str | None:
            return self._column_from_tree_event(tree, columns, event)

        def show_header_menu(event) -> None:
            header_col = _column_from_event(event)
            if not header_col:
                return
            active_header_col["name"] = header_col
            header_menu.tk_popup(event.x_root, event.y_root)
            header_menu.grab_release()

        tree.bind("<Button-3>", show_header_menu)
        self._bind_treeview_shortcuts(tree, columns)
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self._with_errors(lambda: sort_by_column(c)))

        def clear_all_filters() -> None:
            batch_view_state["filters"].clear()
            refresh_view()

        tk.Button(
            table_toolbar,
            text="Alle Filter löschen",
            command=lambda: self._with_errors(clear_all_filters),
            width=18,
        ).pack(side="right")

        def run_batch_export() -> None:
            self.batch_export_count_var.set(batch_size_var.get())
            self.output_file_var.set("Batch-Upload.csv")
            self.state.output_file = "Batch-Upload.csv"
            self._with_errors(lambda: self._export_next_batch_from_df(df_snapshot))
            refresh_view()

        def _restore_output_field_on_close() -> None:
            # Restore the main UI field value after closing the batch window.
            self.output_file_var.set(previous_output_csv)
            win.destroy()

        footer = tk.Frame(container)
        footer.pack(fill="x", pady=(8, 0))
        tk.Button(footer, text="Batch exportieren", command=run_batch_export, width=16).pack(side="right", padx=(6, 0))
        tk.Button(footer, text="Abbrechen", command=_restore_output_field_on_close, width=16).pack(side="right")

        win.protocol("WM_DELETE_WINDOW", _restore_output_field_on_close)

        tk.Button(
            controls,
            text="Anzeige aktualisieren",
            command=lambda: self._with_errors(refresh_view),
            width=36,
        ).grid(row=3, column=3, padx=4, pady=4, sticky="w")

        # Enter in batch fields refreshes the proposed batch preview.
        for child in controls.winfo_children():
            if isinstance(child, tk.Entry):
                child.bind("<Return>", lambda _e: self._with_errors(refresh_view))

        self._with_errors(refresh_view)


def run_ui() -> None:
    root = tk.Tk()
    TreeSolutionHelperUI(root)
    root.mainloop()

