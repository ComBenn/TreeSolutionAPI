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
    DEFAULT_USERS_FILE,
    DEFAULT_USERS_SHEET,
    DEFAULT_KEYWORDS_FILE,
    DEFAULT_OUTPUT_FILE,
)
from io_utils import load_table, load_keywords_txt
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
        self.flag_cache: dict[str, pd.DataFrame] = {}
        self.batch_export_tracker_file = self.runtime_dir / "batch_export_tracker.json"
        self.ui_state_file = self.runtime_dir / "ui_state.json"
        self.batch_exported_ids: set[str] = set()
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
        self.flag_cache = {}

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


class TreeSolutionUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("TreeSolution API Helper")
        self.root.geometry("1100x760")

        self.state = AppState()

        self.users_file_var = tk.StringVar(value=self.state.users_file)
        self.users_sheet_var = tk.StringVar(value=self.state.users_sheet or "")
        self.keywords_file_var = tk.StringVar(value=self.state.keywords_file)
        self.output_file_var = tk.StringVar(value=self.state.output_file)
        self.export_department_override_var = tk.StringVar(value="")
        self.batch_export_count_var = tk.StringVar(value="250")
        self.employee_list_file_var = tk.StringVar(value="")
        self.employee_list_sheet_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Bereit.")

        self._build_ui()
        self._refresh_status()
        self._load_ui_state()
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
        tk.Button(actions, text="Nur technische auswählen", command=self.keep_technical, width=24).grid(row=1, column=0, padx=4, pady=4, sticky="w")
        tk.Button(actions, text="Technische ausschliessen", command=self.exclude_technical, width=24).grid(row=1, column=1, padx=4, pady=4, sticky="w")

        employee = tk.LabelFrame(self.root, text="Mitarbeiterliste", padx=10, pady=10)
        employee.pack(fill="x", padx=10, pady=(0, 10))

        self._build_file_row(employee, "Liste", self.employee_list_file_var, self._pick_employee_list_file, row=0, on_enter=self.mark_employee_list)
        self._build_entry_row(employee, "Liste Sheet", self.employee_list_sheet_var, row=1, on_enter=self.mark_employee_list)

        tk.Button(employee, text="Nach Liste filtern", command=self.mark_employee_list, width=28).grid(row=0, column=3, padx=4, pady=4, sticky="w")

        kw_box = tk.LabelFrame(self.root, text="Keywords", padx=10, pady=10)
        kw_box.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(kw_box, text="Keyword-Datei öffnen", command=self.show_keywords, width=22).grid(row=0, column=0, padx=4, pady=4, sticky="w")

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

        if users_file:
            self.users_file_var.set(users_file)
        if users_sheet is not None:
            self.users_sheet_var.set(str(users_sheet))
        if keywords_file:
            self.keywords_file_var.set(keywords_file)
        if output_file:
            self.output_file_var.set(output_file)
        self.export_department_override_var.set(export_department)
        self._sync_state_paths()

    def _save_ui_state(self) -> None:
        self._sync_state_paths()
        payload = {
            "users_file": self.users_file_var.get().strip(),
            "users_sheet": self.users_sheet_var.get().strip(),
            "keywords_file": self.keywords_file_var.get().strip(),
            "output_file": self.output_file_var.get().strip(),
            "export_department_override": self.export_department_override_var.get().strip(),
        }
        try:
            self.state.ui_state_file.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

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
        return build_upload_export(df_source, department_override=department_override or None)

    def _log_export_result(self, rows: int) -> None:
        department_override = self.export_department_override_var.get().strip()
        if department_override:
            self._log(
                f"Export geschrieben: {self.state.output_file} | Zeilen: {rows} | "
                f"Department Override: {department_override}"
            )
        else:
            self._log(f"Export geschrieben: {self.state.output_file} | Zeilen: {rows}")

    def _export_regular_from_df(self, df_source: pd.DataFrame, initialfile: str = "Upload.csv") -> None:
        save_path = self._ask_export_save_path(initialfile=initialfile)
        if not save_path:
            self._log("Export abgebrochen.")
            return
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

    def _refresh_technical_flags_from_keywords(self) -> int:
        self._sync_state_paths()
        if self.state.original_df is None:
            raise RuntimeError("Zuerst Benutzerdatei laden.")
        keywords = load_keywords_txt(self.state.keywords_file)
        marked_df = mark_technical_accounts(self.state.original_df, keywords)
        self.state.flag_cache["flag_technical_account"] = marked_df.copy()
        self.state.current_df = marked_df.copy()
        hits = int(marked_df["flag_technical_account"].sum())
        self._log(f"Technische Markierung aktualisiert. Keywords: {len(keywords)} | Treffer: {hits}")
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

    def load_users(self) -> None:
        def _run() -> None:
            self._sync_state_paths()
            self.state.load_users()
            self._log(f"Benutzer geladen: {len(self.state.current_df)} aus {self.state.users_file}")
            self._refresh_technical_flags_from_keywords()
            self.preview_current()
        self._with_errors(_run)

    def reset_users(self) -> None:
        def _run() -> None:
            self.state.reset()
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

    def mark_employee_list(self) -> None:
        def _run() -> None:
            if self.state.original_df is None:
                self._sync_state_paths()
                self.state.load_users()
                self._log(f"Benutzer automatisch geladen: {len(self.state.current_df)} aus {self.state.users_file}")
                self._refresh_technical_flags_from_keywords()
            df = self.state.original_df
            list_file = self.employee_list_file_var.get().strip()
            if not list_file:
                raise RuntimeError("Bitte Mitarbeiterliste auswählen.")
            list_sheet = self.employee_list_sheet_var.get().strip() or None
            if Path(list_file).suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
                list_sheet = None

            flag_name = "flag_employee_list"
            df_list = load_table(list_file, list_sheet)
            marked_df, stats = mark_by_employee_list(
                df, df_list, flag_name=flag_name, return_stats=True
            )
            self.state.flag_cache[flag_name] = marked_df.copy()
            self.state.current_df = marked_df.copy()
            hits = int(marked_df[flag_name].sum())
            self._log(f"Mitarbeiterliste geprüft ({flag_name}). Treffer: {hits}")
            self._log(
                "Listeneinträge ohne Treffer: "
                f"{stats['employee_entries_unmatched']} "
                f"(gematcht: {stats['employee_entries_matched']} / berücksichtigt: {stats['employee_entries_total']})"
            )

            self.state.current_df = marked_df[marked_df[flag_name] == True].copy()
            self._log(f"Nach Liste gefiltert. Verbleibend: {len(self.state.current_df)}")
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

    def show_current_table(self) -> None:
        df = self.state.current_df
        if df is None:
            messagebox.showinfo("Keine Daten", "Noch keine Benutzerdatei geladen.")
            return
        view_state = {"df": df.copy()}
        previous_output_csv = self.output_file_var.get().strip()
        self.output_file_var.set("Upload.csv")
        self.state.output_file = "Upload.csv"

        win = tk.Toplevel(self.root)
        win.title(f"Aktuelle Auswahl ({len(df)} Zeilen)")
        win.geometry("1200x700")
        self._make_modal(win)

        container = tk.Frame(win, padx=8, pady=8)
        container.pack(fill="both", expand=True)

        info = tk.Label(container, text=f"{len(df)} Zeilen | {len(df.columns)} Spalten", anchor="w")
        info.pack(fill="x", pady=(0, 6))

        export_controls = tk.LabelFrame(container, text="Export für diese Auswahl", padx=8, pady=8)
        export_controls.pack(fill="x", pady=(0, 8))
        self._build_entry_row(
            export_controls,
            "Output CSV",
            self.output_file_var,
            row=0,
            on_enter=lambda: self._with_errors(lambda: self._export_regular_from_df(view_state["df"], initialfile="Upload.csv")),
        )
        self._build_entry_row(
            export_controls,
            "Export department",
            self.export_department_override_var,
            row=1,
            on_enter=lambda: self._with_errors(lambda: self._export_regular_from_df(view_state["df"])),
        )
        tk.Button(
            export_controls,
            text="Exportieren",
            command=lambda: self._with_errors(lambda: self._export_regular_from_df(view_state["df"], initialfile="Upload.csv")),
            width=28,
        ).grid(row=0, column=3, padx=4, pady=4, sticky="w")

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

        def refresh_selection_table() -> None:
            current_df = view_state["df"]
            tree.delete(*tree.get_children())
            for i, (_, row) in enumerate(current_df.fillna("").astype(str).iterrows()):
                values = [row.get(col, "") for col in columns]
                tree.insert("", "end", iid=str(i), values=values, tags=("odd" if i % 2 else "even",))
            self.state.current_df = current_df.copy()
            self._refresh_status()

        menu = tk.Menu(win, tearoff=0)

        def remove_selected_entries() -> None:
            selected = tree.selection()
            if not selected:
                return
            positions = sorted({int(iid) for iid in selected if str(iid).isdigit()}, reverse=True)
            if not positions:
                return
            current_df = view_state["df"]
            if current_df.empty:
                return
            valid_positions = [p for p in positions if 0 <= p < len(current_df)]
            if not valid_positions:
                return
            drop_idx = current_df.iloc[valid_positions].index
            view_state["df"] = current_df.drop(index=drop_idx).copy()
            self.state.current_df = view_state["df"].copy()
            self._log(f"Einträge aus Auswahl entfernt: {len(valid_positions)} | Verbleibend: {len(view_state['df'])}")
            refresh_selection_table()

        menu.add_command(label="Eintrag entfernen", command=lambda: self._with_errors(remove_selected_entries))

        def show_context_menu(event) -> None:
            row_id = tree.identify_row(event.y)
            if row_id:
                current_selection = tree.selection()
                if row_id not in current_selection:
                    tree.selection_set(row_id)
                tree.focus(row_id)
            if tree.selection():
                menu.tk_popup(event.x_root, event.y_root)
            menu.grab_release()

        tree.bind("<Button-3>", show_context_menu)
        refresh_selection_table()

        def _restore_output_field_on_close() -> None:
            self.output_file_var.set(previous_output_csv)
            win.destroy()

        footer = tk.Frame(container)
        footer.pack(fill="x", pady=(8, 0))
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
        df_snapshot = df.copy()
        batch_size_var = tk.StringVar(value=str(len(df_snapshot)))
        previous_output_csv = self.output_file_var.get().strip()
        self.output_file_var.set("Batch-Upload.csv")
        self.state.output_file = "Batch-Upload.csv"

        win = tk.Toplevel(self.root)
        win.title(f"Batch-Export ({len(df_snapshot)} Zeilen in aktueller Auswahl)")
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
        self._build_entry_row(controls, "Export department", self.export_department_override_var, row=1)
        self._build_entry_row(controls, "Batch-Grösse", batch_size_var, row=2)

        stats_var = tk.StringVar(value="")
        details_var = tk.StringVar(value="")
        tracker_var = tk.StringVar(value="")
        tk.Label(controls, textvariable=stats_var, anchor="w").grid(row=3, column=0, columnspan=3, padx=4, pady=(8, 2), sticky="w")
        tk.Label(controls, textvariable=details_var, anchor="w").grid(row=4, column=0, columnspan=3, padx=4, pady=2, sticky="w")
        tk.Label(controls, textvariable=tracker_var, anchor="w").grid(row=5, column=0, columnspan=3, padx=4, pady=2, sticky="w")

        table_frame = tk.Frame(container)
        table_frame.pack(fill="both", expand=True)

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

        def refresh_view() -> None:
            _, eligible_df, already_df, remaining_df = self._get_batch_remaining_df(df_snapshot)
            try:
                batch_size = int(batch_size_var.get().strip())
            except ValueError:
                raise RuntimeError("Batch-Grösse muss eine ganze Zahl sein.")
            if batch_size <= 0:
                raise RuntimeError("Batch-Grösse muss grösser als 0 sein.")

            selected_batch_df = remaining_df.head(batch_size).copy()
            tree.delete(*tree.get_children())
            display_df = selected_batch_df.drop(columns=["__batch_id"], errors="ignore")
            for i, (_, row) in enumerate(display_df.fillna("").astype(str).iterrows()):
                values = [row.get(col, "") for col in columns]
                tree.insert("", "end", values=values, tags=("odd" if i % 2 else "even",))

            stats_var.set(
                f"Aktuelle Auswahl: {len(df_snapshot)} | Mit ID (batch-fähig): {len(eligible_df)} | "
                f"Noch nicht exportiert: {len(remaining_df)}"
            )
            details_var.set(
                f"Batch-Auswahl aktuell angezeigt: {len(selected_batch_df)} | "
                f"Bereits exportiert (in dieser Auswahl): {len(already_df)} | "
                f"Batch-Merkliste gesamt: {len(self.state.batch_exported_ids)} IDs"
            )
            tracker_var.set(f"Merkliste-Datei: {self.state.batch_export_tracker_file.name}")

        def run_batch_export() -> None:
            self.batch_export_count_var.set(batch_size_var.get())
            self.output_file_var.set("Batch-Upload.csv")
            self.state.output_file = "Batch-Upload.csv"
            self._with_errors(lambda: self._export_next_batch_from_df(df_snapshot))
            refresh_view()

        def run_reset_and_refresh() -> None:
            self.reset_batch_export_tracker()
            refresh_view()

        def _restore_output_field_on_close() -> None:
            # Restore the main UI field value after closing the batch window.
            self.output_file_var.set(previous_output_csv)
            win.destroy()

        footer = tk.Frame(container)
        footer.pack(fill="x", pady=(8, 0))
        tk.Button(footer, text="Abbrechen", command=_restore_output_field_on_close, width=16).pack(side="right")

        win.protocol("WM_DELETE_WINDOW", _restore_output_field_on_close)

        tk.Button(
            controls,
            text="Batch exportieren",
            command=run_batch_export,
            width=36,
        ).grid(row=0, column=3, padx=4, pady=4, sticky="w")
        tk.Button(
            controls,
            text="Anzeige aktualisieren",
            command=lambda: self._with_errors(refresh_view),
            width=36,
        ).grid(row=1, column=3, padx=4, pady=4, sticky="w")
        tk.Button(
            controls,
            text="Batch-Merkliste zurücksetzen",
            command=run_reset_and_refresh,
            width=36,
        ).grid(row=2, column=3, padx=4, pady=4, sticky="w")

        # Enter in batch fields refreshes the proposed batch preview.
        for child in controls.winfo_children():
            if isinstance(child, tk.Entry):
                child.bind("<Return>", lambda _e: self._with_errors(refresh_view))

        self._with_errors(refresh_view)


def run_ui() -> None:
    root = tk.Tk()
    TreeSolutionUI(root)
    root.mainloop()

