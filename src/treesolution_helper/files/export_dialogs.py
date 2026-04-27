from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd


def open_current_table_dialog(
    ui,
    df_override: pd.DataFrame | None = None,
    title_override: str | None = None,
    sync_state: bool = True,
    delete_callback=None,
    delete_confirm_text: str | None = None,
) -> None:
    """Oeffnet die allgemeine Tabellenansicht mit Filter-, Loesch- und Exportfunktionen."""
    df = df_override if df_override is not None else ui.state.current_df
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
    previous_output_csv = ui.output_file_var.get().strip()
    ui.output_file_var.set("Upload.csv")
    ui.state.output_file = "Upload.csv"

    win = tk.Toplevel(ui.root)
    win.title(title_override or f"Aktuelle Auswahl ({len(df)} Zeilen)")
    win.geometry("1200x700")
    ui._make_modal(win)

    container = tk.Frame(win, padx=8, pady=8)
    container.pack(fill="both", expand=True)

    info_var = tk.StringVar(value=f"{len(df)} Zeilen | {len(df.columns)} Spalten")
    info = tk.Label(container, textvariable=info_var, anchor="w")
    info.pack(fill="x", pady=(0, 6))

    export_controls = tk.LabelFrame(container, text="Export für diese Auswahl", padx=8, pady=8)
    export_controls.pack(fill="x", pady=(0, 8))
    ui._build_entry_row(
        export_controls,
        "Output CSV",
        ui.output_file_var,
        row=0,
        on_enter=lambda: ui._with_errors(lambda: ui._export_regular_from_df(view_state["base_df"].copy(), initialfile="Upload.csv")),
    )
    ui._build_department_override_controls(
        export_controls,
        start_row=1,
        on_enter=lambda: ui._with_errors(lambda: ui._export_regular_from_df(view_state["base_df"].copy())),
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
        view_state["display_df"] = ui._filter_and_sort_df(
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
            ui.state.current_df = view_state["base_df"].copy()
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
        ui._refresh_status()

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
            ui.state.current_df = view_state["base_df"].copy()
        ui._log(f"Einträge aus Auswahl entfernt: {len(valid_indices)} | Verbleibend: {len(view_state['base_df'])}")
        refresh_selection_table()

    row_menu.add_command(label="Eintrag entfernen", command=lambda: ui._with_errors(remove_selected_entries))

    def _open_header_filter_dialog(col: str) -> None:
        ui._open_contains_filter_dialog(
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
        command=lambda: ui._with_errors(lambda: _open_header_filter_dialog(str(active_header_col["name"] or ""))),
    )
    header_menu.add_command(
        label="Filter dieser Spalte löschen",
        command=lambda: ui._with_errors(lambda: _clear_header_filter(str(active_header_col["name"] or ""))),
    )

    def sort_by_column(col: str) -> None:
        if view_state["sort_col"] == col:
            view_state["sort_asc"] = not view_state["sort_asc"]
        else:
            view_state["sort_col"] = col
            view_state["sort_asc"] = True
        refresh_selection_table()

    def _column_from_event(event) -> str | None:
        return ui._column_from_tree_event(tree, columns, event)

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
    ui._bind_treeview_shortcuts(tree, columns)

    for col in columns:
        tree.heading(col, text=col, command=lambda c=col: ui._with_errors(lambda: sort_by_column(c)))

    def clear_all_filters() -> None:
        view_state["filters"].clear()
        refresh_selection_table()

    tk.Button(
        table_toolbar,
        text="Alle Filter löschen",
        command=lambda: ui._with_errors(clear_all_filters),
        width=18,
    ).pack(side="right")
    refresh_selection_table()

    def _restore_output_field_on_close() -> None:
        ui._reset_export_department_override_fields()
        ui.output_file_var.set(previous_output_csv)
        ui._save_ui_state()
        win.destroy()

    footer = tk.Frame(container)
    footer.pack(fill="x", pady=(8, 0))
    tk.Button(
        footer,
        text="Exportieren",
        command=lambda: ui._with_errors(lambda: ui._export_regular_from_df(view_state["base_df"].copy(), initialfile="Upload.csv")),
        width=16,
    ).pack(side="right", padx=(6, 0))
    tk.Button(footer, text="Abbrechen", command=_restore_output_field_on_close, width=16).pack(side="right")

    win.protocol("WM_DELETE_WINDOW", _restore_output_field_on_close)


def open_batch_export_window(ui) -> None:
    """Zeigt die noch nicht exportierten Batch-Kandidaten mit Exportsteuerung an."""
    df = ui.state.current_df
    if df is None:
        messagebox.showinfo("Keine Daten", "Noch keine Benutzerdatei geladen.")
        return
    _, eligible_df_start, already_df_start, remaining_df_start = ui._get_batch_remaining_df(df)
    df_snapshot = remaining_df_start.drop(columns=["__batch_id"], errors="ignore").copy()
    if df_snapshot.empty:
        messagebox.showinfo(
            "Keine neuen IDs",
            "Es sind keine neuen batch-fähigen IDs mehr vorhanden.\n"
            "Alle IDs aus der aktuellen Auswahl wurden bereits exportiert oder haben keine ID.",
        )
        return
    batch_size_var = tk.StringVar(value=str(len(df_snapshot)))
    previous_output_csv = ui.output_file_var.get().strip()
    ui.output_file_var.set("Batch-Upload.csv")
    ui.state.output_file = "Batch-Upload.csv"

    win = tk.Toplevel(ui.root)
    win.title(f"Batch-Export ({len(df_snapshot)} Zeilen nach Ausschluss gemerkter IDs)")
    win.geometry("1250x760")
    ui._make_modal(win)

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
    ui._build_entry_row(controls, "Output CSV", ui.output_file_var, row=0)
    ui._build_department_override_controls(controls, start_row=1, on_enter=lambda: ui._with_errors(refresh_view))
    ui._build_entry_row(controls, "Batch-Grösse", batch_size_var, row=3)

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
        return ui._filter_and_sort_df(
            df_in,
            batch_view_state["filters"],
            batch_view_state["sort_col"],
            batch_view_state["sort_asc"],
        )

    def refresh_view() -> None:
        _, eligible_df, already_df, remaining_df = ui._get_batch_remaining_df(df_snapshot)
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
            f"Batch-Merkliste gesamt: {len(ui.state.batch_exported_ids)} IDs"
            f"{filter_info}"
        )
        tracker_var.set(f"Merkliste-Datei: {ui.state.batch_export_tracker_file.name}")

    active_header_col = {"name": None}
    header_menu = tk.Menu(win, tearoff=0)

    def _open_header_filter_dialog(col: str) -> None:
        ui._open_contains_filter_dialog(
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
        command=lambda: ui._with_errors(lambda: _open_header_filter_dialog(str(active_header_col["name"] or ""))),
    )
    header_menu.add_command(
        label="Filter dieser Spalte löschen",
        command=lambda: ui._with_errors(lambda: _clear_header_filter(str(active_header_col["name"] or ""))),
    )

    def sort_by_column(col: str) -> None:
        if batch_view_state["sort_col"] == col:
            batch_view_state["sort_asc"] = not batch_view_state["sort_asc"]
        else:
            batch_view_state["sort_col"] = col
            batch_view_state["sort_asc"] = True
        refresh_view()

    def _column_from_event(event) -> str | None:
        return ui._column_from_tree_event(tree, columns, event)

    def show_header_menu(event) -> None:
        header_col = _column_from_event(event)
        if not header_col:
            return
        active_header_col["name"] = header_col
        header_menu.tk_popup(event.x_root, event.y_root)
        header_menu.grab_release()

    tree.bind("<Button-3>", show_header_menu)
    ui._bind_treeview_shortcuts(tree, columns)
    for col in columns:
        tree.heading(col, text=col, command=lambda c=col: ui._with_errors(lambda: sort_by_column(c)))

    def clear_all_filters() -> None:
        batch_view_state["filters"].clear()
        refresh_view()

    tk.Button(
        table_toolbar,
        text="Alle Filter löschen",
        command=lambda: ui._with_errors(clear_all_filters),
        width=18,
    ).pack(side="right")

    def run_batch_export() -> None:
        ui.batch_export_count_var.set(batch_size_var.get())
        ui.output_file_var.set("Batch-Upload.csv")
        ui.state.output_file = "Batch-Upload.csv"
        ui._with_errors(lambda: ui._export_next_batch_from_df(df_snapshot))
        refresh_view()

    def _restore_output_field_on_close() -> None:
        ui._reset_export_department_override_fields()
        ui.output_file_var.set(previous_output_csv)
        ui._save_ui_state()
        win.destroy()

    footer = tk.Frame(container)
    footer.pack(fill="x", pady=(8, 0))
    tk.Button(footer, text="Batch exportieren", command=run_batch_export, width=16).pack(side="right", padx=(6, 0))
    tk.Button(footer, text="Abbrechen", command=_restore_output_field_on_close, width=16).pack(side="right")

    win.protocol("WM_DELETE_WINDOW", _restore_output_field_on_close)

    tk.Button(
        controls,
        text="Anzeige aktualisieren",
        command=lambda: ui._with_errors(refresh_view),
        width=36,
    ).grid(row=3, column=3, padx=4, pady=4, sticky="w")

    for child in controls.winfo_children():
        if isinstance(child, tk.Entry):
            child.bind("<Return>", lambda _e: ui._with_errors(refresh_view))

    ui._with_errors(refresh_view)
