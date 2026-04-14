from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from config import COL_EMAIL, COL_FIRSTNAME, COL_ID, COL_LASTNAME, COL_USERNAME


def _normalize_sort_value(value: str) -> tuple[int, object]:
    text = str(value).strip()
    if text == "":
        return (2, "")
    try:
        return (0, float(text.replace(",", ".")))
    except ValueError:
        return (1, text.casefold())


def _filter_row_records(row_records: list[dict], filter_column: str, filter_text: str) -> list[dict]:
    needle = str(filter_text).strip().casefold()
    if not needle:
        return list(row_records)
    return [
        record
        for record in row_records
        if needle in str(record["data"].get(filter_column, "")).casefold()
    ]


def _sort_row_records(row_records: list[dict], sort_column: str, descending: bool) -> list[dict]:
    return sorted(
        row_records,
        key=lambda record: _normalize_sort_value(record["data"].get(sort_column, "")),
        reverse=descending,
    )


def open_duplicate_review_dialog(ui) -> None:
    marked_df = ui._get_marked_duplicate_df()
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

    win = tk.Toplevel(ui.root)
    win.title(f"Duplikate prüfen ({len(group_ids)} Gruppen)")
    win.geometry("1320x760")
    ui._make_modal(win)

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

    sort_state = {"column": "flag_duplicate_group", "descending": False}

    def _heading_text(col: str) -> str:
        text = labels.get(col, col)
        if sort_state["column"] == col:
            return f"{text} {'▼' if sort_state['descending'] else '▲'}"
        return text

    for col in columns:
        tree.heading(col, text=_heading_text(col), command=lambda c=col: None)
        width = 130
        if col == "__exclude":
            width = 110
        elif col == "flag_duplicate_reason":
            width = 240
        tree.column(col, width=width, minwidth=90, stretch=True)

    tree.tag_configure("odd", background="#f6f8fb")
    tree.tag_configure("even", background="#ffffff")
    filter_status_var = tk.StringVar(value="Kein Filter aktiv")
    tk.Label(container, textvariable=filter_status_var, anchor="w", justify="left").pack(fill="x", pady=(8, 0))

    row_state: dict[str, dict] = {}
    row_records: list[dict] = []
    for i, (_, row) in enumerate(duplicate_df.fillna("").astype(str).iterrows()):
        row_id = str(row.get(COL_ID, "")).strip()
        if not row_id:
            continue
        excluded = row_id in ui.duplicate_excluded_ids
        iid = f"dup-{i}"
        data = {
            "id": row_id,
            "flag_duplicate_group": str(row.get("flag_duplicate_group", "")).strip(),
            COL_ID: row_id,
            COL_USERNAME: row.get(COL_USERNAME, ""),
            COL_EMAIL: row.get(COL_EMAIL, ""),
            COL_FIRSTNAME: row.get(COL_FIRSTNAME, ""),
            COL_LASTNAME: row.get(COL_LASTNAME, ""),
            "flag_duplicate_reason": row.get("flag_duplicate_reason", ""),
        }
        record = {
            "iid": iid,
            "id": row_id,
            "data": data,
            "excluded": excluded,
            "group": data["flag_duplicate_group"],
        }
        row_records.append(record)
        row_state[iid] = record

    def _record_values(record: dict) -> list[str]:
        return [
            "☑" if record["excluded"] else "☐",
            record["data"]["flag_duplicate_group"],
            record["data"][COL_ID],
            record["data"][COL_USERNAME],
            record["data"][COL_EMAIL],
            record["data"][COL_FIRSTNAME],
            record["data"][COL_LASTNAME],
            record["data"]["flag_duplicate_reason"],
        ]

    active_filter = {"column": "", "text": ""}

    def _refresh_table() -> None:
        filtered_records = _filter_row_records(row_records, active_filter["column"], active_filter["text"])
        visible_records = _sort_row_records(
            filtered_records,
            sort_state["column"],
            bool(sort_state["descending"]),
        )
        tree.delete(*tree.get_children())
        for i, record in enumerate(visible_records):
            tree.insert(
                "",
                "end",
                iid=record["iid"],
                values=_record_values(record),
                tags=("odd" if i % 2 else "even",),
            )
        for col in columns:
            tree.heading(col, text=_heading_text(col))

    def _update_filter_status() -> None:
        if active_filter["column"] and active_filter["text"]:
            filter_status_var.set(
                f"Filter aktiv: {labels.get(active_filter['column'], active_filter['column'])} enthält '{active_filter['text']}'"
            )
        else:
            filter_status_var.set("Kein Filter aktiv")

    def _clear_filter() -> None:
        active_filter["column"] = ""
        active_filter["text"] = ""
        _update_filter_status()
        _refresh_table()

    def _on_sort_column(col: str) -> None:
        if sort_state["column"] == col:
            sort_state["descending"] = not bool(sort_state["descending"])
        else:
            sort_state["column"] = col
            sort_state["descending"] = False
        _refresh_table()

    for col in columns:
        tree.heading(col, text=_heading_text(col), command=lambda c=col: _on_sort_column(c))

    def _toggle_selected_row(event=None) -> None:
        current_iid = tree.focus()
        if event is not None and not current_iid:
            current_iid = tree.identify_row(event.y)
        if not current_iid or current_iid not in row_state:
            return
        row_state[current_iid]["excluded"] = not bool(row_state[current_iid]["excluded"])
        _refresh_table()

    def _toggle_checkbox_click(event) -> None:
        iid = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if iid and col == "#1":
            tree.focus(iid)
            tree.selection_set(iid)
            _toggle_selected_row()

    header_menu = tk.Menu(win, tearoff=0)

    def _open_filter_dialog(column: str) -> None:
        if column == "__exclude":
            return
        current_text = active_filter["text"] if active_filter["column"] == column else ""
        filter_value = simpledialog.askstring(
            "Filter setzen",
            f"Filter für Spalte '{labels.get(column, column)}' eingeben:",
            initialvalue=current_text,
            parent=win,
        )
        if filter_value is None:
            return
        filter_value = str(filter_value).strip()
        if not filter_value:
            _clear_filter()
            return
        active_filter["column"] = column
        active_filter["text"] = filter_value
        _update_filter_status()
        _refresh_table()

    def _show_header_menu(event) -> None:
        if tree.identify_region(event.x, event.y) != "heading":
            return
        column_ref = tree.identify_column(event.x)
        if not column_ref or not column_ref.startswith("#"):
            return
        column_index = int(column_ref[1:]) - 1
        if column_index < 0 or column_index >= len(columns):
            return
        column = columns[column_index]
        if column == "__exclude":
            return
        header_menu.delete(0, "end")
        header_menu.add_command(
            label=f"Filter für {labels.get(column, column)} setzen",
            command=lambda c=column: _open_filter_dialog(c),
        )
        header_menu.add_command(label="Filter entfernen", command=_clear_filter)
        header_menu.tk_popup(event.x_root, event.y_root)

    tree.bind("<Button-1>", _toggle_checkbox_click, add="+")
    tree.bind("<Button-3>", _show_header_menu, add="+")
    tree.bind("<Double-1>", _toggle_selected_row)
    tree.bind("<space>", _toggle_selected_row)

    _refresh_table()
    _update_filter_status()

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
            group_records = [record for record in row_records if record["group"] == group]
            group_records.sort(
                key=lambda record: (
                    str(record["data"][COL_ID]),
                    str(record["data"][COL_USERNAME]),
                    str(record["data"][COL_EMAIL]),
                    str(record["data"][COL_FIRSTNAME]),
                )
            )
            for pos, record in enumerate(group_records):
                record["excluded"] = pos > 0
        _refresh_table()

    def _open_selection_in_export_window(export_excluded: bool) -> None:
        selected_ids = [
            record["data"]["id"]
            for record in row_state.values()
            if bool(record["excluded"]) == export_excluded and str(record["data"]["id"]).strip()
        ]
        if not selected_ids:
            if export_excluded:
                raise RuntimeError("Es ist kein ausgeschlossenes Duplikat für den Export ausgewählt.")
            raise RuntimeError("Es ist kein eingeschlossenes Duplikat für den Export ausgewählt.")
        export_df = duplicate_df[
            duplicate_df[COL_ID].fillna("").astype(str).str.strip().isin(selected_ids)
        ].copy()
        ui.show_current_table(
            df_override=export_df,
            title_override=(
                f"Ausgeschlossene Duplikate ({len(export_df)} Zeilen)"
                if export_excluded
                else f"Eingeschlossene Duplikate ({len(export_df)} Zeilen)"
            ),
            sync_state=False,
        )

    def _save_selection() -> None:
        _validate_selection()
        ui.duplicate_excluded_ids = {
            state["id"]
            for state in row_state.values()
            if bool(state["excluded"]) and str(state["id"]).strip()
        }
        ui._ensure_duplicate_template_present(marked_df=marked_df)
        ui._save_ui_state()
        ui._refresh_employee_templates_view()

        df_base = ui._ensure_original_users_loaded()
        all_indices = list(range(len(ui.employee_list_templates)))
        selected_df, include_count, exclude_count = ui._apply_employee_templates(
            df_base,
            all_indices,
            label="Alle Vorlagen",
        )
        ui.state.current_df = selected_df
        ui._log(
            f"Duplikat-Entscheidungen gespeichert. Gruppen: {len(group_ids)} | "
            f"Ausgeschlossene IDs: {len(ui.duplicate_excluded_ids)} | "
            f"Vorlagen aktiv: {len(all_indices)} | Einschliessen: {include_count} | Ausschliessen: {exclude_count}"
        )
        ui.preview_current()
        win.destroy()

    tk.Button(
        footer,
        text="Alle außer erster ausschließen",
        command=lambda: ui._with_errors(_exclude_all_but_first),
        width=28,
    ).pack(side="left")
    tk.Button(
        footer,
        text="Eingeschlossene exportieren",
        command=lambda: ui._with_errors(lambda: _open_selection_in_export_window(False)),
        width=26,
    ).pack(side="left", padx=(6, 0))
    tk.Button(
        footer,
        text="Ausgeschlossene exportieren",
        command=lambda: ui._with_errors(lambda: _open_selection_in_export_window(True)),
        width=26,
    ).pack(side="left", padx=(6, 0))
    tk.Button(
        footer,
        text="Speichern",
        command=lambda: ui._with_errors(_save_selection),
        width=14,
    ).pack(side="right", padx=(6, 0))
    tk.Button(footer, text="Abbrechen", command=win.destroy, width=14).pack(side="right")
