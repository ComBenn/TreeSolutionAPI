from __future__ import annotations

import re
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd

from config import COL_DEPARTMENT, COL_EMAIL, COL_FIRSTNAME, COL_ID, COL_LASTNAME, COL_USERNAME


_DEPARTMENT_COLUMN_RE = re.compile(r"^department(\d+)?$", re.IGNORECASE)
_DUPLICATES_DEPARTMENT = "duplicates"


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


def _department_column_sort_key(column: str) -> tuple[int, int, str]:
    text = str(column).strip()
    if text.casefold() == COL_DEPARTMENT.casefold():
        return (0, 0, "")
    match = _DEPARTMENT_COLUMN_RE.fullmatch(text)
    if match and match.group(1):
        return (1, int(match.group(1)), "")
    return (2, 0, text.casefold())


def _extract_department_values(row: dict) -> list[str]:
    department_columns = sorted(
        [str(col) for col in row.keys() if _DEPARTMENT_COLUMN_RE.fullmatch(str(col).strip())],
        key=_department_column_sort_key,
    )
    values: list[str] = []
    seen: set[str] = set()
    for column in department_columns:
        value = str(row.get(column, "")).strip()
        if not value:
            continue
        value_cf = value.casefold()
        if value_cf in seen:
            continue
        seen.add(value_cf)
        values.append(value)
    return values


def _resolve_initial_excluded_ids(
    row_records: list[dict],
    saved_excluded_ids: set[str],
    reviewed_ids: set[str],
) -> set[str]:
    initial_excluded_ids: set[str] = set()
    for record in row_records:
        row_id = str(record.get("id", "")).strip()
        if not row_id:
            continue
        if row_id in reviewed_ids:
            if row_id in saved_excluded_ids:
                initial_excluded_ids.add(row_id)
            continue
        department_values = record.get("department_values", [])
        if any(str(value).strip().casefold() == _DUPLICATES_DEPARTMENT for value in department_values):
            initial_excluded_ids.add(row_id)
    return initial_excluded_ids


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
            "Es koennen auch alle Accounts einer Duplikat-Gruppe ausgeschlossen werden. "
            "Accounts mit Department 'duplicates' sind vorausgewaehlt."
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
        "departments",
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
        "departments": "Departments",
        COL_FIRSTNAME: "firstname",
        COL_LASTNAME: "lastname",
        "flag_duplicate_reason": "Grund",
    }
    info_var = tk.StringVar(value=f"{len(duplicate_df)} Zeilen gesamt | {len(group_ids)} Gruppen")
    tk.Label(container, textvariable=info_var, anchor="w").pack(fill="x", pady=(0, 6))

    table_toolbar = tk.Frame(container)
    table_toolbar.pack(fill="x", pady=(0, 8))
    tk.Label(
        table_toolbar,
        text="Sortieren: Klick auf Spaltenkopf | Filtern: Rechtsklick auf Spaltenkopf",
        anchor="w",
    ).pack(side="left")

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

    sort_state = {"column": "flag_duplicate_group", "ascending": True}

    def _heading_text(col: str) -> str:
        text = labels.get(col, col)
        if sort_state["column"] == col:
            return f"{text} {'▲' if sort_state['ascending'] else '▼'}"
        return text

    for col in columns:
        tree.heading(col, text=_heading_text(col), command=lambda c=col: None)
        width = 130
        if col == "__exclude":
            width = 110
        elif col == "departments":
            width = 240
        elif col == "flag_duplicate_reason":
            width = 240
        tree.column(col, width=width, minwidth=90, stretch=True)

    tree.tag_configure("odd", background="#f6f8fb")
    tree.tag_configure("even", background="#ffffff")

    row_state: dict[str, dict] = {}
    row_records: list[dict] = []
    for i, (_, row) in enumerate(duplicate_df.fillna("").astype(str).iterrows()):
        row_id = str(row.get(COL_ID, "")).strip()
        if not row_id:
            continue
        source_row = row.to_dict()
        department_values = _extract_department_values(source_row)
        iid = f"dup-{i}"
        data = {
            "id": row_id,
            "flag_duplicate_group": str(row.get("flag_duplicate_group", "")).strip(),
            COL_ID: row_id,
            COL_USERNAME: row.get(COL_USERNAME, ""),
            COL_EMAIL: row.get(COL_EMAIL, ""),
            "departments": ", ".join(department_values),
            COL_FIRSTNAME: row.get(COL_FIRSTNAME, ""),
            COL_LASTNAME: row.get(COL_LASTNAME, ""),
            "flag_duplicate_reason": row.get("flag_duplicate_reason", ""),
        }
        record = {
            "iid": iid,
            "id": row_id,
            "data": data,
            "excluded": False,
            "group": data["flag_duplicate_group"],
            "department_values": department_values,
            "source_row": source_row,
        }
        row_records.append(record)
        row_state[iid] = record

    current_duplicate_ids = {
        str(record["id"]).strip()
        for record in row_records
        if str(record["id"]).strip()
    }
    saved_excluded_ids = {
        str(row_id).strip()
        for row_id in ui.duplicate_excluded_ids
        if str(row_id).strip() and str(row_id).strip() in current_duplicate_ids
    }
    reviewed_ids = {
        str(row_id).strip()
        for row_id in getattr(ui, "duplicate_reviewed_ids", set())
        if str(row_id).strip() and str(row_id).strip() in current_duplicate_ids
    }
    initial_excluded_ids = _resolve_initial_excluded_ids(row_records, saved_excluded_ids, reviewed_ids)
    for record in row_records:
        record["excluded"] = str(record["id"]).strip() in initial_excluded_ids

    source_rows = []
    for record in row_records:
        source_rows.append(
            {
                "__iid": record["iid"],
                "__exclude": "1" if record["excluded"] else "0",
                "flag_duplicate_group": record["data"]["flag_duplicate_group"],
                COL_ID: record["data"][COL_ID],
                COL_USERNAME: record["data"][COL_USERNAME],
                COL_EMAIL: record["data"][COL_EMAIL],
                "departments": record["data"]["departments"],
                COL_FIRSTNAME: record["data"][COL_FIRSTNAME],
                COL_LASTNAME: record["data"][COL_LASTNAME],
                "flag_duplicate_reason": record["data"]["flag_duplicate_reason"],
            }
        )
    if source_rows:
        source_df = pd.DataFrame(source_rows).set_index("__iid", drop=False)
    else:
        source_df = pd.DataFrame(columns=["__iid", *columns]).set_index("__iid", drop=False)

    def _record_values(record: dict) -> list[str]:
        return [
            "☑" if record["excluded"] else "☐",
            record["data"]["flag_duplicate_group"],
            record["data"][COL_ID],
            record["data"][COL_USERNAME],
            record["data"][COL_EMAIL],
            record["data"]["departments"],
            record["data"][COL_FIRSTNAME],
            record["data"][COL_LASTNAME],
            record["data"]["flag_duplicate_reason"],
        ]

    view_state = {
        "source_df": source_df,
        "filters": {},
        "sort_col": "flag_duplicate_group",
        "sort_asc": True,
    }

    def _update_source_df_for_record(record: dict) -> None:
        iid = str(record["iid"])
        if iid in view_state["source_df"].index:
            view_state["source_df"].at[iid, "__exclude"] = "1" if record["excluded"] else "0"

    def _refresh_table() -> None:
        display_df = ui._filter_and_sort_df(
            view_state["source_df"],
            view_state["filters"],
            view_state["sort_col"],
            view_state["sort_asc"],
        )
        visible_records = [
            row_state[str(iid)]
            for iid in display_df["__iid"].fillna("").astype(str).tolist()
            if str(iid) in row_state
        ]
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
        filter_info = ""
        if view_state["filters"]:
            parts = [f"{labels.get(col, col)}='{term}'" for col, term in view_state["filters"].items()]
            filter_info = " | Filter: " + " ; ".join(parts[:3])
            if len(parts) > 3:
                filter_info += f" (+{len(parts) - 3})"
        info_var.set(
            f"{len(row_records)} Zeilen gesamt | "
            f"{len(visible_records)} Zeilen angezeigt | "
            f"{len(group_ids)} Gruppen"
            f"{filter_info}"
        )

    def _clear_filter() -> None:
        view_state["filters"].clear()
        _refresh_table()

    def _on_sort_column(col: str) -> None:
        if view_state["sort_col"] == col:
            view_state["sort_asc"] = not bool(view_state["sort_asc"])
        else:
            view_state["sort_col"] = col
            view_state["sort_asc"] = True
        sort_state["column"] = str(view_state["sort_col"])
        sort_state["ascending"] = bool(view_state["sort_asc"])
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
        _update_source_df_for_record(row_state[current_iid])
        _refresh_table()

    def _toggle_checkbox_click(event) -> None:
        iid = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if iid and col == "#1":
            tree.focus(iid)
            tree.selection_set(iid)
            _toggle_selected_row()

    header_menu = tk.Menu(win, tearoff=0)
    active_header_col = {"name": None}

    def _open_header_filter_dialog(column: str) -> None:
        if not column or column == "__exclude":
            return
        ui._open_contains_filter_dialog(
            parent=win,
            col=column,
            source_df=view_state["source_df"],
            filters=view_state["filters"],
            refresh_callback=_refresh_table,
            columns=columns,
        )

    def _clear_header_filter(column: str) -> None:
        if not column:
            return
        view_state["filters"].pop(column, None)
        _refresh_table()

    def _show_header_menu(event) -> None:
        column = ui._column_from_tree_event(tree, columns, event)
        if not column or column == "__exclude":
            return
        active_header_col["name"] = column
        header_menu.delete(0, "end")
        header_menu.add_command(
            label="Filter setzen...",
            command=lambda: ui._with_errors(
                lambda: _open_header_filter_dialog(str(active_header_col["name"] or ""))
            ),
        )
        header_menu.add_command(
            label="Filter dieser Spalte loeschen",
            command=lambda: ui._with_errors(
                lambda: _clear_header_filter(str(active_header_col["name"] or ""))
            ),
        )
        header_menu.tk_popup(event.x_root, event.y_root)
        header_menu.grab_release()

    tree.bind("<Button-1>", _toggle_checkbox_click, add="+")
    tree.bind("<Button-3>", _show_header_menu, add="+")
    tree.bind("<Double-1>", _toggle_selected_row)
    tree.bind("<space>", _toggle_selected_row)

    _refresh_table()

    tk.Button(
        table_toolbar,
        text="Alle Filter loeschen",
        command=lambda: ui._with_errors(_clear_filter),
        width=18,
    ).pack(side="right")

    footer = tk.Frame(container)
    footer.pack(fill="x", pady=(8, 0))

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
                _update_source_df_for_record(record)
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
        selected_excluded_ids = {
            state["id"]
            for state in row_state.values()
            if bool(state["excluded"]) and str(state["id"]).strip()
        }
        ui.duplicate_excluded_ids = {
            row_id
            for row_id in ui.duplicate_excluded_ids
            if row_id not in current_duplicate_ids
        }
        ui.duplicate_excluded_ids.update(selected_excluded_ids)
        ui.duplicate_reviewed_ids = {
            row_id
            for row_id in getattr(ui, "duplicate_reviewed_ids", set())
            if row_id not in current_duplicate_ids
        }
        ui.duplicate_reviewed_ids.update(current_duplicate_ids)
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
