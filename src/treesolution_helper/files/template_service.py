from __future__ import annotations

from pathlib import Path
from typing import Callable

import pandas as pd

from config import COL_ID
from filters_employee_list import mark_by_employee_list
from io_utils import load_table


def sanitize_employee_templates(templates_raw) -> list[dict]:
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
            for row in internal_rows_raw:
                if isinstance(row, dict):
                    internal_rows.append({str(k): str(v) for k, v in row.items()})
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


def normalize_employee_list_sheet(list_file: str, list_sheet: str | None) -> str | None:
    sheet = (list_sheet or "").strip() or None
    if Path(list_file).suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
        return None
    return sheet


def build_internal_template_data(df_base: pd.DataFrame, file_path: str, sheet: str | None) -> tuple[list[str], list[dict], int]:
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


def apply_employee_templates(
    df_base: pd.DataFrame,
    templates: list[dict],
    template_indices: list[int],
    rebuild_callback: Callable[[dict], tuple[list[str], list[dict], int]],
    log_callback: Callable[[str], None] | None = None,
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
        template = templates[i]
        ids_in_template = {
            str(v).strip()
            for v in template.get("internal_ids", [])
            if str(v).strip()
        }
        if not ids_in_template:
            rows = template.get("internal_rows", [])
            for row in rows if isinstance(rows, list) else []:
                if isinstance(row, dict):
                    value = str(row.get(COL_ID, "")).strip()
                    if value:
                        ids_in_template.add(value)

        if not ids_in_template:
            rebuilt_ids, rebuilt_rows, _hits = rebuild_callback(template)
            ids_in_template = set(rebuilt_ids)
            template["internal_ids"] = rebuilt_ids
            template["internal_rows"] = rebuilt_rows

        row_mask = id_series.isin(ids_in_template)
        hits = int(row_mask.sum())
        mode = str(template.get("mode", "include"))
        mode_label = "einschliessen" if mode == "include" else "ausschliessen"
        if log_callback is not None:
            log_callback(
                f"Vorlage '{template['name']}' ({mode_label}) geprüft über interne Liste. "
                f"Interne IDs: {len(ids_in_template)} | Treffer in Benutzerdatei: {hits}"
            )
        if mode == "include":
            include_mask = include_mask | row_mask
            include_count += 1
        else:
            exclude_mask = exclude_mask | row_mask
            exclude_count += 1

    selected_mask = ~exclude_mask
    if include_count > 0:
        selected_mask = selected_mask | include_mask
    selected_df = df_base[selected_mask].copy()
    return selected_df, include_count, exclude_count
