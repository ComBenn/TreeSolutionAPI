from __future__ import annotations

import pandas as pd


def find_template_index_by_name(templates: list[dict], name: str) -> int | None:
    """Findet eine Vorlage per Name case-insensitive in der UI-Liste."""
    name_cf = str(name).strip().casefold()
    for i, template in enumerate(templates):
        if str(template.get("name", "")).strip().casefold() == name_cf:
            return i
    return None


def build_internal_technical_template_data(
    marked_df: pd.DataFrame,
    col_id: str,
) -> tuple[list[str], list[dict], int]:
    """Leitet aus technisch markierten Zeilen die interne Auto-Vorlage ab."""
    if col_id not in marked_df.columns:
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
            for v in matched_df[col_id].fillna("").astype(str)
            if str(v).strip()
        }
    )
    rows = matched_df.to_dict(orient="records")
    return ids, rows, len(matched_df)


def build_internal_duplicate_template_data(
    marked_df: pd.DataFrame,
    duplicate_excluded_ids: set[str],
    col_id: str,
) -> tuple[list[str], list[dict], int]:
    """Baut die interne Ausschluss-Vorlage aus den gemerkten Duplikat-IDs."""
    if col_id not in marked_df.columns:
        return [], [], 0
    matched_df = marked_df[
        marked_df[col_id].fillna("").astype(str).str.strip().isin(duplicate_excluded_ids)
    ].copy()
    if matched_df.empty:
        return [], [], 0
    matched_df = matched_df.fillna("").astype(str)
    ids = sorted(
        {
            str(v).strip()
            for v in matched_df[col_id].fillna("").astype(str)
            if str(v).strip()
        }
    )
    rows = matched_df.to_dict(orient="records")
    return ids, rows, len(matched_df)


def upsert_auto_template(
    templates: list[dict],
    template_name: str,
    file_marker: str,
    kind: str,
    ids: list[str],
    rows: list[dict],
    insert_at: int,
) -> tuple[bool, dict]:
    """Legt eine Auto-Vorlage an oder aktualisiert sie an ihrer festen Position."""
    payload = {
        "name": template_name,
        "file": file_marker,
        "sheet": "",
        "mode": "exclude",
        "kind": kind,
        "readonly": True,
        "internal_ids": ids,
        "internal_rows": rows,
    }
    idx = find_template_index_by_name(templates, template_name)
    if idx is None:
        templates.insert(insert_at, payload)
        return True, payload
    templates[idx].update(payload)
    return False, templates[idx]
