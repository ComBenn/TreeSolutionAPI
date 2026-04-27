from __future__ import annotations

import pandas as pd

from exporter import build_upload_export


def build_export_df(df_source: pd.DataFrame, department_overrides: list[str]) -> pd.DataFrame:
    """Normalisiert die UI-Parameter und delegiert den finalen Exportaufbau."""
    department_override = department_overrides[0] if department_overrides else None
    return build_upload_export(
        df_source,
        department_override=department_override or None,
        department_overrides=department_overrides or None,
    )


def format_export_log_message(output_file: str, rows: int, department_overrides: list[str]) -> str:
    """Erzeugt die kompakte Log-Zeile fuer regulären und Batch-Export."""
    if department_overrides:
        return (
            f"Export geschrieben: {output_file} | Zeilen: {rows} | "
            f"Departments: {', '.join(department_overrides)}"
        )
    return f"Export geschrieben: {output_file} | Zeilen: {rows}"
