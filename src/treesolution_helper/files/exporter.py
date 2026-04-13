# exporter.py

import pandas as pd
from config import (
    COL_ID,
    COL_USERNAME,
    COL_EMAIL,
    COL_FIRSTNAME,
    COL_LASTNAME,
    COL_INSTITUTION,
    COL_DEPARTMENT,
    COL_AUTH,
    EXPORT_INSTITUTION_VALUE,
    EXPORT_AUTH_VALUE,
)
from io_utils import require_columns


def _is_technical_export_column(col_name: str) -> bool:
    name = str(col_name).strip().lower()
    return name.startswith("flag_") or name.startswith("__")


def build_upload_export(
    df: pd.DataFrame,
    department_override: str | None = None,
    department_overrides: list[str] | None = None,
) -> pd.DataFrame:
    """
    Erzeugt einen Export auf Basis der importierten Struktur:
    - fachliche Spalten bleiben erhalten (Reihenfolge bleibt erhalten)
    - technische Hilfsspalten (z.B. flag_*, __*) werden entfernt
    - Standardfelder werden bei Bedarf gesetzt/ueberschrieben
    - fehlende Upload-Spalten werden ergaenzt
    """
    required = [COL_ID, COL_USERNAME, COL_EMAIL, COL_FIRSTNAME, COL_LASTNAME]
    require_columns(df, required, "Exportquelle")

    keep_cols = [c for c in df.columns if not _is_technical_export_column(c)]
    out = df.loc[:, keep_cols].copy()
    out = out.fillna("")

    # Standardspalten als String normalisieren (falls vorhanden)
    for col in (COL_ID, COL_USERNAME, COL_EMAIL, COL_FIRSTNAME, COL_LASTNAME):
        out[col] = out[col].fillna("").astype(str)

    if COL_INSTITUTION not in out.columns:
        out[COL_INSTITUTION] = ""
    out[COL_INSTITUTION] = EXPORT_INSTITUTION_VALUE

    override_values = [str(v).strip() for v in (department_overrides or []) if str(v).strip()]
    if not override_values and department_override is not None and str(department_override).strip() != "":
        override_values = [str(department_override).strip()]

    if override_values:
        out = out.drop(columns=[COL_DEPARTMENT], errors="ignore")
        for idx, value in enumerate(override_values, start=1):
            out[f"{COL_DEPARTMENT}{idx}"] = value
    else:
        if COL_DEPARTMENT not in out.columns:
            out[COL_DEPARTMENT] = ""
        out[COL_DEPARTMENT] = out[COL_DEPARTMENT].fillna("").astype(str)

    if COL_AUTH not in out.columns:
        out[COL_AUTH] = ""
    out[COL_AUTH] = EXPORT_AUTH_VALUE

    return out


def export_utf8_csv(df: pd.DataFrame, path: str):
    """
    UTF-8 mit BOM (utf-8-sig), damit Excel Umlaute sauber erkennt.
    """
    df.to_csv(path, index=False, encoding="utf-8-sig", sep=";")
