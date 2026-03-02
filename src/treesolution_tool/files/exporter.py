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


def build_upload_export(df: pd.DataFrame, department_override: str | None = None) -> pd.DataFrame:
    """
    Erzeugt einen Export auf Basis der importierten Struktur:
    - alle vorhandenen Spalten bleiben erhalten (Reihenfolge bleibt erhalten)
    - Standardfelder werden bei Bedarf gesetzt/ueberschrieben
    - fehlende Upload-Spalten werden ergaenzt
    """
    required = [COL_ID, COL_USERNAME, COL_EMAIL, COL_FIRSTNAME, COL_LASTNAME]
    require_columns(df, required, "Exportquelle")

    out = df.copy()
    out = out.fillna("")

    # Standardspalten als String normalisieren (falls vorhanden)
    for col in (COL_ID, COL_USERNAME, COL_EMAIL, COL_FIRSTNAME, COL_LASTNAME):
        out[col] = out[col].fillna("").astype(str)

    if COL_INSTITUTION not in out.columns:
        out[COL_INSTITUTION] = ""
    out[COL_INSTITUTION] = EXPORT_INSTITUTION_VALUE

    if COL_DEPARTMENT not in out.columns:
        out[COL_DEPARTMENT] = ""
    if department_override is not None and str(department_override).strip() != "":
        out[COL_DEPARTMENT] = str(department_override).strip()
    else:
        out[COL_DEPARTMENT] = out[COL_DEPARTMENT].fillna("").astype(str)

    if COL_AUTH not in out.columns:
        out[COL_AUTH] = ""
    out[COL_AUTH] = EXPORT_AUTH_VALUE

    return out


def export_utf8_csv(df: pd.DataFrame, path: str):
    """
    UTF-8 mit BOM (utf-8-sig), damit Excel Umlaute sauber erkennt.
    """
    df.to_csv(path, index=False, encoding="utf-8-sig")
