# io_utils.py

from pathlib import Path
import pandas as pd
import re


def norm_text(v) -> str:
    """Normalisiert Text fuer case-insensitive Vergleiche im Projekt."""
    if pd.isna(v):
        return ""
    return str(v).strip().casefold()


def is_numeric_string(v: str) -> bool:
    """Prueft, ob ein Wert ausschliesslich aus Ziffern besteht."""
    return bool(re.fullmatch(r"\d+", v)) if v else False


def load_table(path: str, sheet_name: str | None = None) -> pd.DataFrame:
    """Laedt CSV- oder Excel-Dateien tolerant und liefert immer Strings zurueck."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {path}")

    if p.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                # sep=None lets pandas sniff delimiters like ; , tab automatically.
                return pd.read_csv(p, dtype=str, encoding=enc, sep=None, engine="python").fillna("")
            except UnicodeDecodeError:
                continue
        raise UnicodeDecodeError("utf-8", b"", 0, 1, f"CSV konnte nicht gelesen werden: {path}")

    # pandas returns a dict for Excel when sheet_name=None; force a single sheet DataFrame.
    excel_sheet = sheet_name if (sheet_name is not None and str(sheet_name).strip() != "") else 0
    return pd.read_excel(p, sheet_name=excel_sheet, dtype=str).fillna("")


def require_columns(df: pd.DataFrame, cols: list[str], context: str):
    """Bricht frueh mit einer klaren Meldung bei fehlenden Pflichtspalten ab."""
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Fehlende Spalten in {context}: {missing}")


def load_keywords_txt(path: str) -> set[str]:
    """Laedt die Keyword-Datei zeilenweise und normalisiert auf lowercase."""
    p = Path(path)
    if not p.exists():
        p.write_text("", encoding="utf-8")
        return set()

    keywords = set()
    text = _read_text_with_fallbacks(p)
    for line in text.splitlines():
        k = line.strip()
        if k:
            keywords.add(k.casefold())
    return keywords


def append_keywords_txt(path: str, new_keywords: list[str]) -> int:
    """Ergaenzt neue Keywords ohne bestehende Eintraege doppelt zu schreiben."""
    p = Path(path)
    existing = load_keywords_txt(path)
    cleaned_to_add = []

    for k in new_keywords:
        k2 = k.strip()
        if not k2:
            continue
        if k2.casefold() not in existing:
            cleaned_to_add.append(k2)
            existing.add(k2.casefold())

    if cleaned_to_add:
        prefix = "\n" if p.exists() and p.stat().st_size > 0 else ""
        with open(p, "a", encoding="utf-8") as f:
            f.write(prefix + "\n".join(cleaned_to_add))

    return len(cleaned_to_add)


def _read_text_with_fallbacks(path: Path) -> str:
    """Liest Textdateien mit den im Projekt ueblichen Fallback-Encodings."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("utf-8", b"", 0, 1, f"Datei konnte nicht gelesen werden: {path}")
