# filters_technical.py

import pandas as pd
import re
from io_utils import norm_text, is_numeric_string, require_columns
from config import COL_FIRSTNAME, COL_ID, COL_LASTNAME


MIN_SUBSTRING_KEYWORD_LEN = 5


def _contains_keyword_token(text: str, keywords: set[str]) -> str | None:
    """
    Liefert das erste Keyword zurueck, das als eigenstaendiges Token in text vorkommt.
    Trennzeichen sind alle nicht-alphanumerischen Zeichen.
    """
    if not text or not keywords:
        return None
    tokens = [t for t in re.split(r"[\W_]+", text, flags=re.UNICODE) if t]
    for token in tokens:
        if token in keywords:
            return token
    return None


def _contains_keyword_substring(text: str, keywords: list[str]) -> str | None:
    """
    Liefert das erste laengere Keyword zurueck, das als Teilstring in text vorkommt.
    Kurze Keywords bleiben bei exakten/Token-Treffern, um False Positives zu begrenzen.
    """
    if not text or not keywords:
        return None
    for keyword in keywords:
        if keyword in text:
            return keyword
    return None


def _fullname_variants(firstname: str, lastname: str) -> set[str]:
    first = norm_text(firstname)
    last = norm_text(lastname)
    variants = set()
    if first and last:
        variants.add(f"{first} {last}")
        variants.add(f"{last} {first}")
    return variants


def mark_technical_accounts(df: pd.DataFrame, keywords: set[str]) -> pd.DataFrame:
    """
    Markiert technische Accounts per:
    - exakter Match id gegen Keywordliste
    - exakter Match firstname gegen Keywordliste
    - exakter Match lastname gegen Keywordliste
    - Keyword als Token in firstname/lastname (z.B. "hiller admin" enthaelt "admin")
    - firstname ist reine Zahl
    - lastname ist reine Zahl
    """
    require_columns(df, [COL_ID, COL_FIRSTNAME, COL_LASTNAME], "Benutzerdatei")
    out = df.copy()
    substring_keywords = sorted(
        (kw for kw in keywords if len(kw) >= MIN_SUBSTRING_KEYWORD_LEN),
        key=len,
        reverse=True,
    )

    flags = []
    reasons = []

    for _, row in out.iterrows():
        uid = norm_text(row.get(COL_ID, ""))
        fn = norm_text(row.get(COL_FIRSTNAME, ""))
        ln = norm_text(row.get(COL_LASTNAME, ""))
        fullname_variants = _fullname_variants(fn, ln)

        row_reasons = []

        if uid in keywords:
            row_reasons.append(f"exact_id:{uid}")
        else:
            substring_uid = _contains_keyword_substring(uid, substring_keywords)
            if substring_uid:
                row_reasons.append(f"substring_id:{substring_uid}")
        if fn in keywords:
            row_reasons.append(f"exact_firstname:{fn}")
        if ln in keywords:
            row_reasons.append(f"exact_lastname:{ln}")
        matched_fullname = next((name for name in fullname_variants if name in keywords), None)
        if matched_fullname:
            row_reasons.append(f"exact_fullname:{matched_fullname}")
        if fn not in keywords:
            token_fn = _contains_keyword_token(fn, keywords)
            if token_fn:
                row_reasons.append(f"token_firstname:{token_fn}")
            else:
                substring_fn = _contains_keyword_substring(fn, substring_keywords)
                if substring_fn:
                    row_reasons.append(f"substring_firstname:{substring_fn}")
        if ln not in keywords:
            token_ln = _contains_keyword_token(ln, keywords)
            if token_ln:
                row_reasons.append(f"token_lastname:{token_ln}")
            else:
                substring_ln = _contains_keyword_substring(ln, substring_keywords)
                if substring_ln:
                    row_reasons.append(f"substring_lastname:{substring_ln}")
        if is_numeric_string(fn):
            row_reasons.append(f"numeric_firstname:{fn}")
        if is_numeric_string(ln):
            row_reasons.append(f"numeric_lastname:{ln}")

        flags.append(len(row_reasons) > 0)
        reasons.append(" | ".join(row_reasons))

    out["flag_technical_account"] = flags
    out["flag_technical_reason"] = reasons
    return out
