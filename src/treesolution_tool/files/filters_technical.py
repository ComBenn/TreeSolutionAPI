# filters_technical.py

import pandas as pd
from io_utils import norm_text, is_numeric_string, require_columns
from config import COL_FIRSTNAME, COL_LASTNAME


def mark_technical_accounts(df: pd.DataFrame, keywords: set[str]) -> pd.DataFrame:
    """
    Markiert technische Accounts per:
    - exakter Match firstname gegen Keywordliste
    - exakter Match lastname gegen Keywordliste
    - firstname ist reine Zahl
    - lastname ist reine Zahl
    """
    require_columns(df, [COL_FIRSTNAME, COL_LASTNAME], "Benutzerdatei")
    out = df.copy()

    flags = []
    reasons = []

    for _, row in out.iterrows():
        fn = norm_text(row.get(COL_FIRSTNAME, ""))
        ln = norm_text(row.get(COL_LASTNAME, ""))

        row_reasons = []

        if fn in keywords:
            row_reasons.append(f"exact_firstname:{fn}")
        if ln in keywords:
            row_reasons.append(f"exact_lastname:{ln}")
        if is_numeric_string(fn):
            row_reasons.append(f"numeric_firstname:{fn}")
        if is_numeric_string(ln):
            row_reasons.append(f"numeric_lastname:{ln}")

        flags.append(len(row_reasons) > 0)
        reasons.append(" | ".join(row_reasons))

    out["flag_technical_account"] = flags
    out["flag_technical_reason"] = reasons
    return out