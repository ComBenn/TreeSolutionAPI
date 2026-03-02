# filters_employee_list.py

import pandas as pd
from io_utils import norm_text, build_fullname_key
from config import COL_EMAIL, COL_FIRSTNAME, COL_LASTNAME


def _canon_colname(name: str) -> str:
    text = str(name).replace("\ufeff", "").casefold().strip()
    out = []
    last_was_space = False
    for ch in text:
        if ch.isalnum():
            out.append(ch)
            last_was_space = False
        else:
            if not last_was_space:
                out.append(" ")
                last_was_space = True
    return "".join(out).strip()


def _detect_employee_list_columns(df_list: pd.DataFrame) -> dict[str, str | None]:
    """
    Erlaubt flexible Spaltennamen in der Mitarbeiterliste.
    Unterstützt z.B.:
    - email
    - lastname firstname
    - nachname / vorname
    - vorname / nachname
    - vorname + nachname (getrennte Spalten)
    """
    normalized = {_canon_colname(c): c for c in df_list.columns}
    detected = {
        "email": None,
        "combined_name": None,
        "firstname": None,
        "lastname": None,
    }

    for n, original in normalized.items():
        if n in ("email", "e mail", "mail"):
            detected["email"] = detected["email"] or original

        if n in (
            "lastname firstname",
            "last name first name",
            "nachname vorname",
            "name vorname",
            "vorname nachname",
            "first name last name",
            "firstname lastname",
        ):
            detected["combined_name"] = detected["combined_name"] or original

        if n in ("firstname", "first name", "vorname", "given name", "givenname"):
            detected["firstname"] = detected["firstname"] or original

        if n in ("lastname", "last name", "nachname", "surname", "family name", "familyname", "name"):
            detected["lastname"] = detected["lastname"] or original

    return detected


def _name_tokens(value) -> list[str]:
    text = norm_text(value)
    if not text:
        return []
    tokens = []
    current = []
    for ch in text:
        if ch.isalnum():
            current.append(ch)
        else:
            if current:
                tokens.append("".join(current))
                current = []
    if current:
        tokens.append("".join(current))
    return [t for t in tokens if t]


def _variants_from_combined_name(value) -> set[str]:
    tokens = _name_tokens(value)
    if not tokens:
        return set()

    variants = {" ".join(tokens)}
    if len(tokens) >= 2:
        variants.add(" ".join(reversed(tokens)))
    return {v.strip() for v in variants if v.strip()}


def _variants_from_first_last(first, last) -> set[str]:
    first_n = norm_text(first)
    last_n = norm_text(last)
    first_tokens = _name_tokens(first_n)
    last_tokens = _name_tokens(last_n)
    if not first_tokens and not last_tokens:
        return set()

    first_clean = " ".join(first_tokens).strip()
    last_clean = " ".join(last_tokens).strip()
    first_initial = first_tokens[0][0] if first_tokens and first_tokens[0] else ""

    variants = set()
    if last_clean and first_clean:
        variants.add(f"{last_clean} {first_clean}".strip())
        variants.add(f"{first_clean} {last_clean}".strip())
    if last_clean and first_initial:
        variants.add(f"{last_clean} {first_initial}".strip())
        variants.add(f"{first_initial} {last_clean}".strip())
    if last_clean:
        variants.add(last_clean)
    if first_clean:
        variants.add(first_clean)
    return {v.strip() for v in variants if v.strip()}


def mark_by_employee_list(
    df_users: pd.DataFrame,
    df_employee_list: pd.DataFrame,
    flag_name: str = "flag_employee_list",
    return_stats: bool = False,
):
    """
    Markiert Benutzer, die in einer beliebigen Mitarbeiterliste vorkommen.
    Match über email und/oder fullname-key 'lastname firstname'.
    """
    out = df_users.copy()

    cols = _detect_employee_list_columns(df_employee_list)
    col_email_list = cols["email"]
    col_name_list = cols["combined_name"]
    col_firstname_list = cols["firstname"]
    col_lastname_list = cols["lastname"]

    if not col_email_list and not col_name_list and not (col_firstname_list and col_lastname_list):
        raise ValueError(
            "Mitarbeiterliste benötigt mindestens eine Spalte 'email' oder "
            "eine Namensspalte ('Nachname / Vorname', 'Vorname / Nachname', ...) "
            "oder getrennte Spalten 'Vorname' + 'Nachname'."
        )

    email_set = set()
    name_set = set()
    employee_entry_keys: set[str] = set()

    if col_email_list:
        email_set = {
            norm_text(v)
            for v in df_employee_list[col_email_list].fillna("").astype(str)
            if str(v).strip()
        }
        employee_entry_keys.update({f"email:{v}" for v in email_set if v})

    if col_name_list:
        for v in df_employee_list[col_name_list].fillna("").astype(str):
            if str(v).strip():
                variants = _variants_from_combined_name(v)
                name_set.update(variants)
                employee_entry_keys.update({f"name:{x}" for x in variants if x})

    if col_firstname_list and col_lastname_list:
        for _, row in df_employee_list.iterrows():
            first = row.get(col_firstname_list, "")
            last = row.get(col_lastname_list, "")
            variants = _variants_from_first_last(first, last)
            name_set.update(variants)
            employee_entry_keys.update({f"name:{x}" for x in variants if x})

    flags = []
    reasons = []
    matched_employee_entry_keys: set[str] = set()

    for _, row in out.iterrows():
        row_reasons = []

        em = norm_text(row.get(COL_EMAIL, ""))
        if em and em in email_set:
            row_reasons.append("match_email")
            matched_employee_entry_keys.add(f"email:{em}")

        user_name_variants = _variants_from_first_last(row.get(COL_FIRSTNAME, ""), row.get(COL_LASTNAME, ""))
        matched_name_variants = [v for v in user_name_variants if v in name_set]
        if matched_name_variants:
            row_reasons.append("match_fullname")
            matched_employee_entry_keys.update({f"name:{v}" for v in matched_name_variants})

        flags.append(len(row_reasons) > 0)
        reasons.append(" | ".join(row_reasons))

    out[flag_name] = flags
    out[f"{flag_name}_reason"] = reasons

    if not return_stats:
        return out

    stats = {
        "employee_entries_total": len(employee_entry_keys),
        "employee_entries_matched": len(matched_employee_entry_keys),
        "employee_entries_unmatched": max(0, len(employee_entry_keys) - len(matched_employee_entry_keys)),
    }
    return out, stats
