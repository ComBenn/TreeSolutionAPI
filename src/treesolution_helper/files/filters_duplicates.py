import pandas as pd

from config import COL_EMAIL, COL_FIRSTNAME, COL_ID, COL_LASTNAME, COL_USERNAME
from io_utils import norm_text


def _normalize_name_part(value) -> str:
    text = norm_text(value)
    if not text:
        return ""
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
    return " ".join(tokens).strip()


class _UnionFind:
    def __init__(self, size: int) -> None:
        self.parent = list(range(size))

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        root_a = self.find(a)
        root_b = self.find(b)
        if root_a != root_b:
            self.parent[root_b] = root_a


def _is_true_like(value) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    text = str(value).strip().casefold()
    return text in ("1", "true", "yes", "ja", "x")


def mark_duplicate_accounts(df_users: pd.DataFrame) -> pd.DataFrame:
    out = df_users.copy()
    row_count = len(out)
    if row_count == 0:
        out["flag_duplicate"] = []
        out["flag_duplicate_group"] = []
        out["flag_duplicate_reason"] = []
        return out

    uf = _UnionFind(row_count)
    key_to_indices: dict[str, list[int]] = {}
    row_match_reasons: list[set[str]] = [set() for _ in range(row_count)]
    technical_flags = [False] * row_count
    if "flag_technical_account" in out.columns:
        technical_flags = [_is_true_like(value) for value in out["flag_technical_account"].tolist()]

    for pos, (_, row) in enumerate(out.iterrows()):
        if technical_flags[pos]:
            continue
        email = norm_text(row.get(COL_EMAIL, ""))
        username = norm_text(row.get(COL_USERNAME, ""))
        lastname = _normalize_name_part(row.get(COL_LASTNAME, ""))
        firstname = _normalize_name_part(row.get(COL_FIRSTNAME, ""))
        full_name = f"{lastname} {firstname}".strip() if lastname and firstname else ""

        keys: list[tuple[str, str]] = []
        if email:
            keys.append(("email", email))
        if username:
            keys.append(("username", username))
        if full_name:
            keys.append(("name", full_name))

        for key_type, key_value in keys:
            compound_key = f"{key_type}:{key_value}"
            existing = key_to_indices.get(compound_key, [])
            if existing:
                reason = f"duplicate_{key_type}"
                row_match_reasons[pos].add(reason)
                for other_pos in existing:
                    uf.union(pos, other_pos)
                    row_match_reasons[other_pos].add(reason)
            existing.append(pos)
            key_to_indices[compound_key] = existing

    groups: dict[int, list[int]] = {}
    for pos in range(row_count):
        if technical_flags[pos]:
            continue
        root = uf.find(pos)
        groups.setdefault(root, []).append(pos)

    flags: list[bool] = [False] * row_count
    group_labels: list[str] = [""] * row_count
    reasons: list[str] = [""] * row_count

    duplicate_groups = [indices for indices in groups.values() if len(indices) > 1]
    duplicate_groups.sort(key=lambda indices: min(indices))

    for group_no, indices in enumerate(duplicate_groups, start=1):
        label = f"dup-{group_no:04d}"
        for pos in indices:
            flags[pos] = True
            group_labels[pos] = label
            reasons[pos] = " | ".join(sorted(row_match_reasons[pos]))

    out["flag_duplicate"] = flags
    out["flag_duplicate_group"] = group_labels
    out["flag_duplicate_reason"] = reasons

    if COL_ID in out.columns:
        id_series = out[COL_ID].fillna("").astype(str).str.strip()
        out["flag_duplicate_keep_candidate"] = (~out["flag_duplicate"]) | (id_series != "")
    else:
        out["flag_duplicate_keep_candidate"] = ~out["flag_duplicate"]

    return out
