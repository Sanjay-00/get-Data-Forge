import numpy as np
import pandas as pd

from core.schema import get_unmatched_pairs, normalized_columns


def diff_columns(df_a, df_b):
    """Column-level diff. Aligns df_b's columns onto df_a's naming via the
    same _normalize_columns matcher used by Merge/Concat, so a truncated or
    renamed header is treated as the same column rather than a false
    added+removed pair.

    Returns (result, normalized_b_columns) where result has keys:
      added, removed, common: sorted lists of column names
      unmatched_pairs: output of get_unmatched_pairs on the normalized B
        columns - columns still ambiguous after normalization, for manual
        resolution.
    normalized_b_columns is df_b's column list after normalization, in
    order; apply it with with_columns() only when rows are actually needed,
    so the column-level view never copies df_b.
    """
    cols_a = list(df_a.columns)
    cols_b = normalized_columns(list(df_b.columns), cols_a)

    set_a, set_b = set(cols_a), set(cols_b)
    result = {
        "added": sorted(set_b - set_a),
        "removed": sorted(set_a - set_b),
        "common": sorted(set_a & set_b),
        # get_unmatched_pairs only reads the column labels
        "unmatched_pairs": get_unmatched_pairs(pd.DataFrame(columns=cols_b), cols_a),
    }
    return result, cols_b


def with_columns(df, columns):
    """df relabelled to `columns` (same length, same order). If two labels
    collide (two B columns resolved onto one A column), the first wins,
    the same rule Merge/Concat alignment uses."""
    out = df.set_axis(columns, axis=1)
    mask = out.columns.duplicated()
    return out.loc[:, ~mask] if mask.any() else out


def _values_equal(x, y):
    """A plain x == y flags "100" (text) vs 100 (number) as a real change --
    a false positive that's very plausible in practice, since the same
    column can easily get exported as text in one file version and as a
    real number in another. Tolerate type-only differences: try numeric
    equality first, then a whitespace-stripped string comparison, before
    falling back to strict equality."""
    if pd.isna(x) and pd.isna(y):
        return True
    if x == y:
        return True
    try:
        return float(x) == float(y)
    except (TypeError, ValueError):
        pass
    return str(x).strip() == str(y).strip()


def _equal_mask(a, b):
    """Vectorized _values_equal over two aligned columns. Strict equality
    and both-blank are settled for the whole column at once; only the
    (usually few) remaining cells go through the lenient per-value check,
    so the result is identical to calling _values_equal on every pair."""
    try:
        equal = (a == b).fillna(False).to_numpy(dtype=bool)
    except (TypeError, ValueError):
        equal = np.zeros(len(a), dtype=bool)
    equal |= (a.isna() & b.isna()).to_numpy()
    pending = (~equal).nonzero()[0]
    if len(pending):
        a_vals, b_vals = a.to_numpy(dtype=object), b.to_numpy(dtype=object)
        for i in pending:
            equal[i] = _values_equal(a_vals[i], b_vals[i])
    return equal


def diff_rows(df_a, df_b_normalized, key_columns, common_columns):
    """Row-level diff, matched by key_columns identity (not full-row hash or
    position). common_columns is the set of non-added/removed columns to
    compare for changes (already schema-aligned by diff_columns).

    Returns a dict:
      added: DataFrame of rows only in df_b (keyed columns + compare_cols)
      removed: DataFrame of rows only in df_a
      changed: DataFrame of rows present in both with >=1 differing column,
               with '_A'/'_B' suffixed compare columns and a 'Changed Columns' col
      unchanged_count: int
      duplicate_keys: {"A": [...], "B": [...]} - distinct key tuples that are
        not 1:1 within a file; diff results for these keys may be unreliable.
    """
    compare_cols = [c for c in common_columns if c not in key_columns]

    # Column selection already returns new frames; no extra .copy() needed.
    a = df_a[key_columns + compare_cols]
    b = df_b_normalized[key_columns + compare_cols]

    dup_a = a[a.duplicated(subset=key_columns, keep=False)][key_columns].drop_duplicates()
    dup_b = b[b.duplicated(subset=key_columns, keep=False)][key_columns].drop_duplicates()

    merged = a.merge(b, on=key_columns, how="outer", suffixes=("_A", "_B"), indicator=True)

    removed = merged[merged["_merge"] == "left_only"][key_columns + [f"{c}_A" for c in compare_cols]]
    added = merged[merged["_merge"] == "right_only"][key_columns + [f"{c}_B" for c in compare_cols]]
    both = merged[merged["_merge"] == "both"]

    changed_flags = {
        c: ~_equal_mask(both[f"{c}_A"], both[f"{c}_B"]) for c in compare_cols
    }
    if changed_flags:
        flags = np.column_stack([changed_flags[c] for c in compare_cols])
        row_changed = flags.any(axis=1)
    else:
        flags = np.zeros((len(both), 0), dtype=bool)
        row_changed = np.zeros(len(both), dtype=bool)

    changed = both.loc[row_changed].drop(columns=["_merge"])
    changed["Changed Columns"] = [
        ", ".join(c for c, f in zip(compare_cols, row) if f) for row in flags[row_changed]
    ]
    unchanged_count = int((~row_changed).sum())

    return {
        "added": added.reset_index(drop=True),
        "removed": removed.reset_index(drop=True),
        "changed": changed.reset_index(drop=True),
        "unchanged_count": unchanged_count,
        "duplicate_keys": {
            "A": dup_a.values.tolist(),
            "B": dup_b.values.tolist(),
        },
    }
