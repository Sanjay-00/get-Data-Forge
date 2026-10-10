import re
from difflib import SequenceMatcher

import numpy as np
import pandas as pd


def _is_truncated(short, long):
    s, l = short.lower(), long.lower()
    return len(s) >= 3 and l.startswith(s) and len(s) / len(l) > 0.5


_SEPARATORS = re.compile(r"[\s_\-./\\]+")


def _loose_key(name):
    return _SEPARATORS.sub("", str(name).lower())


def _rename_map(columns, reference_cols):
    """{file column: reference column} for every header the automatic passes
    can match. Works on the column list alone, so callers that only need
    names (summaries, unmatched lists) never copy the DataFrame."""
    ref_set = set(reference_cols)
    rename_map = {}

    # Pass 1: case-insensitive name match
    ref_lower = {col.lower(): col for col in reference_cols}
    for col in columns:
        if col not in ref_set and col.lower() in ref_lower:
            rename_map[col] = ref_lower[col.lower()]

    # Pass 1b: same name once spacing and separators are ignored
    # ("Customer_Name", "customer-name", "Customer  Name" -> "Customer Name").
    # Only applied when the stripped name is unique on both sides, so two
    # genuinely different columns can never be folded onto one.
    taken = (set(columns) & ref_set) | set(rename_map.values())
    ref_keys = {}
    for col in reference_cols:
        ref_keys.setdefault(_loose_key(col), []).append(col)
    file_keys = {}
    for col in columns:
        file_keys.setdefault(_loose_key(col), []).append(col)
    for col in columns:
        if col in ref_set or col in rename_map:
            continue
        key = _loose_key(col)
        refs = ref_keys.get(key, [])
        if key and len(refs) == 1 and len(file_keys[key]) == 1 and refs[0] not in taken:
            rename_map[col] = refs[0]
            taken.add(refs[0])

    # Pass 2: positional match for truncated column names
    exact_matches = set(columns) & ref_set
    matched_ref = exact_matches | set(rename_map.values())
    matched_file = exact_matches | set(rename_map.keys())
    if len(columns) >= len(reference_cols):
        for file_col, ref_col in zip(list(columns)[:len(reference_cols)], reference_cols):
            if file_col not in matched_file and ref_col not in matched_ref:
                shorter, longer = sorted([file_col, ref_col], key=len)
                if _is_truncated(shorter, longer):
                    rename_map[file_col] = ref_col
                    matched_ref.add(ref_col)
                    matched_file.add(file_col)
    return rename_map


def normalized_columns(columns, reference_cols):
    rename_map = _rename_map(columns, reference_cols)
    return [rename_map.get(c, c) for c in columns]


def _normalize_columns(df, reference_cols):
    rename_map = _rename_map(list(df.columns), reference_cols)
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def build_file_summary(filename, df, reference_cols):
    columns = normalized_columns(list(df.columns), reference_cols)
    ref_set = set(reference_cols)
    file_set = set(columns)
    return {
        "File": filename,
        "Rows": len(df),
        "Total Columns": len(columns),
        "Matched Columns": len(file_set & ref_set),
        "Missing Columns": len(ref_set - file_set),
        "Extra Columns (dropped)": len(file_set - ref_set),
    }


def similarity(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def get_unmatched_pairs(df, reference_cols):
    columns = normalized_columns(list(df.columns), reference_cols)
    ref_set = set(reference_cols)
    exact_matches = set(columns) & ref_set
    unmatched_file = [c for c in columns if c not in exact_matches and not c.startswith("Unnamed")]
    unmatched_ref = [c for c in reference_cols if c not in exact_matches]
    pairs = []
    for fc in unmatched_file:
        sorted_refs = sorted(unmatched_ref, key=lambda rc: similarity(fc, rc), reverse=True)
        pairs.append((fc, sorted_refs))
    return pairs


_DTYPE_CATEGORIES = {
    "i": "numeric", "u": "numeric", "f": "numeric",
    "O": "text", "b": "boolean", "M": "datetime", "m": "duration",
}


def detect_dtype_mismatches(frames, columns):
    """pd.concat silently combines a numeric column from one file with a text
    column of the same name from another into one inconsistent object column
    (e.g. real ints mixed with strings) -- no coercion, no warning. This scans
    the already schema-aligned frames and flags any column where more than
    one broad dtype category (numeric / text / boolean / datetime / duration)
    appears across files, so the merge tool can surface it instead of staying
    silent. int64-vs-float64 within "numeric" is NOT flagged -- that's a
    routine, harmless pandas quirk (a blank cell upcasts a column to float),
    not a real type conflict.
    """
    mismatches = {}
    for col in columns:
        categories = set()
        for frame in frames:
            if col not in frame.columns:
                continue
            column = frame[col]
            # notna().any() instead of dropna(): same answer, no copy of the
            # column (dropping nulls never changes a column's dtype).
            if not column.notna().any():
                continue
            categories.add(_DTYPE_CATEGORIES.get(column.dtype.kind, column.dtype.kind))
        if len(categories) > 1:
            mismatches[col] = sorted(categories)
    return mismatches


def align_to_schema(df, reference_cols, manual_map=None):
    """Normalize df's columns onto reference_cols and reindex.

    Returns (aligned_df, dropped) where dropped is a list of
    (original_file_column, reference_column_it_collided_on) tuples for any
    column that had to be dropped because two source columns resolved to the
    same reference column (first occurrence wins, same as before - this is
    now just reported instead of silent).
    """
    original_cols = list(df.columns)
    names = normalized_columns(original_cols, reference_cols)
    if manual_map:
        names = [manual_map.get(n, n) for n in names]

    mask = pd.Index(names).duplicated()
    dropped = [(original_cols[i], names[i]) for i in range(len(names)) if mask[i]]
    position = {n: i for i, n in enumerate(names) if not mask[i]}

    # Built from views of df's columns with copy=False: no data is copied
    # at all (the old rename -> rename -> loc -> reindex chain copied the
    # whole frame four times). The result shares memory with df, which is
    # the read-only cached upload, so it must only be read or have new
    # columns added, never be modified in place; Merge and Concat only
    # read it (and Concat adds its Source column as a new column).
    missing = pd.Series(np.nan, index=df.index)
    aligned = pd.DataFrame(
        {rc: df.iloc[:, position[rc]] if rc in position else missing for rc in reference_cols},
        index=df.index,
        copy=False,
    )
    return aligned, dropped
