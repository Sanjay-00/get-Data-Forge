import re

import pandas as pd

ZERO_WIDTH_SPACE = "​"
ZERO_WIDTH_NON_JOINER = "‌"
ZERO_WIDTH_JOINER = "‍"
BOM = "﻿"
NBSP = "\xa0"

ZERO_WIDTH_CHARS = [ZERO_WIDTH_SPACE, ZERO_WIDTH_NON_JOINER, ZERO_WIDTH_JOINER, BOM]

DEFAULT_PUNCTUATION = ".,;:-–—'\"()[]"  # includes hyphen, en dash, em dash as literal strippable characters

_INVISIBLE_PREVIEW_MARKERS = {
    NBSP: "·",
    ZERO_WIDTH_SPACE: "[ZWSP]",
    ZERO_WIDTH_NON_JOINER: "[ZWNJ]",
    ZERO_WIDTH_JOINER: "[ZWJ]",
    BOM: "[BOM]",
}


def _map_strings(series, func):
    """func applied to every str cell, everything else left as is. Each
    distinct string is cleaned once and mapped back, so a 200k row column
    with a few hundred distinct values does a few hundred calls instead of
    200k. Mapping only str cells also keeps hash-equal non-strings apart
    (1, 1.0 and True share a dict key but aren't touched here)."""
    values = series.tolist()
    is_str = [isinstance(v, str) for v in values]
    if not any(is_str):
        return series
    cache = {}
    out = []
    for v, flag in zip(values, is_str):
        if flag:
            cleaned = cache.get(v)
            if cleaned is None:
                cleaned = cache[v] = func(v)
            out.append(cleaned)
        else:
            out.append(v)
    return pd.Series(out, index=series.index, dtype=series.dtype, name=series.name)


def _collapse_whitespace_value(v):
    return re.sub(r"\s+", " ", v.strip())


def _strip_invisible_value(v):
    for ch in ZERO_WIDTH_CHARS:
        v = v.replace(ch, "")
    return v.replace(NBSP, " ")


def _case_value(v, mode):
    if mode == "title":
        return v.title()
    if mode == "upper":
        return v.upper()
    if mode == "lower":
        return v.lower()
    return v


def collapse_whitespace(series):
    return _map_strings(series, _collapse_whitespace_value)


def strip_invisible_chars(series):
    return _map_strings(series, _strip_invisible_value)


def strip_edge_punctuation(series, chars=DEFAULT_PUNCTUATION):
    return _map_strings(series, lambda v: v.strip(chars))


def normalize_case(series, mode):
    return _map_strings(series, lambda v: _case_value(v, mode))


def _series_equal(a, b):
    both_na = a.isna() & b.isna()
    return (a == b) | both_na


def _is_string_like(series):
    return series.dtype == object or pd.api.types.is_string_dtype(series)


def apply_trim_operations(df, columns, operations, punctuation_chars=None, case_mode=None):
    """Apply the selected composable cleaning operations to `columns` of df.

    operations: subset of {"invisible", "whitespace", "punctuation", "case"},
    always applied in this fixed order regardless of selection order (invisible
    chars normalized to spaces before whitespace-collapse sweeps them up).

    Returns (new_df, change_report) where change_report is
    {column: {"changed": int, "newly_empty": int, "skipped_non_string": bool}}.
    """
    steps = []
    if "invisible" in operations:
        steps.append(_strip_invisible_value)
    if "whitespace" in operations:
        steps.append(_collapse_whitespace_value)
    if "punctuation" in operations and punctuation_chars:
        steps.append(lambda v: v.strip(punctuation_chars))
    if "case" in operations and case_mode:
        steps.append(lambda v: _case_value(v, case_mode))

    def clean(v):
        for step in steps:
            v = step(v)
        return v

    # Shallow copy: only the cleaned columns are replaced (pandas 2 setitem
    # always swaps in a new array), the rest stay shared with df, which may
    # be the read-only cached upload.
    new_df = df.copy(deep=False)
    change_report = {}

    for col in columns:
        original = df[col]
        if not _is_string_like(original):
            change_report[col] = {"changed": 0, "newly_empty": 0, "skipped_non_string": True}
            continue

        working = _map_strings(original, clean) if steps else original

        new_df[col] = working
        changed_mask = ~_series_equal(original, working)
        was_nonempty_str = pd.Series(
            [isinstance(v, str) and v != "" for v in original.tolist()], index=original.index,
        )
        newly_empty_mask = was_nonempty_str & (working == "")

        change_report[col] = {
            "changed": int(changed_mask.sum()),
            "newly_empty": int(newly_empty_mask.sum()),
            "skipped_non_string": False,
        }

    return new_df, change_report


def _visualize_invisible(v):
    if not isinstance(v, str):
        return v
    for ch, marker in _INVISIBLE_PREVIEW_MARKERS.items():
        v = v.replace(ch, marker)
    return v


def build_trim_preview(original_df, new_df, columns, cap=50):
    """Sample of rows where any selected column changed, rendered with
    invisible characters made visible. Returns (preview_df, more_count)."""
    change_mask = pd.Series(False, index=original_df.index)
    for col in columns:
        change_mask = change_mask | ~_series_equal(original_df[col], new_df[col])

    changed_idx = original_df.index[change_mask]
    sample_idx = changed_idx[:cap]

    rows = []
    for idx in sample_idx:
        record = {}
        for col in columns:
            record[f"{col} (before)"] = _visualize_invisible(original_df.at[idx, col])
            record[f"{col} (after)"] = _visualize_invisible(new_df.at[idx, col])
        rows.append(record)

    preview_df = pd.DataFrame(rows)
    more = len(changed_idx) - len(sample_idx)
    return preview_df, more
