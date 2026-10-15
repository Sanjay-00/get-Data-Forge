import io
import re
import zipfile

import pandas as pd
import streamlit as st

from core.excel_writer import to_excel_bytes
from core.memory import release_memory
from core.schema import similarity
from core.session import cached_read, cached_sheet_names, file_key, ns_key, release_unused_uploads
from ui.shared import render_download_section, render_success_banner, render_tool_hero, section_label

TOOL = "split"
BLANK_LABEL = "(blank)"
FUZZY_THRESHOLD = 0.75
MAX_FUZZY_VALUES = 150
HIGH_CARDINALITY_WARNING_THRESHOLD = 100
# Rows of the per-value preview actually sent to the browser; an ID column
# would otherwise ship a 200k-row table on every rerun.
PREVIEW_ROW_CAP = 1000


# Names Windows refuses as a file name, with or without an extension.
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL"} | {f"{p}{i}" for p in ("COM", "LPT") for i in range(1, 10)}


def _safe_filename(value):
    """A file name that extracts cleanly on Windows, macOS and Linux."""
    name = re.sub(r'[\\/*?:"<>|\x00-\x1f]', "_", str(value).strip())
    # Windows silently drops trailing dots/spaces, which would merge
    # "A." and "A" into one file on extraction.
    name = name[:80].rstrip(". ") or "blank"
    if name.split(".")[0].upper() in _WINDOWS_RESERVED:
        name = f"_{name}"
    return name


def _unique_name(base, used):
    """base, or base_2, base_3... whichever isn't taken yet. Compared
    case-insensitively: Windows and macOS treat "ABC.xlsx" and "abc.xlsx"
    as the same file, so extracting the zip would overwrite one with the
    other. Checks final names, so a real value "X_2" can't collide with
    the numbered copy of "X" either."""
    name, n = base, 1
    while name.lower() in used:
        n += 1
        name = f"{base}_{n}"
    used.add(name.lower())
    return name


def _display_value(v):
    """Clean string form of a split-column value, used consistently for
    preview/grouping/filenames regardless of the column's dtype. A whole-number
    float (e.g. 5.0) displays as "5" -- pandas silently upcasts an int column
    to float64 the moment it has any blank cell, so a day-of-month column like
    [5, 10, None] would otherwise show as "5.0"/"10.0" everywhere. A "due date"
    column is just as plausibly a real date -- pandas reads it as a Timestamp,
    which would otherwise print the trivial "00:00:00" on every value even
    when no row actually has a time component."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, pd.Timestamp):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d_%H-%M-%S")
    return str(v)


def _display_column(series):
    """series.map(_display_value) with each distinct value formatted once.
    Typed columns map through a dict built from their unique values; object
    columns key the memo on (type, value) so hash-equal values of different
    types (1, 1.0, True) still get their own display string."""
    if series.dtype != object:
        uniques = series.dropna().drop_duplicates().tolist()
        return series.map({u: _display_value(u) for u in uniques}, na_action="ignore")
    memo = {}

    def display(v):
        key = (type(v), v)
        out = memo.get(key)
        if out is None:
            out = memo[key] = _display_value(v)
        return out

    return series.map(display, na_action="ignore")


def _detect_case_variants(series):
    """Values that differ only by case or surrounding whitespace (e.g. "CHE"
    vs "che") would otherwise be silently split into separate output files.
    Returns {normalized_key: [distinct raw values]} for keys with more than
    one distinct raw value.

    Works on distinct (normalized, raw) pairs and only groups the keys that
    actually repeat: grouping every distinct value took ~5s on a 200k-row ID
    column, which is often the column selected when the page first opens."""
    non_null = series.dropna().astype(str)
    if non_null.empty:
        return {}
    pairs = pd.DataFrame({"norm": non_null.str.strip().str.casefold(), "raw": non_null}).drop_duplicates()
    repeated = pairs["norm"].duplicated(keep=False)
    if not repeated.any():
        return {}
    return {
        norm_key: sorted(raw_group.tolist())
        for norm_key, raw_group in pairs[repeated].groupby("norm")["raw"]
    }


def _build_canonical_map(series, variants):
    """Maps every raw value in a detected variant group to the most frequent
    raw spelling (by row count) in that group -- that spelling becomes the
    group's output filename and grouping key. Row data itself is untouched;
    only which output file a row lands in changes."""
    non_null = series.dropna().astype(str)
    normalized = non_null.str.strip().str.casefold()
    # Only rows in a variant group are grouped, not every distinct value.
    in_variant = normalized.isin(variants.keys())
    canonical = {}
    for norm_key, raw_group in non_null[in_variant].groupby(normalized[in_variant]):
        top_value = raw_group.value_counts().idxmax()
        for v in raw_group.unique():
            canonical[v] = top_value
    return canonical


def _find_fuzzy_candidates(values, threshold=FUZZY_THRESHOLD):
    """Pairs of already-distinct values (post case/whitespace fold) that are
    textually similar enough to plausibly be typos of each other -- e.g.
    "SAKIN" vs "SAKINN". Ranked by similarity, never auto-applied: unlike
    case-fold (provably the same value), textual similarity between short
    codes is a weak signal -- "CHE" vs "CHI" could be a typo or two
    genuinely different valid codes, so a human has to confirm each pair."""
    values = sorted(values)
    candidates = []
    for i in range(len(values)):
        for j in range(i + 1, len(values)):
            a, b = values[i], values[j]
            ratio = similarity(a, b)
            if ratio >= threshold:
                candidates.append((a, b, ratio))
    candidates.sort(key=lambda c: -c[2])
    return candidates


def _build_fuzzy_canonical_map(value_counts, approved_pairs):
    """Union-finds only the pairs a human explicitly approved, then maps
    every value in a merged group to the most frequent spelling (by row
    count) in that group."""
    parent = {v: v for v in value_counts.index}

    def find(v):
        while parent[v] != v:
            v = parent[v]
        return v

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for a, b in approved_pairs:
        union(a, b)

    groups = {}
    for v in value_counts.index:
        groups.setdefault(find(v), []).append(v)

    canonical = {}
    for members in groups.values():
        if len(members) > 1:
            top_value = max(members, key=lambda v: value_counts[v])
            for v in members:
                canonical[v] = top_value
    return canonical


def render_split_page():
    render_tool_hero("✂️", "Split by Column Value", "Split one Excel file into many, one output per distinct column value", "#8B5CF6")

    uploaded_file = st.file_uploader(
        "upload", type=["xlsx", "xls", "xlsb"],
        accept_multiple_files=False, label_visibility="collapsed",
        key=ns_key(TOOL, "uploader"),
    )
    release_unused_uploads(uploaded_file)

    if not uploaded_file:
        st.stop()

    with st.spinner(f"Reading {uploaded_file.name}..."):
        sheets = cached_sheet_names(uploaded_file)
    sheet_choice = 0
    if len(sheets) > 1:
        sheet_choice = st.selectbox("Sheet", options=sheets, index=0, key=ns_key(TOOL, "sheet"))

    with st.spinner("Reading sheet data..."):
        df, err = cached_read(uploaded_file, sheet_choice)
    if err:
        st.error(f"Cannot read file: {err}")
        st.stop()

    section_label("Split Column", f"{len(df.columns)} columns available", "#D97706")
    split_col = st.selectbox(
        "col", options=list(df.columns), label_visibility="collapsed",
        key=ns_key(TOOL, "split_col"),
    )

    # ── Normalize to a clean string column first, regardless of dtype ─────
    # so numeric/date columns (e.g. a day-of-month "due date" column) go
    # through the same case-fold/fuzzy-typo pipeline as text columns instead
    # of crashing it (similarity() expects strings) or displaying "5.0"
    # instead of "5" (pandas upcasts int -> float64 the moment a column has
    # any blank cell).
    raw_col = _display_column(df[split_col])

    # ── Stage 1: case/whitespace variant detection (safe, auto-eligible) ──
    variants = _detect_case_variants(raw_col)

    merge_variants = False
    if variants:
        example_key, example_values = next(iter(variants.items()))
        example_text = " / ".join(f'"{v}"' for v in example_values)
        st.warning(
            f"Found {len(variants)} value(s) in **{split_col}** that differ only by case "
            f"or whitespace (e.g. {example_text}) - these will be split into separate "
            f"files unless merged below."
        )
        merge_variants = st.checkbox(
            "Merge values that differ only by case or whitespace into one output file each",
            value=False, key=ns_key(TOOL, "merge_variants"),
        )

    if merge_variants:
        canonical_map = _build_canonical_map(raw_col, variants)
        case_folded_col = raw_col.copy()
        mask = raw_col.notna()
        case_folded_col.loc[mask] = raw_col.loc[mask].astype(str).map(lambda v: canonical_map.get(v, v))
    else:
        case_folded_col = raw_col

    # ── Stage 2: fuzzy-typo suggestions (human-reviewed, never auto-applied) ──
    case_value_counts = case_folded_col.value_counts(dropna=True)
    distinct_values = list(case_value_counts.index)

    approved_pairs = []
    if len(distinct_values) > MAX_FUZZY_VALUES:
        st.info(
            f"Skipping typo-similarity check - {len(distinct_values)} distinct values "
            "is too many to compare pairwise."
        )
    elif len(distinct_values) > 1:
        candidates = _find_fuzzy_candidates(distinct_values)
        if candidates:
            section_label("Possible Typos", f"{len(candidates)} pair(s) found", "#D97706")
            st.markdown(
                '<span style="color:var(--eq-text-secondary);font-size:0.75rem;">'
                "These values look similar but are not identical after case/whitespace "
                "normalization. Tick any pair that is actually the same value to merge "
                "them - nothing is merged automatically.</span>",
                unsafe_allow_html=True,
            )
            for a, b, ratio in candidates:
                checked = st.checkbox(
                    f'"{a}"  is the same as  "{b}"   ({ratio:.0%} similar)',
                    value=False, key=ns_key(TOOL, f"fuzzy_{(split_col, a, b)!r}"),
                )
                if checked:
                    approved_pairs.append((a, b))

    fuzzy_map = _build_fuzzy_canonical_map(case_value_counts, approved_pairs) if approved_pairs else {}
    effective_col = case_folded_col.map(lambda v: fuzzy_map.get(v, v), na_action="ignore") if fuzzy_map else case_folded_col

    current_sig = (file_key(uploaded_file), sheet_choice, split_col, merge_variants, tuple(sorted(approved_pairs)))
    if st.session_state.get(ns_key(TOOL, "last_sig")) != current_sig:
        st.session_state[ns_key(TOOL, "split_done")] = False
        st.session_state[ns_key(TOOL, "last_sig")] = current_sig

    # ── Per-value preview summary ────────────────────────────────────────
    value_counts = effective_col.value_counts(dropna=True)
    blank_count = int(effective_col.isna().sum())

    preview_rows = [{"Value": str(v), "Rows": int(c)} for v, c in value_counts.items()]
    if blank_count:
        preview_rows.append({"Value": BLANK_LABEL, "Rows": blank_count})

    section_label("Split Preview", f"{len(preview_rows)} output files", "#0EA5E9")
    if len(preview_rows) > HIGH_CARDINALITY_WARNING_THRESHOLD:
        st.warning(
            f"**{split_col}** has {len(preview_rows)} distinct values, so this will "
            f"produce {len(preview_rows)} output files. If you meant to split on a "
            "column with fewer, more meaningful groups (not a near-unique column like "
            "an ID), double-check your column choice before splitting."
        )
    if len(preview_rows) > PREVIEW_ROW_CAP:
        st.caption(f"Showing the {PREVIEW_ROW_CAP:,} largest of {len(preview_rows):,} groups.")
    st.dataframe(pd.DataFrame(preview_rows[:PREVIEW_ROW_CAP]), width="stretch", hide_index=True)

    if st.button("Split File", type="primary", width="stretch", key=ns_key(TOOL, "split_btn")):
        blank_mask = effective_col.isna()
        # Groups are produced one at a time: list(df.groupby(...)) used to
        # materialize every group up front, a full extra copy of the sheet.
        # groupby's default dropna=True leaves the blank rows out (dropna=False
        # raises on some pandas versions when the key has nulls), and they
        # are appended as their own group at the end.
        grouped = df.groupby(effective_col, sort=True)
        total = grouped.ngroups + (1 if blank_mask.any() else 0)

        def iter_groups():
            yield from grouped
            if blank_mask.any():
                yield BLANK_LABEL, df[blank_mask]

        buf = io.BytesIO()
        split_progress = st.progress(0, text="Preparing output files...")
        used_names = set()
        last_pct = -1
        # ZIP_STORED: every entry is an .xlsx, which is already a deflated
        # zip; compressing it again costs CPU for ~0 bytes saved.
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
            for i, (value, group_df) in enumerate(iter_groups()):
                label = str(value)
                pct = int((i + 1) / total * 100)
                # One progress message per percent, not per file: with
                # thousands of groups, per-file updates flood the websocket.
                if pct != last_pct:
                    split_progress.progress(pct, text=f"Writing {label} ({i + 1} of {total})")
                    last_pct = pct
                # Different raw values can sanitize to the same file name
                # ("A/B" and "A_B"), so names are made unique, not overwritten.
                file_name = _unique_name(_safe_filename(label), used_names)
                # Sheet named after the group (safe_sheet_name trims it to Excel's rules).
                zf.writestr(f"{file_name}.xlsx", to_excel_bytes(group_df, sheet_name=label))
        split_progress.empty()

        st.session_state[ns_key(TOOL, "zip_bytes")] = buf.getvalue()
        del buf
        release_memory()
        st.session_state[ns_key(TOOL, "split_done")] = True
        st.session_state[ns_key(TOOL, "split_summary")] = preview_rows

    if st.session_state.get(ns_key(TOOL, "split_done")):
        summary = st.session_state.get(ns_key(TOOL, "split_summary"), [])
        k1, k2 = st.columns([1, 3])
        k1.metric("Output Files", len(summary))
        with k2:
            render_download_section(
                TOOL,
                "split_output.zip",
                st.session_state[ns_key(TOOL, "zip_bytes")],
                "application/zip",
            )
        render_success_banner("Split complete")
