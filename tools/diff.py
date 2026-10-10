import io
import zipfile

import streamlit as st

from core.diff import diff_columns, diff_rows, with_columns
from core.excel_writer import to_excel_bytes
from core.session import cached_read, cached_sheet_names, file_key, ns_key, release_unused_uploads, remember_cache_entry
from ui.shared import render_download_section, render_success_banner, render_tool_hero, render_unmatched_ui, section_label

TOOL = "diff"

# Streamlit's dataframe grid can crash the browser tab when handed a very
# large/wide table in one shot (seen in practice on a 9k-row changed set
# from a duplicate-heavy key column), so cap what's rendered, same pattern
# Merge already uses for its preview. The download is unaffected, since
# it's built from the full, uncapped result.
PREVIEW_ROW_CAP = 500


def _load_one(label, key_ns):
    st.markdown(
        f'<p style="font-size:0.8rem;font-weight:600;color:var(--eq-text);margin-bottom:0.4rem;">{label}</p>',
        unsafe_allow_html=True,
    )
    f = st.file_uploader(
        label, type=["xlsx", "xls", "xlsb"], accept_multiple_files=False,
        label_visibility="collapsed", key=ns_key(TOOL, f"{key_ns}_uploader"),
    )
    if not f:
        return None, None
    with st.spinner(f"Reading {f.name}..."):
        sheets = cached_sheet_names(f)
    sheet_choice = 0
    if len(sheets) > 1:
        sheet_choice = st.selectbox(f"{label} sheet", options=sheets, index=0, key=ns_key(TOOL, f"{key_ns}_sheet"))
    with st.spinner(f"Reading {f.name} data..."):
        df, err = cached_read(f, sheet_choice)
    if err:
        st.error(f"Cannot read {label}: {err}")
        return None, None
    return (file_key(f), sheet_choice), df


@st.cache_resource(scope="session", max_entries=2, ttl="1h", show_spinner=False)
def _cached_row_diff(id_a, id_b, b_columns, key_columns, common, _df_a, _df_b):
    """Row diff for one combination of inputs. The ids stand in for the two
    (read-only, cached) frames in the cache key, so changing an unrelated
    widget, or typing a download filename, never re-runs the comparison."""
    return diff_rows(_df_a, with_columns(_df_b, list(b_columns)), list(key_columns), list(common))


def _build_report_zip(row_result):
    buf = io.BytesIO()
    # ZIP_STORED: the entries are .xlsx files, already deflated.
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
        zf.writestr("added_rows.xlsx", to_excel_bytes(row_result["added"]))
        zf.writestr("removed_rows.xlsx", to_excel_bytes(row_result["removed"]))
        zf.writestr("changed_rows.xlsx", to_excel_bytes(row_result["changed"]))
    return buf.getvalue()


def render_diff_page():
    render_tool_hero("🔍", "Compare / Diff", "Schema-aware diff between two versions of a workbook", "#06B6D4")

    col_a, col_b = st.columns(2, gap="medium")
    with col_a:
        id_a, df_a = _load_one("File A (base)", "a")
    with col_b:
        id_b, df_b = _load_one("File B (comparison)", "b")
    release_unused_uploads(
        st.session_state.get(ns_key(TOOL, "a_uploader")), st.session_state.get(ns_key(TOOL, "b_uploader")),
    )

    if df_a is None or df_b is None:
        st.stop()

    # ── Column-level diff ────────────────────────────────────────────────────
    result, b_columns = diff_columns(df_a, df_b)

    # Filled in after the resolver below runs, so the counts include any
    # manual mappings while still rendering above the resolver.
    column_diff_area = st.container()

    # Anything still ambiguous after normalization - resolve or leave as add/remove
    resolvable = [(fc, rc) for fc, rc in result["unmatched_pairs"] if rc]
    manual_map = {}
    if resolvable:
        manual_maps = render_unmatched_ui({"File B": resolvable}, key_prefix=ns_key(TOOL, "map"))
        manual_map = manual_maps.get("File B", {})
        if manual_map:
            b_columns = [manual_map.get(c, c) for c in b_columns]
            cols_a_set, cols_b_set = set(df_a.columns), set(b_columns)
            result["common"] = sorted(cols_a_set & cols_b_set)
            result["added"] = sorted(cols_b_set - cols_a_set)
            result["removed"] = sorted(cols_a_set - cols_b_set)

    with column_diff_area:
        section_label("Column Diff", f"{len(result['common'])} common", "#0EA5E9")
        c1, c2, c3 = st.columns(3)
        c1.metric("Added Columns", len(result["added"]))
        c2.metric("Removed Columns", len(result["removed"]))
        c3.metric("Common Columns", len(result["common"]))

        if result["added"]:
            st.markdown(f"**Added in B:** {', '.join(result['added'])}")
        if result["removed"]:
            st.markdown(f"**Removed from A:** {', '.join(result['removed'])}")

    # ── Row-level diff (needs a user-designated key column) ──────────────────
    section_label("Row Diff", "requires a key column", "#D97706")
    key_columns = st.multiselect(
        "Key column(s) that uniquely identify a row (e.g. Employee ID, Order Number)",
        options=result["common"], key=ns_key(TOOL, "key_cols"),
    )

    if not key_columns:
        st.info("Select one or more key columns above to enable row-level diff.")
        st.stop()

    diff_args = (id_a, id_b, tuple(b_columns), tuple(key_columns), tuple(result["common"]))
    for upload_key in (id_a[0], id_b[0]):
        remember_cache_entry(upload_key, _cached_row_diff, *diff_args, None, None)
    with st.spinner("Comparing rows by key column(s)..."):
        row_result = _cached_row_diff(*diff_args, df_a, df_b)

    dup_a = row_result["duplicate_keys"]["A"]
    dup_b = row_result["duplicate_keys"]["B"]
    if dup_a or dup_b:
        st.warning(
            f"Key column(s) {', '.join(key_columns)} are not unique - "
            f"{len(dup_a)} duplicate key value(s) in File A, {len(dup_b)} in File B. "
            "Diff results for these rows may be unreliable."
        )

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Added Rows", len(row_result["added"]))
    r2.metric("Removed Rows", len(row_result["removed"]))
    r3.metric("Changed Rows", len(row_result["changed"]))
    r4.metric("Unchanged Rows", row_result["unchanged_count"])

    def _render_capped(label, df, expanded=False):
        with st.expander(label, expanded=expanded):
            if len(df) > PREVIEW_ROW_CAP:
                st.caption(f"Showing first {PREVIEW_ROW_CAP:,} of {len(df):,} rows. Download the full report below to see everything.")
            st.dataframe(df.head(PREVIEW_ROW_CAP), width="stretch", hide_index=True)

    _render_capped("Added rows (only in File B)", row_result["added"])
    _render_capped("Removed rows (only in File A)", row_result["removed"])
    _render_capped("Changed rows", row_result["changed"], expanded=True)

    # Built only when Download is clicked: writing three workbooks on every
    # rerun was most of this page's cost.
    render_download_section(TOOL, "diff_report.zip", lambda: _build_report_zip(row_result), "application/zip")
    render_success_banner("Diff complete")
