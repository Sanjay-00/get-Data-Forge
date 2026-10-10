import pandas as pd
import streamlit as st

from core.excel_writer import to_excel_bytes
from core.memory import release_memory
from core.session import cached_read, cached_sheet_names, file_key, ns_key, release_unused_uploads
from core.trim import DEFAULT_PUNCTUATION, apply_trim_operations, build_trim_preview
from ui.shared import render_download_section, render_success_banner, render_tool_hero, section_label

TOOL = "trim"


def render_trim_page():
    render_tool_hero("🧼", "Trim Columns", "Keep only the columns you need, and clean up whitespace, invisible characters, and casing in the rest", "#F59E0B")

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

    # ── Column picker: keep only what you need, drop the rest ───────────────
    # Default is "everything selected" (a safe no-op for anyone who's only
    # here to clean values, not drop columns). The Select all/Clear all
    # buttons must run *before* the multiselect reads its session-state value
    # -- Streamlit forbids writing to a widget's key after it's instantiated,
    # so this order (buttons first, widget second) is required, not stylistic.
    section_label("Columns to Keep", f"{len(df.columns)} columns available", "#D97706")
    btn_all, btn_none, _ = st.columns([1, 1, 4])
    with btn_all:
        if st.button("Select all", key=ns_key(TOOL, "select_all_cols"), width="stretch"):
            st.session_state[ns_key(TOOL, "keep_columns")] = list(df.columns)
            st.rerun()
    with btn_none:
        if st.button("Clear all", key=ns_key(TOOL, "clear_all_cols"), width="stretch"):
            st.session_state[ns_key(TOOL, "keep_columns")] = []
            st.rerun()

    keep_columns = st.multiselect(
        "keep_col", options=list(df.columns), default=list(df.columns),
        label_visibility="collapsed", key=ns_key(TOOL, "keep_columns"),
    )

    if not keep_columns:
        st.info("Select at least one column to keep (or click \"Select all\" above).")
        st.stop()

    # preserve the file's original column order, not multiselect click order
    keep_columns = [c for c in df.columns if c in keep_columns]
    dropped_columns = [c for c in df.columns if c not in keep_columns]
    if dropped_columns:
        with st.expander(f"Dropping {len(dropped_columns)} column(s)", expanded=False):
            st.caption(", ".join(dropped_columns))

    df = df[keep_columns]

    section_label("Columns to Clean", f"{len(df.columns)} columns kept", "#D97706")
    columns = st.multiselect(
        "col", options=list(df.columns), label_visibility="collapsed",
        key=ns_key(TOOL, "columns"),
    )

    if not columns:
        st.info("Select one or more columns above to clean, or just download with the column selection above applied.")

    section_label("Cleaning Operations", None)
    op1, op2 = st.columns(2)
    with op1:
        do_invisible = st.checkbox("Strip invisible/non-printable characters (incl. non-breaking space)",
                                    value=True, key=ns_key(TOOL, "op_invisible"))
        do_whitespace = st.checkbox("Collapse whitespace (leading/trailing + internal runs)",
                                     value=True, key=ns_key(TOOL, "op_whitespace"))
    with op2:
        do_punctuation = st.checkbox("Strip edge punctuation", value=False, key=ns_key(TOOL, "op_punctuation"))
        punctuation_chars = st.text_input(
            "Punctuation characters (edge-only, never mid-string)",
            value=DEFAULT_PUNCTUATION, key=ns_key(TOOL, "punct_chars"),
            disabled=not do_punctuation,
        )
        do_case = st.checkbox("Normalize case", value=False, key=ns_key(TOOL, "op_case"))
        case_mode = st.selectbox(
            "Case mode", options=["title", "upper", "lower"], key=ns_key(TOOL, "case_mode"),
            disabled=not do_case,
        )

    operations = []
    if do_invisible:
        operations.append("invisible")
    if do_whitespace:
        operations.append("whitespace")
    if do_punctuation:
        operations.append("punctuation")
    if do_case:
        operations.append("case")

    if not operations and columns:
        st.info("Select at least one cleaning operation above, or clear the column selection to skip cleaning.")

    # No cleaning columns and/or no operations selected -> skip cleaning
    # entirely rather than forcing a stop; the column-keep selection above
    # is a complete, useful action on its own (e.g. just reducing 30 columns
    # down to 4, with no value cleanup needed).
    if columns and operations:
        with st.spinner("Applying cleaning operations..."):
            new_df, change_report = apply_trim_operations(
                df, columns, operations,
                punctuation_chars=punctuation_chars if do_punctuation else None,
                case_mode=case_mode if do_case else None,
            )
    else:
        new_df, change_report = df.copy(), {}

    # change_report only has entries for columns that actually went through
    # apply_trim_operations -- if cleaning was skipped (no columns and/or no
    # operations selected), fall back to an empty list instead of indexing
    # into an empty report with the (still-selected) column multiselect.
    cleaned_columns = columns if change_report else []

    # ── Mutation safety report ───────────────────────────────────────────────
    if cleaned_columns:
        section_label("Change Report", "review before downloading", "#0EA5E9")
        report_rows = []
        for col in cleaned_columns:
            r = change_report[col]
            report_rows.append({
                "Column": col,
                "Cells Changed": r["changed"],
                "Newly Empty": r["newly_empty"],
                "Skipped (non-string column)": "Yes" if r["skipped_non_string"] else "No",
            })
        st.dataframe(pd.DataFrame(report_rows), width="stretch", hide_index=True)

        if any(change_report[c]["skipped_non_string"] for c in cleaned_columns):
            skipped = [c for c in cleaned_columns if change_report[c]["skipped_non_string"]]
            st.warning(f"Skipped non-string column(s), no changes applied: {', '.join(skipped)}")

        if any(change_report[c]["newly_empty"] for c in cleaned_columns):
            newly_empty_cols = {c: change_report[c]["newly_empty"] for c in cleaned_columns if change_report[c]["newly_empty"]}
            details = ", ".join(f"{c} ({n} cell{'s' if n != 1 else ''})" for c, n in newly_empty_cols.items())
            st.warning(f"Some cells became empty after cleaning: {details}. Review the preview below before downloading.")

    # ── Mandatory before/after preview ──────────────────────────────────────
    with st.spinner("Building before/after preview..."):
        preview_df, more = build_trim_preview(df, new_df, cleaned_columns)
    if cleaned_columns:
        section_label("Before / After Preview", f"{len(preview_df)} changed row(s) shown" + (f", {more} more not shown" if more else ""), "#6366F1")
    if preview_df.empty:
        if cleaned_columns:
            st.info("No changes to preview with the current selection.")
    else:
        st.dataframe(preview_df, width="stretch", hide_index=True)

    # Any change to the selections invalidates a prepared download; without
    # this, editing columns or operations after "Apply" kept serving the
    # old file.
    current_sig = (
        file_key(uploaded_file), sheet_choice, tuple(keep_columns), tuple(cleaned_columns), tuple(operations),
        punctuation_chars if do_punctuation else None, case_mode if do_case else None,
    )
    if st.session_state.get(ns_key(TOOL, "last_sig")) != current_sig:
        st.session_state[ns_key(TOOL, "trim_done")] = False
        st.session_state[ns_key(TOOL, "last_sig")] = current_sig

    if st.button("Apply and Prepare Download", type="primary", width="stretch", key=ns_key(TOOL, "apply_btn")):
        with st.spinner("Writing cleaned file..."):
            excel_bytes = to_excel_bytes(new_df)
        st.session_state[ns_key(TOOL, "excel_bytes")] = excel_bytes
        st.session_state[ns_key(TOOL, "trim_done")] = True
        del excel_bytes
        release_memory()

    if st.session_state.get(ns_key(TOOL, "trim_done")):
        render_download_section(
            TOOL,
            "trimmed_output.xlsx",
            st.session_state[ns_key(TOOL, "excel_bytes")],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        render_success_banner("Trim complete")
