import pandas as pd
import streamlit as st

from core.excel_writer import frames_to_excel_bytes
from core.memory import release_memory
from core.schema import align_to_schema, build_file_summary, get_unmatched_pairs
from core.session import cached_read_many, cached_sheet_names, file_key, ns_key, release_unused_uploads
from ui.shared import (
    render_download_section,
    render_dropped_columns_warning,
    render_success_banner,
    render_summary_bar,
    render_tool_hero,
    render_unmatched_ui,
    section_label,
    stop_on_duplicate_names,
)

TOOL = "concat"
PREVIEW_ROWS = 100


def _unit_key(workbook_name, sheet_name):
    return f"{workbook_name} :: {sheet_name}"


def render_concat_page():
    render_tool_hero("📚", "Sheet-wise Concat", "Stack sheets from one or more workbooks into one consolidated file", "#3B82F6")

    uploaded_files = st.file_uploader(
        "upload", type=["xlsx", "xls", "xlsb"],
        accept_multiple_files=True, label_visibility="collapsed",
        key=ns_key(TOOL, "uploader"),
    )
    release_unused_uploads(*(uploaded_files or []))

    if uploaded_files:
        render_summary_bar(uploaded_files)

    if not uploaded_files:
        st.stop()

    # Sheets are keyed "workbook :: sheet", so two workbooks with one name
    # would collide just like Merge's files do.
    stop_on_duplicate_names(uploaded_files)

    # ── Per-workbook sheet selection (default: all sheets participate) ─────
    section_label("Sheets to include", "per workbook", "#D97706")
    participating = []  # list of (workbook_name, sheet_name, uploaded_file)
    sheet_cols = st.columns(min(len(uploaded_files), 3))
    with st.spinner("Detecting sheets..."):
        workbook_sheets = {f.name: cached_sheet_names(f) for f in uploaded_files}
    for i, f in enumerate(uploaded_files):
        sheets = workbook_sheets[f.name]
        with sheet_cols[i % 3]:
            chosen = st.multiselect(
                f.name, options=sheets, default=sheets,
                key=ns_key(TOOL, f"sheets_{f.name}"),
            )
        for s in chosen:
            participating.append((f.name, s, f))

    if not participating:
        st.warning("Select at least one sheet to continue.")
        st.stop()

    # ── Read every participating sheet ──────────────────────────────────────
    sheet_data = {}
    read_progress = st.progress(0, text="Reading sheets...")

    def on_progress(done, total):
        read_progress.progress(int(done / total * 100), text=f"Reading sheets ({done} of {total})")

    # New sheets are parsed in parallel; already-parsed ones come from cache.
    results = cached_read_many([(f, s) for _, s, f in participating], progress_cb=on_progress)
    for (wb_name, sheet_name, _), result in zip(participating, results):
        sheet_data[_unit_key(wb_name, sheet_name)] = result
    read_progress.empty()

    # ── Reference sheet + schema preview ─────────────────────────────────────
    unit_keys = [_unit_key(wb, s) for wb, s, _ in participating]

    col_ref, col_schema = st.columns([1, 1], gap="medium")
    with col_ref:
        section_label("Reference Sheet", "sets column schema", "#D97706")
        ref_key = st.selectbox(
            "ref", options=unit_keys, index=0,
            help="All other sheets are aligned to this sheet's columns.",
            label_visibility="collapsed",
            key=ns_key(TOOL, "ref_key"),
        )

    ref_df, ref_err = sheet_data[ref_key]
    if ref_err:
        st.error(f"Cannot read reference sheet: {ref_err}")
        st.stop()

    reference_cols = list(ref_df.columns)

    if not reference_cols:
        st.error(
            f"**{ref_key}** has no columns to align to (the sheet may be empty). "
            "Pick a different reference sheet."
        )
        st.stop()

    with col_schema:
        section_label("Schema Preview", f"{len(reference_cols)} columns", "#6366F1")
        with st.expander(ref_key, expanded=False):
            st.write(reference_cols)

    include_provenance = st.checkbox(
        "Add a source column (workbook + sheet each row came from)",
        value=True, key=ns_key(TOOL, "provenance"),
    )

    # ── Sheet summary ─────────────────────────────────────────────────────────
    section_label("Sheet Summary", f"{len(unit_keys)} sheets", "#0EA5E9")
    summaries = []
    warnings = []
    for key in unit_keys:
        df, err = sheet_data[key]
        if err:
            summaries.append({
                "Sheet": key, "Rows": "-", "Total Columns": "-",
                "Matched Columns": "-", "Missing Columns": "-", "Extra Columns (dropped)": "-",
            })
            warnings.append(key)
        else:
            row = build_file_summary(key, df, reference_cols)
            row["Sheet"] = row.pop("File")
            summaries.append(row)

    for w in warnings:
        st.warning(f"Could not read **{w}**")

    st.dataframe(pd.DataFrame(summaries), width="stretch", hide_index=True)

    # ── Unmatched column resolution ─────────────────────────────────────────
    all_unmatched = {}
    for key in unit_keys:
        if key == ref_key:
            continue
        df, err = sheet_data[key]
        if err:
            continue
        pairs = get_unmatched_pairs(df, reference_cols)
        resolvable = [(fc, rc) for fc, rc in pairs if rc]
        if resolvable:
            all_unmatched[key] = resolvable

    manual_maps = render_unmatched_ui(all_unmatched, key_prefix=ns_key(TOOL, "map")) if all_unmatched else {}

    # Column mappings and file contents are part of the signature, so
    # changing either after a concat invalidates the old result.
    current_sig = (
        tuple(unit_keys), tuple(file_key(f) for f in uploaded_files), ref_key, include_provenance,
        repr(sorted((k, sorted(v.items())) for k, v in manual_maps.items())),
    )
    if st.session_state.get(ns_key(TOOL, "last_sig")) != current_sig:
        st.session_state[ns_key(TOOL, "concat_done")] = False
        st.session_state[ns_key(TOOL, "last_sig")] = current_sig

    # ── Concat ────────────────────────────────────────────────────────────────
    if st.button("Concat Sheets", type="primary", width="stretch", key=ns_key(TOOL, "concat_btn")):
        frames = []
        errors = []
        dropped_by_sheet = {}
        ordered_keys = [ref_key] + [k for k in unit_keys if k != ref_key]

        for key in ordered_keys:
            df, err = sheet_data[key]
            if err:
                errors.append(f"{key}: {err}")
                continue
            if key == ref_key:
                # Shallow: adding the Source column below must not touch the
                # cached upload, but the data itself needs no copy.
                aligned_df = df.copy(deep=False)
            else:
                aligned_df, dropped = align_to_schema(df, reference_cols, manual_maps.get(key))
                if dropped:
                    dropped_by_sheet[key] = dropped
            if include_provenance:
                aligned_df["Source"] = key
            frames.append(aligned_df)

        if not frames:
            st.error("Concat failed, no valid data.")
            st.stop()

        # Stored (not rendered here) so the warning survives later reruns
        # (e.g. clicking Download) instead of vanishing while results persist.
        st.session_state[ns_key(TOOL, "dropped_by_sheet")] = dropped_by_sheet
        concat_progress = st.progress(0, text="Generating Excel file...")

        def update_progress(written, total):
            pct = int(written / total * 100)
            concat_progress.progress(pct, text=f"Writing rows {written:,} of {total:,}")

        # Streamed frame by frame, no pd.concat copy of every sheet.
        try:
            excel_bytes = frames_to_excel_bytes(frames, progress_cb=update_progress)
        except ValueError as e:
            concat_progress.empty()
            st.error(str(e))
            st.stop()
        concat_progress.empty()

        st.session_state[ns_key(TOOL, "preview")] = pd.concat(
            [f.head(PREVIEW_ROWS) for f in frames], ignore_index=True,
        ).head(PREVIEW_ROWS)
        st.session_state[ns_key(TOOL, "shape")] = (sum(len(f) for f in frames), len(frames[0].columns))
        st.session_state[ns_key(TOOL, "sheets_done")] = len(frames)
        st.session_state[ns_key(TOOL, "excel_bytes")] = excel_bytes
        st.session_state[ns_key(TOOL, "concat_done")] = True
        st.session_state[ns_key(TOOL, "concat_errors")] = errors
        del frames, excel_bytes
        release_memory()

    # ── Results ──────────────────────────────────────────────────────────────
    if st.session_state.get(ns_key(TOOL, "concat_done")):
        total_rows, total_cols = st.session_state[ns_key(TOOL, "shape")]
        errors = st.session_state.get(ns_key(TOOL, "concat_errors"), [])

        render_dropped_columns_warning(st.session_state.get(ns_key(TOOL, "dropped_by_sheet"), {}))
        if errors:
            st.warning(
                "Skipped sheet(s) that could not be read:\n\n" + "\n\n".join(f"- {e}" for e in errors)
            )

        k1, k2, k3, k4 = st.columns([1, 1, 1, 2])
        k1.metric("Total Rows", f"{total_rows:,}")
        k2.metric("Columns", total_cols)
        k3.metric("Sheets Concatenated", st.session_state[ns_key(TOOL, "sheets_done")])
        with k4:
            render_download_section(
                TOOL,
                "concatenated_output.xlsx",
                st.session_state[ns_key(TOOL, "excel_bytes")],
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        render_success_banner("Concat complete")

        with st.expander(f"Preview, first {PREVIEW_ROWS} rows", expanded=True):
            st.dataframe(st.session_state[ns_key(TOOL, "preview")], width="stretch", hide_index=True)
