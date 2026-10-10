import pandas as pd
import streamlit as st

from core.excel_writer import frames_to_excel_bytes
from core.memory import release_memory
from core.schema import align_to_schema, build_file_summary, detect_dtype_mismatches, get_unmatched_pairs
from core.session import cached_read_many, cached_sheet_names, file_key, ns_key, release_unused_uploads
from ui.shared import (
    render_download_section,
    render_dropped_columns_warning,
    render_dtype_mismatch_warning,
    render_success_banner,
    render_summary_bar,
    render_tool_hero,
    render_unmatched_ui,
    section_label,
    stop_on_duplicate_names,
)

PREVIEW_ROWS = 100

TOOL = "merge"


def render_merge_page():
    render_tool_hero("📑", "Merge Files", "Drop your Excel files · Get one consolidated master Excel file", "#6366F1")

    uploaded_files = st.file_uploader(
        "upload", type=["xlsx", "xls", "xlsb"],
        accept_multiple_files=True, label_visibility="collapsed",
        key=ns_key(TOOL, "uploader"),
    )
    release_unused_uploads(*(uploaded_files or []))

    # ── Summary bar ──────────────────────────────────────────────────────
    if uploaded_files:
        render_summary_bar(uploaded_files)

    if not uploaded_files:
        st.stop()

    stop_on_duplicate_names(uploaded_files)

    # ── Per-file sheet selection ──────────────────────────────────────────
    with st.spinner("Detecting sheets..."):
        all_sheet_names = {f.name: cached_sheet_names(f) for f in uploaded_files}

    has_multi_sheets = any(len(sheets) > 1 for sheets in all_sheet_names.values())

    file_sheet_map = {}
    if has_multi_sheets:
        with st.expander("Sheet selection (per file)", expanded=False):
            sheet_cols = st.columns(min(len(uploaded_files), 3))
            for i, f in enumerate(uploaded_files):
                sheets = all_sheet_names[f.name]
                if len(sheets) <= 1:
                    file_sheet_map[f.name] = 0
                    continue
                with sheet_cols[i % 3]:
                    file_sheet_map[f.name] = st.selectbox(
                        f.name, options=sheets, index=0,
                        key=ns_key(TOOL, f"sheet_{f.name}"),
                    )
    else:
        for f in uploaded_files:
            file_sheet_map[f.name] = 0

    # ── Read all files ─────────────────────────────────────────────────────
    file_data = {}
    warnings = []
    progress = st.progress(0, text="Analyzing files...")

    def read_progress(done, total):
        progress.progress(int(done / total * 100), text=f"Reading files ({done} of {total})")

    # New uploads are parsed in parallel; already-parsed ones come from cache.
    results = cached_read_many(
        [(f, file_sheet_map[f.name]) for f in uploaded_files], progress_cb=read_progress,
    )
    for f, result in zip(uploaded_files, results):
        file_data[f.name] = result

    progress.empty()

    # ── Reference file + Schema preview ────────────────────────────────────
    col_ref, col_schema = st.columns([1, 1], gap="medium")

    with col_ref:
        section_label("Reference File", "sets column schema", "#D97706")
        ref_name = st.selectbox(
            "ref", options=[f.name for f in uploaded_files], index=0,
            help="All other files are aligned to this file's columns.",
            label_visibility="collapsed",
            key=ns_key(TOOL, "ref_name"),
        )

    ref_df, ref_err = file_data[ref_name]

    if ref_err:
        st.error(f"Cannot read reference file: {ref_err}")
        st.stop()

    reference_cols = list(ref_df.columns)

    if not reference_cols:
        # e.g. the chosen sheet is genuinely empty -- reindexing every other
        # file onto zero columns would silently produce a merged output with
        # rows but no columns at all, and no clue why.
        st.error(
            f"**{ref_name}** has no columns to align to (the sheet may be empty). "
            "Pick a different reference file or sheet."
        )
        st.stop()

    with col_schema:
        section_label("Schema Preview", f"{len(reference_cols)} columns", "#6366F1")
        with st.expander(f"{ref_name}", expanded=False):
            st.write(reference_cols)

    # ── File summary ────────────────────────────────────────────────────────
    section_label("File Summary", f"{len(uploaded_files)} files", "#0EA5E9")

    summaries = []
    for f in uploaded_files:
        df, err = file_data[f.name]
        if err:
            summaries.append({
                "File": f.name, "Rows": "-", "Total Columns": "-",
                "Matched Columns": "-", "Missing Columns": "-", "Extra Columns (dropped)": "-",
            })
            warnings.append(f.name)
        else:
            summaries.append(build_file_summary(f.name, df, reference_cols))

    for w in warnings:
        st.warning(f"Could not read **{w}**")

    st.dataframe(pd.DataFrame(summaries), width="stretch", hide_index=True)

    # ── Unmatched column resolution ─────────────────────────────────────────
    all_unmatched = {}
    for f in uploaded_files:
        if f.name == ref_name:
            continue
        df, err = file_data[f.name]
        if err:
            continue
        pairs = get_unmatched_pairs(df, reference_cols)
        resolvable = [(fc, rc) for fc, rc in pairs if rc]
        if resolvable:
            all_unmatched[f.name] = resolvable

    manual_maps = render_unmatched_ui(all_unmatched, key_prefix=ns_key(TOOL, "map")) if all_unmatched else {}

    # Any input that changes the merged output invalidates a finished merge,
    # including the column mappings above (changing one after merging used
    # to keep showing and serving the old result). file_key, not just the
    # name, so re-uploading a different file under the same name counts.
    current_sig = (
        tuple(file_key(f) for f in uploaded_files), ref_name, str(file_sheet_map),
        repr(sorted((k, sorted(v.items())) for k, v in manual_maps.items())),
    )
    if st.session_state.get(ns_key(TOOL, "last_file_sig")) != current_sig:
        st.session_state[ns_key(TOOL, "merge_done")] = False
        st.session_state[ns_key(TOOL, "last_file_sig")] = current_sig

    # ── Merge ────────────────────────────────────────────────────────────────
    if st.button("Merge Files", type="primary", width="stretch", key=ns_key(TOOL, "merge_btn")):
        frames = []
        errors = []
        dropped_by_file = {}
        ordered_names = [ref_name] + [f.name for f in uploaded_files if f.name != ref_name]

        for name in ordered_names:
            df, err = file_data[name]
            if err:
                errors.append(f"{name}: {err}")
                continue
            if name == ref_name:
                frames.append(df)
            else:
                aligned_df, dropped = align_to_schema(df, reference_cols, manual_maps.get(name))
                frames.append(aligned_df)
                if dropped:
                    dropped_by_file[name] = dropped

        if not frames:
            st.error("Merge failed, no valid data.")
            st.stop()

        # Stored (not rendered here) so the warnings survive the rerun that
        # any later interaction triggers (e.g. clicking Download) instead of
        # vanishing while the results and download button persist.
        st.session_state[ns_key(TOOL, "dropped_by_file")] = dropped_by_file
        st.session_state[ns_key(TOOL, "dtype_mismatches")] = detect_dtype_mismatches(frames, reference_cols)
        merge_progress = st.progress(0, text="Generating Excel file...")

        def update_progress(written, total):
            pct = int(written / total * 100)
            merge_progress.progress(pct, text=f"Writing rows {written:,} of {total:,}")

        # Frames are streamed into the sheet one after another; a pd.concat
        # here would hold a second full copy of every file in memory.
        try:
            excel_bytes = frames_to_excel_bytes(frames, progress_cb=update_progress)
        except ValueError as e:
            merge_progress.empty()
            st.error(str(e))
            st.stop()
        merge_progress.empty()

        # Only what the results view shows is kept per session: the counts
        # and a small preview, never the full merged frame.
        st.session_state[ns_key(TOOL, "preview")] = pd.concat(
            [f.head(PREVIEW_ROWS) for f in frames], ignore_index=True,
        ).head(PREVIEW_ROWS)
        st.session_state[ns_key(TOOL, "shape")] = (sum(len(f) for f in frames), len(reference_cols))
        st.session_state[ns_key(TOOL, "files_merged")] = len(frames)
        st.session_state[ns_key(TOOL, "excel_bytes")] = excel_bytes
        st.session_state[ns_key(TOOL, "merge_done")] = True
        st.session_state[ns_key(TOOL, "merge_errors")] = errors
        del frames, excel_bytes
        release_memory()

    # ── Results ──────────────────────────────────────────────────────────────
    if st.session_state.get(ns_key(TOOL, "merge_done")):
        total_rows, total_cols = st.session_state[ns_key(TOOL, "shape")]
        errors = st.session_state.get(ns_key(TOOL, "merge_errors"), [])

        render_dropped_columns_warning(st.session_state.get(ns_key(TOOL, "dropped_by_file"), {}))
        render_dtype_mismatch_warning(st.session_state.get(ns_key(TOOL, "dtype_mismatches"), {}))
        if errors:
            st.warning(
                "Skipped file(s) that could not be read:\n\n" + "\n\n".join(f"- {e}" for e in errors)
            )

        k1, k2, k3, k4 = st.columns([1, 1, 1, 2])
        k1.metric("Total Rows", f"{total_rows:,}")
        k2.metric("Columns", total_cols)
        k3.metric("Files Merged", st.session_state[ns_key(TOOL, "files_merged")])
        with k4:
            render_download_section(
                TOOL,
                "merged_output.xlsx",
                st.session_state[ns_key(TOOL, "excel_bytes")],
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        render_success_banner()

        with st.expander(f"Preview, first {PREVIEW_ROWS} rows", expanded=True):
            st.dataframe(st.session_state[ns_key(TOOL, "preview")], width="stretch", hide_index=True)
