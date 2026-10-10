import streamlit as st

from core.excel_writer import tables_to_excel_bytes, to_excel_bytes
from core.memory import release_memory
from core.pdf_extract import (
    coerce_numeric_columns,
    extract_pdf_tables,
    ocr_backend,
    rows_to_dataframe,
    stack_tables,
)
from core.session import file_key, ns_key, release_unused_uploads, remember_cache_entry
from ui.shared import render_download_section, render_success_banner, render_tool_hero, section_label

TOOL = "pdf"

PREVIEW_ROW_CAP = 50


ENGINE_LABELS = {
    "Auto (recommended)": "auto",
    "Vector ruling lines only": "lines",
    "Text alignment only (borderless tables)": "text",
}


@st.cache_resource(scope="session", max_entries=4, ttl="1h", show_spinner=False)
def _cached_extract(key, engine, use_ocr, _pdf_bytes):
    # Keyed on the upload's file_key rather than hashing the whole PDF on
    # every rerun; session-scoped and bounded so extracted tables don't
    # outlive the visitor. The result is shared, so callers only read it.
    return extract_pdf_tables(_pdf_bytes, engine=engine, use_ocr=use_ocr)


def render_pdf_page():
    render_tool_hero("📄", "PDF to Excel", "Pull tables out of a text-layer PDF into a real Excel workbook", "#EF4444")

    uploaded_file = st.file_uploader(
        "upload", type=["pdf"],
        accept_multiple_files=False, label_visibility="collapsed",
        key=ns_key(TOOL, "uploader"),
    )
    release_unused_uploads(uploaded_file)

    if not uploaded_file:
        st.stop()

    engine_label = st.selectbox(
        "Table detection engine", options=list(ENGINE_LABELS),
        key=ns_key(TOOL, "engine"),
        help="Auto reads bordered tables straight from the PDF's ruling lines (fast), "
             "and switches to image-based grid detection for merged/spanning cells, "
             "borderless tables and scans. Pick a specific engine if Auto splits or "
             "misses your table.",
    )
    engine = ENGINE_LABELS[engine_label]

    with st.spinner(f"Scanning {uploaded_file.name} for tables..."):
        remember_cache_entry(file_key(uploaded_file), _cached_extract, file_key(uploaded_file), engine, False, None)
        tables, page_count, has_any_text, err = _cached_extract(file_key(uploaded_file), engine, False, uploaded_file.getvalue())

    if err:
        st.error(f"Cannot read PDF: {err}")
        st.stop()

    # ── Scanned PDF: no text layer, so the only route is OCR ────────────────
    ocr_used = False
    if not tables and not has_any_text:
        backend = ocr_backend()
        if backend:
            st.warning(
                "No extractable text found in this PDF - it looks like a scan or "
                "photo of a document. OCR can read it instead (slower, and accuracy "
                "depends on scan quality)."
            )
            run_ocr = st.checkbox(
                f"Run OCR text recognition ({backend})",
                value=False, key=ns_key(TOOL, "run_ocr"),
            )
            if not run_ocr:
                st.stop()
            with st.spinner("Running OCR - this can take a while on multi-page scans..."):
                remember_cache_entry(file_key(uploaded_file), _cached_extract, file_key(uploaded_file), engine, True, None)
                tables, page_count, has_any_text, err = _cached_extract(
                    file_key(uploaded_file), engine, True, uploaded_file.getvalue(),
                )
            if err:
                st.error(f"OCR failed: {err}")
                st.stop()
            ocr_used = True
            if not tables:
                st.error(
                    "OCR ran, but no tables were recognized. The scan may be too "
                    "low-resolution, skewed, or the table may have no visible grid lines."
                )
                st.stop()
            st.info(
                "Tables below were read by OCR from page images. Review them more "
                "carefully than usual: OCR can misread characters (0/O, 1/l, 5/S), "
                "and those errors look exactly like real data."
            )
        else:
            st.error(
                "No extractable text found in this PDF - it looks like a scan or "
                "photo of a document. Reading it requires OCR, and no OCR backend "
                "was found. Install one with: pip install rapidocr onnxruntime"
            )
            st.stop()

    if not tables:
        st.warning(
            f"No tables detected across {page_count} page(s) with the "
            f"'{engine_label}' engine. The PDF has a text layer, so try a "
            "different detection engine above."
        )
        st.stop()

    # ── Options ───────────────────────────────────────────────────────────────
    section_label("Options", None)
    opt1, opt2 = st.columns(2)
    with opt1:
        first_row_header = st.checkbox(
            "First row of each table is the header",
            value=True, key=ns_key(TOOL, "first_row_header"),
        )
        coerce_numeric = st.checkbox(
            "Convert all-numeric text columns to real numbers",
            value=True, key=ns_key(TOOL, "coerce_numeric"),
            help="PDF cells always arrive as text. This converts columns where every "
                 "non-blank cell parses as a number, so Excel gets real numbers. Turn it "
                 "off if you have ID codes with leading zeros (e.g. 007) that must stay text.",
        )
    with opt2:
        output_mode = st.radio(
            "Output", options=["One sheet per table", "Stack all tables into one sheet"],
            key=ns_key(TOOL, "output_mode"),
            help="Stacking is for one logical table that continues across pages: fragments "
                 "are appended positionally under the first table's headers.",
        )

    # ── Per-table preview + selection ────────────────────────────────────────
    section_label("Detected Tables", f"{len(tables)} table(s) on {page_count} page(s)", "#0EA5E9")

    selected = []  # (label, coerced df for per-sheet output, raw all-text df for stacking)
    all_converted = set()
    for t in tables:
        raw_df = rows_to_dataframe(t["rows"], first_row_header)
        df = raw_df
        if coerce_numeric and len(df):
            df, converted = coerce_numeric_columns(df)
            all_converted.update(converted)
        label = f"Page {t['page']} Table {t['index'] + 1}"
        title_part = f"  ·  {t['title']}" if t.get("title") else ""
        with st.expander(f"{label}{title_part}  ({len(df)} rows x {len(df.columns)} columns)", expanded=len(tables) == 1):
            include = st.checkbox(
                "Include in output", value=True,
                key=ns_key(TOOL, f"include_{t['page']}_{t['index']}"),
            )
            if len(df) > PREVIEW_ROW_CAP:
                st.caption(f"Showing first {PREVIEW_ROW_CAP} of {len(df):,} rows.")
            st.dataframe(df.head(PREVIEW_ROW_CAP), width="stretch", hide_index=True)
        if include:
            selected.append((label, df, raw_df))

    if all_converted:
        st.caption(f"Converted to numeric: {', '.join(sorted(all_converted))}")

    if not selected:
        st.info("Select at least one table to include in the output.")
        st.stop()

    current_sig = (
        file_key(uploaded_file), engine, ocr_used, first_row_header, coerce_numeric, output_mode,
        tuple(label for label, _, _ in selected),
    )
    if st.session_state.get(ns_key(TOOL, "last_sig")) != current_sig:
        st.session_state[ns_key(TOOL, "pdf_done")] = False
        st.session_state[ns_key(TOOL, "last_sig")] = current_sig

    # ── Convert ──────────────────────────────────────────────────────────────
    if st.button("Convert to Excel", type="primary", width="stretch", key=ns_key(TOOL, "convert_btn")):
        try:
            if output_mode == "Stack all tables into one sheet":
                # Stack the raw all-text fragments, then coerce the combined
                # result: coercing per fragment first can leave one page's
                # column numeric and another's text (e.g. a TOTAL row on the
                # last page blocks only that fragment's coercion), producing
                # a mixed-dtype stacked column.
                stacked, notes, stack_err = stack_tables([raw for _, _, raw in selected])
                if stack_err:
                    st.error(stack_err)
                    st.stop()
                for note in notes:
                    st.warning(note)
                if coerce_numeric and len(stacked):
                    stacked, _ = coerce_numeric_columns(stacked)
                excel_bytes = to_excel_bytes(stacked, sheet_name="Data")
                total_rows = len(stacked)
            else:
                progress = st.progress(0, text="Writing sheets...")

                def update_progress(done, total):
                    progress.progress(int(done / total * 100), text=f"Writing sheet {done} of {total}")

                excel_bytes = tables_to_excel_bytes(
                    [(label, df) for label, df, _ in selected], progress_cb=update_progress,
                )
                progress.empty()
                total_rows = sum(len(df) for _, df, _ in selected)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        st.session_state[ns_key(TOOL, "excel_bytes")] = excel_bytes
        st.session_state[ns_key(TOOL, "total_rows")] = total_rows
        st.session_state[ns_key(TOOL, "pdf_done")] = True
        del excel_bytes
        release_memory()

    # ── Results ──────────────────────────────────────────────────────────────
    if st.session_state.get(ns_key(TOOL, "pdf_done")):
        k1, k2, k3 = st.columns([1, 1, 2])
        k1.metric("Tables Exported", len(selected))
        k2.metric("Total Rows", f"{st.session_state[ns_key(TOOL, 'total_rows')]:,}")
        with k3:
            render_download_section(
                TOOL,
                "pdf_tables.xlsx",
                st.session_state[ns_key(TOOL, "excel_bytes")],
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        render_success_banner("Conversion complete")
