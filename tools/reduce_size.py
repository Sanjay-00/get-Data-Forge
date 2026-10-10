import streamlit as st

from core.excel_writer import to_excel_bytes
from core.reduce_size import build_reduction_report, downcast_numeric_columns, drop_empty_columns, inspect_workbook_bloat
from core.session import cached_read, cached_sheet_names, file_key, ns_key, release_unused_uploads, remember_cache_entry
from ui.shared import render_download_section, render_success_banner, render_tool_hero, section_label

TOOL = "reduce"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Both steps below used to run on every rerun of the page (any checkbox
# click, or typing a download filename): a full workbook rewrite plus an
# openpyxl load of the original. Cached per session, keyed on the upload and
# the options, they run once per distinct choice.
# Small caps: every reduce entry holds a whole rebuilt file.
_RESULT_CACHE = dict(scope="session", max_entries=3, ttl="1h")


@st.cache_resource(show_spinner="Inspecting original file for the before/after report...", **_RESULT_CACHE)
def _cached_bloat(key, filename, sheet_name, _raw_bytes):
    return inspect_workbook_bloat(_raw_bytes, filename, sheet_name=sheet_name)


@st.cache_resource(show_spinner="Building the reduced file...", **_RESULT_CACHE)
def _cached_reduce(key, sheet_choice, drop_empty, do_downcast, output_format, _df):
    working_df = _df
    columns_dropped = 0
    if drop_empty:
        working_df, columns_dropped = drop_empty_columns(working_df)
    downcast_info = None
    if do_downcast:
        working_df, downcast_info = downcast_numeric_columns(working_df)
    if output_format == "csv":
        # utf-8-sig: Excel only auto-detects UTF-8 CSVs when a BOM is
        # present; plain utf-8 garbles non-ASCII text on open.
        output_bytes = working_df.to_csv(index=False).encode("utf-8-sig")
    else:
        # Maximum compression: the smallest file is this tool's whole job.
        output_bytes = to_excel_bytes(working_df, compresslevel=9)
    return output_bytes, columns_dropped, downcast_info


def render_reduce_size_page():
    render_tool_hero("📉", "Reduce Size", "Shrink an Excel file with an honest before/after breakdown of what was removed", "#10B981")

    uploaded_file = st.file_uploader(
        "upload", type=["xlsx", "xls", "xlsb"],
        accept_multiple_files=False, label_visibility="collapsed",
        key=ns_key(TOOL, "uploader"),
    )
    release_unused_uploads(uploaded_file)

    if not uploaded_file:
        st.stop()

    raw_bytes = uploaded_file.getvalue()
    before_bytes = len(raw_bytes)

    with st.spinner(f"Reading {uploaded_file.name}..."):
        sheets = cached_sheet_names(uploaded_file)

    sheet_choice = 0
    if len(sheets) > 1:
        sheet_choice = st.selectbox("Sheet", options=sheets, index=0, key=ns_key(TOOL, "sheet"))
    sheet_name_for_inspection = sheet_choice if isinstance(sheet_choice, str) else None

    with st.spinner("Reading sheet data..."):
        df, err = cached_read(uploaded_file, sheet_choice)
    if err:
        st.error(f"Cannot read file: {err}")
        st.stop()

    key = file_key(uploaded_file)
    remember_cache_entry(key, _cached_bloat, key, uploaded_file.name, sheet_name_for_inspection, None)
    bloat_info = _cached_bloat(key, uploaded_file.name, sheet_name_for_inspection, raw_bytes)
    if bloat_info["error"]:
        st.warning(f"Could not fully inspect the original file for reporting purposes: {bloat_info['error']}")

    # ── Options ───────────────────────────────────────────────────────────────
    section_label("Reduction Options", None)
    opt1, opt2 = st.columns(2)
    with opt1:
        drop_empty = st.checkbox("Drop fully-empty columns", value=True, key=ns_key(TOOL, "drop_empty"))
        do_downcast = st.checkbox(
            "Downcast numeric dtypes (skips any column where it would lose precision)",
            value=False, key=ns_key(TOOL, "downcast"),
        )
    with opt2:
        output_format = st.radio(
            "Output format", options=["xlsx", "csv"], horizontal=True, key=ns_key(TOOL, "output_format"),
            help="CSV drops the workbook/sheet/formatting concept entirely for maximum size reduction.",
        )

    remember_cache_entry(key, _cached_reduce, key, sheet_choice, drop_empty, do_downcast, output_format, None)
    try:
        output_bytes, columns_dropped, downcast_info = _cached_reduce(
            key, sheet_choice, drop_empty, do_downcast, output_format, df,
        )
    except ValueError as e:
        st.error(str(e))
        st.stop()
    if output_format == "csv":
        file_name, mime = "reduced_output.csv", "text/csv"
    else:
        file_name, mime = "reduced_output.xlsx", XLSX_MIME

    after_bytes = len(output_bytes)
    report = build_reduction_report(before_bytes, after_bytes, bloat_info, columns_dropped, downcast_info, output_format)

    # ── Report ───────────────────────────────────────────────────────────────
    section_label("Before / After", f"{report['saved_pct']}% smaller", "#0EA5E9")
    m1, m2, m3 = st.columns(3)
    m1.metric("Original Size", f"{before_bytes / 1024:,.1f} KB")
    m2.metric("New Size", f"{after_bytes / 1024:,.1f} KB")
    m3.metric("Saved", f"{report['saved_bytes'] / 1024:,.1f} KB")

    st.markdown("**What changed and why:**")
    for line in report["details"]:
        st.markdown(f"- {line}")

    render_download_section(TOOL, file_name, output_bytes, mime)
    render_success_banner("Reduction complete")
