import datetime as _dt
import io
import re
import zipfile

import numpy as np
import pandas as pd

from core.dedupe import dedupe_columns

# Some non-Excel exporters (e.g. loan/ERP report generators) emit an xlsx
# whose styles.xml has a self-closing <fill/> with no patternFill/gradientFill
# child. That's valid-looking XML but violates the schema openpyxl expects,
# so it raises TypeError: expected <class 'openpyxl.styles.fills.Fill'>
# instead of a readable parse error. Patch it in-place before handing the
# bytes to openpyxl/pandas.
_EMPTY_FILL_RE = re.compile(rb"<((?:\w+:)?)fill\s*/>")


def _repair_xlsx_bytes(raw):
    try:
        zin = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        return None
    if "xl/styles.xml" not in zin.namelist():
        return None
    styles = zin.read("xl/styles.xml")

    def repl(m):
        prefix = m.group(1)
        return b"<" + prefix + b'fill><' + prefix + b'patternFill patternType="none"/></' + prefix + b"fill>"

    patched = _EMPTY_FILL_RE.sub(repl, styles)
    if patched == styles:
        return None

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = patched if item.filename == "xl/styles.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    return out.getvalue()


def _with_repair(uploaded_file, func, engine):
    try:
        return func(uploaded_file)
    except Exception:
        if engine != "openpyxl":
            raise
        raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
        repaired = _repair_xlsx_bytes(raw)
        if repaired is None:
            raise
        buf = io.BytesIO(repaired)
        buf.name = getattr(uploaded_file, "name", "")
        return func(buf)


def _get_engine(filename):
    ext = filename.rsplit(".", 1)[-1].lower()
    return {"xls": "xlrd", "xlsb": "pyxlsb"}.get(ext, "openpyxl")


def _rewind(uploaded_file):
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)


def _raw_bytes(uploaded_file):
    if hasattr(uploaded_file, "getvalue"):
        return uploaded_file.getvalue()
    _rewind(uploaded_file)
    return uploaded_file.read()


def _read_with_fallback(uploaded_file, func):
    """calamine (Rust) first, for .xlsx, .xls and .xlsb alike. Anything it
    can't parse falls back to the format's original pure-Python engine,
    which keeps the styles.xml repair path for broken exporter output."""
    try:
        return func(uploaded_file, "calamine")
    except Exception:
        _rewind(uploaded_file)
        engine = _get_engine(uploaded_file.name)
        return _with_repair(uploaded_file, lambda f: func(f, engine), engine)


# ── Header naming ────────────────────────────────────────────────────────
# Both read paths below produce the same column labels pandas' own header
# parsing would: blank header cells become "Unnamed: <position>", and
# repeats get ".1", ".2" suffixes.

def _header_label(value, position):
    if value is None or (isinstance(value, float) and value != value) or value == "":
        return f"Unnamed: {position}"
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _dedupe_like_pandas(labels):
    """pandas' duplicate header mangling: Name, Name.1, Name.2, skipping any
    suffix that is already taken."""
    counts = {}
    out = []
    for label in labels:
        count = counts.get(label, 0)
        while count > 0:
            counts[label] = count + 1
            label = f"{label}.{count}"
            count = counts.get(label, 0)
        out.append(label)
        counts[label] = count + 1
    return out


_FASTEXCEL_DUP_RE = re.compile(r"^(.*)_(\d+)$")


def _fastexcel_labels(column_infos):
    """fastexcel names blank headers __UNNAMED__<i> and repeats Name_1,
    Name_2. Undo that to the raw header values, then label them the pandas
    way so both read paths agree."""
    raw = []
    seen = {}
    for info in column_infos:
        name = info.name
        if info.column_name_from == "generated":
            raw.append(None)
            continue
        m = _FASTEXCEL_DUP_RE.match(name)
        if m and seen.get(m.group(1), 0) == int(m.group(2)):
            name = m.group(1)
        seen[name] = seen.get(name, 0) + 1
        raw.append(name)
    return _dedupe_like_pandas([_header_label(v, i) for i, v in enumerate(raw)])


# ── Fast path: fastexcel ─────────────────────────────────────────────────

_EXCEL_TIME_BASE = {pd.Timestamp("1899-12-30").date(), pd.Timestamp("1899-12-31").date()}


def _to_pandas_semantics(df):
    """fastexcel's Arrow columns -> the dtypes the rest of the app expects
    (the same ones pandas' per-cell parsing produces)."""
    out = {}
    for col in df.columns:
        s = df[col]
        kind = s.dtype.kind
        if kind == "f":
            # Excel stores every number as a float; whole-number columns
            # with no blanks are ints, as pandas reads them.
            vals = s.to_numpy()
            if len(vals) and not np.isnan(vals).any() and (np.abs(vals) < 2**53).all() and (vals == np.floor(vals)).all():
                s = s.astype("int64")
        elif kind == "M":
            non_null = s.dropna()
            if len(non_null) and non_null.dt.normalize().dt.date.isin(_EXCEL_TIME_BASE).all():
                # time-of-day cells arrive as a datetime on Excel's day zero
                s = s.map(lambda v: v.time(), na_action="ignore").astype(object)
            else:
                if getattr(s.dt, "tz", None) is None:
                    s = s.astype("datetime64[ns]")
        elif kind == "O":
            if s.isna().all():
                s = pd.Series(np.nan, index=s.index)
            else:
                first = s.dropna().iloc[0]
                if isinstance(first, _dt.date) and not isinstance(first, _dt.datetime):
                    s = pd.to_datetime(s)
                else:
                    s = s.where(s.notna(), np.nan)
        out[col] = s
    return pd.DataFrame(out, index=df.index)


def _read_fastexcel(raw, sheet):
    """Whole-sheet read through fastexcel (calamine -> Arrow): ~3x faster
    than pandas' calamine reader and about half the peak memory, because
    cells go straight into typed columns instead of one Python object each.

    dtype_coercion="strict" makes any column holding more than one type
    (numbers and text, dates and text...) raise instead of being silently
    turned into text, so those sheets take the per-cell path below."""
    import fastexcel

    sheet_data = fastexcel.read_excel(raw).load_sheet(
        sheet, schema_sample_rows=None, dtype_coercion="strict",
    )
    labels = _fastexcel_labels(sheet_data.available_columns())
    df = sheet_data.to_pandas()
    df.columns = range(len(df.columns))
    # A header-only sheet has no cells to type; pandas leaves those columns
    # as object, so do the same.
    df = df.astype(object) if df.empty else _to_pandas_semantics(df)
    df.columns = labels
    return df


# ── Per-cell path: pandas ────────────────────────────────────────────────

def _read_per_cell(uploaded_file, sheet):
    """Keeps each cell's own type, so mixed columns stay mixed (real numbers
    next to text). Header rules match the fast path: leading blank rows and
    columns are skipped, the first non-empty row is the header. Text is
    never reinterpreted: "007" stays "007" and a cell reading "NA" stays
    "NA" (pandas' defaults would turn them into 7 and a blank)."""
    grid = _read_with_fallback(
        uploaded_file,
        lambda f, engine: pd.read_excel(
            f, sheet_name=sheet, engine=engine, header=None, dtype=object,
            keep_default_na=False, na_values=[""],
        ),
    )
    filled = grid.notna().to_numpy()
    if not filled.any():
        return pd.DataFrame()
    # Only *leading* blank rows/columns are trimmed, matching where
    # calamine's used range starts on the fast path; blank columns in the
    # middle of the table stay (as "Unnamed: n"), same as before.
    first_row = int(filled.any(axis=1).argmax())
    first_col = int(filled.any(axis=0).argmax())
    grid = grid.iloc[first_row:, first_col:]
    header = grid.iloc[0].tolist()
    data = grid.iloc[1:].reset_index(drop=True)
    data.columns = _dedupe_like_pandas([_header_label(v, i) for i, v in enumerate(header)])
    return data.infer_objects()


_HEADER_SCAN_ROWS = 20


def _is_blank(v):
    return v is None or v == "" or (isinstance(v, float) and v != v) or v is pd.NaT


def _promote_header_row(df):
    """Report exports often put a title ("Sales Report - Jan 2024") and a
    few metadata lines above the real header. Read naively, the title
    becomes the header: one named column and the rest "Unnamed: n", so
    Merge/Concat can't match a single column. When most labels are
    placeholders, look in the first rows for the real header (a row that
    is at least half filled and mostly text) and promote it."""
    labels = [str(c) for c in df.columns]
    width = len(labels)
    unnamed = sum(label.startswith("Unnamed: ") for label in labels)
    if width < 2 or unnamed * 2 <= width:
        return df
    head = df.head(_HEADER_SCAN_ROWS)
    for i, row in enumerate(head.itertuples(index=False, name=None)):
        filled = [v for v in row if not _is_blank(v)]
        texts = sum(isinstance(v, str) for v in filled)
        if len(filled) >= max(2, (width + 1) // 2) and texts >= 0.8 * len(filled):
            data = df.iloc[i + 1:].reset_index(drop=True)
            data.columns = _dedupe_like_pandas([_header_label(v, j) for j, v in enumerate(row)])
            # Columns held the title/header text too; with those rows gone,
            # let them settle to their real dtypes (numbers, dates...).
            return data.astype(object).infer_objects()
    return df


def get_sheet_names(uploaded_file):
    try:
        import fastexcel

        return fastexcel.read_excel(_raw_bytes(uploaded_file)).sheet_names
    except Exception:
        pass
    try:
        _rewind(uploaded_file)
        return _read_with_fallback(uploaded_file, lambda f, engine: pd.ExcelFile(f, engine=engine).sheet_names)
    except Exception:
        return []


def read_single_file(uploaded_file, sheet_name_or_index):
    try:
        try:
            df = _read_fastexcel(_raw_bytes(uploaded_file), sheet_name_or_index)
        except Exception:
            _rewind(uploaded_file)
            df = _read_per_cell(uploaded_file, sheet_name_or_index)
        df = _promote_header_row(df)
        # Stringify every label (headers can be numbers) before stripping.
        df.columns = [str(c).strip() for c in df.columns]
        df, _ = dedupe_columns(df)
        df.dropna(how="all", inplace=True)
        return df, None
    except Exception as e:
        return None, str(e)
