"""Streaming .xlsx writer.

Builds the worksheet XML directly, a chunk of rows at a time, instead of
going through xlsxwriter's per-cell write() calls. On a 200k row x 15 column
sheet that per-cell dispatch was ~21s; formatting whole columns at once and
streaming the XML into the zip is several times faster and memory stays
bounded by the chunk size, not the sheet size.

Output is a minimal but complete SpreadsheetML package (content types, rels,
workbook, styles, one XML part per sheet). Strings are written inline, never
as formulas or hyperlinks, so a text cell like "=SUM(A1)" or "http://..."
stays exactly the text it was in the source data.
"""

import datetime as _dt
import decimal
import io
import math
import re
import threading
import zipfile
from xml.sax.saxutils import escape as _xml_escape

import numpy as np
import pandas as pd

# Excel's hard per-sheet row limit (including the header row).
EXCEL_MAX_ROWS = 1_048_576
# Excel refuses to open a file with a longer string in a cell.
EXCEL_MAX_STRING = 32_767
# Last column is XFD; a wider sheet makes a file Excel refuses to open.
EXCEL_MAX_COLUMNS = 16_384

_CHUNK_ROWS = 10_000

# Deflate level for the zip container. Measured on a 200k x 15 sheet:
# level 6 (zip's default) 4.2s / 15.5 MB, level 4 2.8s / 17.2 MB, level 1
# 2.4s / 20.5 MB. 4 is a third faster for an ~11% bigger file; the Reduce
# Size tool, where file size is the whole point, asks for 9 instead.
DEFAULT_COMPRESSLEVEL = 4

# Excel forbids []:*?/\ in sheet names, 31 chars max, and blank names.
_SHEET_NAME_BAD_CHARS = re.compile(r"[\[\]:*?/\\]")

# Control characters are illegal in XML 1.0, so they are written using
# Excel's own _xHHHH_ escape (same thing xlsxwriter does). Literal text that
# already looks like an escape gets its underscore escaped so Excel shows it
# verbatim instead of decoding it.
_ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_LOOKS_LIKE_ESCAPE = re.compile(r"_(x[0-9A-Fa-f]{4}_)")

# Lotus 1-2-3 leap year bug: Excel serial 60 is the nonexistent 1900-02-29,
# so for real dates from 1900-03-01 onward the effective epoch is 1899-12-30.
_EPOCH = pd.Timestamp("1899-12-30")
_EPOCH_DT = _dt.datetime(1899, 12, 30)

# cellXfs indexes in _STYLES_XML.
_STYLE_DATE = 1
_STYLE_TIME = 2
_STYLE_DURATION = 3

_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="3">'
    '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'
    '<numFmt numFmtId="165" formatCode="hh:mm:ss"/>'
    '<numFmt numFmtId="166" formatCode="[h]:mm:ss"/>'
    "</numFmts>"
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="4">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    "</cellXfs>"
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)

_ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>"
)

_SHEET_HEAD = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    "<sheetData>"
)
_SHEET_TAIL = "</sheetData></worksheet>"


def safe_sheet_name(name, fallback="Sheet1"):
    name = _SHEET_NAME_BAD_CHARS.sub("_", str(name)).strip()
    return (name or fallback)[:31]


def _col_letter(idx):
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _escape_text(s):
    if len(s) > EXCEL_MAX_STRING:
        s = s[:EXCEL_MAX_STRING]
    s = _xml_escape(s)
    if "_x" in s:
        s = _LOOKS_LIKE_ESCAPE.sub(r"_x005F_\1", s)
    if _ILLEGAL_XML_CHARS.search(s):
        s = _ILLEGAL_XML_CHARS.sub(lambda m: f"_x{ord(m.group()):04X}_", s)
    return s


# Shared string table for the workbook being written. Each distinct string
# is stored (and XML-escaped) once in xl/sharedStrings.xml and cells hold
# its index, the way Excel itself saves files: repeated values like branch
# codes or statuses no longer repeat in the sheet XML, which makes the XML
# smaller, faster to compress and the file smaller. Thread-local because
# every Streamlit session writes from its own thread.
_local = threading.local()


def _str_cell(ref, s):
    if len(s) > EXCEL_MAX_STRING:
        s = s[:EXCEL_MAX_STRING]
    sst = _local.sst
    idx = sst.get(s)
    if idx is None:
        idx = sst[s] = len(sst)
    return f'<c r="{ref}" t="s"><v>{idx}</v></c>'


def _write_shared_strings(zf, sst):
    with zf.open("xl/sharedStrings.xml", "w", force_zip64=True) as out:
        out.write((
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'uniqueCount="{len(sst)}">'
        ).encode())
        batch = []
        for text in sst:
            batch.append(f'<si><t xml:space="preserve">{_escape_text(text)}</t></si>')
            if len(batch) >= 10_000:
                out.write("".join(batch).encode("utf-8", "replace"))
                batch = []
        out.write("".join(batch).encode("utf-8", "replace"))
        out.write(b"</sst>")


def _num_cell(ref, v):
    if math.isinf(v):
        return f'<c r="{ref}" t="e"><v>#NUM!</v></c>'
    return f'<c r="{ref}"><v>{v!r}</v></c>'


def _serial_from_datetime(v):
    delta = v - _EPOCH_DT
    serial = delta.days + delta.seconds / 86400 + delta.microseconds / 86_400_000_000
    return serial - 1 if serial < 61 else serial


def _object_cell(ref, v):
    """Per-value dispatch for object (mixed type) columns."""
    if type(v) is str:
        return _str_cell(ref, v) if v else ""
    if v is None:
        return ""
    if isinstance(v, str):
        return _str_cell(ref, v) if v else ""
    if isinstance(v, (bool, np.bool_)):
        return f'<c r="{ref}" t="b"><v>{int(v)}</v></c>'
    if isinstance(v, (int, np.integer)):
        return f'<c r="{ref}"><v>{int(v)}</v></c>'
    if isinstance(v, (float, np.floating)):
        v = float(v)
        return "" if v != v else _num_cell(ref, v)
    if isinstance(v, decimal.Decimal):
        return "" if v.is_nan() else _num_cell(ref, float(v))
    if v is pd.NaT or v is pd.NA:
        return ""
    if isinstance(v, _dt.datetime):
        if v.tzinfo is not None:
            v = v.replace(tzinfo=None)
        if isinstance(v, pd.Timestamp):
            v = v.to_pydatetime()
        serial = _serial_from_datetime(v)
        if serial < 0:
            return _str_cell(ref, v.isoformat(sep=" "))
        return f'<c r="{ref}" s="{_STYLE_DATE}"><v>{serial!r}</v></c>'
    if isinstance(v, _dt.date):
        return _object_cell(ref, _dt.datetime(v.year, v.month, v.day))
    if isinstance(v, _dt.time):
        frac = (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
        return f'<c r="{ref}" s="{_STYLE_TIME}"><v>{frac!r}</v></c>'
    if isinstance(v, (_dt.timedelta, pd.Timedelta)):
        return f'<c r="{ref}" s="{_STYLE_DURATION}"><v>{v.total_seconds() / 86400!r}</v></c>'
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return _str_cell(ref, str(v))


def _column_cells(series, letter, first_row):
    """XML for every cell of one column slice, '' for blanks. Typed columns
    are formatted in one vectorized/comprehension pass; only object columns
    fall back to per-value type dispatch."""
    kind = series.dtype.kind
    rows = range(first_row, first_row + len(series))

    # Nullable extension dtypes (Int64, boolean, string...) hold pd.NA, which
    # the fast typed paths below would print literally; tz-aware datetimes
    # are the one extension dtype the "M" path handles itself.
    if not isinstance(series.dtype, np.dtype) and kind != "M":
        return [_object_cell(f"{letter}{r}", v) for r, v in zip(rows, series.tolist())]

    if kind in "iu":
        return [f'<c r="{letter}{r}"><v>{v}</v></c>' for r, v in zip(rows, series.tolist())]

    if kind == "b":
        return [f'<c r="{letter}{r}" t="b"><v>{int(v)}</v></c>' for r, v in zip(rows, series.tolist())]

    if kind == "f":
        values = series.to_numpy()
        if np.isinf(values).any():
            return [
                "" if v != v else _num_cell(f"{letter}{r}", v)
                for r, v in zip(rows, series.tolist())
            ]
        # No infinities (the usual case): format inline, skipping the
        # per-cell function call.
        return [
            "" if v != v else f'<c r="{letter}{r}"><v>{v!r}</v></c>'
            for r, v in zip(rows, series.tolist())
        ]

    if kind == "M":
        if getattr(series.dt, "tz", None) is not None:
            series = series.dt.tz_localize(None)
        serial = ((series - _EPOCH) / pd.Timedelta(days=1)).to_numpy()
        # Dates before 1900-03-01 sit before Excel's phantom 1900-02-29.
        serial = np.where(serial < 61, serial - 1, serial)
        out = []
        for r, v, raw in zip(rows, serial.tolist(), series.tolist()):
            if v != v:
                out.append("")
            elif v < 0:
                out.append(_str_cell(f"{letter}{r}", raw.isoformat(sep=" ")))
            else:
                out.append(f'<c r="{letter}{r}" s="{_STYLE_DATE}"><v>{v!r}</v></c>')
        return out

    if kind == "m":
        days = (series / pd.Timedelta(days=1)).tolist()
        return [
            "" if v != v else f'<c r="{letter}{r}" s="{_STYLE_DURATION}"><v>{v!r}</v></c>'
            for r, v in zip(rows, days)
        ]

    return [_object_cell(f"{letter}{r}", v) for r, v in zip(rows, series.tolist())]


def _write_sheet_xml(out, frames, progress_cb=None):
    """Stream one worksheet's XML into the writable zip member `out`.
    frames is a list of DataFrames sharing the same columns, written one
    after another as a single sheet (so callers stacking many files never
    need a pd.concat copy of all of them)."""
    columns = list(frames[0].columns) if frames else []
    letters = [_col_letter(i) for i in range(len(columns))]
    out.write(_SHEET_HEAD.encode())

    header = "".join(_str_cell(f"{l}1", str(c)) for l, c in zip(letters, columns) if str(c))
    out.write(f'<row r="1">{header}</row>'.encode())

    total = sum(len(f) for f in frames)
    written = 0
    for df in frames:
        for start in range(0, len(df), _CHUNK_ROWS):
            chunk = df.iloc[start:start + _CHUNK_ROWS]
            first_row = written + 2
            cols = [_column_cells(chunk.iloc[:, ci], letters[ci], first_row) for ci in range(len(letters))]
            lines = [
                f'<row r="{first_row + i}">{"".join(cells)}</row>'
                for i, cells in enumerate(zip(*cols))
            ] if cols else [f'<row r="{first_row + i}"/>' for i in range(len(chunk))]
            # "replace": a lone surrogate (from badly decoded source text)
            # becomes "?" instead of invalid UTF-8, which Excel rejects as a
            # corrupt file.
            out.write("".join(lines).encode("utf-8", "replace"))
            written += len(chunk)
            if progress_cb:
                progress_cb(written, total)

    out.write(_SHEET_TAIL.encode())


def _check_column_limit(df, label="Result"):
    if len(df.columns) > EXCEL_MAX_COLUMNS:
        raise ValueError(
            f"{label} has {len(df.columns):,} columns, more than Excel's limit of "
            f"{EXCEL_MAX_COLUMNS:,} per sheet. Export as CSV instead."
        )


def _check_row_limit(df, label="Result"):
    if not isinstance(df, int):
        _check_column_limit(df, label)
    rows = df if isinstance(df, int) else len(df)
    if rows > EXCEL_MAX_ROWS - 1:
        raise ValueError(
            f"{label} has {rows:,} rows, more than Excel's limit of "
            f"{EXCEL_MAX_ROWS - 1:,} data rows per sheet. Export as CSV instead."
        )


def _write_workbook(named_frames, sheet_progress_cb=None, row_progress_cb=None, compresslevel=DEFAULT_COMPRESSLEVEL):
    buf = io.BytesIO()
    sheet_names = [name for name, _ in named_frames]
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=compresslevel) as zf:
        overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/'
            f'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, len(sheet_names) + 1)
        )
        zf.writestr("[Content_Types].xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            f"{overrides}</Types>"
        ))
        zf.writestr("_rels/.rels", _ROOT_RELS_XML)
        sheets_xml = "".join(
            f'<sheet name="{_xml_escape(name, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>'
            for i, name in enumerate(sheet_names, start=1)
        )
        zf.writestr("xl/workbook.xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheets_xml}</sheets></workbook>"
        ))
        rels = "".join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/'
            f'2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(sheet_names) + 1)
        )
        n = len(sheet_names) + 1
        rels += (
            f'<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        )
        zf.writestr("xl/_rels/workbook.xml.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{rels}</Relationships>"
        ))
        zf.writestr("xl/styles.xml", _STYLES_XML)
        _local.sst = {}
        try:
            for i, (_, frames) in enumerate(named_frames, start=1):
                with zf.open(f"xl/worksheets/sheet{i}.xml", "w", force_zip64=True) as out:
                    _write_sheet_xml(out, frames, row_progress_cb)
                if sheet_progress_cb:
                    sheet_progress_cb(i, len(named_frames))
            _write_shared_strings(zf, _local.sst)
        finally:
            _local.sst = None
    return buf.getvalue()


def to_excel_bytes(df, progress_cb=None, sheet_name="Merged", compresslevel=DEFAULT_COMPRESSLEVEL):
    _check_row_limit(df)
    return _write_workbook(
        [(safe_sheet_name(sheet_name), [df])], row_progress_cb=progress_cb, compresslevel=compresslevel,
    )


def frames_to_excel_bytes(frames, progress_cb=None, sheet_name="Merged"):
    """Same output as to_excel_bytes(pd.concat(frames, ignore_index=True)),
    without materializing the concatenated frame. Every frame must have the
    same columns in the same order."""
    if frames:
        _check_column_limit(frames[0])
    _check_row_limit(sum(len(f) for f in frames))
    return _write_workbook([(safe_sheet_name(sheet_name), list(frames))], row_progress_cb=progress_cb)


def tables_to_excel_bytes(named_frames, progress_cb=None):
    """One workbook with one sheet per (name, df) pair. Sheet names are
    sanitized for Excel's rules and de-duplicated by numbering repeats."""
    for name, df in named_frames:
        _check_row_limit(df, label=f"Table '{name}'")

    used = {}
    named = []
    for name, df in named_frames:
        base = safe_sheet_name(name)
        seen = used.get(base.lower(), 0)
        used[base.lower()] = seen + 1
        sheet = base if seen == 0 else safe_sheet_name(f"{base[:27]} ({seen + 1})")
        named.append((sheet, [df]))
    return _write_workbook(named, sheet_progress_cb=progress_cb)
