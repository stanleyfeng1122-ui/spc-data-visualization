"""
Top-of-sheet header / layout detection.

Functions for locating the "Dim. No." anchor cell, scanning the label column
for known metadata row labels, and finding where measurement data starts.
Also exposes :data:`KNOWN_META_HEADERS`, the set of recognised metadata
header names (lowercased for matching).
"""

from __future__ import annotations

import re

from .dimensions import _safe_num

# ---------------------------------------------------------------------------
# Known metadata header names (normalised to lowercase for matching)
# ---------------------------------------------------------------------------

KNOWN_META_HEADERS = {
    "build",
    "shipment date",
    "color",
    "config",
    "vendor serial number",
    "fabric thickness",
    "2d barcode",
    "1d barcode",
    "rm coil",
    "raw material",
    "start point",
}


# ---------------------------------------------------------------------------
# Header / layout detection
# ---------------------------------------------------------------------------


def _find_dim_no_cell(rows, max_scan_rows=50, max_scan_cols=30):
    """
    Scan the top-left area of a sheet looking for a cell that says "Dim. No."
    (case-insensitive).  Returns (row_1based, col_1based) or (None, None).
    """
    for ri in range(min(max_scan_rows, len(rows))):
        row = rows[ri]
        for ci in range(min(max_scan_cols, len(row))):
            val = row[ci].value
            if val is not None and re.match(r"dim\.?\s*no\.?", str(val).strip(), re.IGNORECASE):
                return ri + 1, ci + 1  # 1-based
    return None, None


def _scan_label_rows(rows, label_col, start_row, end_row):
    """
    Scan a label column for known metadata row labels.
    Returns a dict: normalised_label -> row_1based.
    """
    mapping = {}
    for ri in range(start_row - 1, min(end_row, len(rows))):
        row = rows[ri]
        if label_col - 1 >= len(row):
            continue
        val = row[label_col - 1].value
        if val is None:
            continue
        s = str(val).strip().lower()
        # Normalise common variants
        label_map = {
            "dim. no.": "dim_no",
            "dim no.": "dim_no",
            "dim. no": "dim_no",
            "dimension description": "description",
            "dimension type": "dim_type",
            "point number (if applicable)": "point_number",
            "point number": "point_number",
            "point no.": "point_number",
            "nominal dim.": "nominal",
            "nominal": "nominal",
            "tol. max. (+)": "tol_max",
            "tol. max (+)": "tol_max",
            "tol max (+)": "tol_max",
            "tol max": "tol_max",
            "tol. min. (-)": "tol_min",
            "tol. min (-)": "tol_min",
            "tol min (-)": "tol_min",
            "tol min": "tol_min",
            "usl": "usl",
            "lsl": "lsl",
            "start point": "start_point",
            "sn": "sn",
            "process": "process",
        }
        if s in label_map:
            mapping[label_map[s]] = ri + 1  # 1-based
    return mapping


def _find_data_start(rows, label_col, data_col_start, after_row, max_search=60):
    """
    Find where measurement data rows begin by looking for:
    1. A header row containing "Start Point", "SN", or "NO" in the label column,
       or a row with multiple text headers in the pre-data columns
    2. First row after metadata with numeric values in data columns
    Returns (header_row_1based_or_None, data_start_row_1based).
    """
    # Known header indicators at the label column position
    _HEADER_KEYWORDS = {"start point", "sn", "no", "no."}

    # Strategy 1a: look for known header keywords in the label area
    _search_cols = max(label_col + 1, 20)
    for ri in range(after_row - 1, min(after_row + max_search, len(rows))):
        row = rows[ri]
        for ci in range(min(_search_cols, len(row))):
            val = row[ci].value
            if val is None:
                continue
            s = str(val).strip().lower()
            if s in _HEADER_KEYWORDS:
                return ri + 1, ri + 2  # header row, data starts next row

    # Strategy 1b: look for a row with multiple text values before data_col_start.
    # A row with 3+ non-numeric text cells in the metadata area is almost
    # certainly a header row (e.g. "Build | CFG | Color | ... | NO").
    for ri in range(after_row - 1, min(after_row + max_search, len(rows))):
        row = rows[ri]
        text_count = 0
        for ci in range(min(data_col_start, len(row))):
            val = row[ci].value
            if val is not None and isinstance(val, str) and val.strip():
                text_count += 1
        if text_count >= 3:
            return ri + 1, ri + 2

    # Strategy 2: find first row with numeric data in dimension columns
    for ri in range(after_row - 1, min(after_row + max_search, len(rows))):
        row = rows[ri]
        num_count = 0
        for ci in range(data_col_start - 1, min(data_col_start + 10, len(row))):
            val = row[ci].value
            if val is not None and _safe_num(val) is not None:
                num_count += 1
        if num_count >= 2:
            return None, ri + 1  # no header row, data starts here

    return None, None
