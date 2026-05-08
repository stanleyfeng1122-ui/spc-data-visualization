"""
Per-row metadata column extraction and post-processing.

Helpers for building the metadata column map (Build, CFG, Color, Vendor SN
etc.) from the header row, converting Shipment Date to datetime, and
detecting the factory / site code from the parsed data and filename.
"""

from __future__ import annotations

import re
from collections import OrderedDict

import pandas as pd

from .dimensions import ParsedFile, _safe_str


def build_meta_col_map(
    sheet_rows: list,
    header_row_idx: int | None,
    label_col: int,
    data_col_start: int,
    label_rows: dict,
) -> OrderedDict:
    """
    Build the ordered mapping ``column_name -> 1-based column index`` for every
    metadata column present in the sheet.

    Two strategies:

    * If a header row was identified, read every text cell from that row up
      to (but not including) the first measurement column.
    * Otherwise, fall back to the compact-format heuristics: an SN row label
      means there is a serial-number column to the left of the label column,
      and a Process row label means column A holds the process value.
    """
    meta_col_map: OrderedDict = OrderedDict()
    if header_row_idx is not None and header_row_idx - 1 < len(sheet_rows):
        hrow = sheet_rows[header_row_idx - 1]

        # Read ALL columns from the header row up to (and including) the
        # first measurement data column.  This captures every metadata
        # column regardless of its name — no hardcoded list needed.
        sn_col = None  # noqa: F841 — kept for parity with original logic
        for ci, cell in enumerate(hrow, 1):
            val = _safe_str(cell.value).strip()
            if not val:
                continue
            if ci >= data_col_start:
                break  # past the metadata columns — into measurement data
            # Skip if it looks like a dimension label (e.g. "SPC_A")
            if val.upper().startswith("SPC_") or val.upper().startswith("DIM"):
                break
            meta_col_map[val] = ci
            if val.lower() == "sn":
                sn_col = ci
    else:
        # No header row -- check if there's an SN / serial column
        # (compact format has "SN" at label_col-1, serial numbers at label_col-1)
        sn_row_label = label_rows.get("sn")
        if sn_row_label is not None:
            # The SN column is typically one col left of the label col
            sn_col_idx = label_col - 1 if label_col > 1 else 1
            meta_col_map["SN"] = sn_col_idx
        process_row_label = label_rows.get("process")
        if process_row_label is not None:
            meta_col_map["Process"] = 1  # typically col A

    return meta_col_map


def coerce_shipment_date(df: pd.DataFrame) -> None:
    """
    Convert the Shipment Date column (if present) to ``datetime``.
    Mutates the dataframe in place.
    """
    if "Shipment Date" in df.columns:
        df["Shipment Date"] = pd.to_datetime(df["Shipment Date"], errors="coerce")


def detect_factory(result: ParsedFile, filename: str) -> None:
    """
    Detect factory / site code on the given ``ParsedFile`` and assign it to
    ``result.factory``. Tries three strategies in order:

    1. The ``Vendor Serial Number`` column's mode value.
    2. A 2–4 letter prefix at the start of the first ``SN`` value.
    3. The first underscore-delimited token of the filename (e.g. "FX_K116_…").
    """
    if "Vendor Serial Number" in result.data.columns:
        vsn_vals = result.data["Vendor Serial Number"].dropna().astype(str)
        if len(vsn_vals) > 0:
            most_common = vsn_vals.mode()
            if len(most_common) > 0:
                result.factory = str(most_common.iloc[0]).strip()

    # Fallback: extract factory prefix from SN column (e.g. "FJS..." -> "FJS")
    if not result.factory and "SN" in result.data.columns:
        sn_vals = result.data["SN"].dropna().astype(str)
        if len(sn_vals) > 0:
            first_sn = sn_vals.iloc[0]
            m = re.match(r"^([A-Z]{2,4})", first_sn)
            if m:
                result.factory = m.group(1)

    # Fallback: try to extract factory from filename (e.g. "FX_K116_...")
    if not result.factory:
        name_parts = filename.split("_")
        if name_parts and re.match(r"^[A-Z]{2,4}$", name_parts[0]):
            result.factory = name_parts[0]
