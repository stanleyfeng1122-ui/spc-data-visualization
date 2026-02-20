"""
SPC Data Visualization Tool - Excel Parser Module

Parses Apple CPK Data Sheet Excel files (.xlsx) with dynamic column detection.
Handles varying column layouts between different vendor files.
Supports both "Data Input" (summary values) and "Raw data" (full profiles) sheets.
"""

import openpyxl
import pandas as pd
import re
from dataclasses import dataclass, field
from collections import OrderedDict
from typing import Optional, Dict, List, Tuple
import io


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class DimensionMeta:
    """Metadata for a single dimension group (e.g. SPC_AA)."""
    dim_no: str                     # e.g. "SPC_AA"
    description: str                # e.g. "landing to E surface height"
    dim_type: str                   # e.g. "Non-Profile Measurement"
    point_numbers: list             # list of point labels per sub-column
    nominal: list                   # nominal value per sub-column
    tol_max: list                   # tolerance max (+) per sub-column
    tol_min: list                   # tolerance min (-) per sub-column
    usl: list                       # upper spec limit per sub-column
    lsl: list                       # lower spec limit per sub-column
    col_indices: list               # 1-based column indices in the sheet
    col_labels: list                # readable column labels for the dataframe


@dataclass
class ParsedFile:
    """Result of parsing a single Excel file."""
    filename: str
    sheet_name: str
    part_number: Optional[str] = None
    part_description: Optional[str] = None
    revision: Optional[str] = None
    factory: Optional[str] = None   # factory/site code (e.g. "FX", "TY")
    dimensions: OrderedDict = field(default_factory=OrderedDict)  # dim_no -> DimensionMeta
    data: Optional[pd.DataFrame] = None  # measurement rows
    meta_columns: list = field(default_factory=list)  # names of metadata columns present


# ---------------------------------------------------------------------------
# Known metadata header names (normalised to lowercase for matching)
# ---------------------------------------------------------------------------

KNOWN_META_HEADERS = {
    "build", "shipment date", "color", "config",
    "vendor serial number", "fabric thickness",
    "2d barcode", "1d barcode", "rm coil", "raw material",
    "start point",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_str(val) -> str:
    """Convert a cell value to a stripped string, or empty string if None."""
    if val is None:
        return ""
    return str(val).strip()


def _safe_num(val):
    """Return a float if numeric, else None."""
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _is_interval_point(point_label: str) -> bool:
    """Check if a point label is an interval (e.g. 'C11-C12') vs actual (e.g. 'C11')."""
    return bool(re.search(r'C\d+-C\d+', str(point_label)))


# ---------------------------------------------------------------------------
# Core parser
# ---------------------------------------------------------------------------

def parse_excel(file_or_path, sheet_name: str = "Raw data") -> ParsedFile:
    """
    Parse a vendor CPK Excel file and return structured data.

    Parameters
    ----------
    file_or_path : str or file-like
        Path to an .xlsx file, or an in-memory file object (e.g. from
        Streamlit's file_uploader).
    sheet_name : str
        Which sheet to read: "Data Input" or "Raw data".

    Returns
    -------
    ParsedFile
    """
    # Support both file paths and BytesIO objects from Streamlit
    if isinstance(file_or_path, (str,)):
        wb = openpyxl.load_workbook(file_or_path, data_only=True, read_only=True)
        filename = file_or_path.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    else:
        # BytesIO / UploadedFile
        file_or_path.seek(0)
        wb = openpyxl.load_workbook(file_or_path, data_only=True, read_only=True)
        filename = getattr(file_or_path, "name", "uploaded_file")

    if sheet_name not in wb.sheetnames:
        available = ", ".join(wb.sheetnames)
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available sheets: {available}"
        )

    ws = wb[sheet_name]

    # We need random-access rows, so materialise the sheet into a list.
    # read_only worksheets yield rows lazily; convert to list of tuples.
    rows = list(ws.rows)  # list of tuples of Cell objects

    result = ParsedFile(filename=filename, sheet_name=sheet_name)

    # ------------------------------------------------------------------
    # 1. File-level metadata (rows 1-3, column L=12, M=13)
    # ------------------------------------------------------------------
    def _cell(r, c):
        """Get cell value; r and c are 1-based."""
        if r - 1 < len(rows):
            row = rows[r - 1]
            if c - 1 < len(row):
                return row[c - 1].value
        return None

    result.part_number = _safe_str(_cell(2, 13))      # Row 2, Col M
    result.revision = _safe_str(_cell(2, 16))          # Row 2, Col P
    result.part_description = _safe_str(_cell(3, 13))  # Row 3, Col M

    # ------------------------------------------------------------------
    # 2. Dimension metadata (rows 6-20, columns 13+)
    # ------------------------------------------------------------------
    max_col = len(rows[5]) if len(rows) >= 6 else 0  # row 6 (index 5)

    # Collect per-column info first, then group by dim_no
    col_dim_no = {}      # col_idx -> dim_no string
    col_desc = {}        # col_idx -> description
    col_type = {}        # col_idx -> dimension type
    col_point = {}       # col_idx -> point number
    col_nominal = {}     # col_idx -> nominal
    col_tol_max = {}     # col_idx -> tol max
    col_tol_min = {}     # col_idx -> tol min
    col_usl = {}         # col_idx -> USL
    col_lsl = {}         # col_idx -> LSL

    for ci in range(13, max_col + 1):  # 1-based col 13 onward
        dim_no = _safe_str(_cell(6, ci))
        if not dim_no:
            continue
        col_dim_no[ci] = dim_no
        col_desc[ci] = _safe_str(_cell(7, ci))
        col_type[ci] = _safe_str(_cell(8, ci))
        col_point[ci] = _safe_str(_cell(9, ci))
        col_nominal[ci] = _safe_num(_cell(14, ci))
        col_tol_max[ci] = _safe_num(_cell(15, ci))
        col_tol_min[ci] = _safe_num(_cell(16, ci))
        col_usl[ci] = _safe_num(_cell(19, ci))
        col_lsl[ci] = _safe_num(_cell(20, ci))

    # Group by dim_no preserving order
    dim_groups = OrderedDict()  # dim_no -> list of col indices
    for ci, dno in col_dim_no.items():
        dim_groups.setdefault(dno, []).append(ci)

    for dno, cols in dim_groups.items():
        desc = col_desc.get(cols[0], "")
        dtype = col_type.get(cols[0], "")

        # Build readable column labels: "SPC_AA_C39", "SPC_AA_C43", etc.
        col_labels = []
        for ci in cols:
            pt = col_point.get(ci, "")
            if pt:
                col_labels.append(f"{dno}_{pt}")
            else:
                col_labels.append(f"{dno}_col{ci}")

        result.dimensions[dno] = DimensionMeta(
            dim_no=dno,
            description=desc,
            dim_type=dtype,
            point_numbers=[col_point.get(ci, "") for ci in cols],
            nominal=[col_nominal.get(ci) for ci in cols],
            tol_max=[col_tol_max.get(ci) for ci in cols],
            tol_min=[col_tol_min.get(ci) for ci in cols],
            usl=[col_usl.get(ci) for ci in cols],
            lsl=[col_lsl.get(ci) for ci in cols],
            col_indices=cols,
            col_labels=col_labels,
        )

    # ------------------------------------------------------------------
    # 3. Data header row (row 40) -- dynamic column mapping
    # ------------------------------------------------------------------
    header_row_idx = 40  # 1-based

    # Build mapping: header_name -> col_index (1-based)
    header_map = {}  # normalised_name -> col_index
    header_raw = {}  # col_index -> raw_name
    if header_row_idx - 1 < len(rows):
        hrow = rows[header_row_idx - 1]
        for ci, cell in enumerate(hrow, 1):
            val = _safe_str(cell.value).lower()
            if val:
                header_map[val] = ci
                header_raw[ci] = _safe_str(cell.value)

    # Determine which metadata columns exist
    meta_col_map = OrderedDict()  # display_name -> col_index
    for name in ["Build", "Shipment Date", "Color", "Config",
                  "Vendor Serial Number", "Fabric thickness",
                  "2D Barcode", "1D Barcode", "RM Coil", "Raw material",
                  "Start Point"]:
        # Try exact match first, then case-insensitive
        ci = header_map.get(name.lower())
        if ci is not None:
            meta_col_map[name] = ci

    result.meta_columns = list(meta_col_map.keys())

    # ------------------------------------------------------------------
    # 4. Measurement data (rows 41+)
    # ------------------------------------------------------------------
    data_start_row = 41  # 1-based

    records = []
    for ri in range(data_start_row - 1, len(rows)):
        row = rows[ri]

        # Build record dict
        rec = {}

        # Metadata columns
        for name, ci in meta_col_map.items():
            if ci - 1 < len(row):
                rec[name] = row[ci - 1].value
            else:
                rec[name] = None

        # Skip rows that look empty (no Start Point value)
        sp = rec.get("Start Point")
        if sp is None:
            continue

        # Measurement columns (all dimension columns)
        for dno, dmeta in result.dimensions.items():
            for ci, label in zip(dmeta.col_indices, dmeta.col_labels):
                if ci - 1 < len(row):
                    rec[label] = row[ci - 1].value
                else:
                    rec[label] = None

        records.append(rec)

    result.data = pd.DataFrame(records)

    # Convert Shipment Date to datetime if present
    if "Shipment Date" in result.data.columns:
        result.data["Shipment Date"] = pd.to_datetime(
            result.data["Shipment Date"], errors="coerce"
        )

    # ------------------------------------------------------------------
    # 5. Detect factory / site code from Vendor Serial Number column
    # ------------------------------------------------------------------
    if "Vendor Serial Number" in result.data.columns:
        vsn_vals = result.data["Vendor Serial Number"].dropna().astype(str)
        if len(vsn_vals) > 0:
            # The VSN column typically contains a short factory code
            # (e.g. "FX", "TY") repeated for all rows in a file
            most_common = vsn_vals.mode()
            if len(most_common) > 0:
                result.factory = str(most_common.iloc[0]).strip()

    # Fallback: try to extract factory from filename (e.g. "FX_K116_...")
    if not result.factory:
        name_parts = filename.split("_")
        if name_parts:
            result.factory = name_parts[0]

    wb.close()
    return result


# ---------------------------------------------------------------------------
# Dimension grouping by description keywords
# ---------------------------------------------------------------------------

def detect_dimension_groups(dimensions: OrderedDict) -> Dict[str, List[str]]:
    """
    Auto-detect dimension groups by analysing description keywords.

    Returns a dict of group_display_name -> list of dim_no strings.
    Groups dimensions that share a common keyword in their description.
    Dimensions without a matching keyword are placed in individual groups.
    Also includes an "All dimensions" pseudo-group.
    """
    # Build keyword -> list of dim_nos mapping
    keyword_map: Dict[str, List[Tuple[str, str]]] = {}  # keyword -> [(dim_no, description)]
    ungrouped: List[Tuple[str, str]] = []  # dims with no keyword match

    for dno, dmeta in dimensions.items():
        desc = dmeta.description.lower().strip()
        if not desc:
            ungrouped.append((dno, ""))
            continue

        keyword = _extract_group_keyword(desc)
        if keyword:
            keyword_map.setdefault(keyword, []).append((dno, dmeta.description))
        else:
            ungrouped.append((dno, dmeta.description))

    groups: Dict[str, List[str]] = OrderedDict()

    # Keyword-matched groups (2+ dimensions sharing a keyword)
    for keyword, dim_list in keyword_map.items():
        if len(dim_list) >= 2:
            dim_nos = [d[0] for d in dim_list]
            dim_labels = " / ".join(dim_nos)
            display_keyword = keyword.replace("_", " ").title()
            group_label = f"{display_keyword}: {dim_labels}"
            groups[group_label] = dim_nos
        else:
            # Single-member keyword group -> treat as individual
            ungrouped.extend(dim_list)

    # Individual dimension entries
    for dno, desc in ungrouped:
        label = f"{dno} - {desc}" if desc else dno
        groups[label] = [dno]

    # "All dimensions" pseudo-group
    if len(dimensions) > 1:
        all_dim_nos = list(dimensions.keys())
        groups["All dimensions"] = all_dim_nos

    return groups


def _extract_group_keyword(description: str) -> str:
    """
    Extract a grouping keyword from a dimension description.

    Examples:
        "z straightness of left"       -> "z_straightness"
        "z straightness of front"      -> "z_straightness"
        "flatness of Datum A"          -> "flatness"
        "overall length(outer edge)"   -> "overall_length"
        "half length"                  -> "half_length"
        "half width"                   -> "half_width"
        "landing to E surface height"  -> "landing_height"
    """
    desc = description.lower().strip()

    # Try matching known patterns (most specific first)
    patterns = [
        (r"z\s*straightness", "z_straightness"),
        (r"flatness", "flatness"),
        (r"overall\s*length", "overall_length"),
        (r"half\s*length", "half_length"),
        (r"half\s*width", "half_width"),
        (r"landing.*height", "landing_height"),
        (r"height", "height"),
        (r"thickness", "thickness"),
        (r"gap", "gap"),
        (r"offset", "offset"),
        (r"profile", "profile"),
        (r"straightness", "straightness"),
    ]

    for pattern, keyword in patterns:
        if re.search(pattern, desc):
            return keyword

    return ""


def get_filtered_dim_meta(
    dmeta: DimensionMeta,
    exclude_intervals: bool = True,
) -> Tuple[List[str], List[str], List, List, List]:
    """
    Return filtered lists of (col_labels, point_numbers, nominal, usl, lsl)
    optionally excluding interval points (e.g. "C11-C12").

    Parameters
    ----------
    dmeta : DimensionMeta
    exclude_intervals : bool
        If True, exclude interval-type points like "C11-C12".

    Returns
    -------
    (col_labels, point_numbers, nominal, usl, lsl) -- filtered lists
    """
    col_labels = []
    point_numbers = []
    nominal = []
    usl = []
    lsl = []

    for i, pt in enumerate(dmeta.point_numbers):
        if exclude_intervals and _is_interval_point(pt):
            continue
        col_labels.append(dmeta.col_labels[i])
        point_numbers.append(pt)
        nominal.append(dmeta.nominal[i])
        usl.append(dmeta.usl[i])
        lsl.append(dmeta.lsl[i])

    return col_labels, point_numbers, nominal, usl, lsl


# ---------------------------------------------------------------------------
# Convenience helpers for the app layer
# ---------------------------------------------------------------------------

def get_dimension_options(parsed: ParsedFile) -> list:
    """
    Return a list of (display_label, dim_no) tuples for the dimension selector.
    """
    options = []
    for dno, dmeta in parsed.dimensions.items():
        label = f"{dno} - {dmeta.description}" if dmeta.description else dno
        options.append((label, dno))
    return options


def get_groupable_columns(parsed: ParsedFile) -> list:
    """
    Return the list of metadata column names that can be used for
    X-axis grouping or color-by.
    """
    usable = []
    for name in parsed.meta_columns:
        if name == "Start Point":
            continue
        # Only include columns that actually have data
        if name in parsed.data.columns and parsed.data[name].notna().any():
            usable.append(name)
    return usable
