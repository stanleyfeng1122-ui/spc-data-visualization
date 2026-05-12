"""
Measurement value extraction.

Helpers extracted from the per-sheet parser:

* :func:`_is_non_data_sheet` — recognise pivot / chart sheets to skip.
* :func:`extract_dimension_columns` — read per-column dimension metadata
  (dim_no, description, type, point, nominal, tol_max, tol_min, usl, lsl)
  from the header rows.
* :func:`merge_dimension_groups` — group ``SPC_HG``, ``SPC_HG.01``,
  ``SPC_HG.02`` etc. into a single dimension and synthesise point labels
  where they are missing.
* :func:`extract_records` — iterate the data rows and build the list of
  per-row dicts used to construct the dataframe.
"""

from __future__ import annotations

import re
from collections import OrderedDict

from .dimensions import DimensionMeta, _safe_num, _safe_str

# ---------------------------------------------------------------------------
# Sheet name patterns for non-data sheets (to skip during auto-detection)
# ---------------------------------------------------------------------------

_NON_DATA_SHEET_PATTERNS = [
    re.compile(r"^BoxPlotCht", re.IGNORECASE),
    re.compile(r"^Histo\s+Pivot$", re.IGNORECASE),
    re.compile(r"^Histo\s+Listbox$", re.IGNORECASE),
    re.compile(r"^Histo\s+Curve$", re.IGNORECASE),
]


def _is_non_data_sheet(name: str) -> bool:
    """Return True if the sheet name matches a known non-data pattern."""
    for pat in _NON_DATA_SHEET_PATTERNS:
        if pat.search(name):
            return True
    return False


# ---------------------------------------------------------------------------
# Per-column dimension metadata extraction
# ---------------------------------------------------------------------------


def extract_dimension_columns(
    sheet_rows: list,
    dim_no_row: int,
    data_col_start: int,
    label_rows: dict,
) -> tuple[dict, dict, dict, dict, dict, dict, dict, dict, dict]:
    """
    Walk the dimension columns and read per-column metadata.

    Returns nine dicts keyed by 1-based column index:
    ``(col_dim_no, col_desc, col_type, col_point, col_nominal,
       col_tol_max, col_tol_min, col_usl, col_lsl)``.
    """
    desc_row = label_rows.get("description")
    type_row = label_rows.get("dim_type")
    point_row = label_rows.get("point_number")
    nominal_row = label_rows.get("nominal")
    tol_max_row = label_rows.get("tol_max")
    tol_min_row = label_rows.get("tol_min")
    usl_row = label_rows.get("usl")
    lsl_row = label_rows.get("lsl")

    max_col = len(sheet_rows[dim_no_row - 1]) if dim_no_row - 1 < len(sheet_rows) else 0

    def _cell(r, c):
        """Get cell value; r and c are 1-based."""
        if r is None:
            return None
        if r - 1 < len(sheet_rows):
            row = sheet_rows[r - 1]
            if c - 1 < len(row):
                return row[c - 1].value
        return None

    col_dim_no: dict = {}
    col_desc: dict = {}
    col_type: dict = {}
    col_point: dict = {}
    col_nominal: dict = {}
    col_tol_max: dict = {}
    col_tol_min: dict = {}
    col_usl: dict = {}
    col_lsl: dict = {}

    for ci in range(data_col_start, max_col + 1):  # 1-based
        raw_val = _safe_str(_cell(dim_no_row, ci))
        if not raw_val:
            continue
        # Some files embed description in the dim_no cell (e.g. "SPC_G\nCombo Flex flatness")
        if "\n" in raw_val:
            parts = raw_val.split("\n", 1)
            dim_no = parts[0].strip()
            embedded_desc = parts[1].strip()
        else:
            dim_no = raw_val
            embedded_desc = ""
        col_dim_no[ci] = dim_no
        col_desc[ci] = _safe_str(_cell(desc_row, ci)) if desc_row else embedded_desc
        col_type[ci] = _safe_str(_cell(type_row, ci)) if type_row else ""
        col_point[ci] = _safe_str(_cell(point_row, ci)) if point_row else ""
        col_nominal[ci] = _safe_num(_cell(nominal_row, ci)) if nominal_row else None
        col_tol_max[ci] = _safe_num(_cell(tol_max_row, ci)) if tol_max_row else None
        col_tol_min[ci] = _safe_num(_cell(tol_min_row, ci)) if tol_min_row else None
        col_usl[ci] = _safe_num(_cell(usl_row, ci)) if usl_row else None
        col_lsl[ci] = _safe_num(_cell(lsl_row, ci)) if lsl_row else None

    return (
        col_dim_no,
        col_desc,
        col_type,
        col_point,
        col_nominal,
        col_tol_max,
        col_tol_min,
        col_usl,
        col_lsl,
    )


# ---------------------------------------------------------------------------
# Dimension group merging + point-label deduplication
# ---------------------------------------------------------------------------


def merge_dimension_groups(
    col_dim_no: dict,
    col_desc: dict,
    col_type: dict,
    col_point: dict,
    col_nominal: dict,
    col_tol_max: dict,
    col_tol_min: dict,
    col_usl: dict,
    col_lsl: dict,
) -> OrderedDict:
    """
    Merge numbered sub-dimensions (compact format) into a single group and
    synthesise point labels where they are missing.

    For example::

        SPC_HG, SPC_HG.01, SPC_HG.02 ... -> single "SPC_HG" group

    Only applies when individual dims are single-column with no point
    numbers (i.e. the compact format pattern). Returns an
    :class:`OrderedDict` of ``dim_no -> DimensionMeta``.
    """
    # Group by dim_no preserving order
    dim_groups: OrderedDict = OrderedDict()
    for ci, dno in col_dim_no.items():
        dim_groups.setdefault(dno, []).append(ci)

    merged_groups: OrderedDict = OrderedDict()  # parent_name -> list of col indices
    merged_descs: dict = {}  # parent_name -> description
    consumed: set = set()  # dim_nos already merged

    # First pass: find explicit parents with .NNN children
    for dno in list(dim_groups.keys()):
        if dno in consumed:
            continue
        children = []
        for other in dim_groups:
            if other == dno:
                continue
            if re.match(re.escape(dno) + r"\.\d+$", other):
                children.append(other)
        if children:
            all_cols = list(dim_groups[dno])
            for child in children:
                all_cols.extend(dim_groups[child])
                consumed.add(child)
            merged_groups[dno] = all_cols
            merged_descs[dno] = col_desc.get(dim_groups[dno][0], "")
            consumed.add(dno)

    # Second pass: group orphan .NNN siblings with no parent
    # e.g. SPC_1.001, SPC_1.002, ... (no bare SPC_1 exists)
    orphans: OrderedDict = OrderedDict()  # prefix -> list of (dno, cols)
    for dno in list(dim_groups.keys()):
        if dno in consumed:
            continue
        m = re.match(r"^(.+)\.\d+$", dno)
        if m:
            prefix = m.group(1)
            orphans.setdefault(prefix, []).append(dno)
        else:
            # Not a numbered dim, keep standalone
            merged_groups[dno] = list(dim_groups[dno])
            consumed.add(dno)

    for prefix, siblings in orphans.items():
        if len(siblings) >= 2:
            # Merge all siblings under the prefix name
            all_cols = []
            for sib in siblings:
                all_cols.extend(dim_groups[sib])
                consumed.add(sib)
            merged_groups[prefix] = all_cols
            merged_descs[prefix] = col_desc.get(dim_groups[siblings[0]][0], "")
        else:
            # Single orphan, keep as-is
            dno = siblings[0]
            merged_groups[dno] = list(dim_groups[dno])
            consumed.add(dno)

    dimensions: OrderedDict = OrderedDict()
    for dno, cols in merged_groups.items():
        desc = col_desc.get(cols[0], "") if dno not in merged_descs else merged_descs[dno]
        dtype = col_type.get(cols[0], "")

        col_labels = []
        point_numbers = []
        for idx, ci in enumerate(cols):
            pt = col_point.get(ci, "")
            if pt:
                point_numbers.append(pt)
                col_labels.append(f"{dno}_{pt}")
            else:
                # Synthesize point label: P0, P1, P2, ...
                syn_pt = f"P{idx}"
                point_numbers.append(syn_pt)
                col_labels.append(f"{dno}_{syn_pt}")

        dimensions[dno] = DimensionMeta(
            dim_no=dno,
            description=desc,
            dim_type=dtype,
            point_numbers=point_numbers,
            nominal=[col_nominal.get(ci) for ci in cols],
            tol_max=[col_tol_max.get(ci) for ci in cols],
            tol_min=[col_tol_min.get(ci) for ci in cols],
            usl=[col_usl.get(ci) for ci in cols],
            lsl=[col_lsl.get(ci) for ci in cols],
            col_indices=cols,
            col_labels=col_labels,
        )

    return dimensions


# ---------------------------------------------------------------------------
# Per-row measurement extraction
# ---------------------------------------------------------------------------


def extract_records(
    sheet_rows: list,
    data_start_row: int,
    meta_col_map: OrderedDict,
    dimensions: OrderedDict,
) -> list[dict]:
    """
    Iterate the measurement rows and return a list of record dicts ready
    for ``pd.DataFrame``.

    Empty rows are skipped using a sentinel column heuristic: prefer
    "Start Point" or "SN" if present, otherwise require numeric data in
    at least one of the first three columns of any dimension.
    """
    # Determine which column to use for the "is row populated?" check
    # Prefer "Start Point" or "SN", fallback to first dimension column
    check_col = None
    if "Start Point" in meta_col_map:
        check_col = meta_col_map["Start Point"]
    elif "SN" in meta_col_map:
        check_col = meta_col_map["SN"]

    records: list[dict] = []
    for ri in range(data_start_row - 1, len(sheet_rows)):
        row = sheet_rows[ri]
        rec: dict = {}

        # Metadata columns
        for name, ci in meta_col_map.items():
            if ci - 1 < len(row):
                rec[name] = row[ci - 1].value
            else:
                rec[name] = None

        # Skip empty rows: check sentinel column or look for numeric data
        if check_col is not None:
            if check_col - 1 < len(row):
                sentinel = row[check_col - 1].value
            else:
                sentinel = None
            if sentinel is None:
                continue
        else:
            # No sentinel: check if row has any numeric data in dim columns
            has_data = False
            for dno, dmeta in dimensions.items():
                for ci_d in dmeta.col_indices[:3]:
                    if ci_d - 1 < len(row) and _safe_num(row[ci_d - 1].value) is not None:
                        has_data = True
                        break
                if has_data:
                    break
            if not has_data:
                continue

        # Measurement columns
        for dno, dmeta in dimensions.items():
            for ci, label in zip(dmeta.col_indices, dmeta.col_labels):
                if ci - 1 < len(row):
                    rec[label] = row[ci - 1].value
                else:
                    rec[label] = None

        records.append(rec)

    return records
