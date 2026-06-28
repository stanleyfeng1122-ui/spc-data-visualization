"""
Dimension data classes and helpers.

Contains the :class:`DimensionMeta` and :class:`ParsedFile` dataclasses, the
primitive cell-coercion helpers (``_safe_str`` / ``_safe_num``), and the
public dimension-grouping API (``detect_dimension_groups``,
``get_filtered_dim_meta``, ``get_dimension_options``,
``get_groupable_columns``).
"""

from __future__ import annotations

import re
from collections import OrderedDict
from dataclasses import dataclass, field

import pandas as pd

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class DimensionMeta:
    """Metadata for a single dimension group (e.g. SPC_AA)."""

    dim_no: str  # e.g. "SPC_AA"
    description: str  # e.g. "landing to E surface height"
    dim_type: str  # e.g. "Non-Profile Measurement"
    point_numbers: list  # list of point labels per sub-column
    nominal: list  # nominal value per sub-column
    tol_max: list  # tolerance max (+) per sub-column
    tol_min: list  # tolerance min (-) per sub-column
    usl: list  # upper spec limit per sub-column
    lsl: list  # lower spec limit per sub-column
    col_indices: list  # 1-based column indices in the sheet
    col_labels: list  # readable column labels for the dataframe
    source_dim_nos: list | None = None  # for paired dims: the source SPC bubble ids


@dataclass
class ParsedFile:
    """Result of parsing a single Excel file."""

    filename: str
    sheet_name: str
    part_number: str | None = None
    part_description: str | None = None
    revision: str | None = None
    factory: str | None = None  # factory/site code (e.g. "FX", "TY")
    dimensions: OrderedDict = field(default_factory=OrderedDict)  # dim_no -> DimensionMeta
    data: pd.DataFrame | None = None  # measurement rows
    meta_columns: list = field(default_factory=list)  # names of metadata columns present


# ---------------------------------------------------------------------------
# Primitive cell helpers
# ---------------------------------------------------------------------------


def _safe_str(val: object) -> str:
    """Convert a cell value to a stripped string, or empty string if None."""
    if val is None:
        return ""
    return str(val).strip()


def _safe_num(val: object) -> float | None:
    """Return a float if numeric, else None."""
    if val is None:
        return None
    try:
        return float(val)  # type: ignore[arg-type]
    except (ValueError, TypeError):
        return None


def _is_interval_point(point_label: str) -> bool:
    """Check if a point label is an interval (e.g. 'C11-C12') vs actual (e.g. 'C11')."""
    return bool(re.search(r"C\d+-C\d+", str(point_label)))


# ---------------------------------------------------------------------------
# Dimension grouping by description keywords
# ---------------------------------------------------------------------------


def detect_dimension_groups(dimensions: OrderedDict) -> dict[str, list[str]]:
    """Auto-detect dimension groups by analysing description keywords.

    Returns a dict of group_display_name -> list of dim_no strings.
    Groups dimensions that share a common keyword in their description.
    Dimensions without a matching keyword are placed in individual groups.
    Also includes an "All dimensions" pseudo-group.
    """
    # Build keyword -> list of dim_nos mapping
    keyword_map: dict[str, list[tuple[str, str]]] = {}  # keyword -> [(dim_no, description)]
    ungrouped: list[tuple[str, str]] = []  # dims with no keyword match

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

    groups: dict[str, list[str]] = OrderedDict()

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
    """Extract a grouping keyword from a dimension description.

    Examples::

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
    patterns: list[tuple[str, str]] = [
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
) -> tuple[list[str], list[str], list, list, list]:
    """Return filtered lists of (col_labels, point_numbers, nominal, usl, lsl).

    Optionally excludes interval points (e.g. "C11-C12").

    Parameters
    ----------
    dmeta : DimensionMeta
    exclude_intervals : bool
        If True, exclude interval-type points like "C11-C12".

    Returns
    -------
    (col_labels, point_numbers, nominal, usl, lsl) -- filtered lists
    """
    col_labels: list[str] = []
    point_numbers: list[str] = []
    nominal: list = []
    usl: list = []
    lsl: list = []

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


def get_dimension_options(parsed: ParsedFile) -> list[tuple[str, str]]:
    """Return a list of (display_label, dim_no) tuples for the dimension selector."""
    options: list[tuple[str, str]] = []
    for dno, dmeta in parsed.dimensions.items():
        label = f"{dno} - {dmeta.description}" if dmeta.description else dno
        options.append((label, dno))
    return options


def get_groupable_columns(parsed: ParsedFile) -> list[str]:
    """Return the list of metadata column names that can be used for X-axis grouping or color-by."""
    usable: list[str] = []
    if parsed.data is None:
        return usable
    for name in parsed.meta_columns:
        if name == "Start Point":
            continue
        # Only include columns that actually have data
        if name in parsed.data.columns and parsed.data[name].notna().any():
            usable.append(name)
    return usable
