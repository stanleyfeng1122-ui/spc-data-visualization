"""
SPC parser package.

Re-exports the public parser API from the focused submodules and triggers
the openpyxl monkey-patch at import time.
"""

from __future__ import annotations

# Activate openpyxl ExternalReference patch as a side effect of importing
# the package (parity with the original spc_parser module behaviour).
from . import openpyxl_patch  # noqa: F401
from .dimensions import (
    DimensionMeta,
    ParsedFile,
    _extract_group_keyword,
    _is_interval_point,
    _safe_num,
    _safe_str,
    detect_dimension_groups,
    get_dimension_options,
    get_filtered_dim_meta,
    get_groupable_columns,
)
from .excel_reader import (
    _open_workbook,
    _parse_single_sheet,
    parse_excel,
    parse_excel_multi,
)
from .header_detect import (
    KNOWN_META_HEADERS,
    _find_data_start,
    _find_dim_no_cell,
    _scan_label_rows,
)
from .measurements import (
    _is_non_data_sheet,
    extract_dimension_columns,
    extract_records,
    merge_dimension_groups,
)
from .metadata import (
    build_meta_col_map,
    coerce_shipment_date,
    detect_factory,
)
from .pairing import (
    PAIR_DIM_PREFIX,
    build_paired_dimension_map,
    is_paired_dim_id,
    simple_feature_name,
)

__all__ = [
    # Public API
    "DimensionMeta",
    "ParsedFile",
    "parse_excel",
    "parse_excel_multi",
    "PAIR_DIM_PREFIX",
    "build_paired_dimension_map",
    "detect_dimension_groups",
    "get_dimension_options",
    "get_filtered_dim_meta",
    "get_groupable_columns",
    # Lower-level helpers retained for parity with the legacy module
    "KNOWN_META_HEADERS",
    "_find_data_start",
    "_find_dim_no_cell",
    "_is_interval_point",
    "_is_non_data_sheet",
    "_open_workbook",
    "_parse_single_sheet",
    "_safe_num",
    "_safe_str",
    "_scan_label_rows",
    "_extract_group_keyword",
    "build_meta_col_map",
    "coerce_shipment_date",
    "detect_factory",
    "is_paired_dim_id",
    "simple_feature_name",
]
