"""
SPC charts package.

Re-exports the public chart-building, styling, and SPC analytics API
from the focused submodules so existing callers can keep importing from
``spc_viz.charts`` without knowing the internal layout.
"""

from __future__ import annotations

from .base import (
    COLOR_PALETTE,
    MAX_TRACES_PER_GROUP,
    calc_process_capability,
    compute_row_groups,
    compute_sections,
    cusum_analysis,
    get_color_for_group,
    nelson_rules,
    prepare_combined_data,
)
from .box_plot import build_box_plot
from .combined_profile import build_combined_chart
from .histogram import build_histogram
from .range_envelope import build_range_envelope_chart
from .styling import finalize_plotly_style

__all__ = [
    "COLOR_PALETTE",
    "MAX_TRACES_PER_GROUP",
    "build_box_plot",
    "build_combined_chart",
    "build_histogram",
    "build_range_envelope_chart",
    "calc_process_capability",
    "compute_row_groups",
    "compute_sections",
    "cusum_analysis",
    "finalize_plotly_style",
    "get_color_for_group",
    "nelson_rules",
    "prepare_combined_data",
]
