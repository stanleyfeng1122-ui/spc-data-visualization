"""Streamlit UI widgets for the SPC visualization app.

Each submodule takes a *key_prefix* so multiple pages can coexist in the
same Streamlit session without widget-key collisions.

For v1.0 the Summary Statistics analysis features (CPK / ANOVA / Nelson /
CUSUM / EWMA) are intentionally not exposed here — they live in
spc_viz._deferred.analysis until a future release re-wires them.
"""

from .batch_export import render_batch_export
from .chart_view import _build_chart_figure, build_and_render_chart
from .dimension_picker import build_dimension_selector, build_point_filter
from .sidebar import (
    SECTION_FIELDS,
    build_chart_controls,
    build_color_pickers,
)
from .state import ChartControls, prepare_and_clean

__all__ = [
    "ChartControls",
    "SECTION_FIELDS",
    "_build_chart_figure",
    "build_and_render_chart",
    "build_chart_controls",
    "build_color_pickers",
    "build_dimension_selector",
    "build_point_filter",
    "prepare_and_clean",
    "render_batch_export",
]
