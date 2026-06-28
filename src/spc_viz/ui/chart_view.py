"""Chart building + rendering orchestration.

For v1.0 (visualization-only release) the Summary Statistics expander is
intentionally NOT called from build_and_render_chart. The deferred analysis
code lives in spc_viz._deferred.analysis.render_summary_statistics — re-import
and call it after st.plotly_chart to re-enable.
"""

from __future__ import annotations

from collections import OrderedDict

import streamlit as st
from pandas import DataFrame
from plotly.graph_objects import Figure

from spc_viz.charts import (
    build_box_plot,
    build_combined_chart,
    build_histogram,
    build_range_envelope_chart,
    finalize_plotly_style,
)
from spc_viz.parsers.dimensions import DimensionMeta

from .state import ChartControls


def _build_chart_figure(
    df_clean: DataFrame,
    dim_metas: OrderedDict[str, DimensionMeta],
    selected_dim_nos: list[str],
    controls: ChartControls,
    custom_color_map: dict[str, str],
    exclude_intervals: bool,
    selected_group_label: str,
    selected_points: list[str] | None,
) -> Figure | None:
    """Build a Plotly Figure from controls without rendering it.

    Returns the finalised Figure, or None when data is insufficient.
    """
    ct = controls.chart_type
    common: dict = dict(
        df=df_clean,
        dim_metas=dim_metas,
        dim_nos=selected_dim_nos,
        color_by=controls.color_by,
        exclude_intervals=exclude_intervals,
        group_label=selected_group_label,
        row_by=controls.row_by,
        custom_color_map=custom_color_map,
        selected_points=selected_points,
    )

    fig: Figure | None
    if ct == "Combined Profile":
        fig = build_combined_chart(
            **common,
            section_by_fields=controls.section_by_fields,
            y_axis_mode=controls.y_axis_mode,
            custom_yrange=controls.custom_yrange,
            show_average_line=controls.show_average_line,
        )
    elif ct == "Box Plot":
        fig = build_box_plot(
            **common,
            section_by_fields=controls.section_by_fields,
            y_axis_mode=controls.y_axis_mode,
            custom_yrange=controls.custom_yrange,
        )
    elif ct == "Histogram":
        fig = build_histogram(
            **common,
            nbins=controls.hist_nbins,
        )
    elif ct == "Range Envelope":
        fig = build_range_envelope_chart(
            **common,
            section_by_fields=controls.section_by_fields,
            y_axis_mode=controls.y_axis_mode,
            custom_yrange=controls.custom_yrange,
        )
    else:
        fig = None

    if fig is not None:
        finalize_plotly_style(fig)
    return fig


def build_and_render_chart(
    df_clean: DataFrame,
    dim_metas: OrderedDict[str, DimensionMeta],
    selected_dim_nos: list[str],
    controls: ChartControls,
    custom_color_map: dict[str, str],
    exclude_intervals: bool,
    selected_group_label: str,
    selected_points: list[str] | None,
    key_prefix: str = "",
) -> Figure:
    """Build the Plotly figure based on ChartControls and render it."""
    fig = _build_chart_figure(
        df_clean,
        dim_metas,
        selected_dim_nos,
        controls,
        custom_color_map,
        exclude_intervals,
        selected_group_label,
        selected_points,
    )

    if fig is None:
        st.warning("Could not generate chart. Check dimensions have data.")
        st.stop()

    st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}main_chart")
    return fig
