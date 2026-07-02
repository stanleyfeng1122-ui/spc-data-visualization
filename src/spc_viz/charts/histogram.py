"""Histogram chart builder.

Builds the frequency-distribution histogram with optional row/column
faceting and spec-limit overlays. Pure code movement from the original
``chart_utils`` module.
"""

from __future__ import annotations

from collections import OrderedDict

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.graph_objects import Figure
from plotly.subplots import make_subplots

from spc_viz.parsers import get_filtered_dim_meta
from spc_viz.parsers.dimensions import DimensionMeta

from .base import compute_row_groups, get_color_for_group
from .spec_limits import HISTOGRAM_STYLE, SpecSpan, render_spec_limits

# ---------------------------------------------------------------------------
# Chart building -- histogram
# ---------------------------------------------------------------------------


def build_histogram(
    df: pd.DataFrame,
    dim_metas: OrderedDict[str, DimensionMeta],
    dim_nos: list[str],
    color_by: str,
    exclude_intervals: bool,
    group_label: str,
    nbins: int = 40,
    row_by: str = "None",
    custom_color_map: dict[str, str] | None = None,
    selected_points: list[str] | None = None,
) -> Figure | None:
    """Build a histogram showing frequency distribution of measurement values."""
    # Normalise selected_points to a set for O(1) lookup; None/empty means show all
    _point_filter: set[str] | None = set(selected_points) if selected_points else None

    valid_dim_nos = [d for d in dim_nos if d in dim_metas]
    n_dim_cols = len(valid_dim_nos)
    if n_dim_cols == 0:
        return None

    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))
    n_facet_rows = len(unique_rows)

    if color_by != "None" and color_by in df.columns:
        color_series = df[color_by].fillna("Unknown").astype(str)
        unique_colors = sorted(color_series.unique())
    else:
        color_series = pd.Series("All", index=df.index)
        unique_colors = ["All"]

    if custom_color_map:
        color_map = {
            g: custom_color_map.get(g, get_color_for_group(i)) for i, g in enumerate(unique_colors)
        }
    else:
        color_map = {g: get_color_for_group(i) for i, g in enumerate(unique_colors)}

    n_cols = max(n_dim_cols, 1)
    n_rows = max(n_facet_rows, 1)
    use_subplots = n_cols > 1 or n_rows > 1

    fig: Figure
    if use_subplots:
        subplot_titles: list[str] = []
        for r_label in unique_rows:
            for dno in valid_dim_nos:
                if n_rows > 1 and n_cols > 1:
                    subplot_titles.append(f"{r_label} / {dno}")
                elif n_rows > 1:
                    subplot_titles.append(str(r_label))
                else:
                    subplot_titles.append(str(dno))
        fig = make_subplots(
            rows=n_rows,
            cols=n_cols,
            subplot_titles=subplot_titles,
            shared_yaxes=True,
            vertical_spacing=0.08,
        )
    else:
        fig = go.Figure()

    legend_shown: set[str] = set()

    for row_idx, row_label in enumerate(unique_rows):
        plotly_row = row_idx + 1
        row_mask = row_labels == row_label

        for col_idx, dno in enumerate(valid_dim_nos, 1):
            dmeta = dim_metas[dno]
            col_labels, point_nums, nominals, usls, lsls = get_filtered_dim_meta(
                dmeta, exclude_intervals=exclude_intervals
            )
            valid_cols = [
                c
                for c, pn in zip(col_labels, point_nums)
                if c in df.columns and (_point_filter is None or pn in _point_filter)
            ]
            if not valid_cols:
                continue

            usl_val = next((v for v in usls if v is not None), None)
            lsl_val = next((v for v in lsls if v is not None), None)
            nom_val = next((v for v in nominals if v is not None), None)

            for grp_name in unique_colors:
                grp_mask = (color_series == grp_name) & row_mask
                values = (
                    df.loc[grp_mask, valid_cols]
                    .apply(pd.to_numeric, errors="coerce")
                    .values.flatten()
                )
                values = values[~np.isnan(values)]

                if len(values) == 0:
                    continue

                show_legend = grp_name not in legend_shown
                legend_shown.add(grp_name)

                trace = go.Histogram(
                    x=values,
                    name=grp_name,
                    legendgroup=grp_name,
                    marker_color=color_map[grp_name],
                    opacity=0.6,
                    nbinsx=nbins,
                    showlegend=show_legend,
                )

                if use_subplots:
                    fig.add_trace(trace, row=plotly_row, col=col_idx)
                else:
                    fig.add_trace(trace)

            line_kwargs: dict = dict(row=plotly_row, col=col_idx) if use_subplots else {}
            render_spec_limits(
                fig,
                [SpecSpan(usl=usl_val, lsl=lsl_val, nominal=nom_val)],
                style=HISTOGRAM_STYLE,
                orientation="v",
                rows=[line_kwargs],
            )

    chart_height = max(400, 300 * n_rows)
    fig.update_layout(
        title=dict(
            text=f"<b>Histogram: {group_label}</b>", font=dict(size=15), x=0.5, xanchor="center"
        ),
        barmode="overlay",
        height=chart_height,
        margin=dict(l=50, r=120, t=80, b=60),
        legend=dict(
            title=dict(text=color_by if color_by != "None" else ""),
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02,
            font=dict(size=11),
            bgcolor="rgba(255,255,255,0.8)",
        ),
        template="plotly_white",
    )

    if not use_subplots:
        fig.update_xaxes(title_text="Value")
        fig.update_yaxes(title_text="Count")

    return fig
