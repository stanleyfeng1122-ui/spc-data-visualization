"""Box plot chart builder.

Builds the per-point box-plot showing distribution of measurements,
with optional row facets, color grouping, and spec-limit overlays.
Pure code movement from the original ``chart_utils`` module.
"""

from collections import OrderedDict

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from spc_viz.parsers import get_filtered_dim_meta

from .base import compute_row_groups, get_color_for_group


# ---------------------------------------------------------------------------
# Chart building -- box plot
# ---------------------------------------------------------------------------


def build_box_plot(
    df,
    dim_metas: OrderedDict,
    dim_nos: list,
    color_by: str,
    y_axis_mode: str,
    exclude_intervals: bool,
    group_label: str,
    row_by: str = "None",
    custom_color_map: dict = None,
    custom_yrange: list = None,
    selected_points: list = None,
):
    """Build a box plot showing the distribution of measurements at each point."""
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    # Normalise selected_points to a set for O(1) lookup; None/empty means show all
    _point_filter = set(selected_points) if selected_points else None

    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))
    n_rows = len(unique_rows)
    use_row_facets = n_rows > 1

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

    if use_row_facets:
        fig = make_subplots(
            rows=n_rows,
            cols=1,
            shared_xaxes=True,
            row_titles=[str(r) for r in unique_rows],
            vertical_spacing=0.06,
        )
    else:
        fig = go.Figure()

    legend_shown = set()
    multi_dim = len(dim_nos) > 1
    rep_usl, rep_lsl, rep_nom = None, None, None

    for row_idx, row_label in enumerate(unique_rows):
        plotly_row = row_idx + 1 if use_row_facets else None
        row_mask = row_labels == row_label
        row_df = df[row_mask]
        row_colors = color_series[row_mask]

        for dno in dim_nos:
            if dno not in dim_metas:
                continue
            dmeta = dim_metas[dno]
            col_labels, point_nums, nominals, usls, lsls = get_filtered_dim_meta(
                dmeta, exclude_intervals=exclude_intervals
            )
            valid = [
                (cl, pn, n, u, l)
                for cl, pn, n, u, l in zip(col_labels, point_nums, nominals, usls, lsls)
                if cl in df.columns and (_point_filter is None or pn in _point_filter)
            ]
            if not valid:
                continue

            for col_label, point_num, nominal, usl_val, lsl_val in valid:
                x_label = f"{dno}_{point_num}" if multi_dim else point_num
                if rep_usl is None and usl_val is not None:
                    rep_usl, rep_lsl, rep_nom = usl_val, lsl_val, nominal

                for grp_name in unique_colors:
                    grp_mask = row_colors == grp_name
                    values = pd.to_numeric(
                        row_df.loc[grp_mask, col_label], errors="coerce"
                    ).dropna()
                    if deviation_mode and nominal is not None:
                        values = values - nominal

                    show_legend = grp_name not in legend_shown
                    legend_shown.add(grp_name)

                    trace = go.Box(
                        y=values,
                        x=[x_label] * len(values),
                        name=grp_name,
                        legendgroup=grp_name,
                        marker_color=color_map[grp_name],
                        showlegend=show_legend,
                        boxpoints="outliers",
                    )
                    if use_row_facets:
                        fig.add_trace(trace, row=plotly_row, col=1)
                    else:
                        fig.add_trace(trace)

    dash_style = dict(dash="dash", width=1.2)
    row_kwargs_list = [dict(row=i + 1, col=1) for i in range(n_rows)] if use_row_facets else [{}]
    for rk in row_kwargs_list:
        if rep_usl is not None:
            ref_usl = (rep_usl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_usl
            fig.add_hline(
                y=ref_usl,
                line=dict(color="rgba(220,38,38,0.5)", **dash_style),
                annotation_text="USL",
                annotation_position="top right",
                **rk,
            )
        if rep_lsl is not None:
            ref_lsl = (rep_lsl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_lsl
            fig.add_hline(
                y=ref_lsl,
                line=dict(color="rgba(220,38,38,0.5)", **dash_style),
                annotation_text="LSL",
                annotation_position="bottom right",
                **rk,
            )
        if rep_nom is not None:
            ref_nom = 0.0 if deviation_mode else rep_nom
            fig.add_hline(
                y=ref_nom,
                line=dict(color="rgba(34,197,94,0.5)", dash="dot", width=1),
                annotation_text="Nominal",
                annotation_position="top right",
                **rk,
            )
        if rep_usl is not None and rep_lsl is not None:
            band_usl = (rep_usl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_usl
            band_lsl = (rep_lsl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_lsl
            fig.add_hrect(
                y0=band_lsl,
                y1=band_usl,
                fillcolor="rgba(34, 197, 94, 0.10)",
                line_width=0,
                layer="below",
                **rk,
            )

    spec_tickvals = []
    spec_ticktext = []
    if rep_usl is not None:
        ref_usl_v = (rep_usl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_usl
        spec_tickvals.append(ref_usl_v)
        spec_ticktext.append(f"USL-{ref_usl_v:.4g}")
    if rep_lsl is not None:
        ref_lsl_v = (rep_lsl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_lsl
        spec_tickvals.append(ref_lsl_v)
        spec_ticktext.append(f"LSL-{ref_lsl_v:.4g}")

    chart_height = 350 * n_rows if use_row_facets else 620
    y_range_kwargs = dict(range=custom_yrange) if custom_yrange else {}
    fig.update_layout(
        title=dict(
            text=f"<b>Box Plot: {group_label}</b>", font=dict(size=15), x=0.5, xanchor="center"
        ),
        xaxis=dict(
            title="Measurement Point", tickangle=-45, tickfont=dict(size=8, color="#000000")
        ),
        yaxis=dict(title="Deviation from Nominal" if deviation_mode else "Value", **y_range_kwargs),
        boxmode="group",
        height=chart_height,
        margin=dict(l=50, r=120, t=80, b=100),
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
        hovermode="closest",
        template="plotly_white",
    )

    spec_annotations = []
    for val, label in zip(spec_tickvals, spec_ticktext):
        spec_annotations.append(
            dict(
                x=0.0,
                y=val,
                xref="paper",
                yref="y",
                text=f"<b>{label}</b>",
                showarrow=False,
                xanchor="right",
                font=dict(size=10, color="rgba(220,38,38,0.9)", family="Arial Black"),
                bgcolor="rgba(255,255,255,0.7)",
            )
        )
    if spec_annotations:
        existing = list(fig.layout.annotations or [])
        fig.update_layout(annotations=existing + spec_annotations)

    return fig
