"""Combined Profile chart builder.

Builds the per-point profile chart with optional row facets and section
groupings. Pure code movement from the original ``chart_utils`` module.
"""

from collections import OrderedDict

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from spc_viz.parsers import get_filtered_dim_meta

from .base import (
    MAX_TRACES_PER_GROUP,
    compute_row_groups,
    compute_sections,
    get_color_for_group,
)


# ---------------------------------------------------------------------------
# Chart building -- combined profile view
# ---------------------------------------------------------------------------


def build_combined_chart(
    df,
    dim_metas: OrderedDict,
    dim_nos: list,
    section_by_fields: list,
    color_by: str,
    y_axis_mode: str,
    exclude_intervals: bool,
    group_label: str,
    row_by: str = "None",
    custom_color_map: dict = None,
    custom_yrange: list = None,
    selected_points: list = None,
):
    """Build the combined profile chart with section and row facets."""
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    section_labels = compute_sections(df, section_by_fields)
    unique_sections = list(dict.fromkeys(section_labels))

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

    # Normalise selected_points to a set for O(1) lookup; None/empty means show all
    _point_filter = set(selected_points) if selected_points else None

    dim_point_info = OrderedDict()
    for dno in dim_nos:
        if dno not in dim_metas:
            continue
        dmeta = dim_metas[dno]
        info = get_filtered_dim_meta(dmeta, exclude_intervals=exclude_intervals)
        col_labels, point_numbers, nominal, usl, lsl = info
        valid = [
            (cl, pn, n, u, l)
            for cl, pn, n, u, l in zip(col_labels, point_numbers, nominal, usl, lsl)
            if cl in df.columns and (_point_filter is None or pn in _point_filter)
        ]
        if valid:
            cls, pns, noms, usls, lsls = zip(*valid)
            dim_point_info[dno] = (list(cls), list(pns), list(noms), list(usls), list(lsls))

    if not dim_point_info:
        return None

    points_per_section = sum(len(v[0]) for v in dim_point_info.values())
    if points_per_section == 0:
        return None

    section_gap = max(3, int(points_per_section * 0.06))

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

    first_dim_info = list(dim_point_info.values())[0]
    usl_rep = next((v for v in first_dim_info[3] if v is not None), None)
    lsl_rep = next((v for v in first_dim_info[4] if v is not None), None)
    nom_rep = next((v for v in first_dim_info[2] if v is not None), None)

    legend_shown = set()

    all_tick_vals = []
    all_tick_text = []
    section_boundaries = []

    x_offset = 0
    section_x_ranges = {}
    dim_x_positions = OrderedDict()

    for sec_idx, sec_label in enumerate(unique_sections):
        section_start_x = x_offset
        for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
            n_points = len(col_labels)
            x_positions = list(range(x_offset, x_offset + n_points))
            dim_x_positions[(sec_label, dno)] = x_positions

            for xi, pn in zip(x_positions, point_nums):
                all_tick_vals.append(xi)
                all_tick_text.append(pn if pn else "")
            x_offset += n_points

        section_end_x = x_offset
        section_x_ranges[sec_label] = (section_start_x, section_end_x)
        if sec_idx < len(unique_sections) - 1:
            section_boundaries.append(x_offset + section_gap / 2)
            x_offset += section_gap

    for row_idx, row_label in enumerate(unique_rows):
        plotly_row = row_idx + 1 if use_row_facets else None
        row_mask = row_labels == row_label

        for sec_label in unique_sections:
            sec_mask = section_labels == sec_label
            combined_mask = row_mask & sec_mask
            cell_df = df[combined_mask].reset_index(drop=True)
            cell_colors = color_series[combined_mask].reset_index(drop=True)

            if cell_df.empty:
                continue

            for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
                x_positions = dim_x_positions[(sec_label, dno)]
                nom_array = np.array(
                    [n if n is not None else np.nan for n in nominals],
                    dtype=float,
                )

                for grp_name in unique_colors:
                    grp_mask = cell_colors == grp_name
                    grp_df = cell_df.loc[grp_mask, col_labels].reset_index(drop=True)
                    color = color_map[grp_name]

                    n_parts = len(grp_df)
                    if n_parts == 0:
                        continue

                    step = max(1, n_parts // MAX_TRACES_PER_GROUP)

                    # Single-point dimensions (e.g. flatness) need markers;
                    # scattergl lines with 1 point render nothing.
                    is_single_point = len(x_positions) == 1
                    trace_mode = "markers" if is_single_point else "lines"

                    for ri in range(0, n_parts, step):
                        y_vals = pd.to_numeric(grp_df.iloc[ri], errors="coerce").values.copy()
                        if deviation_mode:
                            y_vals = y_vals - nom_array

                        show_legend = grp_name not in legend_shown
                        legend_shown.add(grp_name)

                        trace_kwargs = dict(
                            x=x_positions,
                            y=y_vals,
                            mode=trace_mode,
                            opacity=0.45,
                            name=grp_name,
                            legendgroup=grp_name,
                            showlegend=show_legend,
                            hovertemplate=(
                                "Point: %{text}<br>"
                                "Value: %{y:.4f}<br>"
                                f"{color_by}: {grp_name}<br>"
                                f"Section: {sec_label}<br>"
                                f"Row: {row_label}"
                                "<extra></extra>"
                            ),
                            text=[pn for pn in point_nums],
                        )
                        if is_single_point:
                            trace_kwargs["marker"] = dict(size=6, color=color)
                        else:
                            trace_kwargs["line"] = dict(width=0.7, color=color)

                        trace = go.Scattergl(**trace_kwargs)
                        if use_row_facets:
                            fig.add_trace(trace, row=plotly_row, col=1)
                        else:
                            fig.add_trace(trace)

    row_kwargs_list = [dict(row=i + 1, col=1) for i in range(n_rows)] if use_row_facets else [{}]

    dash_style = dict(dash="dash", width=1.2)
    for rk in row_kwargs_list:
        if usl_rep is not None and lsl_rep is not None:
            band_usl = (usl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else usl_rep
            band_lsl = (lsl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else lsl_rep
            fig.add_hrect(
                y0=band_lsl,
                y1=band_usl,
                fillcolor="rgba(34, 197, 94, 0.15)",
                line_width=0,
                layer="below",
                **rk,
            )

        if usl_rep is not None:
            ref_usl = (usl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else usl_rep
            fig.add_hline(y=ref_usl, line=dict(color="rgba(220,38,38,0.5)", **dash_style), **rk)

        if lsl_rep is not None:
            ref_lsl = (lsl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else lsl_rep
            fig.add_hline(y=ref_lsl, line=dict(color="rgba(220,38,38,0.5)", **dash_style), **rk)

    for bx in section_boundaries:
        fig.add_vline(x=bx, line=dict(color="rgba(100,116,139,0.5)", width=1.5, dash="solid"))

    annotations = []

    is_group = len(dim_nos) > 1
    if is_group:
        dim_names = "/".join(dno.replace("SPC_", "") for dno in dim_nos)
        first_desc = ""
        for dno in dim_nos:
            if dno in dim_metas and dim_metas[dno].description:
                desc = dim_metas[dno].description
                for keyword in [
                    "z straightness",
                    "flatness",
                    "overall length",
                    "half length",
                    "half width",
                    "height",
                ]:
                    if keyword in desc.lower():
                        first_desc = keyword.title()
                        break
                if first_desc:
                    break
        title_text = f"SPC_{dim_names}"
        if first_desc:
            title_text += f", {first_desc}"
    else:
        dno = dim_nos[0]
        dmeta = dim_metas.get(dno)
        desc = dmeta.description if dmeta else ""
        title_text = f"{dno}, {desc}" if desc else dno

    subtitle = ""  # Don't show section field names (e.g. "Factory") as subtitle
    y_title = "Deviation from Nominal" if deviation_mode else ""

    tick_step = max(1, len(all_tick_vals) // 80)
    tick_kwargs = dict(
        tickmode="array",
        tickvals=all_tick_vals[::tick_step],
        ticktext=all_tick_text[::tick_step],
        tickangle=-90,
        tickfont=dict(size=7, color="#000000"),
        showgrid=False,
    )

    chart_height = 350 * n_rows if use_row_facets else 620

    fig.update_layout(
        title=dict(
            text=f"<b>{title_text}</b>"
            + (
                f"<br><span style='font-size:12px;color:#64748B'>{subtitle}</span>"
                if subtitle
                else ""
            ),
            font=dict(size=15),
            x=0.5,
            xanchor="center",
        ),
        height=chart_height,
        margin=dict(l=50, r=120, t=120, b=80),
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
        annotations=annotations,
        hovermode="closest",
        template="plotly_white",
    )

    spec_tickvals = []
    spec_ticktext = []
    if usl_rep is not None:
        ref_usl = (usl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else usl_rep
        spec_tickvals.append(ref_usl)
        spec_ticktext.append(f"USL-{ref_usl:.4g}")
    if lsl_rep is not None:
        ref_lsl = (lsl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else lsl_rep
        spec_tickvals.append(ref_lsl)
        spec_ticktext.append(f"LSL-{ref_lsl:.4g}")

    y_range_kwargs = dict(range=custom_yrange) if custom_yrange else {}
    if use_row_facets:
        for i in range(1, n_rows + 1):
            x_axis_name = f"xaxis{i}" if i > 1 else "xaxis"
            y_axis_name = f"yaxis{i}" if i > 1 else "yaxis"
            show_ticks = i == n_rows
            fig.update_layout(
                **{
                    x_axis_name: dict(**tick_kwargs, showticklabels=show_ticks),
                    y_axis_name: dict(
                        title=y_title if i == (n_rows + 1) // 2 else "",
                        zeroline=True,
                        zerolinecolor="rgba(100,116,139,0.3)",
                        **y_range_kwargs,
                    ),
                }
            )
    else:
        fig.update_layout(
            xaxis=tick_kwargs,
            yaxis=dict(
                title=y_title,
                zeroline=True,
                zerolinecolor="rgba(100,116,139,0.3)",
                **y_range_kwargs,
            ),
        )

    for val, label in zip(spec_tickvals, spec_ticktext):
        annotations.append(
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
    fig.update_layout(annotations=annotations)

    # ----- Factory / section header bands (paper coordinates) -----
    total_x_span = x_offset  # total x-axis data range
    if total_x_span > 0 and len(unique_sections) > 1:
        header_shapes = []
        section_centers = []
        for sec_label, (sx0, sx1) in section_x_ranges.items():
            # Map data x-range to paper coordinates [0, 1]
            px0 = sx0 / total_x_span
            px1 = sx1 / total_x_span
            center_x = (px0 + px1) / 2
            section_centers.append((center_x, sec_label))
            header_shapes.append(
                dict(
                    type="rect",
                    xref="paper",
                    yref="paper",
                    x0=px0,
                    x1=px1,
                    y0=1.01,
                    y1=1.07,
                    fillcolor="#F1F5F9",
                    line=dict(color="#E2E8F0", width=1),
                    layer="above",
                )
            )
        # Merge with existing shapes (USL/LSL lines)
        existing_shapes = list(fig.layout.shapes or [])
        fig.update_layout(shapes=existing_shapes + header_shapes)

        # Add centered section labels
        for cx, sec_label in section_centers:
            annotations.append(
                dict(
                    x=cx,
                    y=1.04,
                    xref="paper",
                    yref="paper",
                    text=f"<b>{sec_label}</b>",
                    showarrow=False,
                    xanchor="center",
                    yanchor="middle",
                    font=dict(size=11, color="#334155"),
                )
            )
        fig.update_layout(annotations=annotations)

    return fig
