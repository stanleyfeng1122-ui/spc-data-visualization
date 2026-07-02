"""Range envelope chart builder.

Shows the per-point min/max band with an overlaid mean line. This is useful
when the engineer wants the process spread by measurement point without
rendering every individual part trace.
"""

from __future__ import annotations

from collections import OrderedDict

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.graph_objects import Figure

from spc_viz.parsers import get_filtered_dim_meta
from spc_viz.parsers.dimensions import DimensionMeta

from .base import (
    compute_sections,
    get_color_for_group,
    has_domain_section_order,
    section_sort_key,
)
from .spec_limits import ENVELOPE_STYLE, SpecSpan, render_spec_limits, spec_axis_annotations


def _hex_to_rgba(color: str, alpha: float) -> str:
    if not color.startswith("#") or len(color) != 7:
        return f"rgba(37,99,235,{alpha})"
    red = int(color[1:3], 16)
    green = int(color[3:5], 16)
    blue = int(color[5:7], 16)
    return f"rgba({red},{green},{blue},{alpha})"


def _field_series(df: pd.DataFrame, field_name: str) -> pd.Series:
    if field_name == "Factory":
        if "_factory" in df.columns:
            return df["_factory"].fillna("?").astype(str)
        if "Factory" in df.columns:
            return df["Factory"].fillna("?").astype(str)
        return pd.Series("?", index=df.index)
    if field_name == "Source File":
        if "_source_file" in df.columns:
            return df["_source_file"].fillna("?").astype(str)
        return pd.Series("?", index=df.index)
    if field_name in df.columns:
        return df[field_name].fillna("?").astype(str)
    return pd.Series("?", index=df.index)


def build_range_envelope_chart(
    df: pd.DataFrame,
    dim_metas: OrderedDict[str, DimensionMeta],
    dim_nos: list[str],
    section_by_fields: list[str],
    color_by: str,
    y_axis_mode: str,
    exclude_intervals: bool,
    group_label: str,
    row_by: str = "None",
    custom_color_map: dict[str, str] | None = None,
    custom_yrange: list[float] | None = None,
    selected_points: list[str] | None = None,
) -> Figure | None:
    """Build a range envelope chart: min/max band plus mean line per group."""
    del row_by  # Row facets can be added later if this view needs them.
    deviation_mode = y_axis_mode == "Deviation from Nominal"
    point_filter: set[str] | None = set(selected_points) if selected_points else None

    dim_point_info: OrderedDict[str, tuple[list[str], list[str], list, list, list]] = OrderedDict()
    for dno in dim_nos:
        if dno not in dim_metas:
            continue
        col_labels, point_nums, nominals, usls, lsls = get_filtered_dim_meta(
            dim_metas[dno], exclude_intervals=exclude_intervals
        )
        valid = [
            (cl, pn, n, u, l)
            for cl, pn, n, u, l in zip(col_labels, point_nums, nominals, usls, lsls)
            if cl in df.columns and (point_filter is None or pn in point_filter)
        ]
        if valid:
            cls, pns, noms, usl_vals, lsl_vals = zip(*valid)
            dim_point_info[dno] = (
                list(cls),
                list(pns),
                list(noms),
                list(usl_vals),
                list(lsl_vals),
            )

    if not dim_point_info:
        return None

    section_labels = compute_sections(df, section_by_fields)
    unique_sections = list(dict.fromkeys(section_labels))
    section_field_values = [_field_series(df, field) for field in section_by_fields]
    section_parts: dict[str, tuple[str, ...]] = {}
    for sec_label in unique_sections:
        matching_rows = section_labels[section_labels == sec_label]
        if matching_rows.empty or not section_field_values:
            section_parts[sec_label] = (str(sec_label),)
            continue
        first_idx = matching_rows.index[0]
        section_parts[sec_label] = tuple(str(values.loc[first_idx]) for values in section_field_values)

    if len(section_by_fields) > 1 or has_domain_section_order(
        section_by_fields, section_parts.values()
    ):
        unique_sections = sorted(
            unique_sections,
            key=lambda label: section_sort_key(
                section_by_fields,
                section_parts.get(label, (str(label),)),
            ),
        )

    if color_by != "None":
        color_series = _field_series(df, color_by)
        unique_colors = sorted(color_series.unique())
    else:
        color_series = pd.Series("All", index=df.index)
        unique_colors = ["All"]

    if custom_color_map:
        color_map = {
            group: custom_color_map.get(group, get_color_for_group(i))
            for i, group in enumerate(unique_colors)
        }
    else:
        color_map = {group: get_color_for_group(i) for i, group in enumerate(unique_colors)}

    multi_dim = len(dim_nos) > 1
    points_per_section = sum(len(info[0]) for info in dim_point_info.values())
    section_gap = 0 if points_per_section <= 1 else max(2, int(points_per_section * 0.04))

    x_offset = 0
    x_positions_by_key: OrderedDict[tuple[str, str], list[float]] = OrderedDict()
    all_tick_vals: list[float] = []
    all_tick_text: list[str] = []
    section_boundaries: list[float] = []
    spec_segments: list[SpecSpan] = []

    for sec_idx, sec_label in enumerate(unique_sections):
        for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
            x_positions = [x_offset + idx + 0.5 for idx in range(len(col_labels))]
            x_positions_by_key[(sec_label, dno)] = x_positions
            for xi, point_num in zip(x_positions, point_nums):
                all_tick_vals.append(xi)
                label = f"{dno}_{point_num}" if multi_dim else point_num
                all_tick_text.append(label if label else "")
            for xi, nominal, usl, lsl in zip(x_positions, nominals, usls, lsls):
                spec_segments.append(
                    SpecSpan(usl=usl, lsl=lsl, nominal=nominal, x0=xi - 0.5, x1=xi + 0.5)
                )
            x_offset += len(col_labels)

        if sec_idx < len(unique_sections) - 1:
            section_boundaries.append(x_offset + section_gap / 2)
            x_offset += section_gap

    fig = go.Figure()
    legend_shown: set[str] = set()

    for sec_label in unique_sections:
        sec_mask = section_labels == sec_label
        for grp_name in unique_colors:
            grp_mask = color_series == grp_name
            row_mask = sec_mask & grp_mask
            if not row_mask.any():
                continue

            group_x: list[float] = []
            mean_vals: list[float | None] = []
            min_vals: list[float | None] = []
            max_vals: list[float | None] = []
            hover_labels: list[str] = []

            for dno, (col_labels, point_nums, nominals, _, _) in dim_point_info.items():
                x_positions = x_positions_by_key[(sec_label, dno)]
                for xi, col_label, point_num, nominal in zip(
                    x_positions, col_labels, point_nums, nominals
                ):
                    values = pd.to_numeric(df.loc[row_mask, col_label], errors="coerce").dropna()
                    if deviation_mode and nominal is not None:
                        values = values - nominal

                    group_x.append(xi)
                    hover_labels.append(f"{sec_label} / {dno}_{point_num}")
                    if values.empty:
                        mean_vals.append(None)
                        min_vals.append(None)
                        max_vals.append(None)
                    else:
                        mean_vals.append(float(values.mean()))
                        min_vals.append(float(values.min()))
                        max_vals.append(float(values.max()))

            color = color_map[grp_name]
            legend_name = grp_name if grp_name != "All" else "Envelope"
            show_legend = grp_name not in legend_shown
            legend_shown.add(grp_name)

            fig.add_trace(
                go.Scatter(
                    x=group_x,
                    y=min_vals,
                    mode="lines",
                    line=dict(width=0, color=color),
                    showlegend=False,
                    hoverinfo="skip",
                    legendgroup=grp_name,
                    name=f"{legend_name} min",
                )
            )
            fig.add_trace(
                go.Scatter(
                    x=group_x,
                    y=max_vals,
                    mode="lines",
                    fill="tonexty",
                    fillcolor=_hex_to_rgba(color, 0.18),
                    line=dict(width=0, color=color),
                    showlegend=False,
                    hoverinfo="skip",
                    legendgroup=grp_name,
                    name=f"{legend_name} max",
                )
            )
            fig.add_trace(
                go.Scatter(
                    x=group_x,
                    y=mean_vals,
                    mode="lines+markers",
                    line=dict(width=2.1, color=color),
                    marker=dict(size=4, color=color),
                    name=legend_name,
                    legendgroup=grp_name,
                    showlegend=show_legend,
                    text=hover_labels,
                    hovertemplate=(
                        "%{text}<br>"
                        "Mean: %{y:.4f}<br>"
                        f"{color_by}: {grp_name}<extra></extra>"
                    ),
                )
            )

    unique_usls = {s.usl for s in spec_segments if s.usl is not None}
    unique_lsls = {s.lsl for s in spec_segments if s.lsl is not None}
    use_stepping = len(unique_usls) > 1 or len(unique_lsls) > 1

    # Uniform spec collapses to one full-width span (a single hline + band);
    # stepping keeps the per-point segments.
    if not use_stepping and spec_segments:
        s0 = spec_segments[0]
        spec_spans = [SpecSpan(usl=s0.usl, lsl=s0.lsl, nominal=s0.nominal)]
    else:
        spec_spans = spec_segments
    render_spec_limits(fig, spec_spans, style=ENVELOPE_STYLE, deviation_mode=deviation_mode)

    for boundary_x in section_boundaries:
        fig.add_vline(x=boundary_x, line=dict(color="rgba(100,116,139,0.45)", width=1.1))

    tick_step = max(1, len(all_tick_vals) // 80)
    y_title = "Deviation from Nominal" if deviation_mode else "Value"
    y_range_kwargs = dict(range=custom_yrange) if custom_yrange else {}
    annotations: list[dict] = spec_axis_annotations(
        spec_spans, style=ENVELOPE_STYLE, deviation_mode=deviation_mode
    )

    fig.update_layout(
        title=dict(
            text=f"<b>Range Envelope: {group_label}</b>",
            font=dict(size=15),
            x=0.5,
            xanchor="center",
        ),
        xaxis=dict(
            tickmode="array",
            tickvals=all_tick_vals[::tick_step],
            ticktext=all_tick_text[::tick_step],
            tickangle=-90,
            tickfont=dict(size=7, color="#000000"),
            showgrid=False,
            range=[0, x_offset] if x_offset > 0 else None,
        ),
        yaxis=dict(title=y_title, zeroline=True, zerolinecolor="rgba(100,116,139,0.3)", **y_range_kwargs),
        height=620,
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
        annotations=annotations,
        hovermode="closest",
        template="plotly_white",
    )

    return fig
