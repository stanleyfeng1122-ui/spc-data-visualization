"""Box plot chart builder.

Builds the per-point box-plot showing distribution of measurements,
with optional row facets, color grouping, and spec-limit overlays.
Pure code movement from the original ``chart_utils`` module.
"""

from __future__ import annotations

from collections import OrderedDict, defaultdict

import pandas as pd
import plotly.graph_objects as go
from plotly.graph_objects import Figure
from plotly.subplots import make_subplots

from spc_viz.parsers import get_filtered_dim_meta
from spc_viz.parsers.dimensions import DimensionMeta

from .base import compute_row_groups, compute_sections, get_color_for_group
from .spec_limits import BOX_STYLE, SpecSpan, render_spec_limits, spec_axis_annotations

# ---------------------------------------------------------------------------
# Chart building -- box plot
# ---------------------------------------------------------------------------


def _add_section_bands(fig: Figure, ordered_cats: list[str], cat_section: dict[str, str | None]) -> None:
    """Draw combined-profile-style section header bands + dividers (paper coords).

    Each section gets a header rect above the plot with a centered label, and a
    vertical divider separates adjacent sections — matching the profile chart.
    Categories must already be in axis order; sections must be contiguous.
    """
    n = len(ordered_cats)
    if n == 0:
        return

    ranges: list[tuple[str, int, int]] = []  # (label, first_idx, last_idx)
    start = 0
    for i in range(1, n + 1):
        if i == n or cat_section[ordered_cats[i]] != cat_section[ordered_cats[start]]:
            label = cat_section[ordered_cats[start]]
            if label is not None:
                ranges.append((label, start, i - 1))
            start = i

    shapes = list(fig.layout.shapes or [])
    annotations = list(fig.layout.annotations or [])
    for label, i0, i1 in ranges:
        px0, px1 = i0 / n, (i1 + 1) / n
        shapes.append(
            dict(
                type="rect", xref="paper", yref="paper",
                x0=px0, x1=px1, y0=1.01, y1=1.07,
                fillcolor="#F1F5F9", line=dict(color="#E2E8F0", width=1), layer="above",
            )
        )
        annotations.append(
            dict(
                x=(px0 + px1) / 2, y=1.04, xref="paper", yref="paper",
                text=f"<b>{label}</b>", showarrow=False, xanchor="center", yanchor="middle",
                font=dict(size=11, color="#334155", family="Arial"),
            )
        )
    # Vertical dividers between adjacent sections, spanning the plot height.
    for _, _, i1 in ranges[:-1]:
        bx = (i1 + 1) / n
        shapes.append(
            dict(
                type="line", xref="paper", yref="paper",
                x0=bx, x1=bx, y0=0, y1=1, line=dict(color="rgba(71,85,105,0.55)", width=1.4),
            )
        )
    fig.update_layout(shapes=shapes, annotations=annotations)


def build_box_plot(
    df: pd.DataFrame,
    dim_metas: OrderedDict[str, DimensionMeta],
    dim_nos: list[str],
    color_by: str,
    y_axis_mode: str,
    exclude_intervals: bool,
    group_label: str,
    row_by: str = "None",
    custom_color_map: dict[str, str] | None = None,
    custom_yrange: list[float] | None = None,
    selected_points: list[str] | None = None,
    section_by_fields: list[str] | None = None,
) -> Figure | None:
    """Build a box plot showing the distribution of measurements at each point.

    Each box overlays all of its data points (jittered) so the sample size is
    visible. ``section_by_fields`` splits the boxes along the x-axis by metadata
    value (e.g. one box per Factory), composing with ``color_by`` (grouped
    boxes at each x) and ``row_by`` (vertical facets).
    """
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    # Normalise selected_points to a set for O(1) lookup; None/empty means show all
    _point_filter: set[str] | None = set(selected_points) if selected_points else None

    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))
    n_rows = len(unique_rows)
    use_row_facets = n_rows > 1

    # Section labels go on the x-axis (None sentinel = no sectioning).
    section_series: pd.Series | None
    if section_by_fields:
        sec = compute_sections(df, section_by_fields)
        section_series = sec
        unique_sections: list = sorted(sec.unique())
    else:
        section_series = None
        unique_sections = [None]

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

    fig: Figure
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

    legend_shown: set[str] = set()
    multi_dim = len(dim_nos) > 1
    rep_usl, rep_lsl, rep_nom = None, None, None

    # Valid (col, point, spec) tuples per dim are row-independent — compute once.
    dim_valids: OrderedDict[str, list] = OrderedDict()
    for dno in dim_nos:
        if dno not in dim_metas:
            continue
        col_labels, point_nums, nominals, usls, lsls = get_filtered_dim_meta(
            dim_metas[dno], exclude_intervals=exclude_intervals
        )
        valid = [
            (cl, pn, n, u, l)
            for cl, pn, n, u, l in zip(col_labels, point_nums, nominals, usls, lsls)
            if cl in df.columns and (_point_filter is None or pn in _point_filter)
        ]
        if valid:
            dim_valids[dno] = valid
    # One measurement point overall -> the section value alone is the x label.
    single_point = sum(len(v) for v in dim_valids.values()) <= 1
    sectioned = section_series is not None

    # Track x categories in axis order so section header bands (drawn in paper
    # coords like the combined profile) can span each section's contiguous slice.
    # Section is the OUTER loop so each section's points stay contiguous.
    ordered_cats: list[str] = []
    cat_section: dict[str, str | None] = {}
    cat_tick: dict[str, str] = {}
    # Per-box stats for the readable mean/median labels.
    box_stats: list[tuple[str, int | None, str, float, float, float, float]] = []

    for row_idx, row_label in enumerate(unique_rows):
        plotly_row = row_idx + 1 if use_row_facets else None
        row_mask = row_labels == row_label
        row_df = df[row_mask]
        row_colors = color_series[row_mask]
        row_sections = section_series[row_mask] if section_series is not None else None

        for sec_value in unique_sections:
            sec_mask = (row_sections == sec_value) if (sectioned and sec_value is not None) else None

            for dno, valid in dim_valids.items():
                for col_label, point_num, nominal, usl_val, lsl_val in valid:
                    base_point = f"{dno}_{point_num}" if multi_dim else point_num
                    if rep_usl is None and usl_val is not None:
                        rep_usl, rep_lsl, rep_nom = usl_val, lsl_val, nominal

                    if not sectioned:
                        cat, tick, sec_lbl = base_point, base_point, None
                    elif single_point:
                        cat, tick, sec_lbl = str(sec_value), "", str(sec_value)
                    else:
                        cat = f"{sec_value} · {base_point}"
                        tick, sec_lbl = base_point, str(sec_value)

                    for grp_name in unique_colors:
                        grp_mask = row_colors == grp_name
                        if sec_mask is not None:
                            grp_mask = grp_mask & sec_mask
                        values = pd.to_numeric(
                            row_df.loc[grp_mask, col_label], errors="coerce"
                        ).dropna()
                        if deviation_mode and nominal is not None:
                            values = values - nominal
                        if len(values) == 0:
                            continue

                        if cat not in cat_section:
                            ordered_cats.append(cat)
                            cat_section[cat] = sec_lbl
                            cat_tick[cat] = tick

                        show_legend = grp_name not in legend_shown
                        legend_shown.add(grp_name)

                        color = color_map[grp_name]
                        trace = go.Box(
                            y=values,
                            x=[cat] * len(values),
                            name=grp_name,
                            legendgroup=grp_name,
                            showlegend=show_legend,
                            boxpoints="all",
                            jitter=0.5,
                            pointpos=0,
                            boxmean=True,
                            marker=dict(color=color, size=3, opacity=0.4),
                            line=dict(color=color, width=1.2),
                        )
                        if use_row_facets:
                            fig.add_trace(trace, row=plotly_row, col=1)
                        else:
                            fig.add_trace(trace)
                        box_stats.append(
                            (
                                cat,
                                plotly_row,
                                grp_name,
                                float(values.mean()),
                                float(values.median()),
                                float(values.max()),
                                float(values.min()),
                            )
                        )

    row_kwargs_list = [dict(row=i + 1, col=1) for i in range(n_rows)] if use_row_facets else [{}]
    spec_spans = [SpecSpan(usl=rep_usl, lsl=rep_lsl, nominal=rep_nom)]
    render_spec_limits(
        fig,
        spec_spans,
        style=BOX_STYLE,
        deviation_mode=deviation_mode,
        rows=row_kwargs_list,
    )

    # Section header bands (profile style) replace tilted x-axis section labels.
    distinct_secs = [s for s in dict.fromkeys(cat_section.values()) if s is not None]
    show_bands = (not use_row_facets) and len(distinct_secs) >= 2

    # Few categories render in a narrower column (see chart_view); shorten the
    # height there so the aspect stays landscape instead of a tall sliver.
    if use_row_facets:
        chart_height = 350 * n_rows
    else:
        chart_height = 460 if len(ordered_cats) <= 6 else 620
    y_range_kwargs = dict(range=custom_yrange) if custom_yrange else {}
    x_title = (
        " / ".join(section_by_fields)
        if (section_by_fields and single_point)
        else "Measurement Point"
    )
    fig.update_layout(
        title=dict(
            text=f"<b>Box Plot: {group_label}</b>",
            font=dict(size=15),
            x=0.5,
            xanchor="center",
            y=0.98 if show_bands else None,
            yanchor="top",
        ),
        xaxis=dict(title=x_title, tickangle=-45, tickfont=dict(size=8, color="#000000")),
        yaxis=dict(title="Deviation from Nominal" if deviation_mode else "Value", **y_range_kwargs),
        boxmode="group",
        height=chart_height,
        margin=dict(l=50, r=120, t=120 if show_bands else 80, b=100),
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

    spec_annotations = spec_axis_annotations(
        spec_spans, style=BOX_STYLE, deviation_mode=deviation_mode
    )
    if spec_annotations:
        existing = list(fig.layout.annotations or [])
        fig.update_layout(annotations=existing + spec_annotations)

    # Lock the x category order so section bands line up; show point-only ticks.
    if ordered_cats:
        fig.update_xaxes(categoryorder="array", categoryarray=ordered_cats)
    if sectioned and ordered_cats and any(cat_tick[c] != c for c in ordered_cats):
        fig.update_xaxes(tickvals=ordered_cats, ticktext=[cat_tick[c] for c in ordered_cats])

    if show_bands:
        _add_section_bands(fig, ordered_cats, cat_section)

    # Readable mean/median labels above each box (mean line is drawn by boxmean).
    if box_stats:
        span = (max(s[5] for s in box_stats) - min(s[6] for s in box_stats)) or 1.0
        multi = len(unique_colors) > 1
        by_cat: dict[tuple[str, int | None], list] = defaultdict(list)
        for cat, prow, grp, mean, med, vmax, _vmin in box_stats:
            by_cat[(cat, prow)].append((grp, mean, med, vmax))
        for (cat, prow), items in by_cat.items():
            base = max(v for *_, v in items)
            for k, (grp, mean, _med, _v) in enumerate(items):
                txt = (f"{grp}: " if multi else "") + f"Avg {mean:.4g}"
                ann = dict(
                    x=cat,
                    y=base + span * 0.03 * (k + 1),
                    text=txt,
                    showarrow=False,
                    yanchor="bottom",
                    align="center",
                    font=dict(size=9, color="#1e3a8a"),
                    bgcolor="rgba(255,255,255,0.7)",
                )
                if use_row_facets:
                    fig.add_annotation(**ann, row=prow, col=1)
                else:
                    fig.add_annotation(**ann)

    return fig
