"""Shared chart-building and SPC analysis utilities.

Extracted from app.py so both the main app and the Quick Test page
can share the same logic without duplication.
"""

import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from collections import OrderedDict
from scipy import stats as scipy_stats

from spc_parser import get_filtered_dim_meta

# ---------------------------------------------------------------------------
# Color palettes (no purple)
# ---------------------------------------------------------------------------
COLOR_PALETTE = [
    "#2563EB",  # blue
    "#DC2626",  # red
    "#059669",  # green
    "#D97706",  # amber
    "#0891B2",  # cyan
    "#E11D48",  # rose
    "#4F46E5",  # indigo
    "#EA580C",  # orange
    "#0D9488",  # teal
    "#64748B",  # slate
]

MAX_TRACES_PER_GROUP = 600


def get_color_for_group(idx: int) -> str:
    return COLOR_PALETTE[idx % len(COLOR_PALETTE)]


# ---------------------------------------------------------------------------
# Data preparation
# ---------------------------------------------------------------------------

def _get_factory(pf):
    """Get factory code for a parsed file dict."""
    return pf.get("factory") or "Unknown"


def prepare_combined_data(parsed_files, dim_nos):
    """
    Combine data from all files for the requested dimensions.
    Returns (df, dim_metas_dict) where df has all rows and a _factory column.
    """
    frames = []
    dim_metas = OrderedDict()

    for pf in parsed_files:
        factory = _get_factory(pf)
        df = pf["data"].copy()
        df["_factory"] = factory
        df["_source_file"] = pf["filename"]

        for dno in dim_nos:
            if dno in pf["dimensions"] and dno not in dim_metas:
                dim_metas[dno] = pf["dimensions"][dno]

        meta_cols = [c for c in pf["meta_columns"] if c in df.columns]
        meas_cols = []
        for dno in dim_nos:
            if dno in pf["dimensions"]:
                dmeta = pf["dimensions"][dno]
                meas_cols.extend([c for c in dmeta.col_labels if c in df.columns])

        keep = meta_cols + meas_cols + ["_factory", "_source_file"]
        keep = [c for c in keep if c in df.columns]
        df = df[keep]
        frames.append(df)

    if not frames:
        return None, None

    combined = pd.concat(frames, ignore_index=True)
    return combined, dim_metas


# ---------------------------------------------------------------------------
# Section / row logic
# ---------------------------------------------------------------------------

def compute_sections(df, section_by_fields):
    """
    Assign a section label to each row based on selected fields.
    Returns a Series of section labels aligned with df index.
    """
    if not section_by_fields:
        return pd.Series("All", index=df.index)

    def _get_col(field_name):
        if field_name == "Factory":
            if "_factory" in df.columns:
                return df["_factory"].fillna("?").astype(str)
            return pd.Series("?", index=df.index)
        elif field_name == "Source File":
            if "_source_file" in df.columns:
                return df["_source_file"].fillna("?").astype(str)
            return pd.Series("?", index=df.index)
        elif field_name in df.columns:
            return df[field_name].fillna("?").astype(str)
        return pd.Series("?", index=df.index)

    parts = [_get_col(f) for f in section_by_fields]
    combined = parts[0]
    for p in parts[1:]:
        combined = combined + " " + p
    return combined


def compute_row_groups(df, row_by):
    """
    Assign a row group label to each row based on row_by field.
    Returns a Series of row labels aligned with df index.
    """
    if row_by == "None" or row_by not in df.columns:
        return pd.Series("All", index=df.index)
    return df[row_by].fillna("?").astype(str)


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
        color_map = {g: custom_color_map.get(g, get_color_for_group(i)) for i, g in enumerate(unique_colors)}
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
        valid = [(cl, pn, n, u, l) for cl, pn, n, u, l in
                 zip(col_labels, point_numbers, nominal, usl, lsl)
                 if cl in df.columns and (_point_filter is None or pn in _point_filter)]
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
            rows=n_rows, cols=1,
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
                    [n if n is not None else np.nan for n in nominals], dtype=float,
                )

                for grp_name in unique_colors:
                    grp_mask = cell_colors == grp_name
                    grp_df = cell_df.loc[grp_mask, col_labels].reset_index(drop=True)
                    color = color_map[grp_name]

                    n_parts = len(grp_df)
                    if n_parts == 0:
                        continue

                    step = max(1, n_parts // MAX_TRACES_PER_GROUP)

                    for ri in range(0, n_parts, step):
                        y_vals = pd.to_numeric(grp_df.iloc[ri], errors="coerce").values.copy()
                        if deviation_mode:
                            y_vals = y_vals - nom_array

                        show_legend = grp_name not in legend_shown
                        legend_shown.add(grp_name)

                        trace = go.Scattergl(
                            x=x_positions,
                            y=y_vals,
                            mode="lines",
                            line=dict(width=0.7, color=color),
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
                        if use_row_facets:
                            fig.add_trace(trace, row=plotly_row, col=1)
                        else:
                            fig.add_trace(trace)

    row_kwargs_list = [dict(row=i+1, col=1) for i in range(n_rows)] if use_row_facets else [{}]

    dash_style = dict(dash="dash", width=1.2)
    for rk in row_kwargs_list:
        if usl_rep is not None and lsl_rep is not None:
            band_usl = (usl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else usl_rep
            band_lsl = (lsl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else lsl_rep
            fig.add_hrect(y0=band_lsl, y1=band_usl,
                          fillcolor="rgba(34, 197, 94, 0.15)", line_width=0, layer="below", **rk)

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
                for keyword in ["z straightness", "flatness", "overall length",
                                "half length", "half width", "height"]:
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

    subtitle = " + ".join(section_by_fields) if section_by_fields else ""
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
            text=f"<b>{title_text}</b>" + (f"<br><span style='font-size:12px;color:#64748B'>{subtitle}</span>" if subtitle else ""),
            font=dict(size=15), x=0.5, xanchor="center",
        ),
        height=chart_height,
        margin=dict(l=50, r=120, t=60, b=80),
        legend=dict(
            title=dict(text=color_by if color_by != "None" else ""),
            orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
            font=dict(size=11), bgcolor="rgba(255,255,255,0.8)",
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
            show_ticks = (i == n_rows)
            fig.update_layout(**{
                x_axis_name: dict(**tick_kwargs, showticklabels=show_ticks),
                y_axis_name: dict(title=y_title if i == (n_rows + 1) // 2 else "",
                                  zeroline=True, zerolinecolor="rgba(100,116,139,0.3)",
                                  **y_range_kwargs),
            })
    else:
        fig.update_layout(
            xaxis=tick_kwargs,
            yaxis=dict(title=y_title, zeroline=True, zerolinecolor="rgba(100,116,139,0.3)",
                       **y_range_kwargs),
        )

    for val, label in zip(spec_tickvals, spec_ticktext):
        annotations.append(dict(
            x=0.0, y=val,
            xref="paper", yref="y",
            text=f"<b>{label}</b>",
            showarrow=False,
            xanchor="right",
            font=dict(size=10, color="rgba(220,38,38,0.9)", family="Arial Black"),
            bgcolor="rgba(255,255,255,0.7)",
        ))
    fig.update_layout(annotations=annotations)

    return fig


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
        color_map = {g: custom_color_map.get(g, get_color_for_group(i)) for i, g in enumerate(unique_colors)}
    else:
        color_map = {g: get_color_for_group(i) for i, g in enumerate(unique_colors)}

    if use_row_facets:
        fig = make_subplots(rows=n_rows, cols=1, shared_xaxes=True,
                            row_titles=[str(r) for r in unique_rows],
                            vertical_spacing=0.06)
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
            valid = [(cl, pn, n, u, l) for cl, pn, n, u, l in
                     zip(col_labels, point_nums, nominals, usls, lsls)
                     if cl in df.columns and (_point_filter is None or pn in _point_filter)]
            if not valid:
                continue

            for col_label, point_num, nominal, usl_val, lsl_val in valid:
                x_label = f"{dno}_{point_num}" if multi_dim else point_num
                if rep_usl is None and usl_val is not None:
                    rep_usl, rep_lsl, rep_nom = usl_val, lsl_val, nominal

                for grp_name in unique_colors:
                    grp_mask = row_colors == grp_name
                    values = pd.to_numeric(row_df.loc[grp_mask, col_label], errors="coerce").dropna()
                    if deviation_mode and nominal is not None:
                        values = values - nominal

                    show_legend = grp_name not in legend_shown
                    legend_shown.add(grp_name)

                    trace = go.Box(
                        y=values, x=[x_label] * len(values),
                        name=grp_name, legendgroup=grp_name,
                        marker_color=color_map[grp_name],
                        showlegend=show_legend, boxpoints="outliers",
                    )
                    if use_row_facets:
                        fig.add_trace(trace, row=plotly_row, col=1)
                    else:
                        fig.add_trace(trace)

    dash_style = dict(dash="dash", width=1.2)
    row_kwargs_list = [dict(row=i+1, col=1) for i in range(n_rows)] if use_row_facets else [{}]
    for rk in row_kwargs_list:
        if rep_usl is not None:
            ref_usl = (rep_usl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_usl
            fig.add_hline(y=ref_usl, line=dict(color="rgba(220,38,38,0.5)", **dash_style),
                          annotation_text="USL", annotation_position="top right", **rk)
        if rep_lsl is not None:
            ref_lsl = (rep_lsl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_lsl
            fig.add_hline(y=ref_lsl, line=dict(color="rgba(220,38,38,0.5)", **dash_style),
                          annotation_text="LSL", annotation_position="bottom right", **rk)
        if rep_nom is not None:
            ref_nom = 0.0 if deviation_mode else rep_nom
            fig.add_hline(y=ref_nom, line=dict(color="rgba(34,197,94,0.5)", dash="dot", width=1),
                          annotation_text="Nominal", annotation_position="top right", **rk)
        if rep_usl is not None and rep_lsl is not None:
            band_usl = (rep_usl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_usl
            band_lsl = (rep_lsl - rep_nom) if (deviation_mode and rep_nom is not None) else rep_lsl
            fig.add_hrect(y0=band_lsl, y1=band_usl,
                          fillcolor="rgba(34, 197, 94, 0.10)", line_width=0, layer="below", **rk)

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
        title=dict(text=f"<b>Box Plot: {group_label}</b>", font=dict(size=15),
                   x=0.5, xanchor="center"),
        xaxis=dict(title="Measurement Point", tickangle=-45, tickfont=dict(size=8, color="#000000")),
        yaxis=dict(title="Deviation from Nominal" if deviation_mode else "Value",
                   **y_range_kwargs),
        boxmode="group",
        height=chart_height,
        margin=dict(l=50, r=120, t=80, b=100),
        legend=dict(title=dict(text=color_by if color_by != "None" else ""),
                    orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
                    font=dict(size=11), bgcolor="rgba(255,255,255,0.8)"),
        hovermode="closest",
        template="plotly_white",
    )

    spec_annotations = []
    for val, label in zip(spec_tickvals, spec_ticktext):
        spec_annotations.append(dict(
            x=0.0, y=val,
            xref="paper", yref="y",
            text=f"<b>{label}</b>",
            showarrow=False,
            xanchor="right",
            font=dict(size=10, color="rgba(220,38,38,0.9)", family="Arial Black"),
            bgcolor="rgba(255,255,255,0.7)",
        ))
    if spec_annotations:
        existing = list(fig.layout.annotations or [])
        fig.update_layout(annotations=existing + spec_annotations)

    return fig


# ---------------------------------------------------------------------------
# Chart building -- histogram
# ---------------------------------------------------------------------------

def build_histogram(
    df,
    dim_metas: OrderedDict,
    dim_nos: list,
    color_by: str,
    exclude_intervals: bool,
    group_label: str,
    nbins: int = 40,
    row_by: str = "None",
    custom_color_map: dict = None,
    selected_points: list = None,
):
    """Build a histogram showing frequency distribution of measurement values."""
    # Normalise selected_points to a set for O(1) lookup; None/empty means show all
    _point_filter = set(selected_points) if selected_points else None

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
        color_map = {g: custom_color_map.get(g, get_color_for_group(i)) for i, g in enumerate(unique_colors)}
    else:
        color_map = {g: get_color_for_group(i) for i, g in enumerate(unique_colors)}

    n_cols = max(n_dim_cols, 1)
    n_rows = max(n_facet_rows, 1)
    use_subplots = n_cols > 1 or n_rows > 1

    if use_subplots:
        subplot_titles = []
        for r_label in unique_rows:
            for dno in valid_dim_nos:
                if n_rows > 1 and n_cols > 1:
                    subplot_titles.append(f"{r_label} / {dno}")
                elif n_rows > 1:
                    subplot_titles.append(str(r_label))
                else:
                    subplot_titles.append(str(dno))
        fig = make_subplots(
            rows=n_rows, cols=n_cols,
            subplot_titles=subplot_titles,
            shared_yaxes=True,
            vertical_spacing=0.08,
        )
    else:
        fig = go.Figure()

    legend_shown = set()

    for row_idx, row_label in enumerate(unique_rows):
        plotly_row = row_idx + 1
        row_mask = row_labels == row_label

        for col_idx, dno in enumerate(valid_dim_nos, 1):
            dmeta = dim_metas[dno]
            col_labels, point_nums, nominals, usls, lsls = get_filtered_dim_meta(
                dmeta, exclude_intervals=exclude_intervals
            )
            valid_cols = [c for c, pn in zip(col_labels, point_nums)
                          if c in df.columns and (_point_filter is None or pn in _point_filter)]
            if not valid_cols:
                continue

            usl_val = next((v for v in usls if v is not None), None)
            lsl_val = next((v for v in lsls if v is not None), None)
            nom_val = next((v for v in nominals if v is not None), None)

            for grp_name in unique_colors:
                grp_mask = (color_series == grp_name) & row_mask
                values = df.loc[grp_mask, valid_cols].apply(
                    pd.to_numeric, errors="coerce"
                ).values.flatten()
                values = values[~np.isnan(values)]

                if len(values) == 0:
                    continue

                show_legend = grp_name not in legend_shown
                legend_shown.add(grp_name)

                trace = go.Histogram(
                    x=values, name=grp_name, legendgroup=grp_name,
                    marker_color=color_map[grp_name], opacity=0.6,
                    nbinsx=nbins, showlegend=show_legend,
                )

                if use_subplots:
                    fig.add_trace(trace, row=plotly_row, col=col_idx)
                else:
                    fig.add_trace(trace)

            line_kwargs = dict(row=plotly_row, col=col_idx) if use_subplots else {}
            dash_style = dict(dash="dash", width=1.5)
            if usl_val is not None:
                fig.add_vline(x=usl_val, line=dict(color="rgba(220,38,38,0.7)", **dash_style),
                              annotation_text="USL", annotation_position="top right", **line_kwargs)
            if lsl_val is not None:
                fig.add_vline(x=lsl_val, line=dict(color="rgba(220,38,38,0.7)", **dash_style),
                              annotation_text="LSL", annotation_position="top left", **line_kwargs)
            if nom_val is not None:
                fig.add_vline(x=nom_val, line=dict(color="rgba(34,197,94,0.7)", dash="dot", width=1.2),
                              annotation_text="Nom", annotation_position="top", **line_kwargs)

    chart_height = max(400, 300 * n_rows)
    fig.update_layout(
        title=dict(text=f"<b>Histogram: {group_label}</b>", font=dict(size=15),
                   x=0.5, xanchor="center"),
        barmode="overlay",
        height=chart_height,
        margin=dict(l=50, r=120, t=80, b=60),
        legend=dict(title=dict(text=color_by if color_by != "None" else ""),
                    orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
                    font=dict(size=11), bgcolor="rgba(255,255,255,0.8)"),
        template="plotly_white",
    )

    if not use_subplots:
        fig.update_xaxes(title_text="Value")
        fig.update_yaxes(title_text="Count")

    return fig


# ---------------------------------------------------------------------------
# Plotly white-background finalizer
# ---------------------------------------------------------------------------

def finalize_plotly_style(fig):
    """Apply consistent white background and black text to a Plotly figure."""
    fig.update_layout(
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#000000", family="SF Pro Display, SF Pro, -apple-system, sans-serif"),
        title=dict(font=dict(color="#000000")),
        legend=dict(font=dict(color="#000000"), title=dict(font=dict(color="#000000"))),
        xaxis=dict(tickfont=dict(color="#000000"), title=dict(font=dict(color="#000000")), color="#000000"),
        yaxis=dict(tickfont=dict(color="#000000"), title=dict(font=dict(color="#000000")), color="#000000"),
    )
    return fig


# ---------------------------------------------------------------------------
# SPC analytics (pure computation, no Streamlit dependency)
# ---------------------------------------------------------------------------

def calc_process_capability(data_series, usl_val, lsl_val):
    """Calculate Cp, Cpk, Pp, Ppk, sigma level, DPMO, and yield %."""
    data = data_series.dropna()
    if len(data) < 2:
        return None
    mean = data.mean()
    std_within = data.std(ddof=1)
    std_overall = data.std(ddof=0)

    result = {"n": len(data), "mean": round(mean, 6), "std": round(std_within, 6)}

    if usl_val is not None and lsl_val is not None and std_within > 0:
        cp = (usl_val - lsl_val) / (6 * std_within)
        cpu = (usl_val - mean) / (3 * std_within)
        cpl = (mean - lsl_val) / (3 * std_within)
        cpk = min(cpu, cpl)
        pp = (usl_val - lsl_val) / (6 * std_overall) if std_overall > 0 else np.nan
        ppu = (usl_val - mean) / (3 * std_overall) if std_overall > 0 else np.nan
        ppl = (mean - lsl_val) / (3 * std_overall) if std_overall > 0 else np.nan
        ppk = min(ppu, ppl) if std_overall > 0 else np.nan
        result.update({"Cp": round(cp, 4), "Cpk": round(cpk, 4),
                        "Pp": round(pp, 4), "Ppk": round(ppk, 4)})
        sigma_level = cpk * 3
        result["Sigma Level"] = round(sigma_level, 2)
        z_upper = (usl_val - mean) / std_within if std_within > 0 else np.inf
        z_lower = (mean - lsl_val) / std_within if std_within > 0 else np.inf
        p_defect = scipy_stats.norm.sf(z_upper) + scipy_stats.norm.cdf(-z_lower)
        dpmo = p_defect * 1_000_000
        yield_pct = (1 - p_defect) * 100
        result["DPMO"] = int(round(dpmo))
        result["Yield %"] = round(yield_pct, 4)
    elif usl_val is not None and std_within > 0:
        cpu = (usl_val - mean) / (3 * std_within)
        result.update({"Cpk (upper)": round(cpu, 4)})
    elif lsl_val is not None and std_within > 0:
        cpl = (mean - lsl_val) / (3 * std_within)
        result.update({"Cpk (lower)": round(cpl, 4)})

    oos = 0
    if usl_val is not None:
        oos += (data > usl_val).sum()
    if lsl_val is not None:
        oos += (data < lsl_val).sum()
    result["OOS Count"] = int(oos)
    result["OOS %"] = round(oos / len(data) * 100, 2) if len(data) > 0 else 0.0

    return result


def nelson_rules(data_series):
    """Detect Nelson rule violations for trend & shift detection.
    Returns a dict of rule_name -> list of violating indices."""
    data = data_series.dropna().values
    n = len(data)
    if n < 9:
        return {}
    mean = np.mean(data)
    std = np.std(data, ddof=1)
    if std == 0:
        return {}

    violations = {}

    r1 = [i for i in range(n) if abs(data[i] - mean) > 3 * std]
    if r1:
        violations["Rule 1: Beyond 3s"] = r1

    r2 = []
    for i in range(n - 8):
        segment = data[i:i+9]
        if all(s > mean for s in segment) or all(s < mean for s in segment):
            r2.extend(range(i, i+9))
    if r2:
        violations["Rule 2: 9 pts same side"] = sorted(set(r2))

    r3 = []
    for i in range(n - 5):
        seg = data[i:i+6]
        diffs = np.diff(seg)
        if all(d > 0 for d in diffs) or all(d < 0 for d in diffs):
            r3.extend(range(i, i+6))
    if r3:
        violations["Rule 3: 6 pts trend"] = sorted(set(r3))

    r4 = []
    for i in range(n - 13):
        seg = data[i:i+14]
        diffs = np.diff(seg)
        alternating = all(diffs[j] * diffs[j+1] < 0 for j in range(len(diffs)-1))
        if alternating:
            r4.extend(range(i, i+14))
    if r4:
        violations["Rule 4: 14 pts alternating"] = sorted(set(r4))

    r5 = []
    for i in range(n - 2):
        seg = data[i:i+3]
        above = sum(1 for s in seg if s > mean + 2*std)
        below = sum(1 for s in seg if s < mean - 2*std)
        if above >= 2 or below >= 2:
            r5.extend(range(i, i+3))
    if r5:
        violations["Rule 5: 2/3 beyond 2s"] = sorted(set(r5))

    r6 = []
    for i in range(n - 14):
        seg = data[i:i+15]
        if all(abs(s - mean) < std for s in seg):
            r6.extend(range(i, i+15))
    if r6:
        violations["Rule 6: 15 pts within 1s"] = sorted(set(r6))

    return violations


def cusum_analysis(data_series, target=None, h=5.0, k=0.5):
    """CUSUM (Cumulative Sum) analysis for shift detection."""
    data = data_series.dropna().values
    n = len(data)
    if n < 5:
        return None, None, []
    mean = target if target is not None else np.mean(data)
    std = np.std(data, ddof=1)
    if std == 0:
        return None, None, []

    cusum_pos = np.zeros(n)
    cusum_neg = np.zeros(n)
    shift_points = []

    for i in range(n):
        zi = (data[i] - mean) / std
        cusum_pos[i] = max(0, cusum_pos[i-1] + zi - k) if i > 0 else max(0, zi - k)
        cusum_neg[i] = max(0, cusum_neg[i-1] - zi - k) if i > 0 else max(0, -zi - k)
        if cusum_pos[i] > h or cusum_neg[i] > h:
            shift_points.append(i)

    return cusum_pos, cusum_neg, shift_points
