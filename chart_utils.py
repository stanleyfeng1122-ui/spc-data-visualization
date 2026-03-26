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

from spc_parser import get_filtered_dim_meta, DimensionMeta

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


def _find_matching_dim(pf_dims, target_dno):
    """Find a dimension in pf_dims that matches target_dno.

    Handles naming variations like SPC_A vs SPC_A-1 by comparing
    the base name (stripping trailing -N suffixes).
    """
    if target_dno in pf_dims:
        return target_dno
    import re as _re
    target_base = _re.sub(r'-\d+$', '', target_dno)
    for candidate_dno in pf_dims:
        if _re.sub(r'-\d+$', '', candidate_dno) == target_base:
            return candidate_dno
    return None


def prepare_combined_data(parsed_files, dim_nos):
    """
    Combine data from all files for the requested dimensions.
    Returns (df, dim_metas_dict) where df has all rows and a _factory column.

    Handles two cross-file mismatches:
    1. Dimension name variations (SPC_A vs SPC_A-1) via fuzzy matching.
    2. Different column labels for the same dimension (SPC_B_G4 vs SPC_B_P0)
       by renaming each file's local columns to positional canonical names
       (SPC_B__0, SPC_B__1, ...) so rows align by point index.
    """
    frames = []
    dim_metas = OrderedDict()

    # First pass: collect canonical dim_metas (first file that has each dim)
    for pf in parsed_files:
        for dno in dim_nos:
            if dno not in dim_metas:
                match = _find_matching_dim(pf["dimensions"], dno)
                if match:
                    dim_metas[dno] = pf["dimensions"][match]

    # Build positional canonical column names per dimension
    # e.g. SPC_B with 44 points → SPC_B__0 .. SPC_B__43
    canonical_labels = OrderedDict()  # dno -> list of canonical col names
    for dno in dim_nos:
        if dno in dim_metas:
            n_cols = len(dim_metas[dno].col_labels)
            canonical_labels[dno] = [f"{dno}__{i}" for i in range(n_cols)]

    for pf in parsed_files:
        factory = _get_factory(pf)
        df = pf["data"].copy()
        df["_factory"] = factory
        df["_source_file"] = pf["filename"]

        meta_cols = [c for c in pf["meta_columns"] if c in df.columns]
        rename_map = {}
        local_meas_cols = []

        for dno in dim_nos:
            match = _find_matching_dim(pf["dimensions"], dno)
            if match is None:
                continue
            local_meta = pf["dimensions"][match]
            canon = canonical_labels.get(dno)
            if canon is None:
                continue

            # Map local col labels → canonical positional labels
            n = min(len(local_meta.col_labels), len(canon))
            for i in range(n):
                local_col = local_meta.col_labels[i]
                if local_col in df.columns:
                    local_meas_cols.append(local_col)
                    if local_col != canon[i]:
                        rename_map[local_col] = canon[i]

        keep = meta_cols + local_meas_cols + ["_factory", "_source_file"]
        seen = set()
        keep_dedup = []
        for c in keep:
            if c not in seen and c in df.columns:
                seen.add(c)
                keep_dedup.append(c)

        df = df[keep_dedup]
        if rename_map:
            df = df.rename(columns=rename_map)
        frames.append(df)

    if not frames:
        return None, None

    combined = pd.concat(frames, ignore_index=True)

    # Update dim_metas col_labels to use canonical names
    for dno in dim_nos:
        if dno in dim_metas and dno in canonical_labels:
            dim_metas[dno] = DimensionMeta(
                dim_no=dim_metas[dno].dim_no,
                description=dim_metas[dno].description,
                dim_type=dim_metas[dno].dim_type,
                point_numbers=dim_metas[dno].point_numbers,
                nominal=dim_metas[dno].nominal,
                tol_max=dim_metas[dno].tol_max,
                tol_min=dim_metas[dno].tol_min,
                usl=dim_metas[dno].usl,
                lsl=dim_metas[dno].lsl,
                col_indices=dim_metas[dno].col_indices,
                col_labels=canonical_labels[dno],
            )

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
    line_width: float = 2.0,
    highlight_groups: set = None,
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

    # Per-dimension spec limits (for stepping USL/LSL lines)
    # Also keep a single representative for backward compat (annotations, etc.)
    first_dim_info = list(dim_point_info.values())[0]
    usl_rep = next((v for v in first_dim_info[3] if v is not None), None)
    lsl_rep = next((v for v in first_dim_info[4] if v is not None), None)
    nom_rep = next((v for v in first_dim_info[2] if v is not None), None)

    # Check if dimensions have different spec limits
    _all_usls = set()
    _all_lsls = set()
    for dno, (_, _, noms, usls, lsls) in dim_point_info.items():
        for u in usls:
            if u is not None:
                _all_usls.add(round(u, 6))
        for l in lsls:
            if l is not None:
                _all_lsls.add(round(l, 6))
    _has_varying_specs = len(_all_usls) > 1 or len(_all_lsls) > 1

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

                        # Muted style for non-highlighted groups
                        # Grey color (from color_map) is enough — keep same width & opacity
                        _is_muted = (highlight_groups is not None
                                     and grp_name not in highlight_groups)
                        _trace_opacity = 0.35 if _is_muted else 0.45
                        _trace_width = line_width

                        trace = go.Scattergl(
                            x=x_positions,
                            y=y_vals,
                            mode="lines",
                            line=dict(width=_trace_width, color=color),
                            opacity=_trace_opacity,
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

    if _has_varying_specs:
        # Build per-point USL/LSL arrays that step with each dimension
        _spec_x = []
        _usl_y = []
        _lsl_y = []
        _band_usl_y = []
        _band_lsl_y = []
        for sec_label in unique_sections:
            for dno, (col_labels, _, nominals, usls, lsls) in dim_point_info.items():
                x_pos = dim_x_positions[(sec_label, dno)]
                _dim_usl = next((u for u in usls if u is not None), None)
                _dim_lsl = next((l for l in lsls if l is not None), None)
                _dim_nom = next((n for n in nominals if n is not None), None)

                for xi in x_pos:
                    _spec_x.append(xi)
                    if _dim_usl is not None:
                        _usl_y.append((_dim_usl - _dim_nom) if (deviation_mode and _dim_nom is not None) else _dim_usl)
                    else:
                        _usl_y.append(None)
                    if _dim_lsl is not None:
                        _lsl_y.append((_dim_lsl - _dim_nom) if (deviation_mode and _dim_nom is not None) else _dim_lsl)
                    else:
                        _lsl_y.append(None)

        for rk in row_kwargs_list:
            # Filled band between USL and LSL
            _band_usl = [v for v in _usl_y]
            _band_lsl = [v for v in _lsl_y]
            if any(v is not None for v in _band_usl) and any(v is not None for v in _band_lsl):
                fig.add_trace(go.Scatter(
                    x=_spec_x + _spec_x[::-1],
                    y=_band_usl + _band_lsl[::-1],
                    fill="toself",
                    fillcolor="rgba(34, 197, 94, 0.12)",
                    line=dict(width=0),
                    showlegend=False, hoverinfo="skip",
                ), **rk)

            # USL line
            if any(v is not None for v in _usl_y):
                fig.add_trace(go.Scatter(
                    x=_spec_x, y=_usl_y,
                    mode="lines", line=dict(color="rgba(220,38,38,0.5)", **dash_style),
                    showlegend=False,
                    hovertemplate="USL: %{y:.4f}<extra></extra>",
                ), **rk)
            # LSL line
            if any(v is not None for v in _lsl_y):
                fig.add_trace(go.Scatter(
                    x=_spec_x, y=_lsl_y,
                    mode="lines", line=dict(color="rgba(220,38,38,0.5)", **dash_style),
                    showlegend=False,
                    hovertemplate="LSL: %{y:.4f}<extra></extra>",
                ), **rk)
    else:
        # Single spec limit — flat horizontal lines (original behavior)
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
            text=f"<b>{title_text}</b>" + (f"<br><span style='font-size:12px;color:#64748B'>{subtitle}</span>" if subtitle else ""),
            font=dict(size=15), x=0.5, xanchor="center",
        ),
        height=chart_height,
        margin=dict(l=50, r=120, t=120, b=80),
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
    if _has_varying_specs:
        # Show all unique USL/LSL values on y-axis
        _seen_vals = set()
        for dno, (_, _, nominals, usls, lsls) in dim_point_info.items():
            _dim_usl = next((u for u in usls if u is not None), None)
            _dim_lsl = next((l for l in lsls if l is not None), None)
            _dim_nom = next((n for n in nominals if n is not None), None)
            if _dim_usl is not None:
                v = (_dim_usl - _dim_nom) if (deviation_mode and _dim_nom is not None) else _dim_usl
                rv = round(v, 6)
                if rv not in _seen_vals:
                    _seen_vals.add(rv)
                    spec_tickvals.append(v)
                    spec_ticktext.append(f"USL-{v:.4g}")
            if _dim_lsl is not None:
                v = (_dim_lsl - _dim_nom) if (deviation_mode and _dim_nom is not None) else _dim_lsl
                rv = round(v, 6)
                if rv not in _seen_vals:
                    _seen_vals.add(rv)
                    spec_tickvals.append(v)
                    spec_ticktext.append(f"LSL-{v:.4g}")
    else:
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
            header_shapes.append(dict(
                type="rect",
                xref="paper", yref="paper",
                x0=px0, x1=px1,
                y0=1.01, y1=1.07,
                fillcolor="#F1F5F9",
                line=dict(color="#E2E8F0", width=1),
                layer="above",
            ))
        # Merge with existing shapes (USL/LSL lines)
        existing_shapes = list(fig.layout.shapes or [])
        fig.update_layout(shapes=existing_shapes + header_shapes)

        # Add centered section labels
        for cx, sec_label in section_centers:
            annotations.append(dict(
                x=cx, y=1.04,
                xref="paper", yref="paper",
                text=f"<b>{sec_label}</b>",
                showarrow=False,
                xanchor="center",
                yanchor="middle",
                font=dict(size=11, color="#334155"),
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
    section_by_fields: list,
    color_by: str,
    y_axis_mode: str,
    exclude_intervals: bool,
    group_label: str,
    row_by: str = "None",
    custom_color_map: dict = None,
    custom_yrange: list = None,
    selected_points: list = None,
    highlight_groups: set = None,
):
    """Build a box plot with section-by, color-by, row-by, and highlight support."""
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    _point_filter = set(selected_points) if selected_points else None

    # Section labels
    section_labels = compute_sections(df, section_by_fields)
    unique_sections = list(dict.fromkeys(section_labels))

    # Row labels
    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))
    n_rows = len(unique_rows)
    use_row_facets = n_rows > 1

    # Color groups
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
    multi_section = len(unique_sections) > 1
    rep_usl, rep_lsl, rep_nom = None, None, None

    # Build dim point info
    dim_point_info = OrderedDict()
    for dno in dim_nos:
        if dno not in dim_metas:
            continue
        dmeta = dim_metas[dno]
        info = get_filtered_dim_meta(dmeta, exclude_intervals=exclude_intervals)
        col_labels, point_nums, nominals, usls, lsls = info
        valid = [(cl, pn, n, u, l) for cl, pn, n, u, l in
                 zip(col_labels, point_nums, nominals, usls, lsls)
                 if cl in df.columns and (_point_filter is None or pn in _point_filter)]
        if valid:
            dim_point_info[dno] = list(zip(*valid))

    if not dim_point_info:
        return None

    # Build ordered x-axis labels: "Section | Point" or just "Point"
    ordered_x_labels = []
    x_label_to_section = {}
    for sec_label in unique_sections:
        for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
            for pn in point_nums:
                if multi_section:
                    x_lbl = f"{sec_label} | {dno}_{pn}" if multi_dim else f"{sec_label} | {pn}"
                else:
                    x_lbl = f"{dno}_{pn}" if multi_dim else pn
                ordered_x_labels.append(x_lbl)
                x_label_to_section[x_lbl] = sec_label

    for row_idx, row_label in enumerate(unique_rows):
        plotly_row = row_idx + 1 if use_row_facets else None
        row_mask = row_labels == row_label

        for sec_label in unique_sections:
            sec_mask = section_labels == sec_label
            combined_mask = row_mask & sec_mask
            cell_df = df[combined_mask]
            cell_colors = color_series[combined_mask]

            if cell_df.empty:
                continue

            for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
                for col_label, point_num, nominal, usl_val, lsl_val in zip(
                    col_labels, point_nums, nominals, usls, lsls
                ):
                    if multi_section:
                        x_label = f"{sec_label} | {dno}_{point_num}" if multi_dim else f"{sec_label} | {point_num}"
                    else:
                        x_label = f"{dno}_{point_num}" if multi_dim else point_num

                    if rep_usl is None and usl_val is not None:
                        rep_usl, rep_lsl, rep_nom = usl_val, lsl_val, nominal

                    for grp_name in unique_colors:
                        grp_mask = cell_colors == grp_name
                        values = pd.to_numeric(
                            cell_df.loc[grp_mask, col_label], errors="coerce"
                        ).dropna()
                        if deviation_mode and nominal is not None:
                            values = values - nominal
                        if values.empty:
                            continue

                        show_legend = grp_name not in legend_shown
                        legend_shown.add(grp_name)

                        _is_muted = (highlight_groups is not None
                                     and grp_name not in highlight_groups)
                        _opacity = 0.4 if _is_muted else 0.8

                        trace = go.Box(
                            y=values, x=[x_label] * len(values),
                            name=grp_name, legendgroup=grp_name,
                            marker_color=color_map[grp_name],
                            opacity=_opacity,
                            showlegend=show_legend, boxpoints="outliers",
                        )
                        if use_row_facets:
                            fig.add_trace(trace, row=plotly_row, col=1)
                        else:
                            fig.add_trace(trace)

    # Spec limit lines
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

    subtitle = " + ".join(section_by_fields) if section_by_fields else ""
    chart_height = 350 * n_rows if use_row_facets else 620
    y_range_kwargs = dict(range=custom_yrange) if custom_yrange else {}
    fig.update_layout(
        title=dict(
            text=f"<b>Box Plot: {group_label}</b>" + (
                f"<br><span style='font-size:12px;color:#64748B'>{subtitle}</span>" if subtitle else ""
            ),
            font=dict(size=15), x=0.5, xanchor="center",
        ),
        xaxis=dict(
            title="Measurement Point", tickangle=-45,
            tickfont=dict(size=8, color="#000000"),
            categoryorder="array", categoryarray=ordered_x_labels,
        ),
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
