"""Shared UI widgets for SPC pages.

Extracts the duplicated sidebar controls, chart rendering, and summary
statistics that were copy-pasted between app.py and 1_Quick_Test.py.

Every function takes a *key_prefix* so multiple pages can coexist in the
same Streamlit session without widget-key collisions.
"""

from collections import OrderedDict

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from scipy import stats as scipy_stats

from chart_utils import (
    build_box_plot,
    build_combined_chart,
    build_histogram,
    calc_process_capability,
    cusum_analysis,
    finalize_plotly_style,
    get_color_for_group,
    nelson_rules,
    prepare_combined_data,
)
from spc_viz.parsers import detect_dimension_groups, get_filtered_dim_meta
from spc_viz.theme import (
    ACCENT,
    BORDER,
    DANGER,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WARNING,
    WHITE,
)

# ---------------------------------------------------------------------------
# 1. Dimension selector (preset groups + multiselect)
# ---------------------------------------------------------------------------


def build_dimension_selector(all_dimensions, key_prefix=""):
    """Render preset + multiselect in the sidebar, return selected dim numbers.

    Returns (selected_dim_nos, selected_group_label, dim_groups).
    """
    dim_groups = detect_dimension_groups(all_dimensions)

    dim_display_map = OrderedDict()
    for dno, dmeta in all_dimensions.items():
        label = f"{dno} — {dmeta.description}" if dmeta.description else dno
        dim_display_map[label] = dno

    dim_no_to_label = {v: k for k, v in dim_display_map.items()}
    dim_display_labels = list(dim_display_map.keys())

    if not dim_display_labels:
        st.warning("No dimensions found.")
        st.stop()

    st.sidebar.markdown("---")
    group_options = ["Custom"] + list(dim_groups.keys())
    selected_preset = st.sidebar.selectbox(
        "Preset",
        options=group_options,
        index=0,
        key=f"{key_prefix}preset",
    )

    if selected_preset != "Custom":
        preset_dim_nos = dim_groups[selected_preset]
        default_labels = [dim_no_to_label[dno] for dno in preset_dim_nos if dno in dim_no_to_label]
    else:
        default_labels = [dim_display_labels[0]] if dim_display_labels else []

    selected_dim_labels = st.sidebar.multiselect(
        "Dimensions",
        options=dim_display_labels,
        default=default_labels,
        key=f"{key_prefix}dims",
    )
    selected_dim_nos = [dim_display_map[lbl] for lbl in selected_dim_labels]
    selected_group_label = (
        " / ".join(dno.replace("SPC_", "") for dno in selected_dim_nos) if selected_dim_nos else ""
    )

    if not selected_dim_nos:
        st.info("Select at least one dimension from the sidebar.")
        st.stop()

    return selected_dim_nos, selected_group_label, dim_groups


# ---------------------------------------------------------------------------
# 2. Point filter
# ---------------------------------------------------------------------------


def build_point_filter(all_dimensions, selected_dim_nos, key_prefix=""):
    """Render exclude-points controls. Returns (exclude_intervals, selected_points)."""
    exclude_intervals = st.sidebar.checkbox(
        "Exclude interval points",
        value=True,
        key=f"{key_prefix}excl",
    )

    all_point_numbers = []
    for dno in selected_dim_nos:
        if dno in all_dimensions:
            meta = all_dimensions[dno]
            _cls, pns, _noms, _usls, _lsls = get_filtered_dim_meta(meta, exclude_intervals=False)
            for pn in pns:
                if pn and pn not in all_point_numbers:
                    all_point_numbers.append(pn)

    excluded_points = st.sidebar.multiselect(
        "Exclude points",
        options=all_point_numbers,
        default=[],
        help="Pick points to hide. Empty = show all.",
        key=f"{key_prefix}points",
    )
    if excluded_points:
        selected_points = [p for p in all_point_numbers if p not in excluded_points]
        if not selected_points:
            selected_points = None
    else:
        selected_points = None

    return exclude_intervals, selected_points


# ---------------------------------------------------------------------------
# 3. Chart type + grouping + Y-axis controls
# ---------------------------------------------------------------------------

_CHART_LABELS = ["Profile", "Box Plot", "Histogram"]
_CHART_MAP = {"Profile": "Combined Profile", "Box Plot": "Box Plot", "Histogram": "Histogram"}

SECTION_FIELDS = [
    "Factory",
    "Build",
    "Config",
    "Raw material",
    "Vendor Serial Number",
    "Source File",
]


def build_chart_controls(parsed_files, key_prefix=""):
    """Render chart-type, grouping, and Y-axis controls.

    Returns dict with keys: chart_type, color_by, section_by_fields, row_by,
    y_axis_mode, hist_nbins, custom_yrange.
    """
    st.sidebar.markdown("---")
    chart_label = st.sidebar.radio(
        "Chart type",
        options=_CHART_LABELS,
        index=0,
        horizontal=True,
        key=f"{key_prefix}chart_type",
    )
    chart_type = _CHART_MAP[chart_label]

    # Determine available metadata columns
    available_meta = set()
    for pf in parsed_files:
        available_meta.update(pf["meta_columns"])
    available_meta.discard("Start Point")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Grouping")

    meta_list = sorted(available_meta)
    groupby_options = [m for m in meta_list if m not in ("Start Point", "SN")] + ["None"]
    color_by = st.sidebar.selectbox(
        "Color-by",
        options=groupby_options,
        index=len(groupby_options) - 1,
        key=f"{key_prefix}color",
    )

    if chart_type in ("Combined Profile", "Box Plot"):
        section_options = [m for m in meta_list if m not in ("Start Point", "SN")]
        section_options += [s for s in ("Factory", "Source File") if s not in section_options]
        section_by_fields = st.sidebar.multiselect(
            "Section-by",
            options=section_options,
            default=["Factory"] if "Factory" in section_options else [],
            key=f"{key_prefix}section",
        )
    else:
        section_by_fields = []

    rowby_options = [m for m in meta_list if m not in ("Start Point", "SN")] + ["None"]
    row_by = st.sidebar.selectbox(
        "Row-by",
        options=rowby_options,
        index=len(rowby_options) - 1,
        key=f"{key_prefix}row",
    )

    if chart_type in ("Combined Profile", "Box Plot"):
        y_axis_mode = st.sidebar.selectbox(
            "Y-axis",
            options=["Measurement values", "Deviation from Nominal"],
            index=0,
            key=f"{key_prefix}yaxis",
        )
    else:
        y_axis_mode = "Measurement values"

    if chart_type == "Histogram":
        hist_nbins = st.sidebar.slider("Bins", 10, 100, 40, key=f"{key_prefix}bins")
    else:
        hist_nbins = 40

    st.sidebar.markdown("---")
    st.sidebar.subheader("Y-axis Range")
    use_custom = st.sidebar.checkbox("Custom Y range", value=False, key=f"{key_prefix}yr")
    if use_custom:
        y_min = st.sidebar.number_input("Min", value=0.0, format="%.4f", key=f"{key_prefix}ymin")
        y_max = st.sidebar.number_input("Max", value=1.0, format="%.4f", key=f"{key_prefix}ymax")
        custom_yrange = [y_min, y_max] if y_min < y_max else None
    else:
        custom_yrange = None

    return {
        "chart_type": chart_type,
        "color_by": color_by,
        "section_by_fields": section_by_fields,
        "row_by": row_by,
        "y_axis_mode": y_axis_mode,
        "hist_nbins": hist_nbins,
        "custom_yrange": custom_yrange,
    }


# ---------------------------------------------------------------------------
# 4. Color pickers
# ---------------------------------------------------------------------------


def build_color_pickers(df_clean, color_by, key_prefix=""):
    """Render per-group color pickers. Returns custom_color_map dict."""
    st.sidebar.markdown("---")
    st.sidebar.subheader("Colors")
    if color_by != "None" and color_by in df_clean.columns:
        groups = sorted(df_clean[color_by].fillna("Unknown").astype(str).unique())
    else:
        groups = ["All"]

    custom_color_map = {}
    for i, grp in enumerate(groups):
        default_color = get_color_for_group(i)
        custom_color_map[grp] = st.sidebar.color_picker(
            f"{grp}",
            value=default_color,
            key=f"{key_prefix}color_{grp}",
        )
    return custom_color_map


# ---------------------------------------------------------------------------
# 5. Data preparation
# ---------------------------------------------------------------------------


def prepare_and_clean(parsed_files, selected_dim_nos):
    """Combine data and drop all-NaN rows. Returns (df_clean, dim_metas, all_meas_cols)."""
    df, dim_metas = prepare_combined_data(parsed_files, selected_dim_nos)
    if df is None or dim_metas is None or df.empty:
        st.warning("No data found for selected dimensions.")
        st.stop()

    all_meas_cols = []
    for dno in selected_dim_nos:
        if dno in dim_metas:
            all_meas_cols.extend([c for c in dim_metas[dno].col_labels if c in df.columns])

    df_clean = (
        df.dropna(subset=all_meas_cols, how="all").reset_index(drop=True) if all_meas_cols else df
    )
    if df_clean.empty:
        st.warning("No measurement data for selected dimensions.")
        st.stop()

    return df_clean, dim_metas, all_meas_cols


# ---------------------------------------------------------------------------
# 6. Chart building + rendering
# ---------------------------------------------------------------------------


def _build_chart_figure(
    df_clean,
    dim_metas,
    selected_dim_nos,
    controls,
    custom_color_map,
    exclude_intervals,
    selected_group_label,
    selected_points,
):
    """Build a Plotly Figure from controls without rendering it.

    Returns the finalised Figure, or None when data is insufficient.
    """
    ct = controls["chart_type"]
    common = dict(
        df=df_clean,
        dim_metas=dim_metas,
        dim_nos=selected_dim_nos,
        color_by=controls["color_by"],
        exclude_intervals=exclude_intervals,
        group_label=selected_group_label,
        row_by=controls["row_by"],
        custom_color_map=custom_color_map,
        selected_points=selected_points,
    )

    if ct == "Combined Profile":
        fig = build_combined_chart(
            **common,
            section_by_fields=controls["section_by_fields"],
            y_axis_mode=controls["y_axis_mode"],
            custom_yrange=controls["custom_yrange"],
        )
    elif ct == "Box Plot":
        fig = build_box_plot(
            **common,
            y_axis_mode=controls["y_axis_mode"],
            custom_yrange=controls["custom_yrange"],
        )
    elif ct == "Histogram":
        fig = build_histogram(
            **common,
            nbins=controls["hist_nbins"],
        )
    else:
        fig = None

    if fig is not None:
        finalize_plotly_style(fig)
    return fig


def build_and_render_chart(
    df_clean,
    dim_metas,
    selected_dim_nos,
    controls,
    custom_color_map,
    exclude_intervals,
    selected_group_label,
    selected_points,
    key_prefix="",
):
    """Build the Plotly figure based on controls dict and render it."""
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


# ---------------------------------------------------------------------------
# 7. Capability card
# ---------------------------------------------------------------------------


def render_capability_card(cap):
    """Render process capability metrics in a dense grid."""
    cpk = cap.get("Cpk", cap.get("Cpk (upper)", cap.get("Cpk (lower)", None)))
    if cpk is not None:
        if cpk >= 1.67:
            rating, color = "EXCELLENT", SUCCESS
        elif cpk >= 1.33:
            rating, color = "GOOD", ACCENT
        elif cpk >= 1.0:
            rating, color = "MARGINAL", WARNING
        else:
            rating, color = "POOR", DANGER
    else:
        rating, color = "N/A", TEXT_MUTED

    cols = st.columns(5)
    if "Cp" in cap:
        cols[0].metric("Cp", cap["Cp"])
    if "Cpk" in cap:
        cols[1].metric("Cpk", cap["Cpk"])
    elif "Cpk (upper)" in cap:
        cols[1].metric("Cpk (upper)", cap["Cpk (upper)"])
    elif "Cpk (lower)" in cap:
        cols[1].metric("Cpk (lower)", cap["Cpk (lower)"])
    if "Pp" in cap:
        cols[2].metric("Pp", cap["Pp"])
    if "Ppk" in cap:
        cols[3].metric("Ppk", cap["Ppk"])
    cols[4].markdown(
        f"<div style='text-align:center;padding:4px;'>"
        f"<span style='font-size:0.65rem;text-transform:uppercase;letter-spacing:0.05em;"
        f"color:{TEXT_MUTED};font-family:IBM Plex Sans,sans-serif;'>Rating</span><br>"
        f"<span style='font-size:1.1rem;font-weight:700;color:{color};"
        f"font-family:JetBrains Mono,monospace;'>{rating}</span></div>",
        unsafe_allow_html=True,
    )

    cols2 = st.columns(4)
    if "Sigma Level" in cap:
        cols2[0].metric("Sigma", f"{cap['Sigma Level']}σ")
    if "DPMO" in cap:
        cols2[1].metric("DPMO", f"{cap['DPMO']:,}")
    if "Yield %" in cap:
        cols2[2].metric("Yield", f"{cap['Yield %']}%")
    if cap.get("OOS Count", 0) > 0:
        cols2[3].markdown(
            f"<div style='padding:4px;font-size:0.82rem;color:{DANGER};'>"
            f"OOS: {cap['OOS Count']} ({cap['OOS %']}%)</div>",
            unsafe_allow_html=True,
        )
    else:
        cols2[3].markdown(
            f"<div style='padding:4px;font-size:0.82rem;color:{SUCCESS};'>"
            f"0 OOS / {cap['n']} pts</div>",
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# 8. Summary Statistics expander (Capability + ANOVA + Trend/CUSUM/EWMA)
# ---------------------------------------------------------------------------


def render_summary_statistics(
    df_clean,
    dim_metas,
    selected_dim_nos,
    exclude_intervals,
    color_by,
    custom_color_map,
    key_prefix="",
):
    """Render the full Summary Statistics expander with 3 tabs per dimension."""
    with st.expander("Summary Statistics", expanded=True):
        for dno in selected_dim_nos:
            if dno not in dim_metas:
                continue
            dmeta = dim_metas[dno]
            col_labels, point_nums, nominals, usls, lsls = get_filtered_dim_meta(
                dmeta, exclude_intervals=exclude_intervals
            )
            valid_cols = [c for c in col_labels if c in df_clean.columns]
            if not valid_cols:
                continue

            st.markdown(
                f"<div style='border-bottom:1px solid {BORDER};padding:4px 0 2px;margin-top:8px;'>"
                f"<span style='font-family:Barlow Condensed,sans-serif;font-size:0.95rem;"
                f"font-weight:600;color:{TEXT_PRIMARY};'>{dno}</span>"
                f"<span style='font-size:0.78rem;color:{TEXT_MUTED};margin-left:8px;'>"
                f"{dmeta.description}</span></div>",
                unsafe_allow_html=True,
            )

            usl_val = next((v for v in usls if v is not None), None)
            nom_val = next((v for v in nominals if v is not None), None)
            lsl_val = next((v for v in lsls if v is not None), None)

            spec_cols = st.columns(3)
            spec_cols[0].metric("USL", f"{usl_val:.4f}" if usl_val is not None else "N/A")
            spec_cols[1].metric("Nominal", f"{nom_val:.4f}" if nom_val is not None else "N/A")
            spec_cols[2].metric("LSL", f"{lsl_val:.4f}" if lsl_val is not None else "N/A")

            all_values = df_clean[valid_cols].values.flatten()
            all_values = pd.Series(all_values).dropna()

            tab_cap, tab_anova, tab_trend = st.tabs(
                ["Process Capability", "ANOVA", "Trend / Shift"]
            )

            # --- Process Capability ---
            with tab_cap:
                _render_tab_capability(
                    df_clean, valid_cols, point_nums, usls, lsls, usl_val, lsl_val, all_values
                )

            # --- ANOVA ---
            with tab_anova:
                _render_tab_anova(
                    df_clean,
                    valid_cols,
                    color_by,
                    all_values,
                    usl_val,
                    lsl_val,
                    custom_color_map,
                    key_prefix=f"{key_prefix}anova_{dno}",
                )

            # --- Trend / Shift ---
            with tab_trend:
                _render_tab_trend(all_values, nom_val, key_prefix=f"{key_prefix}trend_{dno}")

            st.markdown(
                f"<hr style='border:none;border-top:1px solid {BORDER};margin:8px 0;'>",
                unsafe_allow_html=True,
            )


# -- Private tab helpers ----------------------------------------------------


def _render_tab_capability(
    df_clean, valid_cols, point_nums, usls, lsls, usl_val, lsl_val, all_values
):
    if len(all_values) < 2:
        st.info("Not enough data for process capability.")
        return
    cap = calc_process_capability(all_values, usl_val, lsl_val)
    if not cap:
        return
    render_capability_card(cap)

    if len(valid_cols) > 1:
        st.markdown(
            f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
            f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:12px 0 4px;'>"
            f"Per-Point Breakdown</div>",
            unsafe_allow_html=True,
        )
        rows = []
        for ci, col in enumerate(valid_cols):
            col_usl = usls[ci] if ci < len(usls) and usls[ci] is not None else usl_val
            col_lsl = lsls[ci] if ci < len(lsls) and lsls[ci] is not None else lsl_val
            pc = calc_process_capability(df_clean[col], col_usl, col_lsl)
            if pc:
                pt_label = point_nums[ci] if ci < len(point_nums) else col
                rows.append(
                    {
                        "Point": pt_label,
                        **{
                            k: v
                            for k, v in pc.items()
                            if k
                            in [
                                "mean",
                                "std",
                                "Cp",
                                "Cpk",
                                "Pp",
                                "Ppk",
                                "Sigma Level",
                                "DPMO",
                                "Yield %",
                                "OOS Count",
                            ]
                        },
                    }
                )
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def _render_tab_anova(
    df_clean, valid_cols, color_by, all_values, usl_val, lsl_val, custom_color_map, key_prefix=""
):
    if color_by == "None" or color_by not in df_clean.columns:
        st.info("Select a Color-by grouping for group comparison.")
        return

    groups = df_clean[color_by].fillna("Unknown").astype(str)
    unique_groups = sorted(groups.unique())
    if len(unique_groups) < 2:
        st.info("Need 2+ groups for ANOVA.")
        return

    group_data = {}
    for g in unique_groups:
        mask = groups == g
        vals = df_clean.loc[mask, valid_cols].values.flatten()
        vals = pd.Series(vals).dropna()
        if len(vals) > 0:
            group_data[g] = vals

    if len(group_data) < 2:
        st.info("Not enough data in groups.")
        return

    f_stat, p_value = scipy_stats.f_oneway(*group_data.values())

    anova_cols = st.columns(3)
    anova_cols[0].metric("F-statistic", f"{f_stat:.4f}")
    anova_cols[1].metric("p-value", f"{p_value:.6f}")
    sig = "YES" if p_value < 0.05 else "NO"
    sig_color = DANGER if p_value < 0.05 else SUCCESS
    anova_cols[2].markdown(
        f"<div style='text-align:center;padding:4px;'>"
        f"<span style='font-size:0.65rem;text-transform:uppercase;"
        f"letter-spacing:0.05em;color:{TEXT_MUTED};'>Significant</span><br>"
        f"<span style='font-size:1.1rem;font-weight:700;color:{sig_color};"
        f"font-family:JetBrains Mono,monospace;'>{sig}</span></div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
        f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:8px 0 4px;'>"
        f"Group Summary</div>",
        unsafe_allow_html=True,
    )
    summary_rows = []
    for g, vals in group_data.items():
        summary_rows.append(
            {
                "Group": g,
                "n": len(vals),
                "Mean": round(vals.mean(), 6),
                "Std": round(vals.std(ddof=1), 6),
                "Min": round(vals.min(), 6),
                "Max": round(vals.max(), 6),
                "Range": round(vals.max() - vals.min(), 6),
            }
        )
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    grand_mean = all_values.mean()
    ss_between = sum(
        len(group_data[g]) * (group_data[g].mean() - grand_mean) ** 2 for g in group_data
    )
    ss_within = sum(((group_data[g] - group_data[g].mean()) ** 2).sum() for g in group_data)
    ss_total = ss_between + ss_within
    if ss_total > 0:
        var_cols = st.columns(3)
        var_cols[0].metric("SS Between", f"{ss_between:.4f}")
        var_cols[1].metric("SS Within", f"{ss_within:.4f}")
        var_cols[2].metric("% Between", f"{ss_between / ss_total * 100:.1f}%")

    # Box plot per group
    fig_box = go.Figure()
    for g in unique_groups:
        if g in group_data:
            fig_box.add_trace(
                go.Box(
                    y=group_data[g].values,
                    name=g,
                    marker_color=custom_color_map.get(g, ACCENT),
                    boxmean="sd",
                )
            )
    fig_box.update_layout(
        yaxis_title="Value",
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
        height=300,
        margin=dict(l=40, r=20, t=30, b=40),
        xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
        yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
    )
    if usl_val is not None:
        fig_box.add_hline(
            y=usl_val, line_dash="dash", line_color=DANGER, annotation_text=f"USL {usl_val:.4g}"
        )
    if lsl_val is not None:
        fig_box.add_hline(
            y=lsl_val, line_dash="dash", line_color=DANGER, annotation_text=f"LSL {lsl_val:.4g}"
        )
    st.plotly_chart(fig_box, use_container_width=True, key=f"{key_prefix}_box")


def _render_tab_trend(all_values, nom_val, key_prefix=""):
    if len(all_values) < 9:
        st.info("Need 9+ data points for trend/shift analysis.")
        return

    # Nelson Rules
    st.markdown(
        f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
        f"letter-spacing:0.06em;color:{TEXT_MUTED};margin-bottom:4px;'>"
        f"Nelson Rules</div>",
        unsafe_allow_html=True,
    )
    violations = nelson_rules(all_values)
    if not violations:
        st.markdown(
            f"<span style='font-size:0.82rem;color:{SUCCESS};'>No violations — process stable</span>",
            unsafe_allow_html=True,
        )
    else:
        for rule_name, indices in violations.items():
            st.markdown(
                f"<span style='font-size:0.82rem;color:{WARNING};'>"
                f"{rule_name}: {len(indices)} pts</span>",
                unsafe_allow_html=True,
            )
        viol_rows = []
        for rule_name, indices in violations.items():
            viol_rows.append(
                {
                    "Rule": rule_name,
                    "Violations": len(indices),
                    "Indices": str(indices[:20]) + ("..." if len(indices) > 20 else ""),
                }
            )
        st.dataframe(pd.DataFrame(viol_rows), use_container_width=True, hide_index=True)

    # CUSUM
    st.markdown(
        f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
        f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:12px 0 4px;'>"
        f"CUSUM Chart</div>",
        unsafe_allow_html=True,
    )
    cusum_pos, cusum_neg, shift_pts = cusum_analysis(all_values, target=nom_val)
    if cusum_pos is not None:
        fig_cusum = go.Figure()
        x_idx = list(range(len(cusum_pos)))
        fig_cusum.add_trace(
            go.Scatter(
                x=x_idx,
                y=cusum_pos,
                mode="lines",
                name="CUSUM+",
                line=dict(color=ACCENT, width=1.5),
            )
        )
        fig_cusum.add_trace(
            go.Scatter(
                x=x_idx,
                y=cusum_neg,
                mode="lines",
                name="CUSUM−",
                line=dict(color=DANGER, width=1.5),
            )
        )
        fig_cusum.add_hline(y=5.0, line_dash="dash", line_color=TEXT_MUTED, annotation_text="h=5")
        if shift_pts:
            fig_cusum.add_trace(
                go.Scatter(
                    x=shift_pts,
                    y=[max(cusum_pos[i], cusum_neg[i]) for i in shift_pts],
                    mode="markers",
                    name="Shift",
                    marker=dict(color=DANGER, size=6, symbol="x"),
                )
            )
        fig_cusum.update_layout(
            xaxis_title="Observation",
            yaxis_title="Cumulative Sum",
            paper_bgcolor=WHITE,
            plot_bgcolor=WHITE,
            font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
            height=250,
            margin=dict(l=40, r=20, t=20, b=40),
            xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
            yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
        )
        st.plotly_chart(fig_cusum, use_container_width=True, key=f"{key_prefix}_cusum")

        if shift_pts:
            st.markdown(
                f"<span style='font-size:0.82rem;color:{WARNING};'>"
                f"CUSUM: {len(shift_pts)} potential shifts</span>",
                unsafe_allow_html=True,
            )

    # EWMA
    st.markdown(
        f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
        f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:12px 0 4px;'>"
        f"EWMA Chart</div>",
        unsafe_allow_html=True,
    )
    lam = 0.2
    ewma = np.zeros(len(all_values))
    ewma[0] = all_values.iloc[0]
    for i in range(1, len(all_values)):
        ewma[i] = lam * all_values.iloc[i] + (1 - lam) * ewma[i - 1]
    overall_mean = all_values.mean()
    overall_std = all_values.std(ddof=1)
    ewma_ucl = np.array(
        [
            overall_mean
            + 3 * overall_std * np.sqrt(lam / (2 - lam) * (1 - (1 - lam) ** (2 * (i + 1))))
            for i in range(len(all_values))
        ]
    )
    ewma_lcl = np.array(
        [
            overall_mean
            - 3 * overall_std * np.sqrt(lam / (2 - lam) * (1 - (1 - lam) ** (2 * (i + 1))))
            for i in range(len(all_values))
        ]
    )

    fig_ewma = go.Figure()
    x_idx = list(range(len(ewma)))
    fig_ewma.add_trace(
        go.Scatter(x=x_idx, y=ewma, mode="lines", name="EWMA", line=dict(color=ACCENT, width=2))
    )
    fig_ewma.add_trace(
        go.Scatter(
            x=x_idx,
            y=ewma_ucl,
            mode="lines",
            name="UCL",
            line=dict(color=DANGER, dash="dash", width=1),
        )
    )
    fig_ewma.add_trace(
        go.Scatter(
            x=x_idx,
            y=ewma_lcl,
            mode="lines",
            name="LCL",
            line=dict(color=DANGER, dash="dash", width=1),
        )
    )
    fig_ewma.add_hline(
        y=overall_mean, line_dash="dot", line_color=TEXT_MUTED, annotation_text="Center"
    )
    ooc_ewma = [i for i in range(len(ewma)) if ewma[i] > ewma_ucl[i] or ewma[i] < ewma_lcl[i]]
    if ooc_ewma:
        fig_ewma.add_trace(
            go.Scatter(
                x=ooc_ewma,
                y=[ewma[i] for i in ooc_ewma],
                mode="markers",
                name="OOC",
                marker=dict(color=DANGER, size=6, symbol="x"),
            )
        )
    fig_ewma.update_layout(
        xaxis_title="Observation",
        yaxis_title="EWMA",
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
        height=250,
        margin=dict(l=40, r=20, t=20, b=40),
        xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
        yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
    )
    st.plotly_chart(fig_ewma, use_container_width=True, key=f"{key_prefix}_ewma")

    if ooc_ewma:
        st.markdown(
            f"<span style='font-size:0.82rem;color:{WARNING};'>"
            f"EWMA: {len(ooc_ewma)} out-of-control</span>",
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# 9. Batch chart export
# ---------------------------------------------------------------------------


def render_batch_export(
    all_dimensions: OrderedDict,
    parsed_files: list,
    controls: dict,
    exclude_intervals: bool,
    selected_points: list | None,
    custom_color_map: dict,
    key_prefix: str = "",
) -> None:
    """Sidebar expander that batch-exports one chart per selected dimension."""
    import io
    import re
    import zipfile

    from streamlit.runtime.scriptrunner import StopException

    dim_display_map = OrderedDict()
    for dno, dmeta in all_dimensions.items():
        label = f"{dno} — {dmeta.description}" if dmeta.description else dno
        dim_display_map[label] = dno

    with st.sidebar.expander("Batch Chart Export", expanded=False):
        batch_dims = st.multiselect(
            "Dimensions to export",
            options=list(dim_display_map.keys()),
            default=[],
            help="Select dimensions. One chart per dimension.",
            key=f"{key_prefix}batch_dims",
        )
        if not batch_dims:
            st.caption("Pick dimensions above, then click Export.")
            return

        export_btn = st.button(
            f"Export {len(batch_dims)} chart{'s' if len(batch_dims) != 1 else ''}",
            key=f"{key_prefix}batch_export_btn",
        )

        if not export_btn:
            return

        # --- Generate charts ---
        ct = controls["chart_type"]
        progress = st.progress(0, text="Preparing export…")
        images: list[tuple[str, bytes]] = []
        skipped: list[str] = []

        for idx, label in enumerate(batch_dims):
            dno = dim_display_map[label]
            progress.progress(
                (idx) / len(batch_dims),
                text=f"Generating {dno} ({idx + 1}/{len(batch_dims)})…",
            )

            # Build data for this single dimension
            try:
                df_clean, dim_metas, _ = prepare_and_clean(parsed_files, [dno])
            except (StopException, Exception):
                skipped.append(dno)
                continue

            if df_clean is None or df_clean.empty:
                skipped.append(dno)
                continue

            desc = all_dimensions[dno].description or ""
            group_label = f"{dno.replace('SPC_', '')} — {desc}" if desc else dno

            fig = _build_chart_figure(
                df_clean,
                dim_metas,
                [dno],
                controls,
                custom_color_map,
                exclude_intervals,
                group_label,
                selected_points,
            )
            if fig is None:
                skipped.append(dno)
                continue

            # Convert to PNG
            try:
                png_bytes = fig.to_image(
                    format="png",
                    width=1400,
                    height=700,
                    scale=2,
                )
            except Exception as e:
                skipped.append(f"{dno} (image error: {e})")
                continue

            safe_desc = re.sub(r"[^\w\s-]", "", desc).strip().replace(" ", "_")
            fname = (
                f"{ct.replace(' ', '_')}_{dno}_{safe_desc}.png"
                if safe_desc
                else f"{ct.replace(' ', '_')}_{dno}.png"
            )
            images.append((fname, png_bytes))

        progress.progress(1.0, text="Done!")

        if skipped:
            st.warning(f"Skipped {len(skipped)} dimension(s): {', '.join(skipped)}")

        if not images:
            st.error("No charts could be generated.")
            return

        # Bundle into ZIP
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, png_data in images:
                zf.writestr(fname, png_data)
        buf.seek(0)

        st.download_button(
            label=f"Download {len(images)} chart{'s' if len(images) != 1 else ''} (ZIP)",
            data=buf.getvalue(),
            file_name=f"SPC_Charts_{ct.replace(' ', '_')}.zip",
            mime="application/zip",
            key=f"{key_prefix}batch_download",
        )
