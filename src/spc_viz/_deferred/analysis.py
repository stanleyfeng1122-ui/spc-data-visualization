"""Statistical analysis features deferred from v1.0 (visualization-only release).

Re-enable by importing render_summary_statistics in src/spc_viz/ui/chart_view.py
and calling it after the chart renders. See git history for the original wiring.

Originally lived in shared_ui.py sections 7 and 8 (capability card +
Summary Statistics expander with Process Capability / ANOVA / Trend tabs
which include Nelson Rules, CUSUM, and EWMA charts).
"""

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from scipy import stats as scipy_stats

from spc_viz.charts import (
    calc_process_capability,
    cusum_analysis,
    nelson_rules,
)
from spc_viz.parsers import get_filtered_dim_meta
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
# Capability card
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
# Summary Statistics expander (Capability + ANOVA + Trend/CUSUM/EWMA)
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
