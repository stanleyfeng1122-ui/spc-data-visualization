"""
Quick Test Page — Auto-loads local .xlsx files for fast iteration.

No file uploading needed. All Excel files in the project directory are
parsed automatically so you can immediately verify chart and analysis
behaviour after code changes.
"""

import os
import sys
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from collections import OrderedDict
from scipy import stats as scipy_stats

# Ensure project root is on the path so we can import siblings
_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from spc_parser import (
    parse_excel,
    parse_excel_multi,
    detect_dimension_groups,
    get_filtered_dim_meta,
)
from chart_utils import (
    COLOR_PALETTE,
    get_color_for_group,
    prepare_combined_data,
    build_combined_chart,
    build_box_plot,
    build_histogram,
    finalize_plotly_style,
    calc_process_capability,
    nelson_rules,
    cusum_analysis,
)
from ui_theme import inject_theme, FONT_MONO, TEXT_PRIMARY, TEXT_SECONDARY, TEXT_MUTED, ACCENT, DANGER, SUCCESS, WARNING, WHITE, BG_SUBTLE, BORDER

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Quick Test — SPC",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_theme()


# ---------------------------------------------------------------------------
# Auto-discover and parse local .xlsx files
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Parsing local Excel files...")
def load_local_files(data_dir: str, sheet: str):
    """Scan data_dir for .xlsx files (skip temp ~$ files) and parse them."""
    results = []
    xlsx_files = sorted([
        f for f in os.listdir(data_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ])
    for fname in xlsx_files:
        fpath = os.path.join(data_dir, fname)
        try:
            parsed_list = parse_excel_multi(fpath, sheet_name=sheet)
            for parsed in parsed_list:
                results.append({
                    "filename": parsed.filename,
                    "sheet_name": parsed.sheet_name,
                    "part_number": parsed.part_number,
                    "part_description": parsed.part_description,
                    "revision": parsed.revision,
                    "factory": parsed.factory,
                    "dimensions": parsed.dimensions,
                    "data": parsed.data,
                    "meta_columns": parsed.meta_columns,
                })
        except Exception as e:
            st.sidebar.error(f"Error parsing {fname}: {e}")
    return results


def _discover_sheets(data_dir: str):
    """Read sheet names from all .xlsx files in the directory."""
    import openpyxl
    all_sheets = []
    _NON_DATA_PREFIXES = ("BoxPlotCht", "Histo ")
    _NON_DATA_EXACT = {"Histo Pivot", "Histo Listbox", "Histo Curve"}
    xlsx_files = sorted([
        f for f in os.listdir(data_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ])
    for fname in xlsx_files:
        fpath = os.path.join(data_dir, fname)
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True, keep_links=False)
            for sn in wb.sheetnames:
                if sn not in all_sheets and sn not in _NON_DATA_EXACT and not any(sn.startswith(p) for p in _NON_DATA_PREFIXES):
                    all_sheets.append(sn)
            wb.close()
        except Exception:
            pass
    return all_sheets


# ---------------------------------------------------------------------------
# SIDEBAR — dense control panel
# ---------------------------------------------------------------------------
st.sidebar.title("Quick Test")
st.sidebar.caption("Auto-loads .xlsx from project directory")

_available_sheets = _discover_sheets(_project_root)
_sheet_options = ["Auto-detect"] + _available_sheets

sheet_choice = st.sidebar.selectbox(
    "Sheet",
    options=_sheet_options,
    index=0,
    key="qt_sheet",
    help="Auto-detect scans all sheets. Or pick one.",
)
sheet_name = "Raw data" if sheet_choice == "Auto-detect" else sheet_choice

parsed_files = load_local_files(_project_root, sheet_name)

if not parsed_files:
    st.title("Quick Test")
    st.warning("No .xlsx files found or none could be parsed.")
    st.stop()

# Loaded files summary
st.sidebar.markdown("---")
with st.sidebar.expander(f"Files ({len(parsed_files)})", expanded=False):
    for pf in parsed_files:
        n_rows = len(pf["data"]) if pf["data"] is not None else 0
        builds = ""
        if pf["data"] is not None and "Build" in pf["data"].columns:
            builds = ", ".join(sorted(pf["data"]["Build"].dropna().unique().astype(str)))
        factory = pf.get("factory", "?")
        st.markdown(
            f"`{pf['filename']}`  \n"
            f"<span style='font-size:0.72rem;color:{TEXT_MUTED}'>"
            f"{factory} / {pf['part_number']} / {n_rows} rows"
            f"</span>",
            unsafe_allow_html=True,
        )

# ---------------------------------------------------------------------------
# Build unified dimension map
# ---------------------------------------------------------------------------
all_dimensions = OrderedDict()
for pf in parsed_files:
    for dno, dmeta in pf["dimensions"].items():
        if dno not in all_dimensions:
            all_dimensions[dno] = dmeta

dim_groups = detect_dimension_groups(all_dimensions)

dim_display_map = OrderedDict()
for dno, dmeta in all_dimensions.items():
    label = f"{dno} — {dmeta.description}" if dmeta.description else dno
    dim_display_map[label] = dno
dim_no_to_label = {v: k for k, v in dim_display_map.items()}
dim_display_labels = list(dim_display_map.keys())

if not dim_display_labels:
    st.warning("No dimensions found in loaded files.")
    st.stop()

# Dimension selector
st.sidebar.markdown("---")
group_options = ["Custom"] + list(dim_groups.keys())
selected_preset = st.sidebar.selectbox("Preset", options=group_options, index=0, key="qt_preset")

if selected_preset != "Custom":
    preset_dim_nos = dim_groups[selected_preset]
    default_labels = [dim_no_to_label[dno] for dno in preset_dim_nos if dno in dim_no_to_label]
else:
    default_labels = [dim_display_labels[0]] if dim_display_labels else []

selected_dim_labels = st.sidebar.multiselect(
    "Dimensions", options=dim_display_labels, default=default_labels, key="qt_dims",
)
selected_dim_nos = [dim_display_map[lbl] for lbl in selected_dim_labels]
selected_group_label = " / ".join(dno.replace("SPC_", "") for dno in selected_dim_nos) if selected_dim_nos else ""

if not selected_dim_nos:
    st.info("Select at least one dimension from the sidebar.")
    st.stop()

# Options
exclude_intervals = st.sidebar.checkbox("Exclude interval points", value=True, key="qt_excl")

# Points selector — empty = show all
_all_point_numbers = []
for _dno in selected_dim_nos:
    if _dno in all_dimensions:
        _meta = all_dimensions[_dno]
        _cls, _pns, _noms, _usls, _lsls = get_filtered_dim_meta(_meta, exclude_intervals=False)
        for _pn in _pns:
            if _pn and _pn not in _all_point_numbers:
                _all_point_numbers.append(_pn)

excluded_points = st.sidebar.multiselect(
    "Exclude points",
    options=_all_point_numbers,
    default=[],
    help="Pick points to hide. Empty = show all.",
    key="qt_points",
)
if excluded_points:
    selected_points = [p for p in _all_point_numbers if p not in excluded_points]
    if not selected_points:
        selected_points = None
else:
    selected_points = None

st.sidebar.markdown("---")
_CHART_LABELS = ["Profile", "Box Plot", "Histogram"]
_CHART_MAP = {"Profile": "Combined Profile", "Box Plot": "Box Plot", "Histogram": "Histogram"}
_chart_label = st.sidebar.radio(
    "Chart type",
    options=_CHART_LABELS,
    index=0,
    horizontal=True,
    key="qt_chart_type",
)
chart_type = _CHART_MAP[_chart_label]

# Grouping
st.sidebar.markdown("---")
st.sidebar.subheader("Grouping")

available_meta = set()
for pf in parsed_files:
    available_meta.update(pf["meta_columns"])
available_meta.discard("Start Point")

GROUPBY_OPTIONS = ["Raw material", "Build", "Vendor Serial Number", "Config", "None"]
groupby_options_filtered = [g for g in GROUPBY_OPTIONS if g in available_meta or g == "None"]
color_by = st.sidebar.selectbox("Color-by", options=groupby_options_filtered, index=0, key="qt_color")

if chart_type in ("Combined Profile", "Box Plot"):
    SECTION_FIELDS = ["Factory", "Build", "Config", "Raw material", "Vendor Serial Number", "Source File"]
    section_fields_available = [s for s in SECTION_FIELDS if s in available_meta or s in ("Factory", "Source File")]
    section_by_fields = st.sidebar.multiselect(
        "Section-by", options=section_fields_available,
        default=["Factory"] if "Factory" in section_fields_available else [],
        key="qt_section",
    )
else:
    section_by_fields = []

ROWBY_OPTIONS = ["Raw material", "Build", "Factory", "Config", "None"]
rowby_options_filtered = [r for r in ROWBY_OPTIONS if r in available_meta or r == "None"]
row_by = st.sidebar.selectbox("Row-by", options=rowby_options_filtered,
                               index=len(rowby_options_filtered) - 1, key="qt_row")

if chart_type in ("Combined Profile", "Box Plot"):
    y_axis_mode = st.sidebar.selectbox("Y-axis", options=["Measurement values", "Deviation from Nominal"],
                                        index=0, key="qt_yaxis")
else:
    y_axis_mode = "Measurement values"

if chart_type == "Histogram":
    hist_nbins = st.sidebar.slider("Bins", 10, 100, 40, key="qt_bins")
else:
    hist_nbins = 40

st.sidebar.markdown("---")
st.sidebar.subheader("Y-axis Range")
use_custom_yrange = st.sidebar.checkbox("Custom Y range", value=False, key="qt_yr")
if use_custom_yrange:
    y_min = st.sidebar.number_input("Min", value=0.0, format="%.4f", key="qt_ymin")
    y_max = st.sidebar.number_input("Max", value=1.0, format="%.4f", key="qt_ymax")
    custom_yrange = [y_min, y_max] if y_min < y_max else None
else:
    custom_yrange = None


# ---------------------------------------------------------------------------
# MAIN AREA — chart + analysis
# ---------------------------------------------------------------------------

# Header row: title left, spec info right
hdr_left, hdr_right = st.columns([3, 1])
with hdr_left:
    st.markdown(
        f"<h1 style='margin:0;padding:0;font-size:1.3rem;'>{selected_group_label or 'SPC Analysis'}</h1>"
        f"<span style='font-size:0.75rem;color:{TEXT_MUTED};font-family:{FONT_MONO};'>"
        f"{len(parsed_files)} file{'s' if len(parsed_files) != 1 else ''} / "
        f"{chart_type} / {color_by}"
        f"</span>",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Prepare data
# ---------------------------------------------------------------------------
df, dim_metas = prepare_combined_data(parsed_files, selected_dim_nos)

if df is None or dim_metas is None or df.empty:
    st.warning("No data found for selected dimensions.")
    st.stop()

all_meas_cols = []
for dno in selected_dim_nos:
    if dno in dim_metas:
        all_meas_cols.extend([c for c in dim_metas[dno].col_labels if c in df.columns])

df_clean = df.dropna(subset=all_meas_cols, how="all").reset_index(drop=True) if all_meas_cols else df

if df_clean.empty:
    st.warning("No measurement data for selected dimensions.")
    st.stop()

# Color pickers
st.sidebar.markdown("---")
st.sidebar.subheader("Colors")
if color_by != "None" and color_by in df_clean.columns:
    _color_groups = sorted(df_clean[color_by].fillna("Unknown").astype(str).unique())
else:
    _color_groups = ["All"]

custom_color_map = {}
for i, grp in enumerate(_color_groups):
    default_color = get_color_for_group(i)
    custom_color_map[grp] = st.sidebar.color_picker(f"{grp}", value=default_color, key=f"qt_color_{grp}")

# ---------------------------------------------------------------------------
# Build chart
# ---------------------------------------------------------------------------
if chart_type == "Combined Profile":
    fig = build_combined_chart(
        df=df_clean, dim_metas=dim_metas, dim_nos=selected_dim_nos,
        section_by_fields=section_by_fields, color_by=color_by,
        y_axis_mode=y_axis_mode, exclude_intervals=exclude_intervals,
        group_label=selected_group_label, row_by=row_by,
        custom_color_map=custom_color_map, custom_yrange=custom_yrange,
        selected_points=selected_points,
    )
elif chart_type == "Box Plot":
    fig = build_box_plot(
        df=df_clean, dim_metas=dim_metas, dim_nos=selected_dim_nos,
        color_by=color_by, y_axis_mode=y_axis_mode,
        exclude_intervals=exclude_intervals, group_label=selected_group_label,
        row_by=row_by, custom_color_map=custom_color_map, custom_yrange=custom_yrange,
        selected_points=selected_points,
    )
elif chart_type == "Histogram":
    fig = build_histogram(
        df=df_clean, dim_metas=dim_metas, dim_nos=selected_dim_nos,
        color_by=color_by, exclude_intervals=exclude_intervals,
        group_label=selected_group_label, nbins=hist_nbins,
        row_by=row_by, custom_color_map=custom_color_map,
        selected_points=selected_points,
    )

if fig is None:
    st.warning("Could not generate chart. Check dimensions have data.")
    st.stop()

# Apply theme-aligned Plotly styling
fig.update_layout(
    paper_bgcolor=WHITE,
    plot_bgcolor=WHITE,
    font=dict(
        color=TEXT_PRIMARY,
        family="IBM Plex Sans, Barlow, system-ui, sans-serif",
    ),
    title=dict(font=dict(
        color=TEXT_PRIMARY,
        family="Barlow Condensed, system-ui, sans-serif",
    )),
    legend=dict(
        font=dict(color=TEXT_SECONDARY, size=11),
        title=dict(font=dict(color=TEXT_MUTED, size=10)),
    ),
    xaxis=dict(
        tickfont=dict(color=TEXT_SECONDARY, family="JetBrains Mono, monospace", size=9),
        title=dict(font=dict(color=TEXT_SECONDARY)),
        gridcolor="#F0F0F0",
        linecolor=BORDER,
        linewidth=1,
    ),
    yaxis=dict(
        tickfont=dict(color=TEXT_SECONDARY, family="JetBrains Mono, monospace", size=9),
        title=dict(font=dict(color=TEXT_SECONDARY)),
        gridcolor="#F0F0F0",
        linecolor=BORDER,
        linewidth=1,
    ),
)

st.plotly_chart(fig, use_container_width=True, key="qt_main_chart")


# ---------------------------------------------------------------------------
# Summary Statistics — dense, tabbed
# ---------------------------------------------------------------------------

def _render_capability_card(cap):
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
        cols2[0].metric("Sigma", f"{cap['Sigma Level']}s")
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

        tab_cap, tab_anova, tab_trend = st.tabs([
            "Process Capability", "ANOVA", "Trend / Shift"
        ])

        with tab_cap:
            if len(all_values) < 2:
                st.info("Not enough data for process capability.")
            else:
                cap = calc_process_capability(all_values, usl_val, lsl_val)
                if cap:
                    _render_capability_card(cap)

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
                                rows.append({"Point": pt_label, **{k: v for k, v in pc.items()
                                              if k in ["mean", "std", "Cp", "Cpk", "Pp", "Ppk",
                                                        "Sigma Level", "DPMO", "Yield %", "OOS Count"]}})
                        if rows:
                            cap_df = pd.DataFrame(rows)
                            st.dataframe(cap_df, use_container_width=True, hide_index=True)

        with tab_anova:
            if color_by == "None" or color_by not in df_clean.columns:
                st.info("Select a Color-by grouping for group comparison.")
            else:
                groups = df_clean[color_by].fillna("Unknown").astype(str)
                unique_groups = sorted(groups.unique())
                if len(unique_groups) < 2:
                    st.info("Need 2+ groups for ANOVA.")
                else:
                    group_data = {}
                    for g in unique_groups:
                        mask = groups == g
                        vals = df_clean.loc[mask, valid_cols].values.flatten()
                        vals = pd.Series(vals).dropna()
                        if len(vals) > 0:
                            group_data[g] = vals

                    if len(group_data) < 2:
                        st.info("Not enough data in groups.")
                    else:
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
                            summary_rows.append({
                                "Group": g, "n": len(vals),
                                "Mean": round(vals.mean(), 6), "Std": round(vals.std(ddof=1), 6),
                                "Min": round(vals.min(), 6), "Max": round(vals.max(), 6),
                                "Range": round(vals.max() - vals.min(), 6),
                            })
                        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

                        grand_mean = all_values.mean()
                        ss_between = sum(len(group_data[g]) * (group_data[g].mean() - grand_mean)**2
                                         for g in group_data)
                        ss_within = sum(((group_data[g] - group_data[g].mean())**2).sum()
                                        for g in group_data)
                        ss_total = ss_between + ss_within
                        if ss_total > 0:
                            var_cols = st.columns(3)
                            var_cols[0].metric("SS Between", f"{ss_between:.4f}")
                            var_cols[1].metric("SS Within", f"{ss_within:.4f}")
                            var_cols[2].metric("% Between", f"{ss_between/ss_total*100:.1f}%")

                        fig_box = go.Figure()
                        for g in unique_groups:
                            if g in group_data:
                                fig_box.add_trace(go.Box(
                                    y=group_data[g].values, name=g,
                                    marker_color=custom_color_map.get(g, ACCENT),
                                    boxmean="sd",
                                ))
                        fig_box.update_layout(
                            yaxis_title="Value",
                            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
                            font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
                            height=300,
                            margin=dict(l=40, r=20, t=30, b=40),
                            xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                            yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                        )
                        if usl_val is not None:
                            fig_box.add_hline(y=usl_val, line_dash="dash", line_color=DANGER,
                                              annotation_text=f"USL {usl_val:.4g}")
                        if lsl_val is not None:
                            fig_box.add_hline(y=lsl_val, line_dash="dash", line_color=DANGER,
                                              annotation_text=f"LSL {lsl_val:.4g}")
                        st.plotly_chart(fig_box, use_container_width=True, key=f"qt_anova_box_{dno}")

        with tab_trend:
            if len(all_values) < 9:
                st.info("Need 9+ data points for trend/shift analysis.")
            else:
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
                        viol_rows.append({
                            "Rule": rule_name, "Violations": len(indices),
                            "Indices": str(indices[:20]) + ("..." if len(indices) > 20 else ""),
                        })
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
                    fig_cusum.add_trace(go.Scatter(x=x_idx, y=cusum_pos, mode="lines", name="CUSUM+",
                                                    line=dict(color=ACCENT, width=1.5)))
                    fig_cusum.add_trace(go.Scatter(x=x_idx, y=cusum_neg, mode="lines", name="CUSUM-",
                                                    line=dict(color=DANGER, width=1.5)))
                    fig_cusum.add_hline(y=5.0, line_dash="dash", line_color=TEXT_MUTED,
                                        annotation_text="h=5")
                    if shift_pts:
                        fig_cusum.add_trace(go.Scatter(
                            x=shift_pts, y=[max(cusum_pos[i], cusum_neg[i]) for i in shift_pts],
                            mode="markers", name="Shift",
                            marker=dict(color=DANGER, size=6, symbol="x"),
                        ))
                    fig_cusum.update_layout(
                        xaxis_title="Observation", yaxis_title="Cumulative Sum",
                        paper_bgcolor=WHITE, plot_bgcolor=WHITE,
                        font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
                        height=250, margin=dict(l=40, r=20, t=20, b=40),
                        xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                        yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                    )
                    st.plotly_chart(fig_cusum, use_container_width=True, key=f"qt_cusum_{dno}")

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
                    ewma[i] = lam * all_values.iloc[i] + (1 - lam) * ewma[i-1]
                overall_mean = all_values.mean()
                overall_std = all_values.std(ddof=1)
                ewma_ucl = np.array([
                    overall_mean + 3 * overall_std * np.sqrt(lam / (2 - lam) * (1 - (1-lam)**(2*(i+1))))
                    for i in range(len(all_values))
                ])
                ewma_lcl = np.array([
                    overall_mean - 3 * overall_std * np.sqrt(lam / (2 - lam) * (1 - (1-lam)**(2*(i+1))))
                    for i in range(len(all_values))
                ])

                fig_ewma = go.Figure()
                x_idx = list(range(len(ewma)))
                fig_ewma.add_trace(go.Scatter(x=x_idx, y=ewma, mode="lines", name="EWMA",
                                               line=dict(color=ACCENT, width=2)))
                fig_ewma.add_trace(go.Scatter(x=x_idx, y=ewma_ucl, mode="lines", name="UCL",
                                               line=dict(color=DANGER, dash="dash", width=1)))
                fig_ewma.add_trace(go.Scatter(x=x_idx, y=ewma_lcl, mode="lines", name="LCL",
                                               line=dict(color=DANGER, dash="dash", width=1)))
                fig_ewma.add_hline(y=overall_mean, line_dash="dot", line_color=TEXT_MUTED,
                                    annotation_text="Center")
                ooc_ewma = [i for i in range(len(ewma)) if ewma[i] > ewma_ucl[i] or ewma[i] < ewma_lcl[i]]
                if ooc_ewma:
                    fig_ewma.add_trace(go.Scatter(
                        x=ooc_ewma, y=[ewma[i] for i in ooc_ewma],
                        mode="markers", name="OOC",
                        marker=dict(color=DANGER, size=6, symbol="x"),
                    ))
                fig_ewma.update_layout(
                    xaxis_title="Observation", yaxis_title="EWMA",
                    paper_bgcolor=WHITE, plot_bgcolor=WHITE,
                    font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
                    height=250, margin=dict(l=40, r=20, t=20, b=40),
                    xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                    yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                )
                st.plotly_chart(fig_ewma, use_container_width=True, key=f"qt_ewma_{dno}")

                if ooc_ewma:
                    st.markdown(
                        f"<span style='font-size:0.82rem;color:{WARNING};'>"
                        f"EWMA: {len(ooc_ewma)} out-of-control</span>",
                        unsafe_allow_html=True,
                    )

        st.markdown(f"<hr style='border:none;border-top:1px solid {BORDER};margin:8px 0;'>",
                    unsafe_allow_html=True)
