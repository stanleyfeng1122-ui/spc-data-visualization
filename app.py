"""
SPC Data Visualization Tool

Streamlit web application for visualizing vendor CPK / SPC measurement data.
Upload one or more .xlsx files, select dimensions or dimension groups, and
generate interactive combined profile charts with spec limits, color-coded
by raw material or other groupings.

Chart model (combined profile view):
  X-axis  = concatenated measurement points across sections
  Y-axis  = measured value at each point
  Each line = one part (one data row)
  Color   = group-by field (Raw material, Build, etc.)
  Sections = Factory x Build (e.g. FX P1, FX P2, TRM P1, TRM P2)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from collections import OrderedDict

from spc_parser import (
    parse_excel,
    detect_dimension_groups,
    get_filtered_dim_meta,
    ParsedFile,
    DimensionMeta,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SPC Data Visualization",
    layout="wide",
    initial_sidebar_state="expanded",
)

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


def get_color_for_group(idx: int) -> str:
    return COLOR_PALETTE[idx % len(COLOR_PALETTE)]


# ---------------------------------------------------------------------------
# Sidebar: file upload
# ---------------------------------------------------------------------------
st.sidebar.title("SPC Data Visualization")

uploaded_files = st.sidebar.file_uploader(
    "Upload CPK Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Drag and drop one or more vendor CPK Excel files here.",
)

if not uploaded_files:
    st.title("SPC Data Visualization Tool")
    st.info(
        "Upload one or more .xlsx CPK data files using the sidebar to get started."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Sheet selector (default = Raw data for full profiles)
# ---------------------------------------------------------------------------
sheet_name = st.sidebar.selectbox(
    "Sheet to analyse",
    options=["Raw data", "Data Input"],
    index=0,
    help="'Raw data' has full measurement profiles; 'Data Input' has summary values.",
)

# ---------------------------------------------------------------------------
# Parse uploaded files (cached per file + sheet)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Parsing Excel files...")
def _parse_file(file_bytes: bytes, filename: str, sheet: str) -> dict:
    """Parse and return serialisable dict (cache-friendly)."""
    import io
    buf = io.BytesIO(file_bytes)
    buf.name = filename
    parsed = parse_excel(buf, sheet_name=sheet)
    return {
        "filename": parsed.filename,
        "sheet_name": parsed.sheet_name,
        "part_number": parsed.part_number,
        "part_description": parsed.part_description,
        "revision": parsed.revision,
        "factory": parsed.factory,
        "dimensions": parsed.dimensions,
        "data": parsed.data,
        "meta_columns": parsed.meta_columns,
    }


parsed_files = []
for uf in uploaded_files:
    try:
        raw = uf.read()
        result = _parse_file(raw, uf.name, sheet_name)
        parsed_files.append(result)
    except Exception as e:
        st.sidebar.error(f"Error parsing {uf.name}: {e}")

if not parsed_files:
    st.warning("No files could be parsed. Check the sidebar for errors.")
    st.stop()

# ---------------------------------------------------------------------------
# File summaries
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("Loaded Files")
for pf in parsed_files:
    n_rows = len(pf["data"]) if pf["data"] is not None else 0
    builds = ""
    if pf["data"] is not None and "Build" in pf["data"].columns:
        builds = ", ".join(sorted(pf["data"]["Build"].dropna().unique().astype(str)))
    factory = pf.get("factory", "?")
    st.sidebar.markdown(
        f"**{pf['filename']}**  \n"
        f"Factory: {factory} | Part: {pf['part_number']} | Rows: {n_rows} | Builds: {builds}"
    )

# ---------------------------------------------------------------------------
# Build unified dimension map from all files
# ---------------------------------------------------------------------------
all_dimensions = OrderedDict()
for pf in parsed_files:
    for dno, dmeta in pf["dimensions"].items():
        if dno not in all_dimensions:
            all_dimensions[dno] = dmeta

# ---------------------------------------------------------------------------
# Detect dimension groups
# ---------------------------------------------------------------------------
dim_groups = detect_dimension_groups(all_dimensions)

# ---------------------------------------------------------------------------
# Dimension mode selector
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
dim_mode = st.sidebar.radio(
    "Dimension mode",
    options=["Dimension group", "Single dimension"],
    index=0,
    help="Group mode concatenates related dimensions on one chart. Single mode shows one dimension.",
)

selected_dim_nos = []
selected_group_label = ""

if dim_mode == "Dimension group":
    if dim_groups:
        group_labels = list(dim_groups.keys())
        selected_group_label = st.sidebar.selectbox(
            "Dimension group",
            options=group_labels,
            index=0,
        )
        selected_dim_nos = dim_groups[selected_group_label]
    else:
        st.sidebar.warning("No dimension groups detected. Switching to single dimension mode.")
        dim_mode = "Single dimension"

if dim_mode == "Single dimension":
    dim_display = [
        f"{dno} - {dmeta.description}" if dmeta.description else dno
        for dno, dmeta in all_dimensions.items()
    ]
    if not dim_display:
        st.warning("No dimensions found in the uploaded files.")
        st.stop()
    selected_dim_display = st.sidebar.selectbox(
        "Dimension",
        options=dim_display,
        index=0,
    )
    selected_dim_nos = [selected_dim_display.split(" - ")[0]]
    selected_group_label = selected_dim_display

if not selected_dim_nos:
    st.info("Select at least one dimension from the sidebar.")
    st.stop()

# ---------------------------------------------------------------------------
# Exclude interval points option
# ---------------------------------------------------------------------------
exclude_intervals = st.sidebar.checkbox(
    "Exclude interval points (e.g. C11-C12)",
    value=True,
    help="Hide derived interval measurements, showing only actual C-points.",
)

# ---------------------------------------------------------------------------
# Axis and grouping controls
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("Grouping")

# Determine available metadata columns across all files
available_meta = set()
for pf in parsed_files:
    available_meta.update(pf["meta_columns"])
available_meta.discard("Start Point")

# Color-by / Group-by
GROUPBY_OPTIONS = ["Raw material", "Build", "Vendor Serial Number", "None"]
groupby_options_filtered = [g for g in GROUPBY_OPTIONS if g in available_meta or g == "None"]

color_by = st.sidebar.selectbox(
    "Color-by",
    options=groupby_options_filtered,
    index=0,
    help="Choose how to color-code the data traces.",
)

# Section-by: controls how the chart is divided into sections
SECTION_OPTIONS = ["Factory + Build", "Factory", "Build", "None"]
section_by = st.sidebar.selectbox(
    "Section-by",
    options=SECTION_OPTIONS,
    index=0,
    help="How to split the chart into sections with vertical dividers.",
)

# Y-axis mode
YAXIS_OPTIONS = ["Measurement values", "Deviation from Nominal"]
y_axis_mode = st.sidebar.selectbox(
    "Y-axis",
    options=YAXIS_OPTIONS,
    index=0,
)

# ---------------------------------------------------------------------------
# Prepare combined data across files
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

        # Keep metadata + all measurement columns for selected dims
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
# Section logic
# ---------------------------------------------------------------------------

def compute_sections(df, section_by):
    """
    Assign a section label to each row based on section_by mode.
    Returns a Series of section labels aligned with df index.
    """
    if section_by == "Factory + Build":
        factory_col = df["_factory"].fillna("?").astype(str) if "_factory" in df.columns else pd.Series("?", index=df.index)
        build_col = df["Build"].fillna("?").astype(str) if "Build" in df.columns else pd.Series("?", index=df.index)
        return factory_col + " " + build_col
    elif section_by == "Factory":
        if "_factory" in df.columns:
            return df["_factory"].fillna("?").astype(str)
        return pd.Series("All", index=df.index)
    elif section_by == "Build":
        if "Build" in df.columns:
            return df["Build"].fillna("?").astype(str)
        return pd.Series("All", index=df.index)
    else:
        return pd.Series("All", index=df.index)


# ---------------------------------------------------------------------------
# Chart building -- combined profile view
# ---------------------------------------------------------------------------

MAX_TRACES_PER_GROUP = 600  # per color group per section, to limit Plotly load


def build_combined_chart(
    df,
    dim_metas: OrderedDict,
    dim_nos: list,
    section_by: str,
    color_by: str,
    y_axis_mode: str,
    exclude_intervals: bool,
    group_label: str,
):
    """
    Build the combined profile chart matching the screenshot layout.

    Sections are laid out side-by-side along the X-axis.
    Within each section, measurement points from all selected dimensions
    are concatenated. Each part is one line overlaid on the chart.
    """
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    # --- Compute sections ---
    section_labels = compute_sections(df, section_by)
    unique_sections = list(dict.fromkeys(section_labels))  # preserve order

    # --- Determine color grouping ---
    if color_by != "None" and color_by in df.columns:
        color_series = df[color_by].fillna("Unknown").astype(str)
        unique_colors = sorted(color_series.unique())
    else:
        color_series = pd.Series("All", index=df.index)
        unique_colors = ["All"]

    color_map = {g: get_color_for_group(i) for i, g in enumerate(unique_colors)}

    # --- Build dimension point info for x-axis ---
    # For each dimension, get filtered columns and point labels
    dim_point_info = OrderedDict()  # dno -> (col_labels, point_numbers, nominal, usl, lsl)
    for dno in dim_nos:
        if dno not in dim_metas:
            continue
        dmeta = dim_metas[dno]
        info = get_filtered_dim_meta(dmeta, exclude_intervals=exclude_intervals)
        col_labels, point_numbers, nominal, usl, lsl = info
        # Only keep columns that exist in df
        valid = [(cl, pn, n, u, l) for cl, pn, n, u, l in
                 zip(col_labels, point_numbers, nominal, usl, lsl)
                 if cl in df.columns]
        if valid:
            cls, pns, noms, usls, lsls = zip(*valid)
            dim_point_info[dno] = (list(cls), list(pns), list(noms), list(usls), list(lsls))

    if not dim_point_info:
        return None

    # Total measurement points per section
    points_per_section = sum(len(v[0]) for v in dim_point_info.values())
    if points_per_section == 0:
        return None

    # Gap between sections (in x-units)
    section_gap = max(3, int(points_per_section * 0.06))

    # --- Build figure ---
    fig = go.Figure()

    # Track x-axis positioning
    x_offset = 0
    section_centers = []    # (center_x, label) for section header annotations
    dim_centers = []        # (center_x, label, section_label) for dim sub-labels
    section_boundaries = [] # x positions of section dividers

    # For spec band (use first dim's USL/LSL as representative)
    first_dim_info = list(dim_point_info.values())[0]
    usl_rep = next((v for v in first_dim_info[3] if v is not None), None)
    lsl_rep = next((v for v in first_dim_info[4] if v is not None), None)
    nom_rep = next((v for v in first_dim_info[2] if v is not None), None)

    # Collect all x-tick positions and labels
    all_tick_vals = []
    all_tick_text = []

    legend_shown = set()

    for sec_idx, sec_label in enumerate(unique_sections):
        sec_mask = section_labels == sec_label
        sec_df = df[sec_mask].reset_index(drop=True)
        sec_colors = color_series[sec_mask].reset_index(drop=True)

        if sec_df.empty:
            continue

        section_start_x = x_offset

        # --- Plot each dimension within this section ---
        for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
            n_points = len(col_labels)
            dim_start_x = x_offset

            # X positions for this dimension's points in this section
            x_positions = list(range(x_offset, x_offset + n_points))

            # Add tick labels for the measurement points
            for xi, pn in zip(x_positions, point_nums):
                all_tick_vals.append(xi)
                all_tick_text.append(pn if pn else "")

            # Build nominal array for deviation mode
            nom_array = np.array(
                [n if n is not None else np.nan for n in nominals],
                dtype=float,
            )

            # --- Plot traces for each color group ---
            for grp_name in unique_colors:
                grp_mask = sec_colors == grp_name
                grp_df = sec_df.loc[grp_mask, col_labels].reset_index(drop=True)
                color = color_map[grp_name]

                n_parts = len(grp_df)
                if n_parts == 0:
                    continue

                step = max(1, n_parts // MAX_TRACES_PER_GROUP)

                for row_idx in range(0, n_parts, step):
                    y_vals = pd.to_numeric(grp_df.iloc[row_idx], errors="coerce").values.copy()

                    if deviation_mode:
                        y_vals = y_vals - nom_array

                    show_legend = grp_name not in legend_shown
                    legend_shown.add(grp_name)

                    fig.add_trace(go.Scattergl(
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
                            f"Section: {sec_label}"
                            "<extra></extra>"
                        ),
                        text=[pn for pn in point_nums],
                    ))

            # Record dimension center for sub-label
            dim_center_x = x_offset + n_points / 2
            dim_centers.append((dim_center_x, dno, sec_label))

            x_offset += n_points

        # Record section center and boundary
        section_end_x = x_offset
        section_center_x = (section_start_x + section_end_x) / 2
        section_centers.append((section_center_x, sec_label))

        # Add gap before next section
        if sec_idx < len(unique_sections) - 1:
            section_boundaries.append(x_offset + section_gap / 2)
            x_offset += section_gap

    # --- Add spec limit green band ---
    if usl_rep is not None and lsl_rep is not None:
        band_usl = usl_rep
        band_lsl = lsl_rep
        if deviation_mode and nom_rep is not None:
            band_usl = usl_rep - nom_rep
            band_lsl = lsl_rep - nom_rep
        fig.add_hrect(
            y0=band_lsl, y1=band_usl,
            fillcolor="rgba(34, 197, 94, 0.15)",
            line_width=0,
            layer="below",
        )

    # --- Add reference lines ---
    dash_style = dict(dash="dash", width=1.2)

    if usl_rep is not None:
        ref_usl = (usl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else usl_rep
        fig.add_hline(
            y=ref_usl,
            line=dict(color="rgba(220,38,38,0.5)", **dash_style),
        )

    if lsl_rep is not None:
        ref_lsl = (lsl_rep - nom_rep) if (deviation_mode and nom_rep is not None) else lsl_rep
        fig.add_hline(
            y=ref_lsl,
            line=dict(color="rgba(220,38,38,0.5)", **dash_style),
        )

    # --- Add section divider lines ---
    for bx in section_boundaries:
        fig.add_vline(
            x=bx,
            line=dict(color="rgba(100,116,139,0.5)", width=1.5, dash="solid"),
        )

    # --- Add section header annotations (top of chart) ---
    annotations = []

    # Section headers (top level: factory name)
    # Group sections by factory prefix for two-level header
    factory_sections = OrderedDict()  # factory -> [(center_x, build_label)]
    for center_x, sec_label in section_centers:
        parts = sec_label.split(None, 1)
        if len(parts) == 2:
            factory, build = parts
        else:
            factory = sec_label
            build = ""
        factory_sections.setdefault(factory, []).append((center_x, build))

    # Add factory-level header (top)
    for factory, builds in factory_sections.items():
        if len(builds) > 1:
            factory_center = sum(cx for cx, _ in builds) / len(builds)
        else:
            factory_center = builds[0][0]
        annotations.append(dict(
            x=factory_center,
            y=1.12,
            xref="x",
            yref="paper",
            text=f"<b>{factory}</b>",
            showarrow=False,
            font=dict(size=13, color="#1E293B"),
        ))

        # Build-level sub-header
        for cx, build_label in builds:
            if build_label:
                annotations.append(dict(
                    x=cx,
                    y=1.06,
                    xref="x",
                    yref="paper",
                    text=f"<b>{build_label}</b>",
                    showarrow=False,
                    font=dict(size=11, color="#475569"),
                ))

    # --- Build title ---
    is_group = len(dim_nos) > 1
    if is_group:
        dim_names = "/".join(dno.replace("SPC_", "") for dno in dim_nos)
        first_desc = ""
        for dno in dim_nos:
            if dno in dim_metas and dim_metas[dno].description:
                # Extract common keyword from description
                desc = dim_metas[dno].description
                # Simplify: use the keyword portion
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

    # Add subtitle
    subtitle = "Vendor Serial Number" if section_by != "None" else ""

    # --- Layout ---
    y_title = "Deviation from Nominal" if deviation_mode else ""

    fig.update_layout(
        title=dict(
            text=f"<b>{title_text}</b>" + (f"<br><span style='font-size:12px;color:#64748B'>{subtitle}</span>" if subtitle else ""),
            font=dict(size=15),
            x=0.5,
            xanchor="center",
        ),
        xaxis=dict(
            tickmode="array",
            tickvals=all_tick_vals[::max(1, len(all_tick_vals) // 80)],  # Show at most ~80 tick labels
            ticktext=all_tick_text[::max(1, len(all_tick_vals) // 80)],
            tickangle=-90,
            tickfont=dict(size=7),
            showgrid=False,
        ),
        yaxis=dict(
            title=y_title,
            zeroline=True,
            zerolinecolor="rgba(100,116,139,0.3)",
        ),
        height=620,
        margin=dict(l=50, r=120, t=100, b=80),
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


# ---------------------------------------------------------------------------
# Main content area
# ---------------------------------------------------------------------------
st.title("SPC Data Visualization")

# Combine data from all files
df, dim_metas = prepare_combined_data(parsed_files, selected_dim_nos)

if df is None or dim_metas is None or df.empty:
    st.warning("No data found for the selected dimensions in the uploaded file(s).")
    st.stop()

# Drop rows where ALL measurement columns are NaN
all_meas_cols = []
for dno in selected_dim_nos:
    if dno in dim_metas:
        all_meas_cols.extend([c for c in dim_metas[dno].col_labels if c in df.columns])

if all_meas_cols:
    df_clean = df.dropna(subset=all_meas_cols, how="all").reset_index(drop=True)
else:
    df_clean = df

if df_clean.empty:
    st.warning("No measurement data for the selected dimensions.")
    st.stop()

# Build the chart
fig = build_combined_chart(
    df=df_clean,
    dim_metas=dim_metas,
    dim_nos=selected_dim_nos,
    section_by=section_by,
    color_by=color_by,
    y_axis_mode=y_axis_mode,
    exclude_intervals=exclude_intervals,
    group_label=selected_group_label,
)

if fig is None:
    st.warning("Could not generate chart. Check that the selected dimensions have data.")
    st.stop()

st.plotly_chart(fig, use_container_width=True, key="main_chart")

# ---------------------------------------------------------------------------
# Summary statistics
# ---------------------------------------------------------------------------
with st.expander("Summary Statistics"):
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

        st.markdown(f"**{dno} - {dmeta.description}**")

        stats_data = df_clean[valid_cols].describe().T
        stats_data.index.name = "Point"
        st.dataframe(stats_data, use_container_width=True, height=200)

        usl_val = next((v for v in usls if v is not None), None)
        nom_val = next((v for v in nominals if v is not None), None)
        lsl_val = next((v for v in lsls if v is not None), None)

        col1, col2, col3 = st.columns(3)
        col1.metric("USL", usl_val if usl_val is not None else "N/A")
        col2.metric("Nominal", nom_val if nom_val is not None else "N/A")
        col3.metric("LSL", lsl_val if lsl_val is not None else "N/A")

        st.markdown("---")
