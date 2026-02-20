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
from plotly.subplots import make_subplots
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
with st.sidebar.expander(f"Loaded Files ({len(parsed_files)})", expanded=False):
    for pf in parsed_files:
        n_rows = len(pf["data"]) if pf["data"] is not None else 0
        builds = ""
        if pf["data"] is not None and "Build" in pf["data"].columns:
            builds = ", ".join(sorted(pf["data"]["Build"].dropna().unique().astype(str)))
        factory = pf.get("factory", "?")
        st.markdown(
            f"**{pf['filename']}**  \n"
            f"<small>Factory: {factory} | Part: {pf['part_number']} | Rows: {n_rows} | Builds: {builds}</small>",
            unsafe_allow_html=True,
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
# Dimension selector (preset groups + user-editable multi-select)
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")

# Detect preset groups from all dimensions
dim_groups = detect_dimension_groups(all_dimensions)

# Build display labels for each dimension
dim_display_map = OrderedDict()  # display_label -> dim_no
for dno, dmeta in all_dimensions.items():
    label = f"{dno} - {dmeta.description}" if dmeta.description else dno
    dim_display_map[label] = dno

# Reverse map: dim_no -> display_label
dim_no_to_label = {v: k for k, v in dim_display_map.items()}

dim_display_labels = list(dim_display_map.keys())

if not dim_display_labels:
    st.warning("No dimensions found in the uploaded files.")
    st.stop()

# Preset group selector
group_options = ["Custom selection"] + list(dim_groups.keys())
selected_preset = st.sidebar.selectbox(
    "Dimension preset",
    options=group_options,
    index=0,
    help="Pick a preset group to auto-fill dimensions, or choose 'Custom selection' to pick manually.",
)

# Determine default selection from preset
if selected_preset != "Custom selection":
    preset_dim_nos = dim_groups[selected_preset]
    default_labels = [dim_no_to_label[dno] for dno in preset_dim_nos if dno in dim_no_to_label]
else:
    default_labels = [dim_display_labels[0]] if dim_display_labels else []

# Editable multi-select (preset fills the default, user can still modify)
selected_dim_labels = st.sidebar.multiselect(
    "Dimensions",
    options=dim_display_labels,
    default=default_labels,
    help="Select one or more dimensions to plot together. Use the preset above for quick selection.",
)

selected_dim_nos = [dim_display_map[lbl] for lbl in selected_dim_labels]
selected_group_label = " / ".join(
    dno.replace("SPC_", "") for dno in selected_dim_nos
) if selected_dim_nos else ""

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
# Chart type selector
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
chart_type = st.sidebar.radio(
    "Chart type",
    options=["Combined Profile", "Box Plot", "Histogram"],
    index=0,
    help=(
        "Combined Profile: overlay all part traces. "
        "Box Plot: distribution per measurement point. "
        "Histogram: frequency distribution of values."
    ),
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
GROUPBY_OPTIONS = ["Raw material", "Build", "Vendor Serial Number", "Config", "None"]
groupby_options_filtered = [g for g in GROUPBY_OPTIONS if g in available_meta or g == "None"]

color_by = st.sidebar.selectbox(
    "Color-by",
    options=groupby_options_filtered,
    index=0,
    help="Choose how to color-code the data traces.",
)

# Section-by (X grouping): split chart into columns
if chart_type in ("Combined Profile", "Box Plot"):
    SECTION_OPTIONS = ["Factory + Build", "Factory", "Build", "Config", "None"]
    section_options_filtered = [s for s in SECTION_OPTIONS
                                if s in available_meta or s in ("Factory + Build", "None")]
    section_by = st.sidebar.selectbox(
        "Section-by (columns)",
        options=section_options_filtered,
        index=0,
        help="How to split the chart into column sections with vertical dividers.",
    )
else:
    section_by = "None"

# Row-by (Y grouping): split chart into subplot rows
ROWBY_OPTIONS = ["Raw material", "Build", "Factory", "Config", "None"]
rowby_options_filtered = [r for r in ROWBY_OPTIONS if r in available_meta or r == "None"]
row_by = st.sidebar.selectbox(
    "Row-by (rows)",
    options=rowby_options_filtered,
    index=len(rowby_options_filtered) - 1,  # default to "None"
    help="Split the chart into vertically stacked rows by this field.",
)

# Y-axis mode: Combined Profile and Box Plot
if chart_type in ("Combined Profile", "Box Plot"):
    YAXIS_OPTIONS = ["Measurement values", "Deviation from Nominal"]
    y_axis_mode = st.sidebar.selectbox(
        "Y-axis",
        options=YAXIS_OPTIONS,
        index=0,
    )
else:
    y_axis_mode = "Measurement values"

# Histogram-specific controls
if chart_type == "Histogram":
    hist_nbins = st.sidebar.slider("Number of bins", 10, 100, 40)
else:
    hist_nbins = 40

# ---------------------------------------------------------------------------
# Y-axis range controls
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("Y-axis Range")
use_custom_yrange = st.sidebar.checkbox(
    "Set custom Y-axis range",
    value=False,
    help="Manually set the min and max of the Y-axis.",
)
if use_custom_yrange:
    y_min = st.sidebar.number_input("Y-axis min", value=0.0, format="%.4f", key="ymin")
    y_max = st.sidebar.number_input("Y-axis max", value=1.0, format="%.4f", key="ymax")
    if y_min >= y_max:
        st.sidebar.warning("Y-axis min must be less than max.")
        custom_yrange = None
    else:
        custom_yrange = [y_min, y_max]
else:
    custom_yrange = None

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
    elif section_by in ("Build", "Config", "Raw material"):
        if section_by in df.columns:
            return df[section_by].fillna("?").astype(str)
        return pd.Series("All", index=df.index)
    else:
        return pd.Series("All", index=df.index)


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
    row_by: str = "None",
    custom_color_map: dict = None,
    custom_yrange: list = None,
):
    """
    Build the combined profile chart.

    Sections are laid out side-by-side along the X-axis (column facets).
    When row_by is set, the chart is split into vertically stacked subplot
    rows, one per unique value of the row_by field.
    """
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    # --- Compute sections (column facets) ---
    section_labels = compute_sections(df, section_by)
    unique_sections = list(dict.fromkeys(section_labels))  # preserve order

    # --- Compute row facets ---
    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))  # preserve order
    n_rows = len(unique_rows)
    use_row_facets = n_rows > 1

    # --- Determine color grouping ---
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

    # --- Build dimension point info for x-axis ---
    dim_point_info = OrderedDict()
    for dno in dim_nos:
        if dno not in dim_metas:
            continue
        dmeta = dim_metas[dno]
        info = get_filtered_dim_meta(dmeta, exclude_intervals=exclude_intervals)
        col_labels, point_numbers, nominal, usl, lsl = info
        valid = [(cl, pn, n, u, l) for cl, pn, n, u, l in
                 zip(col_labels, point_numbers, nominal, usl, lsl)
                 if cl in df.columns]
        if valid:
            cls, pns, noms, usls, lsls = zip(*valid)
            dim_point_info[dno] = (list(cls), list(pns), list(noms), list(usls), list(lsls))

    if not dim_point_info:
        return None

    points_per_section = sum(len(v[0]) for v in dim_point_info.values())
    if points_per_section == 0:
        return None

    section_gap = max(3, int(points_per_section * 0.06))

    # --- Build figure (with or without row subplots) ---
    if use_row_facets:
        fig = make_subplots(
            rows=n_rows, cols=1,
            shared_xaxes=True,
            row_titles=[str(r) for r in unique_rows],
            vertical_spacing=0.06,
        )
    else:
        fig = go.Figure()

    # For spec band
    first_dim_info = list(dim_point_info.values())[0]
    usl_rep = next((v for v in first_dim_info[3] if v is not None), None)
    lsl_rep = next((v for v in first_dim_info[4] if v is not None), None)
    nom_rep = next((v for v in first_dim_info[2] if v is not None), None)

    legend_shown = set()

    # We need consistent x-axis layout across all rows, so compute it once
    # based on sections (same x positions for every row).
    all_tick_vals = []
    all_tick_text = []
    section_centers = []
    section_boundaries = []

    # Pre-compute x layout
    x_offset = 0
    section_x_ranges = {}  # sec_label -> (start_x, end_x)
    dim_x_positions = OrderedDict()  # (sec_label, dno) -> x_positions

    for sec_idx, sec_label in enumerate(unique_sections):
        section_start_x = x_offset
        for dno, (col_labels, point_nums, nominals, usls, lsls) in dim_point_info.items():
            n_points = len(col_labels)
            x_positions = list(range(x_offset, x_offset + n_points))
            dim_x_positions[(sec_label, dno)] = x_positions

            # Only add ticks once (first pass)
            for xi, pn in zip(x_positions, point_nums):
                all_tick_vals.append(xi)
                all_tick_text.append(pn if pn else "")
            x_offset += n_points

        section_end_x = x_offset
        section_x_ranges[sec_label] = (section_start_x, section_end_x)
        section_centers.append(((section_start_x + section_end_x) / 2, sec_label))

        if sec_idx < len(unique_sections) - 1:
            section_boundaries.append(x_offset + section_gap / 2)
            x_offset += section_gap

    # --- Plot traces for each row facet ---
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

    # --- Add spec limits and bands to all rows ---
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

    # --- Add section divider lines ---
    for bx in section_boundaries:
        fig.add_vline(x=bx, line=dict(color="rgba(100,116,139,0.5)", width=1.5, dash="solid"))

    # --- Section header annotations ---
    annotations = []
    factory_sections = OrderedDict()
    for center_x, sec_label in section_centers:
        parts = sec_label.split(None, 1)
        if len(parts) == 2:
            factory, build = parts
        else:
            factory = sec_label
            build = ""
        factory_sections.setdefault(factory, []).append((center_x, build))

    # Use xref for the first (top) x-axis
    xref = "x" if not use_row_facets else "x"
    for factory, builds in factory_sections.items():
        factory_center = sum(cx for cx, _ in builds) / len(builds)
        annotations.append(dict(
            x=factory_center, y=1.08, xref=xref, yref="paper",
            text=f"<b>{factory}</b>", showarrow=False,
            font=dict(size=13, color="#1E293B"),
        ))
        for cx, build_label in builds:
            if build_label:
                annotations.append(dict(
                    x=cx, y=1.03, xref=xref, yref="paper",
                    text=f"<b>{build_label}</b>", showarrow=False,
                    font=dict(size=11, color="#475569"),
                ))

    # --- Build title ---
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

    subtitle = "Vendor Serial Number" if section_by != "None" else ""
    y_title = "Deviation from Nominal" if deviation_mode else ""

    # --- Tick settings: apply to all x-axes (bottom row shows labels) ---
    tick_step = max(1, len(all_tick_vals) // 80)
    tick_kwargs = dict(
        tickmode="array",
        tickvals=all_tick_vals[::tick_step],
        ticktext=all_tick_text[::tick_step],
        tickangle=-90,
        tickfont=dict(size=7),
        showgrid=False,
    )

    chart_height = 350 * n_rows if use_row_facets else 620

    fig.update_layout(
        title=dict(
            text=f"<b>{title_text}</b>" + (f"<br><span style='font-size:12px;color:#64748B'>{subtitle}</span>" if subtitle else ""),
            font=dict(size=15), x=0.5, xanchor="center",
        ),
        height=chart_height,
        margin=dict(l=50, r=120, t=100, b=80),
        legend=dict(
            title=dict(text=color_by if color_by != "None" else ""),
            orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
            font=dict(size=11), bgcolor="rgba(255,255,255,0.8)",
        ),
        annotations=annotations,
        hovermode="closest",
        template="plotly_white",
    )

    # Build USL/LSL Y-axis tick values for secondary axis
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

    # Apply x-axis tick settings and y-axis title to all subplots
    y_range_kwargs = dict(range=custom_yrange) if custom_yrange else {}
    if use_row_facets:
        # Only show tick labels on bottom row
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

    # Add USL/LSL value annotations on the left side of the Y-axis
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
):
    """
    Build a box plot showing the distribution of measurements at each point.
    Supports row faceting via row_by parameter.
    """
    deviation_mode = y_axis_mode == "Deviation from Nominal"

    # Row facets
    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))
    n_rows = len(unique_rows)
    use_row_facets = n_rows > 1

    # Color grouping
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
                     if cl in df.columns]
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

    # Spec limit lines on all rows
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

    # Build USL/LSL Y-axis tick values for secondary axis
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
        xaxis=dict(title="Measurement Point", tickangle=-45, tickfont=dict(size=8)),
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

    # Add USL/LSL value annotations on the left side of the Y-axis
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
):
    """
    Build a histogram showing frequency distribution of measurement values.

    Grid layout: columns = dimensions, rows = row_by groups.
    Vertical lines for USL, LSL, Nominal.
    Color-by groups are overlaid as separate histogram traces.
    """
    valid_dim_nos = [d for d in dim_nos if d in dim_metas]
    n_dim_cols = len(valid_dim_nos)
    if n_dim_cols == 0:
        return None

    # Row facets
    row_labels = compute_row_groups(df, row_by)
    unique_rows = list(dict.fromkeys(row_labels))
    n_facet_rows = len(unique_rows)

    # Color grouping
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

    # Determine subplot grid
    # Columns = dimensions (or 1 if single dim), Rows = row_by groups (or 1)
    n_cols = max(n_dim_cols, 1)
    n_rows = max(n_facet_rows, 1)
    use_subplots = n_cols > 1 or n_rows > 1

    if use_subplots:
        # Build subplot titles: row_label / dim_no
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
            col_labels, _, nominals, usls, lsls = get_filtered_dim_meta(
                dmeta, exclude_intervals=exclude_intervals
            )
            valid_cols = [c for c in col_labels if c in df.columns]
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

            # Spec limit lines
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

# ---------------------------------------------------------------------------
# Per-group color pickers
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("Group Colors")

if color_by != "None" and color_by in df_clean.columns:
    _color_groups = sorted(df_clean[color_by].fillna("Unknown").astype(str).unique())
else:
    _color_groups = ["All"]

custom_color_map = {}
for i, grp in enumerate(_color_groups):
    default_color = get_color_for_group(i)
    custom_color_map[grp] = st.sidebar.color_picker(
        f"{grp}", value=default_color, key=f"color_{grp}"
    )

# Build the chart based on selected chart type
if chart_type == "Combined Profile":
    fig = build_combined_chart(
        df=df_clean,
        dim_metas=dim_metas,
        dim_nos=selected_dim_nos,
        section_by=section_by,
        color_by=color_by,
        y_axis_mode=y_axis_mode,
        exclude_intervals=exclude_intervals,
        group_label=selected_group_label,
        row_by=row_by,
        custom_color_map=custom_color_map,
        custom_yrange=custom_yrange,
    )
elif chart_type == "Box Plot":
    fig = build_box_plot(
        df=df_clean,
        dim_metas=dim_metas,
        dim_nos=selected_dim_nos,
        color_by=color_by,
        y_axis_mode=y_axis_mode,
        exclude_intervals=exclude_intervals,
        group_label=selected_group_label,
        row_by=row_by,
        custom_color_map=custom_color_map,
        custom_yrange=custom_yrange,
    )
elif chart_type == "Histogram":
    fig = build_histogram(
        df=df_clean,
        dim_metas=dim_metas,
        dim_nos=selected_dim_nos,
        color_by=color_by,
        exclude_intervals=exclude_intervals,
        group_label=selected_group_label,
        nbins=hist_nbins,
        row_by=row_by,
        custom_color_map=custom_color_map,
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
