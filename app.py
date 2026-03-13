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
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from collections import OrderedDict
from scipy import stats as scipy_stats

from spc_parser import (
    parse_excel,
    parse_excel_multi,
    detect_dimension_groups,
    get_filtered_dim_meta,
    ParsedFile,
    DimensionMeta,
)
from chart_utils import (
    COLOR_PALETTE,
    get_color_for_group,
    prepare_combined_data,
    compute_sections,
    compute_row_groups,
    build_combined_chart,
    build_box_plot,
    build_histogram,
    finalize_plotly_style,
    calc_process_capability,
    nelson_rules,
    cusum_analysis,
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
# Global styling: white background + SF Pro font
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.cdnfonts.com/css/sf-pro-display');

/* =================================================================
   1. GLOBAL FONT — target text elements only, NOT icon spans
      Streamlit applies icon font via inline style, so we must
      NOT use * or !important on font-family for all elements.
   ================================================================= */
html, body, div, section, main, aside, header, footer, nav,
h1, h2, h3, h4, h5, h6, p, a, label, li, td, th,
input, textarea, select, button, small, strong, em, code, pre,
[data-testid="stMarkdownContainer"],
[data-testid="stWidgetLabel"],
[data-baseweb="select"],
[data-baseweb="input"] {
    font-family: 'SF Pro Display', 'SF Pro', -apple-system, BlinkMacSystemFont, sans-serif !important;
}
/* Explicitly do NOT override span — Streamlit icon spans get inline font-family */
/* Only override spans that are clearly text content */
p span, label span, h1 span, h2 span, h3 span, h4 span,
[data-testid="stMarkdownContainer"] span,
[data-testid="stWidgetLabel"] span,
[data-testid="stMetricLabel"] span,
[data-testid="stMetricValue"] span {
    font-family: 'SF Pro Display', 'SF Pro', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* =================================================================
   1b. MATERIAL SYMBOLS — ensure icon spans keep their font
   ================================================================= */
span[style*="Material Symbols"] {
    font-family: "Material Symbols Rounded" !important;
    color: #64748B !important;
}
html, body,
.stApp, .main, .main > div, .block-container,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
[data-testid="stVerticalBlock"],
[data-testid="stHorizontalBlock"],
[data-testid="column"],
[data-testid="stMainBlockContainer"],
.element-container, .stMarkdown,
div[data-testid] {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
}

/* =================================================================
   2. HEADER BAR
   ================================================================= */
header, header[data-testid="stHeader"], .stAppHeader,
[data-testid="stHeader"], [data-testid="stToolbar"] {
    background-color: #FFFFFF !important;
    border-bottom: 1px solid #F1F5F9 !important;
}
[data-testid="stDecoration"] { display: none !important; }

/* =================================================================
   3. SIDEBAR — light gray
   ================================================================= */
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div,
section[data-testid="stSidebar"] > div > div,
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"],
section[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"],
section[data-testid="stSidebar"] .element-container {
    background-color: #F8FAFC !important;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] h4,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] small,
section[data-testid="stSidebar"] li {
    color: #334155 !important;
}
section[data-testid="stSidebar"] hr { border-color: #E2E8F0 !important; }

/* =================================================================
   4. TEXT COLORS
   ================================================================= */
h1 { color: #0F172A !important; }
h2, h3, h4 { color: #1E293B !important; }
p, span, label, li, td, th, div, small, strong, a { color: #1E293B !important; }

/* =================================================================
   5. ALERTS (info/warning/error/success)
   ================================================================= */
[data-testid="stAlert"], div[role="alert"],
[data-testid="stNotification"] {
    background: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 8px !important;
}
[data-testid="stAlert"] *, div[role="alert"] * { color: #334155 !important; }

/* =================================================================
   6. EXPANDER
   ================================================================= */
[data-testid="stExpander"],
[data-testid="stExpander"] > div,
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary {
    background: #F8FAFC !important;
    border-color: #E2E8F0 !important;
    color: #1E293B !important;
}
[data-testid="stExpander"] * { color: #1E293B !important; }

/* =================================================================
   7. METRIC CARDS
   ================================================================= */
[data-testid="stMetric"] {
    background: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 8px !important;
    padding: 12px !important;
}
[data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * { color: #64748B !important; }
[data-testid="stMetricValue"], [data-testid="stMetricValue"] * { color: #0F172A !important; }

/* =================================================================
   8. SELECT / INPUT / MULTISELECT / NUMBER INPUT
   ================================================================= */
[data-baseweb="select"],
[data-baseweb="select"] > div,
[data-baseweb="select"] ul,
[data-baseweb="input"],
[data-baseweb="input"] > div {
    background: #FFFFFF !important;
    border: 1px solid #94A3B8 !important;
    border-radius: 6px !important;
    color: #1E293B !important;
}
[data-baseweb="select"] *, [data-baseweb="input"] * { color: #1E293B !important; }
[data-baseweb="tag"] { background: #E2E8F0 !important; color: #1E293B !important; }
[data-testid="stNumberInput"] input {
    background: #FFFFFF !important; color: #1E293B !important;
    border: 1px solid #94A3B8 !important;
    border: 1px solid #CBD5E1 !important;
}

/* =================================================================
   9. FILE UPLOADER
   ================================================================= */
[data-testid="stFileUploader"],
[data-testid="stFileUploader"] > div,
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploaderDropzone"] > div {
    background: #F8FAFC !important;
    border-color: #CBD5E1 !important;
    color: #334155 !important;
}
[data-testid="stFileUploaderDropzone"] { border: 2px dashed #94A3B8 !important; border-radius: 8px !important; }
[data-testid="stFileUploaderDropzone"] * { color: #64748B !important; }
/* "Browse files" button */
[data-testid="stFileUploaderDropzone"] button,
[data-testid="baseButton-secondary"] {
    background: #FFFFFF !important; color: #2563EB !important;
    border: 1px solid #2563EB !important; border-radius: 6px !important;
}
/* Uploaded file items */
[data-testid="stFileUploaderFile"],
[data-testid="stFileUploaderFile"] > div,
[data-testid="stUploadedFile"],
[data-testid="stUploadedFile"] > div,
li[class*="uploadedFile"], li[class*="UploadedFile"],
div[class*="uploadedFile"], div[class*="UploadedFile"],
div[class*="UploadedFileInfo"] {
    background: #F8FAFC !important;
    color: #334155 !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 6px !important;
}

/* =================================================================
   10. DROPDOWNS / POPOVER / LISTBOX
   ================================================================= */
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
[data-baseweb="menu"],
[data-baseweb="listbox"],
[data-baseweb="popover"] ul {
    background: #FFFFFF !important;
    border: 1px solid #E2E8F0 !important;
}
[data-baseweb="menu"] *, [data-baseweb="listbox"] * { color: #1E293B !important; }
[data-baseweb="menu"] li:hover, [data-baseweb="listbox"] li:hover,
[data-baseweb="menu"] li[aria-selected="true"],
[data-baseweb="listbox"] li[aria-selected="true"] {
    background: #F1F5F9 !important;
}

/* =================================================================
   11. CHECKBOX / SLIDER / RADIO / COLOR PICKER
   ================================================================= */
[data-testid="stCheckbox"] *, [data-testid="stSlider"] *,
[data-testid="stRadio"] *, [data-testid="stColorPicker"] label {
    color: #334155 !important;
}

/* =================================================================
   11b. SEGMENTED CONTROL (pill toggle)
   ================================================================= */
[data-testid="stSegmentedControl"] {
    background: #F1F5F9 !important;
    border-radius: 8px !important;
    padding: 3px !important;
    border: 1px solid #E2E8F0 !important;
}
[data-testid="stSegmentedControl"] button {
    border-radius: 6px !important;
    border: none !important;
    background: transparent !important;
    color: #64748B !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
    padding: 6px 12px !important;
    transition: all 0.15s ease !important;
}
[data-testid="stSegmentedControl"] button[aria-checked="true"] {
    background: #FFFFFF !important;
    color: #0F172A !important;
    font-weight: 600 !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08) !important;
}
[data-testid="stSegmentedControl"] button:hover:not([aria-checked="true"]) {
    color: #334155 !important;
    background: rgba(255,255,255,0.5) !important;
}

/* =================================================================
   12. PLOTLY CHART CONTAINER — force white bg on iframe wrapper
   ================================================================= */
[data-testid="stPlotlyChart"],
[data-testid="stPlotlyChart"] > div,
[data-testid="stPlotlyChart"] iframe,
.stPlotlyChart, .stPlotlyChart > div,
iframe[title="streamlit_plotly_events"],
div[class*="plotly"], div[class*="chart"] {
    background: #FFFFFF !important;
}

/* =================================================================
   13. COMPONENT IFRAME CONTAINERS (JS components, etc.)
   ================================================================= */
iframe, [data-testid="stIFrame"],
[data-testid="stCustomComponentV1"],
[data-testid="stCustomComponentV1"] > div,
.stHtml, .stHtml > div {
    background: #FFFFFF !important;
    border: none !important;
}

/* =================================================================
   14. TABLES / DATAFRAMES
   ================================================================= */
[data-testid="stDataFrame"], [data-testid="stTable"],
[data-testid="stDataFrame"] *, [data-testid="stTable"] * {
    background: #FFFFFF !important;
    color: #1E293B !important;
}

/* =================================================================
   15. FOOTER
   ================================================================= */
footer, footer * { background: #FFFFFF !important; color: #94A3B8 !important; }

/* =================================================================
   16. SVG ICONS — visible on white
   ================================================================= */
.stApp svg { fill: #64748B !important; color: #64748B !important; }
[data-testid="stAlert"] svg { fill: #94A3B8 !important; }
/* Don't override Plotly chart SVG colors */
[data-testid="stPlotlyChart"] svg,
.js-plotly-plot svg, .plot-container svg { fill: unset !important; color: unset !important; }

/* =================================================================
   17. BUTTONS
   ================================================================= */
.stApp button { color: #334155 !important; }
.stApp button svg { fill: #64748B !important; }
header button, [data-testid="stToolbar"] button { color: #64748B !important; }

/* =================================================================
   18. CATCH-ALL — nuke any remaining dark backgrounds
   ================================================================= */
.stApp div:not([data-testid="stPlotlyChart"] *) {
    background-color: transparent;
}
.stApp > div, .stApp > div > div,
[data-testid="stVerticalBlockBorderWrapper"],
[data-testid="stVerticalBlockBorderWrapper"] > div {
    background-color: #FFFFFF !important;
}
</style>
""", unsafe_allow_html=True)


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
# Sheet selector – dynamically reads sheets from uploaded files
# ---------------------------------------------------------------------------
import openpyxl, io as _io

_all_sheet_names = []
for _uf in uploaded_files:
    _raw = _uf.getvalue()
    _wb = openpyxl.load_workbook(_io.BytesIO(_raw), read_only=True, data_only=True)
    for _sn in _wb.sheetnames:
        if _sn not in _all_sheet_names:
            _all_sheet_names.append(_sn)
    _wb.close()

# Filter out known non-data sheets (charts, histograms)
_NON_DATA_PREFIXES = ("BoxPlotCht", "Histo ")
_NON_DATA_EXACT = {"Histo Pivot", "Histo Listbox", "Histo Curve"}
_data_sheets = [s for s in _all_sheet_names
                if s not in _NON_DATA_EXACT and not any(s.startswith(p) for p in _NON_DATA_PREFIXES)]

# Add "Auto-detect" as first option — lets parse_excel_multi find data sheets automatically
_sheet_options = ["Auto-detect"] + _data_sheets

sheet_choice = st.sidebar.selectbox(
    "Sheet to analyse",
    options=_sheet_options,
    index=0,
    help="'Auto-detect' scans all sheets for measurement data. Or pick a specific sheet.",
)
sheet_name = "Raw data" if sheet_choice == "Auto-detect" else sheet_choice

# ---------------------------------------------------------------------------
# Parse uploaded files (cached per file + sheet)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Parsing Excel files...")
def _parse_file(file_bytes: bytes, filename: str, sheet: str) -> list:
    """Parse and return list of serialisable dicts (cache-friendly).

    Uses parse_excel_multi to handle files with multiple data sheets
    (e.g. files without a 'Raw data' sheet that have separate sheets
    like 'X3745DH MP', 'X3744DH MP').
    """
    import io
    buf = io.BytesIO(file_bytes)
    buf.name = filename
    parsed_list = parse_excel_multi(buf, sheet_name=sheet)
    results = []
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
    return results


parsed_files = []
for uf in uploaded_files:
    try:
        raw = uf.read()
        results = _parse_file(raw, uf.name, sheet_name)
        parsed_files.extend(results)
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
        sheet_label = f" [{pf['sheet_name']}]" if pf.get('sheet_name') else ""
        st.markdown(
            f"**{pf['filename']}{sheet_label}**  \n"
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
# Points selector – lets the user pick which measurement points to display
# ---------------------------------------------------------------------------
_all_point_numbers = []
for _dno in selected_dim_nos:
    if _dno in all_dimensions:
        _meta = all_dimensions[_dno]
        _cls, _pns, _noms, _usls, _lsls = get_filtered_dim_meta(_meta, exclude_intervals=False)
        for _pn in _pns:
            if _pn and _pn not in _all_point_numbers:
                _all_point_numbers.append(_pn)

selected_points = st.sidebar.multiselect(
    "Points to display",
    options=_all_point_numbers,
    default=_all_point_numbers,
    help="Select which measurement points to include in the chart.",
)
# Empty selection is treated as "show all" for backward compatibility
if not selected_points:
    selected_points = None

# ---------------------------------------------------------------------------
# Chart type selector
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
chart_type = st.sidebar.segmented_control(
    "Chart type",
    options=["Combined Profile", "Box Plot", "Histogram"],
    default="Combined Profile",
    selection_mode="single",
)
if chart_type is None:
    chart_type = "Combined Profile"

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
    SECTION_FIELDS = ["Factory", "Build", "Config", "Raw material",
                      "Vendor Serial Number", "Source File"]
    section_fields_available = [
        s for s in SECTION_FIELDS
        if s in available_meta or s in ("Factory", "Source File")
    ]
    section_by_fields = st.sidebar.multiselect(
        "Section-by (columns)",
        options=section_fields_available,
        default=["Factory"] if "Factory" in section_fields_available else [],
        help="Choose one or more fields to combine for section grouping.",
    )
else:
    section_by_fields = []

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
# Chart rendering
# ---------------------------------------------------------------------------


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
        section_by_fields=section_by_fields,
        color_by=color_by,
        y_axis_mode=y_axis_mode,
        exclude_intervals=exclude_intervals,
        group_label=selected_group_label,
        row_by=row_by,
        custom_color_map=custom_color_map,
        custom_yrange=custom_yrange,
        selected_points=selected_points,
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
        selected_points=selected_points,
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
        selected_points=selected_points,
    )

if fig is None:
    st.warning("Could not generate chart. Check that the selected dimensions have data.")
    st.stop()

finalize_plotly_style(fig)
st.plotly_chart(fig, use_container_width=True, key="main_chart")

# ---------------------------------------------------------------------------
# Click-to-highlight (JMP-style) for Combined Profile chart
# ---------------------------------------------------------------------------
if chart_type == "Combined Profile":
    _highlight_js = """
<script>
(function() {
    function setupClickHighlight() {
        var plotDivs = window.parent.document.querySelectorAll('.js-plotly-plot');
        if (plotDivs.length === 0) {
            setTimeout(setupClickHighlight, 500);
            return;
        }
        var plotDiv = plotDivs[plotDivs.length - 1];
        if (plotDiv._clickHighlightSetup) return;
        plotDiv._clickHighlightSetup = true;

        var highlightedTrace = null;
        var defaultOpacity = 0.45;
        var defaultWidth = 0.7;
        var highlightOpacity = 1.0;
        var highlightWidth = 2.5;
        var dimOpacity = 0.08;
        var dimWidth = 0.4;

        plotDiv.on('plotly_click', function(data) {
            var traceIndex = data.points[0].curveNumber;

            if (highlightedTrace === traceIndex) {
                // Same trace clicked again -- reset all to default
                Plotly.restyle(plotDiv, {'opacity': defaultOpacity, 'line.width': defaultWidth});
                highlightedTrace = null;
            } else {
                // Highlight clicked trace, dim all others
                var nTraces = plotDiv.data.length;
                var opacities = [];
                var widths = [];
                for (var i = 0; i < nTraces; i++) {
                    opacities.push(dimOpacity);
                    widths.push(dimWidth);
                }
                opacities[traceIndex] = highlightOpacity;
                widths[traceIndex] = highlightWidth;
                Plotly.restyle(plotDiv, {'opacity': opacities, 'line.width': widths});
                highlightedTrace = traceIndex;
            }
        });

        plotDiv.on('plotly_doubleclick', function() {
            Plotly.restyle(plotDiv, {'opacity': defaultOpacity, 'line.width': defaultWidth});
            highlightedTrace = null;
        });
    }
    setTimeout(setupClickHighlight, 1000);
})();
</script>
"""
    components.html(_highlight_js, height=0)

# ---------------------------------------------------------------------------
# Summary Statistics – Professional SPC Analytics
# ---------------------------------------------------------------------------



def _render_capability_card(cap):
    """Render a process capability result as styled metrics."""
    # Rating based on Cpk
    cpk = cap.get("Cpk", cap.get("Cpk (upper)", cap.get("Cpk (lower)", None)))
    if cpk is not None:
        if cpk >= 1.67:
            rating, color = "Excellent", "#16A34A"
        elif cpk >= 1.33:
            rating, color = "Good", "#2563EB"
        elif cpk >= 1.0:
            rating, color = "Marginal", "#D97706"
        else:
            rating, color = "Poor", "#DC2626"
    else:
        rating, color = "N/A", "#64748B"

    cols = st.columns(4)
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

    cols2 = st.columns(4)
    if "Sigma Level" in cap:
        cols2[0].metric("Sigma Level", f"{cap['Sigma Level']}σ")
    if "DPMO" in cap:
        cols2[1].metric("DPMO", f"{cap['DPMO']:,}")
    if "Yield %" in cap:
        cols2[2].metric("Yield", f"{cap['Yield %']}%")
    cols2[3].markdown(
        f"<div style='text-align:center;padding:8px;'>"
        f"<span style='font-size:0.8em;color:#64748B;'>Rating</span><br>"
        f"<span style='font-size:1.4em;font-weight:700;color:{color};'>{rating}</span></div>",
        unsafe_allow_html=True,
    )

    # OOS row
    if cap.get("OOS Count", 0) > 0:
        st.warning(f"Out-of-Spec: {cap['OOS Count']} points ({cap['OOS %']}%)")
    else:
        st.success(f"All {cap['n']} points within spec (0 OOS)")


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

        st.markdown(f"### {dno} — {dmeta.description}")

        usl_val = next((v for v in usls if v is not None), None)
        nom_val = next((v for v in nominals if v is not None), None)
        lsl_val = next((v for v in lsls if v is not None), None)

        # Spec limits row
        spec_cols = st.columns(3)
        spec_cols[0].metric("USL", f"{usl_val:.4f}" if usl_val is not None else "N/A")
        spec_cols[1].metric("Nominal", f"{nom_val:.4f}" if nom_val is not None else "N/A")
        spec_cols[2].metric("LSL", f"{lsl_val:.4f}" if lsl_val is not None else "N/A")

        # Combine all measurement columns into one series for overall analysis
        all_values = df_clean[valid_cols].values.flatten()
        all_values = pd.Series(all_values).dropna()

        # --- Tab layout for analyses ---
        tab_cap, tab_anova, tab_trend = st.tabs([
            "Process Capability", "Group Comparison (ANOVA)", "Trend & Shift Detection"
        ])

        # ====== TAB 1: Process Capability ======
        with tab_cap:
            if len(all_values) < 2:
                st.info("Not enough data to calculate process capability.")
            else:
                cap = calc_process_capability(all_values, usl_val, lsl_val)
                if cap:
                    _render_capability_card(cap)

                    # Per-point capability table
                    if len(valid_cols) > 1:
                        st.markdown("**Per-Point Capability**")
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

        # ====== TAB 2: ANOVA Group Comparison ======
        with tab_anova:
            if color_by == "None" or color_by not in df_clean.columns:
                st.info("Select a **Color by** grouping (e.g. Raw Material, Build) to enable group comparison.")
            else:
                groups = df_clean[color_by].fillna("Unknown").astype(str)
                unique_groups = sorted(groups.unique())
                if len(unique_groups) < 2:
                    st.info("Need at least 2 groups for ANOVA comparison.")
                else:
                    # Gather data per group
                    group_data = {}
                    for g in unique_groups:
                        mask = groups == g
                        vals = df_clean.loc[mask, valid_cols].values.flatten()
                        vals = pd.Series(vals).dropna()
                        if len(vals) > 0:
                            group_data[g] = vals

                    if len(group_data) < 2:
                        st.info("Not enough data in groups for ANOVA.")
                    else:
                        # One-way ANOVA
                        f_stat, p_value = scipy_stats.f_oneway(*group_data.values())

                        anova_cols = st.columns(3)
                        anova_cols[0].metric("F-statistic", f"{f_stat:.4f}")
                        anova_cols[1].metric("p-value", f"{p_value:.6f}")
                        sig = "Yes" if p_value < 0.05 else "No"
                        sig_color = "#DC2626" if p_value < 0.05 else "#16A34A"
                        anova_cols[2].markdown(
                            f"<div style='text-align:center;padding:8px;'>"
                            f"<span style='font-size:0.8em;color:#64748B;'>Significant (α=0.05)</span><br>"
                            f"<span style='font-size:1.4em;font-weight:700;color:{sig_color};'>{sig}</span></div>",
                            unsafe_allow_html=True,
                        )

                        if p_value < 0.05:
                            st.warning("Statistically significant difference detected between groups.")
                        else:
                            st.success("No statistically significant difference between groups.")

                        # Group summary table
                        st.markdown("**Group Summary**")
                        summary_rows = []
                        for g, vals in group_data.items():
                            summary_rows.append({
                                "Group": g,
                                "Count": len(vals),
                                "Mean": round(vals.mean(), 6),
                                "Std Dev": round(vals.std(ddof=1), 6),
                                "Min": round(vals.min(), 6),
                                "Max": round(vals.max(), 6),
                                "Range": round(vals.max() - vals.min(), 6),
                            })
                        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

                        # Between-group vs within-group variation
                        grand_mean = all_values.mean()
                        ss_between = sum(len(group_data[g]) * (group_data[g].mean() - grand_mean)**2
                                         for g in group_data)
                        ss_within = sum(((group_data[g] - group_data[g].mean())**2).sum()
                                        for g in group_data)
                        ss_total = ss_between + ss_within
                        if ss_total > 0:
                            var_cols = st.columns(3)
                            var_cols[0].metric("Between-Group SS", f"{ss_between:.4f}")
                            var_cols[1].metric("Within-Group SS", f"{ss_within:.4f}")
                            var_cols[2].metric("% Variation (Between)", f"{ss_between/ss_total*100:.1f}%")

                        # Box plot per group
                        fig_box = go.Figure()
                        for g in unique_groups:
                            if g in group_data:
                                fig_box.add_trace(go.Box(
                                    y=group_data[g].values, name=g,
                                    marker_color=custom_color_map.get(g, "#2563EB"),
                                    boxmean="sd",
                                ))
                        fig_box.update_layout(
                            title="Group Distribution Comparison",
                            yaxis_title="Measured Value",
                            paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                            font=dict(color="#000000"),
                            height=350,
                        )
                        if usl_val is not None:
                            fig_box.add_hline(y=usl_val, line_dash="dash", line_color="red",
                                              annotation_text=f"USL-{usl_val}")
                        if lsl_val is not None:
                            fig_box.add_hline(y=lsl_val, line_dash="dash", line_color="red",
                                              annotation_text=f"LSL-{lsl_val}")
                        st.plotly_chart(fig_box, use_container_width=True,
                                        key=f"anova_box_{dno}")

        # ====== TAB 3: Trend & Shift Detection ======
        with tab_trend:
            if len(all_values) < 9:
                st.info("Need at least 9 data points for trend/shift analysis.")
            else:
                # Nelson Rules
                st.markdown("**Nelson Rules Analysis**")
                violations = nelson_rules(all_values)
                if not violations:
                    st.success("No Nelson rule violations detected — process is stable.")
                else:
                    for rule_name, indices in violations.items():
                        st.warning(f"{rule_name}: {len(indices)} points flagged")

                    # Summary table of violations
                    viol_rows = []
                    for rule_name, indices in violations.items():
                        viol_rows.append({
                            "Rule": rule_name,
                            "Violations": len(indices),
                            "Flagged Indices": str(indices[:20]) + ("..." if len(indices) > 20 else ""),
                        })
                    st.dataframe(pd.DataFrame(viol_rows), use_container_width=True, hide_index=True)

                # CUSUM chart
                st.markdown("**CUSUM (Cumulative Sum) Chart**")
                cusum_pos, cusum_neg, shift_pts = cusum_analysis(all_values, target=nom_val)
                if cusum_pos is not None:
                    fig_cusum = go.Figure()
                    x_idx = list(range(len(cusum_pos)))
                    fig_cusum.add_trace(go.Scatter(
                        x=x_idx, y=cusum_pos, mode="lines", name="CUSUM+",
                        line=dict(color="#2563EB"),
                    ))
                    fig_cusum.add_trace(go.Scatter(
                        x=x_idx, y=cusum_neg, mode="lines", name="CUSUM−",
                        line=dict(color="#DC2626"),
                    ))
                    fig_cusum.add_hline(y=5.0, line_dash="dash", line_color="#94A3B8",
                                        annotation_text="Decision boundary (h=5)")
                    if shift_pts:
                        fig_cusum.add_trace(go.Scatter(
                            x=shift_pts,
                            y=[max(cusum_pos[i], cusum_neg[i]) for i in shift_pts],
                            mode="markers", name="Shift detected",
                            marker=dict(color="#DC2626", size=8, symbol="x"),
                        ))
                    fig_cusum.update_layout(
                        title="CUSUM Control Chart",
                        xaxis_title="Observation", yaxis_title="Cumulative Sum",
                        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                        font=dict(color="#000000"),
                        height=300,
                    )
                    st.plotly_chart(fig_cusum, use_container_width=True,
                                    key=f"cusum_{dno}")

                    if shift_pts:
                        st.warning(f"CUSUM detected potential process shifts at {len(shift_pts)} points.")
                    else:
                        st.success("CUSUM: No significant process shifts detected.")

                # EWMA chart
                st.markdown("**EWMA (Exponentially Weighted Moving Average)**")
                lam = 0.2  # smoothing factor
                ewma = np.zeros(len(all_values))
                ewma[0] = all_values.iloc[0]
                for i in range(1, len(all_values)):
                    ewma[i] = lam * all_values.iloc[i] + (1 - lam) * ewma[i-1]
                overall_mean = all_values.mean()
                overall_std = all_values.std(ddof=1)
                # EWMA control limits
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
                fig_ewma.add_trace(go.Scatter(
                    x=x_idx, y=ewma, mode="lines", name="EWMA",
                    line=dict(color="#2563EB", width=2),
                ))
                fig_ewma.add_trace(go.Scatter(
                    x=x_idx, y=ewma_ucl, mode="lines", name="UCL",
                    line=dict(color="#DC2626", dash="dash", width=1),
                ))
                fig_ewma.add_trace(go.Scatter(
                    x=x_idx, y=ewma_lcl, mode="lines", name="LCL",
                    line=dict(color="#DC2626", dash="dash", width=1),
                ))
                fig_ewma.add_hline(y=overall_mean, line_dash="dot", line_color="#94A3B8",
                                    annotation_text="Center")
                # Flag OOC points
                ooc_ewma = [i for i in range(len(ewma)) if ewma[i] > ewma_ucl[i] or ewma[i] < ewma_lcl[i]]
                if ooc_ewma:
                    fig_ewma.add_trace(go.Scatter(
                        x=ooc_ewma, y=[ewma[i] for i in ooc_ewma],
                        mode="markers", name="Out of Control",
                        marker=dict(color="#DC2626", size=8, symbol="x"),
                    ))
                fig_ewma.update_layout(
                    title="EWMA Control Chart",
                    xaxis_title="Observation", yaxis_title="EWMA Value",
                    paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                    font=dict(color="#000000"),
                    height=300,
                )
                st.plotly_chart(fig_ewma, use_container_width=True,
                                key=f"ewma_{dno}")

                if ooc_ewma:
                    st.warning(f"EWMA: {len(ooc_ewma)} out-of-control points detected.")
                else:
                    st.success("EWMA: Process is in statistical control.")

        st.markdown("---")
