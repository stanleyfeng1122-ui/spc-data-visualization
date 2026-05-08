"""
Quick Test Page — Auto-loads local .xlsx files for fast iteration.

No file uploading needed. All Excel files in the project directory are
parsed automatically so you can immediately verify chart and analysis
behaviour after code changes.
"""

import os
import sys
from collections import OrderedDict

import streamlit as st

# Ensure project root is on the path so we can import siblings
_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from shared_ui import (
    build_and_render_chart,
    build_chart_controls,
    build_color_pickers,
    build_dimension_selector,
    build_point_filter,
    prepare_and_clean,
    render_batch_export,
    render_summary_statistics,
)
from spc_parser import parse_excel_multi
from spc_viz.theme import FONT_MONO, TEXT_MUTED, inject_theme

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
    xlsx_files = sorted(
        [f for f in os.listdir(data_dir) if f.endswith(".xlsx") and not f.startswith("~$")]
    )
    for fname in xlsx_files:
        fpath = os.path.join(data_dir, fname)
        try:
            parsed_list = parse_excel_multi(fpath, sheet_name=sheet)
            for parsed in parsed_list:
                results.append(
                    {
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
                )
        except Exception as e:
            st.sidebar.error(f"Error parsing {fname}: {e}")
    return results


def _discover_sheets(data_dir: str):
    """Read sheet names from all .xlsx files in the directory."""
    import openpyxl

    all_sheets = []
    _NON_DATA_PREFIXES = ("BoxPlotCht", "Histo ")
    _NON_DATA_EXACT = {"Histo Pivot", "Histo Listbox", "Histo Curve"}
    xlsx_files = sorted(
        [f for f in os.listdir(data_dir) if f.endswith(".xlsx") and not f.startswith("~$")]
    )
    for fname in xlsx_files:
        fpath = os.path.join(data_dir, fname)
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True, keep_links=False)
            for sn in wb.sheetnames:
                if (
                    sn not in all_sheets
                    and sn not in _NON_DATA_EXACT
                    and not any(sn.startswith(p) for p in _NON_DATA_PREFIXES)
                ):
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

# ---------------------------------------------------------------------------
# Shared controls
# ---------------------------------------------------------------------------
KP = "qt_"

selected_dim_nos, selected_group_label, _ = build_dimension_selector(all_dimensions, key_prefix=KP)
exclude_intervals, selected_points = build_point_filter(
    all_dimensions, selected_dim_nos, key_prefix=KP
)
controls = build_chart_controls(parsed_files, key_prefix=KP)

# ---------------------------------------------------------------------------
# MAIN AREA — chart + analysis
# ---------------------------------------------------------------------------

# Header
hdr_left, hdr_right = st.columns([3, 1])
with hdr_left:
    st.markdown(
        f"<h1 style='margin:0;padding:0;font-size:1.3rem;'>{selected_group_label or 'SPC Analysis'}</h1>"
        f"<span style='font-size:0.75rem;color:{TEXT_MUTED};font-family:{FONT_MONO};'>"
        f"{len(parsed_files)} file{'s' if len(parsed_files) != 1 else ''} / "
        f"{controls['chart_type']} / {controls['color_by']}"
        f"</span>",
        unsafe_allow_html=True,
    )

df_clean, dim_metas, _ = prepare_and_clean(parsed_files, selected_dim_nos)

custom_color_map = build_color_pickers(df_clean, controls["color_by"], key_prefix=KP)

build_and_render_chart(
    df_clean,
    dim_metas,
    selected_dim_nos,
    controls,
    custom_color_map,
    exclude_intervals,
    selected_group_label,
    selected_points,
    key_prefix=KP,
)

# ---------------------------------------------------------------------------
# Summary Statistics
# ---------------------------------------------------------------------------
render_summary_statistics(
    df_clean,
    dim_metas,
    selected_dim_nos,
    exclude_intervals=exclude_intervals,
    color_by=controls["color_by"],
    custom_color_map=custom_color_map,
    key_prefix=KP,
)

# ---------------------------------------------------------------------------
# Batch Chart Export
# ---------------------------------------------------------------------------
render_batch_export(
    all_dimensions,
    parsed_files,
    controls,
    exclude_intervals,
    selected_points,
    custom_color_map,
    key_prefix=KP,
)
