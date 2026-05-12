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

from collections import OrderedDict

import streamlit as st
import streamlit.components.v1 as components

from spc_viz.parsers import _open_workbook, parse_excel_multi
from spc_viz.theme import inject_theme
from spc_viz.ui import (
    build_and_render_chart,
    build_chart_controls,
    build_color_pickers,
    build_dimension_selector,
    build_point_filter,
    prepare_and_clean,
    render_batch_export,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SPC Data Visualization",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_theme()

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
    st.info("Upload one or more .xlsx CPK data files using the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Sheet selector – scan all files, let user enable sheets to parse
# ---------------------------------------------------------------------------
import io as _io
import warnings as _warnings

_NON_DATA_PREFIXES = ("BoxPlotCht", "Histo ")
_NON_DATA_EXACT = {"Histo Pivot", "Histo Listbox", "Histo Curve"}

_file_sheets = {}  # fname -> [actual sheet names]
_display_to_actual = {}  # display_name -> {fname: actual_name}
_display_sheets = []  # ordered unique display names

for _uf in uploaded_files:
    _raw = _uf.getvalue()
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore")
        _wb, _ = _open_workbook(_io.BytesIO(_raw))
    _sheets = [
        s
        for s in _wb.sheetnames
        if s not in _NON_DATA_EXACT and not any(s.startswith(p) for p in _NON_DATA_PREFIXES)
    ]
    _file_sheets[_uf.name] = _sheets

    for _sn in _sheets:
        _key = _sn.lower().strip()
        _existing = next((d for d in _display_sheets if d.lower().strip() == _key), None)
        if _existing is None:
            _display_sheets.append(_sn)
            _display_to_actual[_sn] = {_uf.name: _sn}
        else:
            if _existing not in _display_to_actual:
                _display_to_actual[_existing] = {}
            _display_to_actual[_existing][_uf.name] = _sn
    _wb.close()

enabled_sheets = st.sidebar.multiselect(
    "Sheets to parse",
    options=_display_sheets,
    default=_display_sheets,
    help="Enable sheets to include. Dimensions with the same name merge across files.",
)

if not enabled_sheets:
    st.info("Select at least one sheet to parse.")
    st.stop()


def _get_actual_sheets_for_file(fname, enabled):
    """Return list of actual sheet names in this file that match enabled display names."""
    actual = []
    for disp_name in enabled:
        mapping = _display_to_actual.get(disp_name, {})
        if fname in mapping:
            actual.append(mapping[fname])
        else:
            for fs in _file_sheets.get(fname, []):
                if fs.lower().strip() == disp_name.lower().strip() and fs not in actual:
                    actual.append(fs)
    return actual


with st.sidebar.expander("Sheet / File map", expanded=False):
    for _fname, _sheets in _file_sheets.items():
        _short = _fname[:30] + "..." if len(_fname) > 30 else _fname
        _active = _get_actual_sheets_for_file(_fname, enabled_sheets)
        if _active:
            st.markdown(
                f"`{_short}`  \n"
                f"<span style='font-size:0.7rem;color:#737373;'>"
                f"{', '.join(_active)}</span>",
                unsafe_allow_html=True,
            )

# ---------------------------------------------------------------------------
# Parse uploaded files
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner="Parsing Excel files...")
def _parse_file_sheets(file_bytes: bytes, filename: str, sheet_names: tuple) -> list:
    """Parse specific sheets from a file and return list of dicts."""
    import io

    results = []
    for sn in sheet_names:
        try:
            buf = io.BytesIO(file_bytes)
            buf.name = filename
            parsed_list = parse_excel_multi(buf, sheet_name=sn)
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
        except Exception:
            pass
    return results


parsed_files = []
for uf in uploaded_files:
    try:
        raw = uf.getvalue()
        _to_parse = tuple(_get_actual_sheets_for_file(uf.name, enabled_sheets))
        if _to_parse:
            results = _parse_file_sheets(raw, uf.name, _to_parse)
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
        n_dims = len(pf["dimensions"])
        factory = pf.get("factory", "?")
        sheet_label = f" [{pf['sheet_name']}]" if pf.get("sheet_name") else ""
        meta_info = f"{n_dims} dims, {n_rows} rows"
        if pf["data"] is not None and "CFG" in pf["data"].columns:
            cfgs = ", ".join(sorted(pf["data"]["CFG"].dropna().unique().astype(str)[:5]))
            meta_info += f", CFG: {cfgs}"
        st.markdown(
            f"`{pf['filename'][:35]}...`{sheet_label}  \n"
            f"<span style='font-size:0.7rem;color:#737373;'>{meta_info}</span>",
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
# Shared controls (dimension selection, chart type, grouping, etc.)
# ---------------------------------------------------------------------------
KP = "main_"  # key prefix

selected_dim_nos, selected_group_label, _ = build_dimension_selector(all_dimensions, key_prefix=KP)
exclude_intervals, selected_points = build_point_filter(
    all_dimensions, selected_dim_nos, key_prefix=KP
)
controls = build_chart_controls(parsed_files, key_prefix=KP)

# ---------------------------------------------------------------------------
# Main content area
# ---------------------------------------------------------------------------
st.title("SPC Data Visualization")

df_clean, dim_metas, _ = prepare_and_clean(parsed_files, selected_dim_nos)

custom_color_map = build_color_pickers(df_clean, controls.color_by, key_prefix=KP)

fig = build_and_render_chart(
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
# Click-to-highlight (JMP-style) for Combined Profile chart
# ---------------------------------------------------------------------------
if controls.chart_type == "Combined Profile":
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
                Plotly.restyle(plotDiv, {'opacity': defaultOpacity, 'line.width': defaultWidth});
                highlightedTrace = null;
            } else {
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
