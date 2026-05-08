"""Sidebar layout helpers: chart-type/grouping/Y-axis controls and color pickers."""

import streamlit as st

from spc_viz.charts import get_color_for_group

# ---------------------------------------------------------------------------
# Chart type + grouping + Y-axis controls
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
# Color pickers
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
