"""Dimension selector — preset groups + multiselect of individual dimensions."""

from collections import OrderedDict

import streamlit as st

from spc_viz.parsers import detect_dimension_groups, get_filtered_dim_meta


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
