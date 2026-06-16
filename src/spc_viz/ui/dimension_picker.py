"""Dimension selector — preset groups + multiselect of individual dimensions."""

from __future__ import annotations

from collections import OrderedDict

import streamlit as st

from spc_viz.parsers import detect_dimension_groups, get_filtered_dim_meta
from spc_viz.parsers.dimensions import DimensionMeta
from spc_viz.parsers.pairing import is_paired_dim_id


def _dimension_display_label(dno: str, dmeta: DimensionMeta) -> str:
    if is_paired_dim_id(dno):
        return dmeta.description or dno
    return f"{dno} — {dmeta.description}" if dmeta.description else dno


def _dimension_group_label(dno: str, dmeta: DimensionMeta) -> str:
    if is_paired_dim_id(dno):
        return dmeta.description or dno
    return dno.replace("SPC_", "")


def build_dimension_selector(
    all_dimensions: OrderedDict[str, DimensionMeta],
    key_prefix: str = "",
) -> tuple[list[str], str, dict[str, list[str]]]:
    """Render preset + multiselect in the sidebar, return selected dim numbers.

    Returns (selected_dim_nos, selected_group_label, dim_groups).
    """
    dim_groups = detect_dimension_groups(all_dimensions)

    dim_display_map: OrderedDict[str, str] = OrderedDict()
    for dno, dmeta in all_dimensions.items():
        label = _dimension_display_label(dno, dmeta)
        dim_display_map[label] = dno

    dim_no_to_label: dict[str, str] = {v: k for k, v in dim_display_map.items()}
    dim_display_labels: list[str] = list(dim_display_map.keys())

    if not dim_display_labels:
        st.warning("No dimensions found.")
        st.stop()

    st.sidebar.markdown("---")
    group_options = ["Custom"] + list(dim_groups.keys())
    selected_preset: str = st.sidebar.selectbox(
        "Preset",
        options=group_options,
        index=0,
        key=f"{key_prefix}preset",
    )

    default_labels: list[str]
    if selected_preset != "Custom":
        preset_dim_nos = dim_groups[selected_preset]
        default_labels = [dim_no_to_label[dno] for dno in preset_dim_nos if dno in dim_no_to_label]
    else:
        default_labels = [dim_display_labels[0]] if dim_display_labels else []

    dims_key = f"{key_prefix}dims"
    preset_state_key = f"{key_prefix}preset_applied"
    current_selected_labels = st.session_state.get(dims_key, default_labels)
    if not isinstance(current_selected_labels, list):
        current_selected_labels = default_labels
    current_selected_labels = [label for label in current_selected_labels if label in dim_display_map]

    if selected_preset != "Custom":
        preset_signature = (selected_preset, tuple(default_labels))
        if st.session_state.get(preset_state_key) != preset_signature:
            current_selected_labels = default_labels
            st.session_state[preset_state_key] = preset_signature
    elif not current_selected_labels:
        current_selected_labels = default_labels

    st.session_state[dims_key] = current_selected_labels

    selected_dim_labels: list[str] = st.sidebar.multiselect(
        "Dimensions",
        options=dim_display_labels,
        key=dims_key,
        placeholder="Type SPC ID or description",
        help="Click here and type to search inside the dropdown, then select matching dimensions.",
    )
    selected_dim_nos = [dim_display_map[lbl] for lbl in selected_dim_labels]
    selected_group_label = (
        " / ".join(
            _dimension_group_label(dno, all_dimensions[dno])
            for dno in selected_dim_nos
            if dno in all_dimensions
        )
        if selected_dim_nos
        else ""
    )

    if not selected_dim_nos:
        st.info("Select at least one dimension from the sidebar.")
        st.stop()

    return selected_dim_nos, selected_group_label, dim_groups


def build_point_filter(
    all_dimensions: OrderedDict[str, DimensionMeta],
    selected_dim_nos: list[str],
    key_prefix: str = "",
) -> tuple[bool, list[str] | None]:
    """Render exclude-points controls. Returns (exclude_intervals, selected_points)."""
    exclude_intervals: bool = st.sidebar.checkbox(
        "Exclude interval points",
        value=True,
        key=f"{key_prefix}excl",
    )

    all_point_numbers: list[str] = []
    for dno in selected_dim_nos:
        if dno in all_dimensions:
            meta = all_dimensions[dno]
            _cls, pns, _noms, _usls, _lsls = get_filtered_dim_meta(meta, exclude_intervals=False)
            for pn in pns:
                if pn and pn not in all_point_numbers:
                    all_point_numbers.append(pn)

    excluded_points: list[str] = st.sidebar.multiselect(
        "Exclude points",
        options=all_point_numbers,
        default=[],
        help="Pick points to hide. Empty = show all.",
        key=f"{key_prefix}points",
    )
    selected_points: list[str] | None
    if excluded_points:
        selected_points = [p for p in all_point_numbers if p not in excluded_points]
        if not selected_points:
            selected_points = None
    else:
        selected_points = None

    return exclude_intervals, selected_points
