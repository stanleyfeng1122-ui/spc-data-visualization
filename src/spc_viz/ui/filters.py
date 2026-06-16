"""Row-level data filters: narrow the combined dataframe by metadata factor values.

The "Filter" sidebar section lets the user pick one or more factor fields
(Factory, Config, Level, Raw material, ...) and choose which values of each to
display. Filters combine with AND. The default (no factors chosen) leaves the
view unchanged, so adding this control never alters existing behaviour until a
selection is made.

``apply_data_filters`` is a pure, Streamlit-free function so it can be unit
tested directly; ``build_data_filters`` is the thin Streamlit wrapper.
"""

from __future__ import annotations

import streamlit as st
from pandas import DataFrame

# Identity-ish columns that are never useful as a categorical filter.
_NON_FILTER_FIELDS = {"Start Point", "SN"}


def _filterable_fields(parsed_files: list[dict], df: DataFrame) -> list[str]:
    """Metadata fields present in ``df`` with more than one distinct value.

    Mirrors the metadata-column source used by the grouping controls so the
    filter offers the same set of factors. A field with a single distinct
    value is omitted because filtering on it can't change the view.
    """
    available: set[str] = set()
    for pf in parsed_files:
        available.update(pf["meta_columns"])

    fields: list[str] = []
    for col in sorted(available - _NON_FILTER_FIELDS):
        if col in df.columns and df[col].dropna().nunique() > 1:
            fields.append(col)
    return fields


def apply_data_filters(df: DataFrame, filters: dict[str, list[str]]) -> DataFrame:
    """Keep rows whose value is in the selected list for every active field.

    Values are compared as strings with ``NaN`` mapped to ``"Unknown"``,
    matching how the color pickers label groups. Fields with an empty
    selection or absent from ``df`` are ignored. AND across fields.
    """
    out = df
    for field, allowed in filters.items():
        if not allowed or field not in out.columns:
            continue
        vals = out[field].fillna("Unknown").astype(str)
        out = out[vals.isin(allowed)]
    return out.reset_index(drop=True)


def build_data_filters(
    parsed_files: list[dict],
    df_clean: DataFrame,
    key_prefix: str = "",
) -> tuple[DataFrame, dict[str, list[str]]]:
    """Render the Filter section; return narrowed df and the active filter spec.

    Returns ``df_clean`` unchanged (and an empty spec) when no factor is
    chosen. The returned filter spec can be replayed against other dataframes
    (e.g. the per-dimension frames in batch export) via ``apply_data_filters``.
    When active filters leave no rows, emits a warning and stops the page
    rather than rendering an empty chart.
    """
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filter")

    fields = _filterable_fields(parsed_files, df_clean)
    if not fields:
        st.sidebar.caption("No multi-value factors to filter on.")
        return df_clean, {}

    chosen = st.sidebar.multiselect(
        "Filter by",
        options=fields,
        default=[],
        help="Pick one or more factors, then choose which values to display. "
        "Leave empty to show everything.",
        key=f"{key_prefix}filter_fields",
    )

    filters: dict[str, list[str]] = {}
    for field in chosen:
        values = sorted(df_clean[field].fillna("Unknown").astype(str).unique())
        filters[field] = st.sidebar.multiselect(
            f"{field} values",
            options=values,
            default=values,
            key=f"{key_prefix}filter_vals_{field}",
        )

    filtered = apply_data_filters(df_clean, filters)
    if filters and filtered.empty:
        st.warning("No rows match the current filter — adjust the selections in the sidebar.")
        st.stop()
    return filtered, filters
