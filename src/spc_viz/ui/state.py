"""Session state helpers, ChartControls dataclass, and data preparation utilities.

Each function takes a *key_prefix* so multiple pages can coexist in the
same Streamlit session without widget-key collisions.
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from typing import Literal

import pandas as pd
import streamlit as st

from spc_viz.charts import prepare_combined_data
from spc_viz.parsers.dimensions import DimensionMeta

# ---------------------------------------------------------------------------
# ChartControls — replaces the ad-hoc dict passed between sidebar and charts
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ChartControls:
    """Immutable snapshot of chart-type, grouping, and Y-axis widget state.

    ``frozen=True`` means instances are hashable and cannot be mutated.
    Use ``dataclasses.replace(controls, field=new_value)`` if a copy with
    a changed field is ever needed.
    """

    chart_type: Literal["Combined Profile", "Box Plot", "Histogram", "Range Envelope"]
    color_by: str  # column name or sentinel "None"
    section_by_fields: list[str]
    row_by: str  # column name or sentinel "None"
    y_axis_mode: Literal["Measurement values", "Deviation from Nominal"]
    custom_yrange: list[float] | None
    hist_nbins: int
    show_average_line: bool = False


# ---------------------------------------------------------------------------
# Data preparation helper
# ---------------------------------------------------------------------------


def prepare_and_clean(
    parsed_files: list[dict],
    selected_dim_nos: list[str],
) -> tuple[pd.DataFrame, OrderedDict[str, DimensionMeta], list[str]]:
    """Combine data and drop all-NaN rows.

    Returns (df_clean, dim_metas, all_meas_cols).
    Calls ``st.stop()`` and emits a warning when data is missing.
    """
    df, dim_metas = prepare_combined_data(parsed_files, selected_dim_nos)
    if df is None or dim_metas is None or df.empty:
        st.warning("No data found for selected dimensions.")
        st.stop()

    all_meas_cols: list[str] = []
    for dno in selected_dim_nos:
        if dno in dim_metas:
            all_meas_cols.extend([c for c in dim_metas[dno].col_labels if c in df.columns])

    df_clean = (
        df.dropna(subset=all_meas_cols, how="all").reset_index(drop=True) if all_meas_cols else df
    )
    if df_clean.empty:
        st.warning("No measurement data for selected dimensions.")
        st.stop()

    return df_clean, dim_metas, all_meas_cols
