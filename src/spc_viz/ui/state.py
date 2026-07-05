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

from spc_viz.parsers.dataset import SpcDataset
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
# Settle gate — defer heavy work until a multi-pick selection stops changing
# ---------------------------------------------------------------------------

SETTLE_SECONDS = 1.2


def _settle_decision(
    prev: tuple | None, value: object, now: float, delay: float
) -> tuple[tuple, float]:
    """Pure timing logic: return (state_to_store, seconds_still_to_wait).

    ``prev`` is the stored (value, first_seen_ts) pair, or None on the very
    first observation — which counts as already settled so a cold start never
    waits.
    """
    if prev is None:
        return (value, now - delay), 0.0
    prev_value, ts = prev
    if value != prev_value:
        return (value, now), delay
    return (prev_value, ts), max(0.0, delay - (now - ts))


def settle(key: str, value: object, delay: float = SETTLE_SECONDS) -> None:
    """Skip the rest of this run until ``value`` has been stable for ``delay``s.

    Streamlit reruns the whole script on every widget click, so picking five
    items in a multiselect used to mean five full parse+render cycles. Placing
    this gate before an expensive stage makes the intermediate reruns cheap:
    while the user is still picking, each click just restarts the quiet
    window; processing starts once — roughly when they close the dropdown.
    """
    import time

    now = time.time()
    state, remaining = _settle_decision(st.session_state.get(key), value, now, delay)
    st.session_state[key] = state
    if remaining > 0:
        time.sleep(remaining)
        st.rerun()


# ---------------------------------------------------------------------------
# Data preparation helper
# ---------------------------------------------------------------------------


def prepare_and_clean(
    dataset: SpcDataset,
    selected_dim_nos: list[str],
) -> tuple[pd.DataFrame, OrderedDict[str, DimensionMeta], list[str]]:
    """Combine data and drop all-NaN rows.

    Returns (df_clean, dim_metas, all_meas_cols).
    Calls ``st.stop()`` and emits a warning when data is missing.
    """
    df, dim_metas = dataset.combined(selected_dim_nos)
    if df is None or dim_metas is None or df.empty:
        st.warning("No data found for selected dimensions.")
        st.stop()
    # st.stop() raises, so this only narrows the Optionals for type checking.
    assert df is not None and dim_metas is not None

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
