"""Session state helpers and data preparation utilities.

Each function takes a *key_prefix* so multiple pages can coexist in the
same Streamlit session without widget-key collisions.
"""

import streamlit as st

from spc_viz.charts import prepare_combined_data


def prepare_and_clean(parsed_files, selected_dim_nos):
    """Combine data and drop all-NaN rows. Returns (df_clean, dim_metas, all_meas_cols)."""
    df, dim_metas = prepare_combined_data(parsed_files, selected_dim_nos)
    if df is None or dim_metas is None or df.empty:
        st.warning("No data found for selected dimensions.")
        st.stop()

    all_meas_cols = []
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
