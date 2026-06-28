"""Sidebar expander for batch chart export — one PNG per selected dimension, zipped."""

from __future__ import annotations

from collections import OrderedDict

import streamlit as st

from spc_viz.parsers.dimensions import DimensionMeta
from spc_viz.parsers.pairing import is_paired_dim_id

from .chart_view import _build_chart_figure
from .dimension_picker import build_display_map
from .filters import apply_data_filters
from .state import ChartControls, prepare_and_clean


def render_batch_export(
    all_dimensions: OrderedDict[str, DimensionMeta],
    parsed_files: list[dict],
    controls: ChartControls,
    exclude_intervals: bool,
    selected_points: list[str] | None,
    custom_color_map: dict[str, str],
    key_prefix: str = "",
    data_filters: dict[str, list[str]] | None = None,
) -> None:
    """Sidebar expander that batch-exports one chart per selected dimension.

    ``data_filters`` is the active sidebar filter spec; when provided, each
    exported chart is narrowed to the same factor values as the on-screen
    chart so the ZIP matches what the user is looking at.
    """
    import io
    import re
    import zipfile

    from streamlit.runtime.scriptrunner import StopException

    dim_display_map: OrderedDict[str, str] = OrderedDict(
        (label, dno) for dno, label in build_display_map(all_dimensions).items()
    )

    with st.sidebar.expander("Batch Chart Export", expanded=False):
        batch_dims: list[str] = st.multiselect(
            "Dimensions to export",
            options=list(dim_display_map.keys()),
            default=[],
            help="Select dimensions. One chart per dimension.",
            key=f"{key_prefix}batch_dims",
        )
        if not batch_dims:
            st.caption("Pick dimensions above, then click Export.")
            return

        export_btn: bool = st.button(
            f"Export {len(batch_dims)} chart{'s' if len(batch_dims) != 1 else ''}",
            key=f"{key_prefix}batch_export_btn",
        )

        if not export_btn:
            return

        # --- Generate charts ---
        ct = controls.chart_type
        progress = st.progress(0, text="Preparing export…")
        images: list[tuple[str, bytes]] = []
        skipped: list[str] = []

        for idx, label in enumerate(batch_dims):
            dno = dim_display_map[label]
            progress.progress(
                (idx) / len(batch_dims),
                text=f"Generating {dno} ({idx + 1}/{len(batch_dims)})…",
            )

            # Build data for this single dimension
            try:
                df_clean, dim_metas, _ = prepare_and_clean(parsed_files, [dno])
            except (StopException, Exception):
                skipped.append(dno)
                continue

            # Apply the same sidebar filter the on-screen chart uses.
            if data_filters:
                df_clean = apply_data_filters(df_clean, data_filters)

            if df_clean is None or df_clean.empty:
                skipped.append(dno)
                continue

            desc = all_dimensions[dno].description or ""
            if is_paired_dim_id(dno):
                group_label = desc or dno
            else:
                group_label = f"{dno.replace('SPC_', '')} — {desc}" if desc else dno

            fig = _build_chart_figure(
                df_clean,
                dim_metas,
                [dno],
                controls,
                custom_color_map,
                exclude_intervals,
                group_label,
                selected_points,
            )
            if fig is None:
                skipped.append(dno)
                continue

            # Convert to PNG
            try:
                png_bytes: bytes = fig.to_image(
                    format="png",
                    width=1400,
                    height=700,
                    scale=2,
                )
            except Exception as e:
                skipped.append(f"{dno} (image error: {e})")
                continue

            safe_dno = re.sub(r"[^\w\s-]", "", dno.replace("SPC_", "")).strip().replace(" ", "_")
            safe_desc = re.sub(r"[^\w\s-]", "", desc).strip().replace(" ", "_")
            fname = (
                f"{ct.replace(' ', '_')}_{safe_dno}_{safe_desc}.png"
                if safe_desc
                else f"{ct.replace(' ', '_')}_{safe_dno}.png"
            )
            images.append((fname, png_bytes))

        progress.progress(1.0, text="Done!")

        if skipped:
            st.warning(f"Skipped {len(skipped)} dimension(s): {', '.join(skipped)}")

        if not images:
            st.error("No charts could be generated.")
            return

        # Bundle into ZIP
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, png_data in images:
                zf.writestr(fname, png_data)
        buf.seek(0)

        st.download_button(
            label=f"Download {len(images)} chart{'s' if len(images) != 1 else ''} (ZIP)",
            data=buf.getvalue(),
            file_name=f"SPC_Charts_{ct.replace(' ', '_')}.zip",
            mime="application/zip",
            key=f"{key_prefix}batch_download",
        )
