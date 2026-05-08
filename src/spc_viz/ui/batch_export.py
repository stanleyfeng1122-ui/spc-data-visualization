"""Sidebar expander for batch chart export — one PNG per selected dimension, zipped."""

from collections import OrderedDict

import streamlit as st

from .chart_view import _build_chart_figure
from .state import prepare_and_clean


def render_batch_export(
    all_dimensions: OrderedDict,
    parsed_files: list,
    controls: dict,
    exclude_intervals: bool,
    selected_points: list | None,
    custom_color_map: dict,
    key_prefix: str = "",
) -> None:
    """Sidebar expander that batch-exports one chart per selected dimension."""
    import io
    import re
    import zipfile

    from streamlit.runtime.scriptrunner import StopException

    dim_display_map = OrderedDict()
    for dno, dmeta in all_dimensions.items():
        label = f"{dno} — {dmeta.description}" if dmeta.description else dno
        dim_display_map[label] = dno

    with st.sidebar.expander("Batch Chart Export", expanded=False):
        batch_dims = st.multiselect(
            "Dimensions to export",
            options=list(dim_display_map.keys()),
            default=[],
            help="Select dimensions. One chart per dimension.",
            key=f"{key_prefix}batch_dims",
        )
        if not batch_dims:
            st.caption("Pick dimensions above, then click Export.")
            return

        export_btn = st.button(
            f"Export {len(batch_dims)} chart{'s' if len(batch_dims) != 1 else ''}",
            key=f"{key_prefix}batch_export_btn",
        )

        if not export_btn:
            return

        # --- Generate charts ---
        ct = controls["chart_type"]
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

            if df_clean is None or df_clean.empty:
                skipped.append(dno)
                continue

            desc = all_dimensions[dno].description or ""
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
                png_bytes = fig.to_image(
                    format="png",
                    width=1400,
                    height=700,
                    scale=2,
                )
            except Exception as e:
                skipped.append(f"{dno} (image error: {e})")
                continue

            safe_desc = re.sub(r"[^\w\s-]", "", desc).strip().replace(" ", "_")
            fname = (
                f"{ct.replace(' ', '_')}_{dno}_{safe_desc}.png"
                if safe_desc
                else f"{ct.replace(' ', '_')}_{dno}.png"
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
