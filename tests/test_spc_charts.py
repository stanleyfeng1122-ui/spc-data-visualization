"""
SPC App — Programmatic Chart & Feature Tests

Run via:  .venv/bin/python -m pytest tests/test_spc_charts.py -v

These tests validate chart rendering logic WITHOUT Streamlit.
They call chart_utils and shared_ui functions directly with synthetic data
to verify trace structure, spec limits, modes, and batch export plumbing.
"""

import os
import sys
from collections import OrderedDict

import numpy as np
import pandas as pd
import pytest

# Ensure project root is importable
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from chart_utils import (
    build_box_plot,
    build_combined_chart,
    build_histogram,
    finalize_plotly_style,
)
from spc_viz.parsers import DimensionMeta

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _make_dim_meta(
    dim_no: str,
    description: str,
    n_points: int,
    nominal: float = 0.0,
    usl: float = 1.0,
    lsl: float = -1.0,
):
    """Helper to create a DimensionMeta with n_points measurement columns."""
    cols = [f"{dim_no}_P{i}" for i in range(n_points)]
    return cols, DimensionMeta(
        dim_no=dim_no,
        description=description,
        dim_type="Profile" if n_points > 1 else "Non-Profile",
        col_labels=cols,
        point_numbers=[str(i) for i in range(n_points)],
        nominal=[nominal] * n_points,
        tol_max=[usl] * n_points,
        tol_min=[lsl] * n_points,
        usl=[usl] * n_points,
        lsl=[lsl] * n_points,
        col_indices=list(range(1, n_points + 1)),
    )


@pytest.fixture
def multi_point_data():
    """20-point profile dimension, 15 parts, 2 factories."""
    cols, meta = _make_dim_meta("SPC_HG", "Height profile", 20, usl=0.8, lsl=0.0)
    n_parts = 15
    df = pd.DataFrame({c: np.random.uniform(0.1, 0.7, n_parts) for c in cols})
    df["Factory"] = ["FJS"] * 8 + ["LYC"] * 7
    dim_metas = OrderedDict([("SPC_HG", meta)])
    return df, dim_metas


@pytest.fixture
def single_point_data():
    """1-point flatness dimension, 30 parts."""
    cols, meta = _make_dim_meta("SPC_C1", "Datum A Flatness", 1, nominal=0.0, usl=0.7, lsl=0.0)
    n_parts = 30
    df = pd.DataFrame({cols[0]: np.random.uniform(0.05, 0.6, n_parts)})
    df["Factory"] = ["FJS"] * 15 + ["LYC"] * 15
    dim_metas = OrderedDict([("SPC_C1", meta)])
    return df, dim_metas


@pytest.fixture
def common_kwargs():
    """Default kwargs shared across chart builders."""
    return dict(
        color_by="None",
        exclude_intervals=False,
        group_label="Test",
        row_by="None",
        custom_color_map=None,
        selected_points=None,
    )


# ---------------------------------------------------------------------------
# 1. Server Health — imports work
# ---------------------------------------------------------------------------


class TestImports:
    def test_shared_ui_imports(self):
        import shared_ui

        assert hasattr(shared_ui, "build_and_render_chart")
        assert hasattr(shared_ui, "_build_chart_figure")
        assert hasattr(shared_ui, "render_batch_export")

    def test_chart_utils_imports(self):
        import chart_utils

        assert hasattr(chart_utils, "build_combined_chart")
        assert hasattr(chart_utils, "build_box_plot")
        assert hasattr(chart_utils, "build_histogram")

    def test_spc_parser_imports(self):
        from spc_viz import parsers as spc_parser

        assert hasattr(spc_parser, "parse_excel_multi")
        assert hasattr(spc_parser, "DimensionMeta")

    def test_kaleido_available(self):
        import kaleido

        # kaleido 1.2+ may not expose __version__ at top level
        assert kaleido is not None


# ---------------------------------------------------------------------------
# 2. Data Loading — local xlsx files parse
# ---------------------------------------------------------------------------


class TestDataLoading:
    def test_xlsx_files_exist(self):
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xlsx = [f for f in os.listdir(root) if f.endswith(".xlsx") and not f.startswith("~$")]
        assert len(xlsx) > 0, "No .xlsx files in project directory"

    def test_parse_first_file(self):
        from spc_viz.parsers import parse_excel_multi

        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xlsx = sorted(
            [f for f in os.listdir(root) if f.endswith(".xlsx") and not f.startswith("~$")]
        )
        fpath = os.path.join(root, xlsx[0])
        results = parse_excel_multi(fpath)
        assert len(results) > 0, "First xlsx parsed to zero results"
        assert results[0].dimensions, "No dimensions found in first file"


# ---------------------------------------------------------------------------
# 3. Combined Profile — multi-point
# ---------------------------------------------------------------------------


class TestCombinedProfileMultiPoint:
    def test_returns_figure(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        assert fig is not None, "build_combined_chart returned None"

    def test_traces_use_lines_mode(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        data_traces = [t for t in fig.data if t.type == "scattergl"]
        assert len(data_traces) > 0, "No scattergl traces"
        assert data_traces[0].mode == "lines", f"Expected lines, got {data_traces[0].mode}"

    def test_traces_have_multiple_x_points(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        data_traces = [t for t in fig.data if t.type == "scattergl"]
        assert len(data_traces[0].x) == 20, f"Expected 20 x-points, got {len(data_traces[0].x)}"

    def test_has_line_width(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        t = [t for t in fig.data if t.type == "scattergl"][0]
        assert t.line is not None and t.line.width == 0.7


# ---------------------------------------------------------------------------
# 4. Combined Profile — single-point (flatness fix)
# ---------------------------------------------------------------------------


class TestSinglePointFix:
    def test_single_point_uses_markers(self, single_point_data, common_kwargs):
        df, dim_metas = single_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_C1"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        data_traces = [t for t in fig.data if t.type == "scattergl"]
        assert len(data_traces) > 0
        assert data_traces[0].mode == "markers", f"Expected markers, got {data_traces[0].mode}"

    def test_single_point_marker_size(self, single_point_data, common_kwargs):
        df, dim_metas = single_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_C1"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        t = [t for t in fig.data if t.type == "scattergl"][0]
        assert t.marker is not None and t.marker.size == 6

    def test_single_point_has_one_x(self, single_point_data, common_kwargs):
        df, dim_metas = single_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_C1"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        t = [t for t in fig.data if t.type == "scattergl"][0]
        assert len(t.x) == 1


# ---------------------------------------------------------------------------
# 5. Spec Limits — USL/LSL lines and tolerance band
# ---------------------------------------------------------------------------


class TestSpecLimits:
    def _build_fig(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        finalize_plotly_style(fig)
        return fig

    def test_has_usl_line(self, multi_point_data, common_kwargs):
        fig = self._build_fig(multi_point_data, common_kwargs)
        shapes = fig.layout.shapes or []
        h_lines = [s for s in shapes if s.type == "line" and s.y0 == s.y1]
        assert len(h_lines) > 0, "No horizontal lines (USL/LSL) found"

    def test_has_tolerance_band(self, multi_point_data, common_kwargs):
        fig = self._build_fig(multi_point_data, common_kwargs)
        shapes = fig.layout.shapes or []
        rects = [s for s in shapes if s.type == "rect"]
        assert len(rects) > 0, "No tolerance band (hrect) found"


# ---------------------------------------------------------------------------
# 6. Box Plot
# ---------------------------------------------------------------------------


class TestBoxPlot:
    def test_returns_figure(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_box_plot(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        assert fig is not None

    def test_has_box_traces(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_box_plot(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        box_traces = [t for t in fig.data if t.type == "box"]
        assert len(box_traces) > 0, "No box traces in Box Plot"


# ---------------------------------------------------------------------------
# 7. Histogram
# ---------------------------------------------------------------------------


class TestHistogram:
    def test_returns_figure(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_histogram(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            nbins=30,
            **common_kwargs,
        )
        assert fig is not None

    def test_has_histogram_traces(self, multi_point_data, common_kwargs):
        df, dim_metas = multi_point_data
        fig = build_histogram(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            nbins=30,
            **common_kwargs,
        )
        hist_traces = [t for t in fig.data if t.type == "histogram"]
        assert len(hist_traces) > 0, "No histogram traces"


# ---------------------------------------------------------------------------
# 8. Batch Export — figure building pipeline
# ---------------------------------------------------------------------------


class TestBatchExport:
    def test_build_chart_figure_combined(self, multi_point_data, common_kwargs):
        from shared_ui import _build_chart_figure

        df, dim_metas = multi_point_data
        controls = {
            "chart_type": "Combined Profile",
            "color_by": "None",
            "row_by": "None",
            "section_by_fields": ["Factory"],
            "y_axis_mode": "Measurement values",
            "custom_yrange": None,
            "hist_nbins": 30,
        }
        fig = _build_chart_figure(
            df,
            dim_metas,
            ["SPC_HG"],
            controls,
            None,
            False,
            "Test",
            None,
        )
        assert fig is not None
        assert len(fig.data) > 0

    def test_build_chart_figure_box(self, single_point_data, common_kwargs):
        from shared_ui import _build_chart_figure

        df, dim_metas = single_point_data
        controls = {
            "chart_type": "Box Plot",
            "color_by": "None",
            "row_by": "None",
            "section_by_fields": ["Factory"],
            "y_axis_mode": "Measurement values",
            "custom_yrange": None,
            "hist_nbins": 30,
        }
        fig = _build_chart_figure(
            df,
            dim_metas,
            ["SPC_C1"],
            controls,
            None,
            False,
            "Test",
            None,
        )
        assert fig is not None

    def test_figure_to_image(self, multi_point_data, common_kwargs):
        """Verify kaleido can convert a figure to PNG bytes."""
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        finalize_plotly_style(fig)
        png_bytes = fig.to_image(format="png", width=1400, height=700, scale=2)
        assert isinstance(png_bytes, bytes)
        assert len(png_bytes) > 1000, f"PNG too small ({len(png_bytes)} bytes)"
        # PNG magic bytes
        assert png_bytes[:4] == b"\x89PNG"


# ---------------------------------------------------------------------------
# 9. Sidebar Controls — color-by grouping
# ---------------------------------------------------------------------------


class TestColorGrouping:
    def test_color_by_factory(self, multi_point_data):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            color_by="Factory",
            y_axis_mode="Measurement values",
            custom_yrange=None,
            exclude_intervals=False,
            group_label="Test",
            row_by="None",
            custom_color_map=None,
            selected_points=None,
        )
        # Should have traces with different colors for FJS and LYC
        colors = set()
        for t in fig.data:
            if t.type == "scattergl" and t.line:
                colors.add(t.line.color)
            elif t.type == "scattergl" and t.marker:
                colors.add(t.marker.color)
        assert len(colors) >= 2, f"Expected 2+ colors for Factory grouping, got {len(colors)}"

    def test_color_by_none(self, multi_point_data):
        df, dim_metas = multi_point_data
        fig = build_combined_chart(
            df=df,
            dim_metas=dim_metas,
            dim_nos=["SPC_HG"],
            section_by_fields=["Factory"],
            color_by="None",
            y_axis_mode="Measurement values",
            custom_yrange=None,
            exclude_intervals=False,
            group_label="Test",
            row_by="None",
            custom_color_map=None,
            selected_points=None,
        )
        colors = set()
        for t in fig.data:
            if t.type == "scattergl" and t.line:
                colors.add(t.line.color)
        assert len(colors) == 1, f"Expected 1 color for None grouping, got {len(colors)}"
