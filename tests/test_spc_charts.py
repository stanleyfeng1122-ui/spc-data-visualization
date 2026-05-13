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

from spc_viz.charts import (
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
        from spc_viz import ui

        assert hasattr(ui, "build_and_render_chart")
        assert hasattr(ui, "_build_chart_figure")
        assert hasattr(ui, "render_batch_export")

    def test_chart_utils_imports(self):
        from spc_viz import charts

        assert hasattr(charts, "build_combined_chart")
        assert hasattr(charts, "build_box_plot")
        assert hasattr(charts, "build_histogram")

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
    def _examples_dir(self):
        from spc_viz.config.paths import EXAMPLES_DIR
        return str(EXAMPLES_DIR)

    def test_xlsx_files_exist(self):
        root = self._examples_dir()
        xlsx = [f for f in os.listdir(root) if f.endswith(".xlsx") and not f.startswith("~$")]
        assert len(xlsx) > 0, f"No .xlsx files in {root}"

    def test_parse_first_file(self):
        from spc_viz.parsers import parse_excel_multi

        root = self._examples_dir()
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
        from spc_viz.ui import ChartControls, _build_chart_figure

        df, dim_metas = multi_point_data
        controls = ChartControls(
            chart_type="Combined Profile",
            color_by="None",
            row_by="None",
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            hist_nbins=30,
        )
        fig = _build_chart_figure(
            df,
            dim_metas,
            ["SPC_HG"],
            controls,
            {},
            False,
            "Test",
            None,
        )
        assert fig is not None
        assert len(fig.data) > 0

    def test_build_chart_figure_box(self, single_point_data, common_kwargs):
        from spc_viz.ui import ChartControls, _build_chart_figure

        df, dim_metas = single_point_data
        controls = ChartControls(
            chart_type="Box Plot",
            color_by="None",
            row_by="None",
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            hist_nbins=30,
        )
        fig = _build_chart_figure(
            df,
            dim_metas,
            ["SPC_C1"],
            controls,
            {},
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


# ---------------------------------------------------------------------------
# 10. Factory detection (metadata.detect_factory) — regression for regex fix
# ---------------------------------------------------------------------------

from spc_viz.parsers.dimensions import ParsedFile
from spc_viz.parsers.metadata import detect_factory


def _pf(filename, data=None):
    df = data if data is not None else pd.DataFrame({"A": [1, 2]})
    pf = ParsedFile(filename=filename, sheet_name="Sheet1", data=df)
    detect_factory(pf, filename)
    return pf


class TestFactoryDetection:
    def test_underscore_separated(self):
        assert _pf("FX_K116_P1.xlsx").factory == "FX"

    def test_space_separated(self):
        # The bug: filename.split("_") broke on space-separated names.
        assert _pf("LK X3745 DH P1 PP&AP CORR CPK&100% Data.xlsx").factory == "LK"

    def test_vendor_serial_number_column_wins(self):
        df = pd.DataFrame({"Vendor Serial Number": ["ABC", "ABC", "XYZ"]})
        assert _pf("FX_K116.xlsx", df).factory == "ABC"

    def test_sn_column_prefix_used(self):
        df = pd.DataFrame({"SN": ["FJS123", "FJS456"]})
        assert _pf("LK_K116.xlsx", df).factory == "FJS"

    def test_weird_lowercase_filename_no_match(self):
        assert _pf("data_export.xlsx").factory is None

    def test_two_letter_factory(self):
        assert _pf("TY_K116.xlsx").factory == "TY"

    def test_three_letter_factory(self):
        assert _pf("TRM_X3083.xlsx").factory == "TRM"

    def test_four_letter_factory(self):
        assert _pf("FXJS_X3744.xlsx").factory == "FXJS"


# ---------------------------------------------------------------------------
# 11. compute_sections
# ---------------------------------------------------------------------------

from spc_viz.charts.base import (
    COLOR_PALETTE,
    compute_row_groups,
    compute_sections,
    get_color_for_group,
)


class TestComputeSections:
    def _df(self):
        return pd.DataFrame({"x": [1, 2, 3]})

    def test_empty_fields_all(self):
        s = compute_sections(self._df(), [])
        assert list(s) == ["All", "All", "All"]

    def test_factory_field_with_column(self):
        df = self._df()
        df["_factory"] = ["FX", "FX", "LK"]
        s = compute_sections(df, ["Factory"])
        assert set(s) == {"FX", "LK"}

    def test_factory_field_no_column(self):
        s = compute_sections(self._df(), ["Factory"])
        assert list(s) == ["?", "?", "?"]

    def test_source_file_field(self):
        df = self._df()
        df["_source_file"] = ["a.xlsx", "a.xlsx", "b.xlsx"]
        s = compute_sections(df, ["Source File"])
        assert set(s) == {"a.xlsx", "b.xlsx"}

    def test_multiple_fields_concat(self):
        df = self._df()
        df["_factory"] = ["FX", "FX", "LK"]
        df["Build"] = ["P1", "P2", "P1"]
        s = compute_sections(df, ["Factory", "Build"])
        assert list(s) == ["FX P1", "FX P2", "LK P1"]

    def test_missing_field_all_q(self):
        s = compute_sections(self._df(), ["DoesNotExist"])
        assert list(s) == ["?", "?", "?"]


# ---------------------------------------------------------------------------
# 12. compute_row_groups
# ---------------------------------------------------------------------------


class TestComputeRowGroups:
    def _df(self):
        return pd.DataFrame({"Color": ["Red", "Blue", "Red"]})

    def test_none_all(self):
        s = compute_row_groups(self._df(), "None")
        assert list(s) == ["All", "All", "All"]

    def test_existing_column(self):
        s = compute_row_groups(self._df(), "Color")
        assert set(s) == {"Red", "Blue"}

    def test_missing_column_all(self):
        s = compute_row_groups(self._df(), "Nope")
        assert list(s) == ["All", "All", "All"]


# ---------------------------------------------------------------------------
# 13. get_color_for_group
# ---------------------------------------------------------------------------


class TestGetColorForGroup:
    def test_distinct_colors(self):
        c0, c1, c2 = get_color_for_group(0), get_color_for_group(1), get_color_for_group(2)
        assert len({c0, c1, c2}) == 3

    def test_wraps_around(self):
        n = len(COLOR_PALETTE)
        assert get_color_for_group(n) == get_color_for_group(0)
        assert get_color_for_group(n + 3) == get_color_for_group(3)

    def test_returns_hex_string(self):
        c = get_color_for_group(0)
        assert isinstance(c, str)
        assert c.startswith("#") and len(c) == 7


# ---------------------------------------------------------------------------
# 14. ChartControls dataclass
# ---------------------------------------------------------------------------

import dataclasses


def _controls(**overrides):
    base = dict(
        chart_type="Combined Profile",
        color_by="None",
        section_by_fields=["Factory"],
        row_by="None",
        y_axis_mode="Measurement values",
        custom_yrange=None,
        hist_nbins=30,
    )
    base.update(overrides)
    from spc_viz.ui import ChartControls
    return ChartControls(**base)


class TestChartControls:
    def test_fields_accessible(self):
        c = _controls()
        assert c.chart_type == "Combined Profile"
        assert c.color_by == "None"
        assert c.section_by_fields == ["Factory"]
        assert c.row_by == "None"
        assert c.y_axis_mode == "Measurement values"
        assert c.custom_yrange is None
        assert c.hist_nbins == 30

    def test_frozen_mutation_raises(self):
        c = _controls()
        with pytest.raises(dataclasses.FrozenInstanceError):
            c.chart_type = "Box Plot"  # type: ignore[misc]

    def test_replace_creates_copy(self):
        c = _controls()
        c2 = dataclasses.replace(c, chart_type="Histogram", hist_nbins=50)
        assert c.chart_type == "Combined Profile"
        assert c2.chart_type == "Histogram"
        assert c2.hist_nbins == 50

    def test_equality(self):
        assert _controls() == _controls()

    def test_chart_type_literal_values(self):
        for ct in ("Combined Profile", "Box Plot", "Histogram"):
            assert _controls(chart_type=ct).chart_type == ct


# ---------------------------------------------------------------------------
# 15. Parser edge cases (real fixtures)
# ---------------------------------------------------------------------------

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")
FX_FIXTURE = os.path.join(FIXTURES_DIR, "FX_X3745.xlsx")
LK_FIXTURE = os.path.join(FIXTURES_DIR, "LK_X3745.xlsx")


class TestParserEdgeCases:
    def test_nonexistent_file_raises(self):
        from spc_viz.parsers import parse_excel_multi

        with pytest.raises(FileNotFoundError):
            parse_excel_multi(os.path.join(FIXTURES_DIR, "no_such_file.xlsx"))

    def test_nonexistent_sheet_autodetects(self):
        from spc_viz.parsers import parse_excel_multi

        # A sheet name that doesn't exist falls through to auto-detect mode,
        # which scans all data sheets and returns >= 1 ParsedFile.
        results = parse_excel_multi(FX_FIXTURE, sheet_name="ZZZ_NoSuchSheet")
        assert isinstance(results, list)
        assert len(results) >= 1
        assert all(r.dimensions for r in results)

    def test_dimension_meta_fields(self):
        from spc_viz.parsers import parse_excel_multi

        results = parse_excel_multi(FX_FIXTURE, sheet_name="Raw Data-PP")
        dmeta = next(iter(results[0].dimensions.values()))
        assert dmeta.dim_no
        assert isinstance(dmeta.description, str)
        assert isinstance(dmeta.col_labels, list) and len(dmeta.col_labels) > 0
        assert isinstance(dmeta.point_numbers, list)
        assert isinstance(dmeta.nominal, list)
        assert isinstance(dmeta.usl, list)
        assert isinstance(dmeta.lsl, list)

    def test_parse_raw_data_pp_sheet(self):
        from spc_viz.parsers import parse_excel_multi

        results = parse_excel_multi(FX_FIXTURE, sheet_name="Raw Data-PP")
        assert len(results) >= 1
        assert results[0].sheet_name == "Raw Data-PP"
        assert len(results[0].dimensions) > 0

    def test_meta_columns_non_empty(self):
        from spc_viz.parsers import parse_excel_multi

        results = parse_excel_multi(FX_FIXTURE, sheet_name="Raw Data-PP")
        assert isinstance(results[0].meta_columns, list)
        assert len(results[0].meta_columns) > 0

    def test_factory_detected_for_both_fixtures(self):
        from spc_viz.parsers import parse_excel_multi

        fx = parse_excel_multi(FX_FIXTURE, sheet_name="Raw Data-PP")
        lk = parse_excel_multi(LK_FIXTURE, sheet_name="Raw Data-PP")
        assert fx[0].factory == "FX"
        assert lk[0].factory == "LK"


# ---------------------------------------------------------------------------
# 16. Single-point dimension handling in build_combined_chart
# ---------------------------------------------------------------------------


class TestSinglePointDimension:
    def test_returns_figure(self, single_point_data, common_kwargs):
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
        assert fig is not None

    def test_uses_markers_not_lines(self, single_point_data, common_kwargs):
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
        traces = [t for t in fig.data if t.type == "scattergl"]
        assert len(traces) > 0
        assert all("markers" in (t.mode or "") for t in traces)
        assert all("lines" not in (t.mode or "") for t in traces)

    def test_marker_size_set(self, single_point_data, common_kwargs):
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
        assert t.marker is not None and t.marker.size is not None and t.marker.size > 0

    def test_single_x_point(self, single_point_data, common_kwargs):
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

    def test_no_crash_synthetic_single_point(self, common_kwargs):
        cols, meta = _make_dim_meta("SPC_X1", "Flatness", 1, nominal=0.0, usl=0.5, lsl=0.0)
        df = pd.DataFrame({cols[0]: np.random.uniform(0.0, 0.4, 12)})
        df["Factory"] = ["A"] * 6 + ["B"] * 6
        fig = build_combined_chart(
            df=df,
            dim_metas=OrderedDict([("SPC_X1", meta)]),
            dim_nos=["SPC_X1"],
            section_by_fields=["Factory"],
            y_axis_mode="Measurement values",
            custom_yrange=None,
            **common_kwargs,
        )
        assert fig is not None and len(fig.data) > 0
