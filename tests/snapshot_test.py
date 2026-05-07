"""Snapshot test runner for the SPC visualization app.

Usage:
    python tests/snapshot_test.py                  # compare against goldens
    python tests/snapshot_test.py --update-goldens # write new goldens
    python tests/snapshot_test.py --list           # list scenarios
    python tests/snapshot_test.py --name <name>    # run one scenario

Renders each scenario via the project's chart builders (no Streamlit needed)
and compares the resulting PNG against a stored golden image.
"""

from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import spc_parser  # noqa: E402  # activates openpyxl monkey-patch
from chart_utils import (  # noqa: E402
    build_box_plot,
    build_combined_chart,
    build_histogram,
    finalize_plotly_style,
    prepare_combined_data,
)
from tests.snapshot_scenarios import SCENARIOS, Scenario  # noqa: E402

_FIXTURES_DIR = Path(__file__).parent / "fixtures"
_GOLDEN_DIR = Path(__file__).parent / "golden"
_DIFF_DIR = Path(__file__).parent / "diff"


def _parsed_file_to_dict(pf):
    """Convert ParsedFile dataclass to the dict format chart_utils expects."""
    return {
        "filename": pf.filename,
        "sheet_name": pf.sheet_name,
        "part_number": pf.part_number,
        "part_description": pf.part_description,
        "revision": pf.revision,
        "factory": pf.factory,
        "dimensions": pf.dimensions,
        "data": pf.data,
        "meta_columns": pf.meta_columns,
    }


def _parse_fixtures(scenario: Scenario) -> list:
    """Parse all xlsx files for a scenario; returns list of dicts."""
    parsed = []
    for fname in scenario.files:
        path = _FIXTURES_DIR / fname
        if not path.exists():
            raise FileNotFoundError(f"Fixture missing: {path}")
        results = spc_parser.parse_excel_multi(str(path), sheet_name=scenario.sheet)
        for r in results:
            parsed.append(_parsed_file_to_dict(r))
    return parsed


def _render_scenario(scenario: Scenario) -> bytes:
    """Render a scenario's chart as PNG bytes."""
    parsed_files = _parse_fixtures(scenario)
    if not parsed_files:
        raise RuntimeError(f"No data parsed for {scenario.name}")

    df, dim_metas = prepare_combined_data(parsed_files, [scenario.dim_no])
    if df is None or df.empty:
        raise RuntimeError(f"Empty dataframe for {scenario.name}")
    if scenario.dim_no not in dim_metas:
        raise RuntimeError(f"Dim {scenario.dim_no} not found in {scenario.files}")

    common = dict(
        df=df,
        dim_metas=dim_metas,
        dim_nos=[scenario.dim_no],
        color_by=scenario.color_by or "None",
        exclude_intervals=scenario.exclude_intervals,
        group_label=scenario.color_by or "All",
        row_by=scenario.row_by or "None",
        custom_color_map=None,
        selected_points=None,
    )

    if scenario.chart_type == "Profile":
        fig = build_combined_chart(
            **common,
            section_by_fields=scenario.section_by,
            y_axis_mode="Measurement values",
            custom_yrange=None,
        )
    elif scenario.chart_type == "Box Plot":
        fig = build_box_plot(
            **common,
            y_axis_mode="Measurement values",
            custom_yrange=None,
        )
    elif scenario.chart_type == "Histogram":
        fig = build_histogram(**common, nbins=40)
    else:
        raise ValueError(f"Unknown chart_type: {scenario.chart_type}")

    if fig is None:
        raise RuntimeError(f"Chart builder returned None for {scenario.name}")

    finalize_plotly_style(fig)
    return fig.to_image(format="png", width=1400, height=700, scale=2)


def _png_hash(png_bytes: bytes) -> str:
    return hashlib.sha256(png_bytes).hexdigest()[:16]


def _compare(scenario: Scenario, *, update: bool) -> tuple[bool, str]:
    golden_path = _GOLDEN_DIR / f"{scenario.name}.png"
    diff_path = _DIFF_DIR / f"{scenario.name}.actual.png"

    try:
        actual = _render_scenario(scenario)
    except Exception as e:
        return False, f"RENDER ERROR: {type(e).__name__}: {e}"

    if update or not golden_path.exists():
        _GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
        golden_path.write_bytes(actual)
        return True, f"GOLDEN WRITTEN ({len(actual):,} bytes, hash {_png_hash(actual)})"

    expected = golden_path.read_bytes()
    if actual == expected:
        return True, f"PASS (hash {_png_hash(actual)})"

    _DIFF_DIR.mkdir(parents=True, exist_ok=True)
    diff_path.write_bytes(actual)
    return False, (
        f"FAIL — got {_png_hash(actual)}, expected {_png_hash(expected)}\n"
        f"  Actual saved to: {diff_path}"
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--update-goldens", action="store_true")
    parser.add_argument("--list", action="store_true")
    parser.add_argument("--name", type=str, default=None)
    args = parser.parse_args()

    if args.list:
        print(f"\n{len(SCENARIOS)} scenarios:")
        for s in SCENARIOS:
            print(f"  {s.name:40s} — {s.description}")
        return 0

    scenarios = [s for s in SCENARIOS if not args.name or s.name == args.name]
    if not scenarios:
        print(f"No scenario matches name={args.name!r}")
        return 1

    mode = " [UPDATE MODE]" if args.update_goldens else ""
    print(f"\nRunning {len(scenarios)} scenario(s){mode}\n")

    results = []
    for s in scenarios:
        passed, msg = _compare(s, update=args.update_goldens)
        status = "✅" if passed else "❌"
        print(f"{status} {s.name:40s} {msg}")
        results.append(passed)

    print(f"\n{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
