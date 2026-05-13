"""Snapshot test scenarios for the SPC visualization app.

Each scenario captures a real workflow the user cares about. The test agent
re-renders each scenario as PNG and compares against a stored golden image.

Adding a new scenario:
  1. Add an entry to SCENARIOS below
  2. Run: python tests/snapshot_test.py --update-goldens
  3. Visually verify the new golden PNG looks right
  4. Commit the new scenario + golden
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Scenario:
    """One snapshot test scenario."""

    name: str  # short identifier — used as PNG filename
    description: str  # human-readable summary
    files: list[str] = field(default_factory=list)  # paths under tests/fixtures/
    sheet: str = "Raw Data-PP"  # sheet name to parse
    dim_no: str = "SPC_A-1"  # dimension to chart (single-dim shorthand)
    chart_type: str = "Profile"  # Profile | Box Plot | Histogram
    color_by: str | None = "Extrusion"  # column name or None
    section_by: list[str] = field(default_factory=list)  # column names
    row_by: str | None = None
    exclude_intervals: bool = True
    # Optional override: when set, used instead of [dim_no] for multi-dim
    # rendering. Existing scenarios leave this empty for back-compat.
    extra_dims: tuple[str, ...] = ()


SCENARIOS: list[Scenario] = [
    # User's primary workflow — captured live via Chrome on 2026-05-07
    Scenario(
        name="01_profile_extrusion_factory",
        description="User's main workflow: 2 vendor files merged, Profile by Extrusion, sectioned by Factory",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        sheet="Raw Data-PP",
        dim_no="SPC_A-1",
        chart_type="Profile",
        color_by="Extrusion",
        section_by=["Factory"],
    ),
    # Auto-variation 1 — same data, Box Plot
    Scenario(
        name="02_boxplot_extrusion_factory",
        description="Same files + dim, Box Plot view",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        sheet="Raw Data-PP",
        dim_no="SPC_A-1",
        chart_type="Box Plot",
        color_by="Extrusion",
        section_by=["Factory"],
    ),
    # Auto-variation 2 — same data, Histogram
    Scenario(
        name="03_histogram_extrusion",
        description="Same files + dim, Histogram view",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        sheet="Raw Data-PP",
        dim_no="SPC_A-1",
        chart_type="Histogram",
        color_by="Extrusion",
        section_by=[],
    ),
    # Auto-variation 3 — Profile with no grouping (sanity check)
    Scenario(
        name="04_profile_no_grouping",
        description="Profile with no color/section grouping (baseline)",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        sheet="Raw Data-PP",
        dim_no="SPC_A-1",
        chart_type="Profile",
        color_by=None,
        section_by=[],
    ),
    # --- Profile chart variations -----------------------------------------
    Scenario(
        name="05_profile_color_none_section_factory",
        description="Profile, no color, sectioned by Factory (catches 'no sections' bug)",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Profile",
        color_by=None,
        section_by=["Factory"],
    ),
    Scenario(
        name="06_profile_row_by_factory",
        description="Profile, row-by Factory faceting",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Profile",
        color_by="Extrusion",
        section_by=[],
        row_by="Factory",
    ),
    Scenario(
        name="07_profile_multi_section",
        description="Profile, section concatenated by [Factory, Build]",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Profile",
        color_by="Extrusion",
        section_by=["Factory", "Build"],
    ),
    Scenario(
        name="08_profile_color_factory_no_section",
        description="Profile colored by Factory with no section",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Profile",
        color_by="Factory",
        section_by=[],
    ),
    # --- Box Plot variations ----------------------------------------------
    Scenario(
        name="09_boxplot_no_color_no_section",
        description="Box Plot with minimal config (no color, no section)",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Box Plot",
        color_by=None,
        section_by=[],
    ),
    Scenario(
        name="10_boxplot_color_factory_section_factory",
        description="Box Plot, same field (Factory) used for both color and section",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Box Plot",
        color_by="Factory",
        section_by=["Factory"],
    ),
    # --- Histogram variations ---------------------------------------------
    Scenario(
        name="11_histogram_no_color",
        description="Histogram with no color grouping",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Histogram",
        color_by=None,
        section_by=[],
    ),
    Scenario(
        name="12_histogram_color_factory",
        description="Histogram colored by Factory",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        chart_type="Histogram",
        color_by="Factory",
        section_by=[],
    ),
    # --- Low-point-count dimensions ---------------------------------------
    # SPC_L-1 has only 2 points (smallest available in fixtures)
    Scenario(
        name="13_profile_low_point_dim",
        description="Profile on SPC_L-1 (2-point dim) — exercises minimal point layout",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_L-1",
        chart_type="Profile",
        color_by="Extrusion",
        section_by=["Factory"],
    ),
    Scenario(
        name="14_boxplot_low_point_dim",
        description="Box Plot on SPC_L-1 (2-point dim)",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_L-1",
        chart_type="Box Plot",
        color_by="Extrusion",
        section_by=["Factory"],
    ),
    # --- Multi-dim --------------------------------------------------------
    Scenario(
        name="15_profile_two_dims",
        description="Profile with two dimensions combined (SPC_A-1 + SPC_B)",
        files=["FX_X3745.xlsx", "LK_X3745.xlsx"],
        dim_no="SPC_A-1",
        extra_dims=("SPC_B",),
        chart_type="Profile",
        color_by="Extrusion",
        section_by=["Factory"],
    ),
]
