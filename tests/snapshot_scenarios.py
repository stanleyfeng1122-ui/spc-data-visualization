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
    dim_no: str = "SPC_A-1"  # dimension to chart
    chart_type: str = "Profile"  # Profile | Box Plot | Histogram
    color_by: str | None = "Extrusion"  # column name or None
    section_by: list[str] = field(default_factory=list)  # column names
    row_by: str | None = None
    exclude_intervals: bool = True


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
]
