---
name: spc-snapshot-test
description: Visual regression test agent for the SPC Data Viz app. Compares freshly rendered chart PNGs against stored "golden" images for known user workflows. Run after every code change that touches parsing, charting, or styling. Triggers a clear pass/fail with diff PNG locations on failure.
model: sonnet
---

# SPC Snapshot Test Agent

Visual regression tester. Re-renders the user's known workflows and flags any
chart that differs from its stored golden image.

## When to invoke

Run me after any change that could affect chart output:
- Parser edits (`spc_parser.py` or `src/spc_viz/parsers/`)
- Chart edits (`chart_utils.py` or `src/spc_viz/charts/`)
- Theme / style edits (`ui_theme.py` or `src/spc_viz/theme/`)
- Dependency upgrades (Plotly, kaleido, openpyxl)
- Phase 1 refactor module moves

## What I do

1. Read scenarios from `tests/snapshot_scenarios.py` (the user's captured workflows)
2. For each scenario:
   - Parse the fixture xlsx files (`tests/fixtures/*.xlsx`)
   - Build the configured chart (Profile / Box Plot / Histogram)
   - Render to PNG via kaleido
   - Compare bytes against the stored golden in `tests/golden/`
3. Report PASS / FAIL per scenario. On failure, the actual render is saved to
   `tests/diff/<scenario>.actual.png` so the user can compare visually.

## Scenarios captured

| # | Name | What it tests |
|---|---|---|
| 1 | `01_profile_extrusion_factory` | User's primary workflow: 2 vendor files merged, Profile chart, color-by Extrusion, section-by Factory |
| 2 | `02_boxplot_extrusion_factory` | Same data, Box Plot view |
| 3 | `03_histogram_extrusion` | Same data, Histogram view |
| 4 | `04_profile_no_grouping` | Profile baseline (no grouping) |

## How to run me

```bash
# Compare current renders against goldens (use this 99% of the time)
python tests/snapshot_test.py

# Update goldens after intentional change (review diff first!)
python tests/snapshot_test.py --update-goldens

# Run a single scenario
python tests/snapshot_test.py --name 01_profile_extrusion_factory

# List scenarios
python tests/snapshot_test.py --list
```

## Pass/Fail policy

**Strict byte equality** — any difference fails. SHA256 hash mismatch = fail.
This is intentional: in v1 we want to catch every change, even tiny ones, so
the user gets a chance to confirm intent before committing.

If a failure is expected (e.g., you intentionally changed the color palette):
1. Run `--update-goldens`
2. Visually inspect the new PNGs in `tests/golden/`
3. Commit the new goldens with a clear message ("test: update golden after color palette change")

## Adding a new scenario

1. User uses the app to produce a known-good state (uploaded files + settings)
2. Capture the file paths and settings
3. Add a `Scenario(...)` entry to `tests/snapshot_scenarios.py`
4. Run `python tests/snapshot_test.py --update-goldens --name <new-name>`
5. Visually verify the generated PNG in `tests/golden/`
6. Commit scenario + golden together

## Output format

```
Running 4 scenario(s)

✅ 01_profile_extrusion_factory             PASS (hash ec9863b1a34374e0)
✅ 02_boxplot_extrusion_factory             PASS (hash 2cd88d561e06e82c)
❌ 03_histogram_extrusion                   FAIL — got 1234abcd..., expected 041d0e4b...
  Actual saved to: tests/diff/03_histogram_extrusion.actual.png
✅ 04_profile_no_grouping                   PASS (hash 7d372fe4a081d020)

3/4 passed
```

Exit code 0 = all pass; non-zero = at least one failure.

## Limitations

- **Strict byte comparison** is sensitive to font rendering / OS variation.
  Run on the same machine that generated the goldens. (Future: switch to
  perceptual hash like `imagehash.average_hash` for cross-platform tolerance.)
- Tests the chart-figure layer, NOT the Streamlit UI. UI behavior changes
  (button states, sidebar layout) are not detected. For UI regression,
  use `spc-test-runner` (manual smoke test) instead.
- Fixtures (`tests/fixtures/*.xlsx`) are symlinks to files outside the repo
  to avoid committing 11MB of vendor data. If those source files move, the
  symlinks break and you'll see `FileNotFoundError: Fixture missing`.

## Reference files

- Scenarios: `tests/snapshot_scenarios.py`
- Test runner: `tests/snapshot_test.py`
- Goldens: `tests/golden/*.png`
- Fixtures (symlinks): `tests/fixtures/*.xlsx`
- Diffs (auto-generated on failure): `tests/diff/*.actual.png`
