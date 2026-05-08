# Test Coverage — v1.0 Baseline

Recorded 2026-05-08 after Phase 1 refactor. Use as the v1.0 starting point;
push specific modules higher only when a real bug points to a missing test.

## Summary

```
TOTAL: 47%   (1575 statements, 827 missed)
```

## Per-module

| Module | % | Note |
|---|---|---|
| `charts/styling.py` | 100% | Trivial wrapper |
| `charts/spec_limits.py` | 100% | Empty placeholder |
| `charts/export.py` | 100% | Empty placeholder |
| `charts/histogram.py` | 80% | Good — exercised by snapshot test |
| `charts/box_plot.py` | ~80% | Good — exercised by snapshot test |
| `charts/combined_profile.py` | ~70% | Good — exercised by snapshot test |
| `charts/base.py` | high | Compute + palette helpers |
| `parsers/measurements.py` | 82% | Good |
| `parsers/header_detect.py` | 63% | Edge cases (no header, exotic layouts) untested |
| `parsers/metadata.py` | 63% | Factory detection branches partial |
| `parsers/excel_reader.py` | 53% | Multi-sheet auto-detect path partial |
| `parsers/dimensions.py` | 52% | Some helpers (group keyword extraction) untested |
| `parsers/openpyxl_patch.py` | 27% | Hard to test — only exercised on broken xlsx files |
| `ui/state.py` | 58% | ChartControls dataclass tested; prepare_and_clean has Streamlit calls |
| `ui/chart_view.py` | 68% | Render path tested via snapshot |
| `ui/batch_export.py` | 11% | Streamlit-dependent — needs harness |
| `ui/sidebar.py` | 19% | Streamlit-dependent — needs harness |
| `ui/dimension_picker.py` | 15% | Streamlit-dependent — needs harness |
| `theme/css.py` | 0% | Only CSS strings; nothing to assert |

## Why we stop at 47%, not 80%

1. **Snapshot regression suite covers what matters most.** The 4 golden-image
   tests at `tests/snapshot_test.py` exercise the full chart pipeline
   end-to-end: parse → combine → render → PNG bytes. Any user-visible chart
   change is caught.

2. **UI layer is Streamlit-dependent.** Unit-testing `st.sidebar.expander`,
   `st.file_uploader`, and `st.plotly_chart` calls requires a Streamlit
   test harness (e.g. `streamlit-testing-library`) we have not adopted.
   The cost/benefit doesn't favor doing it now.

3. **Parser edge-case tests need fixtures.** Lower-covered parser paths
   (header autodetect failures, broken openpyxl files) need malformed
   xlsx fixtures we don't have. Adding them is real work without a real
   bug to motivate the shape of the fixture.

4. **Theme is data, not logic.** CSS strings don't need test coverage.

## When to revisit

Bump targets when:
- A bug ships that an edge-case parser test would have caught → add the test, not blanket coverage
- We adopt a Streamlit testing library → backfill UI tests
- We re-enable analysis features from `_deferred/` → must hit ≥80% on those modules before re-wiring
