# SPC Data Visualization — Architecture (Phase 1)

Current state after refactoring: code reorganized into layered `src/spc_viz/` package.
User-facing behavior unchanged. Tests and examples remain functional.

**Last Updated:** 2026-05-08

---

## 1. Module Structure

```mermaid
graph TB
    subgraph Extern["External Dependencies"]
        ST[Streamlit >= 1.30]
        PL[Plotly >= 5.18]
        OX[openpyxl >= 3.1]
        PD[pandas >= 2.1]
        NP[numpy >= 1.26]
        SC[scipy >= 1.11]
        KA[kaleido >= 1.2]
    end

    subgraph Entry["Entry Points (Streamlit pages at repo root)"]
        APP["app.py<br/>(main page)"]
        QT["pages/1_Quick_Test.py<br/>(auto-load examples/)"]
        SM["pages/2_Sheet_Manager.py<br/>(dim coverage diff)"]
    end

    subgraph Parsers["parsers/ — Excel → Dataclasses"]
        EXL["excel_reader.py<br/>(322 lines)<br/>parse_excel()<br/>parse_excel_multi()"]
        MSR["measurements.py<br/>(319 lines)<br/>extract_measurements()"]
        HDR["header_detect.py<br/>(146 lines)<br/>find dims + data"]
        MTD["metadata.py<br/>(111 lines)<br/>factory detection"]
        DIM["dimensions.py<br/>(242 lines)<br/>DimensionMeta"]
        OXP["openpyxl_patch.py<br/>(91 lines)<br/>ExtRef monkey-patch"]
    end

    subgraph Charts["charts/ — Data → Plotly Figure"]
        BAS["base.py<br/>(330 lines)<br/>data prep, sections<br/>SPC analytics"]
        CPR["combined_profile.py<br/>(416 lines)<br/>main chart type"]
        BXP["box_plot.py<br/>(227 lines)"]
        HST["histogram.py<br/>(193 lines)"]
        STY["styling.py<br/>(29 lines)<br/>finalize_plotly_style()"]
    end

    subgraph UI["ui/ — Streamlit Widgets"]
        SDB["sidebar.py<br/>(136 lines)<br/>chart controls"]
        BEX["batch_export.py<br/>(135 lines)<br/>export panel"]
        CHV["chart_view.py<br/>(100 lines)<br/>orchestrator"]
        DPK["dimension_picker.py<br/>(93 lines)<br/>dim selector"]
        STA["state.py<br/>(~130 lines)<br/>ChartControls"]
    end

    subgraph Theme["theme/ — Design Tokens"]
        CSS["css.py<br/>(~580 lines)<br/>colors, fonts, inject_theme()"]
    end

    subgraph Config["config/ — Constants + Paths"]
        CNS["constants.py<br/>(placeholder)"]
        PTH["paths.py<br/>(REPO_ROOT, EXAMPLES_DIR)"]
    end

    subgraph Hidden["_deferred/ — v1.0 Hidden"]
        ANL["analysis.py<br/>(519 lines)<br/>CPK, ANOVA, Nelson<br/>CUSUM, EWMA"]
    end

    APP --> STA
    APP --> CHV
    APP --> CSS
    QT --> STA
    QT --> CHV
    QT --> CSS
    SM --> OXP

    CHV --> SDB
    CHV --> BEX
    CHV --> CPR
    CHV --> BXP
    CHV --> HST
    SDB --> DPK
    BEX --> EXL

    CPR --> BAS
    BXP --> BAS
    HST --> BAS
    BAS --> MSR
    BAS --> DIM

    EXL --> MSR
    EXL --> HDR
    EXL --> MTD
    EXL --> DIM
    MSR --> DIM
    HDR --> DIM
    MTD --> DIM
    OXP --> OX

    STA --> EXL
    DIM --> PD

    CPR --> PL
    BXP --> PL
    HST --> PL
    STY --> PL
    BAS --> PL
    BAS --> NP
    BAS --> SC

    EXL --> OX
    EXL --> PD
    STA --> PD
    CSS --> ST
    SDB --> ST
    BEX --> ST
    CHV --> ST

    BEX --> KA
```

**Legend:**
- **Parsers layer** — stateless Excel→dataclass pipeline. No Streamlit/Plotly deps.
- **Charts layer** — pure Plotly Figure construction + SPC math (base.py). No Streamlit.
- **UI layer** — Streamlit widgets + session state. Calls parsers + charts, renders with st.*.
- **Theme** — CSS tokens. Injected by ui/sidebar.py.
- **Config** — app-wide constants and path helpers.
- **_deferred/** — Analysis features disabled in v1.0 (CPK, Nelson rules, CUSUM, EWMA, ANOVA).
  Kept in source for Phase 2 re-enablement. Not imported by active code.

---

## 2. Data Flow — Upload to Rendered Chart

```mermaid
sequenceDiagram
    actor User
    participant Browser
    participant Streamlit as app.py
    participant State as ui/state.py
    participant Parser as parsers/excel_reader.py
    participant Charts as charts/base.py & variants
    participant UI as ui/chart_view.py

    User->>Browser: drag-drop .xlsx files
    Browser->>Streamlit: st.file_uploader bytes

    Streamlit->>Parser: parse_excel_multi(file, sheets)
    Note over Parser: openpyxl_patch<br/>find_data_start, detect headers<br/>extract metadata, measurements<br/>build DimensionMeta per dim
    Parser-->>Streamlit: ParsedFile{df, dimensions{}, factory, ...}

    Streamlit->>State: prepare_and_clean(parsed_files, dim_nos)
    Note over State: filter dims<br/>trim measurements<br/>clean outliers
    State-->>Streamlit: (df_clean, DimensionMeta[])

    User->>Browser: pick chart type + color_by
    Browser->>Streamlit: sidebar multiselect

    Streamlit->>UI: build_and_render_chart(...)
    UI->>Charts: prepare_combined_data<br/>build_combined_chart / build_box_plot / build_histogram
    Note over Charts: add spec limits<br/>calc SPC sections<br/>finalize Plotly style
    Charts-->>UI: plotly.graph_objects.Figure
    UI->>Streamlit: st.plotly_chart(fig, ...)
    Streamlit-->>Browser: Plotly HTML + JSON
    Browser-->>User: interactive chart

    Note over Browser: app.py injects JS shim<br/>for Combined Profile mode<br/>click-to-highlight
```

---

## 3. Layered Architecture

### Parsers (`src/spc_viz/parsers/`)

**Contract:** Excel bytes → `DimensionMeta` dataclass + `pandas.DataFrame`

- `openpyxl_patch.py` — Monkey-patch `ExternalReference` load to handle malformed XLNK refs.
- `dimensions.py` — Data models: `DimensionMeta` (spec, unit, target), `ParsedFile` envelope.
- `header_detect.py` — Scan first 50 rows for dimension number + data start row.
- `metadata.py` — Extract metadata columns (part #, serial, date); auto-detect factory.
- `measurements.py` — Extract measurement rows; deduplicate identical rows.
- `excel_reader.py` — Orchestrator: `parse_excel()` single sheet, `parse_excel_multi()` all sheets.

**Key property:** No Streamlit or Plotly imports. Can be used in scripts or tests without UI.

### Charts (`src/spc_viz/charts/`)

**Contract:** Cleaned DataFrame + `DimensionMeta[]` → `plotly.graph_objects.Figure`

- `base.py` — Shared data prep (sections by date, color palettes, SPC calcs: mean, sigma, control limits).
- `combined_profile.py` — Main multi-dimension SPC chart (lines per dim, color by factory/serial).
- `box_plot.py` — Box plot by factory or serial.
- `histogram.py` — Histogram with mean/spec overlay.
- `styling.py` — `finalize_plotly_style()` — margin, font, legend, hovertemplate tune.

**Key property:** No Streamlit imports. Figures are plain Plotly; testable in REPL.

### UI (`src/spc_viz/ui/`)

**Contract:** Streamlit widgets + session state + orchestration

- `state.py` — `ChartControls` dataclass; `prepare_and_clean()` filters/cleans data.
- `sidebar.py` — Chart type + color_by + threshold selectors + color pickers.
- `batch_export.py` — Batch export panel (Ctrl+click to select; export PNG/XLSX).
- `dimension_picker.py` — Multi-select dimensions + point filter (date range, value bounds).
- `chart_view.py` — `build_and_render_chart()` — calls parsers + charts + `st.plotly_chart()`.

**Key property:** Only layer that imports Streamlit. Session-state keys use `key_prefix` per page.

### Theme (`src/spc_viz/theme/`)

- `css.py` — CSS class + design tokens (Slate palette, Inter/Courier Prime fonts, custom scrollbar).
  Exported as `THEME_CSS` string; injected by sidebar via `st.markdown(..., unsafe_allow_html=True)`.

### Config (`src/spc_viz/config/`)

- `paths.py` — `REPO_ROOT`, `EXAMPLES_DIR` helpers (used by Quick Test to auto-discover `.xlsx`).
- `constants.py` — Reserved for app-wide settings (v1.0: empty).

### Hidden (`src/spc_viz/_deferred/`)

- `analysis.py` — CPK, ANOVA, Nelson rules, CUSUM, EWMA. Disabled in v1.0 (no imports from active code).
  Kept for Phase 2 re-enablement without file recreation.

---

## 4. Entry Points

All three remain at repo root (Streamlit requirement):

| File | Purpose | Dependencies |
|------|---------|--------------|
| `app.py` | Upload vendor files → pick sheets → chart dims | parsers, charts, ui, theme |
| `pages/1_Quick_Test.py` | Auto-load `.xlsx` from `examples/` → fast iteration | parsers, charts, ui, theme |
| `pages/2_Sheet_Manager.py` | Diff dimension coverage between two files (set logic only) | parsers only |

**Session isolation:** Each page passes `key_prefix = "main_"`, `"qt_"`, `"sm_"` to UI functions.
Prevents widget-key collisions across pages in shared Streamlit session state.

---

## 5. External Dependencies

| Library    | Purpose | Version |
|------------|---------|---------|
| streamlit  | Web framework, widgets, session state, multi-page routing | ≥ 1.30 |
| plotly     | Interactive chart construction (Figure, Scatter, Box, Histogram) | ≥ 5.18 |
| openpyxl   | `.xlsx` parsing (isolated to `parsers/`) | ≥ 3.1 |
| pandas     | DataFrames, groupby, aggregations | ≥ 2.1 |
| numpy      | Numeric helpers (percentile, std, mean) | ≥ 1.26 |
| scipy      | `scipy.stats` for capability / Nelson rule analysis | ≥ 1.11 |
| kaleido    | Server-side PNG export (Batch Export feature) | ≥ 1.2 |

**Dev tools** (in `pyproject.toml[project.optional-dependencies.dev]`):
- pytest ≥ 8.0, pytest-cov ≥ 4.1
- ruff ≥ 0.4 (linter/formatter)
- mypy ≥ 1.10 (type checker)
- pre-commit ≥ 3.7

---

## 6. Migration from Phase 0

**What changed:** Code reorganized from flat root (7 modules) to layered `src/spc_viz/` package (18 modules, 4 subpackages).

**Why:** Improve maintainability, enable future packaging/distribution, allow isolated testing of parsers and charts without Streamlit.

**What stayed the same:** User-facing behavior (all chart types, export, UI flow). Test suite still passes (snapshot + unit tests remain unaffected). Examples/ directory moved but still auto-discovered.

**Breaking changes:** None for end users. Developers must update imports (`from spc_parser import ...` → `from src.spc_viz.parsers.excel_reader import ...`).

---

**Test Coverage:**
- `tests/test_spc_charts.py` — 24 unit tests (parsers, charts)
- `tests/snapshot_test.py` — 4 visual regression scenarios
- `tests/fixtures/` — Sample `.xlsx` files (symlinked to vendor data)
- `tests/golden/` — Reference PNG images for visual baseline

**Tooling:**
- `.python-version` → 3.10
- `pyproject.toml` → hatchling build, ruff/mypy/pytest config
- `.pre-commit-config.yaml` → ruff format/lint, file checks
