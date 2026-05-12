# Phase 1 Implementation Plan — Code Refactor

> **For Claude:** REQUIRED SUB-SKILL: Use `superpowers:executing-plans` (or `subagent-driven-development` for in-session execution) to implement this plan task-by-task.

**Goal:** Reorganize the SPC Data Viz codebase into a proper Python package (`src/spc_viz/`) with modern tooling — without changing any user-visible behavior.

**Architecture:** Layered package: `parsers/` (data), `charts/` (domain), `ui/` (presentation), `theme/` (CSS), `config/` (constants). Big files split into ~50-200 line modules grouped by responsibility. `pages/` stays at repo root (Streamlit requirement). All entry points become thin wrappers over the package.

**Tech Stack:** Python 3.10, uv, pyproject.toml, ruff (lint+format), mypy (type check), pre-commit (auto-run on commit), pytest + pytest-cov, Streamlit, Plotly, openpyxl, kaleido.

**Branch:** `refactor/phase-1-structure` (worktree at `../Data_Visualiztion_refactor/`)

**Source of truth design:** [`docs/plans/2026-04-18-spc-restructure-design.md`](./2026-04-18-spc-restructure-design.md) and [`docs/02-architecture.md`](../02-architecture.md)

**Constraint from user:**
- Visualization-only — no behavior changes the user would notice
- Statistical analysis features (CPK, ANOVA, Nelson Rules, CUSUM, EWMA in `shared_ui.py` lines 328-709) → keep in code but move to a separate `_deferred/` module, disconnect from active UI (will be re-enabled in a future version)
- Sample `.xlsx` files move from repo root to `examples/`

---

## Timeline

| Group | Tasks | Est. time |
|---|---|---|
| Pre-flight | PF1-PF5 | ~30 min |
| Restructure | R1-R8 | ~3 hours |
| Type Hints | T1-T2 | ~30 min |
| Coverage | C1-C2 | ~30 min |
| Final | F1-F3 | ~30 min |
| **Total** | **20 tasks** | **~5 hours** |

After every task: tests pass + commit. Worktree branch is always green.

---

## Audit Checkpoints

Lead pauses for user input at:
- **AC1** (after PF5): "Tooling set up. Continue with restructure?"
- **AC2** (before R6): "About to disconnect Summary Statistics expander from UI. Confirm?" (analysis features decision)
- **AC3** (after R8): "Restructure done. All tests pass. Continue with type hints?"
- **AC4** (before F3): "Ready to merge to main and tag v1.0?"

Otherwise tasks run continuously, agents commit per task.

---

## Pre-flight: Tooling Setup

### Task PF1: Create `pyproject.toml`

**Files:**
- Create: `pyproject.toml`

**Step 1: Write `pyproject.toml`**

```toml
[project]
name = "spc-viz"
version = "1.0.0"
description = "SPC Data Visualization for manufacturing quality engineers"
readme = "README.md"
requires-python = ">=3.10,<3.11"
dependencies = [
    "streamlit>=1.30",
    "plotly>=5.18",
    "openpyxl>=3.1",
    "pandas>=2.1",
    "numpy>=1.26",
    "scipy>=1.11",
    "kaleido>=1.2",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "pytest-cov>=4.1",
    "ruff>=0.4",
    "mypy>=1.10",
    "pre-commit>=3.7",
]

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/spc_viz"]

[tool.pytest.ini_options]
testpaths = ["tests"]
markers = [
    "unit: unit tests",
    "integration: integration tests",
]

[tool.ruff]
line-length = 100
target-version = "py310"

[tool.ruff.lint]
select = ["E", "F", "I", "B", "UP", "N", "SIM"]
ignore = ["E501"]  # line length handled by formatter

[tool.mypy]
python_version = "3.10"
strict = false  # gradually tighten
ignore_missing_imports = true
```

**Step 2: Verify `uv` can resolve**

```bash
cd /Users/zhefeng/Desktop/Vibe\ Coding/Data_Visualiztion_refactor
uv pip install -e ".[dev]"
```
Expected: install completes; `streamlit`, `pytest`, `ruff`, `mypy` available in venv.

**Step 3: Run all tests to confirm nothing breaks**

```bash
pytest tests/ -q
```
Expected: 24 passed.

**Step 4: Commit**
```bash
git add pyproject.toml
git commit -m "build: add pyproject.toml with deps + dev tools (uv-managed)"
```

---

### Task PF2: Pin Python version + ruff config

**Files:**
- Create: `.python-version` (single line: `3.10`)
- Create: `ruff.toml` (move ruff section out of pyproject.toml for visibility)

**Step 1: Write `.python-version`**
```
3.10
```

**Step 2: Verify ruff runs cleanly**
```bash
ruff check src/ tests/  # source dir doesn't exist yet — OK
ruff check .  # whole project
```
Expected: a list of issues to fix in PF4.

**Step 3: Commit**
```bash
git add .python-version
git commit -m "build: pin Python to 3.10"
```

---

### Task PF3: Pre-commit hooks

**Files:**
- Create: `.pre-commit-config.yaml`

**Step 1: Write `.pre-commit-config.yaml`**
```yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.4.10
    hooks:
      - id: ruff
        args: [--fix]
      - id: ruff-format
  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v4.6.0
    hooks:
      - id: trailing-whitespace
      - id: end-of-file-fixer
      - id: check-yaml
      - id: check-added-large-files
        args: ['--maxkb=500']
  - repo: local
    hooks:
      - id: pytest-quick
        name: pytest (smoke)
        entry: pytest tests/ -q --tb=no -x
        language: system
        pass_filenames: false
        always_run: true
        stages: [pre-push]
```

**Step 2: Install hooks**
```bash
pre-commit install
pre-commit install --hook-type pre-push
```
Expected: `.git/hooks/pre-commit` and `pre-push` created.

**Step 3: Run on all files to baseline**
```bash
pre-commit run --all-files
```
Expected: format fixes applied; trailing whitespace fixed; etc.

**Step 4: Commit any auto-fixes**
```bash
git add -A
git commit -m "style: apply ruff/whitespace fixes from pre-commit baseline"
```

---

### Task PF4: Apply ruff format/lint to existing code

**Files:** all `.py` files

**Step 1: Run ruff format**
```bash
ruff format .
```

**Step 2: Run ruff check with auto-fix**
```bash
ruff check . --fix
```

**Step 3: Manual review of un-auto-fixed issues**
- Note remaining warnings (e.g., unused imports the linter can't auto-remove)
- Fix obvious ones; defer complex ones

**Step 4: Verify tests still pass**
```bash
pytest tests/ -q
```
Expected: 24 passed.

**Step 5: Commit**
```bash
git add -A
git commit -m "style: ruff format + lint baseline on flat structure"
```

---

### Task PF5: Audit checkpoint (AC1)

**No code change.** Lead asks user:

> Phase 1 — Pre-flight done.
> Tooling installed (pyproject, ruff, mypy, pre-commit). Code formatted. Tests still green (24/24).
> Ready to start the module restructure (R1-R8)?
> A) Yes, continue
> B) Pause

---

## Restructure: Move flat files into `src/spc_viz/` package

### Task R1: Create package skeleton

**Files:**
- Create: `src/spc_viz/__init__.py` (empty)
- Create: `src/spc_viz/parsers/__init__.py` (empty)
- Create: `src/spc_viz/charts/__init__.py` (empty)
- Create: `src/spc_viz/ui/__init__.py` (empty)
- Create: `src/spc_viz/theme/__init__.py` (empty)
- Create: `src/spc_viz/config/__init__.py` (empty)
- Create: `src/spc_viz/_deferred/__init__.py` (empty, for analysis features later)

**Step 1: Make directories and empty `__init__.py` files**
```bash
mkdir -p src/spc_viz/{parsers,charts,ui,theme,config,_deferred}
touch src/spc_viz/__init__.py
for d in parsers charts ui theme config _deferred; do
    touch "src/spc_viz/$d/__init__.py"
done
```

**Step 2: Verify tests still pass (no behavior change yet)**
```bash
pytest tests/ -q
```

**Step 3: Commit**
```bash
git add src/
git commit -m "refactor: scaffold src/spc_viz package skeleton"
```

---

### Task R2: Move `ui_theme.py` → `src/spc_viz/theme/`

**Files:**
- Move: `ui_theme.py` → `src/spc_viz/theme/css.py` (CSS strings) and `src/spc_viz/theme/plotly_theme.py` (Plotly theming functions)
- Update imports in: `app.py`, `pages/1_Quick_Test.py`, `pages/2_Sheet_Manager.py`, `shared_ui.py`, `chart_utils.py`

**Step 1: Read `ui_theme.py` and split**
- CSS strings (the `inject_theme` function and CSS constants) → `src/spc_viz/theme/css.py`
- Plotly figure theming (`finalize_plotly_style` etc.) → `src/spc_viz/theme/plotly_theme.py`
- Re-export both from `src/spc_viz/theme/__init__.py`:
  ```python
  from spc_viz.theme.css import inject_theme  # noqa: F401
  from spc_viz.theme.plotly_theme import finalize_plotly_style  # noqa: F401
  ```

**Step 2: Update all imports across the codebase**
- Find: `from ui_theme import X`
- Replace: `from spc_viz.theme import X`
- Same for `import ui_theme`

**Step 3: Delete the old `ui_theme.py`**

**Step 4: Run tests**
```bash
pytest tests/ -q
```
Expected: 24 passed.

**Step 5: Sanity-check the running app (manual or via Quick Test page)**
- Open Streamlit, verify theme still looks identical.

**Step 6: Commit**
```bash
git add -A
git commit -m "refactor: split ui_theme.py into theme/css.py + theme/plotly_theme.py"
```

---

### Task R3: Move `spc_parser.py` → `src/spc_viz/parsers/`

**Split into 6 files** based on responsibility:

- `header_detect.py` — `_find_data_start`, header keyword logic
- `metadata.py` — metadata column extraction
- `dimensions.py` — `DimensionMeta` dataclass, dimension spec extraction
- `measurements.py` — measurement value extraction, deduplication
- `openpyxl_patch.py` — `ExternalReference` monkey-patch + strict-OOXML rewrite
- `excel_reader.py` — `parse_excel_multi`, `_open_workbook`, top-level orchestrator

**Step 1: Read `spc_parser.py` and identify the 6 logical sections**

**Step 2: Move each section to its new file**
- Keep public API in `excel_reader.py` (the entry function `parse_excel_multi`)
- Re-export from `src/spc_viz/parsers/__init__.py`:
  ```python
  from spc_viz.parsers.excel_reader import parse_excel_multi  # noqa: F401
  from spc_viz.parsers.dimensions import DimensionMeta  # noqa: F401
  ```

**Step 3: Update imports in callers**
- `from spc_parser import X` → `from spc_viz.parsers import X` (where applicable)
- For internal cross-module references inside the new `parsers/` files, use relative imports

**Step 4: Make sure the openpyxl monkey-patch runs at import time**
Add to `src/spc_viz/__init__.py`:
```python
from spc_viz.parsers import openpyxl_patch  # noqa: F401  # apply patch on package import
```

**Step 5: Delete the old `spc_parser.py`**

**Step 6: Run tests**
```bash
pytest tests/ -q
```
Expected: 24 passed.

**Step 7: Run the app and upload a test xlsx to confirm parsing still works**

**Step 8: Commit**
```bash
git add -A
git commit -m "refactor: split spc_parser.py into 6 files under parsers/"
```

---

### Task R4: Move `chart_utils.py` → `src/spc_viz/charts/`

**Split into 7 files:**

- `base.py` — `_build_chart_figure` core helpers, finalize wrappers
- `combined_profile.py` — `build_combined_chart`, `prepare_combined_data`
- `box_plot.py` — `build_box_plot`
- `histogram.py` — `build_histogram`
- `spec_limits.py` — USL/LSL/Nominal line drawing helpers
- `styling.py` — `finalize_plotly_style` and color mapping
- `export.py` — kaleido PNG export, ZIP packaging for batch export

**Step 1-7:** Same pattern as R3.

**Step 8: Run tests + smoke check the app + commit**
```bash
git add -A
git commit -m "refactor: split chart_utils.py into 7 files under charts/"
```

---

### Task R5: Extract config constants → `src/spc_viz/config/`

**Files:**
- Create: `src/spc_viz/config/constants.py` — color palette (blue/red/green/amber/cyan/rose/indigo/orange/teal/slate), default chart sizes, regex patterns for sheet names
- Create: `src/spc_viz/config/paths.py` — log paths, default directories

**Step 1: Identify hardcoded constants currently scattered across `chart_utils.py`, `shared_ui.py`, `spc_parser.py`**

**Step 2: Move them to `config/`**

**Step 3: Update imports in callers**

**Step 4: Run tests + commit**

---

### Task R6: AC2 — User audit on Summary Statistics decision

**No code change yet.** Lead asks user:

> Phase 1 — About to refactor `shared_ui.py`.
>
> I found code for CPK / ANOVA / Nelson Rules / CUSUM / EWMA inside the "Summary Statistics" expander (lines 328-709). You said earlier "leave it for this version, just visualization for now."
>
> What should happen to that expander in the UI?
>
> A) Hide the expander entirely (move code to `_deferred/`, no menu entry)
> B) Keep the expander visible but show "Coming soon" / disable it
> C) Keep it as-is (no changes, looks weird if app is described as visualization-only)
>
> A is recommended for cleanest user experience.

---

### Task R7: Move `shared_ui.py` → `src/spc_viz/ui/`

**Split into 6 files based on responsibility:**

- `sidebar.py` — `build_chart_controls`, `build_color_pickers`, sidebar layout
- `upload.py` — file uploader + sheet-selection multiselect
- `dimension_picker.py` — `build_dimension_selector`, presets, custom selection
- `chart_view.py` — `build_and_render_chart` orchestration
- `batch_export.py` — `render_batch_export` (the sidebar expander)
- `state.py` — session-state helpers + the `key_prefix` convention (KP="main_" / "qt_")

**Plus the deferred analysis code** based on R6 decision:
- → `src/spc_viz/_deferred/analysis.py` (CPK, ANOVA, Nelson, CUSUM, EWMA)

**Step 1-7:** Same pattern as R3/R4.

**Step 8: Run tests + smoke check + commit**
```bash
git add -A
git commit -m "refactor: split shared_ui.py into 6 files under ui/ + move analysis to _deferred/"
```

---

### Task R8: Update entry points & move sample xlsx files

**Files:**
- Modify: `app.py` — make it a thin wrapper (~30-50 lines) that imports from `spc_viz` and runs the main flow
- Modify: `pages/1_Quick_Test.py` — thin wrapper (~30-50 lines)
- Modify: `pages/2_Sheet_Manager.py` — thin wrapper (~30-50 lines)
- Move: `*.xlsx` from repo root → `examples/` (and update `Quick_Test.py` to scan `examples/` instead of root)
- Add: `scripts/run_dev.sh` (for local dev) and `scripts/run_server.sh` (for launchd)

**Step 1: Refactor `app.py`**
```python
# app.py (after refactor — illustrative)
import streamlit as st
from spc_viz.theme import inject_theme
from spc_viz.ui.sidebar import build_chart_controls
from spc_viz.ui.upload import build_uploader
from spc_viz.ui.chart_view import build_and_render_chart
from spc_viz.ui.batch_export import render_batch_export

KP = "main_"

def main() -> None:
    st.set_page_config(...)
    inject_theme()
    parsed_files = build_uploader(KP)
    if not parsed_files:
        st.info("Upload xlsx files to begin.")
        return
    controls = build_chart_controls(parsed_files, KP)
    build_and_render_chart(parsed_files, controls, KP)
    render_batch_export(parsed_files, controls, KP)

if __name__ == "__main__":
    main()
```

**Step 2: Refactor `pages/1_Quick_Test.py` similarly with `KP = "qt_"`**

**Step 3: Move sample xlsx files**
```bash
mkdir -p examples
mv *.xlsx examples/
```

**Step 4: Update Quick_Test.py to scan `examples/` instead of repo root**

**Step 5: Update launchd plist path** (or `scripts/run_server.sh` indirection so plist points to script, not python directly)

**Step 6: Run tests + smoke check + commit**
```bash
git add -A
git commit -m "refactor: thin entry points + move sample xlsx to examples/"
```

---

## Type Hints

### Task T1: Add type hints to public APIs

**Files:** every `__init__.py` re-export and every public function in the new modules

**Step 1: Use `mypy --strict` on the package (start lenient, tighten gradually)**
```bash
mypy src/spc_viz --ignore-missing-imports
```

**Step 2: Add `pandas`/`plotly`/`streamlit` types (pip-install `pandas-stubs`, `plotly-stubs` if available; otherwise `# type: ignore` at import time)**

**Step 3: Add type hints to public function signatures**
Example:
```python
def build_combined_chart(
    df: DataFrame,
    dim_metas: list[DimensionMeta],
    controls: ChartControls,
) -> Figure: ...
```

**Step 4: Run mypy again, iterate until zero errors on `src/`**

**Step 5: Commit per module to keep PRs small**

---

### Task T2: `ChartControls` dataclass

**Files:**
- Create: `src/spc_viz/ui/state.py` — define `@dataclass ChartControls` capturing all sidebar selections (chart_type, color_by, section_by, ...)

This replaces the dict-of-strings currently passed around. Improves type safety + IDE help.

---

## Coverage

### Task C1: Run pytest --cov, identify gaps

```bash
pytest tests/ --cov=src/spc_viz --cov-report=term-missing
```

Identify modules below 80% coverage. Likely candidates: parser edge cases, batch export error paths.

### Task C2: Add tests to reach 80%+ on critical modules

Priority modules: `parsers/header_detect.py`, `parsers/metadata.py`, `charts/spec_limits.py`, `charts/export.py`.

For each gap:
- Write failing test (TDD red)
- Run to confirm fail
- Most cases the code already exists, so the test will go straight to green
- Commit per gap

---

## Final

### Task F1: Full smoke test via spc-test-runner

Dispatch the existing `spc-test-runner` agent. It will:
- Run pytest with coverage
- Render every chart type for a sample xlsx
- Visually compare to baseline (golden PNGs from before refactor)
- Report any visual diffs

**Acceptance:** zero visual diffs, all 24+ tests pass, coverage ≥80% on critical modules.

### Task F2: Update `docs/02-architecture.md`

Module diagram + data flow now reflects the post-refactor structure (`src/spc_viz/` layout). Use `doc-updater` agent.

### Task F3: AC4 — Final audit + merge to main

Lead asks user:

> Phase 1 — All refactor tasks complete.
>
> ✅ Tests: N/N passing
> ✅ Coverage: X% on src/spc_viz
> ✅ Visual regression: zero diffs
> ✅ Mypy: clean
> ✅ Architecture doc updated
>
> Ready to merge `refactor/phase-1-structure` → `main` and tag `v1.0`?
> A) Yes
> B) Hold for me to test the live app first

If A: merge + tag + push. Update launchd plist if needed to point at new entry path.

---

## Success Criteria

- [ ] All 24+ tests pass
- [ ] Coverage ≥80% on `src/spc_viz/parsers/`, `charts/`, `config/`
- [ ] Mypy clean (no errors) on `src/spc_viz/`
- [ ] Visual regression: zero chart diffs
- [ ] User runs the app and confirms identical UI/UX
- [ ] `git tag v1.0` exists
- [ ] Working app at `http://localhost:8504` continues running on the main branch (untouched during Phase 1)

## Rollback

If anything goes wrong:
```bash
git checkout v0.2-designed
# Or
git worktree remove ../Data_Visualiztion_refactor
# main repo is unaffected, app still runs
```
