# Design Doc: SPC Data Visualization App

Status: Current state, branch `feature/code-quality-refactor`
Author: Stanley Feng
Date: 2026-04-18

## 1. Context

This app exists to absorb the mechanical portion of a manufacturing quality engineer's workflow: turning stacks of vendor CPK inspection workbooks (CORR / POR / PP / AP stages, five or more `.xlsx` per launch, mixed sheet conventions) into a publishable pack of SPC charts. Before the tool, the QE opened files side by side in Excel, redrew USL/LSL/Nominal reference lines by hand for each dimension, screenshot-and-pasted every chart into a report, and manually reconciled inconsistent headers across vendors — hours of repeatable work before analysis could begin. It replaces that layer with a local Streamlit app: drag-drop the workbooks, auto-detect dimensions and metadata, render interactive Plotly charts with spec limits, and batch-export the whole pack as PNGs. See `docs/00-product-brief.md` for the fuller user-problem framing.

## 2. Goals

**Primary.** Automate repetitive chart production from vendor CPK Excel files: ingest mixed-format workbooks, auto-detect the dimension layout (headers, metadata columns, data start row), and emit SPC-style charts (Combined Profile, Box Plot, Histogram) with USL/LSL/Nominal reference lines pre-drawn.

**Secondary.** Support a three-page workflow that matches how the user actually works: a main multi-file charting page (`app.py`), a scratchpad Quick Test page that auto-loads local xlsx for fast iteration (`pages/1_Quick_Test.py`), and a Sheet Manager that diffs dimension coverage across sheets so nothing is silently dropped (`pages/2_Sheet_Manager.py`). Batch export every selected chart to a ZIP of PNGs for direct paste into reports.

**Non-goals (from Product Brief).** Not multi-user — no accounts, sharing, or permissions. Not a CPK calculator — capability scoring remains owned by the QE team's existing tooling; this app visualizes measurements against limits, it does not certify them. Not a statistical analysis suite — no ANOVA, regression, or hypothesis testing. Not cloud-hosted — data never leaves the laptop.

## 3. Architecture Overview

See `docs/02-architecture.md` for the module-dependency Mermaid diagram and the upload-to-chart sequence diagram. In one paragraph: the codebase is three clean layers. The **parser layer** (`spc_parser.py`) is the only module that touches `openpyxl`; it returns plain `DimensionMeta` dataclasses and a `ParsedFile` per sheet, with zero awareness of Streamlit or Plotly. The **UI layer** has two modules — `shared_ui.py` owns every `st.*` widget call (dimension selector, chart controls, batch export, summary stats) and is the fan-in hub used by all three pages, while `chart_utils.py` is pure Plotly figure construction and never imports Streamlit. The **Streamlit entry layer** (`app.py` plus the two `pages/` files) wires pages to `shared_ui` using a per-page `key_prefix` to isolate session state. The two critical seams — parser/UI and UI/charts — keep parsing and chart logic unit-testable without a browser.

## 4. Tech Stack Decisions

| Choice | Role | Rationale |
|---|---|---|
| **Streamlit** | Web framework, widgets, session state, page routing | Single-file Python turns into a web app; no separate frontend build, no React, no FastAPI+Jinja glue. Matches single-user local-first deployment. |
| **Plotly** | Interactive charts | Hover/zoom/legend-toggle out of the box, click-to-highlight via a JS shim, exports cleanly to PNG via kaleido. First-class Streamlit integration (`st.plotly_chart`). |
| **openpyxl** | `.xlsx` parsing (parser module only) | Pure-Python, streaming `read_only=True` mode for large vendor files, introspectable workbook model needed to scan for "Dim. No." marker cells and dynamic header rows. Required a local monkey-patch plus a strict-OOXML namespace rewrite to handle vendor files. |
| **pandas** | DataFrame for measurement rows, groupby for color/section controls | Standard for tabular data in Python; integrates with Plotly and Streamlit natively. |
| **kaleido** | Server-side PNG export for Batch Export | Pure-Python, zero system deps (no headless Chrome, no orca binary); Plotly's officially recommended static export engine. |
| **launchd** | Persistence on `localhost:8504` | Native macOS per-user supervisor; no Docker, no systemd-on-mac emulation. Keeps the Streamlit server alive across reboots so the user can bookmark the URL. |

## 5. Data Model

Two dataclasses in `spc_parser.py` form the contract between parser and UI.

**`DimensionMeta`** — one instance per dimension group (e.g. `SPC_AA`) on a sheet.
- `dim_no: str` — dimension identifier (e.g. `SPC_AA`).
- `description: str` — human-readable label (e.g. "landing to E surface height").
- `dim_type: str` — vendor-declared type (e.g. "Non-Profile Measurement").
- `point_numbers: list` — label per sub-column; synthesized as `P0..Pn` when absent.
- `nominal: list`, `tol_max: list`, `tol_min: list`, `usl: list`, `lsl: list` — parallel lists of spec values, one entry per sub-column.
- `col_indices: list` — 1-based sheet column indices, used to slice measurement rows.
- `col_labels: list` — readable column labels used as DataFrame column names.

**`ParsedFile`** — result of parsing a single sheet.
- `filename`, `sheet_name`, `part_number`, `part_description`, `revision`, `factory` — file-level metadata (factory is detected from Vendor Serial Number mode, SN prefix, or filename).
- `dimensions: OrderedDict[str, DimensionMeta]` — insertion-ordered map keyed by `dim_no`.
- `data: pd.DataFrame` — measurement rows; metadata columns plus one column per `col_label`.
- `meta_columns: list` — names of metadata columns detected in the header row.

The app layer refers to a list of these as a "ParsedWorkbook" — there is no separate dataclass; it is just `list[ParsedFile]` returned by `parse_excel_multi`.

## 6. Alternatives Considered

**FastAPI + separate SPA frontend (rejected).** Would have given a cleaner API boundary and easier remote hosting, but required a second codebase (React/Next.js), a build pipeline, and CORS handling. For one local user the cost/benefit was wrong — Streamlit collapses frontend and backend into a single Python file and ships widgets for free.

**Matplotlib (rejected) and Bokeh (rejected).** Matplotlib renders static PNGs with no hover or zoom, which defeats the "interactive exploration" part of the workflow and forces the user back to spreadsheet-like poking. Bokeh is interactive but has weaker Streamlit integration and a smaller library of SPC-shaped chart primitives than Plotly.

**SQLite or a database-backed session store (rejected).** There is no multi-user or cross-session requirement; state lives in `st.session_state` for the duration of a browser tab and is discarded on reload. Adding SQLite would mean schema migrations, file-lock handling, and persistence semantics the user explicitly did not want. The launchd-managed always-on Streamlit process covers "come back tomorrow and the URL still works" without a database.

## 7. Known Technical Debt

- **openpyxl `ExternalReference` monkey-patch** (`spc_parser.py` lines 23–30) — patches a strict-OOXML deserialization bug upstream in openpyxl 3.1.x. Brittle: a future openpyxl release may change the signature and silently break the patch. Upstream fix is the real resolution.
- **LK-file parse failure not covered by the patch.** A distinct error (`'NoneType' has no 'Target'`) occurs on some LK vendor files and is not handled by the current monkey-patch; needs a second fix path.
- **Vendor-typo column-name detection** (US-019, planned) — the `_scan_label_rows` label map is a hand-curated list of exact strings; typos like "Nomial Dim." silently miss. A fuzzy-match layer is planned.
- **Statistical analysis code in `shared_ui.py` / `chart_utils.py`** (`calc_process_capability`, `nelson_rules`, `cusum_analysis`) — deferred from the non-goals list; currently unreferenced by the active workflow and at risk of becoming dead code.
- **`requirements.txt` has no version pins** — any `pip install -r` today resolves to whatever the registry returns, so reproducibility is not guaranteed across machines or across time. Action: pin or switch to a lockfile (`uv`, `pip-tools`, `poetry`).
- **launchd plist not checked into the repo** — it lives only in `~/Library/LaunchAgents/` on the author's machine. Rebuilding the environment requires reconstructing it from `logs/streamlit.log` hints and memory. Action: commit the plist template under `ops/` or `scripts/`.

## 8. Constraints

- **Python 3.10**, pinned by the local `.venv` and assumed by match-statement usage in dependencies. No 3.11+ features in the codebase.
- **macOS-only deployment.** launchd is the supervisor; no Windows service or systemd equivalent is provided or tested.
- **Single user, local-only.** No authentication, no authorization, no rate limiting; the app binds to `localhost:8504` and assumes the operating system's user boundary is the only security perimeter.
- **No cloud, no database, no auth.** Measurement data stays on the laptop; state is session-scoped in Streamlit; there is no persistent store beyond the user's original `.xlsx` files.
