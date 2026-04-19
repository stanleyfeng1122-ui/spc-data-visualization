# SPC Data Visualization — User Stories

> Agile user stories enumerating every user-facing feature of the SPC (Statistical
> Process Control) visualization app. Stories describe user BEHAVIOR, not code
> structure. SPC domain terms are defined inline on first use.
>
> **Primary user:** Manufacturing Quality Engineer (QE) analyzing dimension
> measurement data from vendor inspection reports.

---

## US-001: Upload multiple vendor CPK Excel files

**As a** quality engineer
**I want** to drag-and-drop one or more `.xlsx` files (CPK / SPC reports from vendors like FX, TY, TRM, LK, FXJS) into the sidebar uploader
**So that** I can analyze measurements from several vendors or several build stages in one session without exporting/merging data manually

**Acceptance criteria:**
- [ ] Sidebar file uploader accepts multiple `.xlsx` files at once
- [ ] Non-xlsx file types are rejected by the uploader
- [ ] When no file is uploaded, the main area shows an info banner and a landing title instead of an empty chart
- [ ] After upload, the loaded files appear in a "Loaded Files" sidebar expander with row count, dim count, and (when present) CFG values

**Status:** Implemented

---

## US-002: Select which sheets to parse across uploaded files

**As a** quality engineer
**I want** to pick one or more sheets (e.g. "Raw data", "Data Input-PP", "Data Input-AP") to parse, shown as a unified multi-select across all uploaded files
**So that** I can focus on a specific CPK production stage (CORR / POR / PP / AP) even when different vendors label the same sheet slightly differently

**Acceptance criteria:**
- [ ] Sheet multi-select appears in the sidebar after files are uploaded
- [ ] Non-data sheets (`BoxPlotCht*`, `Histo Pivot`, `Histo Listbox`, `Histo Curve`) are automatically hidden
- [ ] Sheets with the same name (case-insensitive) across different files are merged into one option
- [ ] All detected sheets are enabled by default
- [ ] A collapsible "Sheet / File map" expander shows which actual sheet in each file is being parsed for the enabled selection
- [ ] Clearing the selection halts parsing with an instructional message

**Status:** Implemented

---

## US-003: Auto-detect sheet layout and header rows

**As a** quality engineer
**I want** the parser to automatically locate the `Dim. No.` marker, the header row, and the first data row regardless of where they sit on the sheet
**So that** I don't have to pre-format or normalize vendor files before uploading — the app adapts to layout variations

**Acceptance criteria:**
- [ ] Parser scans the top-left region of each sheet for a cell matching `Dim. No.` (case-insensitive, with or without dots/spaces)
- [ ] Header row is identified by presence of keywords (`Start Point`, `SN`, `NO`, `no.`) or by a row containing 3+ non-numeric text cells in the metadata area
- [ ] When no header keywords are found, the parser falls back to finding the first row with numeric data in dimension columns
- [ ] Metadata rows (description, nominal, USL, LSL, tol.max, tol.min, point number) are detected from the label column dynamically, not from fixed positions
- [ ] A sheet that contains no `Dim. No.` marker is skipped without crashing

**Status:** Implemented

---

## US-004: Auto-detect metadata columns from the header row

**As a** quality engineer
**I want** every metadata column on the header row (Build, CFG, Color, Vendor Serial Number, Raw material, Fabric thickness, Shipment Date, etc.) to be picked up automatically — whatever the vendor labeled it
**So that** I can group/color/section charts by any metadata column without the developer having to hardcode column lists

**Acceptance criteria:**
- [ ] Parser reads every non-empty cell in the header row up to the first measurement column as a metadata column
- [ ] Columns whose name starts with `SPC_` or `DIM` are treated as measurement columns, not metadata
- [ ] `Shipment Date` values are converted to datetime automatically
- [ ] Detected metadata column names appear in the sidebar Color-by, Section-by, and Row-by selectors

**Status:** Implemented

---

## US-005: Open strict-OOXML xlsx files produced by some vendor tools

**As a** quality engineer
**I want** files exported by tools that use the strict OOXML namespace (or that contain broken `ExternalReference` entries) to open without errors
**So that** I am not blocked by "openpyxl cannot read this file" when a vendor sends me a slightly nonstandard xlsx

**Acceptance criteria:**
- [ ] A file that openpyxl opens with zero sheets is transparently rewritten from strict-OOXML namespaces to transitional namespaces in memory and retried
- [ ] `openpyxl.packaging.workbook.ExternalReference` is monkey-patched so a missing positional `id` argument does not crash parsing
- [ ] The `conformance="strict"` attribute is stripped during the rewrite

**Status:** Implemented

---

## US-006: Auto-detect vendor / factory from serial numbers

**As a** quality engineer
**I want** the app to identify the vendor (FX, TY, TRM, LK, FXJS…) for each file automatically from the Vendor Serial Number, SN prefix, or filename
**So that** I can section-by or color-by Factory without manually tagging each upload

**Acceptance criteria:**
- [ ] Factory is derived from the mode of `Vendor Serial Number` values when that column exists
- [ ] When `Vendor Serial Number` is missing, factory is extracted as the leading 2–4 uppercase letters of the SN column
- [ ] When neither column helps, factory is inferred from the first underscore-separated token of the filename if it matches `[A-Z]{2,4}`
- [ ] Detected factory code shows up next to each file in the "Loaded Files" summary

**Status:** Implemented

---

## US-007: Select dimensions individually or via preset groups

**As a** quality engineer
**I want** to pick which dimensions to chart, either one-by-one or via auto-detected keyword groups (e.g. "Flatness", "Z Straightness", "Overall Length", "All dimensions")
**So that** I can quickly compare every flatness dim at once without ticking each box manually

**Acceptance criteria:**
- [ ] Sidebar exposes a "Preset" dropdown with keyword-based groups plus "Custom" and "All dimensions"
- [ ] Selecting a preset auto-fills the dimension multi-select with matching dim numbers
- [ ] "Custom" mode lets the user check individual dimensions
- [ ] Each dimension label shows `SPC_XX — description` when a description is available
- [ ] At least one dimension must be selected or an informational message is shown

**Status:** Implemented

---

## US-008: Switch between Combined Profile, Box Plot, and Histogram

**As a** quality engineer
**I want** a radio selector to flip between three chart types for the same selected dimensions:
  - **Combined Profile** — one line per part across concatenated measurement points
  - **Box Plot** — distribution per measurement point
  - **Histogram** — frequency distribution of all values
**So that** I can look at the same data through per-unit profile, per-point spread, and overall distribution lenses

**Acceptance criteria:**
- [ ] Sidebar radio `Profile | Box Plot | Histogram` changes the main chart in place
- [ ] Combined Profile draws one line per unit across X-axis points; single-point dimensions (e.g. flatness with only 1 observation) render as markers instead of invisible lines
- [ ] Box Plot shows one box per measurement point, colored by the group-by selection
- [ ] Histogram shows overlaid frequency bars with a bin slider (10–100, default 40)
- [ ] Section-by controls only appear for Profile and Box Plot; bin slider only appears for Histogram

**Status:** Implemented

---

## US-009: See USL / LSL / Nominal spec limits on every chart

**As a** quality engineer
**I want** Upper Spec Limit (USL), Lower Spec Limit (LSL), and Nominal (target) lines drawn on every chart type, with a shaded in-spec band
**So that** I can visually judge whether measurements are within tolerance without cross-referencing the drawing

**Acceptance criteria:**
- [ ] Combined Profile draws dashed red horizontal lines for USL / LSL and a shaded green in-spec band when both exist
- [ ] Box Plot draws dashed USL / LSL lines and a dotted Nominal line
- [ ] Histogram draws vertical USL / LSL / Nominal reference lines with inline annotations
- [ ] USL / LSL numeric labels render on the Y-axis margin for Profile and Box Plot charts
- [ ] When only one spec bound exists, only that line is drawn (no crash)

**Status:** Implemented

---

## US-010: Group and facet the chart by any metadata column

**As a** quality engineer
**I want** independent controls to color-by (e.g. Raw material), section-by (e.g. Factory × Build — adds vertical dividers and a header band), and row-by (vertical faceting into subplots)
**So that** I can see cross-vendor or cross-build patterns at a glance

**Acceptance criteria:**
- [ ] Color-by dropdown lists every detected metadata column plus "None"
- [ ] Section-by is a multi-select; selecting multiple fields concatenates their values into one section label (e.g. `FX P1`, `TRM P2`)
- [ ] Section boundaries render as vertical gridlines with a gray header band showing the section label
- [ ] Row-by splits the chart into stacked subplots with shared X-axis when more than one row value exists
- [ ] Box Plot and Combined Profile support Section-by; Histogram does not (instead it facets by dimension/row)

**Status:** Implemented

---

## US-011: Exclude specific measurement points from the view

**As a** quality engineer
**I want** a multi-select of point numbers (C1, C2, C11-C12, …) that removes them from all chart types and statistics
**So that** I can ignore known bad probe points or interval-style aggregate points when comparing unit profiles

**Acceptance criteria:**
- [ ] "Exclude interval points" sidebar checkbox (on by default) drops labels matching `C\d+-C\d+`
- [ ] "Exclude points" multi-select lists every remaining point number across selected dimensions
- [ ] Excluded points disappear from Combined Profile X-axis, Box Plot categories, and Histogram input values
- [ ] Excluded points remain excluded when a new chart type is chosen (persistent across chart switches)
- [ ] Emptying the exclusion list returns to showing all points

**Status:** Implemented

---

## US-012: Custom Y-axis range and Deviation-from-Nominal mode

**As a** quality engineer
**I want** to toggle Y-axis between raw measurement values and deviation from nominal, and optionally pin a custom Y min/max
**So that** I can zoom into tolerance band behavior or normalize dims with different nominals onto one comparable scale

**Acceptance criteria:**
- [ ] Y-axis selector offers "Measurement values" and "Deviation from Nominal" (Combined Profile and Box Plot only)
- [ ] In deviation mode, USL, LSL, and the green band are shifted by the nominal so that `0` is target
- [ ] "Custom Y range" checkbox reveals numeric Min/Max inputs (4-decimal precision)
- [ ] Chart respects the custom range only when `min < max`; otherwise falls back to auto-scaling

**Status:** Implemented

---

## US-013: Choose custom colors per group

**As a** quality engineer
**I want** a color picker per group value (e.g. FX vs TY vs TRM, or RM coil A vs B) that overrides the default palette
**So that** I can keep consistent vendor colors across many screenshots in a report

**Acceptance criteria:**
- [ ] Sidebar "Colors" section shows one color picker per distinct value in the Color-by column
- [ ] Picker defaults to the next color in the built-in palette (blue, red, green, amber, cyan, rose, indigo, orange, teal, slate — no purple)
- [ ] Chosen colors apply to Combined Profile lines, Box Plot fill, and Histogram bars
- [ ] When Color-by is "None", a single "All" picker is shown

**Status:** Implemented

---

## US-014: Click a profile line to highlight one unit (JMP-style)

**As a** quality engineer
**I want** to click any line on the Combined Profile chart to bold it and dim all other lines
**So that** I can isolate an outlier unit's profile to read off individual point values in the hover tooltip

**Acceptance criteria:**
- [ ] Single-clicking a line on the Combined Profile chart raises its opacity to 1.0 and line width to 2.5, while dimming all other lines to opacity 0.08 / width 0.4
- [ ] Clicking the same line a second time restores the default view (all lines at opacity 0.45, width 0.7)
- [ ] Double-clicking anywhere on the plot resets the highlight
- [ ] Highlight behavior is only active on the Combined Profile chart, not Box Plot or Histogram

**Status:** Implemented

---

## US-015: Batch-export one chart per dimension as a ZIP of PNGs

**As a** quality engineer
**I want** to multi-select dimensions from a sidebar "Batch Chart Export" panel and download a ZIP containing one PNG per dimension, honoring all current controls (chart type, color, section, excludes)
**So that** I can paste a large number of per-dim charts into a weekly vendor review deck without clicking each one

**Acceptance criteria:**
- [ ] Sidebar expander "Batch Chart Export" with a dimension multi-select and an Export button
- [ ] Progress bar shows `Generating SPC_XX (i/N)` during export
- [ ] Exported PNGs are rendered via kaleido at 1400×700 @ scale=2
- [ ] Filenames follow `{ChartType}_{dim_no}_{description}.png` (safe chars only)
- [ ] Dimensions with no data, or that fail image conversion, are listed in a "Skipped" warning instead of aborting the whole export
- [ ] Final download button produces `SPC_Charts_{ChartType}.zip`

**Status:** Implemented

---

## US-016: Quick Test page — auto-load local files for dev iteration

**As a** quality engineer (or developer of this app)
**I want** a `Quick Test` page that auto-parses every `.xlsx` in the project directory with one click, using the same sidebar controls and chart components as the main page
**So that** I can verify chart behavior after code changes without repeatedly re-uploading the same files

**Acceptance criteria:**
- [ ] `Quick Test` appears as a Streamlit page in the multi-page sidebar
- [ ] Page scans the project root for `.xlsx` files, skipping Excel lock files (`~$*.xlsx`)
- [ ] A sidebar "Sheet" selector offers "Auto-detect" plus every detected sheet name
- [ ] Chart, dimension selector, exclude controls, and batch export all behave identically to the main page
- [ ] Widget state is kept separate from the main page via the `qt_` key prefix (no collision)

**Status:** Implemented

---

## US-017: Sheet Manager page — compare dimension coverage across two sheets

**As a** quality engineer
**I want** a `Sheet Manager` page where I upload files, pick any two (file / sheet) combos, and see which dimensions are shared, only-in-A, only-in-B, or fuzzy-matched
**So that** I can quickly confirm a new CORR file reports the same dimensions as the previous POR file before trusting the comparison

**Acceptance criteria:**
- [ ] Page accepts multi-file upload independent of the main page
- [ ] Two "File A" / "File B" selectors list every data sheet across uploaded files
- [ ] Summary metrics show Total, Shared, Only in A, Only in B
- [ ] Table rows sort Fuzzy → Missing → OK and color-code the status badge
- [ ] Dimensions close-but-not-identical (e.g. `SPC_A` vs `SPC_A-1`) are flagged Fuzzy with a hint showing the candidate match
- [ ] Filter radio (All / Mismatched / Fuzzy) and a free-text Search box narrow the table
- [ ] Fuzzy match block at the bottom lists each `dim_A ↔ dim_B` pair for manual review

**Status:** Implemented

---

## US-018: Persistent locally-hosted service on port 8504

**As a** quality engineer
**I want** the Streamlit app to be running permanently on `http://localhost:8504` as a launchd service
**So that** I can open a bookmark and start analyzing without remembering terminal commands or restarting the server after reboots

**Acceptance criteria:**
- [ ] A macOS launchd plist keeps the Streamlit server up on port 8504
- [ ] Service auto-restarts on crash
- [ ] Service survives machine reboot
- [ ] Logs are written to a known location for debugging

> [!question]
> Confirm exact plist path (`~/Library/LaunchAgents/...`), log file location,
> and whether the service should also run on login vs always.

**Status:** Partial

---

## US-019: Tolerate vendor typos and formatting variants in column / dimension names

**As a** quality engineer
**I want** the parser to match column and sheet names tolerantly — ignoring case, whitespace, punctuation, and minor spelling variants (e.g. `RM` vs `R.M.`, `Raw Data-PP` vs `Raw data-AP`, `Data Input-PP` vs `Data Input - PP`)
**So that** I don't have to manually rename columns or sheets every time a vendor slightly changes their template — the user called this "one of the most annoying things" when files come in

**Acceptance criteria:**
- [ ] Sheet-name matching ignores case and leading/trailing whitespace (already partially handled for the sheet multi-select merge)
- [ ] Metadata column matching ignores case, whitespace, and punctuation differences (e.g. `R.M.` matches `RM`, `Raw material` matches `raw_material`)
- [ ] Dimension description keyword grouping is robust to whitespace and punctuation variations
- [ ] A fuzzy match (edit-distance or normalized-token similarity) is attempted before declaring a column "missing"
- [ ] Log / debug info surfaces when a fuzzy match was used so the engineer can spot misreads

> [!question]
> Scope of tolerance: should the fuzzy match be silent, or should the app
> prompt the user to confirm before merging two near-identical column names?
> What's the acceptable similarity threshold (e.g. 0.85)?

**Status:** Planned

---

## Story summary

- **Total stories:** 19 (US-001 through US-019)
- **Implemented:** 17
- **Partial:** 1 (US-018 — launchd service details unconfirmed)
- **Planned:** 1 (US-019 — column/sheet-name fuzzy matching)

Stories flagged for clarification:
- **US-018** — launchd plist location, log path, on-login vs always-on
- **US-019** — fuzzy-match silent vs confirm, similarity threshold

## Deferred (future version)

The codebase contains CPK/PPK calculations, ANOVA, Nelson Rules, CUSUM, and EWMA inside a "Summary Statistics" expander in `shared_ui.py` (lines ~328-709). **Scope decision:** this app is a visualization tool only. Analysis features are deferred to a future version. The code is retained for now but should be considered dormant; Phase 1 refactor can either gate it behind a feature flag or move it to a separate module.
