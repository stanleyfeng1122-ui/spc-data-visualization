# SPC App Work Cards

Small ledger of bugs and improvements captured during real app use.

## Done

### WC-001: Default sheet selection should start with Data Input only

- **Type:** UX improvement
- **Priority:** P2
- **Lane:** Workflow Speed
- **Issue:** The app selected every detected sheet by default, forcing manual deletion before normal use.
- **Expected:** Uploaded files may still be scanned, but `Sheets to parse` should start with only `Data Input` when available.
- **Pass condition:** New uploads default to `Data Input`; users can still add other sheets manually.
- **Status:** Implemented, tests passing.

### WC-002: Dimension dropdown is too long to browse manually

- **Type:** UX improvement
- **Priority:** P2
- **Lane:** Workflow Speed
- **Issue:** The dimension multiselect can contain many dimensions, making manual scrolling slow.
- **Expected:** Search should happen inside the existing `Dimensions` dropdown field, not in a separate sidebar input.
- **Pass condition:** The separate `Find dimension` box is removed; the `Dimensions` multiselect remains the single place to type/search/select.
- **Status:** Corrected after user feedback, tests passing.

### WC-003: Refresh/code updates force repeated file upload

- **Type:** UX improvement
- **Priority:** P2
- **Lane:** Workflow Speed
- **Issue:** During development, code updates or browser refreshes require re-uploading files and repeating setup.
- **Expected:** The main app can load `.xlsx` files from a local folder so the same workbooks are available after refresh/code updates.
- **Pass condition:** Sidebar supports `Local folder` source, remembers the last folder path, and loads all non-temporary `.xlsx` files from that folder.
- **Status:** Implemented, tests passing.

### WC-004: Spec labels overlap y-axis tick labels

- **Type:** Chart rendering issue
- **Priority:** P1
- **Lane:** Chart Correctness
- **Issue:** Red `USL-*` and `LSL-*` labels are anchored at the plot's left edge with right alignment, so they extend into the y-axis tick-label area.
- **Expected:** Spec labels should stay on the left side, remain readable, and not collide with numeric y-axis labels.
- **Pass condition:** Spec labels are anchored just inside the left plot area with left alignment and a small x-shift.
- **Status:** Revised after user feedback, tests passing.

### WC-005: Multi-factor Section-by should render as nested JMP-style headers

- **Type:** Chart rendering improvement
- **Priority:** P1
- **Lane:** Chart Correctness
- **Issue:** Selecting two `Section-by` factors currently flattens them into one combined header label such as `LK INN`.
- **Expected:** Selected section factors should render as stacked table-like header bands, similar to JMP.
- **Pass condition:** `Section-by = [Factory, RM]` renders separate field rows and value rows such as `Factory`, `LK` / `FJS`, `RM`, `INN` / `OUT`; 3-level groupings also render all levels.
- **Status:** Implemented, tests passing.

### WC-006: Single-point profile dots should be centered in each section

- **Type:** Chart rendering issue
- **Priority:** P1
- **Lane:** Chart Correctness
- **Issue:** Single-point dimensions plot each dot stack at the left edge of its section, making the chart look unbalanced under the section header.
- **Expected:** Single-point dot stacks should sit at the center of their section/point slot.
- **Pass condition:** X positions use centered point slots, so the first single-point stack renders at `0.5` instead of `0`.
- **Status:** Implemented, tests passing.

### WC-007: Polish JMP-style nested header readability

- **Type:** Chart rendering improvement
- **Priority:** P2
- **Lane:** Chart Correctness / Visual polish
- **Issue:** Nested headers work structurally, but field names, row hierarchy, and dividers need clearer visual hierarchy.
- **Expected:** Field-name rows should be visually distinct from value rows, row spacing should be readable, and major/minor dividers should have different weights.
- **Pass condition:** Multi-factor headers use taller bands, stronger field-row styling, dynamic top margin, and hierarchy-weighted vertical dividers.
- **Status:** Implemented, tests passing.
