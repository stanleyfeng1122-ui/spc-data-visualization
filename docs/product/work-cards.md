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
- **Expected:** Spec labels should stay readable and not collide with numeric y-axis labels.
- **Pass condition:** Spec labels are anchored just inside the plot area with left alignment.
- **Status:** Implemented, tests passing.
