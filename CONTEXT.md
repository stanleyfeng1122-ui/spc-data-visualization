# Domain Glossary — SPC Data Visualization

The ubiquitous language of this codebase. Use these terms in code, tests,
docs, and commit messages; don't invent synonyms.

- **SPC bubble (bubble id)** — the vendor's dimension identifier on the
  drawing, e.g. `SPC_GG`. One bubble usually owns a contiguous run of
  measurement-point columns in a sheet. Vendors sometimes typo the
  description on some of a bubble's columns (the bubble id governs), and
  occasionally reuse a bubble id for an unrelated feature far away in the
  same sheet (a non-contiguous column is foreign — dropped).
- **Dimension** — one measured feature (e.g. "Left Side Edge Straightness"),
  identified by a bubble id and a description, carrying per-point specs.
  `DimensionMeta` is its record.
- **Measurement point** — one probed location within a dimension (`C76`,
  `P1`, `PS35`). A dimension's points form the x-axis of profile charts.
- **Level (PP / AP)** — process stage a sheet's data comes from: PP
  (pre-process) vs AP (after-process). Detected from the sheet name.
- **Condition (POR / CORR)** — the build condition a sheet covers: plan of
  record vs correlation. Also detected from the sheet name.
- **Pairing** — merging the *same physical feature* measured at PP and AP
  under different bubble ids (e.g. `SPC_BA` in PP + `SPC_AU` in AP) into one
  virtual dimension (`PAIR::…`) so the PP→AP change can be tracked. Pairing
  fires only when the members span both levels and never co-occur in one
  sheet; two bubbles sharing a description inside one sheet are different
  features and are never merged.
- **Section** — a vertical slice of a chart's x-axis grouping parts by a
  metadata factor (Factory, CFG, Level…). Rendered as header bands.
- **Spec limits** — USL / LSL / Nominal for a point or dimension.
- **Spec span** — a spec over one x-range of a chart. A uniform spec is a
  single full-width span; a *stepping* spec is several spans with different
  USL/LSL (e.g. multiple dims or per-section specs). `SpecSpan` in
  `charts/spec_limits.py` is the record; all four chart builders render
  specs through that one module.
- **Factory / vendor code** — short site prefix (FXJS, FXVN, LK…) detected
  from vendor serial numbers or the filename.
- **SpcDataset** — the one deep interface over everything parsed from the
  uploaded workbooks (`parsers/dataset.py`): `load_dataset`/`parse_sheets` +
  `assemble_dataset` own parsing, pairing and source metadata; the app layer
  consumes `.dimensions`, `.display_labels()`, `.meta_columns`,
  `.combined(dim_nos)` and `.file_summaries()` and never touches the raw
  parsed-file records.
