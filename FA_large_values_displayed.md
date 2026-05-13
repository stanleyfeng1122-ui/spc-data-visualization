# Failure Analysis: Large Values Displayed on SPC Chart

## Symptom
Chart shows measurement values (e.g., 13, 19) that appear far above the specification limits (USL 0.8, LSL -0.8) for dimensions such as "z straightness of left". Values appear to be 10x-20x the spec range.

## How the Data Is Read

### 1. Excel Parsing (spc_parser.py)
- **Layout detection**: Scans for "Dim. No." row, then "Point Number", "Nominal", "USL", "LSL" rows. First column with `SPC_` starts dimension data.
- **Per-column metadata**: For each column `ci` in the dimension block, reads:
  - `col_point[ci]` = Point label (e.g. C11, C24)
  - `col_nominal[ci]` = Nominal value
  - `col_usl[ci]` = Upper spec limit
  - `col_lsl[ci]` = Lower spec limit
- **Data rows**: Starting after "Start Point" (or "SN") row, each row is one part. Cell `(row, ci)` = measured value for that point.
- **Stat row exclusion**: Rows whose "Vendor Serial Number" or "Start Point" equals "Mean", "Std Dev", "Cp", "Cpk", etc. are skipped.

### 2. Overall Column Filter (lines 414-425)
If a dimension has 3+ columns and the **first** column:
- Has no point number, and
- Has different USL from the second column,

then the first column is dropped (e.g. "overall flatness" vs individual points).

### 3. Chart Rendering (chart_utils.py)
- Y-values: `grp_df.iloc[ri]` for selected columns (col_labels). No scaling.
- Deviation mode: `y_vals = y_vals - nom_array`.
- USL/LSL lines: Use `usl_rep` / `lsl_rep` from the first non-None value across all points of the dimension.

## Root Cause Hypotheses

| # | Hypothesis | Likelihood | How to Verify |
|---|------------|------------|---------------|
| 1 | **Wrong column mapping** – Merged cells or misaligned headers cause Point/USL/LSL to be read from different columns than the data. | High | Inspect Excel: compare `Point` row, `USL` row, and data column indices. |
| 2 | **First "overall" column included** – First column has large values (e.g. overall flatness) but same USL as points, so it is not dropped. Chart uses point USL (0.8) but plots overall column data (13). | High | Check if first point has same USL as others; compare first-column values vs rest. |
| 3 | **Units mismatch** – Data in microns (13000) or different scale; specs in mm. Parser reads raw cell values. | Medium | Check units in Excel header/footer; compare numeric scale to spec. |
| 4 | **Profile dimension layout** – "z straightness" uses a different layout (e.g. row-per-point, or transposed). Parser assumes row-per-part, column-per-point. | Medium | Compare Excel layout to parser assumptions. |
| 5 | **Multiple dimension blocks** – Dim. No. row has SPC_A in cols 10–30, but Point/USL rows align to a different block. Column indices map to wrong specs. | Medium | Verify Dim. No., Point, USL/LSL columns align for each SPC block. |
| 6 | **Stat row not excluded** – A row (e.g. "Maximum") has large values and is not in STAT_ROW_LABELS, so it is treated as data. | Low | Add "maximum", "minimum" to STAT_ROW_LABELS if needed; check ID column values. |

## Recommended Actions

1. **Add Data Inspection panel** – Show for selected dimension: Point, Nominal, USL, LSL, and sample min/max from parsed data. Compare specs vs actual value range.
2. **Expand overall-column filter** – When first column has a point label like "Overall", "0", or "Max" and its USL differs from the median of other points, consider dropping it.
3. **Inspect source Excel** – For the problematic dimension, confirm:
   - Which columns belong to that dimension.
   - Whether Point, USL, LSL rows align to those columns.
   - Whether the first column is an overall/summary column.
4. **Check STAT_ROW_LABELS** – Ensure all summary rows (Mean, Max, Min, Std Dev, etc.) used in the file format are excluded.
