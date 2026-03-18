"""
Sheet Manager — Guided workflow for controlled sheet selection, validation,
dimension aliasing, and vendor comparison.

Steps: Upload → Select Sheets → Validate → Map Dims → Compare Vendors
"""

import io
import os
import sys
import hashlib
from collections import OrderedDict
from difflib import SequenceMatcher

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from spc_parser import parse_excel_multi, get_filtered_dim_meta
from chart_utils import (
    prepare_combined_data,
    calc_process_capability,
    get_color_for_group,
)
from ui_theme import (
    inject_theme, FONT_MONO, FONT_BODY, FONT_HEADING,
    TEXT_PRIMARY, TEXT_SECONDARY, TEXT_MUTED,
    ACCENT, ACCENT_HOVER, DANGER, SUCCESS, WARNING,
    WHITE, BG_SUBTLE, BORDER, BORDER_LIGHT,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Sheet Manager — SPC", layout="wide", initial_sidebar_state="collapsed")
inject_theme()

# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------
_DEFAULTS = {
    "sm_uploaded_bytes": {},       # {filename: bytes}
    "sm_sheet_infos": [],          # list of SheetInfo dicts
    "sm_parsed_files": [],         # list of parsed file dicts (after selection)
    "sm_validation_issues": [],    # list of Issue dicts
    "sm_dim_aliases": {},          # {old_name: canonical_name}
    "sm_step": 0,                  # current workflow step
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ===================================================================
# HELPER FUNCTIONS
# ===================================================================

def scan_sheets(file_bytes: bytes, filename: str) -> list:
    """Lightweight scanner: detect data sheets, count dims/rows, read part#.

    Returns list of SheetInfo dicts without doing a full parse.
    """
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    _NON_DATA_PREFIXES = ("BoxPlotCht", "Histo ")
    _NON_DATA_EXACT = {"Histo Pivot", "Histo Listbox", "Histo Curve"}
    _NON_DATA_KEYWORDS = {"color", "cosmetic", "waive"}

    results = []
    for sn in wb.sheetnames:
        # Skip known non-data sheets
        if sn in _NON_DATA_EXACT or any(sn.startswith(p) for p in _NON_DATA_PREFIXES):
            continue

        ws = wb[sn]
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0

        if max_row < 5 or max_col < 5:
            continue

        # Find "Dim. No." row (scan first 30 rows × 30 cols)
        dim_row = None
        dim_label_col = None
        for r in range(1, min(31, max_row + 1)):
            for c in range(1, min(31, max_col + 1)):
                v = ws.cell(r, c).value
                if v and isinstance(v, str) and v.strip().lower().replace(".", "").replace(" ", "") in (
                    "dimno", "dim no", "dimno"
                ):
                    dim_row = r
                    dim_label_col = c
                    break
            if dim_row:
                break

        if dim_row is None:
            # Check if it's a flat table (e.g., "color" sheet with SN, Point, Gloss columns)
            # Still useful to list but mark as non-standard
            if any(kw in sn.lower() for kw in _NON_DATA_KEYWORDS):
                continue
            # No Dim. No. row — skip
            continue

        # Read unique dims from dim row
        dims = []
        first_values = []
        for c in range(dim_label_col + 1, min(max_col + 1, dim_label_col + 500)):
            v = ws.cell(dim_row, c).value
            if v is not None and str(v).strip():
                s = str(v).strip()
                if s.lower() not in ("dim. no.", "dim.no.", "dim no"):
                    dims.append(s)
        unique_dims = list(dict.fromkeys(dims))  # preserve order, dedupe

        # Read part number (usually row 2, 1-2 cols after "Part Number :")
        part_number = ""
        for r in range(1, min(6, max_row + 1)):
            for c in range(1, min(31, max_col + 1)):
                v = ws.cell(r, c).value
                if v and isinstance(v, str) and "part number" in v.lower():
                    # Part number is typically in the next column
                    pn = ws.cell(r, c + 1).value
                    if pn:
                        part_number = str(pn).strip()
                    break

        # Count data rows (rows after dim_row + ~30 header rows that have data)
        data_start = dim_row + 25  # approximate: header metadata is ~25 rows
        data_row_count = 0
        empty_streak = 0
        for r in range(data_start, max_row + 1):
            has_data = False
            for c in range(dim_label_col, min(dim_label_col + 10, max_col + 1)):
                v = ws.cell(r, c).value
                if v is not None and v != "":
                    has_data = True
                    break
            if has_data:
                data_row_count += 1
                empty_streak = 0
            else:
                empty_streak += 1
                if empty_streak > 10:
                    break

        # Read a few sample values for fingerprinting
        sample_vals = []
        if data_row_count > 0:
            for c in range(dim_label_col + 1, min(dim_label_col + 6, max_col + 1)):
                v = ws.cell(data_start, c).value
                if v is not None:
                    sample_vals.append(str(v)[:20])

        results.append({
            "filename": filename,
            "sheet_name": sn,
            "dim_count": len(unique_dims),
            "dims_list": unique_dims,
            "dims_set": set(unique_dims),
            "data_rows": data_row_count,
            "part_number": part_number,
            "sample_vals": sample_vals,
            "selected": True,
            "duplicate_of": None,
        })

    wb.close()
    return results


def detect_duplicates(sheet_infos: list) -> list:
    """Flag sheets with >90% dimension overlap within the same file."""
    # Group by filename
    by_file = {}
    for si in sheet_infos:
        by_file.setdefault(si["filename"], []).append(si)

    for fname, sheets in by_file.items():
        if len(sheets) < 2:
            continue
        for i in range(len(sheets)):
            for j in range(i + 1, len(sheets)):
                a, b = sheets[i], sheets[j]
                if not a["dims_set"] or not b["dims_set"]:
                    continue
                overlap = len(a["dims_set"] & b["dims_set"])
                union = len(a["dims_set"] | b["dims_set"])
                if union == 0:
                    continue
                jaccard = overlap / union
                if jaccard > 0.85:
                    # Flag the one with fewer dims or fewer rows
                    if a["dim_count"] < b["dim_count"] or (
                        a["dim_count"] == b["dim_count"] and a["data_rows"] < b["data_rows"]
                    ):
                        a["duplicate_of"] = b["sheet_name"]
                        a["selected"] = False
                    else:
                        b["duplicate_of"] = a["sheet_name"]
                        b["selected"] = False
    return sheet_infos


def validate_parsed_files(parsed_files: list) -> list:
    """Run validation checks across parsed files. Returns list of Issue dicts."""
    issues = []

    # Collect all dims across all files
    dim_specs = {}  # dim_no → list of {filename, sheet, usl, lsl, point_count}
    for pf in parsed_files:
        for dno, dmeta in pf["dimensions"].items():
            usl_vals = [v for v in dmeta.usl if v is not None]
            lsl_vals = [v for v in dmeta.lsl if v is not None]
            usl = usl_vals[0] if usl_vals else None
            lsl = lsl_vals[0] if lsl_vals else None
            dim_specs.setdefault(dno, []).append({
                "filename": pf["filename"],
                "sheet": pf["sheet_name"],
                "usl": usl,
                "lsl": lsl,
                "point_count": len(dmeta.col_labels),
            })

    for dno, entries in dim_specs.items():
        # Missing specs
        for e in entries:
            if e["usl"] is None and e["lsl"] is None:
                issues.append({
                    "level": "warning",
                    "category": "Missing Specs",
                    "dimension": dno,
                    "message": f"No USL or LSL in {e['filename']} / {e['sheet']}",
                })

        # Conflicting specs across files
        if len(entries) > 1:
            usls = [e["usl"] for e in entries if e["usl"] is not None]
            lsls = [e["lsl"] for e in entries if e["lsl"] is not None]
            if usls and len(set(round(v, 6) for v in usls)) > 1:
                vals_str = ", ".join(f"{e['filename']}: {e['usl']}" for e in entries if e["usl"] is not None)
                issues.append({
                    "level": "error",
                    "category": "Conflicting USL",
                    "dimension": dno,
                    "message": f"USL differs across files: {vals_str}",
                })
            if lsls and len(set(round(v, 6) for v in lsls)) > 1:
                vals_str = ", ".join(f"{e['filename']}: {e['lsl']}" for e in entries if e["lsl"] is not None)
                issues.append({
                    "level": "error",
                    "category": "Conflicting LSL",
                    "dimension": dno,
                    "message": f"LSL differs across files: {vals_str}",
                })

            # Point count mismatch
            counts = [e["point_count"] for e in entries]
            if max(counts) > 2 * min(counts) and min(counts) > 0:
                issues.append({
                    "level": "warning",
                    "category": "Point Count Mismatch",
                    "dimension": dno,
                    "message": f"Point counts vary: " + ", ".join(
                        f"{e['filename']}: {e['point_count']} pts" for e in entries
                    ),
                })

    # Zero-row sheets
    for pf in parsed_files:
        if pf["data"] is None or len(pf["data"]) == 0:
            issues.append({
                "level": "error",
                "category": "Empty Sheet",
                "dimension": "—",
                "message": f"{pf['filename']} / {pf['sheet_name']} has 0 data rows",
            })

    # Sort: errors first, then warnings
    order = {"error": 0, "warning": 1, "info": 2}
    issues.sort(key=lambda x: order.get(x["level"], 9))
    return issues


def find_similar_dims(dim_list: list, threshold: float = 0.80) -> list:
    """Find dimension name pairs that look similar (likely vendor typos)."""
    pairs = []
    for i in range(len(dim_list)):
        for j in range(i + 1, len(dim_list)):
            a, b = dim_list[i], dim_list[j]
            if a == b:
                continue
            # Normalize for comparison
            a_norm = a.replace("-", "_").replace(" ", "").upper()
            b_norm = b.replace("-", "_").replace(" ", "").upper()
            if a_norm == b_norm:
                pairs.append((a, b, 1.0))
                continue
            score = SequenceMatcher(None, a_norm, b_norm).ratio()
            if score >= threshold:
                pairs.append((a, b, round(score, 3)))
    pairs.sort(key=lambda x: -x[2])
    return pairs


def apply_dim_aliases(parsed_files: list, aliases: dict) -> list:
    """Apply dimension aliases: rename dims in parsed data before merge.
    Returns modified copies (does not mutate originals).
    """
    if not aliases:
        return parsed_files

    result = []
    for pf in parsed_files:
        pf = dict(pf)  # shallow copy
        new_dims = OrderedDict()
        rename_map = {}  # old_col_label → new_col_label

        for dno, dmeta in pf["dimensions"].items():
            canonical = aliases.get(dno, dno)
            if canonical != dno:
                # Need to rename column labels too
                for cl in dmeta.col_labels:
                    new_label = cl.replace(dno, canonical, 1)
                    rename_map[cl] = new_label

            if canonical not in new_dims:
                new_dims[canonical] = dmeta
            # If canonical already exists, we skip (first occurrence wins)

        pf["dimensions"] = new_dims
        if rename_map and pf["data"] is not None:
            pf["data"] = pf["data"].rename(columns=rename_map)
        result.append(pf)
    return result


# ===================================================================
# MAIN PAGE
# ===================================================================

st.markdown(
    f"<h1 style='font-family:{FONT_HEADING};font-size:1.5rem;font-weight:700;"
    f"color:{TEXT_PRIMARY};margin-bottom:0;'>Sheet Manager</h1>"
    f"<p style='font-size:0.8rem;color:{TEXT_MUTED};margin-top:2px;'>"
    f"Upload → Select Sheets → Validate → Map Dims → Compare</p>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Step 0: File Upload
# ---------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload CPK Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="sm_uploader",
)

if not uploaded_files:
    st.info("Upload one or more .xlsx files to begin.")
    st.stop()

# Read file bytes (cache in session state to avoid re-reading)
file_bytes_map = {}
for uf in uploaded_files:
    raw = uf.read()
    uf.seek(0)
    file_bytes_map[uf.name] = raw

# ---------------------------------------------------------------------------
# Scan all sheets
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Scanning sheets...")
def _scan_all(file_data: dict) -> list:
    all_infos = []
    for fname, fbytes in file_data.items():
        infos = scan_sheets(fbytes, fname)
        all_infos.extend(infos)
    return detect_duplicates(all_infos)

# Convert to hashable for cache
_file_data_for_cache = {k: v for k, v in file_bytes_map.items()}
sheet_infos = _scan_all(_file_data_for_cache)

if not sheet_infos:
    st.warning("No data sheets found in uploaded files.")
    st.stop()


# ===================================================================
# TABBED WORKFLOW
# ===================================================================
tab_select, tab_validate, tab_alias, tab_compare = st.tabs([
    "1. Select Sheets",
    "2. Validate",
    "3. Map Dimensions",
    "4. Compare Vendors",
])


# ---------------------------------------------------------------------------
# TAB 1: Sheet Selection Panel
# ---------------------------------------------------------------------------
with tab_select:
    st.markdown(
        f"<p style='font-size:0.82rem;color:{TEXT_SECONDARY};margin-bottom:8px;'>"
        f"Found <b>{len(sheet_infos)}</b> data sheets across "
        f"<b>{len(file_bytes_map)}</b> files. Uncheck sheets to exclude.</p>",
        unsafe_allow_html=True,
    )

    # Build editable dataframe
    df_sheets = pd.DataFrame([
        {
            "Include": si["selected"],
            "File": si["filename"][:40] + "..." if len(si["filename"]) > 40 else si["filename"],
            "Sheet": si["sheet_name"],
            "Rows": si["data_rows"],
            "Dims": si["dim_count"],
            "Part #": si["part_number"],
            "Note": f"Possible dup of '{si['duplicate_of']}'" if si["duplicate_of"] else "",
        }
        for si in sheet_infos
    ])

    edited_df = st.data_editor(
        df_sheets,
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", default=True),
            "Rows": st.column_config.NumberColumn("Rows", format="%d"),
            "Dims": st.column_config.NumberColumn("Dims", format="%d"),
        },
        disabled=["File", "Sheet", "Rows", "Dims", "Part #", "Note"],
        hide_index=True,
        use_container_width=True,
        key="sm_sheet_editor",
    )

    # Update selections
    for i, si in enumerate(sheet_infos):
        if i < len(edited_df):
            si["selected"] = bool(edited_df.iloc[i]["Include"])

    selected_count = sum(1 for si in sheet_infos if si["selected"])
    st.caption(f"{selected_count} of {len(sheet_infos)} sheets selected")

    # Peek expanders
    with st.expander("Peek at sheet dimensions", expanded=False):
        for si in sheet_infos:
            if si["selected"]:
                dims_str = ", ".join(si["dims_list"][:15])
                if len(si["dims_list"]) > 15:
                    dims_str += f" ... (+{len(si['dims_list']) - 15} more)"
                st.markdown(
                    f"<div style='padding:4px 0;border-bottom:1px solid {BORDER_LIGHT};'>"
                    f"<b style='font-size:0.78rem;'>{si['sheet_name']}</b>"
                    f"<span style='color:{TEXT_MUTED};font-size:0.72rem;margin-left:8px;'>"
                    f"{si['filename'][:30]}</span><br>"
                    f"<span style='font-family:{FONT_MONO};font-size:0.7rem;color:{TEXT_SECONDARY};'>"
                    f"{dims_str}</span></div>",
                    unsafe_allow_html=True,
                )


# ---------------------------------------------------------------------------
# Parse selected sheets
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Parsing selected sheets...")
def _parse_selected(file_data: dict, selections: list) -> list:
    """Parse only the selected sheets."""
    parsed = []
    # Group selections by filename
    by_file = {}
    for sel in selections:
        by_file.setdefault(sel["filename"], []).append(sel["sheet_name"])

    for fname, sheet_names in by_file.items():
        fbytes = file_data.get(fname)
        if not fbytes:
            continue
        buf = io.BytesIO(fbytes)
        buf.name = fname
        for sn in sheet_names:
            try:
                parsed_list = parse_excel_multi(buf, sheet_name=sn)
                for p in parsed_list:
                    parsed.append({
                        "filename": p.filename,
                        "sheet_name": p.sheet_name,
                        "part_number": p.part_number,
                        "part_description": p.part_description,
                        "revision": p.revision,
                        "factory": p.factory,
                        "dimensions": p.dimensions,
                        "data": p.data,
                        "meta_columns": p.meta_columns,
                    })
                buf.seek(0)
            except Exception as e:
                st.warning(f"Error parsing {fname} / {sn}: {e}")
                buf.seek(0)
    return parsed

selected_sheets = [si for si in sheet_infos if si["selected"]]
if not selected_sheets:
    st.warning("No sheets selected. Check at least one sheet in the Select tab.")
    st.stop()

# Create a hashable key from selections
_sel_key = tuple((s["filename"], s["sheet_name"]) for s in selected_sheets)
parsed_files = _parse_selected(_file_data_for_cache, selected_sheets)

if not parsed_files:
    st.warning("No data parsed from selected sheets.")
    st.stop()


# ---------------------------------------------------------------------------
# TAB 2: Pre-Merge Validation
# ---------------------------------------------------------------------------
with tab_validate:
    issues = validate_parsed_files(parsed_files)

    errors = [i for i in issues if i["level"] == "error"]
    warnings = [i for i in issues if i["level"] == "warning"]
    infos = [i for i in issues if i["level"] == "info"]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Errors", len(errors))
    col2.metric("Warnings", len(warnings))
    col3.metric("Files", len(parsed_files))
    all_dims = set()
    for pf in parsed_files:
        all_dims.update(pf["dimensions"].keys())
    col4.metric("Dimensions", len(all_dims))

    if not issues:
        st.markdown(
            f"<div style='padding:12px;background:{BG_SUBTLE};border:1px solid {BORDER_LIGHT};"
            f"border-left:3px solid {SUCCESS};border-radius:2px;'>"
            f"<span style='color:{SUCCESS};font-weight:600;'>All checks passed</span>"
            f"<span style='color:{TEXT_MUTED};font-size:0.82rem;margin-left:8px;'>"
            f"No spec conflicts, missing data, or anomalies detected.</span></div>",
            unsafe_allow_html=True,
        )
    else:
        for issue in issues:
            color = DANGER if issue["level"] == "error" else WARNING if issue["level"] == "warning" else TEXT_MUTED
            icon = "!!" if issue["level"] == "error" else "!" if issue["level"] == "warning" else "i"
            st.markdown(
                f"<div style='padding:6px 10px;margin:4px 0;background:{BG_SUBTLE};"
                f"border:1px solid {BORDER_LIGHT};border-left:3px solid {color};border-radius:2px;'>"
                f"<span style='font-weight:600;color:{color};font-size:0.78rem;'>"
                f"[{issue['category']}]</span> "
                f"<span style='font-family:{FONT_MONO};font-size:0.72rem;color:{ACCENT};'>"
                f"{issue['dimension']}</span><br>"
                f"<span style='font-size:0.75rem;color:{TEXT_SECONDARY};'>{issue['message']}</span>"
                f"</div>",
                unsafe_allow_html=True,
            )


# ---------------------------------------------------------------------------
# TAB 3: Dimension Aliasing
# ---------------------------------------------------------------------------
with tab_alias:
    all_dim_names = sorted(all_dims)

    if len(all_dim_names) < 2:
        st.info("Need 2+ unique dimension names for alias detection.")
    else:
        similar_pairs = find_similar_dims(all_dim_names, threshold=0.78)

        if not similar_pairs:
            st.markdown(
                f"<div style='padding:12px;background:{BG_SUBTLE};border:1px solid {BORDER_LIGHT};"
                f"border-left:3px solid {SUCCESS};border-radius:2px;'>"
                f"<span style='color:{SUCCESS};font-weight:600;'>No similar dimension names found</span>"
                f"<span style='color:{TEXT_MUTED};font-size:0.82rem;margin-left:8px;'>"
                f"All {len(all_dim_names)} dimensions have distinct names.</span></div>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"<p style='font-size:0.82rem;color:{TEXT_SECONDARY};'>"
                f"Found <b>{len(similar_pairs)}</b> potential matches. "
                f"Check the ones you want to merge.</p>",
                unsafe_allow_html=True,
            )

            aliases = dict(st.session_state.get("sm_dim_aliases", {}))
            for idx, (a, b, score) in enumerate(similar_pairs):
                pct = int(score * 100)
                merge = st.checkbox(
                    f"Merge '{a}' → '{b}' ({pct}% match)",
                    value=a in aliases,
                    key=f"sm_alias_{idx}",
                )
                if merge:
                    aliases[a] = b
                elif a in aliases:
                    del aliases[a]

            st.session_state["sm_dim_aliases"] = aliases

            if aliases:
                st.markdown(
                    f"<p style='font-size:0.75rem;color:{TEXT_MUTED};margin-top:8px;'>"
                    f"Active aliases: {', '.join(f'{k}→{v}' for k, v in aliases.items())}</p>",
                    unsafe_allow_html=True,
                )

    # Show full dimension inventory
    with st.expander(f"All dimensions ({len(all_dim_names)})", expanded=False):
        # Show which files have each dimension
        dim_file_map = {}
        for pf in parsed_files:
            for dno in pf["dimensions"]:
                dim_file_map.setdefault(dno, []).append(pf["filename"][:25])

        rows = []
        for dno in all_dim_names:
            files = dim_file_map.get(dno, [])
            rows.append({
                "Dimension": dno,
                "Files": len(set(files)),
                "Sources": ", ".join(sorted(set(files))),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# TAB 4: Vendor Comparison View
# ---------------------------------------------------------------------------
with tab_compare:
    # Apply aliases
    aliases = st.session_state.get("sm_dim_aliases", {})
    working_files = apply_dim_aliases(parsed_files, aliases)

    # Rebuild dimension map after aliasing
    compare_dims = OrderedDict()
    for pf in working_files:
        for dno, dmeta in pf["dimensions"].items():
            if dno not in compare_dims:
                compare_dims[dno] = dmeta

    dim_options = list(compare_dims.keys())
    if not dim_options:
        st.info("No dimensions available for comparison.")
        st.stop()

    selected_dim = st.selectbox("Dimension", options=dim_options, key="sm_compare_dim")

    if selected_dim:
        # Merge data for this dimension
        df, dim_metas = prepare_combined_data(working_files, [selected_dim])

        if df is None or df.empty:
            st.warning(f"No data for {selected_dim}.")
        else:
            dmeta = dim_metas.get(selected_dim)
            if not dmeta:
                st.warning("Dimension metadata not found.")
                st.stop()

            col_labels = dmeta.col_labels
            valid_cols = [c for c in col_labels if c in df.columns]
            if not valid_cols:
                st.warning("No measurement columns found.")
                st.stop()

            # Get USL/LSL
            usl_vals = [v for v in dmeta.usl if v is not None]
            lsl_vals = [v for v in dmeta.lsl if v is not None]
            usl = usl_vals[0] if usl_vals else None
            lsl = lsl_vals[0] if lsl_vals else None

            # Group by factory/source
            group_col = "_factory"
            if group_col not in df.columns:
                group_col = "_source_file"
            groups = df[group_col].fillna("Unknown").astype(str)
            unique_groups = sorted(groups.unique())

            # --- Stats table ---
            st.markdown(
                f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
                f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:12px 0 4px;'>"
                f"Process Capability by Vendor</div>",
                unsafe_allow_html=True,
            )

            stats_rows = []
            for g in unique_groups:
                mask = groups == g
                vals = df.loc[mask, valid_cols].values.flatten()
                vals = pd.Series(vals).dropna()
                if len(vals) < 2:
                    stats_rows.append({"Vendor": g, "n": len(vals), "Mean": "—", "Std": "—",
                                       "Cpk": "—", "Ppk": "—", "Rating": "—"})
                    continue
                cap = calc_process_capability(vals, usl, lsl)
                if cap:
                    cpk = cap.get("Cpk", cap.get("Cpk (upper)", cap.get("Cpk (lower)", None)))
                    ppk = cap.get("Ppk", None)
                    if cpk is not None:
                        if cpk >= 1.67:
                            rating = "EXCELLENT"
                        elif cpk >= 1.33:
                            rating = "GOOD"
                        elif cpk >= 1.0:
                            rating = "MARGINAL"
                        else:
                            rating = "POOR"
                    else:
                        rating = "N/A"
                    stats_rows.append({
                        "Vendor": g,
                        "n": cap.get("n", len(vals)),
                        "Mean": cap.get("mean", "—"),
                        "Std": cap.get("std", "—"),
                        "Cpk": cpk if cpk is not None else "—",
                        "Ppk": ppk if ppk is not None else "—",
                        "Rating": rating,
                    })
                else:
                    stats_rows.append({"Vendor": g, "n": len(vals), "Mean": round(vals.mean(), 6),
                                       "Std": round(vals.std(ddof=1), 6), "Cpk": "—", "Ppk": "—", "Rating": "N/A"})

            st.dataframe(pd.DataFrame(stats_rows), use_container_width=True, hide_index=True)

            # --- Overlapping histograms ---
            st.markdown(
                f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
                f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:12px 0 4px;'>"
                f"Distribution Comparison</div>",
                unsafe_allow_html=True,
            )

            fig_hist = go.Figure()
            for gi, g in enumerate(unique_groups):
                mask = groups == g
                vals = df.loc[mask, valid_cols].values.flatten()
                vals = pd.Series(vals).dropna()
                if len(vals) == 0:
                    continue
                color = get_color_for_group(gi)
                fig_hist.add_trace(go.Histogram(
                    x=vals.values,
                    name=g,
                    opacity=0.6,
                    marker_color=color,
                    nbinsx=40,
                ))

            # Add spec lines
            if usl is not None:
                fig_hist.add_vline(x=usl, line_dash="dash", line_color=DANGER,
                                   annotation_text=f"USL {usl:.4g}")
            if lsl is not None:
                fig_hist.add_vline(x=lsl, line_dash="dash", line_color=DANGER,
                                   annotation_text=f"LSL {lsl:.4g}")

            fig_hist.update_layout(
                barmode="overlay",
                xaxis_title="Value",
                yaxis_title="Count",
                paper_bgcolor=WHITE,
                plot_bgcolor=WHITE,
                font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
                height=350,
                margin=dict(l=40, r=20, t=30, b=40),
                xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                legend=dict(font=dict(size=11)),
            )
            st.plotly_chart(fig_hist, use_container_width=True, key="sm_hist")

            # --- Box plot ---
            st.markdown(
                f"<div style='font-size:0.72rem;font-weight:600;text-transform:uppercase;"
                f"letter-spacing:0.06em;color:{TEXT_MUTED};margin:12px 0 4px;'>"
                f"Box Plot by Vendor</div>",
                unsafe_allow_html=True,
            )

            fig_box = go.Figure()
            for gi, g in enumerate(unique_groups):
                mask = groups == g
                vals = df.loc[mask, valid_cols].values.flatten()
                vals = pd.Series(vals).dropna()
                if len(vals) == 0:
                    continue
                color = get_color_for_group(gi)
                fig_box.add_trace(go.Box(
                    y=vals.values,
                    name=g,
                    marker_color=color,
                    boxmean="sd",
                ))

            if usl is not None:
                fig_box.add_hline(y=usl, line_dash="dash", line_color=DANGER,
                                  annotation_text=f"USL {usl:.4g}")
            if lsl is not None:
                fig_box.add_hline(y=lsl, line_dash="dash", line_color=DANGER,
                                  annotation_text=f"LSL {lsl:.4g}")

            fig_box.update_layout(
                yaxis_title="Value",
                paper_bgcolor=WHITE, plot_bgcolor=WHITE,
                font=dict(color=TEXT_PRIMARY, family="IBM Plex Sans, sans-serif"),
                height=300,
                margin=dict(l=40, r=20, t=30, b=40),
                xaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
                yaxis=dict(linecolor=BORDER, linewidth=1, gridcolor="#F0F0F0"),
            )
            st.plotly_chart(fig_box, use_container_width=True, key="sm_box")
