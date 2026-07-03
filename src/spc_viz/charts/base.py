"""Shared chart building helpers and SPC analytics.

Common helpers used across chart types (combined profile, box plot,
histogram), plus pure-computation SPC analytics functions used by the UI
layer. Extracted from the original ``chart_utils`` module without
behavioural changes.
"""

from __future__ import annotations

import re
from collections import OrderedDict

import numpy as np
import pandas as pd
from scipy import stats as scipy_stats

from spc_viz.parsers.pairing import (
    SOURCE_META_COLUMNS,
    detect_source_condition,
    detect_source_level,
    is_paired_dim_id,
)

# ---------------------------------------------------------------------------
# Color palettes (no purple)
# ---------------------------------------------------------------------------
COLOR_PALETTE: list[str] = [
    "#2563EB",  # blue
    "#DC2626",  # red
    "#059669",  # green
    "#D97706",  # amber
    "#0891B2",  # cyan
    "#E11D48",  # rose
    "#4F46E5",  # indigo
    "#EA580C",  # orange
    "#0D9488",  # teal
    "#64748B",  # slate
]

MAX_TRACES_PER_GROUP: int = 600


def get_color_for_group(idx: int) -> str:
    """Return a hex color string for a given group index (cycles through palette)."""
    return COLOR_PALETTE[idx % len(COLOR_PALETTE)]


def _natural_sort_key(value: str) -> tuple:
    parts = re.split(r"(\d+(?:\.\d+)?)", str(value).lower())
    key: list[tuple[int, object]] = []
    for part in parts:
        if not part:
            continue
        try:
            key.append((0, float(part)))
        except ValueError:
            key.append((1, part))
    return tuple(key)


def _section_value_sort_key(field_name: str, value: str) -> tuple:
    field = str(field_name or "").strip().lower()
    text = str(value or "").strip()
    upper = text.upper()

    if field in {"level", "source level"} or upper in {"PP", "AP"}:
        level_order = {"PP": 0, "AP": 1}
        if upper in level_order:
            return (0, level_order[upper], upper)

    if field in {"source sheet", "sheet"}:
        if re.search(r"(^|[^A-Z])PP([^A-Z]|$)", upper):
            return (0, 0, upper)
        if re.search(r"(^|[^A-Z])AP([^A-Z]|$)", upper):
            return (0, 1, upper)

    if field in {"source condition", "condition"}:
        if "POR" in upper:
            return (0, 0, upper)
        if "CORR" in upper:
            return (0, 1, upper)

    return (1, _natural_sort_key(text))


def section_sort_key(
    section_by_fields: list[str],
    section_parts: tuple[str, ...],
) -> tuple:
    """Return a stable sort key for section labels.

    Domain-specific order currently puts PP before AP, then POR before CORR.
    Other values fall back to natural sorting.
    """
    return tuple(
        _section_value_sort_key(field, value)
        for field, value in zip(section_by_fields, section_parts)
    )


def has_domain_section_order(
    section_by_fields: list[str],
    section_parts_values,
) -> bool:
    """Whether a section set contains PP/AP or POR/CORR ordering semantics."""
    for parts in section_parts_values:
        for field, value in zip(section_by_fields, parts):
            sort_key = _section_value_sort_key(field, value)
            if sort_key and sort_key[0] == 0:
                return True
    return False


# ---------------------------------------------------------------------------
# Data preparation
# ---------------------------------------------------------------------------


def _get_factory(pf: dict) -> str:
    """Get factory code for a parsed file dict."""
    return pf.get("factory") or "Unknown"


def _find_matching_dim(pf_dims: OrderedDict, target_dno: str) -> str | None:
    """Find a dimension in pf_dims that matches target_dno.

    Handles naming variations like SPC_A vs SPC_A-1 by comparing
    the base name (stripping trailing -N suffixes) and checking if
    column labels overlap.
    """
    if target_dno in pf_dims:
        return target_dno

    # Strip trailing dash-number suffix for fuzzy matching
    # e.g. "SPC_A-1" base is "SPC_A", "SPC_A" base is "SPC_A"
    target_base = re.sub(r"-\d+$", "", target_dno)

    for candidate_dno in pf_dims:
        candidate_base = re.sub(r"-\d+$", "", candidate_dno)
        if candidate_base == target_base:
            return candidate_dno

    return None


def _resolve_dimension_for_file(pf: dict, target_dno: str) -> str | None:
    """Resolve a selected dimension to the local dimension in one parsed file."""
    if is_paired_dim_id(target_dno):
        pair_info = pf.get("_paired_dimensions", {}).get(target_dno)
        if pair_info:
            return pair_info.get("local_dim_no")
        return None
    return _find_matching_dim(pf["dimensions"], target_dno)


def _canonical_meta_for_target(pf: dict, target_dno: str, local_dno: str) -> object:
    if is_paired_dim_id(target_dno):
        pair_info = pf.get("_paired_dimensions", {}).get(target_dno)
        if pair_info and pair_info.get("meta") is not None:
            return pair_info["meta"]
    return pf["dimensions"][local_dno]


def prepare_combined_data(
    parsed_files: list[dict],
    dim_nos: list[str],
) -> tuple[pd.DataFrame | None, OrderedDict | None]:
    """Combine data from all files for the requested dimensions.

    Returns (df, dim_metas_dict) where df has all rows and a _factory column.

    Handles dimension name variations between files (e.g. SPC_A vs SPC_A-1)
    by fuzzy-matching on base dimension name and renaming columns to align.
    """
    frames: list[pd.DataFrame] = []
    dim_metas: OrderedDict = OrderedDict()

    # First pass: collect canonical dim_metas from the first file that has each dim
    for pf in parsed_files:
        for dno in dim_nos:
            if dno not in dim_metas:
                match = _resolve_dimension_for_file(pf, dno)
                if match:
                    dim_metas[dno] = _canonical_meta_for_target(pf, dno, match)

    for pf in parsed_files:
        factory = _get_factory(pf)
        df = pf["data"].copy()
        df["_factory"] = factory
        df["_source_file"] = pf["filename"]
        sheet_name = pf.get("sheet_name") or "Unknown"
        df["Source Sheet"] = sheet_name
        df["Source Level"] = detect_source_level(sheet_name)
        df["Source Condition"] = detect_source_condition(sheet_name)

        meta_cols = [c for c in pf["meta_columns"] if c in df.columns]
        for col in SOURCE_META_COLUMNS:
            if col != "Original Dimension" and col in df.columns and col not in meta_cols:
                meta_cols.append(col)
        meas_cols: list[str] = []
        rename_map: dict[str, str] = {}
        resolved_dim_names: list[str] = []

        for dno in dim_nos:
            match = _resolve_dimension_for_file(pf, dno)
            if match is None:
                continue
            resolved_dim_names.append(match)

            local_meta = pf["dimensions"][match]
            canonical_meta = dim_metas.get(dno)

            if canonical_meta and (match != dno or is_paired_dim_id(dno)):
                # Rename local columns to canonical names so they align
                for local_label, canon_label in zip(
                    local_meta.col_labels, canonical_meta.col_labels
                ):
                    if local_label in df.columns and local_label != canon_label:
                        rename_map[local_label] = canon_label

            # Always add local labels — rename_map will convert them to canonical names later
            meas_cols.extend([c for c in local_meta.col_labels if c in df.columns])

        # Deduplicate while preserving order
        seen: set[str] = set()
        meas_cols_dedup: list[str] = []
        for c in meas_cols:
            if c not in seen:
                seen.add(c)
                meas_cols_dedup.append(c)
        meas_cols = meas_cols_dedup

        keep = meta_cols + meas_cols + ["_factory", "_source_file"]
        if resolved_dim_names:
            df["Original Dimension"] = " / ".join(dict.fromkeys(resolved_dim_names))
            keep.append("Original Dimension")
        keep = [c for c in keep if c in df.columns]
        df = df[keep]

        if rename_map:
            df = df.rename(columns=rename_map)

        frames.append(df)

    if not frames:
        return None, None

    combined = pd.concat(frames, ignore_index=True)
    return combined, dim_metas


# ---------------------------------------------------------------------------
# Section / row logic
# ---------------------------------------------------------------------------


def compute_sections(df: pd.DataFrame, section_by_fields: list[str]) -> pd.Series:
    """Assign a section label to each row based on selected fields.

    Returns a Series of section labels aligned with df index.
    """
    if not section_by_fields:
        return pd.Series("All", index=df.index)

    def _get_col(field_name: str) -> pd.Series:
        if field_name == "Factory":
            if "_factory" in df.columns:
                return df["_factory"].fillna("?").astype(str)
            if "Factory" in df.columns:
                return df["Factory"].fillna("?").astype(str)
            return pd.Series("?", index=df.index)
        elif field_name == "Source File":
            if "_source_file" in df.columns:
                return df["_source_file"].fillna("?").astype(str)
            return pd.Series("?", index=df.index)
        elif field_name in df.columns:
            return df[field_name].fillna("?").astype(str)
        return pd.Series("?", index=df.index)

    parts = [_get_col(f) for f in section_by_fields]
    combined: pd.Series = parts[0]
    for p in parts[1:]:
        combined = combined + " " + p
    return combined


def compute_row_groups(df: pd.DataFrame, row_by: str) -> pd.Series:
    """Assign a row group label to each row based on row_by field.

    Returns a Series of row labels aligned with df index.
    """
    if row_by == "None" or row_by not in df.columns:
        return pd.Series("All", index=df.index)
    return df[row_by].fillna("?").astype(str)


# ---------------------------------------------------------------------------
# SPC analytics (pure computation, no Streamlit dependency)
# ---------------------------------------------------------------------------


def calc_process_capability(
    data_series: pd.Series,
    usl_val: float | None,
    lsl_val: float | None,
) -> dict | None:
    """Calculate Cp, Cpk, Pp, Ppk, sigma level, DPMO, and yield %."""
    data = data_series.dropna()
    if len(data) < 2:
        return None
    mean = data.mean()
    std_within = data.std(ddof=1)
    std_overall = data.std(ddof=0)

    result: dict = {"n": len(data), "mean": round(mean, 6), "std": round(std_within, 6)}

    if usl_val is not None and lsl_val is not None and std_within > 0:
        cp = (usl_val - lsl_val) / (6 * std_within)
        cpu = (usl_val - mean) / (3 * std_within)
        cpl = (mean - lsl_val) / (3 * std_within)
        cpk = min(cpu, cpl)
        pp = (usl_val - lsl_val) / (6 * std_overall) if std_overall > 0 else np.nan
        ppu = (usl_val - mean) / (3 * std_overall) if std_overall > 0 else np.nan
        ppl = (mean - lsl_val) / (3 * std_overall) if std_overall > 0 else np.nan
        ppk = min(ppu, ppl) if std_overall > 0 else np.nan
        result.update(
            {"Cp": round(cp, 4), "Cpk": round(cpk, 4), "Pp": round(pp, 4), "Ppk": round(ppk, 4)}
        )
        sigma_level = cpk * 3
        result["Sigma Level"] = round(sigma_level, 2)
        z_upper = (usl_val - mean) / std_within if std_within > 0 else np.inf
        z_lower = (mean - lsl_val) / std_within if std_within > 0 else np.inf
        p_defect = scipy_stats.norm.sf(z_upper) + scipy_stats.norm.cdf(-z_lower)
        dpmo = p_defect * 1_000_000
        yield_pct = (1 - p_defect) * 100
        result["DPMO"] = int(round(dpmo))
        result["Yield %"] = round(yield_pct, 4)
    elif usl_val is not None and std_within > 0:
        cpu = (usl_val - mean) / (3 * std_within)
        result.update({"Cpk (upper)": round(cpu, 4)})
    elif lsl_val is not None and std_within > 0:
        cpl = (mean - lsl_val) / (3 * std_within)
        result.update({"Cpk (lower)": round(cpl, 4)})

    oos = 0
    if usl_val is not None:
        oos += (data > usl_val).sum()
    if lsl_val is not None:
        oos += (data < lsl_val).sum()
    result["OOS Count"] = int(oos)
    result["OOS %"] = round(oos / len(data) * 100, 2) if len(data) > 0 else 0.0

    return result


def nelson_rules(data_series: pd.Series) -> dict[str, list[int]]:
    """Detect Nelson rule violations for trend & shift detection.

    Returns a dict of rule_name -> list of violating indices.
    """
    data = data_series.dropna().values
    n = len(data)
    if n < 9:
        return {}
    mean = np.mean(data)
    std = np.std(data, ddof=1)
    if std == 0:
        return {}

    violations: dict[str, list[int]] = {}

    r1 = [i for i in range(n) if abs(data[i] - mean) > 3 * std]
    if r1:
        violations["Rule 1: Beyond 3s"] = r1

    r2: list[int] = []
    for i in range(n - 8):
        segment = data[i : i + 9]
        if all(s > mean for s in segment) or all(s < mean for s in segment):
            r2.extend(range(i, i + 9))
    if r2:
        violations["Rule 2: 9 pts same side"] = sorted(set(r2))

    r3: list[int] = []
    for i in range(n - 5):
        seg = data[i : i + 6]
        diffs = np.diff(seg)
        if all(d > 0 for d in diffs) or all(d < 0 for d in diffs):
            r3.extend(range(i, i + 6))
    if r3:
        violations["Rule 3: 6 pts trend"] = sorted(set(r3))

    r4: list[int] = []
    for i in range(n - 13):
        seg = data[i : i + 14]
        diffs = np.diff(seg)
        alternating = all(diffs[j] * diffs[j + 1] < 0 for j in range(len(diffs) - 1))
        if alternating:
            r4.extend(range(i, i + 14))
    if r4:
        violations["Rule 4: 14 pts alternating"] = sorted(set(r4))

    r5: list[int] = []
    for i in range(n - 2):
        seg = data[i : i + 3]
        above = sum(1 for s in seg if s > mean + 2 * std)
        below = sum(1 for s in seg if s < mean - 2 * std)
        if above >= 2 or below >= 2:
            r5.extend(range(i, i + 3))
    if r5:
        violations["Rule 5: 2/3 beyond 2s"] = sorted(set(r5))

    r6: list[int] = []
    for i in range(n - 14):
        seg = data[i : i + 15]
        if all(abs(s - mean) < std for s in seg):
            r6.extend(range(i, i + 15))
    if r6:
        violations["Rule 6: 15 pts within 1s"] = sorted(set(r6))

    return violations


def cusum_analysis(
    data_series: pd.Series,
    target: float | None = None,
    h: float = 5.0,
    k: float = 0.5,
) -> tuple[np.ndarray | None, np.ndarray | None, list[int]]:
    """CUSUM (Cumulative Sum) analysis for shift detection."""
    data = data_series.dropna().values
    n = len(data)
    if n < 5:
        return None, None, []
    mean = target if target is not None else np.mean(data)
    std = np.std(data, ddof=1)
    if std == 0:
        return None, None, []

    cusum_pos = np.zeros(n)
    cusum_neg = np.zeros(n)
    shift_points: list[int] = []

    for i in range(n):
        zi = (data[i] - mean) / std
        cusum_pos[i] = max(0, cusum_pos[i - 1] + zi - k) if i > 0 else max(0, zi - k)
        cusum_neg[i] = max(0, cusum_neg[i - 1] - zi - k) if i > 0 else max(0, -zi - k)
        if cusum_pos[i] > h or cusum_neg[i] > h:
            shift_points.append(i)

    return cusum_pos, cusum_neg, shift_points
