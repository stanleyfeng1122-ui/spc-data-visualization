"""SpcDataset — the one deep interface over everything parsed from Excel.

Callers used to receive a list of raw dicts from parsing and pass it around
(28 touch points across three packages, plus in-place mutation during
pairing). This module owns that lifecycle instead: parse the requested
sheets, apply PP/AP pairing and source metadata, and expose only what the
app layer actually needs. The parsed-file records stay internal.

``load_dataset`` is the constructor; ``SpcDataset.parsed_files`` remains as
a transitional accessor while callers migrate (C1-3 … C1-5), and should not
gain new users.
"""

from __future__ import annotations

import re
from collections import Counter, OrderedDict
from dataclasses import dataclass, field

import pandas as pd

from .dimensions import DimensionMeta
from .excel_reader import FileOrPath, parse_excel_multi
from .pairing import (
    SOURCE_META_COLUMNS,
    build_paired_dimension_map,
    detect_source_condition,
    detect_source_level,
    is_paired_dim_id,
)

# ---------------------------------------------------------------------------
# Display labels (UI-independent; the picker re-exports these)
# ---------------------------------------------------------------------------


def _dimension_display_label(dno: str, dmeta: DimensionMeta) -> str:
    if is_paired_dim_id(dno):
        ids = " / ".join(dmeta.source_dim_nos or [])
        if ids and dmeta.description:
            return f"{ids} — {dmeta.description}"
        return ids or dmeta.description or dno
    return f"{dno} — {dmeta.description}" if dmeta.description else dno


def _bubble_ids(dno: str, dmeta: DimensionMeta) -> list[str]:
    """Source SPC bubble id(s) for a dimension: the paired sources, else dno."""
    if is_paired_dim_id(dno):
        return list(dmeta.source_dim_nos or [])
    return [dno]


def build_display_labels(
    all_dimensions: OrderedDict[str, DimensionMeta],
) -> OrderedDict[str, str]:
    """Map ``dno -> unique display label``.

    Two genuinely different dimensions can share a description (e.g. SPC_CY
    and SPC_CW both "Top ply corrugate bottom to datum A offset"). A plain
    description label would collide and hide one in the picker, so when 2+
    dimensions share a label we append ``(bubble id · point span)`` — falling
    back to the span alone for paired dims that span more than one bubble id.
    Labels that don't collide stay clean.
    """
    base = {dno: _dimension_display_label(dno, dm) for dno, dm in all_dimensions.items()}
    collisions = {lbl for lbl, n in Counter(base.values()).items() if n > 1}

    out: OrderedDict[str, str] = OrderedDict()
    used: set[str] = set()
    for dno, dm in all_dimensions.items():
        label = base[dno]
        if label in collisions:
            pts = [str(p).strip() for p in dm.point_numbers]
            span = f"{pts[0]}–{pts[-1]}" if pts else ""
            ids = _bubble_ids(dno, dm)
            if len(ids) == 1 and span:
                label = f"{label} ({ids[0]} · {span})"
            elif len(ids) == 1:
                label = f"{label} ({ids[0]})"
            elif span:
                label = f"{label} ({span})"
        # Guarantee uniqueness even if a suffix still collides.
        unique, n = label, 2
        while unique in used:
            unique = f"{label} #{n}"
            n += 1
        used.add(unique)
        out[dno] = unique
    return out


# ---------------------------------------------------------------------------
# Combining engine (moved from charts.base — it reads the parsed records, so
# it lives behind the parsers seam; charts re-exports it for back-compat)
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
# The dataset
# ---------------------------------------------------------------------------


@dataclass
class SpcDataset:
    """Everything the app needs from the parsed workbooks, behind one seam."""

    dimensions: OrderedDict[str, DimensionMeta]
    _parsed_files: list[dict] = field(repr=False)

    @property
    def parsed_files(self) -> list[dict]:
        """Transitional: raw parsed-file records for not-yet-migrated callers."""
        return self._parsed_files

    @property
    def meta_columns(self) -> list[str]:
        """Union of metadata column names across all parsed sheets, ordered."""
        seen: OrderedDict[str, None] = OrderedDict()
        for pf in self._parsed_files:
            for col in pf.get("meta_columns") or []:
                seen.setdefault(col, None)
        return list(seen)

    def display_labels(self) -> OrderedDict[str, str]:
        """Unique picker label per dimension (bubble id + description)."""
        return build_display_labels(self.dimensions)

    def combined(
        self, dim_nos: list[str]
    ) -> tuple[pd.DataFrame | None, OrderedDict | None]:
        """Combined measurement frame for the selected dimensions."""
        return prepare_combined_data(self._parsed_files, dim_nos)

    def file_summaries(self) -> list[dict]:
        """Per-source display summaries for the pages' "Loaded Files" panel."""
        summaries: list[dict] = []
        for pf in self._parsed_files:
            data = pf.get("data")
            cfg_values: list[str] = []
            if data is not None and "CFG" in data.columns:
                cfg_values = sorted(data["CFG"].dropna().unique().astype(str)[:5])
            summaries.append(
                {
                    "filename": pf.get("filename") or "?",
                    "sheet_name": pf.get("sheet_name") or "",
                    "part_number": pf.get("part_number"),
                    "factory": pf.get("factory") or "?",
                    "n_rows": len(data) if data is not None else 0,
                    "n_dims": len(pf.get("dimensions") or {}),
                    "cfg_values": cfg_values,
                }
            )
        return summaries


def parse_sheets(source: FileOrPath, sheet_names: tuple[str, ...]) -> list[dict]:
    """Parse the requested sheets of ONE source into raw parsed-file records.

    This is the cache-friendly stage: callers may wrap it in ``st.cache_data``
    per file so unchanged files skip reparsing. The returned records are
    opaque tokens meant only to be fed into :func:`assemble_dataset` — do not
    index into them. Sheets that fail to parse are skipped (historical app
    behavior).
    """
    parsed_files: list[dict] = []
    for sheet_name in sheet_names:
        # In-memory buffers are consumed by openpyxl; rewind between sheets.
        if hasattr(source, "seek"):
            source.seek(0)
        try:
            results = parse_excel_multi(source, sheet_name=sheet_name)
        except Exception:
            continue
        for p in results:
            parsed_files.append(
                {
                    "filename": p.filename,
                    "sheet_name": p.sheet_name,
                    "part_number": p.part_number,
                    "part_description": p.part_description,
                    "revision": p.revision,
                    "factory": p.factory,
                    "dimensions": p.dimensions,
                    "data": p.data,
                    "meta_columns": p.meta_columns,
                }
            )
    return parsed_files


def assemble_dataset(parsed_files: list[dict]) -> SpcDataset:
    """Build one SpcDataset from the records of all sources together.

    Pairing must see every source at once (same-bubble across files, PP/AP
    across sheets), which is why this stage is separate from the per-source
    :func:`parse_sheets`.
    """
    dimensions = build_paired_dimension_map(parsed_files)
    return SpcDataset(dimensions=dimensions, _parsed_files=parsed_files)


def load_dataset(sources: list[tuple[FileOrPath, tuple[str, ...]]]) -> SpcDataset:
    """Parse the requested sheets of each source and build one SpcDataset.

    ``sources`` is a list of ``(file_or_path, sheet_names)`` pairs — a path or
    in-memory buffer plus the sheet names to parse from it.
    """
    parsed_files: list[dict] = []
    for source, sheet_names in sources:
        parsed_files.extend(parse_sheets(source, sheet_names))
    return assemble_dataset(parsed_files)
