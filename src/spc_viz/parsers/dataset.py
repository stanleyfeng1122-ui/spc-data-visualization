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

from collections import Counter, OrderedDict
from dataclasses import dataclass, field

import pandas as pd

from .dimensions import DimensionMeta
from .excel_reader import FileOrPath, parse_excel_multi
from .pairing import build_paired_dimension_map, is_paired_dim_id

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
        """Combined measurement frame for the selected dimensions.

        Same contract as ``charts.base.prepare_combined_data`` (imported
        lazily to keep parsers import-light and avoid a cycle).
        """
        from spc_viz.charts.base import prepare_combined_data

        return prepare_combined_data(self._parsed_files, dim_nos)


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
