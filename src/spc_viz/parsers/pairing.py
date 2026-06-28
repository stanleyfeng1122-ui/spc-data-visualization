"""Cross-sheet dimension pairing helpers.

Some vendor workbooks use different SPC bubble IDs for the same physical
feature at different levels, for example PP uses ``SPC_AR`` while AP uses
``SPC_AT`` for the same front-edge point sequence. These helpers build a
virtual dimension layer keyed by physical feature + measurement points.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from copy import copy

from .dimensions import DimensionMeta

PAIR_DIM_PREFIX = "PAIR::"
SOURCE_META_COLUMNS = ["Source Sheet", "Source Level", "Source Condition", "Original Dimension"]


def is_paired_dim_id(dim_no: str) -> bool:
    return str(dim_no).startswith(PAIR_DIM_PREFIX)


def simple_feature_name(description: str) -> str:
    """Normalize a feature description into one stable, user-facing name."""
    text = re.sub(r"\s+", " ", str(description or "")).strip().lower()
    text = text.replace(" edge straightness", " edge straightness")
    if not text:
        return "Paired measurement"
    return text[0].upper() + text[1:]


def _slug(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or "paired-measurement"


def _point_signature(dmeta: DimensionMeta) -> tuple[str, ...]:
    return tuple(str(point).strip() for point in dmeta.point_numbers)


def _detect_source_level(sheet_name: str) -> str:
    text = f" {sheet_name.upper()} "
    if re.search(r"(^|[^A-Z])PP([^A-Z]|$)", text):
        return "PP"
    if re.search(r"(^|[^A-Z])AP([^A-Z]|$)", text):
        return "AP"
    return "Unknown"


def _detect_source_condition(sheet_name: str) -> str:
    text = sheet_name.upper()
    if "CORR" in text:
        return "CORR"
    if "POR" in text:
        return "POR"
    return "Unknown"


def _merged_spec(values: list, mode: str) -> list:
    """Merge point-wise spec values across paired dimensions."""
    if not values:
        return []
    merged: list = []
    for per_point_values in zip(*values):
        numeric = [v for v in per_point_values if v is not None]
        if not numeric:
            merged.append(None)
        elif mode == "max":
            merged.append(max(numeric))
        elif mode == "min":
            merged.append(min(numeric))
        else:
            merged.append(numeric[0])
    return merged


def _build_pair_meta(pair_id: str, feature_name: str, members: list[dict]) -> DimensionMeta:
    first_meta: DimensionMeta = members[0]["meta"]
    point_numbers = list(first_meta.point_numbers)
    col_labels = [f"{pair_id}__{point}" for point in point_numbers]
    return DimensionMeta(
        dim_no=pair_id,
        description=feature_name,
        dim_type=first_meta.dim_type,
        point_numbers=point_numbers,
        nominal=_merged_spec([m["meta"].nominal for m in members], "first"),
        tol_max=_merged_spec([m["meta"].tol_max for m in members], "max"),
        tol_min=_merged_spec([m["meta"].tol_min for m in members], "min"),
        usl=_merged_spec([m["meta"].usl for m in members], "max"),
        lsl=_merged_spec([m["meta"].lsl for m in members], "min"),
        col_indices=list(range(1, len(point_numbers) + 1)),
        col_labels=col_labels,
        source_dim_nos=sorted({m["dim_no"] for m in members}),
    )


def _add_source_metadata(parsed_file: dict) -> None:
    data = parsed_file.get("data")
    if data is None:
        return

    sheet_name = parsed_file.get("sheet_name") or ""
    source_level = _detect_source_level(sheet_name)
    source_condition = _detect_source_condition(sheet_name)

    data["Source Sheet"] = sheet_name or "Unknown"
    data["Source Level"] = source_level
    data["Source Condition"] = source_condition

    meta_columns = list(parsed_file.get("meta_columns") or [])
    for col in SOURCE_META_COLUMNS:
        if col not in meta_columns:
            meta_columns.append(col)
    parsed_file["meta_columns"] = meta_columns


def build_paired_dimension_map(parsed_files: list[dict]) -> OrderedDict[str, DimensionMeta]:
    """Return the user-facing dimension map with stackable features paired.

    Mutates each parsed-file dict by adding:
    - source metadata columns to ``data`` / ``meta_columns``
    - ``_paired_dimensions`` mapping virtual pair IDs to local SPC dimensions
    """
    groups: OrderedDict[tuple[str, tuple[str, ...]], list[dict]] = OrderedDict()

    for pf_idx, pf in enumerate(parsed_files):
        _add_source_metadata(pf)
        pf["_paired_dimensions"] = {}
        for dno, dmeta in pf.get("dimensions", {}).items():
            signature = (simple_feature_name(dmeta.description), _point_signature(dmeta))
            groups.setdefault(signature, []).append(
                {"pf_idx": pf_idx, "dim_no": dno, "meta": dmeta, "sheet_name": pf.get("sheet_name")}
            )

    paired_dimensions: OrderedDict[str, DimensionMeta] = OrderedDict()
    covered_dim_nos: set[str] = set()

    for (feature_name, points), members in groups.items():
        distinct_dims = {m["dim_no"] for m in members}
        distinct_sheets = {m["sheet_name"] for m in members}
        if len(members) < 2 or (len(distinct_dims) < 2 and len(distinct_sheets) < 2):
            continue

        point_span = f"{points[0]}-{points[-1]}" if points else "points"
        pair_id = f"{PAIR_DIM_PREFIX}{_slug(feature_name)}::{_slug(point_span)}"
        pair_meta = _build_pair_meta(pair_id, feature_name, members)
        paired_dimensions[pair_id] = pair_meta

        for member in members:
            pf = parsed_files[member["pf_idx"]]
            pf.setdefault("_paired_dimensions", {})[pair_id] = {
                "local_dim_no": member["dim_no"],
                "meta": copy(pair_meta),
            }
            covered_dim_nos.add(member["dim_no"])

    dimensions: OrderedDict[str, DimensionMeta] = OrderedDict()
    dimensions.update(paired_dimensions)

    for pf in parsed_files:
        for dno, dmeta in pf.get("dimensions", {}).items():
            if dno in covered_dim_nos:
                continue
            if dno not in dimensions:
                dimensions[dno] = dmeta

    return dimensions
