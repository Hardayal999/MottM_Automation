from __future__ import annotations

from typing import Dict, Iterable, List, Tuple

from .config import CompareConfig
from .excel_utils import normalise_value, row_is_blank
from .row_model import RowRecord


def build_merged_lookup(ws, columns: Iterable[int]) -> Dict[Tuple[int, int], Tuple[int, int]]:
    """
    Return lookup for cells inside actual merged ranges only.
    Key: (row, col), value: top-left cell coordinate.

    This avoids blind fill-down. Only cells that are genuinely merged inherit
    the merged range's top-left value.
    """
    wanted_cols = set(columns)
    lookup: Dict[Tuple[int, int], Tuple[int, int]] = {}

    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        affected_cols = [c for c in wanted_cols if min_col <= c <= max_col]
        if not affected_cols:
            continue
        for row in range(min_row, max_row + 1):
            for col in affected_cols:
                lookup[(row, col)] = (min_row, min_col)
    return lookup


def resolved_cell_value(ws, row: int, col: int, merged_lookup: Dict[Tuple[int, int], Tuple[int, int]]):
    if (row, col) in merged_lookup:
        top_row, top_col = merged_lookup[(row, col)]
        return ws.cell(top_row, top_col).value
    return ws.cell(row, col).value


def sheet_to_records(ws, id_columns: List[int], max_col: int, config: CompareConfig) -> List[RowRecord]:
    merged_lookup = build_merged_lookup(ws, id_columns)
    records: List[RowRecord] = []

    for row in range(config.data_start_row, ws.max_row + 1):
        values = {}
        for col in range(1, max_col + 1):
            if col in id_columns:
                values[col] = resolved_cell_value(ws, row, col, merged_lookup)
            else:
                values[col] = ws.cell(row, col).value

        if row_is_blank(values, max_col):
            continue

        id_path = tuple(normalise_value(values.get(col)) for col in id_columns)
        fingerprint = tuple(normalise_value(values.get(col)) for col in range(1, max_col + 1))

        # Ignore rows that have no root/top-level ID. They cannot be safely placed
        # in the hierarchy comparison.
        if not id_path or normalise_value(id_path[0]) == "":
            continue

        records.append(
            RowRecord(
                source_sheet=ws.title,
                source_row=row,
                values=values,
                id_path=id_path,
                fingerprint=fingerprint,
            )
        )
    return dedupe_records(records)


def dedupe_records(records: List[RowRecord]) -> List[RowRecord]:
    """Collapse exact duplicate normalised records while preserving order."""
    seen = set()
    result: List[RowRecord] = []
    for record in records:
        key = (record.id_path, record.fingerprint)
        if key in seen:
            continue
        seen.add(key)
        result.append(record)
    return result
