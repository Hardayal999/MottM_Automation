from dataclasses import dataclass
from .excel_utils import is_empty_row, normalise_value


@dataclass(frozen=True)
class Record:
    source: str
    source_row: int
    values: tuple
    ids: tuple


def build_merge_lookup(ws):
    """
    Map every cell inside an actual merged range to that range's top-left value.
    This avoids blind fill-down and prevents false relationships.
    """
    lookup = {}
    for rng in ws.merged_cells.ranges:
        value = ws.cell(rng.min_row, rng.min_col).value
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                lookup[(row, col)] = value
    return lookup


def effective_value(ws, merge_lookup, row, col):
    value = ws.cell(row, col).value
    if value is None and (row, col) in merge_lookup:
        return merge_lookup[(row, col)]
    return value


def make_records(ws, source_name, id_cols, max_col, header_rows=1):
    merge_lookup = build_merge_lookup(ws)
    records = []
    seen = set()

    for row in range(header_rows + 1, ws.max_row + 1):
        values = tuple(
            normalise_value(effective_value(ws, merge_lookup, row, col))
            for col in range(1, max_col + 1)
        )
        if is_empty_row(values):
            continue

        ids = tuple(values[col - 1] for col in id_cols)
        if all(x == "" for x in ids):
            continue

        # Remove exact duplicated normalised physical rows.
        signature = (ids, values)
        if signature in seen:
            continue
        seen.add(signature)
        records.append(Record(source_name, row, values, ids))

    return records
