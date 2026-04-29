from copy import copy
from .excel_utils import resolve_simple_reference


def build_merged_value_map(ws) -> dict[tuple[int, int], object]:
    value_map = {}
    for merged_range in ws.merged_cells.ranges:
        top_left = ws.cell(merged_range.min_row, merged_range.min_col)
        value = resolve_simple_reference(ws, top_left.value)
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                value_map[(row, col)] = value
    return value_map


def extract_filled_matrix(ws, id_columns: list[int], header_row: int) -> list[list[object]]:
    merged_values = build_merged_value_map(ws)
    matrix = []
    last_id_values = {col: None for col in id_columns}
    for row in range(1, ws.max_row + 1):
        row_values = []
        for col in range(1, ws.max_column + 1):
            raw_value = merged_values.get((row, col), ws.cell(row, col).value)
            value = resolve_simple_reference(ws, raw_value)
            if row > header_row and col in id_columns:
                if value not in (None, ""):
                    last_id_values[col] = value
                else:
                    value = last_id_values[col]
            row_values.append(value)
        matrix.append(row_values)
    return matrix


def first_source_row_for_output(ws, source_row: int) -> int:
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= source_row <= merged_range.max_row:
            return merged_range.min_row
    return source_row
