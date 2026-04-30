
def validate_schema(new_ws, old_ws, header_rows=1):
    if new_ws.max_column != old_ws.max_column:
        raise ValueError(f"Column count mismatch: new={new_ws.max_column}, old={old_ws.max_column}")
    max_col = new_ws.max_column
    for row in range(1, header_rows + 1):
        new_headers = [new_ws.cell(row, c).value for c in range(1, max_col + 1)]
        old_headers = [old_ws.cell(row, c).value for c in range(1, max_col + 1)]
        if new_headers != old_headers:
            raise ValueError(f"Header mismatch in row {row}. New={new_headers}. Old={old_headers}")
    return max_col
