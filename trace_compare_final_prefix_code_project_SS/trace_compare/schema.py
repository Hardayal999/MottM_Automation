from __future__ import annotations

from .config import CompareConfig
from .excel_utils import normalise_value


def validate_schema(new_ws, old_ws, config: CompareConfig) -> int:
    """Validate both sheets have the same number of columns and identical header row."""
    max_col = max(new_ws.max_column, old_ws.max_column)

    new_headers = [normalise_value(new_ws.cell(config.header_row, c).value) for c in range(1, max_col + 1)]
    old_headers = [normalise_value(old_ws.cell(config.header_row, c).value) for c in range(1, max_col + 1)]

    if len(new_headers) != len(old_headers):
        raise ValueError("New and old sheets do not have the same number of columns.")

    mismatches = []
    for idx, (new_h, old_h) in enumerate(zip(new_headers, old_headers), start=1):
        if new_h != old_h:
            mismatches.append((idx, new_h, old_h))

    if mismatches:
        msg_lines = ["New and old sheets have different column headers/order:"]
        for col, new_h, old_h in mismatches[:20]:
            msg_lines.append(f"  Column {col}: new='{new_h}' old='{old_h}'")
        if len(mismatches) > 20:
            msg_lines.append(f"  ... and {len(mismatches) - 20} more mismatch(es)")
        raise ValueError("\n".join(msg_lines))

    return max_col
