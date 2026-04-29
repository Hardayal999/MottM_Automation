from .config import CompareConfig
from .excel_utils import normalise_header


def get_headers(ws, header_row: int) -> list:
    return [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]


def validate_schema(new_ws, old_ws, config: CompareConfig) -> list[str]:
    new_headers = get_headers(new_ws, config.header_row)
    old_headers = get_headers(old_ws, config.header_row)
    if len(new_headers) != len(old_headers):
        raise ValueError(f"Column count mismatch: new sheet has {len(new_headers)} columns, old sheet has {len(old_headers)} columns.")
    mismatches = []
    for idx, (new_h, old_h) in enumerate(zip(new_headers, old_headers), start=1):
        if normalise_header(new_h, config.case_sensitive_headers) != normalise_header(old_h, config.case_sensitive_headers):
            mismatches.append(f"Column {idx}: new='{new_h}' old='{old_h}'")
    if mismatches:
        message = "Column names/order do not match.\n" + "\n".join(mismatches[:50])
        if len(mismatches) > 50:
            message += f"\n... plus {len(mismatches) - 50} more mismatches"
        raise ValueError(message)
    return new_headers
