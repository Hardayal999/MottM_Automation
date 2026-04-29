from openpyxl import load_workbook
from .config import CompareConfig
from .workbook_loader import find_version_sheets
from .schema import validate_schema
from .preprocess import extract_filled_matrix
from .id_detection import detect_id_columns, describe_id_columns
from .row_model import build_records
from .diff_engine import compare_records
from .writer import write_comparison_sheet


def run_comparison(config: CompareConfig) -> dict:
    workbook = load_workbook(config.input_path)
    new_ws, old_ws = find_version_sheets(workbook, config)
    headers = validate_schema(new_ws, old_ws, config)

    detection_matrix_new = extract_filled_matrix(new_ws, [], config.header_row)
    detection_matrix_old = extract_filled_matrix(old_ws, [], config.header_row)
    id_columns = detect_id_columns(new_ws, old_ws, headers, detection_matrix_new, detection_matrix_old, config)

    new_matrix = extract_filled_matrix(new_ws, id_columns, config.header_row)
    old_matrix = extract_filled_matrix(old_ws, id_columns, config.header_row)
    new_records = build_records(new_matrix, "new", id_columns, config.header_row)
    old_records = build_records(old_matrix, "old", id_columns, config.header_row)
    diff = compare_records(new_records, old_records)

    write_comparison_sheet(workbook, new_ws, old_ws, headers, diff, config)
    output_path = config.resolved_output_path()
    workbook.save(output_path)

    return {
        "output_path": str(output_path),
        "new_sheet": new_ws.title,
        "old_sheet": old_ws.title,
        "id_columns": id_columns,
        "id_columns_described": describe_id_columns(headers, id_columns),
        "new_rows": len(new_records),
        "old_rows": len(old_records),
        "output_rows": len(diff.rows),
        "added_rows": diff.added_count,
        "deleted_rows": diff.deleted_count,
        "unchanged_rows": diff.unchanged_count,
    }
