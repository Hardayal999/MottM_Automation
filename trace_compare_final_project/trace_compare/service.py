from pathlib import Path
from openpyxl import load_workbook
from .config import CompareConfig
from .workbook_loader import load_pair
from .schema import validate_schema
from .id_detection import detect_id_columns
from .preprocess import make_records
from .diff_engine import compare_records
from .writer import write_output


def default_output_path(input_path):
    path = Path(input_path)
    return str(path.with_name(path.stem + "_comparison_output.xlsx"))


def run_comparison(input_path, output_path=None, id_columns=None, deleted_strike_mode="relationship"):
    config = CompareConfig(deleted_strike_mode=deleted_strike_mode)
    wb, new_ws, old_ws = load_pair(input_path, config.new_suffix, config.old_suffix)
    max_col = validate_schema(new_ws, old_ws, config.header_rows)
    ids = id_columns or detect_id_columns(new_ws, old_ws, config.header_rows)
    if ids[0] < 1 or ids[-1] > max_col:
        raise ValueError(f"ID columns {ids} are outside the sheet width 1..{max_col}")
    new_records = make_records(new_ws, "new", ids, max_col, config.header_rows)
    old_records = make_records(old_ws, "old", ids, max_col, config.header_rows)
    output_records = compare_records(new_records, old_records, ids, max_col, config.deleted_strike_mode)
    write_output(wb, new_ws, old_ws, output_records, max_col, config)
    out = output_path or default_output_path(input_path)
    wb.save(out)
    return {
        "output_path": out,
        "id_columns": ids,
        "new_records": len(new_records),
        "old_records": len(old_records),
        "output_records": len(output_records),
        "added": sum(1 for r in output_records if r.status == "added"),
        "deleted": sum(1 for r in output_records if r.status == "deleted"),
        "unchanged": sum(1 for r in output_records if r.status == "unchanged"),
    }
