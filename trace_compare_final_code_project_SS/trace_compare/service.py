from pathlib import Path
from .config import CompareConfig
from .diff_engine import compare_records
from .excel_utils import column_letters
from .id_detection import detect_id_columns
from .preprocess import make_records
from .schema import validate_schema
from .workbook_loader import load_pair
from .writer import write_output


def default_output_path(input_path):
    path = Path(input_path)
    return str(path.with_name(f"{path.stem}_comparison_output.xlsx"))


def run_comparison(input_path, output_path=None, id_cols=None, deleted_strike_mode=None):
    config = CompareConfig(
        deleted_strike_mode=deleted_strike_mode or CompareConfig.deleted_strike_mode
    )

    wb, new_ws, old_ws = load_pair(input_path, config.new_suffix, config.old_suffix)
    max_col = validate_schema(new_ws, old_ws, config.header_rows)

    if id_cols is None:
        id_cols = detect_id_columns(new_ws, old_ws, config.header_rows)

    new_records = make_records(new_ws, "new", id_cols, max_col, config.header_rows)
    old_records = make_records(old_ws, "old", id_cols, max_col, config.header_rows)

    output_records = compare_records(
        new_records,
        old_records,
        id_cols=id_cols,
        max_col=max_col,
        deleted_strike_mode=config.deleted_strike_mode,
    )

    write_output(wb, new_ws, old_ws, output_records, max_col, config, id_cols)

    if output_path is None:
        output_path = default_output_path(input_path)
    wb.save(output_path)

    return {
        "output_path": output_path,
        "new_sheet": new_ws.title,
        "old_sheet": old_ws.title,
        "id_columns": column_letters(id_cols),
        "new_records": len(new_records),
        "old_records": len(old_records),
        "output_records": len(output_records),
        "added": sum(1 for r in output_records if r.status == "added"),
        "deleted": sum(1 for r in output_records if r.status == "deleted"),
        "unchanged": sum(1 for r in output_records if r.status == "unchanged"),
    }
