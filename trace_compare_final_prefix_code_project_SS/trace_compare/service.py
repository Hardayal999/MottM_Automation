from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from .config import CompareConfig
from .diff_engine import compare_records
from .excel_utils import col_letters
from .id_detection import detect_id_columns
from .preprocess import sheet_to_records
from .schema import validate_schema
from .workbook_loader import load_workbook_and_sheets
from .writer import write_output_workbook


def default_output_path(input_path: str) -> str:
    p = Path(input_path)
    return str(p.with_name(p.stem + "_comparison_output.xlsx"))


def run_comparison(input_path: str, id_columns: Optional[List[int]] = None, output_path: Optional[str] = None) -> str:
    config = CompareConfig()
    wb, new_ws, old_ws = load_workbook_and_sheets(input_path, config)
    max_col = validate_schema(new_ws, old_ws, config)

    if not id_columns:
        id_columns = detect_id_columns(new_ws, max_col)
        if not id_columns:
            raise ValueError(
                "Could not auto-detect ID columns. Please provide them, e.g. --id-columns A,D,G,J"
            )

    print(f"Using ID columns: {col_letters(id_columns)}")

    new_records_raw = sheet_to_records(new_ws, id_columns, max_col, config)
    old_records_raw = sheet_to_records(old_ws, id_columns, max_col, config)

    records, old_prefix_sets, new_prefix_sets = compare_records(new_records_raw, old_records_raw, id_columns)

    if output_path is None:
        output_path = default_output_path(input_path)

    result = write_output_workbook(
        wb=wb,
        new_ws=new_ws,
        old_ws=old_ws,
        records=records,
        old_prefix_sets=old_prefix_sets,
        new_prefix_sets=new_prefix_sets,
        id_columns=id_columns,
        max_col=max_col,
        output_path=output_path,
        config=config,
    )

    added = sum(1 for r in records if r.status == "added")
    deleted = sum(1 for r in records if r.status == "deleted")
    unchanged = sum(1 for r in records if r.status == "unchanged")
    print("Comparison complete")
    print(f"  New records read: {len(new_records_raw)}")
    print(f"  Old records read: {len(old_records_raw)}")
    print(f"  Output records: {len(records)}")
    print(f"  Added: {added}")
    print(f"  Deleted: {deleted}")
    print(f"  Unchanged: {unchanged}")
    print(f"  Output: {result}")
    return result
