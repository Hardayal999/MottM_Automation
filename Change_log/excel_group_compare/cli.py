import argparse
from pathlib import Path
from .config import CompareConfig
from .excel_utils import parse_column_list
from .service import run_comparison


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Compare grouped Excel traceability matrix sheets ending with _new and _old.")
    parser.add_argument("input_file", help="Workbook containing one *_new sheet and one *_old sheet")
    parser.add_argument("--output", "-o", help="Output workbook path")
    parser.add_argument("--id-columns", help="Optional manual ID columns, e.g. A,D,G,J or 1,4,7,10")
    parser.add_argument("--header-row", type=int, default=1, help="Header row number. Default: 1")
    parser.add_argument("--new-suffix", default="_new", help="New sheet suffix. Default: _new")
    parser.add_argument("--old-suffix", default="_old", help="Old sheet suffix. Default: _old")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    config = CompareConfig(
        input_path=Path(args.input_file),
        output_path=Path(args.output) if args.output else None,
        id_columns=parse_column_list(args.id_columns),
        header_row=args.header_row,
        new_sheet_suffix=args.new_suffix,
        old_sheet_suffix=args.old_suffix,
    )
    result = run_comparison(config)
    print("Comparison completed")
    print(f"Output: {result['output_path']}")
    print(f"New sheet: {result['new_sheet']}")
    print(f"Old sheet: {result['old_sheet']}")
    print(f"Detected ID columns: {', '.join(result['id_columns_described'])}")
    print(f"Rows: new={result['new_rows']}, old={result['old_rows']}, output={result['output_rows']}")
    print(f"Added={result['added_rows']}, deleted={result['deleted_rows']}, unchanged={result['unchanged_rows']}")


if __name__ == "__main__":
    main()
