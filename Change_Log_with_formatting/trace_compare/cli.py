import argparse
from .excel_utils import parse_columns
from .service import run_comparison


def build_parser():
    parser = argparse.ArgumentParser(description="Fast hierarchical traceability Excel comparison.")
    parser.add_argument("input_file", help="Workbook containing one *_new sheet and one *_old sheet")
    parser.add_argument("--output", "-o", help="Output .xlsx path")
    parser.add_argument("--id-columns", help="Comma-separated ID columns, e.g. A,D,G,J or 1,4,7,10")
    parser.add_argument(
        "--deleted-strike-mode",
        choices=["relationship", "missing-level"],
        default="relationship",
        help="relationship = strike deleted relationship chain from second ID level if parent exists; missing-level = strike only from first missing hierarchy level",
    )
    return parser


def main():
    args = build_parser().parse_args()
    result = run_comparison(
        input_path=args.input_file,
        output_path=args.output,
        id_columns=parse_columns(args.id_columns),
        deleted_strike_mode=args.deleted_strike_mode,
    )
    print("Comparison complete")
    print(f"Output: {result['output_path']}")
    print(f"ID columns: {result['id_columns']}")
    print(f"New records: {result['new_records']}")
    print(f"Old records: {result['old_records']}")
    print(f"Output records: {result['output_records']}")
    print(f"Added: {result['added']}")
    print(f"Deleted: {result['deleted']}")
    print(f"Unchanged: {result['unchanged']}")
