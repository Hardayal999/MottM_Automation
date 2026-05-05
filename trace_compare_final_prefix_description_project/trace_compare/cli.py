from __future__ import annotations

import argparse
import sys

from .excel_utils import parse_column_list
from .service import run_comparison


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compare hierarchical traceability matrix sheets ending with _new and _old."
    )
    parser.add_argument("input_file", help="Excel workbook containing one *_new sheet and one *_old sheet")
    parser.add_argument(
        "--id-columns",
        required=False,
        help="Comma-separated hierarchy ID columns, e.g. A,D,G,J or 1,4,7,10",
    )
    parser.add_argument(
        "--output",
        required=False,
        help="Optional output .xlsx path. Default: input filename + _comparison_output.xlsx",
    )
    return parser


def main(argv=None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        id_columns = parse_column_list(args.id_columns) if args.id_columns else None
        run_comparison(args.input_file, id_columns=id_columns, output_path=args.output)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
