import argparse
import sys
from .excel_utils import parse_columns
from .service import run_comparison


def build_parser():
    parser = argparse.ArgumentParser(
        description="Compare hierarchical traceability matrix sheets ending _new and _old."
    )
    parser.add_argument("input_file", help="Excel workbook containing one *_new sheet and one *_old sheet")
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help="Optional output .xlsx path. Default: input filename with _comparison_output.xlsx suffix.",
    )
    parser.add_argument(
        "--id-columns",
        default=None,
        help="Comma-separated hierarchy ID columns, e.g. A,D,G,J or 1,4,7,10. Recommended for production.",
    )
    parser.add_argument(
        "--deleted-strike-mode",
        choices=["relationship", "missing-level"],
        default="relationship",
        help="relationship = keep root context and strike deleted chain from level 2 when root exists. missing-level = strike from exact first missing level.",
    )
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        id_cols = parse_columns(args.id_columns)
        result = run_comparison(
            input_path=args.input_file,
            output_path=args.output,
            id_cols=id_cols,
            deleted_strike_mode=args.deleted_strike_mode,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print("Comparison complete")
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0
