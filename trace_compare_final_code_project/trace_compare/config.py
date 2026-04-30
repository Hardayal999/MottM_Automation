from dataclasses import dataclass


@dataclass(frozen=True)
class CompareConfig:
    new_suffix: str = "_new"
    old_suffix: str = "_old"
    output_sheet_name: str = "Comparison_Output"
    header_rows: int = 1

    # Excel colours
    added_fill: str = "C6EFCE"
    added_font: str = "006100"
    deleted_fill: str = "FFC7CE"
    deleted_font: str = "9C0006"
    repeated_root_font: str = "A6A6A6"
    repeated_root_border: str = "D9D9D9"

    # relationship = if root/parent still exists, strike deleted chain from the 2nd ID level onward
    # missing-level = strike from the exact first hierarchy level that disappears
    deleted_strike_mode: str = "relationship"
