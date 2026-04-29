from dataclasses import dataclass

@dataclass(frozen=True)
class CompareConfig:
    new_suffix: str = "_new"
    old_suffix: str = "_old"
    output_sheet_name: str = "Comparison_Output"
    added_fill: str = "C6EFCE"
    added_font: str = "006100"
    deleted_fill: str = "FFC7CE"
    deleted_font: str = "9C0006"
    deleted_strike_mode: str = "relationship"
    header_rows: int = 1
