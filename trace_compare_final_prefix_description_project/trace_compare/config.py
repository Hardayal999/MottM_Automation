from dataclasses import dataclass


@dataclass(frozen=True)
class CompareConfig:
    new_sheet_suffix: str = "_new"
    old_sheet_suffix: str = "_old"
    output_sheet_name: str = "Comparison_Output"

    # Excel ARGB colours
    added_fill: str = "FFC6EFCE"       # light green
    deleted_fill: str = "FFFFC7CE"     # light red
    changed_fill: str = "FFFFEB9C"     # light yellow for changed detail/description cells
    repeated_root_fill: str = "FFD9E1F2"  # light blue-grey
    light_border: str = "FFBFBFBF"

    header_row: int = 1
    data_start_row: int = 2
