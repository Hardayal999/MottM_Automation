from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class CompareConfig:
    input_path: Path
    output_path: Optional[Path] = None
    new_sheet_suffix: str = "_new"
    old_sheet_suffix: str = "_old"
    output_sheet_name: str = "Comparison_Output"
    header_row: int = 1
    id_columns: Optional[list[int]] = None
    preserve_original_sheets: bool = True
    case_sensitive_headers: bool = False
    added_fill: str = "C6EFCE"
    deleted_fill: str = "FFC7CE"
    deleted_font: str = "9C0006"
    added_font: str = "006100"
    min_pattern_score: float = 0.55
    max_rows_for_detection: int = 2500
    backup_existing_output_sheet: bool = False

    def resolved_output_path(self) -> Path:
        if self.output_path:
            return self.output_path
        return self.input_path.with_name(f"{self.input_path.stem}_comparison.xlsx")
