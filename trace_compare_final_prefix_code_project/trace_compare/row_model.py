from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Tuple

from .excel_utils import normalise_value


@dataclass
class RowRecord:
    source_sheet: str
    source_row: int
    values: Dict[int, object]
    id_path: Tuple[str, ...]
    fingerprint: Tuple[str, ...]
    status: str = "unchanged"  # unchanged, added, deleted

    @property
    def key(self) -> Tuple[str, ...]:
        return self.id_path

    @property
    def root_id(self) -> str:
        return self.id_path[0] if self.id_path else ""

    def is_parent_only(self) -> bool:
        """True when row has root/top-level ID but no lower child IDs."""
        if not self.id_path:
            return False
        return normalise_value(self.id_path[0]) != "" and all(
            normalise_value(v) == "" for v in self.id_path[1:]
        )

    def has_child_link(self) -> bool:
        return bool(self.id_path) and any(normalise_value(v) != "" for v in self.id_path[1:])
