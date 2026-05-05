from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from .excel_utils import normalise_value


@dataclass
class RowRecord:
    source_sheet: str
    source_row: int
    values: Dict[int, object]
    id_path: Tuple[str, ...]
    fingerprint: Tuple[str, ...]
    status: str = "unchanged"  # unchanged, added, deleted

    # Non-ID/detail columns that changed while the relevant hierarchy prefix still exists.
    # Example: parent ID still exists but parent description changed while child relation was deleted.
    changed_detail_columns: List[int] = field(default_factory=list)

    # Values to display for changed detail cells. Normally this is the newer value.
    changed_detail_values: Dict[int, object] = field(default_factory=dict)

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
