from dataclasses import dataclass
from difflib import SequenceMatcher
from .row_model import RowRecord


@dataclass(frozen=True)
class DiffRow:
    status: str
    record: RowRecord


@dataclass(frozen=True)
class DiffResult:
    rows: list[DiffRow]
    unchanged_count: int
    added_count: int
    deleted_count: int


def compare_records(new_records: list[RowRecord], old_records: list[RowRecord]) -> DiffResult:
    old_keys = [r.row_key for r in old_records]
    new_keys = [r.row_key for r in new_records]
    matcher = SequenceMatcher(a=old_keys, b=new_keys, autojunk=False)
    output_rows = []
    unchanged = added = deleted = 0

    for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if tag == "equal":
            for record in new_records[new_start:new_end]:
                output_rows.append(DiffRow("unchanged", record))
                unchanged += 1
        elif tag == "delete":
            for record in old_records[old_start:old_end]:
                output_rows.append(DiffRow("deleted", record))
                deleted += 1
        elif tag == "insert":
            for record in new_records[new_start:new_end]:
                output_rows.append(DiffRow("added", record))
                added += 1
        elif tag == "replace":
            old_block = old_records[old_start:old_end]
            new_block = new_records[new_start:new_end]
            output_rows.extend(DiffRow("deleted", record) for record in old_block)
            output_rows.extend(DiffRow("added", record) for record in new_block)
            deleted += len(old_block)
            added += len(new_block)

    return DiffResult(output_rows, unchanged, added, deleted)
