from __future__ import annotations

from collections import defaultdict
from copy import copy
from typing import Dict, Iterable, List, Set, Tuple

from .excel_utils import normalise_value
from .row_model import RowRecord


def build_key_map(records: Iterable[RowRecord]) -> Dict[Tuple[str, ...], RowRecord]:
    """Map each hierarchy path to the first matching record."""
    result: Dict[Tuple[str, ...], RowRecord] = {}
    for record in records:
        result.setdefault(record.key, record)
    return result


def build_root_sets(records: Iterable[RowRecord]):
    root_ids: Set[str] = set()
    roots_with_child_links: Set[str] = set()
    for record in records:
        if not record.root_id:
            continue
        root_ids.add(record.root_id)
        if record.has_child_link():
            roots_with_child_links.add(record.root_id)
    return root_ids, roots_with_child_links


def filter_parent_only_noise(
    records: List[RowRecord],
    other_root_ids: Set[str],
    other_roots_with_child_links: Set[str],
) -> List[RowRecord]:
    """
    Suppress parent-only/orphan rows when the same parent exists in the other
    version with meaningful child links.

    Keep parent-only rows when the whole parent is absent in the other version.
    """
    filtered: List[RowRecord] = []
    for record in records:
        if not record.is_parent_only():
            filtered.append(record)
            continue

        root = record.root_id
        if root not in other_root_ids:
            filtered.append(record)
            continue

        if root in other_roots_with_child_links:
            # Not a meaningful relationship change; suppress it.
            continue

        filtered.append(record)
    return filtered


def build_prefix_sets(records: Iterable[RowRecord], id_columns: List[int]) -> List[Set[Tuple[str, ...]]]:
    """
    Prefix sets by hierarchy level.

    For id_columns A,D,G,J and path (A, '', '', J), this includes:
    - level 0: (A,)
    - skips blank D and blank G
    - level 3: (A, '', '', J)

    This allows added highlighting to start at the deepest genuinely new level.
    """
    prefix_sets: List[Set[Tuple[str, ...]]] = [set() for _ in id_columns]
    for record in records:
        path = tuple(normalise_value(v) for v in record.id_path)
        for idx, value in enumerate(path):
            if value == "":
                continue
            prefix_sets[idx].add(path[: idx + 1])
    return prefix_sets


def hierarchy_level_ranges(id_columns: List[int], max_col: int) -> List[Tuple[int, int]]:
    """
    Convert ID columns into dynamic hierarchy blocks.

    Example id_columns A,D,G,J with max_col O:
    - level 0: A-C
    - level 1: D-F
    - level 2: G-I
    - level 3: J-O
    """
    ranges: List[Tuple[int, int]] = []
    for idx, start_col in enumerate(id_columns):
        if idx + 1 < len(id_columns):
            end_col = id_columns[idx + 1] - 1
        else:
            end_col = max_col
        ranges.append((start_col, end_col))
    return ranges


def detail_columns_for_level(id_columns: List[int], max_col: int, level_idx: int) -> List[int]:
    """Return non-ID/detail columns belonging to one hierarchy level."""
    id_set = set(id_columns)
    ranges = hierarchy_level_ranges(id_columns, max_col)
    start_col, end_col = ranges[level_idx]
    return [col for col in range(start_col, end_col + 1) if col not in id_set]


def build_prefix_detail_index(
    records: Iterable[RowRecord],
    id_columns: List[int],
    max_col: int,
) -> List[Dict[Tuple[str, ...], Dict[int, object]]]:
    """
    Build detail/description values for each hierarchy prefix.

    This is the key part for yellow cells. It lets the script detect that a
    parent description changed even when a child relationship was added/deleted.

    Example:
    - Prefix level 0, key (ParentID,) stores B-C values.
    - Prefix level 1, key (ParentID, ObjID) stores E-F values.
    - Prefix level 2, key (ParentID, ObjID, BrsID) stores H-I values.
    - Prefix level 3, key (ParentID, ObjID, BrsID, SrsID) stores K-last values.
    """
    index: List[Dict[Tuple[str, ...], Dict[int, object]]] = [dict() for _ in id_columns]

    for record in records:
        path = tuple(normalise_value(v) for v in record.id_path)
        for level_idx, level_value in enumerate(path):
            if level_value == "":
                continue

            prefix = path[: level_idx + 1]
            detail_cols = detail_columns_for_level(id_columns, max_col, level_idx)
            detail_values = {col: record.values.get(col) for col in detail_cols}

            # Keep the most informative representative for this prefix.
            # Merged descriptions often repeat; if there is a conflict, prefer
            # the version with more nonblank detail cells.
            existing = index[level_idx].get(prefix)
            if existing is None:
                index[level_idx][prefix] = detail_values
            else:
                new_nonblank = sum(1 for v in detail_values.values() if normalise_value(v) != "")
                old_nonblank = sum(1 for v in existing.values() if normalise_value(v) != "")
                if new_nonblank > old_nonblank:
                    index[level_idx][prefix] = detail_values

    return index


def apply_prefix_detail_changes(
    record: RowRecord,
    new_prefix_details: List[Dict[Tuple[str, ...], Dict[int, object]]],
    old_prefix_details: List[Dict[Tuple[str, ...], Dict[int, object]]],
    id_columns: List[int],
    max_col: int,
) -> None:
    """
    Mark changed non-ID/detail cells using hierarchy prefixes, not only full rows.

    This fixes cases like:
    - Old: parent + child relation.
    - New: parent still exists but child relation is deleted, and parent description changed.

    Even though the full child path is deleted, the parent prefix exists in both
    versions, so parent description/detail changes must be shown in yellow.
    """
    changed_cols: List[int] = []
    changed_values: Dict[int, object] = {}
    path = tuple(normalise_value(v) for v in record.id_path)

    for level_idx, level_value in enumerate(path):
        if level_value == "":
            continue

        prefix = path[: level_idx + 1]
        new_details = new_prefix_details[level_idx].get(prefix)
        old_details = old_prefix_details[level_idx].get(prefix)
        if new_details is None or old_details is None:
            continue

        for col in detail_columns_for_level(id_columns, max_col, level_idx):
            new_value = new_details.get(col)
            old_value = old_details.get(col)
            if normalise_value(new_value) != normalise_value(old_value):
                if col not in changed_cols:
                    changed_cols.append(col)
                changed_values[col] = new_value

    record.changed_detail_columns = sorted(changed_cols)
    record.changed_detail_values = changed_values


def added_start_col(record: RowRecord, old_prefix_sets: List[Set[Tuple[str, ...]]], id_columns: List[int]) -> int:
    """
    Decide where green highlighting starts for an added row.

    Rule:
    - If root/top-level ID is new, highlight from root column.
    - If root existed, highlight only from the first genuinely new child level.

    Example id_columns=[A,D,G,J]:
    Old: A exists only.
    New: same A + J child relation.
    Result: green starts at J, not A.
    """
    path = tuple(normalise_value(v) for v in record.id_path)
    if not path:
        return 1

    # Root did not exist in old version: whole root-level record is new.
    if not old_prefix_sets or path[:1] not in old_prefix_sets[0]:
        return id_columns[0]

    # Find first nonblank hierarchy level whose exact prefix did not exist.
    for idx in range(1, len(id_columns)):
        if idx >= len(path):
            break
        if path[idx] == "":
            continue
        if path[: idx + 1] not in old_prefix_sets[idx]:
            return id_columns[idx]

    # Fallback: root existed but the exact full path was still considered added,
    # usually due to row-level details. Start from second level if available.
    return id_columns[1] if len(id_columns) > 1 else id_columns[0]


def deleted_start_col(record: RowRecord, new_prefix_sets: List[Set[Tuple[str, ...]]], id_columns: List[int]) -> int:
    """
    Decide where red/strikethrough starts for a deleted row.

    If the root is gone, strike from root. If the root still exists but an old
    relationship is gone, keep root as context and strike from the second level.
    """
    path = tuple(normalise_value(v) for v in record.id_path)
    if not path:
        return 1
    if not new_prefix_sets or path[:1] not in new_prefix_sets[0]:
        return id_columns[0]
    return id_columns[1] if len(id_columns) > 1 else id_columns[0]


def group_by_root(records: Iterable[RowRecord]) -> Dict[str, List[RowRecord]]:
    grouped = defaultdict(list)
    for record in records:
        grouped[record.root_id].append(record)
    return grouped


def compare_records(new_records_raw: List[RowRecord], old_records_raw: List[RowRecord], id_columns: List[int], max_col: int):
    """
    Main comparison engine.

    Returns:
    - ordered output records with status set
    - old prefix sets for added highlighting
    - new prefix sets for deleted highlighting
    """
    old_root_ids_raw, old_roots_with_child_raw = build_root_sets(old_records_raw)
    new_root_ids_raw, new_roots_with_child_raw = build_root_sets(new_records_raw)

    # Filter parent-only noise for comparison display, but build prefix sets from
    # raw records so parent-only existence still affects highlight start logic.
    new_records = filter_parent_only_noise(new_records_raw, old_root_ids_raw, old_roots_with_child_raw)
    old_records = filter_parent_only_noise(old_records_raw, new_root_ids_raw, new_roots_with_child_raw)

    new_map = build_key_map(new_records)
    old_map = build_key_map(old_records)
    new_keys = set(new_map)
    old_keys = set(old_map)

    old_prefix_sets = build_prefix_sets(old_records_raw, id_columns)
    new_prefix_sets = build_prefix_sets(new_records_raw, id_columns)

    # Prefix detail indexes must be built from raw records so parent-only rows can
    # still tell us whether parent descriptions changed.
    new_prefix_details = build_prefix_detail_index(new_records_raw, id_columns, max_col)
    old_prefix_details = build_prefix_detail_index(old_records_raw, id_columns, max_col)

    deleted = []
    for key in old_keys - new_keys:
        rec = copy(old_map[key])
        rec.status = "deleted"
        apply_prefix_detail_changes(rec, new_prefix_details, old_prefix_details, id_columns, max_col)
        deleted.append(rec)

    deleted_by_root = group_by_root(deleted)
    used_deleted_keys = set()
    output: List[RowRecord] = []

    previous_root = None
    for new_record in new_records:
        root_changed = new_record.root_id != previous_root and previous_root is not None
        if root_changed:
            for d in deleted_by_root.get(previous_root, []):
                if d.key not in used_deleted_keys:
                    output.append(d)
                    used_deleted_keys.add(d.key)

        rec = copy(new_record)
        if rec.key in old_keys:
            rec.status = "unchanged"
        else:
            rec.status = "added"
        apply_prefix_detail_changes(rec, new_prefix_details, old_prefix_details, id_columns, max_col)
        output.append(rec)
        previous_root = new_record.root_id

    if previous_root is not None:
        for d in deleted_by_root.get(previous_root, []):
            if d.key not in used_deleted_keys:
                output.append(d)
                used_deleted_keys.add(d.key)

    # Deleted roots not present in new, or anything not placed above, go at end.
    for d in deleted:
        if d.key not in used_deleted_keys:
            output.append(d)
            used_deleted_keys.add(d.key)

    return output, old_prefix_sets, new_prefix_sets
