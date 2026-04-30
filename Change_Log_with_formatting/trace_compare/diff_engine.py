from collections import defaultdict
from dataclasses import dataclass

@dataclass(frozen=True)
class OutputRecord:
    status: str
    record: object
    change_start_col: int


def level_ranges(id_cols, max_col):
    ranges = []
    for i, start in enumerate(id_cols):
        end = id_cols[i + 1] - 1 if i + 1 < len(id_cols) else max_col
        ranges.append((start, end))
    return ranges


def cumulative_paths(record):
    paths = []
    current = []
    for value in record.ids:
        current.append(value)
        paths.append(tuple(current))
    return paths


def first_missing_level(record, target_prefixes):
    for idx, path in enumerate(cumulative_paths(record)):
        if path not in target_prefixes:
            return idx
    return len(record.ids) - 1


def full_key(record):
    return record.ids


def build_prefixes(records):
    prefixes = set()
    for record in records:
        prefixes.update(cumulative_paths(record))
    return prefixes


def group_by_root(records):
    grouped = defaultdict(list)
    order = []
    for record in records:
        root = record.ids[0] if record.ids else ""
        if root not in grouped:
            order.append(root)
        grouped[root].append(record)
    return grouped, order


def choose_deleted_start(record, new_prefixes, id_cols, mode):
    if mode == "missing-level":
        idx = first_missing_level(record, new_prefixes)
    elif mode == "relationship":
        root_exists = (record.ids[0],) in new_prefixes if record.ids else False
        idx = 1 if root_exists and len(id_cols) > 1 else 0
    else:
        raise ValueError("deleted_strike_mode must be 'relationship' or 'missing-level'")
    return id_cols[idx]


def compare_records(new_records, old_records, id_cols, max_col, deleted_strike_mode="relationship"):
    old_keys = {full_key(r) for r in old_records}
    new_keys = {full_key(r) for r in new_records}
    old_prefixes = build_prefixes(old_records)
    new_prefixes = build_prefixes(new_records)
    new_by_key = {full_key(r): r for r in new_records}
    old_by_root, old_root_order = group_by_root(old_records)
    new_by_root, new_root_order = group_by_root(new_records)
    roots = []
    for root in new_root_order + old_root_order:
        if root not in roots:
            roots.append(root)
    output = []
    emitted_new = set()
    for root in roots:
        deleted_here = [r for r in old_by_root.get(root, []) if full_key(r) not in new_keys]
        new_here = new_by_root.get(root, [])
        for old_record in deleted_here:
            start_col = choose_deleted_start(old_record, new_prefixes, id_cols, deleted_strike_mode)
            output.append(OutputRecord("deleted", old_record, start_col))
        for new_record in new_here:
            key = full_key(new_record)
            if key in emitted_new:
                continue
            if key in old_keys:
                output.append(OutputRecord("unchanged", new_record, max_col + 1))
            else:
                idx = first_missing_level(new_record, old_prefixes)
                output.append(OutputRecord("added", new_record, id_cols[idx]))
            emitted_new.add(key)
    return output
