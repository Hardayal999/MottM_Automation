from __future__ import annotations

from typing import List

from .excel_utils import normalise_value


def detect_id_columns(ws, max_col: int) -> List[int]:
    """
    Best-effort ID column detection.

    Manual --id-columns is strongly recommended for production use. This exists
    as a fallback only.
    """
    candidates = []
    keywords = ("id", "identifier", "ref", "reference", "code", "key")

    for col in range(1, max_col + 1):
        header = normalise_value(ws.cell(1, col).value).lower()
        score = 0
        if any(k in header for k in keywords):
            score += 5
        merged_hits = 0
        for mr in ws.merged_cells.ranges:
            min_col, min_row, max_col_r, max_row = mr.bounds
            if min_col <= col <= max_col_r and max_row > min_row:
                merged_hits += 1
        score += min(merged_hits, 5)

        sample_code_like = 0
        for row in range(2, min(ws.max_row, 100) + 1):
            value = normalise_value(ws.cell(row, col).value)
            if any(ch.isalpha() for ch in value) and any(ch.isdigit() for ch in value):
                sample_code_like += 1
        if sample_code_like >= 3:
            score += 3

        if score >= 5:
            candidates.append((col, score))

    candidates.sort(key=lambda x: x[0])
    return [col for col, _score in candidates]
