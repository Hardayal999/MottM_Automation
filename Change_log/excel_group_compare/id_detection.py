import re
from collections import Counter
from .config import CompareConfig
from .excel_utils import normalise_cell

_HEADER_PATTERN = re.compile(r"(^|\b|_)(id|ids|identifier|reference|ref|code|key)(\b|_|$)", re.I)
_ID_VALUE_PATTERN = re.compile(r"^(?=.*\d)(?=.*[A-Za-z])[A-Za-z0-9&_.:/\\ -]{3,}$")
_STRONG_CODE_PATTERN = re.compile(r"[A-Za-z]{2,}[A-Za-z0-9&]*[-_/][A-Za-z0-9&_.-]*\d{2,}")


def _header_score(header) -> float:
    text = normalise_cell(header).lower()
    if not text:
        return 0.0
    if _HEADER_PATTERN.search(text):
        return 0.85
    return 0.0


def _value_pattern_score(values: list) -> float:
    cleaned = [normalise_cell(v) for v in values if normalise_cell(v)]
    if len(cleaned) < 2:
        return 0.0
    pattern_hits = sum(1 for v in cleaned if _ID_VALUE_PATTERN.match(v))
    strong_hits = sum(1 for v in cleaned if _STRONG_CODE_PATTERN.search(v))
    unique_ratio = len(set(cleaned)) / max(len(cleaned), 1)
    repeated_ratio = 1 - unique_ratio
    code_score = (pattern_hits / len(cleaned)) * 0.35 + (strong_hits / len(cleaned)) * 0.35
    uniqueness_score = min(unique_ratio, 0.95) * 0.20
    repeated_group_score = min(repeated_ratio, 0.60) * 0.10
    return code_score + uniqueness_score + repeated_group_score


def _merged_column_score(ws, col: int, header_row: int) -> float:
    hits = 0
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_col <= col <= merged_range.max_col and merged_range.max_row > max(header_row, merged_range.min_row):
            hits += 1
    return min(hits / 3, 1.0) * 0.35


def _structure_score(headers: list, col: int) -> float:
    idx = col - 1
    next_one = normalise_cell(headers[idx + 1]).lower() if idx + 1 < len(headers) else ""
    next_two = normalise_cell(headers[idx + 2]).lower() if idx + 2 < len(headers) else ""
    score = 0.0
    if any(word in next_one for word in ["name", "title", "requirement", "description"]):
        score += 0.10
    if any(word in next_two for word in ["description", "desc", "text", "statement"]):
        score += 0.10
    return score


def detect_id_columns(new_ws, old_ws, headers: list, new_matrix_preview: list[list], old_matrix_preview: list[list], config: CompareConfig) -> list[int]:
    if config.id_columns:
        return sorted(set(config.id_columns))

    max_col = len(headers)
    scores = {}
    start_row = config.header_row + 1
    end_row = min(max(len(new_matrix_preview), len(old_matrix_preview)), config.max_rows_for_detection)

    for col in range(1, max_col + 1):
        values = []
        for matrix in [new_matrix_preview, old_matrix_preview]:
            for row_idx in range(start_row, min(len(matrix), end_row) + 1):
                values.append(matrix[row_idx - 1][col - 1])
        score = 0.0
        score += _header_score(headers[col - 1])
        score += _value_pattern_score(values)
        score += max(_merged_column_score(new_ws, col, config.header_row), _merged_column_score(old_ws, col, config.header_row))
        score += _structure_score(headers, col)
        scores[col] = min(score, 1.0)

    header_columns = [col for col, score in scores.items() if _header_score(headers[col - 1]) >= 0.85]
    if header_columns:
        return header_columns

    detected = [col for col, score in scores.items() if score >= config.min_pattern_score]
    if detected:
        return detected

    raise ValueError(
        "Could not confidently detect ID/group columns. Run again with --id-columns, for example --id-columns A,D,G,J."
    )


def describe_id_columns(headers: list, id_columns: list[int]) -> list[str]:
    return [f"{col}: {headers[col - 1]}" for col in id_columns]
