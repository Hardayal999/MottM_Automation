import re

ID_WORDS = ("id", "identifier", "requirement", "reference", "ref", "code", "key")
CODE_RE = re.compile(r"^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9_&/ .:-]+$")


def _merged_score(ws, col):
    score = 0
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.max_row > rng.min_row:
            score += rng.max_row - rng.min_row + 1
    return score


def _header_score(ws, col, header_rows):
    text = " ".join(str(ws.cell(r, col).value or "") for r in range(1, header_rows + 1)).lower()
    return 6 if any(word in text for word in ID_WORDS) else 0


def _value_score(ws, col, start_row, sample_rows=300):
    total = 0
    code_like = 0
    repeated_or_blank = 0
    last = object()
    for row in range(start_row, min(ws.max_row, start_row + sample_rows - 1) + 1):
        value = ws.cell(row, col).value
        text = "" if value is None else str(value).strip()
        if text:
            total += 1
            if CODE_RE.match(text):
                code_like += 1
            if text == last:
                repeated_or_blank += 1
            last = text
        else:
            repeated_or_blank += 1
    if total == 0:
        return 0
    return int((code_like / total) * 8) + min(4, repeated_or_blank // 20)


def detect_id_columns(new_ws, old_ws, header_rows=1):
    start_row = header_rows + 1
    scored = []
    for col in range(1, new_ws.max_column + 1):
        score = (
            _header_score(new_ws, col, header_rows)
            + _value_score(new_ws, col, start_row)
            + _value_score(old_ws, col, start_row)
            + min(10, _merged_score(new_ws, col) + _merged_score(old_ws, col))
        )
        if score >= 9:
            scored.append((col, score))
    cols = [col for col, _ in scored]
    filtered = []
    for col in cols:
        if not filtered or col - filtered[-1] >= 2:
            filtered.append(col)
    if len(filtered) < 2:
        raise ValueError("Could not confidently detect ID columns. Re-run with --id-columns A,D,G,J or your actual ID columns.")
    return filtered
