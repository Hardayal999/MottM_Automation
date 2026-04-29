from dataclasses import dataclass
from .excel_utils import normalise_cell


@dataclass(frozen=True)
class RowRecord:
    source: str
    source_row: int
    values: tuple[str, ...]
    raw_values: tuple
    id_key: tuple[str, ...]

    @property
    def row_key(self) -> tuple[str, ...]:
        return self.values


def build_records(matrix: list[list], source: str, id_columns: list[int], header_row: int) -> list[RowRecord]:
    records = []
    for row_number, values in enumerate(matrix, start=1):
        if row_number <= header_row:
            continue
        normalised = tuple(normalise_cell(v) for v in values)
        if not any(normalised):
            continue
        id_key = tuple(normalised[col - 1] for col in id_columns)
        records.append(RowRecord(source=source, source_row=row_number, values=normalised, raw_values=tuple(values), id_key=id_key))
    return records
