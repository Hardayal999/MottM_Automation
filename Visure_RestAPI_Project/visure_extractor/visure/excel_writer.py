"""
excel_writer.py
===============
Takes the raw element list from the Visure API and writes a structured
Excel workbook.

The shape problem
-----------------
The API returns a list of elements (Element14). Each element looks like:

    {
      "id": 12345,
      "code": "REQ-001",
      "name": "System shall...",
      "description": "Long HTML-y description",
      "descriptionTxt": "Long plain-text description",
      "depth": 0,
      "isRequirement": "true",
      "author": "...",
      "creationDate": "...",
      "attributes": [
         {"name": "Status",   "valueStr": "Approved"},
         {"name": "Priority", "valueStr": "High"},
         ...
      ]
    }

Different specs have different sets of user-defined attributes. Power BI
wants a FLAT TABLE — one row per element, one column per attribute, with
empty cells where a given element doesn't have that attribute.

So we:
  1. Walk every element and collect the union of all attribute names seen.
  2. Build a wide DataFrame: standard columns + one column per attribute.
  3. Write three sheets: Requirements, Summary, Run Log.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
import pandas as pd


# Columns we always want from the Element14 object, in display order.
# These appear LEFT of the dynamic attribute columns.
STANDARD_COLUMNS = [
    "project_name",
    "project_id",
    "spec_name",
    "spec_id",
    "element_id",
    "code",
    "name",
    "chapter",
    "depth",
    "is_requirement",
    "description_txt",
    "author",
    "creation_date",
    "last_modification_date",
    "version_number",
]


@dataclass
class ExtractedSpec:
    """One specification's worth of extracted data.

    The extractor builds a list of these and hands them to write_workbook().
    """
    project_id: int
    project_name: str
    spec_id: int
    spec_name: str
    elements: list[dict]           # raw Element14 dicts from the API
    error: str | None = None       # if extraction failed, set this and skip


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _attribute_value(attr: dict) -> str | None:
    """Pick the most useful value off a UserAttribute14 dict.

    The API returns attributes with several possible value fields:
      - valueStr   (text)
      - valueInt   (integer)
      - name / code (for enumerated/dropdown attributes)

    For Excel/Power BI display we want a single string per cell.
    """
    # valueStr is filled for free-text attributes and most enumerations
    if attr.get("valueStr"):
        return attr["valueStr"]
    # `name` is the display value of a dropdown selection
    if attr.get("name"):
        return attr["name"]
    if attr.get("valueInt") is not None and attr.get("valueInt") != 0:
        return str(attr["valueInt"])
    return None


def _attribute_name(attr: dict) -> str:
    """The attribute label. `attribute` is Visure's preferred field; `code` is fallback."""
    return attr.get("attribute") or attr.get("code") or "Unknown attribute"


def _flatten_element(
    element: dict,
    *,
    project_id: int,
    project_name: str,
    spec_id: int,
    spec_name: str,
) -> dict:
    """Turn one Element14 dict into one flat row.

    Standard fields are pulled out by hand. Then every attribute on the
    element becomes its own column, keyed by attribute name.
    """
    row: dict = {
        "project_name": project_name,
        "project_id": project_id,
        "spec_name": spec_name,
        "spec_id": spec_id,
        "element_id": element.get("id"),
        "code": element.get("code"),
        "name": element.get("name"),
        "chapter": element.get("chapter"),
        "depth": element.get("depth"),
        "is_requirement": element.get("isRequirement"),
        # descriptionTxt is the plain-text version. `description` may contain
        # HTML markup which is noisy in Excel — we prefer the txt version.
        "description_txt": element.get("descriptionTxt") or element.get("description"),
        "author": element.get("author"),
        "creation_date": element.get("creationDate"),
        "last_modification_date": element.get("lastModificationDate"),
        "version_number": element.get("versionNumber"),
    }

    for attr in element.get("attributes") or []:
        col = _attribute_name(attr)
        # If the same attribute name appears twice (shouldn't, but be safe),
        # the second value wins. We log nothing because it's harmless.
        row[col] = _attribute_value(attr)

    return row


def _build_requirements_dataframe(extracted: Iterable[ExtractedSpec]) -> pd.DataFrame:
    """Flatten every element from every spec into one wide DataFrame."""
    rows: list[dict] = []
    for spec in extracted:
        if spec.error:
            continue
        for element in spec.elements:
            rows.append(_flatten_element(
                element,
                project_id=spec.project_id,
                project_name=spec.project_name,
                spec_id=spec.spec_id,
                spec_name=spec.spec_name,
            ))

    if not rows:
        return pd.DataFrame(columns=STANDARD_COLUMNS)

    df = pd.DataFrame(rows)

    # Column ordering: standard columns first (in our preferred order),
    # then attribute columns alphabetically. This gives Power BI a stable,
    # predictable schema across runs.
    standard_present = [c for c in STANDARD_COLUMNS if c in df.columns]
    attribute_cols = sorted(c for c in df.columns if c not in STANDARD_COLUMNS)
    return df[standard_present + attribute_cols]


def _build_summary_dataframe(
    extracted: Iterable[ExtractedSpec],
    run_started_at: datetime,
) -> pd.DataFrame:
    """One row per spec — for a quick "did it work" overview."""
    rows = []
    for s in extracted:
        rows.append({
            "project_name": s.project_name,
            "project_id": s.project_id,
            "spec_name": s.spec_name,
            "spec_id": s.spec_id,
            "element_count": 0 if s.error else len(s.elements),
            "requirement_count": (
                0 if s.error else sum(
                    1 for e in s.elements
                    if str(e.get("isRequirement", "")).lower() == "true"
                )
            ),
            "status": "ERROR" if s.error else "OK",
            "error_message": s.error or "",
        })
    df = pd.DataFrame(rows)
    # Add a single metadata row showing when the extraction ran
    df.attrs["run_started_at"] = run_started_at.isoformat(timespec="seconds")
    return df


def _build_runlog_dataframe(
    extracted: Iterable[ExtractedSpec],
    run_started_at: datetime,
    run_finished_at: datetime,
) -> pd.DataFrame:
    """A timestamped log of the extraction itself."""
    total_specs = sum(1 for _ in extracted)
    extracted = list(extracted)  # we consumed the iterable above

    failed = [s for s in extracted if s.error]
    succeeded = [s for s in extracted if not s.error]
    total_elements = sum(len(s.elements) for s in succeeded)

    rows = [
        {"key": "run_started_at",  "value": run_started_at.isoformat(timespec="seconds")},
        {"key": "run_finished_at", "value": run_finished_at.isoformat(timespec="seconds")},
        {"key": "duration_seconds","value": f"{(run_finished_at - run_started_at).total_seconds():.1f}"},
        {"key": "specs_attempted", "value": len(extracted)},
        {"key": "specs_succeeded", "value": len(succeeded)},
        {"key": "specs_failed",    "value": len(failed)},
        {"key": "elements_total",  "value": total_elements},
    ]
    for f in failed:
        rows.append({"key": f"failed_spec_{f.spec_id}", "value": f"{f.spec_name}: {f.error}"})

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def write_workbook(
    extracted: list[ExtractedSpec],
    output_dir: Path,
    run_started_at: datetime,
    filename: str | None = None,
) -> Path:
    """Write the 3-sheet workbook and return the path it was written to.

    Sheet 1 'Requirements' — flat element rows
    Sheet 2 'Summary'      — one row per spec
    Sheet 3 'Run Log'      — extraction metadata
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    run_finished_at = datetime.now()
    if not filename:
        # Timestamped filename so successive runs don't overwrite each other.
        # Format: visure_extract_YYYYMMDD_HHMMSS.xlsx
        filename = f"visure_extract_{run_finished_at:%Y%m%d_%H%M%S}.xlsx"

    output_path = output_dir / filename

    requirements_df = _build_requirements_dataframe(extracted)
    summary_df = _build_summary_dataframe(extracted, run_started_at)
    runlog_df = _build_runlog_dataframe(extracted, run_started_at, run_finished_at)

    # ExcelWriter with openpyxl gives us multi-sheet output in one file.
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        requirements_df.to_excel(writer, sheet_name="Requirements", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        runlog_df.to_excel(writer, sheet_name="Run Log", index=False)

    return output_path
