"""
excel_writer.py
===============
Takes the raw element list + links from the Visure API and writes a
structured Excel workbook.

Output sheets
-------------
Sheet 1 'Requirements'   — one row per element (flat table with attributes)
Sheet 2 'Traceability'   — one row per link per anchor spec
Sheet 3 'Summary'        — one row per spec
Sheet 4 'Run Log'        — extraction metadata

Critical API rule
-----------------
When Visure returns links for an element we queried, link.get("code") is
always the OTHER side's code — the side that isn't our anchored element.
Not the target's code. Not the source's code. The OTHER side's.

So:
  - For outgoing links (our element is source): link.code = target's code
  - For incoming links (our element is target): link.code = source's code

After we flip incoming links so our element sits in source_element_id, the
other element becomes our new target. In both directions, link.get("code")
is the right fallback for target_code.

Orientation
-----------
Each link is anchored to the spec it was fetched from. Whichever element
belongs to that spec sits in source_element_id. The other endpoint goes in
target_element_id. Direction column records outgoing / incoming.

Deduplication
-------------
Key = (source, target, link_type, anchor_spec_id). When extracting multiple
specs, the same relationship appears once per anchor — both perspectives
are kept on purpose.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable
import pandas as pd


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

TRACEABILITY_COLUMNS = [
    "project_name",
    "project_id",
    "source_spec_name",
    "source_spec_id",
    "source_element_id",
    "source_code",
    "source_name",
    "target_element_id",
    "target_code",
    "target_name",
    "target_spec_name",
    "target_spec_id",
    "target_project",
    "link_type",
    "direction",
    "is_suspect",
    "suspect_reason",
]


@dataclass
class ExtractedSpec:
    """One specification's worth of extracted data."""
    project_id: int
    project_name: str
    spec_id: int
    spec_name: str
    elements: list[dict]
    links: list[dict] = field(default_factory=list)
    error: str | None = None


# ---------------------------------------------------------------------------
# Element flattening
# ---------------------------------------------------------------------------

def _attribute_value(attr: dict) -> str | None:
    values = attr.get("values") or []
    clean = [str(v) for v in values if v not in (None, "")]
    if not clean:
        return None
    if len(clean) == 1:
        return clean[0]
    return " | ".join(clean)


def _attribute_name(attr: dict) -> str:
    return attr.get("name") or attr.get("attribute") or attr.get("code") or "Unknown attribute"


def _flatten_element(
    element: dict,
    *,
    project_id: int,
    project_name: str,
    spec_id: int,
    spec_name: str,
) -> dict:
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
        "description_txt": element.get("descriptionTxt") or element.get("description"),
        "author": element.get("author"),
        "creation_date": element.get("creationDate"),
        "last_modification_date": element.get("lastModificationDate"),
        "version_number": element.get("versionNumber"),
    }
    for attr in element.get("attributes") or []:
        row[_attribute_name(attr)] = _attribute_value(attr)
    return row


# ---------------------------------------------------------------------------
# Traceability flattening
# ---------------------------------------------------------------------------

def _build_element_lookup(extracted: list[ExtractedSpec]) -> dict[int, dict]:
    """Build {element_id -> {code, name, spec_name, spec_id}} across all
    extracted specs. Used to enrich the other side of each link when that
    element happens to be in our extract."""
    lookup: dict[int, dict] = {}
    for spec in extracted:
        if spec.error:
            continue
        for el in spec.elements:
            el_id = el.get("id")
            if el_id is None:
                continue
            lookup[el_id] = {
                "code": el.get("code"),
                "name": el.get("name"),
                "spec_id": spec.spec_id,
                "spec_name": spec.spec_name,
            }
    return lookup


def _flatten_link(
    link: dict,
    source_spec: ExtractedSpec,
    source_spec_element_ids: set[int],
    element_lookup: dict[int, dict],
) -> dict:
    """Turn one LinkedItem14 dict into one flat traceability row.

    Orientation anchored to source_spec — our element sits in source_element_id
    regardless of direction. Direction column tells you which way the
    relationship actually flows.

    Critical: link.get("code") is ALWAYS the OTHER side's code (the side
    that isn't our anchored element). This holds for both outgoing and
    incoming. We use it as the fallback for target_code in both branches.
    """
    raw_source_id = link.get("sourceItemID")
    raw_target_id = link.get("targetItemID")

    if raw_source_id in source_spec_element_ids:
        # OUTGOING — our element is the source, target is the other side
        our_id = raw_source_id
        other_id = raw_target_id
        direction_label = "outgoing"
    elif raw_target_id in source_spec_element_ids:
        # INCOMING — flip so our element is the source. Other side becomes
        # our new target.
        our_id = raw_target_id
        other_id = raw_source_id
        direction_label = "incoming"
    else:
        # Neither side anchored — shouldn't happen, fall back
        our_id = raw_source_id
        other_id = raw_target_id
        direction_label = "unknown"

    # Look up both elements. link.code/link.name is always the OTHER side's,
    # which IS our target after the flip — use it as fallback.
    our_info = element_lookup.get(our_id, {}) if our_id else {}
    other_info = element_lookup.get(other_id, {}) if other_id else {}

    return {
        "project_name": source_spec.project_name,
        "project_id": source_spec.project_id,
        "source_spec_name": our_info.get("spec_name") or source_spec.spec_name,
        "source_spec_id": our_info.get("spec_id") or source_spec.spec_id,
        "source_element_id": our_id,
        "source_code": our_info.get("code"),
        "source_name": our_info.get("name"),
        "target_element_id": other_id,
        # link.code is the OTHER side's code — used as fallback in both
        # outgoing and incoming. element_lookup wins when richer info available.
        "target_code": other_info.get("code") or link.get("code"),
        "target_name": other_info.get("name") or link.get("name"),
        "target_spec_name": other_info.get("spec_name"),
        "target_spec_id": other_info.get("spec_id"),
        "target_project": link.get("project"),
        "link_type": link.get("linkType"),
        "direction": direction_label,
        "is_suspect": bool(link.get("isSuspect")),
        "suspect_reason": link.get("reason"),
    }


# ---------------------------------------------------------------------------
# DataFrame builders
# ---------------------------------------------------------------------------

def _build_requirements_dataframe(extracted: Iterable[ExtractedSpec]) -> pd.DataFrame:
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
    standard_present = [c for c in STANDARD_COLUMNS if c in df.columns]
    attribute_cols = sorted(c for c in df.columns if c not in STANDARD_COLUMNS)
    return df[standard_present + attribute_cols]


def _build_traceability_dataframe(extracted: list[ExtractedSpec]) -> pd.DataFrame:
    """Build the traceability table anchored per spec."""
    element_lookup = _build_element_lookup(extracted)

    rows: list[dict] = []
    seen_keys: set[tuple[int, int, str, int]] = set()

    for spec in extracted:
        if spec.error:
            continue

        spec_element_ids: set[int] = {
            el["id"] for el in spec.elements if el.get("id") is not None
        }

        for link in spec.links:
            src = link.get("sourceItemID")
            tgt = link.get("targetItemID")
            ltype = link.get("linkType") or ""
            if src is None or tgt is None:
                continue

            key = (src, tgt, ltype, spec.spec_id)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            rows.append(_flatten_link(link, spec, spec_element_ids, element_lookup))

    if not rows:
        return pd.DataFrame(columns=TRACEABILITY_COLUMNS)

    df = pd.DataFrame(rows)
    cols_present = [c for c in TRACEABILITY_COLUMNS if c in df.columns]
    extras = [c for c in df.columns if c not in TRACEABILITY_COLUMNS]
    return df[cols_present + extras]


def _build_summary_dataframe(
    extracted: Iterable[ExtractedSpec],
    run_started_at: datetime,
) -> pd.DataFrame:
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
            "link_count_raw": 0 if s.error else len(s.links),
            "status": "ERROR" if s.error else "OK",
            "error_message": s.error or "",
        })
    df = pd.DataFrame(rows)
    df.attrs["run_started_at"] = run_started_at.isoformat(timespec="seconds")
    return df


def _build_runlog_dataframe(
    extracted: Iterable[ExtractedSpec],
    run_started_at: datetime,
    run_finished_at: datetime,
) -> pd.DataFrame:
    extracted = list(extracted)
    failed = [s for s in extracted if s.error]
    succeeded = [s for s in extracted if not s.error]
    total_elements = sum(len(s.elements) for s in succeeded)
    total_links_raw = sum(len(s.links) for s in succeeded)

    rows = [
        {"key": "run_started_at",   "value": run_started_at.isoformat(timespec="seconds")},
        {"key": "run_finished_at",  "value": run_finished_at.isoformat(timespec="seconds")},
        {"key": "duration_seconds", "value": f"{(run_finished_at - run_started_at).total_seconds():.1f}"},
        {"key": "specs_attempted",  "value": len(extracted)},
        {"key": "specs_succeeded",  "value": len(succeeded)},
        {"key": "specs_failed",     "value": len(failed)},
        {"key": "elements_total",   "value": total_elements},
        {"key": "links_raw_total",  "value": total_links_raw},
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
    """Write the 4-sheet workbook and return the path it was written to."""
    output_dir.mkdir(parents=True, exist_ok=True)

    run_finished_at = datetime.now()
    if not filename:
        filename = f"visure_extract_{run_finished_at:%Y%m%d_%H%M%S}.xlsx"

    output_path = output_dir / filename

    requirements_df = _build_requirements_dataframe(extracted)
    traceability_df = _build_traceability_dataframe(extracted)
    summary_df      = _build_summary_dataframe(extracted, run_started_at)
    runlog_df       = _build_runlog_dataframe(extracted, run_started_at, run_finished_at)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        requirements_df.to_excel(writer, sheet_name="Requirements",  index=False)
        traceability_df.to_excel(writer, sheet_name="Traceability",  index=False)
        summary_df.to_excel(     writer, sheet_name="Summary",       index=False)
        runlog_df.to_excel(      writer, sheet_name="Run Log",       index=False)

    return output_path
