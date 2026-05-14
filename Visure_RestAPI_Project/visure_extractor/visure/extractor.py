"""
extractor.py
============
Orchestrator. Drives user prompts, calls the client, hands data to the writer.
No HTTP code here, no Excel code here — both are delegated.
"""

from __future__ import annotations

from datetime import datetime
from typing import Sequence

from visure.client import (
    VisureClient,
    VisureAPIError,
    Project,
    Specification,
)
from visure.config import settings
from visure.excel_writer import ExtractedSpec, write_workbook


# ---------------------------------------------------------------------------
# CLI helpers
# ---------------------------------------------------------------------------

def _print_table(title: str, items: Sequence) -> None:
    print(f"\n{title}")
    print("-" * len(title))
    for i, item in enumerate(items, start=1):
        print(f"  [{i:>2}] {item.name}  (id={item.id})")


def _prompt_choice(prompt: str, max_choice: int) -> int:
    while True:
        raw = input(f"{prompt} (1-{max_choice}): ").strip()
        if not raw:
            continue
        if not raw.isdigit():
            print("  Please enter a number.")
            continue
        n = int(raw)
        if not 1 <= n <= max_choice:
            print(f"  Please enter a number between 1 and {max_choice}.")
            continue
        return n


def _prompt_spec_choice(specs: Sequence[Specification]) -> list[Specification]:
    print("\nDo you want to extract:")
    print("  [1] All specifications in this project")
    print("  [2] A single specification")
    choice = _prompt_choice("Choice", 2)
    if choice == 1:
        return list(specs)

    _print_table("Specifications in selected project:", specs)
    idx = _prompt_choice("Pick a specification", len(specs))
    return [specs[idx - 1]]


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def _extract_one_spec(
    client: VisureClient,
    project: Project,
    spec: Specification,
) -> ExtractedSpec:
    """Fetch elements AND their traceability links for one spec.

    Strategy:
      1. Get all elements (with attributes) in one call.
      2. Collect their element IDs.
      3. Send those IDs to the bulk linkeditems endpoint to get every link.

    Any failure is wrapped in ExtractedSpec.error so one bad spec doesn't
    abort the whole run.
    """
    try:
        elements = client.get_specification_items(spec.id, include_all_attributes=True)

        # Pull element IDs and ask the bulk link endpoint for everything at once.
        element_ids = [el["id"] for el in elements if el.get("id") is not None]
        links = client.get_linked_items_bulk(element_ids) if element_ids else []

        return ExtractedSpec(
            project_id=project.id,
            project_name=project.name,
            spec_id=spec.id,
            spec_name=spec.name,
            elements=elements,
            links=links,
        )
    except VisureAPIError as e:
        return ExtractedSpec(
            project_id=project.id,
            project_name=project.name,
            spec_id=spec.id,
            spec_name=spec.name,
            elements=[],
            links=[],
            error=str(e),
        )


def run_interactive_extraction() -> None:
    run_started_at = datetime.now()
    print("=" * 60)
    print("Visure API → Excel extractor")
    print(f"Base URL: {settings.base_url}")
    print(f"User:     {settings.username}")
    print("=" * 60)

    with VisureClient() as client:
        # ---- 1. Authenticate ------------------------------------------------
        print("\n[1/4] Authenticating...")
        client.authenticate()
        print("      OK")

        # ---- 2. Pick a project ---------------------------------------------
        projects = client.list_projects_from_auth()
        if not projects:
            print("\nERROR: This account has no projects assigned.")
            return

        _print_table("[2/4] Available projects:", projects)
        idx = _prompt_choice("Pick a project", len(projects))
        chosen_project = projects[idx - 1]
        print(f"      Selected: {chosen_project.name}")

        # ---- 3. Switch session, list specs ---------------------------------
        print(f"\n[3/4] Loading specifications for {chosen_project.name!r}...")
        client.set_current_project(chosen_project)
        specs = client.list_specifications()
        if not specs:
            print("      This project has no specifications.")
            return
        print(f"      Found {len(specs)} specification(s).")

        chosen_specs = _prompt_spec_choice(specs)
        print(f"      Will extract {len(chosen_specs)} specification(s).")

        # ---- 4. Pull elements + traceability links -------------------------
        print(f"\n[4/4] Extracting elements, attributes, and links...")
        results: list[ExtractedSpec] = []
        for i, spec in enumerate(chosen_specs, start=1):
            print(f"      ({i}/{len(chosen_specs)}) {spec.name} ...", end="", flush=True)
            result = _extract_one_spec(client, chosen_project, spec)
            if result.error:
                print(f" FAILED ({result.error[:60]}...)")
            else:
                print(f" {len(result.elements)} elements, {len(result.links)} links")
            results.append(result)

    # ---- 5. Write Excel -----------------------------------------------------
    print("\nWriting Excel workbook...")
    output_path = write_workbook(
        extracted=results,
        output_dir=settings.output_dir,
        run_started_at=run_started_at,
    )
    print(f"      Wrote: {output_path}")

    total_elements = sum(len(r.elements) for r in results if not r.error)
    total_links = sum(len(r.links) for r in results if not r.error)
    failures = sum(1 for r in results if r.error)
    print(
        f"\nDone. {len(results) - failures}/{len(results)} specs OK, "
        f"{total_elements} elements, {total_links} raw links."
    )
