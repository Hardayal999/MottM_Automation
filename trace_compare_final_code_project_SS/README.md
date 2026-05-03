# Traceability Matrix Hierarchical Compare

This project compares one Excel workbook containing:

- one sheet ending `_new`
- one sheet ending `_old`

It is built for hierarchical traceability matrices where IDs appear in multiple hierarchy columns such as `A,D,G,J`, but the code is **not hard-coded** to those columns.

## Install once

```powershell
pip install -r requirements.txt
```

## Recommended run

```powershell
python run_compare.py "traceability_matrix_export.xlsx" --id-columns A,D,G,J
```

## Output

The output workbook contains:

- `Comparison_Output`
- original `_new` sheet
- original `_old` sheet

## Main rules implemented

1. Uses actual merged-cell ranges only. It does not blindly fill blanks down.
2. Checks both sheets have the same number of columns and matching headers.
3. Compares hierarchy relationships using the ID columns you provide.
4. New/added relationships are highlighted green from the first new hierarchy level onward.
5. Deleted relationships are red with strikethrough.
6. Parent-only/orphan rows are handled:
   - Old parent-only + new same parent with child links: suppress old parent-only row.
   - Old parent with child links + new parent-only: show deleted child links, suppress new parent-only row.
   - Parent-only exists only in one version and parent is absent in the other: show it as added/deleted.
7. Repeated top/root IDs are greyed out in `Comparison_Output` for easier scrolling.
8. Exact duplicate normalised records are removed.

## Different ID structures

For 3 ID columns:

```powershell
python run_compare.py "file.xlsx" --id-columns A,D,G
```

For 10 ID columns:

```powershell
python run_compare.py "file.xlsx" --id-columns A,D,G,J,M,P,S,V,Y,AB
```

The columns must be listed from left to right in hierarchy order.

## Deleted strike mode

Default:

```powershell
python run_compare.py "file.xlsx" --id-columns A,D,G,J --deleted-strike-mode relationship
```

This keeps the top/root parent as context if it still exists and strikes the relationship chain from the second hierarchy level onward.

Alternative:

```powershell
python run_compare.py "file.xlsx" --id-columns A,D,G,J --deleted-strike-mode missing-level
```

This strikes from the exact first hierarchy level that is missing.

## Folder structure

```text
run_compare.py
requirements.txt
README.md
trace_compare/
  __init__.py
  cli.py
  config.py
  diff_engine.py
  excel_utils.py
  id_detection.py
  preprocess.py
  schema.py
  service.py
  workbook_loader.py
  writer.py
```
