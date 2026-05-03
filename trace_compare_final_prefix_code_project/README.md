# Traceability Matrix Comparator

Compares two sheets in one Excel workbook where the newer sheet ends with `_new` and the older sheet ends with `_old`.

It is designed for hierarchical traceability matrices where ID/group columns appear left-to-right, for example:

```text
A, D, G, J
```

The tool:

- validates that both sheets have matching headers
- uses actual merged-cell ranges only, not blind fill-down
- handles grouped/merged hierarchy IDs internally
- compares relationships by hierarchy path
- highlights added relationships in green from the first genuinely new hierarchy level
- highlights deleted relationships in red with strikethrough
- suppresses parent-only/orphan noise rows when the parent exists with child links in the other version
- greys repeated top-level/root IDs for readability
- creates `Comparison_Output` plus keeps the original `_new` and `_old` sheets

## Install

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python run_compare.py "your_file.xlsx" --id-columns A,D,G,J
```

For a different hierarchy:

```powershell
python run_compare.py "another_file.xlsx" --id-columns B,E,H,K,N
```

Optional output path:

```powershell
python run_compare.py "your_file.xlsx" --id-columns A,D,G,J --output "comparison_result.xlsx"
```
