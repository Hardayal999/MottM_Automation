# Traceability Matrix Comparison Tool

Compares two Excel sheets in one workbook:

- sheet ending with `_new`
- sheet ending with `_old`

The output workbook keeps the original sheets and adds `Comparison_Output`.

## What it detects

- Green: new rows / new relationships
- Red + strikethrough: deleted rows / deleted relationships
- Yellow: changed description/detail cells where the ID relationship or hierarchy prefix still exists
- Grey repeated root IDs for easier scanning
- Parent-only/orphan rows are handled so blank parent-only rows do not create noisy false changes

## Install

Run once inside the project folder:

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python run_compare.py "your_workbook.xlsx" --id-columns A,D,G,J
```

Optional output path:

```powershell
python run_compare.py "your_workbook.xlsx" --id-columns A,D,G,J --output "result.xlsx"
```

## Different ID columns

The ID columns are dynamic. For a different project, pass the hierarchy ID columns left-to-right:

```powershell
python run_compare.py "another_workbook.xlsx" --id-columns B,E,H,K,N
```

The program converts these into hierarchy blocks. For example `A,D,G,J` means:

- A-C: level 1 / parent
- D-F: level 2
- G-I: level 3
- J-last column: level 4

## Important notes

- The code uses actual Excel merged-cell ranges only. It does not blindly fill every blank downward.
- Descriptions are compared by hierarchy prefix. This means a parent description change is still detected even if a child relationship was deleted.
- Green/red formatting takes priority over yellow, so newly added/deleted relationship sections remain clear.
