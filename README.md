# Excel Group Comparison Tool

This tool compares one Excel workbook containing one sheet ending with `_new` and one sheet ending with `_old`.

It is designed for grouped traceability matrices where ID cells may be merged vertically. The program automatically reads merged ID/group cells, fills the blank cells downward in memory, and then creates one output sheet called `Comparison_Output`.

## What it does

- Checks that both sheets have the same number of columns.
- Checks that column headers match in the same order.
- Automatically detects ID/group columns.
- Supports manual ID column override when a different project has unusual headers.
- Does not compare ID columns character-by-character.
- Keeps deleted old rows in the comparison sheet with red fill and strikethrough.
- Keeps new added rows in the comparison sheet with green fill.
- Keeps original `_new` and `_old` sheets in the workbook.
- Removes older helper sheets such as `Formatted_Comparison`, `Comparison_Summary`, `Change_Log`, and `Deleted_IDs` if they exist.

## Install

```bash
pip install -r requirements.txt
```

## Run with automatic ID detection

```bash
python run_compare.py "traceability_matrix_export.xlsx"
```

## Run with manual ID columns

Use this when a different project has different ID naming or unusual column layout.

```bash
python run_compare.py "traceability_matrix_export.xlsx" --id-columns A,D,G,J
```

You can also use numbers:

```bash
python run_compare.py "traceability_matrix_export.xlsx" --id-columns 1,4,7,10
```

## Custom output path

```bash
python run_compare.py "traceability_matrix_export.xlsx" --output "compared_output.xlsx"
```

## How ID detection works

The automatic detector does not depend on one fixed prefix such as `Project-BRS`. It scores each column using:

1. Header names containing words such as `ID`, `Identifier`, `Reference`, `Ref`, `Code`, or `Key`.
2. Merged vertical cell structure, because group ID columns are often merged in traceability matrices.
3. Code-like values containing both letters and numbers, including values such as `Project-BRS-0001`, `REQ_012`, `SYS/ABC/010`, or similar.
4. Nearby column structure such as `ID → Name/Title → Description`.

For your current file, it detects columns `A, D, G, J`. For another project, it should detect different ID columns automatically if the headers or values clearly indicate ID/group columns. If it cannot detect them confidently, it stops and asks you to provide `--id-columns` so it does not produce a misleading comparison.
