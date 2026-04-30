# Traceability Matrix Hierarchical Compare

This compares one Excel workbook containing a sheet ending `_new` and another sheet ending `_old`.

It is built for traceability matrices where IDs appear in hierarchy columns such as A, D, G and J, but the ID columns are not hard-coded. You can provide any ID columns on the command line.

## Install once

```powershell
pip install -r requirements.txt
```

## Recommended run

```powershell
python run_compare.py "traceability_matrix_export.xlsx" --id-columns A,D,G,J
```

## Auto-detection run

```powershell
python run_compare.py "traceability_matrix_export.xlsx"
```

If auto-detection is unsure, use `--id-columns`.

## Output rules

- Uses actual merged-cell ranges only.
- Does not blindly fill every blank cell down.
- Removes exact duplicate normalised records.
- Added relationships are highlighted green from the first newly added hierarchy level onward.
- Deleted relationships are red with strikethrough.
- Default deleted rule is dynamic: if the parent/root ID still exists, it keeps the root level as context and strikes the relationship chain from the second ID level onward. If the root no longer exists, it strikes from the first ID level.

## Different ID structure examples

```powershell
python run_compare.py "file.xlsx" --id-columns B,E,H,K
python run_compare.py "file.xlsx" --id-columns 2,5,8,11
```

## Optional deleted strike mode

Default:

```powershell
python run_compare.py "file.xlsx" --id-columns A,D,G,J --deleted-strike-mode relationship
```

Alternative, stricter minimal highlighting:

```powershell
python run_compare.py "file.xlsx" --id-columns A,D,G,J --deleted-strike-mode missing-level
```


## Repeated top-level ID formatting

The output sheet now makes repeated root/top-level IDs easier to scan. The first occurrence of a root ID stays normal, while consecutive repeated root IDs are shown in a softer grey font with lighter borders. This is dynamic and uses the first column listed in `--id-columns`.

Example:

```powershell
python run_compare.py "traceability_matrix_export.xlsx" --id-columns A,D,G,J
```

Here, column A is treated as the root/top-level ID column. If another project uses `B,E,H,K`, then column B becomes the root/top-level ID column automatically.
