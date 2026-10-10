# Stage 1 user workflow

Stage 1 accepts a Manual DATA export as CSV and writes a new workbook from the supplied VJR template. It leaves both inputs unchanged. The project Reference is optional and is never selected automatically.

## Inputs

1. Export the Manual workbook's `DATA` sheet to CSV. Keep its column headers, including `Subject`, `Length`, `Length Unit`, `Count`, `Label`, `DUCT SIZE`, `TYPE`, `ACCESSORIES1`, `ACCESSORIES2`, `REMARK`, and `UNITS`.
2. Select the clean VJR template workbook.
3. Choose a new output path separate from the CSV and template.

## Run

From the repository root:

```powershell
python manual_csv_to_vjr.py "path\to\Manual.csv" `
  --template "path\to\VJR template.xlsx" `
  --output "path\to\generated VJR.xlsx"
```

The clean template defaults to `source_files/Ventilation Jean Roy excel.xlsx`. To compare against or use evidence from a project-specific Reference VJR, supply it explicitly with `--reference "path\to\Reference VJR.xlsx"`. Do not pass another project's Reference.

## Outputs

The run creates:

- the generated VJR workbook;
- a validation report beside it;
- a record audit CSV;
- an unresolved-fields CSV;
- a normalized reference comparison CSV (it contains no comparison records unless `--reference` is supplied).

Every Manual CSV record is counted. Records with missing source fields or unsupported mappings appear in the unresolved CSV with their source row and reason. The writer continues with other supported fields and does not guess.

## Template and capacity

The engine reads the supplied template's sheet names and spiral diameter values. It writes to the established template input areas: TAKE OFF rows 20–1004 and the preformatted SPIRALE diameter rows 11–63. These are template-structure coordinates, not Manual source row numbers or project-specific output row assignments. The writer allocates rows from current input data and aggregates records dynamically; it does not assume a fixed Manual record count.

The generator patches the workbook package from the template and validates formulas and structure. Review the run's validation report before using the generated file.

## Validate and test

Run the complete automated suite:

```powershell
python -m unittest discover -s tests -v
```

`test_input/stage1_second_project_sample.csv` is a small synthetic fixture for checking variable record counts and the generic workflow. It is not a real client project. Replace it with the next project's authorized CSV when that input is available.

For a new project, use the direct `manual_csv_to_vjr.py` command above. The older `manual_to_vjr.py` dispatcher is retained for compatibility with earlier workbook/manifest workflows and contains legacy project-specific routing; it is not the new-project CSV entry point.

## Stage 1 baseline

The frozen baseline bundle and SHA-256 manifest are under `baselines/`. Changes to business rules or mappings require an explicit Stage 1 update; routine project CSV processing does not change the baseline.
