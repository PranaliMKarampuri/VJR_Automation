# Manual to VJR

This application provides a small upload-and-download interface over the validated Stage 1 Manual CSV → VJR backend. Mapping logic stays in `manual_csv_to_vjr.py`; `app.py` contains only UI code.

## Install

Use Python 3.10 or newer. From the repository root, install the runtime dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Fixed template

The application automatically uses:

```text
source_files/Ventilation Jean Roy excel.xlsx
```

The user does not upload a VJR template. Keep the fixed template at that path. The application checks that it exists and contains the required worksheets before enabling generation.

## Start

```powershell
python -m streamlit run app.py
```

Open the local address shown by Streamlit, upload a Manual file, then select **Generate VJR**. Each completed job is kept in its own folder under `output/jobs/` so its download links remain available during the session.

## Input format

Supported formats:

- `.csv`, exported from the Manual's DATA sheet;
- `.xlsx`, with a `DATA` worksheet.

The input must use the Stage 1 Manual DATA fields, including `Subject`, `Length`, `Length Unit`, `Count`, `Label`, `DUCT SIZE`, `TYPE`, `ACCESSORIES1`, `ACCESSORIES2`, `REMARK`, and `UNITS`. Legacy `.xls` is not supported; save it as `.xlsx` or export DATA to CSV first.

The application never selects a Reference VJR automatically. Stage 1 rules and mappings come from the existing backend. Missing-source-data, unsupported, and ambiguous records appear in the unresolved CSV; the app does not guess.

## Output files

Each job folder contains:

- generated VJR Excel workbook;
- validation report;
- record audit CSV;
- unresolved records CSV;
- normalized Reference comparison CSV (empty of comparison records unless a Reference is explicitly provided through the backend API).

Both the uploaded Manual and fixed template are hash-checked by the backend. Output is created from a separate working copy; the fixed template is never modified. Workbook structure/formula preservation checks must pass before the job becomes downloadable.

## Tests

Run all tests with:

```powershell
python -m unittest discover -s tests -v
```

The Stage 1 baseline and the synthetic second-project fixture are described in `reports/stage1_baseline_readiness.md` and `docs/stage1_user_workflow.md`.
