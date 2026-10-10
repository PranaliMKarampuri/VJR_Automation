# Stage 1 web frontend implementation report

## Implementation

- Added `app.py`, a Streamlit upload-and-download interface for Manual CSV and Excel files.
- Added `stage1_app_backend.py`, a thin adapter that converts an Excel `DATA` sheet to the CSV schema and invokes the existing `manual_csv_to_vjr.generate` engine. No mapping rules were copied or changed.
- The app always uses `source_files/Ventilation Jean Roy excel.xlsx`; users do not upload a template. Each run is staged separately and then stored in a unique folder under `output/jobs/`.
- The app provides the generated VJR, validation report, record audit, and unresolved-record CSV, plus processed/write/unresolved counts.
- Input and template SHA-256 hashes are checked before and after processing. A job is exposed only if the existing report confirms zero calculation-formula changes and preservation checks pass.
- `.xls` is rejected with an instruction to save as `.xlsx` or export the `DATA` sheet as `.csv`.
- Added `requirements.txt` and updated `README.md` with installation, launch, input, template, output, and test instructions.

## Verification

- Full suite: `python -m unittest discover -s tests -v` — **69 tests passed** in 588.8 seconds.
- Focused frontend/backend suite: **4 tests passed**. It covers the injected UI flow, a different-size three-record CSV, an equivalent `.xlsx` `DATA` sheet, invalid input, fixed-template use, generated artifacts, and preservation checks.
- Synthetic second-project CSV: 3 records processed, 3 with writes, 21 cells written; output created from the fixed template.
- Existing KATIVIK CSV and original Manual `.xlsx` were each run through the frontend adapter: 661 records processed, 480 with writes, 181 zero-write, 1,180 cells written, 24 unsupported, 341 records with unresolved fields, 0 formula changes, and workbook preservation checks passed. The adapter uses the clean template without automatically applying the KATIVIK populated Reference VJR; it therefore reports the generic Stage 1 result for that input.
- The original source and template are hash-checked during each job. The clean template also remained unchanged in the integration tests.

## Runtime note

Streamlit is listed in `requirements.txt` but was not installed in this environment. The UI flow was tested through the app's injected Streamlit test double, and the app module keeps the Streamlit import lazy. After installing requirements, launch with `python -m streamlit run app.py`.

## Outcome

The web workflow is ready for local testing with a second project CSV or `.xlsx` Manual. Missing source data and unsupported values remain visible in the unresolved CSV and are not guessed. The Stage 1 business rules and mappings were not modified.
