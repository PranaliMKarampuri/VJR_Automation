# Stage 1 baseline readiness

## Baseline

- Frozen archive: `baselines/stage1_baseline_2026-10-05.zip`
- SHA-256 manifest: `baselines/stage1_baseline_2026-10-05.sha256.json`
- Archive digest: `f609c94d6702c484efe5cec9d7467742912ff58e6e57d0c31d59e33bef5ff462`
- The archive includes the Stage 1 CSV engine, current compatibility modules, full tests, workflow guide, existing baseline output and validation artifacts, and the synthetic second-project fixture/output.
- All 26 archived file hashes were verified against the archive.

## Portability checks

- The active generic CSV engine (`manual_csv_to_vjr.py`) contains no project name, project-specific Manual row selection, or fixed Manual record count. The automated test checks for project-name and source-row gates.
- The engine enumerates whatever records are present in the supplied CSV. A two-record CSV and a three-record CSV were generated with the same clean VJR template; their record counts and generated rows were checked.
- TAKE OFF rows 20–1004 and SPIRALE diameter rows 11–63 are fixed coordinates of the existing template architecture, not project-specific Manual row numbers. Record-to-output row allocation is calculated for each run.
- The generic command uses the supplied template (default: `source_files/Ventilation Jean Roy excel.xlsx`) and never selects a Reference VJR automatically. Reference-based enrichment/comparison occurs only when an explicit `--reference` is supplied.
- `manual_to_vjr.py` remains as a compatibility dispatcher for older workbook/manifest runs and contains legacy project-specific routing. Use the documented direct CSV command for a new project.

## Second-project smoke run

`test_input/stage1_second_project_sample.csv` is synthetic, with three records, two sizes, and a same-size repeated group. It was run through the generic CLI with the clean template:

- Records processed: 3
- Written cells: 21
- Populated records: 3
- Unsupported records: 0
- Rectangular source and generated lengths: 8.50 m and 8.50 m
- Template sheet structure and validation report: generated successfully

This proves the workflow handles a different CSV record count. The actual second project's CSV has not yet been supplied.

## Test result

`python -m unittest discover -s tests -v` completed successfully: **65 tests passed** in 535 seconds. This includes the new two-record/three-record CSV pipeline test and workbook integration/preservation tests.

## User workflow

See `docs/stage1_user_workflow.md` for the CSV export, command, output files, unresolved-case handling, and validation steps. No established business mapping was changed for this readiness work.
