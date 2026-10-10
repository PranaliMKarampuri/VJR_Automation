# Reusable Manual workbook inputs

Place a compatible Manual Team workbook (`.xlsx`) here. Run:

```powershell
python manual_to_vjr.py <filename>.xlsx --test
```

When exactly one `.xlsx` file is present, the filename may be omitted. The
runner reads each workbook; it does not modify the Manual source.

An estimator-approved mapping manifest is required before a record can be
written. See `reports/manual_to_vjr_approved_mapping_template.csv` for its
column contract. The first writer release can write only rectangular TAKE OFF
records whose dimensions, length, gauge, BP/MP/HP type, installation, accessory
codes, exact destination row, and estimator sign-off all pass validation.
Other documented item families remain review-only until their record-level
destination and transformations have a complete writer validator. Test mode
uses the same safety checks as full mode.

The manifest's `manifest_scope` must be `TEST_ONLY` or `APPROVED_PROJECT`.
`TEST_ONLY` rows are accepted only with `--test` and when both the Manual and
VJR test template are in a directory named `test_input`. The current project
has no approved manifest rows; do not label test fixtures as project approvals.
