# rtl_lot1_normalized_canonical → VJR validation

## Run totals

- Total CSV records: **128**
- Records processed: **128**
- Records with at least one populated input: **104**
- Records aggregated into an existing group: **0**
- Unsupported records: **0**
- Records with review items (including partially populated): **24**
- Records entirely skipped for review: **24**
- Cells written: **363**
- Pressure types inferred from a unique matching Reference row: **0**
- Riser records populated from a unique same-size Reference row: **0**
- Row-level audit: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-hg_ywqe2\rtl_lot1_normalized_canonical_record_audit.csv`
- Unresolved records/fields: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-hg_ywqe2\rtl_lot1_normalized_canonical_unresolved.csv`
- Reference comparison: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-hg_ywqe2\rtl_lot1_normalized_canonical_reference_comparison.csv`

## Source/VJR reconciliation

- Rectangular Manual length total: **174.5**; grouped/planned TAKE OFF total: **174.5** (unit preserved per source row).
- Riser length imported from unique Reference matches: **0 m**; unmatched risers remain unresolved for missing project drop/direction data.
- Spiral Manual length total: **7.9**; grouped/planned SPIRALE total: **7.9**.
- Elbow Manual count total: **1**; grouped/planned 45°/90° count: **1**.
- Flex: **0** source count; not written where product block is unspecified.
- Accessory counts are represented as repeated PPT codes in TAKE OFF G:J except turning-vane codes 7/71, which are written once per grouped accessory type; overflow still uses continuation rows with identical dimensions and zero length.

## Reference comparison

- Exact generated TAKE OFF rows matching Reference values: **0**.
- Comparison classifications: **{}**.
- The Reference has no Unit/Label fields, so exact project identity cannot be proven from row values alone.
- Every non-match and its field-level difference/reason is in the comparison CSV.

| Input | SHA-256 before generation | SHA-256 after generation |
|---|---|---|
| CSV | `a9ddad27653a9ccc09fcf2e7c5fe711023ca96b402889c54cf1f95132e523164` | `a9ddad27653a9ccc09fcf2e7c5fe711023ca96b402889c54cf1f95132e523164` |
| template | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` |

## Template integrity

- Source template used as base; workbook structure was not recreated.
- PASS: Workbook package membership and untouched package parts verified.
- PASS: Formula semantics preserved except 175 explicitly written TAKE OFF input formulas (A/B/D/E/F, rows 21:1004).
- PASS: Shared-formula masters and followers audited; touched groups were safely expanded.
- PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved.
- Calculation formulas modified: **0**; carry-forward input formulas replaced on allocated rows: **175**.
- Original CSV/template/Reference hashes unchanged.

## Errors and warnings

- CSV parse errors: **0**.
- Unsupported or missing project facts are itemized by source CSV row in the unresolved CSV; other supported records continue through generation.

## Frontend input provenance

- Processing run ID: `6d5d3a4fb2d845f285550a900f0752b6`
- Original input: `rtl_lot1_normalized.csv` (.csv)
- Original input SHA-256: `f54bf87b1a97df558ba9df6f51877e1b8f7242de448ccee08e889ea7762967c8`
- Generated VJR: `rtl_lot1_normalized_VJR.xlsx`
- Generated VJR SHA-256: `cf888337035a34451b245570ecc2a35085bfe131208faa383cbfdc5c9bac4955`
- Fixed VJR template: `Ventilation Jean Roy excel.xlsx`
- Fixed VJR template SHA-256 before/after: `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` / `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8`
- Detected input format: **CSV**.
- Detected data sheet(s): **CSV**.
- Detected source columns: `{"CSV": ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "DUCT SIZE", "TYPE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "DAMPER TYPE", "source_row", "source_sheet", "source_fields", "source_physical_row", "cross_reference_rows"]}`.
- Canonical categories: `{"elbows": 1, "accessories": 27, "dampers": 48, "riser": 20, "spiral": 1, "rectangular": 31}`.
- Original DATA records/categories before authoritative grouping: **128** / `{"elbows": 1, "accessories": 27, "dampers": 48, "riser": 20, "spiral": 1, "rectangular": 31}`.
- Canonical input records passed to the existing mapper: **128**. Missing source TYPE remains blank; input TYPE is not a required field.
- Mapping engine: existing `manual_csv_to_vjr.generate`; no VJR mapping rules were changed.
