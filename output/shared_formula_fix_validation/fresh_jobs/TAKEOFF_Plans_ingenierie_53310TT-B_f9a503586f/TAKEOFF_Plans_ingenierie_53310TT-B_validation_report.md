# TAKEOFF_Plans_ingenierie_53310TT-B_canonical → VJR validation

## Run totals

- Total CSV records: **252**
- Records processed: **252**
- Records with at least one populated input: **224**
- Records aggregated into an existing group: **109**
- Unsupported records: **12**
- Records with review items (including partially populated): **18**
- Records entirely skipped for review: **16**
- Cells written: **535**
- Pressure types inferred from a unique matching Reference row: **0**
- Riser records populated from a unique same-size Reference row: **0**
- Row-level audit: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-qn5vpo36\TAKEOFF_Plans_ingenierie_53310TT-B_canonical_record_audit.csv`
- Unresolved records/fields: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-qn5vpo36\TAKEOFF_Plans_ingenierie_53310TT-B_canonical_unresolved.csv`
- Reference comparison: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-qn5vpo36\TAKEOFF_Plans_ingenierie_53310TT-B_canonical_reference_comparison.csv`

## Source/VJR reconciliation

- Rectangular Manual length total: **342.5000000000000000**; grouped/planned TAKE OFF total: **342.5000000000000000** (unit preserved per source row).
- Riser length imported from unique Reference matches: **0 m**; unmatched risers remain unresolved for missing project drop/direction data.
- Spiral Manual length total: **692.4**; grouped/planned SPIRALE total: **692.4**.
- Elbow Manual count total: **58**; grouped/planned 45°/90° count: **58**.
- Flex: **13** source count; not written where product block is unspecified.
- Accessory counts are represented as repeated PPT codes in TAKE OFF G:J except turning-vane codes 7/71, which are written once per grouped accessory type; overflow still uses continuation rows with identical dimensions and zero length.

## Reference comparison

- Exact generated TAKE OFF rows matching Reference values: **0**.
- Comparison classifications: **{}**.
- The Reference has no Unit/Label fields, so exact project identity cannot be proven from row values alone.
- Every non-match and its field-level difference/reason is in the comparison CSV.

| Input | SHA-256 before generation | SHA-256 after generation |
|---|---|---|
| CSV | `959c9dc7ecd7eae1f4661787f51e7abb4b398a40aa7da73755c3c8267a7819f8` | `959c9dc7ecd7eae1f4661787f51e7abb4b398a40aa7da73755c3c8267a7819f8` |
| template | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` |

## Template integrity

- Source template used as base; workbook structure was not recreated.
- PASS: Workbook package membership and untouched package parts verified.
- PASS: Formula semantics preserved except 250 explicitly written TAKE OFF input formulas (A/B/D/E/F, rows 21:1004).
- PASS: Shared-formula masters and followers audited; touched groups were safely expanded.
- PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved.
- Calculation formulas modified: **0**; carry-forward input formulas replaced on allocated rows: **250**.
- Original CSV/template/Reference hashes unchanged.

## Errors and warnings

- CSV parse errors: **0**.
- Unsupported or missing project facts are itemized by source CSV row in the unresolved CSV; other supported records continue through generation.

## Frontend input provenance

- Processing run ID: `f9a503586f2a4262a224d422353eb5dc`
- Original input: `TAKEOFF_Plans_ingenierie_53310TT-B.xlsx` (.xlsx)
- Original input SHA-256: `e24c224d5bd5be1bcad01122f9bde2f5073038dbbf2a5369cc355c139bac3baf`
- Generated VJR: `TAKEOFF_Plans_ingenierie_53310TT-B_VJR.xlsx`
- Generated VJR SHA-256: `6e0fd5bdb7e857ccbea9f46b3293e645228f6d6737dbe54e6ca23914cdb6b1c8`
- Fixed VJR template: `Ventilation Jean Roy excel.xlsx`
- Fixed VJR template SHA-256 before/after: `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` / `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8`
- Detected input format: **Excel workbook**.
- Detected data sheet(s): **DATA, TAKEOFF-DUCTS, DATA-DUCTS**.
- Detected source columns: `{"DATA": ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS"], "TAKEOFF-DUCTS": ["UNITS", "Subject", "DUCT SIZE", "Total", "Length Unit", "ACCESSORIES1", "REMARK"], "DATA-DUCTS": ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS"]}`.
- Canonical categories: `{"elbows": 24, "dampers": 36, "other/unknown": 12, "riser": 4, "flex": 10, "accessories": 11, "spiral": 104, "rectangular": 51}`.
- Original DATA records/categories before authoritative grouping: **295** / `{"elbows": 24, "dampers": 36, "other/unknown": 12, "riser": 4, "flex": 10, "accessories": 11, "spiral": 104, "rectangular": 94}`.
- Canonical input records passed to the existing mapper: **252**. Missing source TYPE remains blank; input TYPE is not a required field.

## Evidence-based authoritative source selection

- Rectangular authority: `TAKEOFF-DUCTS`; detailed cross-reference: `DATA-DUCTS`.
- Detailed rectangular records: **94**; grouped authority records: **51**.
- Length reconciliation: detail **342.5**; grouped **342.5**; PASS: every grouped rectangular total reconciles to detailed and DATA totals.
- Explicit pressure evidence in authority sheet: `{"BP": 51}`.
- Detailed rectangular rows were cross-referenced and excluded from mapper input; only the reconciled grouped rows were sent for rectangular mapping.
- Material was unavailable; Metal Code D was left blank. Installation F remains calculated by the existing dimensional rule.
- Canvas associations without a matching system/size row in the detailed accessory sheet were withheld from VJR accessory cells and sent for review.
- Mapping engine: existing `manual_csv_to_vjr.generate`; no VJR mapping rules were changed.
