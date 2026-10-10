# TAKEOFF_Plans_Pav_C_source_canonical → VJR validation

## Run totals

- Total CSV records: **4809**
- Records processed: **4809**
- Records with at least one populated input: **3204**
- Records aggregated into an existing group: **2294**
- Unsupported records: **0**
- Records with review items (including partially populated): **2009**
- Records entirely skipped for review: **1605**
- Cells written: **5358**
- Pressure types inferred from a unique matching Reference row: **0**
- Riser records populated from a unique same-size Reference row: **0**
- Row-level audit: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-3oxiac2z\TAKEOFF_Plans_Pav_C_source_canonical_record_audit.csv`
- Unresolved records/fields: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-3oxiac2z\TAKEOFF_Plans_Pav_C_source_canonical_unresolved.csv`
- Reference comparison: `C:\Users\91739\Desktop\VJR_Automation\output\shared_formula_fix_validation\fresh_jobs\.stage1-3oxiac2z\TAKEOFF_Plans_Pav_C_source_canonical_reference_comparison.csv`

## Source/VJR reconciliation

- Rectangular Manual length total: **7324.40000000000009524**; grouped/planned TAKE OFF total: **7324.40000000000009524** (unit preserved per source row).
- Riser length imported from unique Reference matches: **0 m**; unmatched risers remain unresolved for missing project drop/direction data.
- Spiral Manual length total: **5782.2**; grouped/planned SPIRALE total: **5782.2**.
- Elbow Manual count total: **909**; grouped/planned 45°/90° count: **909**.
- Flex: **1688** source count; not written where product block is unspecified.
- Accessory counts are represented as repeated PPT codes in TAKE OFF G:J except turning-vane codes 7/71, which are written once per grouped accessory type; overflow still uses continuation rows with identical dimensions and zero length.

## Reference comparison

- Exact generated TAKE OFF rows matching Reference values: **0**.
- Comparison classifications: **{}**.
- The Reference has no Unit/Label fields, so exact project identity cannot be proven from row values alone.
- Every non-match and its field-level difference/reason is in the comparison CSV.

| Input | SHA-256 before generation | SHA-256 after generation |
|---|---|---|
| CSV | `599c8e939862d391771126f8d843811760f36f7e672a2804ffc0f6efc50e1fe1` | `599c8e939862d391771126f8d843811760f36f7e672a2804ffc0f6efc50e1fe1` |
| template | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` |

## Template integrity

- Source template used as base; workbook structure was not recreated.
- PASS: Workbook package membership and untouched package parts verified.
- PASS: Formula semantics preserved except 2660 explicitly written TAKE OFF input formulas (A/B/D/E/F, rows 21:1004).
- PASS: Shared-formula masters and followers audited; touched groups were safely expanded.
- PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved.
- Calculation formulas modified: **0**; carry-forward input formulas replaced on allocated rows: **2660**.
- Original CSV/template/Reference hashes unchanged.

## Errors and warnings

- CSV parse errors: **0**.
- Unsupported or missing project facts are itemized by source CSV row in the unresolved CSV; other supported records continue through generation.

## Frontend input provenance

- Processing run ID: `9ab183ee73014b95a7a161fb716d72c5`
- Original input: `TAKEOFF_Plans_Pav_C_source.xlsx` (.xlsx)
- Original input SHA-256: `d137df9a578f65069957a52e7560e845dfe73ff1ec8f6a27222c3bafe5231d18`
- Generated VJR: `TAKEOFF_Plans_Pav_C_source_VJR.xlsx`
- Generated VJR SHA-256: `1d604e5d8f5419e9360562fce9f9b1492e2d28610729821b7e8385491aee1dcf`
- Fixed VJR template: `Ventilation Jean Roy excel.xlsx`
- Fixed VJR template SHA-256 before/after: `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` / `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8`
- Detected input format: **Excel workbook**.
- Detected data sheet(s): **DATA, TAKEOFF-DUCT, DATA-DUCT**.
- Detected source columns: `{"DATA": ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "DAMPER TYPE"], "TAKEOFF-DUCT": ["UNITS", "Label", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "Total", "Length Unit"], "DATA-DUCT": ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "DAMPER TYPE"]}`.
- Canonical categories: `{"elbows": 658, "accessories": 132, "dampers": 274, "flex": 1369, "riser": 197, "spiral": 1659, "rectangular": 520}`.
- Original DATA records/categories before authoritative grouping: **6289** / `{"elbows": 658, "accessories": 132, "dampers": 274, "flex": 1369, "rectangular": 2000, "riser": 197, "spiral": 1659}`.
- Canonical input records passed to the existing mapper: **4809**. Missing source TYPE remains blank; input TYPE is not a required field.

## Evidence-based authoritative source selection

- Rectangular authority: `TAKEOFF-DUCT`; detailed cross-reference: `DATA-DUCT`.
- Detailed rectangular records: **2000**; grouped authority records: **652**.
- Length reconciliation: detail **7324.4**; grouped **7324.4**; PASS: every grouped rectangular total reconciles to detailed and DATA totals.
- Explicit pressure evidence in authority sheet: `{"": 597, "BP": 27, "MP": 28}`.
- Detailed rectangular rows were cross-referenced and excluded from mapper input; only the reconciled grouped rows were sent for rectangular mapping.
- Material was unavailable; Metal Code D was left blank. Installation F remains calculated by the existing dimensional rule.
- Canvas associations without a matching system/size row in the detailed accessory sheet were withheld from VJR accessory cells and sent for review.
- Mapping engine: existing `manual_csv_to_vjr.generate`; no VJR mapping rules were changed.
