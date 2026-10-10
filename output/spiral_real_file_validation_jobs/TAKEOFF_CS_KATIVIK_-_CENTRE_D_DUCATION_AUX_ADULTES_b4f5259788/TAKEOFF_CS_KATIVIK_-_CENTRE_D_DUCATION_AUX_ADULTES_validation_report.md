# TAKEOFF_CS_KATIVIK_-_CENTRE_D_DUCATION_AUX_ADULTES_canonical → VJR validation

## Run totals

- Total CSV records: **661**
- Records processed: **661**
- Records with at least one populated input: **478**
- Records aggregated into an existing group: **270**
- Unsupported records: **24**
- Records with review items (including partially populated): **295**
- Records entirely skipped for review: **135**
- Cells written: **1143**
- Pressure types inferred from a unique matching Reference row: **0**
- Riser records populated from a unique same-size Reference row: **0**
- Row-level audit: `C:\Users\91739\Desktop\VJR_Automation\output\spiral_real_file_validation_jobs\.stage1-oekxvi16\TAKEOFF_CS_KATIVIK_-_CENTRE_D_DUCATION_AUX_ADULTES_canonical_record_audit.csv`
- Unresolved records/fields: `C:\Users\91739\Desktop\VJR_Automation\output\spiral_real_file_validation_jobs\.stage1-oekxvi16\TAKEOFF_CS_KATIVIK_-_CENTRE_D_DUCATION_AUX_ADULTES_canonical_unresolved.csv`
- Reference comparison: `C:\Users\91739\Desktop\VJR_Automation\output\spiral_real_file_validation_jobs\.stage1-oekxvi16\TAKEOFF_CS_KATIVIK_-_CENTRE_D_DUCATION_AUX_ADULTES_canonical_reference_comparison.csv`

## Source/VJR reconciliation

- Rectangular Manual length total: **1151.7**; grouped/planned TAKE OFF total: **1151.7** (unit preserved per source row).
- Riser length imported from unique Reference matches: **0 m**; unmatched risers remain unresolved for missing project drop/direction data.
- Spiral Manual length total: **863.4**; grouped/planned SPIRALE total: **863.4**.
- Elbow Manual count total: **172**; grouped/planned 45°/90° count: **172**.
- Flex: **73** source count; not written where product block is unspecified.
- Accessory counts are represented as repeated PPT codes in TAKE OFF G:J except turning-vane codes 7/71, which are written once per grouped accessory type; overflow still uses continuation rows with identical dimensions and zero length.

## Reference comparison

- Exact generated TAKE OFF rows matching Reference values: **0**.
- Comparison classifications: **{}**.
- The Reference has no Unit/Label fields, so exact project identity cannot be proven from row values alone.
- Every non-match and its field-level difference/reason is in the comparison CSV.

| Input | SHA-256 before generation | SHA-256 after generation |
|---|---|---|
| CSV | `d2651b97c8a75f76eda8c300f17ebf20298531ce738119dfe26e0e3ddef68b6f` | `d2651b97c8a75f76eda8c300f17ebf20298531ce738119dfe26e0e3ddef68b6f` |
| template | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` |

## Template integrity

- Source template used as base; workbook structure was not recreated.
- PASS: Workbook package membership and untouched package parts verified.
- PASS: All formula XML preserved except 555 explicitly written TAKE OFF input cells (A/B/D/E/F, rows 21:1004).
- PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved.
- Calculation formulas modified: **0**; carry-forward input formulas replaced on allocated rows: **555**.
- Original CSV/template/Reference hashes unchanged.

## Errors and warnings

- CSV parse errors: **0**.
- Unsupported or missing project facts are itemized by source CSV row in the unresolved CSV; other supported records continue through generation.

## Frontend input provenance

- Original input: `TAKEOFF_CS KATIVIK - CENTRE D'ÉDUCATION AUX ADULTES.xlsx` (.xlsx)
- Original input SHA-256: `2e1c810f26945e1c02b288dcd9bd8c54accd5e4be20f6b28c91f6317725a8342`
- Fixed VJR template: `Ventilation Jean Roy excel.xlsx`
- Fixed VJR template SHA-256 before/after: `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` / `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8`
- Detected input format: **Excel workbook**.
- Detected data sheet(s): **DATA**.
- Detected source columns: `{"DATA": ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "BRAND", "MODEL", "DUCT SIZE", "TYPE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "DAMPER TYPE"]}`.
- Canonical categories: `{"elbows": 75, "dampers": 90, "accessories": 110, "flex": 41, "rectangular": 203, "riser": 21, "spiral": 116, "other/unknown": 5}`.
- Original DATA records/categories before authoritative grouping: **661** / `{"elbows": 75, "dampers": 90, "accessories": 110, "flex": 41, "rectangular": 203, "riser": 21, "spiral": 116, "other/unknown": 5}`.
- Canonical input records passed to the existing mapper: **661**. Missing source TYPE remains blank; input TYPE is not a required field.
- Mapping engine: existing `manual_csv_to_vjr.generate`; no VJR mapping rules were changed.
