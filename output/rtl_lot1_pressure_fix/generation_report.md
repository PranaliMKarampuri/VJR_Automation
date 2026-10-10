# rtl_lot1_normalized → VJR validation

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
- Row-level audit: `C:\Users\91739\Desktop\VJR_Automation\output\rtl_lot1_pressure_fix\rtl_lot1_normalized_record_audit.csv`
- Unresolved records/fields: `C:\Users\91739\Desktop\VJR_Automation\output\rtl_lot1_pressure_fix\rtl_lot1_normalized_unresolved.csv`
- Reference comparison: `C:\Users\91739\Desktop\VJR_Automation\output\rtl_lot1_pressure_fix\rtl_lot1_normalized_reference_comparison.csv`

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
| CSV | `f54bf87b1a97df558ba9df6f51877e1b8f7242de448ccee08e889ea7762967c8` | `f54bf87b1a97df558ba9df6f51877e1b8f7242de448ccee08e889ea7762967c8` |
| template | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` |

## Template integrity

- Source template used as base; workbook structure was not recreated.
- PASS: Workbook package membership and untouched package parts verified.
- PASS: All formula XML preserved except 175 explicitly written TAKE OFF input cells (A/B/D/E/F, rows 21:1004).
- PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved.
- Calculation formulas modified: **0**; carry-forward input formulas replaced on allocated rows: **175**.
- Original CSV/template/Reference hashes unchanged.

## Errors and warnings

- CSV parse errors: **0**.
- Unsupported or missing project facts are itemized by source CSV row in the unresolved CSV; other supported records continue through generation.
