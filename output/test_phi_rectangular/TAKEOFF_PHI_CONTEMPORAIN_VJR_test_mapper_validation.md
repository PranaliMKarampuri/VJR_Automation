# TAKEOFF_PHI_CONTEMPORAIN_VJR_test_normalized_internal → VJR validation

## Run totals

- Total CSV records: **320**
- Records processed: **320**
- Records with at least one populated input: **320**
- Records aggregated into an existing group: **104**
- Unsupported records: **0**
- Records with review items (including partially populated): **298**
- Records entirely skipped for review: **0**
- Cells written: **2161**
- Pressure types inferred from a unique matching Reference row: **0**
- Riser records populated from a unique same-size Reference row: **0**
- Row-level audit: `C:\Users\91739\Desktop\VJR_Automation\output\test_phi_rectangular\TAKEOFF_PHI_CONTEMPORAIN_VJR_test_normalized_internal_record_audit.csv`
- Unresolved records/fields: `C:\Users\91739\Desktop\VJR_Automation\output\test_phi_rectangular\TAKEOFF_PHI_CONTEMPORAIN_VJR_test_normalized_internal_unresolved.csv`
- Reference comparison: `C:\Users\91739\Desktop\VJR_Automation\output\test_phi_rectangular\TAKEOFF_PHI_CONTEMPORAIN_VJR_test_normalized_internal_reference_comparison.csv`

## Source/VJR reconciliation

- Rectangular Manual length total: **3337.9**; grouped/planned TAKE OFF total: **3337.9** (unit preserved per source row).
- Riser length imported from unique Reference matches: **0 m**; unmatched risers remain unresolved for missing project drop/direction data.
- Spiral Manual length total: **0**; grouped/planned SPIRALE total: **0**.
- Elbow Manual count total: **0**; grouped/planned 45°/90° count: **0**.
- Flex: **0** source count; not written where product block is unspecified.
- Accessory counts are represented as repeated PPT codes in TAKE OFF G:J; continuation rows repeat size and use length zero.

## Reference comparison

- Exact generated TAKE OFF rows matching Reference values: **0**.
- Comparison classifications: **{}**.
- The Reference has no Unit/Label fields, so exact project identity cannot be proven from row values alone.
- Every non-match and its field-level difference/reason is in the comparison CSV.

| Input | SHA-256 before generation | SHA-256 after generation |
|---|---|---|
| CSV | `029da889565695252cab037f84e9fcefd7e3fed7c15bae667d0fab5ab7e19f95` | `029da889565695252cab037f84e9fcefd7e3fed7c15bae667d0fab5ab7e19f95` |
| template | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` | `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8` |

## Template integrity

- Source template used as base; workbook structure was not recreated.
- PASS: Workbook package membership and untouched package parts verified.
- PASS: All calculation formulas preserved; formula differences are limited to explicitly populated carry-forward input cells.
- PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved.
- Calculation formulas modified: **0**; carry-forward input formulas replaced on allocated rows: **1075**.
- Original CSV/template/Reference hashes unchanged.

## Errors and warnings

- CSV parse errors: **0**.
- Unsupported or missing project facts are itemized by source CSV row in the unresolved CSV; other supported records continue through generation.
