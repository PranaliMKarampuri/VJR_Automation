# Pavillon C generated vs manual VJR comparison

## Scope and identity

- Manual/reference workbook: `Downloads/TAKEOFF_Campus MIL Pavillon C/TAKEOFF_Campus MIL Pavillon C/TAKEOFF_Plans Pav. C -révisé add1à8_VJR.xlsx`.
- Generated workbook: `output/jobs/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_2e354dcc6b/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_VJR.xlsx`, processing run `2e354dcc6b704f8a881f209612e389cf`.
- Same-project evidence: both VJR filenames contain the same Pavillon C source project name, and the manual VJR is beside the matching source workbook and PDF in the Pavillon C project folder. The generated workbook is the one recorded by the Pavillon C validation report. No other project's workbook was used.
- Generated workbook SHA-256 is `1d604e5d8f5419e9360562fce9f9b1492e2d28610729821b7e8385491aee1dcf`, matching the documented generation run. The manual workbook was open in another process; it was read without saving. Its Windows hash command was blocked by the file lock, so no manual hash is reported.
- Both files have the same 13 sheet names. The shared relevant sheets are `TAKE OFF`, `SPIRALE`, `ROND ET FLEX`, and `COMPARATIF FLEX` (the latter was not used for row-level quantity matching because it is a derived comparison/price sheet). `SOMM. T.OFF` is a calculated summary sheet.

## Result summary

Counts below are row/diameter/block comparisons from `row_level_comparison.csv`.

| Sheet | MATCH | MISSING IN GENERATED | EXTRA IN GENERATED | QUANTITY DIFFERENCE | MAPPING DIFFERENCE | UNRESOLVED |
|---|---:|---:|---:|---:|---:|---:|
| TAKE OFF | 88 | 499 | 436 | 0 | 9 | 123 |
| SPIRALE | 43 | 0 | 0 | 10 | 0 | 0 |
| ROND ET FLEX | 30 | 0 | 0 | 9 | 0 | 0 |

TAKE OFF has 436 additional generated records in the row-level CSV; they are classified as `EXTRA IN GENERATED`.

Manual TAKE OFF contains 123 rows with Installation `RI`; 122 have a nonzero length and one has zero length. Generated TAKE OFF contains zero `RI` rows. The 123 manual entries are listed as unresolved because the Pavillon C source workbook has zero length on all 197 riser records, so these reference lengths cannot be tied to a specific source record from the available identifiers. Direction was not used as a requirement. The manual values remain evidence, not proof of source quantity.

For SPIRALE, 10 diameter rows differ in the direct input cells B:P (length/fittings/damper inputs); 43 agree. For ROND ET FLEX, 9 direct quantity/type blocks differ and 30 agree. The manual workbook has nonzero flex/round quantities where the generated workbook has blanks or zeroes in several blocks. Formula-derived price totals and summary totals are not treated as source quantities: the generated workbook's cached formula results include zeros and may not reflect a recalculation by Excel.

## Comparison method and limits

- `TAKE OFF`: compared multisets using width, height, length, metal, TYPE (BP/MP/HP), Installation, then compared G:J accessory codes. This avoids row-number-only matching and captures accessory-only differences. The VJR sheet does not expose system/unit or source record IDs, so identical fingerprints can collide across systems. Consequently, these are content-fingerprint comparisons, not proven system-level associations. Large missing/extra counts must be reviewed against source identifiers before interpreting them as true omissions.
- `SPIRALE`: matched the diameter row and compared direct takeoff inputs B:P (spiral length, elbow/reducer/tap/damper quantities and related fields); calculated columns Q:AP were excluded.
- `ROND ET FLEX`: compared direct quantity/type inputs for round duct and the flexible/semi-rigid product blocks; price and cost formulas were excluded.
- Same-size/same-length alone was not used. Where a matching source record could not be established, particularly risers, the comparison remains unresolved.
- `SOURCE_DATA_MISSING` matters for all 197 riser source rows: their source lengths are zero. No length was inferred from duct dimensions or unrelated duct lengths. The manual VJR's 123 RI rows include nonzero values, but there is no reliable row-to-record/system join in the files compared.
- Openpyxl emitted its known warning that the workbook's Data Validation extension is unsupported. Workbooks were opened read-only and not saved; no repair behavior was triggered or tested in Excel during this comparison.

## Artifacts

- `row_level_comparison.csv`: detailed sheet, classification, manual/generated row, content key, field evidence, and reason.
- `summary.json`: machine-readable counts and riser evidence.
- `workbook_contents.json`: read-only extracted populated cells for audit.
- `inspect_workbooks.py`, `compare_vjrs.py`, `preview_rows.py`: diagnostic scripts only; no production code or workbook was changed.

This comparison is evidence for review, not a claim that the generated workbook is equivalent to the manual VJR. The missing system/source IDs and zero source riser lengths prevent a fully authoritative record-level reconciliation.
