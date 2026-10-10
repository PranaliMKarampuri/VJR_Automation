# Pavillon C riser source investigation

## Source and method

- Read the original XLSX member directly from `C:\Users\91739\Downloads\TAKEOFF_Campus MIL Pavillon C.zip`: `TAKEOFF_Campus MIL Pavillon C/TAKEOFF_Plans Pav. C -révisé add1à8.xlsx`.
- Original archive member SHA-256: `d137df9a578f65069957a52e7560e845dfe73ff1ec8f6a27222c3bafe5231d18`. It is byte-identical to the prior extracted source copy used for comparison.
- `DATA` has no explicit record-ID column. The stable investigation identifiers below use `DATA!<physical row>`.
- Opened the source with formula and cached-value views. No workbook was saved. Compared all source riser rows with the downloaded canonical record audit.
- Full per-record evidence: [197 riser rows CSV](pavillon_c_197_riser_source_rows.csv). It includes source sheet/row and investigation ID, original size, Length, Length Unit, Count, all original source columns, canonical values, exact system/size related rows, formulas, and merged-cell evidence.

## Finding: the zero lengths originate in the source

- The original `DATA` sheet has **197** riser rows. Every one has numeric `Length = 0` and `Length Unit = "."`; none has a formula in the row. `Count` is quantity: 193 rows have 1, two have 2, and two have 3 (sum **203**).
- The canonical record audit contains the same 197 physical rows. Original and canonical `Length`, `Length Unit`, `Count`, and `DUCT SIZE` match for **197/197** records. The normalizer did not replace a positive length with zero.
- The original workbook contains **zero formula cells**, no defined names, no Excel tables, and no merged cells on DATA. No formula, merged-cell carry, or cached formula result can account for the zeros.
- `DATA-ACCESSORIES` repeats the riser source records and related accessories; its matching rows also have zero Length and `Length Unit = "."`.

## Other sheets and related rows

The workbook has no dedicated riser-length column or separate riser length schedule. Exact `(UNITS, DUCT SIZE)` matches exist, but their labels and fields show why they cannot supply a riser length:

| Sheet | Exact system/size matches | What the rows contain | Can supply riser length? |
|---|---:|---|---|
| `DATA-ACCESSORIES` | All 197 source rows have a match; 673 same-system/size candidate rows including other accessories | Riser detail copies with Length 0, unit `.` and Count; other rows may be CANEVAS or other accessories | No |
| `TAKEOFF-ACCESSORIES` | 196 source rows have a match; 323 candidate rows including other accessories; 139 unique `RISERS` lines | `Label = RISERS`, `Total` quantity; no Length or Length Unit columns | No |
| `DATA-DUCT` | 96 source rows have a match; 1,232 candidate detail rows | Detail records labeled `RECT. DUCT` or `SPIRAL DUCT`, with their own Length values | No safe riser identity; same system/size alone does not identify a riser drop |
| `TAKEOFF-DUCT` | 91 source rows have a match; 128 grouped rows | Grouped `RECT. DUCT`/`SPIRAL DUCT` rows; `Total` plus Length Unit is duct takeoff length | No safe riser identity; these are ordinary duct rows, not labeled RISERS |
| `TO-DUCT` | 12 source rows have a match; 12 grouped rows | Derived/grouped ordinary duct totals | No |
| `TAKEOFF-SPIRAL DUCT (FAN)` | 5 source rows have a match; 5 grouped rows | `SPIRAL DUCT` totals in feet | No; not labeled RISERS |

The `TAKEOFF-ACCESSORIES` RISERS quantities total **202** across 139 unique system/size groups, versus source `Count` total **203** across 140 groups. The group `VARCF-305 / 1575X1981` has source count 1 but no matching RISERS row there. This is a quantity reconciliation gap, not a length source.

## Fifteen single-number sizes and the 35-inch cases

The 15 single-number `DUCT SIZE` values are `125` (6 records), `250` (2), `305` (5), and `881` (2). They are in the size field, not Length or Count. The source has no separate diameter-unit column; `Length Unit = "."` is not evidence for the size unit.

The two `881` records are `DATA!4554` (`VE-331-VE-0001`) and `DATA!4555` (`VE-331-VE-0002`). Both have Length 0, Length Unit `.`, Count 1, and size 881. Their exact `DATA-ACCESSORIES` rows include both `RISERS` and `CANEVAS`; matching `TAKEOFF-ACCESSORIES` rows show `Total = 1` for each label. These totals are counts, not lengths.

The prior 35-inch finding converts 881 mm to approximately 34.69 inches and rounds to nominal 35. That conversion depends on treating the numeric size as millimeters; the riser source row itself does not state the size unit. In addition, the VJR SPIRALE diameter list contains 34 and 36 but not 35. Do not snap it to a neighboring size. Regardless of diameter interpretation, these two rows still lack a positive riser length.

## Conclusion and safe next step

The lengths are **absent from the source riser records**, not lost by extraction or normalization. Matching regular duct records contain lengths, but they have different record identities and duct labels; transferring those values to risers would invent a riser drop. Zero cannot be assumed to mean a valid zero-length riser.

Ask the source owner/estimator to provide an explicit positive riser length for each affected `DATA!` row and confirm the dimension unit for the 15 single-number sizes, especially the two 881 entries. Keep those rows unresolved until those facts are supplied. The 35-inch size also needs a confirmed nominal-diameter treatment/reference row before spiral mapping. No code or workbook change is recommended from this evidence alone.
