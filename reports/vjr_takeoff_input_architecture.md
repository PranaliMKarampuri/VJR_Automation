# VJR TAKE OFF input architecture

## Evidence inspected

Compared `source_files/Ventilation Jean Roy excel.xlsx` (clean template) with `source_files/TAKEOFF_CS KATIVIK - CENTRE D'Ã‰DUCATION AUX ADULTES_VJR.xlsx` (populated reference), without saving either source.

The clean template preallocates rows 20-1004 (985 rows). Every row has 39 calculation formulas in K:AY. C and G:J are blank inputs throughout. In A/B, 983 of the 985 cells are carry-forward formulas (rows 21-1004); D has 982; E/F each have 984. Row 20 has blank A:F. These formulas in input columns are carry-forward defaults, not takeoff calculation formulas.

The populated reference uses rows 20-67 and has no new project entries below row 67. A/B/C contain literal inputs on rows starting duct entries, replacing the clean template carry formulas in A/B. E/F have literal values where type or installation is set or changes; carry-forward formulas remain where the prior value applies. K:AY calculation formulas remain. This confirms A/B and D:F are input destinations when a new size or classification must be stated. The reference did not insert rows.

## A:J roles

| Columns | Intended use | Clean-template pattern |
|---|---|---|
| A, B | Width and height. Enter literal values when starting a different duct size; continuation formulas repeat the prior size. | Row 20 blank; later rows carry-forward formulas. |
| C | Length. Literal user input; a continuation accessory row uses zero length. | Blank input cells. |
| D | Metal/material code when specified. Otherwise the carry-forward default may be retained only when it is still correct. | Blank in row 20; carry-forward formula later. |
| E | BP/MP/HP type when established by Scope/Remarks. | Blank in row 20; carry-forward formula later. |
| F | HA/RE/RI installation. | Blank in row 20; carry-forward formula later. |
| G:J | Up to four accessory codes. | Blank input cells. |

K:AY are calculation/output columns and must not receive Manual values. Their existing formulas should remain intact on every used row.

## Safe row creation and formula handling

1. Use the clean template and allocate the next unused preformatted row from 20 through 1004. Do not insert rows while capacity remains; the reference uses this preallocated area.
2. Treat A:J as the input band. At a new group start, write literal A/B dimensions and explicit D/E/F values (or a blank when the source does not establish one); write the Manual length in C and accessory codes in G:J. A continuation row repeats A/B, writes C=0 and uses G:J for remaining accessories.
3. The A/B/D/E/F formulas are carry-forward conveniences and may be replaced by explicit user inputs in those input cells. Preserve carry-forward formulas on unused rows. Preserve all K:AY formulas on used and unused rows.
4. No formula copy-down is needed for rows 20â€“1004: the clean template already contains the row formulas and styles. If future projects exceed row 1004, extend the sheet by cloning a complete preformatted input row: translate K:AY formulas to the new row, retain row styles/height and worksheet validations, and initialize A:J using the same carry-forward pattern before applying the new group values. Package-level OOXML editing is required to preserve unsupported Excel extensions.

## Recommended writer strategy

Build a separate output from the clean template. Group rectangular Manual records by Unit, size, and any distinct pressure/material classification; allocate one preformatted row per group and continuation rows for accessory overflow. Allow writes over carry-forward formulas only in A/B/D:F, only when those fields are explicitly assigned for a used row. Write C and G:J only to their blank inputs. Keep K:AY byte-for-byte/formula-identical and leave all unused row carry formulas untouched. Validate that formula differences are confined to the explicitly populated carry-forward input cells, while every calculation formula, style, merged range, validation, drawing, named range, and workbook package part remains preserved.
