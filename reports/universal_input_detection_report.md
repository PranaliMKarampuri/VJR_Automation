# Universal Stage 1 input detection report

## Implementation

Stage 1 now inspects CSV headers and workbook sheets instead of requiring a fixed uploaded schema or a literal `TYPE` column. The detector scores header rows and aliases, prefers a canonical `DATA` sheet over presentation/pivot tabs, and records detected sheets and source columns. The normalizer emits canonical records with category, dimensions, length/unit, type/pressure-material, system unit, accessories, and source sheet/row/raw fields. It adapts those records to the existing mapper CSV boundary. The VJR mapping and writer were not rewritten. When uploaded `TYPE` is absent, the adapter supplies an empty internal mapper field; it does not infer material or pressure.

The frontend result now shows detected input format, selected data sheet(s), category counts, supported-category count, processed/written/unresolved counts, and downloads for the VJR, validation report, record audit, and unresolved fields. The fixed VJR template is unchanged and backend preservation checks remain required.

## Test inputs and detection

| Input | Detected format / sheet | Records normalized | Category counts |
|---|---|---:|---|
| KATIVIK Manual Excel: the KATIVIK workbook in `source_files/` | Excel / `DATA` | 661 | elbows 75, dampers 90, accessories 110, flex 41, rectangular 203, riser 21, spiral 116, other/unknown 5 |
| KATIVIK Manual CSV: `output/TAKEOFF_CS_KATIVIK_Manual.csv` | CSV / `CSV` | 661 | Same counts as Excel |
| Local PHI workbook: `test_input/TAKEOFF_PHI CONTEMPORAIN(1).xlsx` | PHI adapter / `DATA-DUCT`, exact rectangular rows | 320 | rectangular 320 |
| Requested `TAKEOFF_Plans_ingenierie_53310TT-B.xlsx` | **Not present in the workspace; exact workbook run could not be performed.** | — | — |

The local PHI workbook validates the demonstrated field interpretation `Subject`, `DUCT SIZE`, `Length`, `Length Unit`, `ACCESSORIES1/2`, and `UNITS`, but has `DATA-DUCT` and uses PHI inches in that specific adapter. It is not claimed as a substitute for the named Plans workbook. In the generic normalizer, `RECT. DUCT m` maps to `rectangular`; `350X250` maps to width 350 and height 250; and length/unit are preserved. A regression test covers this with reordered headers and no uploaded `TYPE` column.

### Column mapping observed

- KATIVIK `DATA`: `Subject` → category; `DUCT SIZE` → dimensions; `Length` + `Length Unit` → length and unit; `TYPE` → type/material; `ACCESSORIES1/2` → accessories; `UNITS` → system unit; `Count`, `Label`, `REMARK`, `DAMPER TYPE`, and `Page Index` → mapper fields/source context. `BRAND` and `MODEL` are retained in raw source fields.
- Local PHI `DATA-DUCT`: `Subject`, `Page Index`, `Label`, `Count`, `DUCT SIZE`, `ACCESSORIES1/2`, `REMARK`, `UNITS`, `Length`, `Length Unit`, and `DAMPER TYPE` were read by the existing PHI adapter. That adapter selects only exact rectangular subject rows and passes them to the same mapper entry point.
- Generic inputs accept recognized aliases for category/subject, size/dimensions, length and unit, count/quantity, label/name, type/material, accessory fields, remark/notes, and system/unit. A `TYPE` heading is optional.

## Mapper and unresolved results

| Input | Existing mapper records | Records with VJR writes | Unresolved field findings |
|---|---:|---:|---:|
| KATIVIK Excel | 661 | 480 | 341 |
| KATIVIK CSV | 661 | 480 | 341 |
| Local PHI rectangular sample | 320 | 320 | 371 findings across 298 records |

KATIVIK unresolved findings are chiefly: pressure class absent (160); accessory has no unique same-system and size rectangular group (63); flex product is not identified sufficiently to choose VJR cells (41); accessory lacks size/type route (32); riser drop length missing or nonpositive (21); and damper codes RFCU/RCFU/BDD lack exact approved destination (19 combined). Five `vent cap` rows have no established destination. The remaining findings are listed with exact reasons in each job's unresolved CSV. These are mapper review results; category detection alone does not certify that the VJR mapping is known.

For the local PHI rectangular run, missing pressure is reported for 289 source lines; 52 acoustic insulation descriptions lack a supported thickness; 29 `18G` descriptions have no approved mapping; and one thermal insulation description has no approved mapping. No missing value was guessed. The unresolved CSVs and source-row audit are retained with the generated jobs under `output/universal_input_validation/`.

## Canonical boundary and preservation

KATIVIK Excel and CSV normalize to the same 661 records, category totals, mapper columns, and downstream result counts. Both feed `manual_csv_to_vjr.generate`; the approved writer/mapping layer remains the sole VJR writer. Generic inputs without a `TYPE` header are accepted and the internal mapper field is blank. The local PHI adapter is retained for backward-compatible PHI rectangular normalization; its exact scope is recorded above.

Template SHA-256 is checked before and after generation, and generated-workbook validation requires zero calculation-formula changes and preservation checks to pass. The generated VJR and its validation report are included in each retained job directory.

## Limitations

This detector supports recognized aliases, evidence in values, sheet names, and header samples; it does not promise automatic support for every unknown workbook. Ambiguous or unmapped items continue to the established mapper for explicit review. The named Plans workbook is absent, so its exact `DATA` sheet, column mapping, and record totals remain unverified. Supply that workbook to complete the requested project-specific run.
