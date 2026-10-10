# VENTILATION_DUCTING.pptx rule layer

## Implementation scope

Added `vjr_ppt_rules.py` as a pure decision layer. It does not open, edit, or serialize Excel workbooks, and it does not call or change the Phase 1–6A pipelines. Each helper returns a structured `RuleResult` with `status` (`SAFE`, `NEEDS_CONFIRMATION`, or `UNSUPPORTED`), a reason, and the supporting PPT slide/rule.

The PPTX was inspected as OOXML in presentation order. The rules below cite its slide numbers.

## Rules implemented

The implementation covers the ten requested rule sets, with safe gates where the PPT requires information or client confirmation:

1. **Rectangular duct** — orders the larger dimension first; preserves the supplied dimensions/unit; validates documented BP/MP/HP pressure labels; evaluates HA/RE using width and height; returns `NEEDS_CONFIRMATION` at exactly 36 inches; sets risers to RI. Rectangular elbow generation is intentionally absent. Metric dimensions are retained in mm for rectangular records.
2. **Accessories** — supports only documented codes 1, 21, 7, 71, 8, 10, 100, 11, 12, 13, 131, and 14. Unknown codes are `UNSUPPORTED`. More than four supported accessories split into continuation rows that repeat size and set continuation length to zero (PPT slides 26 and 29). Code 9 is not treated as a generally available accessory.
3. **Risers** — rectangular risers retain size and meter length in TAKE OFF with RI. Spiral drops are added in feet at the given diameter. An already represented spiral riser is flagged for review to avoid duplicate counting.
4. **Spiral ducts** — sums repeated exact diameters, supports 45°/90° elbow counts, and converts metric diameter/length inputs to inches/feet. No diameter rounding or inferred grouping is performed. Y-type is safe only when two connecting ducts and two 45° elbows are explicitly supplied.
5. **Flex** — ordinary flex is five feet per count. ERV flex is ten feet per connection only with explicit client confirmation. Equipment-specific flex remains confirmation-required.
6. **Canvas** — code 8 is available. Automatic eligibility requires explicit facts for direct fan/AHU/RTU connection and the kitchen hood/dryer exception; otherwise the result is not safe.
7. **MBD** — explicit rectangular MBD maps to code 13 or 131 at the documented >30-inch boundary. Explicit spiral MBD targets the SPIRALE damper field. Topology alone never adds an MBD.
8. **Non-balancing spiral damper** — conversion to rectangular, zero length, and the supplied damper code is available only when width, height, and a documented accessory code are explicit. Missing size returns `NEEDS_CONFIRMATION`.
9. **Multiple systems** — requires the explicit system/unit name and returns it separately with its associated accessory rows. Missing names return `NEEDS_CONFIRMATION`; duct size is not used to infer a system.
10. **Special items** — reducers, taps, caps, unions, demolition, plenum, kitchen hood special handling, equipment flex, and unrecognized items remain `UNSUPPORTED` in this layer.

## Deliberate blocks and confirmation requirements

- Exactly 36 inches has no HA/RE boundary instruction in the PPT; the function does not select either.
- ERV flex requires client confirmation (slide 13); equipment-specific flex requires a confirmed type (slide 38).
- Canvas inference requires both direct-connection and exception facts (slides 5 and 12).
- Rigid metal round MBD code 9 is described on slide 18, but its rectangular row/size representation is not defined in the accessory mapping. The layer does not generate a code-9 rectangular record.
- Spiral non-balancing damper conversion needs actual rectangular width and height; no dimensions are invented.
- Y-type branches require actual dimensions and the two stated elbows; no branch dimensions or allocations are invented.
- Special-item rules are documented but not sufficiently specified here for automatic conversion; they are blocked rather than guessed.
- Workbook row insertion, destination coordinates, formulas, formatting, merged cells, drawings, validations, and preservation behavior are outside this rule layer and remain unimplemented.

## Verification

- New rule-engine tests: **19 passed** (`python -m unittest discover -s tests -p test_vjr_ppt_rules.py -v`). The requested cases are covered, including subtests for all documented accessory codes and all three pressure labels.
- Full existing plus new test suite: **43 passed** (`python -m unittest discover -s tests -v`; 499.696 seconds).

## Files created

- `vjr_ppt_rules.py`
- `tests/test_vjr_ppt_rules.py`
- `reports/vjr_ppt_rules_implementation.md`
