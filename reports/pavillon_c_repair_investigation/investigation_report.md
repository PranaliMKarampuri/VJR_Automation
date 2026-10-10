# Pavillon C VJR Excel repair warning investigation

## Scope and preservation

Read-only investigation of the exact latest frontend output and its downloaded copy. No workbook was overwritten, repaired, or deleted, and no production code was changed. The workbook was only opened through openpyxl in read-only and normal load modes; it was never saved.

## Exact file identification

- Latest frontend job: `output/jobs/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_8bc5caddc1/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_VJR.xlsx` (job run ID `8bc5caddc10c4a82bcc160ecbfd9d1d2`).
- Downloaded copy: `%USERPROFILE%/Downloads/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_VJR.xlsx`.
- Both files are 962,887 bytes and SHA-256 `a9e2f1209e107f1f6e15a59a005b2f1bddd03268faaa0e498fdc901cca29e921`.
- The two files have the same 83 ZIP members, and every member is byte-identical. The downloaded workbook is exactly the latest generated workbook.
- Original VJR template: `source_files/Ventilation Jean Roy excel.xlsx`, 1,039,276 bytes, SHA-256 `a6775a49eb639efd3e3d34805ee2091dd5305cbd839a2ceccb8aa1c053cdcea8`.

## Package, worksheet, and XML findings

- ZIP CRC validation succeeded for generated output, downloaded copy, and template.
- All XML and relationship XML parts parse. All internal OPC relationship targets resolve after URI path normalization.
- Every `mc:Ignorable` prefix in the generated worksheets is declared on that root element. In particular, edited `sheet2.xml` and `sheet3.xml` both declare `x14ac` and retain `mc:Ignorable="x14ac"`. The earlier known defect (an undeclared `x14ac` prefix after XML serialization) is not present in this output. Namespace declaration order differs; `sheet3.xml` drops an unused `xmlns:r` declaration. This has no effect on element names or relationships in that part.
- Generated package membership exactly matches the template. Only `xl/worksheets/sheet2.xml` and `xl/worksheets/sheet3.xml` differ byte-for-byte. All other parts, including styles, workbook metadata, defined names, relationships, drawings, image assets, tables, comments, VML, threaded comments, validations, and printer settings, are byte-identical to the template.
- The generated workbook retains the same worksheet feature inventory as the template: 13 worksheets, 1,389 merged cells total, 3 standard data-validation elements, 3 x14 data-validation elements, 4 drawing references, 1 table reference, four drawing XML parts, ten images, and one table. Defined name `ESTIMATEUR` remains present.
- `sheet2.xml` has 49,625 cell records in both template and generated versions; 3,567 cell records differ semantically across 534 rows. This includes input values written into template input/formula cells and 2,660 carry-forward formulas deliberately replaced by inputs. `sheet3.xml` has 1,954 cells in both versions; 27 differ across 9 rows. The complete per-cell records, including cell type, formula text and attributes, value, and inline string, are in `cell_xml_diff.json`.
- Both changed worksheets retain their merge counts (74 on `TAKE OFF`, 1 on `SPIRALE`) and have no worksheet extension list or data-validation elements. Differences also include ElementTree serialization changes (XML declaration style, empty-element syntax, attribute ordering, and whitespace). The raw diffs are retained separately.
- The worksheet child elements remain in SpreadsheetML schema order, and every row has strictly ordered, unique cell references in both template and generated XML. These XML tree and ordering checks pass; the shared-formula cross-cell relationship check below does not.

## Confirmed structural defect: shared formula masters removed

The generated `TAKE OFF` XML has five shared-formula groups whose master cells were replaced by generated input values while surviving follower cells still refer to the groups. In the template each group has a master `<f t="shared" ref="..." si="...">formula</f>`. In the generated worksheet the master `<f>` is gone but followers retain `<f t="shared" si="..."/>`. Microsoft’s SpreadsheetML documentation defines the first formula in a shared group as its master and says the `si` identifies the group whose formula is held at that master. The generated worksheet therefore has shared-formula followers with no corresponding master; its XML is well-formed, but the shared-formula structure is semantically inconsistent with SpreadsheetML. [Microsoft CellFormula documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1)

| Shared index | Template master and range | Template master formula | Generated master cell | Remaining generated followers |
|---:|---|---|---|---|
| 417 | A546; A546:A609 | `IF(A545="","",IF(C546="","",A545))` | A546 becomes numeric value `450` | A553:A609 still have `t="shared" si="417"` |
| 418 | B546; B546:B609 | `IF(B545="","",IF(C546="","",B545))` | B546 becomes numeric value `300` | B553:B609 still have `t="shared" si="418"` |
| 419 | D546; D546:D609 | `IF(A546="","",IF(D545="","",D545))` | D546 formula removed; no value | D553:D609 still have `t="shared" si="419"` |
| 420 | E546; E546:E609 | `IF(A546="","",IF(E545="","",E545))` | E546 becomes inline string `BP` | E553:E609 still have `t="shared" si="420"` |
| 421 | F546; F546:F609 | `IF(A546="","",IF(F545="","",F545))` | F546 becomes inline string `HA` | F553:F609 still have `t="shared" si="421"` |

The changed input rows consume the template formula masters and followers at rows 546:552. The remaining follower runs start at row 553. The same five shared indices have valid masters in the template; these orphan follower groups appear only in the generated workbook. No orphan shared-formula indices were found in `sheet3.xml`.

This is the concrete OOXML defect found and the strongest evidence-based explanation for Excel's repair prompt: Excel must handle five formula groups whose defining master formula has been removed. The report does not claim direct Excel confirmation because that could not be obtained in this environment.

## Data-validation extension assessment

- The x14 data-validation extensions are on template sheets 1 and 4, not on edited sheets 2 and 3. Those worksheet parts are byte-identical between template and generated output.
- Openpyxl emits twice `Data Validation extension is not supported and will be removed` for both the template and generated workbook in normal loading mode. This is a warning about a future openpyxl save; this generation path patches the copied worksheet XML directly and did not remove those extensions.
- Since the same extensions and warnings exist in the template and the generated x14-validation parts are unchanged, the unsupported-extension warning is not supported as the cause of this output's repair prompt. This does not rule out separate Excel compatibility concerns with x14 validations.

## openpyxl findings and extension risk

- `load_workbook(..., read_only=True, data_only=False)` loads the generated workbook without warnings.
- Normal read-only-in-effect inspection with `read_only=False` also loads it, but emits twice: `Data Validation extension is not supported and will be removed`.
- The same two warnings occur for the original template and downloaded copy. The x14 data-validation extensions are still present in the generated ZIP and are byte-identical to the template because the production writer patches worksheet XML in a copied package rather than saving the workbook through openpyxl.
- Therefore openpyxl's unsupported-extension warning is a real risk if some future path opens and saves this workbook via openpyxl, but it does not explain loss of those extensions in this generated file, nor is it evidence that Excel caused the reported repair warning.

## Excel reproduction status and confidence

Excel is installed (`Office16/EXCEL.EXE`), but this execution session could not drive it interactively. Excel COM automation failed at `Workbooks.Open` with `Unable to get the Open property of the Workbooks class`; a separate `/x` launch created a process with no main-window handle. No recovery prompt or recovery log was captured. The attempt used only `diagnostic_copy.xlsx`, whose SHA-256 remains identical to the original output. The downloaded workbook and source template hashes also remain unchanged.

The shared-formula defect is confirmed by direct template/generated XML comparison and contradicts the shared-formula master semantics in Microsoft’s SpreadsheetML documentation. It is therefore the identified structural root cause candidate and the best-supported explanation for repair. Excel’s own repair log and an Excel open-without-repair confirmation remain unverified.

## Prior checks

The earlier `formula changes: 0` and `workbook structure preserved: Yes` checks were **not sufficient** to detect an Excel repair prompt. Package-membership and feature-count checks do not verify shared-formula master/follower consistency. The check counted formula text changes but did not ensure every surviving `t="shared"` formula references a master. It also cannot establish Excel compatibility by itself.

## Artifacts

- `workbook_package_analysis.json`: hashes, ZIP inventories, CRC results, relationship checks, XML checks, openpyxl warnings, extension markers, and template part comparison.
- `sheet2.xml.diff.txt`, `sheet3.xml.diff.txt`: byte-level unified diffs. Each very long worksheet is serialized on one line, so these diffs are large.
- `cell_xml_diff.json`: exact semantic cell-level changes for both worksheets.
- `shared_formula_audit.json`: shared-formula groups and the five generated orphan group indices.
- `diagnostic_copy.xlsx`: separate copy used for the attempted Excel reproduction; its hash matches the generated file.
- `excel_open_session.txt`, `open_in_excel.ps1`: Excel automation attempt and script; no repair log was produced.
- `analyze_workbooks.py`: reproducible read-only analysis script.

## Production changes

No production code or original workbook was changed. The smallest safe correction is to repair the five affected shared-formula groups while applying writes: move each master to the first remaining follower row (553), write the correctly row-adjusted formula there, and change the master `ref` to the surviving range ending at row 609, leaving followers 554:609 with the same `si`. This preserves carry-forward formulas in unused rows while retaining generated values in rows 546:552. An alternative is to expand the remaining followers into equivalent standalone formulas. Either correction must be tested on a diagnostic output and confirmed by opening it in Excel without a repair prompt before declaring the issue fixed.
