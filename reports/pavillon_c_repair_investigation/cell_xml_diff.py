from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).parent.parent.parent
OUT = Path(__file__).parent
TEMPLATE = ROOT / "source_files" / "Ventilation Jean Roy excel.xlsx"
GENERATED = ROOT / "output/jobs/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_8bc5caddc1/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_VJR.xlsx"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"}


def cell_map(root):
    return {c.attrib["r"]: c for c in root.findall(".//m:sheetData/m:row/m:c", NS)}


def serial(el):
    return ET.tostring(el, encoding="unicode") if el is not None else None


def record(cell):
    if cell is None:
        return None
    return {"attrs": dict(cell.attrib), "formula": cell.findtext("m:f", namespaces=NS), "formula_attrs": dict(cell.find("m:f", NS).attrib) if cell.find("m:f", NS) is not None else None,
            "value": cell.findtext("m:v", namespaces=NS), "inline_string": "".join(cell.find("m:is", NS).itertext()) if cell.find("m:is", NS) is not None else None}


def worksheet(z, part):
    data = z.read(part)
    root = ET.fromstring(data)
    cells = cell_map(root)
    rows = root.findall("m:sheetData/m:row", NS)
    row_map = {r.attrib["r"]: r for r in rows}
    root_start = data.find(b"<worksheet")
    root_end = data.find(b">", root_start) + 1 if root_start >= 0 else 0
    root_tag = data[root_start:root_end].decode("utf-8", "replace") if root_end else ""
    prefixes = dict(re.findall(r"xmlns:([\w.-]+)=['\"]([^'\"]+)", root_tag))
    ign = re.search(r"(?:\w+:)?Ignorable=['\"]([^'\"]*)", root_tag)
    ext = root.find("m:extLst", NS)
    child_order = [x.tag.rsplit("}", 1)[-1] for x in root]
    schema_order = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]
    ranks = [schema_order.index(name) if name in schema_order else 999 for name in child_order]
    invalid_row_cell_order = []
    for row in rows:
        refs = [c.attrib.get("r", "") for c in row.findall("m:c", NS)]
        cols = [sum((ord(ch.upper()) - 64) * (26 ** i) for i, ch in enumerate(ref.rstrip("0123456789")[::-1])) for ref in refs]
        if cols != sorted(cols) or len(cols) != len(set(cols)):
            invalid_row_cell_order.append({"row": row.attrib.get("r"), "refs": refs})
    return root, cells, row_map, {"root_start": root_tag, "namespace_prefixes": prefixes, "mc_ignorable": ign.group(1) if ign else None,
                                  "worksheet_extensions": serial(ext), "merge_count": len(root.findall("m:mergeCells/m:mergeCell", NS)),
                                  "standard_data_validation_count": len(root.findall("m:dataValidations/m:dataValidation", NS)),
                                  "x14_validation_count": len(root.findall(".//x14:dataValidations/x14:dataValidation", NS)),
                                  "worksheet_child_order": child_order, "worksheet_child_order_valid": ranks == sorted(ranks),
                                  "rows_with_duplicate_or_unsorted_cell_refs": invalid_row_cell_order}


result = {}
with zipfile.ZipFile(TEMPLATE) as t, zipfile.ZipFile(GENERATED) as g:
    for part in ("xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"):
        tr, tc, trm, tm = worksheet(t, part)
        gr, gc, grm, gm = worksheet(g, part)
        cell_diffs = []
        for ref in sorted(tc.keys() | gc.keys()):
            if serial(tc.get(ref)) != serial(gc.get(ref)):
                cell_diffs.append({"cell": ref, "template": record(tc.get(ref)), "generated": record(gc.get(ref))})
        row_diffs = []
        for ref in sorted(trm.keys() | grm.keys(), key=lambda x: int(x)):
            if serial(trm.get(ref)) != serial(grm.get(ref)):
                row_diffs.append({"row": ref, "template_attrs": dict(trm[ref].attrib) if ref in trm else None,
                                  "generated_attrs": dict(grm[ref].attrib) if ref in grm else None})
        result[part] = {"template_features": tm, "generated_features": gm, "template_cell_count": len(tc), "generated_cell_count": len(gc),
                        "changed_cell_count": len(cell_diffs), "changed_cells": cell_diffs, "changed_row_count": len(row_diffs), "changed_rows": row_diffs,
                        "root_element_namespace_and_attributes_equal": tm["namespace_prefixes"] == gm["namespace_prefixes"] and tm["mc_ignorable"] == gm["mc_ignorable"]}

out = OUT / "cell_xml_diff.json"
out.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
print(json.dumps({part: {k: v for k, v in val.items() if k not in ("changed_cells", "changed_rows")} | {"sample_changed_cells": val["changed_cells"][:30], "changed_cells_file": str(out)} for part, val in result.items()}, indent=2, ensure_ascii=False))
