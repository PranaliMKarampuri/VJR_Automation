#!/usr/bin/env python3
"""Populate documented round and standard flexible duct inputs in the VJR workbook.

Uses openpyxl to inspect and populate a workbook model, then patches only the
ROND ET FLEX worksheet XML in a copy of the original XLSX package. This retains
unsupported Excel extensions, drawings, validations, formulas, and other package
parts byte-for-byte.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import warnings
import zipfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple


SHEET_NAME = "ROND ET FLEX"
ROUND_FIRST_ROW, ROUND_LAST_ROW = 16, 24
FLEX_FIRST_ROW, FLEX_LAST_ROW = 4, 33
HEADERS = ("Category", "DiameterIn", "Quantity", "FlexType", "FlexBlock")
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"


class InputValidationError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors
        super().__init__("Input validation failed")


def _decimal(value: Any, label: str, record: int, errors: list[str]) -> Decimal | None:
    try:
        result = Decimal(str(value).strip())
    except (InvalidOperation, TypeError, ValueError, AttributeError):
        errors.append(f"Record {record}: {label} must be numeric; got {value!r}")
        return None
    if not result.is_finite():
        errors.append(f"Record {record}: {label} must be finite")
        return None
    return result


def read_records(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    records: list[dict[str, str]] = []
    errors: list[str] = []
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            missing_headers = [name for name in HEADERS if name not in (reader.fieldnames or [])]
            if missing_headers:
                return [], ["CSV is missing required columns: " + ", ".join(missing_headers)]
            for line, raw in enumerate(reader, start=2):
                record = {name: (raw.get(name) or "").strip() for name in HEADERS}
                if not any(record.values()):
                    continue
                category = record["Category"].upper()
                if category not in {"ROUND", "FLEX"}:
                    errors.append(f"CSV row {line}: Category must be ROUND or FLEX")
                for name in ("DiameterIn", "Quantity"):
                    if not record[name]:
                        errors.append(f"CSV row {line}: missing required field {name}")
                if category == "ROUND":
                    if record["FlexType"] or record["FlexBlock"]:
                        errors.append(f"CSV row {line}: round duct cannot specify FlexType or FlexBlock")
                elif category == "FLEX":
                    if not record["FlexType"]:
                        errors.append(f"CSV row {line}: missing required field FlexType")
                    if record["FlexBlock"].upper() not in {"FG", "JK"}:
                        errors.append(f"CSV row {line}: FlexBlock must explicitly be FG or JK")
                records.append(record)
    except OSError as exc:
        errors.append(f"Cannot read input CSV {path}: {exc}")
    if not records and not errors:
        errors.append("CSV contains no duct records")
    return records, errors


def _valid_selectors(template: Path) -> list[str]:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        workbook = load_workbook(template, data_only=False, read_only=False)
    # Formula lookup points to the existing VJR list, not a newly invented map.
    lookup_sheet = next((ws for ws in workbook.worksheets if ws.title.upper().startswith("DONN")), None)
    if lookup_sheet is None:
        raise InputValidationError(["Cannot find the existing DONNÉES lookup sheet"])
    return [str(lookup_sheet.cell(row, 1).value).strip() for row in range(67, 85)
            if lookup_sheet.cell(row, 1).value is not None]


def map_records(records: list[dict[str, str]], template: Path) -> tuple[dict[str, Decimal | str], list[str]]:
    errors: list[str] = []
    selectors = _valid_selectors(template)
    selector_lookup = {item.casefold(): item for item in selectors}
    round_keys: dict[Decimal, int] = {}
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        wb = load_workbook(template, data_only=False, read_only=False)
    if SHEET_NAME not in wb.sheetnames:
        raise InputValidationError([f"Template does not contain {SHEET_NAME!r}"])
    sheet = wb[SHEET_NAME]
    for row in range(ROUND_FIRST_ROW, ROUND_LAST_ROW + 1):
        try:
            key = Decimal(str(sheet[f"B{row}"].value))
        except (InvalidOperation, TypeError, ValueError):
            continue
        round_keys[key] = row

    round_totals: dict[int, Decimal] = defaultdict(Decimal)
    flex_totals: dict[str, dict[tuple[str, Decimal], Decimal]] = {"FG": {}, "JK": {}}
    flex_types: dict[str, str] = {}
    for idx, record in enumerate(records, start=1):
        category = record["Category"].upper()
        diameter = _decimal(record["DiameterIn"], "DiameterIn", idx, errors)
        quantity = _decimal(record["Quantity"], "Quantity", idx, errors)
        if diameter is None or quantity is None:
            continue
        if diameter <= 0 or quantity < 0:
            errors.append(f"Record {idx}: DiameterIn must be positive and Quantity cannot be negative")
            continue
        if category == "ROUND":
            row = round_keys.get(diameter)
            if row is None:
                errors.append(f"Record {idx}: {diameter} inch round diameter does not exactly match a VJR key in B16:B24; rejected without rounding")
                continue
            round_totals[row] += quantity  # Round Quantity is feet, as specified.
        elif category == "FLEX":
            block = record["FlexBlock"].upper()
            supplied = record["FlexType"]
            flex_type = selector_lookup.get(supplied.casefold())
            if flex_type is None:
                errors.append(f"Record {idx}: FlexType {supplied!r} is not an existing VJR dropdown value; rejected")
                continue
            # Only the two documented ordinary standard types are in scope.
            if flex_type.casefold() not in {"acoustique standard", "isolé standard"}:
                errors.append(f"Record {idx}: FlexType {flex_type!r} is not one of the explicitly supported ordinary standard flex types")
                continue
            previous = flex_types.get(block)
            if previous is not None and previous.casefold() != flex_type.casefold():
                errors.append(f"Record {idx}: FlexBlock {block} contains multiple types ({previous!r} and {flex_type!r}); one selector cell cannot represent both")
                continue
            flex_types[block] = flex_type
            key = (flex_type, diameter)
            flex_totals[block][key] = flex_totals[block].get(key, Decimal(0)) + quantity * Decimal(5)

    for block in ("FG", "JK"):
        if len(flex_totals[block]) > FLEX_LAST_ROW - FLEX_FIRST_ROW + 1:
            errors.append(f"FlexBlock {block}: aggregated rows exceed the available F{FLEX_FIRST_ROW}:G{FLEX_LAST_ROW} or J{FLEX_FIRST_ROW}:K{FLEX_LAST_ROW} input slots")

    values: dict[str, Decimal | str] = {}
    for row, total in round_totals.items():
        values[f"A{row}"] = total
    for block, selector_cell, col_qty, col_diameter in (("FG", "G2", "F", "G"), ("JK", "K2", "J", "K")):
        if block in flex_types:
            values[selector_cell] = flex_types[block]
        for offset, ((_, diameter), feet) in enumerate(sorted(flex_totals[block].items(), key=lambda item: (item[0][1], item[0][0].casefold()))):
            row = FLEX_FIRST_ROW + offset
            values[f"{col_qty}{row}"] = feet
            values[f"{col_diameter}{row}"] = diameter
    return values, errors


def _sheet_part(archive: zipfile.ZipFile) -> str:
    root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{{{NS_REL_PKG}}}Relationship")}
    for sheet in root.find(f"{{{NS_MAIN}}}sheets"):
        if sheet.attrib.get("name") == SHEET_NAME:
            target = rel_map[sheet.attrib[f"{{{NS_REL_DOC}}}id"]].replace("\\", "/")
            return target.lstrip("/") if target.startswith("/") else target if target.startswith("xl/") else "xl/" + target
    raise ValueError(f"Worksheet {SHEET_NAME!r} not found in package")


def _set_cell(root: ET.Element, coordinate: str, value: Decimal | str) -> None:
    row_number, _ = coordinate_to_tuple(coordinate)
    data = root.find(f"{{{NS_MAIN}}}sheetData")
    row = next((r for r in data.findall(f"{{{NS_MAIN}}}row") if int(r.attrib["r"]) == row_number), None)
    if row is None:
        raise ValueError(f"Input row {row_number} is absent from the template")
    cell = next((c for c in row.findall(f"{{{NS_MAIN}}}c") if c.attrib.get("r") == coordinate), None)
    if cell is None:
        raise ValueError(f"Input cell {coordinate} has no existing cell/style in the template")
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if isinstance(value, str):
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
        text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
        if value != value.strip():
            text.set(f"{{{NS_XML}}}space", "preserve")
        text.text = value
    else:
        v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        v.text = str(value)


def _formula_map(path: Path) -> dict[tuple[str, str], str]:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        wb = load_workbook(path, data_only=False, read_only=False)
    return {(ws.title, cell.coordinate): cell.value for ws in wb.worksheets for row in ws.iter_rows()
            for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")}


def generate(template: Path, input_csv: Path, output: Path) -> list[str]:
    if not template.is_file():
        raise InputValidationError([f"Template workbook not found: {template}"])
    if template.resolve() == output.resolve():
        raise InputValidationError(["Output path must not be the original template workbook"])
    records, errors = read_records(input_csv)
    if errors:
        raise InputValidationError(errors)
    values, errors = map_records(records, template)
    if errors:
        raise InputValidationError(errors)

    # Confirm assignments in an openpyxl model before producing the package.
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        model = load_workbook(template, data_only=False, read_only=False)
    for cell, value in values.items():
        model[SHEET_NAME][cell] = float(value) if isinstance(value, Decimal) else value

    before_formulas = _formula_map(template)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template, "r") as src:
        part = _sheet_part(src)
        sheet_root = ET.fromstring(src.read(part))
        for coordinate, value in values.items():
            _set_cell(sheet_root, coordinate, value)
        patched = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as dst:
            for member in src.infolist():
                dst.writestr(member, patched if member.filename == part else src.read(member.filename))

    with zipfile.ZipFile(template) as src, zipfile.ZipFile(output) as dst:
        if set(src.namelist()) != set(dst.namelist()):
            raise RuntimeError("XLSX package member list changed")
        changed = [name for name in sorted(src.namelist()) if src.read(name) != dst.read(name)]
        if changed != [part]:
            raise RuntimeError(f"Unexpected XLSX package changes: {changed}")
        if dst.testzip() is not None:
            raise RuntimeError("Output workbook ZIP integrity check failed")

    after_formulas = _formula_map(output)
    if before_formulas != after_formulas:
        raise RuntimeError("Workbook formulas changed")
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        source = load_workbook(template, data_only=False, read_only=False)
        result = load_workbook(output, data_only=False, read_only=False)
    ss, osheet = source[SHEET_NAME], result[SHEET_NAME]
    for coordinate, expected in values.items():
        actual = osheet[coordinate].value
        if isinstance(expected, Decimal):
            valid = actual is not None and Decimal(str(actual)) == expected
        else:
            valid = actual == expected
        if not valid:
            raise RuntimeError(f"Written value mismatch at {SHEET_NAME}!{coordinate}: {actual!r} != {expected!r}")
    for row in ss.iter_rows():
        for cell in row:
            if cell.coordinate not in values and osheet[cell.coordinate].value != cell.value:
                raise RuntimeError(f"Unexpected cell value change at {SHEET_NAME}!{cell.coordinate}")
            if cell.style_id != osheet[cell.coordinate].style_id:
                raise RuntimeError(f"Style changed at {SHEET_NAME}!{cell.coordinate}")
    if set(map(str, ss.merged_cells.ranges)) != set(map(str, osheet.merged_cells.ranges)):
        raise RuntimeError("Merged cells changed")
    if len(ss._images) != len(osheet._images):
        raise RuntimeError("Image count changed")

    written = ", ".join(f"{SHEET_NAME}!{c}={v}" for c, v in values.items())
    return [
        f"PASS: parsed and validated {len(records)} ROUND/FLEX records; unmatched round diameters, unsupported flex types, and ambiguous block types are rejected.",
        f"PASS: wrote only confirmed input cells: {written}.",
        f"PASS: preserved all {len(before_formulas)} workbook formulas, including all ROND ET FLEX formulas; formula map unchanged.",
        f"PASS: workbook opens with openpyxl; XLSX ZIP integrity check passed; only package part {part} differs from the template.",
        "PASS: all worksheet cell values outside the listed input cells and all styles/merged cells are unchanged; image count is unchanged.",
        "PASS: all package parts other than the ROND ET FLEX worksheet XML are byte-identical, including drawings and raw validation extensions.",
        f"OUTPUT: {output}",
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=Path("source_files/Ventilation Jean Roy excel.xlsx"))
    parser.add_argument("--input", type=Path, default=Path("data/rond_et_flex_phase4_test_cases.csv"))
    parser.add_argument("--output", type=Path, default=Path("output/Ventilation Jean Roy - rond et flex phase4 output.xlsx"))
    parser.add_argument("--report", type=Path, default=Path("output/rond_et_flex_phase4_validation_report.txt"))
    args = parser.parse_args()
    try:
        lines = generate(args.template, args.input, args.output)
    except InputValidationError as exc:
        print("VALIDATION FAILED")
        for error in exc.errors:
            print(f"ERROR: {error}")
        return 2
    except Exception as exc:
        print(f"GENERATION FAILED: {exc}", file=sys.stderr)
        return 1
    report = "ROND ET FLEX PHASE 4 VALIDATION REPORT\n" + "\n".join(lines) + "\n"
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(report, encoding="utf-8")
    print(report, end="")
    print(f"REPORT: {args.report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
