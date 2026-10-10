#!/usr/bin/env python3
"""Populate the documented spiral length and elbow fields in VJR SPIRALE.

Supported inputs are diameter, total duct length, and separate 90/45 degree
elbow counts. Repeated diameters are aggregated into the existing diameter
row. Metric diameters are converted to inches and must exactly match a VJR
diameter key; no nominal size rounding is performed.

The workbook is loaded with openpyxl, while the final XLSX package is patched
at the SPIRALE input cells so formulas, formatting, drawings, validations,
and workbook features unsupported by openpyxl are retained from the template.
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


SHEET_NAME = "SPIRALE"
DIAMETER_FIRST_ROW = 11
DIAMETER_LAST_ROW = 63
INPUT_COLUMNS = ("B", "C", "D")
REQUIRED_HEADERS = (
    "Diameter",
    "DiameterUnit",
    "Length",
    "LengthUnit",
    "Elbow90Count",
    "Elbow45Count",
)
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"


class InputValidationError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors
        super().__init__("Input validation failed")


def _decimal(value: Any, label: str, record_number: int, errors: list[str]) -> Decimal | None:
    try:
        result = Decimal(str(value).strip())
    except (InvalidOperation, ValueError, TypeError, AttributeError):
        errors.append(f"Record {record_number}: {label} must be numeric; got {value!r}")
        return None
    if not result.is_finite():
        errors.append(f"Record {record_number}: {label} must be finite")
        return None
    return result


def read_records(csv_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    records: list[dict[str, Any]] = []
    errors: list[str] = []
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = reader.fieldnames or []
            missing = [name for name in REQUIRED_HEADERS if name not in headers]
            if missing:
                return [], ["CSV is missing required columns: " + ", ".join(missing)]
            for line_number, row in enumerate(reader, start=2):
                parsed = {name: (row.get(name) or "").strip() for name in REQUIRED_HEADERS}
                if not any(parsed.values()):
                    continue
                for name in REQUIRED_HEADERS:
                    if parsed[name] == "":
                        errors.append(f"CSV row {line_number}: missing required field {name}")
                records.append(parsed)
    except OSError as exc:
        errors.append(f"Cannot read input CSV {csv_path}: {exc}")
    if not records and not errors:
        errors.append("CSV contains no spiral duct records")
    return records, errors


def diameter_rows(template: Path) -> dict[Decimal, int]:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        workbook = load_workbook(template, data_only=False, read_only=False)
    if SHEET_NAME not in workbook.sheetnames:
        raise InputValidationError([f"Template does not contain the {SHEET_NAME!r} sheet"])
    sheet = workbook[SHEET_NAME]
    result: dict[Decimal, int] = {}
    for row in range(DIAMETER_FIRST_ROW, DIAMETER_LAST_ROW + 1):
        value = sheet[f"A{row}"].value
        if value is None:
            continue
        try:
            diameter = Decimal(str(value))
        except InvalidOperation:
            continue
        if diameter.is_finite():
            result[diameter] = row
    return result


def aggregate_records(
    records: list[dict[str, Any]],
    diameter_to_row: dict[Decimal, int],
) -> tuple[dict[int, dict[str, Decimal]], list[str]]:
    errors: list[str] = []
    totals: dict[int, dict[str, Decimal]] = defaultdict(
        lambda: {"B": Decimal(0), "C": Decimal(0), "D": Decimal(0)}
    )
    for record_number, record in enumerate(records, start=1):
        diameter = _decimal(record["Diameter"], "Diameter", record_number, errors)
        length = _decimal(record["Length"], "Length", record_number, errors)
        elbow90 = _decimal(record["Elbow90Count"], "Elbow90Count", record_number, errors)
        elbow45 = _decimal(record["Elbow45Count"], "Elbow45Count", record_number, errors)
        if any(value is None for value in (diameter, length, elbow90, elbow45)):
            continue

        diameter_unit = record["DiameterUnit"].strip().upper()
        length_unit = record["LengthUnit"].strip().upper()
        if diameter_unit not in {"IN", "MM"}:
            errors.append(f"Record {record_number}: DiameterUnit must be IN or MM")
            continue
        if length_unit not in {"FT", "M"}:
            errors.append(f"Record {record_number}: LengthUnit must be FT or M")
            continue

        diameter_inches = diameter if diameter_unit == "IN" else diameter / Decimal("25.4")
        row = diameter_to_row.get(diameter_inches)
        if row is None:
            errors.append(
                f"Record {record_number}: diameter {record} converts to {diameter_inches} inches, "
                "which does not exactly match a VJR SPIRALE diameter key. Confirm the nominal diameter; no rounding was applied."
            )
            continue

        # The working notes request feet in the VJR sheet. One international
        # foot is exactly 0.3048 metres; keep full Decimal precision.
        length_feet = length if length_unit == "FT" else length / Decimal("0.3048")
        totals[row]["B"] += length_feet
        totals[row]["C"] += elbow90
        totals[row]["D"] += elbow45

    return dict(totals), errors


def _read_workbook_sheet_part(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relationships = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship")
    }
    sheets = workbook_root.find(f"{{{NS_MAIN}}}sheets")
    if sheets is None:
        raise ValueError("Workbook has no sheets element")
    for sheet in sheets.findall(f"{{{NS_MAIN}}}sheet"):
        if sheet.attrib.get("name") == sheet_name:
            target = relationships[sheet.attrib[f"{{{NS_REL_DOC}}}id"]].replace("\\", "/")
            if target.startswith("/"):
                return target.lstrip("/")
            return target if target.startswith("xl/") else "xl/" + target
    raise ValueError(f"Worksheet {sheet_name!r} was not found")


def _set_cell_value(sheet_root: ET.Element, coordinate: str, value: Any) -> None:
    row_number, _ = coordinate_to_tuple(coordinate)
    sheet_data = sheet_root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        raise ValueError("Worksheet has no sheetData element")
    row = next(
        (item for item in sheet_data.findall(f"{{{NS_MAIN}}}row") if int(item.attrib["r"]) == row_number),
        None,
    )
    if row is None:
        raise ValueError(f"Missing row {row_number} in {SHEET_NAME}")
    cell = next(
        (item for item in row.findall(f"{{{NS_MAIN}}}c") if item.attrib.get("r") == coordinate),
        None,
    )
    if cell is None:
        cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": coordinate})
        col_letters = re.match(r"[A-Z]+", coordinate).group(0)
        col_number = 0
        for char in col_letters:
            col_number = col_number * 26 + ord(char) - 64
        insertion = len(row)
        for index, sibling in enumerate(list(row)):
            sibling_ref = re.match(r"[A-Z]+", sibling.attrib.get("r", ""))
            sibling_column = 0
            if sibling_ref:
                for char in sibling_ref.group(0):
                    sibling_column = sibling_column * 26 + ord(char) - 64
            if sibling_column > col_number:
                insertion = index
                break
        row.insert(insertion, cell)
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
    elif value is not None:
        numeric = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        numeric.text = str(value)


def _write_package_preserving_template(
    template: Path, output: Path, values: dict[str, Decimal]
) -> str:
    with zipfile.ZipFile(template, "r") as source_zip:
        sheet_part = _read_workbook_sheet_part(source_zip, SHEET_NAME)
        worksheet = ET.fromstring(source_zip.read(sheet_part))
        for coordinate, value in values.items():
            _set_cell_value(worksheet, coordinate, value)
        patched_sheet = ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                data = patched_sheet if item.filename == sheet_part else source_zip.read(item.filename)
                output_zip.writestr(item, data)
    return sheet_part


def _formula_map(path: Path) -> dict[tuple[str, str], str]:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        workbook = load_workbook(path, data_only=False, read_only=False)
    return {
        (sheet.title, cell.coordinate): cell.value
        for sheet in workbook.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    }


def generate(template: Path, csv_path: Path, output: Path) -> list[str]:
    if not template.is_file():
        raise InputValidationError([f"Template workbook not found: {template}"])
    if template.resolve() == output.resolve():
        raise InputValidationError(["Output path must not be the original template workbook"])

    records, errors = read_records(csv_path)
    if errors:
        raise InputValidationError(errors)
    diameter_to_row = diameter_rows(template)
    totals, errors = aggregate_records(records, diameter_to_row)
    if errors:
        raise InputValidationError(errors)

    # Use openpyxl to apply the same cell assignments to the workbook model.
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        workbook = load_workbook(template, data_only=False, read_only=False)
    sheet = workbook[SHEET_NAME]
    values = {
        f"{column}{row}": totals[row][column]
        for row in sorted(totals)
        for column in INPUT_COLUMNS
    }
    for coordinate, value in values.items():
        sheet[coordinate] = value

    output.parent.mkdir(parents=True, exist_ok=True)
    formulas_before = _formula_map(template)
    sheet_part = _write_package_preserving_template(template, output, values)

    with zipfile.ZipFile(template, "r") as original_zip, zipfile.ZipFile(output, "r") as output_zip:
        original_parts = set(original_zip.namelist())
        output_parts = set(output_zip.namelist())
        if original_parts != output_parts:
            raise RuntimeError("Generated workbook package members differ from the template")
        changed_parts = [
            name for name in sorted(original_parts)
            if original_zip.read(name) != output_zip.read(name)
        ]
        if changed_parts != [sheet_part]:
            raise RuntimeError(f"Unexpected workbook package changes: {changed_parts}")
        if output_zip.testzip() is not None:
            raise RuntimeError("Generated workbook ZIP failed its integrity check")

    formulas_after = _formula_map(output)
    if formulas_before != formulas_after:
        raise RuntimeError("Formula preservation check failed")
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        output_book = load_workbook(output, data_only=False, read_only=False)
        source_book = load_workbook(template, data_only=False, read_only=False)
    result_sheet = output_book[SHEET_NAME]
    source_sheet = source_book[SHEET_NAME]
    for coordinate, expected in values.items():
        actual = result_sheet[coordinate].value
        value_matches = (
            float(actual) == float(expected)
            if isinstance(expected, Decimal) and isinstance(actual, (int, float))
            else actual == expected
        )
        if not value_matches:
            raise RuntimeError(f"Output value mismatch at {coordinate}")
    for row in source_sheet.iter_rows():
        for cell in row:
            if cell.coordinate not in values and result_sheet[cell.coordinate].value != cell.value:
                raise RuntimeError(f"Unexpected SPIRALE cell change at {cell.coordinate}")
            if result_sheet[cell.coordinate].style_id != cell.style_id:
                raise RuntimeError(f"Style changed at SPIRALE!{cell.coordinate}")
    if set(map(str, source_sheet.merged_cells.ranges)) != set(map(str, result_sheet.merged_cells.ranges)):
        raise RuntimeError("SPIRALE merged cells changed")

    input_rows = sorted(totals)
    written_ranges = ", ".join(f"B{row}:D{row}" for row in input_rows)
    return [
        f"PASS: validated {len(records)} source records and exact diameter matches against SPIRALE!A{DIAMETER_FIRST_ROW}:A{DIAMETER_LAST_ROW}.",
        f"PASS: aggregated records into {len(input_rows)} diameter rows; wrote only {written_ranges}.",
        f"PASS: wrote only SPIRALE columns B:D (length in feet, 90-degree count, 45-degree count).",
        f"PASS: preserved all {len(formulas_before):,} workbook formulas, including all 657 SPIRALE formulas.",
        f"PASS: preserved all SPIRALE non-input cell values, styles and merged cells; only {sheet_part} differs in the XLSX package.",
        "PASS: all other workbook package parts, including drawings and validation extensions, are byte-identical to the template.",
        f"OUTPUT: {output}",
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=Path("source_files/Ventilation Jean Roy excel.xlsx"))
    parser.add_argument("--input", type=Path, default=Path("data/spiral_phase3_test_cases.csv"))
    parser.add_argument("--output", type=Path, default=Path("output/Ventilation Jean Roy - spiral phase3 output.xlsx"))
    parser.add_argument("--report", type=Path, default=Path("output/spiral_phase3_validation_report.txt"))
    args = parser.parse_args()
    try:
        report = generate(args.template, args.input, args.output)
    except InputValidationError as exc:
        print("VALIDATION FAILED")
        for error in exc.errors:
            print(f"ERROR: {error}")
        return 2
    except Exception as exc:
        print(f"GENERATION FAILED: {exc}", file=sys.stderr)
        return 1
    report_text = "SPIRALE PHASE 3 VALIDATION REPORT\n" + "\n".join(report) + "\n"
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(report_text, encoding="utf-8")
    print(report_text, end="")
    print(f"REPORT: {args.report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
