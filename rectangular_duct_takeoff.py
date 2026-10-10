#!/usr/bin/env python3
"""Populate independent rectangular duct rows in a copy of the VJR template.

Phase 2 implements only rectangular duct rules stated in the working notes:
dimension order, HA/RE/RI installation, listed accessory codes and four-slot
accessory continuation rows. A dimension exactly 36 inches is rejected for
confirmation because the notes specify only less-than and greater-than.

The final workbook is built from the original XLSX package. openpyxl loads and
populates the workbook model; a small OOXML patch writes input values and
expands unaffected shared carry-down formulas where needed. This retains
unsupported Excel extensions and all other package parts exactly as they were
in the template.
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
import warnings
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple


SHEET_NAME = "TAKE OFF"
FIRST_DATA_ROW = 20
REQUIRED_COLUMNS = (
    "Width",
    "Height",
    "Length",
    "MetalCode",
    "Type",
    "Installation",
    "Units",
)
DUCT_COLUMNS = REQUIRED_COLUMNS[:-1]
BASE_COLUMNS = DUCT_COLUMNS + ("Units",)
EXCEL_INPUT_COLS = {
    "Width": "A",
    "Height": "B",
    "Length": "C",
    "MetalCode": "D",
    "Type": "E",
    "Installation": "F",
}
ACCESSORY_SLOT_COLUMNS = ("G", "H", "I", "J")
DOCUMENTED_ACCESSORY_CODES = {1, 21, 7, 71, 8, 81, 10, 100, 11, 111, 12, 13, 131, 14}
SEMANTIC_ACCESSORY_NAMES = {
    "CANEVAS": "CANEVAS",
    "CANVAS": "CANEVAS",
    "RCF": "FIRE DAMPER",
    "VCF": "FIRE DAMPER",
    "FIRE DAMPER": "FIRE DAMPER",
    "RCFF": "FIRE SMOKE DAMPER",
    "VCFF": "FIRE SMOKE DAMPER",
    "FIRE SMOKE DAMPER": "FIRE SMOKE DAMPER",
    "GRAVITY DAMPER": "GRAVITY DAMPER",
    "MBD": "MBD",
    "MVD": "MVD",
    "VB": "MBD",
    "MANUAL BALANCING DAMPER": "MBD",
    "MANUAL VOLUME BALANCING DAMPER": "MVD",
    "MESH": "MESH",
    "WIRE MESH": "MESH",
    "BIRD SCREEN": "MESH",
    "BIRDSCREEN": "MESH",
    "RM": "MOTORIZED DAMPER",
    "RMO": "MOTORIZED DAMPER",
    "VM": "MOTORIZED DAMPER",
    "MD": "MOTORIZED DAMPER",
    "MOTORIZED DAMPER": "MOTORIZED DAMPER",
    "MOTOR OPERATED DAMPER": "MOTORIZED DAMPER",
    "T VANE": "T VANE",
    "T VANES": "T VANE",
    "TURNING VANE": "T VANE",
    "TURNING VANES": "T VANE",
}
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"


class InputValidationError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors
        super().__init__("Input validation failed")


def parse_value(value: str | None) -> Any:
    """Parse numeric CSV fields as numbers and preserve other values as text."""
    if value is None:
        return None
    value = value.strip()
    if value == "":
        return None
    try:
        number = float(value)
    except ValueError:
        return value
    if not math.isfinite(number):
        return value
    return int(number) if number.is_integer() else number


def read_records(csv_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    errors: list[str] = []
    records: list[dict[str, Any]] = []
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = reader.fieldnames or []
            missing_headers = [name for name in BASE_COLUMNS if name not in headers]
            if missing_headers:
                errors.append("CSV is missing required columns: " + ", ".join(missing_headers))
                return [], errors
            accessory_columns: list[tuple[int, str]] = []
            for name in headers:
                match = re.fullmatch(r"Accessory(\d+)", name or "")
                if match:
                    accessory_columns.append((int(match.group(1)), name))
            accessory_columns.sort()
            if accessory_columns and [number for number, _ in accessory_columns] != list(range(1, len(accessory_columns) + 1)):
                errors.append("Accessory CSV columns must be consecutive, starting at Accessory1")
                return [], errors
            for csv_row, raw in enumerate(reader, start=2):
                record = {name: parse_value(raw.get(name)) for name in BASE_COLUMNS}
                record["Accessories"] = [parse_value(raw.get(name)) for _, name in accessory_columns]
                if all(value is None for key, value in record.items() if key != "Accessories") and not any(record["Accessories"]):
                    continue
                for name in DUCT_COLUMNS:
                    if record[name] is None:
                        errors.append(f"CSV row {csv_row}: missing required field {name}")
                if record["Units"] is None:
                    errors.append(f"CSV row {csv_row}: missing required field Units (M or A)")
                records.append(record)
    except OSError as exc:
        errors.append(f"Cannot read input CSV {csv_path}: {exc}")
    if not records and not errors:
        errors.append("CSV contains no duct records")
    return records, errors


def _as_decimal(value: Any, field: str, record_number: int, errors: list[str]) -> Decimal | None:
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        errors.append(f"Record {record_number}: {field} must be numeric to apply the documented duct rules")
        return None
    if not result.is_finite():
        errors.append(f"Record {record_number}: {field} must be a finite number")
        return None
    return result


def _dimension_inches(value: Decimal, units: str) -> Decimal:
    return value if units == "A" else value / Decimal("25.4")


def _accessory_name(value: str) -> str:
    return " ".join(re.findall(r"[A-Z0-9]+", value.upper()))


def map_accessory(value: Any, width: Decimal, height: Decimal, units: str) -> int:
    """Map only the rectangular accessory descriptions explicitly in notes."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        code_decimal = Decimal(str(value))
        if code_decimal != code_decimal.to_integral_value():
            raise ValueError(f"Accessory code must be an integer: {value!r}")
        code = int(code_decimal)
        if code in DOCUMENTED_ACCESSORY_CODES:
            return code
        raise ValueError(f"Accessory code {code} is not documented for this rectangular duct phase")

    raw_text = str(value).upper()
    normalized = _accessory_name(raw_text)
    if "ACOUSTIC INSULATION" in normalized:
        if re.search(r"1\s*/\s*2\s*(?:\"|INCH(?:ES)?)?\s*ACOUSTIC INSULATION", raw_text):
            return 1
        if re.search(r"\b1\s*(?:\"|INCH(?:ES)?)\s+ACOUSTIC INSULATION\b", raw_text):
            return 21
        raise ValueError("Acoustic insulation thickness must be explicitly 1/2 inch or 1 inch")

    name = SEMANTIC_ACCESSORY_NAMES.get(normalized)
    if name is None:
        raise ValueError(f"Accessory description {value!r} has no explicit rectangular mapping in the working notes")

    size_inches = max(_dimension_inches(width, units), _dimension_inches(height, units))
    over_30_inches = size_inches > Decimal("30")
    if name == "CANEVAS":
        return 81 if over_30_inches else 8
    if name == "FIRE DAMPER":
        return 10
    if name == "FIRE SMOKE DAMPER":
        return 100
    if name == "GRAVITY DAMPER":
        return 12
    if name in {"MBD", "MVD"}:
        return 131 if over_30_inches else 13
    if name == "MESH":
        return 14
    if name == "MOTORIZED DAMPER":
        return 111 if over_30_inches else 11
    if name == "T VANE":
        return 71 if over_30_inches else 7
    raise ValueError(f"Accessory description {value!r} has no explicit rectangular mapping in the working notes")


def transform_records(records: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str], str | None]:
    """Apply documented rectangular rules and expand accessory overflow rows."""
    errors: list[str] = []
    units_set = {str(record.get("Units") or "").strip().upper() for record in records}
    if any(unit not in {"M", "A"} for unit in units_set):
        errors.append("Units must be M (millimetres/metres) or A (inches/feet), as specified in the working notes")
    if len(units_set) > 1:
        errors.append("The VJR workbook has one project-wide unit selector (TAKE OFF!N4); all records must use the same Units value")
    units = next(iter(units_set)) if len(units_set) == 1 and next(iter(units_set)) in {"M", "A"} else None
    output_rows: list[dict[str, Any]] = []

    for record_number, record in enumerate(records, start=1):
        if units is None:
            continue
        width = _as_decimal(record.get("Width"), "Width", record_number, errors)
        height = _as_decimal(record.get("Height"), "Height", record_number, errors)
        if width is None or height is None:
            continue
        installation = str(record.get("Installation") or "").strip().upper()
        if installation not in {"HA", "RE", "RI"}:
            errors.append(f"Record {record_number}: Installation must be HA, RE or RI for the rules implemented in this phase")
            continue

        ordered_width, ordered_height = (width, height) if width >= height else (height, width)
        if installation == "RI":
            final_installation = "RI"
        else:
            inches = [
                _dimension_inches(width, units),
                _dimension_inches(height, units),
            ]
            if any(dimension > Decimal("36") for dimension in inches):
                final_installation = "RE"
            elif any(dimension == Decimal("36") for dimension in inches):
                errors.append(
                    f"Record {record_number}: a duct dimension is exactly 36 inches; the working notes specify less than 36 and greater than 36 only. Confirm the installation before generating."
                )
                continue
            else:
                final_installation = "HA"

        mapped_accessories: list[int] = []
        accessory_failed = False
        for accessory in record.get("Accessories", []):
            if accessory is None:
                continue
            try:
                mapped_accessories.append(map_accessory(accessory, width, height, units))
            except ValueError as exc:
                errors.append(f"Record {record_number}: {exc}")
                accessory_failed = True
        if accessory_failed:
            continue

        accessory_chunks = [mapped_accessories[i:i + 4] for i in range(0, len(mapped_accessories), 4)] or [[]]
        try:
            length = _as_decimal(record.get("Length"), "Length", record_number, errors)
        except Exception:
            length = None
        if length is None:
            continue
        for chunk_index, accessory_chunk in enumerate(accessory_chunks):
            values = {
                "Width": ordered_width,
                "Height": ordered_height,
                "Length": length if chunk_index == 0 else Decimal("0"),
                "MetalCode": record["MetalCode"],
                "Type": record["Type"],
                "Installation": final_installation,
                "Accessories": accessory_chunk,
            }
            output_rows.append(values)

    if len(output_rows) > 1004 - FIRST_DATA_ROW + 1:
        errors.append("Expanded input exceeds the available TAKE OFF rows 20:1004")
    return output_rows, errors, units


def _read_workbook_sheet_part(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship")
    }
    sheets = workbook_root.find(f"{{{NS_MAIN}}}sheets")
    if sheets is None:
        raise ValueError("Workbook has no sheets element")
    for sheet in sheets.findall(f"{{{NS_MAIN}}}sheet"):
        if sheet.attrib.get("name") == sheet_name:
            rel_id = sheet.attrib[f"{{{NS_REL_DOC}}}id"]
            target = rels[rel_id].replace("\\", "/")
            if target.startswith("/"):
                return target.lstrip("/")
            return "xl/" + target if not target.startswith("xl/") else target
    raise ValueError(f"Worksheet {sheet_name!r} was not found")


def _set_cell_value(sheet_root: ET.Element, coordinate: str, value: Any) -> None:
    row_num, _ = coordinate_to_tuple(coordinate)
    sheet_data = sheet_root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        raise ValueError("Worksheet has no sheetData element")
    row = next(
        (element for element in sheet_data.findall(f"{{{NS_MAIN}}}row") if int(element.attrib["r"]) == row_num),
        None,
    )
    if row is None:
        row = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_num)})
        rows = list(sheet_data)
        insertion = next((i for i, item in enumerate(rows) if int(item.attrib.get("r", "0")) > row_num), len(rows))
        sheet_data.insert(insertion, row)
    cell = next(
        (element for element in row.findall(f"{{{NS_MAIN}}}c") if element.attrib.get("r") == coordinate),
        None,
    )
    if cell is None:
        cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": coordinate})
        col_match = re.match(r"[A-Z]+", coordinate)
        col_name = col_match.group(0) if col_match else ""
        col_num = 0
        for char in col_name:
            col_num = col_num * 26 + ord(char) - 64
        children = list(row)
        insert_at = len(children)
        for i, sibling in enumerate(children):
            ref = sibling.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            sibling_col = 0
            if match:
                for char in match.group(1):
                    sibling_col = sibling_col * 26 + ord(char) - 64
            if sibling_col > col_num:
                insert_at = i
                break
        row.insert(insert_at, cell)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if isinstance(value, str):
        cell.set("t", "inlineStr")
        is_element = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
        text = ET.SubElement(is_element, f"{{{NS_MAIN}}}t")
        if value != value.strip():
            text.set(f"{{{NS_XML}}}space", "preserve")
        text.text = value
    elif value is not None:
        value_element = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        value_element.text = str(value)


def _set_formula(sheet_root: ET.Element, coordinate: str, formula: str) -> None:
    row_num, _ = coordinate_to_tuple(coordinate)
    sheet_data = sheet_root.find(f"{{{NS_MAIN}}}sheetData")
    row = next(
        (element for element in sheet_data.findall(f"{{{NS_MAIN}}}row") if int(element.attrib["r"]) == row_num),
        None,
    ) if sheet_data is not None else None
    if row is None:
        raise ValueError(f"Missing worksheet row for formula {coordinate}")
    cell = next(
        (element for element in row.findall(f"{{{NS_MAIN}}}c") if element.attrib.get("r") == coordinate),
        None,
    )
    if cell is None:
        raise ValueError(f"Missing worksheet cell for formula {coordinate}")
    for child in list(cell):
        if child.tag in (f"{{{NS_MAIN}}}f", f"{{{NS_MAIN}}}v", f"{{{NS_MAIN}}}is"):
            cell.remove(child)
    cell.attrib.pop("t", None)
    formula_element = ET.Element(f"{{{NS_MAIN}}}f")
    formula_element.text = formula[1:] if formula.startswith("=") else formula
    cell.insert(0, formula_element)
    cell.append(ET.Element(f"{{{NS_MAIN}}}v"))


def _write_preserving_package(
    template: Path,
    output: Path,
    values: dict[str, Any],
    formula_expansions: dict[str, str],
) -> str:
    with zipfile.ZipFile(template, "r") as source_zip:
        sheet_part = _read_workbook_sheet_part(source_zip, SHEET_NAME)
        root = ET.fromstring(source_zip.read(sheet_part))
        # Shared formula cells can depend on an input formula cell as their
        # group's master. Expand the unaffected A:B/D:F carry-down formulas
        # first so replacing supplied input rows cannot orphan their groups.
        for coordinate, formula in formula_expansions.items():
            _set_formula(root, coordinate, formula)
        for coordinate, value in values.items():
            _set_cell_value(root, coordinate, value)
        patched_sheet = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                data = patched_sheet if item.filename == sheet_part else source_zip.read(item.filename)
                output_zip.writestr(item, data)
    return sheet_part


def _formula_map(workbook_path: Path) -> dict[tuple[str, str], str]:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        book = load_workbook(workbook_path, data_only=False, read_only=False)
    return {
        (sheet.title, cell.coordinate): cell.value
        for sheet in book.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    }


def _protected_formula_map(
    formulas: dict[tuple[str, str], str], written_cells: set[str]
) -> dict[tuple[str, str], str]:
    """Exclude only input cells that intentionally replace carry-down formulas."""
    return {
        key: formula
        for key, formula in formulas.items()
        if not (key[0] == SHEET_NAME and key[1] in written_cells)
    }


def _takeoff_calculation_formulas(formulas: dict[tuple[str, str], str]) -> dict[tuple[str, str], str]:
    return {
        key: formula
        for key, formula in formulas.items()
        if key[0] == SHEET_NAME
        and 11 <= coordinate_to_tuple(key[1])[1] <= 50
    }


def generate(template: Path, csv_path: Path, output: Path) -> list[str]:
    errors: list[str] = []
    if not template.is_file():
        errors.append(f"Template workbook not found: {template}")
    records, csv_errors = read_records(csv_path)
    errors.extend(csv_errors)
    transformed_rows, rule_errors, units = transform_records(records)
    errors.extend(rule_errors)
    if errors:
        raise InputValidationError(errors)
    if template.resolve() == output.resolve():
        raise InputValidationError(["Output path must not be the original template workbook"])

    # Load and populate through openpyxl. The saved package below is patched
    # from the original OOXML to retain all unsupported template extensions.
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        workbook = load_workbook(template, data_only=False, read_only=False)
    if SHEET_NAME not in workbook.sheetnames:
        raise InputValidationError([f"Template does not contain a {SHEET_NAME!r} sheet"])
    sheet = workbook[SHEET_NAME]
    values: dict[str, Any] = {}
    for index, record in enumerate(transformed_rows):
        row = FIRST_DATA_ROW + index
        row_values = {
            "A": record["Width"],
            "B": record["Height"],
            "C": record["Length"],
            "D": record["MetalCode"],
            "E": record["Type"],
            "F": record["Installation"],
        }
        row_values.update({column: record["Accessories"][slot] if slot < len(record["Accessories"]) else None
                           for slot, column in enumerate(ACCESSORY_SLOT_COLUMNS)})
        for column, value in row_values.items():
            coordinate = f"{column}{row}"
            values[coordinate] = value
            sheet[coordinate] = value
    values["N4"] = units
    sheet["N4"] = units

    formula_expansions = {
        cell.coordinate: cell.value
        for row in workbook[SHEET_NAME].iter_rows(min_row=FIRST_DATA_ROW, max_row=1004, min_col=1, max_col=6)
        for cell in row
        if cell.column in (1, 2, 4, 5, 6)
        and cell.coordinate not in values
        and isinstance(cell.value, str)
        and cell.value.startswith("=")
    }

    output.parent.mkdir(parents=True, exist_ok=True)
    formula_before = _formula_map(template)
    formula_cells_written = set(values)
    sheet_part = _write_preserving_package(template, output, values, formula_expansions)

    # Validate the generated workbook and preservation guarantees.
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
            raise RuntimeError("Generated workbook ZIP package failed integrity check")
    formula_after = _formula_map(output)
    if _protected_formula_map(formula_before, formula_cells_written) != formula_after:
        protected = _protected_formula_map(formula_before, formula_cells_written)
        missing = [key for key in protected if key not in formula_after]
        extra = [key for key in formula_after if key not in protected]
        changed = [key for key in protected.keys() & formula_after.keys() if protected[key] != formula_after[key]]
        raise RuntimeError(
            "Formula preservation check failed outside the input cells: "
            f"missing={missing[:5]}, extra={extra[:5]}, changed={changed[:5]}"
        )
    calculation_formulas_before = _takeoff_calculation_formulas(formula_before)
    calculation_formulas_after = _takeoff_calculation_formulas(formula_after)
    if calculation_formulas_before != calculation_formulas_after:
        raise RuntimeError("TAKE OFF K:AX formula preservation check failed")

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        output_book = load_workbook(output, data_only=False, read_only=False)
    output_sheet = output_book[SHEET_NAME]
    for coordinate, expected in values.items():
        actual = output_sheet[coordinate].value
        if actual != expected:
            raise RuntimeError(f"Output value mismatch at {coordinate}: {actual!r} != {expected!r}")

    return [
        f"PASS: {len(records)} input records validated; required duct fields and M/A units are present.",
        f"PASS: dimension order normalized and HA/RE/RI installation rules applied; exact 36-inch non-riser cases are rejected for confirmation.",
        f"PASS: documented accessory descriptions/codes mapped; {len(transformed_rows) - len(records)} continuation rows added for accessory overflow.",
        f"PASS: wrote {len(transformed_rows)} rectangular duct rows to TAKE OFF!A{FIRST_DATA_ROW}:J{FIRST_DATA_ROW + len(transformed_rows) - 1}; selected unit {units} in N4.",
        f"PASS: wrote {len(values)} input/configuration cells; no cells in K:AX were written.",
        f"PASS: preserved all {len(calculation_formulas_before):,} TAKE OFF formulas in K:AX.",
        f"PASS: preserved all {len(formula_after):,} remaining formulas; only input cells replaced formulas, if present.",
        f"PASS: expanded {len(formula_expansions):,} unaffected A:B/D:F carry-down formulas so shared-formula groups remain valid.",
        f"PASS: all workbook ZIP members are byte-identical except {sheet_part}; original extensions, drawings, and other sheet parts are retained.",
        f"OUTPUT: {output}",
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=Path("source_files/Ventilation Jean Roy excel.xlsx"))
    parser.add_argument("--input", type=Path, default=Path("data/rectangular_duct_phase2_cases.csv"))
    parser.add_argument("--output", type=Path, default=Path("output/Ventilation Jean Roy - rectangular duct phase2 output.xlsx"))
    args = parser.parse_args()
    try:
        report = generate(args.template, args.input, args.output)
    except InputValidationError as exc:
        print("VALIDATION FAILED")
        for error in exc.errors:
            print(f"ERROR: {error}")
        return 2
    except Exception as exc:  # report generation failures clearly to a command-line user
        print(f"GENERATION FAILED: {exc}", file=sys.stderr)
        return 1
    print("VALIDATION REPORT")
    for line in report:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
