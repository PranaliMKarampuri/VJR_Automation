#!/usr/bin/env python3
"""Write explicitly supplied spiral balancing damper counts to SPIRALE!M.

Diameter conversion and exact VJR diameter-row matching are delegated to the
Phase 3 spiral aggregation function. This phase does not infer dampers from
branches, GRDs, VAVs, valves, or plan topology.
"""

from __future__ import annotations

import argparse
import csv
import sys
import warnings
import zipfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import load_workbook

import spiral_duct_takeoff as spiral


ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "source_files" / "Ventilation Jean Roy excel.xlsx"
DEFAULT_INPUT = ROOT / "data" / "spiral_damper_phase6a_test_cases.csv"
DEFAULT_OUTPUT = ROOT / "output" / "Ventilation Jean Roy - spiral damper phase6a output.xlsx"
DEFAULT_REPORT = ROOT / "output" / "spiral_damper_phase6a_validation_report.txt"
SHEET_NAME = spiral.SHEET_NAME
DIAMETER_COUNT_COLUMN = "M"
REQUIRED_HEADERS = ("Diameter", "DiameterUnit", "DamperCount")


class InputValidationError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors
        super().__init__("Input validation failed")


def read_records(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    records: list[dict[str, str]] = []
    errors: list[str] = []
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = reader.fieldnames or []
            missing_headers = [name for name in REQUIRED_HEADERS if name not in headers]
            if missing_headers:
                return [], ["CSV is missing required columns: " + ", ".join(missing_headers)]
            for line_number, raw in enumerate(reader, start=2):
                row = {name: (raw.get(name) or "").strip() for name in REQUIRED_HEADERS}
                if not any(row.values()):
                    continue
                for name, value in row.items():
                    if not value:
                        errors.append(f"CSV row {line_number}: missing required field {name}")
                if row["DiameterUnit"].upper() not in {"IN", "MM"}:
                    errors.append(f"CSV row {line_number}: DiameterUnit must be IN or MM")
                else:
                    row["DiameterUnit"] = row["DiameterUnit"].upper()
                try:
                    count = Decimal(row["DamperCount"])
                    if not count.is_finite() or count < 0 or count != count.to_integral_value():
                        raise InvalidOperation
                except (InvalidOperation, ValueError):
                    if row["DamperCount"]:
                        errors.append(f"CSV row {line_number}: DamperCount must be a non-negative whole number")
                records.append(row)
    except OSError as exc:
        errors.append(f"Cannot read input CSV {path}: {exc}")
    if not records and not errors:
        errors.append("CSV contains no spiral damper records")
    return records, errors


def build_values(records: list[dict[str, str]], template: Path) -> tuple[dict[str, Decimal], list[str]]:
    """Aggregate explicit counts by the exact Phase 3 diameter row."""
    errors: list[str] = []
    try:
        diameter_to_row = spiral.diameter_rows(template)
    except InputValidationError as exc:  # compatible if Phase 3 raises its own validation error
        return {}, list(exc.errors)
    except Exception as exc:
        return {}, [f"Could not read SPIRALE diameter keys: {exc}"]

    totals: dict[int, Decimal] = defaultdict(Decimal)
    for record_number, record in enumerate(records, start=1):
        count = Decimal(record["DamperCount"])
        # Route diameter parsing, metric conversion, and exact-key matching
        # through the established Phase 3 logic using neutral zero-valued
        # length/elbow fields. Only the resulting row key is used here.
        match_record = {
            "Diameter": record["Diameter"],
            "DiameterUnit": record["DiameterUnit"],
            "Length": "0",
            "LengthUnit": "FT",
            "Elbow90Count": "0",
            "Elbow45Count": "0",
        }
        matched, match_errors = spiral.aggregate_records([match_record], diameter_to_row)
        if match_errors:
            errors.extend(f"Record {record_number}: {error.split(': ', 1)[-1]}" for error in match_errors)
            continue
        row = next(iter(matched))
        totals[row] += count
    return {f"{DIAMETER_COUNT_COLUMN}{row}": count for row, count in sorted(totals.items())}, errors


def generate(template: Path, input_csv: Path, output: Path) -> list[str]:
    if not template.is_file():
        raise InputValidationError([f"Template workbook not found: {template}"])
    if template.resolve() == output.resolve():
        raise InputValidationError(["Output path must not be the original template workbook"])
    records, errors = read_records(input_csv)
    if errors:
        raise InputValidationError(errors)
    values, errors = build_values(records, template)
    if errors:
        raise InputValidationError(errors)

    formulas_before = spiral._formula_map(template)
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet_part = spiral._write_package_preserving_template(template, output, values)

    with zipfile.ZipFile(template, "r") as source_zip, zipfile.ZipFile(output, "r") as output_zip:
        if set(source_zip.namelist()) != set(output_zip.namelist()):
            raise RuntimeError("Output workbook package members differ from the template")
        changed_parts = [name for name in sorted(source_zip.namelist())
                         if source_zip.read(name) != output_zip.read(name)]
        if changed_parts != [sheet_part]:
            raise RuntimeError(f"Unexpected workbook package changes: {changed_parts}")
        if output_zip.testzip() is not None:
            raise RuntimeError("Generated workbook ZIP failed its integrity check")
        source_root = ET.fromstring(source_zip.read(sheet_part))
        output_root = ET.fromstring(output_zip.read(sheet_part))
        for tag in ("dataValidations", "extLst"):
            source_node = source_root.find(f"{{{spiral.NS_MAIN}}}{tag}")
            output_node = output_root.find(f"{{{spiral.NS_MAIN}}}{tag}")
            source_xml = ET.tostring(source_node, encoding="utf-8") if source_node is not None else None
            output_xml = ET.tostring(output_node, encoding="utf-8") if output_node is not None else None
            if source_xml != output_xml:
                raise RuntimeError(f"SPIRALE {tag} XML changed")

    formulas_after = spiral._formula_map(output)
    if formulas_before != formulas_after:
        raise RuntimeError("Formula preservation check failed")
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        source_book = load_workbook(template, data_only=False, read_only=False)
        output_book = load_workbook(output, data_only=False, read_only=False)
    if source_book.sheetnames != output_book.sheetnames:
        raise RuntimeError("Worksheet structure changed")
    source_sheet, output_sheet = source_book[SHEET_NAME], output_book[SHEET_NAME]
    for coordinate, expected in values.items():
        actual = output_sheet[coordinate].value
        if Decimal(str(actual)) != expected:
            raise RuntimeError(f"Output value mismatch at {SHEET_NAME}!{coordinate}")
    for row in source_sheet.iter_rows():
        for cell in row:
            if cell.coordinate not in values and output_sheet[cell.coordinate].value != cell.value:
                raise RuntimeError(f"Unexpected SPIRALE cell change at {cell.coordinate}")
            if cell.style_id != output_sheet[cell.coordinate].style_id:
                raise RuntimeError(f"Style changed at SPIRALE!{cell.coordinate}")
    if set(map(str, source_sheet.merged_cells.ranges)) != set(map(str, output_sheet.merged_cells.ranges)):
        raise RuntimeError("SPIRALE merged cells changed")
    if len(source_sheet._images) != len(output_sheet._images):
        raise RuntimeError("SPIRALE image count changed")

    cells = ", ".join(f"{cell}={value}" for cell, value in values.items()) or "none"
    return [
        f"PASS: validated {len(records)} explicit spiral damper records using Phase 3 exact diameter matching against SPIRALE!A{spiral.DIAMETER_FIRST_ROW}:A{spiral.DIAMETER_LAST_ROW}.",
        f"PASS: aggregated repeated diameters; wrote only SPIRALE!M input cells: {cells}.",
        "PASS: no branch, GRD, VAV/valve, or drawing-topology inference was applied.",
        f"PASS: preserved all {len(formulas_before):,} workbook formulas; formula recalculation was not performed.",
        f"PASS: preserved sheet names, SPIRALE non-input values, styles, merged cells and image count; only {sheet_part} differs in the XLSX package.",
        "PASS: all other workbook package parts, including drawing and validation-extension parts, are byte-identical to the template.",
        f"OUTPUT: {output}",
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
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
    report = "SPIRAL DAMPER PHASE 6A VALIDATION REPORT\n" + "\n".join(lines) + "\n"
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(report, encoding="utf-8")
    print(report, end="")
    print(f"REPORT: {args.report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
