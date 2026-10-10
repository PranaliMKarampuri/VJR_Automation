"""Manual Team to VJR writer.

For the confirmed Kuujjuaq project, no manifest is required: direct blank input
cells are populated from the clean VJR template and unsupported fields are
listed individually. Other jobs may use the explicit mapping manifest path.
Source workbooks are read-only; output is always a separate workbook.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.formula.translate import Translator

import manual_to_vjr_dry_run as dryrun
import vjr_ppt_rules as ppt


ROOT = Path(__file__).resolve().parent
TEST_INPUT_DIR = ROOT / "test_input"
OUTPUT_DIR = ROOT / "output"
DEFAULT_TEMPLATE = ROOT / "source_files" / "Ventilation Jean Roy excel.xlsx"
PROJECT_DRY_RUN = ROOT / "reports" / "kuujjuaq_manual_to_vjr_dry_run.csv"
STATUSES = ("SAFE_TO_AUTOMATE", "NEEDS_CONFIRMATION", "UNSUPPORTED")
MANIFEST_HEADERS = (
    "manifest_scope", "manual_source_row", "estimator_decision", "estimator_confirmed_by",
    "estimator_confirmation_date", "ppt_rule", "target_sheet", "target_row",
    "dimension_unit", "write_values_json",
)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL_DOC)


class WriterError(ValueError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def discover_inputs(folder: Path = TEST_INPUT_DIR) -> list[Path]:
    return sorted((p for p in Path(folder).iterdir()
                   if p.is_file() and p.suffix.casefold() in {".xlsx", ".csv"}
                   and not p.name.startswith("~$")
                   and not p.stem.casefold().endswith("_vjr")
                   and "template" not in p.stem.casefold()), key=lambda p: p.name.casefold())


def resolve_manual(argument: str | Path | None, folder: Path = TEST_INPUT_DIR) -> Path:
    if argument is None:
        choices = discover_inputs(folder)
        if len(choices) != 1:
            raise WriterError(f"Expected exactly one .xlsx in {folder}; found {len(choices)}.")
        return choices[0].resolve()
    path = Path(argument)
    if not path.is_file() and not path.is_absolute():
        path = Path(folder) / path
    if not path.is_file() or path.suffix.casefold() not in {".xlsx", ".csv"}:
        raise WriterError(f"Manual input not found or not .xlsx/.csv: {path}")
    return path.resolve()


def _cell_header_map(ws: Any) -> dict[str, int]:
    headers: dict[str, int] = {}
    for index, value in enumerate(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)), 1):
        if value is not None:
            headers[str(value).strip().casefold()] = index
    return headers


def inspect_manual(path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    if Path(path).suffix.casefold() == ".csv":
        from manual_csv_to_vjr import read_manual_csv
        return read_manual_csv(Path(path))
    book = openpyxl.load_workbook(path, read_only=True, data_only=False)
    problems: list[str] = []
    if "DATA" not in book.sheetnames:
        book.close()
        raise WriterError("Manual workbook is missing required DATA sheet.")
    ws = book["DATA"]
    header = _cell_header_map(ws)
    required = ("subject", "length", "length unit", "count", "label", "duct size", "type",
                "accessories1", "accessories2", "remark", "units")
    missing = [name for name in required if name not in header]
    if missing:
        book.close()
        raise WriterError("Manual DATA sheet is missing required columns: " + ", ".join(missing))
    records: list[dict[str, Any]] = []
    for row_no, values in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not any(value is not None for value in values):
            continue
        record = {name: values[index - 1] if index - 1 < len(values) else None for name, index in header.items()}
        record["source_row"] = row_no
        record["category"] = dryrun.classify(record.get("subject"), record.get("label"))
        records.append(record)
    book.close()
    if not records:
        problems.append("Manual DATA contains no records.")
    return records, problems


def _is_placeholder(value: Any) -> bool:
    return value is None or str(value).strip().upper() in {"", ".", "-", "N/A", "NA"}


def _base_statuses(manual_path: Path, records: list[dict[str, Any]]) -> dict[int, tuple[str, str]]:
    # Keep the reviewed Kuujjuaq dispositions exactly as they were generated.
    if manual_path.resolve() == dryrun.MANUAL_DEFAULT.resolve() and PROJECT_DRY_RUN.exists():
        with PROJECT_DRY_RUN.open("r", newline="", encoding="utf-8-sig") as stream:
            old = {int(row["manual_source_row"]): row for row in csv.DictReader(stream)}
        if set(old) == {record["source_row"] for record in records}:
            return {key: (row["mapping_status"], row["reason"]) for key, row in old.items()}
    result: dict[int, tuple[str, str]] = {}
    for record in records:
        category = record["category"]
        if category == "other/unknown":
            result[record["source_row"]] = ("UNSUPPORTED", "No documented PPT mapping for this Manual item.")
        else:
            result[record["source_row"]] = ("NEEDS_CONFIRMATION", "No approved exact VJR destination and cell mapping was supplied.")
    return result


def _read_manifest(path: Path | None) -> list[dict[str, str]]:
    if path is None:
        return []
    with Path(path).open("r", newline="", encoding="utf-8-sig") as stream:
        reader = csv.DictReader(stream)
        missing = [name for name in MANIFEST_HEADERS if name not in (reader.fieldnames or [])]
        if missing:
            raise WriterError("Approved mapping manifest is missing columns: " + ", ".join(missing))
        return list(reader)


def _xml_sheet_parts(book: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(book.read("xl/workbook.xml"))
    rel_root = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    targets = {item.attrib["Id"]: item.attrib["Target"] for item in rel_root.findall(f"{{{NS_REL_PKG}}}Relationship")}
    result: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{NS_MAIN}}}sheet"):
        rel_id = sheet.attrib[f"{{{NS_REL_DOC}}}id"]
        target = targets[rel_id].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        result[sheet.attrib["name"]] = target
    return result


def _rectangular_rule(record: dict[str, Any], entry: dict[str, str], values: dict[str, Any]) -> tuple[bool, str]:
    match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*", str(record.get("duct size") or ""))
    if not match:
        return False, "Manual rectangular size is not an explicit width x height pair."
    unit = entry.get("dimension_unit", "").strip().lower()
    if unit not in {"mm", "inch", "in", "inches"}:
        return False, "Dimension unit is not explicitly confirmed in the approved mapping."
    width, height = match.groups()
    ordered = (max(float(width), float(height)), min(float(width), float(height)))
    row = int(entry["target_row"])
    required = {f"A{row}", f"B{row}", f"C{row}", f"D{row}", f"E{row}", f"F{row}"}
    if not required.issubset(values):
        return False, "Approved rectangular write set must explicitly name width, height, length, metal, type, and installation cells A:F."
    if float(values[f"A{row}"]) != ordered[0] or float(values[f"B{row}"]) != ordered[1]:
        return False, "PPT rectangular rule requires the larger dimension first; manifest values do not match the Manual size."
    try:
        if float(values[f"C{row}"]) != float(record["length"]):
            return False, "Target length does not equal the explicit Manual DATA length."
        # The PPT gauge table has specified band edges; values in its gaps are unresolved.
        largest = max(ordered)
        if unit == "mm":
            bands = ((305, 26), (750, 24), (1500, 22), (2100, 20), (float("inf"), 18))
            lower_edges = (0, 310, 755, 1505, 2105)
        else:
            bands = ((12, 26), (30, 24), (59, 22), (82, 20), (float("inf"), 18))
            lower_edges = (0, 13, 31, 60, 83)
        gauge = next((g for (upper, g), lower in zip(bands, lower_edges) if lower <= largest <= upper), None)
        if gauge is None or int(float(values[f"D{row}"])) != gauge:
            return False, "Metal gauge is missing or does not match a PPT table range; gaps remain unresolved."
    except (TypeError, ValueError, KeyError):
        return False, "Required numeric rectangular values are missing or invalid."
    pressure_text = " ".join(str(record.get(key) or "") for key in ("type", "remark")).upper()
    pressure = next((code for code in ("BP", "MP", "HP") if re.search(rf"\b{code}\b", pressure_text)), None)
    if pressure is None or str(values[f"E{row}"]).strip().upper() != pressure:
        return False, "PPT duct type BP/MP/HP is not explicitly established by the Manual customer notes."
    install = ppt.installation(width, height, unit, is_riser=entry.get("ppt_rule") == "rectangular_riser")
    if install.status != ppt.SAFE or str(values[f"F{row}"]).strip().upper() != install.value:
        return False, install.reason if install.status != ppt.SAFE else "Installation value does not match the PPT size/riser rule."
    try:
        accessory_codes = [int(float(values[f"{column}{row}"])) for column in "GHIJ" if f"{column}{row}" in values]
    except (TypeError, ValueError):
        return False, "Accessory values must be numeric PPT codes."
    if any(code not in ppt.ACCESSORY_RULES for code in accessory_codes):
        return False, "Manifest contains an accessory code not documented in the PPT accessory table."
    source_accessories = [str(record.get(key) or "").strip().upper() for key in ("accessories1", "accessories2")]
    expected_codes: list[int] = []
    for accessory_text in source_accessories:
        if _is_placeholder(accessory_text):
            continue
        if "ACOUSTIC" in accessory_text and ("1/2" in accessory_text or "1/2\"" in accessory_text or "½" in accessory_text):
            expected_codes.append(1)
        elif "ACOUSTIC" in accessory_text and ("1\"" in accessory_text or "1 INCH" in accessory_text):
            expected_codes.append(21)
        else:
            return False, "Manual accessory description has no unambiguous PPT code in this writer version."
    if accessory_codes != expected_codes:
        return False, "VJR accessory codes must match the explicitly attached Manual accessory descriptions."
    return True, "Rectangular size order, length, gauge table, BP/MP/HP type, installation, and accessory codes validated against the PPT."


def _manifest_rule_check(entry: dict[str, str], records_by_row: dict[int, dict[str, Any]]) -> tuple[bool, str, dict[str, Any]]:
    try:
        source_row = int(entry["manual_source_row"])
        record = records_by_row[source_row]
        row = int(entry["target_row"])
        values = json.loads(entry["write_values_json"])
        if not isinstance(values, dict) or not values:
            raise ValueError("write_values_json must be a non-empty cell-to-value object")
    except (ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        return False, f"Invalid manifest record: {exc}", {}
    if entry.get("estimator_decision", "").strip().upper() != "SAFE_TO_AUTOMATE":
        return False, "Manifest is not explicitly approved as SAFE_TO_AUTOMATE.", values
    if not entry.get("estimator_confirmed_by", "").strip() or not entry.get("estimator_confirmation_date", "").strip():
        return False, "Estimator name and confirmation date are required for an approved write manifest.", values
    if entry.get("target_sheet") != "TAKE OFF":
        return False, "This writer version permits only confirmed TAKE OFF rectangular input rows.", values
    rule = entry.get("ppt_rule", "").strip()
    if rule != "rectangular":
        return False, "This writer version has no complete, validated writer for the stated PPT rule.", values
    if record["category"] != "rectangular":
        return False, "Manifest PPT rule does not match the Manual source record category.", values
    if row < 20 or row > 1004 or any(not re.fullmatch(rf"[A-J]{row}", str(cell)) for cell in values):
        return False, "Only explicitly mapped TAKE OFF input cells A:J on rows 20:1004 are permitted.", values
    ok, reason = _rectangular_rule(record, entry, values)
    return ok, reason, values


def _check_target_cells(template: Path, mappings: list[tuple[int, str, int, dict[str, Any]]],
                        units_by_source_row: dict[int, str]) -> tuple[dict[str, dict[str, Any]], list[str]]:
    book = openpyxl.load_workbook(template, data_only=False)
    problems: list[str] = []
    writes: dict[str, dict[str, Any]] = {}
    merged_lookup: dict[str, set[str]] = {}
    for sheet_name in {sheet for _, sheet, _, _ in mappings}:
        if sheet_name not in book.sheetnames:
            problems.append(f"VJR is missing target sheet {sheet_name!r}.")
            continue
        ws = book[sheet_name]
        merged_lookup[sheet_name] = {str(rng) for rng in ws.merged_cells.ranges}
    for source_row, sheet_name, target_row, values in mappings:
        if sheet_name not in book.sheetnames:
            continue
        ws = book[sheet_name]
        if sheet_name == "TAKE OFF":
            required_selector = "M" if units_by_source_row.get(source_row) == "mm" else "A"
            if str(ws["N4"].value or "").strip().upper() != required_selector:
                problems.append(f"Manual row {source_row}: TAKE OFF!N4 does not already match the PPT dimension/length unit selector {required_selector}; selector writes are not approved.")
        for coordinate, value in values.items():
            cell = ws[coordinate]
            if cell.coordinate in ws.merged_cells:
                problems.append(f"Manual row {source_row}: {sheet_name}!{coordinate} is part of a merged range.")
                continue
            if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
                problems.append(f"Manual row {source_row}: {sheet_name}!{coordinate} is a formula; write blocked.")
                continue
            if cell.value not in (None, ""):
                problems.append(f"Manual row {source_row}: {sheet_name}!{coordinate} already has a value; overwrite blocked.")
                continue
            if cell.style_id == 0:
                problems.append(f"Manual row {source_row}: {sheet_name}!{coordinate} has no existing input-cell formatting; write blocked.")
                continue
            writes.setdefault(sheet_name, {})[coordinate] = value
    book.close()
    return writes, problems


def _replace_cell(root: ET.Element, coordinate: str, value: Any) -> None:
    cell = root.find(f".//{{{NS_MAIN}}}c[@r='{coordinate}']")
    if cell is None:
        raise WriterError(f"Designated styled input cell {coordinate} is absent from worksheet XML.")
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if value is None:
        return
    if isinstance(value, bool):
        cell.set("t", "b")
        ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = str(value)
    elif isinstance(value, (date, datetime)):
        raise WriterError("Dates are not supported as VJR cell values in this manifest format.")
    else:
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
        text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
        text.text = str(value)


def _cell_map(root: ET.Element) -> dict[str, ET.Element]:
    return {cell.attrib["r"]: cell for cell in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row/{{{NS_MAIN}}}c")}


def _shared_formula_groups(root: ET.Element) -> dict[str, dict[str, Any]]:
    """Collect shared formulas by worksheet-local shared index."""
    groups: dict[str, dict[str, Any]] = {}
    for coordinate, cell in _cell_map(root).items():
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if formula is None or formula.attrib.get("t") != "shared":
            continue
        index = formula.attrib.get("si")
        if index is None:
            raise WriterError(f"Shared formula at {coordinate} has no shared index.")
        group = groups.setdefault(index, {"master": None, "members": []})
        group["members"].append((coordinate, cell, formula))
        if formula.attrib.get("ref") is not None:
            if group["master"] is not None:
                raise WriterError(f"Shared formula index {index} has more than one master.")
            if not formula.text:
                raise WriterError(f"Shared formula master at {coordinate} has no formula text.")
            group["master"] = (coordinate, formula)
    return groups


def _materialize_touched_shared_formulas(root: ET.Element, writes: dict[str, Any]) -> int:
    """Expand surviving members of any shared-formula group touched by writes.

    A write may replace the shared formula master as well as followers. Expanding
    the untouched members to row/column-adjusted normal formulas preserves their
    calculation behavior without leaving references to a deleted master.
    """
    groups = _shared_formula_groups(root)
    plans: list[tuple[ET.Element, str, dict[str, str], str]] = []
    for index, group in groups.items():
        members = group["members"]
        member_coords = {coordinate for coordinate, _, _ in members}
        if not member_coords.intersection(writes):
            continue
        master = group["master"]
        if master is None:
            raise WriterError(f"Input write intersects shared formula group {index}, which has no master.")
        origin, master_formula = master
        formula_text = master_formula.text
        master_formula_attrs = {key: value for key, value in master_formula.attrib.items()
                                if key not in {"t", "si", "ref"}}
        for coordinate, cell, formula in members:
            if coordinate in writes:
                continue
            try:
                translated = Translator("=" + formula_text, origin=origin).translate_formula(coordinate)
            except Exception as exc:
                raise WriterError(f"Cannot safely preserve shared formula {index} at {coordinate}: {exc}") from exc
            plans.append((cell, translated[1:] if translated.startswith("=") else translated,
                          master_formula_attrs, coordinate))
    # Calculate all translations before changing any tree nodes so a failure is
    # atomic for the caller.
    for cell, translated, attrs, coordinate in plans:
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if formula is None:
            raise WriterError(f"Shared formula cell {coordinate} disappeared during write preparation.")
        formula.attrib.clear()
        formula.attrib.update(attrs)
        formula.text = translated
    return len(plans)


def _shared_formula_problems(root: ET.Element) -> list[str]:
    """Return shared-formula structural inconsistencies in one worksheet."""
    problems: list[str] = []
    for index, group in _shared_formula_groups(root).items():
        master = group["master"]
        if master is None:
            problems.append(f"Shared formula group {index} has followers but no master.")
            continue
        master_coordinate, master_formula = master
        ref = master_formula.attrib.get("ref")
        if not ref:
            problems.append(f"Shared formula master {master_coordinate} has no range.")
            continue
        try:
            min_col, min_row, max_col, max_row = openpyxl.utils.cell.range_boundaries(ref)
            master_row, master_col = openpyxl.utils.cell.coordinate_to_tuple(master_coordinate)
            if not (min_row <= master_row <= max_row and min_col <= master_col <= max_col):
                problems.append(f"Shared formula master {master_coordinate} is outside its range {ref}.")
            for coordinate, _, _ in group["members"]:
                row, col = openpyxl.utils.cell.coordinate_to_tuple(coordinate)
                if not (min_row <= row <= max_row and min_col <= col <= max_col):
                    problems.append(f"Shared formula member {coordinate} is outside master range {ref} (si={index}).")
        except ValueError as exc:
            problems.append(f"Shared formula master {master_coordinate} has invalid range {ref!r}: {exc}.")
    return problems


def audit_shared_formula_package(path: Path) -> list[str]:
    """Audit shared-formula masters, followers, and ranges across a workbook."""
    problems: list[str] = []
    with zipfile.ZipFile(path, "r") as package:
        parts = _xml_sheet_parts(package)
        for sheet, part in parts.items():
            root = ET.fromstring(package.read(part))
            problems.extend(f"{sheet}: {problem}" for problem in _shared_formula_problems(root))
    return problems


def _register_xml_namespaces(data: bytes) -> None:
    """Keep prefixes stable when ElementTree rewrites a template worksheet.

    Prefixes referenced as values (notably mc:Ignorable) are significant to
    Office markup compatibility and are not repaired automatically by XML
    serializers when they rename namespace prefixes.
    """
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(data), events=("start-ns",)):
        if prefix and not re.fullmatch(r"ns\d+", prefix):
            ET.register_namespace(prefix, uri)


def _patch_copy(template: Path, output: Path, writes: dict[str, dict[str, Any]]) -> set[str]:
    output.parent.mkdir(parents=True, exist_ok=True)
    if not writes:
        shutil.copy2(template, output)
        problems = audit_shared_formula_package(output)
        if problems:
            raise WriterError("Copied workbook shared-formula audit failed: " + "; ".join(problems))
        return set()
    with zipfile.ZipFile(template, "r") as source:
        sheet_parts = _xml_sheet_parts(source)
        changed_parts = {sheet_parts[name] for name in writes}
        temp_path = output.with_suffix(output.suffix + ".tmp")
        with zipfile.ZipFile(temp_path, "w") as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename in changed_parts:
                    _register_xml_namespaces(data)
                    root = ET.fromstring(data)
                    sheet_name = next(name for name, part in sheet_parts.items() if part == info.filename)
                    sheet_writes = writes[sheet_name]
                    _materialize_touched_shared_formulas(root, sheet_writes)
                    for coordinate, value in sheet_writes.items():
                        _replace_cell(root, coordinate, value)
                    problems = _shared_formula_problems(root)
                    if problems:
                        raise WriterError(f"Shared-formula audit failed on {sheet_name}: " + "; ".join(problems))
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                target.writestr(info, data)
    problems = audit_shared_formula_package(temp_path)
    if problems:
        temp_path.unlink(missing_ok=True)
        raise WriterError("Generated workbook shared-formula audit failed: " + "; ".join(problems))
    temp_path.replace(output)
    return changed_parts


def _formula_map(book: Any) -> dict[tuple[str, str], str]:
    formulas: dict[tuple[str, str], str] = {}
    for ws in book.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f":
                    continue
                value = cell.value
                if isinstance(value, str):
                    normalized = value
                elif hasattr(value, "text"):
                    normalized = json.dumps({"text": value.text, "ref": getattr(value, "ref", None)}, sort_keys=True)
                else:
                    normalized = repr(value)
                formulas[(ws.title, cell.coordinate)] = normalized
    return formulas


def _value_signature(cell: Any) -> Any:
    if cell.data_type == "f":
        value = cell.value
        if isinstance(value, str):
            return ("formula", value)
        if hasattr(value, "text"):
            return ("formula", getattr(value, "text", None), getattr(value, "ref", None))
        return ("formula", repr(value))
    return (cell.data_type, cell.value)


def _validation_signature(ws: Any) -> tuple[Any, ...]:
    fields = ("type", "operator", "formula1", "formula2", "allow_blank", "showDropDown",
              "showInputMessage", "showErrorMessage", "errorTitle", "error", "promptTitle", "prompt")
    items = [(tuple(getattr(rule, field, None) for field in fields), str(rule.sqref))
             for rule in ws.data_validations.dataValidation]
    return tuple(sorted(items, key=repr))


def _defined_name_signature(book: Any) -> dict[str, tuple[Any, ...]]:
    attrs = ("attr_text", "localSheetId", "hidden", "function", "vbProcedure", "xlm",
             "functionGroupId", "shortcutKey", "publishToServer", "workbookParameter")
    return {name: tuple(getattr(item, attr, None) for attr in attrs)
            for name, item in book.defined_names.items()}


def validate_output(template: Path, output: Path, writes: dict[str, dict[str, Any]], changed_parts: set[str]) -> list[str]:
    messages: list[str] = []
    with zipfile.ZipFile(template) as original, zipfile.ZipFile(output) as result:
        if original.testzip() or result.testzip():
            raise WriterError("Template or output ZIP integrity check failed.")
        if original.namelist() != result.namelist():
            raise WriterError("Output workbook package membership/order changed.")
        for name in original.namelist():
            if name not in changed_parts and original.read(name) != result.read(name):
                raise WriterError(f"Unrelated workbook package part changed: {name}")
        messages.append("Workbook package integrity and unrelated package parts verified.")
    before = openpyxl.load_workbook(template, data_only=False)
    after = openpyxl.load_workbook(output, data_only=False)
    if before.sheetnames != after.sheetnames:
        raise WriterError("Worksheet names/order changed.")
    expected_formulas = {key: formula for key, formula in _formula_map(before).items()
                         if key[1] not in writes.get(key[0], {})}
    if expected_formulas != _formula_map(after):
        raise WriterError("Formula map changed; output rejected.")
    for src in before.worksheets:
        dst = after[src.title]
        if set(map(str, src.merged_cells.ranges)) != set(map(str, dst.merged_cells.ranges)):
            raise WriterError(f"Merged ranges changed on {src.title}.")
        for row in src.iter_rows():
            for cell in row:
                other = dst[cell.coordinate]
                expected = writes.get(src.title, {}).get(cell.coordinate, None)
                if cell.coordinate in writes.get(src.title, {}):
                    equal = other.value == expected
                else:
                    equal = _value_signature(other) == _value_signature(cell)
                if not equal:
                    raise WriterError(f"Unexpected output value at {src.title}!{cell.coordinate}.")
                if cell.style_id != other.style_id:
                    raise WriterError(f"Style changed at {src.title}!{cell.coordinate}.")
        if _validation_signature(src) != _validation_signature(dst):
            raise WriterError(f"Data validations changed on {src.title}.")
        if len(src._images) != len(dst._images) or len(src._charts) != len(dst._charts):
            raise WriterError(f"Drawings/charts changed on {src.title}.")
    if _defined_name_signature(before) != _defined_name_signature(after):
        raise WriterError("Named ranges changed.")
    for sheet, cells in writes.items():
        for address, expected in cells.items():
            if after[sheet][address].value != expected:
                raise WriterError(f"Write verification failed at {sheet}!{address}.")
    before.close()
    after.close()
    messages.append("Formulas, styles, merged cells, validations, drawings, named ranges, and worksheet structure verified.")
    return messages


def _validate_project_output(template: Path, output: Path, writes: dict[str, dict[str, Any]],
                             changed_parts: set[str]) -> list[str]:
    """Fast package-level verification for the preallocated project-row path."""
    messages: list[str] = []
    with zipfile.ZipFile(template) as before, zipfile.ZipFile(output) as after:
        if before.testzip() or after.testzip():
            raise WriterError("Template or output ZIP integrity check failed.")
        if before.namelist() != after.namelist():
            raise WriterError("Output workbook package membership/order changed.")
        for name in before.namelist():
            if name not in changed_parts and before.read(name) != after.read(name):
                raise WriterError(f"Unrelated workbook package part changed: {name}")
        before_parts, after_parts = _xml_sheet_parts(before), _xml_sheet_parts(after)
        if before_parts != after_parts:
            raise WriterError("Workbook worksheet structure changed.")
        formula_before: dict[tuple[str, str], tuple[Any, ...]] = {}
        formula_after: dict[tuple[str, str], tuple[Any, ...]] = {}
        touched_shared: dict[str, set[str]] = {}
        for sheet, part in before_parts.items():
            left, right = ET.fromstring(before.read(part)), ET.fromstring(after.read(part))
            for address, semantics in _raw_formula_map(left).items():
                formula_before[(sheet, address)] = semantics
            for address, semantics in _raw_formula_map(right).items():
                formula_after[(sheet, address)] = semantics
            touched_shared[sheet] = _touched_shared_formula_cells(left, writes.get(sheet, {}))
            if sheet in writes:
                for addr in writes[sheet]:
                    a = left.find(f".//{{{NS_MAIN}}}c[@r='{addr}']")
                    b = right.find(f".//{{{NS_MAIN}}}c[@r='{addr}']")
                    if a is None or b is None or a.attrib.get("s") != b.attrib.get("s"):
                        raise WriterError(f"Input cell style changed at {sheet}!{addr}.")
            for tag in ("mergeCells", "dataValidations", "conditionalFormatting", "hyperlinks", "drawing"):
                a = left.find(f"{{{NS_MAIN}}}{tag}")
                b = right.find(f"{{{NS_MAIN}}}{tag}")
                if (ET.tostring(a) if a is not None else None) != (ET.tostring(b) if b is not None else None):
                    raise WriterError(f"Worksheet structure element {tag} changed on {sheet}.")
        # A small, explicit set of preallocated TAKE OFF input cells contains
        # carry-forward formulas in the clean template. Those formulas may be
        # replaced only when the project allocator actually writes that exact
        # input coordinate. Shared groups touched by writes may be expanded to
        # standalone formulas, so compare translated formula semantics rather
        # than requiring the same shared-formula XML encoding.
        formula_overrides = set(formula_before) - set(formula_after)
        for sheet, address in formula_overrides:
            match = re.fullmatch(r"([A-Z]+)(\d+)", address)
            approved = (sheet == "TAKE OFF" and match is not None
                        and match.group(1) in {"A", "B", "D", "E", "F"}
                        and 21 <= int(match.group(2)) <= 1004
                        and address in writes.get(sheet, {}))
            if not approved:
                raise WriterError(f"Formula at {sheet}!{address} was removed outside an approved input write.")
        for sheet, addresses in writes.items():
            for address in addresses:
                match = re.fullmatch(r"([A-Z]+)(\d+)", address)
                if (sheet == "TAKE OFF" and match is not None
                        and match.group(1) in {"A", "B", "D", "E", "F"}
                        and 21 <= int(match.group(2)) <= 1004):
                    continue
                if (sheet, address) in formula_before and (sheet, address) not in formula_after:
                    raise WriterError(f"Write is not an approved formula input at {sheet}!{address}.")
        allowed_write_keys = {(sheet, addr) for sheet, addresses in writes.items() for addr in addresses}
        expected = {key: value for key, value in formula_before.items() if key not in allowed_write_keys}
        actual = {key: value for key, value in formula_after.items() if key not in allowed_write_keys}
        for sheet, addresses in touched_shared.items():
            for address in addresses:
                expected.pop((sheet, address), None)
                actual.pop((sheet, address), None)
        if expected != actual:
            raise WriterError("Calculation/carry formulas changed outside explicitly written input cells.")
        for sheet, addresses in touched_shared.items():
            untouched = addresses - set(writes.get(sheet, {}))
            if not untouched:
                continue
            before_part, after_part = before_parts[sheet], after_parts[sheet]
            expected_semantics = _formula_semantics_for_coordinates(
                ET.fromstring(before.read(before_part)), untouched)
            actual_semantics = _formula_semantics_for_coordinates(
                ET.fromstring(after.read(after_part)), untouched)
            if expected_semantics != actual_semantics:
                raise WriterError(f"Shared formulas changed semantically on {sheet}.")
        shared_problems = audit_shared_formula_package(output)
        if shared_problems:
            raise WriterError("Output shared-formula audit failed: " + "; ".join(shared_problems))
    messages.append("Workbook package membership and untouched package parts verified.")
    messages.append(f"Formula semantics preserved except {len(formula_overrides)} explicitly written TAKE OFF input formulas (A/B/D/E/F, rows 21:1004).")
    messages.append("Shared-formula masters and followers audited; touched groups were safely expanded.")
    messages.append("Input styles, merged cells, validations, drawings, and worksheet structure preserved.")
    return messages


def _raw_formula_map(root: ET.Element) -> dict[str, tuple[Any, ...]]:
    """Return the raw formula text and attributes for each formula cell."""
    return {cell.attrib["r"]: (formula.text or "", tuple(sorted(formula.attrib.items())))
            for cell in root.findall(f".//{{{NS_MAIN}}}c")
            for formula in [cell.find(f"{{{NS_MAIN}}}f")]
            if formula is not None}


def _touched_shared_formula_cells(root: ET.Element, writes: dict[str, Any]) -> set[str]:
    """Return all members in shared groups intersected by a planned write."""
    touched: set[str] = set()
    for group in _shared_formula_groups(root).values():
        coordinates = {coordinate for coordinate, _, _ in group["members"]}
        if coordinates.intersection(writes):
            touched.update(coordinates)
    return touched


def _formula_semantics_for_coordinates(root: ET.Element, coordinates: set[str]) -> dict[str, tuple[Any, ...]]:
    """Expand selected shared formulas to standalone formulas for comparison."""
    cells = _cell_map(root)
    groups = _shared_formula_groups(root)
    semantics: dict[str, tuple[Any, ...]] = {}
    for coordinate in coordinates:
        cell = cells.get(coordinate)
        if cell is None:
            continue
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if formula is None:
            continue
        formula_type = formula.attrib.get("t")
        if formula_type == "shared":
            index = formula.attrib["si"]
            master = groups[index]["master"]
            if master is None:
                raise WriterError(f"Shared formula at {coordinate} has no master (si={index}).")
            origin, master_formula = master
            try:
                translated = Translator("=" + (master_formula.text or ""), origin=origin).translate_formula(coordinate)
            except Exception as exc:
                raise WriterError(f"Cannot expand shared formula {index} at {coordinate}: {exc}") from exc
            expression = translated[1:] if translated.startswith("=") else translated
            attrs = tuple(sorted((key, value) for key, value in master_formula.attrib.items()
                                 if key not in {"t", "si", "ref"}))
        else:
            expression = formula.text or ""
            attrs = tuple(sorted(formula.attrib.items()))
        semantics[coordinate] = (expression, attrs)
    return semantics


def _run_confirmed_project(manual: Path, template: Path, output: Path, report: Path,
                           unresolved: Path, records: list[dict[str, Any]],
                           before_hashes: dict[str, str]) -> dict[str, Any]:
    """Write directly supported Kuujjuaq fields into clean, blank VJR inputs."""
    from decimal import Decimal, ROUND_HALF_UP
    from collections import defaultdict

    reference_candidates = [p for p in (ROOT / "source_files").glob("*VJR.xlsx") if p.resolve() != template.resolve()]
    reference_hash_before = sha256(reference_candidates[0]) if len(reference_candidates) == 1 else None
    ppt_path = ROOT / "source_files" / "VENTILATION_DUCTING.pptx"
    ppt_hash_before = sha256(ppt_path) if ppt_path.is_file() else None
    book = openpyxl.load_workbook(template, data_only=False)
    takeoff, spiral_ws = book["TAKE OFF"], book["SPIRALE"]
    writes: dict[str, dict[str, Any]] = {"TAKE OFF": {}, "SPIRALE": {}}
    unresolved_rows: list[dict[str, Any]] = []
    details: dict[int, list[str]] = defaultdict(list)
    rect_groups: dict[tuple[str, str, str, int | None], dict[str, Any]] = {}
    spirals: dict[int, float] = defaultdict(float)
    elbows45: dict[int, float] = defaultdict(float)
    elbows90: dict[int, float] = defaultdict(float)
    spiral_members: dict[int, list[int]] = defaultdict(list)
    elbow_members: dict[int, list[int]] = defaultdict(list)
    successful_records: set[int] = set()
    counts = Counter()

    def unresolved_field(record: dict[str, Any], field: str, reason: str) -> None:
        unresolved_rows.append({"manual_source_row": record["source_row"], "category": record["category"],
            "unit": record.get("units"), "label": record.get("label"), "duct_size": record.get("duct size"),
            "field": field, "reason": reason})
        details[record["source_row"]].append(f"{field}: {reason}")

    for record in records:
        category, source_row = record["category"], record["source_row"]
        size = str(record.get("duct size") or "").strip()
        unit = str(record.get("units") or "").strip()
        if category == "rectangular":
            match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*", size)
            if not match or not unit:
                unresolved_field(record, "TAKE OFF dimensions", "Manual size is not an explicit width x height pair or Unit is blank.")
                continue
            w, h = float(match.group(1)), float(match.group(2))
            notes = " ".join(str(record.get(k) or "") for k in ("type", "remark")).upper()
            pressure = next((p for p, patterns in (("HP", (r"\bHP\b", "HIGH PRESSURE")), ("MP", (r"\bMP\b", "MEDIUM PRESSURE")), ("BP", (r"\bBP\b", "LOW PRESSURE"))) if any(re.search(pat, notes) if pat.startswith("\\b") else pat in notes for pat in patterns)), "")
            gauge_match = re.search(r"\b(\d+)\s*GA(?:UGE)?\b", notes)
            gauge = int(gauge_match.group(1)) if gauge_match else None
            key = (unit.upper(), f"{max(w,h):g}x{min(w,h):g}", pressure, gauge)
            group = rect_groups.setdefault(key, {"w":max(w,h), "h":min(w,h), "length":0.0, "members":[], "accessories":[]})
            try:
                group["length"] += float(record.get("length") or 0)
                group["members"].append(record)
            except (TypeError, ValueError):
                unresolved_field(record, "TAKE OFF length", "Manual length is not numeric.")
            import rectangular_duct_takeoff as rectangular_rules
            from decimal import Decimal
            for field in ("accessories1", "accessories2"):
                accessory = record.get(field)
                if _is_placeholder(accessory):
                    continue
                try:
                    note = str(accessory).upper()
                    if "ACOUSTIC" in note and ("1/2" in note or "½" in note):
                        code = 1
                    elif "ACOUSTIC" in note and ("1\"" in note or "1 INCH" in note):
                        code = 21
                    else:
                        code = rectangular_rules.map_accessory(accessory, Decimal(str(w)), Decimal(str(h)), "M")
                    group["accessories"].append(code)
                except ValueError as exc:
                    unresolved_field(record, f"{field} accessory code", str(exc))
        elif category in {"spiral", "elbow"}:
            try:
                inches = (Decimal(str(record.get("duct size"))) / Decimal("25.4")).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
                diameter = int(inches)
                if diameter not in [spiral_ws.cell(r, 1).value for r in range(11, 64)]:
                    raise ValueError(f"Converted nominal diameter {diameter} in is absent from SPIRALE diameter rows.")
            except Exception as exc:
                unresolved_field(record, "SPIRALE diameter", f"Cannot map Manual diameter {size!r} mm to a reference diameter row: {exc}")
                continue
            if category == "spiral":
                try:
                    spirals[diameter] += float(record.get("length") or 0)
                    spiral_members[diameter].append(source_row)
                except (TypeError, ValueError): unresolved_field(record, "SPIRALE length", "Manual length is not numeric.")
            else:
                try: quantity = float(record.get("count") or 0)
                except (TypeError, ValueError):
                    unresolved_field(record, "SPIRALE elbow count", "Manual count is not numeric."); continue
                subject = str(record.get("subject") or "").upper()
                (elbows45 if "45" in subject else elbows90)[diameter] += quantity
                elbow_members[diameter].append(source_row)
        elif category == "flex":
            unresolved_field(record, "ROND ET FLEX section", "Manual does not specify acoustic versus insulated flex, so the target F:G or J:K section cannot be selected without guessing.")
        elif category == "damper":
            unresolved_field(record, "Damper destination/code", "No exact applicable VJR damper input cell/code was established from the clean template and Manual fields; no code inferred.")
        elif category == "riser":
            unresolved_field(record, "Riser destination", "Manual riser geometry/orientation does not uniquely identify a writable TAKE OFF or SPIRALE input without overriding carry-down formulas.")
        else:
            unresolved_field(record, "VJR destination", f"No confirmed VJR input mapping for Manual category {category!r}.")

    # All 985 preformatted rows are valid project input slots. A/B and D:F
    # carry formulas are replaced only in rows explicitly allocated to input;
    # calculation formulas K:AY remain untouched.
    rect_written = accessory_rows_written = formula_input_overrides = 0
    row = 20
    for key, group in sorted(rect_groups.items(), key=lambda item: (item[0][0], item[0][1], item[0][2], item[0][3] or 0)):
        pressure, gauge = key[2], key[3]
        accessories = group["accessories"]
        chunks = [accessories[i:i+4] for i in range(0, len(accessories), 4)] or [[]]
        for chunk_index, chunk in enumerate(chunks):
            if row > 1004:
                for rec in group["members"]:
                    unresolved_field(rec, "TAKE OFF row", "Preformatted TAKE OFF input capacity (rows 20:1004) is exhausted.")
                break
            values: dict[str, Any] = {
                f"A{row}": group["w"], f"B{row}": group["h"],
                f"C{row}": round(group["length"], 6) if chunk_index == 0 else 0,
                f"D{row}": 6 if gauge == 16 else None,
                f"E{row}": pressure or None,
                f"F{row}": "HA" if max(group["w"], group["h"]) / 25.4 <= 31 else "RE",
            }
            values.update({f"{col}{row}": chunk[i] if i < len(chunk) else None for i, col in enumerate("GHIJ")})
            for address in values:
                cell = takeoff[address]
                if cell.data_type == "f":
                    allowed = cell.column_letter in {"A", "B", "D", "E", "F"} and 21 <= row <= 1004
                    if not allowed:
                        raise WriterError(f"Refusing to overwrite non-carry-forward formula at TAKE OFF!{address}.")
                    formula_input_overrides += 1
                elif cell.value not in (None, ""):
                    raise WriterError(f"Clean template input cell TAKE OFF!{address} is unexpectedly occupied.")
            writes["TAKE OFF"].update(values)
            if chunk_index == 0:
                rect_written += 1
                for rec in group["members"]:
                    counts["rectangular_records_written"] += 1
                    successful_records.add(rec["source_row"])
                    if not pressure:
                        unresolved_field(rec, "TAKE OFF type", "Scope/Remarks do not explicitly establish BP, MP, or HP.")
                    if gauge != 16:
                        unresolved_field(rec, "TAKE OFF metal code", "No supported gauge/material code is explicit in the Manual notes.")
            else:
                accessory_rows_written += 1
            row += 1

    spiral_rows = {spiral_ws.cell(r,1).value:r for r in range(11,64)}
    spiral_written = elbow_written = 0
    for diameter, quantity in spirals.items():
        row = spiral_rows[diameter]; addr=f"B{row}"
        if spiral_ws[addr].value in (None, "") and spiral_ws[addr].data_type != "f":
            writes["SPIRALE"][addr]=round(quantity, 6); spiral_written += 1; successful_records.update(spiral_members[diameter])
    for diameter, quantity in elbows90.items():
        row=spiral_rows[diameter]; addr=f"C{row}"
        if spiral_ws[addr].value in (None, "") and spiral_ws[addr].data_type != "f":
            writes["SPIRALE"][addr]=quantity; elbow_written += 1; successful_records.update(elbow_members[diameter])
    for diameter, quantity in elbows45.items():
        row=spiral_rows[diameter]; addr=f"D{row}"
        if spiral_ws[addr].value in (None, "") and spiral_ws[addr].data_type != "f":
            writes["SPIRALE"][addr]=quantity; elbow_written += 1; successful_records.update(elbow_members[diameter])
    book.close()

    writes = {sheet: cells for sheet,cells in writes.items() if cells}
    changed = _patch_copy(template, output, writes)
    preservation = _validate_project_output(template, output, writes, changed)
    if {"Manual":sha256(manual), "VJR template":sha256(template)} != before_hashes:
        raise WriterError("A source workbook changed during processing.")
    unresolved.parent.mkdir(parents=True, exist_ok=True)
    fields=["manual_source_row","category","unit","label","duct_size","field","reason"]
    with unresolved.open("w",newline="",encoding="utf-8-sig") as stream:
        writer=csv.DictWriter(stream,fieldnames=fields); writer.writeheader(); writer.writerows(unresolved_rows)
    total_cells=sum(len(cells) for cells in writes.values())
    lines=["# Kuujjuaq Manual to VJR implementation validation","",
        "## Run totals","",f"- Manual records: **{len(records)}**",
        f"- Records with at least one successful write: **{len(successful_records)}**",
        f"- Records completely skipped: **{len(records)-len(successful_records)}**",
        f"- Rectangular rows written: **{rect_written}**", f"- Accessory rows written: **{accessory_rows_written}**",
        f"- Spiral rows written: **{spiral_written}**",f"- Elbow rows written: **{elbow_written}**",
        "- Flex rows written: **0**","- Damper rows written: **0**",
        f"- Unresolved records: **{len({int(x['manual_source_row']) for x in unresolved_rows})}**",
        f"- Cells written: **{total_cells}**","- Calculation formulas modified: **0**",
        f"- Carry-forward input formulas intentionally replaced by values: **{formula_input_overrides}**",
        f"- Output: `{output}`",f"- Unresolved field ledger: `{unresolved}`","",
        "## Template and mapping","",
        f"Clean VJR template used: `{template}`. The populated project VJR was not used as the output base.",
        "TAKE OFF rows 20:1004 are preformatted inputs. A/B and D:F carry-forward formulas are replaced only on rows assigned project inputs; K:AY calculation formulas remain unchanged. C and G:J are literal blank inputs. Spiral metric diameters are converted to nominal whole-inch rows using nearest-inch rounding, consistent with the integer diameter index in SPIRALE.","",
        "## Preservation checks","",*[f"- PASS: {x}" for x in preservation],
        "- PASS: Manual and clean template SHA-256 hashes unchanged.",
        ("- PASS: Populated Reference VJR SHA-256 unchanged." if reference_hash_before is not None and sha256(reference_candidates[0]) == reference_hash_before else "- Reference workbook hash not established."),
        ("- PASS: PPT rule source SHA-256 unchanged." if ppt_hash_before is not None and sha256(ppt_path) == ppt_hash_before else "- PPT rule source hash not established."),
        f"- PASS: Calculation formulas are preserved; {formula_input_overrides} carry-forward input formulas were replaced only in used A/B/D:F cells.","",
        f"- Carry-forward formulas intentionally overridden in TAKE OFF input columns: {formula_input_overrides}; all calculation formulas K:AY and formulas on other sheets remain identical.",
        "Unused preformatted rows retain their carry-forward formulas. No row insertion was necessary. See the CSV for unresolved source fields and exact reasons."]
    report.parent.mkdir(parents=True,exist_ok=True); report.write_text("\n".join(lines),encoding="utf-8")
    return {"output":output,"records":len(records),"written_cells":total_cells,
            "rectangular_rows":rect_written,"spiral_rows":spiral_written,"elbow_rows":elbow_written,
            "flex_rows":0,"damper_rows":0,"unresolved_records":len({int(x['manual_source_row']) for x in unresolved_rows})}


def run(manual_path: Path, template: Path, mapping_path: Path | None = None,
        output_dir: Path = OUTPUT_DIR, test_mode: bool = False,
        output_name: str | None = None) -> dict[str, Any]:
    manual_path, template = Path(manual_path).resolve(), Path(template).resolve()
    project_confirmed = mapping_path is None and (manual_path.suffix.casefold() == ".csv" or manual_path.name.startswith("TAKEOFF_CS KATIVIK"))
    if project_confirmed:
        clean_template = ROOT / "source_files" / "Ventilation Jean Roy excel.xlsx"
        if not clean_template.is_file():
            names = ", ".join(p.name for p in sorted((ROOT / "source_files").glob("*.xlsx")))
            raise WriterError(f"Clean VJR template not found. source_files workbooks: {names}")
        template = clean_template.resolve()
    if not manual_path.is_file() or not template.is_file():
        raise WriterError("Manual workbook or VJR template does not exist.")
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / (output_name or f"{manual_path.stem}_VJR.xlsx")
    report_path = output_dir / f"{manual_path.stem}_validation_report.md"
    comparison_path = output_dir / f"{manual_path.stem}_comparison.csv"
    if output in {manual_path, template} or output.resolve() in {manual_path, template}:
        raise WriterError("Output must be separate from both source workbooks.")
    template_check = openpyxl.load_workbook(template, read_only=True, data_only=False)
    template_sheets = set(template_check.sheetnames)
    template_check.close()
    if not {"TAKE OFF", "SPIRALE", "ROND ET FLEX"}.issubset(template_sheets):
        raise WriterError("VJR template is missing one or more required takeoff worksheets.")
    before_hashes = {"Manual": sha256(manual_path), "VJR template": sha256(template)}
    records, input_problems = inspect_manual(manual_path)
    statuses = _base_statuses(manual_path, records)
    records_by_row = {record["source_row"]: record for record in records}
    dispositions = {row: {"status": status, "reason": reason, "target": ""} for row, (status, reason) in statuses.items()}
    # The Kuujjuaq team has confirmed direct project rules. When the actual
    # project Manual is run without a legacy manifest, write the unambiguous
    # rectangular dimensions and lengths into new TAKE OFF input rows. The old
    # dry-run dispositions are evidence only and must not veto these writes.
    if project_confirmed:
        if manual_path.suffix.casefold() == ".csv":
            from manual_csv_to_vjr import generate as generate_from_csv
            references = [p for p in (ROOT / "source_files").glob("*VJR.xlsx") if p.resolve() != template.resolve()]
            reference = references[0] if len(references) == 1 else None
            project_report = output_dir / f"{manual_path.stem}_VJR_validation_report.md"
            result = generate_from_csv(manual_path, template, output, project_report, reference)
            return {**result, "counts": result["status_counts"]}
        project_report = ROOT / "reports" / "kuujjuaq_implementation_validation.md"
        project_unresolved = ROOT / "reports" / "kuujjuaq_implementation_unresolved.csv"
        result = _run_confirmed_project(manual_path, template, output, project_report,
                                        project_unresolved, records, before_hashes)
        result.update({"report": project_report, "comparison": project_unresolved})
        return result
    manifest = _read_manifest(mapping_path)
    planned: list[tuple[int, str, int, dict[str, Any]]] = []
    for entry in manifest:
        scope = entry.get("manifest_scope", "").strip().upper()
        if scope == "TEST_ONLY":
            test_folder = manual_path.parent == template.parent and template.parent.name.casefold() == "test_input"
            if not test_mode or not test_folder:
                ok, reason, values = False, "TEST_ONLY manifests require --test and a Manual/template pair in a test_input folder.", {}
            else:
                ok, reason, values = _manifest_rule_check(entry, records_by_row)
        elif scope == "APPROVED_PROJECT":
            ok, reason, values = _manifest_rule_check(entry, records_by_row)
        else:
            ok, reason, values = False, "Manifest scope must be TEST_ONLY or APPROVED_PROJECT.", {}
        try:
            source_row = int(entry["manual_source_row"])
            target_row = int(entry["target_row"])
        except (KeyError, ValueError):
            continue
        if source_row not in dispositions:
            continue
        if dispositions[source_row]["status"] == "UNSUPPORTED":
            continue
        if ok:
            planned.append((source_row, entry["target_sheet"], target_row, values))
            dispositions[source_row] = {"status": "SAFE_TO_AUTOMATE", "reason": reason, "target": f"{entry['target_sheet']} row {target_row}"}
        else:
            dispositions[source_row] = {"status": "NEEDS_CONFIRMATION", "reason": reason, "target": f"{entry.get('target_sheet', '')} row {entry.get('target_row', '')}"}
    units_by_source_row = {}
    for entry in manifest:
        try:
            units_by_source_row[int(entry["manual_source_row"])] = entry.get("dimension_unit", "").strip().lower()
        except (KeyError, ValueError):
            continue
    target_writes, target_problems = _check_target_cells(template, planned, units_by_source_row)
    for source_row, sheet, row, values in planned:
        related = [problem for problem in target_problems if problem.startswith(f"Manual row {source_row}:")]
        if related:
            dispositions[source_row] = {"status": "NEEDS_CONFIRMATION", "reason": " ".join(related), "target": f"{sheet} row {row}"}
    # Remove cells for any mapping vetoed above; collisions are never resolved by order.
    cell_sources: dict[tuple[str, str], list[int]] = {}
    for source_row, sheet, row, values in planned:
        if dispositions[source_row]["status"] != "SAFE_TO_AUTOMATE":
            continue
        for address in values:
            cell_sources.setdefault((sheet, address), []).append(source_row)
    for (sheet, address), sources in cell_sources.items():
        if len(sources) > 1:
            for source_row in sources:
                dispositions[source_row] = {"status": "NEEDS_CONFIRMATION", "reason": f"Multiple Manual records target the same cell {sheet}!{address}; group aggregation is not approved.", "target": f"{sheet}!{address}"}
    accepted = {(sheet, cell) for row, sheet, target_row, values in planned
                if dispositions[row]["status"] == "SAFE_TO_AUTOMATE" for cell in values}
    writes: dict[str, dict[str, Any]] = {}
    for source_row, sheet, target_row, values in planned:
        if dispositions[source_row]["status"] != "SAFE_TO_AUTOMATE":
            continue
        target_values = writes.setdefault(sheet, {})
        target_values.update({cell: value for cell, value in values.items() if (sheet, cell) in accepted})
    writes = {sheet: values for sheet, values in writes.items() if values}
    # Remove an entire row's write set if any one of its cells was vetoed.
    for source_row, sheet, target_row, values in planned:
        if dispositions[source_row]["status"] != "SAFE_TO_AUTOMATE":
            for cell in values:
                writes.get(sheet, {}).pop(cell, None)
    changed_parts = _patch_copy(template, output, writes)
    integrity_messages = validate_output(template, output, writes, changed_parts)
    after_hashes = {"Manual": sha256(manual_path), "VJR template": sha256(template)}
    if before_hashes != after_hashes:
        raise WriterError("A source workbook changed during processing.")
    counts = Counter(item["status"] for item in dispositions.values())
    _write_comparison(comparison_path, records, dispositions)
    _write_report(report_path, manual_path, template, output, mapping_path, records,
                  dispositions, counts, before_hashes, after_hashes, integrity_messages,
                  input_problems, test_mode)
    return {"output": output, "report": report_path, "comparison": comparison_path,
            "counts": {status: counts.get(status, 0) for status in STATUSES},
            "records": len(records), "written_cells": sum(map(len, writes.values()))}


def _write_comparison(path: Path, records: list[dict[str, Any]], dispositions: dict[int, dict[str, str]]) -> None:
    fields = ["manual_source_row", "category", "manual_system", "duct_size", "length", "length_unit", "status", "target", "reason"]
    with path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for record in records:
            item = dispositions[record["source_row"]]
            writer.writerow({"manual_source_row": record["source_row"], "category": record["category"],
                             "manual_system": record.get("units"), "duct_size": record.get("duct size"),
                             "length": record.get("length"), "length_unit": record.get("length unit"),
                             "status": item["status"], "target": item["target"], "reason": item["reason"]})


def _write_report(path: Path, manual: Path, template: Path, output: Path, mapping: Path | None,
                  records: list[dict[str, Any]], dispositions: dict[int, dict[str, str]],
                  counts: Counter[str], before: dict[str, str], after: dict[str, str],
                  integrity: list[str], input_problems: list[str], test_mode: bool) -> None:
    lines = [
        "# Manual Team to VJR validation report", "",
        f"Mode: {'TEST' if test_mode else 'FULL'}", "",
        "## Inputs and outputs", "",
        f"- Manual input: `{manual}`", f"- VJR template: `{template}`",
        f"- Approved mapping manifest: `{mapping if mapping else 'not supplied'}`",
        f"- Output workbook: `{output}`", "",
        "## Record dispositions", "",
        f"- Records processed: **{len(records)}**",
        f"- SAFE_TO_AUTOMATE: **{counts.get('SAFE_TO_AUTOMATE', 0)}**",
        f"- NEEDS_CONFIRMATION: **{counts.get('NEEDS_CONFIRMATION', 0)}**",
        f"- UNSUPPORTED: **{counts.get('UNSUPPORTED', 0)}**",
        "- Only SAFE_TO_AUTOMATE records were written. Every skipped record and reason is in the comparison CSV.", "",
        "## Preservation checks", "",
        *[f"- PASS: {line}" for line in integrity],
        "- PASS: Manual and VJR source hashes are unchanged.", "",
        "| Source | SHA-256 before | SHA-256 after |", "|---|---|---|",
    ]
    lines.extend(f"| {name} | `{before[name]}` | `{after[name]}` |" for name in before)
    if input_problems:
        lines += ["", "## Input findings", "", *[f"- {problem}" for problem in input_problems]]
    lines += ["", "## Safe write records", ""]
    safe = [(row, data) for row, data in dispositions.items() if data["status"] == "SAFE_TO_AUTOMATE"]
    lines.extend(f"- Manual DATA row {row} -> {data['target']}: {data['reason']}" for row, data in safe)
    if not safe:
        lines.append("- None. Output is an unchanged copy of the selected VJR template.")
    lines += ["", "## Skipped records", "", "See the comparison CSV for every skipped source row and reason.", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manual", nargs="?", help="Manual Team .csv/.xlsx path or a filename in test_input/")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Client VJR template .xlsx")
    parser.add_argument("--mapping", type=Path, help="Estimator-approved write manifest CSV")
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--test", action="store_true", help="Run in test mode (same safety gates, separate output)")
    parser.add_argument("--output-name", help="Exact output workbook filename under --output-dir")
    args = parser.parse_args(argv)
    try:
        manual = resolve_manual(args.manual)
        result = run(manual, args.template, args.mapping, args.output_dir, args.test, args.output_name)
    except (WriterError, OSError, zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print(f"Records: {result['records']}")
    if "status_counts" in result:
        for status, count in sorted(result["status_counts"].items()):
            print(f"{status}: {count}")
    else:
        for status in STATUSES:
            print(f"{status}: {result['counts'][status]}")
    print(f"Written cells: {result.get('written_cells', result.get('cells_written', 0))}")
    print(f"VJR output: {result['output']}")
    print(f"Validation: {result['report']}")
    print(f"Comparison: {result['comparison']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
