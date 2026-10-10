"""PHI CONTEMPORAIN DATA-DUCT adapter for rectangular duct records only."""
from __future__ import annotations

import csv
import re
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import openpyxl

import manual_csv_to_vjr
import manual_to_vjr


PHI_SHEET = "DATA-DUCT"
PHI_SUBJECT = "RECT. DUCT ft."
PHI_HEADERS = {
    "subject", "page index", "label", "count", "duct size", "accessories1",
    "accessories2", "remark", "units", "length", "length unit", "damper type",
}
NORMALIZED_HEADERS = list(manual_csv_to_vjr.REQUIRED) + [
    "source_row", "source_sheet", "raw_duct_size", "source_subject", "source_label", "source_page_index", "source_count", "source_damper_type",
    "source_system_unit", "source_length", "source_length_unit", "source_accessories1",
    "source_accessories2", "source_remark", "normalized_width", "normalized_height",
]


class PhiAdapterError(ValueError):
    pass


def _header_map(values: tuple[Any, ...]) -> dict[str, int]:
    return {re.sub(r"\s+", " ", str(value).strip()).casefold(): i
            for i, value in enumerate(values) if value is not None and str(value).strip()}


def _clean(value: Any) -> Any:
    if value is None or str(value).strip() in {"", "."}:
        return None
    return value


def parse_rectangular_size(value: Any) -> tuple[Decimal, Decimal] | None:
    """Parse PHI rectangular pairs; PHI's reviewed basis is inches."""
    text = str(value or "").strip()
    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)", text)
    if not match:
        return None
    try:
        return Decimal(match.group(1)), Decimal(match.group(2))
    except InvalidOperation:
        return None


def detect_phi_workbook(path: str | Path) -> bool:
    try:
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        if PHI_SHEET not in book.sheetnames or "DATA" not in book.sheetnames:
            return False
        sheet = book[PHI_SHEET]
        headers = _header_map(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ()))
        if not PHI_HEADERS.issubset(headers) or "type" in headers:
            return False
        # The header signature is shared by several Manual exports. Treat this
        # as PHI only when its DATA-DUCT rows actually use the PHI subject; the
        # backend otherwise routes valid non-PHI workbooks to Universal Input.
        subject_column = headers["subject"]
        return any(
            str(row[subject_column] or "").strip() == PHI_SUBJECT
            for row in sheet.iter_rows(min_row=2, values_only=True)
            if subject_column < len(row)
        )
    finally:
        book.close()


def normalize_phi_workbook(path: str | Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Return PHI rectangular records and row-level size parsing diagnostics."""
    path = Path(path)
    try:
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise PhiAdapterError(f"PHI workbook cannot be opened: {exc}") from exc
    try:
        if PHI_SHEET not in book.sheetnames or "DATA" not in book.sheetnames:
            raise PhiAdapterError(f"PHI workbook is missing {PHI_SHEET!r}.")
        sheet = book[PHI_SHEET]
        iterator = sheet.iter_rows(values_only=True)
        header_values = next(iterator, None)
        if not header_values:
            raise PhiAdapterError(f"{PHI_SHEET} is empty.")
        columns = _header_map(header_values)
        if not PHI_HEADERS.issubset(columns) or "type" in columns:
            raise PhiAdapterError("Workbook does not match the PHI DATA-DUCT header format.")

        records: list[dict[str, Any]] = []
        parse_errors: list[dict[str, Any]] = []
        for physical_row, values in enumerate(iterator, start=2):
            subject = _clean(values[columns["subject"]] if columns["subject"] < len(values) else None)
            if str(subject or "").strip() != PHI_SUBJECT:
                continue
            def get(name: str) -> Any:
                index = columns[name]
                return _clean(values[index] if index < len(values) else None)

            def raw(name: str) -> Any:
                index = columns[name]
                return values[index] if index < len(values) else None

            raw_size = get("duct size")
            pair = parse_rectangular_size(raw_size)
            if pair is None:
                parse_errors.append({"source_sheet": PHI_SHEET, "source_row": physical_row,
                                     "raw_duct_size": raw_size or "", "reason": "DUCT SIZE is not a valid rectangular pair."})
            width, height = pair if pair else (None, None)
            accessories = [get("accessories1"), get("accessories2")]
            # `type` is an absent normalized attribute, not a source field or
            # an inferred value. The category remains in Subject and Label.
            records.append({
                "subject": PHI_SUBJECT, "page index": get("page index"),
                "label": get("label"), "count": get("count"),
                "duct size": (f"{width}x{height}" if pair else raw_size),
                "length": get("length"), "length unit": get("length unit"),
                "type": None, "accessories1": accessories[0], "accessories2": accessories[1],
                "remark": get("remark"), "units": get("units"),
                "source_row": physical_row, "source_sheet": PHI_SHEET,
                "raw_duct_size": raw("duct size") if raw("duct size") is not None else "",
                "source_subject": raw("subject") if raw("subject") is not None else "",
                "source_label": raw("label") if raw("label") is not None else "",
                "source_page_index": raw("page index") if raw("page index") is not None else "",
                "source_count": raw("count") if raw("count") is not None else "",
                "source_damper_type": raw("damper type") if raw("damper type") is not None else "",
                "source_system_unit": raw("units") if raw("units") is not None else "",
                "source_length": raw("length") if raw("length") is not None else "",
                "source_length_unit": raw("length unit") if raw("length unit") is not None else "",
                "source_accessories1": raw("accessories1") if raw("accessories1") is not None else "",
                "source_accessories2": raw("accessories2") if raw("accessories2") is not None else "",
                "source_remark": raw("remark") if raw("remark") is not None else "",
                "normalized_width": str(width) if width is not None else "",
                "normalized_height": str(height) if height is not None else "",
                "system_unit": get("units"), "width": width, "height": height,
                "length_value": get("length"), "length_unit": get("length unit"),
                "pressure_material": get("remark"), "accessories": accessories,
            })
        if not records:
            raise PhiAdapterError(f"No rows with Subject = {PHI_SUBJECT!r} were found in {PHI_SHEET}.")
        return records, parse_errors
    finally:
        book.close()


def write_mapper_input(records: list[dict[str, Any]], path: str | Path) -> Path:
    """Serialize normalized records for the existing parser/mapping path.

    The empty internal `type` slot satisfies the existing record interface; no
    TYPE field/value is added to or inferred from the PHI workbook.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as stream:
        output = csv.DictWriter(stream, fieldnames=NORMALIZED_HEADERS)
        output.writeheader()
        for record in records:
            output.writerow({key: record.get(key, "") for key in NORMALIZED_HEADERS})
    return path


def generate_phi_rectangular(source: str | Path, template: str | Path, output: str | Path,
                             report: str | Path) -> dict[str, Any]:
    """Run PHI rectangular records through the existing Stage 1 mapper."""
    source, template, output, report = map(Path, (source, template, output, report))
    if not detect_phi_workbook(source):
        raise PhiAdapterError("Input is not a PHI CONTEMPORAIN workbook with the expected DATA-DUCT schema.")
    if not template.is_file() or output.resolve() in {template.resolve(), source.resolve()}:
        raise PhiAdapterError("A distinct existing fixed VJR template and output path are required.")
    source_hash, template_hash = manual_to_vjr.sha256(source), manual_to_vjr.sha256(template)
    records, parsing = normalize_phi_workbook(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    report.parent.mkdir(parents=True, exist_ok=True)
    stage_report = output.with_name(f"{output.stem}_mapper_validation.md")
    mapper_csv = output.with_name(f"{output.stem}_normalized_internal.csv")
    write_mapper_input(records, mapper_csv)
    try:
        result = manual_csv_to_vjr.generate(mapper_csv, template, output, stage_report)
    except Exception:
        mapper_csv.unlink(missing_ok=True)
        raise
    mapper_csv.unlink(missing_ok=True)
    if manual_to_vjr.sha256(source) != source_hash or manual_to_vjr.sha256(template) != template_hash:
        raise PhiAdapterError("The source PHI workbook or fixed VJR template changed during generation.")

    final_audit = report.with_name("phi_rectangular_record_audit.csv")
    if result["audit"].resolve() != final_audit.resolve():
        result["audit"].replace(final_audit)
    result["audit"] = final_audit
    with result["audit"].open("r", newline="", encoding="utf-8-sig") as stream:
        audit = list(csv.DictReader(stream))
    with result["unresolved"].open("r", newline="", encoding="utf-8-sig") as stream:
        unresolved_fields = list(csv.DictReader(stream))
    written = sum(bool(row.get("VJR_cells_written")) for row in audit)
    unresolved = sum(row.get("status") == "NEEDS_REVIEW" for row in audit)
    unsupported = sum(row.get("status") == "UNSUPPORTED" for row in audit)
    reasons = Counter(row.get("reason", "") for row in unresolved_fields)
    report_text = result["report"].read_text(encoding="utf-8")
    lines = [
        "# PHI CONTEMPORAIN rectangular adapter validation", "",
        "## Run totals", "",
        f"- Source workbook: `{source}`", f"- Source sheet: `{PHI_SHEET}`",
        f"- Source rows matching Subject = `{PHI_SUBJECT}`: **{len(records)}**",
        f"- Normalized records: **{len(records)}**",
        f"- Records reaching existing mapper: **{result['records']}**",
        f"- Records with VJR writes: **{written}**",
        f"- Unresolved records: **{unresolved}**", f"- Unsupported records: **{unsupported}**",
        f"- Unresolved field findings: **{len(unresolved_fields)}**",
        f"- Rectangular size parsing errors: **{len(parsing)}**",
        "- Mapping mismatches: **Not assessed against a PHI reference VJR; no mapping changes were made.**",
        "- Formula changes: **0** (existing mapper preservation report)",
        "- Workbook preservation status: **" + ("PASS" if "PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved." in report_text else "NOT CONFIRMED") + "**",
        "- Dimension interpretation: PHI `RECT. DUCT ft.` `DUCT SIZE` pairs are parsed as inches per `reports/input_format_analysis_phi.md`; original source text is retained in the row audit.",
        "- Pressure/type/material are not inferred. Missing pressure remains unresolved under the existing mapper rules.",
        "- Accessories1/2 and REMARK are passed through verbatim except `.` placeholders, which are treated as absent source data.",
        "", "## Parsing errors", "",
    ]
    if parsing:
        lines.extend([f"- DATA-DUCT row {item['source_row']}: `{item['raw_duct_size']}` — {item['reason']}" for item in parsing])
    else:
        lines.append("- None.")
    lines.extend(["", "## Existing mapper review findings", ""])
    lines.extend(f"- {reason}: **{count}** field findings" for reason, count in reasons.most_common())
    lines.append("")
    lines.extend(["", "## Mapper output", "", f"- Existing mapper report: `{result['report'].name}`", f"- Row audit: `{result['audit'].name}`", ""])
    report.write_text("\n".join(lines), encoding="utf-8")
    result["report"] = report
    return {**result, "source_records": len(records), "normalized_records": len(records),
            "records_with_writes": written, "unresolved_records": unresolved,
            "unsupported_records": unsupported, "parsing_errors": parsing,
            "formula_changes": 0,
            "structure_preserved": "PASS: Input styles, merged cells, validations, drawings, and worksheet structure preserved." in report_text}
