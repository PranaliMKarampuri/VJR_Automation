"""Thin upload/job wrapper around the validated Stage 1 CSV generator."""
from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Any

import openpyxl

import manual_csv_to_vjr
import manual_to_vjr
import phi_rectangular_adapter
import universal_input


ROOT = Path(__file__).resolve().parent
FIXED_TEMPLATE = ROOT / "source_files" / "Ventilation Jean Roy excel.xlsx"
DEFAULT_OUTPUT_ROOT = ROOT / "output" / "jobs"


class Stage1InputError(ValueError):
    """User-correctable input or template problem."""


def preview_input(input_file: str | Path) -> dict[str, Any]:
    """Inspect an uploaded file with the same normalizer used by generation.

    This is detection only: it does not invoke or reproduce VJR mapping rules.
    """
    source = Path(input_file)
    if not source.is_file() or source.stat().st_size == 0:
        raise Stage1InputError("The uploaded Manual file is empty or could not be read.")
    if source.suffix.casefold() not in {".csv", ".xlsx"}:
        if source.suffix.casefold() == ".xls":
            raise Stage1InputError("Legacy .xls files are not supported. Save as .xlsx or export the DATA sheet as .csv.")
        raise Stage1InputError("Upload a Manual .csv or .xlsx file.")
    try:
        if source.suffix.casefold() == ".xlsx" and phi_rectangular_adapter.detect_phi_workbook(source):
            records, _ = phi_rectangular_adapter.normalize_phi_workbook(source)
            detection = {"format": "PHI CONTEMPORAIN Excel", "sheets": ["DATA-DUCT"],
                         "columns": {"DATA-DUCT": ["Subject", "Page Index", "Label", "Count", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "Length", "Length Unit", "DAMPER TYPE"]},
                         "category_counts": {"rectangular": len(records)}, "total_records": len(records),
                         "detected_input_records": len(records), "detected_input_category_counts": {"rectangular": len(records)},
                         "source_selection": None, "records": records}
        else:
            detection = universal_input.inspect_file(source)
    except (ValueError, OSError, KeyError) as exc:
        raise Stage1InputError(f"Manual structure detection failed: {exc}") from exc
    categories = detection.get("detected_input_category_counts", detection["category_counts"])
    warnings = []
    if categories.get("other/unknown", 0):
        warnings.append(f"{categories['other/unknown']} record(s) have an unsupported or unrecognized category.")
    records = detection.get("records", [])
    if any(not (r.get("type") or r.get("pressure_material")) for r in records):
        warnings.append("Some records do not provide pressure/material or TYPE information; applicable output fields may remain blank for review.")
    if any(r.get("category") == "accessories" and not r.get("system_unit") for r in records):
        warnings.append("Some accessory records have no system association; they may remain unresolved instead of being attached by guesswork.")
    if detection.get("source_selection"):
        selection = detection["source_selection"]
        warnings.append(f"Rectangular rows will use {selection['authoritative_sheet']} ({selection['grouped_records']} reconciled groups); {selection['detail_records']} detail records are cross-reference only.")
        reviews = selection.get("source_review_items", [])
        if any(item.get("field") == "ACCESSORIES1" for item in reviews):
            warnings.append("Some rectangular accessory associations could not be corroborated by system and size; those codes are held for review.")
        if any(item.get("field") == "D (Metal Code)" for item in reviews):
            warnings.append("Rectangular duct material is not stated in the source; Metal Code D will remain blank.")
    if detection.get("detected_input_records", detection["total_records"]) != detection["total_records"]:
        warnings.append(f"{detection['detected_input_records']} source DATA rows detected; {detection['total_records']} canonical records will be sent to the mapper after authoritative grouping and detail reconciliation.")
    return {**detection, "warnings": warnings,
            "required_information_available": bool(records) and bool(detection.get("columns"))}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _safe_stem(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")
    return (stem or "Manual")[:80]


def _validate_template(template: Path) -> None:
    if not template.is_file():
        raise Stage1InputError(f"The fixed VJR template was not found: {template}")
    try:
        book = openpyxl.load_workbook(template, read_only=True, data_only=False)
    except Exception as exc:
        raise Stage1InputError(f"The fixed VJR template cannot be opened: {exc}") from exc
    try:
        missing = {"TAKE OFF", "SPIRALE", "ROND ET FLEX"}.difference(book.sheetnames)
        if missing:
            raise Stage1InputError("The fixed VJR template is missing required sheets: " + ", ".join(sorted(missing)))
    finally:
        book.close()


def _summary(result: dict[str, Any], audit_path: Path, unresolved_path: Path) -> dict[str, Any]:
    with audit_path.open("r", newline="", encoding="utf-8-sig") as stream:
        audit = list(csv.DictReader(stream))
    with unresolved_path.open("r", newline="", encoding="utf-8-sig") as stream:
        unresolved = list(csv.DictReader(stream))
    missing_reasons = ("missing", "absent", "unavailable", "does not identify", "does not specify")
    missing_rows = {
        row["csv_row"] for row in unresolved
        if row.get("field_status") == "NEEDS_REVIEW"
        and any(term in row.get("reason", "").casefold() for term in missing_reasons)
    }
    unsupported_rows = {row["csv_row"] for row in unresolved if row.get("field_status") == "UNSUPPORTED"}
    unresolved_rows = {row["csv_row"] for row in unresolved}
    written_records = sum(bool(row.get("VJR_cells_written")) for row in audit)
    return {
        "records_processed": result["records"],
        "records_with_writes": written_records,
        "zero_write_records": len(audit) - written_records,
        "cells_written": result["cells_written"],
        "missing_source_data_records": len(missing_rows),
        "unsupported_records": len(unsupported_rows),
        "other_unresolved_records": len(unresolved_rows - missing_rows - unsupported_rows),
        "unresolved_records": len(unresolved_rows),
        "unresolved_field_count": len(unresolved),
    }


def _apply_source_reviews(audit_path: Path, unresolved_path: Path,
                         review_items: list[dict[str, Any]]) -> None:
    """Add normalizer/source-selection reviews without changing mapper rules."""
    if not review_items:
        return
    with audit_path.open("r", newline="", encoding="utf-8-sig") as stream:
        audit_reader = csv.DictReader(stream)
        audit_fields = list(audit_reader.fieldnames or [])
        audit = list(audit_reader)
    with unresolved_path.open("r", newline="", encoding="utf-8-sig") as stream:
        unresolved_reader = csv.DictReader(stream)
        unresolved_fields = list(unresolved_reader.fieldnames or [])
        unresolved = list(unresolved_reader)
    for field in ("field", "source_sheet", "source_row"):
        if field not in unresolved_fields:
            unresolved_fields.append(field)
    audit_by_id = {row.get("csv_row"): row for row in audit}
    for item in review_items:
        mapper_id = str(item["source_row"])
        audit_row = audit_by_id.get(mapper_id)
        if audit_row:
            audit_row["status"] = "NEEDS_REVIEW"
            message = f"{item['field']}: {item['reason']}"
            current = audit_row.get("reason", "")
            if message not in current:
                audit_row["reason"] = "; ".join(x for x in (current, message) if x)
        unresolved.append({"csv_row": mapper_id, "category": item.get("category", "rectangular"),
            "field_status": "NEEDS_REVIEW", "field": item["field"], "reason": item["reason"],
            "source_sheet": item.get("source_sheet", ""), "source_row": item.get("source_physical_row", "")})
    for path, fields, rows in ((audit_path, audit_fields, audit),
                               (unresolved_path, unresolved_fields, unresolved)):
        with path.open("w", newline="", encoding="utf-8-sig") as stream:
            writer = csv.DictWriter(stream, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)
def generate_vjr(input_file: str | Path, template_file: str | Path | None = None,
                 output_dir: str | Path | None = None) -> dict[str, Any]:
    """Generate a job-isolated VJR using the existing Stage 1 mapping engine.

    Excel inputs are converted from their DATA sheet to the engine's supported
    Manual CSV schema. Legacy .xls files are rejected with a clear conversion
    instruction because the project does not include an .xls reader.
    """
    source = Path(input_file).resolve()
    template = Path(template_file or FIXED_TEMPLATE).resolve()
    root = Path(output_dir or DEFAULT_OUTPUT_ROOT).resolve()
    if not source.is_file():
        raise Stage1InputError("The uploaded Manual file could not be found.")
    if source.suffix.casefold() not in {".csv", ".xlsx"}:
        if source.suffix.casefold() == ".xls":
            raise Stage1InputError("Legacy .xls files are not supported. Save the Manual as .xlsx or export its DATA sheet as .csv.")
        raise Stage1InputError("Upload a Manual .csv or .xlsx file.")
    if source.stat().st_size == 0:
        raise Stage1InputError("The uploaded Manual file is empty.")
    _validate_template(template)

    input_hash = _sha256(source)
    template_hash = _sha256(template)
    root.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    stem = _safe_stem(source.name)
    staging = Path(tempfile.mkdtemp(prefix=".stage1-", dir=root)).resolve()
    final_dir = root / f"{stem}_{job_id[:10]}"
    try:
        csv_input = staging / f"{stem}_canonical.csv"
        try:
            if source.suffix.casefold() == ".xlsx" and phi_rectangular_adapter.detect_phi_workbook(source):
                phi_records, phi_errors = phi_rectangular_adapter.normalize_phi_workbook(source)
                phi_rectangular_adapter.write_mapper_input(phi_records, csv_input)
                detection = {"format": "PHI CONTEMPORAIN Excel", "sheets": ["DATA-DUCT"],
                    "columns": {"DATA-DUCT": ["Subject", "Page Index", "Label", "Count", "DUCT SIZE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "Length", "Length Unit", "DAMPER TYPE"]},
                    "category_counts": {"rectangular": len(phi_records)}, "total_records": len(phi_records)}
            else:
                detection = universal_input.normalize_file(source, csv_input)
        except (ValueError, OSError) as exc:
            raise Stage1InputError(f"Manual structure detection failed: {exc}") from exc
        output = staging / f"{stem}_VJR.xlsx"
        report = staging / f"{stem}_validation_report.md"
        try:
            result = manual_csv_to_vjr.generate(csv_input, template, output, report)
        except (manual_to_vjr.WriterError, OSError, ValueError) as exc:
            raise Stage1InputError(f"VJR generation stopped: {exc}") from exc
        if not output.is_file() or output.stat().st_size == 0:
            raise Stage1InputError("Generation did not produce a valid VJR workbook.")
        verify_book = None
        try:
            verify_book = openpyxl.load_workbook(output, read_only=True, data_only=False)
            if not {"TAKE OFF", "SPIRALE", "ROND ET FLEX"}.issubset(verify_book.sheetnames):
                raise Stage1InputError("The generated VJR is missing required worksheets.")
        except Stage1InputError:
            raise
        except Exception as exc:
            raise Stage1InputError(f"Generated VJR validation failed: {exc}") from exc
        finally:
            if verify_book is not None:
                verify_book.close()
        if _sha256(source) != input_hash:
            raise Stage1InputError("The uploaded Manual changed during processing; output was discarded.")
        if _sha256(template) != template_hash:
            raise Stage1InputError("The fixed VJR template changed during processing; output was discarded.")
        output_hash = _sha256(output)

        if csv_input != source:
            csv_input.unlink(missing_ok=True)
        with report.open("a", encoding="utf-8") as validation:
            validation.write("\n## Frontend input provenance\n\n")
            validation.write(f"- Processing run ID: `{job_id}`\n")
            validation.write(f"- Original input: `{source.name}` ({source.suffix.casefold()})\n")
            validation.write(f"- Original input SHA-256: `{input_hash}`\n")
            validation.write(f"- Generated VJR: `{output.name}`\n")
            validation.write(f"- Generated VJR SHA-256: `{output_hash}`\n")
            validation.write(f"- Fixed VJR template: `{template.name}`\n")
            validation.write(f"- Fixed VJR template SHA-256 before/after: `{template_hash}` / `{_sha256(template)}`\n")
            validation.write(f"- Detected input format: **{detection['format']}**.\n")
            validation.write(f"- Detected data sheet(s): **{', '.join(detection['sheets'])}**.\n")
            validation.write(f"- Detected source columns: `{json.dumps(detection['columns'], ensure_ascii=False)}`.\n")
            validation.write(f"- Canonical categories: `{json.dumps(detection['category_counts'], ensure_ascii=False)}`.\n")
            validation.write(f"- Original DATA records/categories before authoritative grouping: **{detection.get('detected_input_records', detection['total_records'])}** / `{json.dumps(detection.get('detected_input_category_counts', detection['category_counts']), ensure_ascii=False)}`.\n")
            validation.write(f"- Canonical input records passed to the existing mapper: **{detection['total_records']}**. Missing source TYPE remains blank; input TYPE is not a required field.\n")
            selection = detection.get("source_selection")
            if selection:
                validation.write("\n## Evidence-based authoritative source selection\n\n")
                validation.write(f"- Rectangular authority: `{selection['authoritative_sheet']}`; detailed cross-reference: `{selection['detail_sheet']}`.\n")
                validation.write(f"- Detailed rectangular records: **{selection['detail_records']}**; grouped authority records: **{selection['grouped_records']}**.\n")
                validation.write(f"- Length reconciliation: detail **{selection['detail_length_total']}**; grouped **{selection['grouped_length_total']}**; {selection['reconciliation']}.\n")
                validation.write(f"- Explicit pressure evidence in authority sheet: `{json.dumps(selection['pressure_counts'])}`.\n")
                validation.write("- Detailed rectangular rows were cross-referenced and excluded from mapper input; only the reconciled grouped rows were sent for rectangular mapping.\n")
                validation.write("- Material was unavailable; Metal Code D was left blank. Installation F remains calculated by the existing dimensional rule.\n")
                validation.write("- Canvas associations without a matching system/size row in the detailed accessory sheet were withheld from VJR accessory cells and sent for review.\n")
            validation.write("- Mapping engine: existing `manual_csv_to_vjr.generate`; confirmed rectangular and spiral riser mapping rules are applied by the existing pipeline.\n")

        if detection.get("source_selection"):
            _apply_source_reviews(staging / f"{csv_input.stem}_record_audit.csv",
                                  staging / f"{csv_input.stem}_unresolved.csv",
                                  detection["source_selection"].get("source_review_items", []))
        summary = _summary(result, staging / f"{csv_input.stem}_record_audit.csv",
                           staging / f"{csv_input.stem}_unresolved.csv")
        # Fail closed if preservation checks are missing or do not report zero
        # calculation-formula changes.
        report_text = report.read_text(encoding="utf-8")
        if ("Calculation formulas modified: **0**" not in report_text
                or "Input styles, merged cells, validations, drawings, and worksheet structure preserved." not in report_text
                or "Shared-formula masters and followers audited; touched groups were safely expanded." not in report_text):
            raise Stage1InputError("Workbook preservation checks did not pass; output was discarded.")

        os.replace(staging, final_dir)
        staging = None
        return {
            **summary,
            "job_id": job_id,
            "run_id": job_id,
            "input_name": source.name,
            "input_sha256": input_hash,
            "template_name": template.name,
            "detected_format": detection["format"],
            "detected_sheets": detection["sheets"],
            "category_counts": detection["category_counts"],
            "detected_input_records": detection.get("detected_input_records", detection["total_records"]),
            "detected_input_category_counts": detection.get("detected_input_category_counts", detection["category_counts"]),
            "supported_records": summary["records_with_writes"],
            "formula_changes": 0,
            "structure_preserved": True,
            "output_sha256": output_hash,
            "output": final_dir / output.name,
            "report": final_dir / report.name,
            "audit": final_dir / f"{csv_input.stem}_record_audit.csv",
            "unresolved": final_dir / f"{csv_input.stem}_unresolved.csv",
            "comparison": final_dir / f"{csv_input.stem}_reference_comparison.csv",
        }
    except Exception:
        if staging is not None and staging.exists() and staging.parent == root:
            shutil.rmtree(staging)
        raise
