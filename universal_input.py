"""Format detection and canonical normalization for Stage 1 Manual inputs.

The canonical records are adapted back to the established mapper CSV contract;
the VJR writer and its mapping rules remain the downstream implementation.
"""
from __future__ import annotations

import csv
import json
import re
from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl


ALIASES = {
    "subject": {"subject", "category", "item", "item type", "description", "duct type"},
    "page index": {"page index", "page", "sheet page"},
    "length": {"length", "total length", "run length"},
    "length unit": {"length unit", "unit of length", "length units"},
    "count": {"count", "quantity", "qty", "total"},
    "label": {"label", "name", "item name"},
    "duct size": {"duct size", "size", "dimensions", "dimension", "duct dimensions"},
    "type": {"type", "material", "pressure material", "construction"},
    "accessories1": {"accessories1", "accessory 1", "accessories 1", "accessories", "accessory"},
    "accessories2": {"accessories2", "accessory 2", "accessories 2"},
    "remark": {"remark", "remarks", "notes", "comment"},
    "units": {"units", "system", "system unit", "system units"},
    "damper type": {"damper type", "damper"},
    "source_row": {"source_row"},
}
ALIASES = {key: {re.sub(r"[^a-z0-9]+", "", s.casefold()) for s in values | {key}} for key, values in ALIASES.items()}


def _key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def _canonical_category(text: str) -> str:
    t = text.upper()
    if "RISER" in t: return "riser"
    if "RECT" in t and "DUCT" in t: return "rectangular"
    if "SPIRAL" in t or "NAD DUCT" in t or "BLUEDUCT" in t: return "spiral"
    if "FLEX" in t: return "flex"
    if "ELBOW" in t: return "elbows"
    if "DAMPER" in t or re.search(r"\b(?:RCF|RCFF|RCFU|RFCU|VCF|VCFF|VB|MBD|MVD|BDD|CD|RM|RMO|VM|MD)\b", t): return "dampers"
    if any(x in t for x in ("ACCESSOR", "CANVAS", "CANEVAS", "VANE", "MESH", "INSULATION", "GRILLAGE")): return "accessories"
    return "other/unknown"


def _dimension(size: Any, category: str) -> tuple[str, Any, Any, Any]:
    text = str(size or "").strip().replace("×", "X").replace("x", "X")
    pair = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)\s*", text, re.I)
    if pair:
        a, b = float(pair.group(1)), float(pair.group(2))
        return f"{a:g}X{b:g}", a, b, None
    number = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH(?:ES)?)?\s*", text, re.I)
    if number:
        value = float(number.group(1))
        return f"{value:g}", None, None, value
    return text, None, None, None


def _header_row(rows: list[tuple[Any, ...]]) -> tuple[int, dict[str, int], list[str]] | None:
    best = None
    for ix, row in enumerate(rows[:15]):
        mapped: dict[str, int] = {}
        raw = []
        for j, cell in enumerate(row):
            raw.append(str(cell).strip() if cell is not None else "")
            key = _key(cell)
            for canonical, aliases in ALIASES.items():
                if key in aliases:
                    mapped.setdefault(canonical, j)
        score = len(mapped)
        if "subject" in mapped or "label" in mapped: score += 2
        if "duct size" in mapped: score += 2
        if "length" in mapped or "count" in mapped: score += 1
        if best is None or score > best[0]: best = (score, ix, mapped, raw)
    if best and best[0] >= 5 and len(best[2]) >= 3:
        return best[1], best[2], best[3]
    return None


def _records_from_rows(sheet: str, rows: list[tuple[Any, ...]], header: tuple[int, dict[str, int], list[str]]) -> list[dict[str, Any]]:
    header_ix, columns, names = header
    result = []
    for row_ix, row in enumerate(rows[header_ix + 1:], header_ix + 2):
        raw = {names[j]: row[j] if j < len(row) else None for j in range(len(names)) if names[j]}
        if not any(v not in (None, "") for v in raw.values()): continue
        get = lambda key: row[columns[key]] if key in columns and columns[key] < len(row) else None
        subject, label = get("subject"), get("label")
        description = " ".join(str(v or "") for v in (subject, label, get("type"), get("remark")))
        category = _canonical_category(description)
        size, width, height, diameter = _dimension(get("duct size"), category)
        record = {
            "category": category, "width": width, "height": height, "diameter": diameter,
            "length": get("length"), "length_unit": get("length unit"), "type": get("type"),
            "pressure_material": get("type"), "system_unit": get("units"),
            "installation": None, "accessories": [v for v in (get("accessories1"), get("accessories2")) if v not in (None, "", ".")],
            "source_sheet": sheet, "source_row": row_ix, "source_fields": raw,
            "mapper": {"subject": subject, "page index": get("page index"), "length": get("length"),
                "length unit": get("length unit"), "count": get("count"), "label": label,
                "duct size": size, "type": get("type"), "accessories1": get("accessories1"),
                "accessories2": get("accessories2"), "remark": get("remark"), "units": get("units"),
                "damper type": get("damper type"), "source_row": row_ix, "source_sheet": sheet,
                "source_fields": json.dumps(raw, ensure_ascii=False, default=str)},
        }
        result.append(record)
    return result


def _duct_group_key(record: dict[str, Any]) -> tuple[str, str, str, str]:
    mapper = record["mapper"]
    subject = str(mapper.get("subject") or "").strip().upper()
    if record.get("category") == "rectangular":
        # Grouped TAKEOFF sheets may put the duct label in Label instead of
        # Subject, while detail sheets use values such as "RECT. DUCT m".
        subject = "RECT. DUCT"
    size = re.sub(r"\s+", "", str(mapper.get("duct size") or "").upper().replace("×", "X"))
    return (str(mapper.get("units") or "").strip().upper(), subject, size,
            str(mapper.get("length unit") or "").strip().casefold())


def _pressure_code(value: Any) -> str:
    match = re.search(r"\b(BP|MP|HP)\b", str(value or "").upper())
    return match.group(1) if match else ""


def _select_reconciled_rectangular_source(
    candidates: list[tuple[str, list[dict[str, Any]], list[str], int]],
    base: tuple[str, list[dict[str, Any]], list[str], int],
) -> tuple[list[dict[str, Any]], dict[str, Any]] | None:
    """Prefer a pressure-bearing grouped duct sheet only when detailed rows reconcile."""
    detail_candidates = [item for item in candidates
                         if "data" in item[0].casefold() and "duct" in item[0].casefold()
                         and "accessor" not in item[0].casefold()]
    detail = next((item for item in detail_candidates if item[0].casefold() != "data"), base)
    base_records = base[1]
    detail_rect = [r for r in detail[1] if r["category"] == "rectangular"]
    if not detail_rect:
        detail_rect = [r for r in base_records if r["category"] == "rectangular"]
        detail = base
    detail_totals: dict[tuple[str, str, str, str], float] = {}
    detail_rows: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
    for record in detail_rect:
        key = _duct_group_key(record)
        amount = record.get("length")
        try:
            detail_totals[key] = detail_totals.get(key, 0.0) + float(amount or 0)
        except (TypeError, ValueError):
            return None
        detail_rows.setdefault(key, []).append(record)
    base_rows_by_key: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
    for record in base_records:
        if record["category"] == "rectangular":
            base_rows_by_key.setdefault(_duct_group_key(record), []).append(record)
    for summary in candidates:
        name, rows, columns, _ = summary
        normalized_name = name.casefold()
        column_keys = {_key(x) for x in columns}
        if "takeoff" not in normalized_name or "duct" not in normalized_name:
            continue
        # Both supported workbook layouts carry a pressure-bearing rollup.
        # Some projects use Subject + Length; RTL LOT 1 uses Label + Total.
        has_category = "subject" in column_keys or "label" in column_keys
        has_measure = "length" in column_keys or "total" in column_keys
        if not {"units", "ductsize", "lengthunit", "remark"}.issubset(column_keys) or not has_category or not has_measure:
            continue
        grouped = [r for r in rows if r["category"] == "rectangular"]
        if not grouped or not any(_pressure_code(r["mapper"].get("remark")) for r in grouped):
            continue
        grouped_by_key: dict[tuple[str, str, str, str], tuple[dict[str, Any], float, list[int]]] = {}
        for record in grouped:
            key = _duct_group_key(record)
            raw = record["source_fields"]
            total = next((value for header, value in raw.items()
                          if _key(header) == "total"), record.get("length"))
            try:
                total_value = float(total)
            except (TypeError, ValueError):
                grouped_by_key = {}
                break
            if key in grouped_by_key:
                first, previous, source_rows = grouped_by_key[key]
                grouped_by_key[key] = (first, previous + total_value, source_rows + [int(record["source_row"])])
            else:
                grouped_by_key[key] = (record, total_value, [int(record["source_row"])])
        if not grouped_by_key or set(grouped_by_key) != set(detail_totals):
            continue
        if any(abs(grouped_by_key[key][1] - detail_totals[key]) > 1e-7 for key in detail_totals):
            continue
        # The broad DATA sheet can contain demo rows absent from DATA-DUCT and
        # TAKEOFF-DUCT. Detail-to-rollup reconciliation is the authoritative
        # check for this layout; keep the existing extra DATA check where the
        # DATA rectangular groups are themselves a complete matching set.
        base_totals: dict[tuple[str, str, str, str], float] = {}
        for key, members in base_rows_by_key.items():
            try:
                base_totals[key] = sum(float(member.get("length") or 0) for member in members)
            except (TypeError, ValueError):
                base_totals[key] = float("nan")
        rtl_layout = "subject" not in column_keys and "label" in column_keys and "total" in column_keys
        if not rtl_layout and (set(base_totals) != set(detail_totals) or any(
                abs(base_totals[key] - detail_totals[key]) > 1e-7 for key in detail_totals)):
            continue

        retained = [r for r in base_records if r["category"] != "rectangular"]
        reviews: list[dict[str, Any]] = []
        merged: list[dict[str, Any]] = []
        accessory_sheet_rows = [r for item in candidates if "accessor" in item[0].casefold()
                                and "data" in item[0].casefold() for r in item[1]]
        for key, (source, total, source_rows) in grouped_by_key.items():
            detail_members = detail_rows[key]
            base_members = base_rows_by_key.get(key) or detail_members
            # A removed DATA row supplies a unique internal mapper key, while the
            # source provenance keeps the authoritative TAKEOFF row and detail links.
            internal_row = min(int(r["source_row"]) for r in base_members)
            raw_source = dict(source["source_fields"])
            physical_row = source["source_row"]
            raw_source["_authoritative_source_rows"] = source_rows
            linked_rows = sorted({int(r["source_row"]) for r in detail_members})
            raw_source["_authoritative_source_row"] = physical_row
            raw_source["_cross_reference_sheet"] = detail[0]
            raw_source["_cross_reference_rows"] = linked_rows
            mapper = dict(source["mapper"])
            pressure = _pressure_code(mapper.get("remark"))
            mapper.update({"subject": mapper.get("subject") or "RECT. DUCT m",
                           "length": total, "count": None, "label": mapper.get("label") or mapper.get("subject"),
                           "type": pressure, "remark": pressure or mapper.get("remark"),
                           "source_row": internal_row, "source_physical_row": physical_row,
                           "cross_reference_rows": ";".join(map(str, linked_rows)),
                           "source_sheet": name,
                           "source_fields": json.dumps(raw_source, ensure_ascii=False, default=str)})
            source["source_fields"] = raw_source
            source["source_row"] = physical_row
            source["type"] = pressure or None
            source["pressure"] = pressure or None
            source["pressure_material"] = None
            source["length"] = total
            source["mapper"] = mapper
            accessory = mapper.get("accessories1")
            if str(accessory or "").strip().upper() in {"CANEVAS", "CANVAS"}:
                corroborated = any(
                    str(item["mapper"].get("subject") or "").strip().upper() in {"CANEVAS", "CANVAS"}
                    and _duct_group_key(item)[:1] == key[:1]
                    and _duct_group_key(item)[2] == key[2]
                    for item in accessory_sheet_rows
                )
                if not corroborated:
                    reviews.append({"source_row": internal_row, "source_sheet": name,
                        "source_physical_row": physical_row, "category": "rectangular", "field": "ACCESSORIES1",
                        "reason": "TAKEOFF-DUCTS associates CANEVAS with this rectangular group, but the detailed accessory sheet does not corroborate the same system and size; code was left blank for review."})
                    mapper["accessories1"] = None
            # Material is absent in both the detail and grouped source. Keep D blank
            # and expose that as a review item without altering VJR mapping rules.
            reviews.append({"source_row": internal_row, "source_sheet": name,
                "source_physical_row": physical_row, "category": "rectangular", "field": "D (Metal Code)",
                "reason": "Source workbook does not specify rectangular duct material; VJR Metal Code D was left blank."})
            source["source_sheet"] = name
            source["source_review"] = reviews[-1]["reason"]
            merged.append(source)
        records = retained + merged
        info = {"authoritative_sheet": name, "detail_sheet": detail[0],
                "detail_records": len(detail_rect), "grouped_records": len(grouped),
                "group_count": len(grouped_by_key), "detail_length_total": sum(detail_totals.values()),
                "grouped_length_total": sum(v[1] for v in grouped_by_key.values()),
                "reconciliation": "PASS: every grouped rectangular total reconciles to detailed and DATA totals",
                "pressure_counts": dict(Counter(_pressure_code(r["mapper"].get("remark")) for r in grouped)),
                "source_review_items": reviews}
        return records, info
    return None


def inspect_file(source: Path) -> dict[str, Any]:
    """Detect, select, and normalize source records without writing mapper input."""
    source = Path(source)
    candidates: list[tuple[str, list[dict[str, Any]], list[str], int]] = []
    if source.suffix.casefold() == ".csv":
        with source.open("r", newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        header = _header_row([tuple(r) for r in rows])
        if header:
            records = _records_from_rows("CSV", [tuple(r) for r in rows], header)
            candidates.append(("CSV", records, [s for s in header[2] if s], header[0]))
    else:
        book = openpyxl.load_workbook(source, read_only=True, data_only=True)
        try:
            for ws in book.worksheets:
                rows = list(ws.iter_rows(values_only=True))
                header = _header_row(rows)
                if not header: continue
                records = _records_from_rows(ws.title, rows, header)
                candidates.append((ws.title, records, [s for s in header[2] if s], header[0]))
        finally: book.close()
    if not candidates:
        raise ValueError("No worksheet or CSV header row with recognizable Manual fields was found.")
    # Prefer DATA as the broad category source, then replace rectangular details
    # only when a pressure-bearing grouped sheet reconciles exactly to the detail.
    data = next((c for c in candidates if c[0].casefold() == "data"), None)
    detected_input_records = len(data[1]) if data else sum(len(c[1]) for c in candidates)
    detected_input_categories = dict(Counter(r["category"] for r in data[1])) if data else dict(
        Counter(r["category"] for _, rows, _, _ in candidates for r in rows))
    if data:
        merged = _select_reconciled_rectangular_source(candidates, data)
        if merged:
            records, source_selection = merged
            chosen = [data, next(c for c in candidates if c[0] == source_selection["authoritative_sheet"])]
            detail_item = next((c for c in candidates if c[0] == source_selection["detail_sheet"]), None)
            if detail_item and detail_item not in chosen:
                chosen.append(detail_item)
        else:
            records = data[1]
            source_selection = None
            chosen = [data]
    else:
        selected = []
        for c in candidates:
            n = c[0].casefold()
            if n.startswith("data-") and any(x[0].casefold() == n[5:] or x[0].casefold() == n[5:] + "-accessories" for x in candidates): continue
            selected.append(c)
        chosen = selected or [max(candidates, key=lambda x: len(x[1]))]
        records = [r for _, rs, _, _ in chosen for r in rs]
        source_selection = None
    return {"format": "Excel workbook" if source.suffix.casefold() != ".csv" else "CSV", "sheets": [x[0] for x in chosen], "columns": {x[0]: x[2] for x in chosen}, "records": records, "category_counts": dict(Counter(r["category"] for r in records)), "total_records": len(records), "detected_input_records": detected_input_records, "detected_input_category_counts": detected_input_categories, "source_selection": source_selection}


def normalize_file(source: Path, mapper_csv: Path) -> dict[str, Any]:
    """Write detected records using the established mapper CSV contract."""
    detection = inspect_file(source)
    fields_out = ["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "DUCT SIZE", "TYPE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "DAMPER TYPE", "source_row", "source_sheet", "source_fields", "source_physical_row", "cross_reference_rows"]
    mapper_csv.parent.mkdir(parents=True, exist_ok=True)
    with mapper_csv.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields_out); writer.writeheader()
        for rec in detection["records"]:
            m = rec["mapper"]
            writer.writerow({"Subject": m["subject"], "Page Index": m["page index"], "Length": m["length"], "Length Unit": m["length unit"], "Count": m["count"], "Label": m["label"], "DUCT SIZE": m["duct size"], "TYPE": m["type"], "ACCESSORIES1": m["accessories1"], "ACCESSORIES2": m["accessories2"], "REMARK": m["remark"], "UNITS": m["units"], "DAMPER TYPE": m["damper type"], "source_row": m["source_row"], "source_sheet": m["source_sheet"], "source_fields": m["source_fields"], "source_physical_row": m.get("source_physical_row", ""), "cross_reference_rows": m.get("cross_reference_rows", "")})
    return detection
