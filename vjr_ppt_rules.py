"""Pure business-rule helpers transcribed from VENTILATION_DUCTING.pptx.

This module makes decisions about individual takeoff facts. It does not read or
write workbooks; callers remain responsible for workbook layout and protection.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable, Sequence


SAFE = "SAFE"
NEEDS_CONFIRMATION = "NEEDS_CONFIRMATION"
UNSUPPORTED = "UNSUPPORTED"


@dataclass(frozen=True)
class RuleResult:
    status: str
    reason: str
    source_rule: str
    value: Any = None
    details: dict[str, Any] | None = None


def _d(value: Any) -> Decimal:
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise ValueError(f"Not a numeric value: {value!r}") from exc


def rectangular_dimensions(width: Any, height: Any, unit: str = "mm") -> RuleResult:
    """Return the larger dimension first; retain the supplied dimension unit."""
    w, h = _d(width), _d(height)
    if w <= 0 or h <= 0:
        return RuleResult(NEEDS_CONFIRMATION, "Width and height must be positive.", "PPT slide 4: include duct size as shown; slide 28: provide both dimensions", details={"width": str(w), "height": str(h)})
    if unit.lower() not in {"mm", "inch", "in", "inches"}:
        return RuleResult(NEEDS_CONFIRMATION, "Dimension unit is not established.", "PPT slides 4 and 28", details={"unit": unit})
    a, b = max(w, h), min(w, h)
    return RuleResult(SAFE, "Dimensions ordered larger first; dimensions and unit preserved.", "PPT slide 5: larger dimension on left; slide 28: metric/imperial size tags", (a, b, unit))


def pressure_type(pressure: str | None) -> RuleResult:
    mapping = {"BP": "low pressure", "MP": "medium pressure", "HP": "high pressure"}
    if pressure is None or pressure.strip().upper() not in mapping:
        return RuleResult(NEEDS_CONFIRMATION, "Pressure/type is not one of the documented BP, MP, or HP values.", "PPT slide 30: customer notes specify BP/MP/HP", details={"input": pressure})
    key = pressure.strip().upper()
    return RuleResult(SAFE, f"{key} means {mapping[key]}.", "PPT slide 30: BP low, MP medium, HP high", mapping[key])


def installation(width: Any, height: Any, unit: str = "inch", *, is_riser: bool = False) -> RuleResult:
    if is_riser:
        return RuleResult(SAFE, "Risers use RI installation.", "PPT slide 30: for risers select RI", "RI")
    if unit.lower() not in {"inch", "in", "inches", "mm"}:
        return RuleResult(NEEDS_CONFIRMATION, "Dimension unit is unknown; the 36-inch threshold cannot be evaluated.", "PPT slide 30: 36-inch installation threshold")
    w, h = _d(width), _d(height)
    if unit.lower() == "mm":
        w, h = w / Decimal("25.4"), h / Decimal("25.4")
    if w == 36 or h == 36:
        return RuleResult(NEEDS_CONFIRMATION, "A duct dimension is exactly 36 inches; the PPT does not specify the boundary case.", "PPT slide 30: below 36 inches => HA; exceeds 36 inches => RE", details={"width_inches": str(w), "height_inches": str(h)})
    if w > 36 or h > 36:
        return RuleResult(SAFE, "At least one duct dimension exceeds 36 inches.", "PPT slide 30: any width or height exceeding 36 inches => RE", "RE")
    return RuleResult(SAFE, "Both duct dimensions are below 36 inches.", "PPT slide 30: all duct sizes below 36 inches => HA", "HA")


ACCESSORY_RULES = {
    1: ("1/2-inch acoustic insulation", "PPT slide 26"),
    21: ("1-inch acoustic insulation", "PPT slide 26"),
    7: ("T-vane", "PPT slides 7 and 26"),
    71: ("T-vane for duct size greater than 30 inches", "PPT slide 26"),
    8: ("canvas", "PPT slides 12 and 26"),
    10: ("fire damper", "PPT slide 26"),
    100: ("fire-smoke damper", "PPT slide 26"),
    11: ("motorized damper", "PPT slide 26"),
    12: ("gravity damper", "PPT slide 26"),
    13: ("manual balancing damper", "PPT slide 26"),
    131: ("manual balancing damper for duct size greater than 30 inches", "PPT slide 26"),
    14: ("mesh", "PPT slide 26"),
}


def accessory(code: int) -> RuleResult:
    if code == 9:
        return RuleResult(UNSUPPORTED, "Code 9 is described only for MBD on rigid metal round duct; its rectangular representation is not specified.", "PPT slide 18")
    if code not in ACCESSORY_RULES:
        return RuleResult(UNSUPPORTED, "No mapping is documented for this accessory code.", "PPT slide 26 lists supported codes", details={"code": code})
    label, source = ACCESSORY_RULES[code]
    return RuleResult(SAFE, f"Accessory code {code}: {label}.", source, code, {"label": label})


def split_accessories(codes: Sequence[int], *, riser: bool = False) -> RuleResult:
    invalid = [c for c in codes if c not in ACCESSORY_RULES]
    if invalid:
        return RuleResult(UNSUPPORTED, "One or more accessory codes lack a documented mapping.", "PPT slide 26 lists supported codes", details={"unsupported_codes": invalid})
    chunks = [list(codes[i:i + 4]) for i in range(0, len(codes), 4)] or [[]]
    rows = []
    for i, chunk in enumerate(chunks):
        rows.append({"accessories": chunk, "length": 0 if i > 0 else None, "repeat_size": i > 0, "installation": "RI" if riser else None})
    if len(codes) > 4:
        return RuleResult(SAFE, "Accessories fit four per row; continuation rows repeat the duct size and use zero length.", "PPT slide 29: repeat size, zero total length, continue accessories", rows)
    return RuleResult(SAFE, "Accessories fit within the four documented columns.", "PPT rule: up to four VJR accessory columns", rows)


def rectangular_riser(width: Any, height: Any, length_m: Any, accessories: Sequence[int] = (), unit: str = "mm") -> RuleResult:
    dims = rectangular_dimensions(width, height, unit)
    if dims.status != SAFE:
        return dims
    try:
        length = _d(length_m)
    except ValueError:
        return RuleResult(NEEDS_CONFIRMATION, "Riser length in meters is required.", "PPT slides 14, 29")
    if length < 0:
        return RuleResult(NEEDS_CONFIRMATION, "Riser length cannot be negative.", "PPT slides 14, 29")
    split = split_accessories(accessories, riser=True)
    if split.status != SAFE:
        return split
    rows = split.value
    rows[0].update({"width": dims.value[0], "height": dims.value[1], "unit": unit, "length_m": length, "installation": "RI"})
    return RuleResult(SAFE, "Rectangular riser uses TAKE OFF, RI, preserved size, meter length; no extra riser is generated.", "PPT slides 14, 15, 29", rows)


def spiral_riser_drop(diameter: Any, drop_ft: Any, *, already_included: bool = False) -> RuleResult:
    if already_included:
        return RuleResult(NEEDS_CONFIRMATION, "The riser is already represented; avoid duplicate counting and verify the existing included drop.", "PPT slide 15: include riser for only one floor to avoid duplicate counting")
    d, length = _d(diameter), _d(drop_ft)
    if d <= 0 or length < 0:
        return RuleResult(NEEDS_CONFIRMATION, "Diameter must be positive and drop length nonnegative.", "PPT slide 32")
    return RuleResult(SAFE, "Add riser drop length in feet to the spiral duct column for this diameter.", "PPT slide 32", {"diameter": d, "add_length_ft": length})


def spiral_length_by_diameter(records: Iterable[tuple[Any, Any]], unit: str = "ft") -> RuleResult:
    """Aggregate lengths, preserving exact diameter; convert mm and meters to inches/feet."""
    totals: dict[Decimal, Decimal] = {}
    unit_key = unit.lower()
    for diameter, length in records:
        d, l = _d(diameter), _d(length)
        if d <= 0 or l < 0:
            return RuleResult(NEEDS_CONFIRMATION, "Spiral diameter must be positive and length nonnegative.", "PPT slide 31")
        if unit_key == "mm":
            d = d / Decimal("25.4")
            l = l * Decimal("3.280839895")
        elif unit_key == "m":
            l = l * Decimal("3.280839895")
        elif unit_key not in {"ft", "inch", "in", "inches"}:
            return RuleResult(NEEDS_CONFIRMATION, "Unsupported or ambiguous spiral units.", "PPT slide 31")
        totals[d] = totals.get(d, Decimal(0)) + l
    return RuleResult(SAFE, "Lengths are totaled by exact diameter; metric values are converted to inches/feet.", "PPT slide 31", dict(sorted(totals.items())))


def spiral_elbows(degrees: int, count: int, diameter: Any) -> RuleResult:
    if degrees not in (45, 90):
        return RuleResult(UNSUPPORTED, "Only 45- and 90-degree spiral elbows are documented.", "PPT slides 9 and 33")
    if count < 0 or _d(diameter) <= 0:
        return RuleResult(NEEDS_CONFIRMATION, "Elbow count must be nonnegative and diameter positive.", "PPT slides 9 and 33")
    return RuleResult(SAFE, f"Record {count} spiral {degrees}-degree elbows at the supplied diameter.", "PPT slides 9 and 33", {"degrees": degrees, "count": count, "diameter": _d(diameter)})


def y_type(connecting_ducts: Sequence[Any], elbow_45_count: int) -> RuleResult:
    if len(connecting_ducts) != 2 or any(x in (None, "") for x in connecting_ducts) or elbow_45_count != 2:
        return RuleResult(NEEDS_CONFIRMATION, "Y-type requires two connecting ducts and two 45-degree elbows; missing dimensions/allocation are not inferred.", "PPT slide 10", details={"connecting_ducts": list(connecting_ducts), "elbow_45_count": elbow_45_count})
    return RuleResult(SAFE, "Two connecting ducts and two 45-degree elbows are explicitly supplied.", "PPT slide 10", {"connecting_ducts": list(connecting_ducts), "elbow_45_count": 2})


def flex_length(count: int, *, kind: str = "ordinary", client_confirmed: bool = False, equipment_type: str | None = None) -> RuleResult:
    if count < 0:
        return RuleResult(NEEDS_CONFIRMATION, "Flex count cannot be negative.", "PPT slides 36–38")
    if kind.lower() == "ordinary":
        return RuleResult(SAFE, "Ordinary flex is five feet per count.", "PPT slide 36", Decimal(count) * Decimal(5))
    if kind.lower() == "erv":
        if not client_confirmed:
            return RuleResult(NEEDS_CONFIRMATION, "ERV flex is ten feet per connection only after client confirmation.", "PPT slide 13")
        return RuleResult(SAFE, "Confirmed ERV flex is ten feet per connection.", "PPT slide 13", Decimal(count) * Decimal(10))
    if kind.lower() == "equipment":
        if not equipment_type:
            return RuleResult(NEEDS_CONFIRMATION, "Equipment-specific flex type requires VJR/client confirmation.", "PPT slide 38")
        return RuleResult(NEEDS_CONFIRMATION, "Equipment flex type must follow the confirmed equipment-specific VJR instruction.", "PPT slide 38", details={"equipment_type": equipment_type})
    return RuleResult(UNSUPPORTED, "Flex category is not documented.", "PPT slides 36–38")


def canvas_eligibility(*, direct_to_fan_ahu_rtu: bool | None, kitchen_hood_or_dryer: bool | None) -> RuleResult:
    if direct_to_fan_ahu_rtu is None or kitchen_hood_or_dryer is None:
        return RuleResult(NEEDS_CONFIRMATION, "Canvas inference requires explicit direct-connection and kitchen hood/dryer exception facts.", "PPT slides 5 and 12")
    if not direct_to_fan_ahu_rtu or kitchen_hood_or_dryer:
        return RuleResult(UNSUPPORTED, "Automatic canvas addition is not eligible under the supplied connection/exception facts.", "PPT slides 5 and 12")
    return RuleResult(SAFE, "Direct connection to fan/AHU/RTU is established and kitchen hood/dryer exception is false; use accessory 8.", "PPT slides 12 and 26", 8)


def mbd(*, explicit: bool, duct_kind: str, size_in: Any | None = None, size_available: bool = True) -> RuleResult:
    if not explicit:
        return RuleResult(NEEDS_CONFIRMATION, "MBD is not inferred from topology; explicit association is required.", "PPT slide 18")
    kind = duct_kind.lower()
    if kind == "rectangular":
        if size_in is None:
            return RuleResult(NEEDS_CONFIRMATION, "Rectangular size is required to choose code 13 or 131.", "PPT slide 26")
        code = 131 if _d(size_in) > 30 else 13
        return RuleResult(SAFE, f"Explicit rectangular MBD maps to code {code} by the greater-than-30-inch rule.", "PPT slide 26", code)
    if kind == "spiral":
        return RuleResult(SAFE, "Explicit spiral MBD belongs in the SPIRALE worksheet damper column.", "PPT slide 18", "SPIRALE.damper")
    if kind in {"rigid_round", "rigid metal round"}:
        if not size_available:
            return RuleResult(NEEDS_CONFIRMATION, "PPT calls for rectangular accessory code 9, but the required rectangular size representation is unavailable.", "PPT slide 18")
        return RuleResult(UNSUPPORTED, "Code 9 representation is not defined by the PPT accessory code list; do not emit an assumed rectangular row.", "PPT slide 18")
    return RuleResult(UNSUPPORTED, "Duct kind has no documented MBD destination.", "PPT slides 17–18")


def spiral_nonbalancing_damper(*, width: Any | None = None, height: Any | None = None, damper_code: int | None = None) -> RuleResult:
    if width is None or height is None:
        return RuleResult(NEEDS_CONFIRMATION, "Spiral non-balancing damper converts to rectangular with zero length, but rectangular width and height are not supplied.", "PPT slide 17")
    if damper_code not in {10, 100, 11, 12}:
        return RuleResult(NEEDS_CONFIRMATION, "Confirm the explicit non-balancing damper accessory code before conversion.", "PPT slides 16–17")
    dims = rectangular_dimensions(width, height, "inch")
    if dims.status != SAFE:
        return dims
    return RuleResult(SAFE, "Convert to rectangular, use zero duct length, and attach the explicitly supplied damper accessory.", "PPT slide 17", {"width": dims.value[0], "height": dims.value[1], "length": 0, "accessory": damper_code})


def system_assignment(system: str | None, accessories: Sequence[int] = ()) -> RuleResult:
    if not system or not system.strip():
        return RuleResult(NEEDS_CONFIRMATION, "System/unit is missing; it cannot be inferred from duct size.", "PPT slide 27")
    split = split_accessories(accessories)
    if split.status != SAFE:
        return split
    return RuleResult(SAFE, "Use the explicitly supplied system/unit and keep it separate with its corresponding accessories.", "PPT slide 27", {"system": system.strip(), "accessory_rows": split.value})


def rectangular_record(*, width: Any, height: Any, length: Any, unit: str,
                       metal_material_code: Any, pressure: str | None,
                       system: str | None, accessories: Sequence[int] = (),
                       is_riser: bool = False) -> RuleResult:
    """Validate a fully explicit rectangular record without guessing codes."""
    missing = [name for name, value in (
        ("width", width), ("height", height), ("length", length),
        ("metal/material code", metal_material_code), ("pressure/type", pressure),
        ("system/unit", system),
    ) if value is None or value == ""]
    if missing:
        return RuleResult(NEEDS_CONFIRMATION, "Required rectangular fields are missing: " + ", ".join(missing), "PPT slides 27–30")
    try:
        length_value = _d(length)
    except ValueError:
        return RuleResult(NEEDS_CONFIRMATION, "Duct length must be numeric and explicitly supplied.", "PPT slide 28")
    if length_value < 0:
        return RuleResult(NEEDS_CONFIRMATION, "Duct length cannot be negative.", "PPT slide 28")
    dims = rectangular_dimensions(width, height, unit)
    if dims.status != SAFE:
        return dims
    typ = pressure_type(pressure)
    if typ.status != SAFE:
        return typ
    install = installation(width, height, unit, is_riser=is_riser)
    if install.status != SAFE:
        return install
    system_result = system_assignment(system, accessories)
    if system_result.status != SAFE:
        return system_result
    return RuleResult(SAFE, "All required rectangular fields were explicitly supplied and validated; material code is preserved without deriving a gauge.", "PPT slides 27–30", {
        "width": dims.value[0], "height": dims.value[1], "dimension_unit": unit,
        "length": length_value, "length_unit": "m" if unit.lower() == "mm" else "ft",
        "metal_material_code": metal_material_code, "type": typ.value,
        "installation": install.value, "system": system.strip(),
        "accessory_rows": system_result.value["accessory_rows"],
    })


def special_item(item: str) -> RuleResult:
    name = item.strip().lower()
    blocked = {"reducer", "tap", "cap", "union", "demolition", "plenum", "kitchen hood", "equipment flex"}
    if name in blocked:
        return RuleResult(UNSUPPORTED, f"{item} is outside this rule engine's automatic scope.", "PPT slides 23, 34, 38, 40; implementation request special-item rule")
    return RuleResult(UNSUPPORTED, "No explicit mapping is available for this item.", "PPT must provide a sufficient explicit mapping")

