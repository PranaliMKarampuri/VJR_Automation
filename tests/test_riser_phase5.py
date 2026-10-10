import unittest
from decimal import Decimal
from unittest.mock import patch

from riser_takeoff import build_values, deduplicate_floor_records


def rectangular_riser(riser_id="R-1", floor="Level 1", accessories=None, length="4.5"):
    return {
        "RiserID": riser_id,
        "Floor": floor,
        "System": "RECTANGULAR",
        "Width": 300,
        "Height": 200,
        "Length": length,
        "LengthUnit": "M",
        "MetalCode": 6,
        "Type": "BP",
        "Units": "M",
        "Diameter": "",
        "DiameterUnit": "",
        "Installation": "RI",
        "Accessories": accessories or [],
    }


def spiral_riser(riser_id="S-1", floor="Level 1", diameter="8", length="15"):
    return {
        "RiserID": riser_id,
        "Floor": floor,
        "System": "SPIRAL",
        "Width": "",
        "Height": "",
        "Length": length,
        "LengthUnit": "FT",
        "MetalCode": "",
        "Type": "",
        "Units": "",
        "Diameter": diameter,
        "DiameterUnit": "IN",
        "Accessories": [],
    }


class RiserPhase5Tests(unittest.TestCase):
    def test_duplicate_floor_listing_counts_once_by_explicit_id(self):
        records, suppressed, errors = deduplicate_floor_records([
            rectangular_riser(floor="Level 1"),
            rectangular_riser(floor="Level 2"),
        ])
        self.assertEqual(errors, [])
        self.assertEqual(suppressed, 1)
        self.assertEqual(len(records), 1)

    def test_conflicting_floor_records_are_rejected(self):
        records, suppressed, errors = deduplicate_floor_records([
            rectangular_riser(floor="Level 1"),
            rectangular_riser(floor="Level 2", length="7"),
        ])
        self.assertEqual(records, [])
        self.assertEqual(suppressed, 0)
        self.assertTrue(any("conflicting riser data" in error for error in errors))

    def test_rectangular_riser_accessories_continue_with_zero_length_and_ri(self):
        record = rectangular_riser(accessories=[7, 8, 10, 100, 12])
        with patch("riser_takeoff.diameter_rows", return_value={}):
            values, _, errors = build_values([record], "unused-template.xlsx")
        self.assertEqual(errors, [])
        self.assertEqual(values["TAKE OFF"]["C20"], Decimal("4.5"))
        self.assertEqual([values["TAKE OFF"][f"{col}20"] for col in "GHIJ"], [7, 8, 10, 100])
        self.assertEqual(values["TAKE OFF"]["C21"], Decimal("0"))
        self.assertEqual(values["TAKE OFF"]["F20"], "RI")
        self.assertEqual(values["TAKE OFF"]["F21"], "RI")
        self.assertEqual(values["TAKE OFF"]["A20"], values["TAKE OFF"]["A21"])

    def test_spiral_riser_feet_aggregate_to_column_b_only(self):
        records = [spiral_riser("S-1"), spiral_riser("S-2", length="12")]
        with patch("riser_takeoff.diameter_rows", return_value={Decimal("8"): 16}):
            values, _, errors = build_values(records, "unused-template.xlsx")
        self.assertEqual(errors, [])
        self.assertEqual(values["SPIRALE"], {"B16": Decimal("27")})

    def test_spiral_nonmatching_diameter_is_rejected_without_rounding(self):
        with patch("riser_takeoff.diameter_rows", return_value={Decimal("8"): 16}):
            values, _, errors = build_values([spiral_riser(diameter="8.01")], "unused-template.xlsx")
        self.assertEqual(values, {})
        self.assertTrue(any("does not exactly match" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
