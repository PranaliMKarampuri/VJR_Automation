import unittest
from decimal import Decimal

from rectangular_duct_takeoff import map_accessory, transform_records


def record(width, height, length=5, installation="HA", units="M", accessories=None):
    return {
        "Width": width,
        "Height": height,
        "Length": length,
        "MetalCode": 6,
        "Type": "BP",
        "Installation": installation,
        "Units": units,
        "Accessories": accessories or [],
    }


class RectangularDuctRuleTests(unittest.TestCase):
    def test_larger_dimension_is_written_first(self):
        rows, errors, units = transform_records([record(250, 400)])
        self.assertEqual(errors, [])
        self.assertEqual(units, "M")
        self.assertEqual((rows[0]["Width"], rows[0]["Height"]), (Decimal("400"), Decimal("250")))

    def test_installation_ha_and_re_from_documented_36_inch_boundary(self):
        rows, errors, _ = transform_records([record(900, 500), record(1000, 500)])
        self.assertEqual(errors, [])
        self.assertEqual([row["Installation"] for row in rows], ["HA", "RE"])

    def test_riser_installation_is_ri_even_above_36_inches(self):
        rows, errors, _ = transform_records([record(1200, 500, installation="RI")])
        self.assertEqual(errors, [])
        self.assertEqual(rows[0]["Installation"], "RI")

    def test_exactly_36_inches_is_flagged_for_confirmation(self):
        rows, errors, _ = transform_records([record(36, 20, units="A")])
        self.assertEqual(rows, [])
        self.assertTrue(any("exactly 36 inches" in error for error in errors))

    def test_documented_accessory_mapping_and_four_slot_row_split(self):
        accessories = [
            "T-VANES",
            "Canevas",
            "Fire damper",
            "Fire smoke damper",
            "Gravity damper",
            "Mesh",
        ]
        rows, errors, _ = transform_records([record(400, 250, accessories=accessories)])
        self.assertEqual(errors, [])
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Accessories"], [7, 8, 10, 100])
        self.assertEqual(rows[0]["Length"], Decimal("5"))
        self.assertEqual(rows[1]["Accessories"], [12, 14])
        self.assertEqual(rows[1]["Length"], Decimal("0"))
        self.assertEqual((rows[0]["Width"], rows[0]["Height"]), (rows[1]["Width"], rows[1]["Height"]))

    def test_size_dependent_accessory_codes_above_30_inches(self):
        self.assertEqual(map_accessory("T-VANES", Decimal("800"), Decimal("400"), "M"), 71)
        self.assertEqual(map_accessory("MBD", Decimal("800"), Decimal("400"), "M"), 131)
        self.assertEqual(map_accessory("T-VANES", Decimal("762"), Decimal("400"), "M"), 7)
        self.assertEqual(map_accessory("MBD", Decimal("762"), Decimal("400"), "M"), 13)

    def test_motorized_name_uses_size_rule_and_explicit_codes_are_accepted(self):
        self.assertEqual(map_accessory(11, Decimal("400"), Decimal("250"), "M"), 11)
        self.assertEqual(map_accessory(111, Decimal("400"), Decimal("250"), "M"), 111)
        rows, errors, _ = transform_records([record(400, 250, accessories=["Motorized damper"])])
        self.assertEqual(errors, [])
        self.assertEqual(rows[0]["Accessories"], [11])

    def test_project_unit_must_be_consistent(self):
        rows, errors, _ = transform_records([record(400, 250, units="M"), record(20, 10, units="A")])
        self.assertEqual(rows, [])
        self.assertTrue(any("one project-wide unit selector" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
