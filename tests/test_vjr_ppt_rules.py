import unittest
from decimal import Decimal

import vjr_ppt_rules as rules


class RectangularRulesTests(unittest.TestCase):
    def test_dimension_ordering(self):
        result = rules.rectangular_dimensions(250, 400, "mm")
        self.assertEqual(result.status, rules.SAFE)
        self.assertEqual(result.value, (Decimal(400), Decimal(250), "mm"))

    def test_bp_mp_hp(self):
        for pressure, label in (("BP", "low pressure"), ("MP", "medium pressure"), ("HP", "high pressure")):
            with self.subTest(pressure=pressure):
                self.assertEqual(rules.pressure_type(pressure).value, label)

    def test_installation_ha_re_and_boundary(self):
        self.assertEqual(rules.installation(24, 30).value, "HA")
        self.assertEqual(rules.installation(37, 20).value, "RE")
        self.assertEqual(rules.installation(914.4, 500, "mm").status, rules.NEEDS_CONFIRMATION)
        self.assertEqual(rules.installation(100, 200, "mm", is_riser=True).value, "RI")

    def test_all_documented_accessory_codes(self):
        expected = {1, 21, 7, 71, 8, 10, 100, 11, 12, 13, 131, 14}
        for code in expected:
            with self.subTest(code=code):
                self.assertEqual(rules.accessory(code).status, rules.SAFE)
        self.assertEqual(rules.accessory(999).status, rules.UNSUPPORTED)

    def test_accessory_continuation_rows(self):
        result = rules.split_accessories([1, 7, 8, 10, 11, 12], riser=True)
        self.assertEqual(result.status, rules.SAFE)
        self.assertEqual(result.value[0]["accessories"], [1, 7, 8, 10])
        self.assertEqual(result.value[1], {"accessories": [11, 12], "length": 0, "repeat_size": True, "installation": "RI"})

    def test_rectangular_riser(self):
        result = rules.rectangular_riser(250, 400, 4.5, [1, 8], "mm")
        self.assertEqual(result.status, rules.SAFE)
        self.assertEqual(result.value[0]["installation"], "RI")
        self.assertEqual(result.value[0]["length_m"], Decimal("4.5"))

    def test_complete_rectangular_record_preserves_explicit_material_and_length(self):
        result = rules.rectangular_record(width=500, height=300, length=12.5,
            unit="mm", metal_material_code="22", pressure="MP",
            system="AHU-1", accessories=[1, 7])
        self.assertEqual(result.status, rules.SAFE)
        self.assertEqual(result.value["metal_material_code"], "22")
        self.assertEqual(result.value["length_unit"], "m")
        self.assertEqual(result.value["length"], Decimal("12.5"))
        self.assertEqual(result.value["system"], "AHU-1")
        self.assertEqual(rules.rectangular_record(width=500, height=300,
            length=12.5, unit="mm", metal_material_code="22",
            pressure="MP", system=None).status, rules.NEEDS_CONFIRMATION)


class SpiralRulesTests(unittest.TestCase):
    def test_spiral_riser_drop(self):
        result = rules.spiral_riser_drop(10, 15)
        self.assertEqual(result.value["add_length_ft"], Decimal(15))
        self.assertEqual(rules.spiral_riser_drop(10, 15, already_included=True).status, rules.NEEDS_CONFIRMATION)

    def test_spiral_length_aggregation_and_metric_conversion(self):
        result = rules.spiral_length_by_diameter([(8, 20), (8, 5), (10, 7)])
        self.assertEqual(result.value[Decimal(8)], Decimal(25))
        metric = rules.spiral_length_by_diameter([(203.2, 3)], "mm")
        self.assertEqual(metric.status, rules.SAFE)
        self.assertAlmostEqual(float(metric.value[Decimal(8)]), 9.842519685, places=6)

    def test_45_and_90_elbows(self):
        for angle in (45, 90):
            with self.subTest(angle=angle):
                result = rules.spiral_elbows(angle, 3, 250)
                self.assertEqual(result.status, rules.SAFE)
                self.assertEqual(result.value["degrees"], angle)

    def test_y_type_validation(self):
        self.assertEqual(rules.y_type([8, 10], 2).status, rules.SAFE)
        self.assertEqual(rules.y_type([8], 2).status, rules.NEEDS_CONFIRMATION)
        self.assertEqual(rules.y_type([8, 10], 1).status, rules.NEEDS_CONFIRMATION)


class FlexAndAccessoryAssociationTests(unittest.TestCase):
    def test_ordinary_flex_five_feet_each(self):
        self.assertEqual(rules.flex_length(4).value, Decimal(20))

    def test_erv_flex_requires_explicit_confirmation(self):
        self.assertEqual(rules.flex_length(2, kind="erv").status, rules.NEEDS_CONFIRMATION)
        self.assertEqual(rules.flex_length(2, kind="erv", client_confirmed=True).value, Decimal(20))

    def test_equipment_flex_confirmation(self):
        self.assertEqual(rules.flex_length(1, kind="equipment").status, rules.NEEDS_CONFIRMATION)

    def test_canvas_requires_both_explicit_facts(self):
        self.assertEqual(rules.canvas_eligibility(direct_to_fan_ahu_rtu=None, kitchen_hood_or_dryer=False).status, rules.NEEDS_CONFIRMATION)
        self.assertEqual(rules.canvas_eligibility(direct_to_fan_ahu_rtu=True, kitchen_hood_or_dryer=False).value, 8)
        self.assertEqual(rules.canvas_eligibility(direct_to_fan_ahu_rtu=True, kitchen_hood_or_dryer=True).status, rules.UNSUPPORTED)

    def test_explicit_mbd_destinations(self):
        self.assertEqual(rules.mbd(explicit=True, duct_kind="rectangular", size_in=32).value, 131)
        self.assertEqual(rules.mbd(explicit=True, duct_kind="rectangular", size_in=30).value, 13)
        self.assertEqual(rules.mbd(explicit=True, duct_kind="spiral").value, "SPIRALE.damper")
        self.assertEqual(rules.mbd(explicit=False, duct_kind="spiral").status, rules.NEEDS_CONFIRMATION)

    def test_rigid_round_mbd_code_9_is_blocked_without_size(self):
        self.assertEqual(rules.mbd(explicit=True, duct_kind="rigid_round", size_available=False).status, rules.NEEDS_CONFIRMATION)
        self.assertEqual(rules.accessory(9).status, rules.UNSUPPORTED)

    def test_spiral_nonbalancing_damper_needs_rectangular_size(self):
        result = rules.spiral_nonbalancing_damper(damper_code=10)
        self.assertEqual(result.status, rules.NEEDS_CONFIRMATION)
        explicit = rules.spiral_nonbalancing_damper(width=12, height=8, damper_code=10)
        self.assertEqual(explicit.value, {"width": Decimal(12), "height": Decimal(8), "length": 0, "accessory": 10})

    def test_system_unit_required_and_separate(self):
        self.assertEqual(rules.system_assignment(None).status, rules.NEEDS_CONFIRMATION)
        assigned = rules.system_assignment("EF-1", [1, 8])
        self.assertEqual(assigned.status, rules.SAFE)
        self.assertEqual(assigned.value["system"], "EF-1")

    def test_special_items_stay_unsupported(self):
        for item in ("reducer", "tap", "cap", "union", "demolition", "plenum", "kitchen hood", "equipment flex"):
            with self.subTest(item=item):
                self.assertEqual(rules.special_item(item).status, rules.UNSUPPORTED)


if __name__ == "__main__":
    unittest.main()
