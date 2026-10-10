import csv
import hashlib
import tempfile
import unittest
from pathlib import Path

import openpyxl

import manual_to_vjr_dry_run as dryrun
import vjr_ppt_rules as rules


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


class ManualToVjrDryRunTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.manual = self.root / "manual.xlsx"
        self.vjr = self.root / "vjr.xlsx"
        self.mapping = self.root / "mapping.csv"
        self.crosswalk = self.root / "crosswalk.csv"
        self.decisions = self.root / "decisions.csv"
        self.out_csv = self.root / "out.csv"
        self.out_md = self.root / "out.md"
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "DATA"
        sheet.append(["Subject", "Page Index", "Length", "Length Unit", "Count", "Label", "BRAND", "MODEL", "DUCT SIZE", "TYPE", "ACCESSORIES1", "ACCESSORIES2", "REMARK", "UNITS", "DAMPER TYPE"])
        sheet.append(["RECT. DUCT M", 1, 2.0, "m", 0, "RECT. DUCT", ".", ".", "200X300", ".", ".", ".", ".", "UTA-1", "NO"])
        sheet.append(["RECT. DUCT M", 1, 3.0, "m", 0, "RECT. DUCT", ".", ".", "200X300", ".", ".", ".", ".", ".", "NO"])
        sheet.append(["VENT CAP", 2, 0, ".", 1, "VENT CAP", ".", ".", 150, ".", ".", ".", ".", ".", "NO"])
        book.save(self.manual)
        vjr_book = openpyxl.Workbook()
        vjr_sheet = vjr_book.active
        vjr_sheet.title = "TAKE OFF"
        vjr_sheet["A1"] = "Header"
        vjr_sheet["A6"] = 200
        vjr_sheet["B6"] = 300
        vjr_sheet["C6"] = 8
        vjr_book.save(self.vjr)
        self._write_csv(self.mapping, [
            {"Source Row": "2", "Record Type": "rectangular duct", "Status": "NEEDS_CONFIRMATION", "Match Status": "UNMATCHED", "Matching VJR Record/Cell if found": "TAKE OFF!A6:C6", "VJR Sheet": "TAKE OFF", "VJR Cell/Column": "TAKE OFF!A6:C6", "VJR Input Field": "width/height/length", "Rule/Logic": "explicit candidate only", "Confidence": "Low", "Notes": "Await estimator."},
            {"Source Row": "3", "Record Type": "rectangular duct", "Status": "NEEDS_CONFIRMATION", "Match Status": "UNMATCHED", "Matching VJR Record/Cell if found": "", "VJR Sheet": "UNRESOLVED", "VJR Cell/Column": "", "VJR Input Field": "", "Rule/Logic": "", "Confidence": "Low", "Notes": "Await estimator."},
            {"Source Row": "4", "Record Type": "other", "Status": "DO_NOT_AUTOMATE_YET", "Match Status": "UNMATCHED", "Matching VJR Record/Cell if found": "", "VJR Sheet": "UNRESOLVED", "VJR Cell/Column": "", "VJR Input Field": "", "Rule/Logic": "unsupported", "Confidence": "Low", "Notes": "Unsupported item."},
        ])
        self._write_csv(self.crosswalk, [{"manual_source_row": "2", "manual_source_sheet": "DATA", "mapping_basis": "GROUP_ROLLUP", "vjr_candidate_sheet": "TAKE OFF", "vjr_candidate_row": "6", "vjr_width": "", "vjr_height": "", "vjr_length": "", "vjr_metal": "", "vjr_type": "", "vjr_installation": "", "vjr_accessories": "", "confidence": "Medium", "estimator_confirmation_needed": "Confirm group membership."}])
        self._write_csv(self.decisions, [{"decision_id": "D-TEST", "manual_source_rows": "2; 3", "status": "AWAITING_ESTIMATOR"}])

    def tearDown(self):
        self.temp.cleanup()

    @staticmethod
    def _write_csv(path, rows):
        with path.open("w", newline="", encoding="utf-8-sig") as stream:
            writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)

    def run_pipeline(self):
        return dryrun.run_dry_run(self.manual, self.vjr, self.mapping,
            self.crosswalk, self.decisions, self.out_csv, self.out_md)

    def test_one_rectangular_record_is_never_safe_without_approval(self):
        result = self.run_pipeline()
        record = result["records"][0]
        self.assertEqual(record["manual_duct_size"], "200X300")
        self.assertEqual(record["mapping_status"], rules.NEEDS_CONFIRMATION)
        self.assertEqual(record["rule_status"], rules.NEEDS_CONFIRMATION)
        self.assertEqual(result["summary"]["safe_to_automate_count"], 0)

    def test_dispositions_use_required_names(self):
        result = self.run_pipeline()
        self.assertEqual(
            {record["mapping_status"] for record in result["records"]},
            {"NEEDS_CONFIRMATION", "UNSUPPORTED"},
        )
        self.assertIn("SAFE_TO_AUTOMATE", self.out_md.read_text(encoding="utf-8"))

    def test_grouped_rectangular_record_retains_group_rollup(self):
        result = self.run_pipeline()
        self.assertEqual(result["summary"]["grouped_rollup_count"], 1)
        self.assertEqual(result["records"][0]["mapping_basis"], "GROUP_ROLLUP")
        self.assertIn("not an exact row mapping", result["records"][0]["reason"])

    def test_missing_system_and_accessory_association_are_reported(self):
        result = self.run_pipeline()
        record = result["records"][1]
        self.assertEqual(record["mapping_status"], rules.NEEDS_CONFIRMATION)
        self.assertIn("system/unit is blank", record["reason"])
        self.assertIn("No accessory association", record["reason"])

    def test_missing_installation_remains_unfilled(self):
        result = self.run_pipeline()
        self.assertEqual(result["records"][0]["proposed_installation"], "")

    def test_exactly_36_inch_boundary_needs_confirmation(self):
        result = rules.installation(36, 20, "inch")
        self.assertEqual(result.status, rules.NEEDS_CONFIRMATION)

    def test_unsupported_special_case(self):
        result = self.run_pipeline()
        self.assertEqual(result["records"][2]["rule_category"], "other/unknown")
        self.assertEqual(result["records"][2]["mapping_status"], rules.UNSUPPORTED)

    def test_existing_vjr_candidate_row_is_protected(self):
        result = self.run_pipeline()
        self.assertIn("row 6: existing data/formula present", result["records"][0]["candidate_row_protection"])
        self.assertIn("never be overwritten", result["records"][0]["candidate_row_protection"])

    def test_vjr_workbook_bytes_are_unchanged(self):
        before = sha256(self.vjr)
        self.run_pipeline()
        self.assertEqual(sha256(self.vjr), before)

    def test_protects_against_using_source_workbook_as_output(self):
        with self.assertRaises(ValueError):
            dryrun.run_dry_run(self.manual, self.vjr, self.mapping,
                self.crosswalk, self.decisions, self.vjr, self.out_md)


if __name__ == "__main__":
    unittest.main()
