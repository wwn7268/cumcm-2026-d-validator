"""问题四附件：间隔调整、完整周期与指标回归。"""
from pathlib import Path
import json
import sys
import unittest

from openpyxl import Workbook

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
from engine import validate_submission  # noqa: E402

Q4_HEADER = ["用频装备编号", "调整后频段范围", "调整后时间区间", "调整后间隔时长", "是否撤销用频计划"]


class QuestionFourTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = APP / "tests/_generated"
        cls.directory.mkdir(parents=True, exist_ok=True)
        cls.original = json.loads((APP / "data/original_plans.json").read_text(encoding="utf-8"))
        cls.all_canceled = [[p["id"], None, None, None, "是"] for p in cls.original]
        cls.q4 = APP / "examples/result4_可行示例.xlsx"

    def table(self, rows, header=Q4_HEADER):
        path = self.directory / (self._testMethodName + ".xlsx")
        book = Workbook()
        sheet = book.active
        sheet.title = "问题四"
        sheet.append(header)
        for row in rows:
            sheet.append(row)
        book.save(path)
        return path

    def isolated(self, replacement, *, header=Q4_HEADER, **kwargs):
        rows = [row for row in self.all_canceled if row[0] != replacement[0]]
        return validate_submission(4, self.table(rows + [replacement], header), **kwargs)

    def assert_rejected(self, result, code):
        self.assertFalse(result["passed"], result)
        self.assertIn(code, {error["code"] for error in result["errors"]}, result["errors"])

    def test_known_q4_feasible_and_metrics(self):
        result = validate_submission(4, self.q4)
        self.assertTrue(result["passed"], result["errors"])
        summary = result["summary"]
        self.assertEqual(summary["objective_six"], [4, 0, 2, 140, 18, 34])
        self.assertEqual(summary["frequency_adjusted"], 66)
        self.assertEqual(summary["time_adjusted"], 33)
        self.assertEqual(summary["gap_adjusted"], 41)
        self.assertEqual(summary["occupied_area"], 16176)
        self.assertEqual(summary["demand_area"], 16176)
        self.assertEqual(summary["conflict_events"], 0)
        self.assertEqual(result["conflicts"], [])
        self.assertAlmostEqual(summary["utilization"], 16176 / 64300)

    def test_q4_independent_of_q2_base(self):
        result = validate_submission(4, self.q4, base_q2=APP / "does_not_exist.xlsx")
        self.assertTrue(result["passed"], result["errors"])
        self.assertNotIn("base_validation", result)

    def test_zero_gap_is_valid_and_counts_amplitude(self):
        result = self.isolated(["C001", None, None, 0, None])
        self.assertTrue(result["passed"], result["errors"])
        self.assertEqual(result["summary"]["occupied_area"], 72)
        self.assertEqual(result["summary"]["max_end"], 64)
        self.assertEqual(result["summary"]["objective_tuple"][3], 1)
        self.assertEqual(result["summary"]["objective_tuple"][6], 8)
        self.assertEqual(result["summary"]["by_category"]["C"]["gap_adjusted"], 1)

    def test_gap_eighteen_and_cumulative_shift_over_five_valid(self):
        result = self.isolated(["C001", None, None, 18, "否"])
        self.assertTrue(result["passed"], result["errors"])
        plan = next(p for p in result["plans"] if p["id"] == "C001")
        self.assertEqual(plan["t0"], 40)
        self.assertEqual(plan["dt"], 0)
        self.assertEqual(plan["dg"], 10)
        self.assertEqual(result["summary"]["max_end"], 262)
        self.assertEqual(result["summary"]["objective_tuple"][6], 10)

    def test_gap_nineteen_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, None, 19, None]), "gap_shift_limit")

    def test_negative_gap_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, None, -1, None]), "invalid_repetition")

    def test_fractional_gap_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, None, 8.5, None]), "invalid_integer")

    def test_a_gap_change_rejected(self):
        self.assert_rejected(self.isolated(["A001", None, None, 61, None]), "gap_class_limit")

    def test_b_gap_change_rejected(self):
        self.assert_rejected(self.isolated(["B001", None, None, 39, None]), "gap_class_limit")

    def test_gap_and_time_change_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, "[41,43)", 9, None]), "multiple_actions")

    def test_gap_and_frequency_change_rejected(self):
        self.assert_rejected(self.isolated(["C001", "[91,94)", None, 9, None]), "multiple_actions")

    def test_canceled_and_gap_change_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, None, 9, "是"]), "canceled_and_adjusted")

    def test_repeat_count_change_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, None, None, None, 11],
                                          header=Q4_HEADER + ["使用次数"]), "changed_count")

    def test_gap_change_last_use_must_fit_horizon(self):
        self.assert_rejected(self.isolated(["C083", None, None, 9, None]), "time_boundary")

    def test_unchanged_intervals_with_gap_change_are_one_action(self):
        result = self.isolated(["C001", "[90,93)", "[40,42)", 9, None])
        self.assertTrue(result["passed"], result["errors"])
        self.assertEqual(result["summary"]["gap_adjusted"], 1)
        self.assertEqual(result["summary"]["objective_tuple"][3], 1)
        self.assertEqual(result["summary"]["objective_tuple"][6], 1)

    def test_blank_values_preserve_original(self):
        result = self.isolated(["C001", None, None, None, None])
        self.assertTrue(result["passed"], result["errors"])
        plan = next(p for p in result["plans"] if p["id"] == "C001")
        self.assertEqual((plan["f0"], plan["t0"], plan["gap"], plan["count"]), (90, 40, 8, 12))
        self.assertEqual(result["summary"]["objective_tuple"][3], 0)

    def test_omitted_device_preserves_original(self):
        rows = [row for row in self.all_canceled if row[0] != "C001"]
        result = validate_submission(4, self.table(rows))
        self.assertTrue(result["passed"], result["errors"])
        self.assertEqual(result["summary"]["occupied_area"], 72)
        self.assertEqual(result["summary"]["objective_tuple"][3], 0)

    def test_empty_q4_keeps_original_conflicts(self):
        result = validate_submission(4, self.table([]))
        self.assert_rejected(result, "resource_conflict")
        self.assertEqual(result["summary"]["occupied_area"], 14698)
        self.assertEqual(result["summary"]["demand_area"], 16680)

    def test_duplicate_device_rejected(self):
        result = validate_submission(4, self.table(self.all_canceled + [["C001", None, None, None, "是"]]))
        self.assert_rejected(result, "duplicate_id")

    def test_formula_gap_rejected(self):
        self.assert_rejected(self.isolated(["C001", None, None, "=8+1", None]), "formula_cell")

    def test_nonconforming_example_rejected(self):
        result = validate_submission(4, APP / "examples/result4_间隔超限示例.xlsx")
        self.assert_rejected(result, "gap_shift_limit")


if __name__ == "__main__":
    unittest.main()
