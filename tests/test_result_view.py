"""结果摘要使用真实附件；检查平移口径与未完成检查时的显示。"""
from copy import deepcopy
from pathlib import Path
import sys
import unittest

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
from engine import validate_submission  # noqa: E402
from result_view import build_result_view  # noqa: E402


class ResultViewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.results = {
            q: validate_submission(q, APP / f"examples/result{q}_可行示例.xlsx",
                                   base_q2=APP / "examples/result2_可行示例.xlsx" if q == 3 else None)
            for q in (2, 3, 4)
        }

    def test_question_two_metrics(self):
        view = build_result_view(self.results[2])
        self.assertEqual(view["status"], "feasible")
        self.assertEqual([m["value"] for m in view["metrics"]], ["全部", "6", "126", "678"])
        self.assertEqual((view["frequency_distance"], view["time_distance"]), (581, 97))
        self.assertEqual(view["total_amplitude"], 678)

    def test_question_three_metrics_belong_to_base(self):
        view = build_result_view(self.results[3])
        self.assertEqual((view["cancelled"], view["adjusted"], view["translation_distance"]), (6, 126, 678))
        self.assertEqual(view["added_count"], 136)
        self.assertIn("第二问基础方案", view["metrics"][1]["hint"])
        self.assertTrue(any("新增 136 台" in note for note in view["notes"]))

    def test_question_four_separates_gap_from_translation(self):
        view = build_result_view(self.results[4])
        self.assertEqual([m["value"] for m in view["metrics"]], ["全部", "4", "140", "537"])
        self.assertEqual((view["frequency_distance"], view["time_distance"], view["gap_distance"]), (438, 99, 238))
        self.assertEqual(view["translation_distance"] + view["gap_distance"], view["total_amplitude"])
        self.assertEqual(view["total_amplitude"], self.results[4]["summary"]["objective_tuple"][6])
        self.assertTrue(any("41 台间隔调整" in note for note in view["notes"]))

    def test_parse_failure_does_not_display_partial_zero(self):
        result = validate_submission(2, APP / "examples/result2_重复编号示例.xlsx")
        self.assertEqual(result["status"], "invalid")
        self.assertTrue(result["summary"]["analysis_complete"])
        view = build_result_view(result)
        self.assertEqual([m["value"] for m in view["metrics"]], ["未完成", "—", "—", "—"])
        for key in ("cancelled", "adjusted", "translation_distance", "total_amplitude", "added_count"):
            self.assertIsNone(view[key])

    def test_incomplete_analysis_or_missing_plan_hides_metrics(self):
        for mode in ("analysis", "plans"):
            with self.subTest(mode=mode):
                result = deepcopy(self.results[2])
                if mode == "analysis":
                    result["summary"]["analysis_complete"] = False
                else:
                    result["plans"].pop()
                view = build_result_view(result)
                self.assertEqual(view["status"], "invalid")
                self.assertIsNone(view["translation_distance"])

    def test_infeasible_complete_submission_shows_submission_statistics(self):
        result = validate_submission(3, APP / "examples/result3_新增重叠示例.xlsx",
                                     base_q2=APP / "examples/result2_可行示例.xlsx")
        self.assertEqual(result["status"], "infeasible")
        view = build_result_view(result)
        self.assertEqual(view["title"], "验证失败")
        self.assertEqual(view["metrics"][0]["value"], "未通过")
        self.assertEqual(view["translation_distance"], 678)
        self.assertIn("当前提交统计", view["notes"][0])
        self.assertIn(str(len(result["conflicts"])), view["message"])

    def test_canceled_and_new_devices_do_not_contribute_distance(self):
        result = deepcopy(self.results[3])
        for plan in result["plans"]:
            if plan["canceled"] or plan["is_new"]:
                plan.update(df=999, dt=-999, dg=999)
            else:
                plan["count"] *= 2
        view = build_result_view(result)
        self.assertEqual(view["translation_distance"], 678)
        self.assertEqual(view["gap_distance"], 0)
        self.assertEqual(view["adjusted"], 126)

    def test_all_canceled_is_a_known_zero_not_missing(self):
        result = deepcopy(self.results[2])
        for plan in result["plans"]:
            plan["canceled"] = True
        view = build_result_view(result)
        self.assertEqual([m["value"] for m in view["metrics"]], ["全部", "150", "0", "0"])


if __name__ == "__main__":
    unittest.main()
