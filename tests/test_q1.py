"""Regression checks for complete, unordered original conflict-pair lists."""
from pathlib import Path
from uuid import uuid4
import csv
import shutil
import sys
import unittest

from openpyxl import Workbook, load_workbook

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
from engine import validate_submission
from q1 import HEADERS, conflict_truth


class QuestionOneTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.correct = APP / 'examples/result1_正确示例.xlsx'
        workbook = load_workbook(cls.correct, read_only=True)
        try:
            cls.rows = [list(row) for row in list(workbook.active.values)[1:]]
        finally:
            workbook.close()

    def setUp(self):
        generated = (APP / 'tests/_generated/q1').resolve()
        self.folder = generated / uuid4().hex
        self.folder.mkdir(parents=True)

        def cleanup():
            target = self.folder.resolve()
            if target.parent != generated or not target.is_relative_to(APP.resolve()):
                raise AssertionError('Refusing to remove data outside the test folder.')
            shutil.rmtree(target)

        self.addCleanup(cleanup)

    def table(self, rows, *, header=HEADERS, filename='result1.xlsx'):
        target = self.folder / filename
        workbook = Workbook()
        page = workbook.active
        page.title = '冲突清单'
        page.append(header)
        for row in rows:
            page.append(row)
        workbook.save(target)
        return target

    def test_correct_pair_list_passes_despite_original_conflicts(self):
        result = validate_submission(1, self.correct)
        self.assertTrue(result['passed'], result['errors'])
        self.assertEqual(result['status'], 'feasible')
        self.assertEqual(result['summary']['expected_pair_count'], 297)
        self.assertEqual(result['summary']['submitted_pair_count'], 297)
        self.assertEqual(result['summary']['correct_pair_count'], 297)
        self.assertEqual(result['summary']['expected_event_count'], 431)
        self.assertEqual(len(result['conflicts']), 431)
        self.assertEqual(result['summary']['missing_pair_count'], 0)
        self.assertEqual(result['summary']['extra_pair_count'], 0)
        self.assertEqual(result['summary']['duplicate_pair_count'], 0)
        self.assertTrue(result['summary']['analysis_complete'])

    def test_reversed_pairs_and_reordered_rows_are_accepted(self):
        rows = [[serial, right, left] for serial, left, right in reversed(self.rows)]
        result = validate_submission(1, self.table(rows))
        self.assertTrue(result['passed'], result['errors'])

    def test_missing_pair_is_explicit(self):
        result = validate_submission(1, self.table(self.rows[1:]))
        self.assertEqual(result['status'], 'infeasible')
        self.assertEqual(result['summary']['correct_pair_count'], 296)
        self.assertEqual(result['summary']['missing_pair_count'], 1)
        _, left, right = self.rows[0]
        self.assertEqual(result['missing_pairs'], [{'id1': left, 'id2': right}])
        self.assertIn('missing_pairs', {error['code'] for error in result['errors']})

    def test_extra_nonconflicting_pair_is_explicit(self):
        result = validate_submission(1, self.table(self.rows + [[298, 'A001', 'A002']]))
        self.assertEqual(result['status'], 'infeasible')
        self.assertEqual(result['summary']['submitted_pair_count'], 298)
        self.assertEqual(result['summary']['extra_pair_count'], 1)
        self.assertEqual(result['extra_pairs'], [{'id1': 'A001', 'id2': 'A002'}])
        self.assertIn('extra_pair', {error['code'] for error in result['errors']})

    def test_duplicates_and_reverse_duplicates_count_extra_rows(self):
        _, left, right = self.rows[0]
        result = validate_submission(1, self.table(self.rows + [[298, left, right], [299, right, left]]))
        self.assertEqual(result['status'], 'infeasible')
        self.assertEqual(result['summary']['submitted_row_count'], 299)
        self.assertEqual(result['summary']['submitted_pair_count'], 297)
        self.assertEqual(result['summary']['duplicate_pair_count'], 2)
        self.assertTrue(all(item['first_row'] == 2 for item in result['duplicate_pairs']))

    def test_self_conflict_and_unknown_equipment_are_rejected(self):
        for right, code in [('A001', 'self_conflict'), ('A999', 'unknown_id')]:
            with self.subTest(code=code):
                result = validate_submission(1, self.table(self.rows + [[298, 'A001', right]]))
                self.assertEqual(result['status'], 'invalid')
                self.assertFalse(result['passed'])
                self.assertIn(code, {error['code'] for error in result['errors']})

    def test_formula_error_value_wrong_header_and_corrupt_workbook(self):
        cases = [('formula', '=1+1', 'formula_cell'), ('excel_error', '#DIV/0!', 'excel_error_cell')]
        for name, value, code in cases:
            with self.subTest(code=code):
                path = self.table([[value, 'A001', 'B001']], filename=name + '.xlsx')
                result = validate_submission(1, path)
                self.assertEqual(result['status'], 'invalid')
                self.assertIn(code, {error['code'] for error in result['errors']})
        wrong = self.table([], header=['用频装备编号', '调整后频段区间', '调整后时间区间', '是否撤销用频计划'])
        result = validate_submission(1, wrong)
        self.assertIn('missing_header', {error['code'] for error in result['errors']})
        broken = self.folder / 'broken.xlsx'
        broken.write_bytes(b'not an Excel workbook')
        result = validate_submission(1, broken)
        self.assertIn('unreadable_workbook', {error['code'] for error in result['errors']})

    def test_long_chinese_headers_and_csv(self):
        header = ['序号', '发生冲突的用频装备编号1', '发生冲突的用频装备编号2']
        path = self.folder / '第一问.csv'
        with path.open('w', encoding='utf-8-sig', newline='') as handle:
            writer = csv.writer(handle)
            writer.writerow(header)
            writer.writerows(self.rows)
        result = validate_submission(1, path)
        self.assertTrue(result['passed'], result['errors'])

    def test_short_requested_horizon_does_not_clip_truth(self):
        result = validate_submission(1, self.correct, horizon=10)
        self.assertTrue(result['passed'], result['errors'])
        self.assertEqual(result['settings']['requested_horizon'], 10)
        self.assertEqual(result['settings']['horizon'], 643)
        self.assertEqual(result['summary']['expected_pair_count'], 297)
        self.assertEqual(len(result['conflicts']), 431)

    def test_truth_checks_later_repetitions_and_half_open_boundaries(self):
        a = dict(id='A001', f0=0, f1=2, t0=0, t1=2, gap=8, count=2)
        b = dict(id='B001', f0=0, f1=2, t0=10, t1=12, gap=0, count=1)
        pairs, events = conflict_truth([a, b])
        self.assertEqual(pairs, {('A001', 'B001')})
        self.assertEqual((events[0]['use1'], events[0]['t0'], events[0]['t1']), (2, 10, 12))
        self.assertEqual(conflict_truth([a, dict(b, t0=2, t1=4)])[0], set())
        self.assertEqual(conflict_truth([a, dict(b, f0=2, f1=4)])[0], set())

    def test_questions_two_to_four_dispatch_unchanged(self):
        base = APP / 'examples/result2_可行示例.xlsx'
        for question in (2, 3, 4):
            with self.subTest(question=question):
                result = validate_submission(question, APP / f'examples/result{question}_可行示例.xlsx',
                                             base_q2=base if question == 3 else None)
                self.assertTrue(result['passed'], result['errors'])


if __name__ == '__main__':
    unittest.main()
