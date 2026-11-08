"""Directory validation, dependency isolation, and summary export regressions."""
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from uuid import uuid4
import csv
import io
import json
import shutil
import stat
import sys
import unittest

from openpyxl import load_workbook

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
import batch
import batch_reporting
import validate


class DirectoryTests(unittest.TestCase):
    def setUp(self):
        generated = (APP / 'tests/_generated/batch').resolve()
        self.root = generated / uuid4().hex
        self.root.mkdir(parents=True)

        def cleanup():
            target = self.root.resolve()
            if target.parent != generated or not target.is_relative_to(APP.resolve()):
                raise AssertionError('Refusing to remove a test directory outside this app.')
            shutil.rmtree(target)

        self.addCleanup(cleanup)
        self.folder = self.root / '学生提交'
        self.folder.mkdir()

    def example(self, question, name=None, source=None, folder=None):
        source = source or ('result1_正确示例.xlsx' if question == 1 else f'result{question}_可行示例.xlsx')
        target = (folder or self.folder) / (name or f'result{question}.xlsx')
        shutil.copyfile(APP / 'examples' / source, target)
        return target

    def complete(self, folder=None):
        for question in (1, 2, 3, 4):
            self.example(question, folder=folder)

    def test_four_questions_pass_and_q3_uses_same_directory_base(self):
        self.complete()
        result = batch.validate_directory(self.folder)
        self.assertTrue(result['passed'], result['results'])
        self.assertEqual(set(result['results']), {1, 2, 3, 4})
        self.assertEqual(result['results'][1]['summary']['correct_pair_count'], 297)
        self.assertEqual(len(result['results'][1]['conflicts']), 431)
        self.assertEqual(result['results'][3]['settings']['base_file'], result['files'][2])
        self.assertEqual(result['results'][3]['summary']['added_count'], 136)
        for item in result['results'].values():
            self.assertEqual(item['status'], 'feasible')
            self.assertNotIn('batch_state', item)

    def test_bad_q2_blocks_q3_but_q4_still_runs(self):
        self.example(2, source='result2_频移超限示例.xlsx')
        self.example(3)
        self.example(4)
        result = batch.validate_directory(self.folder)
        self.assertFalse(result['passed'])
        self.assertEqual(result['results'][2]['status'], 'infeasible')
        self.assertEqual(result['results'][3]['status'], 'invalid')
        self.assertEqual(result['results'][3]['batch_state'], 'blocked')
        self.assertTrue(result['results'][4]['passed'])

    def test_missing_q2_blocks_present_q3_and_missing_q4_is_explicit(self):
        self.example(3)
        nested = self.folder / '嵌套目录'
        nested.mkdir()
        shutil.copyfile(APP / 'examples/result4_可行示例.xlsx', nested / 'result4.xlsx')
        result = batch.validate_directory(self.folder)
        self.assertEqual(result['results'][2]['batch_state'], 'missing')
        self.assertEqual(result['results'][3]['batch_state'], 'blocked')
        self.assertEqual(result['results'][4]['batch_state'], 'missing')
        self.assertIsNone(result['files'][2])
        self.assertIsNone(result['files'][4])

    def test_ambiguous_q2_is_not_guessed(self):
        self.complete()
        self.example(2, '问题二_最终.xlsx')
        result = batch.validate_directory(self.folder)
        item = result['results'][2]
        self.assertEqual(item['batch_state'], 'ambiguous')
        self.assertEqual(len(item['candidates']), 2)
        self.assertIn('问题二_最终.xlsx', item['errors'][0]['message'])
        self.assertIsNone(result['files'][2])
        self.assertEqual(result['results'][3]['batch_state'], 'blocked')
        self.assertTrue(result['results'][4]['passed'])

    def test_corrupt_q2_is_invalid_and_q4_is_independent(self):
        self.complete()
        (self.folder / 'result2.xlsx').write_bytes(b'not an Excel workbook')
        result = batch.validate_directory(self.folder)
        self.assertEqual(result['results'][2]['status'], 'invalid')
        self.assertEqual(result['results'][2]['batch_state'], 'error')
        self.assertEqual(result['results'][3]['batch_state'], 'blocked')
        self.assertTrue(result['results'][4]['passed'])

    def test_chinese_csv_xlsm_names_and_custom_original(self):
        self.example(1, '问题一_最终清单.xlsx')
        original = APP / 'data/original_plans.json'
        plans = json.loads(original.read_text(encoding='utf-8'))
        with (self.folder / '第二问_答案.csv').open('w', encoding='utf-8-sig', newline='') as handle:
            writer = csv.writer(handle)
            writer.writerow(['用频装备编号', '调整后频段区间', '调整后时间区间', '是否撤销用频计划'])
            writer.writerows([p['id'], '', '', '是'] for p in plans)
        with (self.folder / '问题三.csv').open('w', encoding='utf-8-sig', newline='') as handle:
            writer = csv.writer(handle)
            writer.writerow(['新增用频装备序号', '调整后频段区间', '调整后时间区间'])
            writer.writerow([1, '[0,3)', '[0,2)'])
        self.example(4, '第四问_提交.xlsm')
        result = batch.validate_directory(self.folder, original_path=original)
        self.assertTrue(result['passed'], result['results'])
        self.assertEqual(result['results'][3]['summary']['added_count'], 1)
        self.assertEqual(result['results'][2]['settings']['original_file'], str(original))

    def test_unrelated_and_office_lock_names_do_not_create_ambiguity(self):
        self.complete()
        for name in ('~$result2.xlsx', 'results10.xlsx', 'result20.xlsx', 'draft_result2.xlsx', '问题二三四.xlsx'):
            (self.folder / name).write_bytes(b'unrelated')
        result = batch.validate_directory(self.folder)
        self.assertTrue(result['passed'], result['results'])

    def test_unexpected_engine_exception_is_isolated(self):
        self.complete()
        real_validate = batch.validate_submission

        def isolated(question, *args, **kwargs):
            if question == 2:
                raise RuntimeError('模拟单问异常')
            return real_validate(question, *args, **kwargs)

        with patch('batch.validate_submission', side_effect=isolated) as call:
            result = batch.validate_directory(self.folder)
        self.assertEqual([args.args[0] for args in call.call_args_list], [1, 2, 4])
        self.assertEqual(result['results'][2]['batch_state'], 'error')
        self.assertEqual(result['results'][3]['batch_state'], 'blocked')
        self.assertTrue(result['results'][4]['passed'])

    def test_missing_directory_returns_explanations_for_all_questions(self):
        result = batch.validate_directory(self.root / '不存在')
        self.assertFalse(result['passed'])
        self.assertTrue(all(item['batch_state'] == 'error' for item in result['results'].values()))

    def test_summary_export_is_unique_and_keeps_missing_explanations(self):
        result = batch.validate_directory(self.folder)
        output = self.root / '报告'
        first = batch_reporting.export_directory_report(result, output)
        second = batch_reporting.export_directory_report(result, output)
        self.assertNotEqual(first, second)
        for name in ('目录验证报告.json', '目录验证摘要.txt', '目录验证汇总.xlsx'):
            self.assertTrue((first / name).is_file())
        for question in (1, 2, 3, 4):
            self.assertIn('未找到', (first / f'问题{question}/验证摘要.txt').read_text(encoding='utf-8-sig'))
        workbook = load_workbook(first / '目录验证汇总.xlsx', read_only=True)
        try:
            self.assertEqual(workbook['四问汇总'].max_row, 5)
        finally:
            workbook.close()

    def test_per_question_export_failure_does_not_stop_remaining_reports(self):
        self.complete()
        result = batch.validate_directory(self.folder)
        calls = []

        def exporter(item, directory):
            calls.append(item['question'])
            if item['question'] == 2:
                raise ValueError('模拟导出失败')
            (directory / '已导出.txt').write_text('ok', encoding='utf-8')

        with patch('batch_reporting.reporting.export_report', side_effect=exporter):
            destination = batch_reporting.export_directory_report(result, self.root / '报告')
        saved = json.loads((destination / '目录验证报告.json').read_text(encoding='utf-8'))
        self.assertEqual(calls, [1, 2, 3, 4])
        self.assertEqual(len(saved['export_errors']), 1)
        self.assertTrue((destination / '问题4/已导出.txt').is_file())
        self.assertNotIn('export_errors', result)
        workbook = load_workbook(destination / '目录验证汇总.xlsx', read_only=True)
        try:
            row = list(workbook['四问汇总'].iter_rows(min_row=2, max_row=2, values_only=True))[0]
            self.assertEqual(row[2], '冲突清单准确')
            self.assertEqual(row[4:10], (297, 297, 297, 0, 0, 0))
            self.assertIsNone(row[12])  # The original conflicts are not remaining errors.
        finally:
            workbook.close()
        text = (destination / '目录验证摘要.txt').read_text(encoding='utf-8-sig')
        self.assertIn('问题1：冲突清单准确', text)
        self.assertIn('431（作为清单真值', text)

    def test_cli_supports_directory_and_legacy_modes_and_rejects_mixing(self):
        self.complete()
        with redirect_stdout(io.StringIO()), patch('validate.export_directory_report', return_value=self.root / '报告'):
            self.assertEqual(validate.main(['--directory', str(self.folder)]), 0)
        with redirect_stdout(io.StringIO()), patch('validate.export_report'):
            self.assertEqual(validate.main(['2', str(self.folder / 'result2.xlsx')]), 0)
            self.assertEqual(validate.main(['1', str(self.folder / 'result1.xlsx')]), 0)
        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
            validate.main(['2', str(self.folder / 'result2.xlsx'), '--directory', str(self.folder)])
        self.assertEqual(error.exception.code, 2)

    def test_results_and_result_names_are_both_recognized(self):
        for question, name in enumerate(('results1.xlsx', '第二问_最终稿.xlsx',
                                         'RESULTS_3_新增方案.xlsx', '问题四-调整结果.xlsx'), 1):
            self.example(question, name)
        result = batch.validate_directory(self.folder)
        self.assertTrue(result['passed'], result['results'])

    def test_complete_root_wins_without_scanning_children(self):
        self.complete()
        child = self.folder / '另一套'
        child.mkdir()
        self.complete(child)
        with patch('batch._scan_directory', wraps=batch._scan_directory) as scan:
            result = batch.validate_directory(self.folder)
        self.assertTrue(result['passed'])
        self.assertEqual(result['directory'], str(self.folder))
        self.assertEqual(scan.call_count, 1)

    def test_unique_nested_complete_set_wins_over_partial_root(self):
        self.example(2, source='result2_频移超限示例.xlsx')
        child = self.folder / '压缩包解压' / '正式提交'
        child.mkdir(parents=True)
        self.complete(child)
        result = batch.validate_directory(self.folder)
        self.assertTrue(result['passed'], result['results'])
        self.assertEqual(result['requested_directory'], str(self.folder))
        self.assertEqual(result['directory'], str(child))
        self.assertIn('唯一', result['search_note'])
        self.assertTrue(all(Path(path).parent == child for path in result['files'].values()))
        self.assertEqual(result['results'][3]['settings']['base_file'], str(child / 'result2.xlsx'))

    def test_multiple_complete_subfolders_are_directory_ambiguity(self):
        for name in ('甲同学', '乙同学'):
            child = self.folder / name
            child.mkdir()
            self.complete(child)
        with patch('batch.validate_submission') as validate_file:
            result = batch.validate_directory(self.folder)
        validate_file.assert_not_called()
        self.assertEqual(len(result['candidate_directories']), 2)
        self.assertTrue(all(item['batch_state'] == 'ambiguous' for item in result['results'].values()))
        self.assertTrue(all(value is None for value in result['files'].values()))

    def test_single_partial_subfolder_is_selected_with_missing_questions(self):
        child = self.folder / '唯一部分结果'
        child.mkdir()
        self.example(1, folder=child)
        self.example(4, folder=child)
        result = batch.validate_directory(self.folder)
        self.assertEqual(result['directory'], str(child))
        self.assertTrue(result['results'][1]['passed'])
        self.assertTrue(result['results'][4]['passed'])
        self.assertEqual(result['results'][2]['batch_state'], 'missing')
        self.assertEqual(result['results'][3]['batch_state'], 'missing')

    def test_partial_subfolders_are_not_combined_into_one_set(self):
        for name, questions in (('前两问', (1, 2)), ('后两问', (3, 4))):
            child = self.folder / name
            child.mkdir()
            for question in questions:
                self.example(question, folder=child)
        result = batch.validate_directory(self.folder)
        self.assertFalse(result['passed'])
        self.assertEqual(len(result['candidate_directories']), 2)
        self.assertTrue(all(item['batch_state'] == 'ambiguous' for item in result['results'].values()))

    def test_read_error_is_not_hidden_by_an_accessible_complete_subfolder(self):
        complete = self.folder / 'a_完整结果'
        complete.mkdir()
        self.complete(complete)
        denied = self.folder / 'z_不可读'
        denied.mkdir()
        real_scan = batch._scan_directory

        def scan(folder):
            if folder == denied:
                raise PermissionError(f'拒绝读取：{folder}')
            return real_scan(folder)

        with patch('batch._scan_directory', side_effect=scan):
            result = batch.validate_directory(self.folder)
        self.assertTrue(result['search_errors'])
        self.assertIn(str(denied), result['search_note'])
        self.assertTrue(all(item['batch_state'] == 'error' for item in result['results'].values()))

    def test_junction_and_symlink_are_skipped(self):
        self.example(4)
        linked = self.folder / '模拟目录联接'
        linked.mkdir()
        self.complete(linked)
        real_lstat = Path.lstat

        def lstat(path):
            if path == linked:
                return SimpleNamespace(st_mode=stat.S_IFDIR, st_file_attributes=0x400)
            return real_lstat(path)

        with patch.object(Path, 'lstat', lstat):
            result = batch.validate_directory(self.folder)
        self.assertEqual(result['directory'], str(self.folder))
        self.assertEqual(result['skipped_links'], [str(linked)])
        self.assertEqual(result['results'][1]['batch_state'], 'missing')
        self.assertTrue(result['results'][4]['passed'])
        self.assertTrue(batch._is_link(SimpleNamespace(st_mode=stat.S_IFLNK)))


if __name__ == '__main__':
    unittest.main()
