"""Read valid Excel files without worksheet dimensions and retain input limits."""
from pathlib import Path
import re
import shutil
import sys
import unittest
from uuid import uuid4
from zipfile import ZipFile

from openpyxl import load_workbook

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
import engine


class MissingExcelDimensionsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        generated = (APP / 'tests' / '_generated' / 'excel_dimensions').resolve()
        cls.folder = generated / uuid4().hex
        cls.folder.mkdir(parents=True)

        def cleanup():
            target = cls.folder.resolve()
            if target.parent != generated or not target.is_relative_to(APP.resolve()):
                raise AssertionError('Refusing to remove files outside the generated test folder.')
            shutil.rmtree(target)

        cls.addClassCleanup(cleanup)
        cls.sources, cls.dimensionless = {}, {}
        for question in (1, 2, 3, 4):
            suffix = '正确示例' if question == 1 else '可行示例'
            cls.sources[question] = APP / 'examples' / f'result{question}_{suffix}.xlsx'
            target = cls.folder / f'result{question}_no_dimensions.xlsx'
            cls.remove_dimensions(cls.sources[question], target)
            cls.dimensionless[question] = target

    @staticmethod
    def remove_dimensions(source, target):
        with ZipFile(source) as old, ZipFile(target, 'w') as new:
            removed = 0
            for item in old.infolist():
                content = old.read(item.filename)
                if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', item.filename):
                    content, count = re.subn(rb'<dimension\b[^>]*/>', b'', content)
                    removed += count
                new.writestr(item, content)
        if not removed:
            raise AssertionError('The fixture must originally contain a worksheet dimension.')

    def changed_fixture(self, question, name, edit):
        intermediate = self.folder / f'q{question}_{name}_with_dimensions.xlsx'
        workbook = load_workbook(self.sources[question])
        try:
            edit(workbook.active)
            workbook.save(intermediate)
        finally:
            workbook.close()
        target = self.folder / f'q{question}_{name}_no_dimensions.xlsx'
        self.remove_dimensions(intermediate, target)
        return target

    def validate(self, question, path):
        return engine.validate_submission(
            question, path,
            base_q2=self.dimensionless[2] if question == 3 else None)

    def assert_rejected(self, result, code):
        self.assertEqual('invalid', result['status'])
        self.assertFalse(result['passed'])
        self.assertIn(code, {error['code'] for error in result['errors']})

    def test_all_four_official_formats_without_dimensions_are_read(self):
        for question in (1, 2, 3, 4):
            with self.subTest(question=question):
                path = self.dimensionless[question]
                workbook = load_workbook(path, read_only=True)
                try:
                    self.assertIsNone(workbook.active.max_column)
                    self.assertIsNone(workbook.active.max_row)
                finally:
                    workbook.close()
                result = self.validate(question, path)
                self.assertTrue(result['passed'], result['errors'])
                expected = self.validate(question, self.sources[question])
                self.assertEqual(expected['summary'], result['summary'])

    def test_formulas_and_excel_errors_still_rejected(self):
        for question in (1, 2, 3, 4):
            for name, value, code in (
                    ('formula', '=1+1', 'formula_cell'),
                    ('excel_error', '#DIV/0!', 'excel_error_cell')):
                with self.subTest(question=question, case=name):
                    path = self.changed_fixture(
                        question, name, lambda sheet, value=value: sheet.cell(2, 1, value))
                    self.assert_rejected(self.validate(question, path), code)

    def test_actual_row_limit_is_enforced_without_dimensions(self):
        for question in (1, 2, 3, 4):
            with self.subTest(question=question):
                row = engine.MAX_ROWS + 21
                path = self.changed_fixture(
                    question, 'row_limit', lambda sheet: sheet.cell(row, 1, 'extra'))
                result = self.validate(question, path)
                self.assert_rejected(result, 'table_limit')
                self.assertTrue(any(error['code'] == 'table_limit' and error['row'] == row
                                    for error in result['errors']))

    def test_actual_column_limit_is_enforced_without_dimensions(self):
        for question in (1, 2, 3, 4):
            with self.subTest(question=question):
                path = self.changed_fixture(
                    question, 'column_limit',
                    lambda sheet: sheet.cell(10, engine.MAX_COLUMNS + 1, 'extra'))
                result = self.validate(question, path)
                self.assert_rejected(result, 'table_limit')
                self.assertTrue(any(error['code'] == 'table_limit' and error['row'] == 10
                                    for error in result['errors']))

    def test_header_is_required_without_dimensions(self):
        for question in (1, 2, 3, 4):
            with self.subTest(question=question):
                path = self.changed_fixture(
                    question, 'no_header', lambda sheet: sheet.delete_rows(1))
                self.assert_rejected(self.validate(question, path), 'missing_header')


if __name__ == '__main__':
    unittest.main(verbosity=2)
