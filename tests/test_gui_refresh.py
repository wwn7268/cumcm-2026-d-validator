"""Hidden-window regression checks for replacing a rendered Tk plot.

Run from the application directory: python -X utf8 tests/test_gui_refresh.py
Requires a desktop-capable Tk installation; the test window remains withdrawn.
"""
from pathlib import Path
import sys
import tkinter as tk
import traceback
import unittest

HERE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(HERE))

from app import ValidatorApp
from engine import validate_submission


class PlotRefreshTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        base = HERE / 'examples' / 'result2_可行示例.xlsx'
        cls.q2 = validate_submission(2, base)
        cls.q3 = validate_submission(
            3, HERE / 'examples' / 'result3_可行示例.xlsx', base_q2=base)
        cls.conflicting = validate_submission(
            3, HERE / 'examples' / 'result3_新增重叠示例.xlsx', base_q2=base)
        cls.q4 = validate_submission(4, HERE / 'examples' / 'result4_可行示例.xlsx')
        if not (cls.q2['passed'] and cls.q3['passed'] and cls.q4['passed']):
            raise AssertionError('The feasible input fixtures must pass validation.')
        if cls.conflicting['passed'] or not cls.conflicting['conflicts']:
            raise AssertionError('The conflict fixture must contain a detected overlap.')

    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.callback_errors = []
        self.root.report_callback_exception = lambda *exc: self.callback_errors.append(
            ''.join(traceback.format_exception(*exc)))
        # Tcl idle callbacks can fail before entering Python's callback handler.
        self.root.createcommand('record_test_bgerror', self.callback_errors.append)
        self.root.tk.eval('proc bgerror {message} {record_test_bgerror $message}')
        self.app = ValidatorApp(self.root)

    def tearDown(self):
        try:
            self.app._dispose_plot()
            for after_id in self.root.tk.call('after', 'info'):
                self.root.after_cancel(after_id)
        finally:
            self.root.destroy()

    def drain_callbacks(self):
        self.root.update_idletasks()
        self.root.update()
        self.assertEqual([], self.callback_errors)

    def queue_navigation_draw(self):
        toolbar = self.app.toolbar
        toolbar.push_current()
        self.app.figure.axes[0].set_xlim(30, 120)
        toolbar.push_current()
        toolbar.back()
        toolbar.forward()
        toolbar.home()
        self.app.canvas.draw_idle()

    def replace_result(self, result):
        old_canvas = self.app.canvas
        old_toolbar = self.app.toolbar
        old_widget = old_canvas.get_tk_widget() if old_canvas else None
        self.app.question.set(result['question'])
        self.app._question_changed()
        self.app.show_result(result)
        self.drain_callbacks()
        self.assertIs(result, self.app.result)
        self.assertTrue(self.app.canvas.get_tk_widget().winfo_exists())
        self.assertTrue(self.app.toolbar.winfo_exists())
        self.assertEqual(1, len(self.app.plot_holder.winfo_children()))
        self.assertEqual(1, len(self.app.toolbar_holder.winfo_children()))
        expected_state = 'normal' if result['question'] == 3 else 'disabled'
        self.assertEqual(expected_state, str(self.app.base_entry.cget('state')))
        if old_widget is not None:
            self.assertFalse(old_widget.winfo_exists())
            self.assertFalse(old_toolbar.winfo_exists())

    def test_repeated_result_after_navigation_and_pending_draw(self):
        self.replace_result(self.q2)
        for _ in range(2):
            self.queue_navigation_draw()
            # Replace immediately, before the queued draw has run.
            self.replace_result(self.q2)

    def test_switch_questions_conflict_location_and_return_to_feasible(self):
        self.replace_result(self.q2)
        for result in (self.q3, self.conflicting, self.q4, self.q2):
            self.queue_navigation_draw()
            self.replace_result(result)
            if result is self.conflicting:
                self.assertGreater(int(self.app.counts.max()), 1)
                self.app.conflict_tree.selection_set(
                    self.app.conflict_tree.get_children()[0])
                self.app._locate_conflict()
                self.drain_callbacks()
                self.assertEqual(str(self.app.plot_page), self.app.tabs.select())
            else:
                self.assertLessEqual(int(self.app.counts.max()), 1)

    def test_fourth_question_example_statistics_and_matrix(self):
        self.app.question.set(4)
        self.app._question_changed()
        self.app._example()
        self.assertEqual('result4_可行示例.xlsx', Path(self.app.submission.get()).name)
        self.assertTrue(Path(self.app.submission.get()).is_file())
        self.assertEqual('disabled', str(self.app.base_entry.cget('state')))
        self.replace_result(self.q4)
        self.assertEqual(16176, int((self.app.counts > 0).sum()))
        text = self.app.summary_box.get('1.0', 'end')
        self.assertIn('第四问指标', text)
        self.assertIn('改间隔41', text)
        self.assertIn('(4, 0, 2, 140, 18, 34,', text)


if __name__ == '__main__':
    unittest.main(verbosity=2)
