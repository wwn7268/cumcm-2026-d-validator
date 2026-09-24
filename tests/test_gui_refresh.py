"""Hidden-window regression checks for replacing a rendered Tk plot.

Run from the application directory: python -X utf8 tests/test_gui_refresh.py
Requires a desktop-capable Tk installation; the test window remains withdrawn.
"""
from pathlib import Path
import sys
import tkinter as tk
import time
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
        expected_title = '验证通过' if result['status'] == 'feasible' else '验证失败'
        self.assertEqual(expected_title, self.app.result_panel.title_label.cget('text'))
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
                conflict = result['conflicts'][0]
                detail = self.app.result_panel.conflict_details[0]
                for value in (conflict['id1'], conflict['id2'],
                              f'[{conflict["t0"]}, {conflict["t1"]})',
                              f'[{conflict["f0"]}, {conflict["f1"]})'):
                    self.assertIn(value, detail)
                self.app.result_panel.conflict_buttons[0].invoke()
                self.drain_callbacks()
                self.assertEqual(str(self.app.plot_page), self.app.tabs.select())
                left, right = self.app.figure.axes[0].get_xlim()
                self.assertLessEqual(left, conflict['t0'])
                self.assertGreaterEqual(right, conflict['t1'])
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
        self.assertEqual(['全部', '4', '140', '537'],
                         [label.cget('text') for label in self.app.result_panel.metric_labels])

    def test_incomplete_input_replaces_passed_dashboard(self):
        self.replace_result(self.q2)
        self.assertEqual(['全部', '6', '126', '678'],
                         [label.cget('text') for label in self.app.result_panel.metric_labels])
        invalid = validate_submission(2, HERE / 'examples' / 'result2_重复编号示例.xlsx')
        self.app.show_result(invalid)
        self.drain_callbacks()
        self.assertEqual('验证未完成', self.app.result_panel.title_label.cget('text'))
        self.assertEqual(['未完成', '—', '—', '—'],
                         [label.cget('text') for label in self.app.result_panel.metric_labels])
        self.assertEqual([], self.app.result_panel.conflict_buttons)

    def test_start_button_replaces_previous_results_in_worker(self):
        for question, filename, expected_status in (
                (4, 'result4_可行示例.xlsx', 'feasible'),
                (2, 'result2_可行示例.xlsx', 'feasible'),
                (3, 'result3_新增重叠示例.xlsx', 'infeasible')):
            self.app.question.set(question)
            self.app._question_changed()
            self.app.submission.set(str(HERE / 'examples' / filename))
            self.app.base.set(str(HERE / 'examples' / 'result2_可行示例.xlsx'))
            self.app.run_button.invoke()
            self.assertIsNone(self.app.result)
            self.assertEqual('正在验证…', self.app.result_panel.title_label.cget('text'))
            deadline = time.monotonic() + 15
            def finished():
                if not self.app.busy or time.monotonic() >= deadline:
                    self.root.quit()
                else:
                    self.root.after(20, finished)
            self.root.after(20, finished)
            self.root.mainloop()
            self.assertFalse(self.app.busy)
            self.assertEqual(expected_status, self.app.result['status'])
            self.drain_callbacks()


if __name__ == '__main__':
    unittest.main(verbosity=2)
