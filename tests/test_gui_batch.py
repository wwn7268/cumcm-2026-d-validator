"""Hidden-window regression for the directory-only four-question workflow.

Run from the application directory: python -X utf8 tests/test_gui_batch.py
"""
from pathlib import Path
import shutil
import sys
import tempfile
import time
import tkinter as tk
import traceback
import unittest
from unittest.mock import patch

HERE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(HERE))

from app import ValidatorApp


class DirectoryWorkflowTests(unittest.TestCase):
    def setUp(self):
        generated = HERE / 'tests' / '_generated'
        generated.mkdir(exist_ok=True)
        self.temporary = tempfile.TemporaryDirectory(prefix='gui_batch_', dir=generated)
        self.complete = Path(self.temporary.name) / 'complete'
        self.missing_q3 = Path(self.temporary.name) / 'missing_q3'
        self.complete.mkdir()
        self.missing_q3.mkdir()
        self.nested = self.complete / 'submission' / '附件2'
        self.nested.mkdir(parents=True)
        for question in (1, 2, 3, 4):
            suffix = '正确示例' if question == 1 else '可行示例'
            source = HERE / 'examples' / f'result{question}_{suffix}.xlsx'
            shutil.copy2(source, self.nested / f'result{question}.xlsx')
            if question != 3:
                shutil.copy2(source, self.missing_q3 / f'result{question}.xlsx')

        self.root = tk.Tk()
        self.root.withdraw()
        self.callback_errors = []
        self.root.report_callback_exception = lambda *exc: self.callback_errors.append(
            ''.join(traceback.format_exception(*exc)))
        self.root.createcommand('record_batch_test_bgerror', self.callback_errors.append)
        self.root.tk.eval('proc bgerror {message} {record_batch_test_bgerror $message}')
        self.dialogs = []
        self.dialog_patches = []
        for method in ('showerror', 'showwarning', 'showinfo'):
            replacement = patch(
                f'app.messagebox.{method}',
                side_effect=lambda *args, **kwargs: self.dialogs.append((args, kwargs)))
            replacement.start()
            self.dialog_patches.append(replacement)
        self.app = ValidatorApp(self.root)

    def tearDown(self):
        try:
            self.app._dispose_plot()
            for after_id in self.root.tk.call('after', 'info'):
                self.root.after_cancel(after_id)
        finally:
            self.root.destroy()
            for replacement in reversed(self.dialog_patches):
                replacement.stop()
            self.temporary.cleanup()

    def drain_callbacks(self):
        self.root.update_idletasks()
        self.root.update()
        self.assertEqual([], self.callback_errors)
        self.assertEqual([], self.dialogs)

    def run_worker(self):
        self.app.run_button.invoke()
        self.assertTrue(self.app.busy)
        self.assertIsNone(self.app.result)
        self.assertIsNone(self.app.batch_result)
        self.assertFalse(self.app.export_button.instate(['!disabled']))
        self.assertEqual(['等待结果'] * 4,
                         [self.app.batch_panel.status_labels[q].cget('text')
                          for q in (1, 2, 3, 4)])

        deadline = time.monotonic() + 20

        def finished():
            if not self.app.busy or time.monotonic() >= deadline:
                self.root.quit()
            else:
                self.root.after(20, finished)

        self.root.after(20, finished)
        self.root.mainloop()
        self.assertFalse(self.app.busy, 'The validation worker did not finish in 20 s.')
        self.drain_callbacks()

    def test_nested_directory_details_missing_file_and_retry(self):
        self.assertEqual(str(self.app.batch_page), self.app.tabs.select())
        self.app.directory.set(str(self.complete))
        self.run_worker()

        completed_batch = self.app.batch_result
        self.assertIsNotNone(completed_batch)
        self.assertEqual({1, 2, 3, 4}, set(completed_batch['results']))
        self.assertTrue(all(result['passed'] for result in completed_batch['results'].values()))
        self.assertEqual(['验证通过'] * 4,
                         [self.app.batch_panel.status_labels[q].cget('text')
                          for q in (1, 2, 3, 4)])
        self.assertEqual(str(self.app.batch_page), self.app.tabs.select())
        self.assertIn('4 / 4', self.app.batch_panel.banner.cget('text'))
        for question in (1, 2, 3, 4):
            self.assertEqual((self.nested / f'result{question}.xlsx').resolve(),
                             Path(completed_batch['files'][question]).resolve())

        previous_widget = None
        for question in (1, 2, 3, 4):
            if self.app.canvas is not None:
                self.app.toolbar.push_current()
                self.app.figure.axes[0].set_xlim(20, 100)
                self.app.toolbar.push_current()
                self.app.toolbar.back()
                self.app.canvas.draw_idle()
            self.app.batch_panel.detail_buttons[question].invoke()
            self.drain_callbacks()
            self.assertIs(completed_batch, self.app.batch_result)
            self.assertIs(completed_batch['results'][question], self.app.result)
            self.assertEqual(question, self.app.result['question'])
            self.assertEqual('验证通过', self.app.result_panel.title_label.cget('text'))
            if question == 1:
                self.assertEqual(297, self.app.result['summary']['correct_pair_count'])
                self.assertEqual(431, len(self.app.result['conflicts']))
                self.assertGreater(int(self.app.counts.max()), 1)
                metrics = [label.cget('text') for label in self.app.result_panel.metric_labels]
                self.assertEqual(['297', '297', '0', '0'], metrics)
            self.assertEqual(1, len(self.app.plot_holder.winfo_children()))
            self.assertEqual(1, len(self.app.toolbar_holder.winfo_children()))
            if previous_widget is not None:
                self.assertFalse(previous_widget.winfo_exists())
            previous_widget = self.app.canvas.get_tk_widget()

        self.app.directory.set(str(self.missing_q3))
        self.run_worker()
        incomplete_batch = self.app.batch_result
        self.assertIsNot(completed_batch, incomplete_batch)
        self.assertEqual('invalid', incomplete_batch['results'][3]['status'])
        self.assertFalse(incomplete_batch['results'][3]['passed'])
        self.assertTrue(incomplete_batch['results'][1]['passed'])
        self.assertTrue(incomplete_batch['results'][2]['passed'])
        self.assertTrue(incomplete_batch['results'][4]['passed'])
        self.assertEqual('缺少结果', self.app.batch_panel.status_labels[3].cget('text'))
        self.assertIn('3 / 4', self.app.batch_panel.banner.cget('text'))
        self.assertEqual(str(self.app.batch_page), self.app.tabs.select())
        self.app.batch_panel.detail_buttons[3].invoke()
        self.drain_callbacks()
        self.assertEqual('invalid', self.app.result['status'])
        self.assertEqual('验证未完成', self.app.result_panel.title_label.cget('text'))

        self.app.directory.set(str(self.complete))
        self.run_worker()
        self.assertIsNone(self.app.result)
        self.assertIsNot(incomplete_batch, self.app.batch_result)
        self.assertTrue(all(result['passed'] for result in self.app.batch_result['results'].values()))
        self.assertEqual(str(self.app.batch_page), self.app.tabs.select())
        self.assertIn('4 / 4', self.app.batch_panel.banner.cget('text'))


if __name__ == '__main__':
    unittest.main(verbosity=2)
