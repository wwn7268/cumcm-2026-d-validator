"""Offline desktop validator for D-question submissions."""
from pathlib import Path
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from engine import validate_submission, read_sheet_names
from reporting import summary_text, export_report, new_output_dir, STATUS, text_of
from visuals import make_figure, occupants

HERE = Path(__file__).resolve().parent
AUTO = '自动识别'


class ValidatorApp:
    def __init__(self, root):
        self.root = root
        root.title('2026年高教社杯数学建模竞赛D题 · 解答验证器 v1.1.2')
        root.geometry('1210x870')
        root.minsize(940, 690)
        style = ttk.Style(root)
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        root.option_add('*Font', ('Microsoft YaHei', 10))
        style.configure('Heading.TLabel', font=('Microsoft YaHei', 17, 'bold'))
        self.question = tk.IntVar(value=2)
        self.submission = tk.StringVar()
        self.base = tk.StringVar()
        self.sheet = tk.StringVar(value=AUTO)
        self.base_sheet = tk.StringVar(value=AUTO)
        self.horizon = tk.StringVar(value='643')
        self.question_note = tk.StringVar()
        self.verdict = tk.StringVar(value='请选择附件，然后开始验证')
        self.pointer = tk.StringVar(value='矩阵支持缩放、平移；鼠标移到占用格可查看装备及使用次数。')
        self.result = None
        self.last_output = None
        self.canvas = self.figure = self.toolbar = None
        self.hover_connection = None
        self.job_queue = queue.Queue()
        self.busy = False
        self.conflict_items = {}
        self.last_hover = None
        self._build()
        self._question_changed()
        root.after(100, self._poll)

    def _build(self):
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill='both', expand=True)
        ttk.Label(outer, text='D题学生附件 · 可行性验证', style='Heading.TLabel').pack(anchor='w')
        ttk.Label(outer, text='按题目模板读取 Excel / CSV，检查规则及全部重复使用区间，绘制时频占用矩阵。').pack(anchor='w', pady=(4, 10))
        form = ttk.LabelFrame(outer, text='附件与验证设置', padding=10)
        form.pack(fill='x')
        form.columnconfigure(1, weight=1)
        radios = ttk.Frame(form)
        radios.grid(row=0, column=1, sticky='w', pady=(0, 5))
        ttk.Label(form, text='验证问题').grid(row=0, column=0, sticky='w', padx=(0, 10))
        for q in (2, 3, 4):
            ttk.Radiobutton(radios, text=f'问题{q}', variable=self.question, value=q, command=self._question_changed).pack(side='left', padx=(0, 18))
        ttk.Label(form, text='学生提交附件').grid(row=1, column=0, sticky='w')
        ttk.Entry(form, textvariable=self.submission).grid(row=1, column=1, sticky='ew', padx=5, pady=4)
        ttk.Button(form, text='选择文件…', command=lambda: self._browse(False)).grid(row=1, column=2, padx=4)
        self.sheet_combo = ttk.Combobox(form, textvariable=self.sheet, values=[AUTO], state='readonly', width=17)
        self.sheet_combo.grid(row=1, column=3, padx=4)
        ttk.Label(form, text='对应第二问方案').grid(row=2, column=0, sticky='w')
        self.base_entry = ttk.Entry(form, textvariable=self.base)
        self.base_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=4)
        self.base_button = ttk.Button(form, text='选择文件…', command=lambda: self._browse(True))
        self.base_button.grid(row=2, column=2, padx=4)
        self.base_combo = ttk.Combobox(form, textvariable=self.base_sheet, values=[AUTO], state='readonly', width=17)
        self.base_combo.grid(row=2, column=3, padx=4)
        line = ttk.Frame(form)
        line.grid(row=3, column=0, columnspan=4, sticky='ew', pady=(6, 0))
        ttk.Label(line, text='时间窗 [0, H)，H =').pack(side='left')
        ttk.Entry(line, textvariable=self.horizon, width=7).pack(side='left', padx=4)
        ttk.Label(line, text='默认643沿用原计划最晚结束时刻；请与学生建模假设核对（整数1～10000）。').pack(side='left', padx=8)
        ttk.Label(form, textvariable=self.question_note, foreground='#5c6070').grid(row=4, column=0, columnspan=4, sticky='w', pady=(5, 0))
        buttons = ttk.Frame(outer)
        buttons.pack(fill='x', pady=10)
        self.run_button = ttk.Button(buttons, text='开始验证', command=self._start)
        self.run_button.pack(side='left')
        ttk.Button(buttons, text='载入可行示例', command=self._example).pack(side='left', padx=7)
        self.export_button = ttk.Button(buttons, text='导出报告…', command=self._export, state='disabled')
        self.export_button.pack(side='left')
        self.open_button = ttk.Button(buttons, text='打开结果目录', command=self._open_output, state='disabled')
        self.open_button.pack(side='left', padx=7)
        ttk.Button(buttons, text='使用说明', command=lambda: os.startfile(str(HERE / '使用说明.md'))).pack(side='right')
        self.verdict_label = ttk.Label(outer, textvariable=self.verdict, font=('Microsoft YaHei', 12, 'bold'))
        self.verdict_label.pack(anchor='w', pady=(0, 8))
        self.tabs = ttk.Notebook(outer)
        self.tabs.pack(fill='both', expand=True)
        self.summary_page = ttk.Frame(self.tabs)
        self.issue_page = ttk.Frame(self.tabs, padding=6)
        self.plot_page = ttk.Frame(self.tabs)
        self.tabs.add(self.summary_page, text=' 判定与统计 ')
        self.tabs.add(self.issue_page, text=' 问题与冲突 ')
        self.tabs.add(self.plot_page, text=' 时频矩阵 ')
        self.summary_box = ScrolledText(self.summary_page, wrap='word', padx=12, pady=10, state='disabled')
        self.summary_box.pack(fill='both', expand=True)
        ttk.Label(self.issue_page, text='附件格式与调整规则').pack(anchor='w')
        self.error_tree = self._tree(self.issue_page, [('sheet', '工作表', 120), ('row', '行号', 55), ('id', '装备/序号', 100), ('message', '问题', 760)], height=5)
        ttk.Label(self.issue_page, text='时频交叠事件（双击定位到矩阵；同一装备对可能发生多次交叠）').pack(anchor='w', pady=(8, 3))
        self.conflict_tree = self._tree(self.issue_page, [('kind', '类型', 110), ('id1', '装备1', 120), ('use1', '次数1', 55), ('id2', '装备2', 120), ('use2', '次数2', 55), ('time', '交叠时间', 150), ('freq', '交叠频段', 150)], height=9)
        self.conflict_tree.bind('<Double-1>', self._locate_conflict)
        self.plot_holder = ttk.Frame(self.plot_page)
        self.toolbar_holder = ttk.Frame(self.plot_page)
        # Reserve the controls before the canvas takes the remaining height.
        ttk.Label(self.plot_page, textvariable=self.pointer, padding=5).pack(side='bottom', fill='x')
        self.toolbar_holder.pack(side='bottom', fill='x')
        self.plot_holder.pack(fill='both', expand=True)

    @staticmethod
    def _tree(parent, columns, height):
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show='headings', height=height)
        for key, label, width in columns:
            tree.heading(key, text=label)
            tree.column(key, width=width, minwidth=45, stretch=True)
        scroll = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky='nsew')
        scroll.grid(row=0, column=1, sticky='ns')
        horizontal.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _question_changed(self):
        question = self.question.get()
        enabled = question == 3
        self.base_entry.configure(state='normal' if enabled else 'disabled')
        self.base_button.configure(state='normal' if enabled else 'disabled')
        self.base_combo.configure(state='readonly' if enabled else 'disabled')
        notes = {
            2: '第二问以原始150台装备为基准；每台可调频、调时或撤销，未列出的装备保持原计划。',
            3: '第三问固定所选第二问方案；新增C类与原有装备均须无冲突。',
            4: '第四问以原始150台为基准，无需第二问附件；仅C类可改间隔，每台只能调整一个参数或撤销。',
        }
        self.question_note.set(notes[question])

    def _browse(self, base):
        path = filedialog.askopenfilename(title='选择第二问执行方案' if base else '选择学生提交附件', filetypes=[('Excel或CSV附件', '*.xlsx *.xlsm *.csv'), ('所有文件', '*.*')])
        if not path:
            return
        (self.base if base else self.submission).set(path)
        variable, combo = (self.base_sheet, self.base_combo) if base else (self.sheet, self.sheet_combo)
        variable.set(AUTO)
        try:
            names = read_sheet_names(Path(path))
            combo.configure(values=[AUTO, *names])
        except Exception as exc:
            combo.configure(values=[AUTO])
            messagebox.showwarning('无法读取工作表', str(exc))

    def _example(self):
        q = self.question.get()
        self.submission.set(str(HERE / 'examples' / f'result{q}_可行示例.xlsx'))
        self.base.set(str(HERE / 'examples' / 'result2_可行示例.xlsx'))
        self.sheet.set(AUTO)
        self.base_sheet.set(AUTO)
        self.horizon.set('643')
        self.verdict.set('已载入示例，点击“开始验证”。')

    def _start(self):
        if self.busy:
            return
        try:
            horizon = int(self.horizon.get())
            if not 1 <= horizon <= 10000:
                raise ValueError()
        except ValueError:
            messagebox.showerror('时间窗格式错误', 'H须为1～10000的整数。')
            return
        if not self.submission.get().strip():
            messagebox.showinfo('未选择附件', '请先选择学生附件。')
            return
        if self.question.get() == 3 and not self.base.get().strip():
            messagebox.showinfo('缺少第二问方案', '验证第三问时，需要选择学生采用的第二问执行方案。')
            return
        args = dict(question=self.question.get(), submission=Path(self.submission.get().strip()),
                    base_q2=Path(self.base.get().strip()) if self.question.get() == 3 else None,
                    horizon=horizon, sheet=None if self.sheet.get() == AUTO else self.sheet.get(),
                    base_sheet=None if self.base_sheet.get() == AUTO else self.base_sheet.get())
        self.busy = True
        self.run_button.configure(state='disabled')
        self.export_button.configure(state='disabled')
        self.verdict.set('正在解析附件并检查完整使用计划…')
        self.verdict_label.configure(foreground='#37475b')
        def worker():
            try:
                self.job_queue.put(('result', validate_submission(**args)))
            except Exception as exc:
                self.job_queue.put(('error', f'{type(exc).__name__}: {exc}'))
        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            kind, payload = self.job_queue.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            self.run_button.configure(state='normal')
            if kind == 'error':
                self.result = None
                self.verdict.set('校验未完成，不能据此认定方案可行')
                self.verdict_label.configure(foreground='#a6323c')
                messagebox.showerror('验证失败', payload)
            else:
                try:
                    self.show_result(payload)
                except Exception as exc:
                    self.verdict.set('验证结束，但界面绘图失败；可导出文字结果')
                    self.verdict_label.configure(foreground='#b62f40')
                    self.result = payload
                    self.export_button.configure(state='normal')
                    messagebox.showerror('显示失败', str(exc))
        self.root.after(100, self._poll)

    def show_result(self, result):
        self.result = result
        status = result.get('status')
        q = result.get('question')
        self.verdict.set(f'问题{q}：{STATUS.get(status, status)}　｜　该判定不证明最优性')
        self.verdict_label.configure(foreground='#197456' if status == 'feasible' else '#b62f40')
        self.summary_box.configure(state='normal')
        self.summary_box.delete('1.0', 'end')
        self.summary_box.insert('end', summary_text(result))
        self.summary_box.configure(state='disabled')
        for tree in (self.error_tree, self.conflict_tree):
            tree.delete(*tree.get_children())
        for e in result.get('errors', []):
            if isinstance(e, dict):
                values = [e.get(k, '') for k in ('sheet', 'row', 'id', 'message')]
            else:
                values = ['', '', '', text_of(e)]
            self.error_tree.insert('', 'end', values=values)
        self.conflict_items.clear()
        for c in result.get('conflicts', []):
            values = [c.get(k, '') for k in ('kind', 'id1', 'use1', 'id2', 'use2')]
            values += [f'[{c.get("t0")}, {c.get("t1")})', f'[{c.get("f0")}, {c.get("f1")})']
            item = self.conflict_tree.insert('', 'end', values=values)
            self.conflict_items[item] = c
        self._dispose_plot()
        self.figure, self.counts, _ = make_figure(result)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.plot_holder)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.toolbar_holder, pack_toolbar=False)
        self.toolbar.pack(fill='x')
        self.hover_connection = self.canvas.mpl_connect('motion_notify_event', self._hover)
        self.canvas.draw()
        self.last_hover = None
        self.export_button.configure(state='normal')
        self.tabs.select(self.summary_page)

    def _dispose_plot(self):
        if self.canvas is not None:
            if self.hover_connection is not None:
                self.canvas.mpl_disconnect(self.hover_connection)
            pending_draw = getattr(self.canvas, '_idle_draw_id', None)
            if pending_draw is not None:
                self.canvas.get_tk_widget().after_cancel(pending_draw)
                self.canvas._idle_draw_id = None
            # Figure.clear() updates canvas.toolbar; detach it before destroying Tk widgets.
            self.canvas.toolbar = None
        if self.figure is not None:
            self.figure.clear()
        if self.toolbar is not None:
            self.toolbar.destroy()
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
        self.canvas = self.figure = self.toolbar = None
        self.hover_connection = None

    def _hover(self, event):
        if not self.result or event.inaxes not in self.figure.axes[:2] or event.xdata is None or event.ydata is None:
            return
        t, f = int(event.xdata), int(event.ydata)
        if not (0 <= t < self.counts.shape[0] and 0 <= f < 100) or (t, f) == self.last_hover:
            return
        self.last_hover = (t, f)
        names = occupants(self.result.get('plans', []), t, f)
        shown = '；'.join(names[:8]) or '空闲'
        if len(names) > 8:
            shown += f' 等{len(names)}个占用'
        self.pointer.set(f'时间 [{t},{t+1}) × 频段 [{f},{f+1})　占用数 {self.counts[t,f]}：{shown}')

    def _locate_conflict(self, event=None):
        chosen = self.conflict_tree.selection()
        if not chosen or not self.figure:
            return
        c = self.conflict_items[chosen[0]]
        horizon = self.result.get('settings', {}).get('horizon', 643)
        for ax in self.figure.axes[:2]:
            ax.set_xlim(max(0, c['t0'] - 12), min(horizon, c['t1'] + 12))
            ax.set_ylim(max(0, c['f0'] - 8), min(100, c['f1'] + 8))
        self.tabs.select(self.plot_page)
        self.canvas.draw_idle()

    def _export(self):
        if not self.result:
            return
        root = filedialog.askdirectory(title='选择报告保存目录（将在其中新建独立结果文件夹）', initialdir=str(HERE))
        if not root:
            return
        self.root.configure(cursor='watch')
        self.root.update_idletasks()
        try:
            source = self.result.get('settings', {}).get('input_file') or '学生附件'
            if not isinstance(source, (str, Path)):
                source = '学生附件'
            out = new_output_dir(root, source, self.result.get('question', 2))
            self.last_output = export_report(self.result, out)
            self.open_button.configure(state='normal')
            messagebox.showinfo('报告已保存', f'{self.last_output}\n\n包含矩阵图、占用计数CSV、Excel明细、JSON及文字摘要。')
        except Exception as exc:
            messagebox.showerror('导出失败', str(exc))
        finally:
            self.root.configure(cursor='')

    def _open_output(self):
        if self.last_output:
            os.startfile(str(self.last_output))


if __name__ == '__main__':
    window = tk.Tk()
    ValidatorApp(window)
    window.mainloop()
