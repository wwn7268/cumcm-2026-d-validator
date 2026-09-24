"""Scrollable result dashboard for the desktop validator."""
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from result_view import build_result_view


class ResultPanel(ttk.Frame):
    BACKGROUND = '#f2f5fa'
    INK = '#182b45'
    MUTED = '#617187'
    THEMES = {
        'feasible': ('#064e3b', '#a7f3d0', '#10b981'),
        'infeasible': ('#821d32', '#fecdd3', '#f43f5e'),
        'invalid': ('#783d12', '#fde68a', '#f59e0b'),
        'pending': ('#173e6d', '#bfdbfe', '#60a5fa'),
        'ready': ('#173e6d', '#bfdbfe', '#60a5fa'),
    }

    def __init__(self, parent, locate_conflict, show_issues):
        super().__init__(parent)
        self.locate_conflict = locate_conflict
        self.show_issues = show_issues
        self.canvas = tk.Canvas(self, bg=self.BACKGROUND, highlightthickness=0)
        scroll = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        self.body = tk.Frame(self.canvas, bg=self.BACKGROUND)
        self.window = self.canvas.create_window(0, 0, anchor='nw', window=self.body)
        self.body.bind('<Configure>', self._resize_body)
        self.canvas.bind('<Configure>', self._resize_canvas)
        self.canvas.bind('<MouseWheel>', self._wheel)
        self.wrap_labels = []
        self.progress = None
        self.summary_box = None
        self.view = None
        self.show_waiting()

    def _resize_body(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _resize_canvas(self, event):
        self.canvas.itemconfigure(self.window, width=event.width)
        for label, inset in self.wrap_labels:
            label.configure(wraplength=max(240, event.width - inset))

    def _wheel(self, event):
        if self.body.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_scroll(-int(event.delta / 120), 'units')
        return 'break'

    def _bind_scroll(self, widget):
        # Bind only this panel's widgets, so the matrix keeps its own mouse actions.
        if not isinstance(widget, (tk.Text, ttk.Scrollbar)):
            widget.bind('<MouseWheel>', self._wheel)
            for child in widget.winfo_children():
                self._bind_scroll(child)

    def _reset(self):
        if self.progress is not None:
            self.progress.stop()
            self.progress = None
        for child in self.body.winfo_children():
            child.destroy()
        self.wrap_labels = []
        self.summary_box = None
        self.canvas.yview_moveto(0)

    def _label(self, parent, text, size=11, color=None, bg='white', bold=False,
               wrap=None, **pack):
        label = tk.Label(parent, text=text, bg=bg, fg=color or self.INK,
                         font=('Microsoft YaHei', size, 'bold' if bold else 'normal'),
                         anchor='w', justify='left', borderwidth=0)
        if wrap:
            label.configure(wraplength=max(240, self.canvas.winfo_width() - wrap))
            self.wrap_labels.append((label, wrap))
        label.pack(**pack)
        return label

    def _hero(self, status, title, message, question=None):
        dark, light, accent = self.THEMES[status]
        hero = tk.Frame(self.body, bg=dark, padx=22, pady=16)
        hero.pack(fill='x', padx=16, pady=(16, 12))
        emblem = tk.Canvas(hero, width=84, height=84, bg=dark, highlightthickness=0)
        emblem.pack(side='right', padx=(16, 4))
        emblem.create_oval(5, 5, 79, 79, outline=accent, width=3)
        if status == 'feasible':
            emblem.create_line(24, 43, 37, 56, 61, 30, fill=light, width=6,
                               capstyle='round', joinstyle='round')
        elif status == 'infeasible':
            emblem.create_line(28, 28, 56, 56, fill=light, width=5, capstyle='round')
            emblem.create_line(28, 56, 56, 28, fill=light, width=5, capstyle='round')
        else:
            emblem.create_line(42, 24, 42, 46, fill=light, width=5, capstyle='round')
            emblem.create_oval(39, 56, 45, 62, fill=light, outline=light)
        text = tk.Frame(hero, bg=dark)
        text.pack(side='left', fill='both', expand=True)
        self._label(text, f'问题 {question}  /  附件可行性检查' if question else '解答验证器',
                    size=10, color=light, bg=dark, anchor='w')
        self.title_label = self._label(text, title, size=30, color='white', bg=dark,
                                      bold=True, anchor='w', pady=(2, 2))
        self._label(text, message, size=11, color=light, bg=dark, wrap=205, fill='x')

    def show_waiting(self, busy=False):
        self._reset()
        self.view = None
        self._hero('pending' if busy else 'ready', '正在验证…' if busy else '准备检查你的方案',
                   '正在逐项核对调整规则，并检查全部重复使用区间。' if busy else
                   '选好问题和结果附件，点击“开始验证”。也可以先载入一份可行示例。')
        if busy:
            self.progress = ttk.Progressbar(self.body, mode='indeterminate')
            self.progress.pack(fill='x', padx=20, pady=10)
            self.progress.start(12)
        self._bind_scroll(self.body)

    def render(self, result, full_summary):
        self._reset()
        view = self.view = build_result_view(result)
        self._hero(view['status'], view['title'], view['message'], result.get('question'))
        metrics = tk.Frame(self.body, bg=self.BACKGROUND)
        metrics.pack(fill='x', padx=16)
        self.metric_labels = []
        self.metric_breakdown_labels = []
        for index, metric in enumerate(view['metrics']):
            metrics.columnconfigure(index, weight=1, uniform='metric')
            card = tk.Frame(metrics, bg='white', highlightbackground='#dfe7f0', highlightthickness=1,
                            padx=14, pady=12)
            card.grid(row=0, column=index, sticky='nsew', padx=(0 if index == 0 else 5, 0 if index == 3 else 5))
            self._label(card, metric['label'], size=11, color=self.MUTED, fill='x')
            color = self.THEMES[view['status']][0] if index == 0 else self.INK
            value = self._label(card, metric['value'], size=28, color=color, bold=True,
                                anchor='w', pady=(4, 2))
            self.metric_labels.append(value)
            breakdown = None
            if metric.get('breakdown'):
                breakdown = self._label(card, metric['breakdown'], size=9, color=self.INK,
                                        fill='x', pady=(0, 4))
            self.metric_breakdown_labels.append(breakdown)
            hint = self._label(card, metric['hint'], size=9, color=self.MUTED, fill='x')
            def resize_card(event, labels=tuple(label for label in (breakdown, hint) if label is not None)):
                for label in labels:
                    label.configure(wraplength=max(100, event.width - 30))
            card.bind('<Configure>', resize_card)
        notes = tk.Frame(self.body, bg=self.BACKGROUND)
        notes.pack(fill='x', padx=20, pady=(10, 12))
        for note in view['notes']:
            self._label(notes, note, size=10, color=self.MUTED, bg=self.BACKGROUND,
                        wrap=60, fill='x', pady=1)
        conflicts = result.get('conflicts', [])
        issues = [issue for issue in result.get('errors', [])
                  if not isinstance(issue, dict) or issue.get('code') != 'resource_conflict']
        self.conflict_buttons = []
        self.conflict_details = []
        if conflicts:
            pairs = {}
            for index, conflict in enumerate(conflicts):
                key = tuple(sorted((conflict['id1'], conflict['id2'])))
                if key not in pairs:
                    pairs[key] = [index, 0]
                pairs[key][1] += 1
            self._section_heading('这些装备发生了冲突',
                                  f'已发现 {len(conflicts)} 次交叠 · 下方展示前 {min(3, len(pairs))} 对装备的首处冲突')
            for index, pair_count in list(pairs.values())[:3]:
                conflict = conflicts[index]
                row = tk.Frame(self.body, bg='#fff1f2', padx=16, pady=11,
                               highlightbackground='#fecdd3', highlightthickness=1)
                row.pack(fill='x', padx=16, pady=(0, 7))
                button = ttk.Button(row, text='定位到矩阵 →',
                                    command=lambda i=index: self.locate_conflict(i))
                button.pack(side='right', padx=(12, 0))
                self.conflict_buttons.append(button)
                words = tk.Frame(row, bg='#fff1f2')
                words.pack(side='left', fill='both', expand=True)
                title = (f'{conflict["id1"]}（第{conflict["use1"]}次）  ↔  '
                         f'{conflict["id2"]}（第{conflict["use2"]}次）')
                detail = (f'重叠时间 [{conflict["t0"]}, {conflict["t1"]}) Δt    ·    '
                          f'重叠频段 [{conflict["f0"]}, {conflict["f1"]}) Δf')
                self.conflict_details.append(title + '\n' + detail)
                self._label(words, title, size=11, color='#9f1239', bg='#fff1f2', bold=True,
                            wrap=250, fill='x')
                self._label(words, detail, size=10, color='#9f1239', bg='#fff1f2',
                            wrap=250, fill='x', pady=(3, 0))
                self._label(words, f'这两台装备共有 {pair_count} 次交叠记录', size=9,
                            color='#9f1239', bg='#fff1f2', anchor='w', pady=(3, 0))
        if issues:
            self._section_heading('需要修改的地方', f'共 {len(issues)} 条格式或规则提示')
            for issue in issues[:3]:
                if isinstance(issue, dict):
                    parts = [str(issue[k]) for k in ('sheet', 'id') if issue.get(k)]
                    if issue.get('row') is not None:
                        parts.append(f'第{issue["row"]}行')
                    text = (' / '.join(parts) + '：' if parts else '') + str(issue.get('message', issue))
                else:
                    text = str(issue)
                self._label(self.body, text, size=10, color='#8d3e13', bg=self.BACKGROUND,
                            wrap=64, fill='x', padx=20, pady=(0, 6))
        if conflicts or issues:
            line = tk.Frame(self.body, bg=self.BACKGROUND)
            line.pack(fill='x', padx=16, pady=(2, 12))
            ttk.Button(line, text='查看全部问题与冲突 →', command=self.show_issues).pack(side='left')
        self._section_heading('完整统计与验证说明', '面积利用率、分类统计和九项指标。')
        self.summary_box = ScrolledText(self.body, wrap='word', height=12, padx=14, pady=12,
                                       bg='white', fg=self.INK, relief='flat', borderwidth=0,
                                       font=('Microsoft YaHei', 10))
        self.summary_box.pack(fill='both', expand=True, padx=16, pady=(0, 16))
        self.summary_box.insert('1.0', full_summary)
        self.summary_box.configure(state='disabled')
        self._bind_scroll(self.body)
        self.canvas.yview_moveto(0)
        return self.summary_box

    def _section_heading(self, title, caption):
        box = tk.Frame(self.body, bg=self.BACKGROUND)
        box.pack(fill='x', padx=20, pady=(3, 8))
        self._label(box, title, size=12, bold=True, bg=self.BACKGROUND, anchor='w')
        self._label(box, caption, size=9, color=self.MUTED, bg=self.BACKGROUND,
                    wrap=60, fill='x', pady=(2, 0))
