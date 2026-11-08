"""Four-question overview for checking one submission directory."""
from pathlib import Path
import tkinter as tk
from tkinter import ttk

from result_view import build_result_view


class BatchResultPanel(ttk.Frame):
    COLORS = {
        'feasible': ('#065f46', '#ecfdf5'),
        'infeasible': ('#9f1239', '#fff1f2'),
        'invalid': ('#92400e', '#fffbeb'),
        'pending': ('#1e40af', '#eff6ff'),
    }

    def __init__(self, parent, open_question):
        super().__init__(parent, padding=12)
        self.open_question = open_question
        self.banner = tk.Label(self, anchor='w', padx=16, pady=12,
                               font=('Microsoft YaHei', 23, 'bold'))
        self.banner.pack(fill='x')
        self.note = ttk.Label(self, wraplength=1050)
        self.note.pack(fill='x', pady=(7, 12))
        self.cards = ttk.Frame(self)
        self.cards.pack(fill='both', expand=True)
        self.cards.rowconfigure(0, weight=1)
        self.status_labels = {}
        self.detail_buttons = {}
        self.labels = {}
        for column, q in enumerate((1, 2, 3, 4)):
            self.cards.columnconfigure(column, weight=1, uniform='questions')
            card = tk.Frame(self.cards, padx=14, pady=12, bg='white',
                            highlightbackground='#dbe3ef', highlightthickness=1)
            card.grid(row=0, column=column, sticky='nsew', padx=(0 if column==0 else 8, 0))
            tk.Label(card, text=f'问题 {q}', bg='white', fg='#334155',
                     anchor='w', font=('Microsoft YaHei', 13, 'bold')).pack(fill='x')
            status = tk.Label(card, bg='white', anchor='w', justify='left',
                              font=('Microsoft YaHei', 20, 'bold'))
            status.pack(fill='x', pady=(6, 9))
            detail = tk.Label(card, bg='white', fg='#334155', anchor='nw',
                              justify='left', font=('Microsoft YaHei', 10))
            detail.pack(fill='both', expand=True)
            filename = tk.Label(card, bg='white', fg='#64748b', anchor='w',
                                justify='left', font=('Microsoft YaHei', 9))
            filename.pack(fill='x', pady=(8, 6))
            button = ttk.Button(card, text=f'查看问题 {q} 详情',
                                command=lambda question=q: self.open_question(question))
            button.pack(anchor='w')
            card.bind('<Configure>', lambda event, labels=(detail, filename):
                      [label.configure(wraplength=max(120, event.width-30)) for label in labels])
            self.status_labels[q] = status
            self.detail_buttons[q] = button
            self.labels[q] = (detail, filename)
        self.bind('<Configure>', lambda event: self.note.configure(wraplength=max(250,event.width-28)))
        self.show_waiting()

    def show_waiting(self, busy=False):
        ink, background = self.COLORS['pending']
        self.banner.configure(text='正在验证四份附件…' if busy else '一次检查问题 1、2、3、4',
                              fg=ink, bg=background)
        self.note.configure(text='正在查找同一套四问结果；第三问使用同一目录的第二问方案。' if busy else
                            '选择放有 results1～4 的文件夹或其上级目录，再点击“开始验证”。')
        for q in (1,2,3,4):
            self.status_labels[q].configure(text='等待结果' if busy else '尚未验证',fg=ink)
            detail,filename = self.labels[q]
            detail.configure(text=('核对全部原始冲突对。' if q == 1 else '检查规则、边界与全部重复使用区间。') if busy else '')
            filename.configure(text='')
            self.detail_buttons[q].configure(state='disabled')

    def render(self, batch):
        results = batch['results']
        passed = sum(bool(results[q].get('passed')) for q in (1,2,3,4))
        failed = sum(results[q].get('status')=='infeasible' for q in (1,2,3,4))
        state = 'feasible' if passed==4 else ('infeasible' if failed else 'invalid')
        title = '全部验证通过' if passed==4 else ('存在未通过的题目' if failed else '部分题目未完成验证')
        ink,background=self.COLORS[state]
        self.banner.configure(text=f'{title}  ·  {passed} / 4 通过',fg=ink,bg=background)
        search_note = batch.get('search_note', '')
        self.note.configure(text=f"结果目录：{batch['directory']}\n{search_note}\n第一问核对清单准确性；第二至四问检查可行性，不证明最优。")
        special = {'missing':'缺少结果', 'ambiguous':'结果不唯一',
                   'blocked':'基准不可用', 'error':'读取未完成'}
        for q in (1,2,3,4):
            result=results[q]
            view=build_result_view(result)
            status=result.get('status','invalid')
            title=special.get(result.get('batch_state'),view['title'])
            self.status_labels[q].configure(text=title,fg=self.COLORS.get(status,self.COLORS['invalid'])[0])
            if result.get('passed'):
                if q == 1:
                    s = result.get('summary', {})
                    details = [f"应有冲突对 {s.get('expected_pair_count')} 对",
                               f"正确报告 {s.get('correct_pair_count')} 对",
                               f"漏报 {s.get('missing_pair_count')} · 多报 {s.get('extra_pair_count')}",
                               f"重复记录 {s.get('duplicate_pair_count')} 行",
                               '冲突清单完整、准确。']
                else:
                    details=[f"撤销 {view['cancelled']} 台　调整 {view['adjusted']} 台",
                         f"总平移距离 {view['translation_distance']} 步"]
                    for key, label in [('cancelled', '撤销'), ('adjusted', '调整'), ('translation_distance', '平移')]:
                        details.append(label + '：' + ' / '.join(f"{cls} {view['by_category'][cls][key]}" for cls in 'ABC'))
                if q==3:
                    details.insert(0,f"新增 C 类 {view['added_count']} 台")
                    details.append('撤销与调整统计来自第二问基准。')
                elif q==4:
                    details.append(f"间隔改变量 {view['gap_distance']} 步（另计）")
                if q != 1:
                    details.append('未发现规则、边界或时频冲突问题。')
            else:
                errors=result.get('errors',[])
                details=[str(e.get('message','') if isinstance(e,dict) else e) for e in errors[:2]]
                conflicts=result.get('conflicts',[])
                if conflicts and q != 1:
                    c=conflicts[0]
                    details.append(f"{c.get('id1')} 与 {c.get('id2')} 冲突\n时间 [{c.get('t0')}, {c.get('t1')})\n频段 [{c.get('f0')}, {c.get('f1')})")
                if not details:
                    details=[view['message']]
                if q == 1:
                    for key, label in [('missing_pairs', '漏报'), ('extra_pairs', '多报')]:
                        if result.get(key):
                            pair = result[key][0]
                            details.append(f"{label}：{pair['id1']} ↔ {pair['id2']}")
            detail,filename=self.labels[q]
            text='\n'.join(details)
            detail.configure(text=text if len(text)<=220 else text[:217]+'…')
            path=batch.get('files',{}).get(q)
            name=Path(path).name if path else '未识别到唯一文件'
            filename.configure(text=name if len(name)<=65 else name[:62]+'…')
            self.detail_buttons[q].configure(state='normal')
