"""Directory summaries and isolated per-question reports."""
from datetime import datetime
from pathlib import Path
import json
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

import reporting
from batch import QUESTIONS


BATCH_STATES = {'missing': '附件缺失', 'ambiguous': '候选附件或目录有歧义',
                'blocked': '第二问基准不可用', 'error': '输入错误或检查未完成'}
Q1_METRICS = {'expected_pair_count': '应报冲突对数', 'submitted_pair_count': '提交不同装备对数',
              'correct_pair_count': '正确冲突对数', 'missing_pair_count': '漏报冲突对数',
              'extra_pair_count': '多报装备对数', 'duplicate_pair_count': '重复行数'}


def _q1_metric(result, key):
    summary = result.get('summary', {})
    if summary.get('analysis_complete') or (key == 'expected_pair_count' and result.get('expected_pairs')):
        return summary.get(key)
    return None


def _status(result):
    if result.get('batch_state') in BATCH_STATES:
        return BATCH_STATES[result['batch_state']]
    if result.get('question') == 1 and result.get('status') in ('feasible', 'infeasible'):
        return '冲突清单准确' if result['status'] == 'feasible' else '冲突清单不准确'
    return reporting.STATUS.get(result.get('status'), '未知')


def _result(batch, question):
    # Accept both the live integer-key form and a saved JSON report.
    return batch['results'].get(question, batch['results'].get(str(question), {}))


def directory_summary_text(batch):
    lines = ['目录验证汇总', '请求目录：' + str(batch.get('requested_directory', batch['directory'])),
             '实际验证目录：' + str(batch['directory']),
             '目录选择：' + batch.get('search_note', '使用所选目录顶层附件。'),
             f'问题2—4时间窗口：[0,{batch["horizon"]})；频率窗口：[0,100)',
             '总体结果：' + ('四问均通过' if batch['passed'] else '存在未通过或未完成的项目'),
             '四问附件来自同一目录；第三问固定同目录第二问方案；第一、四问独立验证。',
             '问题1按全部原始周期核对冲突清单；问题2—4检查可行性，不证明最优性。']
    for question in QUESTIONS:
        result = _result(batch, question)
        lines += ['', f'问题{question}：{_status(result)}']
        filename = batch['files'].get(question, batch['files'].get(str(question)))
        if filename:
            lines.append('  附件：' + str(filename))
        for error in result.get('errors', []):
            lines.append('  ' + reporting.text_of(error))
        summary = result.get('summary', {})
        metrics = Q1_METRICS if question == 1 else {
            key: reporting.METRIC_LABELS[key]
            for key in ('executed_count', 'added_count', 'conflict_pairs', 'objective_tuple')}
        for key, label in metrics.items():
            if key in summary:
                value = _q1_metric(result, key) if question == 1 else summary[key]
                lines.append('  ' + label + '：' + (str(value) if value is not None else '—'))
        if question == 1 and 'expected_event_count' in summary:
            lines.append(f'  原始计划交叠事件数：{summary["expected_event_count"]}（作为清单真值，不表示清单验证失败）')
    for error in batch.get('export_errors', []):
        lines += ['', '报告导出说明：' + error]
    return '\n'.join(lines)


def _new_directory(root, submission_directory):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r'[^\w\-\u4e00-\u9fff]+', '_', Path(submission_directory).name)[:70] or '提交目录'
    prefix = f'{datetime.now():%Y%m%d_%H%M%S_%f}_目录验证_{safe}'
    index = 0
    while True:
        directory = root / (prefix if index == 0 else f'{prefix}_{index}')
        try:
            directory.mkdir()
            return directory
        except FileExistsError:
            index += 1


def export_directory_report(batch, output_root):
    """Create a unique report directory and keep each question's errors visible."""
    destination = _new_directory(output_root, batch['directory'])
    export_errors = []
    for question in QUESTIONS:
        result = _result(batch, question)
        child = destination / f'问题{question}'
        child.mkdir()
        if result.get('batch_state') in ('missing', 'ambiguous', 'blocked'):
            (child / '验证报告.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            (child / '验证摘要.txt').write_text(reporting.summary_text(result), encoding='utf-8-sig')
            continue
        try:
            reporting.export_report(result, child)
        except Exception as exc:
            message = f'问题{question}明细导出未完成：{type(exc).__name__}: {exc}'
            export_errors.append(message)
            (child / '导出说明.txt').write_text(message, encoding='utf-8-sig')
            (child / '验证报告.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            (child / '验证摘要.txt').write_text(reporting.summary_text(result), encoding='utf-8-sig')
    saved = dict(batch)
    if export_errors:
        saved['export_errors'] = export_errors
    (destination / '目录验证报告.json').write_text(json.dumps(saved, ensure_ascii=False, indent=2), encoding='utf-8')
    (destination / '目录验证摘要.txt').write_text(directory_summary_text(saved), encoding='utf-8-sig')
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '四问汇总'
    sheet.append(['问题', '附件', '判定', '目录状态', *Q1_METRICS.values(),
                  '执行装备数（问题2—4）', '新增C类', '剩余冲突对数（问题2—4）', '原有装备指标（问题2—4）', '说明'])
    for question in QUESTIONS:
        result = _result(batch, question)
        summary = result.get('summary', {})
        filename = batch['files'].get(question, batch['files'].get(str(question)))
        values = [question, filename, _status(result), BATCH_STATES.get(result.get('batch_state'), ''),
                  *(_q1_metric(result, key) if question == 1 else None for key in Q1_METRICS),
                  *(summary.get(key) if question != 1 else None
                    for key in ('executed_count', 'added_count', 'conflict_pairs', 'objective_tuple')),
                  '\n'.join(reporting.text_of(error) for error in result.get('errors', []))]
        sheet.append([reporting.safe_value(value) for value in values])
    notes = workbook.create_sheet('验证设置与导出说明')
    notes.append(['项目', '内容'])
    for name, value in [('请求目录', batch.get('requested_directory', batch['directory'])),
                        ('实际验证目录', batch['directory']), ('目录选择', batch.get('search_note', '')),
                        ('问题2—4时间上限', batch['horizon']), ('总体通过', batch['passed']),
                        ('判定范围', '问题1核对冲突清单；问题2—4验证可行性，不证明最优性'),
                        ('问题1周期范围', '全部原始周期；原始交叠事件不表示清单错误'),
                        ('第三问基准', '同目录唯一且通过验证的第二问附件')]:
        notes.append([name, reporting.safe_value(value)])
    for error in export_errors:
        notes.append(['导出说明', reporting.safe_value(error)])
    for page in workbook:
        page.freeze_panes = 'A2'
        page.auto_filter.ref = page.dimensions
        for cell in page[1]:
            cell.font = Font(name='Microsoft YaHei', bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='245F87')
        for column in page.columns:
            page.column_dimensions[column[0].column_letter].width = min(75, max(14, max(len(str(c.value or '')) for c in column)))
        for row in page:
            for cell in row:
                cell.alignment = Alignment(vertical='top', wrap_text=True)
    workbook.save(destination / '目录验证汇总.xlsx')
    return destination
