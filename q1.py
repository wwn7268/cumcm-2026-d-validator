"""Check the original-plan conflict-pair list, independently of optimization."""
from pathlib import Path

from openpyxl import load_workbook

import engine


HEADERS = ['序号', '冲突装备1', '冲突设备2']
_ALIASES = {
    'serial': {'序号', '冲突序号'},
    'id1': {'冲突装备1', '冲突设备1', '发生冲突的用频装备编号1', '用频装备编号1', '装备编号1', '设备编号1'},
    'id2': {'冲突装备2', '冲突设备2', '发生冲突的用频装备编号2', '用频装备编号2', '装备编号2', '设备编号2'},
    'note': {'备注', '说明', 'note'},
}
_HEADER_MAP = {engine._header_key(alias): key for key, aliases in _ALIASES.items() for alias in aliases}


def _header(rows):
    for index, row in enumerate(rows[:20]):
        mapping = {}
        for column, value in enumerate(row):
            if not engine._blank(value):
                key = _HEADER_MAP.get(engine._header_key(value))
                if key:
                    if key in mapping and key != 'note':
                        mapping['__duplicate__'] = column
                    mapping[key] = column
        if {'serial', 'id1', 'id2'} <= mapping.keys():
            return index, mapping
    return None


def _read_table(path, sheet):
    engine._check_file(path)
    if path.suffix.lower() == '.csv':
        if sheet not in (None, '', 'CSV'):
            raise engine.InputError('sheet_not_found', 'CSV 的表名为 CSV。')
        rows = engine._csv_rows(path)
        found = _header(rows)
        if not found:
            raise engine.InputError('missing_header', '前20行未找到问题1的完整附件表头。', sheet='CSV')
        return 'CSV', rows, *found
    try:
        workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception as exc:
        raise engine.InputError('unreadable_workbook', f'无法打开 Excel：{exc}') from exc
    try:
        if sheet and sheet not in workbook.sheetnames:
            raise engine.InputError('sheet_not_found', f'找不到工作表：{sheet}')
        matched = []
        for name in [sheet] if sheet else workbook.sheetnames:
            page = workbook[name]
            preview_columns = min(page.max_column or engine.MAX_COLUMNS + 1, engine.MAX_COLUMNS + 1)
            preview = [[cell.value for cell in row] for row in page.iter_rows(
                min_row=1, max_row=20, max_col=preview_columns)]
            found = _header(preview)
            if found:
                matched.append((name, found))
        if not matched:
            raise engine.InputError('missing_header', '未找到问题1的完整附件表头。', sheet=sheet)
        if len(matched) > 1:
            raise engine.InputError('ambiguous_sheet', '多个工作表符合第一问格式，请明确选择：' + '、'.join(name for name, _ in matched))
        name, found = matched[0]
        page = workbook[name]
        if (page.max_row or 0) > engine.MAX_ROWS + 20 or (page.max_column or 0) > engine.MAX_COLUMNS:
            raise engine.InputError('table_limit', f'附件超过{engine.MAX_ROWS}条数据或{engine.MAX_COLUMNS}列，未完成检查。', sheet=name)
        rows = []
        for rowno, cells in enumerate(page.iter_rows(), 1):
            if rowno > engine.MAX_ROWS + 20 or len(cells) > engine.MAX_COLUMNS:
                raise engine.InputError('table_limit', f'附件超过{engine.MAX_ROWS}条数据或{engine.MAX_COLUMNS}列，未完成检查。', sheet=name, row=rowno)
            for cell in cells:
                if cell.data_type == 'f':
                    raise engine.InputError('formula_cell', f'单元格{cell.coordinate}含公式，请先粘贴为数值。', sheet=name, row=rowno)
                if cell.data_type == 'e':
                    raise engine.InputError('excel_error_cell', f'单元格{cell.coordinate}含Excel错误值。', sheet=name, row=rowno)
            rows.append([cell.value for cell in cells])
        return name, rows, *found
    finally:
        workbook.close()


def conflict_truth(plans):
    """Expand all repetitions, then independently compare positive intersections."""
    times = {}
    for plan in plans:
        duration = plan['t1'] - plan['t0']
        times[plan['id']] = [(plan['t0'] + k * (duration + plan['gap']),
                              plan['t1'] + k * (duration + plan['gap']))
                             for k in range(plan['count'])]
    pairs, events = set(), []
    for i, left in enumerate(plans):
        for right in plans[i + 1:]:
            f0, f1 = max(left['f0'], right['f0']), min(left['f1'], right['f1'])
            if f0 >= f1:
                continue
            for use1, (a, b) in enumerate(times[left['id']], 1):
                for use2, (c, d) in enumerate(times[right['id']], 1):
                    t0, t1 = max(a, c), min(b, d)
                    if t0 < t1:
                        pairs.add(tuple(sorted((left['id'], right['id']))))
                        events.append(dict(id1=left['id'], id2=right['id'], use1=use1, use2=use2,
                                           f0=f0, f1=f1, t0=t0, t1=t1, kind='原有—原有'))
    return pairs, events


def _pair_records(pairs):
    return [dict(id1=left, id2=right) for left, right in sorted(pairs)]


def validate_q1(submission, *, horizon=643, sheet=None, original_path=None):
    """A passing result means an accurate pair list, not conflict-free original plans."""
    result = dict(question=1, status='invalid', passed=False, errors=[], warnings=[], plans=[],
                  conflicts=[], missing_pairs=[], extra_pairs=[], duplicate_pairs=[],
                  expected_pairs=[], submitted_pairs=[],
                  summary=dict(analysis_complete=False, expected_pair_count=0, submitted_pair_count=0,
                               submitted_row_count=0, correct_pair_count=0, missing_pair_count=0,
                               extra_pair_count=0, duplicate_pair_count=0),
                  settings=dict(horizon=horizon, requested_horizon=horizon, frequency_range=[0, 100],
                                interval_convention='half-open', fixed_base=False, input_file=str(Path(submission)),
                                original_file=str(original_path or engine.DEFAULT_ORIGINAL),
                                scope='核对原始计划的全部无向冲突装备对；通过表示清单准确，并不表示原计划无冲突。',
                                horizon_note='问题一不按输入时间窗裁剪冲突真值；绘图时间窗自动覆盖原始完整计划。'))
    try:
        if isinstance(horizon, bool) or not isinstance(horizon, int) or not 1 <= horizon <= 10000:
            raise engine.InputError('invalid_horizon', '时间窗上限必须是1..10000的整数。')
        original = engine._load_original(Path(original_path) if original_path else engine.DEFAULT_ORIGINAL)
        plans = list(original.values())
        for plan in plans:
            if not (0 <= plan['f0'] < plan['f1'] <= 100 and 0 <= plan['t0'] < plan['t1']
                    and plan['gap'] >= 0 and 1 <= plan['count'] <= 1000):
                raise engine.InputError('original_data_error', f"原始装备{plan['id']}的周期或边界参数不合法。")
        last = max(p['t1'] + (p['count'] - 1) * (p['t1'] - p['t0'] + p['gap']) for p in plans)
        if last > 10000 or sum(p['count'] for p in plans) > engine.MAX_RECTANGLES:
            raise engine.InputError('original_data_error', '原始计划超过完整展开的时间或矩形上限。')
        drawing_horizon = max(horizon, last)
        result['settings']['horizon'] = drawing_horizon
        if horizon < last:
            result['warnings'].append(f'输入时间窗{horizon}小于原始末次结束{last}，问题一仍检查全部周期，绘图时间窗已扩展至{last}。')
        expected, events = conflict_truth(plans)
        geometry = engine.analyze_plans(plans, drawing_horizon)
        if geometry['errors'] or not geometry['summary'].get('analysis_complete'):
            raise engine.InputError('original_data_error', '原始计划的几何信息无法完整分析。')
        result['plans'] = plans
        result['conflicts'] = events
        result['expected_pairs'] = _pair_records(expected)
        result['summary'].update(geometry['summary'])
        # These counters refer to original geometry, not mistakes in the submitted list.
        result['summary'].update(expected_pair_count=len(expected), expected_event_count=len(events),
                                 conflict_pairs=len(expected), conflict_events=len(events), analysis_complete=False)
        result['summary'].pop('objective_tuple', None)
        result['summary'].pop('objective_six', None)
        result['summary'].pop('metrics_valid_for_feasible_plan', None)
        name, rows, header_no, columns = _read_table(Path(submission), sheet)
        if '__duplicate__' in columns:
            raise engine.InputError('duplicate_header', '表头包含重复字段，请每个字段只保留一列。', sheet=name, row=header_no + 1)
        if len(rows) - header_no - 1 > engine.MAX_ROWS:
            raise engine.InputError('table_limit', f'数据行超过{engine.MAX_ROWS}，未完成检查。', sheet=name)
        incomplete = False
        for index, value in enumerate(rows[header_no]):
            if index not in columns.values() and not engine._blank(value):
                result['errors'].append(engine._error('unsupported_column', f'无法解释字段“{engine._text(value)}”。', sheet=name, row=header_no + 1))
                incomplete = True
        submitted, first_rows = set(), {}
        for rowno, row in enumerate(rows[header_no + 1:], header_no + 2):
            if all(engine._blank(value) for value in row):
                continue
            result['summary']['submitted_row_count'] += 1
            try:
                if any(isinstance(value, str) and value.lstrip().startswith('=') for value in row):
                    raise engine.InputError('formula_cell', '数据行含公式文本，请先粘贴为数值。')
                def cell(key):
                    index = columns[key]
                    return row[index] if index < len(row) else None
                if engine._integer(cell('serial'), '序号') < 1:
                    raise engine.InputError('invalid_serial', '序号必须为正整数；可任意调整行的排列顺序。')
                left = engine._normalize_id(cell('id1'), 2)
                right = engine._normalize_id(cell('id2'), 2)
                unknown = [ident for ident in (left, right) if ident not in original]
                if unknown:
                    raise engine.InputError('unknown_id', '原始附件不存在装备：' + '、'.join(unknown))
                if left == right:
                    raise engine.InputError('self_conflict', f'{left}不能与自身组成装备冲突对。')
                pair = tuple(sorted((left, right)))
                if pair in submitted:
                    result['duplicate_pairs'].append(dict(id1=pair[0], id2=pair[1], sheet=name,
                                                          row=rowno, first_row=first_rows[pair]))
                    result['errors'].append(engine._error('duplicate_pair',
                        f'{pair[0]}与{pair[1]}重复出现（含反向重复）；首次出现在第{first_rows[pair]}行。', sheet=name, row=rowno))
                else:
                    submitted.add(pair)
                    first_rows[pair] = rowno
            except engine.InputError as exc:
                result['errors'].append(engine._error(exc.code, str(exc), sheet=name, row=rowno))
                incomplete = True
        missing, extra = expected - submitted, submitted - expected
        result['submitted_pairs'] = _pair_records(submitted)
        result['missing_pairs'] = _pair_records(missing)
        result['extra_pairs'] = _pair_records(extra)
        if missing:
            result['errors'].append(engine._error('missing_pairs', f'漏报{len(missing)}对真实冲突装备，详见漏报清单。'))
        for left, right in sorted(extra):
            result['errors'].append(engine._error('extra_pair', f'{left}与{right}的全部周期不存在时频冲突，属于多报。',
                                                  sheet=name, row=first_rows[(left, right)]))
        result['summary'].update(analysis_complete=not incomplete, submitted_pair_count=len(submitted),
                                 correct_pair_count=len(submitted & expected), missing_pair_count=len(missing),
                                 extra_pair_count=len(extra), duplicate_pair_count=len(result['duplicate_pairs']))
        result['status'] = 'invalid' if incomplete else 'infeasible' if result['errors'] else 'feasible'
        result['passed'] = result['status'] == 'feasible'
        result['warnings'].append('原始计划的交叠事件用于定位真实冲突；不会把这些已正确报告的冲突作为第一问清单错误。')
    except engine.InputError as exc:
        result['errors'].append(engine._error(exc.code, str(exc), **exc.location))
    except Exception as exc:
        result['errors'].append(engine._error('read_error', f'问题一检查未完成：{type(exc).__name__}: {exc}'))
    return result
