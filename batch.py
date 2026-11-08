"""Validate a single submission directory without changing any input files."""
from pathlib import Path
import re
import stat
import unicodedata

from engine import validate_submission


QUESTIONS = (1, 2, 3, 4)
EXTENSIONS = {'.xlsx', '.xlsm', '.csv'}
_NUMBERS = {'1': 1, '2': 2, '3': 3, '4': 4, '一': 1, '二': 2, '三': 3, '四': 4}
_ENGLISH = re.compile(r'^results?[\s_-]*([1234])(?=$|[\s_.()\-]|[\u4e00-\u9fff])', re.I)
_CHINESE = re.compile(r'^(?:问题[\s_-]*([1234一二三四])|第?([1234一二三四])[问题])(?=$|[^0-9一二三四])')


def _question(path):
    if path.name.startswith('~$') or path.suffix.lower() not in EXTENSIONS:
        return None
    stem = unicodedata.normalize('NFKC', path.stem).strip()
    match = _ENGLISH.match(stem) or _CHINESE.match(stem)
    if match:
        return _NUMBERS[next(group for group in match.groups() if group)]
    return None


def _invalid(question, state, message, horizon, submission=None, base=None):
    settings = dict(horizon=horizon, frequency_range=[0, 100], fixed_base=question == 3,
                    interval_convention='half-open',
                    scope='问题1检验冲突清单准确性；问题2—4检验可行性与指标，不证明最优性。')
    if submission is not None:
        settings['input_file'] = str(submission)
    if base is not None:
        settings['base_file'] = str(base)
    return dict(question=question, status='invalid', passed=False, batch_state=state,
                errors=[dict(code='batch_' + state, message=message)], warnings=[],
                plans=[], conflicts=[], summary={}, settings=settings)


def _is_link(info):
    # Windows junctions are reparse points, even when is_symlink() is false.
    return stat.S_ISLNK(info.st_mode) or bool(
        getattr(info, 'st_file_attributes', 0) & getattr(stat, 'FILE_ATTRIBUTE_REPARSE_POINT', 0x400))


def _scan_directory(folder):
    """Read one folder without following symlinks or Windows junctions."""
    info = folder.lstat()
    if _is_link(info):
        raise OSError(f'不跟随符号链接或目录联接：{folder}')
    if not stat.S_ISDIR(info.st_mode):
        raise NotADirectoryError(f'指定路径不是目录：{folder}')
    candidates = {question: [] for question in QUESTIONS}
    children, skipped = [], []
    for path in sorted(folder.iterdir(), key=lambda item: item.name.casefold()):
        info = path.lstat()
        if _is_link(info):
            skipped.append(str(path))
        elif stat.S_ISDIR(info.st_mode):
            children.append(path)
        elif stat.S_ISREG(info.st_mode):
            question = _question(path)
            if question is not None:
                candidates[question].append(path)
    return candidates, children, skipped


def _select_directory(folder):
    root_candidates, children, skipped = _scan_directory(folder)
    if all(root_candidates.values()):
        return folder, root_candidates, '根目录四类附件齐全，使用根目录；未搜索子目录。', [], skipped
    complete, partial = [], []
    pending = list(reversed(children))
    while pending:
        child = pending.pop()
        candidates, descendants, links = _scan_directory(child)
        skipped.extend(links)
        if all(candidates.values()):
            complete.append((child, candidates))
        elif any(candidates.values()):
            partial.append((child, candidates))
        pending.extend(reversed(descendants))
    if len(complete) == 1:
        selected, candidates = complete[0]
        return selected, candidates, f'根目录不齐全，找到唯一四类附件齐全的子目录：{selected}', [], skipped
    if len(complete) > 1:
        matches = [str(path) for path, _ in complete]
        return folder, None, '找到多个四类附件齐全的子目录，无法确定使用哪一套；请直接选择其中一个目录。', matches, skipped
    if any(root_candidates.values()):
        return folder, root_candidates, '未找到四类附件齐全的子目录；使用根目录已有附件并报告缺失，不跨目录拼接。', [], skipped
    if len(partial) == 1:
        selected, candidates = partial[0]
        return selected, candidates, f'根目录无可识别附件，找到唯一含部分附件的子目录：{selected}；其余附件按缺失报告。', [], skipped
    if len(partial) > 1:
        matches = [str(path) for path, _ in partial]
        return folder, None, '找到多个含部分附件的子目录，无法确定使用哪一套；请直接选择其中一个目录，不跨目录拼接。', matches, skipped
    return folder, root_candidates, '根目录及可读取的子目录均未找到可识别附件，四问均按缺失报告。', [], skipped


def validate_directory(directory, horizon=643, original_path=None):
    """Select one Q1–Q4 directory; Q3 uses only that directory's passing Q2.

    Results/files have integer keys 1–4. Complete root folders take priority;
    otherwise recursively choose a unique complete (or lone partial) subfolder.
    Never combine files from different directories or guess ambiguous matches.
    """
    folder = Path(directory).expanduser().absolute()
    batch = dict(directory=str(folder), requested_directory=str(folder), search_note='',
                 horizon=horizon, passed=False, results={},
                 files={question: None for question in QUESTIONS})
    try:
        selected, candidates, note, matches, skipped = _select_directory(folder)
        batch.update(directory=str(selected), search_note=note)
        if skipped:
            batch['skipped_links'] = skipped
            batch['search_note'] += f' 已跳过{len(skipped)}个链接或目录联接。'
        if matches:
            batch['candidate_directories'] = matches
            for question in QUESTIONS:
                result = _invalid(question, 'ambiguous', note + '\n候选目录：' + '；'.join(matches), horizon)
                result['candidate_directories'] = matches
                batch['results'][question] = result
            return batch
    except OSError as exc:
        message = f'目录搜索未完成，不能确定唯一提交目录：{type(exc).__name__}: {exc}'
        batch.update(search_note=message, search_errors=[message])
        for question in QUESTIONS:
            batch['results'][question] = _invalid(question, 'error', message, horizon)
        return batch

    for question in QUESTIONS:
        matches = candidates[question]
        if not matches:
            batch['results'][question] = _invalid(
                question, 'missing', f'所选目录顶层未找到问题{question}附件；请使用 results{question}.xlsx、'
                f'result{question}.xlsx（亦支持 .xlsm、.csv）或对应中文问题名称。', horizon)
            continue
        if len(matches) > 1:
            result = _invalid(question, 'ambiguous',
                              f'问题{question}找到多个候选附件，无法自动确定：' + '；'.join(p.name for p in matches), horizon)
            result['candidates'] = [str(path) for path in matches]
            batch['results'][question] = result
            continue
        submission = matches[0]
        batch['files'][question] = str(submission)
        base = batch['files'][2] if question == 3 else None
        if question == 3 and not batch['results'][2].get('passed'):
            result = _invalid(3, 'blocked', '同目录第二问基准缺失、存在歧义或未通过验证，'
                              '问题三暂不能验证；请先修正第二问附件。', horizon, submission, base)
            result['base_validation'] = {key: batch['results'][2].get(key)
                                         for key in ('status', 'passed', 'errors', 'summary', 'settings')}
            batch['results'][3] = result
            continue
        try:
            result = validate_submission(question, submission, base_q2=base,
                                         horizon=horizon, original_path=original_path)
        except Exception as exc:
            result = _invalid(question, 'error',
                              f'问题{question}检查未完成：{type(exc).__name__}: {exc}',
                              horizon, submission, base)
        if result.get('status') == 'invalid':
            result.setdefault('batch_state', 'error')
        batch['results'][question] = result
    batch['passed'] = all(batch['results'][q].get('passed', False) for q in QUESTIONS)
    return batch
