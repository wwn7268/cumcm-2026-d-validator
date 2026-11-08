"""Create seven reusable directory examples and check their actual outcomes."""
from pathlib import Path
import hashlib
import json
import shutil
import sys

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
from batch import validate_directory
from batch_reporting import directory_summary_text


def example(q, name=None, source=None):
    return (source or f'result{q}_可行示例.xlsx', name or f'result{q}.xlsx')


CASES = [
    ('01_文件名不标准但可识别',
     [example(2,'第二问_最终稿.xlsx'), example(3,'RESULT_3_新增方案.xlsx'), example(4,'问题四-调整结果.xlsx')],
     ['feasible','feasible','feasible'],
     '三份内容都可行，仅改为中文、大小写和带后缀的文件名。应识别并全部通过。'),
    ('02_第二问文件名拼错',
     [example(2,'reslut2.xlsx'), example(3), example(4)],
     ['missing','blocked','feasible'],
     '第二问把 result 拼成 reslut。程序不猜测文件身份：问题二未找到，问题三基准不可用，问题四照常检查。'),
    ('03_第二问有两个版本',
     [example(2), example(2,'问题二_最终版.xlsx'), example(3), example(4)],
     ['ambiguous','blocked','feasible'],
     '两份第二问文件内容相同，但都有可识别的文件名。程序应提示文件不唯一，不能自行挑选一个版本。'),
    ('04_第二问频移超限',
     [example(2,source='result2_频移超限示例.xlsx'), example(3), example(4)],
     ['infeasible','blocked','feasible'],
     '第二问 A002 的频段由 [5,15) 改为 [16,26)，平移11格，超过10格。问题三因此不能使用该基准。'),
    ('05_第三问新增装备冲突',
     [example(2), example(3,source='result3_新增重叠示例.xlsx'), example(4)],
     ['feasible','infeasible','feasible'],
     '第二问正常。第三问故意新增一台与第1台位置完全相同的装备，检查能否找到全部重复区间的冲突。'),
    ('06_第四问间隔超限',
     [example(2), example(3), example(4,source='result4_间隔超限示例.xlsx')],
     ['feasible','feasible','infeasible'],
     '第四问 C001 的新间隔填成19，比原值8大11格，超过允许的10格；第二、三问应照常通过。'),
    ('07_文件名正确但内容放错',
     [example(2), example(3,source='result2_可行示例.xlsx'), example(4)],
     ['feasible','error','feasible'],
     '把第二问工作簿的内容放进 result3.xlsx。程序应发现表格内容不适用于第三问，不能因为文件名正确就判为通过。'),
]

LABELS = {'feasible':'通过','infeasible':'失败','missing':'未找到文件',
          'ambiguous':'文件不唯一','blocked':'第二问基准不可用','error':'格式错误／未完成'}


def main():
    output = APP/'examples/整目录测试示例'
    # Refuse to replace examples that the user may have edited while experimenting.
    output.mkdir(exist_ok=False)
    records = output/'实测记录'
    records.mkdir()
    source_files = {source for _,files,_,_ in CASES for source,_ in files}
    hashes = {name:hashlib.sha256((APP/'examples'/name).read_bytes()).hexdigest() for name in source_files}
    overview = []
    details = []
    for name,files,expected,description in CASES:
        folder = output/name
        folder.mkdir()
        for source,target in files:
            shutil.copy2(APP/'examples'/source, folder/target)
        batch = validate_directory(folder)
        actual = [batch['results'][q].get('batch_state',batch['results'][q]['status']) for q in (2,3,4)]
        record = dict(case=name,description=description,expected=expected,actual=actual,
                      matches_expectation=actual==expected,files=[target for _,target in files],batch=batch)
        (records/(name+'.json')).write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
        (records/(name+'.txt')).write_text(directory_summary_text(batch),encoding='utf-8-sig')
        overview.append(record)
        details += ['### '+name,'',description,'',
                    '实际结果：'+'；'.join(f'问题{q}：{LABELS[actual[q-2]]}' for q in (2,3,4))+'。','']
        for q in (2,3,4):
            result = batch['results'][q]
            if result.get('passed'):
                continue
            details += [f'**问题{q}的提示：**','']
            details += ['- '+error.get('message',str(error)) for error in result.get('errors',[])[:4]]
            conflicts = result.get('conflicts',[])
            if conflicts:
                c=conflicts[0]
                details += [f'- 第一条冲突：{c["id1"]}（第{c["use1"]}次）与 {c["id2"]}（第{c["use2"]}次），'
                            f'时间 [{c["t0"]},{c["t1"]})，频段 [{c["f0"]},{c["f1"]})。',
                            f'- 共记录 {len(conflicts)} 次交叠事件。']
            details += ['']
        print(name,actual,'OK' if actual==expected else 'MISMATCH',flush=True)
    for name,digest in hashes.items():
        assert hashlib.sha256((APP/'examples'/name).read_bytes()).hexdigest()==digest
    summary=['# 整目录验证示例：实际运行结果','',
             '以下结果使用当前验证器实际运行得到，时间窗 H=643。原始可行示例和错误示例均未改动。','',
             '| 示例目录 | 问题2 | 问题3 | 问题4 | 符合预期 |',
             '|---|---|---|---|---|']
    for r in overview:
        summary.append('| '+' | '.join([r['case'],*[LABELS[x] for x in r['actual']],
                                          '是' if r['matches_expectation'] else '否'])+' |')
    summary += ['', '“未找到文件”“文件不唯一”“基准不可用”表示尚未完成可行性判断，不等于已证明对应方案不可行。'
                '“失败”表示已经发现规则违反或冲突；文件名变体能识别时，仍须完成内容检查。','',*details]
    (output/'测试结果.md').write_text('\n'.join(summary),encoding='utf-8')
    (output/'使用方法.md').write_text('''# 七套整目录验证示例

1. 打开验证器，选择“整目录验证”。
2. 点击“选择目录”，选中一个以 01～07 开头的子文件夹。
3. 时间窗保持 H=643，点击“开始验证”。
4. 对照“测试结果.md”；点击某问的详情按钮查看具体错误或矩阵。

**不要直接选外层“整目录测试示例”文件夹**，验证器不会递归查找子目录。

01 的三问应全部通过。02、03、07用于演示文件识别或内容错误；04、05、06用于演示不可行方案。
每套目录里都放齐了三个问题的文件；03额外放了一份第二问，02故意把第二问的文件名拼错。

这些示例使用程序自带方案：可行的第二问撤销6台、调整126台；配套第三问新增136台。
它们与四层模型 I 报告中新增129台的那套布局不同，请勿混用第三问基准。

可以另复制一套目录再尝试修正：02把 reslut2.xlsx 改为 result2.xlsx；03只保留一份第二问附件；
其他错误的预期与实际提示见“测试结果.md”。“实测记录”保存逐套的 JSON 和文字结果。
''',encoding='utf-8')
    assert all(r['matches_expectation'] for r in overview), 'Inspect mismatched cases before delivering.'
    print('All 7 directories / 21 question outcomes matched the expectations.')


if __name__=='__main__':
    main()
