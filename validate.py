"""Command-line interface using the same rules as the desktop app."""
from pathlib import Path
import argparse
from engine import validate_submission
from reporting import export_report, new_output_dir, summary_text


def main():
    parser = argparse.ArgumentParser(description='D题第二、三、四问学生附件可行性验证')
    parser.add_argument('question', type=int, choices=(2, 3, 4), help='问题编号')
    parser.add_argument('submission', type=Path, help='学生Excel或CSV附件')
    parser.add_argument('--base-q2', type=Path, help='第三问对应的第二问附件')
    parser.add_argument('--horizon', type=int, default=643, help='时间窗口上限，默认643')
    parser.add_argument('--sheet', help='学生附件工作表名，默认自动识别')
    parser.add_argument('--base-sheet', help='第二问附件工作表名')
    parser.add_argument('--output', type=Path, default=Path(__file__).resolve().parent / 'results', help='报告输出根目录')
    args = parser.parse_args()
    result = validate_submission(args.question, args.submission, args.base_q2,
                                 horizon=args.horizon, sheet=args.sheet, base_sheet=args.base_sheet)
    print(summary_text(result))
    destination = new_output_dir(args.output, args.submission, args.question)
    export_report(result, destination)
    print(f'\n完整报告：{destination}')
    return {'feasible': 0, 'infeasible': 1, 'invalid': 2}.get(result.get('status'), 2)


if __name__ == '__main__':
    raise SystemExit(main())
