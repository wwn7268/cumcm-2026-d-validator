"""Command-line interface using the same rules as the desktop app."""
from pathlib import Path
import argparse
from engine import validate_submission
from reporting import export_report, new_output_dir, summary_text
from batch import validate_directory
from batch_reporting import export_directory_report, directory_summary_text


def main(argv=None):
    parser = argparse.ArgumentParser(description='基于2026年高教社杯数学建模竞赛D题，核对第一问冲突清单并验证第二至四问附件可行性')
    parser.add_argument('question', type=int, choices=(1, 2, 3, 4), nargs='?', help='问题编号')
    parser.add_argument('submission', type=Path, nargs='?', help='提交的Excel或CSV附件')
    parser.add_argument('--directory', type=Path, help='验证同目录四问附件；根目录不齐全时搜索唯一完整子目录，不跨目录拼接')
    parser.add_argument('--base-q2', type=Path, help='第三问对应的第二问附件')
    parser.add_argument('--horizon', type=int, default=643, help='时间窗口上限，默认643')
    parser.add_argument('--sheet', help='提交附件工作表名，默认自动识别')
    parser.add_argument('--base-sheet', help='第二问附件工作表名')
    parser.add_argument('--output', type=Path, default=Path(__file__).resolve().parent / 'results', help='报告输出根目录')
    args = parser.parse_args(argv)
    if args.directory is not None:
        if any(value is not None for value in (args.question, args.submission, args.base_q2, args.sheet, args.base_sheet)):
            parser.error('--directory 不能与单问编号、附件、基准附件或工作表选项混用。')
        result = validate_directory(args.directory, horizon=args.horizon)
        print(directory_summary_text(result))
        destination = export_directory_report(result, args.output)
        print(f'\n完整报告：{destination}')
        if result['passed']:
            return 0
        return 2 if any(item.get('status') == 'invalid' for item in result['results'].values()) else 1
    if args.question is None or args.submission is None:
        parser.error('请提供问题编号和提交附件，或使用 --directory <目录>。')
    result = validate_submission(args.question, args.submission, args.base_q2,
                                 horizon=args.horizon, sheet=args.sheet, base_sheet=args.base_sheet)
    print(summary_text(result))
    destination = new_output_dir(args.output, args.submission, args.question)
    export_report(result, destination)
    print(f'\n完整报告：{destination}')
    return {'feasible': 0, 'infeasible': 1, 'invalid': 2}.get(result.get('status'), 2)


if __name__ == '__main__':
    raise SystemExit(main())
