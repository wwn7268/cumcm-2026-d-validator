"""保留旧版 PDF 对应的第四问附件，并制作一个明确违反间隔限值的示例。"""
from pathlib import Path
from shutil import copyfile

from openpyxl import load_workbook

APP = Path(__file__).resolve().parents[1]


def build():
    source = APP.parent / "四问完整稿_PDF与LaTeX/对应结果附件/result4.xlsx"
    output = APP / "examples"
    output.mkdir(exist_ok=True)
    valid = output / "result4_可行示例.xlsx"
    copyfile(source, valid)
    workbook = load_workbook(valid)
    sheet = workbook.active
    for row in sheet.iter_rows(min_row=2):
        if row[0].value == "C001":
            row[3].value = 19  # 原 gap=8，变化 11，超过问题四允许的 10。
            break
    else:
        raise ValueError("未找到示例中的 C001，不能生成间隔超限示例。")
    workbook.save(output / "result4_间隔超限示例.xlsx")
    print("已复制第四问可行附件，并生成间隔超限示例。")


if __name__ == "__main__":
    build()
