"""从已有求解结果生成教学示例，不重新运行优化，也不修改原始附件。"""
from pathlib import Path
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


APP = Path(__file__).resolve().parents[1]
PROJECT = APP.parent
Q2_HEADER = ["用频装备编号", "调整后频段区间", "调整后时间区间", "是否撤销用频计划"]
Q3_HEADER = ["新增用频装备序号", "调整后频段区间", "调整后时间区间"]


def save_table(path, headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "提交结果"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="205A78")
        cell.alignment = Alignment(horizontal="center")
    for name, width in zip("ABCD", [22, 26, 26, 24]):
        sheet.column_dimensions[name].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def interval(start, end):
    return f"[{start},{end})"


def build():
    output = APP / "examples"
    original = json.loads((PROJECT / "original_plans.json").read_text(encoding="utf-8"))
    original = {row["id"]: row for row in original}
    q2_path = PROJECT / "九层优先级_幅度优化/q2/九层重算/solution.json"
    q3_path = PROJECT / "九层优先级_幅度优化/q3/九层重算/q3_solution.json"
    q2 = json.loads(q2_path.read_text(encoding="utf-8"))
    q3 = json.loads(q3_path.read_text(encoding="utf-8"))
    q2_rows = []
    for row in q2:
        previous = original[row["id"]]
        if row.get("canceled"):
            q2_rows.append([row["id"], None, None, "是"])
        elif (row["f0"], row["f1"]) != (previous["f0"], previous["f1"]):
            q2_rows.append([row["id"], interval(row["f0"], row["f1"]), None, None])
        elif (row["t0"], row["t1"]) != (previous["t0"], previous["t1"]):
            q2_rows.append([row["id"], None, interval(row["t0"], row["t1"]), None])
    q3_rows = [[index, interval(row["f0"], row["f1"]), interval(row["t0"], row["t1"])]
               for index, row in enumerate(q3, 1)]
    save_table(output / "result2_可行示例.xlsx", Q2_HEADER, q2_rows)
    save_table(output / "result3_可行示例.xlsx", Q3_HEADER, q3_rows)
    save_table(output / "result2_重复编号示例.xlsx", Q2_HEADER, q2_rows + [q2_rows[0]])
    save_table(output / "result2_频移超限示例.xlsx", Q2_HEADER,
               [row for row in q2_rows if row[0] != "A002"] + [["A002", "[16,26)", None, None]])
    save_table(output / "result3_新增重叠示例.xlsx", Q3_HEADER,
               q3_rows + [[137, q3_rows[0][1], q3_rows[0][2]]])
    print(f"已生成 5 个示例：第二问 {len(q2_rows)} 行，第三问 {len(q3_rows)} 行。")


if __name__ == "__main__":
    build()
