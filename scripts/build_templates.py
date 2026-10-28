# -*- coding: utf-8 -*-
"""生成考勤导入模板与示例考勤数据（供仓库 templates/ 使用）"""
import datetime as dt
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BRAND = "10B981"
DEEP = "064E3B"
MIST = "ECFDF5"

HEADERS = ["Person ID", "Name", "Date", "Check-In", "Check-out"]


def style_sheet(ws, widths, freeze="A2"):
    header_fill = PatternFill("solid", fgColor=MIST)
    header_font = Font(bold=True, color=DEEP, size=11)
    thin = Side(style="thin", color="D8F0E3")
    border = Border(bottom=thin)
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = freeze


def build_template(out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "考勤明细"
    ws.append(HEADERS)
    demo = [
        ["'1001", "张三", "2026-09-01", "08:00:00", "17:05:00"],
        ["'1001", "张三", "2026-09-02", "08:12:00", "17:00:00"],
        ["'1001", "张三", "2026-09-03", "-", "-"],
        ["'1002", "李四", "2026-09-01", "07:56:00", "16:40:00"],
    ]
    for row in demo:
        ws.append(row)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.number_format = "@"  # 工号列强制文本
            cell.alignment = Alignment(horizontal="left")
    style_sheet(ws, [12, 12, 14, 13, 13])

    # 字段说明
    gi = wb.create_sheet("字段说明")
    gi.append(["列名", "中文名", "是否必填", "填写说明"])
    rows = [
        ["Person ID", "工号", "必填", "员工唯一编号。纯数字工号建议在单元格内以英文单引号开头（如 '1001），避免被 Excel 变成科学计数法"],
        ["Name", "姓名", "必填", "员工姓名；同一工号在不同行应保持一致"],
        ["Date", "日期", "必填", "考勤日期，推荐 YYYY-MM-DD，如 2026-09-01；也兼容 2026/9/1"],
        ["Check-In", "上班打卡", "选填", "当天首次打卡时间，格式 HH:mm 或 HH:mm:ss；未打卡填 - （计为缺卡）"],
        ["Check-out", "下班打卡", "选填", "当天末次打卡时间，格式 HH:mm 或 HH:mm:ss；未打卡填 - （计为缺卡）"],
        ["—", "行结构", "—", "一行 = 一名员工一天。同一员工的多天数据写多行，工具会按工号自动合并汇总"],
        ["—", "表头要求", "—", "表头放在前 20 行内即可，程序会自动定位（兼容考勤机导出常见的标题行/说明行）"],
        ["—", "自定义列名", "—", "兼容中英文表头：工号/姓名/日期/上班打卡/下班打卡，Person ID/Name/Date/Check-In/Check-out，只要命中关键字即可"],
        ["—", "周末规则", "—", "默认周六周日整行跳过（双休或加班），不判迟到早退；需要时可勾选「周末打卡也参与核算」"],
    ]
    for r in rows:
        gi.append(r)
    style_sheet(gi, [14, 11, 10, 74])

    # 用示例说明核算效果
    ex = wb.create_sheet("示例核算结果")
    ex.append(["工号", "姓名", "有效天数", "异常缺卡", "迟到(分)", "早退(分)", "扣款(元)", "说明"])
    ex.append(["'1001", "张三", 3, 1, 12, 0, 12.0,
               "09-01 正常；09-02 迟到 12 分钟；09-03 上下班均未打卡 → 记 1 次缺卡"])
    ex.append(["'1002", "李四", 1, 0, 0, 20, 20.0,
               "07:56 到岗不奖励；16:40 下班，早退 20 分钟 × 1 元 = 20 元"])
    style_sheet(ex, [10, 10, 11, 11, 11, 11, 11, 62])

    wb.save(out_path)


def build_sample(out_path, seed=20260901):
    random.seed(seed)
    staff = [
        ("1001", "张三"), ("1002", "李四"), ("1003", "王五"),
        ("1004", "赵六"), ("1005", "钱七"), ("1006", "孙八"),
    ]
    # 每人设定：迟到概率、早退概率、缺卡概率
    profile = {
        "1001": (0.10, 0.05, 0.00),
        "1002": (0.35, 0.20, 0.05),
        "1003": (0.00, 0.00, 0.00),
        "1004": (0.20, 0.35, 0.10),
        "1005": (0.05, 0.10, 0.15),
        "1006": (0.25, 0.05, 0.03),
    }

    wb = Workbook()
    ws = wb.active
    ws.title = "考勤明细"
    # 模拟考勤机导出的说明行，用于验证工具能自动跳过
    ws.append(["Check-In/Check-out Report 2026-09"])
    ws.append(HEADERS)

    start = dt.date(2026, 9, 1)
    for offset in range(30):
        day = start + dt.timedelta(days=offset)
        weekend = day.weekday() >= 5
        for pid, name in staff:
            late_p, early_p, miss_p = profile[pid]
            roll = random.random()

            if weekend:
                # 周末：绝大多数人无记录（应被自动跳过），1003 偶尔加班打卡
                if pid == "1003" and day.day in (5, 12):
                    ws.append([f"'{pid}", name, day.isoformat(), "09:00:00", "18:00:00"])
                else:
                    ws.append([f"'{pid}", name, day.isoformat(), "-", "-"])
                continue

            if roll < miss_p:
                # 整日未打卡
                ws.append([f"'{pid}", name, day.isoformat(), "-", "-"])
                continue

            # 上班：基准 08:00
            if random.random() < late_p:
                mins = random.choice([3, 5, 8, 12, 17, 25, 41])
                h, m = divmod(8 * 60 + mins, 60)
                cin = f"{h:02d}:{m:02d}:{random.randint(0, 59):02d}"
            else:
                mins_before = random.randint(0, 15)
                total = 8 * 60 - mins_before
                h, m = divmod(total, 60)
                cin = f"{h:02d}:{m:02d}:{random.randint(0, 59):02d}"

            # 下班：基准 17:00；偶发只打上班卡
            if random.random() < 0.04:
                cout = "-"
            elif random.random() < early_p:
                mins = random.choice([5, 10, 18, 30, 55])
                total = 17 * 60 - mins
                h, m = divmod(total, 60)
                cout = f"{h:02d}:{m:02d}:{random.randint(0, 59):02d}"
            else:
                extra = random.choice([0, 1, 3, 8, 20, 45])
                total = 17 * 60 + extra
                h, m = divmod(total, 60)
                cout = f"{h:02d}:{m:02d}:{random.randint(0, 59):02d}"

            ws.append([f"'{pid}", name, day.isoformat(), cin, cout])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.number_format = "@"
    style_sheet(ws, [12, 12, 14, 13, 13], freeze="A2")

    # 第二个工作表当作干扰项，验证工具只读第一个 sheet
    note = wb.create_sheet("使用说明")
    note["A1"] = "本文件为示例数据，直接拖入 index.html 即可看到核算效果。"
    note["A1"].font = Font(bold=True, color=DEEP, size=12)
    note["A3"] = "包含：正常出勤、迟到、早退、整日缺卡、只打上班卡、周末空记录、周末加班"
    note["A4"] = "第二、三个工作表不会被读取（工具只解析第一个工作表）"
    note.column_dimensions["A"].width = 80

    wb.save(out_path)


if __name__ == "__main__":
    build_template("templates/考勤导入模板.xlsx")
    build_sample("templates/示例考勤数据.xlsx")
    print("done")
