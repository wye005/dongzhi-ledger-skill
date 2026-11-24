# -*- coding: utf-8 -*-
"""在已派生的 F0 上，为 9 月新增客户插入客户行（洁宇再生资源 / 宝鑫再生资源）。
- 子表 3 张：在块内「总计/合计」行前插入，填 B=客户名、A=类别(沿用上一行)、
  30日列(AG=3+30)、本月累计列（签约表 AK37 / 产发表 AJ36 / 回款开票表 AJ36）
- 分项汇总 5 块：在块内「其他*统计」或「合计」行前插入，填累计列
插行后 Excel 自动扩展 SUM 范围，故合计会自动包含新行。
"""
import sys

import win32com.client as win32

F0 = sys.argv[1]

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = None

DAY = 30
DAYCOL = 3 + DAY          # 30日 → AG(33)
CUM_COL = {"签约": 37, "产出": 36, "回款开票": 36}   # AK / AJ / AJ

try:
    wb = xl.Workbooks.Open(F0)

    def find_block(ws, title_kw, start=1):
        for r in range(start, ws.UsedRange.Rows.Count + 1):
            v = ws.Cells(r, 1).Value
            if isinstance(v, str) and title_kw in v.replace(" ", ""):
                return r
        return None

    def insert_client(ws, row, name, day_val, cum_val, cum_col):
        ws.Rows(row).Insert(-4121, 0)          # 整行插入，格式沿用
        ws.Cells(row, 1).Value = ws.Cells(row - 1, 1).Value
        ws.Cells(row, 2).Value = name
        if day_val:
            ws.Cells(row, DAYCOL).Value = day_val
        ws.Cells(row, cum_col).Value = cum_val
        return row

    # ---------- 1) 订单签约跟踪表 ----------
    ws = wb.Worksheets("订单签约跟踪表")
    b1 = find_block(ws, "1、26年9月签约金额")
    b2 = find_block(ws, "2、26年9月签约数量")
    # 金额块：在「总计」行前插（block 内最后一个 总计）
    for r in range(b2 - 1, b1, -1):
        if str(ws.Cells(r, 1).Value or "").strip() == "总计":
            insert_client(ws, r, "洁宇再生资源", 118.0, 118.0, CUM_COL["签约"])
            print("签约金额块插入 @", r)
            break
    # 数量块：数量 0（SMT/物料待确认），也要有行
    b3 = find_block(ws, "3、2026年1-12月度签约金额合计")
    for r in range(b3 - 1, b2 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() == "总计":
            insert_client(ws, r, "洁宇再生资源", None, 0.0, CUM_COL["签约"])
            print("签约数量块插入 @", r)
            break

    # ---------- 2) 产出发货跟踪表 ----------
    ws = wb.Worksheets("产出发货跟踪表")
    t1 = find_block(ws, "1、26年9月产出数量")
    t2 = find_block(ws, "2、26年9月发货数量")
    t3 = find_block(ws, "3、26年9月发货金额")
    t4 = find_block(ws, "4、2026年1-12月产出数量月度合计")
    # 发货数量块（金额块前）：数量 0
    for r in range(t3 - 1, t2 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() in ("总计", "合计"):
            insert_client(ws, r, "洁宇再生资源", None, 0.0, CUM_COL["产出"])
            print("发货数量块插入 @", r)
            break
    # 发货金额块：金额 118
    for r in range(t4 - 1, t3 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() in ("总计", "合计"):
            insert_client(ws, r, "洁宇再生资源", 118.0, 118.0, CUM_COL["产出"])
            print("发货金额块插入 @", r)
            break

    # ---------- 3) 回款&开票统计表 ----------
    ws = wb.Worksheets("回款&开票统计表")
    h1 = find_block(ws, "1、2026年9月回款计划跟踪表")
    h2 = find_block(ws, "2、2026年9月开票计划跟踪表")
    h3 = find_block(ws, "3、2026年1-12月度回款统计")
    # 回款块：洁宇 113
    for r in range(h2 - 1, h1 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() in ("总计", "合计"):
            insert_client(ws, r, "洁宇再生资源", 113.0, 113.0, CUM_COL["回款开票"])
            print("回款块插入 @", r)
            break
    # 开票块：宝鑫 2.823
    for r in range(h3 - 1, h2 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() in ("总计", "合计"):
            insert_client(ws, r, "宝鑫再生资源", 2.823, 2.823, CUM_COL["回款开票"])
            print("开票块插入 @", r)
            break

    # ---------- 4) 分项汇总 ----------
    ws = wb.Worksheets("分项汇总")
    # 块1 订单：B=计划数量 C=计划金额 D=累计数量 E=累计金额
    b1 = find_block(ws, "1、9月订单计划")
    b2 = find_block(ws, "2、9月产出计划")
    for r in range(b2 - 1, b1 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() == "合计":
            ws.Rows(r).Insert(-4121, 0)
            ws.Cells(r, 1).Value = "洁宇再生资源"
            ws.Cells(r, 2).Value = 0
            ws.Cells(r, 3).Value = 0
            ws.Cells(r, 4).Value = 0
            ws.Cells(r, 5).Value = 118.0
            print("分项汇总 订单块插入 @", r)
            break
    # 块3 出货：B=计划数量 C=计划金额 D=累计数量 E=累计金额
    b3 = find_block(ws, "3、9月出货计划")
    b4 = find_block(ws, "4、9月开票计划")
    for r in range(b4 - 1, b3 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() == "合计":
            ws.Rows(r).Insert(-4121, 0)
            ws.Cells(r, 1).Value = "洁宇再生资源"
            ws.Cells(r, 2).Value = 0
            ws.Cells(r, 3).Value = 0
            ws.Cells(r, 4).Value = 0
            ws.Cells(r, 5).Value = 118.0
            print("分项汇总 出货块插入 @", r)
            break
    # 块4 开票：B=发货待开票 C=开票计划 D=累计开票 E=完成率
    b5 = find_block(ws, "5、9月回款计划")
    for r in range(b5 - 1, b4 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() == "合计":
            ws.Rows(r).Insert(-4121, 0)
            ws.Cells(r, 1).Value = "宝鑫再生资源"
            ws.Cells(r, 2).Value = 0
            ws.Cells(r, 3).Value = 0
            ws.Cells(r, 4).Value = 2.823
            ws.Cells(r, 5).Value = "-"
            print("分项汇总 开票块插入 @", r)
            break
    # 块5 回款：B/C=?  D=回款计划 E=累计回款 F=完成率 G=待回款
    last = ws.UsedRange.Rows.Count
    for r in range(last, b5 + 3, -1):
        if str(ws.Cells(r, 1).Value or "").strip() == "合计":
            ws.Rows(r).Insert(-4121, 0)
            ws.Cells(r, 1).Value = "洁宇再生资源"
            ws.Cells(r, 2).Value = 0
            ws.Cells(r, 3).Value = 0
            ws.Cells(r, 4).Value = 0
            ws.Cells(r, 5).Value = 113.0
            ws.Cells(r, 6).Value = "-"
            ws.Cells(r, 7).Value = 0
            print("分项汇总 回款块插入 @", r)
            break

    xl.CalculateFullRebuild()
    wb.Save()
    print("已保存", F0)
finally:
    if wb:
        wb.Close(SaveChanges=False)
    xl.Quit()
