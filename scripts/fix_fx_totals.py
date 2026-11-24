# -*- coding: utf-8 -*-
"""分项汇总 5 块合计行「累计值」按台账真值回填（该表合计已被清成静态值）。
计划列（各块的计划格）保持不动。"""
import sys

import win32com.client as win32

F0 = sys.argv[1]
# (块标题关键字, {列号: 值})
VALS = [
    ("1、9月订单计划", {4: 30639, 5: 1528.4307}),
    ("2、9月产出计划", {3: 39097}),
    ("3、9月出货计划", {4: 27148, 5: 3428.5298}),
    ("4、9月开票计划", {4: 3784.4273}),
    ("5、9月回款计划", {5: 2202.9844}),
]

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(F0)
    ws = wb.Worksheets("分项汇总")
    n = ws.UsedRange.Rows.Count
    titles = []
    for r in range(1, n + 1):
        v = str(ws.Cells(r, 1).Value or "").replace(" ", "")
        for kw, _ in VALS:
            if kw.replace(" ", "") in v:
                titles.append((kw, r))
    print("块标题:", titles)
    for i, (kw, r0) in enumerate(titles):
        r_end = titles[i + 1][1] - 1 if i + 1 < len(titles) else n
        total = None
        for r in range(r0, r_end + 1):
            if str(ws.Cells(r, 1).Value or "").strip() == "合计":
                total = r
        if not total:
            print("  ! 无合计行", kw)
            continue
        for col, val in dict(VALS)[kw].items():
            old = ws.Cells(total, col).Value
            ws.Cells(total, col).Value = val
            print("  %s 合计行%d 列%d: %r -> %r" % (kw, total, col, old, val))
    xl.CalculateFullRebuild()
    wb.Save()
    print("已保存")
finally:
    if wb:
        wb.Close(SaveChanges=False)
    xl.Quit()
