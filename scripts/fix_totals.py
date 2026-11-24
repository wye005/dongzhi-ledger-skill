# -*- coding: utf-8 -*-
"""插行后修补：让各合计/总计行的公式把新插入的客户行纳入；
并把「分项汇总」的合计（已被清成静态值）按台账真值回填。
用法: python fix_totals.py <F0> <新行所在工作表>:<新行号> ...
"""
import re
import sys

import win32com.client as win32

F0 = sys.argv[1]
TARGETS = []
for spec in sys.argv[2:]:
    sheet, row = spec.rsplit(":", 1)
    TARGETS.append((sheet, int(row)))

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(F0)
    for sheet, newrow in TARGETS:
        ws = wb.Worksheets(sheet)
        maxc = ws.UsedRange.Columns.Count
        # 合计/总计行 = 新行下方的第一处「总计」或「合计」
        found = None
        for r in range(newrow + 1, newrow + 6):
            v = str(ws.Cells(r, 1).Value or "").strip()
            v2 = str(ws.Cells(r, 2).Value or "").strip()
            if v in ("总计", "合计") or v2 in ("总计", "合计"):
                found = r
                break
        if not found:
            print("  ! 未找到合计行", sheet, newrow)
            continue
        n = 0
        for c in range(3, maxc + 1):
            f = ws.Cells(found, c).Formula
            if not isinstance(f, str) or not f.startswith("="):
                continue
            col = ws.Cells(found, c).Address.replace("$", "")
            col = re.sub(r"\d+", "", col)
            if ("%s%d" % (col, newrow)) in f:
                continue
            m = re.fullmatch(r"=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", f)
            if m and m.group(1) == col and m.group(3) == col:
                ws.Cells(found, c).Formula = "=SUM(%s%s:%s%d)" % (col, m.group(2), col, newrow)
                n += 1
            elif re.fullmatch(r"=[A-Z]+\d+(\+[A-Z]+\d+)+", f):
                ws.Cells(found, c).Formula = f + "+%s%d" % (col, newrow)
                n += 1
        print("  %s 合计行 %d：修补 %d 个公式" % (sheet, found, n))

    xl.CalculateFullRebuild()
    wb.Save()
    print("已保存")
finally:
    if wb:
        wb.Close(SaveChanges=False)
    xl.Quit()
