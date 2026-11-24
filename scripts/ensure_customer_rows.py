# -*- coding: utf-8 -*-
"""为台账里出现、但 F0 没有落点的客户自动新增一行（默认只演练，--apply 才写）。

用法：
    python ensure_customer_rows.py <F0.xlsx> <客户名> [--apply]

覆盖：
  · 订单签约跟踪表：金额块(4-85) + 数量块(87-168)，插在「总计」前，AJ=SUM(D:AG)
  · 分项汇总：订单块(5-71)，插在「合计」前，B/C/D/E 指向刚插入的两行
其他块（产出/发货/开票/回款）若也缺行，脚本会列出来，提示按同样方式补。
"""
import sys, os, shutil, datetime
import win32com.client as win32

F0 = sys.argv[1]
NAME = sys.argv[2]
APPLY = '--apply' in sys.argv
XL_UP = -4162


def openpyxl_col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    bak = F0.replace('.xlsx', '.bak-%s.xlsx' % datetime.datetime.now().strftime('%Y%m%d%H%M%S'))
    if APPLY:
        shutil.copy2(F0, bak)
        print('备份：%s' % bak)

    xl = win32.DispatchEx('KET.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(F0)
    try:
        ws = wb.Worksheets('订单签约跟踪表')
        n_rows = ws.UsedRange.Rows.Count

        def find_total(start, end):
            for r in range(end, start, -1):
                v = str(ws.Cells(r, 2).Value or '').strip()
                a = str(ws.Cells(r, 1).Value or '').strip()
                if v == '总计' or a == '总计':
                    return r
            return None

        tot_amt = find_total(4, 86)
        tot_qty = find_total(87, 169)
        print('金额块总计行=%s  数量块总计行=%s' % (tot_amt, tot_qty))
        if not tot_amt or not tot_qty:
            raise SystemExit('找不到订单签约跟踪表的「总计」行')

        # 已存在就不插
        for r in range(4, 169):
            if str(ws.Cells(r, 2).Value or '').strip() == NAME:
                print('已存在：%s @ R%d，无需新增' % (NAME, r))
                return

        if not APPLY:
            print('[演练] 将在 订单签约跟踪表 R%d / R%d 前各插一行「%s」' % (tot_amt, tot_qty, NAME))
            return

        # 先插数量块（行号大的），再插金额块，避免行号漂移
        ws.Rows(tot_qty).Insert(XL_UP, 0)
        ws.Cells(tot_qty, 2).Value = NAME
        ws.Cells(tot_qty, 3).Value = 0
        ws.Cells(tot_qty, 36).Formula = '=SUM(D%d:AG%d)' % (tot_qty, tot_qty)   # AJ 累计
        ws.Cells(tot_qty, 37).Formula = '=IFERROR(AJ%d/C%d,"-")' % (tot_qty, tot_qty)  # AK 完成率
        qty_row = tot_qty
        # 原「总计」是分类小计相加，新行不在任何分类里 → 必须显式加进总计
        for c in range(3, 37):
            f = ws.Cells(tot_qty + 1, c).Formula
            if isinstance(f, str) and f.startswith('='):
                ws.Cells(tot_qty + 1, c).Formula = f + '+%s%d' % (openpyxl_col(c), tot_qty)

        ws.Rows(tot_amt).Insert(XL_UP, 0)
        ws.Cells(tot_amt, 2).Value = NAME
        ws.Cells(tot_amt, 3).Value = 0
        ws.Cells(tot_amt, 36).Formula = '=SUM(D%d:AG%d)' % (tot_amt, tot_amt)
        ws.Cells(tot_amt, 37).Formula = '=IFERROR(AJ%d/C%d,"-")' % (tot_amt, tot_amt)
        amt_row = tot_amt
        for c in range(3, 37):
            f = ws.Cells(tot_amt + 1, c).Formula
            if isinstance(f, str) and f.startswith('='):
                ws.Cells(tot_amt + 1, c).Formula = f + '+%s%d' % (openpyxl_col(c), tot_amt)

        fx = wb.Worksheets('分项汇总')
        tot = None
        for r in range(5, 80):
            if str(fx.Cells(r, 1).Value or '').strip() == '合计':
                tot = r
                break
        if not tot:
            raise SystemExit('分项汇总找不到订单块「合计」行')
        fx.Rows(tot).Insert(XL_UP, 0)
        fx.Cells(tot, 1).Value = NAME
        fx.Cells(tot, 2).Formula = "='订单签约跟踪表'!C%d" % qty_row
        fx.Cells(tot, 3).Formula = "='订单签约跟踪表'!C%d" % amt_row
        fx.Cells(tot, 4).Formula = "='订单签约跟踪表'!AJ%d" % qty_row
        fx.Cells(tot, 5).Formula = "='订单签约跟踪表'!AJ%d" % amt_row
        fx.Cells(tot, 6).Formula = '=IFERROR(E%d/C%d,"-")' % (tot, tot)
        fx.Cells(tot, 7).Formula = '=B%d-D%d' % (tot, tot)
        fx.Cells(tot, 8).Formula = '=C%d-E%d' % (tot, tot)
        fx.Cells(tot, 9).Formula = '=D%d+G%d' % (tot, tot)
        fx.Cells(tot, 10).Formula = '=E%d+H%d' % (tot, tot)
        fx.Cells(tot, 11).Formula = '=IFERROR(J%d/C%d,"-")' % (tot, tot)
        # 合计行原为 SUM(r1:tot-1)，新行在 tot → 重写范围把新行纳入
        for c in range(2, 6):          # B/C/D/E 才是求和
            fx.Cells(tot + 1, c).Formula = '=SUM(%s5:%s%d)' % (openpyxl_col(c), openpyxl_col(c), tot)
        t = tot + 1
        fx.Cells(t, 6).Formula = '=IFERROR(E%d/C%d,"-")' % (t, t)
        fx.Cells(t, 7).Formula = '=B%d-D%d' % (t, t)
        fx.Cells(t, 8).Formula = '=C%d-E%d' % (t, t)
        fx.Cells(t, 9).Formula = '=D%d+G%d' % (t, t)
        fx.Cells(t, 10).Formula = '=E%d+H%d' % (t, t)
        fx.Cells(t, 11).Formula = '=IFERROR(J%d/C%d,"-")' % (t, t)
        print('已新增：订单签约跟踪表 金额R%d / 数量R%d；分项汇总 R%d' % (amt_row, qty_row, tot))

        xl.CalculateFullRebuild()
        wb.Save()
        print('已保存')
    finally:
        wb.Close(False)
        xl.Quit()


main()
