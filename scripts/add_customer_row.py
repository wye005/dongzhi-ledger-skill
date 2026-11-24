# -*- coding: utf-8 -*-
"""东智老宽表 F0 · 新增客户行（写值版，适配「清公式」后无公式的 F0）

改动两处 + 重算：
  1) `回款&开票统计表` 回款块：在「同业客户合计」行**之前**插入一行，B列=客户名，18日列填值，
     累计 AJ 沿用上一行 SUM（子表尚未清公式，SUM 仍有效）。
  2) `分项汇总` 回款块：在「其他类回款统计」行**之前**插入一行，A列=客户名，E 列直接写台账累计值
     （分项汇总已清公式，不再写公式）。
  3) 插行后行号漂移 → 重新跑 compute_fx（动态定位）+ compute_hz + flatten 固化。

用法：python add_customer_row.py <F0文件> <客户名> [当日列值] [日号]
例：  python add_customer_row.py "...9月18日...汇总表.xlsx" 鼎泰恒华 1.9998 18
"""
import sys, os, re, sqlite3
import win32com.client as win32
from dz_derive import load_month_cum, compute_fx, compute_hz, flatten_sheet

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
DB = r'E:\商务部\东智日常表更新\东智台账.db'

F0 = sys.argv[1]
NAME = sys.argv[2]
VAL = float(sys.argv[3]) if len(sys.argv) > 3 else None
DAY = int(sys.argv[4]) if len(sys.argv) > 4 else 18

# 本月前缀：从 F0 文件名「截止到2026年9月18日」解析
m = re.search(r'截止到(\d{4})年(\d{1,2})月', os.path.basename(F0))
month_prefix = ('%04d%02d' % (int(m.group(1)), int(m.group(2)))) if m else '202609'

# 台账累计（该客户本月回款）
conn = sqlite3.connect(DB)
cum_val = conn.execute(
    "SELECT COALESCE(SUM(回款金额),0) FROM 台账明细 WHERE 客户名称=? AND COALESCE(作废,0)=0 AND 导入批次 LIKE ?",
    (NAME, month_prefix + '%')).fetchone()[0]
conn.close()
print('台账累计回款 %s = %r' % (NAME, cum_val))

xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(F0)
    ws = wb.Worksheets('回款&开票统计表')

    # --- 回款表加行（子表尚未清公式，沿用上一行 SUM）---
    anchor = None
    for r in range(5, 96):
        if str(ws.Cells(r, 1).Value or '').strip() == '同业客户合计':
            anchor = r
            break
    if not anchor:
        print('! 未找到「同业客户合计」锚点行'); sys.exit(1)
    ws.Rows(anchor).Insert(-4121, 0)
    for c in range(3, 36):
        f = ws.Cells(anchor - 1, c).Formula
        if isinstance(f, str) and f.startswith('='):
            ws.Cells(anchor, c).Formula = f
        else:
            ws.Cells(anchor, c).Value = 0
    ws.Cells(anchor, 36).Formula = '=SUM(D%d:AI%d)' % (anchor, anchor)   # 累计（相对引用要显式写）
    ws.Cells(anchor, 2).Value = NAME
    if VAL is not None:
        col = 3 + DAY
        ws.Cells(anchor, col).Value = VAL
    # 同业客户合计 SUM 范围扩到新行
    sumrow = anchor + 1
    for c in range(3, 37):
        f = ws.Cells(sumrow, c).Formula
        if isinstance(f, str) and f.startswith('=SUM('):
            new = re.sub(r'([A-Z]{1,3})\d+\)', r'\g<1>%d)' % anchor, f)
            if new != f:
                ws.Cells(sumrow, c).Formula = new
    print('回款表新行 R%d = %s' % (anchor, NAME))

    # --- 分项汇总加行（已清公式 → 直接写值）---
    ws2 = wb.Worksheets('分项汇总')
    anchor2 = None
    for r in range(280, 362):
        if str(ws2.Cells(r, 1).Value or '').strip() == '其他类回款统计':
            anchor2 = r
            break
    if not anchor2:
        print('! 未找到「其他类回款统计」锚点行'); sys.exit(1)
    ws2.Rows(anchor2).Insert(-4121, 0)
    ws2.Cells(anchor2, 1).Value = NAME
    ws2.Cells(anchor2, 2).Value = 0                       # 开票应收
    ws2.Cells(anchor2, 3).Value = 0                       # 到期应收
    ws2.Cells(anchor2, 4).Value = 0                       # 回款计划
    ws2.Cells(anchor2, 5).Value = cum_val                 # E 累计回款（写值）
    ws2.Cells(anchor2, 6).Value = '-'                     # F 完成率
    ws2.Cells(anchor2, 7).Value = -cum_val                # G 待回款
    ws2.Cells(anchor2, 8).Value = 0.0                     # H 预计可回款
    ws2.Cells(anchor2, 9).Value = '-'                     # I 预计总回款率
    print('分项汇总新行 R%d = %s，E=%r' % (anchor2, NAME, cum_val))

    # --- 插行后行号漂移 → 重新代码算分项汇总 + 汇总 + 固化 ---
    xl.CalculateFullRebuild()
    cum = load_month_cum(month_prefix)
    fx_info = compute_fx(wb, cum)
    hz_info = compute_hz(wb, cum)
    flatten_sheet(wb, '分项汇总')
    flatten_sheet(wb, '汇总')
    print('重算：分项汇总改名 %d / 汇总收入 %r' % (fx_info['改名'], hz_info['收入']))

    wb.Save()
    print('已保存')
finally:
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception as e:
        print('close warn:', str(e)[:60])
    try:
        xl.Quit()
    except Exception:
        pass
print('DONE size=%d' % os.path.getsize(F0))
