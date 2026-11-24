# -*- coding: utf-8 -*-
"""东智 F0 · 「预计收入」列(K) 修正（用户 2026-09-18："按实际的来"）

规则（以基准文件的原始公式为准，逐行判定）：
  · 原 K = 0            → 保持 0（有意为 0，如 光缆事业部（自主接单）/特发信息（自主接单）：收入计 0）
  · T01                → `=MAX(E,J)/1.13*0.019`（收入系数 0.019；且"实际累计 vs 预计总出货"取大）
  · 原 K = `=J/1.13`    → `=MAX(E,J)/1.13`（已超计划的按实际算，未超的不变）
用法：python fix_income_col.py <F0文件> <基准F0文件>
"""
import sys, os, re
import win32com.client as win32

F0 = sys.argv[1]
BASE = sys.argv[2] if len(sys.argv) > 2 else None


def orig_k(base, r):
    """取基准文件 K 列的原始值/公式"""
    if not base:
        return None
    return base.Worksheets('分项汇总').Cells(r, 11).Formula


xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
wb = bwb = None
try:
    wb = xl.Workbooks.Open(F0)
    if BASE and os.path.exists(BASE):
        bwb = xl.Workbooks.Open(BASE)
    fx = wb.Worksheets('分项汇总')
    n0 = nmax = nfix = 0
    for r in range(137, 205):
        nm = str(fx.Cells(r, 1).Value or '').strip()
        if not nm or '合计' in nm:
            continue
        o = orig_k(bwb, r)
        if o == 0 or o == '0':
            new = 0
            n0 += 1
        elif nm == 'T01':
            new = '=MAX(E%d,J%d)/1.13*0.019' % (r, r)
            nfix += 1
        elif isinstance(o, str) and re.match(r'^=J\d+/1\.13$', o):
            new = '=MAX(E%d,J%d)/1.13' % (r, r)
            nmax += 1
        else:
            new = o                      # 其他一律恢复原样
            nfix += 1
        if fx.Cells(r, 11).Formula != new:
            print('   R%-3d %-20s %r -> %r' % (r, nm[:20], fx.Cells(r, 11).Formula, new))
            fx.Cells(r, 11).Formula = new
    print('保持0 %d 行 / 改为MAX %d 行 / 其他 %d 行' % (n0, nmax, nfix))
    xl.CalculateFullRebuild()
    print('   K207 = %r ；E208(收入实际) = %r' % (fx.Cells(207, 11).Value, fx.Cells(208, 5).Value))
    wb.Save()
    print('已重算并保存')
finally:
    try:
        if bwb is not None:
            bwb.Close(SaveChanges=False)
    except Exception:
        pass
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
