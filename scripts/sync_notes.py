# -*- coding: utf-8 -*-
"""东智 F0 · 收尾修正（幂等，每周可重跑）
  1) 「汇总」预计待完成(C6)：套 MAX(0, …) —— 完成量超过计划时"待完成"应为 0，不该出现负数
     （C7 月度预计总完成 = D+F，自动跟着变正确）
  2) 「分项汇总」开票/回款备注：按客户名把本周明细续写进去（**不限「自主接单」客户**，已存在则跳过）
用法：python sync_notes.py <F0文件> <导入批次>
"""
import sys, os, sqlite3
import win32com.client as win32

F0 = sys.argv[1]
BATCH = sys.argv[2]
DB = r'E:\商务部\东智日常表更新\东智台账.db'


def fmt(ds):
    try:
        y, m, d = ds.split('-')
        return '%d/%d' % (int(m), int(d))
    except Exception:
        return ds


xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(F0)

    # ---------- 1) 汇总 C6 套 MAX(0, …) ----------
    hz = wb.Worksheets('汇总')
    for r in range(4, 12):
        f = hz.Cells(r, 6).Formula
        if isinstance(f, str) and f.startswith('=') and not f.upper().startswith('=MAX('):
            hz.Cells(r, 6).Formula = '=MAX(0,' + f[1:] + ')'
            print('   C6 R%d -> %r  值=%r' % (r, hz.Cells(r, 6).Formula, hz.Cells(r, 6).Value))

    # ---------- 2) 分项汇总 备注 ----------
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute("SELECT 时间,客户名称,开票金额 FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 导入批次=? "
                "AND 开票金额 IS NOT NULL AND 开票金额<>0 ORDER BY 时间", (BATCH,))
    kp = {}
    for ds, nm, amt in cur.fetchall():
        kp[(nm, ds)] = kp.get((nm, ds), 0) + amt
    cur.execute("SELECT 时间,客户名称,回款金额,回款类型 FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 导入批次=? "
                "AND 回款金额 IS NOT NULL AND 回款金额>0 ORDER BY 时间", (BATCH,))
    hk = {}
    for ds, nm, amt, typ in cur.fetchall():
        hk[(nm, ds, typ)] = hk.get((nm, ds, typ), 0) + amt
    conn.close()

    fx = wb.Worksheets('分项汇总')

    def append_note(rr, col, txt):
        old = fx.Cells(rr, col).Value
        if isinstance(old, str) and old.strip():
            if txt in old:            # 幂等：已有就不重复写
                return False
            fx.Cells(rr, col).Value = old + '，' + txt
        else:
            fx.Cells(rr, col).Value = txt
        return True

    def base_of(s):
        return str(s or '').split('（')[0].strip()

    def pick_rows(r1, r2):
        """返回 {客户基名: 行号}。同一客户若同时有「X」和「X（自主接单）」两行，
        **以（自主接单）为准**（否则会重复写两处备注）。"""
        plain, auto = {}, {}
        for r in range(r1, r2 + 1):
            a = str(fx.Cells(r, 1).Value or '').strip()
            if not a or '合计' in a or '统计' in a:
                continue
            if '（自主接单）' in a or '(自主接单)' in a:
                auto[base_of(a)] = r
            else:
                plain[a] = r
        out = dict(plain)
        out.update(auto)
        return out

    n1 = n2 = 0

    def block_range(title_kw):
        """2026-10-22：动态定位块范围（新增客户行会让硬编码行号漂移）。"""
        titles = ('1、9月订单计划', '2、9月产出计划', '3、9月出货计划', '4、9月开票计划', '5、9月回款计划')
        n = fx.UsedRange.Rows.Count
        start = None
        for r in range(1, n + 1):
            v = str(fx.Cells(r, 1).Value or '').replace(' ', '')
            if title_kw.replace(' ', '') in v:
                start = r
                break
        if start is None:
            return 0, -1
        end = n
        for r in range(start + 1, n + 1):
            v = str(fx.Cells(r, 1).Value or '').replace(' ', '')
            if any(t in v for t in titles):
                end = r - 1
                break
        return start + 1, end

    # 开票块 → H(8) 列
    a1, a2 = block_range('4、9月开票计划')
    for b, r in pick_rows(a1, a2).items():
        for (nm, ds), amt in sorted(kp.items(), key=lambda x: x[0][1]):
            if nm == b and append_note(r, 8, '%s开票%s万元' % (fmt(ds), round(amt, 2))):
                print('   ✓ 开票备注 R%d %s：%s开票%s万元' % (r, b, fmt(ds), round(amt, 2)))
                n1 += 1
    # 回款块 → J(10) 列
    b1, b2 = block_range('5、9月回款计划')
    for b, r in pick_rows(b1, b2).items():
        for (nm, ds, typ), amt in sorted(hk.items(), key=lambda x: x[0][1]):
            if nm == b and append_note(r, 10, '%s回款%s万元%s' % (fmt(ds), round(amt, 2), typ)):
                print('   ✓ 回款备注 R%d %s：%s回款%s万元%s' % (r, b, fmt(ds), round(amt, 2), typ))
                n2 += 1
    print('   新增备注：开票 %d 条 / 回款 %d 条' % (n1, n2))

    xl.CalculateFullRebuild()
    wb.Save()
    print('已重算并保存')
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
