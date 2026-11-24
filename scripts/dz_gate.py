# -*- coding: utf-8 -*-
"""东智台账 · 闭合校验（出表前强制关卡，FAIL 即禁止出表）

用法：python dz_gate.py <周期目录> <起> <止>

三层对账（源文件 / 台账 / F0）：
  A. 源 → 台账：OA签约、发票开票、现金+票据回款、NC出库单订单号
  B. 台账 → F0：客户覆盖（台账有F0无=FAIL）、客户级数值、合计
  C. F0 内部：分项合计 vs 汇总页实际
附加：数量/金额配比、源快照时点
退出码 0=PASS / 1=FAIL
"""
import os, sys, glob, sqlite3, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dz_common import f0_path, RULES, find_file, num_strict
from openpyxl import load_workbook

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

DB = r'E:\商务部\东智日常表更新\东智台账.db'
SRC = sys.argv[1]
D1 = sys.argv[2]
D2 = sys.argv[3]
BATCH = os.path.basename(os.path.normpath(SRC))
MONTH = D2[:7].replace('-', '')
AMT_TOL = 0.005
QTY_TOL = 0.5

fails, warns = [], []


def fail(m): fails.append(m)


def warn(m): warns.append(m)


KNOWN = set()


def can(name):
    """客户名归一：别名 → 去括号 → 关键字 → 与已知名做包含匹配（源全称↔台账简称）"""
    s = str(name or '').strip()
    for k, v in RULES['ALIAS'].items():
        if s == k:
            return v
    s = s.split('（')[0].split('(')[0].strip()
    for kw, tgt in RULES['SUBSTR_RULES']:
        if kw in s:
            return tgt
    if s in KNOWN:
        return s
    for k in KNOWN:
        if len(k) >= 2 and (k in s or s in k):
            return k
    return s


def N(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or '').strip().replace(',', '').replace('，', '')
    if s in ('', '-', '—', '/', '无'):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


# ---------------- A1. OA 签约源 ----------------
def load_half():
    """半成品清单（已确认计数量）"""
    p2 = r'E:\商务部\东智日常表更新\东智台账.xlsx'
    out = set()
    try:
        wb = load_workbook(p2, read_only=True, data_only=True)
        ws = wb['半成品清单']
        for row in ws.iter_rows(values_only=True):
            v = str(row[0] or '').strip()
            if v and v not in ('物料名称', '名称'):
                out.add(v)
        wb.close()
    except Exception:
        pass
    return out


HALF = load_half()


def read_oa():
    p = find_file(SRC, '合约订单评审表*.xlsx')
    if not p:
        warn('未找到 OA 合约订单评审表，跳过源对账')
        return None, {}
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(min_row=2, values_only=True)
    hdr = list(next(rows))
    ix = {str(h).strip(): i for i, h in enumerate(hdr) if h}
    need = ['制单日期', '评审状态', '客户名称', '公司订单号', '订单行数量', '含税金额', '物料名称']
    miss = [k for k in need if k not in ix]
    if miss:
        wb.close()
        raise SystemExit('OA 表头缺列：%s' % miss)
    agg, orders, pending = {}, set(), {}
    for r in rows:
        d = str(r[ix['制单日期']] or '')[:7]
        if d != D2[:7]:
            continue
        st = str(r[ix['评审状态']] or '')
        if any(k in st for k in RULES['VOID_KEYWORDS']):
            continue
        nm = can(r[ix['客户名称']])
        amt = N(r[ix['含税金额']]) / 10000.0
        qty = 0.0 if nm == '光缆事业部' else N(r[ix['订单行数量']])
        a = agg.setdefault(nm, dict(签约金额=0.0, 签约数量=0.0))
        a['签约金额'] += amt
        a['签约数量'] += qty
        mat = str(r[ix['物料名称']] or '').strip()
        if '半成品' in mat and mat not in HALF and qty > 0:
            pk = (nm, mat)
            p0 = pending.setdefault(pk, [0.0, 0.0])
            p0[0] += qty
            p0[1] += amt
        o = str(r[ix['公司订单号']] or '').strip()
        if o:
            orders.add(o)
    wb.close()
    return p, dict(agg=agg, orders=orders, pending=pending)


# ---------------- A2. 发票/回款源 ----------------
def read_invoice_register():
    """登记表源口径：按明细页「月份」列汇总，不读汇总页。
    2026-11-23 事故：汇总页漏计玖耀启铭 11/14 回款 18,190.7（其客户行公式缺失），
    但明细页 现金回款明细 里月份=11 是齐的。汇总页只是人工维护的看板，不能当源。"""
    p = find_file(SRC, '*发票及回款登记表*.xlsx')
    if not p:
        warn('未找到发票回款登记表，跳过源对账')
        return None
    wb = load_workbook(p, read_only=True, data_only=True)
    mon = int(D2[5:7])
    tot = {'开票金额': 0.0, '现金回款': 0.0, '票据回款': 0.0}
    ws = wb['2026年销售发票登记明细']
    for row in ws.iter_rows(min_row=3, values_only=True):
        d = row[7] if len(row) > 7 else None
        if isinstance(d, datetime.datetime):
            ds = d.strftime('%Y-%m')
        else:
            ds = str(d or '')[:7].replace('/', '-')
        if ds != D2[:7]:
            continue
        tot['开票金额'] += N(row[13] if len(row) > 13 else 0) / 10000.0
    for sheet, key in (('2026年现金回款明细', '现金回款'), ('2026年应收票据', '票据回款')):
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=3, values_only=True):
            m = row[10] if len(row) > 10 else None
            try:
                mm = int(float(m))
            except Exception:
                mm = None
            if mm != mon:
                continue
            tot[key] += N(row[5] if len(row) > 5 else 0) / 10000.0
    wb.close()
    tot['回款金额'] = tot['现金回款'] + tot['票据回款']
    return p, tot


# ---------------- A3. NC 出库单（本周）订单号 ----------------
def read_nc_orders():
    p = find_file(SRC, '销售出库单明细表*.csv')
    if p:
        import csv
        orders, amt = set(), 0.0
        with open(p, encoding='utf-8-sig', newline='') as f:
            for r in csv.DictReader(f):
                o = str(r.get('订单单号') or '').strip()
                if o:
                    orders.add(o)
                amt += N(r.get('金额'))
        return p, dict(orders=orders, amount=amt / 10000.0)
    p = find_file(SRC, '销售出库单明细表*.xls*', prefer_xls=True)
    if not p:
        return None, None
    try:
        import xlrd
        sh = xlrd.open_workbook(p).sheet_by_index(0)
        hdr = [str(sh.cell_value(1, c)).strip() for c in range(sh.ncols)]
        i_o = hdr.index('订单单号') if '订单单号' in hdr else None
        i_a = hdr.index('金额') if '金额' in hdr else None
        orders, amt = set(), 0.0
        for r in range(2, sh.nrows):
            o = str(sh.cell_value(r, i_o)).strip() if i_o is not None else ''
            if o:
                orders.add(o)
            if i_a is not None:
                amt += N(sh.cell_value(r, i_a))
        return p, dict(orders=orders, amount=amt / 10000.0)
    except Exception as e:
        warn('NC 出库单读取失败：%s' % e)
        return p, None


# ---------------- B. 台账 ----------------
def load_db_month():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute('''SELECT 客户名称,
        COALESCE(SUM(签约金额),0), COALESCE(SUM(签约数量),0), COALESCE(SUM(产出数量),0),
        COALESCE(SUM(发货数量),0), COALESCE(SUM(发货金额),0), COALESCE(SUM(开票金额),0), COALESCE(SUM(回款金额),0)
        FROM 台账明细
        WHERE COALESCE(作废,0)=0
          AND COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?
          AND COALESCE(NULLIF(时间,''),'')<=?
        GROUP BY 客户名称''', (D2[:7], D2))
    out = {}
    for nm, qa, qq, oq, sq, sa, kq, rq in cur.fetchall():
        k = can(nm)
        d = out.setdefault(k, dict(签约金额=0.0, 签约数量=0.0, 产出数量=0.0,
                                   发货数量=0.0, 发货金额=0.0, 开票金额=0.0, 回款金额=0.0))
        d['签约金额'] += qa; d['签约数量'] += qq; d['产出数量'] += oq
        d['发货数量'] += sq; d['发货金额'] += sa; d['开票金额'] += kq; d['回款金额'] += rq
    conn.close()
    return out


# ---------------- C. F0 ----------------
BLOCKS = {
    '订单计划': ('D', 'E', '签约数量', '签约金额'),
    '产出计划': ('C', None, '产出数量', None),
    '出货计划': ('D', 'E', '发货数量', '发货金额'),
    '开票计划': (None, 'D', None, '开票金额'),
    '回款计划': (None, 'E', None, '回款金额'),
}


def read_f0(path):
    wb = load_workbook(path, data_only=True)
    ws = wb['分项汇总']
    maxr = ws.max_row
    titles = {}
    for r in range(1, maxr + 1):
        a = str(ws.cell(r, 1).value or '')
        for k in BLOCKS:
            if k in a and k not in titles:
                titles[k] = r
    # 先收集 F0 里的客户名，作为模糊匹配的基准（源文件常用工商全称）
    for r in range(1, maxr + 1):
        a = str(ws.cell(r, 1).value or '').strip()
        if a and '合计' not in a and '统计' not in a:
            KNOWN.add(a.split('（')[0].split('(')[0].strip())
    cust, totals, rows_by_block = {}, {m: 0.0 for m in
        ('签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '开票金额', '回款金额')}, {}
    for key, trow in titles.items():
        qcol, acol, qm, am = BLOCKS[key]
        rows, r, tot_row = [], trow + 2, None
        while r <= maxr:
            a = str(ws.cell(r, 1).value or '').strip()
            if not a:
                r += 1
                continue
            if '合计' in a:
                tot_row = r
                break
            if '统计' in a:
                r += 1
                continue
            rows.append((r, a))
            r += 1
        rows_by_block[key] = rows
        for r, nm in rows:
            k = can(nm)
            d = cust.setdefault(k, {})
            if qcol:
                d[qm] = N(ws.cell(r, ord(qcol) - 64).value)
            if acol:
                d[am] = N(ws.cell(r, ord(acol) - 64).value)
        # 合计以 F0「合计行」为准（它包含「其他类回款统计」等汇总行），不再逐客户累加
        if tot_row:
            if qcol:
                totals[qm] = N(ws.cell(tot_row, ord(qcol) - 64).value)
            if acol:
                totals[am] = N(ws.cell(tot_row, ord(acol) - 64).value)
    hz = wb['汇总']
    summary = {}
    lab = {'签约订单金额': '签约金额', '签约订单数量': '签约数量', '产出数量': '产出数量',
           '发货数量': '发货数量', '发货金额': '发货金额', '开票金额': '开票金额', '回款金额': '回款金额'}
    for r in range(4, 12):
        t = str(hz.cell(r, 2).value or '')
        for k, v in lab.items():
            if k in t:
                summary[v] = N(hz.cell(r, 4).value)
    wb.close()
    return cust, totals, summary, rows_by_block


def main():
    target = f0_path(SRC, D2)
    if not os.path.exists(target):
        c = [p for p in glob.glob(os.path.join(SRC, '截止到*汇总表*.xlsx')) if '~$' not in p]
        if not c:
            print('FATAL: 周期目录无 F0'); return 1
        target = max(c, key=os.path.getmtime)

    db = load_db_month()
    f0_cust, f0_tot, summary, _ = read_f0(target)
    db_tot = {m: sum(v.get(m, 0.0) for v in db.values()) for m in f0_tot}

    print('=' * 80)
    print('东智闭合校验  %s ~ %s  批次 %s' % (D1, D2, BATCH))
    print('=' * 80)

    # ---- A. 源 → 台账 ----
    print('\n[A] 源文件 → 台账')
    p, oa = read_oa()
    if oa:
        for m in ('签约金额', '签约数量'):
            s = sum(v[m] for v in oa['agg'].values())
            tol = QTY_TOL if m.endswith('数量') else AMT_TOL
            flag = 'OK' if abs(s - db_tot[m]) <= tol else 'FAIL'
            print('  OA %-6s 源=%-14s 台账=%-14s 差=%-12s %s' %
                  (m, round(s, 6), round(db_tot[m], 6), round(db_tot[m] - s, 6), flag))
            if flag == 'FAIL':
                fail('OA→台账 %s 不平：源 %s / 台账 %s' % (m, round(s, 6), round(db_tot[m], 6)))
                for k, v in sorted(oa['agg'].items()):
                    sv = v[m]; dv = db.get(k, {}).get(m, 0.0)
                    if abs(sv - dv) > tol:
                        print('       - %-12s 源=%-12s 台账=%-12s 差=%s' %
                              (k, round(sv, 6), round(dv, 6), round(dv - sv, 6)))
                for k, v in sorted(db.items()):
                    if k not in oa['agg'] and abs(v.get(m, 0)) > tol:
                        print('       + %-12s 源=0  台账=%s（源里没有此人，需确认来源）' % (k, round(v[m], 6)))

    reg = read_invoice_register()
    if reg:
        _, rt = reg
        for m, srcv in (('开票金额', rt['开票金额']), ('回款金额', rt['回款金额'])):
            flag = 'OK' if abs(srcv - db_tot[m]) <= AMT_TOL else 'FAIL'
            print('  登记表 %-5s 源=%-14s 台账=%-14s 差=%-12s %s' %
                  (m, round(srcv, 6), round(db_tot[m], 6), round(db_tot[m] - srcv, 6), flag))
            if flag == 'FAIL':
                fail('登记表→台账 %s 不平：源 %s / 台账 %s' % (m, round(srcv, 6), round(db_tot[m], 6)))
                print('       现金=%s 票据=%s' % (round(rt['现金回款'], 6), round(rt['票据回款'], 6)))

    if oa and oa.get('pending'):
        print('\n[需要你确认] 半成品是否计数量（未确认前不得出表）')
        for (cust, mat), (q, a) in sorted(oa['pending'].items()):
            print('    %-10s / %-12s %s 台 / %s 万' % (cust, mat, round(q, 0), round(a, 4)))
            fail('半成品待确认：%s / %s（%s 台 / %s 万）→ 请确认是否计数量，确认后加入「半成品清单」'
                 % (cust, mat, round(q, 0), round(a, 4)))

    ncp, nc = read_nc_orders()
    if nc and oa:
        extra = sorted(nc['orders'] - oa['orders'])
        if extra:
            warn('NC 有出库、OA 无评审的订单号 %d 个：%s' % (len(extra), ', '.join(extra[:10])))
        print('  NC 出库单 本周=%s万 订单号=%d 个；OA 无评审=%d 个' %
              (round(nc['amount'], 6), len(nc['orders']), len(extra)))

    # ---- B. 台账 → F0 ----
    print('\n[B] 台账 → F0')
    other_bucket = set(RULES.get('F0_OTHER_BUCKET', ()))
    for k, v in sorted(db.items()):
        if k not in f0_cust:
            if k in other_bucket:
                print('  · 其他类回款承载：%s（计入「其他类回款统计」）' % k)
                continue
            nz = {a: round(b, 6) for a, b in v.items() if abs(b) > 1e-9}
            if nz:
                fail('客户覆盖：台账有、F0 无 → %s %s' % (k, nz))
                print('  x 客户丢失：%-12s %s' % (k, nz))
    for k, v in sorted(f0_cust.items()):
        dv = db.get(k, {})
        for m, got in v.items():
            want = dv.get(m, 0.0)
            tol = QTY_TOL if m.endswith('数量') else AMT_TOL
            if abs(got - want) > tol:
                fail('客户级 %s %s：F0=%s 台账=%s' % (k, m, round(got, 6), round(want, 6)))
                print('  x %-12s %-6s F0=%-12s 台账=%-12s 差=%s' %
                      (k, m, round(got, 6), round(want, 6), round(got - want, 6)))
    for m in f0_tot:
        tol = QTY_TOL if m.endswith('数量') else AMT_TOL
        if abs(f0_tot[m] - db_tot[m]) > tol:
            fail('合计 %s 不平：F0 %s / 台账 %s' % (m, round(f0_tot[m], 6), round(db_tot[m], 6)))
        if m in summary and abs(summary[m] - f0_tot[m]) > tol:
            fail('汇总页脱节 %s：汇总 %s / 分项 %s' % (m, round(summary[m], 6), round(f0_tot[m], 6)))

    # ---- C. 配比 ----
    for k, v in sorted(db.items()):
        if k == '光缆事业部':
            continue
        for q, a in (('签约数量', '签约金额'), ('发货数量', '发货金额')):
            if abs(v.get(q, 0)) > 0 and abs(v.get(a, 0)) <= 1e-9:
                fail('配比：%s %s=%s 但 %s=0（金额没进来）' % (k, q, round(v[q], 6), a))
            if abs(v.get(a, 0)) > 0 and abs(v.get(q, 0)) <= 1e-9:
                fail('配比：%s %s=%s 但 %s=0（数量被置0：半成品未确认/物料类？）'
                     % (k, a, round(v[a], 6), q))

    # ---- D. 时点 ----
    srcs = [p for p in glob.glob(os.path.join(SRC, '*')) if os.path.isfile(p)
            and '~$' not in p and not p.lower().endswith(('.md', '.pptx'))]
    if srcs:
        newest = max(os.path.getmtime(p) for p in srcs)
        if newest < os.path.getmtime(target):
            warn('源快照早于 F0：源最新 %s / F0 %s' % (
                datetime.datetime.fromtimestamp(newest).strftime('%m-%d %H:%M'),
                datetime.datetime.fromtimestamp(os.path.getmtime(target)).strftime('%m-%d %H:%M')))

    # ---- 输出 ----
    print('\n【合计三层对照】')
    for m in ('签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '开票金额', '回款金额'):
        print('  %-6s F0分项=%-16s 台账=%-16s 汇总页=%s' %
              (m, round(f0_tot.get(m, 0), 6), round(db_tot.get(m, 0), 6), round(summary.get(m, 0), 6)))
    if warns:
        print('\n【告警 %d】' % len(warns))
        for w in warns:
            print('  ! ' + w)
    if fails:
        print('\n【失败 %d】' % len(fails))
        for f in fails:
            print('  x ' + f)
        print('\n结论：FAIL —— 禁止出表')
        return 1
    print('\n结论：PASS —— 可以出表')
    return 0


if __name__ == '__main__':
    sys.exit(main())
