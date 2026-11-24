# -*- coding: utf-8 -*-
r"""东智月度关账（v0.2）

一条命令完成规范月结：
  1) （可选）用 台账明细 按「归属月」重建本月 月度事实
  2) 校验 月度事实（客户-月唯一、8 项指标齐全）
  3) 生成 CANON 快照（dz_month.build_canonical）
  4) 监督对账（dz_month.check_canonical）
  5) （可选）与 F0 实绩合计逐项比对，写 校验结果
  6) 输出统一的 monthly.json，供 F2/F3/F4/PPT 消费

用法：
  python dz_close.py 2026-11 --from-ledger --f0 "…汇总表.xlsx"
  python dz_close.py 2026-10
  python dz_close.py 2026-10 --from-ledger --income finance

口径（用户 2026-11-23 定）：
  - 周更/月更同一套补录/作废规则；月结只是多出 F2–F6，并生成该月报表快照（不锁库）。
  - F2–F6 只读 monthly.json / 月度快照，不互相抄数。
  - 月结不写批次水印。
  - 已知差异（差异登记 里有的）不判 FAIL，新差异判 FAIL。
  - 收入唯一源 = 财务《特发东智收入明细表（YYYYMM）》（dz_income 入库）；F2 上那一格是手填常量，不是源。
    `--income auto`（默认）：有财务数据就用财务口径，没有就沿用台账（发货/1.13）并打印提示。
    `--income finance`：强制财务口径，缺数据直接失败。
    `--income ledger`：强制沿用台账口径（仅用于复现旧批次）。
"""
import os
import sys
import json
import shutil
import sqlite3
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dz_month import conn, build_canonical, check_canonical, METRICS  # noqa: E402

DEFAULT_JSON_DIR = r'E:\商务部\东智日常表更新\monthly'
DB_P = r'E:\商务部\东智日常表更新\东智台账.db'
BAK_DIR = r'E:\商务部\东智日常表更新\_backup'

AMT_TOL = 0.005
QTY_TOL = 0.5

# F0 分项汇总 5 块：块名 -> (数量列, 金额列, 数量指标, 金额指标)
F0_BLOCKS = {
    '订单计划': ('D', 'E', '签约数量', '签约金额'),
    '产出计划': ('C', None, '产出数量', None),
    '出货计划': ('D', 'E', '发货数量', '发货金额'),
    '开票计划': (None, 'D', None, '开票金额'),
    '回款计划': (None, 'E', None, '回款金额'),
}


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or '').strip().replace(',', '').replace('，', '')
    if s in ('', '-', '—', '/', '无'):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _backup(tag):
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    bak = os.path.join(BAK_DIR, '%s_%s' % (tag, ts))
    os.makedirs(bak, exist_ok=True)
    shutil.copy(DB_P, os.path.join(bak, os.path.basename(DB_P)))
    print('已备份 -> %s' % bak)
    return bak


def _apply_finance_income(cur, month):
    """用 收入明细（财务口径）覆盖本月 收入；签约/产出/发货/开票/回款仍取台账。"""
    inc = cur.execute(
        'SELECT 客户ID, 客户名称, ROUND(SUM(金额万元),6) FROM 收入明细 '
        'WHERE 年月=? GROUP BY 客户ID, 客户名称', (month,)).fetchall()
    if not inc:
        return None
    cur.execute("UPDATE 月度事实 SET 收入=0 WHERE 月份=? AND 版本='最新'", (month,))
    std = {cid: name for cid, name in cur.execute('SELECT 客户ID, 标准名 FROM 客户主数据')}
    updated = inserted = 0
    for cid, name, amt in inc:
        row = None
        if cid:
            row = cur.execute(
                "SELECT 客户名称 FROM 月度事实 WHERE 月份=? AND 版本='最新' AND 客户ID=?",
                (month, cid)).fetchone()
        if not row:
            row = cur.execute(
                "SELECT 客户名称 FROM 月度事实 WHERE 月份=? AND 版本='最新' AND 客户名称=?",
                (month, name)).fetchone()
        if row:
            cur.execute(
                "UPDATE 月度事实 SET 收入=? WHERE 月份=? AND 版本='最新' AND 客户名称=?",
                (amt, month, row[0]))
            updated += 1
        else:
            cust = std.get(cid, name)
            cur.execute('''INSERT INTO 月度事实
                (月份,客户ID,客户名称,签约金额,签约数量,产出数量,发货数量,发货金额,
                 收入,开票金额,回款金额,版本,来源,备注)
                VALUES(?,?,?,0,0,0,0,0,?,'最新','收入明细(财务)','财务表有、台账无该客户')''',
                (month, cid, cust, amt))
            inserted += 1
    total = cur.execute(
        "SELECT ROUND(SUM(收入),6) FROM 月度事实 WHERE 月份=? AND 版本='最新'",
        (month,)).fetchone()[0]
    return {'updated': updated, 'inserted': inserted, 'total': total}


def sync_from_ledger(month, income_source='auto'):
    """以 台账明细 的「归属月」为准，重建该月 月度事实（版本=最新）。"""
    _backup('close_from_ledger_' + month.replace('-', ''))
    c = conn(); cur = c.cursor()
    rows = cur.execute('''
        SELECT 客户ID, 客户名称,
               SUM(COALESCE(签约金额,0)), SUM(COALESCE(签约数量,0)),
               SUM(COALESCE(产出数量,0)), SUM(COALESCE(发货数量,0)),
               SUM(COALESCE(发货金额,0)), SUM(COALESCE(收入,0)),
               SUM(COALESCE(开票金额,0)), SUM(COALESCE(回款金额,0))
        FROM 台账明细
        WHERE COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?
          AND COALESCE(作废,0)=0 AND COALESCE(客户名称,'')<>''
        GROUP BY 客户ID, 客户名称
    ''', (month,)).fetchall()
    if not rows:
        print('! %s 台账没有可归集的明细' % month)
        c.close()
        return 1
    cur.execute("DELETE FROM 月度事实 WHERE 月份=? AND 版本='最新'", (month,))
    for r in rows:
        cur.execute('''INSERT INTO 月度事实
            (月份,客户ID,客户名称,签约金额,签约数量,产出数量,发货数量,发货金额,
             收入,开票金额,回款金额,版本,来源,备注)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,'最新',?,?)''',
            (month, r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9],
             '台账明细(归属月)', 'dz_close --from-ledger'))
    fin = None
    if income_source in ('auto', 'finance'):
        fin = _apply_finance_income(cur, month)
        if fin is None:
            if income_source == 'finance':
                print('! %s 收入明细 没有数据，无法按财务口径关账' % month)
                c.close()
                return 1
            print('  · 收入明细 无 %s 数据，收入沿用台账（发货/1.13 口径）' % month)
        else:
            print('OK 收入改用财务源表：%d 户更新，%d 户补建，合计 %.6f 万元'
                  % (fin['updated'], fin['inserted'], fin['total']))
    c.commit()
    n = cur.execute("SELECT COUNT(*) FROM 月度事实 WHERE 月份=? AND 版本='最新'", (month,)).fetchone()[0]
    c.close()
    print('OK 从台账重建 %s 月度事实：%d 个客户' % (month, n))
    return 0


def _check_completeness(cur, month):
    row = cur.execute('''
        SELECT COUNT(*), COUNT(DISTINCT 客户名称),
               SUM(CASE WHEN 签约金额 IS NULL OR 签约数量 IS NULL OR 产出数量 IS NULL
                         OR 发货数量 IS NULL OR 发货金额 IS NULL OR 收入 IS NULL
                         OR 开票金额 IS NULL OR 回款金额 IS NULL THEN 1 ELSE 0 END)
        FROM 月度事实
        WHERE 月份=? AND 版本='最新' AND COALESCE(客户名称,'')<>''
    ''', (month,)).fetchone()
    rows, customers, missing = int(row[0] or 0), int(row[1] or 0), int(row[2] or 0)
    dup = 0
    for _cust, n in cur.execute('''
        SELECT 客户名称, COUNT(*) FROM 月度事实
        WHERE 月份=? AND 版本='最新' GROUP BY 客户名称 HAVING COUNT(*)>1
    ''', (month,)):
        dup += int(n) - 1
    return {'事实行数': rows, '客户数': customers, '缺失指标格': missing, '重复客户行': dup}


def _valid_batch(cur, month):
    return cur.execute('''
        SELECT 批次ID, 月份, 截止时间, 源文件哈希, 生成时间, 脚本版本
        FROM 报告批次 WHERE 月份=? AND 状态='有效'
        ORDER BY 生成时间 DESC LIMIT 1
    ''', (month,)).fetchone()


def _snapshot(cur, batch):
    out = {}
    for cust, metric, val, n in cur.execute('''
        SELECT 客户名称, 指标, 数值, 来源数 FROM 月度快照 WHERE 批次ID=?
    ''', (batch,)):
        out.setdefault(cust, {})[metric] = float(val or 0)
    return out


def _snapshot_totals(cur, batch):
    return {m: float(v or 0) for m, v in cur.execute(
        'SELECT 指标, SUM(数值) FROM 月度快照 WHERE 批次ID=? GROUP BY 指标', (batch,))}


def _read_f0_totals(path):
    """读 F0「分项汇总」5 块的合计行实绩。返回 {指标: 值}。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb['分项汇总']
    maxr, titles = ws.max_row, {}
    for r in range(1, maxr + 1):
        a = str(ws.cell(r, 1).value or '')
        for k in F0_BLOCKS:
            if k in a and k not in titles:
                titles[k] = r
    out = {}
    for key, trow in titles.items():
        qcol, acol, qm, am = F0_BLOCKS[key]
        r, tot = trow + 2, None
        while r <= maxr:
            a = str(ws.cell(r, 1).value or '').strip()
            if not a:
                r += 1
                continue
            if '合计' in a:
                tot = r
                break
            r += 1
        if not tot:
            continue
        if qcol:
            out[qm] = _num(ws.cell(tot, ord(qcol) - 64).value)
        if acol:
            out[am] = _num(ws.cell(tot, ord(acol) - 64).value)
    wb.close()
    return out


def _ledger_totals(cur, month):
    row = cur.execute('''SELECT
        SUM(COALESCE(签约金额,0)), SUM(COALESCE(签约数量,0)),
        SUM(COALESCE(产出数量,0)), SUM(COALESCE(发货数量,0)),
        SUM(COALESCE(发货金额,0)), SUM(COALESCE(开票金额,0)),
        SUM(COALESCE(回款金额,0))
        FROM 台账明细
        WHERE COALESCE(NULLIF(归属月,''),substr(时间,1,7))=? AND COALESCE(作废,0)=0''', (month,)).fetchone()
    keys = ('签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '开票金额', '回款金额')
    return {k: float(v or 0) for k, v in zip(keys, row)}


def _known_diff(cur, month, metric, diff):
    for (v,) in cur.execute('SELECT 差异 FROM 差异登记 WHERE 月份=? AND 指标=?', (month, metric)):
        if v is not None and abs(float(v) - float(diff)) <= 0.002:
            return True
    return False


def _compare_f0(cur, month, batch, f0_path):
    f0 = _read_f0_totals(f0_path)
    snap = _snapshot_totals(cur, batch)
    led = _ledger_totals(cur, month)
    rows, fails = [], []
    for m in ('签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '开票金额', '回款金额'):
        if m not in f0:
            continue
        tol = QTY_TOL if m.endswith('数量') else AMT_TOL
        # 快照（月度事实）对 F0；台账原始明细对 F0 —— 两条都留痕
        for stage, src in (('月结F0', snap), ('台账F0', led)):
            diff = float(src.get(m, 0)) - float(f0[m])
            known = abs(diff) > tol and _known_diff(cur, month, m, diff)
            verdict = 'PASS' if abs(diff) <= tol else ('PASS(已登记差异)' if known else 'FAIL')
            if verdict == 'FAIL':
                fails.append((stage, m, src.get(m, 0), f0[m], diff))
            cur.execute('''INSERT OR REPLACE INTO 校验结果
                (批次ID,月份,环节,检查项,源值,台账值,看板值,差异,容差,结论,详情,生成时间)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,datetime('now','localtime'))''',
                (batch, month, stage, m, None, round(float(src.get(m, 0)), 10),
                 round(float(f0[m]), 10), round(diff, 10), tol, verdict, os.path.basename(f0_path)))
            rows.append({'环节': stage, '指标': m, '本表': round(float(src.get(m, 0)), 6),
                         'F0': round(float(f0[m]), 6), '差异': round(diff, 6), '结论': verdict})
    return rows, fails


def _build_json(cur, month, batch_row, completeness, snapshot_ok, f0_rows=None, f0_fails=None):
    batch, _m, cutoff, src_hash, gen_at, script = batch_row
    snap = _snapshot(cur, batch)
    totals = {m: round(sum(row.get(m, 0) for row in snap.values()), 6) for m in METRICS}
    customers = []
    for cust, row in snap.items():
        item = {'客户名称': cust}
        item.update({m: round(float(row.get(m, 0)), 6) for m in METRICS})
        customers.append(item)
    customers.sort(key=lambda x: (-x['签约金额'], -x['收入'], x['客户名称']))

    top_metrics = ['签约金额', '发货金额', '收入', '开票金额', '回款金额']
    top5 = {m: [{'客户名称': c['客户名称'], '数值': c[m]} for c in sorted(customers, key=lambda x: -x[m])[:5]]
            for m in top_metrics}

    control = []
    for m, ctrl, fact, diff, major in cur.execute('''
        SELECT 指标, 控制值, 事实值, 差异, 重大差异
        FROM v_month_recon WHERE 月份=? ORDER BY 指标
    ''', (month,)):
        control.append({'指标': m, '控制值': ctrl, '事实值': fact, '差异': diff, '重大差异': major})

    qty = []
    for m, actual, ctrl, diff in cur.execute('''
        SELECT 指标, 实际值, 控制值, 差异 FROM v_month_quantity_recon
        WHERE 月份=? ORDER BY 指标
    ''', (month,)):
        qty.append({'指标': m, '实际值': actual, '控制值': ctrl, '差异': diff})

    restatements = []
    for m, old, new, diff, note, source in cur.execute('''
        SELECT 指标, 当月定稿, 最新重述, 差异, 说明, 来源 FROM 月度版本
        WHERE 月份=? ORDER BY 指标
    ''', (month,)):
        restatements.append({'指标': m, '当月定稿': old, '最新重述': new, '差异': diff, '说明': note, '来源': source})

    audits = []
    for m, indicator, old, new, diff, owner, reason, action in cur.execute('''
        SELECT 月份, 指标, 原值, 新值, 差异, 责任方, 原因, 处理
        FROM v_audit_norm WHERE 月份=? ORDER BY 指标
    ''', (month,)):
        audits.append({'指标': indicator, '原值': old, '新值': new, '差异': diff,
                       '责任方': owner, '原因': reason, '处理': action})

    raw_diffs = []
    for r in cur.execute('SELECT * FROM v_month_raw_vs_canonical WHERE 月份=?', (month,)):
        raw_diffs.append({
            '月份': r[0], '签约金额差': r[1], '签约数量差': r[2], '产出数量差': r[3],
            '发货数量差': r[4], '发货金额差': r[5], '收入差': r[6],
            '开票金额差': r[7], '回款金额差': r[8], '说明': r[9]
        })

    major = [x for x in control if x['重大差异'] == '是']
    qty_major = [x for x in qty if x['差异'] is not None and abs(float(x['差异'])) > 0.5]
    checks = {
        '规范范围': month >= '2026-09',
        '事实完整': completeness['缺失指标格'] == 0 and completeness['重复客户行'] == 0,
        '控制重大差异数': len(major),
        '数量重大差异数': len(qty_major),
        '快照与事实一致': bool(snapshot_ok),
        'F0比对': '未执行' if f0_rows is None else 'PASS' if not f0_fails else 'FAIL',
    }
    return {
        'schema_version': '1.1',
        'month': month,
        'generated_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'batch': {'批次ID': batch, '截止时间': cutoff, '源文件哈希': src_hash,
                  '生成时间': gen_at, '脚本版本': script},
        'completeness': completeness,
        'metrics': METRICS,
        'totals': totals,
        'customers': customers,
        'top5': top5,
        'control': control,
        'quantity_control': qty,
        'restatements': restatements,
        'audit_flags': audits,
        'raw_ledger_diffs': raw_diffs,
        'f0_check': f0_rows,
        'checks': checks,
    }


def close(month, json_dir=DEFAULT_JSON_DIR, from_ledger=False, f0_path=None, income_source='auto'):
    if from_ledger:
        rc = sync_from_ledger(month, income_source)
        if rc:
            return rc

    c = conn(); cur = c.cursor()
    completeness = _check_completeness(cur, month)
    c.close()
    if completeness['事实行数'] == 0:
        print('! %s 没有 月度事实 行，不能关账' % month)
        return 1
    if completeness['缺失指标格'] or completeness['重复客户行']:
        print('! %s 月度事实不完整：%s' % (month, completeness))
        return 1

    rc = build_canonical(month)
    if rc:
        return rc
    rc = check_canonical(month)

    c = conn(); cur = c.cursor()
    batch_row = _valid_batch(cur, month)
    if not batch_row:
        print('! %s 关账后没有有效批次' % month)
        c.close()
        return 1

    f0_rows, f0_fails = None, None
    if f0_path:
        if not os.path.exists(f0_path):
            print('! F0 文件不存在：%s' % f0_path)
            c.close()
            return 1
        f0_rows, f0_fails = _compare_f0(cur, month, batch_row[0], f0_path)
        print('F0 比对：%s' % ('PASS' if not f0_fails else 'FAIL'))
        for x in f0_rows:
            print('   %-6s %-6s 本表=%-14s F0=%-14s 差=%-12s %s' %
                  (x['环节'], x['指标'], x['本表'], x['F0'], x['差异'], x['结论']))

    payload = _build_json(cur, month, batch_row, completeness, rc == 0, f0_rows, f0_fails)
    c.close()

    os.makedirs(json_dir, exist_ok=True)
    path = os.path.join(json_dir, '%s_monthly.json' % month)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)

    checks = payload['checks']
    ok = (checks['控制重大差异数'] == 0 and checks['数量重大差异数'] == 0
          and checks['事实完整'] and checks['快照与事实一致']
          and checks['F0比对'] in ('PASS', '未执行'))
    print(('OK' if ok and rc == 0 else 'WARN') + ' %s 关账' % month)
    print('   客户=%d 指标=%d 批次=%s' % (completeness['客户数'], len(METRICS), batch_row[0]))
    print('   控制重大差异=%d 数量重大差异=%d F0=%s' %
          (checks['控制重大差异数'], checks['数量重大差异数'], checks['F0比对']))
    print('   JSON=%s' % path)
    return 0 if ok and rc == 0 else 1


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return
    month = args[0]
    json_dir = DEFAULT_JSON_DIR
    if '--json-dir' in args:
        json_dir = args[args.index('--json-dir') + 1]
    f0 = None
    if '--f0' in args:
        f0 = args[args.index('--f0') + 1]
    income = 'auto'
    if '--income' in args:
        income = args[args.index('--income') + 1]
        if income not in ('auto', 'finance', 'ledger'):
            print('! --income 只能是 auto / finance / ledger')
            sys.exit(2)
    sys.exit(close(month, json_dir, '--from-ledger' in args, f0, income))


if __name__ == '__main__':
    main()
