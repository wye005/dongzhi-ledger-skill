# -*- coding: utf-8 -*-
r"""东智历史月度入库 + 2025/2026 月度对比

只取《2025年&2026年销售数据统计-明细》里的 2025 年客户级月数据
（订单/开票/回款/收入/发货，万元）落到 SQLite；再与 `月度事实`（已关账月份）
和 `台账明细`（当月实时）拼成统一的对比视图，做对比表时不用再开 Excel。

注意：该文件的 2026 块是重复数据，用户 2026-11-23 明确不维护、不比对；
2026 一律以 `台账明细` / `月度事实` 为准。

用法：
  python dz_history.py import "…2025年&2026年销售数据统计-明细（更新至26年11月）.xlsx"
  python dz_history.py check
  python dz_history.py report [--out …]
  python dz_history.py detail 中兴 [--year 2025]
"""
import os
import re
import sys
import io
import sqlite3
import argparse
import datetime

import openpyxl
from openpyxl.utils import get_column_letter

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DB = r'E:\商务部\东智日常表更新\东智台账.db'
SHEET = '2025-2026年数据比较'
METRICS = ['订单', '开票', '回款', '收入', '发货']
UNIT = '万元'
# 历史表里的叫法与主数据不完全一致，先做一次显式归一
ALIAS = {
    '慧捷': '慧捷（上海）科技',
    '新基建事业部': '新基建',
    '特发信息股份有限公司': '特发信息',
}

DDL = """
CREATE TABLE IF NOT EXISTS 历史月度(
    年份 INTEGER NOT NULL,
    月份 INTEGER NOT NULL,
    客户ID TEXT,
    客户名称 TEXT NOT NULL,
    指标 TEXT NOT NULL,
    数值 REAL NOT NULL,
    单位 TEXT DEFAULT '万元',
    来源文件 TEXT,
    来源单元格 TEXT,
    导入时间 TEXT,
    PRIMARY KEY(年份, 月份, 客户名称, 指标)
);

CREATE TABLE IF NOT EXISTS 历史月度合计(
    年份 INTEGER NOT NULL,
    月份 INTEGER NOT NULL,
    指标 TEXT NOT NULL,
    数值 REAL NOT NULL,
    单位 TEXT DEFAULT '万元',
    来源文件 TEXT,
    来源单元格 TEXT,
    导入时间 TEXT,
    PRIMARY KEY(年份, 月份, 指标)
);

CREATE INDEX IF NOT EXISTS idx_hist_年月指标 ON 历史月度(年份, 月份, 指标);
CREATE INDEX IF NOT EXISTS idx_hist_客户 ON 历史月度(客户ID);
"""

VIEW_DDL = """
DROP VIEW IF EXISTS v_台账月度;
CREATE VIEW v_台账月度 AS
SELECT CAST(substr(归属月,1,4) AS INTEGER) AS 年份,
       CAST(substr(归属月,6,2) AS INTEGER) AS 月份,
       '订单' AS 指标, round(sum(COALESCE(签约金额,0)),6) AS 数值
FROM 台账明细
WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]'
GROUP BY 归属月
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       '开票', round(sum(COALESCE(开票金额,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' GROUP BY 归属月
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       '回款', round(sum(COALESCE(回款金额,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' GROUP BY 归属月
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       '收入', round(sum(COALESCE(收入,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' GROUP BY 归属月
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       '发货', round(sum(COALESCE(发货金额,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' GROUP BY 归属月;

DROP VIEW IF EXISTS v_台账月度_客户;
CREATE VIEW v_台账月度_客户 AS
SELECT CAST(substr(归属月,1,4) AS INTEGER) AS 年份,
       CAST(substr(归属月,6,2) AS INTEGER) AS 月份,
       COALESCE(客户ID,'') AS 客户ID, 客户名称, '订单' AS 指标,
       round(sum(COALESCE(签约金额,0)),6) AS 数值
FROM 台账明细
WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' AND COALESCE(客户名称,'')<>''
GROUP BY 归属月, 客户ID, 客户名称
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       COALESCE(客户ID,''), 客户名称, '开票', round(sum(COALESCE(开票金额,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' AND COALESCE(客户名称,'')<>'' GROUP BY 归属月, 客户ID, 客户名称
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       COALESCE(客户ID,''), 客户名称, '回款', round(sum(COALESCE(回款金额,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' AND COALESCE(客户名称,'')<>'' GROUP BY 归属月, 客户ID, 客户名称
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       COALESCE(客户ID,''), 客户名称, '收入', round(sum(COALESCE(收入,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' AND COALESCE(客户名称,'')<>'' GROUP BY 归属月, 客户ID, 客户名称
UNION ALL
SELECT CAST(substr(归属月,1,4) AS INTEGER), CAST(substr(归属月,6,2) AS INTEGER),
       COALESCE(客户ID,''), 客户名称, '发货', round(sum(COALESCE(发货金额,0)),6)
FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 归属月 GLOB '20[0-9][0-9]-[0-9][0-9]' AND COALESCE(客户名称,'')<>'' GROUP BY 归属月, 客户ID, 客户名称;

DROP VIEW IF EXISTS v_月度对比;
CREATE VIEW v_月度对比 AS
SELECT 年份, 月份, 指标, round(sum(数值),6) AS 数值, '历史月度' AS 来源
FROM 历史月度 GROUP BY 年份, 月份, 指标
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), '订单', round(sum(COALESCE(签约金额,0)),6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026' GROUP BY 月份
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), '开票', round(sum(COALESCE(开票金额,0)),6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026' GROUP BY 月份
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), '回款', round(sum(COALESCE(回款金额,0)),6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026' GROUP BY 月份
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), '收入', round(sum(COALESCE(收入,0)),6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026' GROUP BY 月份
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), '发货', round(sum(COALESCE(发货金额,0)),6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026' GROUP BY 月份
UNION ALL
SELECT 年份, 月份, 指标, 数值, '台账实时'
FROM v_台账月度 t
WHERE 年份=2026
  AND NOT EXISTS (SELECT 1 FROM 月度事实 f
                  WHERE substr(f.月份,1,7)=printf('%04d-%02d', t.年份, t.月份));

DROP VIEW IF EXISTS v_月度对比_客户;
CREATE VIEW v_月度对比_客户 AS
SELECT 年份, 月份, COALESCE(客户ID,''), 客户名称, 指标, 数值, '历史月度' AS 来源 FROM 历史月度
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), COALESCE(客户ID,''), 客户名称, '订单', round(签约金额,6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026'
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), COALESCE(客户ID,''), 客户名称, '开票', round(开票金额,6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026'
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), COALESCE(客户ID,''), 客户名称, '回款', round(回款金额,6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026'
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), COALESCE(客户ID,''), 客户名称, '收入', round(收入,6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026'
UNION ALL
SELECT 2026, CAST(substr(月份,6,2) AS INTEGER), COALESCE(客户ID,''), 客户名称, '发货', round(发货金额,6), '月度事实'
FROM 月度事实 WHERE substr(月份,1,4)='2026'
UNION ALL
SELECT t.年份, t.月份, t.客户ID, t.客户名称, t.指标, t.数值, '台账实时'
FROM v_台账月度_客户 t
WHERE t.年份=2026
  AND NOT EXISTS (SELECT 1 FROM 月度事实 f
                  WHERE substr(f.月份,1,7)=printf('%04d-%02d', t.年份, t.月份));

DROP VIEW IF EXISTS v_月度同比;
CREATE VIEW v_月度同比 AS
SELECT a.月份, a.指标, a.数值 AS "2025", b.数值 AS "2026",
       round(COALESCE(b.数值,0)-a.数值,6) AS 差异,
       CASE WHEN a.数值<>0 AND b.数值 IS NOT NULL
            THEN round((b.数值-a.数值)/a.数值,6) END AS 增幅
FROM v_月度对比 a
LEFT JOIN v_月度对比 b
  ON b.年份=2026 AND b.月份=a.月份 AND b.指标=a.指标
WHERE a.年份=2025;
"""


def connect(db=DB):
    c = sqlite3.connect(db)
    ensure_schema(c)
    return c


def ensure_schema(c):
    c.executescript(DDL)
    c.executescript(VIEW_DDL)
    c.commit()


def _now():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def norm_name(s):
    s = str(s or '').strip()
    s = re.sub(r'[（(][^）)]*[）)]\s*$', '', s)
    s = s.replace('股份有限公司', '').strip()
    return s


def load_master(c):
    m = {}
    for cid, std, alias in c.execute('SELECT 客户ID, 标准名, 别名 FROM 客户主数据'):
        m[norm_name(std)] = cid
        for a in re.split(r'[、,，;；/]', alias or ''):
            a = a.strip()
            if a:
                m.setdefault(norm_name(a), cid)
    return m


def _new_customer(c, name):
    n = c.execute('SELECT COALESCE(MAX(CAST(substr(客户ID,2) AS INTEGER)),0) FROM 客户主数据').fetchone()[0]
    cid = 'C%03d' % (n + 1)
    c.execute(
        'INSERT INTO 客户主数据(客户ID,标准名,别名,类别,客户状态,生效月,备注,录入时间) VALUES(?,?,?,?,?,?,?,?)',
        (cid, name, '', '', '待确认', '2025-01', '历史导入自动建', _now()))
    print('  + 新客户(待确认) %s %s' % (cid, name))
    return cid


def _resolve(master, c, name):
    key = norm_name(ALIAS.get(name, name))
    if key in master:
        return master[key]
    return _new_customer(c, name)


def _find_title(ws, keyword):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and keyword in v:
            return r
    return None


def import_year(path, year=2025, db=DB):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[SHEET]
    title = _find_title(ws, '%d年销售数据统计表' % year)
    if not title:
        raise SystemExit('找不到 %d 年数据块：%s' % (year, path))
    header = title + 2
    if str(ws.cell(header, 2).value or '').strip() != '总计':
        raise SystemExit('表头位置不符：第 %d 行 B 列应为“总计”' % header)
    first = title + 4
    if str(ws.cell(header + 1, 7).value or '').strip() != '订单':
        raise SystemExit('指标表头不符：第 %d 行 G 列应为“订单”' % (header + 1))

    c = sqlite3.connect(db)
    ensure_schema(c)
    master = load_master(c)
    src = os.path.basename(path)
    now = _now()

    rows = []
    r = first
    total_ref = ''
    while r <= ws.max_row:
        name = ws.cell(r, 1).value
        if name is None or str(name).strip() == '':
            break
        name = str(name).strip()
        if name.startswith('总计'):
            total_ref = 'B%d' % r
            break
        if name == '实际回款':
            r += 1
            continue
        vals = []
        for m in range(1, 13):
            base = 7 + (m - 1) * 5
            for i, metric in enumerate(METRICS):
                col = base + i
                v = ws.cell(r, col).value
                if isinstance(v, (int, float)) and abs(float(v)) > 1e-9:
                    vals.append((m, metric, float(v), '%s%d' % (get_column_letter(col), r)))
        if vals:
            cid = _resolve(master, c, name)
            for m, metric, v, ref in vals:
                rows.append((year, m, cid, name, metric, v, UNIT, src, ref, now))
        r += 1
    c.execute('DELETE FROM 历史月度 WHERE 年份=?', (year,))
    c.execute('DELETE FROM 历史月度合计 WHERE 年份=?', (year,))
    c.executemany(
        'INSERT INTO 历史月度(年份,月份,客户ID,客户名称,指标,数值,单位,来源文件,来源单元格,导入时间) '
        'VALUES(?,?,?,?,?,?,?,?,?,?)', rows)
    # 总计行按月拆开：从明细逐月求和，保证和客户行同口径
    agg = {}
    for y, m, cid, n, metric, v, *rest in rows:
        agg[(m, metric)] = agg.get((m, metric), 0.0) + v
    tot = [(year, m, metric, round(agg.get((m, metric), 0.0), 6), UNIT, src,
            total_ref, now)
           for m in range(1, 13) for metric in METRICS]
    c.executemany(
        'INSERT INTO 历史月度合计(年份,月份,指标,数值,单位,来源文件,来源单元格,导入时间) '
        'VALUES(?,?,?,?,?,?,?,?)', tot)
    c.commit()

    print('导入 %d 年：客户月记录 %d 条，合计 %d 条（12 月 × 5 指标）' % (year, len(rows), len(tot)))
    print('  客户数 %d，来源 %s' % (len(set(x[3] for x in rows)), src))
    for metric in METRICS:
        s = sum(agg.get((m, metric), 0.0) for m in range(1, 13))
        print('  %-3s 全年 %14.6f 万元' % (metric, s))
    c.close()


def check(db=DB):
    c = sqlite3.connect(db)
    ensure_schema(c)
    print('== 1. 历史月度：客户行 vs 合计行 ==')
    bad = 0
    for year, m, metric, total in c.execute(
            'SELECT 年份,月份,指标,数值 FROM 历史月度合计 ORDER BY 年份,月份,指标'):
        s = c.execute('SELECT COALESCE(SUM(数值),0) FROM 历史月度 WHERE 年份=? AND 月份=? AND 指标=?',
                      (year, m, metric)).fetchone()[0]
        if abs(s - total) > 0.000001:
            bad += 1
            print('  FAIL %d-%02d %s 客户行=%.6f 合计行=%.6f 差=%.6f' % (year, m, metric, s, total, s - total))
    print('  %s（%d 项）' % ('PASS' if bad == 0 else 'FAIL', 60))
    print('== 2. 2026 已关账月份 vs 台账（同口径对照） ==')
    metric_map = {'订单': '签约金额', '开票': '开票金额', '回款': '回款金额',
                  '收入': '收入', '发货': '发货金额'}
    diffs = 0
    checked = 0
    for (month,) in c.execute("SELECT DISTINCT 月份 FROM 月度事实 WHERE substr(月份,1,4)='2026' ORDER BY 月份"):
        y = int(month[:4]); m = int(month[5:7])
        has = c.execute('SELECT COUNT(*) FROM v_台账月度 WHERE 年份=? AND 月份=?', (y, m)).fetchone()[0]
        if not has:
            continue
        for metric, col in (('订单', '签约金额'), ('开票', '开票金额'), ('回款', '回款金额'),
                            ('收入', '收入'), ('发货', '发货金额')):
            checked += 1
            fact = c.execute('SELECT round(COALESCE(SUM(%s),0),6) FROM 月度事实 WHERE 月份=?' % col,
                             (month,)).fetchone()[0]
            # 收入自 2026-11-23 起以财务《收入明细》为唯一源，不再与台账（发货/1.13 推导）比
            if metric == '收入':
                src = c.execute('SELECT ROUND(SUM(金额万元),6) FROM 收入明细 WHERE 年月=?',
                                (month,)).fetchone()[0]
                if src is not None:
                    d = fact - src
                    if abs(d) <= 0.005:
                        continue
                    reg = c.execute(
                        'SELECT 差异 FROM 差异登记 WHERE 月份=? AND 指标=? '
                        'AND 状态 IN (\'已解释\',\'已归档\',\'已按源表更正\') '
                        'AND ABS(ABS(差异)-?)<=0.005 LIMIT 1',
                        (month, '收入', abs(d))).fetchone()
                    if reg:
                        print('  PASS(已登记差异) %s %s 月度事实=%.6f 收入明细=%.6f 差=%.6f'
                              % (month, metric, fact, src, d))
                    else:
                        diffs += 1
                        print('  DIFF %s %s 月度事实=%.6f 收入明细=%.6f 差=%.6f'
                              % (month, metric, fact, src, d))
                    continue
            row = c.execute('SELECT 数值 FROM v_台账月度 WHERE 年份=? AND 月份=? AND 指标=?',
                            (y, m, metric)).fetchone()
            led = row[0] if row else 0.0
            if abs(fact - led) <= 0.000001:
                continue
            reg = c.execute(
                'SELECT 差异 FROM 差异登记 WHERE 月份=? AND 指标=? '
                'AND 状态 IN (\'已解释\',\'已归档\',\'已按源表更正\') '
                'AND ABS(ABS(差异)-?)<=0.005 LIMIT 1',
                (month, metric_map[metric], abs(fact - led))).fetchone()
            if reg:
                print('  PASS(已登记差异) %s %s 月度事实=%.6f 台账=%.6f 差=%.6f' % (month, metric, fact, led, fact - led))
            else:
                diffs += 1
                print('  DIFF %s %s 月度事实=%.6f 台账=%.6f 差=%.6f' % (month, metric, fact, led, fact - led))
    print('  %s（%d 项）' % ('PASS' if diffs == 0 else 'FAIL', checked))
    print('== 3. 对比视图覆盖 ==')
    for r in c.execute('SELECT 年份, COUNT(DISTINCT 月份), COUNT(*) FROM v_月度对比 GROUP BY 年份 ORDER BY 年份'):
        print('  %d 年：%d 个月，%d 条' % r)
    c.close()


def report(db=DB, out=None):
    c = sqlite3.connect(db)
    ensure_schema(c)
    rows = list(c.execute('SELECT 月份,指标,"2025","2026",差异,增幅 FROM v_月度同比'))
    c.close()
    latest = max([r[0] for r in rows if r[3] is not None], default=12)
    lines = []
    lines.append('# 东智 2025 / 2026 月度对比')
    lines.append('')
    lines.append('数据源：`历史月度`（2025 客户级明细）+ `月度事实`（2026 已关账）+ `台账明细`（当月实时）。单位：万元。')
    lines.append('')
    lines.append('> 2026 年 %d 月为台账实时数（未关账）；2026-10 收入/发货两处为 `差异登记` 已登记口径差异。' % latest)
    lines.append('')
    y25 = {m: 0.0 for m in METRICS}
    y25_same = {m: 0.0 for m in METRICS}
    y26 = {m: 0.0 for m in METRICS}
    for metric in METRICS:
        lines.append('## %s' % metric)
        lines.append('')
        lines.append('| 月份 | 2025 | 2026 | 差异 | 增幅 |')
        lines.append('|---:|---:|---:|---:|---:|')
        for m in range(1, 13):
            hit = [r for r in rows if r[0] == m and r[1] == metric]
            if not hit:
                lines.append('| %d月 | - | - | - | - |' % m)
                continue
            _, _, a, b, d, g = hit[0]
            a = a or 0.0
            y25[metric] += a
            if m <= latest:
                y25_same[metric] += a
            if b is not None:
                y26[metric] += b
            lines.append('| %d月 | %s | %s | %s | %s |' % (
                m, '%.6f' % a,
                '-' if b is None else '%.6f' % b,
                '-' if b is None else '%.6f' % d,
                '-' if g is None else '%.2f%%' % (g * 100)))
        d_same = y26[metric] - y25_same[metric]
        lines.append('| **2025 同期(1-%d月)** | **%.6f** |  |  |  |' % (latest, y25_same[metric]))
        lines.append('| **2026 累计(1-%d月)** |  | **%.6f** | **%.6f** | **%s** |' % (
            latest, y26[metric], d_same,
            '-' if y25_same[metric] == 0 else '%.2f%%' % (d_same / y25_same[metric] * 100)))
        lines.append('| 2025 全年 | %.6f |  |  |  |' % y25[metric])
        lines.append('')
    text = '\n'.join(lines)
    print(text)
    if out:
        with open(out, 'w', encoding='utf-8', newline='') as f:
            f.write(text + '\n')
        print('\n已写出 %s' % out)


def detail(name, year=2025, db=DB):
    """打印单个客户 12 个月 × 5 指标，用来核对“每个月每个客户都有”。"""
    c = sqlite3.connect(db)
    ensure_schema(c)
    exact = c.execute('SELECT DISTINCT 客户名称 FROM 历史月度 WHERE 年份=? AND 客户名称=?',
                      (year, name)).fetchone()
    if not exact:
        key = norm_name(name)
        cands = [r[0] for r in c.execute('SELECT DISTINCT 客户名称 FROM 历史月度 WHERE 年份=?', (year,))]
        cands = [x for x in cands if norm_name(x) == key]
        if not cands:
            print('未找到客户：%s（%d 年）' % (name, year))
            print('有数据的客户：' + '、'.join(sorted(set(
                r[0] for r in c.execute('SELECT DISTINCT 客户名称 FROM 历史月度 WHERE 年份=?', (year,))))))
            c.close()
            return
        name = cands[0]
    rows = c.execute('SELECT 月份,指标,数值 FROM 历史月度 WHERE 年份=? AND 客户名称=?',
                     (year, name)).fetchall()
    c.close()
    data = {(m, k): v for m, k, v in rows}
    print('%d 年 · %s（万元，空白=该月无发生）' % (year, name))
    print('| 月份 | ' + ' | '.join(METRICS) + ' |')
    print('|---:|' + '---:|' * len(METRICS))
    for m in range(1, 13):
        cells = ['%.6f' % data[(m, k)] if (m, k) in data else '' for k in METRICS]
        print('| %d月 | %s |' % (m, ' | '.join(cells)))
    tot = {k: sum(v for (mm, kk), v in data.items() if kk == k) for k in METRICS}
    print('| **合计** | ' + ' | '.join('**%.6f**' % tot[k] for k in METRICS) + ' |')


def main():
    ap = argparse.ArgumentParser(description='东智历史月度入库与同比对比')
    sub = ap.add_subparsers(dest='cmd')
    p1 = sub.add_parser('import')
    p1.add_argument('path')
    p1.add_argument('--year', type=int, default=2025)
    sub.add_parser('check')
    p3 = sub.add_parser('report')
    p3.add_argument('--out')
    p4 = sub.add_parser('detail')
    p4.add_argument('name')
    p4.add_argument('--year', type=int, default=2025)
    a = ap.parse_args()
    if a.cmd == 'import':
        import_year(a.path, a.year)
    elif a.cmd == 'check':
        check()
    elif a.cmd == 'report':
        report(out=a.out)
    elif a.cmd == 'detail':
        detail(a.name, a.year)
    else:
        ap.print_help()


if __name__ == '__main__':
    main()
