# -*- coding: utf-8 -*-
"""东智台账 · 月度快照与监督模块（v0.1）

设计口径（用户 2026-11-07 定）：
  - 以月份为分区；每月重新生成一次，覆盖该月旧快照（不是追加）。
  - 台账事实层保留历史；作废是状态（作废=1），统计一律排除，状态可更新。
  - 所有月度视图（F0/F2/F3/F4/PPT）从同一个「报告批次」快照出，星型，不互相读。
  - 本脚本提供：build（生成/覆盖月度快照）+ check（独立重算并对比，监督模块）。

用法：
  python dz_month.py build 2026-10
  python dz_month.py check 2026-10
  python dz_month.py build-canonical 2026-10
  python dz_month.py check-canonical 2026-10
  python dz_month.py list 2026-10
"""
import os, sys, csv, sqlite3, datetime, hashlib, shutil
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dz_store import DB_P, BAK_DIR  # noqa: E402

OUT_DIR = r'E:\商务部\东智日常表更新\监督'
METRICS = ['签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '开票金额', '回款金额', '收入']
TOL = 0.005  # 万元


def conn():
    return sqlite3.connect(DB_P)


def ensure_schema(cur):
    cur.executescript('''
    CREATE TABLE IF NOT EXISTS 报告批次(
      批次ID TEXT PRIMARY KEY,
      月份 TEXT NOT NULL,
      截止时间 TEXT,
      源文件哈希 TEXT,
      生成时间 TEXT,
      脚本版本 TEXT,
      状态 TEXT DEFAULT '有效'
    );
    CREATE TABLE IF NOT EXISTS 月度快照(
      批次ID TEXT NOT NULL,
      月份 TEXT NOT NULL,
      客户ID TEXT,
      客户名称 TEXT NOT NULL,
      指标 TEXT NOT NULL,
      数值 REAL,
      来源数 INTEGER,
      PRIMARY KEY(批次ID, 客户名称, 指标)
    );
    CREATE INDEX IF NOT EXISTS idx_月度快照_月份 ON 月度快照(月份);
    DROP VIEW IF EXISTS v_month_recon;
    CREATE VIEW v_month_recon AS
    SELECT r.*,
      CASE WHEN r.差异 IS NULL THEN NULL
           WHEN abs(r.差异) > 0.01 AND NOT EXISTS (
                SELECT 1 FROM 差异登记 d
                WHERE d.月份 = r.月份 AND d.指标 = r.指标
                  AND d.状态 IN ('已解释','已归档','已按源表更正')
                  AND abs(abs(d.差异) - abs(r.差异)) <= 0.01
           ) THEN '是' ELSE '否' END AS 重大差异,
      CASE WHEN EXISTS (
                SELECT 1 FROM 差异登记 d
                WHERE d.月份 = r.月份 AND d.指标 = r.指标
                  AND d.状态 IN ('已解释','已归档','已按源表更正')
                  AND abs(abs(d.差异) - abs(r.差异)) <= 0.01
           ) THEN 1 ELSE 0 END AS 已登记
    FROM (
      SELECT c.月份, c.指标, c.数值 AS 控制值,
        round(CASE c.指标
          WHEN '签约金额' THEN a.签约金额 WHEN '签约数量' THEN a.签约数量
          WHEN '产出数量' THEN a.产出数量 WHEN '发货数量' THEN a.发货数量
          WHEN '发货金额' THEN a.发货金额 WHEN '收入' THEN a.收入
          WHEN '开票金额' THEN a.开票金额 WHEN '回款金额' THEN a.回款金额 END, 6) AS 事实值,
        round((CASE c.指标
          WHEN '签约金额' THEN a.签约金额 WHEN '签约数量' THEN a.签约数量
          WHEN '产出数量' THEN a.产出数量 WHEN '发货数量' THEN a.发货数量
          WHEN '发货金额' THEN a.发货金额 WHEN '收入' THEN a.收入
          WHEN '开票金额' THEN a.开票金额 WHEN '回款金额' THEN a.回款金额 END) - c.数值, 6) AS 差异
      FROM 月度控制 c JOIN v_month_canonical a ON a.月份 = c.月份
      WHERE c.版本 = '最新'
    ) r;
    ''')
    sc = [r[1] for r in cur.execute('PRAGMA table_info(月度快照)').fetchall()]
    if '客户ID' not in sc:
        cur.execute('ALTER TABLE 月度快照 ADD COLUMN 客户ID TEXT')


def _month_rows(cur, month):
    q = '''SELECT 客户ID, 客户名称, 来源ID, COALESCE(签约金额,0), COALESCE(签约数量,0),
                  COALESCE(产出数量,0), COALESCE(发货数量,0), COALESCE(发货金额,0),
                  COALESCE(开票金额,0), COALESCE(回款金额,0), COALESCE(收入,0)
           FROM 台账明细
           WHERE COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?
             AND COALESCE(作废,0)=0 AND COALESCE(客户名称,'')<>'' '''
    return cur.execute(q, (month,)).fetchall()


def _canonical_rows(cur, month):
    q = '''SELECT 客户ID, 客户名称, COALESCE(签约金额,0), COALESCE(签约数量,0),
                  COALESCE(产出数量,0), COALESCE(发货数量,0), COALESCE(发货金额,0),
                  COALESCE(开票金额,0), COALESCE(回款金额,0), COALESCE(收入,0)
           FROM 月度事实
           WHERE 月份=? AND 版本='最新' AND COALESCE(客户名称,'')<>'' '''
    return cur.execute(q, (month,)).fetchall()


def build(month):
    c = conn(); cur = c.cursor(); ensure_schema(cur)
    rows = _month_rows(cur, month)
    if not rows:
        print('! %s 没有有效台账行' % month); c.close(); return 1

    agg = defaultdict(lambda: defaultdict(float))
    cnt = defaultdict(int)
    cid_of = {}
    ids = []
    for r in rows:
        cid, cust = r[0], r[1]
        cid_of[cust] = cid
        ids.append(str(r[2] or ''))
        cnt[cust] += 1
        for i, m in enumerate(METRICS):
            agg[cust][m] += float(r[3 + i] or 0)

    batch = datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + month
    h = hashlib.sha256('|'.join(sorted(ids)).encode('utf-8')).hexdigest()
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    bak = os.path.join(BAK_DIR, 'month_build_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
    os.makedirs(bak, exist_ok=True)
    shutil.copy(DB_P, os.path.join(bak, os.path.basename(DB_P)))

    cur.execute("UPDATE 报告批次 SET 状态='已归档' WHERE 月份=? AND 状态<>'已归档'", (month,))
    cur.execute("DELETE FROM 月度快照 WHERE 月份=?", (month,))
    cur.execute("INSERT INTO 报告批次(批次ID,月份,截止时间,源文件哈希,生成时间,脚本版本,状态) VALUES(?,?,?,?,?,?,?)",
                (batch, month, ts, h, ts, 'dz_month v0.1', '有效'))
    for cust, m in agg.items():
        for metric in METRICS:
            cur.execute("INSERT INTO 月度快照(批次ID,月份,客户ID,客户名称,指标,数值,来源数) VALUES(?,?,?,?,?,?,?)",
                        (batch, month, cid_of.get(cust), cust, metric, round(m[metric], 10), cnt[cust]))
    c.commit()
    c.close()
    print('OK 月份=%s 批次=%s' % (month, batch))
    print('   客户=%d 台账行=%d 源哈希=%s' % (len(agg), len(rows), h[:16]))
    print('   备份=%s' % bak)
    return 0


def build_canonical(month):
    c = conn(); cur = c.cursor(); ensure_schema(cur)
    rows = _canonical_rows(cur, month)
    if not rows:
        print('! %s 没有 canonical 月度事实行' % month); c.close(); return 1

    agg = defaultdict(lambda: defaultdict(float))
    cnt = defaultdict(int)
    cid_of = {}
    for r in rows:
        cid, cust = r[0], r[1]
        cid_of[cust] = cid
        cnt[cust] += 1
        for i, m in enumerate(METRICS):
            agg[cust][m] += float(r[2 + i] or 0)

    batch = 'CANON-' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + month
    h = hashlib.sha256('|'.join('%s:%.10f' % (c, agg[c][m]) for c in sorted(agg) for m in METRICS).encode('utf-8')).hexdigest()
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    bak = os.path.join(BAK_DIR, 'month_canonical_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
    os.makedirs(bak, exist_ok=True)
    shutil.copy(DB_P, os.path.join(bak, os.path.basename(DB_P)))

    cur.execute("UPDATE 报告批次 SET 状态='已归档' WHERE 月份=? AND 状态<>'已归档'", (month,))
    cur.execute("DELETE FROM 月度快照 WHERE 月份=?", (month,))
    cur.execute("INSERT INTO 报告批次(批次ID,月份,截止时间,源文件哈希,生成时间,脚本版本,状态) VALUES(?,?,?,?,?,?,?)",
                (batch, month, ts, h, ts, 'dz_month canonical v0.2', '有效'))
    for cust, m in agg.items():
        for metric in METRICS:
            cur.execute("INSERT INTO 月度快照(批次ID,月份,客户ID,客户名称,指标,数值,来源数) VALUES(?,?,?,?,?,?,?)",
                        (batch, month, cid_of.get(cust), cust, metric, round(m[metric], 10), cnt[cust]))
    c.commit()
    c.close()
    print('OK canonical 月份=%s 批次=%s' % (month, batch))
    print('   客户=%d 事实行=%d 源哈希=%s' % (len(agg), len(rows), h[:16]))
    print('   备份=%s' % bak)
    return 0


def _independent_agg(cur, month):
    q = '''SELECT 客户名称,
                  SUM(COALESCE(签约金额,0)), SUM(COALESCE(签约数量,0)),
                  SUM(COALESCE(产出数量,0)), SUM(COALESCE(发货数量,0)),
                  SUM(COALESCE(发货金额,0)), SUM(COALESCE(开票金额,0)),
                  SUM(COALESCE(回款金额,0)), SUM(COALESCE(收入,0)), COUNT(*)
           FROM 台账明细
           WHERE COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?
             AND COALESCE(作废,0)=0 AND COALESCE(客户名称,'')<>''
           GROUP BY 客户名称'''
    out = {}
    for r in cur.execute(q, (month,)):
        out[r[0]] = {METRICS[i]: float(r[1 + i] or 0) for i in range(8)}
        out[r[0]]['_count'] = int(r[9])
    return out


def check(month):
    c = conn(); cur = c.cursor(); ensure_schema(cur)
    b = cur.execute("SELECT 批次ID,截止时间,源文件哈希,生成时间 FROM 报告批次 WHERE 月份=? AND 状态='有效' ORDER BY 生成时间 DESC LIMIT 1", (month,)).fetchone()
    if not b:
        print('! %s 没有有效报告批次，先跑 build' % month); c.close(); return 1
    batch, cutoff, src_hash, gen_at = b
    snap = {}
    for cust, metric, val, n in cur.execute("SELECT 客户名称,指标,数值,来源数 FROM 月度快照 WHERE 批次ID=?", (batch,)):
        snap.setdefault(cust, {})[metric] = float(val or 0)
        snap[cust]['_count'] = int(n or 0)
    ind = _independent_agg(cur, month)
    c.close()

    diffs = []
    for cust in sorted(set(snap) | set(ind)):
        s = snap.get(cust, {}); i = ind.get(cust, {})
        for m in METRICS:
            sv = float(s.get(m, 0)); iv = float(i.get(m, 0))
            if abs(sv - iv) > TOL:
                diffs.append((cust, m, sv, iv, sv - iv))
        sc, ic = int(s.get('_count', 0)), int(i.get('_count', 0))
        if sc != ic:
            diffs.append((cust, '来源行数', sc, ic, sc - ic))

    os.makedirs(OUT_DIR, exist_ok=True)
    md = os.path.join(OUT_DIR, '%s_监督.md' % month)
    csvp = os.path.join(OUT_DIR, '%s_差异.csv' % month)
    totals = {}
    for m in METRICS:
        totals[m] = (sum(v.get(m, 0) for v in snap.values()), sum(v.get(m, 0) for v in ind.values()))
    with open(md, 'w', encoding='utf-8') as f:
        f.write('# %s 月度监督报告\n\n' % month)
        f.write('- 批次：`%s`\n- 生成：%s\n- 截止：%s\n- 源哈希：`%s`\n' % (batch, gen_at, cutoff, src_hash))
        f.write('- 客户数：快照 %d / 独立重算 %d\n\n' % (len(snap), len(ind)))
        f.write('## 合计对比（万元）\n\n| 指标 | 快照 | 独立重算 | 差异 |\n|---|---:|---:|---:|\n')
        for m in METRICS:
            a, b2 = totals[m]
            f.write('| %s | %.6f | %.6f | %.6f |\n' % (m, a, b2, a - b2))
        f.write('\n## 结果\n\n')
        if diffs:
            f.write('❌ 发现 %d 项差异（容差 %.3f 万元），明细见 `%s`\n' % (len(diffs), TOL, os.path.basename(csvp)))
        else:
            f.write('✅ 全部一致（%d 个客户 × %d 指标，容差 %.3f 万元）\n' % (len(snap), len(METRICS), TOL))
    with open(csvp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['客户名称', '指标', '快照值', '独立重算值', '差异'])
        w.writerows(diffs)
    print(('OK' if not diffs else 'FAIL') + ' %s：%d 项差异' % (month, len(diffs)))
    print('   报告=%s' % md)
    print('   差异=%s' % csvp)
    for d in diffs[:10]:
        print('   - %s / %s: %.6f vs %.6f (%.6f)' % d)
    return 0 if not diffs else 1


def check_canonical(month):
    c = conn(); cur = c.cursor(); ensure_schema(cur)
    b = cur.execute("SELECT 批次ID,截止时间,源文件哈希,生成时间 FROM 报告批次 WHERE 月份=? AND 状态='有效' ORDER BY 生成时间 DESC LIMIT 1", (month,)).fetchone()
    if not b:
        print('! %s 没有有效报告批次，先跑 build-canonical' % month); c.close(); return 1
    batch, cutoff, src_hash, gen_at = b
    snap = {}
    for cust, metric, val, n in cur.execute("SELECT 客户名称,指标,数值,来源数 FROM 月度快照 WHERE 批次ID=?", (batch,)):
        snap.setdefault(cust, {})[metric] = float(val or 0)
        snap[cust]['_count'] = int(n or 0)

    ind = defaultdict(lambda: defaultdict(float))
    cnt = defaultdict(int)
    for r in _canonical_rows(cur, month):
        cust = r[1]; cnt[cust] += 1
        for i, m in enumerate(METRICS):
            ind[cust][m] += float(r[2 + i] or 0)
    c.close()

    diffs = []
    for cust in sorted(set(snap) | set(ind)):
        s = snap.get(cust, {}); i = ind.get(cust, {})
        for m in METRICS:
            sv = float(s.get(m, 0)); iv = float(i.get(m, 0))
            if abs(sv - iv) > TOL:
                diffs.append((cust, m, sv, iv, sv - iv))
        sc, ic = int(s.get('_count', 0)), int(cnt.get(cust, 0))
        if sc != ic:
            diffs.append((cust, '来源行数', sc, ic, sc - ic))

    c = conn(); cur = c.cursor()
    ctrl = list(cur.execute("SELECT 指标,控制值,事实值,差异,重大差异 FROM v_month_recon WHERE 月份=? AND 重大差异='是' ORDER BY 指标", (month,)))
    raw = list(cur.execute("SELECT * FROM v_month_raw_vs_canonical WHERE 月份=?", (month,)))
    c.close()

    os.makedirs(OUT_DIR, exist_ok=True)
    md = os.path.join(OUT_DIR, '%s_监督_规范.md' % month)
    csvp = os.path.join(OUT_DIR, '%s_差异_规范.csv' % month)
    totals = {m: (sum(v.get(m, 0) for v in snap.values()), sum(v.get(m, 0) for v in ind.values())) for m in METRICS}
    with open(md, 'w', encoding='utf-8') as f:
        f.write('# %s 规范口径监督报告\n\n' % month)
        f.write('- 批次：`%s`\n- 生成：%s\n- 截止：%s\n- 源哈希：`%s`\n' % (batch, gen_at, cutoff, src_hash))
        f.write('- 客户数：快照 %d / 月度事实 %d\n\n' % (len(snap), len(ind)))
        f.write('## 快照 vs 月度事实（万元）\n\n| 指标 | 快照 | 月度事实 | 差异 |\n|---|---:|---:|---:|\n')
        for m in METRICS:
            a, b2 = totals[m]
            f.write('| %s | %.6f | %.6f | %.6f |\n' % (m, a, b2, a - b2))
        f.write('\n## 控制数重大差异（>100 元）\n\n')
        if ctrl:
            f.write('| 指标 | 控制值 | 事实值 | 差异 |\n|---|---:|---:|---:|\n')
            for r in ctrl:
                f.write('| %s | %.6f | %.6f | %.6f |\n' % (r[0], r[1], r[2], r[3]))
        else:
            f.write('无。\n')
        f.write('\n## 台账原始口径差异（仅供解释）\n\n')
        if raw:
            for r in raw:
                f.write('- 签约 %.6f / 产出 %.2f / 发货量 %.2f / 发货额 %.6f / 收入 %.6f / 开票 %.6f / 回款 %.6f；%s\n'
                        % (r[1] or 0, r[3] or 0, r[4] or 0, r[5] or 0, r[6] or 0, r[7] or 0, r[8] or 0, r[9] or ''))
        else:
            f.write('无。\n')
        f.write('\n## 结果\n\n')
        f.write(('✅ 快照与月度事实一致，控制数无重大差异。\n' if not diffs and not ctrl else '❌ 发现差异，见 `%s`。\n' % os.path.basename(csvp)))
    with open(csvp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['客户名称', '指标', '快照值', '月度事实值', '差异'])
        w.writerows(diffs)
    ok = not diffs and not ctrl
    print(('OK' if ok else 'FAIL') + ' %s canonical：快照差异 %d 项，控制重大差异 %d 项' % (month, len(diffs), len(ctrl)))
    print('   报告=%s' % md)
    return 0 if ok else 1


def list_batches(month):
    c = conn(); cur = c.cursor(); ensure_schema(cur)
    for r in cur.execute("SELECT 批次ID,月份,截止时间,源文件哈希,生成时间,状态 FROM 报告批次 WHERE 月份=? ORDER BY 生成时间 DESC", (month,)):
        print(r)
    print('快照行数：', cur.execute("SELECT COUNT(*) FROM 月度快照 WHERE 月份=?", (month,)).fetchone()[0])
    c.close()


def main():
    a = sys.argv[1:]
    if len(a) < 2:
        print(__doc__); return
    cmd, month = a[0], a[1]
    if cmd == 'build':
        sys.exit(build(month))
    elif cmd == 'check':
        sys.exit(check(month))
    elif cmd == 'build-canonical':
        sys.exit(build_canonical(month))
    elif cmd == 'check-canonical':
        sys.exit(check_canonical(month))
    elif cmd == 'list':
        list_batches(month)
    else:
        print('未知命令：%s' % cmd); print(__doc__)


if __name__ == '__main__':
    main()
