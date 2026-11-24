# -*- coding: utf-8 -*-
"""东智台账 · 结构迁移（v1.0，2026-11-23）

一次性/幂等：给「东智台账.db」补上语义层基础设施。
  - 客户主数据（客户ID / 标准名 / 别名 / 状态）
  - 期间归属规则（源 → 归属字段 → 落账日规则）
  - 台账明细 增加 客户ID / 业务日期 / 归属月 / 归属依据
  - 批次清单 / 校验结果 / 差异登记 / 月度计划 / 批次水印
  - 视图：v_customer_dim / v_month_actual_by_period / v_plan_vs_actual

用法：
  python dz_schema.py                 # 迁移 + 回填
  python dz_schema.py --f0 <F0.xlsx>  # 同时把 F0 里的客户并入主数据
"""
import os
import sys
import re
import json
import glob
import shutil
import sqlite3
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dz_common import RULES  # noqa: E402

DB = r'E:\商务部\东智日常表更新\东智台账.db'
BAK_DIR = r'E:\商务部\东智日常表更新\_backup'
DEFAULT_F0 = r'E:\商务部\东智日常表更新\20261114-20261120\截止到2026年11月20日特发东智订单、产出、发货、开票及回款汇总表.xlsx'

DDL = [
    '''CREATE TABLE IF NOT EXISTS 客户主数据(
        客户ID TEXT PRIMARY KEY,
        标准名 TEXT NOT NULL UNIQUE,
        别名 TEXT DEFAULT '',
        类别 TEXT DEFAULT '',
        客户状态 TEXT DEFAULT '在用',
        生效月 TEXT DEFAULT '2026-09',
        备注 TEXT DEFAULT '',
        录入时间 TEXT DEFAULT (datetime('now','localtime'))
    )''',
    '''CREATE TABLE IF NOT EXISTS 期间归属规则(
        源 TEXT PRIMARY KEY,
        归属字段 TEXT NOT NULL,
        归属月规则 TEXT NOT NULL,
        落账日规则 TEXT NOT NULL,
        说明 TEXT,
        生效月 TEXT DEFAULT '2026-09',
        更新时间 TEXT DEFAULT (datetime('now','localtime'))
    )''',
    '''CREATE TABLE IF NOT EXISTS 批次清单(
        批次ID TEXT NOT NULL,
        文件类型 TEXT NOT NULL,
        文件名 TEXT NOT NULL,
        路径 TEXT,
        字节数 INTEGER,
        修改时间 TEXT,
        sha256 TEXT,
        归档时间 TEXT,
        PRIMARY KEY(批次ID, 文件类型, 文件名)
    )''',
    '''CREATE TABLE IF NOT EXISTS 校验结果(
        批次ID TEXT NOT NULL,
        月份 TEXT NOT NULL,
        环节 TEXT NOT NULL,
        检查项 TEXT NOT NULL,
        源值 REAL,
        台账值 REAL,
        看板值 REAL,
        差异 REAL,
        容差 REAL,
        结论 TEXT NOT NULL,
        详情 TEXT,
        生成时间 TEXT DEFAULT (datetime('now','localtime')),
        PRIMARY KEY(批次ID, 环节, 检查项)
    )''',
    '''CREATE TABLE IF NOT EXISTS 差异登记(
        登记ID INTEGER PRIMARY KEY AUTOINCREMENT,
        发现日期 TEXT NOT NULL,
        月份 TEXT NOT NULL,
        指标 TEXT NOT NULL,
        我方值 REAL,
        对方值 REAL,
        差异 REAL,
        对方 TEXT,
        责任方 TEXT,
        原因 TEXT,
        处理 TEXT,
        状态 TEXT DEFAULT '已解释',
        来源 TEXT,
        录入时间 TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(月份, 指标, 对方, 原因)
    )''',
    '''CREATE TABLE IF NOT EXISTS 月度计划(
        月份 TEXT NOT NULL,
        客户ID TEXT NOT NULL,
        指标 TEXT NOT NULL,
        计划值 REAL NOT NULL,
        来源 TEXT,
        备注 TEXT,
        录入时间 TEXT DEFAULT (datetime('now','localtime')),
        PRIMARY KEY(月份, 客户ID, 指标)
    )''',
    '''CREATE TABLE IF NOT EXISTS 批次水印(
        对象 TEXT NOT NULL,
        批次ID TEXT NOT NULL,
        水印 TEXT NOT NULL,
        数据截止 TEXT,
        生成时间 TEXT DEFAULT (datetime('now','localtime')),
        说明 TEXT,
        PRIMARY KEY(对象, 批次ID)
    )''',
]

RULES_ROWS = [
    ('①合约订单评审表', '制单日期(D列)', '业务日期所在月',
     '落账日=制单日期', '签约按 OA 制单日期归属；每周整月重取回补，靠来源ID幂等去重', '2026-09'),
    ('②订单交付跟踪表', '快照日', '快照日所在月',
     '落账日=快照日', '产出源是当月累计，按相邻快照相减得每日增量', '2026-09'),
    ('③销售出库单明细表', '单据日期(billdate)', '业务日期所在月',
     '落账日=单据日期', '发货按 NC 单据日期归属；仓库非成品仓只计金额不计数量', '2026-09'),
    ('④销售发票登记明细', '开票日期(H列)', '业务日期所在月',
     '落账日=开票日期', '开票按发票日期归属', '2026-09'),
    ('④现金回款明细', '月份列(K列)', '月份列优先；空则收款日期',
     '收款日在归属月内=收款日；跨月=归属月1日', '2026-11 玖耀/宝鑫两笔即此规则：收款日期在9/10月，但登记表月份=11', '2026-09'),
    ('④应收票据', '月份列(K列)', '月份列优先；空则收款日期',
     '收款日在归属月内=收款日；跨月=归属月1日', '与现金回款同口径', '2026-09'),
]


def backup():
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    bak = os.path.join(BAK_DIR, 'schema_' + ts)
    os.makedirs(bak, exist_ok=True)
    if os.path.exists(DB):
        shutil.copy(DB, os.path.join(bak, os.path.basename(DB)))
    print('已备份 -> %s' % bak)
    return bak


def cols(cur, table):
    return [r[1] for r in cur.execute('PRAGMA table_info("%s")' % table)]


def add_col(cur, table, name, typ):
    if name not in cols(cur, table):
        cur.execute('ALTER TABLE "%s" ADD COLUMN "%s" %s' % (table, name, typ))
        print('  + %s.%s' % (table, name))


def f0_customers(path):
    """F0 分项汇总里的客户行（去掉合计/统计/表头）"""
    if not path or not os.path.exists(path):
        return []
    from openpyxl import load_workbook
    out = []
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb['分项汇总']
    for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
        a = str(row[0] or '').strip()
        if not a or a == '客户':
            continue
        if '合计' in a or '统计' in a:
            continue
        if re.match(r'^[1-9]\s*[、.]', a):
            continue
        out.append(a)
    wb.close()
    return out


def norm_std(name):
    s = str(name or '').strip()
    for suf in ('（物料售卖）', '(物料售卖)', '（自主接单）', '(自主接单)'):
        if s.endswith(suf):
            s = s[:-len(suf)].strip()
    return s


def build_aliases(std, all_names):
    al = set()
    for k, v in RULES.get('ALIAS', {}).items():
        if v == std:
            al.add(k)
    for kw, tgt in RULES.get('SUBSTR_RULES', []):
        if tgt == std:
            al.add(kw)
    for n in all_names:
        if n != std and norm_std(n) == std:
            al.add(n)
    al.discard(std)
    return '；'.join(sorted(al))


def main():
    f0 = None
    if '--f0' in sys.argv:
        f0 = sys.argv[sys.argv.index('--f0') + 1]
    elif os.path.exists(DEFAULT_F0):
        f0 = DEFAULT_F0

    backup()
    c = sqlite3.connect(DB)
    cur = c.cursor()

    print('== 建表 ==')
    for ddl in DDL:
        cur.execute(ddl)
    print('  报告批次/月度快照/月度事实/月度控制/月度版本/月度数量实际：已存在则沿用')

    print('== 台账明细 加列 ==')
    add_col(cur, '台账明细', '客户ID', 'TEXT')
    add_col(cur, '台账明细', '业务日期', 'TEXT')
    add_col(cur, '台账明细', '归属月', 'TEXT')
    add_col(cur, '台账明细', '归属依据', 'TEXT')
    add_col(cur, '月度事实', '客户ID', 'TEXT')
    add_col(cur, '月度快照', '客户ID', 'TEXT')

    print('== 期间归属规则 ==')
    for r in RULES_ROWS:
        cur.execute('''INSERT INTO 期间归属规则(源,归属字段,归属月规则,落账日规则,说明,生效月)
                       VALUES(?,?,?,?,?,?)
                       ON CONFLICT(源) DO UPDATE SET 归属字段=excluded.归属字段,
                         归属月规则=excluded.归属月规则, 落账日规则=excluded.落账日规则,
                         说明=excluded.说明, 更新时间=datetime('now','localtime')''', r)
        print('  · %s' % r[0])

    print('== 客户主数据 ==')
    active = [r[0] for r in cur.execute(
        "SELECT DISTINCT 客户名称 FROM 台账明细 WHERE COALESCE(客户名称,'')<>''")]
    fact = [r[0] for r in cur.execute(
        "SELECT DISTINCT 客户名称 FROM 月度事实 WHERE COALESCE(客户名称,'')<>''")]
    active = sorted(set(active) | set(fact))
    f0n = [norm_std(n) for n in f0_customers(f0)] if f0 else []
    all_names = sorted(set(active) | set(f0n))

    existing = {r[0]: r[1] for r in cur.execute('SELECT 标准名, 客户ID FROM 客户主数据')}
    next_id = 1
    if existing:
        nums = [int(x[1:]) for x in existing.values() if re.match(r'^C\d+$', x)]
        next_id = (max(nums) + 1) if nums else 1

    # 先活跃客户（在用），再仅计划客户（仅计划）
    ordered = [(n, '在用') for n in active] + [(n, '仅计划') for n in f0n if n not in active]
    seen = set()
    added = 0
    for std, status in ordered:
        if std in seen:
            continue
        seen.add(std)
        if std in existing:
            continue
        cid = 'C%03d' % next_id
        next_id += 1
        cat = ''
        row = cur.execute("SELECT 类别 FROM 台账明细 WHERE 客户名称=? AND COALESCE(类别,'')<>'' LIMIT 1", (std,)).fetchone()
        if row:
            cat = row[0]
        note = 'F0 有行、台账无数据' if status == '仅计划' else ''
        cur.execute('''INSERT INTO 客户主数据(客户ID,标准名,别名,类别,客户状态,生效月,备注)
                       VALUES(?,?,?,?,?,?,?)''',
                    (cid, std, build_aliases(std, all_names), cat, status, '2026-09', note))
        existing[std] = cid
        added += 1
    print('  客户 %d 家（新增 %d）' % (len(existing), added))

    # 清理：F0 表头误当客户；新基建事业部 合并到 新基建
    bad = [r[0] for r in cur.execute("SELECT 客户ID FROM 客户主数据 WHERE 标准名 LIKE '%汇总表%'")]
    for cid in bad:
        cur.execute('DELETE FROM 客户主数据 WHERE 客户ID=?', (cid,))
        print('  - 清理非客户行 %s' % cid)
    dup = cur.execute("SELECT 客户ID FROM 客户主数据 WHERE 标准名='新基建事业部'").fetchone()
    std = cur.execute("SELECT 客户ID FROM 客户主数据 WHERE 标准名='新基建'").fetchone()
    if dup and std:
        cur.execute("UPDATE 台账明细 SET 客户名称='新基建', 客户ID=? WHERE 客户名称='新基建事业部'", (std[0],))
        cur.execute("UPDATE 月度事实 SET 客户名称='新基建', 客户ID=? WHERE 客户名称='新基建事业部'", (std[0],))
        cur.execute("UPDATE 月度快照 SET 客户名称='新基建', 客户ID=? WHERE 客户名称='新基建事业部'", (std[0],))
        cur.execute("DELETE FROM 客户主数据 WHERE 客户ID=?", (dup[0],))
        cur.execute("UPDATE 客户主数据 SET 别名=trim(COALESCE(别名,'') || '；新基建事业部','；') WHERE 客户ID=?", (std[0],))
        print('  · 合并 新基建事业部 -> 新基建')
    cur.execute("UPDATE 客户主数据 SET 别名=replace(别名,'洁字再生资源','洁宇再生资源')")

    # 特殊桶
    for std in ('其他', '其他开票收入'):
        if std in existing:
            cur.execute("UPDATE 客户主数据 SET 客户状态='特殊行' WHERE 标准名=?", (std,))

    print('== 回填 台账明细 ==')
    n_id = n_bd = n_am = 0
    for _id, cust, t in cur.execute('SELECT id, 客户名称, 时间 FROM 台账明细').fetchall():
        cid = existing.get(cust)
        if not cid:
            cid = 'C%03d' % next_id
            next_id += 1
            cur.execute('''INSERT INTO 客户主数据(客户ID,标准名,别名,类别,客户状态,生效月,备注)
                           VALUES(?,?,?,?,?,?,?)''', (cid, cust, '', '', '待确认', '2026-09', '回填时发现的新客户'))
            existing[cust] = cid
        cur.execute('UPDATE 台账明细 SET 客户ID=? WHERE id=? AND (客户ID IS NULL OR 客户ID=?)', (cid, _id, ''))
        if cur.rowcount:
            n_id += 1
        bd = (t or '')[:10]
        am = (t or '')[:7]
        cur.execute('''UPDATE 台账明细 SET 业务日期=COALESCE(NULLIF(业务日期,''),?),
                       归属月=COALESCE(NULLIF(归属月,''),?),
                       归属依据=COALESCE(NULLIF(归属依据,''),?)
                       WHERE id=?''', (bd, am, '历史回填：归属月=时间列', _id))
    print('  补客户ID %d 行；业务日期/归属月 已回填' % n_id)

    # 两笔跨月回款/票据：来源ID 改成规范键，保证下次整月重跑幂等命中
    fixes = [
        ('PAY:M11#宝鑫再生资源#50000#841',
         'PAY:2026-09-17#惠州市宝鑫再生资源回收有限公司#50000#841',
         '2026-09-17', '2026-11-01',
         '登记表月份=11；收款日期2026-09-17跨月→落账11-01'),
        ('PAY:M11#鑫合盛#37835.26#295',
         'PAY:2026-10-11#深圳鑫合盛电子有限公司#37835.26#295',
         '2026-10-11', '2026-11-01',
         '登记表月份=11；收款日期2026-10-11跨月→落账11-01'),
    ]
    src = os.path.join(r'E:\商务部\东智日常表更新\20261114-20261120',
                       '2026年11月-2026年12月东智销售发票及回款登记表.xlsx')
    for old, new, bd, t, basis in fixes:
        row = cur.execute('SELECT id FROM 台账明细 WHERE 来源ID=?', (old,)).fetchone()
        if row:
            cur.execute('''UPDATE 台账明细 SET 来源ID=?,信息来源路径=?,业务日期=?,时间=?,
                           归属月='2026-11',归属依据=?
                           WHERE id=?''', (new, src, bd, t, basis, row[0]))
            print('  · 修正跨月行 %s -> %s' % (old, new))

    print('== 回填 月度事实/月度快照 客户ID ==')
    for tbl, key in (('月度事实', '客户名称'), ('月度快照', '客户名称')):
        for std, cid in existing.items():
            cur.execute('UPDATE "%s" SET 客户ID=? WHERE %s=? AND (客户ID IS NULL OR 客户ID=?)' % (tbl, key),
                        (cid, std, ''))

    print('== 差异登记：迁移审计差异 + 补 3 条已知差异 ==')
    n_seed = 0
    audit_rows = cur.execute('''SELECT 审计日期,月份,指标,原值,新值,差异,责任方,原因,处理,范围
                                 FROM 审计差异''').fetchall()
    for r in audit_rows:
        try:
            cur.execute('''INSERT OR IGNORE INTO 差异登记
                (发现日期,月份,指标,我方值,对方值,差异,对方,责任方,原因,处理,状态,来源)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                (r[0], r[1], r[2], r[3], r[4], r[5], r[9] or '历史审计', r[6], r[7], r[8],
                 '已归档', '审计差异迁移'))
            n_seed += cur.rowcount
        except Exception:
            pass
    known = [
        ('2026-11-23', '2026-11', '回款金额', 1825.580415, 1823.761345, 1.81907,
         '官方F0汇总页', '官方F0', '玖耀启铭11/14回款18,190.7在汇总页客户行未挂公式，整笔漏计',
         '以我方台账为准；官方下月整改', '已解释', '20261120周更复盘'),
        ('2026-11-23', '2026-10', '发货金额', 5880.969062, 5880.980762, -0.0117,
         'F0日矩阵', '我方F0', 'F0日矩阵取整，鑫合盛少117元', '差异0.0002%，保留说明', '已解释', '10月月结复盘'),
        ('2026-11-23', '2026-09', '回款金额', 2979.2566, 2979.2567, -0.0001,
         'F3', 'F3', 'F2四位小数合计2979.2567，F3填2979.2566，差1元', '以F2四位小数合计为准', '已解释', '9月月结复盘'),
        ('2026-11-23', '2026-09', '签约金额', 1528.430713, 1629.590833, 101.16012,
         'OA源', '双方', 'OA周批未补T01 193.66、中兴30.12等；事实层含手工补录', '台账明细仅作原始痕迹，月度事实已固化', '已解释', '9月月结复盘'),
    ]
    for r in known:
        cur.execute('''INSERT OR IGNORE INTO 差异登记
            (发现日期,月份,指标,我方值,对方值,差异,对方,责任方,原因,处理,状态,来源)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''', r)
        n_seed += cur.rowcount
    print('  差异登记新增 %d 条' % n_seed)

    print('== 视图 ==')
    cur.executescript('''
    DROP VIEW IF EXISTS v_customer_dim;
    CREATE VIEW v_customer_dim AS
      SELECT 客户ID, 标准名, 别名, 类别, 客户状态, 生效月 FROM 客户主数据;

    DROP VIEW IF EXISTS v_month_actual_by_period;
    CREATE VIEW v_month_actual_by_period AS
      SELECT COALESCE(NULLIF(归属月,''), substr(时间,1,7)) AS 月份,
             客户ID, 客户名称,
             round(sum(COALESCE(签约金额,0)),6) AS 签约金额,
             round(sum(COALESCE(签约数量,0)),2) AS 签约数量,
             round(sum(COALESCE(产出数量,0)),2)  AS 产出数量,
             round(sum(COALESCE(发货数量,0)),2)  AS 发货数量,
             round(sum(COALESCE(发货金额,0)),6)  AS 发货金额,
             round(sum(COALESCE(收入,0)),6)      AS 收入,
             round(sum(COALESCE(开票金额,0)),6)  AS 开票金额,
             round(sum(COALESCE(回款金额,0)),6)  AS 回款金额
      FROM 台账明细
      WHERE COALESCE(作废,0)=0 AND COALESCE(客户名称,'')<>''
      GROUP BY 1,2,3;

    DROP VIEW IF EXISTS v_plan_vs_actual;
    CREATE VIEW v_plan_vs_actual AS
      SELECT p.月份, p.客户ID, m.标准名, p.指标, p.计划值,
             COALESCE(s.实际值,0) AS 实际值,
             round(COALESCE(s.实际值,0)-p.计划值,6) AS 差异,
             round(CASE WHEN p.计划值<>0 THEN COALESCE(s.实际值,0)/p.计划值 ELSE NULL END,4) AS 完成率,
             p.来源
      FROM 月度计划 p
      LEFT JOIN 客户主数据 m ON m.客户ID=p.客户ID
      LEFT JOIN (
        SELECT 月份, 客户ID, 指标, sum(数值) AS 实际值
        FROM 月度快照 s
        WHERE 批次ID=(SELECT 批次ID FROM 报告批次 b
                      WHERE b.月份=s.月份 AND b.状态='有效'
                      ORDER BY 生成时间 DESC LIMIT 1)
        GROUP BY 月份, 客户ID, 指标
      ) s ON s.月份=p.月份 AND s.客户ID=p.客户ID AND s.指标=p.指标;

    DROP VIEW IF EXISTS v_month_snapshot_totals;
    CREATE VIEW v_month_snapshot_totals AS
      SELECT 月份, 指标, round(sum(数值),6) AS 合计, count(DISTINCT 客户名称) AS 客户数
      FROM 月度快照
      WHERE 批次ID=(SELECT 批次ID FROM 报告批次 b
                    WHERE b.月份=月度快照.月份 AND b.状态='有效'
                    ORDER BY 生成时间 DESC LIMIT 1)
      GROUP BY 月份, 指标;
    ''')

    # 历史月度（2025 等）与同比视图；DDL 只在 dz_history.py 维护一份
    from dz_history import ensure_schema as ensure_history_schema
    ensure_history_schema(c)

    cur.execute('CREATE INDEX IF NOT EXISTS idx_led_归属月 ON 台账明细(归属月)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_led_客户ID ON 台账明细(客户ID)')
    c.commit()

    print('== 结果 ==')
    for t in ('客户主数据', '期间归属规则', '差异登记', '批次清单', '校验结果', '月度计划', '批次水印'):
        print('  %-8s %d' % (t, cur.execute('SELECT COUNT(*) FROM "%s"' % t).fetchone()[0]))
    print('  归属月分布：')
    for r in cur.execute("SELECT 归属月, count(*) FROM 台账明细 GROUP BY 归属月 ORDER BY 归属月"):
        print('    %s  %d' % r)
    c.close()
    print('DONE')


if __name__ == '__main__':
    main()
