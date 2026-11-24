# -*- coding: utf-8 -*-
"""东智台账 · 落库层（v0.3）
读 out/台账预览.csv（一行一笔业务、23 列），双写：
  1) SQLite 真身：E:\\商务部\\东智日常表更新\\东智台账.db
     —— 只增不改：INSERT OR IGNORE（来源ID 唯一索引去重）+ 作废标记 UPDATE；不再 DELETE 窗口。
  2) Excel 镜像：`东智台账.xlsx`「东智台账」sheet —— 由 SQLite 全量回写（镜像，永远与真身一致）。
另：跑前自动备份 db 与 xlsx 到 _backup\\<时间戳>\\；跑完把本次源文件归档到 _snapshots\\<周期>\\。
用法：python dz_store.py [窗口起] [窗口止]
"""
import os, sys, re, csv, glob, sqlite3, datetime, shutil, json, hashlib
from dz_common import COLS, NUM_COLS, ledger_xlsx, batch_of

HERE = os.path.dirname(os.path.abspath(__file__))
CSV_P = os.path.join(HERE, 'out', '台账预览.csv')
DB_P = os.environ.get('DZ_DB') or r'E:\商务部\东智日常表更新\东智台账.db'
XLSX_P = os.environ.get('DZ_XLSX') or ledger_xlsx()   # 「东智台账.xlsx」（兼容旧名「母表.xlsx」）
BAK_DIR = r'E:\商务部\东智日常表更新\_backup'
SNAP_DIR = r'E:\商务部\东智日常表更新\_snapshots'

D1 = sys.argv[1] if len(sys.argv) > 1 else '2026-09-05'
D2 = sys.argv[2] if len(sys.argv) > 2 else '2026-09-11'

# 归档的源文件类型（防误删；产出快照最关键）
SNAP_PATTERNS = ['合约订单评审表*', '销售出库单明细表*', '订单交付跟踪表*',
                 '*东智销售发票及回款登记表*']


def to_val(c, s):
    if s is None or s == '':
        return None
    if c in NUM_COLS:
        try:
            return float(str(s).replace(',', '').strip())
        except Exception:
            return None
    return str(s).strip()


def load_csv():
    rows = list(csv.DictReader(open(CSV_P, encoding='utf-8-sig')))
    out = []
    for r in rows:
        out.append({c: to_val(c, r.get(c)) for c in COLS})
    return out


def backup():
    """跑前备份 db + xlsx 到 _backup\\<时间戳>\\"""
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    bak = os.path.join(BAK_DIR, ts)
    os.makedirs(bak, exist_ok=True)
    n = 0
    for src in (DB_P, XLSX_P):
        if os.path.exists(src):
            shutil.copy(src, os.path.join(bak, os.path.basename(src)))
            n += 1
    print('已备份 %d 个文件 -> %s' % (n, bak))
    return bak


def archive_sources(rows):
    """把本次用到的源文件归档到 _snapshots\\<周期目录名>\\（增量复制、已存在则跳过）。
    防误删；产出快照最关键（产出=相邻快照相减，丢一份就归错日）。"""
    dirs = sorted({os.path.normpath(os.path.dirname(str(r.get('信息来源路径') or '').strip()))
                   for r in rows if r.get('信息来源路径')})
    dirs = [d for d in dirs if d and os.path.isdir(d)]
    if not dirs:
        print('源文件归档：跳过（台账行里没有可用的信息来源路径）')
        return
    for d in dirs:
        tag = os.path.basename(d.rstrip('\\/')) or 'unknown'
        dest = os.path.join(SNAP_DIR, tag)
        os.makedirs(dest, exist_ok=True)
        total = copied = 0
        for pat in SNAP_PATTERNS:
            for p in sorted(glob.glob(os.path.join(d, pat))):
                if '~$' in os.path.basename(p):
                    continue
                total += 1
                t = os.path.join(dest, os.path.basename(p))
                if not os.path.exists(t):
                    shutil.copy(p, t)
                    copied += 1
        print('源文件归档：%s -> %s（新增 %d / 共 %d 个）' % (tag, dest, copied, total))


# ============ 0) 客户ID / 批次 ============
def resolve_customer_ids(cur, rows):
    """按客户主数据落 客户ID；主数据没有的自动建一行（状态=待确认），绝不静默丢客户。"""
    master = {r[0]: r[1] for r in cur.execute('SELECT 标准名, 客户ID FROM 客户主数据')}
    nums = [int(x[1:]) for x in master.values() if re.match(r'^C\d+$', x)]
    next_id = (max(nums) + 1) if nums else 1
    added = []
    mapping = {}
    for r in rows:
        nm = str(r.get('客户名称') or '').strip()
        if not nm:
            continue
        cid = master.get(nm)
        if not cid:
            cid = 'C%03d' % next_id
            next_id += 1
            cur.execute('''INSERT INTO 客户主数据(客户ID,标准名,别名,类别,客户状态,生效月,备注)
                           VALUES(?,?,?,?,?,?,?)''',
                        (cid, nm, '', r.get('类别') or '', '待确认', '2026-09', '落库时自动发现'))
            master[nm] = cid
            added.append(nm)
        mapping[nm] = cid
    if added:
        print('  + 客户主数据新增待确认：%s' % '、'.join(added))
    return mapping


def create_batch(cur, rows, d1, d2):
    """生成本次周更批次 + 源文件清单（sha256）→ out/批次.json，供 F0 水印使用。"""
    now = datetime.datetime.now()
    batch = 'WK-%s-%s' % (now.strftime('%Y%m%d_%H%M%S'), d2.replace('-', ''))
    files = sorted({str(r.get('信息来源路径') or '').strip() for r in rows if r.get('信息来源路径')})
    h = hashlib.sha256()
    n = 0
    for p in files:
        if not os.path.exists(p):
            continue
        st = os.stat(p)
        cur.execute('''INSERT OR REPLACE INTO 批次清单
            (批次ID,文件类型,文件名,路径,字节数,修改时间,sha256,归档时间)
            VALUES(?,?,?,?,?,?,?,?)''',
            (batch, os.path.splitext(os.path.basename(p))[1].lstrip('.').lower(),
             os.path.basename(p), p, st.st_size,
             datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
             file_sha256(p), now.strftime('%Y-%m-%d %H:%M:%S')))
        h.update(file_sha256(p).encode('ascii'))
        n += 1
    watermark = 'DZ-%s-%s' % (batch, h.hexdigest()[:8] if n else 'nosrc')
    meta = {'批次ID': batch, '水印': watermark, '数据截止': d2,
            '窗口起': d1, '源文件数': n, '生成时间': now.strftime('%Y-%m-%d %H:%M:%S')}
    with open(os.path.join(HERE, 'out', '批次.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('  批次：%s 水印：%s 源文件 %d 个' % (batch, watermark, n))
    return meta


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


# ============ 1.5) 业务键比对（治行号漂移） ============
def _sig(r):
    """返回 (业务类型, 金额, 数量) —— 用于判断「是不是同一笔业务」。"""
    g = lambda k: float(r.get(k) or 0)
    if g('签约金额') or g('签约数量'):
        return ('签约', round(g('签约金额'), 4), round(g('签约数量'), 2))
    if g('发货金额') or g('发货数量'):
        return ('发货', round(g('发货金额'), 4), round(g('发货数量'), 2))
    if g('回款金额'):
        return ('回款', round(g('回款金额'), 4), 0.0)
    if g('开票金额'):
        return ('开票', round(g('开票金额'), 4), 0.0)
    if g('产出数量'):
        return ('产出', 0.0, round(g('产出数量'), 2))
    return ('其他', 0.0, 0.0)


def _load_existing(cur, month):
    out = []
    for r in cur.execute('''SELECT id,来源ID,时间,客户名称,签约金额,签约数量,发货金额,发货数量,
                                  回款金额,开票金额,产出数量
                            FROM 台账明细
                            WHERE COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?''', (month,)):
        out.append(dict(id=r[0], sid=str(r[1] or ''), time=r[2], cust=r[3],
                        sig=_sig(dict(签约金额=r[4], 签约数量=r[5], 发货金额=r[6],
                                      发货数量=r[7], 回款金额=r[8], 开票金额=r[9], 产出数量=r[10]))))
    return out


def _same_biz(old, r):
    """同一天、同一客户、同一源类型、同一业务签名 → 视为同一笔（行号漂移）。"""
    if old['time'] != r.get('时间') or old['cust'] != r.get('客户名称'):
        return False
    a = old['sid'].split(':', 1)[0]
    b = str(r.get('来源ID') or '').split(':', 1)[0]
    if a != b:
        return False
    so, sn = old['sig'], _sig(r)
    if so[0] != sn[0]:
        return False
    # 金额差 <= 0.005 万、数量差 <= 0.5（容忍源表末位四舍五入）
    return abs(so[1] - sn[1]) <= 0.005 and abs(so[2] - sn[2]) <= 0.5


# ============ 1) SQLite（只增不改）============
def store_sqlite(rows, d1, d2):
    os.makedirs(os.path.dirname(DB_P), exist_ok=True)
    conn = sqlite3.connect(DB_P)
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS 台账明细 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            客户名称 TEXT NOT NULL,
            类别 TEXT,
            时间 TEXT NOT NULL,
            签约金额预测 REAL, 签约金额 REAL, 签约数量预测 REAL, 签约数量 REAL,
            产出计划 REAL, 产出数量 REAL,
            发货数量交付计划 REAL, 发货数量 REAL, 发货金额交付计划 REAL, 发货金额 REAL,
            收入 REAL,
            回款计划 REAL, 回款金额 REAL, 回款类型 TEXT,
            开票计划 REAL, 开票金额 REAL, 信息来源路径 TEXT,
            来源ID TEXT, 作废 INTEGER DEFAULT 0,
            录入时间 TEXT DEFAULT (datetime('now','localtime'))
        )
    ''')
    # 兼容旧表：缺列则自动补
    existing = [r[1] for r in cur.execute('PRAGMA table_info(台账明细)').fetchall()]
    for c in COLS:
        if c not in existing:
            typ = 'REAL' if c in NUM_COLS else 'TEXT'
            cur.execute('ALTER TABLE 台账明细 ADD COLUMN %s %s' % (c, typ))
    cur.execute('CREATE INDEX IF NOT EXISTS idx_time ON 台账明细(时间)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cust ON 台账明细(客户名称)')
    cur.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_src ON 台账明细(来源ID)')

    cid_map = resolve_customer_ids(cur, rows)
    existing = _load_existing(cur, d2[:7])
    used = set()
    n = nskip = nrekey = nvoid = n_noid = 0
    for r in rows:
        sid = r.get('来源ID')
        if not sid:
            n_noid += 1
        r['导入批次'] = batch_of(r.get('信息来源路径'))      # 首次入库的周期目录名（派生层按它切周）
        cid = cid_map.get(str(r.get('客户名称') or '').strip())

        # 1) 来源ID 已存在 → 跳过（正常幂等）
        hit = next((e for e in existing if e['sid'] == sid), None)
        if hit:
            nskip += 1
            if cid:
                cur.execute('UPDATE 台账明细 SET 客户ID=? WHERE id=?', (cid, hit['id']))
            continue

        # 2) 来源ID 没命中，但同一笔业务已存在 → 只改键 + 刷新数值（治行号漂移）
        match = next((e for e in existing if e['id'] not in used and _same_biz(e, r)), None)
        if match:
            cur.execute('UPDATE 台账明细 SET 来源ID=?, 签约金额=?, 签约数量=?, 发货数量=?, '
                        '发货金额=?, 回款金额=?, 开票金额=?, 产出数量=?, 业务日期=?, 归属月=?, '
                        '归属依据=?, 客户ID=? WHERE id=?',
                        (sid, r.get('签约金额'), r.get('签约数量'), r.get('发货数量'),
                         r.get('发货金额'), r.get('回款金额'), r.get('开票金额'), r.get('产出数量'),
                         r.get('业务日期'), r.get('归属月'), r.get('归属依据'), cid, match['id']))
            used.add(match['id'])
            match['sid'] = sid
            nrekey += 1
            continue

        # 3) 真正的新行 → 插入
        before = cur.connection.total_changes
        vals = [r[c] for c in COLS]
        cur.execute('INSERT OR IGNORE INTO 台账明细 (%s) VALUES (%s)' % (
            ','.join(COLS), ','.join('?' * len(COLS))), vals)
        if cur.connection.total_changes > before:
            n += 1
            if cid:
                cur.execute('UPDATE 台账明细 SET 客户ID=? WHERE 来源ID=?', (cid, sid))
        else:
            nskip += 1
        # 作废单：留痕标记（已存在的旧行也要翻成作废）
        if r.get('作废'):
            cur.execute('UPDATE 台账明细 SET 作废=1 WHERE 来源ID=?', (sid,))
            nvoid += 1

    # 回填历史行的「导入批次」（按 信息来源路径 的父目录名）
    nb = 0
    for _id, p in cur.execute("SELECT id, 信息来源路径 FROM 台账明细 "
                              "WHERE 导入批次 IS NULL OR 导入批次=''").fetchall():
        b = batch_of(p)
        if b:
            cur.execute('UPDATE 台账明细 SET 导入批次=? WHERE id=?', (b, _id))
            nb += 1
    create_batch(cur, rows, d1, d2)
    conn.commit()
    total = cur.execute('SELECT COUNT(*) FROM 台账明细').fetchone()[0]
    conn.close()
    return dict(新增=n, 跳过=nskip, 改键=nrekey, 标作废=nvoid, 无来源ID=n_noid, 回填批次=nb, 累计=total)


# ============ 2) Excel 镜像（SQLite 全量回写）============
def _norm_hdr(s):
    import re
    if s is None:
        return ''
    return re.sub(r'[（(][^）)]*[）)]', '', str(s)).strip()


def store_excel():
    import openpyxl
    conn = sqlite3.connect(DB_P)
    cur = conn.cursor()
    cur.execute('SELECT %s FROM 台账明细 ORDER BY 时间, id' % ','.join(COLS))
    all_rows = cur.fetchall()
    conn.close()

    wb = openpyxl.load_workbook(XLSX_P)
    ws = wb['东智台账'] if '东智台账' in wb.sheetnames else wb[wb.sheetnames[0]]
    # 表头校正：缺列才插（忽略单位后缀，保留用户已写的「（万元）」）
    for j, c in enumerate(COLS):
        if _norm_hdr(ws.cell(1, j + 1).value) == c:
            continue
        row1n = [_norm_hdr(ws.cell(1, k + 1).value) for k in range(ws.max_column)]
        if c in row1n:
            print('   ! 表头错位：第 %d 列期望[%s]，实际[%s]' % (j + 1, c, ws.cell(1, j + 1).value))
            continue
        ws.insert_cols(j + 1)
        ws.cell(1, j + 1).value = c
        print('   + 插入表头列：%s（第 %d 列）' % (c, j + 1))
    # 清空数据区，写全量镜像
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for i, row in enumerate(all_rows):
        rr = 2 + i
        for j, v in enumerate(row):
            ws.cell(rr, j + 1).value = v
    # 台账镜像也带批次水印（写文档属性，不可见、不改表结构）
    try:
        import json as _json
        mp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out', '批次.json')
        meta = _json.load(open(mp, encoding='utf-8'))
        wb.properties.keywords = 'DZ批次水印=%s;批次ID=%s;截止=%s' % (
            meta.get('水印', ''), meta.get('批次ID', ''), meta.get('数据截止', ''))
    except Exception:
        pass
    wb.save(XLSX_P)


# ============ 对账 ============
def check(rows):
    agg = {}
    for r in rows:
        k = r['客户名称']
        e = agg.setdefault(k, {})
        for f in ('签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '收入', '回款金额', '开票金额'):
            if r[f] not in (None, 0):
                e[f] = e.get(f, 0) + r[f]
    print('=== 台账非零汇总（按客户）===')
    for k, e in sorted(agg.items(), key=lambda t: -max(t[1].values() or [0])):
        if any(e.values()):
            print('  %-12s %s' % (k, {x: round(e[x], 10) for x in e}))


def main():
    if not os.path.exists(CSV_P):
        print('! 找不到 %s ，先跑 build_ledger.py' % CSV_P)
        return
    rows = load_csv()
    print('读入台账行：%d 行' % len(rows))
    if rows and len(rows[0]) != len(COLS):
        print('! 列数异常：%d（应 %d）' % (len(rows[0]), len(COLS)))
        return
    print()
    check(rows)
    print()
    backup()
    st = store_sqlite(rows, D1, D2)
    print('SQLite：新增 %d / 跳过(已存在) %d / 改键(行号漂移) %d / 标作废 %d / 无来源ID %d / 回填批次 %d / 库内累计 %d -> %s' % (
        st['新增'], st['跳过'], st['改键'], st['标作废'], st['无来源ID'], st['回填批次'], st['累计'], DB_P))
    store_excel()
    print('Excel ：已由 SQLite 全量回写 -> %s' % XLSX_P)
    mp = os.path.join(HERE, 'out', '批次.json')
    if os.path.exists(mp):
        meta = json.load(open(mp, encoding='utf-8'))
        conn = sqlite3.connect(DB_P); cur = conn.cursor()
        cur.execute('''INSERT OR REPLACE INTO 批次水印(对象,批次ID,水印,数据截止,说明)
                       VALUES(?,?,?,?,?)''',
                    ('台账镜像', meta['批次ID'], meta['水印'], meta.get('数据截止', ''),
                     '写 Excel 文档属性 keywords'))
        conn.commit(); conn.close()
        print('  台账水印：%s' % meta['水印'])
    print()
    archive_sources(rows)
    print()
    print('DONE')


if __name__ == '__main__':
    main()
