# -*- coding: utf-8 -*-
"""东智台账 · 落库行构建（v0.2）
读 out/明细.json，把「产出」行换成"每天增量"，再拼出台账行 → out/台账预览.csv
用法：python build_ledger.py [周期目录] [起] [止]
默认：E:\\商务部\\东智日常表更新\\20260912-20260918  2026-09-12  2026-09-18
"""
import os, sys, json, csv, re, glob, openpyxl, datetime
from dz_common import COLS, find_prev_f0, RULES

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'out')
BASE_DIR = r'E:\商务部\东智日常表更新'
SRC = sys.argv[1] if len(sys.argv) > 1 else r'E:\商务部\东智日常表更新\20260912-20260918'
D1 = sys.argv[2] if len(sys.argv) > 2 else '2026-09-12'
D2 = sys.argv[3] if len(sys.argv) > 3 else '2026-09-18'

recs = json.load(open(os.path.join(OUT, '明细.json'), encoding='utf-8'))
rows = [r for r in recs if r['业务线'] != '产出']          # 产出先剔除，改按天重算

# ---- 按天重算产出 ----
YEAR = D1[:4]
fs = [f for f in sorted(glob.glob(os.path.join(SRC, '订单交付跟踪表*.xls*')))
      if '~$' not in os.path.basename(f)]
snaps = {}
snap_path = {}
for f in fs:
    m = re.search(r'(\d{2})(\d{2})\.xls', os.path.basename(f))
    if not m:
        continue
    d = '%s-%s-%s' % (YEAR, m.group(1), m.group(2))
    ws = openpyxl.load_workbook(f, data_only=True)['交付状况']
    cur = {}
    for r0 in range(3, ws.max_row + 1):
        c = ws.cell(r0, 3).value
        if c and c != '合计':
            cur[str(c).strip()] = ws.cell(r0, 8).value or 0
    snaps[d] = cur
    snap_path[d] = f
days = sorted(snaps)
print('产出快照：%s' % days)

# 产出源用简称 → 台账名
OUT_ALIAS = {'博联': '博联通信', 'T01': 'T01', '中兴': '中兴PON', '鑫合盛': '鑫合盛',
             '优光谱': '优光谱', '上海引望': '引望', '多模尔': '多模尔', '玖耀启铭': '玖耀启铭'}

# 快照 key 统一映射成台账名
snaps_ledger = {}
for d, cur in snaps.items():
    snaps_ledger[d] = {OUT_ALIAS.get(nm, nm): v for nm, v in cur.items()}


def load_prev_cum(path):
    """读上一版 F0「产出发货跟踪表」块1（产出数量）的「本月累计」AJ(36)，作为本周产出增量基线"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['产出发货跟踪表']
    base = {}
    for r in range(5, 70):                 # 只扫产出块（R5-R69），别串到发货块
        nm = ws.cell(r, 2).value
        if not (nm and str(nm).strip()):
            continue
        v = ws.cell(r, 36).value
        if isinstance(v, (int, float)) and v:
            base[str(nm).strip()] = v
    return base


# ---------- 跨月守卫（2026-11-06 新增，皇上"犯过的错写死成代码"）----------
# 交付跟踪表的「产出量」是【本月累计】，每月重置。月初第一周最容易踩两个坑：
#   ① 上一版 F0 是上月 30/31 日的（新月首周必然如此）→ 直接拿它的本月累计当基线，会把上月数
#      当成本月起点，产出全部算成 0 或负数；
#   ② 周期目录里常放着上月末那份快照（做基线用），若不过滤，会把上月累计当本月增量。
MONTH = D1[:7]
_all_days = sorted(snaps)
month_days = [d for d in _all_days if d[:7] == MONTH]
prev_days = [d for d in _all_days if d[:7] < MONTH]
stale = set()
if prev_days:
    sig_prev = {k: v for k, v in snaps_ledger[prev_days[-1]].items() if v}
    for d in month_days:
        sig = {k: v for k, v in snaps_ledger[d].items() if v}
        if sig_prev and sig == sig_prev:      # 表未刷新，仍是上月累计
            stale.add(d)
            print('  ! 快照 %s 的产出量与上月最后一份 %s 完全一致（表未刷新），已跳过' % (d, prev_days[-1]))
month_days = [d for d in month_days if d not in stale]
if prev_days:
    print('  · 跨月：上月快照 %s 只作参照，不参与本月增量' % prev_days)
days = month_days

BASE_F0, BASE_DATE = find_prev_f0(D1)
if BASE_F0 and BASE_DATE[:7] == MONTH:
    BASE = load_prev_cum(BASE_F0)
else:
    # 新月首周：上月 F0 的「本月累计」是上月数，不能当基线 → 从 0 起算
    BASE = {}
    print('  ★ 新月首周：基准 F0 日期 %s 不在本月（%s），产出基线按 0 起算' % (BASE_DATE or '未找到', MONTH))
print('产出基线来源：%s（%s 累计）' % (os.path.basename(BASE_F0) if BASE_F0 else '未找到', BASE_DATE or '-'))

# 2026-11-23 修：本周目录可能只有「本周」几天的快照（如 11/17-11/20），
# BASE 为空时会把第一天的【当月累计】当成【当日增量】。用台账已入库的
# 本月产出（首张快照日之前）补齐基线，避免重复计。
if days:
    import sqlite3
    conn = sqlite3.connect(r'E:\商务部\东智日常表更新\东智台账.db')
    db_base = {}
    for nm, q in conn.execute(
            "SELECT 客户名称, SUM(COALESCE(产出数量,0)) FROM 台账明细 "
            "WHERE COALESCE(NULLIF(归属月,''), substr(时间,1,7))=? "
            "AND COALESCE(作废,0)=0 AND 时间 < ? GROUP BY 客户名称",
            (MONTH, days[0])):
        if q:
            db_base[nm] = q
    conn.close()
    if db_base:
        for nm, v in db_base.items():
            BASE[nm] = max(float(BASE.get(nm, 0) or 0), float(v))
        print('  · 台账补齐产出基线 %d 家（%s 之前累计）' % (len(db_base), days[0]))
print('产出基线：%s' % BASE)

prev = BASE
for d in days:
    for nm, v in snaps_ledger[d].items():
        pv = prev.get(nm, 0) or 0
        inc = (v or 0) - pv
        if inc == 0:
            continue
        if inc < 0:                        # 本月内累计回退（改单/重算），不计负值，留痕
            print('  ! %s %s 累计回退 %s ← %s，跳过' % (d, nm, v, pv))
            continue
        rows.append({'源': '②订单交付跟踪表', '时间': d,
                     '客户名称': nm, '类别': '', '业务线': '产出',
                     '产出数量': inc, '来源文件': snap_path[d],
                     '业务日期': d, '归属月': d[:7], '归属依据': '快照日',
                     '来源ID': 'OUT:%s#%s' % (d, nm), '作废': 0,
                     '明细': '当日增量（累计%s ← 前值%s）' % (v, pv)})
    prev = snaps_ledger[d]

# 类别补全
master = {}
for r in recs:
    if r['类别']:
        master[r['客户名称']] = r['类别']
for r in rows:
    if not r['类别']:
        r['类别'] = master.get(r['客户名称'], '')

# ---- 拼 20 列 ----
TALK = {'签约': ['签约金额', '签约数量'], '产出': ['产出数量'], '发货': ['发货数量', '发货金额'],
        '回款': ['回款金额', '回款类型'], '开票': ['开票金额']}
FIELDS = ['签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '回款金额', '回款类型', '开票金额']

# 收入规则：唯一真源在 dz_common.RULES（用户 2026-09-15 定、2026-09-18 确认"按台账算"）
VAT = RULES['VAT']
INCOME_ZERO = set(RULES['INCOME_ZERO'])
INCOME_RATE = RULES['INCOME_RATE']


def calc_income(cust, amt):
    """由发货金额推导收入"""
    name = str(cust or '').strip()
    if name in INCOME_ZERO:
        return 0.0
    v = (amt or 0) / VAT
    r = INCOME_RATE.get(name)
    if r:
        v *= r
    return round(v, 10)
out = []
for r in rows:
    row = {c: '' for c in COLS}
    row['客户名称'] = r.get('客户名称', '')
    row['类别'] = r.get('类别', '')
    row['产品类别'] = r.get('产品类别', '')
    row['时间'] = r.get('时间', '')
    row['业务日期'] = r.get('业务日期', r.get('时间', ''))
    row['归属月'] = r.get('归属月', str(r.get('时间', ''))[:7])
    row['归属依据'] = r.get('归属依据', '业务日期所在月')
    row['信息来源路径'] = r.get('来源文件', '')
    row['来源ID'] = r.get('来源ID', '')
    row['作废'] = r.get('作废', 0) or 0
    for f in TALK.get(r['业务线'], []):
        if f in r and r[f] not in (None, ''):
            row[f] = r[f]
    # 收入：仅"发货"行产生（由发货金额推导）
    if r['业务线'] == '发货':
        amt = r.get('发货金额') or 0
        if isinstance(amt, str):
            try:
                amt = float(amt)
            except Exception:
                amt = 0.0
        row['收入'] = calc_income(row['客户名称'], amt)
    out.append(row)

out.sort(key=lambda x: (x['时间'], x['客户名称'], x['签约金额'] or 0))
with open(os.path.join(OUT, '台账预览.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
    w = csv.DictWriter(fh, fieldnames=COLS)
    w.writeheader()
    w.writerows(out)

print()
print('=== 台账预览：%d 行 ===' % len(out))
by = {}
for r in out:
    by.setdefault(r['时间'], 0)
    by[r['时间']] += 1
print('  按日期：%s' % by)
print()
SHOW = ['时间', '客户名称', '签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '收入',
        '回款金额', '开票金额']
print(' | '.join(SHOW))
for r in out:
    if r['时间'] >= '2026-09-08':
        print(' | '.join(str(r.get(c, '')) for c in SHOW))
print()
print('明细行数分布（把产出/发货等同日同客户的多笔都展平了）：')
cnt = {}
for r in out:
    k = (r['时间'], r['客户名称'])
    cnt[k] = cnt.get(k, 0) + 1
for k, v in sorted(cnt.items()):
    if v > 1:
        print('   %s %-10s %d 行' % (k[0], k[1], v))
print('输出：%s' % os.path.join(OUT, '台账预览.csv'))
