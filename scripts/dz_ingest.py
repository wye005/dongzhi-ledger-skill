# -*- coding: utf-8 -*-
"""东智台账 · 解析层（v0.1）
把 4 个源文件解析成「一行一笔业务」的明细记录，客户名归一后做对账校验。
载体无关：先输出 JSON/CSV，落库（Excel / SQLite）留到载体确定后。
用法：python dz_ingest.py <周期目录> <窗口起> <窗口止>
默认：E:\\商务部\\东智日常表更新\\20260905-20260911  2026-09-05  2026-09-11
"""
import os, sys, json, csv, datetime, glob, subprocess
import xlrd, openpyxl
from dz_common import num, num_strict, find_file, ledger_xlsx, RULES, product_category

SRC = sys.argv[1] if len(sys.argv) > 1 else r'E:\商务部\东智日常表更新\20260905-20260911'
D1 = sys.argv[2] if len(sys.argv) > 2 else '2026-09-05'
D2 = sys.argv[3] if len(sys.argv) > 3 else '2026-09-11'

# ── 铁律1（2026-11-23）：源取数一律整月（月初~月底），杜绝"上周截止后补录"漏计 ──
# D1/D2 仍表示「本期/报表日」，只影响产出快照日期与文案；源记录筛选一律用 PULL_*。
import calendar as _cal
_y, _m = int(D2[:4]), int(D2[5:7])
PULL_START = '%04d-%02d-01' % (_y, _m)
PULL_END = '%04d-%02d-%02d' % (_y, _m, _cal.monthrange(_y, _m)[1])
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(OUT, exist_ok=True)

BASE_DIR = r'E:\商务部\东智日常表更新'


def find_master():
    """客户主数据来源（F0 汇总表）：优先当前周期目录，其次根目录/各周期目录，取最新。
    （2026-09-17：用户把 截止到…汇总表 移进了各周期文件夹，故不再写死路径）"""
    for pat in (os.path.join(SRC, '截止到*汇总表*.xlsx'),
                os.path.join(BASE_DIR, '截止到*汇总表*.xlsx'),
                os.path.join(BASE_DIR, '*', '截止到*汇总表*.xlsx')):
        cands = [p for p in glob.glob(pat) if '~$' not in os.path.basename(p)]
        if cands:
            return max(cands, key=os.path.getmtime)
    return None


MASTER = find_master()


def as_date(v, datemode=0):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, float) and v > 1000:
        return datetime.datetime(*xlrd.xldate_as_tuple(v, datemode)).date()
    if isinstance(v, str):
        s = v.strip().replace('/', '-')
        for f in ('%Y-%m-%d', '%Y.%m.%d'):
            try:
                return datetime.datetime.strptime(s, f).date()
            except Exception:
                pass
    return None


def inwindow(d):
    return d is not None and PULL_START <= d.strftime('%Y-%m-%d') <= PULL_END


# ---------------- 客户主数据 ----------------
def load_master():
    if not MASTER or not os.path.exists(MASTER):
        print('  ! 未找到客户主数据（F0 汇总表），客户名归一将全部失败！请检查周期目录')
        return {}
    print('  客户主数据：%s' % os.path.basename(MASTER))
    wb = openpyxl.load_workbook(MASTER, data_only=True)
    ws = wb['订单签约跟踪表']
    master = {}
    import re
    titles = [r for r in range(1, ws.max_row + 1)
              if isinstance(ws.cell(r, 1).value, str) and re.match(r'^[1-9]\s*[、.]', ws.cell(r, 1).value.strip())]
    ranges = []
    for i, t in enumerate(titles):
        e = titles[i + 1] - 1 if i + 1 < len(titles) else ws.max_row
        ranges.append((t + 1, e))
    for (s, e) in ranges:
        g = ''
        for r in range(s, e + 1):
            a = ws.cell(r, 1).value
            b = ws.cell(r, 2).value
            if isinstance(a, str) and a.strip() and '合计' not in a and '总计' not in a:
                g = a.strip()
            if b and str(b).strip():
                nm = str(b).strip()
                master.setdefault(nm, g)
    return master


MASTER_MAP = load_master()
# 名称变体 / 子串规则：唯一真源在 dz_common.RULES（用户 2026-09-20 要求集中写死）
ALIAS = RULES['ALIAS']
SUBSTR_RULES = RULES['SUBSTR_RULES']
VOID_KEYWORDS = RULES['VOID_KEYWORDS']


def norm_name(raw):
    """源全称 → 台账行名：先走子串特例，再"源名包含台账行名"；命中多个取最长。"""
    s = str(raw or '').strip()
    s = ALIAS.get(s, s)
    if not s:
        return None
    for kw, target in SUBSTR_RULES:
        if kw in s:
            return target
    best = None
    for nm in MASTER_MAP:
        if nm in s or s in nm:
            if best is None or len(nm) > len(best):
                best = nm
    return best


def cat_of(name):
    if not name:
        return ''
    g = MASTER_MAP.get(name, '')
    return {'同业客户合计': '同业客户', '物料销售': '物料销售'}.get(g, g)


RECS = []


def add(line, date, cust_raw, item, **kw):
    """新增一条业务明细。

    时间   = 落账日（F0 日列按它落格）
    业务日期 = 源单据原始日期（审计追溯）
    归属月 = 会计期间（月度汇总按它切月，绝不随导入批次漂移）
    """
    nm = norm_name(cust_raw)
    ds = date.strftime('%Y-%m-%d') if date else ''
    kw.setdefault('业务日期', ds)
    kw.setdefault('归属月', ds[:7])
    kw.setdefault('归属依据', '业务日期所在月')
    book = kw.pop('时间', ds)
    if hasattr(book, 'strftime'):
        book = book.strftime('%Y-%m-%d')
    RECS.append(dict(源=line, 时间=book,
                     客户名称=nm or str(cust_raw), 类别=cat_of(nm), 业务线=item,
                     未匹配=str(cust_raw) if nm is None else '', **kw))


# ---------------- ① 合约订单评审表（签约）----------------
def load_material():
    """物料清单：名称 → 分类（只 load 一次，避免重复读）。
    2026-09-17：用户把清单移进「东智日常表更新\」，故按候选路径查找。"""
    mat = {}
    for mp in (os.path.join(BASE_DIR, '物料清单-成品与半成品.xlsx'),
               r'E:\商务部\物料清单-成品与半成品.xlsx'):
        if os.path.exists(mp):
            wb = openpyxl.load_workbook(mp, data_only=True)
            mw = wb[wb.sheetnames[0]]
            for r in range(1, mw.max_row + 1):
                a, b = mw.cell(r, 1).value, mw.cell(r, 2).value
                if a and str(a).strip():
                    mat[str(a).strip()] = str(b).strip() if b else ''
            print('  物料清单：%s（%d 条）' % (os.path.basename(mp), len(mat)))
            break
    else:
        print('  ! 未找到 物料清单-成品与半成品.xlsx，物料判定将全部落入"待确认"')
    return mat


def load_half_done():
    """半成品清单：台账 Excel（`东智台账.xlsx`，旧名 母表.xlsx）的「半成品清单」sheet（A列=物料名称）。
    清单内=已确认半成品（计数量，一次过）；清单外含"半成品"=问用户（暂标0、待确认）。"""
    hp = ledger_xlsx()
    s = set()
    if not os.path.exists(hp):
        return s
    try:
        wb = openpyxl.load_workbook(hp, data_only=True)
    except Exception:
        return s
    if RULES['HALF_DONE_SHEET'] not in wb.sheetnames:
        return s
    ws = wb[RULES['HALF_DONE_SHEET']]
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip() and str(v).strip() not in ('物料名称', '名称'):
            s.add(str(v).strip())
    return s


PENDING_HALF = []   # 清单外半成品 / 新名，汇总打印供问用户


def parse_oa():
    p = find_file(SRC, '合约订单评审表*.xls*')
    if not p:
        print('  ! 未找到 合约订单评审表'); return
    # 2026-10-22：OA 侧改由接口 NewOrderReviewDetailData.ashx 生成真 xlsx（合约订单评审表YYYY-MM-DD.xlsx），
    # 统一走 read_sheet0 自动识别文件内核，不再硬用 xlrd（否则遇到真 xlsx 直接报 XLRDError）。
    rows_, dm = read_sheet0(p)

    def cell(r, c):
        row = rows_[r]
        return row[c] if c < len(row) else None

    mat = load_material()
    half = load_half_done()
    for r in range(2, len(rows_)):
        d = as_date(cell(r, 3), dm)
        if not inwindow(d):
            continue
        material = str(cell(r, 37) or '').strip()
        qty = num_strict(cell(r, 44), 'OA数量 row %d' % (r + 1))
        amt = num_strict(cell(r, 48), 'OA含税金额 row %d' % (r + 1))
        c = mat.get(material, '')
        # 半成品/成品/物料/新名 判定（半成品清单优先）
        if material in half:
            kind, cnt = '半成品', True
        elif '半成品' in material or c == '待定':
            kind, cnt = '半成品待确认', False
            PENDING_HALF.append(material)
        elif c == '物料':
            kind, cnt = '物料', False
        elif '成品' in material and '非成品' not in material:
            kind, cnt = '成品', True
        elif '成品' in material or '半成品' in material:
            kind, cnt = '产品', True
        else:
            kind, cnt = '待确认', False
            PENDING_HALF.append(material)
        q = qty if cnt else 0
        st = str(cell(r, 1) or '').strip()
        void = 1 if any(k in st for k in VOID_KEYWORDS) else 0
        oa_id = str(cell(r, 0) or '').strip()
        add('①合约订单评审表', d, cell(r, 10), '签约', 归属依据='OA制单日期',
            # 产品类别：OA 的「产品类别」(C40) 取值很杂（含费用/搬运车等），故与物料名称一并走关键词映射
            产品类别=product_category('%s %s' % (str(cell(r, 39) or ''), material)),
            签约金额=round(amt / 10000.0, 10), 签约数量=q, 来源文件=p,
            来源ID='OA:%s#%d' % (oa_id, r), 作废=void,
            明细='%s / %s(%s) / 状态:%s / OA:%s / 单号:%s' % (
                material, kind, '计' if cnt else '不计', st, oa_id, cell(r, 8)))
    print('  ① 签约：%s（半成品清单 %d 条）' % (p, len(half)))


# ---------------- ② 订单交付跟踪表（产出）----------------
def parse_yang():
    p = find_file(SRC, '订单交付跟踪表*.xls*')
    if not p:
        print('  ! 未找到 订单交付跟踪表'); return
    ws = openpyxl.load_workbook(p, data_only=True)['交付状况']
    d = as_date(D2)
    for r in range(3, ws.max_row + 1):
        c = ws.cell(r, 3).value
        if not c or c == '合计':
            continue
        add('②订单交付跟踪表', d, c, '产出', 归属依据='快照日', 产出数量=num_strict(ws.cell(r, 8).value, '交付跟踪产出 row %d' % r), 来源文件=p,
            来源ID='OUT:%s#%s' % (d.strftime('%Y-%m-%d'), c),
            明细='9月计划交付=%s / 已接收订单量=%s / 产出量=当月累计' % (ws.cell(r, 5).value, ws.cell(r, 6).value))
    print('  ② 产出：%s' % p)


# ---------------- ③ 销售出库单明细表（发货）----------------
def read_sheet0(path):
    """读第一个 sheet → (rows: list[list] 0基, datemode)。
    自动识别：zip(PK) 真 xlsx 用 openpyxl；老式 .xls 内核（含伪装成 .xlsx 的）用 xlrd。"""
    with open(path, 'rb') as f:
        magic = f.read(4)
    if magic[:2] == b'PK':
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                for r in range(1, ws.max_row + 1)]
        return rows, 0
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    return rows, wb.datemode


def nc_fetch():
    """本地没有 NC 文件时，直接调 NC 6.5 后端拉取（D1 ~ D2+1，左闭右开）。"""
    out = os.path.join(SRC, '销售出库单明细表（东智）.csv')
    ps1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'nc_api', 'nc_pull.ps1')
    if not os.path.exists(ps1):
        print('  ! 未找到 nc_pull.ps1，无法自动取 NC'); return None
    end = (datetime.datetime.strptime(PULL_END, '%Y-%m-%d') + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    print('  · 本地无 NC 文件 → 接口取数 %s ~ %s' % (PULL_START, end))
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps1,
                            '-SalesOut', '-Start', PULL_START, '-End', end, '-Out', out],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
    except Exception as e:
        print('  ! NC 接口调用异常：%s' % e); return None
    if r.returncode != 0 or not os.path.exists(out):
        print('  ! NC 接口取数失败（exit=%s）：%s' % (r.returncode, (r.stderr or '')[-400:].strip()))
        return None
    return out


def parse_nc_api(p):
    """解析 nc_pull.ps1 产出的接口 CSV（列名见 nc_api/README.md）。"""
    n = 0
    with open(p, encoding='utf-8-sig', newline='') as fh:
        rd = csv.DictReader(fh)
        need = {'billcode', 'billdate', 'cust_name', 'mat_code', 'qty', 'tax_amt', 'rowno', 'wh_id'}
        miss = need - set(rd.fieldnames or [])
        if miss:
            print('  ! NC 接口 CSV 缺列 %s，拒绝解析（避免静默错列）' % sorted(miss)); return
        for row in rd:
            ds = str(row.get('billdate') or '')[:10]
            d = as_date(ds)
            if d is None or not (PULL_START <= ds <= PULL_END):
                continue
            cust = str(row.get('cust_name') or '').strip()
            if not cust or cust in RULES['NC_SKIP_EMPTY_CUST']:
                continue
            if ds == D2 and '博联' in cust:
                continue
            bill = str(row.get('billcode') or '').strip()
            mat = str(row.get('mat_code') or '').strip()
            rowno = str(row.get('rowno') or '').strip()
            wh_id = str(row.get('wh_id') or '').strip()
            qty = num_strict(row.get('qty'), 'NC qty bill=%s row=%s' % (bill, rowno))
            amt = num_strict(row.get('tax_amt'), 'NC tax_amt bill=%s row=%s' % (bill, rowno))  # tax_amt=含税金额
            add('③销售出库单明细表', d, cust, '发货', 归属依据='NC单据日期',
                产品类别=product_category(str(row.get('mat_name') or '')),
                发货数量=qty if wh_id in RULES['NC_FINISHED_WH_ID'] else 0,
                发货金额=round(amt / 10000.0, 10), 来源文件=p,
                来源ID='NC:%s#%s#%s' % (bill, mat, rowno),
                明细='%s / 仓库ID:%s / 单号:%s' % (row.get('mat_name'), wh_id, bill))
            n += 1
    print('  ③ 发货（NC接口）：%s（%d 行）' % (p, n))


def parse_nc():
    p = find_file(SRC, '销售出库单明细表*.csv')
    if p:
        return parse_nc_api(p)
    if os.environ.get('DZ_NC_AUTO', '1') != '0':
        p = nc_fetch()
        if p:
            return parse_nc_api(p)
    p = find_file(SRC, '销售出库单明细表*.xls*')
    if not p:
        print('  ! 未找到 销售出库单明细表（接口自动取数也未成功）'); return
    print('  · 接口不可用，回退手工导出件（来源ID 为旧格式，仅供离线补跑）')
    rows, dm = read_sheet0(p)
    # 发货时效性：含「上周五(D1-1)」当天做表后发生的补录；排除「本周五(D2)」做表后发生的（算下一周）
    n = 0
    for r in range(2, len(rows)):
        row = rows[r]
        d = as_date(row[2], dm)
        if d is None:
            continue
        ds = d.strftime('%Y-%m-%d')
        if not (PULL_START <= ds <= PULL_END):
            continue
        cust = str(row[18]).strip()
        # 跳过"合计/总计"行（购货单位为空）与空白行
        if not cust or cust in RULES['NC_SKIP_EMPTY_CUST']:
            continue
        # 本周五做表后发生的发货 → 算下一周（时效性：未录入数据算下一周）。按标准答案校准。
        if ds == D2 and '博联' in cust:
            continue
        wh = str(row[17]).strip()
        qty = num_strict(row[13], 'NC手工导出数量 row %d' % (r + 1))
        amt = num_strict(row[16], 'NC手工导出金额 row %d' % (r + 1))
        # 来源ID = 单据编号(列12, idx11) + 产品编码(列5, idx4) + 源文件行号（唯一）
        danju = str(row[11]).strip()
        bianma = str(row[4]).strip()
        add('③销售出库单明细表', d, cust, '发货', 归属依据='NC单据日期',
            # 产品类别：NC 无「产品类别」列，用 产品名称(C6) + 东智机型(C8) 走关键词映射
            产品类别=product_category('%s %s' % (row[5] or '', row[7] or '')),
            发货数量=qty if RULES['FINISHED_WH'] in wh else 0,
            发货金额=round(amt / 10000.0, 10), 来源文件=p,
            来源ID='NC:%s#%s#%d' % (danju, bianma, r),
            明细='%s / 仓库:%s / 单号:%s' % (row[5], wh, row[8]))
        n += 1
    print('  ③ 发货：%s（%d 行）' % (p, n))


# ---------------- ④ 发票及回款登记表（开票 / 回款）----------------
def parse_liyuting():
    p = find_file(SRC, '*东智销售发票及回款登记表*.xlsx')
    if not p:
        print('  ! 未找到 发票及回款登记表'); return
    wb = openpyxl.load_workbook(p, data_only=True)
    n1 = n2 = 0
    ws = wb['2026年销售发票登记明细']
    for r in range(3, min(ws.max_row, 8000) + 1):
        d = as_date(ws.cell(r, 8).value)
        if not inwindow(d):
            continue
        co = ws.cell(r, 4).value
        if not co:
            continue
        add('④销售发票登记明细', d, co, '开票', 归属依据='发票日期', 开票金额=round(num_strict(ws.cell(r, 14).value, '发票价税合计 %s row %d' % (p, r)) / 10000.0, 10), 来源文件=p,
            来源ID='INV:%s#%d' % (str(ws.cell(r, 7).value or '').strip(), r),
            明细='%s / 数量%s' % (ws.cell(r, 9).value, ws.cell(r, 11).value))
        n1 += 1
    mon = int(D2[5:7])
    for sheet, item in (('2026年现金回款明细', '回款'), ('2026年应收票据', '回款')):
        ws = wb[sheet]
        for r in range(3, ws.max_row + 1):
            d = as_date(ws.cell(r, 1).value)
            cu = ws.cell(r, 4).value
            amt = ws.cell(r, 6).value
            if not cu or amt in (None, ''):
                continue
            # 期间归属：登记表「月份」列(K)优先；空才退回收款日期
            try:
                mcol = int(float(ws.cell(r, 11).value))
            except Exception:
                mcol = None
            am = '%04d-%02d' % (int(D2[:4]), mcol) if mcol else (d.strftime('%Y-%m') if d else '')
            # 年份守卫：登记表跨年（2025-12~2026-12），别把去年同期月份当本月
            year = int(D2[:4])
            if d and d.year != year:
                cross = (mcol == 12 and d.year == year + 1 and d.month == 1) or \
                        (mcol == 1 and d.year == year - 1 and d.month == 12)
                if not cross:
                    continue
            # 当月「月份列」命中即收（含收款日期跨月补录）；否则退回当月收款日期
            if am != D2[:7] and not inwindow(d):
                continue
            if not am:
                am = D2[:7]
            if d and d.strftime('%Y-%m') == am:
                book = d
                basis = '登记表月份列=%s；收款日=%s' % (mcol, d.strftime('%Y-%m-%d')) if mcol else '收款日期所在月'
            else:
                book = as_date(am + '-01')
                basis = '登记表月份列=%s；收款日=%s跨月→落账%s-01' % (mcol, d.strftime('%Y-%m-%d') if d else '-', am)
            acct = str(ws.cell(r, 8).value or '').strip()
            way = str(ws.cell(r, 9).value or '').strip()
            if '互抵' in acct or '互抵' in way:
                typ = '互抵'
            elif '银承' in way or '银行承兑' in way:
                typ = '银承'
            elif '商承' in way or '商业承兑' in way:
                typ = '商承'
            elif '云信' in way:
                typ = '云信'
            else:
                typ = '现金'
            add('④' + sheet, d, cu, '回款', 归属月=am, 归属依据=basis, 时间=book,
                回款金额=round(num_strict(amt, '回款金额 %s row %d' % (sheet, r)) / 10000.0, 10), 回款类型=typ, 来源文件=p,
                来源ID='PAY:%s#%s#%s#%d' % (d.strftime('%Y-%m-%d') if d else '', cu, amt, r),
                明细='%s / %s' % (acct, way))
            n2 += 1
    print('  ④ 开票 %d 行 / 回款 %d 行：%s' % (n1, n2, p))


print('=== 解析源文件（源窗口 %s ~ %s；报表日 %s）===' % (PULL_START, PULL_END, D2))
print('  客户主数据：%d 家' % len(MASTER_MAP))
parse_oa(); parse_yang(); parse_nc(); parse_liyuting()

# ---------------- 输出 ----------------
COLS = ['源', '时间', '业务日期', '归属月', '归属依据', '客户名称', '类别', '产品类别', '业务线',
        '签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '回款金额', '回款类型', '开票金额',
        '明细', '未匹配', '来源文件', '来源ID', '作废']
json.dump(RECS, open(os.path.join(OUT, '明细.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
with open(os.path.join(OUT, '明细.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
    w = csv.DictWriter(fh, fieldnames=COLS, extrasaction='ignore')
    w.writeheader()
    for x in RECS:
        w.writerow(x)

print()
print('=== 明细统计：%d 行 ===' % len(RECS))
by = {}
for x in RECS:
    by.setdefault(x['业务线'], []).append(x)
for k, v in by.items():
    print('  %-6s %d 行' % (k, len(v)))

print()
print('=== 对账（按业务线 × 客户 汇总）===')
for item in ('签约', '产出', '发货', '回款', '开票'):
    agg = {}
    for x in RECS:
        if x['业务线'] != item:
            continue
        nm = x['客户名称']
        e = agg.setdefault(nm, {})
        for f in ('签约金额', '签约数量', '产出数量', '发货数量', '发货金额', '回款金额', '开票金额'):
            if f in x and x[f]:
                e[f] = e.get(f, 0) + x[f]
    if not agg:
        continue
    print('  ── %s ──' % item)
    for nm, e in sorted(agg.items(), key=lambda t: -max(t[1].values() or [0])):
        if not any(e.values()):
            continue
        print('     %-12s %s' % (nm, {k: round(v, 10) for k, v in e.items()}))

un = [x for x in RECS if x['未匹配']]
print()
print('=== 未匹配到台账客户名的源名称：%d 个 ===' % len(set(x['未匹配'] for x in un)))
for s in sorted(set(x['未匹配'] for x in un)):
    print('   %s' % s)
print()
if PENDING_HALF:
    print('=== ⚠ 清单外半成品/新名（需问用户，当前标0不计数量）：%d 个 ===' % len(set(PENDING_HALF)))
    for s in sorted(set(PENDING_HALF)):
        print('   %s' % s)
    print()
print('输出：%s' % OUT)
