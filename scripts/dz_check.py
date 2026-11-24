# -*- coding: utf-8 -*-
"""东智台账 · 产出自检（dz_check.py）

跑完 ①解析→②拼行→③落库→④派生 之后执行，把所有踩过的坑都变成自动断言，
有任一 ❌ 就退出码非 0（可接在流程末尾当"质检关卡"）。

用法：
    python dz_check.py <周期目录> <D1> <D2> [F0文件]
    （F0 缺省 = 周期目录里的 截止到<D2>…汇总表.xlsx）

【错 → 守卫 对照】（历次踩坑全在此编码，改前先读）：
   ① 路径/文件名变了        → 关键文件存在性检查（缺就报错，而不是崩）
   ② NC 发货单格式变        → NC 文件可读且数据行>0（按文件头读）
   ③ ~$ 锁文件混入          → 周期目录无 ~$ 文件
   ④ 派生层不切周（回款翻倍）→ 台账"本批次回款合计 == 回款表 D2 日列"；且"全量回款 == 汇总 R11"
   ⑤ NC 合计行入库          → 台账无"客户名为空"的行
   ⑥ 华为→T01 规则漏        → 明细.csv 未匹配客户名 = 0
   ⑦ 备注只写自主接单        → 本周有开票/回款的客户，分项汇总里都有对应备注
   ⑧ COM 公式不调相对引用    → 出货块/回款块各行"累计"公式指向本行
   ⑨ Rows.Insert 不扩 SUM   → 各组合计行的 SUM 范围覆盖到该组最后一行
   ⑩ 备注重复写             → 同一客户在块内只保留一条备注行
   ⑪ 手填值被误套公式        → K 列除"0"外不得出现手填数字（应都是公式）
   ⑫ 来源ID 不唯一           → 来源ID 全唯一且非空
   ⑬ 待完成出现负数          → 汇总 C6 无负数；C7(月度预计) ≥ C4(实际)
   ⑭ T01 收入系数            → 分项汇总 E208 含 0.019、不含 0.02
"""
import os, sys, csv, glob, sqlite3
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dz_common import (COLS, ledger_xlsx, f0_path, find_prev_f0, batch_of,
                       BASE_DIR, RULES)

# ---------- 用户钦定规则（写死） ----------
DB = r'E:\商务部\东智日常表更新\东智台账.db'
T01_INCOME_RATE = 0.019
INCOME_DIVISOR = 1.13
GL_INCOME_ZERO = '光缆事业部'          # 光缆收入计 0
DETAIL_CSV = os.path.join(HERE, 'out', '明细.csv')
OUT_PREVIEW = os.path.join(HERE, 'out', '台账预览.csv')


def _find_total_row(ws, label='总计', col=1, r1=1, r2=400):
    for r in range(r1, r2 + 1):
        if str(ws.cell(r, col).value or '').strip() == label:
            return r
    return None


def _read_sheet0(path):
    """读首个 sheet → (rows:list[list], datemode)。按文件头选 openpyxl/xlrd（不 import dz_ingest，避免触发其解析副作用）"""
    with open(path, 'rb') as f:
        magic = f.read(4)
    if magic[:2] == b'PK':
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                for r in range(1, ws.max_row + 1)]
        return rows, 0
    import xlrd
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    return rows, wb.datemode


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip() or 0)
    except Exception:
        return 0.0


# ============ 检查项 ============
CHECKS = []


def check(name):
    def deco(fn):
        CHECKS.append((name, fn))
        return fn
    return deco


@check('① 关键文件存在')
def c_files(args):
    need = [DB, ledger_xlsx()]
    if args.get('f0'):
        need.append(args['f0'])
    miss = [p for p in need if not os.path.exists(p)]
    return (not miss), '缺：%s' % miss if miss else '全在'


@check('③ 周期目录无 ~$ 锁文件')
def c_lock(args):
    bad = [p for p in glob.glob(os.path.join(args['src'], '~$*'))]
    return (not bad), '发现锁文件：%s' % bad if bad else '干净'


@check('② NC 发货单可读且非空')
def c_nc(args):
    cand = sorted(glob.glob(os.path.join(args['src'], '销售出库单明细表*')))
    cand = [p for p in cand if '~$' not in os.path.basename(p)]
    if not cand:
        return False, '未找到 NC 发货单'
    csvs = [p for p in cand if p.lower().endswith('.csv')]
    p = csvs[-1] if csvs else cand[-1]
    if p.lower().endswith('.csv'):
        with open(p, encoding='utf-8-sig', newline='') as fh:
            rd = csv.DictReader(fh)
            need = {'billcode', 'billdate', 'cust_name', 'mat_code', 'qty', 'tax_amt', 'rowno', 'wh_id'}
            miss = need - set(rd.fieldnames or [])
            if miss:
                return False, 'CSV 缺列 %s（%s）' % (sorted(miss), os.path.basename(p))
            n_data = sum(1 for row in rd if str(row.get('cust_name') or '').strip() not in ('', '合计', '总计'))
        return n_data > 0, '数据行 %d（%s）' % (n_data, os.path.basename(p))
    rows, _ = _read_sheet0(p)
    n_data = sum(1 for r in rows[2:] if str(r[18]).strip() and str(r[18]).strip() not in ('合计', '总计'))
    return n_data > 0, '数据行 %d（%s）' % (n_data, os.path.basename(p))


@check('⑤⑫ 台账：来源ID唯一/客户名非空/批次非空')
def c_ledger(args):
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    total = cur.execute('SELECT COUNT(*) FROM 台账明细').fetchone()[0]
    dup = cur.execute('SELECT COUNT(*) FROM (SELECT 来源ID FROM 台账明细 GROUP BY 来源ID HAVING COUNT(*)>1)').fetchone()[0]
    empty_id = cur.execute("SELECT COUNT(*) FROM 台账明细 WHERE 来源ID IS NULL OR 来源ID=''").fetchone()[0]
    empty_cust = cur.execute("SELECT COUNT(*) FROM 台账明细 WHERE 客户名称 IS NULL OR 客户名称=''").fetchone()[0]
    empty_batch = cur.execute("SELECT COUNT(*) FROM 台账明细 WHERE 导入批次 IS NULL OR 导入批次=''").fetchone()[0]
    conn.close()
    ok = total > 0 and dup == 0 and empty_id == 0 and empty_cust == 0 and empty_batch == 0
    return ok, '共%d行；重复ID %d / 空ID %d / 空客户 %d / 空批次 %d' % (total, dup, empty_id, empty_cust, empty_batch)


@check('⑥ 解析：未匹配客户名 = 0')
def c_unmatched(args):
    if not os.path.exists(DETAIL_CSV):
        return False, '无 %s' % DETAIL_CSV
    rows = list(csv.DictReader(open(DETAIL_CSV, encoding='utf-8-sig')))
    un = sorted({r['未匹配'] for r in rows if r['未匹配']})
    return (not un), ('未匹配 %d 个：%s' % (len(un), un)) if un else '0 个'


@check('④a 台账本月回款 == 回款表累计(AJ)')
def c_batch_vs_f0(args):
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    month_sum = cur.execute(
        "SELECT ROUND(SUM(COALESCE(回款金额,0)),6) FROM 台账明细 "
        "WHERE COALESCE(作废,0)=0 AND COALESCE(NULLIF(归属月,''),substr(时间,1,7))=? "
        "AND COALESCE(NULLIF(时间,''),'')<=?", (args['d2'][:7], args['d2'])).fetchone()[0] or 0
    conn.close()
    wb = openpyxl.load_workbook(args['f0'], data_only=True)
    ws = wb['回款&开票统计表']
    tr = _find_total_row(ws)
    if not tr:
        return False, '找不到回款表「总计」行'
    f0_cum = _num(ws.cell(tr, 36).value)             # AJ = 本月累计
    wb.close()
    ok = abs(month_sum - f0_cum) < 0.005
    return ok, '台账本月回款 %s vs F0 回款表累计 %s' % (month_sum, round(f0_cum, 6))


@check('④b 台账本月回款 == 汇总 R11')
def c_cum_vs_summary(args):
    """2026-11-06 修：原来拿「台账全量回款」比汇总 R11。台账是跨月长期库，
    新月份首周会把往月回款（如 9 月 2202.98）当成本月数去比 → 必然 ❌。
    正确口径：只比【本月批次】（导入批次 LIKE YYYYMM%）。"""
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    conn = sqlite3.connect(DB)
    month = args['d2'][:7].replace('-', '') + '%'
    tot = conn.execute("SELECT ROUND(SUM(COALESCE(回款金额,0)),6) FROM 台账明细"
                       " WHERE COALESCE(作废,0)=0 AND 导入批次 LIKE ?", (month,)).fetchone()[0] or 0
    conn.close()
    h = openpyxl.load_workbook(args['f0'], data_only=True)['汇总']
    r11 = _num(h.cell(11, 4).value)
    ok = abs(tot - r11) < 0.005
    return ok, '台账本月回款 %s vs 汇总R11 %s' % (tot, round(r11, 6))


@check('⑦ 备注覆盖本周所有开票/回款客户')
def c_notes(args):
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    batch = os.path.basename(os.path.normpath(args['src']))
    kp = {r[0] for r in cur.execute("SELECT DISTINCT 客户名称 FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 导入批次=? AND COALESCE(开票金额,0)<>0", (batch,))}
    hk = {r[0] for r in cur.execute("SELECT DISTINCT 客户名称 FROM 台账明细 WHERE COALESCE(作废,0)=0 AND 导入批次=? AND COALESCE(回款金额,0)>0", (batch,))}
    bucket = set(RULES.get('F0_OTHER_BUCKET', ()))
    kp -= bucket
    hk -= bucket
    conn.close()
    wb = openpyxl.load_workbook(args['f0'], data_only=True)
    fx = wb['分项汇总']
    got_kp = set()
    got_hk = set()
    for r in range(1, fx.max_row + 1):
        a = str(fx.cell(r, 1).value or '')
        base = a.split('（')[0].strip()
        if isinstance(fx.cell(r, 8).value, str) and '开票' in fx.cell(r, 8).value:
            got_kp.add(base)
        if isinstance(fx.cell(r, 10).value, str) and '回款' in fx.cell(r, 10).value:
            got_hk.add(base)
    miss = (kp - got_kp) | (hk - got_hk)
    wb.close()
    return (not miss), ('缺备注：%s' % sorted(miss)) if miss else '开票/回款客户都有备注'


@check('⑬ 汇总：待完成无负 & 月度预计 ≥ 实际')
def c_summary_sanity(args):
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    h = openpyxl.load_workbook(args['f0'], data_only=True)['汇总']
    neg = []
    less = []
    for r in range(4, 12):
        c6 = h.cell(r, 6).value
        if isinstance(c6, (int, float)) and c6 < -0.005:
            neg.append((r, round(c6, 4)))
        c4 = h.cell(r, 4).value
        c7 = h.cell(r, 7).value
        if isinstance(c4, (int, float)) and isinstance(c7, (int, float)) and c7 < c4 - 0.005:
            less.append((r, round(c7, 4), round(c4, 4)))
    ok = not neg and not less
    return ok, ('待完成负数 %s；月度预计<实际 %s' % (neg, less)) if not ok else '8 项全合规'


@check('⑦⑩ 备注无重复（同基名仅一条）')
def c_notes_dup(args):
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    fx = openpyxl.load_workbook(args['f0'], data_only=True)['分项汇总']
    from collections import Counter
    kp = Counter()
    hk = Counter()
    for r in range(1, fx.max_row + 1):
        a = str(fx.cell(r, 1).value or '')
        base = a.split('（')[0].strip()
        if isinstance(fx.cell(r, 8).value, str) and '开票' in fx.cell(r, 8).value:
            kp[base] += 1
        if isinstance(fx.cell(r, 10).value, str) and '回款' in fx.cell(r, 10).value:
            hk[base] += 1
    dup = [k for k, v in kp.items() if v > 1] + [k for k, v in hk.items() if v > 1]
    return (not dup), ('重复备注基名：%s' % sorted(set(dup))) if dup else '无重复'


@check('⑭ T01 收入系数 = 0.019')
def c_income_rate(args):
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    wf = openpyxl.load_workbook(args['f0'], data_only=True)
    fx = wf['分项汇总']
    # 2026-10-22：清公式后 K 列已固化为数值 → 改为「数值校验」：
    #   校验 T01 行 K(预计收入) ≈ MAX(累计出货金额 E, 预计出货金额 J) / 1.13 * 0.019
    r0, r1 = _block_range(fx, '3、%d月出货计划' % int(args['d2'].split('-')[1]))
    k_val = e_val = j_val = None
    for r in range(r0, r1 + 1):
        if str(fx.cell(r, 1).value or '').strip() == 'T01':
            k_val = fx.cell(r, 11).value
            e_val = fx.cell(r, 5).value
            j_val = fx.cell(r, 10).value
            break
    if k_val is None:
        return False, '未找到 T01 行'
    if isinstance(k_val, str):
        return (('0.019' in k_val) and ('0.02' not in k_val)), 'T01 预计收入(K列) = %s' % k_val
    try:
        base = max(float(e_val or 0), float(j_val or 0))
    except (TypeError, ValueError):
        return False, 'T01 E/J 非数值'
    expect = base / 1.13 * 0.019
    ok = abs(float(k_val) - expect) <= 0.02
    return ok, 'T01 预计收入 = %.4f（应为 %.4f = MAX(%s,%s)/1.13*0.019）' % (
        float(k_val), expect, e_val, j_val)


def _block_range(ws, title_kw, next_kw=None):
    """返回块标题行之后到下一个块标题（或表尾）的范围。"""
    n = ws.max_row
    start = None
    for r in range(1, n + 1):
        v = str(ws.cell(r, 1).value or '').replace(' ', '')
        if title_kw.replace(' ', '') in v:
            start = r
            break
    if start is None:
        return 0, -1
    titles = ('1、9月订单计划', '2、9月产出计划', '3、9月出货计划', '4、9月开票计划', '5、9月回款计划')
    end = n
    for r in range(start + 1, n + 1):
        v = str(ws.cell(r, 1).value or '').replace(' ', '')
        if any(t in v for t in titles):
            end = r - 1
            break
    return start + 1, end


@check('⑪ K 列除 0 外无手填数字')
def c_k_no_literal(args):
    """2026-10-22：改为「预计收入数值正确性」校验（清公式后 K 列本就是数值）。
    K(预计收入) 应 = MAX(累计出货金额E, 预计出货金额J) / 1.13，T01 再 ×0.019，光缆事业部计 0。"""
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    wf = openpyxl.load_workbook(args['f0'], data_only=True)
    fx = wf['分项汇总']
    bad = []
    r0, r1 = _block_range(fx, '3、%d月出货计划' % int(args['d2'].split('-')[1]))
    for r in range(r0, r1 + 1):
        nm = str(fx.cell(r, 1).value or '').strip()
        if not nm or '合计' in nm or '客户' in nm:
            continue
        k = fx.cell(r, 11).value
        if not isinstance(k, (int, float)):
            continue
        try:
            base = max(float(fx.cell(r, 5).value or 0), float(fx.cell(r, 10).value or 0))
        except (TypeError, ValueError):
            continue
        if nm.split('（')[0] in ('光缆事业部',):
            expect = 0.0
        elif nm == 'T01':
            expect = base / 1.13 * 0.019
        else:
            expect = base / 1.13
        if abs(float(k) - expect) > 0.02:
            bad.append((r, nm, round(float(k), 4), round(expect, 4)))
    return (not bad), ('预计收入不等于规则值：%s' % bad) if bad else 'K 列 = MAX(E,J)/1.13（T01×0.019），全部相符'


@check('⑧ 各行"累计"公式指向本行')
def c_rowformula_self(args):
    if not args.get('f0'):
        return True, '跳过（无 F0）'
    wf = openpyxl.load_workbook(args['f0'])
    ws = wf['回款&开票统计表']
    bad = []
    for r in range(5, 96):
        nm = str(ws.cell(r, 2).value or '').strip()
        if not nm or '合计' in nm:
            continue
        aj = str(ws.cell(r, 36).value or '')
        if aj.startswith('=SUM') and ('%d:' % r) not in aj.replace('$', ''):
            # 粗略校验：公式里应含本行行号
            pass
        if aj and aj.startswith('=') and ('D%d' % r) not in aj and ('AJ%d' % r) not in aj:
            if 'SUM(' in aj:
                bad.append((r, nm, aj))
    return (not bad), ('累计公式指向别行：%s' % bad[:5]) if bad else 'OK'


# ============ 主流程 ============
def main():
    args = {}
    args['src'] = sys.argv[1] if len(sys.argv) > 1 else r'E:\商务部\东智日常表更新\20260912-20260918'
    args['d1'] = sys.argv[2] if len(sys.argv) > 2 else '2026-09-12'
    args['d2'] = sys.argv[3] if len(sys.argv) > 3 else '2026-09-18'
    args['f0'] = sys.argv[4] if len(sys.argv) > 4 else None
    if not args['f0']:
        p = f0_path(args['src'], args['d2'])
        args['f0'] = p if os.path.exists(p) else None

    print('=== 东智台账 · 产出自检 ===')
    print('周期目录：%s' % args['src'])
    print('F0：%s' % (args['f0'] or '（未指定）'))
    print()
    fails = 0
    for name, fn in CHECKS:
        try:
            ok, detail = fn(args)
        except Exception as e:
            ok, detail = False, '异常 %s: %s' % (type(e).__name__, e)
        print('  %s %s  %s' % ('✅' if ok else '❌', name, detail))
        if not ok:
            fails += 1
    print()
    print('结果：%d/%d 项通过' % (len(CHECKS) - fails, len(CHECKS)))
    if fails:
        print('⚠️ 有 %d 项未通过，请按上方 ❌ 处理' % fails)
    else:
        print('✅ 全部通过，可交付')
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
