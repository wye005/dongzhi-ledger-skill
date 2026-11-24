# -*- coding: utf-8 -*-
"""东智台账 · 派生层（v0.3）
把台账（SQLite 真身）按客户聚合本周数据，派生成老宽表 F0（本周五版）：
  1) 从上一版 F0（上周五基准）复制一份 → 本周五版
  2) Excel COM 填三张子表的「本周五列」（本周合计 → 周五列）
  3) 修正分项汇总「文件原有」引用错位（多模尔/云鹏锡业）
  4) CalculateFullRebuild 重算（分项汇总/汇总自动联动）
  5) 续写分项汇总备注 + 重写汇总 H4:H11 备注
  6) 回读校验
用法：python dz_derive.py [周期目录] [起] [止]
默认：E:\\商务部\\东智日常表更新\\20260912-20260918  2026-09-12  2026-09-18
"""
import os, sys, re, shutil, sqlite3, datetime, calendar, subprocess, json
import win32com.client as win32
from dz_common import find_prev_f0, f0_path, f0_date, SUB_SHEETS, FX_BLOCKS, RULES

DB = r'E:\商务部\东智日常表更新\东智台账.db'
SRC = sys.argv[1] if len(sys.argv) > 1 else r'E:\商务部\东智日常表更新\20260912-20260918'
D1 = sys.argv[2] if len(sys.argv) > 2 else '2026-09-12'
D2 = sys.argv[3] if len(sys.argv) > 3 else '2026-09-18'

BASE, BASE_DATE = find_prev_f0(D1)          # 上周五基准 F0
# 可选第 4 参：显式指定基准（2026-11-06 新增）。新月首周的基准 = 新月底表，
# 但它的文件名日期是上月底，自动定位会取到"上月的 F0"→ 会把上月日列/累计带进本月。故允许显式传入。
if len(sys.argv) > 4 and sys.argv[4]:
    _b = sys.argv[4]
    if os.path.exists(_b):
        BASE, BASE_DATE = _b, (f0_date(_b) or '月初基准')
    else:
        print('  ! 指定基准不存在：%s（继续用自动定位）' % _b)
TARGET = f0_path(SRC, D2)                    # 本周五版 F0（写在周期目录里）

# 2026-11-23：重跑同一周时，用本周已产出的 F0 作基准。
# 否则会从上周末的 F0 重放，把本周已修正的补录/补行又打回去。
if not (len(sys.argv) > 4 and sys.argv[4]) and os.path.exists(TARGET) \
        and f0_date(TARGET) == D2 and os.path.abspath(TARGET) != os.path.abspath(BASE):
    BASE, BASE_DATE = TARGET, D2
    print('  重跑同一周：以现有 F0 为基准（%s）' % TARGET)

# 日列：签约表 4+日，其余表 3+日  →  本周五列
_DAY = int(D2.split('-')[2])
# 三张子表的日列都从 D 列（1日）开始：日列 = 3 + 日。
# 2026-11-11 修：签约表原来用 4+日，会把 6日 写进 7日，遇到已填底表就重复累计。
C_SIGN, C_OTHER = 3 + _DAY, 3 + _DAY
# 默认保留公式（只写输入格）；设 DZ_FLATTEN=1 可回到旧的“代码算值并清公式”模式。
KEEP_FORMULAS = os.environ.get('DZ_FLATTEN', '0') != '1'

# 三张子表的写入映射：业务字段 -> (sheet, 块起, 块止, 本周五列号)
WRITE_MAP = [
    ('签约金额', '订单签约跟踪表', 4,  85, C_SIGN),
    ('签约数量', '订单签约跟踪表', 87, 168, C_SIGN),
    ('产出数量', '产出发货跟踪表', 4,  69, C_OTHER),
    ('发货数量', '产出发货跟踪表', 71, 152, C_OTHER),
    ('发货金额', '产出发货跟踪表', 154, 235, C_OTHER),
    ('回款金额', '回款&开票统计表', 4,  94, C_OTHER),
    ('开票金额', '回款&开票统计表', 96, 178, C_OTHER),
]


def load_ledger(batch):
    """读台账：按客户聚合**本批次**数据（= 本次周更导入的行）+ 开票/回款明细（用于备注）。
    ⚠️ 必须按「导入批次」切周：台账是长期库，若不切周会把往月/往周的数据重复计进本周列
       （2026-09-18 实测：光缆回款被重复计成 1512.69 = 1163.04 + 349.66）。"""
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cond = "COALESCE(作废,0)=0 AND 导入批次=?"
    agg = {}
    cur.execute('''SELECT 客户名称,
        COALESCE(SUM(签约金额),0), COALESCE(SUM(签约数量),0), COALESCE(SUM(产出数量),0),
        COALESCE(SUM(发货数量),0), COALESCE(SUM(发货金额),0), COALESCE(SUM(回款金额),0), COALESCE(SUM(开票金额),0),
        COALESCE(SUM(收入),0)
        FROM 台账明细 WHERE ''' + cond + ' GROUP BY 客户名称', (batch,))
    for nm, qa, qq, oq, sq, sa, rq, kq, inc in cur.fetchall():
        agg[nm] = dict(签约金额=qa, 签约数量=qq, 产出数量=oq, 发货数量=sq, 发货金额=sa,
                       回款金额=rq, 开票金额=kq, 收入=inc)
    detail = {'开票': [], '回款': []}
    # 开票：按 (客户, 日期) 合并
    cur.execute("SELECT 时间,客户名称,开票金额 FROM 台账明细 WHERE " + cond +
                " AND 开票金额 IS NOT NULL AND 开票金额>0 ORDER BY 时间", (batch,))
    kp = {}
    for ds, nm, amt in cur.fetchall():
        kp[(nm, ds)] = kp.get((nm, ds), 0) + amt
    detail['开票'] = [(ds, nm, amt) for (nm, ds), amt in sorted(kp.items(), key=lambda x: x[0][1])]
    # 回款：按 (客户, 日期, 类型) 合并
    cur.execute("SELECT 时间,客户名称,回款金额,回款类型 FROM 台账明细 WHERE " + cond +
                " AND 回款金额 IS NOT NULL AND 回款金额>0 ORDER BY 时间", (batch,))
    hk = {}
    for ds, nm, amt, typ in cur.fetchall():
        hk[(nm, ds, typ)] = hk.get((nm, ds, typ), 0) + amt
    detail['回款'] = [(ds, nm, amt, typ) for (nm, ds, typ), amt in sorted(hk.items(), key=lambda x: x[0][1])]
    conn.close()
    return agg, detail


def ensure_payment_rows(wb, names):
    """回款日块缺客户行 → 在「其他类回款统计」前插一行，并同步分项汇总。

    2026-11-23：玖耀/宝鑫这类本周新出现的回款客户，基准 F0 里没有行。
    以前直接中止、要人工插行；现在自动补，且把 总计 的显式加总公式补上新行。
    """
    ws = wb.Worksheets('回款&开票统计表')
    bucket = None
    for r in range(4, 96):
        if str(ws.Cells(r, 2).Value or '').strip() == '其他类回款统计':
            bucket = r
            break
    if not bucket:
        return []
    cats = {}
    try:
        con = sqlite3.connect(DB)
        cats = {n: (c or '其他') for n, c in con.execute(
            "SELECT 客户名称, MAX(类别) FROM 台账明细 GROUP BY 客户名称")}
        con.close()
    except Exception:
        pass
    added = []
    for nm in names:
        ws.Rows(bucket).Insert(-4121, 0)          # xlDown, 沿用上一行格式
        ws.Cells(bucket, 1).Value = cats.get(nm, '其他')
        ws.Cells(bucket, 2).Value = nm
        for c in range(3, 35):
            ws.Cells(bucket, c).ClearContents()
        ws.Cells(bucket, 36).Formula = '=SUM(D%d:AI%d)' % (bucket, bucket)
        total = bucket + 2
        for c in range(3, 37):
            f = ws.Cells(total, c).Formula
            if not isinstance(f, str) or not f.startswith('='):
                continue
            col = re.sub(r'\d+', '', ws.Cells(total, c).Address.replace('$', ''))
            if re.fullmatch(r'=[A-Z]+\d+(\+[A-Z]+\d+)+', f) and ('%s%d' % (col, bucket)) not in f:
                ws.Cells(total, c).Formula = f + '+%s%d' % (col, bucket)

        fx = wb.Worksheets('分项汇总')
        fx_bucket = None
        for r in range(282, 365):
            if str(fx.Cells(r, 1).Value or '').strip() == '其他类回款统计':
                fx_bucket = r
                break
        if fx_bucket:
            fx.Rows(fx_bucket).Insert(-4121, 0)
            fx.Cells(fx_bucket, 1).Value = nm
            fx.Cells(fx_bucket, 2).Value = 0
            fx.Cells(fx_bucket, 3).Formula = "='回款&开票统计表'!C%d" % bucket
            fx.Cells(fx_bucket, 4).Formula = "='回款&开票统计表'!C%d" % bucket
            fx.Cells(fx_bucket, 5).Formula = "='回款&开票统计表'!AJ%d" % bucket
            for r in range(fx_bucket + 1, fx_bucket + 6):
                if str(fx.Cells(r, 1).Value or '').strip() in ('合计', '总计'):
                    for c in range(3, 6):
                        f = fx.Cells(r, c).Formula
                        m = re.fullmatch(r'=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', f or '')
                        if m:
                            fx.Cells(r, c).Formula = '=SUM(%s%s:%s%d)' % (
                                m.group(1), m.group(2), m.group(3), fx_bucket)
                    break
        added.append(nm)
        bucket += 1
    return added


def fix_refs(wb):
    """修正分项汇总「文件原有」引用错位（按客户名匹配 + 公式替换）"""
    ws = wb.Worksheets('分项汇总')
    fixed = []
    last = ws.UsedRange.Rows.Count
    for r in range(1, last + 1):
        a = str(ws.Cells(r, 1).Value or '')
        # 多模尔（自主接单）累计开票 AI141 → AI133
        if '多模尔' in a and '自主' in a:
            f = ws.Cells(r, 4).Formula
            if 'AI141' in str(f):
                ws.Cells(r, 4).Formula = str(f).replace('AI141', 'AI133')
                fixed.append(('多模尔 累计开票 AI141→AI133', r))
        # 云鹏锡业 本月开票计划 C164 → C161
        if '云鹏锡业' in a:
            f = ws.Cells(r, 3).Formula
            if 'C164' in str(f):
                ws.Cells(r, 3).Formula = str(f).replace('C164', 'C161')
                fixed.append(('云鹏锡业 开票计划 C164→C161', r))
    # T01 收入系数统一为 0.019（用户 2026-09-18 定"按台账算"；原公式写的 0.02）
    for r in range(1, last + 1):
        for c in range(2, 12):
            f = ws.Cells(r, c).Formula
            if isinstance(f, str) and '/1.13*0.02' in f:
                ws.Cells(r, c).Formula = f.replace('/1.13*0.02', '/1.13*0.019')
                fixed.append(('T01 收入系数 0.02→0.019', r))

    # 汇总「预计待完成」C6 套 MAX(0,…)：完成量超计划时待完成应为 0，不该出负数
    hz = wb.Worksheets('汇总')
    for r in range(4, 12):
        f = hz.Cells(r, 6).Formula
        if isinstance(f, str) and f.startswith('=') and not f.upper().startswith('=MAX('):
            hz.Cells(r, 6).Formula = '=MAX(0,' + f[1:] + ')'
            fixed.append(('汇总 待完成 MAX(0,…)', r))

    # 出货块「预计收入」K 列：按"实际累计 vs 预计总出货"取大（用户 2026-09-18："按实际的来"）
    #   · 原 K=0 → 保持 0（光缆/特发信息：收入计 0）
    #   · T01   → MAX(E,J)/1.13*0.019
    #   · 原 =J/1.13 → =MAX(E,J)/1.13（未超计划的客户结果不变）
    for r in range(137, 205):
        nm_ = str(ws.Cells(r, 1).Value or '').strip()
        if not nm_ or '合计' in nm_:
            continue
        f = ws.Cells(r, 11).Formula
        fs = f if isinstance(f, str) else str(f)
        if nm_ == 'T01' and fs != '=MAX(E%d,J%d)/1.13*0.019' % (r, r):
            ws.Cells(r, 11).Formula = '=MAX(E%d,J%d)/1.13*0.019' % (r, r)
            fixed.append(('预计收入 T01 MAX(E,J)*0.019', r))
        elif re.match(r'^=J\d+/1\.13$', fs):
            ws.Cells(r, 11).Formula = '=MAX(E%d,J%d)/1.13' % (r, r)
            fixed.append(('预计收入 MAX(E,J)', r))
    return fixed


def load_month_cum(month_prefix, through=None):
    """台账 → 本月累计（每客户每线 + 收入）。month_prefix 如 '202609'。
    按「归属月」切月；through='YYYY-MM-DD' 时只算到该日（做表日之后的不进本表）。"""
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    agg = {}
    cur.execute('''SELECT 客户名称,
        COALESCE(SUM(签约金额),0), COALESCE(SUM(签约数量),0), COALESCE(SUM(产出数量),0),
        COALESCE(SUM(发货数量),0), COALESCE(SUM(发货金额),0), COALESCE(SUM(收入),0),
        COALESCE(SUM(回款金额),0), COALESCE(SUM(开票金额),0)
        FROM 台账明细
        WHERE COALESCE(作废,0)=0
          AND COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?
          AND (? IS NULL OR COALESCE(NULLIF(时间,''),'')<=?)
        GROUP BY 客户名称''', (month_prefix[:4] + '-' + month_prefix[4:6], through, through))
    for nm, qa, qq, oq, sq, sa, inc, rq, kq in cur.fetchall():
        agg[nm] = dict(签约金额=qa, 签约数量=qq, 产出数量=oq, 发货数量=sq, 发货金额=sa,
                       收入=inc, 回款金额=rq, 开票金额=kq)
    conn.close()
    return agg


def read_aj(wb):
    """读 F0 各块「本月累计」AJ(36) 列 → 写周增量的基线。"""
    out = {}
    for field, sheet, r1, r2, col in WRITE_MAP:
        ws = wb.Worksheets(sheet)
        d = {}
        cumcol = 35 if field == '开票金额' else 36   # 开票块止于 AI(35)
        for r in range(r1, r2 + 1):
            nm = str(ws.Cells(r, 2).Value or '').strip()
            if not nm or nm in d:
                continue
            v = ws.Cells(r, cumcol).Value
            d[nm] = float(v) if isinstance(v, (int, float)) else 0.0
        out[field] = d
    return out


def fix_titles(wb):
    """分项汇总 6 处「待做」列标题动态化（用户 2026-09-20 定）：
    非月末做表 → 「{做表日+1}-{月底}日待XXX」；月末做表（做表日==月底）→「本月待XXX」。"""
    y, m, dd = (int(x) for x in D2.split('-'))
    month_end = calendar.monthrange(y, m)[1]
    prefix = '本月' if dd == month_end else '%d/%d-%d/%d日' % (m, dd + 1, m, month_end)
    ws = wb.Worksheets('分项汇总')
    for ref, label in RULES['TITLE_待']:
        ws.Range(ref).Value = prefix + label
    return prefix


def repair_template(wb):
    """把分项汇总「签约块」里历史手工写死的派生格恢复成标准公式（已有公式不动）。"""
    fx = wb.Worksheets('分项汇总')
    std = {
        6: '=IFERROR(E{r}/C{r},"-")',
        7: '=B{r}-D{r}',
        8: '=C{r}-E{r}',
        9: '=D{r}+G{r}',
        10: '=E{r}+H{r}',
        11: '=IFERROR(J{r}/C{r},"-")',
    }
    n = 0
    for rr in range(5, 72):
        a = str(fx.Cells(rr, 1).Value or '').strip()
        if not a or '合计' in a or '统计' in a:
            continue
        for c, t in std.items():
            f = fx.Cells(rr, c).Formula
            if not (isinstance(f, str) and f.startswith('=')) and f is not None:
                fx.Cells(rr, c).Formula = t.format(r=rr)
                n += 1
    for rr in range(1, fx.UsedRange.Rows.Count + 1):
        a = fx.Cells(rr, 1).Value
        if isinstance(a, str) and '（自主接单）' in a:
            fx.Cells(rr, 1).Value = a.replace('（自主接单）', '')
    # 收入「月度预计」沿用既有口径：按台账实际收入展示（不是出货推算）
    wb.Worksheets('汇总').Cells(9, 7).Formula = '=D9'
    return n


def compute_hz(wb, cum):
    """汇总 sheet 8 行：代码算值并贴值（清公式），不依赖分项汇总公式。
    - 实际(D)：台账 SUM（真源，含新增客户）；产出额外加「月初基线」（台账产出是每日增量，缺月初累计）。
    - 计划(C)：读分项汇总合计行（公式已重算，含全部客户，避免读子表时混入「总计/实际发货」等特殊行）。
    - 收入：台账（子表无收入列）。"""
    def N(v):
        if v in (None, ''):
            return 0.0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.0

    def RATE(numer, denom):
        return '-' if not denom else numer / denom

    def sum_cum(field):
        return sum(c.get(field, 0.0) for c in cum.values())

    def output_base():
        """产出月初基线 = 产出块客户行「日期 < 台账本月最早产出日期」的日列合计（如 9/4 的 1787）。
        ⚠️ 2026-11-06 修：必须按【本月】取最早日期。原来取全库最早（9月），ed=9 → 会把本周
           写在「6日」列的数当成"月初基线"再加一遍（实测产出 17269 被加成 34538）。"""
        sheet, r1, r2, _, _ = SUB_SHEETS['产出数量']
        ws = wb.Worksheets(sheet)
        conn = sqlite3.connect(DB)
        earliest = conn.execute(
            "SELECT MIN(时间) FROM 台账明细 WHERE COALESCE(产出数量,0)>0 AND COALESCE(作废,0)=0"
            " AND COALESCE(NULLIF(归属月,''), substr(时间,1,7))=?", (D2[:7],)).fetchone()[0]
        conn.close()
        base = 0.0
        if earliest:
            ed = int(earliest.split('-')[2])
            for r in range(r1, r2 + 1):
                b = str(ws.Cells(r, 2).Value or '').strip()
                if b not in cum:            # 只读台账客户行，跳过总计/实际发货等特殊行
                    continue
                for d in range(1, ed):      # 日列 = 3 + 日（1日=D=4）
                    base += N(ws.Cells(r, 3 + d).Value)
        return base

    CUM = {
        '签约数量': sum_cum('签约数量'),
        '签约金额': sum_cum('签约金额'),
        '产出数量': sum_cum('产出数量') + output_base(),
        '发货数量': sum_cum('发货数量'),
        '发货金额': sum_cum('发货金额'),
        '开票金额': sum_cum('开票金额'),
        '回款金额': sum_cum('回款金额'),
    }

    fx = wb.Worksheets('分项汇总')
    hz = wb.Worksheets('汇总')
    # (汇总行, 分项汇总计划合计格, 台账累计字段)
    # 2026-10-22 修正：汇总表 R4=签约金额、R5=签约数量（原来写反了，导致金额列填成数量）；
    #                  回款计划合计格 = C356（原写 D357 是空行，回款计划被清零）。
    HZ = [
        (4,  'C72',  '签约金额'),
        (5,  'B72',  '签约数量'),
        (6,  'B133', '产出数量'),
        (7,  'B207', '发货数量'),
        (8,  'C207', '发货金额'),
        (10, 'C278', '开票金额'),
        (11, 'D356', '回款金额'),
    ]
    for r, plan_ref, cum_field in HZ:
        P = N(fx.Range(plan_ref).Value)     # C 计划（读分项汇总合计，公式已算好）
        D = CUM[cum_field]                  # D 实际（台账真源）
        F = max(0.0, P - D)                 # F 待完成 = MAX(0, 计划-累计)
        hz.Cells(r, 1).Value = r - 3
        hz.Cells(r, 3).Value = P
        hz.Cells(r, 4).Value = D
        hz.Cells(r, 5).Value = RATE(D, P)
        hz.Cells(r, 6).Value = F
        hz.Cells(r, 7).Value = D + F

    # R9 收入：计划 C / 待完成 F 是手填历史值（读现状固化不清值）；实际 D / 月度预计 G 用台账收入
    inc = sum_cum('收入')
    C9 = N(hz.Cells(9, 3).Value)
    hz.Cells(9, 1).Value = 6
    hz.Cells(9, 4).Value = inc
    hz.Cells(9, 5).Value = RATE(inc, C9)
    hz.Cells(9, 7).Value = inc
    hz.Cells(9, 3).Value = C9
    hz.Cells(9, 6).Value = hz.Cells(9, 6).Value
    return dict(收入=inc)


def compute_fx(wb, cum):
    """分项汇总 5 块：代码算值贴值（清公式）。
    - 计划：读子表计划列（块⑥填的静态值）；累计：台账（base 映射，真源）；派生：代码算。
    - 客户名统一去「（自主接单）」后缀（用户 2026-09-20：同一客户、以台账为准）。
    - 「两行」处理：同一 base 若同时有「X」与「X（自主接单）」，则带后缀行填台账值、不带后缀占位行填 0。
    - 特殊行（其他统计 R277/R356、收入 R208）不在此处理，由 flatten_sheet 兜底固化。"""
    fx = wb.Worksheets('分项汇总')

    def N(v):
        if v in (None, ''):
            return 0.0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.0

    def RATE(n, d):
        return '-' if not d else n / d

    def base_of(s):
        return str(s or '').split('（')[0].strip()

    def income_val(nm, E, J):
        base = base_of(nm)
        if base in RULES['INCOME_ZERO']:
            return 0.0
        if nm == 'T01':
            return max(E, J) / RULES['VAT'] * RULES['INCOME_RATE']['T01']
        return max(E, J) / RULES['VAT']

    def sub_plan(line, base):
        sheet, r1, r2, plan_col, _ = SUB_SHEETS[line]
        ws = wb.Worksheets(sheet)
        for r in range(r1, r2 + 1):
            if str(ws.Cells(r, 2).Value or '').strip() == base:
                return N(ws.Cells(r, plan_col).Value)
        return 0.0

    def dup_bases(r1, r2):
        d = {}
        for r in range(r1, r2 + 1):
            a = str(fx.Cells(r, 1).Value or '').strip()
            if a and '合计' not in a and '统计' not in a:
                d.setdefault(base_of(a), []).append(r)
        return {b: rs for b, rs in d.items() if len(rs) > 1}

    def sumcol(col, r1, r2):
        return sum(N(fx.Cells(r, col).Value) for r in range(r1, r2 + 1))

    def is_ph(base, nm, dups):
        return base in dups and '自主接单' not in nm

    # 产出月初基线（台账产出是每日增量，缺月初累计，如 9/4 的 1787）——按子表客户名预计算
    # ⚠️ 2026-11-06：与 compute_hz.output_base 同源，必须限定【本月】，否则把本周数当基线重复加
    conn = sqlite3.connect(DB)
    earliest = conn.execute(
        "SELECT MIN(时间) FROM 台账明细 WHERE COALESCE(产出数量,0)>0 AND COALESCE(作废,0)=0"
        " AND 导入批次 LIKE ?", (D2[:7].replace('-', '') + '%',)).fetchone()[0]
    conn.close()
    base_cust = {}
    if earliest:
        ed = int(earliest.split('-')[2])
        sheet_o, r1_o, r2_o, _, _ = SUB_SHEETS['产出数量']
        ws_out = wb.Worksheets(sheet_o)
        for r in range(r1_o, r2_o + 1):
            b = str(ws_out.Cells(r, 2).Value or '').strip()
            if b:
                base_cust[b] = sum(N(ws_out.Cells(r, 3 + d).Value) for d in range(1, ed))

    # ---- 块1 订单计划 ----
    r1, r2, tot = FX_BLOCKS['订单计划']
    dups = dup_bases(r1, r2)
    for r in range(r1, r2 + 1):
        nm = str(fx.Cells(r, 1).Value or '').strip()
        if not nm or '合计' in nm or '统计' in nm:
            continue
        base = base_of(nm); ph = is_ph(base, nm, dups); c = cum.get(base, {})
        # 计划口径（2026-11-06 修）：优先读子表计划列；子表为空时【保留分项汇总原值】。
        # 新月底表的计划是直接填在分项汇总里的（子表计划列空），原来一律取子表 → 计划被清零，
        # 汇总 8 行计划全变 0（实测签约计划 722.60 被算成 0）。
        B = sub_plan('签约数量', base) or N(fx.Cells(r, 2).Value)
        C = sub_plan('签约金额', base) or N(fx.Cells(r, 3).Value)
        D = 0.0 if ph else c.get('签约数量', 0.0)
        E = 0.0 if ph else c.get('签约金额', 0.0)
        G, H = B - D, C - E
        I, J = D + G, E + H
        for col, v in ((2, B), (3, C), (4, D), (5, E), (6, RATE(E, C)), (7, G), (8, H), (9, I), (10, J), (11, RATE(J, C))):
            fx.Cells(r, col).Value = v
    for col in (2, 3, 4, 5, 7, 8, 9, 10):
        fx.Cells(tot, col).Value = sumcol(col, r1, r2)
    fx.Cells(tot, 6).Value = '-'
    fx.Cells(tot, 11).Value = '-'

    # ---- 块2 产出计划 ----
    r1, r2, tot = FX_BLOCKS['产出计划']
    dups = dup_bases(r1, r2)
    for r in range(r1, r2 + 1):
        nm = str(fx.Cells(r, 1).Value or '').strip()
        if not nm or '合计' in nm or '统计' in nm:
            continue
        base = base_of(nm); ph = is_ph(base, nm, dups); c = cum.get(base, {})
        B = sub_plan('产出数量', base) or N(fx.Cells(r, 2).Value)
        C = 0.0 if ph else (c.get('产出数量', 0.0) + base_cust.get(base, 0.0))
        E, F = B - C, C + (B - C)
        for col, v in ((2, B), (3, C), (4, RATE(C, B)), (5, E), (6, F), (7, RATE(F, B))):
            fx.Cells(r, col).Value = v
    for col in (2, 3, 5, 6):
        fx.Cells(tot, col).Value = sumcol(col, r1, r2)
    fx.Cells(tot, 4).Value = RATE(N(fx.Cells(tot, 3).Value), N(fx.Cells(tot, 2).Value))
    fx.Cells(tot, 7).Value = RATE(N(fx.Cells(tot, 6).Value), N(fx.Cells(tot, 2).Value))

    # ---- 块3 出货计划 ----
    r1, r2, tot = FX_BLOCKS['出货计划']
    dups = dup_bases(r1, r2)
    for r in range(r1, r2 + 1):
        nm = str(fx.Cells(r, 1).Value or '').strip()
        if not nm or '合计' in nm or '统计' in nm:
            continue
        base = base_of(nm); ph = is_ph(base, nm, dups); c = cum.get(base, {})
        B = sub_plan('发货数量', base) or N(fx.Cells(r, 2).Value)
        C = sub_plan('发货金额', base) or N(fx.Cells(r, 3).Value)
        D = 0.0 if ph else c.get('发货数量', 0.0)
        E = 0.0 if ph else c.get('发货金额', 0.0)
        G, H = B - D, C - E
        I, J = D + G, E + H
        K = income_val(nm, E, J)
        for col, v in ((2, B), (3, C), (4, D), (5, E), (6, RATE(E, C)), (7, G), (8, H), (9, I), (10, J), (11, K)):
            fx.Cells(r, col).Value = v
    for col in (2, 3, 4, 5, 7, 8, 9, 10, 11):
        fx.Cells(tot, col).Value = sumcol(col, r1, r2)
    fx.Cells(tot, 6).Value = RATE(N(fx.Cells(tot, 5).Value), N(fx.Cells(tot, 3).Value))

    # ---- 块4 开票计划 ----
    r1, r2, tot = FX_BLOCKS['开票计划']
    dups = dup_bases(r1, r2)
    for r in range(r1, r2 + 1):
        nm = str(fx.Cells(r, 1).Value or '').strip()
        if not nm or '合计' in nm or '统计' in nm:
            continue
        base = base_of(nm); ph = is_ph(base, nm, dups); c = cum.get(base, {})
        C = sub_plan('开票金额', base) or N(fx.Cells(r, 3).Value)
        D = 0.0 if ph else c.get('开票金额', 0.0)
        F, G = C - D, (C - D) + D
        for col, v in ((3, C), (4, D), (5, RATE(D, C)), (6, F), (7, G)):
            fx.Cells(r, col).Value = v
    for col in (3, 4, 6, 7):
        fx.Cells(tot, col).Value = sumcol(col, r1, r2)
    fx.Cells(tot, 5).Value = RATE(N(fx.Cells(tot, 4).Value), N(fx.Cells(tot, 3).Value))

    # ---- 块5 回款计划 ----
    r1, r2, tot = FX_BLOCKS['回款计划']
    dups = dup_bases(r1, r2)
    for r in range(r1, r2 + 1):
        nm = str(fx.Cells(r, 1).Value or '').strip()
        if not nm or '合计' in nm or '统计' in nm:
            continue
        base = base_of(nm); ph = is_ph(base, nm, dups); c = cum.get(base, {})
        D = sub_plan('回款金额', base) or N(fx.Cells(r, 4).Value)
        E = 0.0 if ph else c.get('回款金额', 0.0)
        G, H = D - E, E + (D - E)
        for col, v in ((4, D), (5, E), (6, RATE(E, D)), (7, G), (8, H), (9, RATE(H, D))):
            fx.Cells(r, col).Value = v
    for col in (4, 5, 7, 8):
        fx.Cells(tot, col).Value = sumcol(col, r1, r2)
    fx.Cells(tot, 6).Value = RATE(N(fx.Cells(tot, 5).Value), N(fx.Cells(tot, 4).Value))
    fx.Cells(tot, 9).Value = RATE(N(fx.Cells(tot, 8).Value), N(fx.Cells(tot, 4).Value))

    # 客户名统一：删「（自主接单）」后缀（同一客户，以台账为准）
    n_rename = 0
    for r in range(1, fx.UsedRange.Rows.Count + 1):
        a = fx.Cells(r, 1).Value
        if isinstance(a, str) and '（自主接单）' in a:
            fx.Cells(r, 1).Value = a.replace('（自主接单）', '')
            n_rename += 1
    return dict(改名=n_rename)


def load_batch_meta():
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out', '批次.json')
    if not os.path.exists(p):
        return None
    try:
        with open(p, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def set_custom_prop(wb, name, value):
    """把批次水印写进 Excel 自定义文档属性（不可见，不改表结构）。"""
    try:
        props = wb.CustomDocumentProperties
        try:
            props.Item(name).Value = value
        except Exception:
            props.Add(name, False, 4, value)   # 4 = msoPropertyTypeString
        return True
    except Exception as e:
        print('  ! 自定义属性写入失败：%s' % e)
        return False


def record_weekly_checks(meta, month, exp, got):
    """把「台账 → F0 汇总」的 7 项校验写进 校验结果（监督模块留痕）。"""
    if not meta:
        return
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    label = {4: '签约金额', 5: '签约数量', 6: '产出数量', 7: '发货数量',
             8: '发货金额', 10: '开票金额', 11: '回款金额'}
    for r, want in exp.items():
        g = got.get(r) or 0
        diff = float(g) - float(want)
        cur.execute('''INSERT OR REPLACE INTO 校验结果
            (批次ID,月份,环节,检查项,源值,台账值,看板值,差异,容差,结论,详情,生成时间)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,datetime('now','localtime'))''',
            (meta['批次ID'], month, '周更F0', label.get(r, 'D%d' % r),
             None, round(float(want), 10), round(float(g), 10), round(diff, 10), 0.0005,
             'PASS' if abs(diff) <= 0.0005 else 'FAIL', '汇总!D%d' % r))
    conn.commit()
    conn.close()


def write_f0_watermark(wb, meta):
    """批次水印只进 DB + Excel 自定义属性，不上看板/备注格。"""
    if not meta:
        return
    ok = set_custom_prop(wb, 'DZ批次水印', meta.get('水印', ''))
    set_custom_prop(wb, 'DZ批次ID', meta.get('批次ID', ''))
    set_custom_prop(wb, 'DZ数据截止', meta.get('数据截止', ''))
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute('''INSERT OR REPLACE INTO 批次水印(对象,批次ID,水印,数据截止,说明)
                   VALUES(?,?,?,?,?)''',
                ('F0周更', meta['批次ID'], meta['水印'], meta.get('数据截止', ''),
                 'Excel自定义属性=%s' % ('已写' if ok else '未写')))
    conn.commit()
    conn.close()
    print('  批次水印：%s（DB%s）' % (meta['水印'], ' + Excel属性' if ok else ''))


def flatten_sheet(wb, sheet, max_col=12):
    """把某 sheet 里残留的公式单元格读值写回（清公式、值不变）。代码算不到的边角公式用这个兜底。"""
    ws = wb.Worksheets(sheet)
    n = 0
    last = ws.UsedRange.Rows.Count
    for r in range(1, last + 1):
        for c in range(1, max_col + 1):
            cell = ws.Cells(r, c)
            if isinstance(cell.Formula, str) and cell.Formula.startswith('='):
                cell.Value = cell.Value
                n += 1
    return n


def fmt_date(ds):
    """2026-09-10 → 9/10"""
    try:
        y, m, d = ds.split('-')
        return '%d/%d' % (int(m), int(d))
    except Exception:
        return ds


def main():
    BATCH = os.path.basename(os.path.normpath(SRC))
    agg, detail = load_ledger(BATCH)
    print('台账聚合（批次 %s）客户数：%d' % (BATCH, len(agg)))
    for k, v in sorted(agg.items()):
        nz = {a: round(b, 6) for a, b in v.items() if b}
        if nz:
            print('   %-12s %s' % (k, nz))

    # 2026-11-23：一律先写临时文件，成功才覆盖 TARGET。
    # 旧逻辑在失败时也会把 BASE 覆盖到 TARGET，重跑一次就把本周 F0 打回上周。
    BUILD = TARGET + '.building'
    if os.path.abspath(BASE) == os.path.abspath(TARGET):
        shutil.copy2(TARGET, BUILD)
        print('就地更新：先写临时文件 %s' % BUILD)
    else:
        shutil.copy2(BASE, BUILD)
        print('已复制基准 → 临时文件 %s' % BUILD)

    xl = win32.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    XL_MANUAL, XL_AUTO = -4135, -4105
    # 2026-11-06：默认自动重算下，每写一格就触发全表 7000+ 公式重算 → Excel 频繁崩
    # （实测 pywintypes.com_error "远程过程调用失败"）。批量写值期间一律手动计算，只在需要时重算。
    # 注意：Calculation 属性必须在【打开工作簿之后】才能设置，否则报"不能设置 Application 的 Calculation 属性"。
    wb = None
    ok = False
    try:
        wb = xl.Workbooks.Open(BUILD)
        xl.Calculation = XL_MANUAL

        # 2026-11-23：写「到本表日为止的月累计 - 基准 F0 已有累计」的差额。
        # 这样补录/跨日期回补/重跑都幂等：基准里已有的不再写第二遍。
        cum_all = load_month_cum(D2[:7].replace('-', ''), through=D2)
        base_cum = read_aj(wb)
        # 本表日那一列马上要清空重写，基线必须扣掉它，否则 delta 会被算成 0。
        for field, sheet, r1, r2, col in WRITE_MAP:
            ws_b = wb.Worksheets(sheet)
            seen = set()
            for r in range(r1, r2 + 1):
                nm = str(ws_b.Cells(r, 2).Value or '').strip()
                if not nm or nm in seen:
                    continue
                seen.add(nm)
                v = ws_b.Cells(r, col).Value
                if isinstance(v, (int, float)):
                    base_cum[field][nm] = base_cum[field].get(nm, 0.0) - float(v)

        def delta(nm, field):
            land = ('其他类回款统计' if (field == '回款金额' and nm in RULES['F0_OTHER_BUCKET'])
                    else nm)
            return ((cum_all.get(nm, {}).get(field, 0) or 0)
                    - (base_cum.get(field, {}).get(land, 0) or 0))

        # 2.0) 回款日块缺客户行 → 自动补（插在「其他类回款统计」前，并同步分项汇总）
        ws_hk0 = wb.Worksheets('回款&开票统计表')
        names_hk0 = {str(ws_hk0.Cells(r, 2).Value or '').strip() for r in range(4, 96)}
        missing_pay = [nm for nm in cum_all
                       if delta(nm, '回款金额') and nm not in names_hk0
                       and nm not in RULES['F0_OTHER_BUCKET']]
        if missing_pay:
            added_rows = ensure_payment_rows(wb, missing_pay)
            print('  ✓ 回款块自动补行：%s' % '、'.join(added_rows))

        # 2) 先清本周日列输入值（只清数值格，不动公式），再按台账写入
        cleared = 0
        n = 0
        for field, sheet, r1, r2, col in WRITE_MAP:
            ws = wb.Worksheets(sheet)
            for r in range(r1, r2 + 1):
                cell = ws.Cells(r, col)
                if not cell.HasFormula and isinstance(cell.Value, (int, float)):
                    cell.ClearContents()
                    cleared += 1
        for field, sheet, r1, r2, col in WRITE_MAP:
            ws = wb.Worksheets(sheet)
            for nm in cum_all:
                v = delta(nm, field)
                if abs(v) < 1e-9:
                    continue
                land = ('其他类回款统计' if (field == '回款金额' and nm in RULES['F0_OTHER_BUCKET'])
                        else nm)
                for r in range(r1, r2 + 1):
                    if str(ws.Cells(r, 2).Value or '').strip() == land:
                        ws.Cells(r, col).Value = v
                        n += 1
                        break
        print('清空日列 %d 格，已填 %d 格' % (cleared, n))

        # 覆盖校验（2026-11-23）：台账有数、子表找不到客户行 → 直接失败，禁止静默丢弃
        unmapped = []
        for field, sheet, r1, r2, col in WRITE_MAP:
            ws_chk = wb.Worksheets(sheet)
            names = {str(ws_chk.Cells(r, 2).Value or '').strip() for r in range(r1, r2 + 1)}
            for nm in cum_all:
                v = delta(nm, field)
                land = ('其他类回款统计' if (field == '回款金额' and nm in RULES['F0_OTHER_BUCKET'])
                        else nm)
                if abs(v) > 1e-9 and land not in names:
                    unmapped.append((field, nm, round(v, 6)))
        if unmapped:
            for field, nm, v in unmapped:
                print('  [无落点] %-6s %-16s %s' % (field, nm, v))
            raise SystemExit('台账客户在子表里没有对应行，已中止（未保存）。请先用 add_client_rows.py 加行')

        # 3) 修正引用错位
        print('  fix_refs 修正引用中…', flush=True)
        fixed = fix_refs(wb)
        for msg, r in fixed:
            print('  ✓ 修正引用：%s (R%d)' % (msg, r))
        repaired = repair_template(wb)
        if repaired:
            print('  ✓ 修复历史写死的派生格：%d 处' % repaired)

        # 4) 重算
        print('  重算中（CalculateFullRebuild）…', flush=True)
        xl.Calculation = XL_AUTO
        xl.CalculateFullRebuild()
        print('  重算完成', flush=True)
        xl.Calculation = XL_MANUAL          # 后续代码写值阶段继续手动，避免反复全表重算

        # 4.5) 标题动态化；默认保留公式，只重算
        prefix = fix_titles(wb)
        cum = cum_all
        if KEEP_FORMULAS:
            xl.Calculation = XL_AUTO
            xl.CalculateFullRebuild()
            xl.Calculation = XL_MANUAL
            print('标题前缀：%s ｜ 保留公式模式（分项汇总/汇总不固化）' % prefix)
        else:
            print('  分项汇总 5 块代码算…', flush=True)
            fx_info = compute_fx(wb, cum)
            print('  汇总 8 行代码算…', flush=True)
            hz_info = compute_hz(wb, cum)
            print('标题前缀：%s ｜ 分项汇总改名 %d 处 ｜ 汇总收入=%r' % (prefix, fx_info['改名'], hz_info['收入']))

        # 5) 备注
        ws_hz = wb.Worksheets('汇总')
        # 读代码算后的汇总表 D 列（实际完成）
        D = {r: ws_hz.Cells(r, 4).Value for r in range(4, 12)}
        if KEEP_FORMULAS:
            # 汇总 D 列是「本月累计」，不是「本批次增量」；第二周起必须用月度累计比对
            # 收入(9) 不由台账写值（分项汇总 E208 纯公式，客户范围与台账不同），比对会误报
            # 回款(11) 只比「F0 回款块里有行的客户」；F0 没有客户行的回款（如 2026-11 玖耀）
            # 只告警、不阻断——那是模板缺行，不是本次写入漏数
            ws_hk = wb.Worksheets('回款&开票统计表')
            names_hk = {str(ws_hk.Cells(r, 2).Value or '').strip() for r in range(4, 95)}
            unmapped_r = sorted(k for k, v in cum.items()
                                if k not in names_hk and (v.get('回款金额') or 0))
            exp = {
                4: sum(v.get('签约金额', 0) or 0 for v in cum.values()),
                5: sum(v.get('签约数量', 0) or 0 for v in cum.values()),
                6: sum(v.get('产出数量', 0) or 0 for v in cum.values()),
                7: sum(v.get('发货数量', 0) or 0 for v in cum.values()),
                8: sum(v.get('发货金额', 0) or 0 for v in cum.values()),
                10: sum(v.get('开票金额', 0) or 0 for v in cum.values()),
                11: sum((v.get('回款金额', 0) or 0) for k, v in cum.items()
                        if k in names_hk or k in RULES['F0_OTHER_BUCKET']),
            }
            bad = []
            for r, want in exp.items():
                got = D.get(r) or 0
                if abs(got - want) > max(0.0005, abs(want) * 1e-9):
                    bad.append((r, got, want))
            if bad:
                for r, got, want in bad:
                    print('  [校验失败] 汇总D%d 实际=%r 台账=%r' % (r, got, want))
                raise SystemExit('汇总与台账不一致，已中止（未保存）')
            if unmapped_r:
                print('  [口径提示] 以下客户有回款但 F0 回款块无客户行，未计入汇总：%s'
                      % {k: round(cum[k].get('回款金额', 0) or 0, 6) for k in unmapped_r})
            inc_diff = (sum(v.get('收入', 0) or 0 for v in cum.values())
                        - (D.get(9) or 0))
            if abs(inc_diff) > 0.0005:
                print('  [口径提示] 汇总D9(收入)=%r 与台账收入合计差 %r（D9 为分项汇总 E208 公式口径，非台账写值）'
                      % (D.get(9), round(inc_diff, 6)))
            print('  ✓ 校验通过：汇总 8 项与台账一致')
            record_weekly_checks(load_batch_meta(), D2[:7], exp, D)

        def cell_by_cust(ws, name, col, r1, r2_):
            """按客户名找行再取值（不要写死行号：插行会让行号漂移）"""
            for r in range(r1, r2_ + 1):
                if str(ws.Cells(r, 2).Value or '').strip() == name:
                    return ws.Cells(r, col).Value
            return None

        # 光缆事业部：改用【台账本月累计】取数（2026-11-06 修）。
        # 原来读子表累计列，光缆那行累计是公式判空返回的 "-"，备注会写成"计划外签约-万元"。
        gl = cum.get('光缆事业部', {})
        gl_qy = gl.get('签约金额', 0.0) or 0.0
        gl_kp = gl.get('开票金额', 0.0) or 0.0
        gl_hk = gl.get('回款金额', 0.0) or 0.0
        r2 = lambda v: round(v, 2) if isinstance(v, (int, float)) else v
        ws_hz.Cells(4, 8).Value = '本月已签约%s万元；\n    光缆事业部（自主接单）计划外签约%s万元；' % (r2(D[4]), r2(gl_qy))
        ws_hz.Cells(6, 8).Value = '本月已产出%d台；' % int(D[6])
        ws_hz.Cells(7, 8).Value = '本月已发货%d台，发货金额%s万元；' % (int(D[7]), r2(D[8]))
        ws_hz.Cells(9, 8).Value = '本月收入约%s万元（预估）；' % r2(D[9])
        ws_hz.Cells(10, 8).Value = '本月已开票%s万元；\n     光缆事业部（自主接单）已开票%s万元;' % (r2(D[10]), r2(gl_kp))
        ws_hz.Cells(11, 8).Value = '本月已回款%s万元（含互抵回款0万元）；\n    光缆事业部已回款（自主接单）%s万元现金；' % (r2(D[11]), r2(gl_hk))

        # 分项汇总备注：开票 H(8) 列在「块4 开票计划」；回款 J(10) 列在「块5 回款计划」
        ws_fx = wb.Worksheets('分项汇总')

        def append_note(r, col, txt):
            old = ws_fx.Cells(r, col).Value
            if isinstance(old, str) and old.strip():
                if txt in old:                      # 幂等：重跑不重复写
                    return False
                ws_fx.Cells(r, col).Value = old + '，' + txt
            else:
                ws_fx.Cells(r, col).Value = txt
            return True

        def base_of(s):
            return str(s or '').split('（')[0].strip()

        def pick_rows(r1, r2_):
            """{客户基名: 行号}；同一客户同有「X」与「X（自主接单）」时以（自主接单）为准，避免重复写"""
            plain, auto = {}, {}
            for r in range(r1, r2_ + 1):
                a = str(ws_fx.Cells(r, 1).Value or '').strip()
                if not a or '合计' in a or '统计' in a:
                    continue
                if '自主接单' in a:
                    auto[base_of(a)] = r
                else:
                    plain[a] = r
            out = dict(plain)
            out.update(auto)
            return out

        # ⚠️ 不要只写「自主接单」行：普客（如浪潮/鼎泰恒华/T01）也要写备注（2026-09-18 用户指出）
        for nm_, r in pick_rows(210, 278).items():
            for ds, nm, amt in detail['开票']:
                if nm == nm_:
                    append_note(r, 8, '%s开票%s万元' % (fmt_date(ds), r2(amt)))
        for nm_, r in pick_rows(281, 360).items():
            for ds, nm, amt, typ in detail['回款']:
                if nm == nm_:
                    append_note(r, 10, '%s回款%s万元%s' % (fmt_date(ds), r2(amt), typ))

        # 6) 旧模式才固化公式；保留公式模式跳过
        if not KEEP_FORMULAS:
            n_flat_fx = flatten_sheet(wb, '分项汇总')
            n_flat_hz = flatten_sheet(wb, '汇总')
            print('固化残留公式：分项汇总 %d 处 / 汇总 %d 处' % (n_flat_fx, n_flat_hz))

        write_f0_watermark(wb, load_batch_meta())
        xl.Calculation = XL_AUTO           # 保存前恢复自动计算，文件交给别人打开时行为不变
        wb.Save()
        ok = True
        print('已重算、写备注并保存（临时文件）')
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass
    if not ok:
        print('FAILED：原 F0 未改动 %s（临时文件保留在 %s）' % (TARGET, BUILD))
        raise SystemExit(1)
    os.replace(BUILD, TARGET)
    print('DONE, size=%d' % os.path.getsize(TARGET))

    # 7) 三层闭合校验（源文件 → 台账 → F0）；FAIL 即禁止出表
    if os.environ.get('DZ_GATE', '1') != '0':
        gate = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dz_gate.py')
        print('\n闭合校验（dz_gate）…', flush=True)
        rc = subprocess.call([sys.executable, gate, SRC, D1, D2],
                             env=dict(os.environ, PYTHONIOENCODING='utf-8'))
        if rc != 0:
            raise SystemExit('闭合校验未通过（退出码 %d）：禁止出表，先修数据再重跑' % rc)


if __name__ == '__main__':
    main()
