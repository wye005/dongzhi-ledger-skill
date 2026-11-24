# -*- coding: utf-8 -*-
"""东智台账 · 公共模块
工具函数 + 台账列定义，供 dz_ingest / build_ledger / dz_store / dz_derive 共用。
"""
import os, glob, datetime


def num(v):
    """宽松转数值：None/空 → 0.0"""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('，', '').replace(',', '').replace(' ', '')
    if not s:
        return 0.0
    # 2026-11-07 事故：9 月登记表华为回款写成 '7,759.543.69'（两个小数点）。
    # 旧逻辑 float() 失败后静默返回 0，整笔 775.95 万被漏掉。
    # 这里把「多个点、最后一段恰好 2 位」解释为千分位点 + 小数点。
    if s.count('.') > 1:
        parts = s.split('.')
        if len(parts[-1]) == 2 and all(len(x) == 3 for x in parts[1:-1]) and 1 <= len(parts[0]) <= 4:
            s = ''.join(parts[:-1]) + '.' + parts[-1]
    try:
        return float(s)
    except Exception:
        return 0.0


def num_strict(v, where=''):
    """严格转数值：非空但无法解析时抛错，禁止静默漏数。"""
    if v is None or (isinstance(v, str) and not v.strip()):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('，', '').replace(',', '').replace(' ', '')
    if s in ('-', '—', '/', '无', 'N/A', 'NA', 'n/a'):
        return 0.0
    if s.count('.') > 1:
        parts = s.split('.')
        if len(parts[-1]) == 2 and all(len(x) == 3 for x in parts[1:-1]) and 1 <= len(parts[0]) <= 4:
            s = ''.join(parts[:-1]) + '.' + parts[-1]
    try:
        return float(s)
    except Exception as exc:
        raise ValueError('无法解析数值 %r %s' % (v, ('@ ' + where) if where else '')) from exc


def _not_lock(p):
    """排除 Office/WPS 的临时锁文件（~$xxx）"""
    return '~$' not in os.path.basename(p)


def find_file(src, pattern, prefer_xls=False):
    """在 src 目录按 pattern 找最新文件；prefer_xls=True 时优先 .xls（避免取到 .xlsx 被 xlrd 拒读）。
    自动排除 `~$` 锁文件（2026-09-18 曾因取到 ~$ 锁文件报 PermissionError）。"""
    if prefer_xls:
        cand = [p for p in sorted(glob.glob(os.path.join(src, pattern.replace('.xls*', '.xls')))) if _not_lock(p)]
        if cand:
            return cand[-1]
    cand = [p for p in sorted(glob.glob(os.path.join(src, pattern))) if _not_lock(p)]
    return cand[-1] if cand else None


# 台账 24 列（含「收入」「信息来源路径」「来源ID」「作废」「作废说明」「导入批次」）
# 「收入」= 发货金额 ÷ 1.13；光缆事业部计 0；T01 再 ×0.019（用户 2026-09-15 定）
# 「来源ID」= 业务天然唯一键（幂等去重用，只增不改）；「作废」= 0/1（留痕标记，统计排除）；「作废说明」= 原因/时间戳
# 「导入批次」= 该行首次入库时的周期目录名（如 20260912-20260918）。派生层按它切周 —— 因为"行属于哪一周"
#             取决于【首次导入时点】，不是【时间列】：做表后补录的行会带着上周的日期进这一周。
COLS = ['客户名称', '类别', '产品类别', '时间', '签约金额预测', '签约金额', '签约数量预测', '签约数量',
        '产出计划', '产出数量', '发货数量交付计划', '发货数量', '发货金额交付计划', '发货金额',
        '收入',
        '回款计划', '回款金额', '回款类型', '开票计划', '开票金额',
        '信息来源路径', '来源ID', '作废', '作废说明', '导入批次',
        # 期间归属（2026-11-23）：时间=落账日；业务日期=源单据原始日期；归属月=会计期间
        '业务日期', '归属月', '归属依据']
TXT_COLS = {'客户名称', '类别', '产品类别', '时间', '回款类型', '信息来源路径', '来源ID', '作废说明', '导入批次',
            '业务日期', '归属月', '归属依据'}
NUM_COLS = [c for c in COLS if c not in TXT_COLS]


def product_category(text):
    """产品名称/机型/类别文本 → 标准产品类别（关键词匹配，写死优先级；未命中归「其他」）。
    用户 2026-09-20 定：产品类别从 4 源提取、写进台账。"""
    s = str(text or '').strip()
    if not s:
        return '其他'
    low = s.lower()
    for cat, kws in RULES['PRODUCT_CATEGORY']:
        for kw in kws:
            if kw.lower() in low:
                return cat
    return '其他'


def batch_of(path):
    """从源文件路径推出「导入批次」= 周期目录名（如 20260912-20260918）"""
    d = os.path.dirname(str(path or '').strip())
    b = os.path.basename(os.path.normpath(d)) if d else ''
    return b if b[:8].isdigit() else ''

# 台账 Excel 镜像路径：优先「东智台账.xlsx」（2026-09-17 用户改名），兼容旧名「母表.xlsx」
LEDGER_XLSX_CANDIDATES = [
    r'E:\商务部\东智日常表更新\东智台账.xlsx',
    r'E:\商务部\东智日常表更新\母表.xlsx',
]


def ledger_xlsx():
    """返回台账 Excel 镜像的实际路径（改名后仍可用）"""
    for p in LEDGER_XLSX_CANDIDATES:
        if os.path.exists(p):
            return p
    return LEDGER_XLSX_CANDIDATES[0]


# ---------------- F0（老宽表）定位 ----------------
import re as _re

BASE_DIR = r'E:\商务部\东智日常表更新'
F0_SUFFIX = '特发东智订单、产出、发货、开票及回款汇总表.xlsx'


def f0_date(path):
    """从文件名「截止到2026年9月11日…」解析出 2026-09-11；解析不到返回 ''"""
    m = _re.search(r'截止到(\d{4})年(\d{1,2})月(\d{1,2})日', os.path.basename(path))
    return '%04d-%02d-%02d' % (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else ''


def f0_path(period_dir, d):
    """按日期拼出 F0 路径（放周期目录内）"""
    y, m, dd = (int(x) for x in d.split('-'))
    return os.path.join(period_dir, '截止到%d年%d月%d日%s' % (y, m, dd, F0_SUFFIX))


def find_prev_f0(d1, base_dir=None):
    """找「上周五基准」F0：各周期目录里、文件日期 < d1 的最新一份。返回 (路径, 日期)"""
    bd = base_dir or BASE_DIR
    cands = glob.glob(os.path.join(bd, '截止到*汇总表*.xlsx'))
    cands += glob.glob(os.path.join(bd, '*', '截止到*汇总表*.xlsx'))
    best, bestd = None, ''
    for p in cands:
        if '~$' in os.path.basename(p):
            continue
        ds = f0_date(p)
        if ds and ds < d1 and ds > bestd:
            best, bestd = p, ds
    return best, bestd


# ================= 用户钦定规则（写死，唯一真源）=================
# 所有「用户定过的事」集中在这里，各脚本从这里取。改口径只改这一处，勿在别处再写一份。
# 每条都标了出处（用户 + 日期）。用户 2026-09-20 明确："我定了的东西一定要写进代码里，写死"。
RULES = {
    # 客户名归一：含关键字的源名 → 直接归到该台账行
    'SUBSTR_RULES': [
        ('华为', 'T01'),              # 用户 2026-08：T01 = 华为技术有限公司 + 深圳华为云计算技术有限公司
        ('鼎泰恒华', '鼎泰恒华'),        # 用户 2026-09-18：源名"深圳鼎泰恒华智慧科技有限公司" → 行名「鼎泰恒华」
        # 用户 2026-10-22 定：9 月新增三家客户，单独建客户行（原 9/30 F0 与定稿表均无此行）
        ('洁宇再生资源', '洁宇再生资源'),  # 深圳市洁宇再生资源回收科技有限公司
        ('洁字再生资源', '洁宇再生资源'),  # 用户确认「洁字」为「洁宇」错字，同一家（签118/发118/回113）
        ('宝鑫再生资源', '宝鑫再生资源'),  # 惠州市宝鑫再生资源回收有限公司
        ('顺迪通讯', '顺迪通讯'),        # 上海顺迪通讯科技有限公司
        # 2026-11-20：F0 开票/回款/出货块有行、但「订单签约跟踪表」主数据没有的客户
        ('新基建', '新基建'),            # 源名「新基建事业部」→ F0 行名「新基建」
        ('凌盛通', '凌盛通'),            # 交付跟踪表有订单(10,000台)、F0 出货块有行
    ],
    'ALIAS': {                         # 名称变体 → 台账行名
        '中兴DSL:无线猫': '中兴DSL',
        '友华物料售卖': '友华',
        '物联网公司': '中移物联网',
    },
    # 产品类别（用户 2026-09-20 定：从 4 源提取、写进台账）
    # 标准类别与 F4 透视表一致：PON / 服务器 / 机顶盒 / 交换机 / 路由器 / 护眠仪 / 其他
    # 顺序=优先级（先匹配先归）；源文件里 OA 的「产品类别」取值很杂（含费用/搬运车等），故一律走关键词映射
    'PRODUCT_CATEGORY': [
        ('护眠仪', ('护眠仪', '睡眠', '脘表', '睡眠雷达', '雷达')),
        ('交换机', ('交换机', 'switch', 'SW-')),
        ('服务器', ('服务器', 'server')),
        ('机顶盒', ('机顶盒', 'STB', '小盒子')),
        ('路由器', ('路由器', 'router', 'WiFi', '无线路由', 'WIFI')),
        ('PON',    ('PON', 'GPON', 'EPON', 'ONU', 'OLT', '光网络单元', '光纤接入', '网关')),
    ],
    # 收入（用户 2026-09-15 定，2026-09-18 再确认"按台账算"）
    'VAT': 1.13,                        # 收入 = 发货金额 ÷ 1.13
    'INCOME_ZERO': ('光缆事业部',),      # 光缆事业部收入计 0
    'INCOME_RATE': {'T01': 0.019},       # T01 = 发货金额 ÷ 1.13 × 0.019（先除后乘，不是 ÷(1.13×0.019)）
    # F0 无独立客户行、由「其他类回款统计」承载的客户（官方表同口径）
    'F0_OTHER_BUCKET': ('宝鑫再生资源',),
    # 作废判定：OA 状态含这些字 → 作废=1（留痕、统计排除）
    'VOID_KEYWORDS': ('撤销', '作废', '中止', '终止', '取消', '废'),
    # 半成品清单（母表子表名）；清单内计数量、清单外含"半成品"→ 标 0 并列出待问用户
    'HALF_DONE_SHEET': '半成品清单',
    # 发货数量：仓库含"成品仓"才计数量；金额全仓计入
    'FINISHED_WH': '成品仓',
    # NC 接口仓库ID（2026-11-07 由 10 月手工导出逐行反查）：成品仓(MES仓)
    # 接口 CSV 没有仓库名称，只有 CBODYWAREHOUSEID，故用 ID 判定；改仓库主数据时同步改这里
    'NC_FINISHED_WH_ID': ('1001A110000000XCQ1SX',),
    # 发货时效性：窗口含「上周五(D1-1)」补录、排除「本周五(D2)」做表后发生的
    'SHIP_WINDOW_D1_MINUS_1': True,
    # NC 合计行防呆：购货单位为空即跳过
    'NC_SKIP_EMPTY_CUST': ('合计', '总计'),
    # 分项汇总 6 处「待做」列标题（用户 2026-09-20 定）：
    #   非月末做表 → 「{做表日+1 M/D}-{月底 M/D}日待XXX」（如 9/19-9/30日待下单数量）
    #   月末做表（做表日==月底）→ 「本月待XXX」
    'TITLE_待': [
        ('G4',   '待下单数量'),
        ('H4',   '待下单金额'),
        ('E75',  '待产出量'),
        ('G136', '待出货数量'),
        ('H136', '待出货金额'),
        ('F210', '待开票金额'),
    ],
}


# ================= F0（老宽表）布局（写死，唯一真源）=================
# 三张子表：线名 -> (sheet, 块起行, 块止行, 计划列, 累计列)
#   计划列：签约表 D(4)，其余表 C(3)；累计列：签约 AK(37)，开票 AI(35)，其余 AJ(36)
# 这些位置由「客户统一名单」固定，插行会让行号漂移 → 一律按客户名在块内找行，不写死行号。
SUB_SHEETS = {
    '签约金额': ('订单签约跟踪表',   4,  85, 4, 37),
    '签约数量': ('订单签约跟踪表',  87, 168, 4, 37),
    '产出数量': ('产出发货跟踪表',   4,  69, 3, 36),
    '发货数量': ('产出发货跟踪表',  71, 152, 3, 36),
    '发货金额': ('产出发货跟踪表', 154, 235, 3, 36),
    '回款金额': ('回款&开票统计表',  4,  94, 3, 36),
    '开票金额': ('回款&开票统计表', 96, 178, 3, 35),
}

# 分项汇总 5 块：块名 -> (客户起, 客户止, 合计行)
# ⚠️ 客户止**不含**「其他统计」特殊行（开票块 R277「其他类型开票统计」、回款块 R355「其他类回款统计」），
#    否则合计 SUM 会把特殊行也算进去造成重复（2026-09-20 实测回款翻倍）。
# 注意：这是「无新增客户」时的行号；dz_derive 用 `compute_fx.find_blocks()` 动态定位，行号漂移时以此为准。
FX_BLOCKS = {
    '订单计划': (5,   71,  72),
    '产出计划': (76, 132, 133),
    '出货计划': (137, 206, 207),
    '开票计划': (211, 276, 278),
    '回款计划': (282, 354, 356),
}
