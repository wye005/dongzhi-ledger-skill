# -*- coding: utf-8 -*-
r"""东智源文件自动落盘（取源自动化第一层）

把 5 类源文件从「邮件附件缓存 / _inbox / 周期目录」收拢到周期目录，并记录 sha256。
不解析、不改数，只负责「找到最新那份、复制进来、留痕」。解析仍由 dz_ingest 做。

源：
  ① OA 合约订单评审表     合约订单评审表*.xls*
  ② NC 销售出库单         销售出库单明细表*.csv / *.xls*   （dz_ingest 也会自动调接口）
  ③ 交付跟踪（产出）      订单交付跟踪表MMDD.xlsx
  ④ 发票回款登记           *东智销售发票及回款登记表*.xlsx
  ⑤ 财务收入明细          特发东智收入明细表（YYYYMM）.xlsx

用法：
  python dz_sources.py scan  [--report 2026-11-20]
  python dz_sources.py stage <周期目录> <报表日> [--strict]
"""
import os
import re
import sys
import io
import glob
import json
import shutil
import hashlib
import datetime

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r'E:\商务部\东智日常表更新'
WEMAIL = r'C:\Users\11447\AppData\Roaming\Tencent\WeMail\cache\attach'
INBOX = os.path.join(BASE, '_inbox')

PATTERNS = {
    'OA': ['合约订单评审表*.xls*'],
    'NC': ['销售出库单明细表*.csv', '销售出库单明细表*.xls*'],
    '交付': ['订单交付跟踪表*.xlsx'],
    '发票回款': ['*东智销售发票及回款登记表*.xlsx'],
    '收入': ['特发东智收入明细表*.xlsx'],
}
REQUIRED_WEEKLY = ['OA', '交付', '发票回款']
REQUIRED_CLOSE = ['OA', '交付', '发票回款', '收入']


def sha256(path, limit=None):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            b = f.read(1024 * 1024)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _roots(period_dir=None):
    roots = [WEMAIL]
    if os.path.isdir(INBOX):
        roots.append(INBOX)
    if period_dir and os.path.isdir(period_dir):
        roots.append(period_dir)
    return roots


def find_all(patterns, period_dir=None):
    out = []
    for root in _roots(period_dir):
        for pat in patterns:
            for p in glob.glob(os.path.join(root, '**', pat), recursive=True):
                if '~$' in os.path.basename(p):
                    continue
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                out.append((st.st_mtime, st.st_size, p))
    seen, uniq = set(), []
    for m, s, p in sorted(out, reverse=True):
        if p in seen:
            continue
        seen.add(p)
        uniq.append((m, s, p))
    return uniq


def _date_of(path):
    b = os.path.basename(path)
    m = re.search(r'(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})', b)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.search(r'(\d{4})(?=\.xls)', b)
    if m:
        y, md = int(m.group(1)[:4]), m.group(1)[4:]
        try:
            return datetime.date(y, int(md[:2]), int(md[2:]))
        except ValueError:
            return None
    m = re.search(r'(\d{2})(\d{2})\.xlsx$', b)
    if m:
        return None
    return None


def _delivery_date(path, year):
    m = re.search(r'订单交付跟踪表(\d{2})(\d{2})', os.path.basename(path))
    if not m:
        return None
    try:
        return datetime.date(year, int(m.group(1)), int(m.group(2)))
    except ValueError:
        return None


def _income_month(path):
    m = re.search(r'（(\d{6})）', os.path.basename(path))
    if not m:
        return None
    return '%s-%s' % (m.group(1)[:4], m.group(1)[4:])


def scan(report=None):
    report_d = datetime.date.fromisoformat(report) if report else datetime.date.today()
    print('== 源文件扫描（报表日 %s）==' % report_d.isoformat())
    found = {}
    for kind, pats in PATTERNS.items():
        hits = find_all(pats)
        if kind == '交付':
            hits = [(m, s, p) for m, s, p in hits
                    if _delivery_date(p, report_d.year) and _delivery_date(p, report_d.year) <= report_d]
        elif kind == '收入':
            hits = [(m, s, p) for m, s, p in hits
                    if _income_month(p) and _income_month(p) <= report_d.strftime('%Y-%m')]
        found[kind] = hits
        print('-- %s --' % kind)
        if not hits:
            print('   （未找到）')
        for m, s, p in hits[:5]:
            print('   %s  %7.1fKB  %s' % (
                datetime.datetime.fromtimestamp(m).strftime('%Y-%m-%d %H:%M'),
                s / 1024.0, p))
    return found


def _copy_if_newer(src, dst):
    if os.path.exists(dst):
        if os.path.getmtime(dst) >= os.path.getmtime(src):
            return 'skip'
        if sha256(src) == sha256(dst):
            return 'same'
    shutil.copy2(src, dst)
    return 'copy'


def stage(period_dir, report, strict=False, close=False):
    report_d = datetime.date.fromisoformat(report)
    d1 = datetime.date.fromisoformat(os.path.basename(period_dir).split('-')[0])
    os.makedirs(period_dir, exist_ok=True)
    manifest = []
    missing = []

    for kind, pats in PATTERNS.items():
        if kind == '交付':
            hits = [(m, s, p) for m, s, p in find_all(pats, period_dir)
                    if _delivery_date(p, report_d.year)
                    and d1 <= _delivery_date(p, report_d.year) <= report_d]
            if not hits:
                # 没有本周快照时，至少把报表日前最新一份带进来当基线
                hits = [(m, s, p) for m, s, p in find_all(pats, period_dir)
                        if _delivery_date(p, report_d.year)
                        and _delivery_date(p, report_d.year) <= report_d][:1]
        elif kind == '收入':
            hits = [(m, s, p) for m, s, p in find_all(pats, period_dir)
                    if _income_month(p) and _income_month(p) <= report_d.strftime('%Y-%m')]
            if hits:
                hits = sorted(hits, key=lambda x: _income_month(x[2]))[-1:]
        else:
            hits = find_all(pats, period_dir)
            if hits:
                hits = hits[:1]

        if not hits:
            missing.append(kind)
            hint = {'OA': '先跑 pull_csv.py + oa2xlsx.py 或手工放合约订单评审表',
                    'NC': 'dz_ingest 会调 nc_pull.ps1 自动取；也可手工放销售出库单',
                    '交付': '等杨艳秀邮件，或把订单交付跟踪表放 _inbox',
                    '发票回款': '等李玉婷邮件，或把登记表放 _inbox',
                    '收入': '等财务《特发东智收入明细表（YYYYMM）》'}.get(kind, '')
            print('  ! 缺 %s  （%s）' % (kind, hint))
        for m, s, p in sorted(hits, key=lambda x: x[0]):
            dst = os.path.join(period_dir, os.path.basename(p))
            act = _copy_if_newer(p, dst)
            if act in ('copy', 'same'):
                print('  %-5s %-6s %s' % (kind, act, os.path.basename(p)))
            manifest.append({
                '类型': kind, '文件名': os.path.basename(dst), '动作': act,
                '来源': p, '字节数': os.path.getsize(dst),
                'mtime': datetime.datetime.fromtimestamp(os.path.getmtime(dst)).strftime('%Y-%m-%d %H:%M:%S'),
                'sha256': sha256(dst),
            })

    required = REQUIRED_CLOSE if close else REQUIRED_WEEKLY
    hard = [k for k in required if k in missing]
    out = os.path.join(period_dir, '源清单.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump({'报表日': report, '周期': os.path.basename(period_dir),
                   '生成时间': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                   '缺源': missing, '文件': manifest}, f, ensure_ascii=False, indent=2)
    print('  源清单 -> %s' % out)
    if hard and strict:
        print('FAIL 缺必需源：%s' % '、'.join(hard))
        return 1
    return 0


def main():
    args = sys.argv[1:]
    if not args or args[0] not in ('scan', 'stage'):
        print(__doc__)
        return
    report = None
    if '--report' in args:
        report = args[args.index('--report') + 1]
    if args[0] == 'scan':
        scan(report)
        return
    rest = [a for a in args[1:] if not a.startswith('--') and a != report]
    if len(rest) < 2:
        print('用法：python dz_sources.py stage <周期目录> <报表日>')
        sys.exit(2)
    sys.exit(stage(rest[0], rest[1], '--strict' in args, '--close' in args))


if __name__ == '__main__':
    main()
