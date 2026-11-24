# -*- coding: utf-8 -*-
r"""东智收入明细入库（财务口径，唯一收入来源）

源表：`特发东智收入明细表（YYYYMM）.xlsx`（每月一张，按客户，单位：元）。
表结构：`汇总` + `202601`…`2026MM`，每月 sheet 第 2 行表头（客户/收入金额），第 3 行起数据。

用法：
  python dz_income.py import "…\特发东智收入明细表（202610）.xlsx"
  python dz_income.py check
  python dz_income.py report
"""
import os
import re
import sys
import io
import sqlite3
import datetime
import openpyxl

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DB = r'E:\商务部\东智日常表更新\东智台账.db'

DDL = """
CREATE TABLE IF NOT EXISTS 收入明细(
    年月 TEXT NOT NULL,
    客户名称 TEXT NOT NULL,
    客户ID TEXT,
    金额元 REAL NOT NULL,
    金额万元 REAL NOT NULL,
    来源文件 TEXT,
    来源工作表 TEXT,
    来源单元格 TEXT,
    导入时间 TEXT,
    PRIMARY KEY(年月, 客户名称)
);
CREATE INDEX IF NOT EXISTS idx_income_年月 ON 收入明细(年月);
CREATE INDEX IF NOT EXISTS idx_income_客户 ON 收入明细(客户ID);
"""


def ensure_schema(c):
    c.executescript(DDL)
    c.commit()


def load_keys(c):
    keys = []
    for cid, std, alias in c.execute('SELECT 客户ID, 标准名, 别名 FROM 客户主数据'):
        for k in [std] + re.split(r'[、,，;；/]', alias or ''):
            k = (k or '').strip()
            if len(k) >= 2:
                keys.append((k, cid))
    keys.sort(key=lambda x: -len(x[0]))
    return keys


def match_cid(keys, name):
    for k, cid in keys:
        if k in name:
            return cid
    return None


def import_file(path, db=DB):
    wb = openpyxl.load_workbook(path, data_only=True)
    c = sqlite3.connect(db)
    ensure_schema(c)
    keys = load_keys(c)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    src = os.path.basename(path)
    months = sorted(n for n in wb.sheetnames if re.fullmatch(r'20\d{4}', n))
    if not months:
        raise SystemExit('工作簿里没有 YYYYMM 形式的月份表：%s' % path)
    total_rows = 0
    missed = set()
    for sheet in months:
        ym = '%s-%s' % (sheet[:4], sheet[4:])
        ws = wb[sheet]
        rows = []
        r = 3
        while r <= ws.max_row:
            name = ws.cell(r, 1).value
            amt = ws.cell(r, 2).value
            if not name:
                break
            name = str(name).strip()
            if isinstance(amt, (int, float)):
                cid = match_cid(keys, name)
                if not cid:
                    missed.add(name)
                rows.append((ym, name, cid, float(amt), round(float(amt) / 10000.0, 6),
                             src, sheet, 'B%d' % r, now))
            r += 1
        c.execute('DELETE FROM 收入明细 WHERE 年月=?', (ym,))
        c.executemany(
            'INSERT INTO 收入明细(年月,客户名称,客户ID,金额元,金额万元,来源文件,来源工作表,来源单元格,导入时间) '
            'VALUES(?,?,?,?,?,?,?,?,?)', rows)
        total_rows += len(rows)
        print('  %s  %2d 户  %14.6f 万元' % (ym, len(rows), sum(x[4] for x in rows)))
    c.commit()
    c.close()
    print('导入完成：%s，%d 行' % (src, total_rows))
    if missed:
        print('未匹配客户ID（保留原名）：' + '、'.join(sorted(missed)))


def check(db=DB):
    c = sqlite3.connect(db)
    ensure_schema(c)
    print('== 收入明细 vs 月度事实 ==')
    bad = 0
    for ym, total in c.execute('SELECT 年月, round(sum(金额万元),6) FROM 收入明细 GROUP BY 年月 ORDER BY 年月'):
        row = c.execute('SELECT round(sum(收入),6) FROM 月度事实 WHERE 月份=?', (ym,)).fetchone()
        fact = row[0] if row and row[0] is not None else None
        if fact is None:
            print('  %s  收入明细=%14.6f  月度事实=<未关账>' % (ym, total))
            continue
        d = round(fact - total, 6)
        flag = 'OK' if abs(d) <= 0.005 else 'DIFF'
        if flag == 'DIFF':
            bad += 1
        print('  %s  收入明细=%14.6f  月度事实=%14.6f  差=%12.6f  %s' % (ym, total, fact, d, flag))
    print('  %s' % ('PASS' if bad == 0 else 'FAIL（%d 个月）' % bad))
    c.close()


def report(db=DB):
    c = sqlite3.connect(db)
    ensure_schema(c)
    print('| 年月 | 收入(万元) | 客户数 | 来源 |')
    print('|---|---:|---:|---|')
    for ym, total, n, src in c.execute(
            'SELECT 年月, round(sum(金额万元),6), count(*), max(来源文件) FROM 收入明细 GROUP BY 年月 ORDER BY 年月'):
        print('| %s | %.6f | %d | %s |' % (ym, total, n, src))
    c.close()


def main():
    import argparse
    ap = argparse.ArgumentParser(description='东智收入明细入库（财务口径）')
    sub = ap.add_subparsers(dest='cmd')
    p1 = sub.add_parser('import'); p1.add_argument('path')
    sub.add_parser('check'); sub.add_parser('report')
    a = ap.parse_args()
    if a.cmd == 'import':
        import_file(a.path)
    elif a.cmd == 'check':
        check()
    elif a.cmd == 'report':
        report()
    else:
        ap.print_help()


if __name__ == '__main__':
    main()
