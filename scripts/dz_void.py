# -*- coding: utf-8 -*-
"""东智台账 · 作废 / 修正工具（v0.1）

背景：台账是「只增不改」（来源ID 唯一键去重），好处是重跑不会污染、历史留痕；
     坏处是**错值一旦入库就改不掉**。本工具补上这个缺口 —— 用「作废 + 重插」修正。

用法：
  python dz_void.py list [关键字]             列出近期行（找 来源ID 用）
  python dz_void.py void <来源ID> [说明...]    标记作废（留痕；来源ID 加 #作废@时间戳 后缀，便于修正行重新入库）
  python dz_void.py unvoid <来源ID>           撤销作废（恢复原 来源ID、作废=0）

原理：
  1) `void` 把该行标 作废=1、写入「作废说明」，并把「来源ID」改名成 `<原ID>#作废@<时间戳>`
     —— 这样原始 来源ID 就被"腾空"，下次重跑时修正后的数据可以正常插入（不会被唯一键挡住）。
  2) 作废行**不删**，仅在统计时排除（`WHERE COALESCE(作废,0)=0`），所以随时能查到"发生过什么"。
  3) 改完自动把 SQLite 全量回写到 Excel 镜像（东智台账.xlsx）。
"""
import os, sys, sqlite3, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dz_store import DB_P, XLSX_P, store_excel   # noqa: E402

VOID_TAG = '#作废@'


def _conn():
    return sqlite3.connect(DB_P)


def do_list(kw=None):
    conn = _conn(); cur = conn.cursor()
    sql = '''SELECT id,时间,客户名称,签约金额,签约数量,产出数量,发货数量,发货金额,
                   回款金额,回款类型,开票金额,来源ID,COALESCE(作废,0),COALESCE(作废说明,'')
             FROM 台账明细'''
    args = ()
    if kw:
        sql += ' WHERE 客户名称 LIKE ? OR 来源ID LIKE ? OR 时间 LIKE ?'
        args = ('%' + kw + '%',) * 3
    sql += ' ORDER BY 时间 DESC, id DESC LIMIT 60'
    rows = cur.execute(sql, args).fetchall()
    conn.close()
    print('共 %d 行（最多显示 60）' % len(rows))
    print('%-5s %-10s %-10s %-9s %-9s %-9s %-11s %-8s %-8s %-4s %s' % (
        'id', '时间', '客户', '签约额', '签约量', '产出', '发货额', '回款', '开票', '作废', '来源ID'))
    for r in rows:
        _id, t, nm, qa, qq, oq, sq, sa, rq, rt, kq, sid, v, note = r
        def f(x):
            return '' if not x else ('%g' % x)
        flag = '作废' if v else ''
        print('%-5s %-10s %-10s %-9s %-9s %-9s %-11s %-8s %-8s %-4s %s' % (
            _id, t, (nm or '')[:10], f(qa), f(qq), f(oq), f(sa), f(rq), f(kq), flag, sid))
        if note:
            print('      └ 说明：%s' % note)


def do_void(sid, reason=''):
    conn = _conn(); cur = conn.cursor()
    rs = cur.execute('SELECT id,时间,客户名称,COALESCE(作废,0) FROM 台账明细 WHERE 来源ID=?', (sid,)).fetchall()
    if not rs:
        print('! 未找到 来源ID = %s' % sid)
        conn.close(); return
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    note = '%s｜作废于 %s' % ((reason or '人工作废'), ts)
    n = 0
    for _id, t, nm, v in rs:
        if v == 1:
            print('  · 已是作废，跳过：%s %s (id=%d)' % (t, nm, _id)); continue
        cur.execute('UPDATE 台账明细 SET 作废=1, 作废说明=?, 来源ID=? WHERE id=?',
                    (note, '%s%s%s' % (sid, VOID_TAG, ts), _id))
        print('  ✓ 已作废：%s %s (id=%d)' % (t, nm, _id))
        n += 1
    conn.commit(); conn.close()
    if n:
        print('  来源ID 已改名加后缀「%s%s」，下次重跑时修正数据可正常插入。' % (VOID_TAG, ts))
        store_excel()
        print('  Excel 镜像已同步 -> %s' % XLSX_P)


def do_unvoid(sid):
    conn = _conn(); cur = conn.cursor()
    rs = cur.execute('SELECT id,时间,客户名称,来源ID,COALESCE(作废,0) FROM 台账明细 '
                     'WHERE 来源ID=? OR 来源ID LIKE ?', (sid, sid + VOID_TAG + '%')).fetchall()
    if not rs:
        print('! 未找到 来源ID = %s（含 #作废 后缀）' % sid)
        conn.close(); return
    n = 0
    for _id, t, nm, cur_sid, v in rs:
        orig = cur_sid.split(VOID_TAG)[0]
        if v != 1 and cur_sid == orig:
            print('  · 本就未作废，跳过：%s %s (id=%d)' % (t, nm, _id)); continue
        cur.execute('UPDATE 台账明细 SET 作废=0, 作废说明=NULL, 来源ID=? WHERE id=?', (orig, _id))
        print('  ✓ 已撤销作废：%s %s (id=%d) → 来源ID 恢复为 %s' % (t, nm, _id, orig))
        n += 1
    conn.commit(); conn.close()
    if n:
        store_excel()
        print('  Excel 镜像已同步 -> %s' % XLSX_P)


USAGE = __doc__


def main():
    a = sys.argv[1:]
    if not a or a[0] in ('-h', '--help', 'help'):
        print(USAGE); return
    cmd = a[0]
    if cmd == 'list':
        do_list(a[1] if len(a) > 1 else None)
    elif cmd == 'void':
        if len(a) < 2:
            print('用法：python dz_void.py void <来源ID> [说明...]'); return
        do_void(a[1], ' '.join(a[2:]))
    elif cmd == 'unvoid':
        if len(a) < 2:
            print('用法：python dz_void.py unvoid <来源ID>'); return
        do_unvoid(a[1])
    else:
        print('未知命令：%s\n' % cmd); print(USAGE)


if __name__ == '__main__':
    main()
