# -*- coding: utf-8 -*-
"""给 xlsx 写不可见的批次水印（docProps/custom.xml），不改任何单元格/公式。

用法：
  python dz_watermark.py <xlsx> <批次ID> [对象名]

水印文本从 DB 批次清单/out/批次.json 推算；同时写 DB 批次水印 表。
"""
import os
import re
import sys
import json
import shutil
import sqlite3
import zipfile
import datetime

DB = os.environ.get('DZ_DB') or r'E:\商务部\东智日常表更新\东智台账.db'
HERE = os.path.dirname(os.path.abspath(__file__))
PROPS = 'docProps/custom.xml'
FMT = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'


def esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def set_props(xml, pairs):
    """在已有 custom.xml 里插入/替换 name=... 的属性，保留原属性。"""
    for name, value in pairs:
        pat = r'<property\b[^>]*name="%s"[^>]*>.*?</property>' % re.escape(name)
        new = ('<property fmtid="%s" pid="99" name="%s"><vt:lpwstr>%s</vt:lpwstr></property>'
               % (FMT, esc(name), esc(value)))
        if re.search(pat, xml, flags=re.S):
            xml = re.sub(pat, new, xml, flags=re.S)
        else:
            xml = xml.replace('</Properties>', new + '</Properties>')
    return xml


def main():
    if len(sys.argv) < 3:
        print(__doc__); return 1
    path, batch = sys.argv[1], sys.argv[2]
    obj = sys.argv[3] if len(sys.argv) > 3 else 'F0周更'
    if not os.path.exists(path):
        print('! 文件不存在：%s' % path); return 1

    conn = sqlite3.connect(DB)
    row = conn.execute('SELECT 截止时间 FROM 报告批次 WHERE 批次ID=?', (batch,)).fetchone()
    conn.close()
    cutoff = (row[0] if row else '')[:10]
    if not cutoff:
        m = re.search(r'(\d{4})(\d{2})(\d{2})$', batch)
        if m:
            cutoff = '%s-%s-%s' % m.groups()

    # 水印 = DZ-<批次ID>-<批次清单哈希前8>；没有清单则用批次ID
    h = os.environ.get('DZ_HASH', '')
    wm = 'DZ-%s-%s' % (batch, h) if h else 'DZ-' + batch

    bak_dir = r'E:\商务部\东智日常表更新\_backup'
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, os.path.basename(path) + '.wm-' +
                       datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
    shutil.copy(path, bak)
    print('已备份 -> %s' % bak)

    tmp = path + '.wm-tmp'
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == PROPS:
                xml = data.decode('utf-8')
                xml = set_props(xml, [('DZ批次水印', wm), ('DZ批次ID', batch),
                                      ('DZ数据截止', cutoff), ('DZ写入时间',
                                       datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))])
                data = xml.encode('utf-8')
            zout.writestr(item, data)
    os.replace(tmp, path)
    print('已写文档属性 -> %s' % path)

    conn = sqlite3.connect(DB)
    conn.execute('''INSERT OR REPLACE INTO 批次水印(对象,批次ID,水印,数据截止,说明)
                    VALUES(?,?,?,?,?)''', (obj, batch, wm, cutoff, 'docProps/custom.xml'))
    conn.commit(); conn.close()
    print('已记 DB 批次水印：%s / %s' % (obj, wm))
    return 0


if __name__ == '__main__':
    sys.exit(main())
