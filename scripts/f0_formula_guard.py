# -*- coding: utf-8 -*-
r"""F0 公式守卫（v0.1）

解决“录入数据后公式被冲掉”的问题：
  - 写前：快照全部公式（Sheet!单元格 -> 公式），并备份原文件
  - 写后：逐格对比公式；缺失/变更即判失败
  - 失败：自动用写前备份回滚，避免坏文件继续流转

用法：
  python f0_formula_guard.py snapshot "F0.xlsx" [--guard-dir DIR]
  python f0_formula_guard.py check "F0.xlsx" [--baseline MAP.json]
  python f0_formula_guard.py run "F0.xlsx" -- python dz_derive.py ...
  python f0_formula_guard.py rollback "F0.xlsx" [--backup BACKUP.xlsx]

铁律：
  - 只用 data_only=False 读取/写入；
  - 不拿 pandas/openpyxl(data_only=True) 回写 F0；
  - 只写输入格，公式格不碰。
"""
import argparse
import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys

from openpyxl import load_workbook

DEFAULT_GUARD_DIR = r'E:\商务部\东智日常表更新\_formula_guard'
FORMULA_PREFIX = '='


def _sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def _safe_stem(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    return ''.join(ch if ch.isalnum() or ch in '-_()（）' else '_' for ch in stem)


def _scan(path):
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    wb = load_workbook(path, data_only=False, read_only=True)
    formulas, counts = {}, {}
    try:
        for ws in wb.worksheets:
            n = 0
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith(FORMULA_PREFIX):
                        formulas['%s!%s' % (ws.title, cell.coordinate)] = value
                        n += 1
            counts[ws.title] = n
    finally:
        wb.close()
    return formulas, counts


def _map_dir(guard_dir):
    d = os.path.join(guard_dir, 'maps')
    os.makedirs(d, exist_ok=True)
    return d


def _backup_dir(guard_dir):
    d = os.path.join(guard_dir, 'backups')
    os.makedirs(d, exist_ok=True)
    return d


def _latest_map(path, guard_dir):
    d = _map_dir(guard_dir)
    stem = _safe_stem(path)
    items = sorted(
        (os.path.join(d, x) for x in os.listdir(d)
         if x.startswith(stem + '_') and x.endswith('_formula_map.json')),
        key=os.path.getmtime, reverse=True)
    return items[0] if items else None


def _latest_backup(path, guard_dir):
    d = _backup_dir(guard_dir)
    stem = _safe_stem(path)
    items = sorted(
        (os.path.join(d, x) for x in os.listdir(d)
         if x.startswith(stem + '_') and x.endswith('.xlsx')),
        key=os.path.getmtime, reverse=True)
    return items[0] if items else None


def snapshot(path, guard_dir):
    formulas, counts = _scan(path)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    stem = _safe_stem(path)
    map_path = os.path.join(_map_dir(guard_dir), '%s_%s_formula_map.json' % (stem, ts))
    payload = {
        'schema_version': '1.0',
        'file': os.path.abspath(path),
        'sha256': _sha256(path),
        'created_at': ts,
        'total': len(formulas),
        'counts': counts,
        'formulas': formulas,
    }
    with open(map_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print('OK 公式快照：%d 条' % len(formulas))
    for sheet, n in counts.items():
        if n:
            print('   %-28s %d' % (sheet, n))
    print('   MAP=%s' % map_path)
    return 0, map_path, payload


def _compare(current, counts, baseline):
    base_formulas = baseline.get('formulas', {})
    missing, changed = [], []
    for key, old in base_formulas.items():
        new = current.get(key)
        if new is None:
            missing.append((key, old))
        elif new != old:
            changed.append((key, old, new))
    extra = [k for k in current if k not in base_formulas]
    base_counts = baseline.get('counts', {})
    count_diffs = []
    for sheet in sorted(set(base_counts) | set(counts)):
        a, b = int(base_counts.get(sheet, 0)), int(counts.get(sheet, 0))
        if a != b:
            count_diffs.append((sheet, a, b, b - a))
    return missing, changed, extra, count_diffs


def check(path, baseline_path):
    if not baseline_path:
        raise ValueError('缺少 --baseline，且未找到自动基线')
    with open(baseline_path, encoding='utf-8') as f:
        baseline = json.load(f)
    current, counts = _scan(path)
    missing, changed, extra, count_diffs = _compare(current, counts, baseline)

    print('公式总数：当前 %d / 基线 %d' % (len(current), baseline.get('total', 0)))
    if count_diffs:
        print('分 Sheet 数量差异：')
        for sheet, old, new, diff in count_diffs:
            print('   %-28s %d -> %d (%+d)' % (sheet, old, new, diff))
    if missing:
        print('[FAIL] 缺失公式 %d 条：' % len(missing))
        for key, old in missing[:20]:
            print('   - %s = %s' % (key, old[:120]))
    if changed:
        print('[FAIL] 公式被改写 %d 条：' % len(changed))
        for key, old, new in changed[:20]:
            print('   - %s: %s -> %s' % (key, old[:80], new[:80]))
    if extra:
        print('[WARN] 新增公式 %d 条（通常可接受）' % len(extra))
        for key in extra[:10]:
            print('   - %s' % key)
    if missing or changed:
        print('FAIL 公式不完整，建议立即回滚：%s' % baseline_path)
        return 1
    print('OK 公式完整，未发现缺失/改写')
    return 0


def _backup(path, guard_dir):
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    dst = os.path.join(_backup_dir(guard_dir), '%s_%s.xlsx' % (_safe_stem(path), ts))
    shutil.copy2(path, dst)
    print('   备份=%s' % dst)
    return dst


def _rollback(path, backup):
    if not backup or not os.path.exists(backup):
        print('FAIL 找不到可回滚备份')
        return 1
    shutil.copy2(backup, path)
    print('已回滚：%s -> %s' % (backup, path))
    return 0


def run_guarded(path, command, guard_dir):
    if not command:
        print('FAIL 没有传入要执行的命令；用法：run "F0.xlsx" -- python ...')
        return 2
    if not os.path.exists(path):
        print('FAIL 文件不存在：%s' % path)
        return 2

    rc, map_path, _ = snapshot(path, guard_dir)
    if rc:
        return rc
    backup = _backup(path, guard_dir)

    print('   执行：%s' % ' '.join(command))
    proc = subprocess.run(command)

    if not os.path.exists(path):
        print('FAIL 子命令执行后文件不存在，自动回滚…')
        _rollback(path, backup)
        return 3
    print('   写后校验公式…')
    check_rc = check(path, map_path)
    if check_rc:
        print('FAIL 公式被破坏，自动回滚…')
        _rollback(path, backup)
        return 3

    if proc.returncode:
        print('WARN 子命令退出码 %d，但公式完整，未回滚' % proc.returncode)
        return proc.returncode
    print('OK 子命令完成，公式完整')
    return 0


def main():
    parser = argparse.ArgumentParser(description='F0 公式守卫')
    parser.add_argument('--guard-dir', default=DEFAULT_GUARD_DIR)
    sub = parser.add_subparsers(dest='cmd', required=True)

    p1 = sub.add_parser('snapshot')
    p1.add_argument('file')

    p2 = sub.add_parser('check')
    p2.add_argument('file')
    p2.add_argument('--baseline', default='')

    p3 = sub.add_parser('run')
    p3.add_argument('file')
    p3.add_argument('command', nargs=argparse.REMAINDER)

    p4 = sub.add_parser('rollback')
    p4.add_argument('file')
    p4.add_argument('--backup', default='')

    args = parser.parse_args()
    try:
        if args.cmd == 'snapshot':
            rc, _, _ = snapshot(args.file, args.guard_dir)
            return rc
        if args.cmd == 'check':
            baseline = args.baseline or _latest_map(args.file, args.guard_dir)
            if not baseline:
                print('FAIL 没有基线；先跑 snapshot')
                return 2
            return check(args.file, baseline)
        if args.cmd == 'run':
            command = args.command
            if command and command[0] == '--':
                command = command[1:]
            return run_guarded(args.file, command, args.guard_dir)
        if args.cmd == 'rollback':
            backup = args.backup or _latest_backup(args.file, args.guard_dir)
            return _rollback(args.file, backup)
    except Exception as exc:
        print('FAIL %s: %s' % (type(exc).__name__, exc))
        return 2
    return 2


if __name__ == '__main__':
    sys.exit(main())
