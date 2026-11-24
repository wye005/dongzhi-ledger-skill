# -*- coding: utf-8 -*-
r"""OA 合约订单评审明细 · 自动取数（取源自动化第二层）

绕开门户页面，直接打行级接口 NewOrderReviewDetailData.ashx。
通道优先级：ego-browser -> agent-browser（都复用已登录浏览器）。
页面侧 fetch 后 POST 到本机临时端口落盘，避免 9MB 走 stdout 分块。

用法：
  python oa_pull.py <周期目录> [--report YYYY-MM-DD] [--timeout 240]

退出码：
  0 = 已取到并生成 xlsx
  3 = 浏览器未登录 / 通道不可用（非致命，沿用周期目录里已有的 OA 文件）
  4 = 接口或落盘失败
"""
import os
import sys
import io
import json
import threading
import subprocess
from http.server import BaseHTTPRequestHandler, HTTPServer

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

API = ('https://bpm.sdgi.com.cn:8443/YZSoft/DZ/SalesOrder/'
       'NewOrderReviewDetailData.ashx?method=GetData&searchType=QuickSearch'
       '&page=1&start=0&limit=6000')
PORT = 18765
EGO = r'C:\Users\11447\.local\share\ego\active_version_dir\ego-browser.exe'
AB = (r'C:\Users\11447\.workbuddy\binaries\node\versions\22.22.2\node_modules'
      r'\agent-browser\bin\agent-browser-win32-x64.exe')

RECV = {'data': None, 'done': threading.Event()}


class Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        n = int(self.headers.get('Content-Length') or 0)
        RECV['data'] = self.rfile.read(n)
        self.send_response(204)
        self.end_headers()
        RECV['done'].set()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()

    def log_message(self, *a):
        pass


def run_js(expr, timeout=180):
    env = dict(os.environ, PLAYWRIGHT_BROWSERS_PATH=r'E:\playwright-browsers')
    cmds = [[EGO, 'nodejs', '-e', expr], [AB, 'eval', expr]]
    last = ''
    for c in cmds:
        if not os.path.exists(c[0]):
            continue
        try:
            p = subprocess.run(c, capture_output=True, text=True,
                               encoding='utf-8', env=env, timeout=timeout)
        except subprocess.TimeoutExpired:
            last = '%s 超时' % os.path.basename(c[0])
            continue
        out = (p.stdout or '').strip()
        err = (p.stderr or '').strip()
        if p.returncode != 0 or not out:
            last = '%s: %s' % (os.path.basename(c[0]), (err or out)[:200])
            continue
        try:
            return json.loads(out)
        except Exception:
            return out.strip('"')
    raise RuntimeError(last or '没有可用的浏览器通道')


PROBE = ('(async function(){try{var r=await fetch(%r,{credentials:"include"});'
         'var t=await r.text();return "LEN"+t.length+"|"+t.slice(0,160);}'
         'catch(e){return "ERR "+e;}})()') % API

FETCH = r'''(async function(){
try{
var r=await fetch(%URL%,{credentials:"include"});
var t=await r.text();
if(t.indexOf('"success": false')>=0||t.indexOf('errorCode')>=0)return "NOTLOGIN";
var s=t.replace(/\\'/g,"'");
var j=new Function('return ('+s+')')();
var rows=j.children||j.rows||[];
if(!rows.length)return "EMPTY";
var keys=Object.keys(rows[0]);
function esc(v){if(v===null||v===undefined)return "";var x=String(v).replace(/"/g,'""');return '"'+x+'"';}
var out=[keys.join(",")];
for(var i=0;i<rows.length;i++){var o=rows[i];var line=[];for(var k=0;k<keys.length;k++){line.push(esc(o[keys[k]]));}out.push(line.join(","));}
var csv="\ufeff"+out.join("\r\n");
await fetch("http://127.0.0.1:%PORT%/oa",{method:"POST",mode:"no-cors",body:csv});
return "OK rows="+rows.length+" bytes="+csv.length;
}catch(e){return "ERR "+String(e).slice(0,200);}
})()'''


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    period = sys.argv[1]
    report = None
    if '--report' in sys.argv:
        report = sys.argv[sys.argv.index('--report') + 1]
    timeout = 240
    if '--timeout' in sys.argv:
        timeout = int(sys.argv[sys.argv.index('--timeout') + 1])
    tag = (report or '').replace('-', '') or 'latest'
    dst = os.path.join(period, 'oa_review_detail_%s.csv' % tag)

    print('· OA 接口探活 ...')
    try:
        probe = run_js(PROBE, timeout=90)
    except Exception as e:
        print('  ! 浏览器通道不可用：%s' % e)
        print('  -> 打开 ego lite 或已登录浏览器后重跑；本次沿用现有 OA 文件')
        return 3

    if 'NOTLOGIN' in probe or '用户未登录' in probe:
        print('  ! OA 未登录（登录已超时）')
        print('  -> 在已打开的浏览器里登录一次 https://bpm.sdgi.com.cn:8443/ ，重跑即可')
        return 3
    if probe.startswith('ERR') or not probe.startswith('LEN'):
        print('  ! 接口异常：%s' % probe[:200])
        return 4
    print('  接口可达：%s' % probe[:90])

    httpd = HTTPServer(('127.0.0.1', PORT), Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    expr = FETCH.replace('%URL%', repr(API)).replace('%PORT%', str(PORT))
    try:
        res = run_js(expr, timeout=timeout)
    finally:
        httpd.shutdown()

    if not res.startswith('OK'):
        print('  ! 取数失败：%s' % res[:200])
        return 4
    if not RECV['done'].wait(30) or not RECV['data']:
        print('  ! 页面已返回，但本机未收到 CSV（可能被浏览器拦截 localhost 回传）')
        return 4
    open(dst, 'wb').write(RECV['data'])
    print('  %s（%d 字节）' % (dst, len(RECV['data'])))

    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'oa2xlsx.py')
    xlsx = os.path.join(period, '合约订单评审表%s.xlsx' % (report or ''))
    r = subprocess.run([sys.executable, script, xlsx],
                       capture_output=True, text=True, encoding='utf-8')
    if r.returncode != 0:
        print('  ! oa2xlsx 失败：%s' % (r.stderr or r.stdout)[-300:])
        return 4
    print('  %s' % (r.stdout or '').strip())
    return 0


if __name__ == '__main__':
    sys.exit(main())
