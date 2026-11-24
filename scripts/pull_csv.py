import json
import os
import subprocess
import sys

AB = r"C:\Users\11447\.workbuddy\binaries\node\versions\22.22.2\node_modules\agent-browser\bin\agent-browser-win32-x64.exe"
ENV = dict(os.environ, PLAYWRIGHT_BROWSERS_PATH=r"E:\playwright-browsers")
OUT = r"E:\workbuddy-temp\tmp\oa_review_detail.csv"

GEN = r"""(async function(){
var r=await fetch('/YZSoft/DZ/SalesOrder/NewOrderReviewDetailData.ashx?method=GetData&searchType=QuickSearch&page=1&start=0&limit=5000');
var txt=await r.text();
var j=new Function('return ('+txt+')')();
var rows=j.children||[];if(!rows.length)return 'empty';
var keys=Object.keys(rows[0]);
var esc=function(v){if(v===null||v===undefined)return '';return '"'+String(v).replace(/"/g,'""').replace(/[\r\n]+/g,' ')+'"'};
var lines=[keys.map(esc).join(',')];
rows.forEach(function(o){lines.push(keys.map(function(k){return esc(o[k])}).join(','))});
var csv='\ufeff'+lines.join('\r\n');
var bytes=new TextEncoder().encode(csv);
var bin='';for(var i=0;i<bytes.length;i+=8192){bin+=String.fromCharCode.apply(null,bytes.subarray(i,i+8192));}
window.__b64=btoa(bin);
return 'rows='+rows.length+' b64len='+window.__b64.length;
})()"""


def ev(expr):
    p = subprocess.run([AB, "eval", expr], capture_output=True, text=True, encoding="utf-8", env=ENV)
    out = (p.stdout or "").strip()
    return out


print("生成:", ev(GEN)[:200])

chunk = 40000
total = None
buf = []
i = 0
while True:
    s = ev(f"window.__b64.slice({i*chunk},{(i+1)*chunk})")
    # 返回值是 JSON 编码的字符串
    try:
        val = json.loads(s)
    except Exception:
        print("解析失败块", i, s[:120])
        break
    if val == "":
        break
    buf.append(val)
    i += 1
    if i > 200:
        break

b64 = "".join(buf)
print("块数", i, "base64长度", len(b64))
import base64
data = base64.b64decode(b64)
open(OUT, "wb").write(data)
print("写出", OUT, len(data))
