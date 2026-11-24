import json
import re

SRC = r"E:\workbuddy-temp\tmp\oa_review.json"
OUT = r"E:\workbuddy-temp\tmp\oa_review_fixed.json"

s = open(SRC, encoding="utf-8").read()
# 修复非法反斜杠转义：只保留合法转义 \ " \ / b f n r t 和 \uXXXX
s2 = re.sub(r'\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})', lambda m: '\\\\', s)
j = json.loads(s2)
rows = j["children"]
print("rows", len(rows), "total", j.get("total"))
print("keys:", " | ".join(rows[0].keys()))
sep = [r for r in rows if str(r.get("App_Date", "")).startswith("2026-09")]
print("9月行数", len(sep), " 9月订单数", len(set(r["SN"] for r in sep)))
print("9月日期范围", min(r["App_Date"] for r in sep), max(r["App_Date"] for r in sep))
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(j, f, ensure_ascii=False)
print("saved", OUT)
