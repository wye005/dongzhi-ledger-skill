# dongzhi-ledger · 特发东智台账

把「东智总台账」从人工 Excel 维护，改成「SQLite 唯一真源 + Excel 看板」的一条管线。
周更与月更是同一套规则，月底只多一步月结与 F1–F6 附表。

## 一条管线

```
OA 评审 / NC 出库 / 交付跟踪 / 发票回款 / 财务收入
        │
        ├─ oa_pull.py        OA 接口自动取数（0a，非致命）
        ├─ dz_sources.py     邮件缓存/_inbox/周期目录 找最新源 + sha256
        ├─ dz_ingest.py      5 类源 → 统一行（来源ID 去重、状态可更新）
        ├─ build_ledger.py   Excel 镜像拼行
        ├─ dz_store.py       落库（增量/作废/补录）
        ├─ dz_derive.py      派生 F0 看板（只写日列，保留原公式）
        ├─ dz_check.py       13 项自检（源→台账→F0 三层）
        └─ dz_gate.py        门禁：PASS 才可交付
```

月底 `-Close`：`dz_close.py` 出 `monthly/YYYY-MM_monthly.json`，F1–F6 只消费这一份。

## 快速开始

```powershell
# 周更
.\scripts\run.ps1 -ReportDate 2026-11-20

# 月底（多出月结 + F1–F6 数据源）
.\scripts\run.ps1 -ReportDate 2026-11-30 -Close

# 跳过某一步
.\scripts\run.ps1 -ReportDate 2026-11-20 -SkipOA -SkipStage -DryRun
```

首次使用：把 `scripts/nc_api/nc.local.ps1.example` 复制为 `nc.local.ps1` 并填入 NC 账号/口令/地址（该文件已 gitignore）。

Python：`C:\Users\11447\.workbuddy\binaries\python\envs\default\Scripts\python.exe`（openpyxl / xlrd / pywin32）。

## 目录

| 路径 | 作用 |
|---|---|
| `SKILL.md` | 全量说明：口径、取数通道、踩坑记录 |
| `scripts/run.ps1` | 唯一入口 |
| `scripts/dz_*.py` | 取源 / 解析 / 落库 / 派生 / 自检 / 门禁 / 月结 |
| `scripts/nc_api/` | NC 6.5 后端直连（凭据见 `nc.local.ps1`） |
| `scripts/oa_pull.py` | OA 行级接口取数（ego-browser → agent-browser 回落） |

## 核心约定

- 数据库是唯一真源，Excel 只做看板；F0 保留原有公式，不锁死。
- 只动东智，其他单位数据不碰。
- 归属月按业务日期落定，不随导入批次漂移；可补录、可作废，作废不参与计算。
- 金额解析一律 `num_strict`，非空但解析失败必须报错，禁止静默当 0。
- 收入实际口径以财务《收入明细表》为准；F2 的累计是「实际 + 当月预计」的预计口径，两者不可混用。
- 9 月及以前为历史遗留，不追溯。
