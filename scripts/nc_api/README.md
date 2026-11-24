# nc_api · 东智 NC 6.5 后端取数

不经过 UClient GUI，直接调 NC `ServiceDispatcherServlet` 执行 SQL。

## 用法

```powershell
# 1) 销售出库明细（东智组织，按单据日期）
powershell -ExecutionPolicy Bypass -File .\nc_pull.ps1 `
  -SalesOut -Start 2026-10-01 -End 2026-11-01 -Out E:\tmp\nc_202610.csv

# 2) 任意 SQL（UTF-8 文件，分号可省略）
powershell -ExecutionPolicy Bypass -File .\nc_pull.ps1 `
  -SqlFile .\my.sql -Out E:\tmp\result.csv
```

账号/口令/地址见本地 `nc.local.ps1`（已 gitignore），或环境变量 `NC_USER` / `NC_PWD` / `NC_URL`；业务中心默认 `001`。
组织 `0001A110000000003V0Q`（深圳特发东智）。

`dz_ingest.py` 在周期目录找不到 `销售出库单明细表*.csv` 时会自动调本脚本（`DZ_NC_AUTO=0` 可关闭），
输出 `销售出库单明细表（东智）.csv` 后直接入库，无需手工导出。

## 输出列口径（2026-11-07）

- `qty` = 实发数量（全仓）；`tax_amt` = **含税金额**（对应台账「发货金额」）；`amt` = 无税金额，别混用。
- `wh_id` = 仓库 ID；台账「发货数量」只计成品仓 `1001A110000000XCQ1SX`（`RULES['NC_FINISHED_WH_ID']`），金额仍全仓计。
- 台账来源ID = `NC:单据编号#产品编码#NC行号`（`rowno`，不是 Excel 行号）。

## 原理（踩过的坑）

1. `POST /ServiceDispatcherServlet/default` 必须带 `X-Internal: true` + `X-Internal-Secret`（否则 403）。
2. `LoginRequest` 必须挂 `AuthenSubject`（key 为类全名），否则服务端 `StaticPWDVerifier` 空指针。
3. 登录成功后把 `NCSession` 回填 `UserExit`/`InvocationInfoProxy`，并 `UserExit.setKeepToken(true)`，否则后续调用报 `User session expired`。
4. `DBILLDATE` 是 `CHAR`，按字符串区间过滤；不要用 `DATE` 字面量。
5. `NcSql.java` 读取 SQL 时会去掉 UTF-8 BOM。
6. 需要桩类 `yonbip.bs.framework.server.util.BipSafeUtil`（已放在本目录）。
7. 发货仓库字段是 `IC_SALEOUT_B.CBODYWAREHOUSEID`（没有仓库名称，只能拿 ID 再映射）。

## 文件

- `NcSql.java` — 登录 + `IUAPQueryBS.executeQuery` + 写 UTF-8 CSV
- `nc_pull.ps1` — 自动定位 UClient CODE、拼 classpath、编译、执行
- `yonbip/.../BipSafeUtil.java|.class` — 反序列化必需桩类
