<#
东智台账 · 唯一入口（2026-11-23）
一条管线：取源 -> 解析 -> 拼行 -> 落库 -> 派生F0 -> 自检 -> 门禁
月底加 -Close：多跑一次关账（月度事实/快照/monthly.json），F1–F6 再基于它出。

用法：
  .\run.ps1 -ReportDate 2026-11-20
  .\run.ps1 -ReportDate 2026-11-30 -Close
  .\run.ps1 -ReportDate 2026-12-04 -BaseF0 "E:\...\截止到2026年11月30日...汇总表.xlsx"

说明：
  - 周更/月更数据规则完全相同（可补录、可作废、状态可更新），不再分两套流程。
  - 源窗口一律整月（dz_ingest 自动），<起> 只影响产出快照区间与文案。
  - 任一环节退出码非 0 立即停，不带着 FAIL 往下走。
#>
param(
  [Parameter(Mandatory = $true)][string]$ReportDate,
  [string]$Period = "",
  [string]$BaseF0 = "",
  [switch]$Close,
  [switch]$SkipStage,
  [switch]$SkipOA,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$env:PYTHONIOENCODING = 'utf-8'
$Py  = 'C:\Users\11447\.workbuddy\binaries\python\envs\default\Scripts\python.exe'
$Dir = 'C:\Users\11447\.workbuddy\skills\dongzhi-ledger\scripts'
$Root = 'E:\商务部\东智日常表更新'

$d2 = [datetime]::ParseExact($ReportDate, 'yyyy-MM-dd', $null)
$month = $d2.ToString('yyyy-MM')

if (-not $Period) {
  $hit = Get-ChildItem -LiteralPath $Root -Directory |
         Where-Object { $_.Name -match '^\d{8}-\d{8}$' -and $_.Name.EndsWith($d2.ToString('yyyyMMdd')) } |
         Sort-Object Name | Select-Object -Last 1
  if ($hit) { $Period = $hit.Name }
  else { $Period = $d2.AddDays(-6).ToString('yyyyMMdd') + '-' + $d2.ToString('yyyyMMdd') }
}
$Src = Join-Path $Root $Period
$d1 = [datetime]::ParseExact($Period.Substring(0, 8), 'yyyyMMdd', $null).ToString('yyyy-MM-dd')

# 上次失败留下的临时 F0（.building）先清掉，正式 F0 不受影响
Get-ChildItem -LiteralPath $Src -Filter '*.building' -File -ErrorAction SilentlyContinue |
  ForEach-Object { [IO.File]::Delete($_.FullName) }

Write-Host ''
Write-Host "=== 东智台账 $Period  (D1=$d1  D2=$ReportDate) ===" -ForegroundColor Cyan

function Invoke-Step([string]$Name, [string]$Exe, [string[]]$Argv) {
  Write-Host ''
  Write-Host "--- $Name ---" -ForegroundColor Yellow
  if ($DryRun) { Write-Host "[dry] $Exe $($Argv -join ' ')" -ForegroundColor DarkGray; return }
  & $Exe @Argv
  if ($LASTEXITCODE -ne 0) { throw "$Name 失败（exit=$LASTEXITCODE），已停止" }
}

# 0a 取 OA（浏览器会话过期时不致命：沿用周期目录里已有的 OA 文件）
if (-not $SkipOA) {
  Write-Host ''
  Write-Host '--- 0a/7 OA 取数 ---' -ForegroundColor Yellow
  if (-not $DryRun) {
    & $Py "$Dir\oa_pull.py" $Src --report $ReportDate
    if ($LASTEXITCODE -ne 0) {
      Write-Host '   OA 未就绪，沿用现有 OA 文件继续' -ForegroundColor DarkYellow
    }
  }
}

if (-not $SkipStage) {
  $stageArgs = @("$Dir\dz_sources.py", 'stage', $Src, $ReportDate)
  if ($Close) { $stageArgs += '--close' }
  Invoke-Step '0/7 取源落盘' $Py $stageArgs
}

Invoke-Step '1/7 解析源文件'   $Py @("$Dir\dz_ingest.py", $Src, $d1, $ReportDate)
Invoke-Step '2/7 拼台账行'     $Py @("$Dir\build_ledger.py", $Src, $d1, $ReportDate)
Invoke-Step '3/7 落库'         $Py @("$Dir\dz_store.py", $d1, $ReportDate)

$deriveArgs = @("$Dir\dz_derive.py", $Src, $d1, $ReportDate)
if ($BaseF0) { $deriveArgs += $BaseF0 }
Invoke-Step '4/7 派生 F0'      $Py $deriveArgs

Invoke-Step '5/7 产出自检'     $Py @("$Dir\dz_check.py", $Src, $d1, $ReportDate)
Invoke-Step '6/7 三层门禁'     $Py @("$Dir\dz_gate.py", $Src, $d1, $ReportDate)

$f0 = Get-ChildItem -LiteralPath $Src -Filter '截止到*汇总表*.xlsx' -File -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime | Select-Object -Last 1

if ($Close) {
  $closeArgs = @("$Dir\dz_close.py", $month, '--from-ledger', '--income', 'finance')
  if ($f0) { $closeArgs += @('--f0', $f0.FullName) }
  Invoke-Step '7/7 月结' $Py $closeArgs
  Write-Host ''
  Write-Host "月结完成。接下来只做报表：F1 -> F2 -> F3 -> F4 -> F5/F6 -> PPT" -ForegroundColor Green
  Write-Host "  数据唯一来源：$Root\monthly\$month`_monthly.json"
} else {
  Write-Host ''
  Write-Host '周更完成（F0 已更新）。月底再加 -Close 出 F1–F6。' -ForegroundColor Green
}

if ($f0) { Write-Host "F0: $($f0.FullName)" -ForegroundColor DarkGray }
