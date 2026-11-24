<#
一键把本 skill 同步到 GitHub（私有库 wye005/dongzhi-ledger-skill）

用法：
  .\sync.ps1
  .\sync.ps1 -Message "补充 F4 产品类别映射"
#>
param([string]$Message = "")

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not $Message) {
  $Message = '同步 ' + (Get-Date -Format 'yyyy-MM-dd HH:mm')
}

git add -A
$staged = @(git diff --cached --name-only)
if ($staged.Count -eq 0) {
  Write-Host '没有变更，无需同步' -ForegroundColor DarkGray
  return
}

git commit -q -m $Message
git push -q origin main
Write-Host ("已同步 {0} 个文件 -> origin/main" -f $staged.Count) -ForegroundColor Green
$staged | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
