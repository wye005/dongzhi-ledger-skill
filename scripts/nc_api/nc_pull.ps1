<#
.SYNOPSIS
  东智 NC 6.5 后端直连取数（绕过 UClient GUI）。

.EXAMPLE
  .\nc_pull.ps1 -SalesOut -Start 2026-10-01 -End 2026-11-01 -Out E:\tmp\nc_202610.csv
  .\nc_pull.ps1 -SqlFile .\my.sql -Out E:\tmp\result.csv
#>
[CmdletBinding(DefaultParameterSetName='Sql')]
param(
  [Parameter(ParameterSetName='Sql', Mandatory=$true)][string]$SqlFile,
  [Parameter(ParameterSetName='SalesOut', Mandatory=$true)][switch]$SalesOut,
  [Parameter(ParameterSetName='SalesOut', Mandatory=$true)][string]$Start,
  [Parameter(ParameterSetName='SalesOut', Mandatory=$true)][string]$End,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$User = '',
  [string]$Pwd  = '',
  [string]$Bc   = '001',
  [string]$Url  = '',
  [string]$Org  = '0001A110000000003V0Q',
  [string]$JavaHome = ''
)

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot

# 凭据：本地 nc.local.ps1（已 gitignore）> 环境变量 NC_USER/NC_PWD/NC_URL
$localCfg = Join-Path $here 'nc.local.ps1'
if (Test-Path -LiteralPath $localCfg) {
  . $localCfg
  if (-not $User -and $NcUser) { $User = $NcUser }
  if (-not $Pwd  -and $NcPwd)  { $Pwd  = $NcPwd }
  if (-not $Url  -and $NcUrl)  { $Url  = $NcUrl }
}
if (-not $User) { $User = $env:NC_USER }
if (-not $Pwd)  { $Pwd  = $env:NC_PWD }
if (-not $Url)  { $Url  = $env:NC_URL }
if (-not $User -or -not $Pwd -or -not $Url) {
  throw '缺少 NC 连接信息：请创建 nc_api\nc.local.ps1（$NcUser/$NcPwd/$NcUrl）或设置环境变量 NC_USER/NC_PWD/NC_URL'
}

if (-not $JavaHome) {
  foreach ($c in @('E:\ideaIU\jbr', $env:JAVA_HOME)) {
    if ($c -and (Test-Path (Join-Path $c 'bin\javac.exe'))) { $JavaHome = $c; break }
  }
}
if (-not $JavaHome) { throw '未找到 JDK（javac.exe）。请用 -JavaHome 指定。' }
$javac = Join-Path $JavaHome 'bin\javac.exe'
$java  = Join-Path $JavaHome 'bin\java.exe'

$code = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'UClient\apps') -Directory -ErrorAction SilentlyContinue |
  ForEach-Object { Join-Path $_.FullName 'nc_client_home\NCCACHE\CODE' } |
  Where-Object { Test-Path (Join-Path $_ 'external\lib\fwpub.jar') } |
  Select-Object -First 1
if (-not $code) { throw '未找到 UClient 客户端 CODE 目录（fwpub.jar）。' }

$parts = @((Join-Path $code 'external\classes'), (Join-Path $code 'modules'), $here)
$parts += Get-ChildItem (Join-Path $code 'external\lib') -Filter *.jar | ForEach-Object { $_.FullName }
$parts += Get-ChildItem (Join-Path $code 'modules') -Recurse -Filter *.jar | ForEach-Object { $_.FullName }
$cp = (($parts -join ';') -replace '\\','/')

$build = Join-Path $here '.build'
New-Item -ItemType Directory -Force $build | Out-Null
$runArgs = Join-Path $build 'java.args'
$cplArgs = Join-Path $build 'javac.args'
Set-Content -Path $runArgs  -Value ('-cp "' + $cp + '"') -Encoding ASCII
Set-Content -Path $cplArgs -Value ('-cp "' + $cp + '"') -Encoding ASCII

$cls = Join-Path $here 'NcSql.class'
$srcFile = Join-Path $here 'NcSql.java'
if (-not (Test-Path $cls) -or (Get-Item $srcFile).LastWriteTime -gt (Get-Item $cls).LastWriteTime) {
  & $javac "@$cplArgs" -encoding UTF-8 -d $here $srcFile
  if ($LASTEXITCODE -ne 0) { throw 'NcSql 编译失败' }
}

if ($SalesOut) {
  $sql = @"
select h.vbillcode as billcode, h.dbilldate as billdate,
       h.ccustomerid as cust_id, cu.code as cust_code, cu.name as cust_name,
       b.crowno as rowno, b.cmaterialvid as mat_id, m.code as mat_code, m.name as mat_name,
       b.nnum as qty, b.norigmny as amt, b.norigtaxmny as tax_amt,
       b.nqtorigprice as price, b.vbatchcode as batch,
       b.cbodywarehouseid as wh_id
from ic_saleout_h h
join ic_saleout_b b on h.cgeneralhid = b.cgeneralhid
left join bd_material m on m.pk_material = b.cmaterialvid
left join bd_customer cu on cu.pk_customer = h.ccustomerid
where h.pk_org = '$Org'
  and h.dbilldate >= '$Start' and h.dbilldate < '$End'
  and h.dr = 0 and b.dr = 0
order by h.vbillcode, b.crowno
"@
  $tmp = Join-Path $build ('salesout_' + [Guid]::NewGuid().ToString('N') + '.sql')
  [IO.File]::WriteAllText($tmp, $sql, (New-Object Text.UTF8Encoding($false)))
  $SqlFile = $tmp
}

$sqlFull = (Resolve-Path $SqlFile).Path
$outFull = [IO.Path]::GetFullPath($Out)
& $java "@$runArgs" NcSql $Url $User $Pwd $Bc $sqlFull $outFull
if ($LASTEXITCODE -ne 0) { throw "NcSql 执行失败，exit=$LASTEXITCODE" }
Write-Host "OK -> $outFull"
