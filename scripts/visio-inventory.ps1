# visio-inventory.ps1 — 扫描本机 Visio 模具库(.vssx) 与模板(.vstx)，产出可检索清单
#
# 用法：
#   $code = Get-Content -Raw -Encoding UTF8 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\visio-inventory.ps1'
#   Invoke-Expression $code
#
# 关键点：**不启动 Visio**。模具库是 OOXML 包，master 名单在 visio/masters/masters.xml 里，
# 直接解包读取即可——362 个模具库 / 9112 个 master 约 3 秒；
# 用 COM 逐个 Documents.Open 反而要 20~30 分钟（还会残留进程）。
#
# 产出（写到 ../references）：
#   visio-masters-full.tsv    模具库 \t 模具名 \t 英文名 \t 类型 \t 隐藏标记
#   visio-stencil-summary.md  每个模具库的 master 数量与名称摘要
#   visio-templates.tsv       模板文件 \t 模板名 \t 配套模具 \t 页面形状数
param(
  [string]$Root   = 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052',
  [string]$OutDir = ''
)
$ErrorActionPreference = 'Continue'

if (-not $OutDir) {
  $skillRoot = if ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot } else { 'C:\Users\thunder\.agents\skills\visio-diagram' }
  $OutDir = Join-Path $skillRoot 'references\data'
}
$DocDir = Join-Path (Split-Path -Parent $OutDir) ''   # 供文档类产物使用
if (-not (Test-Path $Root)) { throw "找不到模具目录: $Root" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ZipText {
  param([string]$ZipPath, [string]$EntryName)
  try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
      $e = $zip.Entries | Where-Object { $_.FullName -eq $EntryName } | Select-Object -First 1
      if (-not $e) { return $null }
      $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
      try { return $sr.ReadToEnd() } finally { $sr.Close() }
    } finally { $zip.Dispose() }
  } catch { return $null }
}

function Get-MasterList {
  <# 从 masters.xml 解析出：名称、英文名、类型、是否隐藏 #>
  param([string]$Xml)
  if (-not $Xml) { return @() }
  $out = New-Object System.Collections.Generic.List[object]
  foreach ($m in [regex]::Matches($Xml, '<Master\b[^>]*>')) {
    $tag = $m.Value
    $u = ''; if ($tag -match '\bNameU="([^"]*)"') { $u = $Matches[1] }
    $rest = $tag -replace '\bNameU="[^"]*"', ''
    $n = ''; if ($rest -match '\bName="([^"]*)"') { $n = $Matches[1] }
    $t = ''; if ($tag -match '\bMasterType="([^"]*)"') { $t = $Matches[1] }
    $h = ''; if ($tag -match '\bHidden="([^"]*)"') { $h = $Matches[1] }
    if (-not $n) { $n = $u }
    $out.Add([pscustomobject]@{ Name = $n; NameU = $u; Type = $t; Hidden = $h })
  }
  return $out
}

# ---------- 1. 模具库 ----------
$tsv = New-Object System.Collections.Generic.List[string]
$tsv.Add("# 模具库`t模具名`t英文名(NameU)`t类型`t隐藏`t# 扫描时间 $stamp`t# 根目录 $Root")
$md  = New-Object System.Collections.Generic.List[string]
$md.Add('# Visio 模具库摘要（本机扫描）')
$md.Add("# 根目录: $Root    扫描时间: $stamp")
$md.Add('')

$stencils = Get-ChildItem $Root -Filter *.vssx | Sort-Object Name
$total = 0
foreach ($f in $stencils) {
  $xml = Get-ZipText -ZipPath $f.FullName -EntryName 'visio/masters/masters.xml'
  $list = Get-MasterList -Xml $xml
  foreach ($m in $list) {
    # 名字里可能带换行，会把 TSV 拆成多行，破坏 grep——先压成空格
    $nm  = ($m.Name  -replace '[\r\n\t]', ' ')
    $nmU = ($m.NameU -replace '[\r\n\t]', ' ')
    $tsv.Add(($f.BaseName + "`t" + $nm + "`t" + $nmU + "`t" + $m.Type + "`t" + $m.Hidden))
  }
  $total += $list.Count
  $names = ($list | ForEach-Object { $_.Name }) -join ' · '
  $md.Add(('## ' + $f.BaseName + '  （' + $list.Count + ' 个）'))
  $md.Add($names)
  $md.Add('')
}
[System.IO.File]::WriteAllLines((Join-Path $OutDir 'visio-masters-full.tsv'), $tsv, (New-Object System.Text.UTF8Encoding $true))
[System.IO.File]::WriteAllLines((Join-Path $OutDir 'visio-stencil-summary.md'), $md, (New-Object System.Text.UTF8Encoding $true))
# 计数以「写出来的文件」为准，避免内存计数器与落盘结果不一致
$written = ([System.IO.File]::ReadAllLines((Join-Path $OutDir 'visio-masters-full.tsv'), [System.Text.Encoding]::UTF8)).Count - 1
Write-Output ("[inv] 模具库 " + $stencils.Count + " 个 / master " + $written + " 个")

# ---------- 2. 模具库一览：文件名 ↔ 英文显示名 ↔ master 数 ----------
$stenTsv = Join-Path $OutDir 'visio-stencils.tsv'
$sl = New-Object System.Collections.Generic.List[string]
$sl.Add("# 模具库文件`t英文显示名(Visio 界面用)`tmaster 数`t# 扫描时间 $stamp")
foreach ($f in $stencils) {
  $cx = Get-ZipText -ZipPath $f.FullName -EntryName 'docProps/custom.xml'
  $disp = $f.BaseName
  if ($cx) {
    if ($cx -match '(?s)_VPID_ALTERNATENAMES[^\>]*>\s*<vt:[^>]*>([^<]*)<') {
      $names = ($Matches[1] -split ';') | Where-Object { $_ -and $_ -notmatch '^(BEGIN|END)' }
      if ($names) { $disp = $names[0] }
    }
  }
  $cnt = 0
  $xml2 = Get-ZipText -ZipPath $f.FullName -EntryName 'visio/masters/masters.xml'
  if ($xml2) { $cnt = ([regex]::Matches($xml2, '<Master\b')).Count }
  $sl.Add(($f.BaseName + "`t" + $disp + "`t" + $cnt))
}
[System.IO.File]::WriteAllLines($stenTsv, $sl, (New-Object System.Text.UTF8Encoding $true))
Write-Output ("[inv] 模具库一览 → " + $stenTsv)

# ---------- 3. 模板（.vstx）：模板 = 预制好的一页 + 配套模具 ----------
$tl = New-Object System.Collections.Generic.List[string]
$tl.Add("# 模板文件`t模板名`t配套模具`t页面形状数`t# 扫描时间 $stamp")
$tpls = Get-ChildItem $Root -Filter *.vstx | Sort-Object Name
foreach ($v in $tpls) {
  $list = Get-MasterList -Xml (Get-ZipText -ZipPath $v.FullName -EntryName 'visio/masters/masters.xml')
  $pgXml = Get-ZipText -ZipPath $v.FullName -EntryName 'visio/pages/page1.xml'
  $shapes = 0
  if ($pgXml) { $shapes = ([regex]::Matches($pgXml, '<Shape\b')).Count }
  $tl.Add(($v.BaseName + "`t" + $v.BaseName + "`t" + (($list | ForEach-Object { $_.Name }) -join ', ') + "`t" + $shapes))
}
[System.IO.File]::WriteAllLines((Join-Path $OutDir 'visio-templates.tsv'), $tl, (New-Object System.Text.UTF8Encoding $true))
Write-Output ("[inv] 模板 " + $tpls.Count + " 个")
Write-Output ('[inv] 输出目录: ' + $OutDir)
