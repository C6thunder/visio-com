# example-2x2-matrix.ps1 — 完整示例：2×2 因子对照矩阵
#
# 运行：
#   $code = Get-Content -Raw -Encoding UTF8 '.\scripts\visio-lib.ps1'
#   . ([scriptblock]::Create($code))
#   $code2 = Get-Content -Raw -Encoding UTF8 '.\scripts\example-2x2-matrix.ps1'
#   Invoke-Expression $code2
#
# 产出：fig-2x2.vsd（或 .vsd 回退）+ fig-2x2.png（1424×801 取决于页面尺寸）

param(
  [string]$OutDir = (Join-Path (Get-Location) 'out'),
  [string]$BaseName = 'fig-2x2'
)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$ctx = Start-VisioSession -Width 13.333 -Height 7.5
$page = $ctx.Page
$st = Open-VisioStencil -App $ctx.App -Names @('BASIC_M', 'ARROWS_M', 'ANNOT_M')

# ---- Okabe-Ito 单色递进 ----
$ink = 'RGB(26,26,26)'; $gray = 'RGB(90,102,110)'; $light = 'RGB(148,157,164)'
$note = 'RGB(107,120,130)'; $axis = 'RGB(110,122,130)'; $blue = 'RGB(0,114,178)'
$fill0 = 'RGB(234,242,247)'; $edge0 = 'RGB(176,199,214)'; $name0 = 'RGB(60,88,106)'
$fill1 = 'RGB(183,213,229)'; $edge1 = 'RGB(0,114,178)';   $name1 = 'RGB(0,90,140)'
$fill2 = 'RGB(0,114,178)';   $edge2 = 'RGB(0,114,178)';   $name2 = 'RGB(255,255,255)'
$body2 = 'RGB(238,245,249)'

Add-VisioPageBackground -Page $page | Out-Null

# ---- 标题 ----
New-VisioText $page 0.60 6.86 1.20 7.18 '图 1' 15 $gray -Bold | Out-Null
New-VisioText $page 1.18 6.86 9.60 7.18 '2×2 因子对照：机制 A × 机制 B' 15 $ink -Bold | Out-Null
New-VisioText $page 9.60 6.88 12.85 7.16 '四组对照 · 基础安全控制统一保留' 8.5 $light -HAlign 2 | Out-Null

# ---- 四张卡 ----
$colL = @(1.60, 6.95); $colR = @(7.40, 12.75)
$rowBot = @(2.55, 4.00); $rowTop = @(4.42, 5.87)

function New-GroupCard($x1, $y1, $x2, $y2, $title, $onA, $lineA, $onB, $lineB, $note2, $fill, $edge, $nameColor, $bodyColor, $ring, $onFill, $offFill) {
  New-VisioCard $page $st $x1 $y1 $x2 $y2 $title @() $fill $edge $nameColor $bodyColor | Out-Null
  $yA = $y2 - 0.61; $yB = $y2 - 0.89
  New-VisioDot $page $st ($x1 + 0.30) $yA 0.13 -On:$onA -Ring $ring -OnFill $onFill -OffFill $offFill | Out-Null
  New-VisioText $page ($x1 + 0.46) ($yA - 0.11) ($x2 - 0.20) ($yA + 0.11) $lineA 9.5 $bodyColor | Out-Null
  New-VisioDot $page $st ($x1 + 0.30) $yB 0.13 -On:$onB -Ring $ring -OnFill $onFill -OffFill $offFill | Out-Null
  New-VisioText $page ($x1 + 0.46) ($yB - 0.11) ($x2 - 0.20) ($yB + 0.11) $lineB 9.5 $bodyColor | Out-Null
  New-VisioText $page ($x1 + 0.22) ($y1 + 0.14) ($x2 - 0.20) ($y1 + 0.36) $note2 8.5 $nameColor | Out-Null
}

New-GroupCard $colL[0] $rowBot[0] $colL[1] $rowBot[1] 'A 组 · 基线' $false '机制 A｜关闭（基线）' $false '机制 B｜关闭' '对照基准' $fill0 $edge0 $name0 $ink $blue $blue 'RGB(255,255,255)'
New-GroupCard $colR[0] $rowBot[0] $colR[1] $rowBot[1] 'B 组 · 单机制' $true '机制 A｜开启' $false '机制 B｜关闭' '对 A：机制 A 纯贡献' $fill1 $edge1 $name1 $ink $blue $blue 'RGB(255,255,255)'
New-GroupCard $colL[0] $rowTop[0] $colL[1] $rowTop[1] 'C 组 · 单机制' $false '机制 A｜关闭' $true '机制 B｜开启' '对 A：机制 B 纯贡献' $fill1 $edge1 $name1 $ink $blue $blue 'RGB(255,255,255)'
New-GroupCard $colR[0] $rowTop[0] $colR[1] $rowTop[1] 'D 组 · 完整版' $true '机制 A｜开启' $true '机制 B｜开启' '对 B、C：交互效应' $fill2 $edge2 $name2 $body2 $blue 'RGB(255,255,255)' $fill2
New-VisioDiamond $page $st ($colR[1] - 0.34) ($rowTop[1] - 0.26) 0.16 'RGB(255,255,255)' '' | Out-Null

# ---- 四向对照箭头 ----
New-VisioArrow $page ($colL[1] + 0.05) 3.275 ($colR[0] - 0.05) 3.275 $note 1.5 | Out-Null
New-VisioArrow $page ($colL[1] + 0.05) 5.145 ($colR[0] - 0.05) 5.145 $note 1.5 | Out-Null
New-VisioArrow $page 4.275 ($rowBot[1] + 0.05) 4.275 ($rowTop[0] - 0.05) $note 1.5 | Out-Null
New-VisioArrow $page 10.075 ($rowBot[1] + 0.05) 10.075 ($rowTop[0] - 0.05) $note 1.5 | Out-Null

# ---- 轴 ----
New-VisioAxisArrow $page $st 7.175 2.13 11.15 -Color $axis | Out-Null
New-VisioAxisArrow $page $st 1.42 4.21 3.50 -Vertical -Color $axis | Out-Null
New-VisioText $page 0.86 5.045 1.36 5.225 '开启' 9 $ink -Bold -HAlign 2 | Out-Null
New-VisioText $page 0.86 3.175 1.36 3.355 '关闭' 9 $ink -Bold -HAlign 2 | Out-Null
New-VisioVerticalText $page 0.61 4.21 1.70 0.26 '机制 B' 9.5 $gray | Out-Null
New-VisioText $page $colL[0] 2.24 $colL[1] 2.44 '水平 1（关闭）' 9.5 $ink -Bold -HAlign 1 | Out-Null
New-VisioText $page $colR[0] 2.24 $colR[1] 2.44 '水平 2（开启）' 9.5 $ink -Bold -HAlign 1 | Out-Null
New-VisioText $page 6.35 1.84 8.05 2.02 '机制 A' 9 $gray -HAlign 1 | Out-Null

# ---- 图例 ----
$leg = $page.DrawRectangle(0.60, 0.66, 4.55, 1.52)
Set-VisioCell $leg 'Rounding' '0.08 in' | Out-Null
Set-VisioCell $leg 'FillPattern' '1' | Out-Null; Set-VisioCell $leg 'FillForegnd' 'RGB(255,255,255)' | Out-Null
Set-VisioCell $leg 'LineColor' 'RGB(201,211,218)' | Out-Null; Set-VisioCell $leg 'LineWeight' '0.75 pt' | Out-Null

New-VisioText $page 0.74 1.28 1.30 1.44 '图例' 8 $gray -Bold | Out-Null
$sw = @(@(0.74, $fill0, $edge0), @(0.98, $fill1, $edge1), @(1.22, $fill2, $edge2))
foreach ($s in $sw) {
  $r = $page.DrawRectangle($s[0], 1.02, $s[0] + 0.20, 1.14)
  Set-VisioCell $r 'Rounding' '0.03 in' | Out-Null
  Set-VisioCell $r 'FillPattern' '1' | Out-Null; Set-VisioCell $r 'FillForegnd' $s[1] | Out-Null
  Set-VisioCell $r 'LineColor' $s[2] | Out-Null; Set-VisioCell $r 'LineWeight' '0.75 pt' | Out-Null
}
New-VisioText $page 1.50 1.00 4.40 1.18 '颜色深浅 = 已开启的机制项数（0 / 1 / 2）' 8 $gray | Out-Null
New-VisioDot $page $st 0.82 0.80 0.12 -On | Out-Null
New-VisioText $page 0.94 0.71 1.26 0.89 '开启' 8 $gray | Out-Null
New-VisioDot $page $st 1.46 0.80 0.12 | Out-Null
New-VisioText $page 1.58 0.71 1.90 0.89 '关闭' 8 $gray | Out-Null
New-VisioArrow $page 2.06 0.80 2.32 0.80 $note 1.25 | Out-Null
New-VisioText $page 2.40 0.71 3.60 0.89 '对照比较方向' 8 $gray | Out-Null

# ---- 图注 ----
New-VisioText $page 4.80 0.72 12.75 1.46 ('注：四组均保留基础安全控制；对照组不得以取消权限校验的方式实现。' + "`n" +
  '箭头表示对照比较方向：A→B 与 A→C 为单因子纯贡献，B→D 与 C→D 为末项增量。') 8.5 $gray | Out-Null

# ---- 交付 ----
$vsd = Join-Path $OutDir ($BaseName + '.vsdx')
$png = Join-Path $OutDir ($BaseName + '.png')
$saved = Save-VisioDoc -Doc $ctx.Doc -Page $page -VsdPath $vsd -PngPath $png

Write-Output ('[example] 形状数: ' + $page.Shapes.Count)
Get-VisioShapeReport -Page $page -Stencils $st |
  Group-Object Family | ForEach-Object { Write-Output ('[example] 模具族 ' + $_.Name + ' × ' + $_.Count) }
Write-Output ('[example] 保存: ' + $saved)
Write-Output ('[example] 导出: ' + $png)

Close-VisioSession -App $ctx.App -Stencils $st
