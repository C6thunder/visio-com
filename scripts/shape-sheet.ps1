# shape-sheet.ps1 — 形状样张：把候选母版的真实外观排成一页
#
# 为什么需要：母版名字看不出长相。MATRIXDOT 的“点 1”带黄色编号徽标、
# PROCESS 的“V 形步骤”改不动尺寸——这些只有拖出来看才知道。
# 选型前先跑一遍这个脚本，别靠名字猜。
#
# 运行：
#   $code = Get-Content -Raw -Encoding UTF8 '.\scripts\visio-lib.ps1'
#   . ([scriptblock]::Create($code))
#   $code2 = Get-Content -Raw -Encoding UTF8 '.\scripts\shape-sheet.ps1'
#   Invoke-Expression $code2

param(
  [string]$OutDir = (Join-Path (Get-Location) 'out'),
  [string]$OutName = 'shape-sheet',
  # 每组：模具名 → 要打样的母版名（有序字典，保证打印顺序稳定）
  $Groups = [ordered]@{
    'BASIC_M'    = @('矩形', '圆形', '菱形', '六边形', '三角形')
    'BASFLO_M'   = @('流程', '判定', '开始/结束', '文档', '数据库')
    'ARROWS_M'   = @('普通箭头', '普通双箭头', '现代箭头', '弧线箭头')
    'ANNOT_M'    = @('圆形标注', '信息线')
    'MATRIXDOT_M' = @('点 1', '菱形 2', '星形 1')
    'PROCESS_M'  = @('箭头步骤', 'V 形步骤')
  }
)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$ctx = Start-VisioSession -Width 13.333 -Height 7.5
$page = $ctx.Page
$st = Open-VisioStencil -App $ctx.App -Names @($Groups.Keys)
Add-VisioPageBackground -Page $page | Out-Null

$white = 'RGB(255,255,255)'; $ink = 'RGB(26,26,26)'; $gray = 'RGB(90,102,110)'; $light = 'RGB(150,158,165)'
$cellFill = 'RGB(234,242,247)'; $cellEdge = 'RGB(176,199,214)'

New-VisioText $page 0.5 6.85 9.0 7.2 'Visio 形状样张（选型前先看这个）' 14 $ink -Bold | Out-Null
New-VisioText $page 9.0 6.90 12.85 7.16 '仅示意外观，不含最终配色' 8 $light -HAlign 2 | Out-Null

# 平铺：每行 6 个
$items = @()
foreach ($k in $Groups.Keys) { foreach ($m in $Groups[$k]) { $items += , @($k, $m) } }

$colW = 2.02; $x0 = 0.62; $rowH = 1.42; $y0 = 5.55
$drops = 0
for ($i = 0; $i -lt $items.Count; $i++) {
  $r = [Math]::Floor($i / 6); $c = $i % 6
  $cx = $x0 + $c * $colW + $colW / 2
  $cy = $y0 - $r * $rowH
  if ($cy -lt 0.2) { break }

  $cell = $page.DrawRectangle($cx - 0.92, $cy - 0.58, $cx + 0.92, $cy + 0.58)
  Set-VisioCell $cell 'Rounding' '0.08 in' | Out-Null
  Set-VisioCell $cell 'FillPattern' '1' | Out-Null
  Set-VisioCell $cell 'FillForegnd' $cellFill | Out-Null
  Set-VisioCell $cell 'LineColor' $cellEdge | Out-Null
  Set-VisioCell $cell 'LineWeight' '0.75 pt' | Out-Null

  $sn = $items[$i][0]; $mn = $items[$i][1]
  $m = Get-VisioMaster $st[$sn] $mn
  if ($m) {
    try {
      $shp = $page.Drop($m, $cx, ($cy + 0.10))
      if ($shp.CellsU('Width').ResultIU -gt 1.2) { Set-VisioCell $shp 'Width' '1.1 in' | Out-Null }
      if ($shp.CellsU('Height').ResultIU -gt 0.8) { Set-VisioCell $shp 'Height' '0.7 in' | Out-Null }
      $drops++
    } catch { Write-Warning ("拖放失败: $sn / $mn") }
  } else { Write-Warning ("母版不存在: $sn / $mn") }

  New-VisioText $page ($cx - 0.92) ($cy - 0.55) ($cx + 0.92) ($cy - 0.29) ($mn + "`n" + $sn.Replace('_M', '')) 7 $gray -HAlign 1 | Out-Null
}

$png = Join-Path $OutDir ($OutName + '.png')
$vsd = Join-Path $OutDir ($OutName + '.vsdx')
$saved = Save-VisioDoc -Doc $ctx.Doc -Page $page -VsdPath $vsd -PngPath $png
Write-Output ('[sheet] 成功拖放 ' + $drops + ' / ' + $items.Count + ' 个母版')
Write-Output ('[sheet] 保存: ' + $saved)
Write-Output ('[sheet] 导出: ' + $png)
Close-VisioSession -App $ctx.App -Stencils $st
