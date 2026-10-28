# new-figure.ps1 — 画图骨架（复制这个文件改，不要从零写）
#
# 用法：
#   $code = Get-Content -Raw -Encoding UTF8 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\new-figure.ps1'
#   Invoke-Expression $code                        # 不传参，用脚本默认参数
#   . ([scriptblock]::Create($code)) -Name 'fig-1'  # 要传参必须用 dot-source + 调用运算符 &
#   （注意：Invoke-Expression 没有 -Args 参数；& ([scriptblock]::Create($code)) @{} 才是带参写法）
#   更常见的做法：把本文件复制到工作目录改名后直接改。
#
# 三段式（每段都对应一个踩过的坑）：
#   阶段 1  绘制 + 断言   → 画之前先确认页面是空的（防"画了两遍"），画完打印形状数
#   阶段 2  存盘 + 导出   → .vsdx 优先、PNG 预览、期刊版再缩到成品尺寸出 PDF
#   阶段 3  高分辨率位图  → **另开一个不可见会话**打开已保存文件来放大导出，
#                          否则会就地改主文档，Visio 窗口里会看到一次放大重绘
param(
  [string]$OutDir = 'C:\Users\thunder\Desktop\DSHTs\out',
  [string]$Name = 'fig-1',
  [ValidateSet('slides','journal')][string]$Preset = 'slides',
  [switch]$HighRes
)

# ---------- 载入工具箱 ----------
$skillDir = if ($PSScriptRoot) { $PSScriptRoot } else { 'C:\Users\thunder\.agents\skills\visio-diagram\scripts' }
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 (Join-Path $skillDir 'visio-lib.ps1'))))

# ---------- 版面参数 ----------
# 母版 13.333×7.5 in（16:9）；期刊成品双栏 183 mm 宽
$MasterW = 13.333; $MasterH = 7.5
$JournalW = 7.2047; $JournalH = 4.2529   # 183 × 108 mm
$ScaleF = $JournalW / $MasterW

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$vsdPath = Join-Path $OutDir ($Name + '.vsdx')
$pngPath = Join-Path $OutDir ($Name + '.png')
$pdfPath = Join-Path $OutDir ($Name + '.pdf')

# ---------- 阶段 1：绘制 ----------
Write-Output '[build] 阶段 1/3：开始绘制'
$ctx = Start-VisioSession -Width $MasterW -Height $MasterH
$app = $ctx.App; $doc = $ctx.Doc; $page = $ctx.Page
$st = Open-VisioStencil -App $app -Names @('BASIC_M', 'ARROWS_M', 'BASFLO_M')

# 防"画完又画一遍"：页面必须从空白开始
if ($page.Shapes.Count -gt 0) { throw ('页面不是空的（' + $page.Shapes.Count + ' 个形状），疑似重复绘制') }

# 满页白底：没有它，Page.Export 只裁图形外接框，四周留白会被切掉
Add-VisioPageBackground -Page $page | Out-Null

# ===== 在这里画图 =====
# 坐标是左下/右上（Visio 的 y 轴向上），单位英寸
# 容器：
#   New-VisioCard  -Page $page -Stencils $st -X1 0.6 -Y1 3.2 -X2 4.0 -Y2 5.6 -Title '标题' -Lines @('第一行','第二行')
# 文字：
#   New-VisioText  -Page $page -X1 0.6 -Y1 6.2 -X2 6.0 -Y2 6.8 -Text '说明' -Size 9.5
# 标记点：
#   New-VisioDot   -Page $page -Stencils $st -CenterX 1.0 -CenterY 2.0 -Diameter 0.13 -On
# 菱形：
#   New-VisioDiamond -Page $page -Stencils $st -CenterX 5.0 -CenterY 2.0 -Size 0.16
# 单向箭头（线 + 实心箭头端；别用 ARROWS_M「普通箭头」，方向被锁死）：
#   New-VisioArrow -Page $page -X1 1.0 -Y1 1.5 -X2 4.0 -Y2 1.5
# 双向轴：
#   New-VisioAxisArrow -Page $page -Stencils $st -CenterX 2.5 -CenterY 1.0 -Length 3.0
# 母版形状（按中文名取，注意同名要带库名）：
#   $m = Get-VisioMaster $st['BASFLO_M'] '判定'; $s = $page.Drop($m, 3.0, 3.0)
# ===== 画图结束 =====

$count = $page.Shapes.Count
Write-Output ('[build] 绘制完成：形状 ' + $count)
Get-VisioShapeReport -Page $page -Stencils $st | Group-Object Family |
  ForEach-Object { Write-Output ('[build]   模具族 ' + $_.Name + ' × ' + $_.Count) }

# ---------- 阶段 2：存盘 + 预览（主会话只保存一次，**不做任何就地缩放**） ----------
Write-Output '[build] 阶段 2/3：存盘与预览'
$saved = Save-VisioDoc -Doc $doc -Page $page -VsdPath $vsdPath -PngPath $pngPath
Write-Output ('[build] 源文件: ' + $saved)
Write-Output ('[build] 预览图: ' + $pngPath)
Close-VisioSession -App $app -Stencils $st

# ---------- 阶段 3：派生期刊件（独立不可见会话，主文档不重绘） ----------
if ($Preset -eq 'journal') {
  Write-Output '[build] 阶段 3/3：派生期刊件（独立会话，主文档不重绘）'
  $deliv = New-VisioJournalDeliverables `
    -SourceVsd $saved `
    -OutDir (Join-Path $OutDir 'journal') `
    -Prefix $Name `
    -WidthIn $JournalW -HeightIn $JournalH -Factor $ScaleF `
    -MinFontPt 5.0 -MinLinePt 0.5 `
    -HighRes:$HighRes -HighResZoom 4
  Write-Output ('[build] 期刊版: ' + $deliv.Vsd)
  Write-Output ('[build] 矢量 PDF: ' + $deliv.Pdf)
  Write-Output ('[build] 预览图: ' + ($deliv.Preview -join ' | '))
  Write-Output ('[build] 高分辨率位图: ' + ($deliv.HiRes -join ' | '))
}

Write-Output '[build] 完成。'
