# visio-lib.ps1 — Visio COM 自动化工具箱（DeepSeek Harness / Windows）
#
# 用法：
#   $lib = "C:\path\to\visio-lib.ps1"
#   . (Resolve-Path $lib)           # dot-source 后即可调用下面的函数
#   $ctx = Start-VisioSession
#   ... 画图 ...
#   Save-VisioDoc -Doc $ctx.Doc -Page $ctx.Page -VsdPath $vsd -PngPath $png
#   Close-VisioSession -App $ctx.App
#
# 设计前提（血泪教训，改动前先读 SKILL.md 的“硬规则”）：
#   * 中文文本必须显式指定字体（默认微软雅黑，本机 ID=245；等线=252），用 Get-VisioFontIndex 运行时查，别硬编码；Arial(21) 会让中文回退宋体
#   * 母版按中文名查找要用 Masters.Item("名字")，ItemU 取不到中文
#   * 组形状（块箭头、多数母版）填色要递归到子形状
#   * 部分母版锁死尺寸/方向，自动化改不动

$ErrorActionPreference = 'Stop'

$script:VisioExe = 'C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE'
$script:StencilRoot = 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052'
$script:CellWarnings = New-Object System.Collections.ArrayList

# ---------- 会话 ----------

function Get-VisioUnsavedDocumentNames {
  <#
    列出当前 Visio 实例里**未保存**的文档名；取不到实例时返回空数组（不报错）。
  #>
  $names = @()
  try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
    foreach ($d in @($app.Documents)) {
      try { if (-not $d.Saved) { $names += $d.Name } } catch {}
    }
  } catch {}
  return $names
}

function Stop-VisioProcess {
  <#
    结束 VISIO 进程，但**不牺牲用户未保存的工作**。
    Visio 是单实例：直接 Stop-Process -Force 会把用户自己开着、
    带未保存修改的文档一起杀掉，丢数据且无法恢复。
    这里先查未保存文档，有就只警告、不杀。
  #>
  param([switch]$Force)
  $procs = Get-Process VISIO -ErrorAction SilentlyContinue
  if (-not $procs) { return }
  $unsaved = @(Get-VisioUnsavedDocumentNames)
  if ($unsaved.Count -gt 0) {
    Write-Warning ('检测到未保存的 Visio 文档（' + ($unsaved -join ', ') + '）——为避免丢失数据，本次不结束 Visio 进程。请先保存或关闭它们。')
    return
  }
  $procs | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Start-VisioSession {
  <#
    启动（或附着）一个 Visio 实例，返回 @{ App; Doc; Page }。
    沙箱会在命令结束时回收子进程，所以同一脚本内要画完存完。

    防「画完又画一遍」：
      0) 若 Visio 里有**未保存**的文档，直接报错停下——不强杀（会丢用户改动），
         也不附着脏实例（会让新旧图同时出现在窗口里）；
      1) 否则杀掉残留 VISIO 并**等到进程真的退出**再启动新的；
      2) 附着到的实例若还开着文档（上一轮残留或用户自己的），全部关掉再新建文档——
         Visio 是单实例应用，附着到脏实例会让窗口里同时出现旧图和新图。
  #>
  param([switch]$Invisible, [double]$Width = 13.333, [double]$Height = 7.5)

  # 有未保存的文档就报错停下，不强杀：Visio 是单实例，脏实例会让新旧图同时出现在窗口里；
  # 而强杀进程会连用户未保存的改动一起丢掉——两害相权，宁可停下来报错。
  $unsaved = @(Get-VisioUnsavedDocumentNames)
  if ($unsaved.Count -gt 0) {
    throw ('Visio 中有未保存的文档（' + ($unsaved -join ', ') + '）。请先保存或关闭它们再运行本脚本。')
  }
  Stop-VisioProcess -Force
  for ($i = 0; $i -lt 10; $i++) {
    if (-not (Get-Process VISIO -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Milliseconds 300
  }
  if (Get-Process VISIO -ErrorAction SilentlyContinue) {
    Write-Warning '仍有 VISIO 进程未退出，后续可能附着到脏实例'
  }
  Start-Process -FilePath $script:VisioExe

  $app = $null
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 300
    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application'); break } catch {}
  }
  if (-not $app) { $app = New-Object -ComObject Visio.Application }
  $app.Visible = (-not $Invisible)
  try { $app.AlertResponse = 7 } catch {}   # 自动关掉模式对话框

  # 关掉所有已打开的文档：Visio 单实例，残留文档会让同一张图在窗口里出现两次
  $closed = 0
  foreach ($d in @($app.Documents)) {
    try { $d.Close(); $closed++ } catch {}
  }
  if ($closed -gt 0) { Write-Output ('[visio] 已关闭 ' + $closed + ' 个残留文档') }

  $doc = $app.Documents.AddEx('')
  $page = $doc.Pages.Item(1)
  Set-VisioCell $page.PageSheet 'PageWidth' ("{0} in" -f $Width) | Out-Null
  Set-VisioCell $page.PageSheet 'PageHeight' ("{0} in" -f $Height) | Out-Null

  return @{ App = $app; Doc = $doc; Page = $page }
}

function Open-VisioStencil {
  <# 打开模具并缓存：$stencils = Open-VisioStencil -App $app -Names @('BASIC_M','ARROWS_M') #>
  param($App, [string[]]$Names)
  $map = @{}
  foreach ($n in $Names) {
    $path = Join-Path $script:StencilRoot ($n + '.VSSX')
    if (-not (Test-Path $path)) { Write-Warning "模具不存在: $path"; continue }
    try { $map[$n] = $App.Documents.OpenEx($path, 4) } catch { Write-Warning "模具打不开: $path" }
  }
  return $map
}

function Close-VisioSession {
  <#
    关文档 + 退出 Visio。可选 -DelaySeconds：先停 N 秒再关，让用户看到成品（默认 0）。
  #>
  param($App, $Stencils, [double]$DelaySeconds = 0)
  if ($DelaySeconds -gt 0) { Start-Sleep -Seconds $DelaySeconds }
  if ($Stencils) { foreach ($k in $Stencils.Keys) { try { $Stencils[$k].Close() } catch {} } }
  try { $App.ActiveDocument.Close() } catch {}
  try { $App.Quit() } catch {}
  Start-Sleep -Seconds 2
  Stop-VisioProcess -Force
  if ($script:CellWarnings.Count -gt 0) {
    Write-Output ('[visio] 单元格写入告警: ' + (($script:CellWarnings | Select-Object -Unique) -join '; '))
  }
}

# ---------- 单元格 / 样式 ----------

function Set-VisioCell {
  <# 写单元格公式，失败只告警不中断（某些母版有 guard，写不进去是常态）。 #>
  param($Shape, [string]$Cell, [string]$Formula)
  try { $Shape.CellsU($Cell).FormulaU = $Formula; return $true }
  catch { [void]$script:CellWarnings.Add("$Cell=$Formula"); return $false }
}

function Get-VisioMaster {
  <#
    按中文名取母版。必须先试 Masters.Item()，ItemU 对中文名一律返回 null，
    而且失败是静默的——这是最容易踩的坑。
  #>
  param($Stencil, [string]$Name)
  if (-not $Stencil) { return $null }
  try { $m = $Stencil.Masters.Item($Name); if ($m) { return $m } } catch {}
  foreach ($m in $Stencil.Masters) { if ($m.Name -eq $Name) { return $m } }
  return $null
}

function Style-VisioTree {
  <# 递归上色：组形状的父级 fill 不生效，必须打到每个子形状上。 #>
  param($Shape, [string]$Fill, [string]$Line, [double]$LineWeight = 0.75, [string]$TextColor = '', [string]$Font = '微软雅黑')
  Set-VisioCell $Shape 'FillPattern' $(if ($Fill -eq '') { '0' } else { '1' }) | Out-Null
  if ($Fill -ne '') { Set-VisioCell $Shape 'FillForegnd' $Fill | Out-Null }
  if ($Line -eq '') {
    Set-VisioCell $Shape 'LinePattern' '0' | Out-Null
  } else {
    Set-VisioCell $Shape 'LinePattern' '1' | Out-Null
    Set-VisioCell $Shape 'LineColor' $Line | Out-Null
    Set-VisioCell $Shape 'LineWeight' ("{0} pt" -f $LineWeight) | Out-Null
  }
  # 字体必须无条件设：中文不指定字体会静默回退宋体。
  # Char.Font 只认字体 ID，运行时用 Get-VisioFontIndex 查，别硬编码（本机微软雅黑=245、等线=252）。
  $fidx = 245
  try { $fidx = Get-VisioFontIndex -Name $Font -Doc $Shape.ContainingPage.Document } catch { }
  if ($fidx -eq 0) { $fidx = 245 }   # 兜底：本机实测微软雅黑
  Set-VisioCell $Shape 'Char.Font' "$fidx" | Out-Null
  # Char.Font 只管西文；中文靠 Char.AsianFont，不设会静默回退宋体(5)
  Set-VisioCell $Shape 'Char.AsianFont' "$fidx" | Out-Null
  if ($TextColor -ne '') {
    Set-VisioCell $Shape 'Char.Color' $TextColor | Out-Null
  }
  if ($Shape.Shapes.Count -gt 0) { foreach ($c in $Shape.Shapes) { Style-VisioTree $c $Fill $Line $LineWeight $TextColor $Font } }
}

# ---------- 文本 / 形状 ----------

function New-VisioText {
  <# 无边框无填充的文本框。多行用 `n 连接；竖排请用 New-VisioVerticalText。 #>
  param(
    $Page, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
    [string]$Text, [double]$Size = 10, [string]$Color = 'RGB(26,26,26)',
    [switch]$Bold, [int]$HAlign = 0, [int]$VAlign = 1, [string]$Font = '微软雅黑'
  )
  $s = $Page.DrawRectangle($X1, $Y1, $X2, $Y2)
  Set-VisioCell $s 'LinePattern' '0' | Out-Null
  Set-VisioCell $s 'FillPattern' '0' | Out-Null
  Set-VisioCell $s 'LeftMargin' '0 in' | Out-Null
  Set-VisioCell $s 'RightMargin' '0 in' | Out-Null
  $s.Text = $Text
  $fidx = Get-VisioFontIndex -Name $Font -Doc $Page.Document
  if ($fidx -eq 0) { $fidx = 245 }
  Set-VisioCell $s 'Char.Font' "$fidx" | Out-Null     # 西文字体
  Set-VisioCell $s 'Char.AsianFont' "$fidx" | Out-Null # 中文字体（不设会回退宋体=5）
  Set-VisioCell $s 'Char.Size' ("{0} pt" -f $Size) | Out-Null
  Set-VisioCell $s 'Char.Color' $Color | Out-Null
  if ($Bold) { Set-VisioCell $s 'Char.Style' '1' | Out-Null }
  Set-VisioCell $s 'Para.HorzAlign' "$HAlign" | Out-Null   # 0 左 1 中 2 右
  Set-VisioCell $s 'VerticalAlign' "$VAlign" | Out-Null    # 0 上 1 中 2 下
  return $s
}

function New-VisioVerticalText {
  <#
    竖排轴名。TxtAngle 在这台机器上不可靠，改用“先建横向框、再整体旋转 270°”。
    注意：传进来的是旋转前的横向框，函数会把它绕中心转 270°。
  #>
  param($Page, [double]$CenterX, [double]$CenterY, [double]$Length, [double]$Thickness,
        [string]$Text, [double]$Size = 9.5, [string]$Color = 'RGB(90,102,110)', [string]$Font = '微软雅黑')
  $half = $Length / 2; $halfT = $Thickness / 2
  $s = $Page.DrawRectangle(($CenterX - $half), ($CenterY - $halfT), ($CenterX + $half), ($CenterY + $halfT))
  Set-VisioCell $s 'LinePattern' '0' | Out-Null
  Set-VisioCell $s 'FillPattern' '0' | Out-Null
  $s.Text = $Text
  $fidx = Get-VisioFontIndex -Name $Font -Doc $Page.Document
  if ($fidx -eq 0) { $fidx = 245 }
  Set-VisioCell $s 'Char.Font' "$fidx" | Out-Null
  Set-VisioCell $s 'Char.AsianFont' "$fidx" | Out-Null
  Set-VisioCell $s 'Char.Size' ("{0} pt" -f $Size) | Out-Null
  Set-VisioCell $s 'Char.Color' $Color | Out-Null
  Set-VisioCell $s 'Para.HorzAlign' '1' | Out-Null
  Set-VisioCell $s 'VerticalAlign' '1' | Out-Null
  Set-VisioCell $s 'Angle' '270 deg' | Out-Null
  return $s
}

function New-VisioCard {
  <#
    圆角卡片：标题 + 若干行描述。坐标是左下/右上（Visio 的 y 向上）。
    Lines 传字符串数组，每行一句。
  #>
  param(
    $Page, $Stencils, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
    [string]$Title, [string[]]$Lines,
    [string]$Fill = 'RGB(243,245,247)', [string]$Edge = 'RGB(150,163,173)',
    [string]$TitleColor = 'RGB(70,82,92)', [string]$BodyColor = 'RGB(26,26,26)',
    [double]$Radius = 0.13, [double]$TitleSize = 12, [double]$BodySize = 9.5
  )
  $card = $null
  $m = Get-VisioMaster $Stencils['BASIC_M'] '矩形'
  if ($m) {
    try {
      $card = $Page.Drop($m, (($X1 + $X2) / 2), (($Y1 + $Y2) / 2))
      Set-VisioCell $card 'Width' ("{0} in" -f ($X2 - $X1)) | Out-Null
      Set-VisioCell $card 'Height' ("{0} in" -f ($Y2 - $Y1)) | Out-Null
    } catch { $card = $null }
  }
  if (-not $card) { $card = $Page.DrawRectangle($X1, $Y1, $X2, $Y2) }
  Set-VisioCell $card 'Rounding' ("{0} in" -f $Radius) | Out-Null
  Set-VisioCell $card 'FillPattern' '1' | Out-Null
  Set-VisioCell $card 'FillForegnd' $Fill | Out-Null
  Set-VisioCell $card 'LinePattern' '1' | Out-Null
  Set-VisioCell $card 'LineColor' $Edge | Out-Null
  Set-VisioCell $card 'LineWeight' '0.75 pt' | Out-Null

  New-VisioText $Page ($X1 + 0.22) ($Y2 - 0.40) ($X2 - 0.22) ($Y2 - 0.12) $Title $TitleSize $TitleColor -Bold | Out-Null
  $body = ($Lines -join "`n")
  New-VisioText $Page ($X1 + 0.22) ($Y1 + 0.16) ($X2 - 0.22) ($Y2 - 0.50) $body $BodySize $BodyColor | Out-Null
  return $card
}

function New-VisioDot {
  <# 实心/空心圆点（状态指示）。 #>
  param($Page, $Stencils, [double]$CenterX, [double]$CenterY, [double]$Diameter = 0.13,
        [switch]$On, [string]$Ring = 'RGB(0,114,178)', [string]$OnFill = 'RGB(0,114,178)', [string]$OffFill = 'RGB(255,255,255)')
  $s = $null
  $m = Get-VisioMaster $Stencils['BASIC_M'] '圆形'
  if ($m) {
    try { $s = $Page.Drop($m, $CenterX, $CenterY); Set-VisioCell $s 'Width' ("{0} in" -f $Diameter) | Out-Null; Set-VisioCell $s 'Height' ("{0} in" -f $Diameter) | Out-Null } catch { $s = $null }
  }
  if (-not $s) { $s = $Page.DrawOval(($CenterX - $Diameter / 2), ($CenterY - $Diameter / 2), ($CenterX + $Diameter / 2), ($CenterY + $Diameter / 2)) }
  Set-VisioCell $s 'LineColor' $Ring | Out-Null
  Set-VisioCell $s 'LineWeight' '1 pt' | Out-Null
  Set-VisioCell $s 'FillPattern' '1' | Out-Null
  Set-VisioCell $s 'FillForegnd' $(if ($On) { $OnFill } else { $OffFill }) | Out-Null
  return $s
}

function New-VisioDiamond {
  <#
    菱形标记（“目标/最优”之类）。
    默认白填充 + 灰色描边 = 空心菱形；**必须留描边**，否则白填充落在白底上等于没画。
  #>
  param($Page, $Stencils, [double]$CenterX, [double]$CenterY, [double]$Size = 0.16,
        [string]$Fill = 'RGB(255,255,255)', [string]$Line = 'RGB(107,120,130)')
  $s = $null
  $m = Get-VisioMaster $Stencils['BASIC_M'] '菱形'
  if ($m) {
    try { $s = $Page.Drop($m, $CenterX, $CenterY); Set-VisioCell $s 'Width' ("{0} in" -f $Size) | Out-Null; Set-VisioCell $s 'Height' ("{0} in" -f $Size) | Out-Null } catch { $s = $null }
  }
  if (-not $s) { $s = $Page.DrawRectangle(($CenterX - $Size / 2), ($CenterY - $Size / 2), ($CenterX + $Size / 2), ($CenterY + $Size / 2)) }
  Set-VisioCell $s 'FillPattern' '1' | Out-Null
  Set-VisioCell $s 'FillForegnd' $Fill | Out-Null
  if ($Line -eq '') { Set-VisioCell $s 'LinePattern' '0' | Out-Null }
  else { Set-VisioCell $s 'LinePattern' '1' | Out-Null; Set-VisioCell $s 'LineColor' $Line | Out-Null; Set-VisioCell $s 'LineWeight' '1 pt' | Out-Null }
  return $s
}

function New-VisioArrow {
  <#
    单向箭头 = 线 + 实心箭头端。
    不要用 ARROWS_M「普通箭头」母版：它的方向被 guard 锁死，Angle/FlipX 都不生效。
  #>
  param($Page, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
        [string]$Color = 'RGB(107,120,130)', [double]$Weight = 1.5)
  $a = $Page.DrawLine($X1, $Y1, $X2, $Y2)
  Set-VisioCell $a 'LineColor' $Color | Out-Null
  Set-VisioCell $a 'LineWeight' ("{0} pt" -f $Weight) | Out-Null
  Set-VisioCell $a 'EndArrow' '4' | Out-Null
  Set-VisioCell $a 'EndArrowSize' '2' | Out-Null
  return $a
}

function New-VisioAxisArrow {
  <# 双向轴：优先用 ARROWS_M「普通双箭头」母版（对称，不受方向锁影响）。 #>
  param($Page, $Stencils, [double]$CenterX, [double]$CenterY, [double]$Length,
        [switch]$Vertical, [string]$Color = 'RGB(110,122,130)')
  $a = $null
  $m = Get-VisioMaster $Stencils['ARROWS_M'] '普通双箭头'
  if ($m) {
    try {
      $a = $Page.Drop($m, $CenterX, $CenterY)
      Set-VisioCell $a 'Width' ("{0} in" -f $Length) | Out-Null
      Set-VisioCell $a 'Height' '0.16 in' | Out-Null
      if ($Vertical) { Set-VisioCell $a 'Angle' '90 deg' | Out-Null }
      Style-VisioTree $a 'RGB(255,255,255)' $Color 0.75 ''
    } catch { $a = $null }
  }
  if (-not $a) {
    if ($Vertical) { $a = $Page.DrawLine($CenterX, ($CenterY - $Length / 2), $CenterX, ($CenterY + $Length / 2)) }
    else { $a = $Page.DrawLine(($CenterX - $Length / 2), $CenterY, ($CenterX + $Length / 2), $CenterY) }
    Set-VisioCell $a 'LineColor' $Color | Out-Null
    Set-VisioCell $a 'LineWeight' '0.75 pt' | Out-Null
    Set-VisioCell $a 'BeginArrow' '4' | Out-Null
    Set-VisioCell $a 'EndArrow' '4' | Out-Null
  }
  return $a
}

function Add-VisioPageBackground {
  <#
    必须先铺一张满页白底：Page.Export 只裁“图形外接框”，
    没有满页底图时导出会切掉四周留白，尺寸不可控。
  #>
  param($Page, [string]$Color = 'RGB(255,255,255)')
  $w = $Page.PageSheet.CellsU('PageWidth').ResultIU
  $h = $Page.PageSheet.CellsU('PageHeight').ResultIU
  $bg = $Page.DrawRectangle(0, 0, $w, $h)
  Set-VisioCell $bg 'FillPattern' '1' | Out-Null
  Set-VisioCell $bg 'FillForegnd' $Color | Out-Null
  Set-VisioCell $bg 'LinePattern' '0' | Out-Null
  try { $bg.SendToBack() } catch {}
  return $bg
}

# ---------- 交付 ----------

function Save-VisioDoc {
  <#
    保存并导出。本机 .vsdx 的 SaveAs 会抛空异常（可见/不可见实例都失败），
    所以要 try .vsdx → fallback .vsd。导出 PNG 分辨率跟随页面尺寸。
    数据安全：SaveAs 会把目标文件锁住（存临时文件再 Move 会失败），
    所以改为"先把旧文件改名成 .bak → 直接存 → 失败就恢复 .bak"。
  #>
  param($Doc, $Page, [string]$VsdPath, [string]$PngPath)
  $saved = ''
  if ($VsdPath) {
    $bak = ''
    if (Test-Path $VsdPath) {
      $bak = "$VsdPath.bak"
      if (Test-Path $bak) { Remove-Item $bak -Force }
      Move-Item $VsdPath $bak -Force
    }
    try { $null = $Doc.SaveAs($VsdPath); $saved = $VsdPath }
    catch {
      # 不许静默：告诉调用方首次存盘失败的真实原因，再回落 .vsd
      Write-Warning ('保存到 ' + [System.IO.Path]::GetExtension($VsdPath) + ' 失败: ' + $_.Exception.Message.Trim() + ' -> 回落 .vsd')
      $fallback = [System.IO.Path]::ChangeExtension($VsdPath, '.vsd')
      try { $null = $Doc.SaveAs($fallback); $saved = $fallback }
      catch {
        if ($bak -and (Test-Path $bak)) { Move-Item $bak $VsdPath -Force }
        Write-Warning ('保存失败: ' + $_.Exception.Message.Trim())
      }
    }
    if ($bak -and (Test-Path $bak)) { Remove-Item $bak -Force }
  }
  if ($PngPath) {
    # Export 会返回一个值，必须吃掉，否则会混进本函数的返回值
    try { $null = $Page.Export($PngPath) } catch { Write-Warning ('导出失败: ' + $_.Exception.Message.Trim()) }
  }
  return $saved
}

function Get-VisioShapeReport {
  <# 统计当前页形状构成，用来确认“不是一个模具打天下”。 #>
  param($Page, $Stencils)
  $names = @{}
  foreach ($k in $Stencils.Keys) {
    try { foreach ($m in $Stencils[$k].Masters) { $names[$m.Name] = $k } } catch {}
  }
  $rows = @()
  foreach ($s in @($Page.Shapes)) {
    $t = ''
    try { $t = $s.Master.Name } catch {}
    $fam = if ($t -and $names.ContainsKey($t)) { $names[$t] } else { '（原生图元）' }
    $rows += [pscustomobject]@{ Shape = $s.Name; Master = $t; Family = $fam }
  }
  return $rows
}

# ---------- 期刊成品尺寸（见 references/journal-figure-standards.md） ----------

function Set-VisioShapeScale {
  <#
    单个形状的等比缩放：几何（PinX/PinY/Width/Height/Rounding）按比例，
    线宽按比例并设下限，字号按比例并设下限。
    组形状的子形状几何随父级自动缩放，所以子形状只处理线宽与文字（-TextOnly）。
  #>
  param($Shape, [double]$Factor, [double]$MinFontPt = 0, [double]$MinLinePt = 0, [switch]$TextOnly)
  if (-not $TextOnly) {
    foreach ($cell in @('PinX', 'PinY', 'Width', 'Height')) {
      try {
        $v = $Shape.CellsU($cell).ResultIU
        if ($v -ne 0) { Set-VisioCell $Shape $cell ("{0} in" -f ($v * $Factor)) | Out-Null }
      } catch {}
    }
    try {
      $r = $Shape.CellsU('Rounding').ResultIU
      if ($r -gt 0) { Set-VisioCell $Shape 'Rounding' ("{0} in" -f ($r * $Factor)) | Out-Null }
    } catch {}
  }
  try {
    # 注意：Char.Size 与 LineWeight 的 ResultIU 单位是英寸（1 pt = 1/72 in），不是磅
    $wIn = $Shape.CellsU('LineWeight').ResultIU
    if ($wIn -gt 0) {
      $nw = $wIn * 72 * $Factor
      if ($MinLinePt -gt 0 -and $nw -lt $MinLinePt) { $nw = $MinLinePt }
      Set-VisioCell $Shape 'LineWeight' ("{0} pt" -f $nw) | Out-Null
    }
  } catch {}
  try {
    $fsIn = $Shape.CellsU('Char.Size').ResultIU
    if ($fsIn -gt 0) {
      $nf = [math]::Round($fsIn * 72 * $Factor, 1)
      if ($MinFontPt -gt 0 -and $nf -lt $MinFontPt) { $nf = $MinFontPt }
      Set-VisioCell $Shape 'Char.Size' ("{0} pt" -f $nf) | Out-Null
    }
  } catch {}
  if ($Shape.Shapes.Count -gt 0) {
    foreach ($c in @($Shape.Shapes)) { Set-VisioShapeScale $c $Factor $MinFontPt $MinLinePt -TextOnly }
  }
}

function Set-VisioPageScale {
  <#
    整页等比缩放：把 13.333 in 母版变成期刊成品尺寸（单栏 89 / 1.5 栏 120 / 双栏 183 mm 宽）。
      -Factor     缩放系数（宽度比例）
      -WidthIn    缩放后页宽（英寸）
      -HeightIn   缩放后页高（英寸）
      -MinFontPt  字号下限（Nature 最小 5 pt）
      -MinLinePt  线宽下限（Elsevier 最小 0.5 pt）
    缩小后几何、线宽、字号同步缩小，字号与线宽由下限兜底。

    ⚠ 它会**就地改动页面**：整页重新排版重绘，Visio 窗口里会看到一次缩放重绘，
      而且文档的页面尺寸被永久改掉。
      所以它只该在**派生的、不可见的会话**里用（见 New-VisioJournalDeliverables）；
      主会话里不要调用，否则会出现"保存时又画了一遍"的观感。
  #>
  param($Page, [double]$Factor, [double]$WidthIn, [double]$HeightIn,
        [double]$MinFontPt = 0, [double]$MinLinePt = 0)
  Set-VisioCell $Page.PageSheet 'PageWidth' ("{0} in" -f $WidthIn) | Out-Null
  Set-VisioCell $Page.PageSheet 'PageHeight' ("{0} in" -f $HeightIn) | Out-Null
  foreach ($s in @($Page.Shapes)) { Set-VisioShapeScale $s $Factor $MinFontPt $MinLinePt }
}

function Export-VisioPageRaster {
  <#
    高分辨率位图导出。Visio 的 Page.Export 跟随默认打印机（本机实测 144 dpi），
    想拿到更高 dpi，就把整页放大 $Zoom 倍再导出——几何、线宽、字号同步放大，
    等价于 $Zoom 倍的矢量渲染。$Zoom 必须是整数（还原时 1/$Zoom 无误差）。

    注意：它会**就地改动页面**，Visio 窗口里会看到一次放大重绘。
    如果当前文档已经保存过，改用「另开一个不可见会话打开已保存文件 → 放大 → 导出 → 不保存」，
    这样主文档和用户窗口都不受影响。
  #>
  param($Page, [string]$Path, [int]$Zoom = 4)
  $w = $Page.PageSheet.CellsU('PageWidth').ResultIU
  $h = $Page.PageSheet.CellsU('PageHeight').ResultIU
  Set-VisioPageScale -Page $Page -Factor $Zoom -WidthIn ($w * $Zoom) -HeightIn ($h * $Zoom) | Out-Null
  if (Test-Path $Path) { Remove-Item $Path -Force }
  try { $null = $Page.Export($Path) } catch { Write-Warning ('导出失败: ' + $_.Exception.Message.Trim()) }
  Set-VisioPageScale -Page $Page -Factor (1.0 / $Zoom) -WidthIn $w -HeightIn $h | Out-Null
  return $Path
}

function Save-VisioPdf {
  <#
    矢量 PDF（期刊首选交付格式，RGB、字体嵌入、文字可编辑）。
    Visio 的 SaveAs 到 .pdf 会抛「参数无效」，必须走 ExportAsFixedFormat，
    且要传满 11 个参数，否则 COM 报「找不到重载」。
      参数：FixedFormat(1=PDF), OutputFileName, Intent, FromPage, ToPage,
            ColorAsBlack, IncludeBackground, IncludeDocumentProperties,
            IncludeStructureTags, UseISO19005_1, FixedFormatExtClassPtr
  #>
  param($Doc, [string]$Path)
  if (Test-Path $Path) { Remove-Item $Path -Force }
  $miss = [System.Reflection.Missing]::Value
  try {
    $Doc.ExportAsFixedFormat(1, $Path, 1, 0, 0, $false, $true, $true, $true, $false, $miss) | Out-Null
    if (Test-Path $Path) { return $Path }
  } catch {
    Write-Warning ('PDF 导出失败(ExportAsFixedFormat): ' + $_.Exception.Message.Trim())
  }
  try { $null = $Doc.SaveAs($Path); if (Test-Path $Path) { return $Path } }
  catch { Write-Warning ('PDF 保存失败(SaveAs): ' + $_.Exception.Message.Trim()) }
  return ''
}

$script:VisioFontCache = @{}

function Get-VisioFontIndex {
  <#
    查文档字体表里某个字体的 ID（Char.Font 只接受字体 ID，不能直接写字体名）。
      -Doc   文档对象（默认取当前活动文档）
      -Name  字体名，如 'Times New Roman' / 'Cambria' / '微软雅黑'
    常见索引（Microsoft 365 本机实测）：微软雅黑 245、Times New Roman 23、Cambria 47、
      Cambria Math 48、Calibri 4、Arial 21、Segoe UI 183、Consolas 65、Georgia 91。
    ⚠ 字体 ID 随文档字体表变化（跨机器/跨版本可能不同），**别硬编码**，运行时查（带会话级缓存）。
    用途：中文一律用 245（微软雅黑）；公式/纯拉丁文字用 Times New Roman 斜体才像数学公式
         （微软雅黑里的 '·' 前后自带空隙，'α·Sim' 会排成 'α ·Sim'，很难看）。
  #>
  param([string]$Name, $Doc = $null)
  if (-not $Doc) {
    try {
      $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
      $Doc = $app.ActiveDocument
    } catch { }
  }
  if (-not $Doc) { throw '请传入 -Doc 文档对象（Char.Font 只接受字体 ID，必须运行时查）' }
  $key = $Name.ToLowerInvariant()
  if ($script:VisioFontCache.ContainsKey($key)) { return $script:VisioFontCache[$key] }
  for ($i = 1; $i -le $Doc.Fonts.Count; $i++) {
    try { if ($Doc.Fonts.Item($i).Name -eq $Name) { $script:VisioFontCache[$key] = $i; return $i } } catch { }
  }
  return 0
}

# ---------- 期刊件派生（从已保存的母版源文件生成，不碰主文档） ----------

function New-VisioJournalDeliverables {
  <#
    从**已保存的母版源文件**派生期刊成品：
      母版 .vsdx（13.333 in） → 缩到成品尺寸 → 期刊版 .vsdx + 矢量 PDF + 预览 PNG（+ 4× 高分辨率 PNG）

    为什么要单独一个会话：
      在主会话里就地缩放页面会触发整页重新排版重绘——用户在窗口里看到的就是
      "保存的时候又画了一遍"，而且主文档会被污染成缩放后的尺寸。
      这里改为：主会话只画一次、只存一次；派生件全部在**不可见会话**里从已存文件生成。

    返回 @{ Vsd; Pdf; Preview = @(); HiRes = @() }
  #>
  param(
    [Parameter(Mandatory = $true)][string]$SourceVsd,   # 母版源文件（已保存）
    [Parameter(Mandatory = $true)][string]$OutDir,
    [Parameter(Mandatory = $true)][string]$Prefix,        # 输出前缀，如 fig-table3-2x2
    [Parameter(Mandatory = $true)][double]$WidthIn,       # 成品宽（英寸，183 mm = 7.2047）
    [Parameter(Mandatory = $true)][double]$HeightIn,      # 成品高（英寸）
    [Parameter(Mandatory = $true)][double]$Factor,        # 缩放系数 = WidthIn / 母版宽
    [double]$MinFontPt = 5.0,
    [double]$MinLinePt = 0.5,
    [switch]$HighRes,
    [int]$HighResZoom = 4
  )
  if (-not (Test-Path $SourceVsd)) { throw ('源文件不存在: ' + $SourceVsd) }
  New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

  $app = New-Object -ComObject Visio.Application
  $app.Visible = $false
  try { $app.AlertResponse = 7 } catch {}
  $doc = $null
  $preview = @(); $hi = @(); $jVsd = ''; $pdf = ''
  try {
    $doc = $app.Documents.Open($SourceVsd)
    $pages = @($doc.Pages)
    for ($i = 0; $i -lt $pages.Count; $i++) {
      $pg = $pages[$i]
      $tag = if ($pages.Count -gt 1) { '-p' + ($i + 1) } else { '' }
      # 1) 缩到成品尺寸
      Set-VisioPageScale -Page $pg -Factor $Factor -WidthIn $WidthIn -HeightIn $HeightIn -MinFontPt $MinFontPt -MinLinePt $MinLinePt | Out-Null
      # 2) 成品尺寸预览图
      $prev = Join-Path $OutDir ($Prefix + $tag + '-preview.png')
      if (Test-Path $prev) { Remove-Item $prev -Force }
      try { $null = $pg.Export($prev); $preview += $prev } catch { Write-Warning ('预览导出失败: ' + $_.Exception.Message.Trim()) }
      # 3) 高分辨率位图（再放大 HighResZoom 倍 ≈ 144×zoom dpi）
      if ($HighRes) {
        $hp = Join-Path $OutDir ($Prefix + $tag + '-' + ($HighResZoom * 144) + 'dpi.png')
        Set-VisioPageScale -Page $pg -Factor $HighResZoom -WidthIn ($WidthIn * $HighResZoom) -HeightIn ($HeightIn * $HighResZoom) -MinFontPt $MinFontPt -MinLinePt $MinLinePt | Out-Null
        if (Test-Path $hp) { Remove-Item $hp -Force }
        try { $null = $pg.Export($hp); $hi += $hp } catch { Write-Warning ('高分辨率导出失败: ' + $_.Exception.Message.Trim()) }
        # 还原到成品尺寸（字号/线宽下限要一起带，否则还原后可能低于 5 pt / 0.5 pt）
        Set-VisioPageScale -Page $pg -Factor (1.0 / $HighResZoom) -WidthIn $WidthIn -HeightIn $HeightIn -MinFontPt $MinFontPt -MinLinePt $MinLinePt | Out-Null
      }
    }
    # 4) 另存期刊版源文件（成品尺寸）
    $jVsd = Join-Path $OutDir ($Prefix + '.vsdx')
    if (Test-Path $jVsd) { Remove-Item $jVsd -Force }
    try { $null = $doc.SaveAs($jVsd) }
    catch {
      $fb = [System.IO.Path]::ChangeExtension($jVsd, '.vsd')
      try { $null = $doc.SaveAs($fb); $jVsd = $fb } catch { Write-Warning ('期刊版保存失败: ' + $_.Exception.Message.Trim()); $jVsd = '' }
    }
    # 5) 矢量 PDF（期刊首选交付格式）
    $pdf = Save-VisioPdf -Doc $doc -Path (Join-Path $OutDir ($Prefix + '.pdf'))
  } finally {
    try { $doc.Saved = $true; $doc.Close() } catch {}
    try { $app.Quit() } catch {}
    Stop-VisioProcess -Force
  }
  return @{ Vsd = $jVsd; Pdf = $pdf; Preview = $preview; HiRes = $hi }
}

# ---------- 模具库检索（绘图前查形状用） ----------

function Get-VisioSkillRoot {
  # dot-source 方式加载时 $PSScriptRoot 为空，需要兜底
  if ($PSScriptRoot) { return (Split-Path -Parent $PSScriptRoot) }
  return 'C:\Users\thunder\.agents\skills\visio-diagram'
}

function Find-VisioMaster {
  <#
    按关键词检索本机所有模具库里的 master（中文名或英文名 NameU）。
    优先读缓存清单 references\data\visio-masters-full.tsv（毫秒级，由 visio-inventory.ps1 生成）；
    没有缓存就直接扫模具目录（慢，约 1~3 分钟）。
    返回 @{ Stencil; Name; NameU } 数组，按 模具库+名称 排序。
  #>
  param(
    [string[]]$Pattern,
    [string]$CachePath = '',
    [string]$Root = 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052',
    [int]$Limit = 80
  )
  if (-not $Pattern -or $Pattern.Count -eq 0) { throw '请至少给一个关键词' }
  if (-not $CachePath) { $CachePath = Join-Path (Get-VisioSkillRoot) 'references\data\visio-masters-full.tsv' }
  $hits = New-Object System.Collections.Generic.List[object]

  if (Test-Path $CachePath) {
    foreach ($line in [System.IO.File]::ReadAllLines($CachePath, [System.Text.Encoding]::UTF8)) {
      if (-not $line -or $line.StartsWith('#')) { continue }
      $parts = $line -split "`t"
      if ($parts.Count -lt 2) { continue }
      $u = if ($parts.Count -gt 2) { $parts[2] } else { '' }
      foreach ($p in $Pattern) {
        if ($parts[1] -like "*$p*" -or $u -like "*$p*") {
          $hits.Add([pscustomobject]@{ Stencil = $parts[0]; Name = $parts[1]; NameU = $u })
          break
        }
      }
    }
    return ($hits | Sort-Object Stencil, Name | Select-Object -First $Limit)
  }

  $app = New-Object -ComObject Visio.Application
  $app.Visible = $false
  try { $app.AlertResponse = 7 } catch {}
  foreach ($f in (Get-ChildItem $Root -Filter *.vssx)) {
    $d = $null
    try {
      $d = $app.Documents.Open($f.FullName)
      foreach ($m in $d.Masters) {
        foreach ($p in $Pattern) {
          if ($m.Name -like "*$p*" -or $m.NameU -like "*$p*") {
            $hits.Add([pscustomobject]@{ Stencil = $f.BaseName; Name = $m.Name; NameU = $m.NameU })
            break
          }
        }
      }
    } catch { } finally { if ($d) { try { $d.Saved = $true; $d.Close() } catch {} } }
  }
  try { $app.Quit() } catch {}
  Stop-VisioProcess -Force
  return ($hits | Sort-Object Stencil, Name | Select-Object -First $Limit)
}

function Get-VisioMasterByName {
  <#
    一次性打开某个模具库并按**中文名**取 master 对象，可直接 $Page.Drop($m, x, y)。
    注意：取到的 master 依赖模具文档保持打开，绘图期间不要 Close。
    中文名一律用 Masters.Item("中文名")；ItemU 只认英文通用名，中文会返回 null。
  #>
  param(
    [Parameter(Mandatory = $true)][string]$Stencil,
    [Parameter(Mandatory = $true)][string]$Name,
    [string]$Root = 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052'
  )
  $file = Join-Path $Root ($Stencil + '.vssx')
  if (-not (Test-Path $file)) { throw "找不到模具库: $file" }
  $app = New-Object -ComObject Visio.Application
  $app.Visible = $false
  try { $app.AlertResponse = 7 } catch {}
  $doc = $app.Documents.Open($file)
  $m = $doc.Masters.Item($Name)
  if (-not $m) { $doc.Close(); $app.Quit(); throw "模具库 $Stencil 里没有名为「$Name」的 master" }
  return [pscustomobject]@{ Master = $m; Doc = $doc; App = $app }
}

# ---------- 从 visio-mcp 吸收的能力（自动布局 / 胶合连接线 / 结构化页检查 / 容器）----------

function Invoke-VisioAutoLayout {
  <#
    Visio 原生自动布局。
      -Style  flowchart_tb | flowchart_lr | tree_tb | tree_lr | radial | circular
      -SpacingIn 同层间距（英寸）
      -ResizePage 布局后把页面缩到内容大小
  #>
  param($Page, [string]$Style = 'flowchart_tb', [double]$SpacingIn = 0.75, [switch]$ResizePage)
  $map = @{
    'flowchart_tb' = @(1, 5); 'flowchart_lr' = @(2, 6)
    'tree_tb'      = @(1, 7); 'tree_lr'      = @(2, 8)
    'radial'       = @(3, 9); 'circular'     = @(6, 16)
  }
  if (-not $map.ContainsKey($Style)) { throw ('Style 必须是: ' + (($map.Keys | Sort-Object) -join ', ')) }
  $sheet = $Page.PageSheet
  $sheet.CellsU('PlaceStyle').ResultIU = $map[$Style][0]
  $sheet.CellsU('RouteStyle').ResultIU = $map[$Style][1]
  $sheet.CellsU('AvenueSizeX').ResultIU = $SpacingIn
  $sheet.CellsU('AvenueSizeY').ResultIU = $SpacingIn
  $Page.Layout()
  if ($ResizePage) { $Page.ResizeToFitContents() }
  return @{ style = $Style; page = [string]$Page.Name }
}

function New-VisioConnector {
  <#
    胶合动态连接线：两端 GlueTo PinX（动态胶合），随形状移动自动重路由。
      -Route right_angle | straight | curved
      -LinePattern solid | dashed | dotted | dash_dot
  #>
  param(
    $App, $Page, $FromShape, $ToShape, [string]$Label = '',
    [string]$Route = 'right_angle', [bool]$EndArrow = $true, [bool]$BeginArrow = $false,
    [string]$LinePattern = 'solid', [double]$LineWeightPt = 0, [string]$LineColor = ''
  )
  $routeMap = @{ 'right_angle' = @(1, 1); 'straight' = @(16, 1); 'curved' = @(1, 2) }
  if (-not $routeMap.ContainsKey($Route)) { throw ('Route 必须是: ' + (($routeMap.Keys | Sort-Object) -join ', ')) }
  $lpMap = @{ 'solid' = 1; 'dashed' = 2; 'dotted' = 3; 'dash_dot' = 4 }
  $conn = $Page.Drop($App.ConnectorToolDataObject, 0.0, 0.0)
  $conn.CellsU('BeginX').GlueTo($FromShape.CellsU('PinX'))
  $conn.CellsU('EndX').GlueTo($ToShape.CellsU('PinX'))
  $conn.CellsU('ShapeRouteStyle').ResultIU = $routeMap[$Route][0]
  $conn.CellsU('ConLineRouteExt').ResultIU   = $routeMap[$Route][1]
  $conn.CellsU('EndArrow').FormulaU   = $(if ($EndArrow) { '5' } else { '0' })
  $conn.CellsU('BeginArrow').FormulaU = $(if ($BeginArrow) { '5' } else { '0' })
  if ($lpMap.ContainsKey($LinePattern)) { $conn.CellsU('LinePattern').ResultIU = $lpMap[$LinePattern] }
  if ($LineWeightPt -gt 0) { $conn.CellsU('LineWeight').FormulaU = ('{0} pt' -f $LineWeightPt) }
  if ($LineColor -ne '') { $conn.CellsU('LineColor').FormulaU = $LineColor }
  if ($Label -ne '') { $conn.Text = $Label }
  return $conn
}

function Get-VisioPageState {
  <# 结构化读取页面所有形状：id / 坐标 / 尺寸 / 文字 / 是否连接线。 #>
  param($Page)
  $rows = @()
  foreach ($sh in @($Page.Shapes)) {
    $r = [ordered]@{}
    try { $r.shape_id = [int]$sh.ID } catch { $r.shape_id = -1 }
    try { $r.name = [string]$sh.NameU } catch { $r.name = '' }
    try { $r.x = [math]::Round([double]$sh.CellsU('PinX').ResultIU, 4) } catch { $r.x = 0 }
    try { $r.y = [math]::Round([double]$sh.CellsU('PinY').ResultIU, 4) } catch { $r.y = 0 }
    try { $r.width_in = [math]::Round([double]$sh.CellsU('Width').ResultIU, 4) } catch { $r.width_in = 0 }
    try { $r.height_in = [math]::Round([double]$sh.CellsU('Height').ResultIU, 4) } catch { $r.height_in = 0 }
    try { $r.text = [string]$sh.Text } catch { $r.text = '' }
    try { $r.is_connector = [bool]([int]$sh.OneD) } catch { $r.is_connector = $false }
    $rows += [pscustomobject]$r
  }
  return $rows
}

function New-VisioContainer {
  <#
    用 Visio 内建容器母版包住一组形状（成员随容器移动）。
      -Members 形状数组；-MasterName 留空用第 1 个容器母版
  #>
  param($App, $Page, [string]$Label, [int[]]$MemberIds, [string]$MasterName = '', [double]$PaddingIn = 0.4)
  if (-not $MemberIds -or $MemberIds.Count -eq 0) { throw 'MemberIds 不能为空' }
  # 按 ID 重新取形状：Page.Layout() 之后旧的 COM 代理会失效
  $mem = @()
  foreach ($id in $MemberIds) { $mem += $Page.Shapes.ItemFromID([int]$id) }
  $path = $App.GetBuiltInStencilFile(2, 1)   # visStencilTypeContainers=2, visMSUS=1
  $base = [System.IO.Path]::GetFileName($path)
  $stencil = $null
  foreach ($d in @($App.Documents)) { try { if ([string]$d.Name -eq $base) { $stencil = $d } } catch { } }
  if (-not $stencil) { $stencil = $App.Documents.OpenEx($path, 6) }   # visOpenRO=2 | visOpenDocked=4
  $master = if ($MasterName -ne '') { $stencil.Masters.Item($MasterName) } else { $stencil.Masters.Item(1) }
  $lefts = @(); $rights = @(); $bottoms = @(); $tops = @()
  foreach ($m in $mem) {
    $l = [double]$m.CellsU('PinX').ResultIU - [double]$m.CellsU('LocPinX').ResultIU
    $b = [double]$m.CellsU('PinY').ResultIU - [double]$m.CellsU('LocPinY').ResultIU
    $lefts += $l; $rights += ($l + [double]$m.CellsU('Width').ResultIU)
    $bottoms += $b; $tops += ($b + [double]$m.CellsU('Height').ResultIU)
  }
  $minX = ($lefts | Measure-Object -Minimum).Minimum
  $maxX = ($rights | Measure-Object -Maximum).Maximum
  $minY = ($bottoms | Measure-Object -Minimum).Minimum
  $maxY = ($tops | Measure-Object -Maximum).Maximum
  $container = $Page.Drop($master, (($minX + $maxX) / 2), (($minY + $maxY) / 2))
  $container.CellsU('Width').ResultIU  = ($maxX - $minX) + 2 * $PaddingIn
  $container.CellsU('Height').ResultIU = ($maxY - $minY) + 2 * $PaddingIn
  $container.Text = $Label
  try {
    $props = $container.ContainerProperties
    foreach ($m in $mem) { $props.AddMember($m, 1) }   # visMemberAddExpandContainer=1
  } catch { Write-Warning '容器成员添加失败（该母版可能不是 Visio 容器）' }
  return $container
}
