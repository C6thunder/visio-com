# selftest.ps1 — 纯逻辑自检，**不需要 Visio**
#
# 覆盖：
#   1) 字体 ID 回退逻辑（Resolve-VisioFontIndex）——目标字体缺失时回退微软雅黑，
#      两者都缺时返回 0（调用方不写 Char.Font），不能静默用错字体 ID。
#   2) 字体 ID 缓存按**文档**隔离——字体 ID 是文档级的，两份文档的字体表顺序不同，
#      不能跨文档复用（回归测试：旧实现会把 A 文档的 ID 用到 B 文档上）。
#   3) Save-VisioDoc 的备份语义——写 .vsdx 失败回落 .vsd 时，**旧 .vsdx 必须还在**
#      （回归测试：旧实现在回退分支把旧 .vsdx 删了，只剩 .vsd）。
#   4) 单元格告警只打印一次（打印后清空），异常路径也能取到。
#
# 用法：powershell -File scripts\selftest.ps1      （退出码 0 = 全过）
param()

$ErrorActionPreference = 'Stop'
$skillRoot = if ($env:VISIO_SKILL_ROOT) { $env:VISIO_SKILL_ROOT }
             elseif ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot }
             else { Join-Path $env:USERPROFILE '.agents\skills\visio-diagram' }
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 (Join-Path $skillRoot 'scripts\visio-lib.ps1'))))

$script:failCount = 0
function Assert-Eq($actual, $expected, [string]$label) {
  if ("$actual" -eq "$expected") { Write-Host ('[OK]   ' + $label + ' = ' + $actual) }
  else { Write-Host ('[FAIL] ' + $label + '：期望 ' + $expected + '，实际 ' + $actual); $script:failCount++ }
}

# ---------- 假对象：不碰 Visio 也能测 COM 调用点 ----------
function New-FakeFontTable {
  # 模拟 Document.Fonts：Visio 是 1-based
  param([string[]]$Names)
  $list = @($Names)
  $t = New-Object PSObject
  $t | Add-Member -MemberType NoteProperty -Name Count -Value $list.Count
  $t | Add-Member -MemberType NoteProperty -Name _list -Value $list
  # Item 必须返回带 .Name 的对象（真 Visio 的 Fonts.Item() 返回 Font 对象，不是字符串）
  $t | Add-Member -MemberType ScriptMethod -Name Item -Value { param($i) return [pscustomobject]@{ Name = $this._list[$i - 1] } }
  return $t
}

function New-FakeDoc {
  param([string]$Id, [string[]]$FontNames)
  return [pscustomobject]@{ ID = $Id; Name = ($Id + '.vsdx'); Fonts = (New-FakeFontTable $FontNames) }
}

# ---------- 2) 字体缓存按文档隔离（用库里的真实现） ----------
$docA = New-FakeDoc -Id 'docA' -FontNames @('宋体', '微软雅黑')
$docB = New-FakeDoc -Id 'docB' -FontNames @('微软雅黑', '宋体')
Assert-Eq (Get-VisioFontIndex -Name '微软雅黑' -Doc $docA) 2 'A 文档里微软雅黑的索引'
Assert-Eq (Get-VisioFontIndex -Name '微软雅黑' -Doc $docB) 1 'B 文档里微软雅黑的索引（不能复用 A 的缓存）'

# ---------- 1) 字体回退：把字体查询换成假的 ----------
$script:FakeFonts = @{ '微软雅黑' = 245; 'Times New Roman' = 23 }
function Get-VisioFontIndex {
  param([string]$Name, $Doc = $null)
  if ($script:FakeFonts.ContainsKey($Name)) { return $script:FakeFonts[$Name] }
  return 0
}

Assert-Eq (Resolve-VisioFontIndex -Name 'Times New Roman' -Doc $null) 23 '存在的字体直接返回索引'
Assert-Eq (Resolve-VisioFontIndex -Name '不存在的字体' -Doc $null) 245 '缺失时回退微软雅黑'
Assert-Eq (Resolve-VisioFontIndex -Name '微软雅黑' -Doc $null) 245 '微软雅黑本身'

$script:FakeFonts = @{}   # 这台机器一个字体都没有
Assert-Eq (Resolve-VisioFontIndex -Name '任意字体' -Doc $null) 0 '两个都缺时返回 0（不写 Char.Font）'

# ---------- 3) Save-VisioDoc 的备份语义 ----------
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('visio-selftest-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
try {
  # 假文档：.vsdx 写不进去（本机常态），.vsd 能写
  $docVsdxFail = New-Object PSObject
  $docVsdxFail | Add-Member -MemberType ScriptMethod -Name SaveAs -Value {
    param($p)
    if ([IO.Path]::GetExtension($p) -eq '.vsdx') { throw 'mock: 本机 .vsdx 不可写' }
    Set-Content -LiteralPath $p -Value 'NEW-CONTENT' -Encoding UTF8
  }
  $x = Join-Path $tmp 'x.vsdx'
  Set-Content -LiteralPath $x -Value 'OLD-CONTENT' -Encoding UTF8
  $out = Save-VisioDoc -Doc $docVsdxFail -Page $null -VsdPath $x -PngPath $null
  Assert-Eq (Test-Path $x) 'True' '回落 .vsd 时旧 .vsdx 仍然存在'
  # 文件不存在时不要直接抛 FileNotFound（EAP=Stop 会中断自检），改成报告缺失
  $oldContent = if (Test-Path $x) { (Get-Content -Raw -LiteralPath $x).Trim() } else { '<文件已丢失>' }
  Assert-Eq $oldContent 'OLD-CONTENT' '旧 .vsdx 内容没被覆盖'
  Assert-Eq (Test-Path (Join-Path $tmp 'x.vsd')) 'True' '回退产物 .vsd 已写出'
  Assert-Eq ([IO.Path]::GetExtension($out)) '.vsd' '返回的是回退后的路径'
  Assert-Eq (Test-Path ($x + '.bak')) 'False' '备份 .bak 已清理'

  # 假文档：什么都写得进去 → .vsdx 写成功，旧文件被覆盖，.bak 删除
  $docOk = New-Object PSObject
  $docOk | Add-Member -MemberType ScriptMethod -Name SaveAs -Value {
    param($p)
    Set-Content -LiteralPath $p -Value 'NEW-CONTENT' -Encoding UTF8
  }
  $y = Join-Path $tmp 'y.vsdx'
  Set-Content -LiteralPath $y -Value 'OLD-CONTENT' -Encoding UTF8
  $out2 = Save-VisioDoc -Doc $docOk -Page $null -VsdPath $y -PngPath $null
  Assert-Eq ([IO.Path]::GetExtension($out2)) '.vsdx' '写成功时返回 .vsdx'
  Assert-Eq ((Get-Content -Raw -LiteralPath $y).Trim()) 'NEW-CONTENT' '内容已写入 .vsdx'
  Assert-Eq (Test-Path ($y + '.bak')) 'False' '写成功后备份已删除'
} finally {
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------- 4) 单元格告警：打印一次后清空 ----------
$mockShape = New-Object PSObject
$mockShape | Add-Member -MemberType ScriptMethod -Name CellsU -Value { param($c) throw 'mock: guard 拦截' }
$null = Set-VisioCell $mockShape 'PinX' '1 in'
$w1 = (Write-VisioCellWarnings | Out-String)
Assert-Eq ($w1 -like '*PinX*') 'True' '单元格写入告警被打印出来'
$w2 = (Write-VisioCellWarnings | Out-String)
Assert-Eq ($w2.Trim()) '' '打印后清空（不会重复打印）'

if ($script:failCount -gt 0) {
  Write-Host ('自检失败 ' + $script:failCount + ' 项')
  exit 1
}
Write-Host '自检全部通过'
exit 0
