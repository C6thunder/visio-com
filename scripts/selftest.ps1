# selftest.ps1 — 纯逻辑自检，**不需要 Visio**
#
# 覆盖：字体 ID 回退逻辑（Resolve-VisioFontIndex）——目标字体缺失时回退微软雅黑，
#      两者都缺时返回 0（调用方不写 Char.Font），不能静默用错字体 ID。
#
# 用法：powershell -File scripts\selftest.ps1      （退出码 0 = 全过）
param()

$ErrorActionPreference = 'Stop'
$skillRoot = if ($env:VISIO_SKILL_ROOT) { $env:VISIO_SKILL_ROOT }
             elseif ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot }
             else { Join-Path $env:USERPROFILE '.agents\skills\visio-diagram' }
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 (Join-Path $skillRoot 'scripts\visio-lib.ps1'))))

# 把字体查询换成假的（dot-source 后再定义同名函数会覆盖库里的实现）
$script:FakeFonts = @{ '微软雅黑' = 245; 'Times New Roman' = 23 }
function Get-VisioFontIndex {
  param([string]$Name, $Doc = $null)
  if ($script:FakeFonts.ContainsKey($Name)) { return $script:FakeFonts[$Name] }
  return 0
}

$script:failCount = 0
function Assert-Eq($actual, $expected, [string]$label) {
  if ("$actual" -eq "$expected") { Write-Host ('[OK]   ' + $label + ' = ' + $actual) }
  else { Write-Host ('[FAIL] ' + $label + '：期望 ' + $expected + '，实际 ' + $actual); $script:failCount++ }
}

Assert-Eq (Resolve-VisioFontIndex -Name 'Times New Roman' -Doc $null) 23 '存在的字体直接返回索引'
Assert-Eq (Resolve-VisioFontIndex -Name '不存在的字体' -Doc $null) 245 '缺失时回退微软雅黑'
Assert-Eq (Resolve-VisioFontIndex -Name '微软雅黑' -Doc $null) 245 '微软雅黑本身'

$script:FakeFonts = @{}   # 这台机器一个字体都没有
Assert-Eq (Resolve-VisioFontIndex -Name '任意字体' -Doc $null) 0 '两个都缺时返回 0（不写 Char.Font）'

if ($script:failCount -gt 0) {
  Write-Host ('自检失败 ' + $script:failCount + ' 项')
  exit 1
}
Write-Host '自检全部通过'
exit 0
