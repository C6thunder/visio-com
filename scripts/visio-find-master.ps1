# visio-find-master.ps1 — 在本机全部模具库里按关键词检索形状（绘图前查模具用）
#
# 用法：
#   $code = Get-Content -Raw -Encoding UTF8 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\visio-find-master.ps1'
#   Invoke-Expression $code
#
#   # 1) 检索（走缓存清单，毫秒级）
#   Find-VisioMaster -Pattern '判定','时间线' | Format-Table -AutoSize
#
#   # 2) 只要某个库里的形状
#   Find-VisioMaster -Pattern '菱形' -Stencil 'BASFLO_M'
#
#   # 3) 直接拿到 master 对象去画
#   $r = Get-VisioMasterByName -Stencil 'BASFLO_M' -Name '判定'
#   $s = $page.Drop($r.Master, 3.0, 3.0)   # 期间不要 Close($r.Doc)
#
# 清单由 scripts/visio-inventory.ps1 生成（references\data\visio-masters-full.tsv）；
# 没有清单时会退化为直接扫描（慢）。
param(
  [string[]]$Pattern = @('流程'),
  [string]$Stencil = '',
  [int]$Limit = 60
)

# 用 Invoke-Expression 执行时 $PSScriptRoot 为空，这里做兜底
$skillDir = if ($PSScriptRoot) { $PSScriptRoot } else { 'C:\Users\thunder\.agents\skills\visio-diagram\scripts' }
$libPath = Join-Path $skillDir 'visio-lib.ps1'
if (-not (Test-Path $libPath)) { $libPath = 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\visio-lib.ps1' }
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 $libPath)))

$hits = Find-VisioMaster -Pattern $Pattern -Limit 500
if ($Stencil) { $hits = $hits | Where-Object { $_.Stencil -like $Stencil } }

Write-Output ('检索词: ' + ($Pattern -join ' / ') + '   命中 ' + @($hits).Count + ' 条')
$hits | Select-Object -First $Limit | Format-Table -AutoSize

# 提示同名形状
$dup = $hits | Group-Object Name | Where-Object { ($_.Group.Stencil | Select-Object -Unique).Count -gt 1 }
if ($dup) {
  Write-Output ''
  Write-Output '⚠ 以下形状名在多个模具库里都有，取用时必须带库名：'
  foreach ($d in $dup) {
    Write-Output ('  ' + $d.Name + ' → ' + (($d.Group.Stencil | Select-Object -Unique) -join ', '))
  }
}
