# compare-render.ps1 — 量 PNG 里指定区域的墨迹包围盒，报告偏移/溢出/空区域
# 用途：“先量后画”——不要肉眼估算文字上移多少 px、标题偏窄百分之几。
#
# 用法：
#   powershell -File compare-render.ps1 -PngPath out\render.png -RegionsJson regions.json
#   powershell -File compare-render.ps1 -PngPath out\render.png -RegionsJson regions.json -OutLog cmp.txt
#
# regions.json 格式（坐标是 PNG 像素坐标，左上原点）：
#   [ {"name":"Title",  "x":35, "y":14, "w":200, "h":30},
#     {"name":"Card01", "x":32, "y":63, "w":194, "h":130} ]
#
# 每个区域会输出：
#   ink      —— 该区域内墨迹包围盒（绝对坐标）+ 宽高
#   center   —— 墨迹中心相对区域中心的偏移 dx/dy（px）
#   overflow —— 墨迹是否触到区域边界（文字溢出/贴边）
#   empty    —— 区域内没有墨迹（内容丢失或位置错了）
param(
    [Parameter(Mandatory)][string]$PngPath,
    [Parameter(Mandatory)][string]$RegionsJson,
    [double]$TolerancePx = 2.0,
    [int]$WhiteThreshold = 245,   # 亮度 >= 此值算背景
    [string]$OutLog = ''
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $PngPath)) { throw "PNG 不存在: $PngPath" }
$regions = Get-Content -Raw -Encoding UTF8 $RegionsJson | ConvertFrom-Json

$bmp = [System.Drawing.Bitmap]::FromFile((Resolve-Path $PngPath))
$W = $bmp.Width; $H = $bmp.Height
$rows = New-Object System.Collections.ArrayList

function Get-InkBox([int]$rx, [int]$ry, [int]$rw, [int]$rh) {
    $x0 = [math]::Max(0, $rx); $y0 = [math]::Max(0, $ry)
    $x1 = [math]::Min($W - 1, $rx + $rw - 1); $y1 = [math]::Min($H - 1, $ry + $rh - 1)
    if ($x1 -lt $x0 -or $y1 -lt $y0) { return $null }
    $minX = [int]::MaxValue; $minY = [int]::MaxValue; $maxX = -1; $maxY = -1
    $count = 0
    for ($y = $y0; $y -le $y1; $y++) {
        for ($x = $x0; $x -le $x1; $x++) {
            $c = $bmp.GetPixel($x, $y)
            $lum = [int](0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B)
            if ($lum -lt $WhiteThreshold) {
                $count++
                if ($x -lt $minX) { $minX = $x }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    if ($count -eq 0) { return $null }
    return [pscustomobject]@{
        x = $minX; y = $minY; w = ($maxX - $minX + 1); h = ($maxY - $minY + 1); pixels = $count
    }
}

foreach ($r in $regions) {
    $name = [string]$r.name; $rx = [int]$r.x; $ry = [int]$r.y; $rw = [int]$r.w; $rh = [int]$r.h
    $box = Get-InkBox $rx $ry $rw $rh
    if ($null -eq $box) {
        [void]$rows.Add([pscustomobject]@{ 区域 = $name; 状态 = 'EMPTY'; 墨迹 = '-'; 偏移dx = '-'; 偏移dy = '-'; 溢出 = '-'; 说明 = '区域内无墨迹' })
        continue
    }
    $icx = $box.x + $box.w / 2.0; $icy = $box.y + $box.h / 2.0
    $rcx = $rx + $rw / 2.0;       $rcy = $ry + $rh / 2.0
    $dx = [math]::Round($icx - $rcx, 2); $dy = [math]::Round($icy - $rcy, 2)

    $ov = @()
    if ($box.x -le $rx) { $ov += 'L' }
    if ($box.x + $box.w - 1 -ge ($rx + $rw - 1)) { $ov += 'R' }
    if ($box.y -le $ry) { $ov += 'T' }
    if ($box.y + $box.h - 1 -ge ($ry + $rh - 1)) { $ov += 'B' }

    $status = 'OK'
    if ([math]::Abs($dx) -gt $TolerancePx -or [math]::Abs($dy) -gt $TolerancePx) { $status = 'SHIFT' }
    if ($ov.Count -gt 0) {
        # 墨迹正好填满整个区域 = 区域本身就是带边框的容器/底框，不算溢出
        $fills = ($box.x -le ($rx + 1)) -and ($box.y -le ($ry + 1)) -and
                 (($box.x + $box.w) -ge ($rx + $rw - 1)) -and
                 (($box.y + $box.h) -ge ($ry + $rh - 1))
        if ($fills) { $status = 'BORDER' } else { $status = 'OVERFLOW' }
    }

    $note = ''
    if ($status -eq 'BORDER') { $note = '墨迹填满区域（容器边框，正常）' }
    elseif ($status -eq 'OVERFLOW') { $note = '触边 ' + ($ov -join '') + '（文字可能被裁切）' }
    elseif ($status -eq 'SHIFT') { $note = ('超出容差 {0}px' -f $TolerancePx) }

    [void]$rows.Add([pscustomobject]@{
        区域 = $name; 状态 = $status
        墨迹 = ('{0},{1} {2}x{3}' -f $box.x, $box.y, $box.w, $box.h)
        偏移dx = $dx; 偏移dy = $dy
        溢出 = $(if ($ov.Count) { ($ov -join '') } else { '-' })
        说明 = $note
    })
}

# 全图墨迹包围盒（必须在 Dispose 之前）
$all = Get-InkBox 0 0 $W $H
$bmp.Dispose()
$lines = @('================ compare-render ================')
$lines += ('PNG: {0}  ({1}x{2})' -f $PngPath, $W, $H)
if ($all) { $lines += ('全图墨迹: {0},{1} {2}x{3}' -f $all.x, $all.y, $all.w, $all.h) }
else { $lines += '全图墨迹: 空（全白）' }
$lines += ''
$rows | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
$lines += ($rows | Format-Table -AutoSize | Out-String -Width 200)

$bad = @($rows | Where-Object { $_.状态 -ne 'OK' -and $_.状态 -ne 'BORDER' })
$lines += ('汇总: 区域={0} 异常={1} 容差={2}px' -f $rows.Count, $bad.Count, $TolerancePx)
Write-Host ('汇总: 区域={0} 异常={1} 容差={2}px' -f $rows.Count, $bad.Count, $TolerancePx)

if ($OutLog) {
    [System.IO.File]::WriteAllText($OutLog, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    Write-Host ('日志: ' + $OutLog)
}
if ($bad.Count -gt 0) { exit 1 } else { exit 0 }
