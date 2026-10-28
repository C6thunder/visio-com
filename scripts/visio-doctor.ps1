# visio-doctor.ps1 — 一个 Visio 会话跑完本机 Visio 环境体检
# 用法：
#   powershell -File visio-doctor.ps1
#   powershell -File visio-doctor.ps1 -TestFile C:\path\existing.vsdx
#   powershell -File visio-doctor.ps1 -OutLog C:\path\doctor.txt
#
# 检查项：Visio 路径/版本、字体表、默认页面、存盘格式(.vsdx/.vsd/.pdf)、
#         TEMP 可写、模具目录、打开既有文件。
param(
    [string]$TestFile = '',
    [string]$OutLog = ''
)
$ErrorActionPreference = 'Stop'
$lib = Join-Path $PSScriptRoot 'visio-lib.ps1'
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 $lib)))

$rows = New-Object System.Collections.ArrayList
function Add-Check([string]$item, [string]$status, [string]$detail) {
    [void]$rows.Add([pscustomobject]@{ 项目 = $item; 结果 = $status; 详情 = $detail })
}
function Say([string]$s) { Write-Host $s }

$app = $null
try {
    # ---------- 1. Visio 可执行文件 ----------
    $exe = $script:VisioExe
    if ($exe -and (Test-Path $exe)) {
        $ver = (Get-Item $exe).VersionInfo.FileVersion
        Add-Check 'Visio 可执行文件' 'OK' "$exe ($ver)"
    } else {
        Add-Check 'Visio 可执行文件' 'FAIL' "未找到：$exe"
    }

    # ---------- 2. 开一个会话（后面所有检查共用）----------
    $s = Start-VisioSession -Width 10 -Height 7
    $app = $s.App; $doc = $s.Doc; $page = $s.Page
    Add-Check '启动 Visio 会话' 'OK' ($app.Version + ' / ' + $app.Name)

    # ---------- 3. 字体表 ----------
    $fonts = @('微软雅黑', '等线', '宋体', '黑体', 'Arial', 'Times New Roman')
    $fp = @()
    foreach ($f in $fonts) {
        $idx = Get-VisioFontIndex -Name $f -Doc $doc
        $fp += ('{0}={1}' -f $f, $idx)
        if ($idx -eq 0) { Add-Check ('字体 ' + $f) 'WARN' '未找到（Get-VisioFontIndex 返回 0）' }
    }
    Add-Check '字体表' 'OK' ($fp -join '  ')

    # ---------- 4. 默认页面 ----------
    Add-Check '页面尺寸' 'OK' ('{0} x {1} in' -f [math]::Round($page.PageSheet.CellsU('PageWidth').ResultIU, 3), [math]::Round($page.PageSheet.CellsU('PageHeight').ResultIU, 3))

    # ---------- 5. 存盘格式 ----------
    $st = Open-VisioStencil -App $app -Names @('BASIC_M')
    $mRect = Get-VisioMaster -Stencil $st['BASIC_M'] -Name '矩形'
    $probe = $page.Drop($mRect, 5, 3.5); $probe.Text = 'doctor'; $probe.CellsU('Width').ResultIU = 2; $probe.CellsU('Height').ResultIU = 1

    $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('visio-doctor-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

    foreach ($ext in @('.vsdx', '.vsd')) {
        $p = Join-Path $tmpDir ('probe' + $ext)
        try { $doc.SaveAs($p); Add-Check ('存盘 ' + $ext) 'OK' ((Get-Item $p).Length.ToString() + ' bytes') }
        catch { Add-Check ('存盘 ' + $ext) 'FAIL' $_.Exception.Message.Trim() }
    }
    $pdf = Join-Path $tmpDir 'probe.pdf'
    try {
        # PDF 必须走 ExportAsFixedFormat（Page.Export 到 .pdf 不生成文件）
        $pdfOut = Save-VisioPdf -Doc $doc -Path $pdf
        if ($pdfOut -and (Test-Path $pdfOut)) { Add-Check '导出 .pdf' 'OK' ((Get-Item $pdfOut).Length.ToString() + ' bytes') }
        else { Add-Check '导出 .pdf' 'FAIL' 'Save-VisioPdf 未生成文件' }
    } catch { Add-Check '导出 .pdf' 'FAIL' $_.Exception.Message.Trim() }
    $png = Join-Path $tmpDir 'probe.png'
    try { $page.Export($png); Add-Check '导出 .png' 'OK' ((Get-Item $png).Length.ToString() + ' bytes') }
    catch { Add-Check '导出 .png' 'FAIL' $_.Exception.Message.Trim() }

    # ---------- 6. TEMP 可写 ----------
    try {
        $tf = Join-Path ([System.IO.Path]::GetTempPath()) ('vd-' + [guid]::NewGuid().ToString('N') + '.tmp')
        [System.IO.File]::WriteAllText($tf, 'x')
        Remove-Item $tf -Force
        Add-Check 'TEMP 可写' 'OK' ([System.IO.Path]::GetTempPath())
    } catch { Add-Check 'TEMP 可写' 'FAIL' $_.Exception.Message.Trim() }

    # ---------- 7. 模具目录 ----------
    try {
        $root = $script:StencilRoot
        if ($root -and (Test-Path $root)) {
            $n = @(Get-ChildItem $root -Filter *.vssx -ErrorAction SilentlyContinue).Count
            Add-Check '模具目录' 'OK' "$root（$n 个 .vssx）"
        } else { Add-Check '模具目录' 'FAIL' "不存在：$root" }
    } catch { Add-Check '模具目录' 'FAIL' $_.Exception.Message.Trim() }

    # ---------- 8. 打开既有文件 ----------
    if ($TestFile) {
        try {
            $d2 = $app.Documents.OpenEx($TestFile, 1)   # visOpenRO=1
            Add-Check '打开既有文件' 'OK' ($d2.Name + ' / pages=' + $d2.Pages.Count)
            try { $d2.Close() } catch {}
        } catch { Add-Check '打开既有文件' 'FAIL' $_.Exception.Message.Trim() }
    } else {
        Add-Check '打开既有文件' 'SKIP' '未传 -TestFile'
    }

    # 清理探针目录
    try { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}
catch {
    Add-Check '整体执行' 'FAIL' ($_.Exception.Message.Trim())
}
finally {
    if ($app) { try { Close-VisioSession -App $app } catch {} }
}

# ---------- 输出 ----------
$fail = @($rows | Where-Object { $_.结果 -eq 'FAIL' })
$warn = @($rows | Where-Object { $_.结果 -eq 'WARN' })
Say ''
Say '================ Visio Doctor ================'
$rows | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Say ("汇总: OK={0} WARN={1} FAIL={2}" -f @($rows | Where-Object { $_.结果 -eq 'OK' }).Count, $warn.Count, $fail.Count)

if ($OutLog) {
    $lines = @('================ Visio Doctor ================')
    foreach ($r in $rows) { $lines += ('[{0}] {1} : {2}' -f $r.结果, $r.项目, $r.详情) }
    $lines += ('汇总: OK={0} WARN={1} FAIL={2}' -f @($rows | Where-Object { $_.结果 -eq 'OK' }).Count, $warn.Count, $fail.Count)
    [System.IO.File]::WriteAllText($OutLog, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    Say ('日志: ' + $OutLog)
}

if ($fail.Count -gt 0) { exit 1 } else { exit 0 }
