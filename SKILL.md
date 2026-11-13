---
name: visio-diagram
description: 用 Visio 画专业图示并通过 PowerShell COM 自动化落盘——流程图、2×2 对照矩阵、服务架构图、研究示意图，输出可编辑的 .vsd/.vsdx 加高分辨率 PNG 与矢量 PDF。只要用户提到 Visio、要 .vsd/.vsdx 源文件、想把框图做成期刊级图件、要求"圆角""多形状拼凑""科研绘图规范""图例配色"，或者要画流程图/矩阵图/架构图/示意图，就用本技能——即使用户没说"Visio"两个字，只要目标是"可编辑的矢量示意图"也适用。
whenToUse: 用户要求用 Visio 画图或修改 Visio 文件；要求把方框示意图做成更专业的图件（圆角、跨模具混用、期刊配色）；要求输出 .vsd/.vsdx 或高分辨率 PNG；需要在 Windows 上无人值守地批量生成 Visio 图形。
metadata:
  platform: windows
  requires: Microsoft Visio 2016+ / PowerShell COM
---

# Visio 示意图

COM 驱动本机 Visio 画图，产出**可编辑**源文件 + 位图（+ 矢量 PDF）。适合：对照矩阵、流程图、架构分层图、研究示意图。
不适合：需要自由手绘插画的场景（那用 SVG/canvas 更省事）。

## 前置检查

1. Visio 可执行文件存在 —— 路径在 `scripts/visio-lib.ps1` 顶部的 `$script:VisioExe`（默认 Office16，装别的版本/位数就改这一处）。
2. 本机模具目录存在 —— 同上，`$script:StencilRoot`（`...\Visio Content\2052\*.vssx`）。
   两项都由 `powershell -File scripts\visio-doctor.ps1` 体检，**别猜路径，先跑它**。

## 三步法

1. **定交付形态**（决定页面尺寸，错了要重画）
   - 投期刊 → 按**成品尺寸**建页：单栏 89 mm / 双栏 183 mm 宽，正文 5–7 pt，线宽 ≥ 0.5 pt
   - 汇报/幻灯 → 16:9（13.333×7.5 in）
   - 细则见 `references/journal-figure-standards.md`
2. **定信息关系，再选形状**（关系决定形状族，不是反过来）
   矩阵=因素×因素、流程=先后顺序、分层=上下级、循环=闭环。
   选 3–5 个模具族按角色分工：**容器** `BASIC_M` 矩形 / **语义** `BASFLO_M`·`VENNDIAG_M`·`TIMELN_M` / **连接** 线+实心箭头端 / **标记** `BASIC_M` 圆形·菱形 / **图例** `LGND_M`。
   形状名与坑 → `references/quick-lookup.md`（**先读这个**）；全量形状名 → `references/stencil-catalog.md`；语义决策 → `references/shape-selection.md`。
3. **写脚本、画、存、导出、亲眼看**
   ```powershell
   # PS 5.1 按系统 ANSI 读文件，UTF-8 无 BOM 的中文脚本会被读坏 → 必须显式 -Encoding UTF8
   # $skillRoot：本机技能根目录；装在默认位置 ~/.agents/skills/visio-diagram 时可省略
   # （PS 5.1 没有 `??`，用 if/else）
   $skillRoot = if ($env:VISIO_SKILL_ROOT) { $env:VISIO_SKILL_ROOT } else { Join-Path $env:USERPROFILE '.agents\skills\visio-diagram' }
   $code = Get-Content -Raw -Encoding UTF8 (Join-Path $skillRoot 'scripts\visio-lib.ps1')
   . ([scriptblock]::Create($code))
   ```
   **别从零写构建脚本**：复制 `scripts/new-figure.ps1`（已含三段式骨架、空白页断言、满页白底、期刊缩放、576 dpi 派生会话导出），在"在这里画图"处填内容即可。
   常用函数：`Start-VisioSession` · `Open-VisioStencil` · `Get-VisioMaster` · `New-VisioCard` · `New-VisioText` · `New-VisioDot` · `New-VisioArrow` · `Add-VisioPageBackground` · `Save-VisioDoc` · `Save-VisioPdf` · `Export-VisioPageRaster` · `Set-VisioPageScale` · `Invoke-VisioAutoLayout` · `New-VisioConnector` · `Get-VisioPageState` · `New-VisioContainer`（全表见 `quick-lookup.md` §5）。

## 硬规则（每条都踩过，代价是重画）

| 规则 | 原因 |
|---|---|
| 中文文本必须**同时**设 `Char.Font` 和 `Char.AsianFont`（默认微软雅黑） | **`Char.Font` 只管西文，中文走 `Char.AsianFont`；不设 AsianFont 中文会静默回退宋体(5)**。两个单元格都要写，用 `Resolve-VisioFontIndex` 运行时查（字体缺失会回退并告警） |
| **公式用 Times New Roman 斜体**，索引运行时查 | 微软雅黑里 '·' 前后自带空隙，`α·Sim` 排成 `α ·Sim`；`Char.Font` 又不认字体名。做法：`Get-VisioFontIndex -Name 'Times New Roman'` + `Char.Style=2`，乘号改 '×'（U+00D7） |
| **Visio 没有阴影/渐变/透明单元格** | `FillGradient`/`Gradient*`/`Shdw.*`/`FillTransparency` 全部不存在，UI 的"效果"也设不了。立体感用**形状叠层**：投影层 + 深色侧面层 + 正面层，偏移方向全图统一（见 `case-notes.md` §2） |
| `Characters.CharProps` 在 PowerShell 里不可用 | 晚绑定报"找不到成员"。想让同一段文字里不同片段不同样式，只能拆成多个文本框 |
| 母版按名查找用 `Masters.Item("中文名")`，**不要 ItemU** | ItemU 走通用名，中文一律返回 null 且不报错 |
| 同名形状跨库撞车 | 「判定」在 AUDIT_M/BASFLO_M/UACTME_M 都有，「圆形」「菱形」更普遍 → 取用必须带库名 |
| 组形状上色要递归 `Shapes` | 块箭头、多数母版是组，父级 `FillForegnd` 不生效 |
| 单向箭头用「线 + 实心箭头端」，别用 `ARROWS_M`「普通箭头」 | 该母版方向被 guard 锁死，`Angle`/`FlipX` 都改不动 |
| `PROCESS_M`「V 形步骤/箭头步骤」不要试图改尺寸 | 宽高被 guard 锁定，设成 2.7×0.32 仍是大方块 |
| `MATRIXDOT_M` 是**带编号徽标的数据点** | 要纯圆点用 `BASIC_M`「圆形」 |
| 竖排文字用形状 `Angle = 270°` | `TxtAngle` 在本机不生效 |
| `Char.Size` / `LineWeight` 的 `ResultIU` 单位是**英寸** | 9.5 pt = 0.132 in，差 72 倍会整页文字隐形 |
| **写长度类单元格必须带单位**（`Char.Size`/`LineWeight`/`Rounding`/`Width`/`Height`…） | `Set-VisioCell $s 'Char.Size' 10.5` 不是 10.5 pt，而是 **10.5 英寸 = 756 pt**：字撑爆页面，`Page.Export` 只报笼统的「发生错误。Visio 无法完成导出」，产物是 **0 字节**文件，EMF 一起坏（header bounds 高 34 in）；PDF 却能正常导出，所以极容易误判成“图太复杂”。一律 `("{0} pt" -f $Size)` / `("{0} in" -f $w)`。同类坑：`Angle` 单元是**弧度**，写 45 = 45 rad |
| **母版最小 9.5 pt（13.333 in 母版）** | 缩到 183 mm 后是 5.1 pt，守住 Nature 的 5 pt 下限 |
| **期刊图不用彩色文字**，线宽 ≥ 0.5 pt | Nature 明令 "Avoid coloured text"；0.75 pt 缩到成品只有 0.41 pt |
| 期刊图必须**最终尺寸建页** | 编辑部按成品排版，大画布缩小后 9 pt 字缩成 2.5 pt |
| **铺满页白底再导出**（`Add-VisioPageBackground`） | 不铺的话 `Page.Export` 只裁图形外接框，四周留白全被切掉 |
| **主会话只画一次、只存一次；缩放别在主会话做** | 就地 `Set-VisioPageScale` 会整页重新排版重绘，窗口里看着像"保存时又画了一遍"，还会把母版污染成缩放后的尺寸。期刊件用 `New-VisioJournalDeliverables` 从已存文件在**独立会话**里派生（自己起的实例不可见；借用用户已开的实例时会在他的窗口里跑） |
| 存盘先试 `.vsdx`，失败落 `.vsd`；PDF 用 `Save-VisioPdf` | `SaveAs .pdf` 抛"参数无效" |
| Visio 是**单实例**；查模具别用 COM 逐个开 | 残留实例会让旧图新图同时出现；362 库用 COM 开要 20~30 分钟（读 XML 只要 3 秒） |
| **变量名大小写不敏感**：`$c` 与 `$C` 是**同一个变量** | `$w = $page.Drop(...)` 会覆盖页面宽 `$W`，后面 `$W - 0.19` 直接报 "op_Subtraction"；用 `$c` 表卡片、`$C` 表颜色则不报错、静默取错值，最难查。约定：**函数参数 PascalCase（`$Page`/`$App`/`$Doc`）、局部变量 camelCase（`$page`/`$app`/`$doc`）**，几何量用 `$PageW/$PageH`，句柄用全称。库已按此约定，新写代码照抄 |
| 脚本里**中文引号“”不能进双引号字符串** | PowerShell 把全角引号也当引号，字符串会被截断，报“缺少终止符”。要么用单引号，要么把文案改写成 `「」` |
| **杀掉 Visio 后要等几秒再开会话** | 立刻重开会偶发"模具打不开 / 对象引用未设置"；`Stop-Process VISIO` 后 `Start-Sleep 4` |
| **脚本不杀你正在编辑的 Visio** | `Start-VisioSession` 发现未保存文档会直接报错停下，`Stop-VisioProcess` 只警告不强杀。避免脚本毁掉用户未保存的图；遇到报错先保存或关闭那些文档 |
| **实例归属：别裸开、别乱 Quit** | Visio 是单实例——`New-Object -ComObject Visio.Application` 不会新起进程，只是**附着**到已在跑的实例；裸开的实例还不设 `AlertResponse`，叠加 `Visible=$false` 后任何模态框（保存确认/兼容性提示）都会让 COM 调用**永久阻塞**（脚本卡死 + 隐形残留进程）；随后 `Quit()` 会关掉用户的窗口，并可能连带丢掉未保存文档（`Stop-VisioProcess` 的未保存检查拦不住已执行过的 `Quit`）。**画新图用 `Start-VisioSession`**（杀残留、等退出、`AlertResponse=7`、单实例）；**派生/检索这类可能撞上用户实例的操作用 `New-VisioAppLease`**（用户已开 Visio 就**借用**：不 Quit、不杀进程、不改可见性，只还原 `AlertResponse`），用完 `Close-VisioAppLease` 释放 |
| **"N 个并排竖框"是版式惯性，可按内容改结构** | 6 段流水线可改 2×3 蛇形（上排向右、下排向左 + 右端垂直箭头）；宽卡能让标题到 16 pt、每条一行、2×2 网格。手法表见 quick-lookup §10 |
| 沙箱会在命令结束时回收子进程 | 一次脚本里画完、存完、导完 |
| **异常路径也要收尾** | 画到一半抛异常 → Visio 进程带着未保存草稿残留 → 下次 `Start-VisioSession` 发现未保存文档直接报错停下（自锁）。`new-figure.ps1` 已用 `trap` 收尾（关会话 + 打单元格告警）；自己写脚本照抄这个模式 |
| **迭代调参在同一个会话里做**，别每轮重跑整个脚本 | 每次 `Start-VisioSession` 要 `Start-Process` 新 Visio（8–12 s）。改参数看一眼的循环应该是：画 → `Page.Export` → 读图/量像素 → 改 cell → 再导出，全程不重启。多轮重跑会白白多花几分钟 |
| **先量后画**：用 `scripts/compare-render.ps1` 量 PNG 里的墨迹包围盒，别肉眼估算 | 文字上移 10 px、标题偏窄 7% 这类偏差肉眼看不准；量化后再改，一轮就能收敛 |
| 绘制默认可见、逐步停顿、绘完停 1 秒再关窗口 | 让用户在 Visio 窗口里实时看到绘制过程与成品：`Start-VisioSession` 不传 `-Invisible`，每个主要步骤之间 `Start-Sleep`，收尾用 `Close-VisioSession -App $app -DelaySeconds 1`；`-Invisible` 只留给批量/期刊派生的独立会话 |
| **函数名别撞 PowerShell 内置别名** | 自定义 `function R` 会被解析成 `Invoke-History` 别名，报「找不到接受实际参数」；`function Measure` 撞 `Measure-Object`、`function Select` 撞 `Select-Object` 同理。**短名一律加动词前缀**（`Invoke-`/`Get-`/`New-`/`Set-`/`Add-`），如 `Add-Check` 而不是 `R` |

## 交付前自检

- [ ] 形状族 3–5 个，同一语义只用一种形状，颜色有图例
- [ ] 关系用连接线，箭头样式全图一种，线宽只有 0.75 / 1.5 pt 两档
- [ ] 导出 PNG **用 read_image 亲眼看过**：无裁字、无重叠、深色底上的圆点用白色
- [ ] 期刊件：页面=成品尺寸、正文 5–7 pt、无彩色文字、灰度 + 色觉障碍模拟各看一遍
- [ ] 交付：可编辑源文件 + 矢量 PDF（RGB、字体嵌入）+ ≥300 dpi 位图

## 文件路由

| 场景 | 打开 |
|---|---|
| 选形状（日常） | `references/quick-lookup.md` — 选型表、坑、函数速查 |
| 某库的全部形状名 / 模板 / 362 库地图 | `references/stencil-catalog.md` — 第 1–6 节高频库全名、第 7 节模板、第 8 节全库地图 |
| 形状语义决策（箭头/流程/标记/混用） | `references/shape-selection.md` |
| 配色与色盲安全 | `references/palettes.md` |
| 期刊尺寸/字号/导出规格 | `references/journal-figure-standards.md` |
| 模具全量清单（**只 grep，别整读**） | `references/data/visio-masters-full.tsv`（9 112 行） |
| 单个项目的实测坐标 / 叠层配方 / 分工示例 | `references/case-notes.md` |

## 脚本

| 脚本 | 调用方式 | 关键参数 | 输出 / 退出码 | 失败了怎么办 |
|---|---|---|---|---|
| `visio-doctor.ps1` | `powershell -File scripts\visio-doctor.ps1` | `-TestFile <既有.vsdx>` `-OutLog <log.txt>` | 表格 `OK/WARN/FAIL/SKIP` + 汇总；**FAIL → exit 1**，否则 0 | 先看 FAIL 行：Visio 路径/字体/存盘格式/模具目录，对症改环境再重跑 |
| `visio-inventory.ps1` | `powershell -File scripts\visio-inventory.ps1` | `-Root <模具目录>` `-OutDir <输出目录>`（默认 `references\data`） | 写 `visio-masters-full.tsv` / `visio-stencils.tsv` / `visio-templates.tsv` / `visio-stencil-summary.md` | 换机器或装了新模具后重跑；不需要 Visio |
| `visio-find-master.ps1` | `powershell -File scripts\visio-find-master.ps1 -Pattern '判定','时间线'` | `-Pattern <关键词[]>` `-Stencil <库名>` `-Limit` | 表格列出 模具库/名称/NameU/类型，末尾提示同名形状；无退出码 | 没命中就换关键词（中文名或 NameU 都试），或先跑 `visio-inventory.ps1` |
| `shape-sheet.ps1` | `powershell -File scripts\shape-sheet.ps1` | `-Groups <库名→母版名[]>` `-OutDir` `-OutName` | 一页候选母版样张 PNG/源文件 | 看样张再定形状；某母版打不开先查 `visio-find-master.ps1` 的库名 |
| `new-figure.ps1` | `powershell -File scripts\new-figure.ps1 -Name 'fig-1' -Preset slides` | `-Name` `-Preset slides\|journal` `-OutDir` `-HighRes` | `[build] 阶段 1/3…` 日志 + 源文件(.vsdx/.vsd) + 预览 PNG；journal 另出 PDF 与 576 dpi 位图 | 形状数不对 → 看"绘制完成：形状 N"；存盘失败 → 先跑 `visio-doctor.ps1` 看存盘格式项 |
| `compare-render.ps1` | `powershell -File scripts\compare-render.ps1 -PngPath out\render.png -RegionsJson regions.json` | `-PngPath` `-RegionsJson`（**必填**）`-TolerancePx` `-WhiteThreshold` `-OutLog` | 每区域 `OK/SHIFT/OVERFLOW/BORDER/EMPTY` + 墨迹包围盒与偏移 dx/dy；**异常 → exit 1** | SHIFT/OVERFLOW → 改**形状属性**（位置/尺寸/字号）；只有整体分辨率不对时才改导出参数 |
| `visio-lib.ps1` | dot-source 后调函数 | — | 见 `references/quick-lookup.md` §5 | — |
| `selftest.ps1` | `powershell -File scripts\selftest.ps1` | — | `[OK]/[FAIL]` 逐项；**失败 → exit 1**。纯逻辑自检，**不需要 Visio**（字体回退、字体缓存按文档隔离、存盘回退不丢旧文件、单元格告警只打一次） | 改完 `visio-lib.ps1` 的公共逻辑后先跑它 |
| `example-2x2-matrix.ps1` | `powershell -File scripts\example-2x2-matrix.ps1` | `-OutDir` `-BaseName` | 2×2 对照矩阵全套产物 | 当模板抄，别当规范 |

> 带 `-Name/-Pattern` 等参数的脚本：**必须 dot-source + 调用运算符**（`. ([scriptblock]::Create($code)) -Name 'fig-1'`），`Invoke-Expression` 传不了参数。
