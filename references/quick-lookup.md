# Visio 绘图速查（入口文件，读这一个就够开始画）

> 这是**画图时手边放着**的速查：选型、函数签名、检索命令、常见坑都在这里。
> 画图流程与硬规则见 `SKILL.md`；深挖某个库的全部形状名见 `stencil-catalog.md`；
> 单个项目的实测坐标与配方见 `case-notes.md`。

## 1. 30 秒选型表

| 想画的东西 | 模具库 | 形状名 | 坑 |
|---|---|---|---|
| 卡片 / 容器 / 圆角框 | `BASIC_M` | 矩形、圆角矩形、圆形、菱形、梯形、五边形、环形、框架 | 圆角用 `Rounding` 单元格，别拖「圆角矩形」母版 |
| 纯圆点（实心 / 空心圈） | `BASIC_M` | 圆形 | `MATRIXDOT_M`「点 1」自带**编号徽标**，不是纯圆点 |
| 箭头（直线） | `ARROWS_M` | 直线箭头、直线、虚线、块状单箭头、普通双箭头、弧线箭头 | 「普通箭头」方向被 guard 锁死；单向箭头用「线 + 实心箭头端」 |
| 流程框 / 判定框 | `BASFLO_M` | 流程、判定、子流程、开始/结束、文档、数据、数据库、外部数据 | 判定是菱形母版，尺寸可改 |
| 步骤条 / 阶段箭头 | `PROCESS_M` | 流程步骤、箭头步骤、连续流程箭头 | V 形/箭头步骤**宽高被锁**，设尺寸无效 → 改用 `BASFLO_M` |
| 泳道 / 网关 | `BPMN_M` | 任务、网关、池、通道、开始事件、结束事件 | 一张图只用一套记法 |
| IDEF0 功能建模 | `IDEF0_M` | 活动框、节点、实心连接线、标题块 | 左→右输入、底→顶控制 |
| 2×2 矩阵 / SWOT | `MATRIXDIAGRAMSHAPES_M` | 基本矩阵 2x2、SWOT 1、定位图、安索夫矩阵 | **自绘更好控制**：格用 `BASIC_M` 矩形拼，母版是固定版式 |
| 矩阵里的点 / 象限标记 | `MATRIXDOT_M` | 点 1~4、菱形 1~5、三角、箭头 1~11、星形 | 带徽标；纯形状改用 `BASIC_M` |
| 统计小图标 | `ANALYTICS_M` | 条形图、靶心、上升/下降趋势、仪表、齿轮、沙漏、灯泡、拼图 | 图标风，装饰性内容慎用 |
| 折线 / 柱状 / 饼 | `CHART_M` | 折线图、条形图 1、饼图、数据点、X-Y 轴 | 是**形状**不是数据绑定图；真数据图用 Excel/Python 出 |
| 时间线 / 里程碑 | `TIMELN_M` | 块状时间线、线形时间线、菱形/针形/圆柱形里程碑、区间 | 实验流程轴用它最合适 |
| 甘特图 | `GANTT_M` | 甘特图框架、任务栏、里程碑、链接线 | 框架是整组，改样式先解散 |
| 组织结构图 | `ORGCH_M` / `ORGPER_M` | 高管、经理、职位、助理、员工、小组框架 | 8 个同族库只换外形，语义相同，别混用 |
| 维恩图 | `VENNDIAG_M` | 2/3 集合维恩图、交集、并集、差集 | 标题在母版里，改色要递归子形状 |
| 金字塔 / 漏斗 | `PYRAMID_SHAPES_M` | 基本棱锥图、分层棱锥图、漏斗、倒棱锥图 | 分层块是整组，逐层改色要递归 |
| 循环 / 迭代 | `CYCLEDIAG_M` | 环形箭头、箭头循环、射线循环、6 步箭头环 | 方向靠 `Angle` 微调 |
| 数据库 / 实体关系 | `DBCHEN_M` / `DBCROW_M` / `DBUML_M` | 实体、弱实体、属性、主键属性、关系、关系连接线 | 三个库名字高度相似，选一个用到底 |
| 鱼骨 / 因果 | `CAUSEF_M` | 鱼骨框架、主要原因 1/2、次要原因 1~6、效果 | 骨架是母版，枝条独立可加可减 |
| 故障树 / 逻辑门 | `FAULT_M` | AND 门、OR 门、异或门、投票门、基本事件 | 用标准符号，别用矩形代替 |
| 注释框 / 引出 | `CALOUT_M` / `ANNOT_M` | 中间文本标注、折线标注、批注、标注-直线、圆形标注 | 一张图最多 1–2 处 |
| 图例条目 / 状态 | `LGND_M` / `LEGEND_M` | 良好、糟糕、重要、疑问、警告、完成、待执行、优先级 1~3 | 状态必须**配形状或图标** |
| 标题栏 / 页眉 | `TITLE_M` | 标题、区域-4、块-顶标签、页码、比例 | 期刊图一般不用 |
| 底纹 / 边框花色 | `BCKGRN_M` / `BORDRS_M` | 实心、垂直/水平渐变、角部渐变、平铺 | 底纹是**图片**，印刷会去色，期刊慎用 |
| 自动跟随的连接线 | `CONNEC_M` | 动态连接线、直线连接线、弯折的连接线、一对多、侧边到侧边 1/2 | 要吸附就用它；静态图用「线 + 箭头端」更快 |
| 云 / 网络 / 服务器图标 | `AWS*` `AZURE*` `NET*` | 数百个图标库 | 图标多为组形状，改色**必须递归**；论文里克制用 |

## 2. 图种 → 模具组合

| 要画的图 | 主形状 | 连接 / 标记 | 备注 |
|---|---|---|---|
| 流程图 | `BASFLO_M` | `CONNEC_M` 动态连接线 | 起止/流程/判定三件套覆盖八成场景 |
| 跨职能流程 | `BASFLO_M` + `TITLE_M`「部分标题」做泳道头 | `CONNEC_M` | 泳道回答"谁负责" |
| 象限 / 对照矩阵 | `BASIC_M` 矩形自绘 | `BASIC_M` 圆形 / 菱形 | 母版「基本矩阵 2x2」塞不进多行中文 |
| 时间线 / 路线图 | `TIMELN_M`、`GANTT_M` | `BLOCK3_M` 箭头 | 连续时间用线形，阶段用块状 |
| 循环 / 闭环 | `CYCLEDIAG_M` | `ARROWS_M` 环形箭头 | 3–6 阶段最清晰 |
| 组织架构 | `ORGCH_M` + `ORGBND_M` | 正交连接线 | 汇报线走直角，避免斜交叉 |
| 系统 / 网络架构 | `SERVER_M`、`NETSYM_M` | `BLOCK_M`「双树枝直角」 | 正交走线 + 同侧汇入 |
| 数据 / ER 模型 | `ENTITY_M`、`DBCROW_M` | — | 主键外键靠字段行，不靠颜色 |
| 金字塔 / 漏斗 | `PYRAMID_SHAPES_M` | — | 层级用棱锥，衰减用漏斗 |
| 维恩关系 | `VENNDIAG_M` | — | ≤3 集合；4+ 换表格/UpSet |
| 期刊示意图（通用） | `BASIC_M` + 1–2 个语义库 | 线 + 实心箭头端 | **3 个族、5 个母版**就够 |

## 3. 选型坑（每条都踩过）

| 坑 | 表现 | 正确做法 |
|---|---|---|
| `Masters.ItemU()` | 中文名返回 null 且不报错 | 用 `Item("中文名")`，ItemU 只认英文通用名 |
| 同名形状跨库撞车 | 「判定」在 AUDIT_M / BASFLO_M / UACTME_M 都有 | 取用必须带库名：`$st['BASFLO_M'].Item('判定')` |
| 组形状改色不生效 | 父级 `FillForegnd` 没反应 | 递归 `$s.Shapes` 给每个子形状上色 |
| 尺寸 / 方向 guard | 设了 Width/Height 仍是大方块 | `PROCESS_M` V 形/箭头步骤、`ARROWS_M`「普通箭头」方向锁死 → 换母版或自绘 |
| 母版带默认文本 | 拖出来自带「在此处输入文本」 | 删文本，或用 `BASIC_M` 原生图元 |
| 母版带徽标 | `MATRIXDOT_M`「点 1」自带数字 | 纯圆点用 `BASIC_M`「圆形」 |
| 竖排文字 | `TxtAngle` 不生效 | 用形状 `Angle = 270°` |
| 中文回退宋体 | `Char.Font` 只接受字体索引 | 中文一律 `Char.Font = '245'`（微软雅黑） |
| 单位换算 | `Char.Size`/`LineWeight` 的 `ResultIU` 是**英寸** | 9.5 pt = 0.132 in，×72 才是 pt |
| 空心形状白填充落在白底上 | 菱形/圆点画完"看不见" | 空心形状**必须留描边**（`New-VisioDiamond` 默认已带灰描边） |
| **保存时整页重绘一次** | 主会话里就地 `Set-VisioPageScale` 缩页面 → 重新排版重绘，看着像"又画了一遍" | 主会话**只画一次、只存一次**；期刊件用 `New-VisioJournalDeliverables` 在不可见会话里从已存文件派生 |

## 4. 我要做 X → 打开哪个文件

**完整索引在 `SKILL.md` 的「文件路由」表**，本节只列本文件覆盖不到的部分：

| 我要做 | 打开 |
|---|---|
| 想知道某个库全部形状（几百个） | `stencil-catalog.md` 第 1–6 节 |
| 想看 238 个模板的全貌 | `stencil-catalog.md` 第 7 节（模板 .vstx）；第 8 节是 **362 个模具库**地图 |
| 查全库形状名（grep，不读全文） | `references/data/visio-masters-full.tsv`（9 112 行，**只 grep 别整读**） |
| 看某个具体项目的复盘（坐标、叠层配方、实测分工） | `case-notes.md` |

## 5. 库函数速查（`scripts/visio-lib.ps1`）

dot-source 之后直接用：

```powershell
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\visio-lib.ps1')))
```

| 函数 | 作用 |
|---|---|
| `Start-VisioSession` | 起 Visio（可见/不可见）并返回干净会话，**会先杀掉残留实例并关掉残留文档** |
| `Close-VisioSession -App $app` | 关文档 + 退出应用 |
| `Open-VisioStencil -App $app -Names @('BASFLO_M')` | 打开模具库，返回 `@{ 'BASFLO_M' = <stencil doc> }` |
| `Get-VisioMaster -Stencil $st -Name '判定'` | 按**中文名**取 master（不要用 ItemU） |
| `Find-VisioMaster -Pattern '判定','时间线'` | 全库检索形状（走缓存清单，毫秒级） |
| `Get-VisioMasterByName -Stencil 'BASFLO_M' -Name '判定'` | 直接拿可 Drop 的 master 对象 |
| `New-VisioCard` | 圆角卡片（容器角色，最常用） |
| `New-VisioText` | 纯文字框 |
| `New-VisioVerticalText` | 竖排文字（`Angle=270°`） |
| `New-VisioDot` | 圆点（实心/空心，标记角色） |
| `New-VisioDiamond` | 菱形（判定/关键项标记） |
| `New-VisioArrow` | 单向箭头（线 + 实心箭头端） |
| `New-VisioAxisArrow` | 坐标轴 / 双向箭头 |
| `Set-VisioCell -Shape $s -Cell 'FillForegnd' -Formula 'RGB(234,242,247)'` | 写单元格。参数是 `-Shape` / `-Cell` / `-Formula`；颜色公式只能写 `RGB(r,g,b)`，**不能写 `#RRGGBB`**（ShapeSheet 不认，写了会被静默丢弃） |
| `Get-VisioFontIndex -Name 'Times New Roman' -Doc $doc` | **查字体 ID**（`Char.Font`/`Char.AsianFont` 只认字体 ID；必须运行时查，别硬编码；带会话级缓存）。**中文必须同时写 `Char.AsianFont`，否则回退宋体(5)** |
| `Style-VisioTree` | 递归整棵形状树（组形状上色/改线用） |
| `Add-VisioPageBackground -Page $page` | 铺满页白底（**导出前必做**，否则四周留白被切）。页面尺寸由 `Start-VisioSession -Width/-Height` 决定，此函数**没有** `-WidthIn/-HeightIn` 参数 |
| `Save-VisioDoc` | 存 `.vsdx`（失败退 `.vsd`）+ 导出 PNG |
| `New-VisioJournalDeliverables` | **从已存的母版文件派生期刊件**（缩到成品尺寸 → 期刊 .vsdx + 矢量 PDF + 预览 + 576 dpi PNG），在不可见会话里做 |
| `Save-VisioPdf -Doc $doc -Path $p` | 导出矢量 PDF（必须用这个，`SaveAs .pdf` 会报"参数无效"） |
| `Export-VisioPageRaster -Page $page -Path $p -Zoom 4` | 导出高分辨率位图（144 dpi × Zoom，参数名是 `-Zoom` 不是 `-Scale`） |
| `Set-VisioPageScale` | 整页等比缩到成品尺寸（如 183 mm） |
| `Set-VisioShapeScale` | 单个形状等比缩放 |
| `Get-VisioShapeReport -Page $page` | 打印形状构成（自查用） |
| `Invoke-VisioAutoLayout -Page $page -Style flowchart_tb -SpacingIn 0.6` | Visio 原生自动布局（flowchart_tb/lr、tree_tb/lr、radial、circular） |
| `New-VisioConnector -App $app -Page $page -FromShape $a -ToShape $b -Label '是'` | 胶合动态连接线（两端 GlueTo，随布局重路由）；`-Route`/`-LinePattern`/`-LineWeightPt`/`-LineColor` |
| `Get-VisioPageState -Page $page` | 结构化读取所有形状：shape_id / 坐标 / 尺寸 / 文字 / is_connector |
| `New-VisioContainer -App $app -Page $page -Label '区' -MemberIds @(1,2)` | Visio 内建容器包装形状（**必须最后建**，见下方注意） |

## 6. 检索命令

```powershell
# 全库关键词检索（毫秒级）+ 同名形状提示
$code = Get-Content -Raw -Encoding UTF8 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\visio-find-master.ps1'
Invoke-Expression $code
# → Find-VisioMaster -Pattern '判定','时间线'

# 直接 grep 全量清单（不要整读 398 KB 文件）
Select-String -Path 'C:\Users\thunder\.agents\skills\visio-diagram\references\data\visio-masters-full.tsv' -Pattern '时间线'
```

## 7. 换机器

跑 `scripts/visio-inventory.ps1` 重扫模具库清单（**不启动 Visio，3 秒**），产出自动落到 `references/data/`。

## 8. 照着参考图重绘（用户给一张图要"用 Visio 重画"）

1. **先量比例**：从参考图的像素宽高比定页面（如 1774×887 → 2:1 → 13.333×6.667 in 母版），别照抄像素当英寸。
2. **抽色板**：按段取主色（浅底 + 中描边 + 深字），定成 `$theme` 数组，一段一色，后面每张卡直接索引。
3. **定栅格**：算卡片宽 =（可用宽 − 间距×(n−1))/n，再给窄卡片单独调宽（内容多的段更宽）。
4. **图标用几何图元近似**：Visio 画不出 3D 图标 → 用圆角方块 + 白色字形（BASIC_M 的 圆形/圆柱形/齿轮/五边形/三角形 + 画线）拼，图标是"识别线索"不是主角。
5. **逐轮渲染核对**：每轮跑完导出 PNG 用 `read_image` 看，重点查**中文竖排**（文本框太窄会被挤成一列）、**词被拆断**（拉丁串如 `ModuleNotFoundError` 要给足宽度）、**标签压字**（徽标列与正文列分开量）。
6. **正文列宽度经验**：中文 9 pt ≈ 0.125 in/字，拉丁 ≈ 0.055 in/字符；算不够就换行或降字号，别让它自动折行。

> **5 pt 下限的例外**：2:1 这类高密度概览图缩到 183 mm 宽后最小字号会掉到 4 pt 以下。默认仍守 5 pt（`new-figure.ps1` 已设 `-MinFontPt 5.0`）；确实放不下时，正确做法是**加大成品尺寸或减少信息密度**，不是直接放行。只有当期刊明确允许小尺寸插图时，才显式把 `-MinFontPt` 设为 0，并在图注里写明。

## 9. 立体感 + 公式排版 + 字号（三件套，本机实测）

### 9.1 Visio 没有阴影 / 渐变 / 透明单元格

`FillGradient`、`Gradient*`、`Shdw.*`、`FillTransparency`、`Bevel*` **在 ShapeSheet 里全部不存在**（`CellsU` 直接抛错）；UI 里的"效果→阴影/发光/棱纹"不落 ShapeSheet，COM 设不了；`Transparency` 单元存在但设了不生效。

→ **立体感只能靠形状叠层**（三层，全图偏移方向统一右下）——具体配方见 `case-notes.md` §2。

### 9.2 公式排版

- `Char.Font` 只接受**索引**，但 `Document.Fonts` 可枚举（本机 258 个）→ 用 `Get-VisioFontIndex -Name 'Times New Roman' -Doc $doc` **运行时查**（索引随字体表顺序变，换机器就错，**别硬编码**）。
  本机实测：微软雅黑 245、Times New Roman 23、Cambria 47、Cambria Math 48、Calibri 4、Arial 21、Segoe UI 183。
- **公式用 Times New Roman + `Char.Style = 2`（斜体）** → 立刻像数学公式（衬线 + 斜体变量）。
- **中点 '·'（U+00B7）在微软雅黑和 Times 里前后都自带空隙**，`α·Sim` 会排成 `α ·Sim`，很难看；U+22C5/U+2219 同理。→ **改用乘号 '×'（U+00D7）**，间距天然对称，最像公式。
- PowerShell 里拼希腊字母用码点，别硬敲：`$a = [char]0x03B1; $b = [char]0x03B2`。
- `Characters.CharProps` 在 PowerShell 晚绑定下**不可用**（"找不到成员"）→ 想让公式里不同片段不同颜色/字体，只能**拆成多个文本框**手工定位；或整条统一 Times 斜体（更省事，够用）。

### 9.3 字号经验（13.333 in 母版 / 6 列卡片）

| 角色 | 母版字号 | 缩到 183 mm 后 |
|---|---|---|
| 卡片标题（6 汉字） | 13 pt | 7.0 pt |
| 副标题 | 9.5–10 pt | 5.1–5.4 pt |
| 正文 / 项目符号 | 10–10.5 pt | 5.4–5.7 pt |
| 次要说明、图例值 | 9.5 pt | 5.1 pt |
| 徽标标签 | 9.5 pt | 5.1 pt |
| 底部公式条 | 19 pt | 10.3 pt |

- **母版字号一律 ≥ 9.5 pt**：183 mm 成品宽度下 9 pt 会缩成 4.6 pt，低于 Nature 的 5 pt 下限（与 `SKILL.md`「母版最小 9.5 pt」一致）。
- **6 个汉字 @13 pt ≈ 1.08 in** → 标题列至少留 1.15 in，否则标题折行；标题起始位置要避开左上角序号徽标（徽标 0.44 in 直径 + 0.1 间隙）。
- 拉丁字符在**微软雅黑**里比 Times 宽（@10 pt ≈ 0.07 in/字符）→ `ModuleNotFound`、`x86_64/LoongArch64` 这类长串要单独给宽度或降号到 9 pt。
- 卡片宽度按"每列内容长度"分配（内容多的段更宽），不要平均分；平均分必然有一列折行。
- 期刊版等比系数 = 183 mm / 13.333 in = **0.5404**。

## 10. 版面重构：按内容改结构（"6 个竖框太难看"时用这招）

单排 N 个等宽竖框是**版式惯性**，不是内容要求。用户说"太难看/允许根据内容改变排版"时，按下表改结构：

| 症状 | 改法 |
|---|---|
| 一排 6 个等宽竖框，又高又窄 | 改 **2×3 蛇形流水线**：上排 1→2→3（从左到右），下排 4→5→6（**从右到左**），右端画一根垂直箭头把 3 接到 4 |
| 某张卡内容很少，卡内大片空 | 别用"小框 + 大空卡"，改成**无框陈述式**：左侧竖色条 + 小标签 + 大字正文（16–21 pt），内容顶满 |
| 某张卡内容极简（只有一个数字/词） | 用**焦点图形**：白色圆形 + 粗描边 + 圈里大字（如 "Top-K"），旁边配一行说明 |
| 一列里多条目被挤成两行 | 卡片变宽后**每条一行**（图标｜名称｜注释｜标签），注释列可给到 1.4–1.5 in |
| 一列里 4 个短定义挤成 4 行 | 改成 **2×2 网格**（每格：粗体名一行 + 释义一行） |
| 两组并列子块被压扁 | 左右并列两个子面板，每个 1.9 in 宽即可容下 10 字标题 + 项目符号 + 标签 + 按钮 |

### 10.1 2×3 蛇形：可迁移的结论

- 竖框变宽后，**标题可以放到 16 pt**（6 汉字 ≈ 1.33 in，宽卡不再挤）。
- 宽卡里 48 字符的公式：11 pt 一行（≈3.5 in）或 13 pt 分两行，别让它自动折行。
- 蛇形下排箭头**方向朝左**，编号 1…6 保持递增即可自证流向。

该项目的实测坐标（列宽、行高、卡内字号）见 `case-notes.md` §1。

### 10.2 脚本与 COM 的坑（每条都踩过）

- **变量名大小写不敏感**：`$w = $page.Drop(...)` 会把页面宽度 `$W` 覆盖成 COM 对象 → 后面 `$W - $Margin` 报"op_Subtraction"。几何量用 `$PageW/$PageH`，形状句柄用 `$tri/$badge` 这种全称。
- **数组字面量里的 `+` 被当成分隔符**：`@(0.45, 0.45 + $W + $gap, ...)` 会展开成 6 个元素（`0.45, 0.45, $W, $gap, ...`），三个面板全叠一起。元素只要含运算就给每个元素单独加括号：`@(0.45, (0.45 + $W + $gap), (0.45 + 2 * ($W + $gap)))`。
- **`[math]::Max` 只收 2 个参数**：3 个以上要嵌套 `[math]::Max($a, [math]::Max($b, $c))`，否则报"找不到 Max 的重载"。
- **Visio 形状宽/高不能设 0**：`Height = 0` 抛"指定的转换无效"。文本框先给 0.2 in 的临时高度、量好行高再定位。
- **`TxtWrap` / `TxtAutoGrow` / `TxtLineSpacing` 不是合法单元格名**：`CellsU` 会抛"意外的文件尾"。换行 / 自动撑高没有对应单元格，得自己按字数估算行高（见 9.3 节）+ 用 `Para.HorzAlign` / `VerticalAlign` / `LeftMargin` 排版。
- **直线母版（ARROWS_M「直线」）的 `BeginX/BeginY/EndX/EndY` 就是页坐标**：直接写目标坐标即可，`PinX/PinY` 自动跟随。画分隔线 / 连线就用它（别用隐藏母版 BPMN_M「分隔符」，Drop 返回空对象）。
- **CFF 容器（BPMN_M）是组形状，改宽高比会失真**：把 6×4 的容器改成 3.9×6.2 后，标题带子形状的宽度公式失效（4 宽的标题带会变成 6.2 宽、伸出容器外）。要标题带就别整体改宽高比，或用普通矩形自绘容器 + 标题条。
- **杀掉 Visio 后立刻重开会偶发"模具打不开"**：`Stop-Process VISIO` 之后 `Start-Sleep 4` 再开新会话；失败就重跑一次整脚本（脚本可重复执行）。
- **脚本不会杀你正在编辑的 Visio**：`Start-VisioSession` 发现有未保存文档会直接报错停下；`Stop-VisioProcess` 也只警告不强杀。遇到报错就先保存或关闭那些文档。

## 11. 移植自 visio-mcp 的 4 个能力（自动布局 / 胶合连接线 / 页检查 / 容器）

这 4 个函数从 jagarlamudisrinath/visio-mcp 移植而来，用 PowerShell COM 直接实现，不依赖 MCP 服务器。

**容器有顺序约束**：
- `New-VisioContainer` 内部走 `ContainerProperties.AddMember`；**成员加入容器后不再属于 `Page.Shapes`（变成容器的子形状）**，所以此后再 `Page.Shapes.ItemFromID(成员ID)` 会取不到。
- 正确顺序：**放形状 → 连线 → 建容器（最后一步）→ 保存**。建容器后不要再调 `Invoke-VisioAutoLayout`。
- 若需要容器内继续操作，用 `$container.Shapes` 找成员。

## 12. 迭代调参与环境体检（省 Visio 启动次数）

**同一个会话里迭代**：改一个参数就重跑整个脚本 = 每次白花 8–12 s 启动 Visio。正确姿势：

```powershell
$s = Start-VisioSession -Width 13.333 -Height 7.5
# ...画...
for ($round = 1; $round -le 5; $round++) {
    $s.Page.Export("iter-$round.png")     # 不重启，直接再导
    # 用 compare-render.ps1 量偏差 → 改 Cell → 下一轮
    Set-VisioCell $shape 'Char.Size' '23 pt'
}
```

- `scripts/visio-doctor.ps1`：**一个会话跑完**字体表 / 存盘格式(.vsdx/.vsd/.pdf) / 临时目录 / 模具目录 / 打开既有文件的体检，替代手写一堆探测脚本。
- `scripts/compare-render.ps1`：量 PNG 里指定区域内的墨迹包围盒，报「偏移 / 溢出 / 空区域」，用于「先量后画」。
