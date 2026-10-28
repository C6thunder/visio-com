# visio-com

用 **PowerShell + COM** 直接驱动本机 Microsoft Visio 画出版级图表的 AI Agent Skill。

不依赖任何 MCP 服务器、不生成外部图片、全部使用 Visio 原生模具；支持中文字体精确控制、原生表格拼装、期刊级导出、以及**像素级渲染验收**。

> 适用于：[Claude Code / Codex / DeepSeek Harness 等](https://github.com) 支持 `SKILL.md` 规约的 Agent。

---

## 能力

| 能力 | 说明 |
|---|---|
| 原生模具 | 362 个 `.vssx` 模具库 + 238 个模板，母版检索、打样张选型 |
| 中文字体 | **同时设 `Char.Font` 和 `Char.AsianFont`**，避免中文静默回退宋体（本机等线=252、微软雅黑=245） |
| 表格 | Visio **没有原生两列表格母版**，用「矩形外框 + 直线分隔 + 文本框」拼装，单元格自动折行 |
| 期刊图 | 576 dpi 导出、满页白底、页宽缩放、矢量 PDF（`ExportAsFixedFormat`） |
| 连接线 | 胶合动态连接线（两端 `GlueTo`）、自动路由、标签 |
| 自动布局 | Visio 原生 `Page.Layout`：flowchart_tb/lr、tree_tb/lr、radial、circular |
| 容器 | Visio 内建容器母版 + `ContainerProperties.AddMember` |
| 像素验收 | `compare-render.ps1` 量 PNG 墨迹包围盒，报 `OK / SHIFT / OVERFLOW / BORDER / EMPTY` |
| 环境体检 | `visio-doctor.ps1` 一个会话跑完字体表 / 存盘格式 / TEMP / 模具目录 |

## 环境要求

- Windows
- Microsoft Visio（2016 / 2019 / 2021 / 365，实测 16.0.20527）
- Windows PowerShell 5.1+
- 不需要 Python，不需要 MCP

## 安装

把本仓库放进 Agent 的 skills 目录即可：

```powershell
git clone https://github.com/<you>/visio-com.git "$HOME\.agents\skills\visio-diagram"
```

或直接复制目录到 `~/.agents/skills/visio-diagram`。

## 快速开始

```powershell
# 1) 先体检，确认环境可用（约 15 秒）
powershell -File scripts\visio-doctor.ps1

# 2) 复制骨架改内容
copy scripts\new-figure.ps1 my-figure.ps1
#    在"在这里画图"处填内容，然后运行
powershell -File my-figure.ps1

# 3) 量一下渲染结果（先量后画）
powershell -File scripts\compare-render.ps1 -PngPath out\render.png -RegionsJson regions.json
```

Agent 读 `SKILL.md`（工作流 + 硬规则）和 `references/quick-lookup.md`（函数速查）即可上手。

## 目录

```
SKILL.md                          Agent 主入口：9 步工作流 + 硬规则
references/
  quick-lookup.md                 函数速查（31 个封装函数 + 用法片段）
  shape-selection.md              形状/母版选型
  stencil-catalog.md              模具库目录
  palettes.md                     配色
  journal-figure-standards.md     期刊图规范
  case-notes.md                   踩坑记录（含所有已证伪的假说）
  data/                           本机 362 库 / 238 模板清单（TSV）
scripts/
  visio-lib.ps1                   核心库：31 个函数（会话/母版/文本/卡片/连接线/容器/导出/缩放）
  new-figure.ps1                  画图骨架，复制它改
  example-2x2-matrix.ps1          完整示例：2×2 对照矩阵
  visio-doctor.ps1                环境一体检
  compare-render.ps1              像素量化验收
  visio-find-master.ps1           全库关键词检索形状
  visio-inventory.ps1             重扫模具清单
  shape-sheet.ps1                 打样张（候选母版真实外观）
```

## 核心函数（节选）

```powershell
Start-VisioSession -Width 13.333 -Height 7.5     # 可见/不可见会话，已处理 AlertResponse、残留进程
Open-VisioStencil -App $app -Names @('BASIC_M')  # 开模具
New-VisioText -Page $page -X1 0 -Y1 0 -X2 5 -Y2 1 -Text '标题' -Size 16 -Bold -Font '等线'
New-VisioCard -Page $page -X1 1 -Y1 1 -X2 3 -Y2 2 -Title '卡片' -Lines @('第一行','第二行')
New-VisioConnector -App $app -Page $page -FromShape $a -ToShape $b -Label '是'
Invoke-VisioAutoLayout -Page $page -Style flowchart_tb -SpacingIn 0.6
New-VisioContainer -App $app -Page $page -Label '容器' -MemberIds @(1,2)
Save-VisioDoc -Doc $doc -Page $page -VsdPath out\fig.vsdx -PngPath out\fig.png
Save-VisioPdf -Doc $doc -Path out\fig.pdf
```

## 硬规则（都是踩过的坑）

1. **中文必须同时设 `Char.Font` 和 `Char.AsianFont`** —— 前者只管西文，不设后者中文会静默回退宋体。
2. **别裸用 `New-Object -ComObject Visio.Application`** —— 不设 `AlertResponse` 再叠加 `Visible=$false`，任何模态框都会让 COM 永久阻塞。
3. **迭代表情在同一个会话里做** —— 每轮重启 Visio 要白花 8–12 秒。
4. **先量后画** —— 用 `compare-render.ps1`，不要肉眼估算偏移。
5. **`.ps1` 必须带 UTF-8 BOM** —— PowerShell 5.1 无 BOM 时按 CP936 读，中文注释直接解析崩。
6. **`.vsdx` 存盘可能失败** —— 用 `Save-VisioDoc`，失败会自动回落 `.vsd` 并给出警告。
7. **Visio 没有阴影/渐变单元格** —— `ShapeShadow`、`ShadowBlur` 等都不存在。
8. **`Page.Layout()` 会让旧 COM 代理失效** —— 布局后按 ID 重新取形状。

## 已知限制

- 纯 Windows + 本机 Visio，不能跨平台。
- Visio 没有原生表格 / 阴影 / 渐变，相关效果需用形状拼装或放弃。
- `Characters.CharProps` 在 PowerShell 晚绑定下不可用，请用 `Char.*` 单元格。

## License

未指定。如需开源协议，建议补一份 MIT / Apache-2.0。
