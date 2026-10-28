# Visio 模具库与模板全量目录（查阅用）

> **这是"库里有什么"的全量目录**；日常选型先看 `quick-lookup.md`（更短、够用）。
> **数据**：本机 `C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052`
> —— **362 个模具库（.vssx）/ 9112 个 master / 238 个模板（.vstx）**。
> **换机器**：跑 `scripts/visio-inventory.ps1` 重扫（3 秒，不启动 Visio）。
> **原始数据**（在 `references\data\`，只 grep 别整读）：`visio-masters-full.tsv`（全量 master）、
> `visio-stencils.tsv`（库↔英文名↔数量）、`visio-stencil-summary.md`（每库形状名摘要）、
> `visio-templates.tsv`（模板↔配套模具）。

---

## 1. 基础形状与容器（最常用，先看这里）

### BASIC_M（58 个）

矩形 · 正方形 · 圆形 · 椭圆形 · 直角三角形 · 三角形 · 旋转三角形 · 五边形 · 六边形 · 七边形 · 八边形 · 十边形 · 圆柱形 · 平行四边形 · 梯形 · 菱形 · 十字形 · V 形 · 立方体 · 水滴 · 半圆 · 半圆 · 圆锥 · 倒圆锥 · 棱锥 · 尖的椭圆形 · 漏斗 · 齿轮 · 四角星 · 五角星 · 六角星 · 七角星 · 十六角星 · 二十四角星 · 三十二角星 · 圆角矩形 · 剪去单角的矩形 · 剪去同侧角的矩形 · 剪去对角的矩形 · 单圆角矩形 · 同侧圆角矩形 · 对角圆角矩形 · 剪去单圆角矩形 · 剪去角的矩形 · 圆角的矩形 · 剪去角的和圆角矩形 · 框架 · 框架角 · L 形状 · 斜纹 · 徽章 · 环形 · 无符号 · 中心拖动圆形 · 左圆括号 · 右圆括号 · 左括号 · 右括号

**用它的理由**：原生图元，Width/Height/Rounding/Fill 全可控，不会带母版的隐藏文本和护栏。
需要圆角矩形时用 `矩形` + `Rounding` 单元格，不要去拖 `圆角矩形` 母版（它带固定版式）。

### ARROWS_M（28 个）/ ARROWS2_M（16 个）

普通双箭头 · 普通箭头 · 现代箭头 · 可伸缩的箭头 · 圆角右箭头 · 手杖形箭头 · 尖圆角右箭头 · 左弧形箭头 · 右弧形箭头 · 多行 · 多箭头 · 2D 多行 · 虚尾箭头 · 燕尾形箭头 · 箭头总汇 · 环形箭头 · 十字箭头 · 丁字箭头 · 左右箭头汇总 · 十字箭头汇总 · 线性双箭头 · 虚线 · 直线 · 块状单箭头 · 块状双箭头 · 直线箭头 · 弧线箭头 · 弧线双箭头

ARROWS2_M 全是**线轮廓/轮廓**系列（线轮廓(顺时针弯曲) · 线轮廓(逆时针弯曲) · 线轮廓(向左旋转) · 线轮廓(向右旋转) · 线轮廓(轻微弯曲) · 线轮廓(直线) · 线轮廓(U 形弯，带尾巴) · 线轮廓(U 形弯) · 轮廓(顺时针弯曲) · 轮廓(逆时针弯曲) · 轮廓(向左旋转) · 轮廓(向右旋转) · 轮廓(轻微弯曲) · 轮廓(直线) · 轮廓(U 形弯，带尾巴) · 轮廓(U 形弯)），画弧形流程线才用得上。

**已知坑**：「普通箭头」「现代箭头」方向被 guard 锁死（`Angle=180`、`FlipX=1` 都改不动）；
「可伸缩的箭头」「多箭头」是组形状，尺寸改的是整体。
**推荐做法**：自己画线 + `LineWeight` + 实心箭头端（`BeginArrow`/`EndArrow`），方向、线宽、箭头样式全部可控。

### BLOCK_M（41 个）/ BLOCK3_M（19 个）/ BLOCKP_M（19 个）

BLOCK_M：圆形 · 框 · 菱形 · 箭头框 · 一维单向箭头 · 一维双向箭头 · 二维单向箭头 · 二维双向箭头 · 双箭头框 · 四向箭头框 · 高度自调节框 · 自动调整大小的框 · 曲线箭头 · 开放式一维单向箭头 · 开放式二维单向箭头 · 三维框 · 三维圆环 · 分层的框 · 开放/闭合条 · 按钮 · 同心圆第 1 层 · 同心圆第 2 层 · 同心圆第 3 层 · 同心圆中心层 · 第 1 层扇环 · 第 2 层扇环 · 第 3 层扇环 · 第 4 层扇环 · 双树枝直角 · 双树枝斜角 · 动态连接线 · 多树枝斜角 · 多树枝直角 · 直线-曲线连接线 · 点式线 · 中间带箭头的直线 · 弧形箭头 · 点箭头线 · 中间带箭头的点式线 · 单箭头 · 双箭头线

BLOCK3_M（3D 块）：正方形块 · 圆形 · 水平条 · 垂直条 · 左箭头 · 右箭头 · 上箭头 · 下箭头 · 左/右箭头 · 上/下箭头 · 开放式上箭头 · 开放式下箭头 · 开放式左箭头 · 开放式右箭头 · 肘形 1 · 肘形 2 · 肘形 3 · 肘形 4 · 框架

BLOCKP_M（平面块）：块 · 圆形 · 向左箭头 · 向右箭头 · 向上箭头 · 向下箭头 · 左/右箭头 · 上/下箭头 · 孔洞 · 浅块 · 三维上/下箭头 · 三维左/右箭头 · 肘形 1 · 肘形 2 · 肘形 3 · 肘形 4 · 线框 1 · 线框 2 · 消失点

**已知坑**：块箭头、多数母版是**组**，父级 `FillForegnd` 不生效，必须递归 `Shapes` 给每个子形状上色。

### CONNEC_M（69 个）——自动连接线

动态连接线 · 直线-弧线连接线 · 曲线连接 1 · 曲线连接 2 · 双树枝直角 · 多树枝直角 · 侧边到侧边 1 · 侧边到侧边 2 · 侧面 · 底部到顶部 1 · 底部到顶部 2 · 顶部或底部 · 侧边到侧边固定 1 · 侧边到侧边固定 2 · 顶部/底部到侧边 · 底部到顶部固定 1 · 底部到顶部固定 2 · 侧边到顶部/底部 · 底部到侧边 1 · 底部到侧边 2 · 底部到侧边 3 · 侧边到顶部 · 直角循环 · 通用连接线 · 布局连接线 · 直线-曲线连接线 · 直线连接线 · 波浪连接线 1 · 波浪连接线 2 · 双树枝斜角 · 多树枝斜角 · 曲线连接 3 · 有向线 1 · 有向线 2 · 一对多 · 弯折的连接线 · 流导向符 1 · 流导向符 2 · 控制传递 · 跨线 · 可变箭头 1 · 一维单向箭头 · 一维双向箭头 · 一维开放端箭头 · 肘形 1 · 肘形 2 · 空心连接 1 · 空心连接 2 · 空心连接 3 · 点式线 · 中间带箭头的直线 · 弧形箭头 · 点箭头线 · 中间带箭头的点式线 · 单箭头 · 双箭头线 · 侧边 - 侧边 · 侧边 - 顶部/底部 · 底部 - 顶部 · 顶部/底部 - 侧边 · 双线 · 双线肘形 · 四线 · 四线肘形 · 通信链路 1 · 通信链路 2 · 总线 · 以太网 · 星形

**用它的理由**：连接线会**自动跟随**端点移动，拖动卡片时连线不会脱节。
代价是连接点吸附规则有学习成本；只画一次、不再拖动的静态图，用「线 + 箭头端」更快。

## 2. 流程图族：五种记法，选一种用到底

同一张图里混用两套记法是最常见的翻车点。

| 记法 | 模具库 | 核心形状 | 适合 |
|---|---|---|---|
| 基本流程 | BASFLO_M | 流程 / 判定 / 子流程 / 开始结束 / 文档 / 数据 | 通用流程，期刊图首选 |
| Microsoft 流程 | MSFLOW_M | 流程步骤 / 流程条件 | Office 风格流程 |
| BPMN | BPMN_M | 任务 / 网关 / 池通道 / 事件 | 跨部门流程、协议图 |
| IDEF0 | IDEF0_M | 活动框 / 节点 / 控制线 | 功能分解、输入-控制-输出 |
| 数据流（Yourdon） | DATFLO_M | 数据流程 / 外部交互方 / 数据存储 | 数据流图 |
| SDL | SDL_M | 起始 / 过程 / 判定 1/2 / 终结符 | 协议状态机 |
| EPC | EPC_M | 事件 / 函数 / 组织单元 / XOR / OR / AND | 事件流程图 |

BASFLO_M（15 个）：流程 · 判定 · 子流程 · 开始/结束 · 文档 · 数据 · 数据库 · 外部数据 · 自定义 1 · 自定义 2 · 自定义 3 · 自定义 4 · 页面内引用 · 跨页引用 · 动态连接线

PROCESS_M（31 个）：V 形步骤 · 彩色 V 形步骤 · 垂直 v 形步骤 · 彩色垂直 V 形 · V 形列表 · 箭头步骤 · 流程箭头 · 3 步箭头流程 · 流程级别 · 多级别 V 形流程 · 简单水平标题 · 简单垂直标题 · 带编号的标题 · 带标题的垂直 V 形列表 · 带编号的垂直列表 · 带编号的水平列表 · 弯曲弧形 · 带标题的弯曲弧形 · 连续流程箭头 · 带标题的连续流程 · 流程大纲 · 垂直流程块 · 带编号的箭头块 · 带编号的块 · 4 步箭头列表 · 弯曲 V 形 · 弯曲箭头 · 标题步骤 · 垂直标题列表 · 流程步骤 · 递阶流程

**已知坑**：`PROCESS_M`「V 形步骤」「箭头步骤」「彩色 V 形步骤」的**宽高被 guard 锁定**，
设成 2.7×0.32 仍显示成大方块；「流程箭头」「多级别 V 形流程」同族问题相同。
要可控就用 BASFLO_M「流程」或自绘矩形。

## 3. 矩阵 / 图表 / 数据可视化

- **MATRIXDIAGRAMSHAPES_M**（15 个）：基本矩阵 2x2 · 矩阵 1 · 矩阵 2 · 矩阵 3 · 基本矩阵 3x3 · 带标签矩阵 3x3 · 带标签的矩阵 · SWOT 1 · SWOT 2 · 三维矩阵 · 定位图 · Ansoff 矩阵 · 波士顿矩阵 · 轴 1 · 轴 2
- **MATRIXDOT_M**（37 个）：点 1 · 点 2 · 点 3 · 点 4 · 菱形 · 菱形 2 · 菱形 3 · 菱形 4 · 菱形 5 · 三角形 1 · 三角形 2 · 圆角三角形 1 · 圆角三角形 2 · 箭头 1 · 箭头 2 · 箭头 3 · 箭头 4 · Arrow 5 · Arrow 6 · Arrow 7 · Arrow 8 · Arrow 9 · Arrow 10 · 箭头 11 · 固定 1 · Pin 2 · 固定 3 · 固定 4 · 固定 5 · 固定 6 · 标志 · 交叉 · 暂存记号 · 星形 1 · 星型 2 · 星型 3 · 星形 4
- **CHART_M**（40 个）：条形图 1 · 条形图 2 · 三维条形图 · 三维轴 · 竖排文本三维条形 · 横排文本三维条形 · 饼图 · 饼图扇区 · 特殊饼图扇区 · 分段条形图 1 · 分段条形图 2 · 流程图 · 部署图 · 功能比较 · 功能开/关 · 行标题 · 列标题 · 网格 · 是/否框 · 标准曲线 · 指数曲线 · 折线图 · 图形线条 · 数据点 · X-Y 轴 · X-Y-Z 轴 · 图形标尺 · X 轴标签 · Y 轴标签 · Z 轴标签 · X 轴 · Y 轴 · Z 轴 · 8pt 的文本块 · 10pt 的文本块 · 12pt 的文本块 · 2-D 文字提示框 · 1-D 文字提示框 · 水平标注 · 批注
- **ANALYTICS_M**（20 个）：条形图 · 靶心 · 计算器 · 数据库 · 下降趋势 · 眼睛 · 筛选 · 仪表 · 齿轮 · 戴装备的头 · 沙漏 · 灯泡 · 放大镜 · 饼图 · 剧本 · 拼图 · 单级齿轮 · 秒表 · 桌子 · 上升趋势
- **CYCLEDIAG_M**（20 个）：环形箭头 · 4 阶段圆周运动 · 饼图弧 · 第 1 层扇环 · 第 2 层扇环 · 第 3 层扇环 · 第 4 层扇环 · 目标线段 · 第 2 级目标线段 · 6 步骤箭头环 · 射线循环 · 箭头循环 · 射线图 · 六边形循环 · 同心圆第 1 层 · 同心圆第 2 层 · 同心圆第 3 层 · 同心圆中心层 · 目标 · 洋葱图
- **VENNDIAG_M**（12 个）：维恩图形状 · 正方形维恩图 · 2 集合正方形维恩图 · 2 集合维恩图 · 带标题的 2 集合维恩图 · 3 集合维恩图 · 带标题的 3 集合维恩图 · 2 集合并集 · 2 集合交集 · 2 集合差集 · 2 集合补集 · 洋葱图
- **PYRAMID_SHAPES_M**（11 个）：基本棱锥图 · 倒棱锥图 · 分层棱锥图 · 倒圆锥 · 带标题的基本棱锥图 · 带标题的倒棱锥图 · 带标签的棱锥图 · 漏斗 · 渐变棱锥图 · 自定义倒圆锥 · 自定义棱锥图

**2×2 对照矩阵的实战结论**：母版「基本矩阵 2x2」只有固定版式，四格里塞不进多行中文；
真正能用的做法是 **BASIC_M 矩形自绘四格 + 独立文字框 + 箭头**，本项目图 1 就是这么画的。

## 4. 层次 / 组织 / 时间 / 集合

- **ORGCH_M**（14 个）：高管 · 经理 · 职位 · 顾问 · 空缺 · 助理 · 小组框架 · 员工 · 多个形状 · 三个职位 · 名称/日期 · 职务 · 动态连接线 · 虚线报告
- **TIMELN_M**（27 个）：块状时间线 · 线形时间线 · 圆柱形日程表 · 线形里程碑 · 菱形里程碑 · 针形里程碑 · 方括号间隔 · 花括号间隔 · 块状间隔 · “今日”标记 · 双三角形里程碑 · X 形里程碑 · 三角形里程碑 · 正方形里程碑 · 圆形里程碑 · 展开的日程表 · 经过的时间 · 同步里程碑 · 同步间隔 · 动态连接线 · 括号间隔 1 · 括号间隔 2 · 规则日程表 · 分段式日程表 · 圆柱形里程碑 · 圆柱形间隔 · 间歇
- **GANTT_M**（18 个）：甘特图框架 · 列 · 行 · 任务栏 · 里程碑 · 链接线 · 标题 · 图例 · 8pt 的文本块 · 10pt 的文本块 · 12pt 的文本块 · 水平标注 · 直角水平 · 动态连接线 · 非工作时间 · 文本输入 · 主要比例单元格 · 次要比例单元格
- **CALNDR_M**（40 个）：月 · 周 · 多周 · 月缩略图 · 年 · 日 · 约会 · 多日事件 · 时钟 · 便笺 · 提醒 · 重要 · 会议 · 图钉 · 注意 · 构思 · 旅行-飞机 · 旅行-火车 · 旅行-汽车 · 旅行-轮船 · 特殊事件 · 假期 · 庆祝 · 生日 · 运动 · 周年纪念日 · 完成 · 里程碑 · 开始 · 圆形标签 · 正方形标签 · 任务 · 月相 · 星形标签 · 晴天 · 晴间多云 · 阴雨 · 闪电 · 暴风雨 · 多云
- **PERT_M**（9 个）：PERT 1 · PERT 2 · 动态连接线 · 直线连接线 · 直线-曲线连接线 · 图例 · 摘要结构 · 水平标注 · 直角水平

**8 个组织图库同构**：ORGBLT（带）/ORGCNS（硬币）/ORGPAN（面板）/ORGPIP（种子）/ORGSTO（石）/
ORGNCH（凹槽）/ORGPET（花瓣）——形状名一模一样（高管、经理、员工…），只是外形不同。
**名称会撞车**：`圆形`、`菱形`、`实体`、`流程` 在多个库里都存在，脚本里必须写成 `Masters.Item('BASFLO_M').Item('判定')` 这种带库名的方式。

## 5. 标注 / 图例 / 注释

- **ANNOT_M**（39 个）：标注 - 曲线 · 标注 - 折弯 · 标注 - 直线 · 圆形标注 · 椭圆形标注 · 框标注 · 12 磅的文本 · 8 磅的文本 · 6 磅的文本 · 已划出 5 行的列 · 已划出 15 行的列 · 信息线 · 中断线 · 截面 1 · 截面 2 · 截面 3 · 基准面(位置) · 基准面 1 · 基准面 2 · 基准 · 水平 · 参考 · 天花板 · 向北箭头 1 · 向北箭头 2 · 向北箭头 3 · 向北箭头 4 · 向北箭头 5 · 修订云形 · 绘图比例 · 比例符号 · 注释符号 · 参考三角形 · 参考矩形 · 参考六边形 · 参考椭圆形 · 参考圆 · 参考标注 1 · 参考标注 2
- **CALOUT_M**（38 个）：侧边外框标注 · 侧边边框线标注 · 侧边文本标注 · 中间框标注 · 中间线条标注 · 中间文本标注 · 侧边肘形框 · 侧边边框线肘形 · 侧边肘形标注 · 中间肘形框 · 中间底线肘形 · 中间肘形标注 · 批注 · 线间文本标注 · 折线标注 · 带文本的大括号 · 椭圆形标注 · 框标注 · 带文本的边框线 · 不完整方括号文本 · 完整方括号文本 · 侧边方括号 · 侧边大括号 · 侧边括号 · 文件 · 黄色便条 · 标记 · 二维文字提示框 · 水平提示框 · 垂直提示框 · 切角图章 · 圆角图章 · 钝爆星 · 锐爆星 · 自动尺寸 · 自定义标注 1 · 自定义标注 2 · 自定义标注 3
- **LGND_M**（24 个）：待执行 · 完成 · 构思 · 需要继续努力 · 重要 · 疑问 · 重要性升高 · 重要性降低 · 图例 · 良好 · 糟糕 · 优先级 1 · 优先级 2 · 优先级 3 · 删除 · 不执行 · 会议 · 信息 · 星形 · 任务 · 注意 · 时间 · 警告 · 注释
- **LEGEND_M**（6 个）：外部列表 · 字段容器 · 内部列表 · CBV 项 · 图标项 · 文本项
- **TITLE_M**（29 个）：区域 - 4 · 区域 - 8 · 边框 · 投影符号 · 已划出 5 行的列 · 已划出 15 行的列 · SI 符号 · 修订块 · 部件清单块 · 部分标题 · 块 - 侧标签 · 块 - 顶标签 · 日期 · 说明 · 制图员 · 文件名 · 文件/路径 · 页码 · 修订日期 · 比例 · 标题 · 大标题块 · 小标题块 · 续页较大 · 续页较小 · 修订区域较大 · 修订区域较小 · 修订较大 · 修订较小

## 6. 图标类模具库（几百个库，基本只做图标）

这些库提供的是**产品/设备/行业图标**，不是结构图元。论文插图里：

1. **慎用**——图标风与线稿混排会显廉价，且大多是位图，印刷质量差；
2. 若必须用（如画系统架构图），优先云厂商图标库，语义明确；
3. **它们几乎全是组形状**，改色/改尺寸必须递归子形状；
4. 单个库动辄 100~290 个 master，**在 Visio 里靠肉眼翻找极慢**，用 `Find-VisioMaster` 检索。

- **云与平台图标（AWS / Azure / …）**：73 个库
- **网络与 IT 设备**：27 个库
- **建筑 / 空间 / 暖通**：23 个库
- **工程 / 电气 / 管道设备**：61 个库
- **行业与生活图标**：16 个库

## 7. 模板（.vstx）是什么、怎么用

本机 **238 个模板**。每个模板 = **预制好的一页 + 配套模具**，例如：

| 模板 | 配套模具 | 页面形状数 |
|---|---|---|
| ACTUAL_FINISHDATES_GANTTCHART_M | 链接线, 甘特图框架, 圆角矩形, 列, 次要比例单元格, 主要比例单元格, 文本输入, 行, 任务栏 | 177 |
| AFFINITY_DIAGRAM_M | 圆角矩形, 正方形, 黄色, 浅绿色, 蓝色, 靛蓝色 | 36 |
| ASSOCIATIONS_UML_CLASS_M | 动态连接线, 复合, 类, 成员, 分隔符, 聚合, 关联, 复合, 圆角矩形 | 76 |
| AUDIT_SEVERAL_OUTCOMES_M | 终止符, 标记的流程, 决策, 动态连接线, 圆角矩形, 事件 | 39 |
| AUDIT_SUBPROCESS_M | 标记的文档, 标记的流程, 决策, 动态连接线, 圆角矩形, 事件, 页面内引用 | 70 |
| AZURELOGANALYTICS_M | 动态连接线, 虚拟机, 齿轮(不透明), 企业(可堆叠), 可扩展的云, Microsoft Azure, 辅助角色(不透明) , Web 角色(不透明) , 存储表, OMS, 通用数据库, 通用数据库(可堆叠), 服务器机架, 服务器/边栏选项卡(可堆叠) | 137 |
| AZUREOMSWORKSPACE_M | 动态连接线, Microsoft Azure, 通用数据库, Web (可堆叠), 应用程序(为添加文本留出的空白) | 34 |
| AZURESITERECOVERY_M | 带编号的标签, 动态连接线, 服务器(边栏选项卡), 齿轮(不透明), 企业(可堆叠), Microsoft Azure, 虚拟机, 通用数据库, Azure 备份, 服务器/边栏选项卡(可堆叠), 存储(Azure), 存储队列 | 338 |
| BASIC_AUDITDIAGRAM_M | 终止符, 标记的文档, 标记的流程, 数据传输, 决策, 动态连接线, 圆角矩形, 事件 | 47 |
| BASIC_GANTTCHART_M | 甘特图框架, 圆角矩形, 列, 次要比例单元格, 主要比例单元格, 非工作时间, 文本输入, 行, 任务栏, 链接线, 水平标注, 动态连接线.51 | 141 |
| BASIC_HOME_NETWORK_M | 无线访问点, 多功能设备, PC, 路由器, 动态连接线, LCD 监视器, 笔记本电脑, 智能手机, 普通.31, 圆角矩形 | 68 |
| BASIC_PESTLE_ANALYSIS_M | 圆角矩形.1000, PESTLE 分析 | 28 |

**怎么用**：

```powershell
# 方式 A：新建文档时直接套模板（COM）
$doc = $app.Documents.AddEx('C:\...\AFFINITY_DIAGRAM_M.VSTX')

# 方式 B：拿模板里的模具库自己画（模板与模具库是同一批 .vssx）
$st = Open-VisioStencil -App $app -Names @('BASFLO_M')
```

**什么时候用模板**：需要「现成的图表骨架」（甘特、亲和图、审计流程），且不打算精细改版式。
**什么时候不用**：期刊图要精确控制字号、颜色、留白——模板的固定版式反而碍事，
更好的做法是**参考模板的语义结构，但用 BASIC_M 自绘**。

完整模板清单见 `references/data/visio-templates.tsv`。

## 8. 模具库地图（全 362 个）

> 数量 = 该库的 master 个数；完整形状名见 `references/data/visio-stencil-summary.md`。

### 云与平台图标（AWS / Azure / …）（73 个）

AWSANALYTICS_M（28） · AWSAPPINT_M（18） · AWSARVR_M（2） · AWSBLK_M（4） · AWSBUSINESSAPP_M（4） · AWSCOMPUTE_M（71） · AWSCONTAINERS_M（18） · AWSCOSTMGMT_M（6） · AWSCUSTENB_M（5） · AWSCUSTENG_M（5） · AWSDB_M（40） · AWSDEVTOOLS_M（13） · AWSENDUSERCOMPUTE_M（5） · AWSGAMETECH_M（2） · AWSGEN_M（22） · AWSIOT_M（60） · AWSMEDIASERV_M（9） · AWSMGMTGOV_M（50） · AWSMIGRATIONTRANSFER_M（14） · AWSML_M（29） · AWSMOBILE_M（7） · AWSNETCONTENTDELIVERY_M（36） · AWSQUANTTECH_M（2） · AWSROBOTICS_M（6） · AWSSAT_M（2） · AWSSECURITYCOMPLIANCE_M（36） · AWSSTORAGE_M（35） · AZUREAIMACHINELEARNING_M（4） · AZUREANALYTICS_M（10） · AZUREAPPSERVICES_M（10） · AZUREBLOCKCHAIN_M（1） · AZURECLOUD_M（205） · AZURECOMMAND_M（168） · AZURECOMPUTE_M（21） · AZURECONTAINERS_M（6） · AZUREDATABASES_M（19） · AZUREDEPRECATED_M（45） · AZUREDEVOPS_M（5） · AZUREDRAWING_M（35） · AZUREECOSYSTEM_M（3） · AZUREENTERPRISE_M（158） · AZUREGENERALSYMBOLS_M（46） · AZUREGENERAL_M（83） · AZUREIDENTITY_M（8） · AZUREINTEGRATION_M（11） · AZUREINTUNE_M（18） · AZUREIOT_M（9） · AZUREMANAGEMENTGOVERNANCE_M（21） · AZUREMENU1_M（157） · AZUREMENU2_M（157） · AZUREMENU3_M（149） · AZUREMIGRATE_M（3） · AZUREMIXEDREALITY_M（2） · AZUREMONITOR_M（1） · AZURENETWORKING_M（25） · AZURENEWICONS_M（48） · AZUREOMS_M（21） · AZUREOPERATIONSMANAGER_M（22） · AZUREOTHERMICROSOFTPRODUCTS_M（25） · AZUREOTHER_M（15） · AZUREPREVIEW_M（9） · AZURESECURITY_M（5） · AZURESTACK_M（8） · AZURESTATUS_M（33） · AZURESTORAGE_M（14） · AZURESYSTEMCENTER_M（10） · AZUREVMSBYFUNCTION_M（17） · AZUREVMWARESOLUTION_M（1） · AZUREWEB_M（7） · DYNAMICS365ICONS_M（39） · KUBERNETESVISIOSTENCIL_M（40） · MICROSOFTPRODUCTS_M（27） · POWERPLATFORMICONS_M（10）

### 其他 / 工具与杂项（71 个）

ACCESSBILITY_NEW_M（11） · APPL_M（21） · APPL_VISIO2013_M（24） · BSTORM_M（5） · BTHKT_M（27） · BTHKT_VISIO2013_M（22） · BUGS_M（14） · COMOLE_M（16） · COMPME_M（12） · COMPS_M（10） · DGICON_M（14） · DRILLD_M（5） · ENTAPP_M（19） · FASTN1_M（36） · FASTN2_M（27） · GANESA_M（4） · GENERAL_M（2） · INTANN_M（9） · LANGLV_M（12） · LOGIC_M（26） · MAP3D_M（39） · MEMOBJ_M（14） · METRO_M（8） · MRKTDG_M（30） · OBJREL_M（10） · OFFACC_M（29） · OFFACC_VISIO2013_M（11） · PRKRD_M（25） · PRKRD_VISIO2013_M（20） · PTSCLD_M（10） · PTUSCL_M（46） · RECRT_M（24） · RECRT_VISIO2013_M（19） · REGSTR_M（6） · ROAD_M（14） · SECURITYJUSTICE_M（12） · SHPMCH_M（25） · SHPREC_M（19） · SHPSTR_M（27） · SIGNS_M（16） · SSHOQ_M（14） · TECHELEC_M（27） · TOOLSBUILDING_M（15） · TRANSP_M（29） · VEHICLES_M（14） · VEHICL_M（28） · VEHICL_VISIO2013_M（27） · VIDSUR_M（20） · VSM_M（30） · WADOWI_M（15） · WEBSIT_M（19） · WFCRS_M（14） · WFCTRL_M（16） · WFDEME_M（43） · WFDEP_M（43） · WFDLGS_M（18） · WFOBJ_M（26） · WFOBME_M（25） · WFSTEP_M（15） · WFSTME_M（22） · WFTLBR_M（19） · WFWICN_M（36） · WWFCONCTRL_M（15） · WWFCRS_M（26） · WWFCTRL_M（20） · WWFMECTRL_M（12） · WWFNAVCTRL_M（12） · WWFTEXTCTRL_M（17） · WWFWICN_M（26） · XFUNC_M（7） · YOURDON_COAD_NOTATION_M（4）

### 工程 / 电气 / 管道设备（61 个）

DIMENG_M（26） · EECHIP_M（35） · EECOMP_M（28） · EEFUND_M（57） · EEFUND_VISIO2013_M（50） · EEMAIN_M（13） · EEMAPS_M（36） · EEMECH_M（19） · EEPATH_M（25） · EEPATH_VISIO2013_M（25） · EEQUAL_M（25） · EEQUAL_VISIO2013_M（25） · EESEMI_M（30） · EESEMI_VISIO2013_M（27） · EESWCH_M（50） · EESWCH_VISIO2013_M（50） · EETCOM_M（28） · EETERM_M（28） · EETRAN_M（21） · EEVHF_M（32） · FPASSM_M（52） · FPEQP_M（35） · FPVALV_M（21） · FREEQP_M（9） · IRRIG_M（11） · LINPAT_M（25） · OFFEQP_M（25） · OFFEQP_VISIO2013_M（24） · OFFFRN_M（29） · OFFFRN_VISIO2013_M（35） · PEANNT_M（22） · PEEQP_M（40） · PEHEAT_M（28） · PEPIPE_M（6） · PEPUMP_M（26） · PERIME_M（32） · PERIPH_M（29） · PEVALV_M（50） · PEVESS_M（23） · PIPE1_M（42） · PIPE2_M（28） · PLUMB_M（22） · RCKEQP_M（19） · RCKSME_M（17） · RCKSVR_M（17） · RECSHP_M（28） · SDCALL_M（14） · SDCONT_M（14） · SITACC_M（35） · SITACC_VISIO2013_M（26） · UACTME_M（15） · UCMNME_M（5） · UCOMME_M（11） · UDEPME_M（21） · USEQME_M（12） · USTAME_M（8） · USTRME_M（16） · UUSEME_M（8） · VALVE1_M（21） · VALVE2_M（17） · WELD_M（27）

### 网络与 IT 设备（27 个）

ADO_M（21） · ADS_M（15） · CISCONETWORKSHAPES_M（292） · DTLNET_M（16） · DTLNME_M（15） · DVCALL_M（57） · ELETEL_M（31） · EXCOBJ_M（17） · LDAPOB_M（21） · MWFCONCTRL_M（22） · MWFCTRL_M（20） · MWFECTRL_M（13） · MWFNAVCTRL_M（15） · MWFTEXTCTRL_M（19） · NETLME_M（8） · NETLOC_M（8） · NETRM_M（4） · NETSME_M（24） · NETSYM_M（24） · PEINST_M（69） · PTSINT_M（75） · PTSINT_VISIO2013_M（53） · SERVER_M（17） · SERVME_M（17） · WEBMAP_M（79） · WFCICN_M（22） · WWFCICN_M（79）

### 建筑 / 空间 / 暖通（23 个）

BLDCOR_M（26） · BLDCOR_VISIO2013_M（25） · BUILDINGS_M（12） · CABNT_M（25） · CABNT_VISIO2013_M（27） · CUBICL_M（16） · CUBICL_VISIO2013_M（18） · FURN_M（33） · FURN_VISIO2013_M（25） · GARDEN_M（42） · GARDEN_VISIO2013_M（17） · HVACCE_M（32） · HVACC_M（24） · HVACD_M（20） · HVACEQ_M（19） · INDOOR_M（58） · LANDSCAPE_M（24） · LNDMRK_M（32） · LOCATION_M（14） · OUTDOOR_M（28） · PLANT_M（25） · PLANT_VISIO2013_M（30） · WALL_M（41）

### 基础形状 / 箭头 / 图例 / 底纹（18 个）

ARROWS2_M（16） · ARROWS_M（28） · BASICORGCHART_M（6） · BASIC_M（58） · BCKGRN_M（11） · BLOCK3_M（19） · BLOCKP_M（19） · BLOCK_M（41） · BORDRS_M（15） · CYCLEDIAG_M（20） · DRAWTL_M（31） · EMBELL_M（41） · GDT_M（37） · LEGEND_M（6） · LGND_M（24） · SYMBOL_M（34） · TITLE_M（29） · Theme_m（13）

### 流程与工作流（5 种记法）（17 个）

BASFLO_M（15） · BPMN_M（21） · EPC_M（12） · FLOSHP_M（56） · IDEF0_M（9） · MSFLOW_M（2） · PROCESS_M（31） · PROCESS_NEW_M（19） · SDL_M（26） · SPWA2_M（43） · SPWA_M（37） · SPWC2_M（9） · SPWC_M（13） · SPWT2_M（12） · SPWT_M（2） · SSFLOW_M（15） · WRKFLO_M（42）

### 业务框架 / 分析方法（16 个）

ALARM_M（15） · AUDIT_M（31） · BA_DEC_BAN_M（26） · BA_DEC_M（17） · BA_GRP_M（10） · BPRES_M（9） · BUSINESS_FRAMEWORKS_M（11） · BUSINESS_M（18） · CAUSEF_M（12） · COMPLN_M（7） · ENTITY_M（9） · EVIDENCE_M（40） · FAULT_M（12） · INTERFACE_M（20） · ITIL_M（11） · TQM_M（52）

### 行业与生活图标（16 个）

ANIMALS_M（22） · APPAREL_M（20） · ARTS_M（18） · CELEBRATION_M（15） · COMMERCE_M（16） · EDUCATION_M（22） · FOODDRINKS_M（29） · HOLIDAYS_M（54） · HOME_M（24） · MEDICAL_M（11） · NATUREOUTDOORS_M（28） · PEOPLE_M（25） · PUZZLESANDGAMES_M（20） · SPORTS_M（28） · WEATHERSEASONS_M（16） · pictogram_m（24）

### 组织结构与人员（14 个）

BODYPARTS_M（18） · FACES_M（20） · ORGBLT_M（15） · ORGBND_M（15） · ORGCH_M（14） · ORGCNS_M（15） · ORGNCH_M（15） · ORGPAN_M（15） · ORGPER_M（15） · ORGPET_M（15） · ORGPIP_M（15） · ORGSTO_M（15） · ORGTAC_M（15） · orgico_m（10）

### 数据可视化 / 矩阵 / 图表（10 个）

ANALYTICS_M（20） · CHART_M（40） · DATAVISUALIZERPREVIEW_M（3） · DATAVISUALIZER_M（5） · MATRIXDIAGRAMSHAPES_M（15） · MATRIXDOT_M（37） · MATRIXSTICKYNOTES_M（12） · SPRBR_M（27） · statistical_infographic_m（18） · step_infographic_shapes_m（24）

### 标注 / 注释 / 引出（5 个）

ANNOT_M（39） · CALOUT_M（38） · COMMUNICATION_M（22） · CONNEC_M（69） · DATFLO_M（21）

### 时间线 / 甘特 / PERT（5 个）

CALNDR_M（40） · GANTT_M（18） · PERT_M（9） · TIMELN_M（27） · timelinetodolist_m（8）

### 数据库与实体关系（4 个）

DBCHEN_M（10） · DBCROW_M（6） · DBIDEF1X_M（5） · DBUML_M（5）

### 集合与分层（维恩 / 金字塔）（2 个）

PYRAMID_SHAPES_M（11） · VENNDIAG_M（12）

## 9. 选型层面的已知坑

完整坑表见 `quick-lookup.md` 第 3 节。本节只留两条**模具库特有**的：

| 坑 | 表现 | 正确做法 |
|---|---|---|
| 隐藏 master | 界面里搜不到，但能 `Item()` 出来 | 扫描结果里 `Hidden=1` 的即此类，可作「藏起来的宝藏」 |
| 母版按 mm 设计 | 改尺寸后比例怪 | 母版宽高用 `Width`/`Height` 单元格直接给英寸值，Visio 内部会换算 |

## 10. 绘图时怎么快速查

```powershell
# 加载库
. ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 'C:\Users\thunder\.agents\skills\visio-diagram\scripts\visio-lib.ps1')))

# ① 全库检索（走缓存清单，毫秒级）
Find-VisioMaster -Pattern '判定','菱形','时间线' | Format-Table -AutoSize

# ② 直接拿到 master 对象去画（自带中文名查找）
$r = Get-VisioMasterByName -Stencil 'BASFLO_M' -Name '判定'
$s = $page.Drop($r.Master, 3.0, 3.0)   # 注意：期间不要 Close($r.Doc)

# ③ 不想等就 grep 全量清单
Select-String -Path '...\references\data\visio-masters-full.tsv' -Pattern '时间线'
```

**建议顺序**：先查 `quick-lookup.md` → 再 `Find-VisioMaster` → 仍不确定就 `scripts/shape-sheet.ps1` 打样张看一眼真实外观。

