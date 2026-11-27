# 工作簿规范（7 表，v1x）

文件命名：`<产品>_ShouldCost_Analysis.xlsx`　工作表名固定为下列 7 个（公式按表名引用，勿改名称）。

| # | 名称 | 作用 |
|---|---|---|
| 1 | `1-Assumptions` | **唯一参数源**：口径声明、年份、币种、产量、材料单价、VA 比例 k、隐含件比例 s |
| 2 | `2-BOM` | 逐件定价（MAT）：数量、单价、区间、价格方法、计入判定 |
| 3 | `3-ProcessCost` | VA 合计、可得分块、方法层级 |
| 4 | `4-Summary` | 第一行结论句、口径块、MAT/VA、功能域透视、方法占比、对账、待取证清单、`PPT_EXPORT` |
| 5 | `5-Opportunities` | 毛/Effort/净分列、MAT/VA/Tooling 归属、防重复计数组、HG |
| 6 | `6-Sensitivity` | 单因子三档、待取证排序 |
| 7 | `7-Evidence` | 数据缺口、待核实事项、版本说明、M5 清单 |

> 本规范随 `SKILL.md` Step 10 同步。与 `should-cost-engineering`（v1 原件）的 43 列 `2-BOM` **不可互读公式**；与 `should-cost-engineering-v2` 的 JSON 驱动模型也不可混读。

## 0. 全局约束

- **参数唯一源**：所有费率、比例、k、s、阈值只出现在 `1-Assumptions`，其他表一律引用单元格，不写死数字。
- 空值必须为空字符串 `""`，**不得写成 0**：`IF(ISBLANK(x),"",...)`。无价件写成 0 会把"缺价"伪装成"零成本"。
- 每个透视/汇总块**必须配一行闭合差公式**（分项之和 − 合计），值域恒为 0。
- 说明性文案里的数字**必须来自生成器变量**，不得手写（手写数字会与公式脱节）。

## 1-Assumptions

列：`参数ID | 参数 | 值 | 单位 | 状态 | 来源/说明 | 备注`

至少包含：

| 组 | 参数 |
|---|---|
| 口径声明 | 报告日期 · 价格基期 · 目标年份 · 情景（竞品默认"现状"/自家默认"含全部措施"）· 币种 · 来源版本 · **边界归一开关状态**（默认关） |
| 产量 | 年产量、生命周期产量、产量来源（实测/同类样本假设值须标 `[假设]`） |
| 材料单价 | 各材料单价及其来源与日期 |
| VA | VA 比例 k（Low/Base/High）与其方法层级（L0/L1/L2/L2x） |
| 隐含件 | 各类隐含件比例 s（Low/High），仅用于定尺寸 |
| 阈值 | M5 占比分档阈值（15% / 40%，**经验设定待校准**）、`RANGE_MAX_RATIO`（1.5）、`[宽>10×]` 阈值（10） |
| 加价 | 供应商 OH / SG&A / 物流 / 利润率（**仅用于零件级 M2 计算器**，不得整机套固定费率） |

## 2-BOM（逐件定价表）

数据自第 5 行起，冻结 `D5`。

| 组 | 列 |
|---|---|
| 层级与归属 | `Level` · `Row Type` · `Function` · `Category` · `Make/Buy`（**默认 Buy**）· `Cost Bucket` |
| 件与数量 | `Part No` · `Name / Model` · `Qty per Unit` · `Quantity Basis` · `Quantity Evidence` |
| **价格** | `Unit Price (= Base)` · `Price Low` · `Price High` · `Price Method`（**M1–M5**）· `Unit Price Basis` · `Price Evidence` · `Fit Check` · `Price Date` · `Purchase Qty` · `Delivery Term` · `Tooling In Price` |
| **计入判定** | `Counted In Total`（Y/N）· `Counted Amount` · `Counted Note` |

- `Row Type` ∈ {汇总件, 参考件, 展开项, **预留行**}。用 `Level` 路径表达层级，不用 `Depth`。
- `Cost Bucket` ∈ {MAT-零部件, MAT-Chem, MAT-Scrap, MAT-外协, MAT-自制并入, 预留}。
- `Unit Price` 是本行**采用的单价**（= Base）；`Price Low` / `Price High` 是**区间**，任何情况下都要保留，不因取了 Base 而丢弃。
- `Counted Amount = IF(Counted In Total="Y", Unit Price × Qty per Unit, "")`。**所有汇总只对 `Counted Amount` 求和**。
- 父件按总成价计入时，子件仍显示参考估价、`Counted In Total = N`、`Counted Note = 已含在 <父行>`。子件之和与父件价不一致时**只列差额，不调平**。
- 同一件**不得**同时按总成价与展开成本计入。
- 预留行：`Counted In Total = Y` 但 `Unit Price` 由 Step 3 的 `Reserve_c` 公式给出；**不计入"已出价行数"**。
- 未定价行：`Unit Price` 留空，`Unit Price Basis` 写清缺什么证据。
- **`Fit Check` 必须能区分四种情形**，这是判断这一行的价"是谁给的"的唯一可见位置：
  `取证` / `取证·未归一` / `子档命中` / `父档命中` / `父档回退（属性缺失）` / `质量轴（件名未命中）` / `材料` / `—`。
- **属性先于价格**：目录行的 `Required Attribute` 必须由属性表的附加列喂入匹配。
  属性**未知时判"不匹配"**，不得当默认符合；退到父档时 `Unit Price Basis` 必须以 `[假设]` 开头并写明缺哪个属性。
- **`[假设]` 是强制前缀**，凡"看不出来的属性取了默认值"（PCB 层数/铜厚、螺纹规格、密封等级…）或质量轴兜底推断，都必须在 `Unit Price Basis` 最前面标出，不得与实测属性混写。
- 取证价缺采购数量 / 阶梯价时，`Unit Price Basis` 必须标"**未完成产量归一**"，不得默认它就是成交价。
- **EE 类别件不影子计价**：`Material Group` 为 EE 类别（EE Components/PCB/Connector/Wire Harness/Sensor…）的件，**即使属性表给了材料分层也绝不走 M2**（材料影子）。价值在功能与供应链，不在原料 kg；分层仅作结构对照。EE 件只能走 M1 / M3–M5 目录，否则未定价，`Unit Price Basis` 须点出"EE 类别件不影子计价"。
- 清单同时出现父件价与其子件明细时，生成器**不自动改 `Row Type`**（会改变交付形态），但必须在表尾显式告警并点名，由人工把其中一侧 `Counted In Total` 置 N。

## 3-ProcessCost（VA）

列：`VA 项 | 分块 | 金额 Low | 金额 Base | 金额 High | 方法层级 | Site | 依据 | 来源` 

- `VA 项` 至少覆盖 PDC（SMT / Soldering / Coating / FA+Testing）、PDOH `[待确认]`、PAC、MOH（LOG / 采购 / PUR tooling / Duty-Freight）。
- **VA 合计必须给 Low/Base/High 并标方法层级**（L0 自下而上 / L1 RB 基准乘因子 / L2 同类样本 `VA = MAT × k` / L2x 跨类）。
- **分块（PDC / PDOH+PAC / MOH）只在有比例依据时给**；只有合并值时不强拆。
- `Site` ∈ {in-house, outsource}；外协 OSAT 计入 **MAT**，不重复进 VA。
- **MOH 已在 VA 内**：MOH/MAT 只用于 VA 内部拆分，**不得在 VA 之上再叠一层 MOH**。
- 设备折旧与工装摊销各只计**一次**：`投资 ÷ 年限 ÷ 产量`；件价 / 机时费 / MOH 已含的不再添加。

## 4-Summary

**第一行必须是结论句**（逐字模板）：

```
> 清单 X 个实物行,已出价 Y,待补 Z;估算 ___ 元/台(区间 ___–___),其中 M5 与预留占 ___%。
```

- 仅当 Z = 0 时才允许显示"逐件已出价"，且**必须同时显示 M5 占比**。"已出价"不等于"可靠"。

块顺序：

1. **口径块**：报告日期 / 价格基期 / 目标年份 · 情景 · 币种 · 来源版本 · 边界归一开关。
2. **MAT / VA / PPC**：`PPC = MAT + VA`。
3. **两个 MAT 口径各带区间**：
   - **已锚定 MAT** = M1–M4 行合计（Low 相加、High 相加）
   - **全口径 MAT** = 已锚定 + M5 + 预留
   - **PPC 标题数用全口径**，同时把已锚定值并列写出。
   - 合计区间声明"假设所有件同向偏离，偏保守"。
4. **功能域透视**：Function × Category 聚合 MAT / VA / Direct / Share%。功能域之和 = MAT；明细之和 = 小计。
5. **方法占比分行披露**（**不得合并成一个数**）：M1+M2 / M3 / M4 / M5 / 预留。非锚定占比 = (M5 + 预留) ÷ 全口径 MAT。
6. **对账结果**：底层保留精度严格对账；取整差异单列（n 个取整明细误差上限 n×0.5，超过不得用舍入解释）。
7. **待取证清单**：按 `(High − Low) × Qty` 降序。
8. **`PPT_EXPORT` 固定块**（见下）。

### `PPT_EXPORT`

- 跟随结论等级导出：**等级 2 以下不导出 Top-N 单件排序**。
- VA 按可得粒度导出，**只有合并值时不强拆**。
- 必须带 `completeness`（`complete` / `partial`）、M5 占比、口径声明、价格基期与目标年份。
- `completeness = partial` 时，图表标题必须含"**不完整小计**"并附未定价清单。
- 除 `completeness` 外必须同时给出 `subtotal_label`：一个**已成品的标题字符串**，
  `complete` 时为 `完整小计`，`partial` 时为 `不完整小计（缺价 n 件未计入）`，
  由公式生成。禁止在幻灯片里手写这个标题，否则件数会与公式脱节。

## 5-Opportunities

列：`ID | 题目 | 现有方案 | 替代方案 | 工程依据 | 毛潜力 | Effort | 净潜力 | MAT Impact | VA Impact | Tooling Impact | 已在 DtC | 已在 RPP | Phase 增量 | 降本机制 | 防重复计数组 | HG | 证据可信度 | 风险 | 验证周期 | 状态 | 优先级`

- `净潜力 [CNY/pc] = 毛潜力 − Effort`，**仅当两者均为 CNY/pc**；单位不一致时不给净潜力。
- 工装（kCNY / mCNY）**另列，不折算进单件**；一次性投入另列，必要时另给摊销与回收期。
- `MAT Impact` 列只接受 `MAT` / `无` / 空；`VA Impact` 列只接受 `VA` / `无` / 空；`Tooling Impact` 列只接受 `Tooling` / `无` / 空。
- **同一成本项目只计一次**：一项设计变化可同时影响 MAT 与 VA，按成本项目分别计入；工时类节省须标注"是否形成实际节省"。
- 同 `防重复计数组` 的措施**不得重复计节省**；跨组叠加前先声明是否可叠加及依据。
- `HG`（0 想法 / 1 毛潜力 / 2 净潜力 / 3 时间表已排 / 4 已实施 / 5 节省已兑现）只记**实施成熟度**；证据可信度（Verified / Modeled / Hypothetical）**另记一列**。HG4/HG5 只由实施状态决定。

## 6-Sensitivity

列：`驱动因素 | 基准值 | Low | Base | High | 影响额(绝对值) | 影响占比 | 方向 | 说明 | 来源参数`

- 只对关键驱动做 Low/Base/High：材料单价、年产量、VA 比例 k、PCB 层数与铜厚假设、M5 区间端点、供应商 SG&A%。**禁止"总成本 ±10%"式假敏感性**。
- **影响额取绝对值**，另用 `方向` 列写清"因子↑→成本↑/↓"。取绝对值后**不得再用符号判断增减**。
- 敏感度表的"影响额"若同行方向不一致，必须取绝对值 + 单列方向说明，否则排序失效。
- 每一项必须能追到 `1-Assumptions` 里的参数；跨表引用**必须从源表的 `make_cols()` 取列字母**，不得假设列位置相同。

## 7-Evidence

列：`项 | 内容 | 来源/位置 | 状态 | 置信度 | 影响成本项 | 待确认问题 | 负责人`

- 来源写到可复核粒度：报告名 + 页码 / 照片文件名 + 拍摄面。
- 必须含：**M5 清单**及其声明"**不得用于定点谈判**"、数据缺口、待核实事项、版本说明。
- 证据状态：Verified / Estimated / Unknown，**竞品成本最高 Estimated**；计算状态 F（公式可复算）/ M（手输），**F 不代表已验证**。
- 待确认问题写"缺什么证据能回答"，不写"待补充"。

## 格式与颜色

| 状态 | 填充 | 用途 |
|---|---|---|
| Verified | 无填充 | 原始资料 |
| Calculated | 浅蓝 `DDEBF7` | 公式结果 |
| Estimated | 浅黄 `FFF2CC` | 工程假设 |
| Unknown | 浅红 `FCE4E4` | 缺失，留空 |

- 表头深蓝底白字；冻结首行；千分位；货币列 2 位小数，质量列 3 位。
- 单位写在列名里，不写在单元格值里。
- **文案以 `=` 开头时必须转全角 `＝`**，否则 Excel 判定为公式并拒开。

## 生成后必查

1. 无 `#REF!` / `#VALUE!` / `#DIV/0!`；无循环引用。
2. **无空白价格行**：未定价行 `Unit Price` 为空且有原因；总额标"**不完整小计**"。
3. 父子件：子件有参考价且 `Counted In Total = N`，汇总未重复；子件之和与父件价的差额已单列。
4. `Counted Amount` 之和 = MAT 分桶之和 = 功能域透视合计（闭合差行恒为 0）。
5. 螺钉逐件估价（不是总量估）；预留行单列且**未计入已出价行数**。
6. M5 中 `High/Low > 10×` 的已先拆子档，拆不开的标 `[宽>10×]` 且进待取证清单前 5。
7. 前 10 大件为 M1–M4，或已标注并进待取证清单。
8. VA 有合计 Low/Base/High 与方法层级；L2 未叠加 MOH；无比例依据时未强拆分块。
9. 已锚定 / 全口径两个 MAT 与 PPC 均带区间；方法占比**分行**披露（M1+M2 / M3 / M4 / M5 / 预留）。
10. 工装只在 MOH 的 PUR tooling 计一次；件价含工装已扣减；设备折旧单次计入。
11. 套用价格点已注明来源型号与适用前提；**芯片价未跨型号套用**（TC375/TC377/TC387/TC389 各按各的价）。
12. `Typical` 只用了带 `Typical Source` 的值。目录里 `Typical` 若来自 `Calibrated` 行且其 `Basis Unit` 不是 `CNY/piece`（如 PCB 为 `CNY/m2`），**必须先按基数换算再当单价用**。
13. 降本毛 / Effort / 净分列；一次性投入另列；HG 与证据可信度分列。
14. 待确认缩写（PDOH、RPP、RBHP、EVOB、PMoC、PM5、HCi、PCi、HSP、CIPG、EAC、DCL）**原样保留**。
15. `4-Summary` 第一行结论句句式为 Step 6 逐字模板，且 Y/Z/M5 占比来自动态变量。
16. **产品无关性回归**：换一个总成跑同一份清单，功能域分类不变，只有 Function 前缀变。
17. 交付前用 Excel COM `CalculateFullRebuild` + 错误值扫描验收；openpyxl 能读 ≠ Excel 能开。
18. 公式长度逐条 < 8192 字符（超长会被 Excel 判定文件损坏）。
