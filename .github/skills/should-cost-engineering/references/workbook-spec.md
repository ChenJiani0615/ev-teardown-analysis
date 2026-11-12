# 工作簿规范（7 表）

文件命名：`<产品>_ShouldCost_Analysis.xlsx`　工作表名固定为下列 7 个（公式引用依赖，勿改中文/空格）。

| # | 名称 | 作用 |
|---|---|---|
| 1 | `1-Assumptions` | 唯一参数源：地区、币种、产量、费率、单价、良率、间接费、利润率 |
| 2 | `2-BOM` | 工程 BOM + 材料成本计算 + 工艺路线 + Make/Buy |
| 3 | `3-ProcessCost` | 逐工序成本 |
| 4 | `4-Summary` | 汇总与排序、成本占比、与报价对比 |
| 5 | `5-Opportunities` | 降本机会与量化 |
| 6 | `6-Sensitivity` | 关键驱动敏感性 |
| 7 | `7-Evidence` | 来源、假设、待确认问题 |

## 1-Assumptions

列：`参数ID | 参数 | 值 | 单位 | 状态 | 来源/说明 | 备注`

规则：**全簿唯一的常数存放处**。其他表的公式只能引用本表单元格（`'1-Assumptions'!$C$5` 这类），不得写死数字。

至少包含：币种、生产地区、年产量、生命周期产量、各材料单价、各设备费率、人工费率、测试费率、材料利用率、废料回收率、工序良率、制造间接费率、SG&A、物流、利润率、汇率。

## 2-BOM

**43 列**（2026-11-04 与 `tools/build_should_cost_workbook_v2.py` 的 `make_cols()` 对齐；列名是公式引用锚点，勿改）。按功能分组：

| 组 | 列 |
|---|---|
| 标识 | `Level · Part No · Name · Function · Category · Qty` |
| 材料属性 | `Material Group · Material Species · Net Weight (kg)` |
| **9 类材料分层** | `AL (kg) · CU (kg) · STEEL (kg) · PLASTIC (kg) · ELASTOMER (kg) · FERRITE (kg) · FOAM (kg) · ADHESIVE (kg) · OTHER (kg)` |
| 材料成本 | `Material Input (CNY)` |
| 工艺 | `Procurement Object · Route · Route Status · Make/Buy · Bottom-up Model (Y/N) · Supplier Process (CNY) · Supplier Mfg Cost (CNY) · Supplier Selling Price (CNY) · Route Cost Basis` |
| 价格取证 | `Price Source · Price Low (CNY) · Price High (CNY) · Price Base (CNY) · Price Class · Price Type · Price Status · Cost Boundary · Price Basis` |
| PCB 专项 | `Bare PCB Proxy (CNY) · Bare PCB Basis` |
| 工装 | `Tooling CapEx (kCNY)` |
| 金额边界 | `Amt Low · Amt High · Amt Base` |

- `Function` / `Category`：功能域两级切割（如 INV→PM 功率级 / DC-Link 电容 / PCB / E-Parts / 功率磁件 / 电流传感；Box→壳体 / 冷却与热 / HV 连接 / 信号连接 / EMC / 密封紧固其他）。**冷却与 TIM 归 Box，不归功率级**；E-Parts（保险/呼吸阀/端子）独立于 PCB。映射由 `tools/domain_classify.py` 的 `function_category()` 判定。**判定看件名主体（`name_subject()` 取逗号前的部件类型词），不看整串**——CS 件名格式为 `<部件类型>, <父件上下文>`，按整串匹配会让父件上下文盖掉部件类型。实测三个反例：`Gasket, … Power Module` 被判功率级、`Ferrite Core, … Current Sensor 2` 被判电流传感、`PCB, … Current Sensor 2` 被判电流传感。只有两类例外允许看全文：
  - **冷却**：`Adapter, … Coolant Inlet/Outlet` 主体是 adapter，但功能上就是冷却接口 → 归 `Box/冷却与热`
  - 其余一律只看主体。**电流传感必须主体是 sensor**，`Ferrite Core … Current Sensor` 归 `Box/EMC`、`PCB … Current Sensor` 归 `PCB / PCBA`
- `Function` 为 `Box` 的件**不带产品前缀**（壳体/冷却/连接器属结构功能域）；只有电气功能域才随产品切换前缀（INV/DCDC/OBC/PDU/BMS，识别不出落 `GEN`）
- **9 类材料分层**：自制件按 `Material Species` 牌号拆分；`Σ9 类 = Net Weight` 必须闭合。**外购件不拆分层**（供应商的材料构成不可见，拆出来是编的），列留空。
- 材料三式（仅自制件）：`Gross = Net / Yield`；`Scrap = (Gross − Net) × Price × 回收率`；`Material Cost = Gross × Price − Scrap`
- 外购件：`Material Input` 留空，只填价格取证组；**`Make/Buy = Purchased` 且 `Bottom-up Model = N`**
- 工艺组三列联动规则：
  - `Supplier Process (CNY)` = 该件**最后一道工序**的 `SUMIFS`（不是全部工序相加，避免重复计）
  - `Supplier Mfg Cost (CNY)` = `N(Material Input) + N(Supplier Process)`
  - `Supplier Selling Price (CNY)` = `Mfg Cost × (1+OH) × (1+SG&A) × (1+物流) ÷ (1−利润率)`
  - `Bottom-up Model (Y/N)`：`Y` 时金额取 `Supplier Selling Price`；`N` 时取价格取证组（High/Low ≤ 1.5× 取 `AVERAGE`，否则取 `GEOMEAN`）
  - `Route Status` ∈ {`Assumed`（有 attrs 证据）, `Inferred`（件名推断）, `Inferred-Weak`（只判到装配级）, `Unknown`（判不出，留空）, `N/A`（不建工艺）}
  - `Route Cost Basis`：写清钱是怎么算出来的；`Inferred` 必须写 `材料分层缺失，Supplier Mfg Cost 仅含加工费`
- **价格取证列**（外购件联网取证后填）：`Price Low` / `Price High`（窄区间 High/Low ≤ 1.5× 时保留）/ `Price Base`（宽区间取均值或几何均值的点值）/ `Price Class`（A 公开源 / B 分销商 / C 外推 / D 先前估计 / **E 基准价目录补位**）/ `Price Basis`（来源 + 日期 + 换算方式的依据文本）
- **价格来源四优先级**（严格自上而下，命中即停）：① `--prices` 取证价（A/B/C/D）→ ② bottom-up 售价链 → ③ 基准价目录严格类别（E・`Comparable`）
  → ④ 基准价目录包络 / 质量轴兜底（E・`Assumption`）。**不允许留空**：③④ 都落不到只允许出现于"既无类别又无净重"行，且该类行数目标为 0。
  基准目录为 `data/benchmark/benchmark-price-catalog.csv`，按**件类别**而非件号建行，含 `Low / High / 计价基数 / 依据 / 来源 / 日期 / Anchor Kind`。
  **`Anchor Kind` 两取值**：`Benchmarked`（有实证锚点 → 走 3a） / `Envelope`（无锚点，给行业最宽可信包络 → 走 3b），
  另有不依赖目录的 3c 质量轴兜底（`净重 × 材料单价 × 加工倍率带`，适用于无 bottom-up 的金属/塑料/弹性体/泡棉结构件）。
  Class E 行必须同时具备：`Price Class = E`、`Price Type = Benchmark Average`、
  `Price Status`（3a `Comparable` / 3b·3c `Assumption`）、
  `Cost Boundary = Benchmark range, not a quotation`、`Price Basis` 内可独立阅读的**层级标记 + 类别 + 基数 + 来源 + 日期 + 取值规则**（层级标记：`[包络]` / `[兜底-质量轴]`，3a 无标记）。
  且 **Low/High 必须成对给出**（不得把均值直接写进 `Price Base`），`Verified Price Coverage` 不得因补位上升。
- **价格状态四列**（外购件必填、自制件留空）：`Price Type`（如 Unit Price / Bottom-up Estimate / **Benchmark Average**）/ `Price Status`（**八类之一**：Actual Procurement / Supplier Quotation / Comparable / Bottom-up Estimate / Parametric Estimate / Assumption / **Benchmark Average** / Unknown；**默认 Unknown**）/ `Cost Boundary`（如 Delivered at assembly plant 到厂价，含包装运输；基准价行写 `Benchmark range, not a quotation`）/ `Quantity Basis`（如 per piece × Qty）。语义约束：Supplier Quotation 是报价不是已发生成本，不得标 Verified；Assumption 值只用于敏感性，不得进 PPC 基线；**Benchmark Average 不得进 Verified Saving 基线、不得用于与报价的差额结论**
- `Tooling CapEx (kCNY)`：工装一次性投入（模具/夹具/检具），与摊销分开记两本账；摊销公式在 Summary（÷ MIN(工装寿命模次, 生命周期产量)）
- `Status` 与 `Source` 每行必填，不允许整表默认 Verified

## 3-ProcessCost

列（18 列，2026-11-04 与 `make_cols()` 对齐）：`Step | Part No | Process | Category | Cost Bearer | Site | Cycle (s) | Cycle Basis | Machine Rate | Rate Basis | Operators | Labor Rate | Labor Cost | Consumable | Step Yield | Step Direct (CNY) | Cumulative Cost (CNY) | Is Last Step (Y/N)`

- `Category` ∈ {加工, 装配, 测试}，Summary 按类别分别 SUMIFS
- `Cost Bearer`：该工序成本最终由谁承担（本件 / 父件装配），防止工序成本挂错件
- `Cycle Basis` / `Rate Basis`：节拍与费率的依据文本；外协工序写 `外协（无设备节拍）`，不是空白
- **`Is Last Step (Y/N)`**：该件最后一道工序标 `Y`（其余留空）。`2-BOM` 的 `Supplier Process (CNY)` 只取 `Y` 那一行（`SUMIFS(..., Is Last Step, "Y")`）。**缺这一列就等于把同一件的所有工序连加多遍**——EX90 后逆变器 v4 就是因为没有这一列，工序费被重复累计。核对口径：`标 Y 的行数 = bottom-up 件数`（DC-DC 实测 10 件 / 10 个 Y）
- `Site` ∈ {in-house, outsource}。**外协工序**：Cycle/Rate/Operators 留空，`Step Direct = 外协单价 × 基数（如重量）÷ 良率`，单价引用 1-Assumptions 参数（如 `PLATE_PRICE`）
- 自制工序：`Step Direct = (Cycle/3600 × Rate + Operators × LaborRate × Cycle/3600 + Consumable) / Yield`
- `Cumulative Cost`：该件从小计到此道的累加，便于查工序顺序错乱
- **外购件/标准件不得出现在本表**；外协**工序**（自制件的某道工序外包）可以

## 4-Summary

列：`Part No | Name | Function | Category | Qty | Material | Purchased | Process | Assembly | Test | Tooling Amort | Direct Cost | Overhead | Factory Cost | Should-Cost Price | Share % | Status`

- 只用**叶件**求和（父件不参与），从结构上避免重复计价
- `Should-Cost Price = (Factory Cost × (1 + 物流率)) / (1 − 利润率)`
- **PPC 口径块**（合计行下方）：
  - `MAT = Σ材料（仅自制件）`——**外购件 `Purchased` 列一律留空**，防重复计价
  - `外购件采购价合计 = Σ价格取证组金额` —— **单列一行，不并入 MAT**
  - `VA = Σ加工 + Σ装配 + Σ测试`
  - `直接成本 = MAT + VA + 外购件采购价合计`
  - 工装 CapEx 合计（一次性）；工装摊销/件；应成本价。空列用 `N()` 包裹避免 `#VALUE!`
  - ⛔ **`MAT = Σ材料 + Σ外购件` 是禁止式**：外购件会既按采购价、又按材料价计两次，且还会被 OH/SG&A/利润再乘一次
- **单层口径三层块**（2026-11-03，PPC 口径块下方）：① 自制件供应商报价 = (直接 + OH + SG&A) × (1 + 物流率) ÷ (1 − 利润率)；② 外购件我方过手 = Σ(采购价 × Qty) × (1 + 物流率)（**不加 OH/SG&A/利润**）；③ 我方增值 = 装配人工 + EOL + 工艺材料 + 设备折旧；整机 = ① + ② + ③
- **成本完整度块**（Phase 2A 任务 B，PPC 口径块下方，公式自动判定）：外购件总数 / 有价外购件数 / Unknown 价格外购件数 / Known/Modeled 外购成本小计 / 成本数据覆盖率 / Complete PPC 判定。**Unknown 计数 > 0 时 PPC 只能标"部分（仅自制件）"**，不得标"完整"；Unknown 金额不得被 N()/SUM 静默当 0
- **基准价补位块（Class E，4-Summary）**：**分三行**——3a 严格类别 / 3b 包络类别 / 3c 质量轴兜底，每行给件数 / 金额（Base）/ 占 Level 1 比例；
  另给“补位前 Unknown 计数 → 补位后 Unknown 计数”。**三行不得合并为一个总数**，否则看不出哪几行最该先去拿报价。
  ★ 「Unknown 计数 = 0」不等于「成本已验证」：补位只是把空白填上，`Verified Price Coverage` 仍为 0。
  两块必须同时存在，只给其中一个就会把「没有空白」读成「成本可信」。
- **功能域透视块**：按 Function×Category 聚合 MAT / VA / Direct Cost / Share%（SUMIFS，求和区与条件区必须同尺寸——H/I/J 三列分别 SUMIFS 相加）。**必须带 `②层 Purchased` 一列**，否则 PCB / PM 功率级 / DC-Link 电容 / 电流传感器这些纯外购功能域会读 0，透视合计对不上明细。合计行 Share% = 透视直接成本合计 ÷ 明细直接成本合计，应 = 100%
- **Quotation 对比区**：报价输入 / 模型应成本价 / Variance / Variance%（报价缺失时留空，不填 0）
- 末尾目标价反推（Top-down）：填目标/竞品单价，反推外购件允许合计

## 5-Opportunities

列：`ID | 题目 | 现有方案 | 替代方案 | 工程依据 | Baseline 单价 | Proposed 单价 | 单件节省 | 净单件节省 | 年产量 | 年度节省 | 一次性投入 | 净生命周期收益 | MAT Impact | VA Impact | Tooling Impact | 降本机制 | 防重复计数组 | HG | 风险 | 验证周期 | 状态 | 优先级`

- 单件节省、净单件节省用公式，不要写死
- `净单件节省 = 单件节省 − 一次性投入 / MIN(工装寿命, 生命周期产量)`（工装投入摊到每件后的真实净收益）
- 状态 ∈ {Verified Saving, Modeled Saving, Hypothetical Opportunity}
- **MAT Impact / VA Impact / Tooling Impact**（Phase 2A 任务 C）：措施级成本归属，金额可空、归属必填；同一降本机制不得同时计入 MAT 与 VA（主归属二选一，另一列写"无"或定性说明）
- **降本机制 + 防重复计数组**：同组措施不得在合计中重复计节省；跨措施叠加前先查组号，同组需声明是否可叠加及依据
- **HG 列**（Hardness Grade，与 Bosch CTM 报告口径一致）：HG0 描述想法 / HG1 毛节省潜力 / HG2 净节省潜力 / HG3 时间表已排 / HG4 已实施 / HG5 节省已兑现。HG 记**实施成熟度**，与状态列（证据可信度）是两个独立坐标：HG4/HG5 只能由实施状态决定，不能由证据可信度决定
- 输入不全时公式返回空，不返回 0

## 6-Sensitivity

列：`驱动因素 | 基准值 | Low | Base | High | 对总成本影响(元) | 影响占比 | 说明`

- 只对关键驱动做，禁止"总成本 ±10%"式假敏感性
- 每一项必须能追到假设表里的参数
- **工装摊销阶梯**：给定工装 CapEx 合计，列出不同生命周期产量（如 50k/100k/200k/500k）下的摊销/件与对 PPC 的影响——同一措施在不同产量下净收益完全不同
- **外购件合计杠杆**（零输入方案）：外购件逐项价格未知时，用假设合计作为单一变量，给阶梯表直接查整机应成本

## 7-Evidence

列：`项 | 内容 | 来源/位置 | 状态 | 置信度 | 影响成本项 | 待确认问题 | 负责人`

- 来源写到可复核粒度：文件名 + 页/行；照片写到文件名 + 拍摄面
- `影响成本项`：该证据缺口卡住哪个成本桶/功能域（如 MAT 外购件 / VA 加工 / 工装摊销 / Quotation 对比）
- 待确认问题写"缺什么证据能回答"，不写"待补充"

## 格式与颜色

| 状态 | 填充 | 用途 |
|---|---|---|
| Verified | 无填充 | 原始资料 |
| Calculated | 浅蓝 `DDEBF7` | 公式结果 |
| Estimated | 浅黄 `FFF2CC` | 工程假设 |
| Unknown | 浅红 `FCE4E4` | 缺失，留空 |

- 表头深蓝底白字；冻结首行；千分位显示；货币列保留 2 位小数，质量列 3 位。
- 单位写在列名里，不写在单元格值里。

## 生成后必查

1. 无 `#REF!` / `#VALUE!` / `#DIV/0!`
2. 无循环引用
3. `4-Summary` 合计 = `3-ProcessCost` 分类合计 + 材料合计 + 外购合计（差值 < 0.5%）
4. 任一零件都不在 Summary 出现两次
5. 随机抽 2 个零件，手算一道工序，与表内一致
6. 空白项确实为空，没有被写成 0
7. PPC 口径闭合：MAT + VA = 直接成本合计；功能域透视合计 = 明细合计，透视 Share% 合计 = 100%
8. 外购件不出现在 3-ProcessCost；外协工序 Site = outsource 且公式为"单价×基数÷良率"
9. 工装两本账：CapEx 列（一次性）与摊销列（÷ MIN(寿命模次, 生命周期产量)）分开，不混记
10. **成本完整度块判定正确**：Unknown 计数 > 0 时 Complete PPC 显示"部分（仅自制件）：N 项外购件价格 Unknown"，不得标"完整"；覆盖率 = 有价外购件数 ÷ 外购件总数
11. **价格状态合规**：外购件 Price Status ∈ 八类且默认 Unknown；Supplier Quotation 未标 Verified；Assumption 未进 PPC 基线
12. **基准补位合规**：Class E 行有 Low/High 成对值、有四元标签（E / Benchmark Average / Comparable / Benchmark range, not a quotation）、
    有含类别+基数+来源+日期的 `Price Basis`；`4-Summary` 与 `7-Evidence` 均有一处披露；`Verified Price Coverage` 未因补位上升
12. **措施归属与防重复**：每条机会有 MAT/VA/Tooling 归属与防重复计数组；同组措施未重复计节省；HG 与状态列分别记录（HG4/HG5 只能由实施状态决定）
13. **单层口径闭合**：`4-Summary` 三层块 ① 自制件供应商报价 + ② 外购件过手 + ③ 我方增值 = 整机应成本（差值 < 0.5%）；② 中不得出现 OH/SG&A/利润项
14. **参数引用计数审计**：1-Assumptions 每个参数在全书公式中被引用 ≥ 1 次；0 次引用的死参数要么接线要么删除（教训：SG&A 定义了却没接进公式，模型直接少一层 5%）
