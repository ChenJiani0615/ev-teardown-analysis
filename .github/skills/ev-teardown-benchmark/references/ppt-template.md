# PPT 交付模板（对标参考版式）

> 适用场景：用户在四表分析完成后，要求把结果做成 **英文 PPT**（.pptx）。
> 交付形态由用户指定：用户给出页序就按页序做，**不得增删章节**；本文件只规定
> 版式语言、数据映射与证据纪律，不规定页数。

## 一、页序映射（用户指定 8 段时的标准映射）

| # | 页面 | 数据来源 | 备注 |
|---|---|---|---|
| 0 | Title + KPI 卡 | should-cost 4-Summary（S1/L1/L2）+ benchmark §0（件数/质量） | KPI 数字必须与源报告逐位一致 |
| 1 | Summary | should-cost 4-Summary 结论 + benchmark 结论段 | 三段式：System Design & BOM / Should-Cost Baseline / Key DTC Levers |
| 2 | Vehicle & Product Intro | 车型公开参数〔模型先验，未核验〕+ benchmark §0 产品画像 | 车型参数无出处时必须标注 |
| 3 | Exploded View | 拆解照片（整总成爆炸/俯视图） | 单张大图；先读像素尺寸算宽高比再放置 |
| 4 | Cost Dissection | should-cost 功能域透视（MAT / VA 两部分） | 图表化：堆叠条 + 功能域横条；标注价格等级降级（C+D 占比） |
| 5 | Cost Insights | benchmark C 表 | 表格或四象限卡 |
| 6 | Mechanical Insights | benchmark A 表 | 表格或四象限卡 |
| 7 | Material Insights | benchmark B 表 | 表格 + 质量分组统计行 |
| 8 | Process Sequence | benchmark D 表 | 3×3 步骤卡网格；**必须带证据边界声明**（推断顺序） |

## 二、版式语言（从参考 PPT 提取的设计 token）

- **画布**：16:9（13.333 × 7.5 in）；字体 **Inter**（缺失时回退 Segoe UI/Arial，全篇统一）。
- **色板**：
  - 主色 teal：`0F766E`（深 `115E59` / `134E4A` / `022C22`）
  - 文字 slate：标题 `0F172A`、正文 `334155`、次要 `64748B`
  - 背景：`FFFFFF` / `F8FAFC` / `F1F5F9`；浅填充 `F0FDFA`（teal tint）/ `ECFDF5` / `E2E8F0`
  - 强调：rose `E11D48`、green `059669`、purple `7E22CE`、blue `0284C7`、amber `92400E`
- **每页家具**（recurring furniture）：
  - 顶部强调线：矩形 (0.33", 0.88") 宽 12.67" 高 0.02"，teal
  - Kicker（小标签）+ 标题：左上 (0.33", 0.25" / 0.42")
  - 右上角 badge：圆角矩形，页主题缩写
  - 页脚：左 "Confidential | <项目名>"，右 "Slide N / <total>"，y≈7.22"
  - 底部 5 段色条：5 × 2.67" 宽矩形，y=7.45"，h=0.04"
- **表格样式**：表头行 fill `115E59`、白字 9pt 加粗；正文行 `FFFFFF`/`F8FAFC` 交替；
  正文 `334155` 9pt；单元格边距 marL/R=63500 marT/B=25400 EMU，垂直居中；**所有边框 noFill**。
- **BRIEF SUMMARY 条**：表格页底部放 "BRIEF SUMMARY — <主题>" pill + 一句话总结带。

## 三、数据纪律（硬规则，继承 SKILL.md 硬性规则 1/2/9）

1. **每个数字必须能在源报告（benchmark md / should-cost md / 工作簿）里找到同一个数**。
   生成器用 CONST 块集中声明全部数字，页面文案只引用变量，不允许手写第二份。
2. **证据状态随行走**：Scenario / Class C+D 降级 / 〔工程推断〕/〔模型先验，未核验〕
   必须出现在对应页面（成本页标降级，工艺页标推断顺序，基准列标出处）。
3. **AC 修正不是降本**；紧固件包的价值在动作数不在单价——这两句要出现在成本启示页。
4. **不新增用户未要求的章节**；参考 PPT 里有而用户没点名的页（如空间 BOM、综合路线图）不做。
5. 交付前校验：用 python-pptx 读回，逐页抽查数字与源报告一致；页数、页脚编号正确。

## 四、生成器约定

- **基准参考稿（用户指定，一律按此走，禁止重新设计版式）**：
  `data/referance/Volvo_EX90_Rear_Inverter_Benchmark_&_Should-Cost_Analysis.pptx`
  （12 页、16:9、Inter、teal `0F766E`；注意目录拼写是 `referance`，不是 reference。）
  做任何总成的 PPT 都用 `shutil.copyfile(此稿, 输出)` → 原地改文字，版式/图标/渐变/表格 1:1 继承。
- **版式继承、内容跟随实际数据**：版式/色板/字体/每页家具照参考稿；但**条目数量与内容以实际数据为准，不删减、不截断**。
  例：参考稿工艺页 5 步、实际 7 步 → 排 7 步（沿用参考稿步骤卡样式）；表格行数、KPI 卡数、分层数同理，实际几条做几条。
- 位置 `tools/build_<assembly>_deck_en.py`；输出 `reports/<Assembly>_Deck_EN.pptx`。
- 复用参考稿的结构模式（CONST 块 + `text/bullets/rect/head/foot/table/bar/kpi` helpers），版式 token 即参考稿本身。
- 照片路径 `photos/<assembly>/...`；放置前用 PIL 或 python-pptx 读原图宽高比，按高度定宽居中。
