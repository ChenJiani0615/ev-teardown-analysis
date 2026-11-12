# ev-teardown-analysis

汽车电子总成拆解对标分析的 Agent Skill。默认产出「机械 / 材料 / 成本 / 总装工艺」四张表 + 结论 + 补充取证需求。

与 `vehicle-data-hub-main` 无耦合，可独立使用。

> 🌐 English version: [README.en.md](README.en.md)

## 目录

```
ev-teardown-analysis/
├── README.md / README.en.md
├── .github/skills/
│   ├── ev-teardown-benchmark/         # 拆解对标分析（四表）
│   │   ├── SKILL.md                   # 主流程（入口）
│   │   ├── references/                # 输出模板、证据规则、件号映射、表面特征矩阵、机理库、基准来源、架构发现
│   │   └── scripts/                   # part_table_check.py 清单自检 / surface_scan.py 表面特征扫描
│   └── should-cost-engineering/       # Should-Cost 核算（7 表工作簿，活公式）
│       ├── SKILL.md
│       └── references/                # 工作簿规格、工艺成本模型
├── docs/                              # 大白话安装使用说明
├── tests/                             # 脚本测试（fixtures 为脱敏演示数据）
└── tools/                             # 通用工具（工作簿生成、清单展开、报告导出、CS 记录解析等）
```

> 拆解照片与零件清单样例**不在仓库中**——`.gitignore` 白名单只上传 skill 与通用工具，使用时自备清单与照片。

## 安装

**装到个人环境（跨项目可用）**

```powershell
# Windows
New-Item -ItemType Directory -Force "$env:USERPROFILE\.copilot\skills" | Out-Null
Copy-Item -Recurse -Force ".github\skills\ev-teardown-benchmark" "$env:USERPROFILE\.copilot\skills\"
```

```bash
# macOS / Linux
mkdir -p ~/.copilot/skills
cp -R .github/skills/ev-teardown-benchmark ~/.copilot/skills/
```

> Skill 只能在 `.github/skills/`、`.agents/skills/`、`.claude/skills/` 或 `~/.copilot/skills/` 等位置被发现；用户级 prompts 目录不承载 skill。

> 📖 **不想读术语？** 看大白话版安装使用说明：`docs/安装使用-大白话.md`

## 使用

1. 在对话里输入 `/ev-teardown-benchmark`，或直接描述任务（描述命中关键词时会自动加载）。
2. 提供：车型/年款/配置 + 部位 + 拆解照片或零件数据。
3. 有 CSV/Excel 时先跑自检：

```bash
python .github/skills/ev-teardown-benchmark/scripts/part_table_check.py <零件清单.csv>
```

输出会给出：件号合计质量、各件质量占比、材料组汇总、尺寸字段错位告警、缺字段件、以及"清单中未出现功率模块/散热关键词"的查漏提示。

## 典型请求

```
/ev-teardown-benchmark Volvo EX90 Ultra MY2025 后逆变器，零件清单见 parts.csv
/ev-teardown-benchmark 某车型 OBC，需要完整版（含逐项深入分析与工艺步骤表）
/ev-teardown-benchmark 核对这份拆解数据的字段一致性和质量占比
```

## 扩展

| 想加的能力 | 改哪里 |
|---|---|
| 新的总成类型（BMS、域控、热管理）常见件类 | `references/part-mapping.md` 第三节 |
| 新的字段陷阱 | `references/evidence-rules.md` 第二节 |
| 新的机理公式（如 NVH、疲劳） | `references/mechanism-library.md` |
| 新的表面特征类型或判读规则 | `references/surface-feature-matrix.md` |
| 输出格式变化 | `references/output-template.md` + `SKILL.md` 的 Step 5/6 |
| 清单列名新增 | `scripts/part_table_check.py` 的 `FIELD_ALIASES` |

## 依赖

- 必需：Python 3.9+（仅标准库）
- 可选：`openpyxl`（读 xlsx）

## 设计取舍

- **默认精简**：四表 + 结论 + 5 项取证。逐项深入分析与工艺步骤表需显式要求，避免每次都产出超长报告。
- **成本表不编数字**：无报价/产量/良率时只给驱动因素与公式模板，避免被当成真实成本基线引用。
- **候选分级**：证据不足时并列 A/B/C 候选并给出"一刀切开"的证据，而不是二选一猜结论。
- **稳态之外必查退化**：热路径结论只覆盖一维稳态；厚 TIM、无基板器件、大平面界面、相变材料、功率循环等组合出现时，必须补泵出/蠕变/疲劳等寿命机制，并声明模型未覆盖项。
- **表面特征逐面判读**：同一纹理对不同对象（冷却液/器件/TIM/密封/结构/工艺）可能有完全不同的功能，只写一侧会系统性漏掉机理。
