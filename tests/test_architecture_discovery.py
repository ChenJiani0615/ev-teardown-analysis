#!/usr/bin/env python3
"""架构创新发现（Step 4.5）行为回归测试（仅标准库）。

这组测试校验的是 **输出结构的强制项与禁止项**，不是数值：
即"skill 要求必须出现的段落是否出现"、"skill 禁止的表述是否泄漏"。

用法：
    python tests/test_architecture_discovery.py

四个场景（对应用户要求）：
    1. substitution-with-evidence  —— 有 E2 证据（版本对比）时，允许且必须写出"替代"
                                     且必须 E 级 + 可信度**双标注**
    2. module-only-no-substitution  —— 只有模块存在（E0），不得虚构替代历史
    3. discrete-not-module          —— 12 个分立 MOSFET，P 轴必须判 P0，不得判 P2；
                                      T 轴不得因"有两块板"直接判 T2
    4. obcm-architecture-baseline    —— 非 MOSFET 对象（磁件/材料路线），也能产出架构发现；
                                      功率级无器件级证据时只给**逐维度上限**

四维度纪律（本轮新增，替换原 L0–L4 单刻度）：
    P 封装与集成 / T 冷却与热路 / E 电路与开关（四栏）/ M 安装与压紧
    四条轴**互相独立**，输出必须逐轴声明等级 + 依据 + 置信度，
    且**不得跨轴比较优劣**（"双面冷却所以集成度更高"必须被判违规）。

证据纪律（本轮新增）：
    E 级（证据类型，客观）与 可信度（结论可靠度，主观）是**两个坐标**。
    任何 E 级标注必须在附近（同一行）带上可信度，
    否则视为"把证据级别当可信度用"。

比较口径：Markdown 文本 + 正则/关键词断言。基线文件是"合格输出样本"，
测试同时做两件事：
    * 正向：样本本身满足全部强制结构项（说明流程可执行）
    * 负向：样本不含禁止表述（说明流程不会滑向错误结论）
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures" / "arch"

# ── 强制结构项：每份合格输出都必须具备 ──────────────────────────────
REQUIRED_SECTIONS = (
    "## 0.",
    "## 1.",
    "## A.",
    "## B.",
    "## C.",
    "## D.",
    "## 结论",
    "## 补充需求",
)

FUNCTION_SLOTS = (
    "电能变换",
    "开关动作",
    "储能与滤波",
    "电流承载与分配",
    "热传导与散热",
    "电气绝缘与隔离",
    "检测与传感",
    "密封与防护",
    "紧固与定位",
)

FIVE_COLUMN_HEADER = re.compile(
    r"\|\s*题目／一句话解释\s*\|\s*业界通用方案／比较基准\s*\|\s*相对优点\s*\|"
    r"\s*相对缺点\s*\|\s*由此得到的启发\s*\|"
)

D_SEVEN_COLUMN_HEADER = re.compile(
    r"\|\s*步骤\s*\|\s*对象／部件\s*\|\s*工艺或操作\s*\|\s*目的及顺序依据\s*\|"
    r"\s*关键控制点\s*\|\s*潜在缺陷及检查方式\s*\|\s*证据状态\s*\|"
)

# ── 四维度强制项 ──────────────────────────────────────────────────
DIMENSION_PATTERNS = {
    "Packaging": re.compile(r"Packaging|\*\*P\*\*\s*封装|封装与集成"),
    "Thermal": re.compile(r"Thermal|\*\*T\*\*\s*冷却|冷却与热路"),
    "Electrical": re.compile(r"Electrical|\*\*E\*\*\s*电路|E-a|E-b|E-c|E-d"),
    "Mounting": re.compile(r"Mounting|\*\*M\*\*\s*安装|安装与压紧"),
}

# 等级刻度：四轴各自的合法等级字母。用于检查"字母用对了轴"。
AXIS_SCALE = {
    "P": re.compile(r"\bP[0-4]\b"),
    "T": re.compile(r"\bT[0-4]\b"),
    "M": re.compile(r"\bM[0-4]\b"),
}

CONFIDENCE_WORDS = ("可信度", "置信度")

# 旧单刻度残留：L0–L4 不得再作为架构等级出现
LEGACY_LEVEL = re.compile(r"(?<![A-Za-z])L[0-4](?![0-9])")


def load(name: str) -> str:
    path = FIXTURES / f"{name}.md"
    if not path.exists():
        raise SystemExit(f"缺少测试基线：{path}")
    return path.read_text(encoding="utf-8")


def substantive_lines(text: str):
    """剔除引用块（`>` 开头）与标题（`#` 开头）后逐行产出。

    引用块里放的是**规范本身的说明**（如"禁止写 X"），不是输出断言，
    若不剔除，所有禁用表述检查都会被规范文本自己触发。
    标题行（`# 场景 N：…（E2 时…）`）是对**场景的描述**而非本体的结论，
    同样既不应参与"E 级必须带可信度"的检查，也不应参与断言检查。
    """
    for line in text.splitlines():
        stripped = line.lstrip()
        if stripped.startswith(">") or stripped.startswith("#"):
            continue
        yield line


def claim_clauses(line: str):
    """按中文标点把一行切成小句，供逐句否定判断。

    必须按**小句**而不是按整行判断否定，否则
    "本产品以模块替代了分立器件，但是否切换无证据" 这种
    既违规又自相矛盾的写法会被整行末尾的"无证据"放过。
    """
    for clause in re.split(r"[，。；、：!?！？]", line):
        if clause.strip():
            yield clause


def check_common(name: str, text: str, failures: list[str]) -> None:
    """所有场景都必须满足的结构纪律。"""
    for section in REQUIRED_SECTIONS:
        if section not in text:
            failures.append(f"{name}: 缺少强制段落 {section}")

    for slot in FUNCTION_SLOTS:
        if slot not in text:
            failures.append(f"{name}: 功能槽 '{slot}' 未出现（九类功能槽不得有沉默项）")

    if not FIVE_COLUMN_HEADER.search(text):
        failures.append(f"{name}: A/B/C 表缺少五列固定表头")

    if not D_SEVEN_COLUMN_HEADER.search(text):
        failures.append(f"{name}: D 表缺少七列固定表头")

    if "证据边界" not in text:
        failures.append(f"{name}: D 表前缺少证据边界声明")

    # ── 四维度必须逐轴点名 ──
    for axis, pattern in DIMENSION_PATTERNS.items():
        if not pattern.search(text):
            failures.append(f"{name}: 架构四维度缺少 {axis} 轴（四轴必须全部显式声明）")

    # ── 等级字母必须落在自己的轴上：P/T/M 各至少出现一次 ──
    for axis, pattern in AXIS_SCALE.items():
        if not pattern.search(text):
            failures.append(f"{name}: 未出现 {axis} 轴等级标注（{axis}0–{axis}4）")

    # ── 旧单刻度不得残留 ──
    for line in substantive_lines(text):
        legacy = LEGACY_LEVEL.search(line)
        if legacy and "L0–L4" not in line and "L0-L4" not in line:
            failures.append(
                f"{name}: 仍在使用旧的单刻度层级 '{legacy.group()}'"
                f"（已改为 P/T/E/M 四维度）: {line.strip()[:60]}"
            )
            break

    # ── 可信度必须显式出现 ──
    if not any(word in text for word in CONFIDENCE_WORDS):
        failures.append(f"{name}: 未给出任何'可信度/置信度'（证据类型 ≠ 结论可信度）")

    # ── 成本表不得出现金额/百分比式降本断言 ──
    for bad in re.finditer(r"降本\s*\d+\s*%|节省\s*\d+\s*%|成本\s*下降\s*\d+\s*%", text):
        failures.append(f"{name}: 出现无数据的降本百分比 '{bad.group()}'")

    # ── 基准列不得出现"传统方案" ──
    if "传统方案" in text:
        failures.append(f"{name}: 出现禁用表述 '传统方案'（必须写具体路线名）")


def check_no_cross_axis_claim(name: str, text: str, failures: list[str]) -> None:
    """禁止跨轴比较优劣。

    这是把 L0–L4 拆成四维度后**最重要的新增纪律**：四条轴没有序关系。
    典型违规：把 T 轴的高值（双面冷却）说成 P 轴的优点（集成度更高）。
    """
    cross_axis = (
        re.compile(
            r"(双面冷却|双面液冷|T2)[^。\n]{0,25}"
            r"(更高集成|集成度更高|更先进|优于.{0,6}分立|比.{0,10}更集成)"
        ),
        re.compile(
            r"(集成度更高|更先进|集成程度更高)[^。\n]{0,25}"
            r"(因为|由于)[^。\n]{0,15}(双面冷却|T2)"
        ),
        re.compile(
            r"\bP[0-4]\b[^。\n]{0,10}(高于|优于|先进于|大于)[^。\n]{0,10}\b(T|M)[0-4]\b"
        ),
        re.compile(
            r"\b(T|M)[0-4]\b[^。\n]{0,10}(高于|优于|先进于|大于)[^。\n]{0,10}\bP[0-4]\b"
        ),
    )
    NEGATIONS = ("不得", "禁止", "不作", "不能", "不应", "不代表", "非", "无法", "不构成")
    # 注意：这里按**整行**而非按小句匹配。
    # 跨轴比较的语义单元本身横跨逗号（"用了双面冷却，所以集成度更高"），
    # 按小句切分会把这一违规切散、两个半句各自都不匹配 -> 漏检。
    # 否定则在**行**级别判定：整行带否定词的视为规范说明/否定句，跳过。
    for line in substantive_lines(text):
        if any(neg in line for neg in NEGATIONS):
            continue
        if any(p.search(line) for p in cross_axis):
            failures.append(
                f"{name}: 出现跨轴优劣比较（四维度无序关系）：{line.strip()[:70]}"
            )
            return


def clause_at(line: str, pos: int) -> str:
    """返回 `pos` 所在的小句（按中文标点切分）。"""
    start = 0
    for match in re.finditer(r"[，。；：!?！？]", line):
        if match.start() < pos:
            start = match.end()
        else:
            return line[start: match.start()]
    return line[start:]


# 门槛式提及：在讨论"E 级要到几级才够"，不是在给某个结论做标注
THRESHOLD_MARKERS = (
    "达到", "需要", "才成立", "门槛", "要求", "最低", "不到", "阶梯",
    "级别", "必须", "限于", "适用范围", "只对", "提到", "升到", "抬到",
)


def check_dual_annotation(name: str, text: str, failures: list[str]) -> None:
    """E 级标注必须带可信度，否则视为把证据级别当可信度用。

    例外只有两类，且都必须能被一句话说清：
    * 概念定义行（同行同时出现"证据类型"与"可信度"）
    * 门槛式提及（在说"E 级要到几级才够"，不是在标注某个结论）
    真正给结论做标注却在附近找不到可信度的，一律判违规。
    """
    e_level = re.compile(r"\bE[0-4]\b")
    for line in substantive_lines(text):
        if "证据类型" in line and "可信度" in line:
            continue  # 概念定义行
        for hit in e_level.finditer(line):
            window = line[max(0, hit.start() - 20): hit.end() + 25]
            if "可信度" in window or "置信度" in window:
                continue
            clause = clause_at(line, hit.start())
            if any(marker in clause for marker in THRESHOLD_MARKERS):
                continue
            failures.append(
                f"{name}: E 级标注 '{hit.group()}' 附近未给可信度"
                f"（证据类型 ≠ 结论可信度，必须双标注）: {line.strip()[:70]}"
            )
            return


def has_substitution_claim(text: str) -> bool:
    """检测是否写出了"替代"性质的历史断言（而非仅描述功能归属）。

    按**分句**判断否定，不按整行——否则"本产品以模块替代了分立器件，但是否切换无证据"
    这种既违规又自相矛盾的写法会被整行的否定词放过。
    """
    NEGATIONS = ("不得", "禁止", "不作", "不能", "无证据", "不应", "不代表",
                 "非历史", "不构成")
    # 疑问/虚拟语气不是断言："模块是否替代了分立" 是在提问，不是在断言
    HEDGES = ("是否", "能否", "有无")
    patterns = (
        r"替代了",
        r"取代了",
        r"改为.{0,12}模块",
        r"由\s*[LP]0\s*升级为",
    )
    for line in substantive_lines(text):
        for clause in claim_clauses(line):
            if any(neg in clause for neg in NEGATIONS + HEDGES):
                continue
            if any(re.search(p, clause) for p in patterns):
                return True
    return False


def has_per_dimension_ceiling(text: str) -> bool:
    """必须逐维度声明上限，而不是只给一个总上限。"""
    if not re.search(r"(上限|不能判定|无法判定)", text):
        return False
    per_axis = 0
    for axis in ("P", "T", "M"):
        if re.search(rf"{axis}\s*(轴)?[^。\n]{{0,60}}(上限|不能判定|无法判定)", text):
            per_axis += 1
    if re.search(r"Electrical[^。\n]{0,100}(不能判定|需原理图)", text):
        per_axis += 1
    return per_axis >= 2


def scenario_1(failures: list[str]) -> None:
    """有 E2 证据时必须能写出替代，且必须双标注（E 级 + 可信度）。"""
    name = "substitution-with-evidence"
    text = load(name)
    check_common(name, text, failures)
    check_no_cross_axis_claim(name, text, failures)
    check_dual_annotation(name, text, failures)

    if not has_substitution_claim(text):
        failures.append(f"{name}: 有 E2 级版本对比证据，却未写出替代关系（过度保守，应报出机会）")

    if "E2" not in text:
        failures.append(f"{name}: 写了替代关系但未标注证据级别 E2")

    if "版本对比" not in text:
        failures.append(f"{name}: 未说明替代结论的证据来源是版本对比")

    # E2 有效性条件必须写出来，否则 E2 会被当成"可靠证据"
    if not re.search(r"(同平台|同功率|同市场|适用范围|限于)", text):
        failures.append(f"{name}: 写了 E2 却未声明其有效性前提（同平台/同功率档/同市场）")

    # 替代必须落在**同一轴**上：本案是 P 轴的 P0 → P2
    if "P0" not in text or "P2" not in text:
        failures.append(f"{name}: 替代关系应落在 P 轴上（P0 → P2），当前未同时写出两端等级")


def scenario_2(failures: list[str]) -> None:
    """只有 E0 证据时，禁止出现历史替代断言。"""
    name = "module-only-no-substitution"
    text = load(name)
    check_common(name, text, failures)
    check_no_cross_axis_claim(name, text, failures)
    check_dual_annotation(name, text, failures)

    if has_substitution_claim(text):
        failures.append(f"{name}: 仅有 E0 证据却写了历史替代断言（违反证据阶梯）")

    if "E0" not in text and "E1" not in text:
        failures.append(f"{name}: 未声明当前证据级别（应写 E0/E1）")

    if not re.search(r"无证据.{0,10}不作判断|不构成替代|不作判断", text):
        failures.append(f"{name}: 未明确写出'无证据、不作替代判断'")

    # E0 + 可信度高 合法，但只对"存在"成立
    if not re.search(r"(存在性|可靠的是.{0,6}存在|不等于.{0,8}替代|不能推出替代)", text):
        failures.append(
            f"{name}: 未区分'存在可靠'与'替代可靠'"
            f"（E0 + 可信度高 只能支撑 A 存在，不能支撑替代）"
        )


def scenario_3(failures: list[str]) -> None:
    """12 个分立 MOSFET 不得被误判为集成模块；T 轴不得仅凭两块板判 T2。"""
    name = "discrete-not-module"
    text = load(name)
    check_common(name, text, failures)
    check_no_cross_axis_claim(name, text, failures)
    check_dual_annotation(name, text, failures)

    if "P0" not in text:
        failures.append(f"{name}: 分立方案未在 P 轴判定为 P0")

    if "12" not in text:
        failures.append(f"{name}: 未体现 12 个独立器件件号这一关键证据")

    # P 轴结论必须是 P0，不得断言本产品属于 P2/P3/P4
    asserts_high_p = re.compile(
        r"(本产品|本总成|本方案)[^。\n]{0,20}"
        r"(判定为|属于|即为|就是|采用了?|使用的是)\s*\**\s*P[234]"
        r"|判定为\s*\**\s*P[234]\s*\**\s*〔?\s*资料事实"
    )
    for line in substantive_lines(text):
        for clause in claim_clauses(line):
            if any(neg in clause for neg in ("不得", "禁止", "不能", "不是", "未")):
                continue
            if asserts_high_p.search(clause):
                failures.append(f"{name}: 把分立方案判定为集成封装：{clause.strip()[:70]}")
                break
        else:
            continue
        break

    # 禁止把本产品说成用了模块（"未使用" 是否定句，合格）
    for line in substantive_lines(text):
        if any(neg in line for neg in ("不得", "禁止", "不能", "未使用", "未采用")):
            continue
        if re.search(r"(本产品|本总成)[^。\n]{0,20}(使用|采用)[^。\n]{0,10}功率模块", line):
            failures.append(f"{name}: 把分立方案误判为使用功率模块：{line.strip()[:70]}")
            break

    # T 轴纪律：有两块冷却板 ≠ T2
    if not re.search(r"(不能|不足以|未证|需|必须)[^。\n]{0,30}(T2|双面)", text):
        failures.append(
            f"{name}: T 轴未写出'有两块板不能直接判 T2'的限定"
            f"（必须排除其中一块为不导通冷却液的压板）"
        )

    # M 轴纪律：有定位件 + 独立压紧件时，应优先 M1
    if "M1" not in text:
        failures.append(f"{name}: 有 Locator/Spacer/Holder 却未在 M 轴给出 M1 候选")


def scenario_4(failures: list[str]) -> None:
    """非 MOSFET 对象也必须能产出架构级发现；无器件级证据时只给逐维度上限。"""
    name = "obcm-architecture-baseline"
    text = load(name)
    check_common(name, text, failures)
    check_no_cross_axis_claim(name, text, failures)
    check_dual_annotation(name, text, failures)

    # 必须有一条实打实落进 A/B/C 表的磁类架构发现行
    magnet_terms = ("磁集成", "磁件", "磁芯", "铁氧体", "共模", "差模")
    row_re = re.compile(r"^\|\s*\*\*[MC]0\d[^\n]*$", re.MULTILINE)
    magnet_rows = [
        row for row in row_re.findall(text)
        if any(term in row for term in magnet_terms)
    ]
    if not magnet_rows:
        failures.append(
            f"{name}: 未在 A/B/C 表内落出磁类架构发现行"
            f"（非 MOSFET 对象必须也能产出架构级发现，不能只覆盖功率器件）"
        )

    table_rows = len(row_re.findall(text))
    if table_rows < 2:
        failures.append(f"{name}: A/B/C 表内实质发现行不足（找到 {table_rows} 行，至少需 2）")

    # 无器件级证据时不得硬给某轴的资料事实级定论
    if re.search(r"判定为\s*P[0-4]（[^）]*）\s*〔资料事实", text):
        failures.append(f"{name}: 无器件级证据却给了 P 轴的资料事实级定论")

    if not has_per_dimension_ceiling(text):
        failures.append(f"{name}: 无器件级证据时未逐维度写出'上限/不能判定'")


SCENARIOS = (
    ("1 有证据的替代（E2，双标注）", scenario_1),
    ("2 只有模块、不得虚构替代（E0）", scenario_2),
    ("3 分立不得误判为模块（P0）", scenario_3),
    ("4 非 MOSFET 对象的架构发现", scenario_4),
)


# ── 变异测试：证明禁止项不是空条款 ─────────────────────────────────
# 做法：把一份合格基线注入一处**典型错误**，断言对应检查器会报错。
# 如果某个检查器对注入的错误无动于衷，说明它是空条款（vacuous），
# 必须修检查器而不是改基线。这是"规则是否真的在起作用"的唯一证据。

def mutation_checks() -> list[str]:
    """返回变异测试的失败项列表。"""
    problems: list[str] = []

    def expect_fire(label: str, base: str, mutate, checker, hint: str) -> None:
        text = load(base)
        if mutate(text) == text:
            problems.append(f"[变异] {label}: 注入失败（替换目标未命中，变异无效）")
            return
        bucket: list[str] = []
        checker(label, mutate(text), bucket)
        if not bucket:
            problems.append(f"[变异] {label}: {hint}——检查器未触发（空条款）")

    # 1. 跨轴优劣比较（把 T 轴高值说成 P 轴优点）
    expect_fire(
        "跨轴比较",
        "discrete-not-module",
        lambda t: t.replace(
            "## 结论\n",
            "## 结论\n\n- 本产品用了双面冷却，所以集成度更高。\n",
            1,
        ),
        check_no_cross_axis_claim,
        "四维度无序关系不得比较优劣",
    )

    # 2. P/T/M 等级之间的书面序关系
    expect_fire(
        "轴间大小比较",
        "discrete-not-module",
        lambda t: t.replace(
            "## 结论\n",
            "## 结论\n\n- P2 高于 T1，属于更先进的封装。\n",
            1,
        ),
        check_no_cross_axis_claim,
        "P 与 T 不在同一刻度",
    )

    # 3. E 级标注缺可信度
    expect_fire(
        "E 级缺可信度",
        "substitution-with-evidence",
        lambda t: t.replace(
            "**P 轴：P0 → P2，模块替代了分立阵列**（E2 / 可信度中高",
            "**P 轴：P0 → P2，模块替代了分立阵列**（E2，",
            1,
        ),
        check_dual_annotation,
        "证据类型 ≠ 结论可信度，必须双标注",
    )

    # 4. 旧单刻度残留
    def inject_legacy(t: str) -> str:
        return t.replace("## 结论\n", "## 结论\n\n- 该总成的集成层级为 L3，属行业主流。\n", 1)

    text = load("discrete-not-module")
    mutated = inject_legacy(text)
    if mutated == text:
        problems.append("[变异] 旧单刻度: 注入失败")
    else:
        bucket: list[str] = []
        check_common("旧单刻度", mutated, bucket)
        if not any("单刻度" in b for b in bucket):
            problems.append("[变异] 旧单刻度: L3 未被检出（已改为四维度的纪律失效）")

    # 5. 缺维度标签
    def strip_thermal(t: str) -> str:
        out = []
        for line in t.splitlines(keepends=True):
            if "Thermal" in line or "**T**" in line or "冷却与热路" in line:
                continue
            out.append(line)
        return "".join(out)

    text = load("discrete-not-module")
    mutated = strip_thermal(text)
    if mutated == text:
        problems.append("[变异] 缺维度标签: 注入失败")
    else:
        bucket = []
        check_common("缺维度标签", mutated, bucket)
        if not any("Thermal" in b for b in bucket):
            problems.append("[变异] 缺维度标签: 删掉 T 轴后未报错（四轴强制项失效）")

    # 6. 分立方案被改判为 P2
    def flip_to_p2(t: str) -> str:
        return t.replace(
            "## 结论\n",
            "## 结论\n\n- 本总成判定为 P2 集成封装。\n",
            1,
        )

    text = load("discrete-not-module")
    mutated = flip_to_p2(text)
    if mutated == text:
        problems.append("[变异] 分立改判 P2: 注入失败")
    else:
        asserts_high_p = re.search(
            r"(本产品|本总成|本方案)[^。\n]{0,20}"
            r"(判定为|属于|即为|就是|采用了?|使用的是)\s*\**\s*P[234]",
            mutated,
        )
        if not asserts_high_p:
            problems.append("[变异] 分立改判 P2: 注入的断言未被识别（正则过窄）")

    # 7. 无证据仍写替代（E0 却写替代了）
    expect_fire(
        "E0 写替代",
        "module-only-no-substitution",
        lambda t: t.replace(
            "## 结论\n",
            "## 结论\n\n- 功率模块替代了分立阵列。\n",
            1,
        ),
        lambda n, txt, b: (
            b.append(f"{n}: 仅有 E0 证据却写了历史替代断言") if has_substitution_claim(txt) else None
        ),
        "E0 不得支撑替代结论",
    )

    # 8. 无数据的降本百分比
    expect_fire(
        "无据降本百分比",
        "discrete-not-module",
        lambda t: t.replace(
            "## 结论\n",
            "## 结论\n\n- 改用模块可降本 18%。\n",
            1,
        ),
        lambda n, txt, b: (
            b.append(f"{n}: 无数据降本百分比") if re.search(r"降本\s*\d+\s*%", txt) else None
        ),
        "成本结论必须给驱动因素，不给百分比",
    )

    return problems


def main() -> int:
    if not FIXTURES.exists():
        raise SystemExit(f"找不到测试基线目录：{FIXTURES}")

    failures: list[str] = []
    for label, scenario in SCENARIOS:
        before = len(failures)
        try:
            scenario(failures)
        except SystemExit as exc:
            failures.append(str(exc))
        status = "ok" if len(failures) == before else "FAIL"
        print(f"[{status}] 场景{label}")

    mutations = mutation_checks()
    status = "ok" if not mutations else "FAIL"
    print(f"[{status}] 变异测试（{8} 处注入，验证禁止项非空条款）")

    failures.extend(mutations)

    if failures:
        print("\n--- 架构发现回归失败 ---")
        for block in failures:
            print(f"- {block}")
        return 1

    print(f"\n全部通过（{len(SCENARIOS)} 个架构发现场景 + 8 处变异测试）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())