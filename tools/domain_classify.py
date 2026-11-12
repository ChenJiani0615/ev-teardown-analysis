#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""产品无关的功能域分类 / 外购判定 / 工艺路由推断 —— 单一真源。

为什么要有这个模块（一次真实的回归）
====================================
v1 `build_should_cost_workbook.py` 用**件名 + 材料牌号**推断 Function/Category 与工艺路由，
因此只给一份叶件清单 CSV 就能建出完整的 3-ProcessCost（实测：EX90 后逆变器 53/53 叶件全部
解析成功，22 外购 / 31 自制，10 条不同路由）。

v2 `build_should_cost_workbook_v2.py` 把这套推断**换成了读外购件属性表**：
    route_key = OBJECT_ROUTE.get(at.get("object", ""))
    bottomup  = bool(at.get("bottomup")) and route_key is not None
外购件属性表（`*-purchased-attributes.csv`）是 DC-DC 项目专有的交付物，逆变器项目**没有**。
于是对逆变器：`obj=""` → `route_key=None` → `bottomup=False` → **一个 bottom-up 工艺行都不生成**，
4-Summary 退化成"仅自制件下限"，2-BOM 的 Route 列全部写成"不建工艺（按取证价采购）"。

实测对照（同一份 53 行逆变器叶件）：
    v1 件名推断        : 53/53 命中路由
    v2 attrs 查表      :  0/53 命中路由

结论：**通用性在旧代码里，是在新代码里被弄丢的。** 这不是"DC-DC 与逆变器的产品差别"，
而是"把证据输入当成了必要条件"。属性表应当只做**增强**（提供九类材料分层、更精确的
采购对象与路由状态），不应是**前置条件**。

三级降级（本模块提供的第 2、3 级）
----------------------------------
    第 1 级  attrs 给出 Procurement Object → 用 OBJECT_ROUTE 精确路由（证据最全）
    第 2 级  无 attrs / Object 未识别   → 用件名+材料牌号推断（route_status = "Inferred"）
    第 3 级  推断也拿不到                → 显式写"未建工艺(Unknown)"，并留空成本
                                          （**绝不静默返回空 steps 当作"不需要工艺"**）

第 2 级必须在输出里可辨认（Route Cost Basis 写明"按件名+材料推断，无 attrs 证据"），
否则读者会把它当成第 1 级的精度。

产品无关性
----------
本模块不出现任何产品名、不使用任何产品专有参数。功能域前缀（INV/DCDC/OBC）与工艺模板
的费率键（RATE_*）都是**入参**，由生成器从 `1-Assumptions` 与 `--product` 传入。
新增产品若命中不了任何规则，应当落到第 3 级并被看见，而不是被硬编码成某个产品的默认值。

（历史坑：v2 的 `function_category(..., prefix="DCDC")` 把 DC-DC 当成了默认值，
逆变器 53 个叶件在没有 attrs 时全部被打上 DCDC 前缀。）
"""
from __future__ import annotations

import re

# ================================================================ 产品 → 功能域前缀
# 功能域前缀只用于 2-BOM 的 Function 列分组，不影响任何金额。
# 直接由 --product 推导，也允许 --function-prefix 显式覆盖（产品名不含关键词时用）。
PREFIX_RULES: list[tuple[tuple[str, ...], str]] = [
    (("dc-dc", "dcdc", "dc dc"), "DCDC"),
    (("obc", "on-board charger", "车载充电"), "OBC"),
    (("inverter", "逆变"), "INV"),
    (("pdu", "power distribution"), "PDU"),
    (("bms", "battery management"), "BMS"),
]


def func_prefix_for(product: str, override: str | None = None) -> str:
    """从产品名推导功能域前缀。无法识别时返回 "GEN"（通用），**不猜成某个产品**。"""
    if override:
        return override.strip().upper()
    p = (product or "").lower()
    for keys, prefix in PREFIX_RULES:
        if any(k in p for k in keys):
            return prefix
    return "GEN"


# ================================================================ 外购判定
# 必须锚定名称开头，否则 "Cover, Rear Inverter PCB" 会被 "PCB" 误判为外购。
# 命中即为外购（Buy）；不命中即为自制候选（Make）。
# 注意：这只是"候选"。真正的 Make/Buy 由生成器的口径决定（v2 的三层口径把所有叶件
# 都当 Purchased）；但材料成本、工艺路由的建立只对 Make 候选有意义。
BUY_RE = re.compile(
    r"^(PCB\b|PCBA\b|SH Film Capacitor|Film Capacitor|Capacitor|MOSFET|IGBT|Diode|"
    r"Ferrite Core|Connector|Fuse|Sensor|Jumper Harness|Thermal Paste|Breather|"
    r"Inductor|Transformer|Gasket|O Ring|O-Ring|NVH Foam|Spring)",
    re.IGNORECASE,
)


def is_buy(name: str) -> bool:
    """件名 → 是否外购候选。锚定开头，大小写不敏感。"""
    return bool(BUY_RE.match((name or "").strip()))


# ================================================================ 功能域分类
# 两级切割：Function（电气功能域，Box 为结构/集成兜底）→ Category（域内类别）。
# 边界公式（防口径漂移，改动须同步 SKILL.md 与 workbook-spec.md）：
#   <product> function = PM 功率级 + DC-Link 电容 + 电流传感 + PCB / PCBA + 功率磁件 + E-Parts
#   Box      function = 壳体 + 冷却与热 + HV 连接 + 信号连接 + EMC + 密封紧固其他
# 原则：
#   * E-Parts 独立于 PCB：电子件是降本主战场，混进 PCB 会被大数淹没；
#   * 冷却板 / TIM 归 Box 不归功率级：热管理是结构件成本，不是器件成本；
#   * 分支顺序有意义：access cover 必须早于通用 cover，coolant adapter 早于通用 adapter。

def function_category(name: str, species: str = "",
                      func_prefix: str = "GEN") -> tuple[str, str]:
    """件名 + 材料牌号 → (Function, Category)。产品无关，纯字典规则。

    匹配顺序（改动须同步 SKILL.md 与 workbook-spec.md）：
        0. 强信号看全文（coolant）：它可能恰好是父级上下文的语义
           ——"Adapter, Rear Inverter Coolant Inlet" 的 coolant 在上下文里，
             但该件确实是冷却件，必须看得见。
        1. 其余规则只看 head（件名主体），避免父级上下文盖掉零件类型。
           电流传感也适用本条：`Ferrite Core, … Current Sensor 2` 是 EMC 件、
           `PCB, … Current Sensor 2` 是 PCB，都不是传感器本体；只有 head 本身
           是 sensor 才算电流传感。
    """
    h, full = name_subject(name)
    s = (species or "").lower()

    # ---- 0. 强信号 ----
    # 电流传感：必须 head 是传感器本体，否则父件上下文会把 Ferrite/PCB/Cover 全拖进来
    if "sensor" in h or h.startswith("current sensor"):
        return func_prefix, "电流传感"
    # 冷却：coolant 允许看全文（"Adapter, … Coolant Inlet" 的主体是 adapter，
    # 但它在功能上就是冷却接口，归冷却与热而非密封紧固其他）
    if "coolant" in full or "cold plate" in full:
        return "Box", "冷却与热"

    # ---- 1. 电气功能域（只看 head）----
    if ("mosfet" in h or "igbt" in h or "power module" in h or "power stage" in h):
        return func_prefix, "PM 功率级"
    if "film capacitor" in h or ("capacitor" in h and "sh" in h):
        return func_prefix, "DC-Link 电容"
    if h.startswith("pcb") or h.startswith("pcba") or "pcb assembly" in h:
        return func_prefix, "PCB / PCBA"
    if "inductor" in h or "transformer" in h or "choke" in h:
        return func_prefix, "功率磁件"
    if "fuse" in h or "breather" in h or "terminal" in h:
        return func_prefix, "E-Parts"

    # ---- 2. Box：壳体（access cover 必须在通用 cover 之前）----
    if ("housing" in h or "casing" in h or "bottom cover" in h
            or "top cover" in h or "access cover" in h):
        return "Box", "壳体"

    # ---- 3. Box：冷却与热 ----
    if ("cooling plate" in h or "cooling pipe" in h or "thermal paste" in h
            or h.strip() == "tim"):
        return "Box", "冷却与热"

    if "busbar" in h:
        return "Box", "HV 连接"
    if "connector" in h or "jumper harness" in h:
        return "Box", "信号连接"
    if "ferrite core" in h:
        return "Box", "EMC"

    # ---- 4. Box：密封紧固其他（含紧固件，按件名判定，不依赖材料牌号）----
    if (h == "cap" or "gasket" in h or "o ring" in h or "o-ring" in h
            or "nvh foam" in h or "spacer" in h or "locator" in h or "seal" in h
            or "bracket" in h or "holder" in h or "cover" in h or "plug" in h
            or "spring" in h or "adapter" in h or "clip" in h or "retainer" in h
            or any(k in h for k in ("screw", "bolt", "nut", "washer", "rivet",
                                    "stud", "clamp", "tie"))):
        return "Box", "密封紧固其他"

    # ---- 5. 兜底：按材料属性归 Box（金属件多为结构/密封/紧固）----
    if (("copper" in s or "aluminium" in s or "aluminum" in s or "steel" in s)
            and "plastic" not in s and "elastomer" not in s):
        return "Box", "密封紧固其他"

    # 无任何命中：归 E-Parts，调用方应把这种行当作"分类待确认"。
    return func_prefix, "E-Parts"


def is_fallback_category(cat: str) -> bool:
    """是否落到了兜底类别（用于自检：这些行的分类未经证据确认）。"""
    return cat in ("E-Parts", "密封紧固其他")


def name_subject(name: str) -> tuple[str, str]:
    """拆件名 → (head, full)。

    CS 件名结构是 `<零件类型>, <父级上下文>`：
        "Busbar, Rear Inverter"                    → head="busbar"
        "Gasket, Rear Inverter Power Module"       → head="gasket"
        "Adapter, Rear Inverter Coolant Inlet"     → head="adapter"
        "Top Cover, DC-DC Converter"               → head="top cover"

    分类**必须优先看 head**。若全文匹配，父级上下文会把零件类型盖掉：
        "Gasket, ... Power Module" 因为上下文里有 "Power Module" 被判成 PM 功率级，
        而 "Gasket, ... Terminal" 被判成 E-Parts —— 同学一个垫片，三种归属。

    这条规则同时解决了产品间不一致：同一零件类型在不同产品里 head 相同，
    因此分类必然相同（跨产品不变量，见 tests/test_workbook_generic.py）。
    只有**强信号**（coolant / current）才允许看全文，因为它们可能正是父级上下文的语义。
    """
    n = (name or "").strip()
    head = n.split(",")[0].strip().lower()
    return head, n.lower()


# ================================================================ 工艺模板
# 工序元组：(工序, 类别, 节拍s, 费率参数ID, 人数, 耗材, 步骤良率)
# 外协行：费率参数为 None，耗材为 ("outsrc", 单价参数ID, 基数)——按外协价×基数计价。
# 这些模板是**通用工艺族**，不含任何产品专有数值；节拍与良率均为模板假设，
# 生成器须在 1-Assumptions 里登记为 Estimated。
CAST: list[tuple] = [
    ("压铸", "加工", 60, "RATE_CAST", 0.5, 0.0, 0.95),
    ("修边", "加工", 20, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("CNC 加工", "加工", 120, "RATE_CNC", 1.0, 3.0, 0.98),
    ("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99),
    ("检漏", "测试", 30, "RATE_TEST", 1.0, 2.0, 0.99),
]
STAMP: list[tuple] = [
    ("冲压", "加工", 8, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("去毛刺", "加工", 15, "RATE_ASSY", 1.0, 0.0, 0.99),
    ("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99),
]
PLATE: list[tuple] = [
    ("冲压", "加工", 10, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("平面磨", "加工", 60, "RATE_CNC", 1.0, 2.0, 0.97),
    ("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99),
]
DIE_CUT: list[tuple] = [("模切", "加工", 10, "RATE_STAMP", 1.0, 0.5, 0.98)]
INJ: list[tuple] = [
    ("注塑", "加工", 45, "RATE_INJ", 1.0, 0.0, 0.97),
    ("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99),
]
INJ_INSERT: list[tuple] = INJ + [("嵌件装配", "装配", 40, "RATE_ASSY", 1.0, 0.5, 0.98)]
OVERMOLD: list[tuple] = [
    ("注塑包胶", "加工", 50, "RATE_INJ", 1.0, 0.0, 0.96),
    ("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99),
]
BUSBAR: list[tuple] = [
    ("冲压", "加工", 12, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("电镀（外协）", "加工", 0, None, 0.0, ("outsrc", "PLATE_PRICE", "weight"), 0.98),
    ("清洗", "加工", 10, "RATE_CNC", 0.5, 1.0, 0.99),
]
ASSY: list[tuple] = [("装配", "装配", 60, "RATE_ASSY", 1.0, 0.5, 0.98)]

# 路由名 → 工序模板（供 route_text 与 steps 一一对应）
ROUTE_TEMPLATES: dict[str, list[tuple]] = {
    "压铸 → 修边 → CNC → 清洗 → 检漏": CAST,
    "冲压 → 去毛刺 → 清洗": STAMP,
    "冲压 → 平面磨 → 清洗": PLATE,
    "模切": DIE_CUT,
    "注塑 → 去浇口": INJ,
    "注塑 → 去浇口 → 嵌件装配": INJ_INSERT,
    "冲压 → 注塑包胶 → 去浇口": OVERMOLD,
    "冲压 → 电镀 → 清洗": BUSBAR,
    "装配": ASSY,
}


# ================================================================ 工艺路由推断
# 塑料判定不能只查子串 "pa"（"pak"、"part" 等会误命中，且 "pa" 不是独立词）。
# 统一的塑料牌号判定，供 route_for 使用。
_PLASTIC_KEYS = ("plastic", "polymer", "pbt", "pps", "ppa", "pa66", "pa6", "pc-abs",
                 "pc/abs", "abs", "pom", "polyamide", "polyester", "pp ", "pp-")


def _is_plastic(species_lower: str) -> bool:
    return any(k in species_lower for k in _PLASTIC_KEYS)


def route_for(name: str, species: str = "") -> tuple[str, list[tuple]]:
    """件名 + 材料牌号 → (路由文本, 工序模板)。第 2 级降级用，零 attrs 依赖。

    返回 ("装配", ASSY) 表示"只能判到装配级"，属于**弱推断**；
    调用方须把这种情况标为 route_status="Inferred-Weak"，不要当成已建立的工艺。
    """
    n = (name or "").strip().lower()
    s = (species or "").lower()
    plastic = _is_plastic(s)

    # 壳体必须先判（"bottom cover" 含 "cover"，顺序反了会被当钣金件）
    if "housing" in n or "casing" in n:
        return "压铸 → 修边 → CNC → 清洗 → 检漏", CAST
    if "bottom cover" in n:
        return "压铸 → 修边 → CNC → 清洗 → 检漏", CAST
    if "top cover" in n or "access cover" in n:
        # 压铸 vs 钣金按牌号分界：浅腔薄板无压铸牌号 → 冲压。
        die_cast = ("aluminium" in s or "aluminum" in s) and any(
            k in s for k in ("ac44300", "ac43500", "enac", "adc", "alsi", "a380", "a360"))
        return (("压铸 → 修边 → CNC → 清洗 → 检漏", CAST) if die_cast
                else ("冲压 → 去毛刺 → 清洗", STAMP))
    if "cooling plate" in n or "cold plate" in n or ("plate" in n and "cooling" in n):
        return "冲压 → 平面磨 → 清洗", PLATE
    if "cooling pipe" in n:
        return "注塑 → 去浇口", INJ
    if n == "cap" and plastic:
        return "注塑 → 去浇口", INJ
    if "gasket" in n or "o ring" in n or "o-ring" in n or "seal" in n:
        return "模切", DIE_CUT
    if "nvh foam" in n or "foam" in n:
        return "模切", DIE_CUT
    if "busbar with ferrite" in n:
        return "冲压 → 注塑包胶 → 去浇口", OVERMOLD
    if "busbar" in n:
        return "冲压 → 电镀 → 清洗", BUSBAR
    if "terminal" in n or "adapter" in n:
        return "注塑 → 去浇口 → 嵌件装配", INJ_INSERT
    if "bracket" in n or "holder" in n:
        return ("注塑 → 去浇口", INJ) if plastic else ("冲压 → 去毛刺", STAMP)
    if "cover" in n or "plug" in n or "clip" in n or "retainer" in n:
        return ("注塑 → 去浇口", INJ) if plastic else ("冲压 → 去毛刺", STAMP)
    if "spacer" in n:
        return "注塑 → 去浇口", INJ
    if "locator" in n:
        return ("注塑 → 去浇口", INJ) if plastic else ("冲压 → 去毛刺", STAMP)
    return "装配", ASSY


def route_strength(route_text: str) -> str:
    """路由推断强度。装配级 = 弱推断（说明该件是总成件或件名无法判族）。"""
    return "Weak" if route_text == "装配" else "Normal"


def to_step_dicts(steps: list[tuple], basis: str = "模板") -> list[dict]:
    """把工序元组转成生成器用的 step dict（v2 的 step() 同构）。

    * rate 为 None → 外协工序：cycle 置 None（无设备节拍），cons 保留 ("outsrc", 参数ID, 基数)；
    * 其余 → cycle 为固定秒数，cons 为固定耗材额。
    """
    out: list[dict] = []
    for proc, cat, cycle, rate, ops, cons, yld in steps:
        if rate is None:
            out.append({"proc": proc, "cat": cat, "cycle": None, "rate": None,
                        "ops": 0.0, "cons": cons, "yld": yld,
                        "basis": "外协（无设备节拍）"})
        else:
            out.append({"proc": proc, "cat": cat, "cycle": cycle, "rate": rate,
                        "ops": ops, "cons": cons, "yld": yld, "basis": basis})
    return out