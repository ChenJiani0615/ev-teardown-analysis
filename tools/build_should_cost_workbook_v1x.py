#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Should-Cost 工作簿生成器 —— **v1x**（PPC = MAT + VA，逐件 M1–M5 计价）。

与 `build_should_cost_workbook_v2.py` 的关系：两者并存，**互不改动**。本文件按
`should-cost-engineering-v1x/SKILL.md` 的 Step 10 重建。

| 项 | v2（旧） | v1x（本文件） |
|---|---|---|
| 价格标签 | `Price Class A–E` + `Price Status` 八类 + `3a/3b/3c` | **`Price Method` M1–M5** |
| 计入判定 | 无 | **`Counted In Total` / `Counted Amount` 两列** |
| 区间 | `Price Base` 一个数 | **`Unit Price(=Base)` + `Price Low` + `Price High` 并存** |
| 汇总 | 全部行相加 | **只对 `Counted Amount` 求和** |
| 结论 | 无 | **`4-Summary` 首行结论句 + 已锚定/全口径双数** |
| 导出 | 无 | **`PPT_EXPORT` 固定块带 `completeness`** |

三条口径决定（写在这里，避免读者自己猜）
=====================================
1. **MAT 与 VA 不重叠。** 外购件（M1/M3/M4/M5）取的是**交付价**，已含供应商加工，
   所以这些行不再生成工序行；自制件（M2）取的是**材料成本**，加工放到 `3-ProcessCost`
   的 PDC。同一笔钱只出现一次。
2. **质量轴价按交付价处理。** 目录 `mass_axis_*` 是 `CNY/kg` 的成品交付价基准。
   件名匹配不上时按材料族回退到它，得到的是完整单价，因此该行 `complete=True`、不再算 PDC。
   若把它也当材料成本再加 PDC，会重复计价。
3. **VA 混合口径。** `VA = PDC已知 + (MAT未建工艺部分 × k) + MOH`。
   `k` 来自 L2 校准比例，只作用于没建工序的那部分 MAT，避免与 PDC 重叠。

工程约束（吃过亏，逐条有出处）
============================
* 空价必须是空字符串 `""`，**不许写 0**：写 0 会把「缺价」伪装成「零成本」。
* 公式长度 < 8192 字符（曾出现 14995 字符单元格 → Excel 拒开，openpyxl 却读得干净）。
* 整列引用必须 `$T:$T`（带冒号）；`$T$T` 非法，Excel 拒开。
* 跨表引用带表名，列字母只从本文件 `make_cols()` 取。
* 说明文字以 `=` 开头会把 Excel 判文件非法 → 统一转全角 `＝`。
* 透视块必须配一行闭合差公式，这是唯一能抓到口径混算的机制。

用法
----
    py -3 -X utf8 tools/build_should_cost_workbook_v1x.py ^
        --bom     data/<product>/derived/<product>-leaves.csv ^
        --attrs   data/<product>/derived/<product>-attributes.csv ^
        --prices  data/<product>/derived/<product>-price-evidence.csv ^
        --catalog data/benchmark/benchmark-price-catalog-v1x.csv ^
        --product "EX90 rear inverter" ^
        --out     reports/EX90_RearInverter_ShouldCost_v1x.xlsx
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from catalog_v1x import load_catalog, match_price  # noqa: E402
from domain_classify import (func_prefix_for, function_category, is_buy,  # noqa: E402
                             route_for, route_strength)

# ================================================================ 样式
HDR_FILL = PatternFill("solid", fgColor="305496")
HDR_FONT = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=9)
BOLD_FONT = Font(name="微软雅黑", size=9, bold=True)
TITLE_FONT = Font(name="微软雅黑", size=13, bold=True, color="1F3864")
SEC_FONT = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
NOTE_FONT = Font(name="微软雅黑", size=9, color="595959")
WARN_FONT = Font(name="微软雅黑", size=9, italic=True, color="C00000")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STATUS_FILL = {
    "Verified": None,
    "Calculated": PatternFill("solid", fgColor="DDEBF7"),
    "Estimated": PatternFill("solid", fgColor="FFF2CC"),
    "Scenario": PatternFill("solid", fgColor="FCE4D6"),
    "Unknown": PatternFill("solid", fgColor="FCE4E4"),
}
MONEY = '#,##0.00'
KG = '0.000'
RATIO = '0.0000'
PCT = '0.0%'

# ================================================================ 参数表
# (ID, 名称, 值, 单位, 状态, 来源/说明)
PARAMS: list[tuple] = [
    # ---- 口径声明（Step 1）
    ("REPORT_DATE", "报告日期", "2026-11-20", "-", "Verified", "生成日"),
    ("PRICE_BASE", "价格基期", 2025, "年", "Estimated", "目录与报告均 @2025；跨年调整须写依据，无依据不调"),
    ("TARGET_YEAR", "目标年份", 2025, "年", "Estimated", "与价格基期一致时无需折现"),
    ("SCENARIO", "情景", "现状", "-", "Estimated", "竞品默认『现状』；自家产品默认『含全部措施』"),
    ("CUR", "币种", "CNY", "-", "Estimated", "目录与报告均为 CNY"),
    ("SRC_VER", "来源版本", "Bosch CTM/P2/PTU + v1x 目录", "-", "Estimated", "每个数字须带报告+页码"),
    ("BOUNDARY_SW", "边界归一开关", "关", "-", "Estimated",
     "默认关。仅当被比较对象同一功能一自制一外购时打开，把自制部分的 VA 并入 MAT 并单列披露"),
    # ---- 产量
    ("VOL_A", "年产量", 20000, "件/年", "Unknown", "待项目输入；影响折旧与摊销"),
    ("VOL_L", "生命周期产量（合格）", 100000, "件", "Unknown", "待项目输入；摊销分母"),
    # ---- 材料单价（只用于 M2 材料影子，不用于目录补位）
    ("AL_PRICE", "铝压铸合金单价", 23.9, "CNY/kg", "Estimated", "长江现货 A00 @2026-10-28"),
    ("CU_PRICE", "铜单价", 110.5, "CNY/kg", "Estimated", "长江现货 1# 电解铜 @2026-10-28"),
    ("ST_PRICE", "钢单价", 8.0, "CNY/kg", "Estimated", "冷轧/镀锌板"),
    ("PLASTIC_PRICE", "通用工程塑料单价", 25.0, "CNY/kg", "Estimated", "未单列牌号时用"),
    ("FERRITE_PRICE", "铁氧体单价", 30.0, "CNY/kg", "Estimated", "烧结磁芯"),
    ("ELASTOMER_PRICE", "弹性体单价", 32.0, "CNY/kg", "Estimated", "TPV/TPS"),
    ("FOAM_PRICE", "PE 泡棉单价", 15.0, "CNY/kg", "Estimated", "模切片"),
    ("ADHESIVE_PRICE", "胶粘剂单价", 60.0, "CNY/kg", "Estimated", "结构胶/密封胶同价"),
    # ---- 利用率与废料
    ("SCRAP_REC", "废料回收率", 0.60, "-", "Estimated", "★ 只对金属生效"),
    ("Y_CAST", "压铸材料利用率", 0.85, "-", "Estimated", "含浇口/溢流"),
    ("Y_STAMP", "冲压材料利用率", 0.72, "-", "Estimated", "含排样废料"),
    ("Y_INJ", "注塑材料利用率", 0.92, "-", "Estimated", "含水口回用"),
    # ---- 设备与人工费率（工序表用）
    ("RATE_CAST", "压铸机综合费率", 180.0, "CNY/h", "Estimated", "全包费率；口径须与报告一致"),
    ("RATE_STAMP", "冲压机综合费率", 95.0, "CNY/h", "Estimated", "同上"),
    ("RATE_CNC", "CNC / 清洗综合费率", 110.0, "CNY/h", "Estimated", "同上"),
    ("RATE_INJ", "注塑机综合费率", 120.0, "CNY/h", "Estimated", "同上"),
    ("RATE_ASSY", "装配工位综合费率", 75.0, "CNY/h", "Estimated", "同上"),
    ("RATE_TEST", "测试工位综合费率", 90.0, "CNY/h", "Estimated", "同上"),
    ("PLATE_PRICE", "电镀外协单价", 22.0, "CNY/kg", "Estimated", "按净重计价"),
    # ---- 加价链（仅 M2 零件级计算器）
    ("OH_RATE", "供应商制造间接费率", 0.12, "-", "Estimated", "仅用于 M2"),
    ("SGA_RATE", "供应商 SG&A 率", 0.05, "-", "Estimated", "仅用于 M2"),
    ("LOG_RATE", "物流包装率", 0.01, "-", "Estimated", "只加这一段"),
    ("PROFIT_RATE", "供应商利润率（按销售额）", 0.08, "-", "Estimated",
     "Price = Cost ÷ (1 − margin)，不是加成"),
    # ---- VA（Step 5）
    ("K_VA_LOW", "VA/MAT 比例 k（Low）", 0.080, "-", "Estimated", "校准文件 B 节：k 8.0%–12.2%"),
    ("K_VA_BASE", "VA/MAT 比例 k（Base）", 0.096, "-", "Estimated", "中位 9.6%，n=6"),
    ("K_VA_HIGH", "VA/MAT 比例 k（High）", 0.122, "-", "Estimated", ""),
    ("MOH_RATIO", "MOH/MAT", 0.012, "-", "Estimated",
     "★ 只用于把 VA 拆成 PDC 与 MOH，不得在 VA 之上再叠一层"),
    # ---- 隐含件比例（Step 3）
    ("S_CHEM", "隐含件比例：Chem+Scrap", 0.0395, "-", "Estimated", "校准文件 C 节 3.4%–4.5% 取中点"),
    ("S_CHEM_LOW", "隐含件比例：Chem+Scrap（Low）", 0.034, "-", "Estimated", ""),
    ("S_CHEM_HIGH", "隐含件比例：Chem+Scrap（High）", 0.045, "-", "Estimated", ""),
    ("S_FAST", "隐含件比例：紧固件", 0.0115, "-", "Estimated", "0.7%–1.6% 取中点"),
    ("S_FAST_LOW", "隐含件比例：紧固件（Low）", 0.007, "-", "Estimated", ""),
    ("S_FAST_HIGH", "隐含件比例：紧固件（High）", 0.016, "-", "Estimated", ""),
    # ---- 阈值（Step 7）
    ("TH_M5_L1", "结论等级 1 的非锚定占比上限", 0.15, "-", "Estimated", "★ 经验设定，待校准"),
    ("TH_M5_L3", "结论等级 3 的非锚定占比下限", 0.40, "-", "Estimated", "★ 经验设定，待校准"),
    ("RANGE_MAX_RATIO", "区间可用上限（High/Low）", 1.5, "倍", "Estimated",
     "≤ 取均值；> 取几何均值并标『区间暂估』"),
    ("WIDE_RATIO", "宽区间阈值（High/Low）", 10.0, "倍", "Estimated", "> 标 [宽>10×] 并进待取证前 5"),
    ("MASS_TOL", "重量闭合容差", 0.20, "-", "Estimated", "> 20% 先复核大件"),
    # ---- 工装
    ("TOOL_LIFE", "工装寿命（模次）", 500000, "模次", "Estimated", "摊销基数 = MIN(寿命, 生命周期产量)"),
    ("DEP_YEARS", "设备折旧年限", 5.0, "年", "Estimated", "直线折旧；只计一次"),
]
PV = {p[0]: p[2] for p in PARAMS}


# ================================================================ 列定义（单一真源）
def make_cols(spec: list[tuple[str, str, int]]) -> tuple[list[str], list[int], dict[str, str]]:
    keys = [k for k, _, _ in spec]
    return ([lab for _, lab, _ in spec], [w for _, _, w in spec],
            {k: get_column_letter(i) for i, k in enumerate(keys, start=1)})


BOM_SPEC = [
    ("level", "Level", 6), ("part", "Part No", 16), ("name", "Name / Model", 30),
    ("rowtype", "Row Type", 9), ("func", "Function", 7), ("cat", "Category", 13),
    ("makebuy", "Make/Buy", 9), ("bucket", "Cost Bucket", 13),
    ("qty", "Qty per Unit", 8), ("netwt", "Net Weight (kg)", 10),
    ("qty_basis", "Quantity Basis", 13), ("qty_ev", "Quantity Evidence", 22),
    ("price", "Unit Price (= Base)", 12), ("p_low", "Price Low", 10),
    ("p_high", "Price High", 10), ("method", "Price Method", 9),
    ("price_basis", "Unit Price Basis", 58), ("price_ev", "Price Evidence", 22),
    ("fit", "Fit Check", 13), ("price_date", "Price Date", 10),
    ("purch_qty", "Purchase Qty", 10), ("deliv", "Delivery Term", 14),
    ("tooling_in", "Tooling In Price", 9),
    ("counted", "Counted In Total", 9), ("counted_amt", "Counted Amount", 13),
    ("counted_note", "Counted Note", 30),
    ("amt_low", "Amt Low", 10), ("amt_high", "Amt High", 10),
]

PROC_SPEC = [
    ("step", "Step", 6), ("part", "Part No", 16), ("proc", "Process", 20),
    ("cat", "Category", 8), ("bearer", "Cost Bearer", 10), ("site", "Site", 10),
    ("cycle", "Cycle (s)", 10), ("cycle_basis", "Cycle Basis", 14),
    ("rate", "Machine Rate", 11), ("rate_basis", "Rate Basis", 20),
    ("cons", "Consumable", 11), ("yld", "Step Yield", 9),
    ("direct", "Step Direct (CNY)", 13), ("cumul", "Cumulative Cost (CNY)", 14),
    ("is_last", "Is Last Step (Y/N)", 11), ("strength", "Route Strength", 12),
]

VA_SPEC = [
    ("item", "VA 项", 32), ("block", "分块", 13),
    ("low", "金额 Low", 12), ("base", "金额 Base", 12), ("high", "金额 High", 12),
    ("level", "方法层级", 9), ("site", "Site", 11),
    ("basis", "依据", 52), ("src", "来源", 26),
]

OPP_SPEC = [
    ("id", "ID", 6), ("title", "题目", 24), ("now", "现有方案", 22),
    ("alt", "替代方案", 22), ("mech", "工程依据", 26),
    ("gross", "毛潜力 (CNY/pc)", 13), ("effort", "Effort (CNY/pc)", 13),
    ("net", "净潜力 (CNY/pc)", 12),
    ("mat_imp", "MAT Impact", 10), ("va_imp", "VA Impact", 10),
    ("tool_imp", "Tooling Impact", 11),
    ("in_dtc", "已在 DtC", 9), ("in_rpp", "已在 RPP", 9),
    ("phase", "Phase 增量", 11), ("mech_tag", "降本机制", 14),
    ("dedup", "防重复计数组", 12), ("hg", "HG", 6),
    ("conf", "证据可信度", 13), ("risk", "风险", 18),
    ("weeks", "验证周期", 9), ("status", "状态", 22), ("prio", "优先级", 8),
]

SENS_SPEC = [
    ("drv", "驱动因素", 26), ("pbase", "基准值", 16),
    ("low", "Low", 16), ("base", "Base", 16), ("high", "High", 18),
    ("impact", "影响额（绝对值）", 15), ("share", "影响占比", 10),
    ("dir", "方向", 18), ("note", "说明", 46), ("src", "来源参数", 22),
]

EVD_SPEC = [
    ("item", "项", 20), ("content", "内容", 56), ("src", "来源/位置", 30),
    ("status", "状态", 11), ("conf", "置信度", 9),
    ("impact", "影响成本项", 16), ("todo", "待确认问题", 42), ("owner", "负责人", 10),
]


# ================================================================ 读数据
SPLIT_COLS = ["AL (kg)", "CU (kg)", "STEEL (kg)", "PLASTIC (kg)", "ELASTOMER (kg)",
              "FERRITE (kg)", "FOAM (kg)", "ADHESIVE (kg)", "OTHER (kg)"]
SPLIT_PARAM = {"AL (kg)": "AL_PRICE", "CU (kg)": "CU_PRICE", "STEEL (kg)": "ST_PRICE",
               "PLASTIC (kg)": "PLASTIC_PRICE", "ELASTOMER (kg)": "ELASTOMER_PRICE",
               "FERRITE (kg)": "FERRITE_PRICE", "FOAM (kg)": "FOAM_PRICE",
               "ADHESIVE (kg)": "ADHESIVE_PRICE", "OTHER (kg)": "PLASTIC_PRICE"}


def _f(v) -> float | None:
    v = (v or "").strip() if isinstance(v, str) else ""
    try:
        return float(v) if v else None
    except ValueError:
        return None


def load_leaves(path: Path) -> list[dict]:
    """读叶件清单。允许 Total Weight 为空（标准件：质量已被父件合并，这是正确口径）。"""
    out = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            qty = _f(r.get("Quantity")) or 1.0
            tw = _f(r.get("Total Weight (kg)"))
            pw = _f(r.get("Part Weight"))
            if pw is None:
                pw = (tw or 0.0) / qty if qty else (tw or 0.0)
            out.append({
                "part": (r.get("CS Part") or "").strip(),
                "name": (r.get("Name") or "").strip(),
                "qty": qty,
                "net_weight": pw,
                "total_weight": tw if tw is not None else pw * qty,
                "group": (r.get("Material Group") or "").strip(),
                "species": (r.get("Material Species") or "").strip(),
                "length": _f(r.get("Length")) or 0.0,
                "width": _f(r.get("Width")) or 0.0,
                "itf": _f(r.get("Number of Interface")) or 0.0,
                "ptype": (r.get("Part Type") or "").strip(),
                "parent": (r.get("Parent Part No") or "").strip(),
            })
    return out


def load_attrs(path: Path | None) -> dict[str, dict]:
    """读属性表。

    ★ 除已知列外的**任意附加列**都按小写列名收集进 `attrs`，供目录
      `Required Attribute` 匹配（如 `layers=4` / `thread=M5` / `vds_V=1200`）。
      这样目录加子档不必改生成器 schema——否则子档永远命中不了（曾是这样）。
    """
    if not path or not path.exists():
        return {}
    KNOWN = {"CS Part", "Procurement Object", "Bottom-up Model (Y/N)",
             "Material Split Basis", "Notes", *SPLIT_COLS}
    out: dict[str, dict] = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        rdr = csv.DictReader(f)
        extra = [c for c in (rdr.fieldnames or []) if c and c not in KNOWN]
        for r in rdr:
            part = (r.get("CS Part") or "").strip()
            if not part:
                continue
            a: dict[str, object] = {}
            for c in extra:
                v = (r.get(c) or "").strip()
                if not v:
                    continue
                try:
                    a[c.strip().lower()] = float(v)
                except ValueError:
                    a[c.strip().lower()] = v
            out[part] = {
                "object": (r.get("Procurement Object") or "").strip(),
                "bottomup": (r.get("Bottom-up Model (Y/N)") or "").strip().upper() == "Y",
                "splits": {c: (_f(r.get(c)) or 0.0) for c in SPLIT_COLS},
                "split_basis": (r.get("Material Split Basis") or "").strip(),
                "notes": (r.get("Notes") or "").strip(),
                "attrs": a,
            }
    return out


def load_prices(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            p = (r.get("CS Part") or "").strip()
            if not p:
                continue
            out[p] = {
                "low": _f(r.get("Price Low (CNY)")),
                "high": _f(r.get("Price High (CNY)")),
                "basis": (r.get("Price Basis") or "").strip(),
                "date": (r.get("Source Date") or "").strip(),
                "deliv": (r.get("Delivery Term") or "").strip(),
                "purch_qty": (r.get("Purchase Qty") or "").strip(),
                "tooling": (r.get("Tooling In Price") or "?").strip() or "?",
            }
    return out


def load_opps(path: Path | None) -> list[dict]:
    if not path or not path.exists():
        return []
    out = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("title") or "").strip():
                continue
            out.append({
                "id": (row.get("id") or "").strip(),
                "title": (row.get("title") or "").strip(),
                "now": (row.get("now") or "").strip(),
                "alt": (row.get("alt") or "").strip(),
                "mech": (row.get("mech") or "").strip(),
                "gross": _f(row.get("gross")),
                "effort": _f(row.get("effort")),
                "mat_imp": (row.get("mat_imp") or "无").strip() or "无",
                "va_imp": (row.get("va_imp") or "无").strip() or "无",
                "tool_imp": (row.get("tool_imp") or "无").strip() or "无",
                "in_dtc": (row.get("in_dtc") or "").strip(),
                "in_rpp": (row.get("in_rpp") or "").strip(),
                "phase": (row.get("phase") or "").strip(),
                "mech_tag": (row.get("mech_tag") or "").strip(),
                "dedup": (row.get("dedup") or "").strip(),
                "hg": (row.get("hg") or "HG0").strip(),
                # ★ 外部条目一律 Hypothetical：不得从 CSV 带进 Verified
                "conf": "Hypothetical",
                "risk": (row.get("risk") or "").strip(),
                "weeks": (row.get("weeks") or "").strip(),
                "status": "Hypothetical Opportunity",
                "prio": (row.get("prio") or "中").strip(),
            })
    return out


# ================================================================ 取值规则
def pick_base(low: float | None, high: float | None) -> tuple:
    """≤ RANGE_MAX_RATIO 取均值；> 取几何均值并标『区间暂估』；> WIDE_RATIO 标 [宽>10×]。"""
    if not low or not high:
        return (low or high), ""
    ratio = high / low
    tag = ""
    if ratio > PV["WIDE_RATIO"]:
        tag = "[宽>10×]"
    elif ratio > PV["RANGE_MAX_RATIO"]:
        tag = "区间暂估"
    base = (low * high) ** 0.5 if ratio > PV["RANGE_MAX_RATIO"] else (low + high) / 2
    return base, tag


MASS_AXIS = (("metal", "mass_axis_metal"), ("aluminium", "mass_axis_metal"),
             ("aluminum", "mass_axis_metal"), ("copper", "mass_axis_metal"),
             ("steel", "mass_axis_metal"), ("plastic", "mass_axis_plastic"),
             ("polymer", "mass_axis_plastic"), ("pbt", "mass_axis_plastic"),
             ("pps", "mass_axis_plastic"), ("elastomer", "mass_axis_elastomer"),
             ("rubber", "mass_axis_elastomer"), ("tpv", "mass_axis_elastomer"),
             ("epdm", "mass_axis_elastomer"), ("foam", "mass_axis_foam"),
             ("pur", "mass_axis_foam"))


def mass_axis_key(group: str, species: str) -> str | None:
    txt = f"{group} {species}".lower()
    for k, key in MASS_AXIS:
        if k in txt:
            return key
    return None


# ★ 主材族 → 材料专属交付价目录 key（质量轴兜底用，看 species **首段**判定）。
#   拆解清单的 species 是「主材在前、嵌件/紧固件/密封胶在后」的逗号串：
#   "Copper Unknown,Steel Unknown" = 铜母线 + 钢紧固件，主材是铜；
#   "Plastic PBT-GF30,Brass Unknown,Aluminium Unknown" = 塑料支架 + 黄铜/铝嵌件。
#   故按首段**第一个 token**（纯等值查表，非子串）判定主材族，绝不做"整串首个关键字命中"。
_PRIMARY_TOKEN_TO_KEY = {
    # 金属（交付价档，非材料价：铜排 79-97 / 铝压铸 38-73 / 钢钣金 25-40）
    "copper": "mass_axis_copper_busbar",
    "brass": "mass_axis_copper_busbar",
    "aluminium": "mass_axis_aluminium_diecast",
    "aluminum": "mass_axis_aluminium_diecast",
    "steel": "mass_axis_steel_sheetmetal",
    "stainless": "mass_axis_steel_sheetmetal",
    # 塑料
    "plastic": "mass_axis_plastic",
    "polymer": "mass_axis_plastic",
    "pbt": "mass_axis_plastic",
    "pps": "mass_axis_plastic",
    "ppa": "mass_axis_plastic",
    "pa66": "mass_axis_plastic",
    "pa6": "mass_axis_plastic",
    "pom": "mass_axis_plastic",
    "para": "mass_axis_plastic",
    # 弹性体
    "elastomer": "mass_axis_elastomer",
    "epdm": "mass_axis_elastomer",
    "aem": "mass_axis_elastomer",
    "tpv": "mass_axis_elastomer",
    "rubber": "mass_axis_elastomer",
    "silicone": "mass_axis_elastomer",
    # 泡棉
    "foam": "mass_axis_foam",
    "pur": "mass_axis_foam",
}


def dominant_material_key(group: str, species: str) -> str | None:
    """按 species **首段**判定主材族，返回材料专属交付价目录 key（质量轴兜底）。

    只认首段第一个 token 作为结构材料词；首段是 EE 类别（EE Components / PCB /
    Connector / Wire Harness …）、密封胶（Sealants / Adhesive）或胶粘剂等
    **非单一可称重材料**词时返回 None —— 这些不能按材料 kg 兜底，只能留白。

    "宁可粗，不可假"：主材族可判定就给交付价兜底，不因 group 标「混合材料」
    或首段带 "Unknown" 修饰词而拒绝；逗号后的嵌件/紧固件/密封胶是附属，不参与判定。

    仅当 species **完全为空**时，才回落到 group 大类判定（单材料族 group 如
    "Metal" / "Plastic" / "Elastomer" 兜底到各自质量轴；"Mixed" / "EE" 不兜底）。
    """
    seg = (species or "").split(",")[0].strip()
    if seg:
        token = re.split(r"[-\s]+", seg.lower())[0]
        if token in _NON_WEIGHABLE_TOKENS:
            return None
        return _PRIMARY_TOKEN_TO_KEY.get(token)
    # species 为空 → 回落到 group 大类
    g = (group or "").strip().lower()
    if g.startswith("mixed") or "混合" in g or g.startswith("ee"):
        return None
    return mass_axis_key(group, species)


# 首段 token 命中这些词 = 不是"单一可称重结构材料"，不能按材料 kg 兜底，
# 只能留白（它们是 EE 类别 / 胶粘密封 / 磁件 / 贴纸 / 纸 / 陶瓷等）。
_NON_WEIGHABLE_TOKENS = {
    "ee", "pcb", "connector", "connectors", "capacitor", "capacitors",
    "sensor", "sensors", "fuse", "fuses", "mosfet", "igbt", "diode",
    "wire", "jumper", "harness", "sealants", "sealant", "adhesive",
    "adhesives", "ferrite", "sticker", "paper", "ceramic", "spring",
    "thermal", "inductor", "transformer", "breather", "ic",
}


def mass_axis_unpriced_reason(group: str, species: str) -> str:
    """说明为何无法按材料 kg 兜底（供未定价行的 reasons 使用）。"""
    seg = (species or "").split(",")[0].strip()
    if not seg:
        g = (group or "").strip().lower()
        if g.startswith("mixed") or "混合" in g:
            return "species 为空且材料族为「混合材料」，无法判定单一主材族"
        if g.startswith("ee"):
            return f"材料族为「{group}」（EE 类别），不能按材料 kg 兜底"
        return f"species 为空且材料族「{group}」无对应质量轴档，无法按材料 kg 兜底"
    token = re.split(r"[-\s]+", seg.lower())[0]
    if token in _NON_WEIGHABLE_TOKENS:
        return (f"species 首段「{seg}」是 EE 类别/胶粘密封/磁件等"
                "非单一可称重材料，不能按材料 kg 兜底")
    return f"species 首段「{seg}」无对应材料族交付价档，无法按材料 kg 兜底"


# 无分层时按材料族取单价参数。★ 必须先按族判，再回落到通用塑料价；
#   否则不锈钢螺钉、铝散热片都会被按"通用塑料 25 元/kg"计价（曾发生）。
SPECIES_PRICE = (("stainless", "ST_PRICE"), ("steel", "ST_PRICE"),
                 ("aluminium", "AL_PRICE"), ("aluminum", "AL_PRICE"),
                 ("alsi", "AL_PRICE"), ("adc", "AL_PRICE"),
                 ("copper", "CU_PRICE"), ("cu ", "CU_PRICE"),
                 ("epdm", "ELASTOMER_PRICE"), ("silicone", "ELASTOMER_PRICE"),
                 ("rubber", "ELASTOMER_PRICE"), ("tpe", "ELASTOMER_PRICE"),
                 ("tpv", "ELASTOMER_PRICE"), ("nbr", "ELASTOMER_PRICE"),
                 ("ferrite", "FERRITE_PRICE"),
                 ("foam", "FOAM_PRICE"), ("pur", "FOAM_PRICE"),
                 ("adhesive", "ADHESIVE_PRICE"), ("resin", "ADHESIVE_PRICE"),
                 ("plastic", "PLASTIC_PRICE"), ("pbt", "PLASTIC_PRICE"),
                 ("pps", "PLASTIC_PRICE"), ("pa66", "PLASTIC_PRICE"),
                 ("pa6", "PLASTIC_PRICE"), ("pom", "PLASTIC_PRICE"),
                 ("pc ", "PLASTIC_PRICE"), ("fr4", "PLASTIC_PRICE"))


def species_price_pid(group: str, species: str) -> str:
    txt = f"{group} {species}".lower()
    for k, pid in SPECIES_PRICE:
        if k in txt:
            return pid
    return "PLASTIC_PRICE"


# ★ 混合材料 / 材料未知的判定。只能用**单一材料族**做质量轴粗估：
#   若 species 是一串并列材料（如 "Copper Unknown, Plastic Unknown, Ferrite,
#   Steel Unknown"），按关键字顺序命中哪一个纯属运气——EX90 真实清单里
#   0.525 kg 的混合件曾被按 ST_PRICE 算成 4.20 元。这类行必须判为未定价。
_MULTI_SPLIT_RE = re.compile(r"[,;，；]")
# ★ 必须带词边界：`n/?a` 不带 \b 会命中 "ENAC43500" 里的 "NA"，
#   把明确的 Aluminium ENAC43500 误判成材料未知（EX90 上发生过）。
_UNKNOWN_RE = re.compile(r"unknown|未知|待确认|\bn/?a\b", re.I)


def species_families(group: str, species: str) -> list[str]:
    """把 species 文本拆成命中的材料族参数 ID 列表（去重，保序）。"""
    out: list[str] = []
    txt = f"{group} {species}".lower()
    for k, pid in SPECIES_PRICE:
        if k in txt and pid not in out:
            out.append(pid)
    return out


def mixed_species_reason(group: str, species: str) -> str:
    """返回非空串 = 该行不能按质量轴粗估，串内容即原因。"""
    txt = f"{group} {species}".strip()
    fams = species_families(group, species)
    if group.strip().lower().startswith("mixed") or "混合" in txt:
        return "材料族为「混合材料」，未做拆解分层，不能按单一材料族单价粗估"
    if _UNKNOWN_RE.search(species or "") or _UNKNOWN_RE.search(group or ""):
        return "材料族含 Unknown，未测明具体合金/牌号，不能按单一材料族单价粗估"
    if len(_MULTI_SPLIT_RE.findall(species or "")) >= 1 and len(fams) > 1:
        return (f"材料族并列 {len(fams)} 种（{'+'.join(fams)}），"
                f"未做拆解分层，不能按单一材料族单价粗估")
    return ""


def material_cost(splits: dict, net_wt: float, group: str = "",
                  species: str = "") -> tuple:
    """M2 材料成本（只算材料；加工由 3-ProcessCost 承担，绝不在此叠加）。"""
    if splits and any(v > 0 for v in splits.values()):
        mat, used = 0.0, []
        for c, pid in SPLIT_PARAM.items():
            kg = splits.get(c, 0.0)
            if kg > 0:
                mat += kg * PV[pid]
                used.append(c.replace(" (kg)", ""))
        if mat > 0:
            return mat, "材料分层：" + "+".join(used)
        return None, "材料分层为 0"
    if net_wt and net_wt > 0:
        bad = mixed_species_reason(group, species)
        if bad:
            return None, bad
        pid = species_price_pid(group, species)
        return (net_wt * PV[pid],
                f"无分层，按净重 × 材料族单价 {pid} = {PV[pid]} 元/kg 粗估")
    return None, "无材料分层且无净重"


def _name_candidates(head: str, catalog: dict) -> list[str]:
    """列出该件名在目录里**前缀命中**的父档 key（用于解释为什么没定上价）。"""
    out = []
    for row in catalog["rows"]:
        if row["use"] == "Calibration" or not row["prefix"]:
            continue
        if row.get("subtier_of"):
            continue
        if any(head.startswith(p) for p in row["prefix"]):
            out.append(row["key"])
    return out


def resolve_price(lf: dict, at: dict, ev: dict | None, catalog: dict) -> dict:
    """取价顺序 M1 取证 → M2 材料影子 → M3/M4/M5 目录 → 质量轴回退 → 未定价。

    `complete=True` 表示该单价已是**交付价**，不再产生工序行。
    """
    # ---- M1 取证价
    if ev and (ev.get("low") or ev.get("high")):
        lo = ev["low"] or ev["high"]
        hi = ev["high"] or ev["low"]
        base, tag = pick_base(lo, hi)
        parts = [ev.get("basis"), tag]
        # ★ 分销价没有采购数量/阶梯价时不得折算，必须明说，不能默认"已经是成交价"
        if not (ev.get("purch_qty") or "").strip():
            parts.append("未完成产量归一（缺采购数量/阶梯价，按单件零售口径，未折算批量）")
            fit = "取证·未归一"
        else:
            fit = "取证"
        return {"method": "M1", "low": lo, "high": hi, "base": base, "complete": True,
                "basis": "；".join(x for x in parts if x),
                "ev": "取证价表", "date": ev.get("date", ""), "deliv": ev.get("deliv", ""),
                "purch_qty": ev.get("purch_qty", ""), "tooling": ev.get("tooling", "?"),
                "tier": tag, "fit": fit, "reason": ""}

    # ---- M2 材料影子（仅自制候选；材料成本 ≠ 交付价 → complete=False）
    # ★ 硬闸门：没有名字/型号时 is_buy("") 恒为 False，会把**每一行**都当成自制件，
    #   于是 PCB 被按塑料价 25 元/kg 算成 0.58 元（EX90 真实清单上发生过）。
    #   无名称时必须有材料分层、或属性表明确写了自制件，才允许走 M2。
    named = bool((lf["name"] or "").strip())
    claims_make = at.get("object", "").strip().startswith("自制")
    has_splits = bool(at.get("splits") and any(v > 0 for v in at["splits"].values()))
    # ★ EE 类别（电子总成/端子/传感器等外购件）即使属性表给了材料分层也**绝不走 M2**：
    #   它们的价值在功能与供应链，不在原料 kg；分层仅作结构对照。若按分层材料成本出价
    #   会系统性严重低估（EX90 端子曾因 OTHER 按 25 元/kg 被算成零头）。EE 件只能走
    #   取证价 / 目录面积或件价，否则留白——「宁可粗，不可假」的反面：假精度不如留白。
    is_ee = (lf.get("group") or "").strip().lower().startswith("ee")
    if (not is_buy(lf["name"]) and not is_ee
            and (named or claims_make or has_splits)):
        mat, why = material_cost(at.get("splits", {}), lf["net_weight"],
                                 lf.get("group", ""), lf.get("species", ""))
        if mat:
            return {"method": "M2", "low": mat, "high": mat, "base": mat, "complete": False,
                    "basis": f"{why}；本行为材料成本，加工见 3-ProcessCost",
                    "ev": "材料分层 / 净重", "date": "", "deliv": "本厂", "purch_qty": "",
                    "tooling": "?", "tier": "", "fit": "材料", "reason": ""}

    # ---- 目录匹配（M3 / M4 / M5）
    # ★ 件名为空时回落到 species 首段：EX90 真实清单只有 Material Species 没有 Name，
    #   而 species 里的 "PCB" / "Capacitor" 正是目录的 Match Head Prefix。
    head = (lf["name"] or "").split(",")[0].strip().lower()
    if not head:
        head = (lf["species"] or "").split(",")[0].strip().lower()
    ctx = {"net_weight": lf["net_weight"], "length": lf["length"], "width": lf["width"]}
    # ★ 必须把属性表的附加列传进去，否则目录 49 个子档永远不可能命中（曾是这样：
    #   传空 dict → 子档 Required Attribute 全部"未知" → 全部退父档，分层信息白填）
    hit = match_price(head, lf["species"], at.get("attrs", {}) or {}, ctx, catalog)
    if hit:
        basis = hit["basis"]
        fit = "子档命中" if hit["subtier"] else "父档命中"
        # ★ 父档本有子档、本件却落在父档 = 属性缺失（层数/铜厚/螺纹/引脚数未知）。
        #   这时用的是**通用档**价，属假设，Basis 必须以 `[假设]` 开头，
        #   否则读者会把"4 层通用面积价"当成"这块板的实测档位"。
        kids = catalog["by_parent"].get(hit["key"]) or []
        if not hit["subtier"] and kids:
            reqs = sorted({r["req_attr"].split("=")[0].split(">")[0].split("<")[0].strip()
                           for r in kids if r["req_attr"]})
            missing = [k for k in reqs
                       if (at.get("attrs", {}) or {}).get(k.lower()) in (None, "")]
            if missing:
                basis = (f"[假设] 属性缺失({'.'.join(missing)})，按父档通用价 {hit['key']}；"
                         + basis)
                fit = "父档回退（属性缺失）"
        if hit["method"] == "M3":
            fit += "·M3 须过 Fit Check"
        return {"method": hit["method"], "low": hit["low"], "high": hit["high"],
                "base": hit["base"], "complete": True, "basis": basis,
                "ev": f"目录 {hit['key']}", "date": hit["row"]["src_date"],
                "deliv": "目录基准（交付价）", "purch_qty": "", "tooling": hit["tooling"],
                "tier": hit["tier"], "fit": fit,
                "reason": ""}

    # ---- 质量轴回退（CNY/kg × 净重；普通金属/塑料结构件唯一的兜底定价通道）
    # ★ 主材族判定只看 species **首段**第一个 token：拆解清单 species 是
    #   "主材在前、嵌件/紧固件/密封胶在后" 的逗号串。group 标「混合材料」或
    #   首段带 "Unknown" 修饰词都不应据此拒绝——主材族依然确定（宁可粗，不可假）。
    #   只有首段本身是 EE 类别/胶粘密封/磁件等非可称重材料时，才不能按 kg 兜底。
    key = dominant_material_key(lf["group"], lf["species"])
    row = catalog["by_key"].get(key) if key else None
    if row and lf["net_weight"] and lf["net_weight"] > 0:
        lo = row["low"] * lf["net_weight"]
        hi = (row["high"] or row["low"]) * lf["net_weight"]
        base, tag = pick_base(lo, hi)
        return {"method": "M5", "low": lo, "high": hi, "base": base, "complete": True,
                # ★ 质量轴是"件名测不出类别"时的兜底推断，一律算假设
                "basis": (f"[兜底-质量轴] 件名无目录前缀命中，按主材族 {key} 交付价推算："
                          f"{row['low']}–{row['high']} CNY/kg × {lf['net_weight']:.3f} kg"
                          + (f"；{tag}" if tag else "")),
                "ev": f"目录 {key}", "date": row["src_date"],
                "deliv": "目录基准（交付价）", "purch_qty": "", "tooling": row["tooling"],
                "tier": tag, "fit": "质量轴（主材族兜底）", "reason": ""}

    # ---- 未定价：写清原因，且必须写出"补什么才能定上价"
    cands = _name_candidates(head, catalog)
    reasons = ["件名为空" if not head else "件名无目录前缀命中"]
    if cands:
        reasons[-1] = (f"件名命中父档 {'/'.join(cands)} 但未过闸门"
                       f"（需补齐目录 Required Attribute 或补 species）")
    if is_ee and (has_splits or claims_make):
        reasons.append("EE 类别件（电子总成/端子/传感器等）即使给了材料分层也不按"
                       "材料 kg 影子计价——其价值在功能与供应链，不在原料；"
                       "分层仅作结构对照")
    if not named and not claims_make and not has_splits:
        reasons.insert(0, "清单无 Name 列且属性表未标自制件 / 未给材料分层，"
                          "无法判定自制还是外购")
    if key is None:
        reasons.append(mass_axis_unpriced_reason(lf["group"], lf["species"]))
    elif not (lf["net_weight"] and lf["net_weight"] > 0):
        reasons.append("净重为 0，质量轴无法换算")
    if not cands and key is None:
        reasons.append(f"species=「{lf['species']}」未通过目录 Species Contains"
                       f"（常见：目录只收 pcb/capacitor/connector 等具体族名，"
                       f"配方代号如 FR4 不匹配）"
                       if lf["species"] else "species 为空，目录父档无法通过 Species 闸门")
    if key is None or (not named and not claims_make and not has_splits):
        reasons.append("补什么：拆解分层称重（各材料 kg）填进属性表，"
                       "并补型号 / 规格书 / 供应商报价")
    else:
        reasons.append("补什么：型号 / 规格书 / 供应商报价，或补拆解尺寸与层数")
    return {"method": "", "low": None, "high": None, "base": None, "complete": False,
            "basis": "未定价：" + "；".join(reasons), "ev": "", "date": "", "deliv": "",
            "purch_qty": "", "tooling": "?", "tier": "", "fit": "—",
            "reason": "；".join(reasons)}


# ================================================================ 写表工具
def write_header(ws, row: int, labels: list[str], widths: list[int]) -> None:
    for i, (lab, w) in enumerate(zip(labels, widths), start=1):
        c = ws.cell(row=row, column=i, value=lab)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 32


def prose(v):
    """★ 只用于标题 / 注释这类散文。

    散文以 '=' 开头时 Excel 会当成公式并报语法错，故转全角 '＝'。
    公式单元格绝不能走这个函数——v1x 早期版本把 write_body 也套了它，
    结果全部公式变成文本，4-Summary 引用后满屏 #VALUE!。
    """
    if isinstance(v, str) and v.lstrip().startswith("="):
        return "＝" + v.lstrip()[1:]
    return v


def write_body(ws, row: int, values: list, status: str | None = None) -> None:
    """数据行：一律原样写入，写在 `=` 开头的字符串就是要变成公式。"""
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = Alignment(vertical="top", wrap_text=True)
    if status and STATUS_FILL.get(status):
        for i in range(1, len(values) + 1):
            ws.cell(row=row, column=i).fill = STATUS_FILL[status]


def note(ws, row: int, text: str, font=NOTE_FONT) -> int:
    ws.cell(row=row, column=1, value=prose(text)).font = font
    return row + 1


def block(ws, start: int, title: str, rows: list[tuple],
          cols=("项目", "金额 / 值", "说明")) -> dict:
    ws.cell(row=start, column=1, value=title).font = SEC_FONT
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=start + 1, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    r = start + 2
    marks: dict[str, int] = {}
    for lab, val, nte, st in rows:
        c1 = ws.cell(row=r, column=1, value=prose(lab))
        c1.font = BODY_FONT
        c1.border = BORDER
        # ★ 第 2 列是数值/公式，绝不能转义
        c2 = ws.cell(row=r, column=2, value=val)
        c2.font = BODY_FONT
        c2.border = BORDER
        if st and STATUS_FILL.get(st):
            c2.fill = STATUS_FILL[st]
        c3 = ws.cell(row=r, column=3, value=prose(nte))
        c3.font = NOTE_FONT
        c3.border = BORDER
        marks[lab] = r
        r += 1
    return {"first": start + 2, "last": r - 1, "next": r + 1, "marks": marks}


# ================================================================ 主构建
def build(leaves: list[dict], out: Path, *, product: str = "Product",
          attrs: dict | None = None, prices: dict | None = None,
          catalog: dict | None = None, opps: list[dict] | None = None,
          src_names: dict | None = None, func_prefix: str | None = None) -> dict:
    attrs = attrs or {}
    prices = prices or {}
    catalog = catalog if catalog is not None else load_catalog()
    sn = src_names or {}
    prefix = func_prefix_for(product, func_prefix)

    wb = Workbook()
    wb.remove(wb.active)
    bom_labels, bom_widths, B = make_cols(BOM_SPEC)
    proc_labels, proc_widths, P3 = make_cols(PROC_SPEC)
    va_labels, va_widths, V = make_cols(VA_SPEC)

    # ============================================================ 1-Assumptions
    ws1 = wb.create_sheet("1-Assumptions")
    ws1["A1"] = f"{product} — 假设与费率（v1x：PPC = MAT + VA，逐件 M1–M5）"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = "参数唯一源：本书所有费率、比例、k、s、阈值只出现在本表，其他表引用单元格，不写死数字。"
    ws1["A2"].font = WARN_FONT
    write_header(ws1, 4, ["参数ID", "参数", "值", "单位", "状态", "来源/说明"],
                 [18, 40, 14, 10, 12, 80])
    ref: dict[str, str] = {}
    r = 5
    for pid, lab, val, unit, st, nte in PARAMS:
        write_body(ws1, r, [pid, lab, val, unit, st, nte], st)
        ws1.cell(row=r, column=3).number_format = (
            RATIO if isinstance(val, float) and abs(val) < 2 else '#,##0.####')
        ref[pid] = f"'1-Assumptions'!$C${r}"
        r += 1

    # 派生参数（隐含件定尺寸；用的是整列引用，2-BOM 建完后自动生效）
    r += 1
    ws1.cell(row=r, column=1,
             value="派生参数（隐含件定尺寸；只用于算预留尺寸，不直接进 MAT）").font = SEC_FONT
    write_header(ws1, r + 1, ["参数ID", "参数", "值", "单位", "状态", "来源/说明"],
                 [18, 40, 14, 10, 12, 80])
    dr = r + 2
    # ★ 2-BOM 尚未建立，公式里的行区间只能先占位，等 2-BOM 建完再回填
    derived_at: dict[str, int] = {}
    # ★ 必须逐条顺序求值：T0 的公式要用到刚写入的行号，不能在一个元组里一起构造
    for pid, lab, unit, nte in (
        ("L_NON_HIDDEN", "非预留行计入金额合计 L", "CNY",
         "排除 Row Type = 预留 的行，避免循环引用"),
        ("S_SUM", "隐含件比例合计 Σs（中点）", "-", "只用于定尺寸"),
        ("T0", "隐含件定尺寸基数 T0 = L ÷ (1 − Σs)", "CNY",
         "T0 只用于定尺寸，不直接进入 MAT"),
    ):
        if pid == "L_NON_HIDDEN":
            formula = None                      # 回填
        elif pid == "S_SUM":
            formula = f'={ref["S_CHEM"]}+{ref["S_FAST"]}'
        else:
            formula = f'=IFERROR({ref["L_NON_HIDDEN"]}/(1-{ref["S_SUM"]}),"")'
        write_body(ws1, dr, [pid, lab, formula, unit, "Calculated", nte], "Calculated")
        ws1.cell(row=dr, column=3).number_format = MONEY if unit == "CNY" else RATIO
        derived_at[pid] = dr
        ref[pid] = f"'1-Assumptions'!$C${dr}"
        dr += 1
    dr += 1
    for t in (
        "★ 隐含件预留：Reserve_c = MAX(0, s_c × T0 − 清单已列该类金额)。"
        "预留计入『全口径』总额，不计入『已出价行数』。",
        "★ 废料回收只对金属生效；塑料/铁氧体/泡棉/胶不减废料。",
        "★ MOH 已在 VA 内：MOH/MAT 只用于 VA 内部拆分，不得在 VA 之上再叠一层。",
        "★ 设备折旧与工装摊销各只计一次：投资 ÷ 年限 ÷ 产量；件价已含的不再添加。",
        "★ 外购件取交付价，已含供应商加工；不要再为它生成工序行，否则重复计价。",
    ):
        dr = note(ws1, dr, t)
    ws1.freeze_panes = "A5"

    # ============================================================ 2-BOM
    ws2 = wb.create_sheet("2-BOM")
    ws2["A1"] = "MAT —— 逐件定价表（Unit Price = 本行采用单价；Price Low/High = 区间，永远保留）"
    ws2["A1"].font = TITLE_FONT
    ws2["A2"] = ("★ 汇总只对 Counted Amount 求和。父件按总成价计入时子件仍显示参考估价且 "
                 "Counted In Total = N；同一件不得同时按总成价与展开成本计入。")
    ws2["A2"].font = WARN_FONT
    write_header(ws2, 4, bom_labels, bom_widths)

    parts: list[dict] = []
    r = 5
    for lf in leaves:
        at = attrs.get(lf["part"], {})
        func, cat = function_category(lf["name"], lf["species"], prefix)
        gp = resolve_price(lf, at, prices.get(lf["part"]), catalog)
        priced = gp["base"] is not None
        route_text, steps = route_for(lf["name"], lf["species"])
        strength = route_strength(route_text)
        # 只有「材料成本」行（M2）才需要工序；交付价行不再生成工序
        need_pdc = bool(priced and not gp["complete"] and strength == "Normal")
        values = [
            None, lf["part"], lf["name"], "汇总件", func, cat,
            "Make" if gp["method"] == "M2" else "Buy",
            "MAT-自制并入" if gp["method"] == "M2" else "MAT-零部件",
            lf["qty"], round(lf["net_weight"], 4) if lf["net_weight"] else None,
            "照片计数" if lf["itf"] else "清单件数",
            sn.get("leaves") or "叶件清单",
            gp["base"], gp["low"], gp["high"], gp["method"] or "未定价",
            gp["basis"], gp["ev"], gp["fit"], gp["date"],
            gp["purch_qty"], gp["deliv"], gp["tooling"],
            "Y" if priced else "N",
            f'=IF(${B["counted"]}{r}="Y",${B["price"]}{r}*${B["qty"]}{r},"")',
            "" if priced else "未定价：不计入合计，列入待取证清单",
            f'=IF(AND(${B["counted"]}{r}="Y",ISNUMBER(${B["p_low"]}{r})),'
            f'${B["p_low"]}{r}*${B["qty"]}{r},"")',
            f'=IF(AND(${B["counted"]}{r}="Y",ISNUMBER(${B["p_high"]}{r})),'
            f'${B["p_high"]}{r}*${B["qty"]}{r},"")',
        ]
        write_body(ws2, r, values, "Estimated" if priced else "Unknown")
        ws2.cell(row=r, column=1, value=r - 4)
        for k in ("price", "p_low", "p_high", "counted_amt", "amt_low", "amt_high"):
            ws2[f"{B[k]}{r}"].number_format = MONEY
        ws2[f"{B['netwt']}{r}"].number_format = KG
        parts.append({"row": r, "part": lf["part"], "name": lf["name"], "qty": lf["qty"],
                      "func": func, "cat": cat, "method": gp["method"], "priced": priced,
                      "tier": gp["tier"], "low": gp["low"], "high": gp["high"],
                      "base": gp["base"], "net_weight": lf["net_weight"],
                      "need_pdc": need_pdc, "steps": steps if need_pdc else [],
                      "strength": strength, "reason": gp["reason"]})
        r += 1

    # 预留行（Row Type = 预留）
    res_rows: list[int] = []
    for key, s_mid, s_lo, s_hi, bucket in (
        ("CHEM", "S_CHEM", "S_CHEM_LOW", "S_CHEM_HIGH", "MAT-Chem"),
        ("FAST", "S_FAST", "S_FAST_LOW", "S_FAST_HIGH", "MAT-零部件"),
    ):
        lo_f = f'=MAX(0,{ref[s_lo]}*{ref["L_NON_HIDDEN"]}/(1-{ref["S_SUM"]}))'
        mid_f = f'=MAX(0,{ref[s_mid]}*{ref["T0"]})'
        hi_f = f'=MAX(0,{ref[s_hi]}*{ref["L_NON_HIDDEN"]}/(1-{ref["S_SUM"]}))'
        values = [
            None, f"RESERVE-{key}", f"隐含件预留（{key}）", "预留", prefix,
            # ★ 必须与功能域透视里那一行完全一致，否则透视不闭合
            "隐含件预留", "Buy", bucket,
            1, None, "比例补全", "校准文件 C 节比例 s",
            mid_f, lo_f, hi_f, "M5",
            f"隐含件预留：s_c × T0；s 见 1-Assumptions {s_mid} / {s_lo} / {s_hi}；"
            f"只补清单可能遗漏的部分",
            "校准文件 C 节", "—", "",
            "", "未记录", "?",
            "Y",
            f'=IF(${B["counted"]}{r}="Y",${B["price"]}{r}*${B["qty"]}{r},"")',
            "预留行：计入全口径总额，不计入已出价行数",
            f'=IF(ISNUMBER(${B["p_low"]}{r}),${B["p_low"]}{r}*${B["qty"]}{r},"")',
            f'=IF(ISNUMBER(${B["p_high"]}{r}),${B["p_high"]}{r}*${B["qty"]}{r},"")',
        ]
        write_body(ws2, r, values, "Scenario")
        ws2.cell(row=r, column=1, value=r - 4)
        for k in ("price", "p_low", "p_high", "counted_amt", "amt_low", "amt_high"):
            ws2[f"{B[k]}{r}"].number_format = MONEY
        res_rows.append(r)
        r += 1
    bom_last = r - 1
    ws2.freeze_panes = "D5"

    # ★ 数据行区间（有界）。绝不能用整列：整列里 100 万空行也会命中
    #   SUMIFS(...,"<>预留") 与 COUNTIF(...,"") —— 曾导致结论句报「1048574 个实物行」。
    RT_B = f"'2-BOM'!${B['rowtype']}$5:${B['rowtype']}${bom_last}"
    AMT_B = f"'2-BOM'!${B['counted_amt']}$5:${B['counted_amt']}${bom_last}"
    MTH_B = f"'2-BOM'!${B['method']}$5:${B['method']}${bom_last}"
    FUNC_B = f"'2-BOM'!${B['func']}$5:${B['func']}${bom_last}"
    CAT_B = f"'2-BOM'!${B['cat']}$5:${B['cat']}${bom_last}"
    rng = {"rt": RT_B, "amt": AMT_B, "m": MTH_B, "func": FUNC_B, "cat": CAT_B,
           "bom_last": bom_last}

    # ★ 回填 1-Assumptions 的 L_NON_HIDDEN（2-BOM 现在才有确定行数）
    ws1.cell(row=derived_at["L_NON_HIDDEN"], column=3,
             value=f'=SUMIFS({AMT_B},{RT_B},"<>预留")')
    ws1.cell(row=derived_at["L_NON_HIDDEN"], column=3).number_format = MONEY

    fn = bom_last + 2
    for t, f_ in (
        ("★ 空价必须为空串（不是 0）。写成 0 会把『缺价』伪装成『零成本』，且 SUM 会静默吞掉。",
         WARN_FONT),
        ("★ 父件展开：子件参考价之和与父件价不一致时只列差额，不调平。", NOTE_FONT),
        ("★ Price Method：M1 取证价 / M2 影子计算（只含材料，加工在 3-ProcessCost）/ "
         "M3 校准锚点 / M4 目录有来源区间 / M5 粗估（含质量轴回退）。", NOTE_FONT),
        ("★ [宽>10×] 行须先进待取证清单前 5；[包络] 表示目录无实证锚点。", NOTE_FONT),
    ):
        fn = note(ws2, fn, t, f_)
    # ★ 父件展开的重复计价闸门：同一份清单里既有父件价又有其子件明细时，
    #   两套都计入 = 重复计价。本生成器不自动改行类型（那会改变交付形态），
    #   但必须**显式告警**并指出是哪几件，不能沉默。
    _ids = {(p["part"] or "").strip() for p in parts}
    _dup = sorted({(lf["part"], (lf.get("parent") or "").strip())
                   for lf in leaves
                   if (lf.get("parent") or "").strip()
                   and (lf.get("parent") or "").strip() in _ids
                   and (lf.get("parent") or "").strip() != (lf["part"] or "").strip()})
    if _dup:
        fn = note(ws2, fn,
                  f"⚠ 父子并存 {len(_dup)} 组（父件与子件同时在本清单内）："
                  f"{', '.join(f'{c}∈{p}' for c, p in _dup[:8])}"
                  f"{' …' if len(_dup) > 8 else ''}。"
                  f"计入时必须二选一（父件总成价 或 子件展开合计），"
                  f"否则重复计价；本表未自动改行类型，请人工确认后把其中一侧 Counted In Total 置 N。",
                  WARN_FONT)

    # ============================================================ 3-ProcessCost
    ws3 = wb.create_sheet("3-ProcessCost")
    ws3["A1"] = "VA —— 增值成本（PDC + PDOH + PAC + MOH）；分块只在有比例依据时给"
    ws3["A1"].font = TITLE_FONT
    ws3["A2"] = ("★ 只有 M2（材料成本）行在这里生成工序，因为它们的单价不含加工；"
                 "M1/M3/M4/M5 行取的是交付价，已含供应商加工，这里不再生成，否则重复计价。")
    ws3["A2"].font = WARN_FONT
    write_header(ws3, 4, proc_labels, proc_widths)

    pr, pdc_rows, routed_parts = 5, [], []
    for p in parts:
        if not p["steps"]:
            continue
        routed_parts.append(p["part"])
        n = len(p["steps"])
        for i, (proc, cat, cycle, rate_pid, ops, cons, yld) in enumerate(p["steps"], start=1):
            if rate_pid is None:
                # 外协：单价 × 基数 ÷ 良率，不走设备费率
                kg = p["net_weight"] or 0.0
                direct = (f'=IF({ref["PLATE_PRICE"]}*{kg:.6f}=0,"",'
                          f'{ref["PLATE_PRICE"]}*{kg:.6f}/{yld})')
                rate_cell, rate_basis = None, "外协：单价 × 基数 ÷ 良率"
                cons_v = None
            else:
                rate_cell = f'={ref[rate_pid]}'
                rate_basis = f"1-Assumptions {rate_pid}"
                cons_v = cons if isinstance(cons, (int, float)) else 0.0
                direct = (f'=IF(AND(ISNUMBER(${P3["cycle"]}{pr}),ISNUMBER(${P3["rate"]}{pr})),'
                          f'(${P3["cycle"]}{pr}/3600*${P3["rate"]}{pr}+N(${P3["cons"]}{pr}))'
                          f'/${P3["yld"]}{pr},"")')
            cumul = (direct if i == 1 else
                     f'=IF(${P3["direct"]}{pr}="","",'
                     f'${P3["direct"]}{pr}+N(${P3["cumul"]}{pr - 1}))')
            write_body(ws3, pr, [
                i, p["part"], proc, cat, "本件", "in-house",
                cycle if rate_pid else None, "模板节拍" if rate_pid else "外协无节拍",
                rate_cell, rate_basis,
                cons_v, yld, direct, cumul,
                "Y" if i == n else "", p["strength"],
            ], "Scenario")
            for k in ("rate", "cons", "direct", "cumul"):
                ws3[f"{P3[k]}{pr}"].number_format = MONEY
            pdc_rows.append(pr)
            pr += 1
    proc_last = pr - 1

    # VA 汇总（独立列块）
    va_top = proc_last + 2
    ws3.cell(row=va_top, column=1, value="VA 合计（Low / Base / High + 方法层级）").font = SEC_FONT
    write_header(ws3, va_top + 1, va_labels, va_widths)
    vr = va_top + 2
    m_rng = MTH_B
    amt_rng = AMT_B
    rt_rng = RT_B
    # ★ k 只覆盖「没有任何 L0 工序」的 MAT。判据必须是「该件是否有工序行」，
    #   不能用 Method="M2" 代替：M2 但 Route Strength=Weak 的件（螺钉、导热垫）
    #   既没有工序行、又会被 M2 判据排除在 k 之外 → 变成零加工的黑洞（曾发生）。
    part_rng = f"'2-BOM'!${B['part']}$5:${B['part']}${bom_last}"
    if routed_parts:
        pdc_covered = "+".join(f'SUMIFS({amt_rng},{part_rng},"{pp}")'
                               for pp in routed_parts)
    else:
        pdc_covered = "0"
    k_base = f'MAX(0,SUMIFS({amt_rng},{rt_rng},"<>预留")-({pdc_covered}))'
    va_first = vr

    if pdc_rows:
        pdc_sum = (f'SUMIFS(${P3["cumul"]}$5:${P3["cumul"]}{proc_last},'
                   f'${P3["is_last"]}$5:${P3["is_last"]}{proc_last},"Y")')
        write_body(ws3, vr, [
            "PDC（L0 工序合计）", "PDC", f"={pdc_sum}", f"={pdc_sum}", f"={pdc_sum}",
            "L0", "in-house", "由工序行累计；区间未建模，故 Low = Base = High",
            "3-ProcessCost 工序行 Is Last Step = Y",
        ], "Calculated")
    else:
        write_body(ws3, vr, [
            "PDC（无 L0 工序输入）", "PDC", None, None, None,
            "—", "—", "无任何行建了工序：本项为空，VA 全由 L2 比例给出，须在 7-Evidence 披露",
            "—",
        ], "Unknown")
    for k in ("low", "base", "high"):
        ws3[f"{V[k]}{vr}"].number_format = MONEY
    vr += 1

    write_body(ws3, vr, [
        "PDC + PDOH + PAC（L2 比例，只覆盖无 L0 工序的 MAT）", "PDC+PDOH+PAC",
        f'=({k_base})*{ref["K_VA_LOW"]}', f'=({k_base})*{ref["K_VA_BASE"]}',
        f'=({k_base})*{ref["K_VA_HIGH"]}',
        "L2", "in-house",
        "k 的基数 = 全部非预留 MAT − 已建 L0 工序件的 MAT，"
        "避免与 L0 工序重复；判据按『是否有工序行』而非 Method=M2",
        "1-Assumptions K_VA_LOW / BASE / HIGH",
    ], "Estimated")
    for k in ("low", "base", "high"):
        ws3[f"{V[k]}{vr}"].number_format = MONEY
    vr += 1

    moh = f'SUMIFS({amt_rng},{rt_rng},"<>预留")*{ref["MOH_RATIO"]}'
    write_body(ws3, vr, [
        "MOH（VA 内部拆分，不叠加）", "MOH", f"={moh}", f"={moh}", f"={moh}",
        "L2", "in-house", "MOH 已在 VA 内，不得在 VA 之上再加一层",
        "1-Assumptions MOH_RATIO",
    ], "Estimated")
    for k in ("low", "base", "high"):
        ws3[f"{V[k]}{vr}"].number_format = MONEY
    vr += 1

    write_body(ws3, vr, [
        "VA 合计", "合计",
        f'=SUM({V["low"]}{va_first}:{V["low"]}{vr - 1})',
        f'=SUM({V["base"]}{va_first}:{V["base"]}{vr - 1})',
        f'=SUM({V["high"]}{va_first}:{V["high"]}{vr - 1})',
        "L0+L2", "—",
        "竞品 VA 一律『折算估计，非实测』；产量是首要敏感性",
        "本表上方各行",
    ], "Calculated")
    for k in ("low", "base", "high"):
        ws3[f"{V[k]}{vr}"].number_format = MONEY
    VA_TOTAL = {k: f"'3-ProcessCost'!${V[k]}${vr}" for k in ("low", "base", "high")}
    va_total_row = vr
    ws3.freeze_panes = "C5"
    note(ws3, vr + 2, "★ 外协工序：单价 × 基数 ÷ 良率，不走设备费率。")
    note(ws3, vr + 3, "★ 外购件不进本表；VA 只含增值动作。")
    note(ws3, vr + 4, "★ Route Strength = Weak 的行表示件名只能判到装配级，属弱推断，不建 L0 工序。")

    # ============================================================ 4-Summary
    ws4 = wb.create_sheet("4-Summary")
    n_priced = sum(1 for p in parts if p["priced"])
    n_unpriced = len(parts) - n_priced
    n_m5 = sum(1 for p in parts if p["method"] == "M5")
    n_res = len(res_rows)

    mat_anch_base = "+".join(
        f'SUMIFS({amt_rng},{rt_rng},"<>预留",{m_rng},"{m}")' for m in ("M1", "M2", "M3", "M4"))
    mat_full = f'SUMIFS({amt_rng},{rt_rng},"<>预留")+SUMIFS({amt_rng},{rt_rng},"预留")'
    lo_rng = f"'2-BOM'!${B['amt_low']}$5:${B['amt_low']}${bom_last}"
    hi_rng = f"'2-BOM'!${B['amt_high']}$5:${B['amt_high']}${bom_last}"
    mat_full_lo = (f'SUMIFS({lo_rng},{rt_rng},"<>预留")+SUMIFS({lo_rng},{rt_rng},"预留")')
    mat_full_hi = (f'SUMIFS({hi_rng},{rt_rng},"<>预留")+SUMIFS({hi_rng},{rt_rng},"预留")')

    ws4["A1"] = ""  # 占位，最后回填结论句
    ws4["A2"] = "★ 『已出价』不等于『可靠』：M5 与预留占比必须同时显示，二者不得合并成一个数。"
    ws4["A2"].font = WARN_FONT

    b1 = block(ws4, 4, "口径块（Step 1 / 一）", [
        ("报告日期", "=" + ref["REPORT_DATE"], "生成日", "Verified"),
        ("价格基期", "=" + ref["PRICE_BASE"], "年", "Estimated"),
        ("目标年份", "=" + ref["TARGET_YEAR"], "年", "Estimated"),
        ("情景", "=" + ref["SCENARIO"], "竞品默认现状", "Estimated"),
        ("币种", "=" + ref["CUR"], "", "Estimated"),
        ("来源版本", "=" + ref["SRC_VER"], "", "Estimated"),
        ("边界归一开关", "=" + ref["BOUNDARY_SW"], "默认关", "Estimated"),
    ])

    b2 = block(ws4, b1["next"] + 1, "MAT / VA / PPC（PPC = MAT + VA）", [
        ("已锚定 MAT（M1–M4）", f"={mat_anch_base}",
         "只含 M1–M4 行；这是可对外引用的部分", "Calculated"),
        ("全口径 MAT（含 M5 与预留）", f"={mat_full}",
         "含 M5 与预留行；PPC 标题数用此口径", "Calculated"),
        ("VA（Base）", "=" + VA_TOTAL["base"], "PDC + PDOH + PAC + MOH", "Calculated"),
        ("PPC（Base，全口径）", None, "= 全口径 MAT + VA", "Calculated"),
        ("PPC（Low）", None, "M1–M4 与 M5、预留全部取 Low", "Calculated"),
        ("PPC（High）", None, "同上，取 High", "Calculated"),
    ])
    m2 = b2["marks"]
    r_anch = m2["已锚定 MAT（M1–M4）"]
    r_full = m2["全口径 MAT（含 M5 与预留）"]
    r_va = m2["VA（Base）"]
    r_ppc = m2["PPC（Base，全口径）"]
    r_plo = m2["PPC（Low）"]
    r_phi = m2["PPC（High）"]
    ws4[f"B{r_ppc}"] = f"=B{r_full}+B{r_va}"
    ws4[f"B{r_plo}"] = f"={mat_full_lo}+{VA_TOTAL['low']}"
    ws4[f"B{r_phi}"] = f"={mat_full_hi}+{VA_TOTAL['high']}"
    for rr in (r_anch, r_full, r_va, r_ppc, r_plo, r_phi):
        ws4[f"B{rr}"].number_format = MONEY
    for rr in (r_anch, r_full, r_ppc):
        ws4[f"B{rr}"].font = BOLD_FONT

    # 方法占比：5 行分行披露（M1+M2 / M3 / M4 / M5 / 预留），不允许合并
    s3 = b2["next"] + 1
    ws4.cell(row=s3, column=1, value="方法占比（分行披露，不得合并成一个数）").font = SEC_FONT
    write_header(ws4, s3 + 1, ["方法", "件数", "金额", "占全口径 MAT", "说明"],
                 [30, 10, 14, 13, 46])
    method_rows = [
        ("M1 取证价", ("M1",), "可追溯到报价 / 发票"),
        ("M2 影子计算（材料成本）", ("M2",), "可在本工作簿复算；加工在 3-ProcessCost"),
        ("M3 校准锚点", ("M3",), "来自三份报告，须过 Fit Check"),
        ("M4 目录有来源区间", ("M4",), "来源与日期可追溯"),
        ("M5 粗估（含质量轴回退）", ("M5",), "区间宽，进待取证清单"),
        ("预留（隐含件）", None, "不计入已出价行数"),
    ]
    mr = s3 + 2
    for lab, mthds, nte in method_rows:
        if mthds:
            cnt = "+".join(f'COUNTIFS({rt_rng},"<>预留",{m_rng},"{m}")' for m in mthds)
            amt = "+".join(f'SUMIFS({amt_rng},{rt_rng},"<>预留",{m_rng},"{m}")' for m in mthds)
        else:
            cnt = f'COUNTIF({rt_rng},"预留")'
            amt = f'SUMIFS({amt_rng},{rt_rng},"预留")'
        ws4.cell(row=mr, column=1, value=lab).font = BODY_FONT
        ws4.cell(row=mr, column=2, value=f"={cnt}").font = BODY_FONT
        c = ws4.cell(row=mr, column=3, value=f"={amt}")
        c.number_format = MONEY
        c.font = BODY_FONT
        c4 = ws4.cell(row=mr, column=4, value=f'=IF($B${r_full}=0,"",C{mr}/$B${r_full})')
        c4.number_format = PCT
        c4.font = BODY_FONT
        ws4.cell(row=mr, column=5, value=nte).font = NOTE_FONT
        for i in range(1, 6):
            ws4.cell(row=mr, column=i).border = BORDER
        mr += 1
    row_m5 = s3 + 2 + 4
    row_res = s3 + 2 + 5

    rn = mr
    ws4.cell(row=rn, column=1, value="非锚定占比 =（M5 + 预留）÷ 全口径 MAT").font = BOLD_FONT
    c = ws4.cell(row=rn, column=4,
                 value=f'=IF($B${r_full}=0,"",(C{row_m5}+C{row_res})/$B${r_full})')
    c.number_format = PCT
    c.font = BOLD_FONT
    c.border = BORDER
    ws4.cell(row=rn, column=5, value="结论等级由这一个数决定；阈值在 1-Assumptions").font = NOTE_FONT
    nonanch = f"$D${rn}"

    # 结论等级（Step 7）
    rl = rn + 2
    ws4.cell(row=rl, column=1,
             value="结论等级（取各项最低档；阈值为经验设定，待校准）").font = SEC_FONT
    write_header(ws4, rl + 1, ["等级", "判据", "可给的结论"], [10, 34, 54])
    for i, (lvl, cond, allowed) in enumerate((
        (1, f'=IF({nonanch}<={ref["TH_M5_L1"]},"非锚定 ≤ 15% → 满足","不满足")',
         "功能域占比、Top-N 单件排序，均带区间"),
        (2, f'=IF({nonanch}<={ref["TH_M5_L3"]},"非锚定 15%–40% → 满足","不满足")',
         "功能域区间与差距方向，不给单件排序"),
        (3, f'=IF({nonanch}>{ref["TH_M5_L3"]},"非锚定 > 40% → 满足","不满足")',
         "定性判断与数量级"),
    )):
        rr = rl + 2 + i
        ws4.cell(row=rr, column=1, value=lvl).font = BODY_FONT
        ws4.cell(row=rr, column=2, value=cond).font = BODY_FONT
        ws4.cell(row=rr, column=3, value=allowed).font = BODY_FONT
        for j in range(1, 4):
            ws4.cell(row=rr, column=j).border = BORDER
    level_formula = (f'=IF({nonanch}<={ref["TH_M5_L1"]},1,'
                     f'IF({nonanch}<={ref["TH_M5_L3"]},2,3))')
    ws4.cell(row=rl + 6, column=1, value="当前等级").font = BOLD_FONT
    ws4.cell(row=rl + 6, column=2, value=level_formula).font = BOLD_FONT
    ws4.cell(row=rl + 6, column=3,
             value="Top-N 出口须与此等级一致；等级 2 及以下不导出单件排序").font = NOTE_FONT

    # 功能域透视
    sp = rl + 9
    ws4.cell(row=sp, column=1,
             value="功能域透视（Function × Category；只对 Counted Amount 求和）").font = SEC_FONT
    write_header(ws4, sp + 1, ["Function", "Category", "MAT", "Share %", "说明"],
                 [12, 20, 14, 11, 40])
    cats = sorted({(p["func"], p["cat"]) for p in parts})
    # ★ 预留行必须也进透视，否则透视永远不闭合（曾差 −76.13，恰好等于预留合计）
    res_func, res_cat = prefix, "隐含件预留"
    sr = sp + 2
    for i, (fu, ca) in enumerate(cats):
        rr = sr + i
        ws4.cell(row=rr, column=1, value=fu).font = BODY_FONT
        ws4.cell(row=rr, column=2, value=ca).font = BODY_FONT
        c = ws4.cell(row=rr, column=3,
                     value=f'=SUMIFS({amt_rng},{FUNC_B},"{fu}",{CAT_B},"{ca}")')
        c.number_format = MONEY
        c.font = BODY_FONT
    rr = sr + len(cats)
    ws4.cell(row=rr, column=1, value=res_func).font = BODY_FONT
    ws4.cell(row=rr, column=2, value=res_cat).font = BODY_FONT
    c = ws4.cell(row=rr, column=3,
                 value=f'=SUMIFS({amt_rng},{RT_B},"预留")')
    c.number_format = MONEY
    c.font = BODY_FONT
    ws4.cell(row=rr, column=5,
             value="隐含件预留（Chem+Scrap / 紧固件），只计全口径").font = NOTE_FONT
    n_piv = len(cats) + 1
    tot_row = sr + n_piv
    ws4.cell(row=tot_row, column=1, value="合计").font = BOLD_FONT
    c = ws4.cell(row=tot_row, column=3, value=f"={mat_full}")
    c.number_format = MONEY
    c.font = BOLD_FONT
    for i in range(n_piv):
        c4 = ws4.cell(row=sr + i, column=4,
                      value=f'=IF($C${tot_row}=0,"",C{sr + i}/$C${tot_row})')
        c4.number_format = PCT
        c4.font = BODY_FONT
    ws4.cell(row=tot_row + 1, column=1,
             value="闭合差（透视各行合计 − 明细全口径合计；必须恒为 0）").font = WARN_FONT
    c = ws4.cell(row=tot_row + 1, column=3,
                 value=f'=SUM(C{sr}:C{tot_row - 1})-C{tot_row}')
    c.number_format = MONEY
    c.font = WARN_FONT

    # 待取证清单
    tp = tot_row + 3
    ws4.cell(row=tp, column=1,
             value="待取证清单（按 (High − Low) × Qty 降序，前 10）").font = SEC_FONT
    write_header(ws4, tp + 1, ["Part No", "Name", "方法", "区间宽度×数量", "优先级"],
                 [16, 34, 10, 15, 10])
    ranked = sorted(parts, key=lambda p: -(((p["high"] or 0) - (p["low"] or 0)) * p["qty"]))
    tr = tp + 2
    for i, p in enumerate([x for x in ranked if x["priced"]][:10], start=1):
        ws4.cell(row=tr, column=1, value=p["part"]).font = BODY_FONT
        ws4.cell(row=tr, column=2, value=p["name"]).font = BODY_FONT
        ws4.cell(row=tr, column=3, value=p["method"]).font = BODY_FONT
        c = ws4.cell(row=tr, column=4,
                     value=f'=IFERROR((\'2-BOM\'!${B["p_high"]}{p["row"]}'
                           f'-\'2-BOM\'!${B["p_low"]}{p["row"]})'
                           f'*\'2-BOM\'!${B["qty"]}{p["row"]},"")')
        c.number_format = MONEY
        ws4.cell(row=tr, column=5, value=i).font = BODY_FONT
        tr += 1
    unpriced = [p for p in parts if not p["priced"]]
    if unpriced:
        tr += 1
        ws4.cell(row=tr, column=1,
                 value="未定价行（不计入合计，需补取证价或型号）").font = WARN_FONT
        tr += 1
        for p in unpriced:
            ws4.cell(row=tr, column=1, value=p["part"]).font = BODY_FONT
            ws4.cell(row=tr, column=2, value=p["name"]).font = BODY_FONT
            ws4.cell(row=tr, column=3, value=p["reason"][:250]).font = NOTE_FONT
            tr += 1

    # PPT_EXPORT
    px = tr + 2
    ws4.cell(row=px, column=1,
             value="PPT_EXPORT（固定块，供幻灯片生成器按名取值）").font = SEC_FONT
    write_header(ws4, px + 1, ["键", "值", "说明"], [26, 32, 58])
    ppr = px + 2
    ppt = [
        ("completeness", f'=IF({nonanch}=0,"complete","partial")',
         "存在 M5 或预留即为 partial；partial 时图表标题须含『不完整小计』并附未定价清单"),
        # ★ 图表标题要用的**成品字符串**：不靠人去拼，避免在 PPT 里手写错口径
        ("subtotal_label",
         f'=IF({nonanch}=0,"完整小计",'
         f'"不完整小计（缺价 "&COUNTIFS({rt_rng},"<>预留",{m_rng},"未定价")&" 件未计入）")',
         "直接粘为图表/表头标题；partial 时缺价件数由公式实时计数（与 A1 同口径）"),
        ("m5_reserve_share", f"={nonanch}", "M5 与预留占全口径 MAT"),
        ("mat_anchored", f"=B{r_anch}", "M1–M4"),
        ("mat_full", f"=B{r_full}", "含 M5 与预留"),
        ("va_base", f"=B{r_va}", "VA 合计（PDC + PDOH + PAC + MOH）"),
        ("ppc_base", f"=B{r_ppc}", "MAT + VA"),
        ("ppc_low", f"=B{r_plo}", "区间下界"),
        ("ppc_high", f"=B{r_phi}", "区间上界"),
        ("caliber", "PPC = MAT + VA（Bosch ME / PCB-CN 口径）", "口径声明"),
        ("price_base_year", "=" + ref["PRICE_BASE"], "价格基期"),
        ("target_year", "=" + ref["TARGET_YEAR"], "目标年份"),
        ("level", level_formula, "结论等级"),
        ("topn", f'=IF({nonanch}<={ref["TH_M5_L1"]},"允许","不导出")',
         "Top-N 单件排序导出开关；等级 2 及以下不导出"),
        ("unpriced_n", n_unpriced, "未定价件数；partial 时必须同屏"),
    ]
    for k, v, nte in ppt:
        ws4.cell(row=ppr, column=1, value=k).font = BODY_FONT
        # ★ 第 2 列可能是公式或纯文本（caliber），公式必须原样写入
        c = ws4.cell(row=ppr, column=2, value=v)
        c.font = BODY_FONT
        if k == "m5_reserve_share":
            c.number_format = PCT
        elif k in ("mat_anchored", "mat_full", "va_base", "ppc_base", "ppc_low", "ppc_high"):
            c.number_format = MONEY
        ws4.cell(row=ppr, column=3, value=nte).font = NOTE_FONT
        ppr += 1

    # 结论句（全部数字来自公式，无手写数）
    ws4["A1"] = (
        f'="清单 "&COUNTIF({rt_rng},"<>预留")&" 个实物行,已出价 "'
        f'&COUNTIFS({rt_rng},"<>预留",{m_rng},"<>未定价")&",待补 "'
        f'&COUNTIFS({rt_rng},"<>预留",{m_rng},"未定价")'
        f'&";估算 "&TEXT(B{r_ppc},"#,##0")&" 元/台(区间 "&TEXT(B{r_plo},"#,##0")'
        f'&"–"&TEXT(B{r_phi},"#,##0")&"),其中 M5 与预留占 "&TEXT({nonanch},"0.0%")&"。"'
    )
    ws4["A1"].font = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
    ws4.freeze_panes = "A3"

    # ============================================================ 5-Opportunities
    ws5 = wb.create_sheet("5-Opportunities")
    opp_labels, opp_widths, O = make_cols(OPP_SPEC)
    ws5["A1"] = "降本机会 —— 毛 / Effort / 净分列；同一成本项目只计一次"
    ws5["A1"].font = TITLE_FONT
    ws5["A2"] = ("★ 净潜力 = 毛潜力 − Effort，仅当两者均为 CNY/pc；工装（kCNY / mCNY）另列，"
                 "不折算进单件。同防重复计数组的措施不得重复计节省。")
    ws5["A2"].font = WARN_FONT
    write_header(ws5, 4, opp_labels, opp_widths)
    allowed = {"mat_imp": ("MAT", "无", ""), "va_imp": ("VA", "无", ""),
               "tool_imp": ("Tooling", "无", "")}
    orow = 5
    for o in (opps or []):
        # 归属列语义校验：值必须与列名一致，不能只判非空
        for k, ok in allowed.items():
            if o[k] not in ok:
                o[k] = "无"
        net = (f'=IF(AND(ISNUMBER(${O["gross"]}{orow}),ISNUMBER(${O["effort"]}{orow})),'
               f'${O["gross"]}{orow}-${O["effort"]}{orow},"")')
        write_body(ws5, orow, [
            o["id"], o["title"], o["now"], o["alt"], o["mech"],
            o["gross"], o["effort"], net, o["mat_imp"], o["va_imp"], o["tool_imp"],
            o["in_dtc"], o["in_rpp"], o["phase"], o["mech_tag"], o["dedup"],
            o["hg"], o["conf"], o["risk"], o["weeks"], o["status"], o["prio"],
        ], "Estimated")
        for k in ("gross", "effort", "net"):
            ws5[f"{O[k]}{orow}"].number_format = MONEY
        orow += 1
    if not opps:
        orow = note(ws5, orow, "未提供机会表：本表为空。降本次目属产品数据，不由生成器编造。",
                    WARN_FONT)
    orow = note(ws5, orow + 1,
                "★ MAT Impact 只接受 MAT/无/空；VA Impact 只接受 VA/无/空；"
                "Tooling Impact 只接受 Tooling/无/空。")
    orow = note(ws5, orow,
                "★ HG 记实施成熟度（0–5），证据可信度（Verified / Modeled / Hypothetical）"
                "另记一列，两坐标独立。")
    orow = note(ws5, orow, "★ 外部条目一律 Hypothetical：不得从 CSV 带进 Verified / HG4 / HG5。")
    ws5.freeze_panes = "C5"

    # ============================================================ 6-Sensitivity
    ws6 = wb.create_sheet("6-Sensitivity")
    sens_labels, sens_widths, S = make_cols(SENS_SPEC)
    ws6["A1"] = "敏感性 —— 单因子 Low/Base/High；影响额取绝对值，方向另列"
    ws6["A1"].font = TITLE_FONT
    ws6["A2"] = ("★ 影响额取绝对值 + 单列方向说明；取绝对值后不得再用符号判断增减。"
                 "禁止『总成本 ±10%』式假敏感性。")
    ws6["A2"].font = WARN_FONT
    write_header(ws6, 4, sens_labels, sens_widths)
    sens = [
        ("材料单价 ±10%", "1-Assumptions 材料单价组", "-10%", "0%", "+10%",
         "因子↑→成本↑", "MAT 占比高时影响最大；只作用于 M2 材料成本",
         "AL_PRICE / CU_PRICE / …"),
        ("年产量 ×0.5 / ×1 / ×2", "VOL_A", "×0.5", "×1", "×2",
         "因子↑（产量）→单件成本↓", "影响折旧与摊销；产量是竞品 VA 的首要敏感性", "VOL_A"),
        ("VA 比例 k", "K_VA_BASE", f'={ref["K_VA_LOW"]}', f'={ref["K_VA_BASE"]}',
         f'={ref["K_VA_HIGH"]}',
         "因子↑→成本↑", "L2 口径下 VA = 未建工序 MAT × k", "K_VA_LOW / BASE / HIGH"),
        ("M5 区间端点", "全口径 MAT", "全部 M5 取 Low", "均值 / 几何均值",
         "全部 M5 取 High", "因子↑→成本↑",
         "本行为区间而非点值，不给单一影响额", "2-BOM Price Low / High"),
        ("PCB 层数与铜厚", "M3 校准锚点", "4L 860–1460 CNY/m²",
         "层数未知 → 父档 900–1800", "8L+HSP / thick-Cu 2916 / 2501",
         "因子↑→成本↑", "层数未知回退父档会系统性低估 6L / HSP / thick-Cu", "pcb_* 子档"),
        ("供应商 SG&A / 利润率", "SGA_RATE / PROFIT_RATE", "0.03 / 0.06", "0.05 / 0.08",
         "0.08 / 0.12", "因子↑→成本↑", "只作用于 M2 零件级计算器",
         "SGA_RATE / PROFIT_RATE"),
    ]
    srow = 5
    for drv, pbase, lo, ba, hi, dirc, nte, src in sens:
        write_body(ws6, srow, [drv, pbase, lo, ba, hi, None, None, dirc, nte, src],
                   "Estimated")
        ws6[f"{S['impact']}{srow}"].number_format = MONEY
        ws6[f"{S['share']}{srow}"].number_format = PCT
        srow += 1
    srow = note(ws6, srow + 1,
                "★ 影响额列留待逐因子建模后填入；未建模前留空并在本行声明，"
                "不得用 0 冒充『无影响』。", WARN_FONT)
    srow = note(ws6, srow,
                "★ 一旦建模，影响额须取 ABS() 并把方向写进方向列——"
                "取绝对值后再用符号判增减会得出相反结论。")
    ws6.freeze_panes = "A5"

    # ============================================================ 7-Evidence
    ws7 = wb.create_sheet("7-Evidence")
    evd_labels, evd_widths, _ = make_cols(EVD_SPEC)
    ws7["A1"] = "证据与缺口 —— 来源写到可复核粒度；待确认问题写『缺什么证据能回答』"
    ws7["A1"].font = TITLE_FONT
    write_header(ws7, 3, evd_labels, evd_widths)
    m5_parts = [p for p in parts if p["method"] == "M5"]
    wide = [p for p in parts if p["tier"] == "[宽>10×]"]
    ev_rows = [
        ("M5 清单", f"{len(m5_parts)} 行使用 M5 粗估（含质量轴回退）；不得用于定点谈判",
         "2-BOM Price Method = M5", "Estimated", "低", "MAT",
         "逐件补取证价，或先补型号 / 规格建立 M2", ""),
        ("[宽>10×] 清单", f"{len(wide)} 行区间宽于 10×",
         "2-BOM Unit Price Basis 含 [宽>10×]", "Estimated", "低", "MAT",
         "先用可见属性拆子档；拆不开再补取证", ""),
        ("未定价行", f"{len(unpriced)} 行无价（不计入合计，原因见 4-Summary）",
         "2-BOM Counted Note", "Unknown", "无", "MAT",
         "缺什么：型号 / 规格 / 供应商报价", ""),
        ("VA 方法层级",
         f"VA = L0（有工序行的件）+ L2（k 覆盖未建工序的 MAT）混合口径；"
         f"k = {PV['K_VA_LOW']:.3f} / {PV['K_VA_BASE']:.3f} / {PV['K_VA_HIGH']:.3f}",
         "3-ProcessCost 方法层级列", "Estimated", "中", "VA",
         "缺什么：供应商 SMT 单点成本、产量、人工工作量（HCi）", ""),
        ("PDC 覆盖", f"{len(routed_parts)} 个件建了 L0 工序；"
                  "其余件分两类：M1/M3/M4/M5 的加工已含在交付价里，"
                  "M2 且无工序行的件由 L2 比例 k 覆盖",
         "3-ProcessCost Step 列", "Estimated", "中", "VA",
         "缺什么：自制 / 外购的真实分工与采购对象", ""),
        ("MOH 口径", "MOH 只用于 VA 内部拆分，未在 VA 之上叠加",
         "1-Assumptions MOH_RATIO", "Estimated", "中", "VA",
         "缺什么：报告的分项间接费比例表", ""),
        ("隐含件预留", "Chem+Scrap 3.4%–4.5%、紧固件 0.7%–1.6%；仅用于定尺寸后算预留",
         "1-Assumptions S_CHEM / S_FAST", "Estimated", "中", "MAT",
         "缺什么：化学品类净重（拆解称重未覆盖）", ""),
        ("工装", "件价是否含工装由 Tooling In Price 承载；未明说时记 '?'；"
                 "采购工装摊销只在 MOH 计一次",
         "2-BOM Tooling In Price", "Estimated", "中", "工装摊销",
         "缺什么：模具费报价、寿命模次", ""),
        ("产量", "年产量与生命周期产量为假设值，折旧与摊销因此不确定",
         "1-Assumptions VOL_A / VOL_L", "Unknown", "无", "折旧 / 摊销",
         "缺什么：项目产量规划", ""),
        ("价格基期", f"全部价格为 @{PV['PRICE_BASE']}，未做跨年调整",
         "1-Assumptions PRICE_BASE", "Estimated", "中", "全部",
         "缺什么：目标年份的物料价格指数", ""),
        ("证据状态口径", "竞品成本最高 Estimated；计算状态 F（公式可复算）/ M（手输），"
                     "F 不代表已验证",
         "本表", "Verified", "高", "—", "无需补充", ""),
        ("版本说明", "本工作簿为 v1x 口径（PPC = MAT + VA，M1–M5）；与 v2（JSON 驱动）不互读",
         "SKILL.md 版本关系", "Verified", "高", "—",
         "两版都建过时须在此写明采用哪一版", ""),
        ("来源保密", "基准数据来自 Bosch Confidential 报告：引用保留来源，"
                 "按报告访问范围存放，不外发",
         "Bosch CTM / P2 / PTU", "Verified", "高", "—",
         "不得把数值放到公开渠道检索", ""),
    ]
    e = 4
    for row in ev_rows:
        write_body(ws7, e, list(row), row[3])
        e += 1
    e = note(ws7, e + 1,
             "★ 待确认缩写原样保留：PDOH、RPP、RBHP、EVOB、PMoC、PM5、HCi、PCi、HSP、"
             "CIPG、EAC、DCL。", WARN_FONT)
    ws7.freeze_panes = "A4"

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"out": str(out), "rows": len(parts), "reserve_rows": n_res,
            "priced": n_priced, "unpriced": n_unpriced, "m5": n_m5,
            "pdc_parts": len(routed_parts), "va_total_row": va_total_row,
            "ppc_row": r_ppc, "nonanch_row": rn, "bom_last": bom_last}


# ================================================================ CLI
def main() -> int:
    ap = argparse.ArgumentParser(description="Should-Cost 工作簿生成器 v1x")
    ap.add_argument("--bom", required=True, help="叶件清单 CSV")
    ap.add_argument("--out", required=True, help="输出 xlsx")
    ap.add_argument("--attrs", help="属性表 CSV（九类材料分层 / 采购对象，可选）")
    ap.add_argument("--prices", help="取证价 CSV（可选，M1）")
    ap.add_argument("--opps", help="降本机会 CSV（可选）")
    ap.add_argument("--catalog", help="价格目录 CSV（默认 v1x 目录）")
    ap.add_argument("--product", default="Product")
    ap.add_argument("--function-prefix", default=None)
    ap.add_argument("--no-catalog-fill", action="store_true",
                    help="关掉目录补价与质量轴回退（这些行恢复未定价）")
    ap.add_argument("--no-benchmark-fill", action="store_true",
                    help="[已废弃] 旧名，等同 --no-catalog-fill")
    a = ap.parse_args()

    if a.no_benchmark_fill:
        print("WARN: --no-benchmark-fill 已废弃，改用 --no-catalog-fill", file=sys.stderr)
        a.no_catalog_fill = True

    leaves = load_leaves(Path(a.bom))
    catalog = load_catalog(Path(a.catalog) if a.catalog else None)
    if a.no_catalog_fill:
        catalog = {**catalog, "rows": [], "by_key": {}, "by_parent": {}}
    info = build(
        leaves, Path(a.out), product=a.product,
        attrs=load_attrs(Path(a.attrs) if a.attrs else None),
        prices=load_prices(Path(a.prices) if a.prices else None),
        catalog=catalog,
        opps=load_opps(Path(a.opps) if a.opps else None),
        src_names={"leaves": Path(a.bom).name},
        func_prefix=a.function_prefix,
    )
    print("OK", info)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
