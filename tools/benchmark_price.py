# -*- coding: utf-8 -*-
"""基准价目录（Class E）匹配器 —— 通用、与产品无关。

★ 设计边界（见 SKILL.md Step 5.8，2026-11-05 修订：**每行都要有价，取消留空**）：

    1. **只在「既无取证价、又没建 bottom-up 模型」时被调用**。
       调用方负责这个门控；本模块不做优先级判断。
    2. **三级降精度，绝不跳类别**：
         3a 严格类别（Anchor Kind = Benchmarked / Derived）→ Price Status = Comparable
         3b 包络类别（Anchor Kind = Envelope）        → Price Status = Assumption
         3c 质量轴兜底（净重 × 材料元/kg 带）        → Price Status = Assumption
       3c **只适用于结构件**（金属/塑料/弹性体/泡棉），
       绝不用于电子件——按重量给 MOSFET 定价得到的是假数，不是粗数。
    3. **只给区间，不给点值**：均值由工作簿的 Price Base 公式自己算（AVERAGE / GEOMEAN）。
    4. **目录里没有产品的件号、产品名、文件名、用量**；
       补目录只增加「类别」，绝不增加件号或产品名。

目录列（data/benchmark/benchmark-price-catalog.csv）：

    Category Key / Category (archetype) / Match Head Prefix / Species Contains /
    Species Excludes / Basis Unit / Quantity Basis / Low / High / Anchor Kind /
    Match Confidence / Use / Rationale / Source / Source Date / Review Due / Notes

`Use` 语义：

    Fill        → 件类别补位（3a/3b）
    MassAxis    → 质量轴兜底（3c）
    Calibration → 只用于反查 bottom-up 模型的合理性，**不参与补位**

`Anchor Kind` 语义：

    Benchmarked → 有实证锚点（检索/取证所得）   → 3a
    Derived     → 由两个已有锚点派生          → 3a
    Envelope    → 工程包络，无实证锚点         → 3b / 3c
"""

from __future__ import annotations

import csv
import io
from pathlib import Path

CATALOG_DEFAULT = Path(__file__).resolve().parent.parent / "data" / "benchmark" / "benchmark-price-catalog.csv"

# 3c 只对这些材料组开放（电子件/线束/胶粘剂不在其列）
MASS_AXIS_GROUPS = {
    "metal": "mass_axis_metal",
    "plastic": "mass_axis_plastic",
    "elastomer": "mass_axis_elastomer",
    "foam": "mass_axis_foam",
}

# 电子件关键词：出现即禁止 3c（按重量给电子件定价 = 假数）
ELECTRONIC_HINTS = (
    "ee components", "sensor", "capacitor", "pcb", "connector",
    "fuse", "mosfet", "igbt", "diode", "ferrite", "inductor", "transformer",
    "wire harness", "adhesive", "sealants", "paper", "sticker", "ceramic",
)


def name_subject(name: str) -> tuple[str, str]:
    """件名 → (head, full)，与 domain_classify.name_subject 同构。

    CS 件名结构为 `<零件类型>, <父级上下文>`：
        "Gasket, Rear Inverter Cooling Plate" → head = "gasket"
        "Cover, Rear Inverter PCB"            → head = "cover"

    ★ 匹配只看 head。若看全文，父级上下文会让
      "Bracket, Rear Inverter Power Module" 去命中「Power Module」类，
      而 "Cover, ... PCB" 会命中「PCB」类 —— 同类件在不同父件下得到不同价。
    """
    n = (name or "").strip()
    head = n.split(",")[0].strip().lower()
    return head, n.lower()


def _split(s: str | None) -> list[str]:
    return [x.strip().lower() for x in (s or "").split(";") if x.strip()]


def load_catalog(path: Path | None = None) -> dict:
    """读目录 → {"fill": [...], "mass": [...]}。

    ★ 只装载 Use ∈ {Fill, MassAxis} 且 Low/High 都有效的行；
      Calibration 行完全不进候选集（它们对应件已建 bottom-up 模型，
      按匹配顺序第 2 步就不该被覆盖）。
    """
    p = Path(path) if path else CATALOG_DEFAULT
    if not p.exists():
        return {"fill": [], "mass": []}
    fill: list[dict] = []
    mass: list[dict] = []
    with io.open(p, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            use = (r.get("Use") or "").strip().lower()
            if use not in ("fill", "massaxis"):
                continue
            lo_raw = (r.get("Low (CNY)") or "").strip()
            hi_raw = (r.get("High (CNY)") or "").strip()
            if not lo_raw or not hi_raw:
                continue
            try:
                lo, hi = float(lo_raw), float(hi_raw)
            except ValueError:
                continue
            if lo <= 0 or hi < lo:
                continue
            row = {
                "key": (r.get("Category Key") or "").strip(),
                "category": (r.get("Category (archetype)") or "").strip(),
                "prefixes": _split(r.get("Match Head Prefix (;-sep, first hit wins)")),
                "sp_in": _split(r.get("Species Contains (optional)")),
                "sp_out": _split(r.get("Species Excludes")),
                "basis_unit": (r.get("Basis Unit") or "").strip(),
                "qty_basis": (r.get("Quantity Basis") or "").strip(),
                "low": lo, "high": hi,
                "anchor_kind": (r.get("Anchor Kind") or "").strip() or "Benchmarked",
                "confidence": (r.get("Match Confidence") or "").strip(),
                "rationale": (r.get("Rationale") or "").strip(),
                "source": (r.get("Source") or "").strip(),
                "src_date": (r.get("Source Date") or "").strip(),
                "review_due": (r.get("Review Due") or "").strip(),
                "notes": (r.get("Notes") or "").strip(),
            }
            (mass if use == "massaxis" else fill).append(row)
    return {"fill": fill, "mass": mass}


def _match_row(head: str, species: str, catalog: list[dict]) -> dict | None:
    """在目录中找一个**语义明确**的类别。命中即停（目录顺序既优先级）。

    命中条件（全部满足）：
        1. head 以某个 Match Head Prefix 开头（不是"包含"，避免
           "Bracket" 命中 "Bracket, ... Busbar" 这种上下文串味）；
        2. Species Excludes 里的词一个都不出现；
        3. Species Contains 为空 → 不限材料；非空 → 至少一个出现。
    """
    s = (species or "").lower()
    for row in catalog:
        if any(x in s for x in row["sp_out"]):
            continue
        if not any(head.startswith(p) for p in row["prefixes"]):
            continue
        if row["sp_in"] and not any(x in s for x in row["sp_in"]):
            continue
        return row
    return None


def _convert(row: dict, net_weight: float, length: float, width: float):
    """按计价基数换算成单件 Low/High。任一必需量缺失 → None（绝不猜）。

    ★ 三种基数：
        CNY/piece  → 直接用
        CNY/kg     → × 净重（净重 ≤ 0 → None，不猜重量）
        CNY/m2     → × (长×宽/1e6)（长或宽缺失 → None，不猜尺寸）
    """
    unit = row["basis_unit"]
    if unit == "CNY/piece":
        return row["low"], row["high"], "按件"
    if unit == "CNY/kg":
        w = net_weight or 0.0
        if w <= 0:
            return None
        return row["low"] * w, row["high"] * w, f"× 净重 {w:.4f} kg"
    if unit == "CNY/m2":
        if not length or not width:
            return None
        area = (length / 1000.0) * (width / 1000.0)
        return (row["low"] * area, row["high"] * area,
                f"× 面积 {area:.5f} m²（{length:g}×{width:g} mm）")
    return None


def _mass_row(group: str, species: str, mass_rows: list[dict]) -> dict | None:
    """3c：按材料组选质量轴。电子件/线束/胶粘剂一律返回 None。"""
    g = (group or "").strip().lower()
    sp = (species or "").lower()
    if g in ("ee components", "wire harness", "adhesives & sealants", "others"):
        return None
    if any(h in sp for h in ELECTRONIC_HINTS):
        return None
    key = MASS_AXIS_GROUPS.get(g)
    if key is None:
        # Mixed Material：按材料牌号猜组，仅限结构材料
        if "foam" in sp:
            key = "mass_axis_foam"
        elif any(x in sp for x in ("elastomer", "rubber", "epdm", "tpv", "aem", "silicone")):
            key = "mass_axis_elastomer"
        elif any(x in sp for x in ("plastic", "pbt", "pa6", "pa66", "pps", "ppa", "para", "pur")):
            key = "mass_axis_plastic"
        elif any(x in sp for x in ("aluminium", "steel", "copper", "brass", "metal")):
            key = "mass_axis_metal"
        else:
            return None
    for row in mass_rows:
        if row["key"] == key:
            return row
    return None


TIER_TAG = {"3a": "", "3b": "[包络]", "3c": "[兜底-质量轴]"}
TIER_STATUS = {"3a": "Comparable", "3b": "Assumption", "3c": "Assumption"}


def benchmark_price(name: str, species: str, net_weight: float,
                    length: float, width: float,
                    catalog: dict, material_group: str = "") -> dict | None:
    """件 → 基准价区间 + 层级。**返回 None 仅当三级全部落不到。**

    返回
        {
          "low", "high",         单件 Low/High（已按基数换算）
          "row",                  命中的目录行
          "tier",                 "3a" / "3b" / "3c"
          "status",               "Comparable" / "Assumption"
          "tag",                  Price Basis 前缀标记（3a 为空）
          "basis_text",           可直接写进 Price Basis 的自解释文本
        }

    ★ 三级顺序固定：3a/3b（件类别）优先于 3c（质量轴）。
      质量轴只在件类别落不到时才用——类别比分档更贴近实际成本。
    """
    head, _ = name_subject(name)

    row = _match_row(head, species, catalog["fill"])
    tier = None
    conv = None
    if row is not None:
        tier = "3a" if row["anchor_kind"] in ("Benchmarked", "Derived") else "3b"
        conv = _convert(row, net_weight, length, width)

    if conv is None:
        mrow = _mass_row(material_group, species, catalog["mass"])
        if mrow is not None:
            conv = _convert(mrow, net_weight, length, width)
            if conv is not None:
                row, tier = mrow, "3c"

    if conv is None or row is None:
        return None
    lo, hi, calc = conv
    if lo <= 0 or hi < lo:
        return None

    tag = TIER_TAG[tier]
    if tier == "3a":
        lead = f"基准价目录 {row['key']}（{row['category']}）；"
    elif tier == "3b":
        lead = (f"[包络] {row['key']}（{row['category']}）—— "
                f"该类别缺实证锚点，区间为行业最宽可信包络，须优先 RFQ；")
    else:
        lead = (f"[兜底-质量轴] {row['key']}（{row['category']}）—— "
                f"件类别无匹配，按该件自身净重定价；")

    basis = (f"{lead}{row['basis_unit']} {row['low']:g}–{row['high']:g} {calc}"
             f" = {lo:.3f}–{hi:.3f} 元/件。"
             f"来源：{row['source']}（{row['src_date']}）。"
             f"★ 行业基准区间，非报价；不得用于定点谈判。")
    if row["notes"]:
        basis += f" 备注：{row['notes']}"

    return {"low": round(lo, 4), "high": round(hi, 4), "row": row,
            "tier": tier, "status": TIER_STATUS[tier], "tag": tag,
            "basis_text": basis}


# ================================================================ 自检
def subset_check(catalog: dict, *, head: str, species: str,
                 expect_key: str | None) -> str:
    """调试用：报出该件命中了哪一类（或为什么没命中），供 tmp 探针打印。"""
    head_l = name_subject(head)[0]
    row = _match_row(head_l, species, catalog["fill"])
    if row is None:
        return f"NO MATCH  head={head_l!r} sp={species[:40]!r}"
    ok = "OK" if expect_key is None or row["key"] == expect_key else f"★MISMATCH expect {expect_key}"
    return f"{row['key']:<26} {ok}"


if __name__ == "__main__":
    import sys
    from collections import Counter

    cat = load_catalog()
    print(f"Fill rows: {len(cat['fill'])}   MassAxis rows: {len(cat['mass'])}")
    for r in cat["fill"]:
        print(f"  {r['anchor_kind']:<12}{r['key']:<26}{r['basis_unit']:<10}"
              f"{r['low']:>7g} - {r['high']:<8g} {r['prefixes'][:2]}")
    for r in cat["mass"]:
        print(f"  {r['anchor_kind']:<12}{r['key']:<26}{r['basis_unit']:<10}"
              f"{r['low']:>7g} - {r['high']:<8g} (mass axis)")
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        leaves = list(csv.DictReader(io.open(p, encoding="utf-8-sig")))
        print(f"\n--- {p.name}: 每件命中情况 ---")

        def _f(x):
            try:
                return float(x)
            except Exception:
                return 0.0

        tiers = Counter()
        for r in leaves:
            total = _f(r.get("Total Weight (kg)"))
            qty = _f(r.get("Quantity")) or 1.0
            res = benchmark_price(r.get("Name") or "", r.get("Material Species") or "",
                                  total / qty if qty else total,
                                  _f(r.get("Length")), _f(r.get("Width")), cat,
                                  r.get("Material Group") or "")
            if res:
                tiers[res["tier"]] += 1
                tag = f"[{res['tier']}] {res['low']:.3f}-{res['high']:.3f}  {res['row']['key']}"
            else:
                tiers["none"] += 1
                tag = "— NONE（三级全落不到）"
            print(f"  {r['Name'][:42]:<44}{tag}")
        print(f"\n层级分布: {dict(tiers)}")