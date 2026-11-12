#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把宽表里"标准件行"（紧固件 / 螺母 / O 圈）补进叶件清单，并同步生成 attrs / price 清单。

为什么需要这一步
================
`build_leaf_list.py` 的叶件过滤是 `r["path"] not in prefix and r["total_weight"]`。
后半个条件会导致**所有 Total Weight 为空的行被静默丢弃**——在 DC-DC 子树里就是
19 行紧固/密封件（77 件实物、14 个料号）。它们的重量已经包含在各自父件的
`Total Weight` 里（父件 Total = Part Weight + Fastener Weight），所以**重量没丢**；
但它们的**成本从来没进过 BOM**。

本脚本做三件事：
1. 生成 `*-leaves-v3.csv`：原叶件 + 19 个标准件行（Total Weight 置 0，避免与父件重复计重）；
2. 生成 `*-purchased-attributes-v3.csv`：标准件行的材料分层全部留空
   （它们的钢/弹性体重量已计入父件的 STEEL / ELASTOMER 列，此处再填就是重复计量）；
3. 生成 `*-price-evidence-v7.csv`：为标准件补 D 类价格（按螺纹规格 / O 圈内径分档）。

用法
----
    py -3 -X utf8 tools/extend_leaves_with_standard_parts.py ^
        --subtree data/ex90-ultra-DCDC/derived/dcdc-subtree.csv ^
        --leaves  data/ex90-ultra-DCDC/derived/dcdc-leaves.csv ^
        --attrs   data/ex90-ultra-DCDC/derived/dcdc-purchased-attributes-v2.csv ^
        --prices  data/ex90-ultra-DCDC/derived/dcdc-price-evidence-v6.csv ^
        --prefix  data/ex90-ultra-DCDC/derived/dcdc
"""
from __future__ import annotations

import argparse
import csv
from pathlib import Path

LEAF_FIELDS = [
    "CS Part", "Name", "Quantity", "Part Weight", "Fastener Weight", "Total Weight (kg)",
    "Material Group", "Material Species", "Material Species (Tested)", "Manufacturer",
    "Marking Number", "Length", "Width", "Height", "Thickness (T)", "Part Type",
    "Number of Interface", "Depth", "Parent Part No", "Mass Accounting",
]

# 标准件价格（元/件，Delivered at assembly plant）。全部 D 类：按几何规格分档类比外推。
# (Low, High, 分档依据)
FASTENER_PRICE = {
    # ---- 螺钉（Torx，钢，Pan/Washer head）
    "S0B00100046ABC": (0.05, 0.14, "M4×8 Pan Torx Flat Point，ISO Class 8；按螺纹规格 M4 档（0.05–0.14 元/件）"),
    "S0B00100046ABK": (0.05, 0.14, "M4×8 Pan Torx Flat Point，无强度等级记录；同 M4 档（0.05–0.14 元/件）"),
    "S0B00100143ABK": (0.08, 0.22, "M5×16 Pan Torx Flat Point；按螺纹规格 M5 档（0.08–0.22 元/件）"),
    "S0B00100074ABK": (0.12, 0.32, "M6×20 Pan Torx Flat Point；按螺纹规格 M6 档（0.12–0.32 元/件）"),
    "S0B01700039ABK": (0.07, 0.20, "M3×12 Pan Torx Flat Point With Washer（自带垫圈 OD 7 mm）；M3 带垫圈档"),
    "S0B50600305A0G": (0.10, 0.28, "M5×12.1 Washer Head Torx，ISO Class 10.9（法兰面整体垫）；M5 法兰档"),
    "S0B50600306ABK": (0.07, 0.20, "M4×8.1 Washer Head Torx；M4 法兰档（法兰面比 Pan head 多一道成形）"),
    "S0B50600307ABK": (0.14, 0.38, "M6×14.1 Washer Head Torx；M6 法兰档"),
    # ---- 自锁螺母
    "S0C15200002MBK": (0.22, 0.60, "M6×1 Hexagonal Spin Washer Nyloc Nut，AF 10 / 法兰 Ø14.2，Self-locking=Yes；"
                                   "尼龙自锁 + 法兰面 + 两种表面处理（Grey / Silver）→ 单件成本高于普通螺母 3–5 倍"),
    # ---- O 圈（弹性体，线径 × 内径决定单价）
    "S0H0010000051D": (0.03, 0.12, "O Ring 内径 7 × 线径 1.5 mm；小截面小内径档"),
    "S0H0010000246D": (0.03, 0.12, "O Ring 内径 5.5 × 线径 2.5 mm；小内径档"),
    "S0H0010000124D": (0.04, 0.15, "O Ring 内径 16 × 线径 2 mm；中内径档"),
    "S0H0010000126D": (0.04, 0.15, "O Ring 内径 17 × 线径 2 mm；中内径档"),
    "S0H0010000274D": (0.05, 0.18, "O Ring 内径 21 × 线径 2.5 mm；大内径档"),
}

# 标准件的 Procurement Object 分类（供 4-Summary 用 SUMIFS 汇总紧固件包）
def object_of(name: str) -> str:
    n = name.strip().lower()
    if "o ring" in n:
        return "Elastomer seal"
    return "Standard fastener"


def read_csv(path: Path) -> tuple[list[str], list[dict]]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        return list(rd.fieldnames or []), list(rd)


def write_csv(path: Path, fields: list[str], rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def main() -> int:
    ap = argparse.ArgumentParser(description="把标准件行补进叶件清单")
    ap.add_argument("--subtree", type=Path, required=True)
    ap.add_argument("--leaves", type=Path, required=True)
    ap.add_argument("--attrs", type=Path, required=True)
    ap.add_argument("--prices", type=Path, required=True)
    ap.add_argument("--prefix", type=Path, required=True, help="输出前缀，如 data/.../derived/dcdc")
    a = ap.parse_args()

    _, sub = read_csv(a.subtree)
    leaf_fields, leaves = read_csv(a.leaves)
    attr_fields, attrs = read_csv(a.attrs)
    price_fields, prices = read_csv(a.prices)

    # --- 1. 定位标准件行，并回溯其父件件号（最近的上一个非 Standard 行）
    parent_part = ""
    std_rows: list[dict] = []
    for r in sub:
        ptype = (r.get("part_type") or "").strip()
        if ptype == "Standard":
            std_rows.append({**r, "_parent": parent_part})
        else:
            cs = (r.get("cs_part") or "").strip()
            if cs:
                parent_part = cs

    if not std_rows:
        raise SystemExit("subtree 里没有 Part Type = Standard 的行，检查输入")

    # --- 2. 按料号归并（14 个料号 / 19 个挂靠位置 / 77 件）
    #     ★ 不按挂靠位置拆行：同一料号拆成三行（如 M5×16 在 Top Cover 10 件、
    #     Terminal Output 2 件、Terminal Input 2 件）会让审阅者怀疑重复计价，
    #     而 Should-Cost 关心的是『料号 × 总用量 × 单价』。挂靠关系写进说明列。
    std_by_part: dict[str, dict] = {}
    for r in std_rows:
        cs = (r.get("cs_part") or "").strip()
        g = std_by_part.setdefault(cs, {
            "name": (r.get("name") or "").strip(),
            "qty": 0.0, "wt": 0.0, "parents": [], "group": (r.get("material_group") or "").strip(),
            "species": (r.get("material_species") or "").strip(),
            "maker": (r.get("manufacturer") or "").strip(),
        })
        g["qty"] += float(r.get("qty") or 0)
        g["wt"] += float(r.get("fastener_weight") or 0)
        p = r["_parent"]
        if p and p not in g["parents"]:
            g["parents"].append(p)

    # --- 3. 生成 leaves v3
    out_leaves: list[dict] = []
    for r in leaves:
        out_leaves.append({**r, "Parent Part No": "",
                           "Mass Accounting": "本行 Total Weight 独立计入质量基线"})
    seen_parts = {str(r["CS Part"]).strip() for r in leaves}
    for cs, g in std_by_part.items():
        if cs in seen_parts:
            continue  # 已作为叶件计价（本数据集里不会发生，留作保险）
        out_leaves.append({
            "CS Part": cs,
            "Name": g["name"],
            "Quantity": int(g["qty"]) if float(g["qty"]).is_integer() else g["qty"],
            "Part Weight": "",
            "Fastener Weight": round(g["wt"], 4),
            # ★ 必须为空：该重量已计入父件 Total Weight，此处再计就是重复计重
            "Total Weight (kg)": "",
            "Material Group": g["group"],
            "Material Species": g["species"],
            "Material Species (Tested)": "",
            "Manufacturer": g["maker"],
            "Marking Number": "",
            "Length": "", "Width": "", "Height": "", "Thickness (T)": "",
            "Part Type": "Standard",
            "Number of Interface": "",
            "Depth": "",
            # 挂靠位置（多父件时以 / 分隔）
            "Parent Part No": " / ".join(g["parents"]),
            "Mass Accounting": "紧固件重量已含在父件 Total Weight 内；本行只计成本，不计量重",
        })

    # --- 4. 生成 attrs v3（标准件不做 bottom-up，材料分层全部留空）
    out_attrs: list[dict] = [dict(r) for r in attrs]
    have = {str(r.get("CS Part") or "").strip() for r in attrs}
    for cs, g in std_by_part.items():
        if cs in have:
            continue
        have.add(cs)
        out_attrs.append({
            "CS Part": cs,
            "Procurement Object": object_of(g["name"]),
            "Bottom-up Model (Y/N)": "N",
            "Route Status": "N/A",
            "Material Split Basis": "不适用：标准件按采购价计价，不做制造破拆；"
                                    "其钢/弹性体重量已计入父件的 STEEL / ELASTOMER 列，此处填数会重复计量",
            "Notes": f"{g['name']}；挂靠 {len(g['parents'])} 个父件：{' / '.join(g['parents'])}",
            "Correction Note": "v3 补齐：v2 因 build_leaf_list.py 的 `and r[\"total_weight\"]` "
                               "过滤而被整行丢弃，导致 0 元计价",
        })

    # --- 5. 生成 price v7
    out_prices: list[dict] = [dict(r) for r in prices]
    havep = {str(r.get("CS Part") or "").strip() for r in prices}
    missing = []
    for cs, g in std_by_part.items():
        if cs in havep:
            continue
        havep.add(cs)
        if cs not in FASTENER_PRICE:
            missing.append(cs)
            continue
        lo, hi, basis = FASTENER_PRICE[cs]
        out_prices.append({
            "CS Part": cs,
            "Procurement Object": object_of(g["name"]),
            "Price Low (CNY)": lo,
            "Price High (CNY)": hi,
            "Price Class": "D",
            "Price Type": "Unsupported Assumption",
            "Price Status": "Estimated",
            "Price Basis": f"{basis}。★ 无供应商报价、无实际采购价；"
                           f"按件号前缀（S0B=标准螺纹件 / S0H=密封件 / S0C=螺母）"
                           f"判定为量产通用标准件，非 OEM 专用件。",
            "Bare PCB Proxy (CNY)": "",
            "Bare PCB Proxy Basis": "",
            "Delivered Price Boundary": "Estimated delivered at assembly plant",
            "Source Date": "2026-11-04",
        })
    if missing:
        raise SystemExit(f"以下标准件料号缺价格，必须先补进 FASTENER_PRICE：{missing}")

    lp = a.prefix
    write_csv(lp.with_name(lp.name + "-leaves-v3.csv"), LEAF_FIELDS, out_leaves)
    write_csv(lp.with_name(lp.name + "-purchased-attributes-v3.csv"), attr_fields, out_attrs)
    write_csv(lp.with_name(lp.name + "-price-evidence-v7.csv"), price_fields, out_prices)

    lo_t = sum(FASTENER_PRICE[cs][0] * g["qty"] for cs, g in std_by_part.items())
    hi_t = sum(FASTENER_PRICE[cs][1] * g["qty"] for cs, g in std_by_part.items())
    tot_qty = sum(g["qty"] for g in std_by_part.values())
    tot_wt = sum(g["wt"] for g in std_by_part.values())
    print(f"标准件：{len(std_rows)} 个挂靠位置 / 归并为 {len(std_by_part)} 个料号 / "
          f"实物 {tot_qty:.0f} 件 / 紧固件重 {tot_wt:.3f} kg（已含在父件内，不再计入基线）")
    print(f"BOM 行数：{len(leaves)} → {len(out_leaves)}")
    print(f"标准件价格合计：{lo_t:.2f} / {hi_t:.2f} 元/台（Base "
          f"{sum((FASTENER_PRICE[c][0]+FASTENER_PRICE[c][1])/2*g['qty'] for c, g in std_by_part.items()):.2f}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())