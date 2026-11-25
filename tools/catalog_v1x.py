#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v1x 价格目录读取与 M1–M5 匹配（供 `build_should_cost_workbook_v1x.py` 使用）。

设计要点（对应 SKILL.md Step 4）
================================
* **按列名读取**，不按列序。17 个必需列缺失 → 直接报错；8 个可选列缺失 → 降级。
* `Method` 缺省时由 `Anchor Kind` 反推：`Calibrated→M3`、`Benchmarked→M4`、
  `Derived/Envelope/MassAxis→M5`。`Use=Calibration` 的行只做反查，不参与补位。
* **子档匹配**：`Subtier Of` 非空的行，只有其 `Required Attribute` **全部已知**才命中；
  任一属性缺失 → 回退父行（绝不猜）。
* `Match Head Prefix` 为空的 `Use=Fill` 行**不可能被名称匹配命中**
  （`startswith("")` 恒真会污染全部件名，因此空前缀列表直接跳过；
  这类行只能经 `Subtier Of` 子档匹配到达）。
* 计价基数换算：`CNY/piece` 直取；`CNY/kg` × 净重；`CNY/m2` × 面积（长×宽/1e6）。
  缺基数 → 返回 None（不猜），交回调用方标为「未定价」。

与 `tools/benchmark_price.py`（v1/v2 的 Class E 三级补位）的差别
--------------------------------------------------------------
旧实现用 `3a/3b/3c` 三级 + `Class E` 四元标签；本模块改为 **M1–M5 方法标签**，
并把「层级标记」换成 `[包络]` / `[宽>10×]` 可读标记。旧模块保持不动。
"""
from __future__ import annotations

import csv
from pathlib import Path

CATALOG_DEFAULT = Path("data/benchmark/benchmark-price-catalog-v1x.csv")

# 17 个必需列（缺任一列 → 报错，不允许静默降级）
REQUIRED_COLS = [
    "Category Key", "Category (archetype)", "Match Head Prefix (;-sep, first hit wins)",
    "Species Contains (optional)", "Species Excludes", "Basis Unit", "Quantity Basis",
    "Low (CNY)", "High (CNY)", "Anchor Kind", "Match Confidence", "Use",
    "Rationale", "Source", "Source Date", "Review Due", "Notes",
]
# 8 个可选列（缺 → 按规则降级，不报错）
OPTIONAL_COLS = [
    "Typical (CNY)", "Typical Source", "Tooling In Price", "Price Year",
    "Scope", "Method", "Required Attribute", "Subtier Of",
]

ANCHOR2METHOD = {
    "calibrated": "M3",
    "benchmarked": "M4",
    "derived": "M5",
    "envelope": "M5",
    "massaxis": "M5",
    "mass axis": "M5",
}
WIDE_RATIO = 10.0


def _f(v):
    try:
        s = (v or "").strip()
        return float(s) if s else None
    except ValueError:
        return None


def _split(v) -> list[str]:
    return [x.strip().lower() for x in (v or "").split(";") if x.strip()]


def load_catalog(path: Path | str | None = None) -> dict:
    """读目录 → {rows, by_key, by_parent, n_rows, missing_cols}"""
    p = Path(path) if path else CATALOG_DEFAULT
    if not p.exists():
        raise FileNotFoundError(f"价格目录不存在：{p}")
    with p.open(encoding="utf-8-sig", newline="") as fh:
        rdr = csv.DictReader(fh)
        hdr = rdr.fieldnames or []
        raw = list(rdr)
    missing = [c for c in REQUIRED_COLS if c not in hdr]
    if missing:
        raise ValueError(f"价格目录缺少必需列：{missing}")
    absent_opt = [c for c in OPTIONAL_COLS if c not in hdr]

    rows, by_key, by_parent = [], {}, {}
    for i, r in enumerate(raw, start=2):
        anchor = (r.get("Anchor Kind") or "").strip()
        method = (r.get("Method") or "").strip().upper()
        if method not in ("M3", "M4", "M5", "CAL"):
            method = ANCHOR2METHOD.get(anchor.lower(), "M5")
        if method == "CAL":
            method = "M3"
        low, high = _f(r.get("Low (CNY)")), _f(r.get("High (CNY)"))
        if low is None or high is None or low <= 0 or high <= 0:
            continue  # 区间不成对的行不参与定价
        key = (r.get("Category Key") or "").strip()
        row = {
            "row": i, "key": key,
            "category": (r.get("Category (archetype)") or "").strip(),
            "prefix": [x.strip().lower() for x in
                       (r.get("Match Head Prefix (;-sep, first hit wins)") or "").split(";")
                       if x.strip()],
            "sp_in": _split(r.get("Species Contains (optional)")),
            "sp_out": _split(r.get("Species Excludes")),
            "basis_unit": (r.get("Basis Unit") or "").strip(),
            "qty_basis": (r.get("Quantity Basis") or "").strip(),
            "low": low, "high": high,
            "typical": _f(r.get("Typical (CNY)")),
            "typical_src": (r.get("Typical Source") or "").strip(),
            "tooling": (r.get("Tooling In Price") or "?").strip() or "?",
            "price_year": (r.get("Price Year") or "").strip(),
            "scope": (r.get("Scope") or "").strip(),
            "method": method, "anchor": anchor,
            "use": (r.get("Use") or "").strip(),
            "rationale": (r.get("Rationale") or "").strip(),
            "source": (r.get("Source") or "").strip(),
            "src_date": (r.get("Source Date") or "").strip(),
            "notes": (r.get("Notes") or "").strip(),
            "req_attr": (r.get("Required Attribute") or "").strip(),
            "subtier_of": (r.get("Subtier Of") or "").strip(),
        }
        # ★ Use=Calibration 的行只做反查（反推模型是否合理），不参与补位定价
        if row["use"].lower() == "calibration":
            rows.append(row)
            by_key.setdefault(key, row)
            continue
        rows.append(row)
        by_key.setdefault(key, row)
        if row["subtier_of"]:
            by_parent.setdefault(row["subtier_of"], []).append(row)
    return {"rows": rows, "by_key": by_key, "by_parent": by_parent,
            "n_rows": len(rows), "missing_cols": absent_opt, "path": str(p)}


def _attr_match(req: str, attrs: dict) -> tuple[bool, str]:
    """Required Attribute 全部已知才命中。格式：`k=v;k2=a-b` / `k>=v` / `k<=v`。

    返回 (是否命中, 说明)。任一属性缺失 → False（回退父行，不猜）。
    """
    if not req:
        return True, ""
    parts = [x.strip() for x in req.split(";") if x.strip()]
    why = []
    for item in parts:
        if ">=" in item:
            k, v = item.split(">=", 1)
            got = attrs.get(k.strip().lower())
            if got is None:
                return False, f"{k.strip()} 未知"
            if float(got) < float(v):
                return False, f"{k.strip()}={got} < {v}"
            why.append(f"{k.strip()}={got}")
        elif "<=" in item:
            k, v = item.split("<=", 1)
            got = attrs.get(k.strip().lower())
            if got is None:
                return False, f"{k.strip()} 未知"
            if float(got) > float(v):
                return False, f"{k.strip()}={got} > {v}"
            why.append(f"{k.strip()}={got}")
        elif "=" in item:
            k, v = item.split("=", 1)
            k, v = k.strip().lower(), v.strip()
            got = attrs.get(k)
            if got is None:
                return False, f"{k} 未知"
            if "-" in v:
                lo, hi = v.split("-", 1)
                try:
                    if not (float(lo) <= float(got) <= float(hi)):
                        return False, f"{k}={got} 不在 {v}"
                    why.append(f"{k}={got}∈{v}")
                    continue
                except ValueError:
                    pass
            if str(got).strip().lower() != v.lower():
                return False, f"{k}={got} ≠ {v}"
            why.append(f"{k}={got}")
        else:
            return False, f"无法解析的条件：{item}"
    return True, ";".join(why)


def _convert(row: dict, ctx: dict) -> tuple[float | None, float | None, str]:
    """按 Basis Unit 把目录区间换成本件单价区间。返回 (low, high, 换算说明)。"""
    unit = row["basis_unit"].lower()
    if "piece" in unit or unit in ("cny/pcs", "cny/pc"):
        return row["low"], row["high"], "CNY/piece 直取"
    if "/kg" in unit:
        kg = ctx.get("net_weight") or 0.0
        if kg <= 0:
            return None, None, "单价按重量计，但本件净重未知"
        return row["low"] * kg, row["high"] * kg, f"CNY/kg × {kg:.3f} kg"
    if "/m2" in unit or "/m²" in unit:
        ln, wd = ctx.get("length") or 0.0, ctx.get("width") or 0.0
        if ln <= 0 or wd <= 0:
            return None, None, "单价按面积计，但本件长/宽未知"
        area = ln * wd / 1e6
        return row["low"] * area, row["high"] * area, f"CNY/m2 × {ln:.1f}×{wd:.1f}mm = {area:.5f} m2"
    return None, None, f"未知计价基数：{row['basis_unit']}"


def _species_ok(row: dict, species: str) -> bool:
    s = (species or "").lower()
    for bad in row["sp_out"]:
        if bad in s:
            return False
    if row["sp_in"]:
        return any(k in s for k in row["sp_in"])
    return True


def match_price(head: str, species: str, attrs: dict, ctx: dict,
                catalog: dict) -> dict | None:
    """按 Step 4 的取价顺序找目录行并换算。

    返回 dict（含 method/low/high/basis/tier/key/row）或 None（目录无命中 → 交回标未定价）。
    ★ 空 `Match Head Prefix` 的 Use=Fill 行不会被名称匹配命中：
      其 `prefix` 列表为空，下面的 `for pre in row["prefix"]` 一次都不执行。
    """
    head_l = (head or "").strip().lower()
    hits = []
    for row in catalog["rows"]:
        if row["use"].lower() == "calibration":
            continue
        if not _species_ok(row, species):
            continue
        # ---- 子档优先：只在 Required Attribute 全部已知时命中
        if row["subtier_of"]:
            ok, why = _attr_match(row["req_attr"], attrs)
            if not ok:
                continue
            parent = catalog["by_key"].get(row["subtier_of"])
            if parent is not None and parent["prefix"]:
                if not any(head_l.startswith(p) for p in parent["prefix"]):
                    continue
            low, high, cnv = _convert(row, ctx)
            if low is None:
                continue
            hits.append((1, row, low, high, cnv, why))
            continue
        # ---- 父档：名称前缀匹配
        if row["req_attr"]:
            # 父档自身带 Required Attribute（如 plastic_holder_GF30）→ 只经子档匹配
            continue
        for pre in row["prefix"]:
            if head_l.startswith(pre):
                low, high, cnv = _convert(row, ctx)
                if low is None:
                    continue
                hits.append((2, row, low, high, cnv, ""))
                break
    if not hits:
        return None
    hits.sort(key=lambda x: (x[0], x[1]["row"]))
    tier, row, low, high, cnv, why = hits[0]

    method = row["method"]
    # Typical 只有带 Typical Source 才可用，且必须是能直接当单价的口径
    basis = f"目录 {row['key']}"
    tag = ""
    ratio = (high / low) if low else 0.0
    if method in ("M4", "M5") and row["typical"] is not None and row["typical_src"]:
        base = row["typical"]
        basis += f"；Typical={base:g}（{row['typical_src']}）"
        # Typical 若是 CNY/m2 等非件口径，必须先换算（PCB 即属此类）
        unit = row["basis_unit"].lower()
        if "piece" not in unit and unit not in ("cny/pcs", "cny/pc"):
            base2, _, cnv2 = _convert({**row, "low": base, "high": base}, ctx)
            if base2 is None:
                base = None
            else:
                base = base2
                basis += f"；按基数换算 → {base:.4f}（{cnv2}）"
    else:
        if ratio > 1.5:
            base = (low * high) ** 0.5
            tag = "区间暂估"
        else:
            base = (low + high) / 2
    if ratio > WIDE_RATIO:
        tag = "[宽>10×]"
    if row["anchor"].lower() == "envelope" and not tag:
        tag = "[包络]"
    basis += f"；{cnv}；{row['source']}"
    if row["src_date"]:
        basis += f"（{row['src_date']}）"
    if why:
        basis += f"；属性：{why}"
    if row["scope"]:
        basis += f"；Scope={row['scope']}"
    return {
        "method": method, "low": low, "high": high,
        "base": base if base is not None else (low + high) / 2,
        "basis": basis, "tier": tag, "key": row["key"], "row": row,
        "tooling": row["tooling"], "subtier": tier == 1,
    }
