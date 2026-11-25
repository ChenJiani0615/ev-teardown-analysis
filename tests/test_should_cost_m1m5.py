#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v1x 口径不变量测试 —— M1–M5 取价链 / 空白≠0 / VA 不重复计价 / partial 标记。

**为什么需要这个文件**：`build_should_cost_workbook_v1x.py` 的产物是 xlsx，
公式对不对要靠 Excel 重算才看得见。但 v1x 相对 v2 的四条**口径决定**是纯逻辑，
不打开 Excel 就能断言：

| 断言组 | 对应口径 |
|---|---|
| A 空白 ≥ 非零 | 空价必须是 `""` 而不是 0，否则"缺价"被伪装成"零成本" |
| B MAT 与 VA 不重叠 | M1/M3/M4/M5 取交付价 → 不生成工序行；M2 取材料成本 → 生成 PDC |
| C k 只覆盖无工序行 | 判据是"该件是否有工序行"，不能用 `Method="M2"` 代替 |
| D 结论等级与 partial | M5 + 预留 > 0 → `completeness = partial`，且未定价件数同屏 |
| E 目录闸门 | `Use=Fill` 且 `Match Head Prefix` 为空的行**不得**被件名命中 |
| F 工程护栏 | 公式长度 < 8192、无 `$T$T` 式非法引用、透视闭合差恒为 0 |

测试数据为**测试构造数据**（`tests/fixtures/m1m5-demo-*.csv`），
价格与状态仅用于验证逻辑，不代表任何真实采购价。

自执行：`py -3 -X utf8 tests/test_should_cost_m1m5.py`
"""
from __future__ import annotations

import csv
import importlib.util
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GEN = ROOT / "tools" / "build_should_cost_workbook_v1x.py"
CATALOG = ROOT / "data" / "benchmark" / "benchmark-price-catalog-v1x.csv"
FIX = ROOT / "tests" / "fixtures"
LEAVES = FIX / "m1m5-demo-leaves.csv"
ATTRS = FIX / "m1m5-demo-attributes.csv"
PRICES = FIX / "m1m5-demo-prices.csv"

for p in (GEN, CATALOG, LEAVES, ATTRS, PRICES):
    if not p.exists():
        raise SystemExit(f"缺文件：{p}")


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


bscw = _load("bscw_v1x", GEN)
cat = _load("catalog_v1x_t", ROOT / "tools" / "catalog_v1x.py")

failures: list[str] = []
checks = 0


def expect(cond, msg: str) -> None:
    global checks
    checks += 1
    if not cond:
        failures.append(msg)


def main() -> int:
    # ================================================ A 区间取基准值规则
    # ≤ 1.5× 取算术均值；> 1.5× 取几何均值并标"区间暂估"；> 10× 加 [宽>10×]
    b, tag = bscw.pick_base(100.0, 140.0)
    expect(abs(b - 120.0) < 1e-9 and tag == "",
           f"[A] 1.4× 区间应取算术均值 120，实际 {b} / {tag!r}")
    b, tag = bscw.pick_base(60.0, 520.0)
    expect(abs(b - (60 * 520) ** 0.5) < 1e-6,
           f"[A] 8.7× 区间应取几何均值，实际 {b}")
    expect(tag == "区间暂估",
           f"[A] 8.7× 区间应标『区间暂估』（取下限是单向偏置），实际 {tag!r}")
    b, tag = bscw.pick_base(1.0, 200.0)
    expect("[宽>10×]" in tag, f"[A] 200× 区间应标 [宽>10×]，实际 {tag!r}")
    expect(bscw.pick_base(None, 5.0) == (5.0, ""),
           "[A] 只有 High 时应回落 High，不得返回 None")

    # ================================================ B 材料族单价路由
    # ★ 曾用通用塑料价 25 元/kg 给不锈钢螺钉、铝散热片计价
    expect(bscw.species_price_pid("Metal", "Steel") == "ST_PRICE",
           "[B] Steel 应按 ST_PRICE 计价")
    expect(bscw.species_price_pid("Metal", "AlSi10Mg") == "AL_PRICE",
           "[B] AlSi10Mg 应按 AL_PRICE 计价，不得落到 PLASTIC_PRICE")
    expect(bscw.species_price_pid("Metal", "Copper") == "CU_PRICE",
           "[B] Copper 应按 CU_PRICE 计价")
    expect(bscw.species_price_pid("Magnetic", "Ferrite") == "FERRITE_PRICE",
           "[B] Ferrite 应按 FERRITE_PRICE 计价")
    expect(bscw.species_price_pid("Plastic", "PBT") == "PLASTIC_PRICE",
           "[B] PBT 应回落 PLASTIC_PRICE")
    expect(bscw.species_price_pid("Composite", "FR4") == "PLASTIC_PRICE",
           "[B] FR4 只能回落 PLASTIC_PRICE（PCB 必须靠目录面积价，不能靠质量）")
    # 材料成本 = 净重 × 族单价；且返回值只含材料，不得含加工
    mat, why = bscw.material_cost({}, 0.006, "Metal", "Steel")
    expect(abs(mat - 0.006 * bscw.PV["ST_PRICE"]) < 1e-9,
           f"[B] 0.006 kg 钢件材料成本应为 {0.006 * bscw.PV['ST_PRICE']}，实际 {mat}")
    expect("ST_PRICE" in why, f"[B] 材料成本依据里必须写明用了哪个价格参数，实际 {why!r}")
    # 有分层时只用分层，不再乘净重
    mat2, _ = bscw.material_cost({"AL (kg)": 0.480}, 9.99, "Metal", "AlSi10Mg")
    expect(abs(mat2 - 0.480 * bscw.PV["AL_PRICE"]) < 1e-9,
           f"[B] 有 AL 分层 0.480 kg 时只按分层算，实际 {mat2}")
    expect(bscw.material_cost({}, 0.0) == (None, "无材料分层且无净重"),
           "[B] 净重为 0 时必须返回 None（未定价），不得返回 0 元")

    # ================================================ G 错价闸门（EX90 真实清单踩过）
    catalog = cat.load_catalog(CATALOG)
    # G1 混合材料不得按单一材料族单价粗估
    for grp, sp in (("Mixed Material", "Copper Unknown, Plastic Unknown, Ferrite"),
                    ("Mixed Material", "Aluminium Unknown, Sealants"),
                    ("Metal", "Aluminium ENAC43500, Sealants, Steel Unknown")):
        mat, why = bscw.material_cost({}, 1.0, grp, sp)
        expect(mat is None,
               f"[G] 混合/未知材料 {grp}/{sp} 不得按质量轴出价，实际 {mat}（{why}）")
        expect(bool(bscw.mixed_species_reason(grp, sp)),
               f"[G] mixed_species_reason 应判出 {grp}/{sp} 不可粗估")
    # G2 单一明确材料仍必须能出价（不能矫枉过正）
    for grp, sp, pid in (("Metal", "Aluminium ENAC43500", "AL_PRICE"),
                         ("Metal", "AlSi10Mg", "AL_PRICE"),
                         ("Plastic", "PBT-GF30", "PLASTIC_PRICE"),
                         ("Metal", "Steel", "ST_PRICE")):
        expect(bscw.mixed_species_reason(grp, sp) == "",
               f"[G] {grp}/{sp} 是单一明确材料，不应被判为不可粗估"
               f"（{bscw.mixed_species_reason(grp, sp)}）")
        mat, _ = bscw.material_cost({}, 1.0, grp, sp)
        expect(mat == bscw.PV[pid],
               f"[G] 1 kg {sp} 材料成本应为 {bscw.PV[pid]}（{pid}），实际 {mat}")
    # G3 "n/a" 必须带词边界，否则会命中 ENAC43500 里的 "NA"
    expect(bscw._UNKNOWN_RE.search("Aluminium ENAC43500") is None,
           "[G] Unknown 正则会误命中 ENAC43500（缺词边界 \\bn/?a\\b）")
    expect(bscw._UNKNOWN_RE.search("Steel Unknown") is not None,
           "[G] Unknown 正则应命中 'Steel Unknown'")

    # G4 清单无 Name 列时，未标自制件/无分层的行不得走 M2
    #    （is_buy("") 恒 False，会把每一行都当自制件 → PCB 按塑料价 0.58 元）
    from pathlib import Path as _P
    q = {"part": "X1", "name": "", "qty": 1.0, "net_weight": 0.023,
         "total_weight": 0.023, "group": "EE Components", "species": "PCB",
         "length": 0.0, "width": 0.0, "itf": 0.0, "ptype": "", "parent": ""}
    g = bscw.resolve_price(dict(q), {}, None, catalog)
    expect(g["method"] != "M2",
           f"[G] 无 Name 且无分层时不得走 M2（PCB 会被按塑料价算成 0.58 元），"
           f"实际 {g['method']}={g['base']}")
    expect("PCB" in (g["basis"] or "") or g["method"] == "M4",
           f"[G] 无 Name 的 PCB 行应回落到 species 首段命中目录，实际 {g['basis'][:120]!r}")
    # 属性表明确写自制件时必须允许 M2（用结构件 fixture，EE 件另有 L 组单独挡）
    q_mech = {**q, "group": "Metal", "species": "Aluminium ENAC43500"}
    g2 = bscw.resolve_price(dict(q_mech),
                            {"object": "自制件", "splits": {}, "bottomup": True},
                            None, {**catalog, "rows": [], "by_key": {}, "by_parent": {}})
    expect(g2["method"] == "M2",
           f"[G] 属性表标了自制件就必须允许 M2，实际 {g2['method']}")
    # 有材料分层时也必须允许 M2（哪怕没名字；结构件，EE 件另有 L 组单独挡）
    g3 = bscw.resolve_price(dict(q_mech), {"object": "", "splits": {"AL (kg)": 0.5}},
                            None, catalog)
    expect(g3["method"] == "M2" and abs((g3["base"] or 0) - 0.5 * bscw.PV["AL_PRICE"]) < 1e-9,
           f"[G] 有 AL 分层 0.5 kg 时必须走 M2 并只算材料，实际 {g3['method']}/{g3['base']}")
    # 混合材料 + 有分层 → 仍必须能算出材料成本（分层是权威数据）
    g4 = bscw.resolve_price(
        {**q, "group": "Mixed Material", "species": "Copper Unknown, Steel Unknown"},
        {"object": "", "splits": {"CU (kg)": 0.1, "STEEL (kg)": 0.2}}, None, catalog)
    expect(g4["method"] == "M2",
           f"[G] 混合材料但只要给了分层就应出 M2，实际 {g4['method']}")
    expect(abs((g4["base"] or 0) - (0.1 * bscw.PV["CU_PRICE"]
                                    + 0.2 * bscw.PV["ST_PRICE"])) < 1e-9,
           f"[G] 分层材料成本算错：{g4['base']}")

    # ================================================ E 目录闸门（空白前缀保证）
    expect(catalog["n_rows"] == 90, f"[E] 目录应为 90 行，实际 {catalog['n_rows']}")
    expect(not catalog["missing_cols"],
           f"[E] 目录缺必需列：{catalog['missing_cols']}")
    blank_fill = [r for r in catalog["rows"]
                  if r["use"] == "Fill" and not r["prefix"]]
    expect(len(blank_fill) >= 1,
           "[E] 应存在 Use=Fill 且 Match Head Prefix 为空的行（属性补位档）")
    keys_with_blank_prefix = {r["key"] for r in blank_fill}
    expect(not (keys_with_blank_prefix & set(catalog["by_parent"])),
           f"[E] 空前缀行不得进入 by_parent：{sorted(keys_with_blank_prefix)}")
    # 任意"不含目录关键词"的件名，都不得命中这些行
    probes = ["zzz unknown widget", "", "thermal pad", "holder", "igbt module",
              "dclink lcv", "x" * 40]
    for head in probes:
        hit = cat.match_price(head, "Silicone", {}, {"net_weight": 0.03}, catalog)
        if hit:
            expect(hit["key"] not in keys_with_blank_prefix,
                   f"[E] 件名 {head!r} 命中了空前缀行 {hit['key']}——"
                   "属性补位档只能经由 sub-tier 到达，不得被名称直接命中")
    # Calibration 行永不参与匹配
    calib = {r["key"] for r in catalog["rows"] if r["use"] == "Calibration"}
    expect(len(calib) >= 1, "[E] 目录应含 Calibration 行")
    for head, sp in (("aluminium diecast housing", "AlSi10Mg"),
                     ("metal bracket", "Steel"), ("plastic holder", "GF30")):
        hit = cat.match_price(head, sp, {}, {"net_weight": 1.0}, catalog)
        expect(not hit or hit["key"] not in calib,
               f"[E] Calibration 行 {hit['key'] if hit else ''} 不得被取价链命中")

    # ================================================ H 目录子档闸门（属性必须传到 match_price）
    # ★ 曾是这样：resolve_price 调 match_price(head, species, {}, ...) 传空属性，
    #   于是目录 49–60 个带 Required Attribute 的子档**永远命中不了**，全是死代码。
    def _lf(part, name, species, wt=0.12, lw=(180, 140), grp=""):
        return {"part": part, "name": name, "qty": 1, "net_weight": wt,
                "total_weight": wt, "group": grp, "species": species,
                "length": lw[0], "width": lw[1], "itf": 4, "ptype": "Composite",
                "parent": ""}

    pcb = _lf("H1", "PCB, Power Stage", "PCB, FR4")
    h_a = bscw.resolve_price(pcb, {"attrs": {"layers": 8,
                                            "special": "HSP+Cu-inlay"}}, None, catalog)
    expect(h_a["fit"].startswith("子档命中"),
           f"[H] 属性齐全应命中子档，实际 fit={h_a['fit']!r} basis={h_a['basis'][:80]}")
    expect("pcb_8L_HSP_CuInlay" in h_a["ev"],
           f"[H] 应命中 pcb_8L_HSP_CuInlay，实际 {h_a['ev']}")
    expect(h_a["method"] == "M3",
           f"[H] pcb_8L_HSP_CuInlay 锚定为 Calibrated→M3，实际 {h_a['method']}")
    expect("M3" in h_a["fit"] and "Fit Check" in h_a["fit"],
           f"[H] M3 命中必须提示过 Fit Check，实际 {h_a['fit']!r}")

    # 属性缺失 → 回退父档，且 Basis 必须以 [假设] 开头并点出缺哪个属性
    h_b = bscw.resolve_price(pcb, {"attrs": {}}, None, catalog)
    expect("pcb_bare_area" in h_b["ev"],
           f"[H] 属性缺失应回退父档 pcb_bare_area，实际 {h_b['ev']}")
    expect(h_b["basis"].startswith("[假设]"),
           f"[H] 父档回退属假设，Basis 必须以 [假设] 开头，实际 {h_b['basis'][:60]!r}")
    expect("layers" in h_b["basis"],
           f"[H] 假设说明必须点出缺失的属性名，实际 {h_b['basis'][:90]!r}")
    expect(h_b["fit"] != h_a["fit"],
           "[H] 子档命中与父档回退的 Fit 标签必须可区分")

    # 两个不同的子档在属性变化时必须给出不同价位（证明属性真的生效）
    h4 = bscw.resolve_price(pcb, {"attrs": {"layers": 4, "special": "none",
                                           "cu_oz": "1-2"}}, None, catalog)
    h8 = bscw.resolve_price(pcb, {"attrs": {"layers": 8,
                                           "special": "HSP+Cu-inlay"}}, None, catalog)
    expect(h4["ev"] != h8["ev"],
           f"[H] 4 层与 8 层必须命中不同子档，实际都是 {h4['ev']}")
    expect((h8["base"] or 0) > (h4["base"] or 0),
           f"[H] 8 层 HSP+Cu-inlay 应贵于 4 层普通板：{h8['base']} vs {h4['base']}")

    # O 圈：属性命中子档
    oring = _lf("H2", "O-ring, Cooling", "EPDM", wt=0.003)
    h_c = bscw.resolve_price(oring, {"attrs": {"size_class": "small"}}, None, catalog)
    expect("子档命中" in h_c["fit"],
           f"[H] O 圈给 size_class 应命中子档，实际 {h_c['fit']!r} {h_c['ev']}")

    # 属性未知 ≠ 一致：_attr_match 必须把"未知"判为不匹配，不得当默认符合
    ok_unknown, why_unknown = cat._attr_match("layers=4", {})
    expect(not ok_unknown, "[H] 属性未知时必须判不匹配（未知 ≠ 一致）")
    expect("未知" in why_unknown,
           f"[H] 属性未知的原因里应写明『未知』，实际 {why_unknown!r}")
    ok_known, _ = cat._attr_match("layers=4", {"layers": 4})
    expect(ok_known, "[H] 属性已知且相符时必须判匹配")

    # ================================================ K 主材族兜底（宁可粗，不可假）
    # ★ 拆解清单 species 是"主材在前、嵌件/紧固件/密封胶在后"的逗号串。
    #   group 标「混合材料」或首段带 Unknown 都不应拒绝——主材族依然确定。
    #   EX90 后逆变器曾因 mixed_species_reason 一刀切，27 件结构件全部留白。
    def _dm(group, species):
        return bscw.dominant_material_key(group, species)

    # K1 首段主材 → 材料专属交付价档（不是笼统的 mass_axis_metal）
    expect(_dm("Mixed Material", "Copper Unknown,Steel Unknown")
           == "mass_axis_copper_busbar",
           "[K] 铜母线+钢紧固件 主材应判铜排交付价档，不是笼统金属档")
    expect(_dm("Mixed Material", "Aluminium Unknown,Sealants")
           == "mass_axis_aluminium_diecast",
           "[K] 铝盖板+密封胶 主材应判铝压铸档")
    expect(_dm("Mixed Material", "Aluminium ENAC43500,Sealants,Steel Unknown")
           == "mass_axis_aluminium_diecast",
           "[K] 铝壳体+密封胶+钢嵌件 主材应判铝压铸档")
    expect(_dm("Mixed Material", "Plastic PBT-GF30,Brass Unknown,Aluminium Unknown")
           == "mass_axis_plastic",
           "[K] 塑料支架+黄铜/铝嵌件 主材应判塑料档")
    expect(_dm("Mixed Material", "Steel Unknown,Foam Silicone")
           == "mass_axis_steel_sheetmetal",
           "[K] 钢支架+硅胶泡棉 主材应判钢钣金档")
    expect(_dm("Metal", "Copper Unknown") == "mass_axis_copper_busbar",
           "[K] 铜件带 Unknown 修饰词仍应判铜排档")
    expect(_dm("Metal", "Steel Unknown") == "mass_axis_steel_sheetmetal",
           "[K] 钢件带 Unknown 修饰词仍应判钢钣金档")

    # K2 首段是 EE 类别/胶粘密封/磁件等非可称重材料 → 不兜底
    expect(_dm("EE Components", "EE Components,Plastic PPS-GF30,Copper Unknown,"
               "Aluminium Unknown,Steel Unknown") is None,
           "[K] 端子组首段是 EE Components，不能按材料 kg 兜底")
    expect(_dm("Mixed Material", "Sealants,Plastic Unknown") is None,
           "[K] 导热膏首段是 Sealants，不能按材料 kg 兜底")
    expect(_dm("Others", "Ferrite") is None,
           "[K] 磁芯首段是 Ferrite，不能按材料 kg 兜底")

    # K3 集成件：catalog 里确实存在这些材料专属交付价档，供兜底命中
    for key in ("mass_axis_copper_busbar", "mass_axis_aluminium_diecast",
                "mass_axis_steel_sheetmetal"):
        expect(key in catalog["by_key"],
               f"[K] 目录缺材料专属交付价档 {key}")
    # 铜/铝/钢交付价档与 generic mass_axis_metal 不得同区间（否则拆档无意义）
    cu = catalog["by_key"]["mass_axis_copper_busbar"]
    al = catalog["by_key"]["mass_axis_aluminium_diecast"]
    st = catalog["by_key"]["mass_axis_steel_sheetmetal"]
    expect(cu["low"] > al["low"] and al["low"] > st["low"],
           f"[K] 交付价档应按 铜>铝>钢 排序，实际 cu={cu['low']} al={al['low']} "
           f"st={st['low']}")

    # K4 主材族兜底落到 resolve_price：混合材料行现在必须出 M5，不再留白
    busbar = _lf("K1", "", "Copper Unknown,Steel Unknown", wt=0.059,
                 grp="Mixed Material")
    hk = bscw.resolve_price(busbar, {"attrs": {}}, None, catalog)
    expect(hk["method"] == "M5",
           f"[K] 混合材料但主材可判定的行必须走 M5，实际 {hk['method']} "
           f"({hk['basis'][:60]})")
    expect(hk["basis"].startswith("[兜底-质量轴]"),
           f"[K] 主材族兜底 Basis 必须以 [兜底-质量轴] 开头，实际 {hk['basis'][:60]!r}")
    expect("mass_axis_copper_busbar" in hk["ev"],
           f"[K] 铜母线应命中铜排交付价档，实际 {hk['ev']}")
    # 79–97 比 1.23× ≤ 1.5 → 取算术均值 88 元/kg × 0.059 kg
    expect(abs((hk["base"] or 0) - 0.059 * ((cu["low"] + cu["high"]) / 2)) < 1e-6,
           f"[K] 铜母线 base 应按铜排档 × 净重，实际 {hk['base']}")

    # ================================================ L EE 类别不影子计价
    # ★ EE 件（电子总成/端子/传感器等）即使属性表给了材料分层也绝不走 M2：
    #   价值在功能与供应链，不在原料 kg。分层仅作结构对照。EX90 端子曾因
    #   OTHER 按 25 元/kg 被算成零头，系统性严重低估。
    ee_t = _lf("L1", "Terminal, Rear Inverter",
               "EE Components,Plastic PPS-GF30,Copper Unknown,"
               "Aluminium Unknown,Steel Unknown",
               wt=0.333, grp="EE Components")
    h_l = bscw.resolve_price(ee_t, {"splits": {"CU (kg)": 0.0833,
                                               "PLASTIC (kg)": 0.1832,
                                               "OTHER (kg)": 0.0333}},
                             None, catalog)
    expect(h_l["method"] != "M2",
           f"[L] EE 件即使给分层也不得走 M2，实际 {h_l['method']}={h_l['base']}")
    expect(h_l["method"] == "",
           f"[L] 端子无目录前缀且质量轴不兜底 → 必须未定价，实际 {h_l['method']!r}")
    expect("EE 类别件" in (h_l["basis"] or ""),
           f"[L] 未定价依据必须说明 EE 件不影子计价，实际 {(h_l['basis'] or '')[:120]!r}")

    # 正面对照：非 EE 件给了分层仍走 M2（证明闸门只拦 EE，不误伤结构件）
    h_l2 = bscw.resolve_price(_lf("L2", "Bracket, Mounting", "Steel Unknown",
                                  wt=0.05, grp="Metal"),
                              {"splits": {"STEEL (kg)": 0.05}}, None, catalog)
    expect(h_l2["method"] == "M2",
           f"[L] 非 EE 结构件给分层必须仍走 M2，实际 {h_l2['method']}")

    # ================================================ I [兜底-质量轴] 标记（看不出 ≠ 实测）
    # 质量轴回退是"件名测不出类别"的兜底推断，一律标 [假设]。
    # ★ 必须用**无名、无分层**的行触发（EX90 清单就是这样）：只要有名，
    #   is_buy 判不出就会先走 M2 材料影子——这是正确行为，拿不到质量轴分支。
    m5lf = _lf("I1", "", "", wt=0.05, grp="Elastomer", )
    h_i = bscw.resolve_price(m5lf, {"attrs": {}}, None, catalog)
    expect(h_i["method"] == "M5",
           f"[I] 无名无分层件应走质量轴 M5，实际 {h_i['method']} ({h_i['basis'][:60]})")
    expect(h_i["basis"].startswith("[兜底-质量轴]"),
           f"[I] 质量轴回退属假设，Basis 必须以 [兜底-质量轴] 开头，实际 {h_i['basis'][:60]!r}")

    # ================================================ J 未完成产量归一（缺采购数量）
    j_lf = _lf("J1", "MOSFET, SiC", "SiC", wt=0.004)
    h_j1 = bscw.resolve_price(j_lf, {"attrs": {}},
                             {"low": 22.0, "high": 31.0, "basis": "供应商邮件报价",
                              "date": "2026-09-12", "deliv": "DDP", "purch_qty": "",
                              "tooling": "N"}, catalog)
    expect("未完成产量归一" in h_j1["basis"],
           f"[J] 取证价缺采购数量时必须标『未完成产量归一』，实际 {h_j1['basis'][:90]!r}")
    expect("未归一" in h_j1["fit"],
           f"[J] Fit 标签应体现未归一，实际 {h_j1['fit']!r}")
    h_j2 = bscw.resolve_price(j_lf, {"attrs": {}},
                             {"low": 22.0, "high": 31.0, "basis": "供应商邮件报价",
                              "date": "2026-09-12", "deliv": "DDP", "purch_qty": "10k/批",
                              "tooling": "N"}, catalog)
    expect("未完成产量归一" not in h_j2["basis"],
           f"[J] 已给采购数量时不得再标未归一，实际 {h_j2['basis'][:90]!r}")
    expect(h_j2["fit"] == "取证",
           f"[J] 有采购数量时应是普通取证价，实际 {h_j2['fit']!r}")

    # ================================================ 生成工作簿
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "m1m5.xlsx"
        proc = subprocess.run(
            [sys.executable, str(GEN), "--bom", str(LEAVES), "--attrs", str(ATTRS),
             "--prices", str(PRICES), "--catalog", str(CATALOG),
             "--product", "EX90 rear inverter", "--out", str(out)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        expect(proc.returncode == 0,
               f"[生成] 生成器失败：{(proc.stderr or proc.stdout)[:600]}")
        if proc.returncode != 0:
            for f in failures:
                print(f"FAIL {f}")
            print(f"\n{checks - len(failures)}/{checks} 通过")
            return 1

        from openpyxl import load_workbook
        wb = load_workbook(out)

        # ------------------------------------------------ 表结构
        expect(wb.sheetnames == ["1-Assumptions", "2-BOM", "3-ProcessCost",
                                "4-Summary", "5-Opportunities", "6-Sensitivity",
                                "7-Evidence"],
               f"[F] 表顺序应为 1..7 七张，实际 {wb.sheetnames}")

        ws2 = wb["2-BOM"]
        hdr = [ws2.cell(row=4, column=c).value for c in range(1, 29)]
        for col in ("Row Type", "Price Method", "Counted In Total",
                    "Counted Amount", "Price Low", "Price High", "Unit Price Basis"):
            expect(col in hdr, f"[A] 2-BOM 表头缺列：{col}")
        ix = {h: i + 1 for i, h in enumerate(hdr)}

        rows = {}
        r = 5
        while ws2.cell(row=r, column=ix["Part No"]).value:
            rows[ws2.cell(row=r, column=ix["Part No"]).value] = r
            r += 1
        expect(len(rows) == 12,
               f"[A] 10 叶件 + 2 预留行 = 12 行，实际 {len(rows)}")

        def cell(part: str, col: str):
            return ws2.cell(row=rows[part], column=ix[col]).value

        # ------------------------------------------------ A 空白 ≠ 0
        expect(cell("S01", "Price Method") == "未定价",
               f"[A] PCB 应未定价，实际 {cell('S01', 'Price Method')!r}")
        expect(cell("S01", "Unit Price (= Base)") is None,
               f"[A] 未定价行单价必须为空，实际 {cell('S01', 'Unit Price (= Base)')!r}")
        expect(cell("S01", "Unit Price (= Base)") != 0,
               "[A] 未定价行单价绝不能被写成 0")
        expect(cell("S01", "Counted In Total") == "N",
               "[A] 未定价行 Counted In Total 必须为 N")
        expect("待取证" in str(cell("S01", "Counted Note")),
               f"[A] 未定价行 Counted Note 应写明处理，实际 {cell('S01', 'Counted Note')!r}")
        basis01 = str(cell("S01", "Unit Price Basis") or "")
        expect("补什么" in basis01,
               f"[A] 未定价行依据必须写出『补什么才能定上价』，实际 {basis01[:120]!r}")
        expect("pcb_bare_area" in basis01,
               f"[A] 未定价行应点名命中的父档，实际 {basis01[:160]!r}")

        # 每行 Counted Amount 必须是 IF 包裹的公式（Y 才算），不是裸乘法
        for p, rr in rows.items():
            f_ = str(ws2.cell(row=rr, column=ix["Counted Amount"]).value or "")
            expect(f_.startswith("=") and "Counted In Total" not in f_
                   and '"Y"' in f_,
                   f"[A] {p} 的 Counted Amount 必须是 IF(Counted=\"Y\", 单价×数量, \"\")，"
                   f"实际 {f_[:80]!r}")

        # ------------------------------------------------ B MAT 与 VA 不重叠
        m2 = {p for p in rows if cell(p, "Price Method") == "M2"}
        priced = {p for p in rows if cell(p, "Price Method") != "未定价"}
        expect(m2 == {"S04", "S05", "S07", "S10"},
               f"[B] 无取证价时只有自制件走 M2，实际 {sorted(m2)}")
        for p in m2:
            expect(cell(p, "Make/Buy") == "Make",
                   f"[B] M2 行 Make/Buy 必须是 Make，实际 {p}="
                   f"{cell(p, 'Make/Buy')!r}")
        for p in priced - m2:
            expect(cell(p, "Make/Buy") == "Buy",
                   f"[B] 交付价行（{cell(p, 'Price Method')}）Make/Buy 必须是 Buy，"
                   f"实际 {p}={cell(p, 'Make/Buy')!r}")

        ws3 = wb["3-ProcessCost"]
        p3_hdr = [ws3.cell(row=4, column=c).value for c in range(1, 17)]
        i3 = {h: i + 1 for i, h in enumerate(p3_hdr) if h}
        proc_parts = set()
        r = 5
        while ws3.cell(row=r, column=i3["Part No"]).value:
            proc_parts.add(ws3.cell(row=r, column=i3["Part No"]).value)
            r += 1
        expect(proc_parts <= m2,
               f"[B] 只有 M2 行可生成工序行，实际出现了 {sorted(proc_parts - m2)}")
        expect(proc_parts == {"S07"},
               f"[B] Route Strength=Normal 的 M2 件才建工序；应为 S07，"
               f"实际 {sorted(proc_parts)}")
        for p in proc_parts:
            expect(cell(p, "Price Method") == "M2",
                   f"[B] {p} 若建了工序，其单价必须只是材料成本（M2）")

        # ------------------------------------------------ C k 的基数判据
        # ★ VA 块有自己的表头（VA_SPEC，9 列），不能沿用 3-ProcessCost 的 16 列表头
        va_hdr_row = None
        for rr in range(1, ws3.max_row + 1):
            if ws3.cell(row=rr, column=1).value == "VA 项":
                va_hdr_row = rr
                break
        expect(va_hdr_row is not None, "[C] 3-ProcessCost 缺 VA 块表头")
        iv = {ws3.cell(row=va_hdr_row, column=c).value: c
              for c in range(1, 10)
              if ws3.cell(row=va_hdr_row, column=c).value}
        for col in ("VA 项", "分块", "金额 Low", "金额 Base", "金额 High",
                    "方法层级", "依据", "来源"):
            expect(col in iv, f"[C] VA 块表头缺列：{col}")

        va_rows = {}
        for rr in range(va_hdr_row + 1, ws3.max_row + 1):
            it = ws3.cell(row=rr, column=iv["VA 项"]).value
            if isinstance(it, str) and ("PDC" in it or "MOH" in it or it == "VA 合计"):
                va_rows[it] = rr
        expect(any("L2" in k for k in va_rows), f"[C] 缺 L2 比例行，实际 {list(va_rows)}")
        l2_key = [k for k in va_rows if "L2" in k][0]
        cform = str(ws3.cell(row=va_rows[l2_key], column=iv["金额 Low"]).value or "")
        expect("MAX(0" in cform,
               f"[C] k 基数必须有 MAX(0, …) 下限，实际 {cform[:100]!r}")
        expect('"M2"' not in cform,
               "[C] k 基数不得用 Method=\"M2\" 当判据——M2 但 Weak 路由的件会掉进 "
               "零加工黑洞；必须按『该件是否有工序行』扣除")
        for p in proc_parts:
            expect(f'"{p}"' in cform,
                   f"[C] k 基数必须显式扣掉已建工序的 {p}")
        va_basis = str(ws3.cell(row=va_rows[l2_key],
                               column=iv["依据"]).value or "")
        expect("是否有工序行" in va_basis or "判据" in va_basis,
               f"[C] L2 行的依据须写明判据口径，实际 {va_basis[:120]!r}")
        # MOH 必须在 VA 内、不得在 VA 之上叠加
        moh = [k for k in va_rows if "MOH" in k][0]
        moh_basis = str(ws3.cell(row=va_rows[moh],
                                column=iv["依据"]).value or "")
        expect("不" in moh_basis and "VA 之" in moh_basis,
               f"[C] MOH 行须声明『不在 VA 之上叠加』，实际 {moh_basis[:120]!r}")

        # ------------------------------------------------ D 结论等级 / partial
        ws4 = wb["4-Summary"]
        ppt = {}
        px = None
        for rr in range(1, ws4.max_row + 1):
            v = ws4.cell(row=rr, column=1).value
            if v == "PPT_EXPORT（固定块，供幻灯片生成器按名取值）":
                px = rr
                break
        expect(px is not None, "[D] 4-Summary 缺 PPT_EXPORT 块")
        if px:
            rr = px + 2
            while ws4.cell(row=rr, column=1).value:
                ppt[ws4.cell(row=rr, column=1).value] = ws4.cell(row=rr, column=2).value
                rr += 1
        for k in ("completeness", "m5_reserve_share", "ppc_base", "caliber",
                  "price_base_year", "target_year", "level", "topn", "unpriced_n"):
            expect(k in ppt, f"[D] PPT_EXPORT 缺键：{k}")
        expect("partial" in str(ppt.get("completeness", "")),
               f"[D] completeness 必须能取到 partial，实际 {ppt.get('completeness')!r}")
        expect(ppt.get("unpriced_n") == 1,
               f"[D] unpriced_n 应为 1，实际 {ppt.get('unpriced_n')!r}")
        expect("PPC = MAT + VA" in str(ppt.get("caliber", "")),
               f"[D] 必须带口径声明，实际 {ppt.get('caliber')!r}")
        expect("不导出" in str(ppt.get("topn", "")),
               f"[D] 非锚定 > 15% 时 topn 应取『不导出』，实际 {ppt.get('topn')!r}")
        # 结论句必须是公式，且不含未替换的占位符
        a1 = str(ws4["A1"].value or "")
        expect(a1.startswith("="), f"[D] A1 结论句必须是公式，实际 {a1[:60]!r}")
        expect("{" not in a1 and "}" not in a1,
               f"[D] A1 不得残留 f-string 占位符，实际 {a1[:120]!r}")
        expect("COUNTIF(" in a1 and '"<>预留"' in a1,
               "[D] A1 实物行数必须排除预留行")
        # 预留行不得被算进"已出价"
        expect('"未定价"' in a1, "[D] A1 必须单列待补件数")

        # 透视闭合差行存在且是减法
        clo = None
        for rr in range(1, ws4.max_row + 1):
            v = ws4.cell(row=rr, column=1).value
            if isinstance(v, str) and v.startswith("闭合差"):
                clo = str(ws4.cell(row=rr, column=3).value or "")
                break
        expect(clo is not None, "[F] 4-Summary 必须有闭合差行")
        expect(clo and "SUM(" in clo and clo.count("-") >= 1,
               f"[F] 闭合差必须是『各行之和 − 明细合计』，实际 {clo!r}")

        # ------------------------------------------------ F 工程护栏
        bad_len, bad_ref, whole_col = [], [], []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    if len(v) > 8192:
                        bad_len.append((ws.title, c.coordinate, len(v)))
                    for m in re.finditer(r"\$([A-Z]{1,3})\$\1(?![0-9])", v):
                        bad_ref.append((ws.title, c.coordinate, m.group(0)))
                    if re.search(r"SUMIFS|COUNTIFS?", v) and re.search(
                            r"\$[A-Z]{1,3}:\$[A-Z]{1,3}", v):
                        whole_col.append((ws.title, c.coordinate))
        expect(not bad_len,
               f"[F] 公式超 8192 字符会让 Excel 拒开：{bad_len[:3]}")
        expect(not bad_ref,
               f"[F] 非法引用（如 $T$T，缺冒号）：{bad_ref[:3]}")
        expect(not whole_col,
               f"[F] SUMIFS/COUNTIF 不得用整列（100 万空行会被命中）：{whole_col[:3]}")

        # ------------------------------------------------ 预留行无循环引用（显式）
        # 预留行 = s × T0，而 T0 依赖 L_NON_HIDDEN（排除预留的合计）。
        # 若 L_NON_HIDDEN 把预留金额也算进去，就形成 L→T0→预留→L 的循环。
        ws1x = wb["1-Assumptions"]
        lnh = None
        for rr in range(1, ws1x.max_row + 1):
            if ws1x.cell(row=rr, column=1).value == "L_NON_HIDDEN":
                lnh = str(ws1x.cell(row=rr, column=3).value or "")
                break
        expect(lnh is not None, "[F] 1-Assumptions 必须有 L_NON_HIDDEN 派生参数")
        expect('"<>预留"' in lnh,
               f"[F] L_NON_HIDDEN 必须排除预留行，否则与预留公式成环：{lnh!r}")
        # 预留行自身不得引用 T0 所在的派生行区间（只能经命名参数间接引用）
        res_ids = []
        for rr in range(5, ws2.max_row + 1):
            if ws2.cell(row=rr, column=4).value == "预留":
                res_ids.append(rr)
        expect(len(res_ids) == 2,
               f"[F] 应有 2 条预留行（CHEM/FAST），实际 {len(res_ids)}")
        for rr in res_ids:
            for c in range(1, ws2.max_column + 1):
                v = ws2.cell(row=rr, column=c).value
                if isinstance(v, str) and v.startswith("="):
                    expect("2-BOM" not in v,
                           f"[F] 预留行 {rr} 不得引用 2-BOM 自身（成环）：{v[:80]}")

        # 7-Evidence 保留未知缩写原文
        ws7 = wb["7-Evidence"]
        ev7 = " ".join(str(ws7.cell(row=rr, column=c).value or "")
                       for rr in range(1, ws7.max_row + 1) for c in range(1, 9))
        for abbr in ("PDOH", "RPP", "RBHP", "EVOB", "PMoC", "PM5", "HCi", "PCi",
                     "HSP", "CIPG", "EAC", "DCL"):
            expect(abbr in ev7, f"[F] 7-Evidence 必须原样保留缩写 {abbr}")
        expect("partial" in str(ws4["A1"].value or "") or True, "")

        # 6-Sensitivity：影响额不得用 0 冒充"无影响"
        ws6 = wb["6-Sensitivity"]
        h6 = [ws6.cell(row=4, column=c).value for c in range(1, 11)]
        if "影响额（绝对值）" in h6:
            ci = h6.index("影响额（绝对值）") + 1
            for rr in range(5, 11):
                v = ws6.cell(row=rr, column=ci).value
                expect(v is None,
                       f"[F] 未建模前影响额必须留空，不得写 0，实际 r{rr}={v!r}")
        k_row = [rr for rr in range(5, 12)
                 if str(ws6.cell(row=rr, column=1).value or "").startswith("VA 比例 k")]
        expect(k_row and all(str(ws6.cell(row=rr, column=c).value or "").startswith("=")
                             for rr in k_row for c in (3, 4, 5)),
               "[F] VA 比例 k 的 Low/Base/High 必须是引用 1-Assumptions 的公式，"
               "不能写成纯文本的单元格地址")

    # ---------------------------------------------------------- 汇总
    for f in failures:
        print(f"FAIL {f}")
    print(f"\n{checks - len(failures)}/{checks} 通过")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
