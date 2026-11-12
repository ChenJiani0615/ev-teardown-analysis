#!/usr/bin/env python3
"""零件清单自检与质量占比核算（只读，不修改输入）。

用法示例:
    python part_table_check.py parts.csv
    python part_table_check.py parts.xlsx
    python part_table_check.py parts.csv --format json --out report.json

设计约束（与 SKILL.md 的硬性规则一致）:
    * 缺失字段标记为 "-"，绝不补 0；
    * 汇总前检查 part + fastener 与 total 的差异，不静默采用；
    * 质量占比只在"本次已列件号集合"内计算，不代表整机；
    * 只做算术与一致性检查，不推断材料配比或成本。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# key -> 可接受的表头别名（已归一化：小写、去标点与空格）
FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "cs_part": ("cspart", "part", "件号", "cs件号", "零件号"),
    "name": ("name", "partname", "零件名称", "名称"),
    "qty": ("quantity", "qty", "数量"),
    "part_weight": ("partweight", "本体重量", "零件重量"),
    "fastener_weight": ("fastenerweight", "紧固件重量"),
    "total_weight": ("totalweightkg", "totalweight", "总重量", "总重"),
    "material_group": ("materialgroup", "材料组"),
    "material_species": ("materialspecies", "材料种类"),
    "material_tested": ("materialspeciestested", "实测材料"),
    "manufacturer": ("manufacturer", "制造商", "供应商"),
    "marking": ("markingnumber", "marking", "标识号"),
    "length": ("length", "长"),
    "width": ("width", "宽"),
    "height": ("height", "高"),
    "thickness": ("thicknesst", "thickness", "壁厚", "厚度"),
    "part_type": ("parttype", "零件类型"),
    "interfaces": ("numberofinterface", "interfaces", "接口数"),
}

REQUIRED = ("cs_part", "total_weight", "material_group", "material_species")

POWER_KEYWORDS = (
    "module", "tim", "thermal", "heatsink", "coldplate", "cold plate",
    "igbt", "sic", "功率", "模块", "冷板", "散热", "导热",
)

# 功能槽：(槽位名, 本体关键词, 限定语关键词 or None)
# 用于"九类功能槽无沉默项"检查：某个槽位在整张清单里一个关键词都没命中，
# 就必须在报告里单列并给区分方法。
#
# **两个必须避开的坑**：
# ① 不要用 "converter"/"inverter" 这类出现在**每个件名后缀**里的词，
#    否则该槽被整张清单命满，检查失去意义。
# ② 必须只匹配件名**逗号前的本体**。CS 件名是 `本体, 所属位置` 倒装格式，
#    用整串匹配会让 "Gasket, Rear Inverter Power Module" 命中"电能变换"槽、
#    "Spacer, Rear Inverter Cooling Plate" 命中"热传导"槽——把装在某总成上的
#    紧固件/密封件当成该总成本身。实测未修正前，rear inverter 的"电能变换"
#    槽被 1 个密封垫 + 2 个支架占满，"紧固与定位"槽反而漏报。
FUNCTION_SLOTS: tuple[tuple[str, tuple[str, ...], tuple[str, ...] | None], ...] = (
    # "电能变换"只认真正的变换级**本体**（模块 / 功率级）。
    # 分立开关管只进"开关动作"槽——因此只有分立器件的产品会在此槽显示"无对应件"，
    # 这不是漏采，而是要显式写一行"变换功能由 N 颗分立器件承担、不存在模块本体"。
    ("电能变换", ("power module", "converter module", "power stage", "功率模块", "功率级",
                  "变换器"), None),
    ("开关动作", ("mosfet", "igbt", "sic", "gan", "transistor", "diode", "开关管", "裸片",
                  "晶闸管"), None),
    # **不能写裸 "cap"**：它是 capacitor 的前缀，也是机械堵盖 cap 的完整词。
    # 实测（Rear E-Drive 盲测）：`End Cap, Shaft Final Drive Rear`、
    # `Cap, Rear Final Drive Clutch`、`Cap, Rear Final Drive Clutch 2` 三个
    # 机械堵盖（钢 + EPDM，合计 0.024 kg）被算成电容，功能槽被假阳性占满。
    # DCDC 的 `Cap`（0.001 kg，纯 HDPE）同属误命中。故只认完整词 capacitor。
    ("储能与滤波", ("capacitor", "inductor", "choke", "ferrite", "core", "电容", "电感",
                    "扼流", "磁"), None),
    ("电流承载与分配", ("busbar", "bus bar", "terminal", "connector", "jumper", "harness",
                        "wire", "cable", "母排", "端子", "线束", "排"), None),
    ("热传导与散热", ("tim", "thermal", "heatsink", "coldplate", "cold plate",
                      "cooling plate", "cooler", "cooling", "radiator",
                      "散热", "冷板", "导热"), None),
    ("电气绝缘与隔离", ("insulat", "isolation", "isolator", "insulation",
                        "绝缘", "隔离"), None),
    ("检测与传感", ("sensor", "shunt", "ntc", "thermistor", "传感", "互感", "检测"), None),
    ("密封与防护", ("seal", "gasket", "oring", "o-ring", "breather", "adhesive", "plug",
                    "密封", "胶", "圈", "呼吸"), None),
    ("紧固与定位", ("screw", "bolt", "nut", "clip", "spring", "locator", "spacer",
                    "holder", "clamp", "retainer", "bracket",
                    "螺钉", "螺栓", "卡扣", "弹片", "弹簧", "定位", "垫块", "保持", "支架"),
     None),
)
# 有一个例外必须靠限定语才能认出来：`Plate, Rear Inverter Cooling` 本体只是 "plate"，
# 必须限定语里出现 cooling 才算冷板。这类"本体泛化 + 限定语定性"的槽单独列出。
SLOT_HEAD_TAIL: tuple[tuple[str, tuple[str, ...], tuple[str, ...]], ...] = (
    ("热传导与散热", ("plate", "panel", "块", "板"), ("cool", "coolant", "冷却", "散热")),
)


def norm(value: object) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", str(value).lower())


ALIAS_TO_KEY = {alias: key for key, aliases in FIELD_ALIASES.items() for alias in aliases}


def to_num(value: object) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text in {"-", "—", "N/A", "n/a"}:
        return None
    text = re.sub(r"(?i)\s*(kg|mm|%)$", "", text).replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def read_rows(path: Path) -> list[dict[str, object]]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        try:
            from openpyxl import load_workbook  # type: ignore
        except ImportError:
            raise SystemExit("读取 xlsx 需要 openpyxl：pip install openpyxl")
        book = load_workbook(path, read_only=True, data_only=True)
        sheet = book[book.sheetnames[0]]
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return []
        header = [norm(c) if c is not None else "" for c in rows[0]]
        raw = [dict(zip(header, r)) for r in rows[1:]]
        book.close()
    else:
        import csv

        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            raw = [{(norm(k) if k else ""): v for k, v in row.items()} for row in reader]

    parsed: list[dict[str, object]] = []
    for row in raw:
        item: dict[str, object] = {}
        for header, value in row.items():
            key = ALIAS_TO_KEY.get(header)
            if key:
                item[key] = value
        if any(str(v).strip() for v in item.values() if v is not None):
            parsed.append(item)
    return parsed


def analyze(rows: list[dict[str, object]]) -> dict[str, object]:
    parts: list[dict[str, object]] = []
    issues: list[dict[str, str]] = []

    for index, row in enumerate(rows, start=1):
        part_id = str(row.get("cs_part") or f"row{index}").strip()
        qty = to_num(row.get("qty")) or 1.0
        part_w = to_num(row.get("part_weight"))
        fast_w = to_num(row.get("fastener_weight"))
        total_w = to_num(row.get("total_weight"))

        missing = [k for k in REQUIRED if not str(row.get(k) or "").strip()]

        weight_source = "total"
        if total_w is None:
            if part_w is None:
                weight_source = "unknown"
            else:
                total_w = (part_w or 0.0) + (fast_w or 0.0)
                weight_source = "derived"
                issues.append(
                    {
                        "part": part_id,
                        "type": "缺 total_weight",
                        "detail": "按 part+fastener 推算，需回查原始记录",
                    }
                )
        elif part_w is not None and abs((part_w + (fast_w or 0.0)) - total_w) > 0.0015:
            issues.append(
                {
                    "part": part_id,
                    "type": "重量不自洽",
                    "detail": f"part+fastener={part_w + (fast_w or 0.0):.4f} vs total={total_w:.4f}",
                }
            )

        length = to_num(row.get("length"))
        width = to_num(row.get("width"))
        height = to_num(row.get("height"))
        thickness = to_num(row.get("thickness"))
        dims = [d for d in (length, width, height) if d]
        if len(dims) == 3 and length and min(width or 0, height or 0) > 0:
            if length < 0.25 * min(width, height):  # type: ignore[arg-type]
                issues.append(
                    {
                        "part": part_id,
                        "type": "尺寸字段疑似错位",
                        "detail": f"length={length} 远小于 width/height={width}/{height}",
                    }
                )
        if thickness and dims and thickness >= min(dims):
            issues.append(
                {
                    "part": part_id,
                    "type": "厚度字段可疑",
                    "detail": f"thickness={thickness} >= 最小外形尺寸 {min(dims)}",
                }
            )

        parts.append(
            {
                "cs_part": part_id,
                "name": str(row.get("name") or "").strip() or None,
                "qty": qty,
                "total_weight": total_w,
                "weight_source": weight_source,
                "material_group": str(row.get("material_group") or "-").strip() or "-",
                "material_species": str(row.get("material_species") or "-").strip() or "-",
                "material_tested": str(row.get("material_tested") or "").strip() or None,
                "manufacturer": str(row.get("manufacturer") or "").strip() or None,
                "part_type": str(row.get("part_type") or "").strip() or None,
                "interfaces": to_num(row.get("interfaces")),
                "missing": missing,
            }
        )

    seen: dict[str, int] = {}
    for part in parts:
        seen[str(part["cs_part"])] = seen.get(str(part["cs_part"]), 0) + 1
    for part_id, count in seen.items():
        if count > 1:
            issues.append(
                {"part": part_id, "type": "重复件号", "detail": f"出现 {count} 次，合计前需去重"}
            )

    known = [p for p in parts if p["total_weight"] is not None]
    total_mass = sum(float(p["total_weight"]) for p in known)  # type: ignore[arg-type]
    for part in known:
        part["share"] = float(part["total_weight"]) / total_mass if total_mass else None

    groups: dict[str, float] = {}
    for part in known:
        group = str(part["material_group"])
        groups[group] = groups.get(group, 0.0) + float(part["total_weight"])  # type: ignore[arg-type]

    fastener_mass = sum(
        float(to_num(r.get("fastener_weight")) or 0.0) for r in rows
    )

    blob = " ".join(str(v) for r in rows for v in r.values()).lower()
    power_gap = not any(k in blob for k in POWER_KEYWORDS)

    # 功率级存在性判定（与 power_gap 不同）：
    # 有 TIM/散热界面件说明存在需要压紧并导热的发热面 -> 功率级一定存在，
    # 只是形态（模块/分立/板载）没有证据。二者必须分开，不能混成一句"缺失"。
    tim_hits = [p for p in parts if any(
        k in " ".join(str(p.get(f) or "") for f in ("cs_part", "name", "material_species", "material_group")).lower()
        for k in ("thermal", "paste", "tim", "导热", "gasket pad")
    )]
    tim_mass = round(sum(float(p["total_weight"] or 0) for p in tim_hits), 4)

    slot_coverage = []
    for slot, keys, _ in FUNCTION_SLOTS:
        hit = []
        for p in parts:
            if any(k in intrinsic(p) for k in keys):
                hit.append(p["cs_part"])
        slot_coverage.append({"slot": slot, "parts": hit, "covered": bool(hit)})
    for slot, hw, tw in SLOT_HEAD_TAIL:
        entry = next(s for s in slot_coverage if s["slot"] == slot)
        for p in parts:
            head, tail = split_name(p)
            if any(k in head for k in hw) and any(k in tail for k in tw):
                if p["cs_part"] not in entry["parts"]:  # type: ignore[index]
                    entry["parts"].append(p["cs_part"])  # type: ignore[index]
        entry["covered"] = bool(entry["parts"])
    # 判断"电能变换"槽时，若分立开关管已在"开关动作"槽，
    # 则该槽空缺不等于功能缺失，必须按"由分立器件承担"写法单列。
    pieces_by_id = {p["cs_part"]: float(p.get("qty") or 1) for p in parts}
    for entry in slot_coverage:
        entry["piece_count"] = round(
            sum(pieces_by_id.get(pid, 1.0) for pid in entry["parts"]), 2)

    return {
        "part_count": len(parts),
        "total_mass_kg": round(total_mass, 4),
        "fastener_mass_kg": round(fastener_mass, 4),
        "parts": parts,
        "groups": {k: round(v, 4) for k, v in sorted(groups.items(), key=lambda x: -x[1])},
        "issues": issues,
        "power_gap_hint": power_gap,
        "tim_evidence": {"parts": [p["cs_part"] for p in tim_hits], "mass_kg": tim_mass},
        "function_slots": slot_coverage,
        "integration_level": assess_architecture(parts),
    }


def split_name(part: dict[str, object]) -> tuple[str, str]:
    """拆分 CS 件名：`本体, 所属位置` → (head, tail)。

    **这是本脚本最容易出错的地方**。CS 件名是倒装格式，
    逗号前才是零件本体，逗号后只是它装在哪：
        "MOSFET, Rear Inverter Power Module"  → head="mosfet"   （本体是 MOSFET）
        "Spacer, Rear Inverter Cooling Plate" → head="spacer"   （本体是垫块，不是冷却板）
        "Holder, Rear Inverter Cooling Plate" → head="holder"   （本体是保持件，不是冷却板）
        "Gasket, Rear Inverter Power Module"  → head="gasket"   （本体是密封垫，不是模块）

    若用整串件名做关键词匹配，12 颗**分立** MOSFET 会因尾部出现 "Power Module"
    被判成模塑模块（P2），这是**方向性错误**：把分立误判为集成。
    因此：**形态判定只用 head，tail 只作限定语**。
    """
    name = str(part.get("name") or "").strip()
    if "," in name:
        head, tail = name.split(",", 1)
        return head.strip().lower(), tail.strip().lower()
    return name.lower(), ""


def intrinsic(part: dict[str, object]) -> str:
    """零件**自身**的形态描述：本体名 + 材料种类 + 材料组。

    为什么要把材料并进来：有些清单（如 samples/ 里的字段陷阱样例）**没有 name 列**，
    只剩件号；此时唯一能说明"这个零件是什么"的就是材料种类。
    `Ferrite`、`Copper`、`Elastomer EPDM` 描述的是**零件自己的材质**，
    而不是它装在谁身上，所以把它们并入是没有方向性风险的。

    绝不能并入的：`name` 的逗号后段（装配位置）、所在 Assembly 名、
    整机名——这些会制造"因为装得靠近所以是同类"的系统性误判。
    """
    head, _ = split_name(part)
    species = str(part.get("material_species") or "").lower()
    group = str(part.get("material_group") or "").lower()
    return f"{head} {species} {group}".strip()


def assess_architecture(parts: list[dict[str, object]]) -> dict[str, object]:
    """按四个**互相独立**的维度粗判架构，只做提示，不做结论。

    维度（不得跨轴比较优劣，不得相加，不得排序）：
      P 封装与集成程度  P0 分立 / P1 裸片 / P2 模塑模块 / P3 多芯片共基板 / P4 板载功率级
      T 冷却与热路      T0 无主动 / T1 单面液冷 / T2 双面液冷 / T3 直接冷却 / T4 基板内流道
      E 电路与开关      四栏：拓扑 / 器件 / 互连 / 驱动检测
      M 安装与压紧      M0 螺钉直压 / M1 弹片压板 / M2 焊接 / M3 压接 / M4 粘接

    刻意保守：拿不到器件级证据时返回 level=None 与 reasons，
    绝不由"有 TIM"直接推到"有模块"——TIM 落在 T 轴，不能反推 P 轴。
    """
    pairs = [(p, *split_name(p)) for p in parts]

    def head_is(p: dict[str, object], words: tuple[str, ...]) -> bool:
        """匹配零件**自身**形态（本体名 or 材料），不看装配位置。"""
        return any(w in intrinsic(p) for w in words)

    def tail_is(p: dict[str, object], words: tuple[str, ...]) -> bool:
        return any(w in split_name(p)[1] for w in words)

    def sel_head(words: tuple[str, ...]) -> list[dict[str, object]]:
        return [p for p in parts if head_is(p, words)]

    def sel_head_tail(hw: tuple[str, ...], tw: tuple[str, ...]) -> list[dict[str, object]]:
        return [p for p in parts if head_is(p, hw) and tail_is(p, tw)]

    def ids(ps: list[dict[str, object]], n: int = 4) -> str:
        """列出件号，并给出每个件号的数量（qty），避免把 1 个件号误读成 1 颗零件。"""
        shown = []
        for p in ps[:n]:
            q = p.get("qty")
            shown.append(f"{p['cs_part']}×{q:g}" if q and float(q) != 1 else str(p["cs_part"]))
        if len(ps) > n:
            shown.append(f"等 {len(ps)} 个件号")
        return ", ".join(shown)

    def pieces(ps: list[dict[str, object]]) -> float:
        """件号数与实物件数不是一回事：报数量必须用 qty 求和。"""
        return round(sum(float(p.get("qty") or 1) for p in ps), 2)

    def mass(ps: list[dict[str, object]]) -> float:
        return round(sum(float(p["total_weight"] or 0) for p in ps), 4)

    asm = [p for p in parts if str(p.get("part_type") or "").lower() == "assembly"]
    asm_mass = mass(asm)
    total = sum(float(p["total_weight"] or 0) for p in parts)
    ceiling = (f"{len(asm)} 个 Assembly 叶子未分解，合计 {asm_mass:.3f} kg"
               + (f"（{asm_mass / total * 100:.1f}%）" if total else ""))

    names_note = ("件名为 `本体, 所属位置` 倒装格式，形态判定只用逗号前本体；"
                  "逗号后位置名不得用于判断零件本身的形态。")

    # ---------- P 封装与集成程度 ----------
    mod = sel_head(("power module", "converter module", "inverter module", "功率模块"))
    # 尾部出现 "Power Module" = 有一个名为"功率模块"的**装配位置**，不是模块本体
    mod_slot_dev = [p for p in parts if tail_is(p, ("power module", "功率模块"))
                    and head_is(p, ("mosfet", "igbt", "sic", "gan", "transistor", "diode"))]
    substrate = sel_head(("substrate", "dbc", "amb", "insulated metal", "基板", "覆铜"))
    dev = sel_head(("mosfet", "igbt", "sic", "gan", "transistor", "开关管", "二极管"))
    p_reasons: list[str] = []
    p_level: str | None = None
    if mod:
        p_level = ("P2 或 P3（有模块本体件号；需区分有无模塑外壳、是否多芯片共基板）")
        p_reasons.append(f"模块本体件 {len(mod)} 个：{ids(mod)}")
        if substrate:
            p_reasons.append(f"另有基板类件 {len(substrate)} 个：{ids(substrate)}")
        p_reasons.append(f"**P 轴判定到此结束**：模块是否双面冷却属 T 轴，不在本轴判")
    elif dev:
        p_level = "P0（分立器件，多个独立器件件号）"
        p_reasons.append(
            f"器件本体件号 {len(dev)} 个，**实物 {pieces(dev):g} 颗**：{ids(dev)}"
            + (f"（合计 {mass(dev):.4f} kg）" if mass(dev) else ""))
        p_reasons.append("无模块本体件号、无基板件号 → 器件为独立封装，未集成到模块内")
        if mod_slot_dev:
            p_reasons.append(
                f"**关键修正**：清单里存在名为\"功率模块\"的装配位置，但装在该位置上的本件是 "
                f"**独立器件件号**（{ids(mod_slot_dev)}，实物 {pieces(mod_slot_dev):g} 颗）"
                "——该位置是**装配位置名，不是模塑模块**。按 P 轴应判 P0，"
                "不得因位置名叫\"功率模块\"就判 P2。")
        p_reasons.append(f"P 轴判定到此结束：{names_note}")
    else:
        p_reasons.append("无任何器件本体件号 → **不能判定 P 轴**，只能给上限")
    if not substrate:
        p_reasons.append("清单无基板/覆铜类件号 → 无 P2/P3 所需的绝缘基板证据（也可能在未展开 Assembly 内）")

    # ---------- T 冷却与热路 ----------
    # 冷板识别有三条路径，**缺一条就会把液冷误判成 T0**：
    #   ① 本体就是冷板词（Plate, ... Cooling / Coldplate）
    #   ② 本体是泛化件（plate/panel）
    #   ③ **本体是结构件、冷却通道写在限定语里**
    # 路径③ 是被 Rear E-Drive 盲测发现的缺口：`Cover, Coolant Channel Rear E-Drive`
    # （0.214 kg、10 接口）本体是 Cover，两条旧规则都不命中，
    # 于是脚本输出"无冷却板 → 倾向 T0（无主动冷却）"，而实物是液冷。
    # 这类"把有说成没有"的假阴性比假阳性危险得多，必须兜住。
    COOL_Q = ("cool", "coolant", "冷却", "液冷", "水冷")
    # 冷却**通道本体**：本体是结构件、冷却写在限定语里（Cover, Coolant Channel → 是通道本体）
    CHANNEL_BODY = ("cover", "housing", "casing", "case", "channel", "manifold",
                    "盖", "壳", "道", "腔")
    # 冷却**管路接口**：本体就是管/接头（Pipe, ... Coolant Inlet → 是接口，不是板）
    # 两者必须分开：若把 Pipe 也算进 plate，"有板 + 有进出口"就退化成"只有板"，
    # T 轴会被降级成"可能风冷"——这是同一类假阴性换个方向再犯一次。
    DUCT_BODY = ("pipe", "tube", "line", "hose", "adapter", "fitting", "nozzle",
                 "inlet", "outlet", "管", "接头", "口")
    _plate_words = ("coldplate", "cold plate", "cooling plate", "cooler",
                    "heatsink", "heat sink", "radiator", "冷板", "散热器")

    plate = [p for p in parts
             if head_is(p, _plate_words)
             or (head_is(p, ("plate", "panel")) and tail_is(p, COOL_Q))
             or (head_is(p, CHANNEL_BODY) and tail_is(p, COOL_Q))]
    plate_ids = {p["cs_part"] for p in plate}
    # 冷却液进出口只认"本体是管/接头"的件，不能把装在冷却板上的垫块算作冷却液件。
    # 本体词表必须包含 pipe/tube/line：只列 adapter/fitting/nozzle
    # 会让 `Pipe, ... Coolant Inlet` 漏掉（盲测实测缺口）。
    port = [p for p in parts
            if p["cs_part"] not in plate_ids
            and head_is(p, DUCT_BODY)
            and tail_is(p, ("coolant", "cooling", "cool", "inlet", "outlet", "water",
                            "冷却", "液"))]
    tim = sel_head(("thermal paste", "thermal pad", "gap filler", "phase change", "tim",
                    "thermal grease", "导热", "界面材料"))
    # 装在冷却板/功率模块上的压紧与定位件：是 T2"需要解耦夹持"的结构旁证
    stack = [p for p in parts
             if head_is(p, ("spacer", "holder", "locator", "shim", "clamp", "clip",
                            "垫块", "定位", "压板"))
             and tail_is(p, ("cool", "power module", "plate", "冷却", "功率模块"))]

    t_reasons: list[str] = []
    t_level: str | None = None
    plate_mass = mass(plate)
    tim_mass = mass(tim)
    if plate and port:
        t_level = ("T1 或 T2（有冷却板本体件号 + 冷却液进出口件号；"
                   "单面/双面必须另证）")
        t_reasons.append(
            f"冷却板本体 {len(plate)} 个件号（{ids(plate)}，合计 {plate_mass:.3f} kg）；"
            f"冷却液接口 {len(port)} 个件号（{ids(port)}）")
        if len({split_name(p)[0] for p in plate}) >= 2:
            t_reasons.append(
                "冷却板有 ≥2 种不同件号（厚度不同）→ **T2 双面冷却有可能**，"
                "但**不能**仅凭\"有两块板\"判 T2：必须排除其中一块是**不导通冷却液的压板/盖板**。"
                "判定方法：看两块板是否都接冷却液进出口，或看器件两侧是否各有 TIM。")
    elif plate:
        t_level = "T1（有冷却板但未见冷却液进出口件号，可能风冷或接口集成在壳体内）"
        t_reasons.append(f"冷却板本体 {len(plate)} 个件号：{ids(plate)}")
    elif port:
        t_reasons.append("检出冷却液接口件但无冷却板本体件号 → 热路形态不明（冷板可能集成在壳体内）")
    else:
        t_reasons.append("无冷却板、无冷却液接口 → 倾向 T0（无主动冷却），但需照片确认")
    if tim:
        t_reasons.append(
            f"界面材料 {len(tim)} 个件号（合计 {tim_mass:.3f} kg）：存在需压紧并导热的发热面。"
            "**这是 T 轴证据，不能反推 P 轴有模块**")
        forms = {split_name(p)[0] for p in tim}
        t_reasons.append(f"TIM 形态（按件名）：{', '.join(sorted(forms))}；层数需按装配位置数点算")
    if stack:
        t_reasons.append(
            f"装在冷却板/功率模块上的垫块·保持·定位件 {len(stack)} 个：{ids(stack)}"
            "——**结构旁证**：双面夹持需要额外 spacer/locator 控制间距与压紧，"
            "单面散热通常不需要这些件。这提高 T2 可能性，但不构成 T2 的直接证据")
    t_reasons.append("T 轴必须另记三属性（脚本读不到，需人工补）：冷板材料 / TIM 层数与形态 / "
                     "压紧力是否解耦")

    # ---------- E 电路与开关（四栏，各给可否判定）----------
    busbar = sel_head(("busbar", "bus bar", "母排"))
    sensor = sel_head(("sensor", "shunt", "current sensor", "霍尔", "分流", "传感器"))
    pcb = sel_head(("pcb",))
    ferrite = sel_head(("ferrite", "铁氧体", "core"))
    e_level = {
        "E-a 拓扑": "不能判定（拓扑不能从件号反推，需原理图或板级反演）",
        "E-b 开关器件": (f"器件本体件号 {len(dev)} 个、实物 {pieces(dev):g} 颗（{ids(dev)}），类型只能取件名关键词，"
                        "**电压等级/并联数/配对要求均不能判定**"
                        if dev else "无器件本体件号 → 不能判定"),
        "E-c 互连": (f"母排本体 {len(busbar)} 个件号（{ids(busbar)}）；"
                     "螺栓母排/叠层母排/PCB 厚铜三者**未区分**"
                     if busbar else "无母排件号 → 可能为 PCB 铜箔或板载，不能排除"),
        "E-d 驱动与检测": (f"PCB {len(pcb)} 个件号、传感/分流 {len(sensor)} 个件号"
                          + (f"（{ids(sensor)}）" if sensor else "")
                          + "；集中/分布式驱动不能判定"
                          if pcb or sensor else "无 PCB 也无传感器件号 → 不能判定"),
    }
    if ferrite:
        e_level["备注"] = (f"检出磁件本体 {len(ferrite)} 个件号（{ids(ferrite)}）→ "
                          "存在共模/差模抑制需求，属 E 轴与储能槽交叉证据")

    # ---------- M 安装与压紧 ----------
    locator = sel_head(("locator", "dowel", "定位销", "定位"))
    clamp = [p for p in parts
             if head_is(p, ("spacer", "holder", "clamp", "clip", "spring", "shim",
                            "垫块", "压板", "弹片", "压簧"))
             and tail_is(p, ("cool", "power module", "plate", "冷却", "功率模块", "器件"))]
    adhesive = sel_head(("adhesive", "structural", "结构胶", "粘接"))
    pressfit = sel_head(("press-fit", "pressfit", "压接"))
    fastener = sel_head(("screw", "bolt", "nut", "screw assembly", "螺钉", "螺栓", "螺母"))

    m_reasons: list[str] = []
    m_level: str | None = None
    if pressfit:
        m_level = "M3（压接/Press-fit）"
        m_reasons.append(f"压接本体件：{ids(pressfit)}")
    elif clamp:
        m_level = "M1（独立压紧件/压板/垫块）——候选"
        m_reasons.append(
            f"压紧类本体件 {len(clamp)} 个：{ids(clamp)}；"
            "压紧点与器件数**不一一对应** → 提示压力可能经独立件解耦，而非螺钉直接压在器件上")
    elif fastener:
        m_level = "M0（螺钉直压）——默认推定，未经界面压力证据确认"
        m_reasons.append(f"紧固件本体 {len(fastener)} 个；未见独立压紧件")
    elif adhesive:
        m_level = "M4（粘接）"
        m_reasons.append(f"粘接类本体件：{ids(adhesive)}")
    else:
        m_reasons.append("清单**未见紧固件/压紧件/焊接/粘接任一证据** → 不能判定 M 轴。"
                         "紧固件常按 Standard Part 另表列出或按整机计重，"
                         "缺它不说明没有螺钉，只说明本清单不含")
    if locator:
        m_reasons.append(
            f"定位件本体 {len(locator)} 个、实物 {pieces(locator):g} 件（{ids(locator)}，"
            f"合计 {mass(locator):.4f} kg）→ 有独立定位销/定位结构，"
            "装配公差链不靠螺钉孔单独保证")
    if clamp and locator:
        m_reasons.append(
            "同时存在**定位件**与**独立压紧件** → M 轴应优先写 M1 而非 M0："
            "定位与压紧由两套不同零件分别承担，属于解耦结构")
    m_reasons.append("M 轴必须另记两属性（脚本读不到，需人工补）：定位方式（销/台阶/仿形槽/无）、"
                     "界面压力是否与器件本体解耦")

    conf_note = ("置信度不由脚本给：脚本只统计件号形态（证据类型），"
                 "结论可信度需分析者按证据一致性判断。两者是不同概念，"
                 "见 architecture-discovery.md §八.零。")

    named = sum(1 for p in parts if str(p.get("name") or "").strip())
    if not named:
        conf_note += (
            " **本清单没有零件名称列**：形态判定只能退化为靠材料种类，"
            "P/T/M 三轴的上限因此**显著降低**——材料种类能说明"
            "\"是铜/铁氧体/弹性体\"，不能说明\"是冷板还是母排\"。"
            "凡本清单给出的 P/T/M 等级，一律只能作为候选，置信度至多中低。")

    return {
        "packaging": {"axis": "P", "level": p_level, "reasons": p_reasons,
                      "evidence_type": "件号本体形态" if (mod or dev) else "无器件级证据",
                      "upper_limit": ceiling},
        "thermal": {"axis": "T", "level": t_level, "reasons": t_reasons,
                    "evidence_type": "件号本体 + 材料组" if (plate or tim) else "无热路件号",
                    "plate_count": len(plate), "plate_mass_kg": plate_mass,
                    "tim_count": len(tim), "tim_mass_kg": tim_mass,
                    "port_count": len(port), "stack_count": len(stack),
                    # 全部被认作冷却件的件号（plate + port 去重）。
                    # 供回归测试断言**单调性**：规则调整后不得比旧规则少认出任何一件。
                    # 少认 = 把液冷误判成 T0（Rear E-Drive 盲测的实际故障）。
                    "cooling_part_ids": sorted(plate_ids | {p["cs_part"] for p in port}),
                    "upper_limit": ceiling},
        "electrical": {"axis": "E", "fields": e_level,
                       "evidence_type": "件名关键词，非器件参数"},
        "mounting": {"axis": "M", "level": m_level, "reasons": m_reasons,
                     "evidence_type": "件号本体 + 紧固/定位件计数",
                     "locator_count": len(locator), "clamp_count": len(clamp),
                     "upper_limit": ceiling},
        "confidence_note": conf_note,
        "assembly_count": len(asm), "assembly_mass_kg": asm_mass,
    }


def render(result: dict[str, object]) -> str:
    lines: list[str] = []
    total = float(result["total_mass_kg"])
    lines.append(f"# 零件清单自检（{result['part_count']} 个件号，合计 {total:.3f} kg）")
    lines.append("")
    lines.append("> 质量与占比仅覆盖本次列出的件号，不代表整机质量；缺失字段以 `-` 标记，未补零。")
    lines.append("")

    section = {"n": 0}

    def head(title: str) -> None:
        section["n"] += 1
        lines.append(f"## {section['n']}. {title}")
        lines.append("")

    head("质量与占比")
    lines.append("| 件号 | 质量 kg | 占比 | 材料组 | 来源 |")
    lines.append("|---|---|---|---|---|")
    for part in result["parts"]:  # type: ignore[union-attr]
        weight = part["total_weight"]
        share = part.get("share")
        lines.append(
            "| {pid} | {w} | {s} | {g} | {src} |".format(
                pid=part["cs_part"],
                w=f"{weight:.3f}" if weight is not None else "-",
                s=f"{share * 100:.1f}%" if share else "-",
                g=part["material_group"],
                src=part["weight_source"],
            )
        )
    lines.append("")

    if result["fastener_mass_kg"]:
        lines.append(
            f"紧固件重量合计 {result['fastener_mass_kg']:.3f} kg"
            f"（{result['fastener_mass_kg'] / total * 100:.1f}%）；不能据此推算螺钉数量。"
        )
        lines.append("")

    head("材料组汇总")
    lines.append("| 材料组 | 质量 kg | 占比 |")
    lines.append("|---|---|---|")
    for group, mass in result["groups"].items():  # type: ignore[union-attr]
        lines.append(f"| {group} | {mass:.3f} | {mass / total * 100:.1f}% |")
    lines.append("")

    head("一致性告警")
    issues = result["issues"]
    if issues:
        lines.append("| 件号 | 类型 | 说明 |")
        lines.append("|---|---|---|")
        for issue in issues:  # type: ignore[union-attr]
            lines.append(f"| {issue['part']} | {issue['type']} | {issue['detail']} |")
    else:
        lines.append("未发现重量/尺寸/重复件号的字段级异常。")
    lines.append("")

    missing_rows = [p for p in result["parts"] if p["missing"]]  # type: ignore[union-attr]
    if missing_rows:
        head("缺失字段（不补零）")
        for part in missing_rows:
            lines.append(f"- `{part['cs_part']}`：缺少 {', '.join(part['missing'])}")
        lines.append("")

    if result["power_gap_hint"]:
        head("提示")
        lines.append(
            "清单中未出现功率模块/散热界面相关关键词。若分析对象是功率电子总成，"
            "先确认是否漏采；**不得据此断言该总成没有功率模块或没有界面材料**。"
        )
        lines.append("")

    tim = result["tim_evidence"]  # type: ignore[index]
    if tim["parts"]:
        head("功率级存在性旁证（不是形态证据）")
        lines.append(
            f"检测到界面材料 {len(tim['parts'])} 件（合计 {tim['mass_kg']:.3f} kg）："
            f"{', '.join(tim['parts'])}。"
        )
        lines.append(
            "界面材料只能证明**存在需要压紧并导热的发热面**，因此功率级**存在**；"
            "它落在 **T（热）轴**，**不能**反推 **P（封装）轴**——不能证明功率级是模块、"
            "分立器件还是板载集成。形态判定必须独立走 P 轴，不得由 TIM 直接推模块。"
        )
        lines.append("")

    arch = result["integration_level"]  # type: ignore[index]
    head("架构四维度粗判（提示，非结论；四轴独立，不得排序比较）")
    lines.append(
        "> 四个维度是**互相独立的设计轴**，等级数字不可相加、不可比较大小。"
        "典型错误：\"用了双面冷却所以集成度更高\"——双面冷却属于 T 轴，"
        "与 P 轴的封装集成度无关。"
    )
    lines.append("")
    lines.append("| 维度 | 候选等级 | 证据类型 |")
    lines.append("|---|---|---|")
    pk = arch["packaging"]
    th = arch["thermal"]
    mt = arch["mounting"]
    lines.append(f"| **P** 封装与集成 | {pk['level'] or '**不能判定**'} | {pk['evidence_type']} |")
    lines.append(f"| **T** 冷却与热路 | {th['level'] or '**不能判定**'} | {th['evidence_type']} |")
    lines.append(f"| **E** 电路与开关 | 见下方四栏 | {arch['electrical']['evidence_type']} |")
    lines.append(f"| **M** 安装与压紧 | {mt['level'] or '**不能判定**'} | {mt['evidence_type']} |")
    lines.append("")

    for title, block in (("P 封装与集成程度", pk), ("T 冷却与热路", th),
                         ("M 安装与压紧", mt)):
        lines.append(f"**{title}**")
        for reason in block["reasons"]:  # type: ignore[union-attr]
            lines.append(f"- {reason}")
        lines.append("")

    lines.append("**E 电路与开关（四栏）**")
    for field, note in arch["electrical"]["fields"].items():  # type: ignore[union-attr]
        lines.append(f"- *{field}*：{note}")
    lines.append("")

    lines.append(
        "**逐维度声明上限时照抄此句式**："
        f"该维度判定上限受 {arch['assembly_count']} 个未分解 Assembly 叶子限制"
        f"（合计 {arch['assembly_mass_kg']:.3f} kg）。"
    )
    lines.append("")
    lines.append(f"> {arch['confidence_note']}")
    lines.append("")

    slots = result["function_slots"]  # type: ignore[index]
    empty = [s["slot"] for s in slots if not s["covered"]]
    head("功能槽覆盖（九类，无对应件必须单列）")
    lines.append("| 功能槽 | 命中件 | 实物数 | 状态 |")
    lines.append("|---|---|---|---|")
    for slot in slots:
        status = "已覆盖" if slot["covered"] else "**无对应件 → 必须单列**"
        parts_cell = ", ".join(slot["parts"][:4]) or "-"
        if len(slot["parts"]) > 4:
            parts_cell += f" 等 {len(slot['parts'])} 个件号"
        pieces = slot.get("piece_count")
        piece_cell = f"{pieces:g}" if pieces else "-"
        lines.append(f"| {slot['slot']} | {parts_cell} | {piece_cell} | {status} |")
    lines.append("")

    # "电能变换"槽空缺但"开关动作"槽有件的特殊情况：
    # 这不是漏采，而是**改造写法**——由分立器件承担变换功能，不存在模块本体。
    switch_slot = next((s for s in slots if s["slot"] == "开关动作"), None)
    conv_slot = next((s for s in slots if s["slot"] == "电能变换"), None)
    if (conv_slot and not conv_slot["covered"] and switch_slot and switch_slot["covered"]):
        lines.append(
            "> **电能变换槽的写法规矩**：本清单有开关器件、无模块本体。"
            "该槽不能写\"电能变换功能缺失\"，必须写成"
            "\"变换功能由 " + f"{switch_slot.get('piece_count', 0):g} 颗分立开关器件承担"
            "（" + ", ".join(switch_slot["parts"][:3]) + "），清单内不存在模块/功率级本体件号\"。"
            "前半句是功能归属（E1），后半句是存在性（E0）；"
            "**不得由此推\"模块已被分立替代\"**。"
        )
        lines.append("")
        empty = [s["slot"] for s in slots if not s["covered"] and s["slot"] != "电能变换"]
    else:
        empty = [s["slot"] for s in slots if not s["covered"]]

    if empty:
        lines.append(
            f"**未覆盖功能槽 {len(empty)} 个**：{', '.join(empty)}。"
            "这些槽位必须在报告里各占一行，写明三种可能解释、区分方法、旁证结论；"
            "**不允许因未命中关键词而省略**（清单缺失 ≠ 产品没有）。"
        )
        lines.append("")
    return "\n".join(lines)


def emit(text: str) -> None:
    """输出到 stdout，遇到编码不支持时退到 UTF-8，而不是直接崩掉。

    Windows 上 stdout 被重定向为管道/文件时，编码取自系统代码页（cp936/cp1252），
    中文会触发 UnicodeEncodeError 让整个脚本以退出码 1 结束。这里只在真的编码
    失败时才切 UTF-8，避免在支持中文的环境里改动原生行为。

    注意：PowerShell 5.1 的 `>` 按 [Console]::OutputEncoding 解码子进程输出，
    即使脚本输出 UTF-8 也可能出现乱码。需要干净文件请用 `--out`，不要用 shell 重定向。
    """
    try:
        print(text)
        return
    except UnicodeEncodeError:
        pass
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if reconfigure is not None:
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            return
    print(text)


def main() -> int:
    parser = argparse.ArgumentParser(description="零件清单自检与质量占比核算")
    parser.add_argument("path", type=Path, help="CSV/XLSX 零件清单")
    parser.add_argument("--format", choices=("md", "json"), default="md")
    parser.add_argument("--out", type=Path, default=None, help="输出文件，缺省打印到标准输出")
    args = parser.parse_args()

    if not args.path.exists():
        raise SystemExit(f"找不到文件：{args.path}")

    result = analyze(read_rows(args.path))
    text = (
        json.dumps(result, ensure_ascii=False, indent=2)
        if args.format == "json"
        else render(result)
    )
    if args.out:
        args.out.write_text(text.rstrip("\n") + "\n", encoding="utf-8")
        emit(f"已写入 {args.out}")
    else:
        emit(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
