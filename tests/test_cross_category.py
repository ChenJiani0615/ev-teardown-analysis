#!/usr/bin/env python3
"""跨品类语义不变量测试（仅标准库）。

**为什么需要这个文件**：`test_part_table_check.py` 是基线快照比对，只能告诉你
"输出变了"，不能告诉你"语义错了"。它无法发现"把液冷判成 T0"这种错误——
只要基线一开始就是错的，快照永远正确。

本文件断言的是**不变量**，不是快照。每一条都对应一次真实故障：

| 断言 | 来源故障 |
|---|---|
| 机械总成的液冷必须被认出（T 轴不得为 None） | Rear E-Drive 盲测：`Cover, Coolant Channel` + `Pipe, Coolant In/Out` 全部漏检 → 输出"倾向 T0（无主动冷却）" |
| 机械堵盖不得计入"储能与滤波" | `End Cap` / `Cap, ... Clutch` 命中裸关键词 `cap` → 3 个机械件被算成电容 |
| 功率电子总成的"储能与滤波"必须仍被覆盖 | 收紧 `cap` → `capacitor` 后，DCDC/OBCM/逆变器不得因此漏掉真实电容 |

**两个故障的共同根因**：关键词表按 3 个功率电子样本调出，对机械总成过拟合。
所有旧 fixture 都来自功率电子，所以永远测不出来。
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".github" / "skills" / "ev-teardown-benchmark" / "scripts" / "part_table_check.py"

# (清单相对路径, 品类说明)
CASES = {
    "mechanical": ROOT / "tests" / "fixtures" / "edrive-rear-mechanical.csv",
    "dcdc": ROOT / "data" / "ex90-ultra-DCDC" / "derived" / "dcdc-leaves.csv",
    "obcm": ROOT / "data" / "ex90-ultra-obcm" / "derived" / "obcm-leaves.csv",
    "inverter": ROOT / "data" / "ex90-rear-inverter" / "derived" / "inverter-assy-leaves.csv",
}


def load(name: str) -> dict:
    path = CASES[name]
    if not path.exists():
        raise SystemExit(f"跨品类测试需要 {path}，但它不存在")
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(path), "--format", "json"],
        capture_output=True, text=True, encoding="utf-8",
    )
    if proc.returncode != 0:
        raise SystemExit(f"脚本在 {name} 上执行失败：\n{proc.stderr}")
    return json.loads(proc.stdout)


def slot(result: dict, name: str) -> dict:
    for entry in result["function_slots"]:
        if entry["slot"] == name:
            return entry
    raise SystemExit(f"功能槽 {name} 不存在")


def main() -> int:
    failures: list[str] = []
    checks = 0

    def expect(cond: bool, msg: str) -> None:
        nonlocal checks
        checks += 1
        if not cond:
            failures.append(msg)

    # ---------- 不变量 1：机械总成的液冷必须被认出 ----------
    mech = load("mechanical")
    t = mech["integration_level"]["thermal"]
    expect(t["level"] is not None,
           f"[mechanical] T 轴为 None —— 液冷被漏检。level={t['level']!r}；reasons={t['reasons']}")
    expect(t.get("plate_count", 0) >= 1,
           f"[mechanical] 未识别到冷却通道本体（plate_count={t.get('plate_count')}）。"
           " `Cover, Coolant Channel Rear E-Drive` 应计入")
    expect(t.get("port_count", 0) >= 1,
           f"[mechanical] 未识别到冷却液进出口（port_count={t.get('port_count')}）。"
           " `Pipe, ... Coolant Inlet/Outlet` 应计入")
    expect(isinstance(t["level"], str) and ("T1" in t["level"] or "T2" in t["level"]),
           f"[mechanical] T 轴等级不是液冷档：{t['level']!r}")

    # ---------- 不变量 2：机械堵盖不得计入储能与滤波 ----------
    stor = slot(mech, "储能与滤波")
    expect(not stor["covered"],
           f"[mechanical] 储能与滤波被假阳性覆盖，命中件号={stor['parts']}。"
           " 这三个应是 End Cap / Cap, Clutch 机械堵盖，不是电容")

    # ---------- 不变量 3：功率电子总成的储能与滤波不得因收紧关键词而丢失 ----------
    for name in ("dcdc", "obcm", "inverter"):
        data = load(name)
        s = slot(data, "储能与滤波")
        expect(s["covered"],
               f"[{name}] 储能与滤波不再被覆盖 —— 收紧 `cap` 时误伤了真实电容/磁件。"
               f" 命中件号={s['parts']}")

    # ---------- 不变量 4：功率电子总成的 T 轴不得比修复前更弱 ----------
    # 注意：DCDC/OBCM 的「冷却板」从未单列，其冷却件形态是 `Cooling Pipe, Inlet`，
    # 本体是 Pipe。旧规则只认 plate 本体 → 这两件被漏检，T 轴为 None（既有缺陷）。
    # 新规则认得它们 → 不入 plate（本体非结构件）但入 port（本体是管）。
    # 这里断言的是**单调性**：新规则认出的冷却件必须是旧规则的超集。
    OLD_PLATE_HEADS = ("coldplate", "cold plate", "cooling plate", "cooler",
                       "heatsink", "heat sink", "radiator", "冷板", "散热器")
    OLD_PORT_HEADS = ("adapter", "fitting", "hose", "inlet", "outlet", "nozzle",
                      "接头", "水管", "进出 口", "进出口")
    for name in ("dcdc", "obcm", "inverter"):
        data = load(name)
        t2 = data["integration_level"]["thermal"]
        hits = set(t2.get("cooling_part_ids", []))
        # 旧规则会命中的件（用旧词表在原始清单上重算）
        legacy = set()
        for p in data["parts"]:
            head, tail = (p["name"].split(",", 1) + [""])[:2]
            head, tail = head.strip().lower(), tail.strip().lower()
            intrin = f'{head} {str(p.get("material_species") or "").lower()}'
            if any(w in intrin for w in OLD_PLATE_HEADS):
                legacy.add(p["cs_part"])
            elif any(w in intrin for w in OLD_PORT_HEADS) and any(
                    w in tail for w in ("coolant", "cooling", "inlet", "outlet", "water", "冷却", "液")):
                legacy.add(p["cs_part"])
        missed = legacy - hits
        expect(not missed,
               f"[{name}] 新规则漏掉了旧规则本已认出的冷却件：{sorted(missed)}")

    # ---------- 不变量 5：DCDC 的 Cooling Pipe 必须被认出（旧规则漏检的真实案例）----------
    if "dcdc" in CASES and CASES["dcdc"].exists():
        dcdc = load("dcdc")
        t3 = dcdc["integration_level"]["thermal"]
        expect(t3.get("port_count", 0) >= 2,
               f"[dcdc] `Cooling Pipe, Inlet` / `Cooling Pipe, Outlet` 未被识别为冷却液接口"
               f"（port_count={t3.get('port_count')}）。这是与 E-Drive 同源的漏检缺陷")

    if failures:
        print("--- 跨品类不变量失败 ---")
        for f in failures:
            print(f)
        return 1

    print(f"全部通过（{checks} 条跨品类不变量，覆盖 4 个品类：机械电驱 / DC-DC / OBCM / 逆变器）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())