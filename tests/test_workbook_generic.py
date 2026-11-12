#!/usr/bin/env python3
"""Should-Cost 工作簿生成器的**产品无关性**回归测试（仅标准库 + openpyxl）。

**为什么需要这个文件**：`tools/build_should_cost_workbook_v2.py` 曾经把 DC-DC 的
三个事实写死进了"通用工具"：

1. `function_category(..., prefix="DCDC")` —— 功能域前缀默认 DCDC，逆变器 53 个叶件
   全被标成 `DCDC/…`；
2. 工艺路线只能由 attrs 表的 `Procurement Object` 查到 —— 没有该列的产品 0/53 命中，
   于是 bottom-up 全为 0，`4-Summary` 直接塌掉；
3. 文档串断言"BOM 21 个叶件全部外购"。

修复后三条都变成**三级降级**（有证据精确路由 → 件名推断 → 显式 Unknown），
产品名只影响 `Function` 前缀。本文件断言的是这些**不变量**，不是快照：
快照只能告诉你"输出变了"，测不出"换一个总成就不工作了"。

| 断言 | 来源故障 |
|---|---|
| 同一件名在两个产品清单里得到相同 Function+Category | 前缀硬编码 DCDC |
| 换产品名只改前缀，其余分类逐条不变 | 前缀硬编码 DCDC |
| 逆变器叶件 100% 有非空工序推断 | attrs 缺失 → 0 命中 |
| 件名主体优先：`Gasket, … Power Module` 不得进功率级 | 整串匹配让父件上下文盖掉部件类型 |
| `is_buy()` 命中的件永不 `Bottom-up = Y` | PCB/MOSFET 被当成冲压建模 |
| `Part Type = Standard` 永不 `Bottom-up = Y` | 紧固件被套"装配"模板 |
| 只给 `--bom` 也能跑出 ≥1 个 bottom-up 行与完整 3-ProcessCost | 无 attrs 时汇总塌掉 |
| 生成器 2-BOM 列数与 `make_cols()` 一致 | 文档滞后两代（30 vs 43 列） |
| 每件 `Supplier Process` 只取其最后一道工序 | 缺 `Is Last Step` → 工序费重复累计 |

运行：`py -3 -X utf8 tests/test_workbook_generic.py`
"""

from __future__ import annotations

import csv
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))

from domain_classify import (  # noqa: E402
    func_prefix_for,
    function_category,
    is_buy,
    name_subject,
    route_for,
    route_strength,
)

GENERATOR = TOOLS / "build_should_cost_workbook_v2.py"
DCDC_LEAVES = ROOT / "data" / "ex90-ultra-DCDC" / "derived" / "dcdc-leaves-v3.csv"
DCDC_ATTRS = ROOT / "data" / "ex90-ultra-DCDC" / "derived" / "dcdc-purchased-attributes-v3.csv"
DCDC_PRICES = ROOT / "data" / "ex90-ultra-DCDC" / "derived" / "dcdc-price-evidence-v6.csv"
INV_LEAVES = ROOT / "data" / "ex90-rear-inverter" / "derived" / "inverter-assy-leaves.csv"
OBCM_LEAVES = ROOT / "data" / "ex90-ultra-obcm" / "derived" / "obcm-leaves.csv"


def read_leaves(path: Path) -> list[dict]:
    """只读 Name / Material Species / Part Type / Total Weight，够断言用。"""
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig", newline="") as f:
        return [
            {
                "name": (r.get("Name") or "").strip(),
                "species": (r.get("Material Species") or "").strip(),
                "ptype": (r.get("Part Type") or "").strip(),
                "total": float((r.get("Total Weight (kg)") or "0").strip() or 0),
            }
            for r in csv.DictReader(f)
        ]


def run_generator(bom: Path, out: Path, product: str,
                  attrs: Path | None = None, prices: Path | None = None,
                  func_prefix: str | None = None) -> str:
    cmd = [sys.executable, "-X", "utf8", str(GENERATOR),
           "--bom", str(bom), "--product", product, "--out", str(out)]
    if attrs:
        cmd += ["--attrs", str(attrs)]
    if prices:
        cmd += ["--prices", str(prices)]
    if func_prefix:
        cmd += ["--function-prefix", func_prefix]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        raise SystemExit(f"生成器执行失败（{product}）：\n{proc.stderr}\n{proc.stdout}")
    return proc.stdout


def rows_of(xlsx: Path, sheet: str) -> tuple[list[str], list[dict]]:
    import openpyxl
    wb = openpyxl.load_workbook(xlsx)
    if sheet not in wb.sheetnames:
        raise SystemExit(f"{xlsx.name} 缺少工作表 {sheet}；实际={wb.sheetnames}")
    ws = wb[sheet]
    hdr = [c.value for c in ws[4]]
    idx = {h: i for i, h in enumerate(hdr) if h}
    out = []
    for r in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
        if not r[idx.get("Part No", 0)]:
            continue
        out.append({h: r[i] for h, i in idx.items()})
    return [h for h in hdr if h], out


def main() -> int:
    failures: list[str] = []
    checks = 0

    def expect(cond: bool, msg: str) -> None:
        nonlocal checks
        checks += 1
        if not cond:
            failures.append(msg)

    dcdc = read_leaves(DCDC_LEAVES)
    inv = read_leaves(INV_LEAVES)
    obcm = read_leaves(OBCM_LEAVES)
    expect(bool(dcdc) and bool(inv),
           f"缺测试数据：dcdc={len(dcdc)} 行, inverter={len(inv)} 行")

    # ---------- 不变量 1：功能域前缀只由产品决定，不由总成内容决定 ----------
    expect(func_prefix_for("EX90 DC-DC Converter") == "DCDC", "DC-DC 产品名前缀应为 DCDC")
    expect(func_prefix_for("Rear Inverter") == "INV", "Rear Inverter 前缀应为 INV")
    expect(func_prefix_for("On-Board Charger") == "OBC", "OBC 前缀应为 OBC")
    expect(func_prefix_for("Rear Drive Unit") == "GEN",
           "未知产品名必须落通用前缀 GEN，不得默认某个具体产品")
    expect(func_prefix_for("Weird Name", "BMS") == "BMS", "显式 override 必须优先于产品名推导")

    # ---------- 不变量 2：换产品名只换前缀，其余分类逐条不变 ----------
    # 注意：Box 功能域的件其 Function 恒为 "Box"，**不跟随产品前缀**（壳体/冷却/连接器
    # 属于结构功能域，不属于某个电气产品）。只有电气功能域的件才带产品前缀。
    names = sorted({e["name"] for e in inv if e["name"]})
    for nm in names:
        sp = next((e["species"] for e in inv if e["name"] == nm), "")
        f_inv, c_inv = function_category(nm, sp, "INV")
        f_obc, c_obc = function_category(nm, sp, "OBC")
        expect(c_inv == c_obc,
               f"换产品改变了功能域分类：{nm!r} → INV/{c_inv} vs OBC/{c_obc}")
        if c_inv == "Box" or f_inv == "Box":
            expect(f_inv == "Box" and f_obc == "Box",
                   f"Box 类功能域被加了产品前缀：{nm!r} → {f_inv}/{f_obc}")
        else:
            expect(f_inv == "INV" and f_obc == "OBC",
                   f"电气功能域前缀未跟随产品：{nm!r} → {f_inv}/{f_obc}")

    # ---------- 不变量 3：同一件名跨产品清单必须同类 ----------
    by_name: dict[str, list[tuple[str, str]]] = {}
    for lst in (dcdc, inv, obcm):
        for e in lst:
            if not e["name"]:
                continue
            by_name.setdefault(e["name"], []).append(
                function_category(e["name"], e["species"], "GEN")
            )
    shared = {k: v for k, v in by_name.items() if len(v) > 1}
    for nm, cats in shared.items():
        expect(len(set(cats)) == 1,
               f"同一件名在不同产品清单里分类不一致：{nm!r} → {cats}")
    expect(bool(shared), "三个产品清单之间没有同名件，跨产品一致性断言形同虚设")

    # ---------- 不变量 4：件名主体优先（父件上下文不得盖掉部件类型）----------
    cases = [
        ("Gasket, Cooling Plate Rear Inverter", "密封紧固其他", False,
         "Gasket 是密封件，'cooling plate' 只是父级上下文"),
        ("Gasket, Power Module", "密封紧固其他", False,
         "Gasket 不得被判为 PM 功率级"),
        ("Ferrite Core, Current Sensor 2", "EMC", False,
         "Ferrite Core 应归 EMC，不得被父件 Current Sensor 拖成电流传感"),
        ("PCB, Current Sensor 2", "PCB / PCBA", False,
         "PCB 应归 PCB / PCBA，不得被父件上下文拖成电流传感"),
        ("MOSFET, Power Module", "PM 功率级", False,
         "MOSFET 本体就是功率级，不能因主体优先而漏判"),
        ("SH Film Capacitor", "DC-Link 电容", False,
         "薄膜电容必须进 DC-Link 电容"),
        ("PCB, Power Module", "PCB / PCBA", False,
         "PCB 必须进 PCB / PCBA，不得因父件是 Power Module 而变功率级"),
        ("Adapter, Rear Inverter Coolant Inlet", "冷却与热", False,
         "冷却接口 adapter 必须归冷却与热，不能被通用 adapter 规则吃掉"),
        ("Sensor, Current", "电流传感", False,
         "传感器本体必须归电流传感（收紧父件规则后不得误伤）"),
    ]
    for nm, want_cat, forbid, why in cases:
        head, _ = name_subject(nm)
        f, c = function_category(nm, "", "INV")
        if forbid and want_cat is None:
            expect(c != "PM 功率级" and c != "电流传感", f"主体优先失效：{nm!r} → INV/{c}（{why}）")
        elif want_cat is not None:
            expect(c == want_cat, f"分类错误：{nm!r} → INV/{c}，期望 {want_cat}（{why}）")
        expect(bool(head), f"name_subject 未能拆出主体：{nm!r}")

    # ---------- 不变量 5：外购件判定 ----------
    for nm in ("PCB", "PCBA Light", "MOSFET", "Fuse", "Connector", "Ferrite Core",
               "Thermal Paste", "NVH Foam, Busbar", "O Ring", "Gasket, Power Module"):
        expect(is_buy(nm), f"外购件未被识别：{nm!r} —— 它会被当成自制件建模，"
                           "PCB 会被算成冲压件")
    for nm in ("Housing, Rear Inverter", "Cooling Plate", "Cold Plate",
               "Bottom Cover", "Busbar, DC", "Access Cover"):
        expect(not is_buy(nm), f"自制件被误判为外购：{nm!r} —— 它的工艺模型会丢失")

    # ---------- 不变量 6：所有叶件都有非空工序推断 ----------
    for tag, lst in (("inverter", inv), ("dcdc", dcdc)):
        if not lst:
            continue
        missing = []
        weak = 0
        for e in lst:
            if not e["name"]:
                continue
            text, steps = route_for(e["name"], e["species"])
            if not text or not steps:
                missing.append(e["name"])
            if route_strength(text) == "Weak":
                weak += 1
        expect(not missing,
               f"[{tag}] 有叶件推断不出任何工序：{missing[:5]} —— 降级链断了")
        tot = sum(1 for e in lst if e["name"])
        expect(weak < tot,
               f"[{tag}] 全部 {tot} 个叶件都只判到装配级 —— 推断规则对该产品无效")

    # ---------- 不变量 7：生成器只给 --bom 也要有 bottom-up 与工序表 ----------
    expect(INV_LEAVES.exists(), f"缺逆变器叶件清单：{INV_LEAVES}")
    tmpdir = Path(tempfile.mkdtemp(prefix="wbtest_"))
    inv_only = tmpdir / "inv_bom_only.xlsx"
    out = run_generator(INV_LEAVES, inv_only, "Rear Inverter")
    hdr, rows = rows_of(inv_only, "2-BOM")
    bu = [r for r in rows if str(r.get("Bottom-up Model (Y/N)") or "").upper() == "Y"]
    expect(len(rows) > 0, "无 attrs 时 2-BOM 为空 —— 降级链没有生效")
    expect(len(bu) > 0, "无 attrs 时 bottom-up 行数为 0 —— attrs 依赖没被解除")
    expect(len(bu) < len(rows),
           f"无 attrs 时全部 {len(rows)} 行都建了工艺模型 —— 外购件/标准件门控失效")
    prefixes = {str(r.get("Function") or "") for r in rows}
    expect("INV" in prefixes,
           f"仅给 --bom 时电气功能域前缀未跟随产品，实际前缀={sorted(prefixes)}")
    expect("DCDC" not in prefixes, f"逆变器工作簿里出现 DCDC 前缀：{sorted(prefixes)}")
    expect("Rear Inverter" in out, "生成器输出未回显产品名，无法确认前缀推导源")

    _ph, p3 = rows_of(inv_only, "3-ProcessCost")
    expect(len(p3) > 0, "无 attrs 时 3-ProcessCost 为空 —— 汇总会塌掉")
    covered = {str(r.get("Part No") or "") for r in p3}
    bu_parts = {str(r.get("Part No") or "") for r in bu}
    expect(covered == bu_parts,
           f"3-ProcessCost 覆盖件号与 bottom-up 件号不一致："
           f"表内多出 {sorted(covered - bu_parts)[:3]}，缺少 {sorted(bu_parts - covered)[:3]}")
    last_flags = [r for r in p3 if str(r.get("Is Last Step (Y/N)") or "").upper() == "Y"]
    expect(len(last_flags) == len(bu_parts),
           f"标 Is Last Step=Y 的工序 {len(last_flags)} 条 ≠ bottom-up 件数 {len(bu_parts)}"
           " —— 同一件的工序会被连加多遍")

    # ---------- 不变量 8：外购件/标准件永不 bottom-up ----------
    bad_buy = [r.get("Part No") for r in bu if is_buy(str(r.get("Name") or ""))]
    expect(not bad_buy, f"外购件被建了工艺模型：{bad_buy}")

    # ---------- 不变量 9：完整输入下标准件也不 bottom-up，且列数正确 ----------
    if DCDC_LEAVES.exists() and DCDC_ATTRS.exists() and DCDC_PRICES.exists():
        full = tmpdir / "dcdc_full.xlsx"
        run_generator(DCDC_LEAVES, full, "EX90 DC-DC Converter",
                      DCDC_ATTRS, DCDC_PRICES)
        h2, rows2 = rows_of(full, "2-BOM")
        expect(len(h2) == 43, f"2-BOM 列数 {len(h2)} ≠ 43（workbook-spec 已同步为 43）")
        bu2 = [r for r in rows2 if str(r.get("Bottom-up Model (Y/N)") or "").upper() == "Y"]
        std_bu = [r.get("Part No") for r in bu2 if is_buy(str(r.get("Name") or ""))]
        expect(not std_bu, f"DC-DC 外购件被建了工艺模型：{std_bu}")
        expect(len(bu2) == 10,
               f"DC-DC bottom-up 件数 {len(bu2)} ≠ 10（已交付基线，漂移说明分类器动了）")
        _ph2, p32 = rows_of(full, "3-ProcessCost")
        y2 = [r for r in p32 if str(r.get("Is Last Step (Y/N)") or "").upper() == "Y"]
        expect(len(p32) == 31, f"DC-DC 工序行数 {len(p32)} ≠ 31（已交付基线）")
        expect(len(y2) == 10, f"DC-DC 末道工序 {len(y2)} 条 ≠ 10")
    else:
        expect(False, "缺 DC-DC 完整输入（leaves-v3 / attrs-v3 / prices-v6），"
                      "无法验证已交付基线未漂移")

    if failures:
        print("--- 产品无关性不变量失败 ---")
        for f in failures:
            print("  • " + f)
        return 1

    print(f"全部通过（{checks} 条产品无关性不变量；数据："
          f"DC-DC {len(dcdc)} 行 / 逆变器 {len(inv)} 行 / OBCM {len(obcm)} 行；"
          f"跨产品同名件 {len(shared)} 个）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())