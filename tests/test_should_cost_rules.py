#!/usr/bin/env python3
"""Should-Cost 规则不变量测试（Phase 2A，仅标准库，自执行）。

**为什么需要这个文件**：生成器 `build_should_cost_workbook.py` 的输出是 xlsx，
公式对错要靠 Excel 重算才能看到。但**规则**（价格状态语义、完整度判定、
防重复计价、HG 定义）是纯 Python 逻辑 + 文档约定，可以在不打开 Excel 的
情况下断言不变量。

每条断言都对应 Phase 2A 的一项用户纠正：

| 断言组 | 来源纠正 |
|---|---|
| 案例 1：已知采购价正确计算标记 | 任务 A——七类价格状态，Actual Procurement 是唯一可标 Verified 的价格类 |
| 案例 2：Unknown 不显示零成本、不完整 PPC 不被标完整 | 任务 B——Unknown 不得被 N()/SUM 静默当 0；Unknown 计数 > 0 时 PPC 只能标"部分" |
| 案例 3：外购 PCBA 与内部展开不重复计价 | 任务 C——外购总成价与内部展开成本不得双重计价 |
| HG 定义 | 纠正 1——HG4=Implemented、HG5=Saving realized，与证据可信度是两个独立坐标 |

测试数据说明：fixture `should-cost-rules-demo.csv` 为**测试构造数据**，
其中的价格与状态列仅用于验证生成器逻辑，不代表任何真实采购价。
"""

from __future__ import annotations

import csv
import importlib.util
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GEN = ROOT / "tools" / "build_should_cost_workbook.py"
FIXTURE = ROOT / "tests" / "fixtures" / "should-cost-rules-demo.csv"
BOSCH_REF = ROOT / "tmp" / "ref_text" / "20241119_China PTU Product Benchmark Final CTM.txt"

# ------------------------------------------------------------------ 加载生成器模块

spec = importlib.util.spec_from_file_location("bscw", GEN)
bscw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bscw)


def main() -> int:
    failures: list[str] = []
    checks = 0

    def expect(cond: bool, msg: str) -> None:
        nonlocal checks
        checks += 1
        if not cond:
            failures.append(msg)

    # ---------- HG 定义验证（纠正 1） ----------
    hg_text = (BOSCH_REF.read_text(encoding="utf-8", errors="replace")
               if BOSCH_REF.exists() else "")
    if not hg_text:
        print(f"[skip] Bosch 参考文本不存在：{BOSCH_REF}")
    else:
        for token in ("Implemented", "Saving realized"):
            expect(token in hg_text,
                   f"[HG] Bosch 原文应含 '{token}'（HG4/HG5 定义出处）")
        # HG4/HG5 不得被映射成证据档位
        expect(bscw.PRICE_STATUS.get("Actual Procurement") == "Verified",
               "[HG] Actual Procurement 应是唯一可标 Verified 的价格类")
        expect("Implemented" not in bscw.PRICE_STATUS.values()
               and "Saving realized" not in bscw.PRICE_STATUS.values(),
               "[HG] HG 档位（Implemented/Saving realized）不得混入价格状态映射——"
               "实施成熟度与证据可信度是两个独立坐标")

    # ---------- 价格状态七类（任务 A） ----------
    seven = {"Actual Procurement", "Supplier Quotation", "Comparable",
             "Bottom-up Estimate", "Parametric Estimate", "Assumption", "Unknown"}
    expect(set(bscw.PRICE_STATUS.keys()) == seven,
           f"[任务A] PRICE_STATUS 应恰好七类，实际：{sorted(bscw.PRICE_STATUS.keys())}")
    expect(bscw.PRICE_STATUS["Supplier Quotation"] != "Verified",
           "[任务A] Supplier Quotation 是报价不是已发生成本，不得标 Verified")
    expect(bscw.PRICE_STATUS["Unknown"] == "Unknown",
           "[任务A] Unknown 价格的证据档必须是 Unknown 本身")
    # 默认 Unknown：无输入行不得自动猜
    expect(bscw.price_status_of({}) == "Unknown",
           "[任务A] 无输入时 price_status_of 必须默认 Unknown，不自动猜")
    expect(bscw.price_status_of({"price_status": "编造的状态"}) == "Unknown",
           "[任务A] 不在七类内的输入必须回落 Unknown，不得放行")

    # ---------- 生成工作簿并检查规则落地（任务 B/C） ----------
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "wb.xlsx"
        proc = subprocess.run(
            [sys.executable, str(GEN), "--bom", str(FIXTURE),
             "--out", str(out), "--product", "规则测试总成"],
            capture_output=True, text=True, encoding="utf-8")
        expect(proc.returncode == 0,
               f"[生成] 生成器在 fixture 上失败：{proc.stderr[:500]}")
        if proc.returncode != 0:
            for f in failures:
                print(f"FAIL {f}")
            return 1

        from openpyxl import load_workbook
        wb = load_workbook(out)

        # ---- 2-BOM 列与状态 ----
        ws2 = wb["2-BOM"]
        hdr = [ws2.cell(row=3, column=c).value for c in range(1, 25)]
        for col in ("Price Type", "Price Status", "Cost Boundary", "Quantity Basis"):
            expect(col in hdr, f"[任务A] 2-BOM 表头缺列：{col}")
        expect(hdr.index("Tooling CapEx (kCNY)") == 22,
               f"[任务C] Tooling CapEx 应在 W 列（第 23 列），实际位置："
               f"{hdr.index('Tooling CapEx (kCNY)') + 1 if 'Tooling CapEx (kCNY)' in hdr else '缺失'}")

        # 读回 BOM 行：fixture 里 TEST-BUY1 有价（Actual Procurement），
        # TEST-BUY2 无价（Unknown），TEST-MAKE1 自制
        rows = {}
        for r in range(4, 3 + 100):
            pn = ws2.cell(row=r, column=2).value
            if not pn:
                break
            rows[pn] = {hdr[c - 1]: ws2.cell(row=r, column=c).value
                        for c in range(1, 25)}
        expect("TEST-BUY1" in rows and "TEST-BUY2" in rows and "TEST-MAKE1" in rows,
               f"[fixture] 三个测试件必须都在 2-BOM：实际 {sorted(rows.keys())}")

        b1, b2, m1 = rows.get("TEST-BUY1", {}), rows.get("TEST-BUY2", {}), rows.get("TEST-MAKE1", {})
        # 案例 1：已知采购价正确计算标记
        expect(b1.get("Purchased Price") == 15.5,
               f"[案例1] TEST-BUY1 采购价应为 15.5，实际 {b1.get('Purchased Price')!r}")
        expect(b1.get("Price Status") == "Actual Procurement",
               f"[案例1] TEST-BUY1 价格状态应为 Actual Procurement，实际 {b1.get('Price Status')!r}")
        expect(b1.get("Cost Boundary") != "" and b1.get("Cost Boundary") is not None,
               "[案例1] 有价外购件必须填成本边界（Cost Boundary）")
        expect(str(b1.get("Quantity Basis") or "").startswith("per piece"),
               f"[案例1] 有价外购件必须填数量口径，实际 {b1.get('Quantity Basis')!r}")
        # 案例 2：Unknown 不显示零成本
        expect(b2.get("Purchased Price") in (None, ""),
               f"[案例2] Unknown 外购件价格必须为空不是 0，实际 {b2.get('Purchased Price')!r}")
        expect(b2.get("Price Status") == "Unknown",
               f"[案例2] 无价外购件状态必须 Unknown，实际 {b2.get('Price Status')!r}")
        expect(b2.get("Cost Boundary") in (None, ""),
               "[案例2] Unknown 外购件成本边界必须留空（口径未定）")
        # 自制件四列留空
        expect(m1.get("Price Status") in (None, ""),
               "[任务A] 自制件价格状态列必须留空（不适用）")

        # ---- 3-ProcessCost：外购件不得出现（案例 3 前半） ----
        ws3 = wb["3-ProcessCost"]
        p3 = {ws3.cell(row=r, column=2).value
              for r in range(4, ws3.max_row + 1)
              if ws3.cell(row=r, column=2).value}
        expect("TEST-BUY1" not in p3 and "TEST-BUY2" not in p3,
               "[案例3] 外购件不得出现在 3-ProcessCost（外购价已含工艺，再算就是双重计价）")
        expect("TEST-MAKE1" in p3,
               "[案例3] 自制件必须有工艺行（Make 但零成本是漏算）")

        # ---- 4-Summary：完整度块（任务 B） ----
        ws4 = wb["4-Summary"]
        labels = {}
        for r in range(4, ws4.max_row + 1):
            v = ws4.cell(row=r, column=1).value
            if isinstance(v, str):
                labels[v] = r
        for key in ("外购件总数", "有价外购件数（Known/Modeled）", "Unknown 价格外购件数",
                    "Known/Modeled 外购成本小计（元）", "成本数据覆盖率（有价外购件数 ÷ 外购件总数）",
                    "Complete PPC 是否可计算"):
            expect(key in labels, f"[任务B] 4-Summary 完整度块缺行：{key}")

        # 公式语义检查：Unknown 计数公式必须引用 T 列（价格状态）+ O 列（Make/Buy）
        # 注：$T:$T 是合法整列引用；无冒号的 $T$T 是非法引用，Excel 会拒绝打开文件
        if "Unknown 价格外购件数" in labels:
            r = labels["Unknown 价格外购件数"]
            f = str(ws4.cell(row=r, column=2).value)
            expect("$T:$T" in f and "$O:$O" in f,
                   f"[任务B] Unknown 计数必须按 Make/Buy（O）+ 价格状态（T）判定，实际公式：{f}")
        if "Complete PPC 是否可计算" in labels:
            r = labels["Complete PPC 是否可计算"]
            f = str(ws4.cell(row=r, column=2).value)
            expect("部分（仅自制件）" in f and "完整" in f,
                   f"[任务B] Complete PPC 判定必须区分完整/部分两种输出，实际公式：{f}")

        # ---- 5-Opportunities：归属列 + 防重复 + HG（任务 C） ----
        ws5 = wb["5-Opportunities"]
        hdr5 = [ws5.cell(row=3, column=c).value for c in range(1, 25)]
        for col in ("MAT Impact", "VA Impact", "Tooling Impact", "降本机制", "防重复计数组", "HG"):
            expect(col in hdr5, f"[任务C] 5-Opportunities 缺列：{col}")
        # 每条机会归属必填（金额可空）
        n_opp = 0
        for r in range(4, ws5.max_row + 1):
            oid = ws5.cell(row=r, column=1).value
            if not (isinstance(oid, str) and oid.startswith("O")):
                break
            n_opp += 1
            mat = ws5.cell(row=r, column=hdr5.index("MAT Impact") + 1).value
            va = ws5.cell(row=r, column=hdr5.index("VA Impact") + 1).value
            expect(mat not in (None, "") or va not in (None, ""),
                   f"[任务C] {oid} 的 MAT/VA 归属至少一列必填")
            expect(not (mat not in (None, "", "无") and va not in (None, "", "无")),
                   f"[任务C] {oid} 同一机制不得同时计入 MAT 与 VA（主归属二选一）")
            # 归属标记必须与列名一致：MAT Impact 列只接受 MAT/无/空，VA Impact 列只接受 VA/无/空
            expect(mat in (None, "", "无", "MAT"),
                   f"[任务C] {oid} MAT Impact 列出现非法标记 {mat!r}（VA 机制的主归属应写在 VA Impact 列）")
            expect(va in (None, "", "无", "VA"),
                   f"[任务C] {oid} VA Impact 列出现非法标记 {va!r}（MAT 机制的主归属应写在 MAT Impact 列）")
            hg = ws5.cell(row=r, column=hdr5.index("HG") + 1).value
            expect(isinstance(hg, str) and hg.startswith("HG") and hg[2:].isdigit() and 0 <= int(hg[2:]) <= 5,
                   f"[HG] {oid} 的 HG 必须是 HG0–HG5，实际 {hg!r}")
        expect(n_opp >= 1, "[任务C] fixture 产品至少应有一条降本机会")

        # ---- 6-Sensitivity：工装 CapEx 引用 W 列（列位移连锁修正） ----
        ws6 = wb["6-Sensitivity"]
        found_w = False
        for r in range(1, ws6.max_row + 1):
            for c in range(1, 8):
                v = ws6.cell(row=r, column=c).value
                if isinstance(v, str) and "'2-BOM'!W" in v:
                    found_w = True
        expect(found_w, "[列位移] 6-Sensitivity 工装阶梯必须引用 2-BOM!W 列（Tooling CapEx 新位置）")
        for ws_name in ("4-Summary",):
            ws = wb[ws_name]
            for r in range(1, ws.max_row + 1):
                for c in range(1, 4):
                    v = ws.cell(row=r, column=c).value
                    if isinstance(v, str) and "'2-BOM'!S4" in v:
                        failures.append(f"[列位移] {ws_name} 仍引用旧 S 列（现为 Price Type）：{v[:60]}")
        checks += 1

    # ---------- 汇总 ----------
    print(f"\n{'=' * 60}")
    print(f"should-cost 规则不变量：{checks - len(failures)}/{checks} 通过")
    if failures:
        print(f"\n失败 {len(failures)} 条：")
        for f in failures:
            print(f"  FAIL {f}")
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
