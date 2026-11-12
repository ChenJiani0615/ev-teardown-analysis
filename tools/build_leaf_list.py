#!/usr/bin/env python3
"""从 Teardown Data 宽表抽出某个总成的叶件清单（CSV），可叠加 Word 记录里的补充字段。

用法:
    python tools/build_leaf_list.py \
        --workbook data/ex90-rear-inverter/VolvoEX90Ultra2025_TeardownData_Overview.xlsx \
        --top OBCM --root "OBCM Assy" \
        --overrides data/ex90-ultra-obcm/derived/obcm-records.csv \
        --out data/ex90-ultra-obcm/derived/obcm-leaves.csv

说明:
    * 叶件 = 不被任何其他行的 Level 路径所包含的行 → 从结构上避免父子重复计重；
    * 保留两类叶件：① 有总重的计量件；② Part Type == Standard 的紧固/密封件（其重量
      已含在父件 Total Weight 内，输出里 Total Weight 留空，只计成本不计重量）；
    * 既无重量又非标准件的叶节点会被跳过并在 stderr 列名单，不静默丢弃；
    * --overrides 里的字段（如 Manufacturer / 长宽高）会覆盖宽表值。
"""

from __future__ import annotations

import argparse
import csv
import re
from pathlib import Path

import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from teardown_xlsx import load_rows  # noqa: E402

OUT_FIELDS = [
    "CS Part", "Name", "Quantity", "Part Weight", "Fastener Weight", "Total Weight (kg)",
    "Material Group", "Material Species", "Material Species (Tested)", "Manufacturer",
    "Marking Number", "Length", "Width", "Height", "Thickness (T)", "Part Type",
    "Number of Interface", "Depth",
]

NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")


def num(value: object) -> float | None:
    if value is None:
        return None
    m = NUM_RE.search(str(value))
    return float(m.group()) if m else None


def load_overrides(path: Path) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            part = (row.get("CS Part") or "").strip()
            if part:
                out[part] = {k: (v or "").strip() for k, v in row.items() if k}
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="宽表 → 叶件清单 CSV")
    ap.add_argument("--workbook", type=Path, required=True)
    ap.add_argument("--top", required=True, help="总成名，对应 Level 1")
    ap.add_argument("--root", default=None, help="可再限定在某个子总成（如 'OBCM Assy'）")
    ap.add_argument("--overrides", type=Path, default=None)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()

    _, rows, _ = load_rows(args.workbook)
    sub = [r for r in rows if str(r["levels"][1] or "") == args.top]
    if args.root:
        sub = [r for r in sub if args.root in [str(x) for x in r["levels"] if x]]
    if not sub:
        raise SystemExit(f"宽表里没有 {args.top} / {args.root}")

    for r in sub:
        r["path"] = tuple(str(x) for x in r["levels"] if x)
    prefix: set[tuple] = set()
    for r in sub:
        p = r["path"]
        for k in range(1, len(p)):
            prefix.add(p[:k])

    # v3 修复（2026-11-04）：原条件 `and r["total_weight"]` 会把所有 Total Weight 为空的行
    # 静默丢弃——紧固件 / O 圈 / 垫片这类标准件的 Total Weight 常为空（重量已含在父件内），
    # 结果 EX90 DC-DC 少算 19 个挂靠位置 / 14 个料号 / 77 件实物（Level 1 少算 9.12 元/台）。
    # 新规则：叶件 = 结构叶节点；若有重量则计量，若是标准件（Part Type == Standard）则保留
    # 但标 Mass Accounting=含在父件，交由后续步骤决定是否计入质量基线。
    leaves = [
        r for r in sub
        if r["path"] not in prefix
        and (r["total_weight"] or (r.get("part_type") or "").strip().lower() == "standard")
    ]
    dropped = [
        r for r in sub
        if r["path"] not in prefix
        and not r["total_weight"]
        and (r.get("part_type") or "").strip().lower() != "standard"
    ]
    if dropped:
        import sys as _sys
        print(
            f"⚠ 有 {len(dropped)} 个叶节点既无重量也非标准件，已跳过（请核对）：",
            file=_sys.stderr,
        )
        for r in dropped[:10]:
            print(f"    {r.get('cs_part')} / {r.get('name')}", file=_sys.stderr)
    overrides = load_overrides(args.overrides) if args.overrides else {}

    args.out.parent.mkdir(parents=True, exist_ok=True)
    hit = 0
    with args.out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=OUT_FIELDS)
        w.writeheader()
        for r in leaves:
            part = str(r["cs_part"]).strip()
            ov = overrides.get(part, {})
            if ov:
                hit += 1
            rec = {
                "CS Part": part,
                "Name": r["name"],
                "Quantity": r["qty"],
                "Part Weight": r["part_weight"],
                "Fastener Weight": r["fastener_weight"],
                "Total Weight (kg)": r["total_weight"],
                "Material Group": r["material_group"] or "Mixed Material",
                "Material Species": r["material_species"],
                "Material Species (Tested)": r["material_tested"],
                "Manufacturer": ov.get("Manufacturer") or r["manufacturer"],
                "Marking Number": ov.get("Marking Number") or r["marking"],
                "Length": num(ov.get("Length")) if ov.get("Length") else r["length"],
                "Width": num(ov.get("Width")) if ov.get("Width") else r["width"],
                "Height": num(ov.get("Height")) if ov.get("Height") else r["height"],
                "Thickness (T)": num(ov.get("Thickness (T)")) if ov.get("Thickness (T)") else r["thickness"],
                "Part Type": r["part_type"],
                "Number of Interface": r["interfaces"],
                "Depth": r["depth"],
            }
            w.writerow(rec)

    # 标准件行输出时 Total Weight 留空（重量已含在父件内），因此求和只统计有重量的行
    weighed = [r for r in leaves if r["total_weight"]]
    total = sum(float(r["total_weight"]) for r in weighed)
    std = len(leaves) - len(weighed)
    print(
        f"{args.top} / {args.root or '-'}：叶件 {len(leaves)} 个"
        f"（计量 {len(weighed)} 个 / 标准件 {std} 个），计量合计 {total:.3f} kg"
    )
    print(f"叠加 Word 记录 {hit} 个 → {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
