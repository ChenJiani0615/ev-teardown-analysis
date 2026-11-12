#!/usr/bin/env python3
"""把「竖排键值表」的 Word 记录导出成 CSV（CS 数据库单件记录常见格式）。

一个表格里可能连续排了多个零件：每遇到 `CS Part` 就开一条新记录。

用法:
    python tools/parse_cs_records.py "data/ex90-ultra-obcm/OBCM.docx" \
        --out data/ex90-ultra-obcm/derived/obcm-records.csv
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

from docx import Document


def norm(text: str) -> str:
    return " ".join(text.split())


KNOWN_KEYS = {
    "CS Part", "Quantity", "Part Weight", "Assembly Weight", "Fastener Weight",
    "Total Weight (kg)", "Material Group", "Material Species",
    "Material Species (Tested)", "Manufacturer", "Manufacturer Origin",
    "Marking Number", "Length", "Width", "Height", "Thickness (T)",
    "Outer Diameter (OD)", "Part Type", "Number of Index", "Number of Interface",
    "Number of Connection", "Types of Index", "Types of Interface",
    "Commodity Type", "Category", "Subcategory", "Polymer Identification Method",
}


def row_pairs(cells: list[str]) -> list[tuple[str, str]]:
    """把一行还原成 (key, value)。兼容三种排版：
    1) 两列：key | value
    2) 四列：key1 | value1 | key2 | value2
    3) 单元格里自带 '|'（横向合并后文本拼在一起）
    """
    ded: list[str] = []
    for c in cells:
        if not ded or ded[-1] != c:
            ded.append(c)
    pairs: list[tuple[str, str]] = []
    i = 0
    while i < len(ded):
        cell = ded[i]
        if "|" in cell:
            key, _, value = cell.partition("|")
            pairs.append((key.strip(), value.strip()))
            i += 1
        elif cell in KNOWN_KEYS and i + 1 < len(ded):
            pairs.append((cell, ded[i + 1]))
            i += 2
        else:
            i += 1
    return pairs


def parse(path: Path) -> list[dict[str, str]]:
    doc = Document(path)
    records: list[dict[str, str]] = []
    current: dict[str, str] | None = None

    for table in doc.tables:
        for row in table.rows:
            cells = [norm(c.text) for c in row.cells]
            for key, value in row_pairs(cells):
                if key == "CS Part" and value:
                    current = {"CS Part": value}
                    records.append(current)
                elif current is not None and key in KNOWN_KEYS:
                    current.setdefault(key, value)
    return records


def main() -> int:
    ap = argparse.ArgumentParser(description="docx 竖排键值表 → CSV")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()

    records = parse(args.docx)
    if not records:
        raise SystemExit(f"{args.docx} 里没找到以 'CS Part' 开头的记录")

    order: list[str] = []
    for rec in records:
        for k in rec:
            if k not in order:
                order.append(k)
    # 让关键字段排在前面
    front = ["CS Part", "Quantity", "Part Weight", "Assembly Weight", "Fastener Weight",
             "Total Weight (kg)", "Material Group", "Material Species",
             "Material Species (Tested)", "Manufacturer", "Manufacturer Origin",
             "Marking Number", "Length", "Width", "Height", "Thickness (T)",
             "Outer Diameter (OD)", "Part Type", "Number of Index",
             "Number of Interface", "Number of Connection"]
    cols = [c for c in front if c in order] + [c for c in order if c not in front]

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with args.out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for rec in records:
            w.writerow({c: rec.get(c, "") for c in cols})

    print(f"记录 {len(records)} 条，字段 {len(cols)} 个 → {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
