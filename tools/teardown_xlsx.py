#!/usr/bin/env python3
"""读取 Cares/拆解数据库导出的 Teardown Data 宽表。

背景：这类导出文件里带 `<autoFilter ref="3:3"/>` 这种非法引用，
openpyxl 直接读会抛 ValueError（filters.py）。这里先复制一份、
去掉 autoFilter 节点，再交给 openpyxl，原始文件不动。

用法（作为库）:
    from teardown_xlsx import load_rows, COLS
    header, rows, idx = load_rows(path)

用法（命令行，快速看结构）:
    python tools/teardown_xlsx.py <xlsx> [--grep 关键词]
"""

from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

SHEET = "Teardown Data"
HEADER_ROW = 3

# 列字母 → 规范字段名。字母来自该导出格式的固定布局，加载时会校验表头文字。
COLS: dict[str, str] = {
    "part_id": "B",
    "depth": "C",
    "component_id": "D",
    "name": "E",
    "group": "F",
    "category": "P",
    "marking": "X",
    "material_group": "Y",
    "material_species": "Z",
    "material_tested": "AA",
    "total_weight_cd": "CD",
    "interfaces_type": "CG",
    "interfaces": "DM",
    "cs_part": "EX",
    "fastener_weight": "FA",
    "length": "IN",
    "manufacturer": "IO",
    "part_weight": "JD",
    "part_type": "JC",
    "qty": "JG",
    "total_weight": "JO",
    "width": "JR",
    "height": "JT",
    "thickness": "JN",
}

LEVELS = [chr(ord("G") + i) for i in range(8)]  # G..N = Level 0..7


def _sanitize(src: Path) -> Path:
    """去掉 autoFilter（含非法 ref），返回临时副本路径。"""
    tmp = Path(tempfile.gettempdir()) / f"_clean_{src.stem}_{abs(hash(src)) % 65536}.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and b"autoFilter" in data:
                data = re.sub(rb"<(?:\w+:)?autoFilter[^>]*/>", b"", data)
            zout.writestr(item, data)
    return tmp


def load_rows(path: Path | str) -> tuple[list[str], list[dict[str, object]], dict[str, int]]:
    """返回 (表头, 数据行, 列名→0 基索引)。空行丢弃。"""
    src = Path(path)
    clean = _sanitize(src)
    try:
        wb = load_workbook(clean, read_only=True, data_only=True)
        if SHEET not in wb.sheetnames:
            raise SystemExit(f"{src.name} 里没有工作表 '{SHEET}'，实际有：{wb.sheetnames}")
        ws = wb[SHEET]
        header = [str(c) if c is not None else "" for c in
                  next(ws.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True))]
        idx = {name: column_index_from_string(letter) - 1 for name, letter in COLS.items()}
        rows: list[dict[str, object]] = []
        for raw in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
            if not any(v not in (None, "") for v in raw):
                continue
            rec = {k: (raw[i] if i < len(raw) else None) for k, i in idx.items()}
            rec["levels"] = [raw[column_index_from_string(c) - 1] for c in LEVELS]
            rows.append(rec)
        wb.close()
    finally:
        clean.unlink(missing_ok=True)

    for name, letter in COLS.items():
        i = column_index_from_string(letter) - 1
        found = header[i] if i < len(header) else ""
        expected = {
            "cs_part": "CS Part", "part_weight": "Part Weight (kg)",
            "fastener_weight": "Fastener Weight (kg)", "total_weight": "Total Weight (kg)",
            "qty": "Quantity", "interfaces": "Number of Interface",
        }.get(name)
        if expected and found and found != expected:
            print(f"[warn] 列 {letter} 表头为 '{found}'，预期 '{expected}'（{name}）")
    return header, rows, idx


def subtree(rows: list[dict[str, object]], keyword: str) -> list[dict[str, object]]:
    kw = keyword.lower()
    out = []
    for r in rows:
        path = [str(x) for x in r["levels"] if x]
        if any(kw in p.lower() for p in path) or kw in str(r["name"] or "").lower():
            out.append(r)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="查看 Teardown Data 宽表")
    parser.add_argument("path", type=Path)
    parser.add_argument("--grep", default=None, help="只看路径或名称包含该关键词的行")
    parser.add_argument("--limit", type=int, default=40)
    args = parser.parse_args()

    _, rows, _ = load_rows(args.path)
    print(f"总行数 {len(rows)}")
    if args.grep:
        rows = subtree(rows, args.grep)
        print(f"命中 {len(rows)} 行")
    for r in rows[: args.limit]:
        path = " > ".join(str(x) for x in r["levels"] if x)
        print(f"d{r['depth']} | {r['name']} | {r['cs_part']} | qty={r['qty']} "
              f"| tot={r['total_weight']} | {r['part_type']}")
        print(f"    {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
