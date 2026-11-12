#!/usr/bin/env python3
"""导出 .docx 里的段落与表格（有些"尺寸数据"Word 里其实是表格，不是图片）。

用法:
    python tools/extract_docx_tables.py "data/ex90-ultra-obcm/OBCM.docx"
    python tools/extract_docx_tables.py <docx> --out data/ex90-ultra-obcm/derived/obcm-tables.md

输出:
    --out 指定的 md（缺省 <docx 同目录>/derived/<stem>-tables.md）
    包含每个表格前的最近标题、表头与全部行，可直接粘给分析用
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(doc: Document):
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def cell_text(cell) -> str:
    return " ".join(cell.text.split())


def main() -> int:
    ap = argparse.ArgumentParser(description="导出 docx 段落与表格")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--out", type=Path, default=None)
    args = ap.parse_args()

    doc = Document(args.docx)
    out = args.out or args.docx.parent / "derived" / f"{args.docx.stem}-tables.md"
    out.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = [f"# {args.docx.name} 抽取结果", ""]
    heading = ""
    n_table = 0
    n_rows = 0
    paragraphs: list[str] = []

    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            text = " ".join(block.text.split())
            if not text:
                continue
            if block.style.name.lower().startswith("heading") or len(text) < 40:
                heading = text
                if block.style.name.lower().startswith("heading"):
                    lines += ["", f"## {text}", ""]
            else:
                paragraphs.append(text)
        else:
            n_table += 1
            rows = [[cell_text(c) for c in r.cells] for r in block.rows]
            if not rows:
                continue
            width = max(len(r) for r in rows)
            rows = [r + [""] * (width - len(r)) for r in rows]
            n_rows += len(rows)
            lines.append(f"### 表 {n_table}｜上文：{heading or '—'}｜{len(rows)} 行 × {width} 列")
            lines.append("")
            lines.append("| " + " | ".join(rows[0]) + " |")
            lines.append("|" + "---|" * width)
            for r in rows[1:]:
                lines.append("| " + " | ".join(x.replace("|", "\\|") for x in r) + " |")
            lines.append("")

    if paragraphs:
        lines += ["## 正文段落", ""] + [f"- {p}" for p in paragraphs]

    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"段落 {len(paragraphs)} 段（长段），表格 {n_table} 个，合计 {n_rows} 行")
    print(f"已写入 {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
