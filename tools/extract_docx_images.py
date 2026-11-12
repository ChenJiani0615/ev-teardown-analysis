#!/usr/bin/env python3
"""把 .docx 里的图片连同上下文文字一起导出，便于逐张读图。

用法:
    python tools/extract_docx_images.py "data/ex90-ultra-obcm/Volvo EX90 Ultra 2025.docx"
    python tools/extract_docx_images.py <docx> --out data/ex90-ultra-obcm/extracted

输出:
    <out>/<image>           所有嵌入图片（按文档顺序重命名：p001_<原名>）
    <out>/index.md          序号 / 文件名 / 周边文字 / 尺寸像素
"""

from __future__ import annotations

import argparse
import io
import json
import re
import zipfile
from pathlib import Path


def png_size(data: bytes) -> tuple[int, int] | None:
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w = int.from_bytes(data[16:20], "big")
        h = int.from_bytes(data[20:24], "big")
        return w, h
    if data[:2] == b"\xff\xd8":
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                h = int.from_bytes(data[i + 5:i + 7], "big")
                w = int.from_bytes(data[i + 7:i + 9], "big")
                return w, h
            seg = int.from_bytes(data[i + 2:i + 4], "big")
            i += 2 + seg
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="导出 docx 内嵌图片及其上下文")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--out", type=Path, default=None)
    args = ap.parse_args()

    src = args.docx
    out = args.out or src.parent / "extracted"
    out.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src) as z:
        document = z.read("word/document.xml").decode("utf-8", "replace")
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8", "replace")
        rel_map = {
            rid: target.split("/")[-1]
            for rid, target in re.findall(
                r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels
            )
        }
        media = {n.split("/")[-1]: z.read(n) for n in z.namelist() if n.startswith("word/media/")}

    paragraphs = re.findall(r"<w:p[ >].*?</w:p>", document, re.S)
    entries: list[dict] = []
    for idx, para in enumerate(paragraphs, start=1):
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", para, re.S)).strip()
        text = re.sub(r"<[^>]+>", "", text)
        for rid in re.findall(r'r:embed="(rId\d+)"', para):
            name = rel_map.get(rid)
            if name and name in media:
                entries.append({"para": idx, "text": text, "file": name})

    lines = ["# 内嵌图片索引", "", f"来源：`{src.name}`　共 {len(entries)} 张", "",
             "| # | 文件 | 文档段 | 周边文字 | 像素 |", "|---|---|---|---|---|"]
    for i, e in enumerate(entries, start=1):
        data = media[e["file"]]
        stem, dot, ext = e["file"].rpartition(".")
        new_name = f"p{i:03d}_{stem}.{ext}" if dot else f"p{i:03d}_{e['file']}"
        (out / new_name).write_bytes(data)
        size = png_size(data)
        dims = f"{size[0]}×{size[1]}" if size else "?"
        e["out"] = new_name
        lines.append(
            f"| {i} | {new_name} | {e['para']} | {e['text'][:60] or '—'} | {dims} |"
        )
        print(f"{i:03d} {new_name:32s} {dims:12s} {e['text'][:50]}")

    (out / "index.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    (out / "index.json").write_text(
        json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n共 {len(entries)} 张，已导出到 {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
