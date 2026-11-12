#!/usr/bin/env python3
"""局部裁切放大：把拆解照片的关键区域切开放大，供结构判读。

为什么需要：全景照片里 MOSFET 只有几十像素，无法判断它是否两面都有
热界面（这是 T 轴 单面/双面 冷却的唯一直接证据）。裁切 + 双三次放大后
才可能分辨"器件两面是否都有压痕/残胶"。

用法：
    py -3 tools/crop_zoom.py <image> <out> --box x0,y0,x1,y1 --scale 4
    py -3 tools/crop_zoom.py <image> <out> --grid 3      # 3x3 瓦片
box 用**相对比例**（0–1），不依赖原始像素尺寸。
"""
from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image


def crop_rel(im: Image.Image, box: str) -> Image.Image:
    x0, y0, x1, y1 = (float(v) for v in box.split(","))
    w, h = im.size
    return im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h)))


def tiles(im: Image.Image, n: int) -> list[Image.Image]:
    w, h = im.size
    out = []
    for r in range(n):
        for c in range(n):
            out.append(im.crop((
                int(c * w / n), int(r * h / n),
                int((c + 1) * w / n), int((r + 1) * h / n),
            )))
    return out


def upscale(im: Image.Image, scale: float) -> Image.Image:
    w, h = im.size
    return im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("out")
    ap.add_argument("--box", default=None, help="相对比例 x0,y0,x1,y1")
    ap.add_argument("--grid", type=int, default=0, help="切成 n×n 瓦片并拼成对照图")
    ap.add_argument("--scale", type=float, default=3.0)
    ap.add_argument("--max", type=int, default=1400, help="输出最长边上限")
    a = ap.parse_args()

    src = Path(a.image)
    im = Image.open(src).convert("RGB")

    if a.grid:
        parts = [upscale(t, a.scale) for t in tiles(im, a.grid)]
        cell = max(p.size[0] for p in parts), max(p.size[1] for p in parts)
        sheet = Image.new("RGB", (cell[0] * a.grid, cell[1] * a.grid), (255, 255, 255))
        for i, p in enumerate(parts):
            sheet.paste(p, ((i % a.grid) * cell[0], (i // a.grid) * cell[1]))
        out_im = sheet
    else:
        piece = crop_rel(im, a.box) if a.box else im
        out_im = upscale(piece, a.scale)

    # 限长边，避免生成过大文件
    w, h = out_im.size
    if max(w, h) > a.max:
        k = a.max / max(w, h)
        out_im = out_im.resize((int(w * k), int(h * k)), Image.LANCZOS)

    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out_im.save(out, quality=92)
    print(f"{out}  {out_im.size[0]}x{out_im.size[1]}  源 {im.size[0]}x{im.size[1]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())