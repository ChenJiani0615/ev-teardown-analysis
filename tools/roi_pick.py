"""ROI 圈选器（人工圈选，非检测程序）。

用途：让你在本地原图上自己拖框圈定 Diamond Grid 区域，输出原生像素坐标。
不做任何纹理识别、不做周期分析、不自动判定。

用法：
    py -3 -X utf8 tools/roi_pick.py "photos/ex90-rear-inverter/Rear Inverter/Rear Inverter/18_Plate2,RearInverterCooling_Front.jpg"

操作：
    鼠标左键拖拽 = 画框
    右键 / Esc   = 取消当前框
    S            = 保存（写入 evidence_check/roi_picks.csv，并导出该框的原生 1:1 裁切）
    R            = 清空重画
    A / D        = 上一张 / 下一张（同一目录内）
    Q            = 退出

输出：
    evidence_check/roi_picks.csv        追加一行：文件名, x0, y0, x1, y1, 宽, 高, 备注
    evidence_check/roi_<文件名>_<x0>_<y0>.png   原生像素、不缩放不增强的裁切

注意：显示窗口是缩放后的预览；保存的是**原生像素坐标**，不是屏幕坐标。
"""
import csv
import sys
import tkinter as tk
from pathlib import Path

from PIL import Image, ImageTk

ROOT = Path(__file__).resolve().parent.parent
OUTDIR = ROOT / "evidence_check"
OUTDIR.mkdir(exist_ok=True)
PICK_CSV = OUTDIR / "roi_picks.csv"


class Picker:
    def __init__(self, root, paths, idx=0):
        self.root = root
        self.paths = paths
        self.idx = idx
        self.box = None
        self.start = None
        self.canvas_w, self.canvas_h = 1400, 860
        self.load()

        self.canvas = tk.Canvas(root, width=self.canvas_w, height=self.canvas_h, bg="#111")
        self.canvas.pack(side="top", fill="both", expand=True)
        self.info = tk.Label(root, anchor="w", justify="left", font=("Consolas", 10),
                             bg="#222", fg="#0f0")
        self.info.pack(side="bottom", fill="x")

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<ButtonPress-3>", lambda e: self.reset())
        root.bind("<Key>", self.on_key)

    def load(self):
        self.path = self.paths[self.idx]
        im = Image.open(self.path)
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        im = im.convert("RGB")
        self.orig = im
        self.W, self.H = im.size
        self.scale = min(self.canvas_w / self.W, self.canvas_h / self.H)
        self.thumb = im.resize((max(1, int(self.W * self.scale)), max(1, int(self.H * self.scale))),
                               Image.LANCZOS)
        self.tkimg = ImageTk.PhotoImage(self.thumb)
        self.box = None
        self.start = None
        self.update_label()

    def update_label(self):
        msg = ("[%d/%d] %s  原生 %dx%d  显示比例 %.3f\n"
               "拖拽画框 | 右键/Esc 取消 | S 保存 | R 重画 | A/D 换图 | Q 退出"
               % (self.idx + 1, len(self.paths), self.path.name, self.W, self.H, self.scale))
        if self.box:
            x0, y0, x1, y1 = self.box
            msg += "\n当前框  原生 x%d..%d  y%d..%d   (宽 %d 高 %d)" % (x0, x1, y0, y1,
                                                                     x1 - x0, y1 - y0)
        self.info.config(text=msg)

    def redraw(self):
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tkimg)
        if self.box:
            x0, y0, x1, y1 = self.box
            cx0, cx1 = x0 * self.scale, x1 * self.scale
            cy0, cy1 = y0 * self.scale, y1 * self.scale
            self.canvas.create_rectangle(cx0, cy0, cx1, cy1, outline="#00ff66", width=2)
        self.update_label()

    def to_native(self, ex, ey):
        x = max(0, min(self.W, int(ex / self.scale)))
        y = max(0, min(self.H, int(ey / self.scale)))
        return x, y

    def on_press(self, e):
        self.start = self.to_native(e.x, e.y)
        self.box = None
        self.redraw()

    def on_drag(self, e):
        if not self.start:
            return
        x, y = self.to_native(e.x, e.y)
        x0, x1 = sorted((self.start[0], x))
        y0, y1 = sorted((self.start[1], y))
        self.box = (x0, y0, x1, y1)
        self.redraw()

    def on_release(self, e):
        if self.box and (self.box[2] - self.box[0] < 10 or self.box[3] - self.box[1] < 10):
            self.box = None
        self.redraw()

    def reset(self):
        self.box = None
        self.start = None
        self.redraw()

    def save(self):
        if not self.box:
            self.info.config(text="没有框，先拖拽画一个")
            return
        x0, y0, x1, y1 = self.box
        crop = self.orig.crop((x0, y0, x1, y1))
        op = OUTDIR / ("roi_%s_%d_%d.png" % (self.path.stem.replace(",", "_")[:40], x0, y0))
        crop.save(op)
        new = not PICK_CSV.exists()
        with open(PICK_CSV, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["照片文件", "x0", "y0", "x1", "y1", "宽", "高", "备注"])
            w.writerow([self.path.name, x0, y0, x1, y1, x1 - x0, y1 - y0, ""])
        self.info.config(text="已保存：%s\n并追加到 %s" % (op.name, PICK_CSV.name))

    def on_key(self, e):
        k = e.keysym.lower()
        if k in ("q", "escape"):
            if k == "escape":
                self.reset()
            else:
                self.root.destroy()
        elif k == "s":
            self.save()
        elif k == "r":
            self.reset()
        elif k in ("a", "d"):
            if k == "a" and self.idx > 0:
                self.idx -= 1
            elif k == "d" and self.idx < len(self.paths) - 1:
                self.idx += 1
            self.load()
            self.redraw()


def main():
    if len(sys.argv) > 1:
        first = Path(sys.argv[1])
        if not first.exists():
            print("找不到文件:", first)
            return
        d = first.parent
        paths = sorted([p for p in d.iterdir() if p.suffix.upper() in (".JPG", ".JPEG", ".PNG")])
        try:
            idx = paths.index(first)
        except ValueError:
            idx = 0
    else:
        # 默认只列冷却板相关照片，避免在 114 张里翻
        d = ROOT / "photos" / "ex90-rear-inverter" / "Rear Inverter" / "Rear Inverter"
        paths = sorted([p for p in d.iterdir()
                        if p.suffix.upper() in (".JPG", ".JPEG", ".PNG")
                        and any(k in p.name for k in ("Cooling", "Locator", "Spacer", "Holder", "MOSFET", "Gasket"))])
        idx = 0

    if not paths:
        print("目录内没有图片")
        return
    print("候选 %d 张；A/D 切换；S 保存" % len(paths))
    for i, p in enumerate(paths):
        print("  %2d  %s" % (i, p.name))

    root = tk.Tk()
    root.title("ROI 圈选器 —— 只圈选，不判定")
    Picker(root, paths, idx)
    root.mainloop()


if __name__ == "__main__":
    main()