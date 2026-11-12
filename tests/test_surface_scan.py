#!/usr/bin/env python3
"""surface_scan.py 回归测试。

覆盖两类东西：
  A. probe 的判别力——用**与规则编写无关**的合成图，验证
     · 真阳性：已知周期的规则交叉纹理必须被检出为多方向
     · 真阴性：纯平面 / 纯噪声 / 单向走刀痕不得被说成"设计特征"
  B. check 的门禁——未查看、缺照片、缺裁切框、缺特征形态、无对照面
     都必须非零退出
"""

from __future__ import annotations

import csv
import os
import shutil
import subprocess
import sys
import tempfile

import numpy as np
from PIL import Image, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCAN = os.path.join(ROOT, ".github", "skills", "ev-teardown-benchmark", "scripts", "surface_scan.py")

N = 900
Y, X = np.mgrid[0:N, 0:N].astype(np.float64)


def _shade(gray=150.0):
    r = np.hypot(Y - N / 2, X - N / 2) / (N * 0.72)
    return np.full((N, N), gray) * (1.0 - 0.35 * np.clip(r, 0, 1) ** 2) + 12


def _save(path, arr):
    Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8)).save(path)


def make_fixtures(d):
    rng = np.random.default_rng(20261030)
    p = 48.0
    d1 = np.cos(2 * np.pi * (X + Y) / (p * np.sqrt(2)))
    d2 = np.cos(2 * np.pi * (X - Y) / (p * np.sqrt(2)))
    f = {}
    f["diamond"] = os.path.join(d, "diamond.png")
    _save(f["diamond"], _shade(150) + 16 * (d1 + d2) / 2)
    f["striation"] = os.path.join(d, "striation.png")
    _save(f["striation"], _shade(145) + 11 * np.cos(2 * np.pi * X / 18.0))
    f["noise"] = os.path.join(d, "noise.png")
    _save(f["noise"], _shade(152) + rng.normal(0, 6, (N, N)))
    f["smooth"] = os.path.join(d, "smooth.png")
    _save(f["smooth"], _shade(158))
    return f


def probe(photo, box="150,150,750,750", lag_max=120):
    r = subprocess.run(
        [sys.executable, SCAN, "probe", "--photo", photo, "--box", box, "--lag-max", str(lag_max)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    return r.returncode, r.stdout + r.stderr


def _clusters(out):
    for line in out.splitlines():
        if "取向聚簇：" in line:
            return int(line.split("取向聚簇：")[1].split("组")[0])
    return 0


def _detected(out):
    return "咨询结论：检出" in out


FAILS = []


def ck(name, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else "  <<< " + detail))
    if not cond:
        FAILS.append(name)


def test_probe(fx):
    print("[A] probe 判别力（合成图，与规则编写无关）")

    code, out = probe(fx["diamond"])
    ck("真阳性·菱形纹理：检出周期结构", code == 0 and _detected(out))
    ck("真阳性·菱形纹理：判为多方向交叉（>=3 组取向）", _clusters(out) >= 3, str(_clusters(out)))
    ck("真阳性·菱形纹理：未按单向加工痕迹处理", "单向/双向条带" not in out)

    code, out = probe(fx["striation"])
    ck("真阴性·单向刀纹：不得判为多方向设计结构", "多方向交叉纹理" not in out)
    ck("真阴性·单向刀纹：提示按加工痕迹处理", "加工痕迹" in out)

    for k, label in (("noise", "纯噪声"), ("smooth", "纯平滑面")):
        code, out = probe(fx[k])
        ck(f"真阴性·{label}：未检出周期结构", (not _detected(out)) and code == 0)


def _write_ledger(path, rows):
    cols = ["接口面", "面向对象", "照片文件", "裁切框_原生像素", "判读结论",
            "特征形态", "节距_px", "是否周期性", "功能假设", "区分依据_缺什么", "备注"]
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})


def check(ledger):
    r = subprocess.run(
        [sys.executable, SCAN, "check", "--ledger", ledger],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    return r.returncode, r.stdout + r.stderr


def test_gate(tmp):
    print("[B] check 门禁")

    ledger = os.path.join(tmp, "blank.csv")
    subprocess.run([sys.executable, SCAN, "init", "--out", ledger, "--force"],
                   capture_output=True, text=True, encoding="utf-8")
    code, out = check(ledger)
    ck("空台账（九行全 未查看）：非零退出", code != 0)
    ck("空台账：逐行指出 未查看", out.count("仍是 未查看") == 9, str(out.count("仍是 未查看")))

    with open(ledger, encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))

    full = [dict(r) for r in rows]
    for i, r in enumerate(full):
        r["照片文件"] = "a.jpg"
        r["裁切框_原生像素"] = "0,0,100,100"
        r["判读结论"] = "无特征" if i == 0 else "有特征"
        r["特征形态"] = "交叉细纹"
    good = os.path.join(tmp, "good.csv")
    _write_ledger(good, full)
    code, out = check(good)
    ck("合规台账（8 有特征 + 1 无特征 对照）：退出 0", code == 0, out)

    allpos = [dict(r) for r in full]
    for r in allpos:
        r["判读结论"] = "有特征"
    p = os.path.join(tmp, "allpos.csv")
    _write_ledger(p, allpos)
    code, out = check(p)
    ck("全部 有特征、零 无特征：非零退出（无对照面）", code != 0, out)

    noimg = [dict(r) for r in full]
    noimg[3]["照片文件"] = ""
    p = os.path.join(tmp, "noimg.csv")
    _write_ledger(p, noimg)
    code, out = check(p)
    ck("缺照片文件：非零退出", code != 0)

    nobox = [dict(r) for r in full]
    nobox[4]["裁切框_原生像素"] = ""
    p = os.path.join(tmp, "nobox.csv")
    _write_ledger(p, nobox)
    code, out = check(p)
    ck("缺裁切框：非零退出", code != 0)

    noshape = [dict(r) for r in full]
    noshape[5]["特征形态"] = ""
    p = os.path.join(tmp, "noshape.csv")
    _write_ledger(p, noshape)
    code, out = check(p)
    ck("有特征但缺特征形态：非零退出", code != 0)

    miss = [dict(r) for r in full if r["接口面"] != "TIM 界面"]
    p = os.path.join(tmp, "miss.csv")
    _write_ledger(p, miss)
    code, out = check(p)
    ck("缺接口面行：非零退出", code != 0)


def main():
    tmp = tempfile.mkdtemp(prefix="surfscan_")
    try:
        fx = make_fixtures(tmp)
        test_probe(fx)
        test_gate(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if FAILS:
        print(f"\n失败 {len(FAILS)} 项：" + "、".join(FAILS))
        return 1
    print("\n全部通过（probe 4 类合成图 + check 7 项门禁）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())