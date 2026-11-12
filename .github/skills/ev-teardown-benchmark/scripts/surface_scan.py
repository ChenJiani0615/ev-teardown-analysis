#!/usr/bin/env python3
"""表面特征发现（Surface Feature Discovery）执行器。

设计目的（针对实际发生过的漏检）：原流程只在"照片里出现无法用装饰/加工痕迹解释
的特征"时才调用判读矩阵，也就是**必须先注意到特征，规则才生效**。但"注意到"这一步
没有任何强制动作、没有留痕、也不允许写"未查看"，结果是：
  - 没人去看 → 矩阵永远不会被调用；
  - 看了但没注意 → 同样不会调用；
  - 报告里无法区分"确实没有特征"与"根本没看"。

本脚本把"注意到"变成一次**必须先于四表执行、必须留痕、允许并强制记录否定结论**的
独立步骤。它只回答"看了没有 / 看到了什么形态 / 有没有周期性"，不预设任何具体纹理
名称（不写死"菱形"这类关键词规则）。

三个子命令
----------
1. ``init``   生成空白台账（每个接口面一行，verdict 默认 `未查看`）。
2. ``probe``  对一张照片的一个裁切区域给出**定量周期性咨询**（去 JPEG 块效应后），
              供人工判读时交叉核对，不做自动定性。
3. ``check``  校验台账是否允许进入出表阶段；不合格直接非零退出。

用法::

    py -3 surface_scan.py init  --out ledger.csv
    py -3 surface_scan.py probe --photo a.jpg --box 2280,1436,2792,1948
    py -3 surface_scan.py check --ledger ledger.csv

只依赖标准库；``probe`` 需要 Pillow 与 numpy（缺失时降级为"未执行定量核查"）。
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from pathlib import Path

# 九类接口面：不写死纹理名称，只写"该去哪里看"。每个面都必须有一行。
SURFACE_CLASSES = [
    ("功率级-器件面", "裸片/模块/基板朝向冷却器的那个面"),
    ("功率级-冷板接触面", "冷板上与器件/TIM 贴合的那一面"),
    ("冷板-流道面", "冷板上面向冷却液的腔/槽/凸包面"),
    ("冷板-背面/结构面", "背向热源、与壳体或支架接触的面"),
    ("TIM 界面", "预成型条、膏、相变材料的可见残留与厚度线索"),
    ("密封面", "垫片/胶槽/止口所在面"),
    ("母排/端子面", "铜排表面、焊接或螺接区域"),
    ("壳体对合面与紧固点", "上/下盖板对合面、螺钉柱、压紧件落座面"),
    ("塑料件-嵌件/包胶面", "支架、连接器、绝缘件的嵌件与包胶面"),
]

VERDICTS = ["有特征", "无特征", "未查看"]
COLUMNS = [
    "接口面",
    "面向对象",
    "照片文件",
    "裁切框_原生像素",
    "判读结论",
    "特征形态",
    "节距_px",
    "是否周期性",
    "功能假设",
    "区分依据_缺什么",
    "备注",
]


def _die(msg: str, code: int = 2) -> None:
    print(f"[surface-scan] {msg}", file=sys.stderr)
    raise SystemExit(code)


def cmd_init(args) -> int:
    out = Path(args.out)
    if out.exists() and not args.force:
        _die(f"{out} 已存在；如需重建加 --force", 3)
    with out.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(COLUMNS)
        for cls, target in SURFACE_CLASSES:
            w.writerow([cls, target, "", "", "未查看", "", "", "", "", "", ""])
    print(f"[surface-scan] 已生成台账 {out}（{len(SURFACE_CLASSES)} 个接口面）")
    print("[surface-scan] 判读结论只允许：有特征 / 无特征 / 未查看")
    print("[surface-scan] 有特征：必须填 照片文件 + 裁切框_原生像素 + 特征形态")
    print("[surface-scan] 无特征：必须填 照片文件 + 裁切框_原生像素（证明真的看过）")
    return 0


def cmd_probe(args) -> int:
    try:
        import numpy as np
        from PIL import Image, ImageFilter
    except Exception as exc:  # pragma: no cover
        print(f"[surface-scan] 定量核查未执行（缺 Pillow/numpy：{exc}）")
        print("[surface-scan] 这不影响人工判读，但报告里必须说明定量核查未执行。")
        return 0

    photo = Path(args.photo)
    if not photo.exists():
        _die(f"照片不存在：{photo}", 4)
    try:
        x0, y0, x1, y1 = (int(v) for v in args.box.split(","))
    except Exception:
        _die("--box 需要 x0,y0,x1,y1（原生像素整数）", 5)

    im = Image.open(photo).convert("L")
    W, H = im.size
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(W, x1), min(H, y1)
    if x1 - x0 < 64 or y1 - y0 < 64:
        _die(f"裁切过小：{(x0, y0, x1, y1)}，至少 64x64 原生像素", 6)

    crop = im.crop((x0, y0, x1, y1))
    # (1) 去 JPEG DCT 块效应（4/5/8 px 整数倍）：sigma>=2 高斯，先抹掉再找周期性
    g = np.asarray(crop.filter(ImageFilter.GaussianBlur(args.deblock)), dtype=np.float64)
    # (2) 去照明梯度/暗角：大尺度低通估计背景后减掉。否则一个抖动渐晕的方差
    #     足以淹没真实纹理，自相关会被背景吃掉。
    bg = np.asarray(
        crop.filter(ImageFilter.GaussianBlur(args.detrend * max(crop.size))), dtype=np.float64
    )

    lag_min, lag_max = args.lag_min, args.lag_max
    h, w = g.shape
    win = np.outer(np.hanning(h), np.hanning(w))
    quad_stats = []  # 四象限的 (主周期, 取向, 强度)，用于空间平稳性检验

    def _peaks_of(sub, win_sub, lmax, r_min):
        """在 sub × win_sub 上求环均值归一化后的方向性峰。
        主裁切与四象限**必须共用同一口径**，否则象限检验会拿绝对 r 最大值的
        最小滞后（总是 lag_min）冒充"主峰"，造成稳定的假通过。"""
        hh, ww = sub.shape
        ss = sub * win_sub
        FF = np.fft.fft2(ss)
        RR = np.fft.fftshift(
            np.real(np.fft.ifft2(FF * np.conj(FF))) / (np.sum(win_sub ** 2) + 1e-12)
        )
        r0 = RR[hh // 2, ww // 2]
        if r0 <= 0:
            return [], r0
        Rq = RR / r0
        lim_q = min(lmax, hh // 2 - 2, ww // 2 - 2)
        if lim_q < lag_min + 8:
            return [], r0
        ry, rx = hh // 2, ww // 2
        rsum = np.zeros(lim_q + 2)
        rcnt = np.zeros(lim_q + 2)
        for dy in range(-lim_q, lim_q + 1):
            for dx in range(-lim_q, lim_q + 1):
                lg = float(np.hypot(dx, dy))
                if lg < lag_min or lg > lim_q:
                    continue
                k = int(round(lg))
                rsum[k] += Rq[ry + dy, rx + dx]
                rcnt[k] += 1
        rmean = np.where(rcnt >= 8, rsum / np.maximum(rcnt, 1), np.nan)
        valid = rmean[np.isfinite(rmean) & (rmean > 0)]
        floor = float(np.median(valid)) * 0.25 if valid.size else 1e-6
        out = []
        for dy in range(-lim_q, lim_q + 1):
            for dx in range(-lim_q, lim_q + 1):
                lg = float(np.hypot(dx, dy))
                if lg < lag_min or lg > lim_q:
                    continue
                k = int(round(lg))
                base = rmean[k]
                if not np.isfinite(base) or base <= 0:
                    continue
                rv = float(Rq[ry + dy, rx + dx])
                if rv < r_min:
                    continue
                out.append((rv / max(base, floor), rv, lg, float(np.degrees(np.arctan2(dy, dx)))))
        out.sort(key=lambda t: t[0], reverse=True)
        return out, r0

    a = g - bg
    aw = a * win
    # 循环自相关：R = ifft2(|FFT|^2) / sum(win^2)，零滞后在 [0,0]
    F = np.fft.fft2(aw)
    R = np.real(np.fft.ifft2(F * np.conj(F))) / (np.sum(win ** 2) + 1e-12)
    R = np.fft.fftshift(R)
    Rc = R[h // 2, w // 2]
    if Rc <= 0:
        _die("去趋势后裁切内无方差（纯色区域），无法做周期分析", 7)

    # 只在 lag_max 半径内取值。自相关本身随滞后衰减（尤其高斯去块后），
    # 纯噪声也会出现 r≈0.97 的"假峰"，因此必须按**半径环均值归一化**：
    # 各向同性衰减被除掉，只剩方向性周期峰才突出。
    # 内圈样本太少（环上只有几点）比值会爆到几千，因此只统计**环上样本>=8**的滞后。
    # 统计必须在**完整圆**上做：若只取 ±lag_max 的方块，最外圈被方边界截断，
    # 环上样本数骤降，会把接近 lag_max 的比值放大 2~3 倍，造出假的"大尺度周期"。
    max_lag_ok = min(lag_max, h // 2 - 2, w // 2 - 2)
    if max_lag_ok < lag_min + 10:
        _die(
            f"裁切太小：{w}x{h} 只能可靠搜索到 {max_lag_ok}px，而 lag_min={lag_min}；"
            "请放大裁切或调小 --lag-min",
            6,
        )
    peaks, _ = _peaks_of(a, win, max_lag_ok, args.r_min)
    keep = []
    for ratio, r, lag, ang in peaks:
        if all(abs(lag - l2) > max(3.0, 0.12 * l2) for _, _, l2, _ in keep):
            keep.append((ratio, r, lag, ang))
        if len(keep) >= 5:
            break

    # ---- 空间平稳性检验 ----
    # 真实的表面纹理是**空间平稳**的：同一个节距在整个裁切内都应可测。
    # 照明梯度、镜头渐晕、以及大尺度走刀方向造成的 1/f 型方向性偏置不是平稳的，
    # 它的"峰"会随搜索上限一起外移（把 --lag-max 从 240 拉到 480，峰跟着跑到 340）。
    # 因此对四个象限分别求主峰，比对周期一致性，作为"是否周期结构"的必要条件。
    # 注意：象限主峰必须用**与主裁切相同的归一化口径**，否则会退化成
    # "绝对 r 最大值 = 最小滞后"，四个象限都报 lag_min，离散度恒为 0%，瞬为假通过。
    for name_q, sub_q in (
        ("TL", a[: h // 2, : w // 2]),
        ("TR", a[: h // 2, w // 2 :]),
        ("BL", a[h // 2 :, : w // 2]),
        ("BR", a[h // 2 :, w // 2 :]),
    ):
        hh, ww = sub_q.shape
        qpeaks, _ = _peaks_of(sub_q, np.outer(np.hanning(hh), np.hanning(ww)), lag_max, args.r_min)
        if qpeaks:
            quad_stats.append((name_q, qpeaks[0][2], qpeaks[0][3], qpeaks[0][0]))

    print(f"[surface-scan] 照片 {photo.name}  {W}x{H}")
    print(f"[surface-scan] 裁切 {x0},{y0},{x1},{y1}  ({x1 - x0}x{y1 - y0} 原生像素)")
    print(
        f"[surface-scan] 去块 sigma={args.deblock}，去照明梯度 sigma={args.detrend * max(h, w):.0f}px，"
        f"搜索周期 {lag_min}-{lag_max} px"
    )
    print("[surface-scan] 前 5 个方向性周期峰（强度=环内峰/环均值，>=1.6 才算各向异性显著）：")
    for ratio, r, lag, ang in keep:
        print(f"    强度={ratio:.2f}  r={r:.3f}  周期={lag:.0f}px  取向={ang:+.0f}deg")
    # 峰是否贴住搜索上限：真实纹理的周期由实物决定，不会随 --lag-max 变化；
    # 1/f 型方向性偏置的峰永远贴在上限上。这是最便宜也最有效的伪影自诊断。
    edge_hit = any(lag > 0.92 * max_lag_ok for _, _, lag, _ in keep[:3])
    if edge_hit:
        print(
            f"[surface-scan] ⚠ 首峰贴近搜索上限（{max_lag_ok}px 的 92%）。请把 --lag-max 加倍重跑："
            "若首峰的周期随之外移、强度反而升高，则该峰是**大尺度方向性梯度/照明偏置**，"
            "不是表面周期纹理，不得计入特征。"
        )
    stationarity = None  # None=无信息量；True=通过；False=不通过
    if quad_stats:
        lags = [q[1] for q in quad_stats]
        med = float(np.median(lags))
        detail = "  ".join(f"{n}:{lg:.0f}px@{ang:+.0f}deg/{q:.2f}" for n, lg, ang, q in quad_stats)
        print(f"[surface-scan] 象限主峰：{detail}")
        # 若象限里多数峰本身不显著（强度 < ratio_min），平稳性检验无信息量，
        # 不得据此得出「一致」或「不一致」的结论。
        sig_q = [q for q in quad_stats if q[3] >= args.ratio_min]
        if len(sig_q) < 3:
            print(
                f"[surface-scan] 象限主峰仅 {len(sig_q)}/4 达到强度阈值，"
                "平稳性检验无信息量（不得当作「纹理存在」或「纹理不存在」的证据）。"
            )
        else:
            # 用**众数聚簇**而不是最大-最小离散度：真实规则的交叉纹理会有一个
            # 稳健主簇（≥40% 象限落在同一周期附近），单个象限被照明/反光带偏是常态。
            # 若改用 max-min，连已知 48px 菱形合成图也会被判「不一致」（假阴性）。
            tol = 0.25 * med if med > 0 else 1.0
            best_cluster = []
            for ref in lags:
                grp = [lg for lg in lags if abs(lg - ref) <= tol]
                if len(grp) > len(best_cluster):
                    best_cluster = grp
            cluster_spread = (
                (max(best_cluster) - min(best_cluster)) / med
                if best_cluster and med > 0
                else 1.0
            )
            if len(best_cluster) >= 3 and cluster_spread <= 0.25:
                stationarity = True
                print(
                    f"[surface-scan] 平稳性检验通过：{len(best_cluster)}/4 象限主峰落在 {med:.0f}px "
                    f"附近（簇内离散度 {cluster_spread * 100:.0f}%），与周期性纹理相容。"
                    "但这只能证明「像纹理」，不能证明「是设计特征」或「有该功能」。"
                )
            else:
                stationarity = False
                reason = (
                    f"最大一致簇仅 {len(best_cluster)}/4"
                    if len(best_cluster) < 3
                    else f"簇内离散度 {cluster_spread * 100:.0f}% > 25%"
                )
                print(
                    f"[surface-scan] ⚠ 平稳性检验不通过（{reason}）："
                    "真实表面纹理是空间平稳的，周期应在多数象限一致。不一致说明该峰由"
                    "**大尺度照明/走刀方向偏置**主导，不足以判定存在规则纹理。"
                )

    strong = [p for p in keep if p[0] >= args.ratio_min]
    if strong and stationarity is False:
        print(
            "[surface-scan] 咨询结论：检出方向性集中，但**未通过平稳性检验**——"
            "更可能是大尺度方向性偏置或成影伪影，不是可判定的表面周期纹理。"
            "若要主张存在纹理，需换侧光/低角度近拍重新取证。"
        )
        strong = []
    elif strong and edge_hit:
        print(
            "[surface-scan] 咨询结论：检出方向性集中，但首峰贴住搜索上限，无法与 1/f "
            "型大尺度偏置区分。请把 --lag-max 加倍重跑：若周期随之外移且强度升高，"
            "则不是表面周期纹理。产物不得据此计入特征。"
        )
        strong = []
    if strong:
        print(
            f"[surface-scan] 咨询结论：检出 {len(strong)} 个各向异性周期峰（强度>="
            f"{args.ratio_min}），形态与功能仍需人工判读（脚本不命名纹理）"
        )
        # 取向聚簇：仅作提示，不作定性。单向条带在实物上多为拉丝/铣削走刀痕。
        angs = sorted((p[3] % 180.0) for p in strong)
        clusters = []
        for a in angs:
            if not clusters or min(abs(a - c) for c in clusters) > 15.0:
                clusters.append(a)
        print(f"[surface-scan] 取向聚簇：{len(clusters)} 组（约 " + "/".join(f"{c:.0f}deg" for c in clusters) + "）")
        if len(clusters) <= 2:
            print(
                "[surface-scan] 提示：单向/双向条带 —— 实物上常见于拉丝、铣削或磨削走刀痕，"
                "默认按**加工痕迹**处理；只有当条带同时满足「节距稳定、覆盖整个功能面、"
                "与邻近未加工面形成明确分界」三条时才可提出设计意图假设，并写入区分依据。"
            )
        else:
            print(
                "[surface-scan] 提示：多方向交叉纹理 —— 不排除是模具/辊压成形的规则结构，"
                "但仍需与加工痕迹区分（看是否局限于功能面、是否与流道或 TIM 边界对齐）。"
            )
    else:
        print(
            "[surface-scan] 咨询结论：未检出可信的方向性周期结构（可能真是平面／纹理由随机"
            "磨损主导／纹理超出可分辨范围／拍摄光照不足）。不得据此断言「没有该特征」，"
            "只能记为「本次裁切未见」。"
        )
    print("[surface-scan] 注意：同一裁切请务必再取一处**已知无特征**的区域做对照；")
    print("[surface-scan] 若对照区也出现同周期同取向的峰，判为成像/压缩伪影，不是表面特征。")
    return 0


def cmd_check(args) -> int:
    ledger = Path(args.ledger)
    if not ledger.exists():
        _die(f"台账不存在：{ledger}", 3)
    rows = list(csv.DictReader(ledger.open(encoding="utf-8-sig")))
    if not rows:
        _die("台账为空", 3)

    problems = []
    seen = {}
    for idx, r in enumerate(rows, 2):
        cls = (r.get("接口面") or "").strip()
        verdict = (r.get("判读结论") or "").strip()
        photo = (r.get("照片文件") or "").strip()
        box = (r.get("裁切框_原生像素") or "").strip()
        shape = (r.get("特征形态") or "").strip()
        seen[cls] = verdict or "未查看"

        if verdict not in VERDICTS:
            problems.append(f"第{idx}行 {cls}: 判读结论非法（{verdict!r}），只能是 有特征/无特征/未查看")
            continue
        if verdict == "未查看":
            problems.append(f"第{idx}行 {cls}: 仍是 未查看 —— 不允许跳过，看完后至少写 无特征")
            continue
        if not photo:
            problems.append(f"第{idx}行 {cls}: {verdict} 但未写照片文件（无法核验真的看过）")
        if not box:
            problems.append(f"第{idx}行 {cls}: {verdict} 但未写裁切框_原生像素")
        if verdict == "有特征" and not shape:
            problems.append(f"第{idx}行 {cls}: 有特征但未写特征形态")

    missing = [c for c, _ in SURFACE_CLASSES if c not in seen]
    if missing:
        problems.append("台账缺少接口面行：" + "、".join(missing))

    have_feature = any(v == "有特征" for v in seen.values())
    have_negative = any(v == "无特征" for v in seen.values())
    if not have_negative and have_feature:
        problems.append(
            "全部判为有特征、没有任何 无特征 记录 —— 说明对照面没看，"
            "无法排除把加工痕迹当成设计特征"
        )

    print(f"[surface-scan] 台账 {ledger}：{len(rows)} 行")
    for c, _ in SURFACE_CLASSES:
        print(f"    {c:<14} {seen.get(c, '缺行')}")

    if problems:
        print(f"\n[surface-scan] 不合格（{len(problems)} 项），不得进入出表阶段：")
        for p in problems:
            print("  - " + p)
        return 1

    print("\n[surface-scan] 合格：每个接口面都留了痕，否定结论也已记录。")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="表面特征发现台账（先于四表执行）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_init = sub.add_parser("init", help="生成空白台账")
    p_init.add_argument("--out", default="surface-ledger.csv")
    p_init.add_argument("--force", action="store_true")
    p_init.set_defaults(func=cmd_init)

    p_probe = sub.add_parser("probe", help="单区域周期性定量咨询")
    p_probe.add_argument("--photo", required=True)
    p_probe.add_argument("--box", required=True, help="x0,y0,x1,y1 原生像素")
    p_probe.add_argument("--lag-min", type=int, default=10)
    p_probe.add_argument("--lag-max", type=int, default=240)
    p_probe.add_argument("--r-min", type=float, default=0.45, help="（保留）绝对相关系数下限")
    p_probe.add_argument("--ratio-min", type=float, default=1.6, help="环内峰/环均值 的显著性下限")
    p_probe.add_argument("--deblock", type=float, default=2.0, help="去 JPEG 块效应的高斯 sigma")
    p_probe.add_argument("--detrend", type=float, default=0.25, help="去照明梯度的背景低通 sigma（占裁切边长比例）")
    p_probe.set_defaults(func=cmd_probe)

    p_check = sub.add_parser("check", help="校验台账是否允许出表")
    p_check.add_argument("--ledger", default="surface-ledger.csv")
    p_check.set_defaults(func=cmd_check)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())