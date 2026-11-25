# 校准数据(Bosch ME/PCB-CN 三份 Benchmark 报告)

> Bosch Confidential。所有估算由 Bosch 同一方法得出,**不是竞品真实成本**,区间是观测范围,不是置信区间。数字取整。
> CTM = LCV Inverter 2nd CTM(2024-04-12);P2 = LCV Inverter Phase 2 Final SBM(2024-08-09);PTU = China PTU Final CTM(2024-11-19)。页码为幻灯片印刷页码。
> 成本页标注 [CNY/pcs] @ 2025;RB 行为"2025 含全部措施"的预测值。数据为 2024 年的,不得自行加通胀系数。
> 三份报告的 Out of scope 页(CTM p3、P2 p3)均含 "Inverter for passenger car":对乘用车牵引逆变器只作部件级锚点。
> 零件价格点在 `data/benchmark/benchmark-price-catalog.csv`(`Anchor Kind = Calibrated`),本文件不重复。

## A. 整机样本(PPC 均为报告原表值;MAT+VA 复算一致,Xin1-M 差 1)

| 类 | 产品 | kg | L(外形) | 配置 | MAT | VA | PPC | 出处 / 备注 |
|---|---|---|---|---|---|---|---|---|
| Xin1 | RB Xin1-M | 25.5 | 45.0 | 5in1;INV 120 kW、DCDC 2.4 kW、DCAC | 6,344 | 743 | 7,088 | CTM p9。P2 p13:6,547 / 540 / 7,087(PMoC VA 在 MAT) |
| Xin1 | Inovance LD35 | 24.5 | 35.6 | 5in1;INV 120 kW、DCDC 3 kW、OBC 选配、DCAC | 4,732 | 392 | 5,124 | CTM p9 |
| Xin1 | RB Xin1-S | ~18 | 12.4 | 4in1;INV 90 kW、DCDC 2.4 kW、OBC 6.6 kW | 3,552 | n/a | — | 含全部措施;"EVOB not available";2025 计划样机 |
| Xin1 | Inovance LD34 | 19 | 24.5 | 4in1;INV 90 kW、DCDC 2 kW、OBC 6.6 kW | 3,540 | 283 | 3,823 | |
| Xin1 | INVT KTZ34X23S0Z | 18 | 29.3 | 3in1;INV 60 kW、DCDC 2 kW | 2,445 | 227 | 2,672 | |
| Xin1 | Sungrow EE30 | 18 | 28.5 | 3in1;INV 60 kW、DCDC 2 kW | 2,331 | 232 | 2,563 | |
| Charcon | VMAX CCU1130(SAIC IM 乘用车) | 7.3 | 9.7 | DCDC 3 kW + OBC 11 kW | 1,795 | 317 | 2,112 | P2 p13;仅 1 个样本 |
| EAC INV | RB INVH200 | 5.9 | 6.1 | INV 45 kW | 3,122 | 276 | 3,398 | P2 p13 |
| EAC INV | Fori FRE120 | 8.2 | 5.1 | INV 45 kW | 2,664 | 324 | 2,988 | |
| EAC INV | Zinsight HS50H | 8.0 | 5.2 | INV 35 kW | 2,121 | 235 | 2,356 | |
| PTU | RB 190F | 32 | 32 | PTU 190 kW | 6,136 | 781 | 6,917 | PTU p7;含全部措施;PMoC VA 在 MAT(622) |
| PTU | Zinsight IPF180HD 1Box | 47 | 38 | PTU 180 kW + INV 45 kW + DCDC 6.6 kW | 6,141 | 911 | 7,052 | 已剔除 EAC+DCL 2,222;重量/体积含多出功能 |
| PTU | Fori | — | — | PTU 220 kW | 8,342 | 1,238 | 9,580 | PTU p30 未给重量/体积 |

## A2. 适用范围

| 类别 | 样本数 | 可用于整机类比 | 可用于部件级锚点 |
|---|---|---|---|
| Xin1 多合一 | 6(竞品 4) | 同功能集合、同功率档 | 功率模块、DC-link、壳体、PCB、连接器 |
| EAC 逆变器 | 3(竞品 2) | 约 35–45 kW 的空压机逆变器 | 同上 |
| PTU | 3(竞品 2) | 燃料电池 PTU | 同上 |
| Charcon | 1 | 无 | OBC/DCDC 部件 |

- 仅 Xin1 类有 ≥ 5 个样本,但全部由同一方法估算,回测不得称为真实误差。
- 重量与体积只做合理性检查,不能定价。
- 不得用 MAT/kW 跨功能组合外推。4in1(90 kW)比 3in1(60 kW)多了 6.6 kW OBC(598 / 633)。

## B. VA 比例(用于 SKILL Step 5)

| 样本 | MAT | VA | VA/MAT | PDC | PDOH+PAC | MOH | MOH/MAT | 出处 |
|---|---|---|---|---|---|---|---|---|
| Inovance LD35 | 4,732 | 392 | 8.3% | 143 | 45 | 204 | 4.3% | CTM p20 |
| Inovance LD34 | 3,540 | 283 | 8.0% | 98 | 34 | 151 | 4.3% | CTM p20 |
| INVT | 2,445 | 227 | 9.3% | 84 | 33 | 109 | 4.5% | CTM p20(分项和 226) |
| Sungrow | 2,331 | 232 | 10.0% | 88 | 37 | 107 | 4.6% | CTM p20 |
| Fori EAC INV | 2,664 | 324 | 12.2% | — | — | — | — | P2 p13 |
| Zinsight EAC INV | 2,121 | 235 | 11.1% | — | — | — | — | P2 p13 |
| VMAX Charcon | 1,795 | 317 | 17.7% | 147 | 70 | 100 | 5.6% | P2 p14(MOH 分项和 99) |
| Zinsight PTU(剔除 Extra) | 6,141 | 911 | 14.8% | — | — | — | — | PTU p7 |
| Fori PTU | 8,342 | 1,238 | 14.8% | — | — | — | — | PTU p7 |
| (RB 参考)RB EAC | 3,122 | 276 | 8.8% | — | — | — | — | P2 p13 |
| (RB 参考)RB 190F | 6,136 | 781 | 12.7% | — | — | — | — | PTU p7 |
| (RB 参考)Xin1-M | 6,344 | 743 | 11.7% | 296 | 133 | 314 | 4.9% | CTM p20 |

- **逆变器/多合一类竞品**(LD35、LD34、INVT、Sungrow、Fori EAC、Zinsight EAC,n = 6):VA/MAT = 8.0%–12.2%,中位约 9.6%。例:MAT = 3,000 → VA ≈ 240–365,Base ≈ 288。
- **跨类竞品**(再加 VMAX、两个 PTU,n = 9):8.0%–17.7%,中位约 11.1%。
- **VA 分块比例**(LCV 四个竞品):PDC 34.6%–37.9%,PDOH+PAC 11.5%–15.9%,MOH 46.1%–53.4%。VMAX 为 46% / 22% / 32%,不同类,不混用。
- VMAX 的 MOH 细分:LOG 18 + Purchasing 17 + PUR tooling 47 + Duty/Freight 17。PDC 细分:Coating 23 + Soldering 14 + SMT 22 + FA+Testing 88。报告注明间接费 170 为 "based on RB factor"。
- 竞品 VA 折算因子(CTM p20),RB 为基准:

  | | Inovance | INVT | Sungrow |
  |---|---|---|---|
  | SMT 单点成本 | -19% | 无信息 | +40% |
  | 预估产量 | 120k | 140k | 200k |
  | HCi workload | +58% | +13% | +16% |
  | 自动化 | 先进 | 无信息 | 先进 |

## C. 隐含件比例(占 MAT,用于 SKILL Step 3)

| 项 | 比例 | 样本 | 出处 |
|---|---|---|---|
| Chem + Scrap | 3.4%–4.5% | Xin1-M 285、LD35 190、Xin1-S 119、LD34 142、INVT 98;Sungrow 94 为按 Box 差额推出 | CTM p13 |
| 螺钉(Screw) | 0.7%–1.6% | 54 / 75 / 36 / 26 / 24 / 24 | CTM p14(螺钉计在 Housing + Cooling 内) |

## D. INV 功能的构成(CTM p11,CNY,LCV Xin1)

| 产品 | PM | DC-link 电容 | 电流传感 | PCB | E-Parts | INV 合计 |
|---|---|---|---|---|---|---|
| RB Xin1-M(SiC) | 1,930 | 214 | 90 | 99 | 442 | 2,775 |
| Inovance LD35 | 912 | 234 | 103 | 62 | 433 | 1,744 |
| RB Xin1-S | 500 | 159 | 58 | 41 | 318 | 1,077 |
| Inovance LD34 | 445 | 142 | 50 | 27 | 268 | 931 |
| INVT | 397 | 151 | 104 | 35 | 315 | 1,003 |
| Sungrow | 302 | 91 | 72 | 93 | 198 | 757 |

- 同一功率档的 PM 差异大(60 kW:302 与 397;90 kW:445 与 500),因技术路线(分立/模块)而异。RB 自制 SiC PM5 不作竞品锚点。
- EAC INV 功能域(P2/PTU p31):RB INV 1,962 + Box 1,042 + PDU+FS 118;Fori 1,715 + 916 + 33;Zinsight 1,199 + 922 + —。RB 的 3-level IGBT 拓扑成本更高。

## E. 工装与设备(PTU p16–17,仅披露,不并入单件)

- 工装总额(kCNY):RB 190F 9,868;Zinsight 一体盒 7,510(壳体 1,620 vs 3,355;盖板 870 vs 1,575;输入 EMC 1,240 vs 170)。
- MAE:RB 自制功率器件(PM5 113 + PMoC 39)152 m CNY;Zinsight OSAT 分立器件 51 m CNY,0.6 s/颗,折算 19.2 s/台(工序工作量,非整线节拍)。
- 采购工装摊销在 MOH:VMAX 的 PUR tooling 47 CNY/pc。

## F. 待核实事项(不是已证明的原文错误)

**F1 差异需解释,建议问报告作者**
1. P2 p25 层数削减 16 + 10 + 10 = 36,标题写 "gross potential: 46"(waterfall 265 → 229 也是 36)。3 行,取整解释不了。
2. CTM p5 功能安全各项相加 106,p21 写 "Saving potential 101"。6 行,取整解释不了。
3. Xin1-M 的 MAT/VA:CTM 6,344/743,P2 6,547/540,两边各变 203,PPC 不变。CTM p20 PMoC 自制 108+24+17+71 = 220,其中 108+24+71 = 203,与变化量相等,但文件未说明原因。

**F2 与取整或版本差异相容,原因尚未确认**
4. Xin1-S 措施合计:P2 p8 逐行相加 99,表上 103;Phase 2 新增逐行 53,摘要 57。该表 13 行,取整误差上限约 6.5,不能排除取整。
5. VMAX MOH 18+17+47+17 = 99,图上 100;INVT VA 分项和 226,表上 227。
6. 同一产品功能域数值随版本变化(Xin1-S DCDC 458 vs 417;LD34 DCDC 418 vs 405;Sungrow 311 vs 325;Xin1-S OBC 598 vs 623;LD34 OBC 633 vs 615),跨版本不混用。
7. P2 p8 与 p15 的措施口径:TC387→TC375 在 Xin1-S 3in1 的措施表为 5,p15 页内毛潜力为 9.4;PMIC 措施表为 5,页内为 4.3。

**F3 仅存疑,需口径确认**
8. Xin1-S 18 kg ÷ 12.4 L ≈ 1.45 kg/L,高于同类约 0.57–0.78。
9. P2 p24 拼板 UT:S2 方案 UT 84.5%,定义与工艺边未给出。
10. CTM p29 "Processing cost 50%":(20.4 − 0.1665 × 34) ÷ 20.4 ≈ 72% 是"售价减净重理论材料费"的剩余,含损耗、嵌件、模具摊销、间接费、利润,不能称加工费。
11. P2 p9 DCDC 图出现两列 "RB Xin1-M"(691 与 458),691 一列额定 4.2 kW,归属不明(691 同时是 VMAX Box 总额),不要使用;458 在 CTM/PTU 中是 Xin1-S。
12. PTU p4 把 190F 与 270F 合并为"互锁连接器颜色 2 → 1",p5 与 p3 写 4 → 1。
13. 设备折旧示例的 7 年是反推。
14. PTU p7 写 Zinsight 的 “DCDC-LV Peak 6 kW”,p30 写 6.6 kW,口径待确认。
