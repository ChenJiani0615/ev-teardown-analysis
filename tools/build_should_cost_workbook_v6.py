#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Should-Cost 工作簿生成器 v6 —— **自制 + 外购混合口径（Make/Buy）**。

与 v2（`build_should_cost_workbook_v2.py`，全外购口径）的差别
================================================
v2 把全部叶件按 Purchased 处理（bottom-up 只用于推导供应商交付价）。
v6 按用户指定口径二分：

    Make       结构件（attrs 给出采购对象且 Bottom-up = Y）按**自制成本**计价：
               价格 = (材料 + 工序) × (1 + 制造间接费率 OH_RATE)
               —— 不含 SG&A / 利润 / 物流（自制件没有供应商加成链）
    Purchased  其余叶件（电子件/密封件/标准件）按交付价；无取证价时按
               基准价目录 Class E 三级补位（行业均价，用户：
               『没有办法算出自制价格的再用行业均值算』）

★ 『自制』是情景假设（用户要求按有自制件的口径建模），不是拆解结论——
  拆解无法判断零件由谁制造。与 v5（全外购口径）的金额不可直接对比。

v6 同时扩展了三条工艺路由（v2 的 OBJECT_ROUTE 表达不了）：
    Stamped part            → 冲压 → 去毛刺 → 清洗（无电镀）
    Machined cooling plate  → 冲压 → 平面磨 → 清洗（冷却板平面度）
    Overmolded busbar       → 冲压 → 注塑包胶 → 去浇口
并把金属材料等效倍率改为**按件选择**：路由含压铸 → K_CAST，否则 K_STAMP
（v2 对铝一律用 K_CAST，冲压铝件的材料成本被低估约 7%）。

以下为 v2 的原始设计说明（三级降级 / 费率去重等约束在 v6 中原样保留）：

与 v1（`build_should_cost_workbook.py`）的根本差别
================================================
v1 按 Make/Buy 二分。v2 改为三层互斥口径：

    Level 1  Purchased Component Cost   ← 每个叶件取供应商交付价
                                         （Bottom-up 推导 或 取证价 A/B/C/D）
    Level 2  Final Integration VA       ← 仅整机总装；全部 Scenario
    Level 3  Complete <product> Cost    ← S0 组件 / S1 我方总装 / S2 整机外购（Unknown）

**产品无关性须知（2026-11-04 修复）**
v2 曾把 "BOM 21 个叶件全部外购" 当成既定事实写进这里，并把工艺路由改成读外购件属性表
（`Procurement Object` → `OBJECT_ROUTE`）。属性表是 DC-DC 项目的交付物，别的产品没有，
于是没有属性表的项目 route_key 全为 None、bottomup 全为 False，**一个 bottom-up 工艺行
都不生成**（实测逆变器 0/53 命中，而 v1 的件名推断是 53/53）。

现行口径：属性表只作**增强**，不是前置条件。分类与路由统一走 `domain_classify`，
三级降级（attrs 精确 → 件名+材料推断 → 显式 Unknown），且降级结果在
`Route Cost Basis` 列里必须可辨认，不允许与第 1 级混同。

本次修复的六处缺陷（逐条对应 7-Evidence 的 Missing Data Audit）
--------------------------------------------------------------
1. 混合材料件按"species 首个命中材料"整件计价 → 改为九类净重分列、各自单价。
   废料回收只对金属生效；塑料/泡棉/胶不减废料。
2. 全包费率行再叠加人工（重复计价）→ RATE_KIND 三分：allin 强制 operators=0，
   machine 另加人工，labor 机时费=0。
3. 裸板面积价冒充 PCBA 交付价 → 裸板单列 `Bare PCB Proxy` 并明标不得作交付价；
   交付价 = 裸板 × PCBA 系数。
4. 供应商制造成本被二次计入 Level 2 → Level 2 只含整机集成动作。
5. 会计修正被当作 Cost Saving → 移入 5-Opportunities 的 Accounting Correction 区。
6. "11 个外购件全部有价，Coverage=100%" → 拆成 Model Coverage 与 Verified Price Coverage。

工程约束（吃过亏的地方）
------------------------
* 所有跨表引用由 `COLS[...]` 生成，**不手写列字母**；禁止 `$X$X` 这类残缺整列引用。
* 良率按工序**累积**串联：C_i = (直接成本_i + C_{i-1}) / 良率_i，
  最后一道工序的累积值才是该件的加工成本（用 Is Last Step 列标记，避免 SUM 全列重复计）。
* Unknown 留空不补零；金额列只对已有值求和。

用法
----
    py -3 -X utf8 tools/build_should_cost_workbook_v6.py ^
        --bom    data/ex90-rear-inverter/derived/inverter-leaves-v5.csv ^
        --attrs  data/ex90-rear-inverter/derived/inverter-selfmade-attributes-v6.csv ^
        --catalog data/benchmark/benchmark-price-catalog.csv ^
        --product "EX90 Rear Inverter" --function-prefix INV ^
        --out    reports/EX90_RearInverter_ShouldCost_Analysis_v6.xlsx

    # 不给 --attrs 时全部叶件按外购处理（等价 v2 行为）；不给 --prices 时
    # 外购件按基准价目录 Class E 补位（行业均价）。
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from domain_classify import (  # noqa: E402  单一真源：分类/外购/路由
    func_prefix_for, function_category, is_buy, route_for, route_strength,
    to_step_dicts,
)
from benchmark_price import (  # noqa: E402  基准价目录（Class E）三级补位
    CATALOG_DEFAULT, benchmark_price, load_catalog,
)

# ================================================================ 样式
HDR_FILL = PatternFill("solid", fgColor="305496")
HDR_FONT = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=9)
BOLD_FONT = Font(name="微软雅黑", size=9, bold=True)
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
SEC_FONT = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
NOTE_FONT = Font(name="微软雅黑", size=9, color="595959")
WARN_FONT = Font(name="微软雅黑", size=9, italic=True, color="C00000")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STATUS_FILL = {
    "Verified": None,
    "Calculated": PatternFill("solid", fgColor="DDEBF7"),
    "Estimated": PatternFill("solid", fgColor="FFF2CC"),
    "Scenario": PatternFill("solid", fgColor="FCE4D6"),
    "Unknown": PatternFill("solid", fgColor="FCE4E4"),
}
MONEY = '#,##0.00'
KG = '0.000'
RATIO = '0.0000'
PCT = '0.0%'

# ================================================================ 费率口径
# 口径比数值更影响结果：同一份人工若既在费率里、又在人数里，就是重复计价。
RATE_KIND = {
    "RATE_CAST": "allin",
    "RATE_STAMP": "allin",
    "RATE_INJ": "allin",
    "RATE_TEST": "allin",
    "RATE_CNC": "machine",
    "RATE_ASSY": "labor",
}
RATE_KIND_LABEL = {
    "allin": "全包（含人工）",
    "machine": "裸机（不含人工）",
    "labor": "纯人工",
    "outsrc": "外协（单价×基数）",
}

# ================================================================ 参数表
# (ID, 名称, 值, 单位, 状态, 来源/说明)
PARAMS: list[tuple] = [
    ("CUR", "币种", "CNY", "-", "Unknown", "未取得报价，币种为假设"),
    ("REG_SUP", "零件供应商生产地区", "中国华东", "-", "Estimated",
     "ECU 产地未取得；只作用于供应商侧制造成本模型"),
    ("REG_INT", "整机集成地点 / 责任方", "Unknown", "-", "Unknown",
     "★ 不填 = 不得把整机总装标为『我方自制』；Level 2 一律 Scenario"),
    ("VOL_A", "年产量", 20000, "件/年", "Unknown", "待项目输入；影响设备折旧"),
    ("VOL_L", "生命周期产量（合格）", 100000, "件", "Unknown", "待项目输入；摊销分母用合格产量"),
    # ---- 材料单价
    ("AL_PRICE", "铝压铸合金单价", 23.9, "CNY/kg", "Verified",
     "长江现货 A00 铝锭 2026-10-28；压铸合金 AC44300 现货通常另加 2–4 元/kg 合金费，未计入"),
    ("CU_PRICE", "铜单价", 110.5, "CNY/kg", "Verified", "长江现货 1# 电解铜 2026-10-28"),
    ("ST_PRICE", "钢单价", 8.0, "CNY/kg", "Estimated", "冷轧/镀锌板"),
    ("PBT_PRICE", "PBT 单价", 25.0, "CNY/kg", "Estimated", ""),
    ("PPS_PRICE", "PPS 单价", 60.0, "CNY/kg", "Estimated", ""),
    ("PPA_PRICE", "PPA 单价", 45.0, "CNY/kg", "Estimated", ""),
    ("PA_PRICE", "PA66/PA6/PP/HDPE 单价", 30.0, "CNY/kg", "Estimated", "含 GF 增强牌号"),
    ("PLASTIC_PRICE", "通用工程塑料单价", 25.0, "CNY/kg", "Estimated",
     "PC-ABS/ABS/POM 等未单独列价的牌号"),
    ("FERRITE_PRICE", "铁氧体单价", 30.0, "CNY/kg", "Estimated", "烧结磁芯"),
    ("ELASTOMER_PRICE", "弹性体单价（TPV/TPS）", 32.0, "CNY/kg", "Estimated",
     "仅用于无独立取证价的模压/模切件；本清单两件垫圈已有取证价，此参数不生效"),
    ("FOAM_PRICE", "PE 泡棉单价", 15.0, "CNY/kg", "Estimated", "模切泡棉片"),
    ("ADHESIVE_PRICE", "胶粘剂单价（结构胶 / 密封胶同价）", 60.0, "CNY/kg", "Estimated",
     "上盖结构胶与下盖密封剂同列此价；若两者牌号价差 > 30% 须拆列"),
    # ---- 利用率与回收
    ("SCRAP_REC", "废料回收率", 0.60, "-", "Estimated",
     "★ 只对金属生效（铝/铜/钢）；塑料、泡棉、胶无回收依据，不参与"),
    ("Y_CAST", "压铸材料利用率", 0.85, "-", "Estimated", "含浇口/溢流"),
    ("Y_STAMP", "冲压材料利用率", 0.72, "-", "Estimated", "含排样废料，须实测"),
    ("Y_INJ", "注塑材料利用率", 0.92, "-", "Estimated", "含水口料回用"),
    ("Y_SINT", "烧结材料利用率", 0.90, "-", "Estimated", "铁氧体磁芯磨削余量"),
    ("Y_DIECUT", "模切利用率", 0.85, "-", "Estimated", "泡棉排样"),
    # ---- 设备 / 人工费率
    ("RATE_CAST", "压铸机费率（含人工）", 180.0, "CNY/h", "Estimated",
     "全包口径：设备折旧+能耗+维护+操作工。★ 不得再加人工"),
    ("RATE_STAMP", "冲压机费率（含人工）", 120.0, "CNY/h", "Estimated", "全包口径。★ 不得再加人工"),
    ("RATE_INJ", "注塑机费率（含人工）", 110.0, "CNY/h", "Estimated", "全包口径。★ 不得再加人工"),
    ("RATE_TEST", "测试/检漏台费率（含人工）", 80.0, "CNY/h", "Estimated", "全包口径。★ 不得再加人工"),
    ("RATE_CNC", "CNC/清洗机费率（不含人工）", 90.0, "CNY/h", "Estimated", "裸机口径。★ 必须另加人工"),
    ("RATE_ASSY", "人工费率", 45.0, "CNY/人·h", "Estimated",
     "纯人工。★ 只能作为人工，不能同时当机时费"),
    ("CAST_BASE_S", "压铸基础节拍（开合模/取件）", 35.0, "s", "Estimated", "与重量无关的固定部分"),
    ("CAST_S_PER_KG", "压铸节拍重量系数", 45.0, "s/kg", "Estimated",
     "浇注重量与冷却时间近似线性；粗算用净重替代浇注重量，偏乐观"),
    ("INJ_BASE_S", "注塑基础节拍", 25.0, "s", "Estimated", "开合模/取件"),
    ("INJ_S_PER_KG", "注塑节拍重量系数", 30.0, "s/kg", "Estimated", ""),
    ("PLATE_PRICE", "电镀外协单价", 20.0, "CNY/kg", "Estimated",
     "外协工序按 单价 × 基数 ÷ 良率 计价，不走设备费率；基数为件净重"),
    # ---- 供应商加价链
    ("OH_RATE", "供应商制造间接费率", 0.12, "-", "Estimated", "对供应商直接成本（材料+加工）"),
    ("SGA_RATE", "供应商 SG&A 率（★ v6 仅 Purchased 件口径参考；Make 件不计）", 0.05, "-", "Estimated", ""),
    ("LOG_RATE", "物流包装率（供应商→集成厂；★ v6 仅 Purchased 件口径参考；Make 件不计）", 0.01, "-", "Estimated",
     "★ 只加这一段；Level 2 之后不得再叠加同段运输"),
    ("PROFIT_RATE", "供应商利润率（按销售额；★ v6 仅 Purchased 件口径参考；Make 件不计）", 0.08, "-", "Estimated",
     "Price = Cost/(1−margin)；不是成本加成。v6 的 Make 件售价 = 成本×(1+OH)，不用本参数"),
    # ---- 工装
    ("TOOL_LIFE_SHOTS", "工装寿命（模次）", 500000, "模次", "Estimated",
     "摊销基数 = MIN(寿命模次, 生命周期产量)"),
    # ---- 价格区间
    ("RANGE_MAX_RATIO", "区间可用上限（High/Low）", 1.5, "倍", "Estimated",
     "≤ 该值取均值并保留区间；> 该值取几何均值、区间仅供参照"),
    ("CENTER_METHOD", "宽区间取值方法", "GEOMEAN", "-", "Estimated",
     "GEOMEAN = 几何均值 √(Low×High)；此处仅作声明，公式逻辑见 2-BOM Price Base"),
    # ---- 基准价补位（Class E，Step 5.8）
    ("BENCH_FILL", "基准价补位开关", "Y", "-", "Estimated",
     "Y = 无取证价且无 bottom-up 的行按基准目录三级补位（3a 严格类别 / 3b 包络类别 / "
     "3c 质量轴兜底）。关掉后这些行恢复留空（Unknown）。"
     "★ 补位不覆盖已有取证价或 bottom-up 模型，也不计入 Verified Price Coverage"),
    ("BENCH_CATALOG", "基准价目录", "data/benchmark/benchmark-price-catalog.csv", "-",
     "Estimated",
     "每条锚点带来源与取证日期、90 天复核期；Calibration 行只用于校验模型、不参与补位"),
    # ---- 整机集成（Level 2）——全部 Scenario
    ("INTEG_IDENTITY", "整机集成方", "Unknown", "-", "Unknown",
     "★ 未取证。S1 情景假设我方总装；S2 情景假设整机外购"),
    ("INTEG_CT_MIN", "整机总装节拍", 12.2, "min/台", "Scenario",
     "★ 原口径 = 68 界面点 × 9 s + 合盖 60 s + 搬运 60 s。"
     "界面点数 ≠ 装配动作数，本值仅作情景，不是实测工时"),
    ("INTEG_EOL_CT_MIN", "EOL 测试节拍（Hi-Pot/功能/气密）", 5.0, "min/台", "Scenario",
     "逐台 EOL，非抽检；无供应商工艺记录"),
    ("INTEG_FIPG_G", "整机密封胶用量", 0.0, "g/台", "Scenario",
     "0 = 假设密封靠成型垫圈（TPV/TPS）；下盖 species 含 Sealants，需拆解确认是否有胶道"),
    ("INTEG_FIPG_PRICE", "密封胶单价", 90.0, "CNY/kg", "Estimated", "车规脱醇型硅酮"),
    ("INTEG_CONS", "整机其他辅料", 2.0, "CNY/台", "Scenario", "清洗剂/擦拭/扎带/标签"),
    ("INTEG_INV", "整机集成设备投资（点胶+拧紧+气密+EOL）", 5000000.0, "CNY", "Unknown",
     "★ 未经确认的设备投资假设 → 由此推导的折旧/件只能标 Scenario，不是实际成本"),
    ("INTEG_LIFE", "集成设备折旧年限", 5.0, "年", "Estimated", "直线折旧"),
]

# ================================================================ 工序模板
def step(proc: str, cat: str, cycle, rate, ops: float = 0.0,
         cons=0.0, yld: float = 0.98, basis: str = "") -> dict:
    return {"proc": proc, "cat": cat, "cycle": cycle, "rate": rate,
            "ops": ops, "cons": cons, "yld": yld, "basis": basis}


# 大件压铸（有密封面 → 需 CNC + 检漏）
CAST_HEAVY = [
    step("压铸", "加工", ("cast", None), "RATE_CAST", 0, 0.0, 0.95, "重量式"),
    step("修边/去浇口", "加工", 20, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("CNC 加工（密封面）", "加工", 120, "RATE_CNC", 1.0, 3.0, 0.98, "模板"),
    step("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
    step("检漏", "测试", 30, "RATE_TEST", 0, 2.0, 0.99, "模板"),
]
# 小件压铸（无密封面 → 不做 CNC 与检漏）
CAST_LIGHT = [
    step("压铸", "加工", ("cast", None), "RATE_CAST", 0, 0.0, 0.95, "重量式"),
    step("修边/去浇口", "加工", 20, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
]
BUSBAR = [
    step("冲压", "加工", 12, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("电镀（外协）", "加工", None, None, 0, ("outsrc", "PLATE_PRICE"), 0.98, "外协"),
    step("清洗", "加工", 10, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
]
INJ = [
    step("注塑", "加工", ("inj", None), "RATE_INJ", 0, 0.0, 0.97, "重量式"),
    step("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99, "模板"),
]
INJ_INSERT = INJ + [
    step("嵌件装配", "装配", 40, "RATE_ASSY", 1.0, 0.5, 0.98, "模板"),
]

# v6 新增工序模板（与 domain_classify 的 STAMP/PLATE/OVERMOLD 同构，但用
# v2 生成器的 step() dict 形态；费率键与 1-Assumptions 一致）
STAMP_V6 = [
    step("冲压", "加工", 8, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("去毛刺", "加工", 15, "RATE_ASSY", 1.0, 0.0, 0.99, "模板"),
    step("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
]
PLATE_V6 = [
    step("冲压", "加工", 10, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("平面磨", "加工", 60, "RATE_CNC", 1.0, 2.0, 0.97, "模板"),
    step("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
]
OVERMOLD_V6 = [
    step("冲压", "加工", 12, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("注塑包胶", "加工", 50, "RATE_INJ", 0, 0.0, 0.96, "模板"),
    step("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99, "模板"),
]

ROUTES: dict[str, tuple[str, list[dict]]] = {
    "Die-cast assembly (heavy)": ("压铸 → 修边 → CNC 密封面 → 清洗 → 检漏", CAST_HEAVY),
    "Die-cast part (light)": ("压铸 → 修边 → 清洗", CAST_LIGHT),
    "Stamped busbar assembly": ("冲压 → 电镀（外协）→ 清洗", BUSBAR),
    "Stamped copper part": ("冲压 → 电镀（外协）→ 清洗", BUSBAR),
    "Overmolded terminal": ("注塑 → 去浇口 → 嵌件装配", INJ_INSERT),
    "Injection molding": ("注塑 → 去浇口", INJ),
    # ---- v6 新增：v2 的六键表达不了的三类结构件 ----
    "Stamped part": ("冲压 → 去毛刺 → 清洗", STAMP_V6),
    "Machined cooling plate": ("冲压 → 平面磨 → 清洗", PLATE_V6),
    "Overmolded busbar": ("冲压 → 注塑包胶 → 去浇口", OVERMOLD_V6),
}

OBJECT_ROUTE = {
    "Die-cast assembly": "Die-cast assembly (heavy)",
    "Die-cast part": "Die-cast part (light)",
    "Stamped busbar assembly": "Stamped busbar assembly",
    "Stamped copper part": "Stamped copper part",
    "Overmolded terminal": "Overmolded terminal",
    "Injection molding": "Injection molding",
    # ---- v6 新增 ----
    "Stamped part": "Stamped part",
    "Machined cooling plate": "Machined cooling plate",
    "Overmolded busbar": "Overmolded busbar",
}

# v6 新增工序模板（与 domain_classify 的 STAMP/PLATE/OVERMOLD 同构，但用
# v2 生成器的 step() dict 形态；费率键与 1-Assumptions 一致）
STAMP_V6 = [
    step("冲压", "加工", 8, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("去毛刺", "加工", 15, "RATE_ASSY", 1.0, 0.0, 0.99, "模板"),
    step("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
]
PLATE_V6 = [
    step("冲压", "加工", 10, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("平面磨", "加工", 60, "RATE_CNC", 1.0, 2.0, 0.97, "模板"),
    step("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99, "模板"),
]
OVERMOLD_V6 = [
    step("冲压", "加工", 12, "RATE_STAMP", 0, 0.0, 0.98, "模板"),
    step("注塑包胶", "加工", 50, "RATE_INJ", 0, 0.0, 0.96, "模板"),
    step("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99, "模板"),
]

# v6：路由 → 金属材料等效倍率键。压铸件用 K_CAST（利用率 0.85），
# 其余金属件（冲压/磨削）用 K_STAMP（利用率 0.72）。
# v2 对铝一律用 K_CAST——冲压铝件的材料成本被低估约 7%。
ROUTE_K_METAL = {
    "Die-cast assembly (heavy)": "K_CAST",
    "Die-cast part (light)": "K_CAST",
}

# ================================================================ 价格状态映射
# 七类价格状态（workbook-spec）：Actual Procurement / Supplier Quotation / Comparable /
# Bottom-up Estimate / Parametric Estimate / Assumption / Unknown
PTYPE2STATUS = {
    "Actual Procurement": "Actual Procurement",
    "Supplier Quotation": "Supplier Quotation",
    "Comparable Market Price": "Comparable",
    "Bottom-up Estimate": "Bottom-up Estimate",
    "Parametric Estimate": "Parametric Estimate",
    "Benchmark Average": "Comparable",
    "Unsupported Assumption": "Assumption",
    "Assumption": "Assumption",
    "Unknown": "Unknown",
}
STATUS2EVIDENCE = {
    "Actual Procurement": "Verified",
    "Supplier Quotation": "Estimated",
    "Comparable": "Estimated",
    "Bottom-up Estimate": "Calculated",
    "Parametric Estimate": "Estimated",
    "Assumption": "Estimated",
    "Unknown": "Unknown",
}

# ================================================================ 功能域
# ★ 已迁到 tools/domain_classify.py（单一真源）。
#   此处原先的 function_category 把 DC-DC 当默认前缀，导致其他产品在无属性表时
#   全部被打上 DCDC 标记；同时它按全文匹配件名，父级上下文会盖掉零件类型
#   （"Gasket, ... Power Module" 被判成 PM 功率级）。两处缺陷已在新模块修正。


def plastic_param(species: str) -> str:
    s = species.lower()
    if "pps" in s:
        return "PPS_PRICE"
    if "ppa" in s:
        return "PPA_PRICE"
    if "pbt" in s:
        return "PBT_PRICE"
    if any(k in s for k in ("pa66", "pa6", "pp-", "hdpe", "pp ")):
        return "PA_PRICE"
    return "PLASTIC_PRICE"


# ================================================================ 读数据
SPLIT_COLS = ["AL (kg)", "CU (kg)", "STEEL (kg)", "PLASTIC (kg)", "ELASTOMER (kg)",
              "FERRITE (kg)", "FOAM (kg)", "ADHESIVE (kg)", "OTHER (kg)"]


def load_leaves(path: Path) -> list[dict]:
    """读叶件清单。**允许 Total Weight 为空的行（标准件）**。

    v3 起叶件清单里多了 19 行紧固/密封件（Part Type = Standard）。它们的 Total Weight
    本来就被父件合并了（父件 Total = Part Weight + Fastener Weight），所以
    `total_weight = 0` 是**正确口径**而不是缺数据：

        * 质量基线只由 total_weight > 0 的行构成（避免与父件重复计重）；
        * 成本独立按件计价（这是 v2 的全部缺口）。
    """
    out = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            qty_raw = (r.get("Quantity") or "").strip()
            qty = float(qty_raw) if qty_raw else 1.0
            tot_raw = (r.get("Total Weight (kg)") or "").strip()
            total = float(tot_raw) if tot_raw else 0.0
            fw_raw = (r.get("Fastener Weight") or "").strip()
            itf_raw = (r.get("Number of Interface") or "").strip()
            ptype = (r.get("Part Type") or "").strip()
            len_raw = (r.get("Length") or "").strip()
            wid_raw = (r.get("Width") or "").strip()
            out.append({
                "part": (r.get("CS Part") or "").strip(),
                "name": (r.get("Name") or "").strip(),
                "qty": qty or 1.0,
                "weight": total / qty if qty else total,
                "total_weight": total,
                "fastener_weight": float(fw_raw) if fw_raw else 0.0,
                "group": (r.get("Material Group") or "").strip(),
                "species": (r.get("Material Species") or "").strip(),
                # ★ 外形尺寸：仅用于面积计价类（PCB 元/m²）。缺失即不给面积价，绝不猜尺寸。
                "length": float(len_raw) if len_raw else 0.0,
                "width": float(wid_raw) if wid_raw else 0.0,
                "supplier": (r.get("Manufacturer") or "").strip(),
                "marking": (r.get("Marking Number") or "").strip(),
                "itf": int(float(itf_raw)) if itf_raw else 0,
                "ptype": ptype,
                "standard": ptype.lower() == "standard" or total == 0.0,
                "parent": (r.get("Parent Part No") or "").strip(),
                "mass_note": (r.get("Mass Accounting") or "").strip(),
            })
    return out


def load_attrs(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    out: dict[str, dict] = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            part = (r.get("CS Part") or "").strip()
            if not part:
                continue
            splits = {}
            for c in SPLIT_COLS:
                v = (r.get(c) or "").strip()
                splits[c] = float(v) if v else 0.0
            out[part] = {
                "object": (r.get("Procurement Object") or "").strip(),
                "bottomup": (r.get("Bottom-up Model (Y/N)") or "").strip().upper() == "Y",
                "route_status": (r.get("Route Status") or "").strip(),
                "split_basis": (r.get("Material Split Basis") or "").strip(),
                "splits": splits,
                "notes": (r.get("Notes") or "").strip(),
                "correction": (r.get("Correction Note") or "").strip(),
            }
    return out


def load_prices(path: Path | None) -> dict[str, dict]:
    if not path:
        return {}
    out: dict[str, dict] = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            part = (r.get("CS Part") or "").strip()
            if not part:
                continue
            low = (r.get("Price Low (CNY)") or "").strip()
            high = (r.get("Price High (CNY)") or "").strip()
            bare = (r.get("Bare PCB Proxy (CNY)") or "").strip()
            out[part] = {
                "low": float(low) if low else None,
                "high": float(high) if high else None,
                "cls": (r.get("Price Class") or "").strip(),
                "ptype": (r.get("Price Type") or "").strip(),
                "pstatus": (r.get("Price Status") or "").strip(),
                "basis": (r.get("Price Basis") or "").strip(),
                "boundary": (r.get("Delivered Price Boundary") or "").strip(),
                "bare": float(bare) if bare else None,
                "bare_basis": (r.get("Bare PCB Proxy Basis") or "").strip(),
                "src_date": (r.get("Source Date") or "").strip(),
            }
    return out


# ================================================================ 列定义（单一真源）
def make_cols(spec: list[tuple[str, str, int]]) -> tuple[list[str], list[int], dict[str, str]]:
    """spec = [(key, label, width)] → (labels, widths, {key: letter})"""
    keys = [k for k, _, _ in spec]
    labels = [lab for _, lab, _ in spec]
    widths = [w for _, _, w in spec]
    cols = {k: get_column_letter(i) for i, k in enumerate(keys, start=1)}
    return labels, widths, cols


def load_opps(path: Path | None) -> list[dict]:
    """读取产品专属机会表（可选）。★ 通用工具不内置任何产品的降本次目：
    某产品的降本分析是**数据**，不是代码。缺省时由 build_opps() 按 BOM 推导
    通用候选（材质归并 / 紧固件精简 / 取证缺口），绝不会凭空写出别的产品的零件。

    CSV 列（缺列按空处理）：id, kind, title, now, alt, mech, s_mat, s_mfg,
    s_tool, price, new, trade, valid, mat_imp, va_imp, tool_imp, mech_tag,
    dedup, hg, risk, weeks, status, prio
    """
    if not path or not path.exists():
        return []
    out = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("title") or "").strip():
                continue
            out.append({
                "id": (row.get("id") or "").strip(),
                "kind": (row.get("kind") or "设计").strip(),
                "title": (row.get("title") or "").strip(),
                "now": (row.get("now") or "").strip(),
                "alt": (row.get("alt") or "").strip(),
                "mech": (row.get("mech") or "").strip(),
                "s_mat": (row.get("s_mat") or "").strip(),
                "s_mfg": (row.get("s_mfg") or "").strip(),
                "s_tool": (row.get("s_tool") or "").strip(),
                "price": (row.get("price") or "").strip(),
                "new": (row.get("new") or "").strip(),
                "trade": (row.get("trade") or "").strip(),
                "valid": (row.get("valid") or "").strip(),
                "mat_imp": (row.get("mat_imp") or "无").strip() or "无",
                "va_imp": (row.get("va_imp") or "无").strip() or "无",
                "tool_imp": (row.get("tool_imp") or "无").strip() or "无",
                "mech_tag": (row.get("mech_tag") or "").strip(),
                "dedup": (row.get("dedup") or "").strip(),
                "hg": (row.get("hg") or "HG0").strip(),
                "risk": (row.get("risk") or "").strip(),
                "weeks": (row.get("weeks") or "").strip(),
                # ★ 任何外部条目一律强制 Hypothetical：不得从 CSV 里带进
                #   Verified Saving / HG4 / HG5 这类未经验证的状态。
                "status": "Hypothetical Opportunity",
                "prio": (row.get("prio") or "中").strip(),
            })
    return out


def load_opps_ac(path: Path | None) -> list[dict]:
    """读取产品专属会计修正表（可选）。列：id, title, before, after, reason, amt, why_not_saving"""
    if not path or not path.exists():
        return []
    out = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("title") or "").strip():
                continue
            out.append({
                "id": (row.get("id") or "").strip(),
                "title": (row.get("title") or "").strip(),
                "before": (row.get("before") or "").strip(),
                "after": (row.get("after") or "").strip(),
                "reason": (row.get("reason") or "").strip(),
                "amt": (row.get("amt") or "").strip(),
                "why": (row.get("why") or "").strip(),
            })
    return out


# ================================================================ 写表工具
def write_header(ws, row: int, labels: list[str], widths: list[int]) -> None:
    for i, (label, width) in enumerate(zip(labels, widths), start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = 32


def _text(v):
    """说明性文本若以 '=' 开头会被 openpyxl 当公式写入，Excel 会判文件非法。

    v1/v2 的第一版就死在这一点上：三处说明文字以 '=' 开头（『= 供应商交付价合计』等），
    结果 4-Summary 整张表让 Excel 拒绝打开（Open method of Workbooks class failed）。
    """
    if isinstance(v, str) and v.lstrip().startswith("="):
        return "＝" + v.lstrip()[1:]  # 全角等号，视觉一致、不再被解析为公式
    return v


def write_body(ws, row: int, values: list, status: str | None = None,
               literal: set[int] | None = None) -> None:
    """literal = 需要强制按文本写入的 1-based 列号集合（说明列）。"""
    for i, v in enumerate(values, start=1):
        if literal and i in literal:
            v = _text(v)
        c = ws.cell(row=row, column=i, value=v)
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = Alignment(vertical="top", wrap_text=True)
    if status and STATUS_FILL.get(status):
        for i in range(1, len(values) + 1):
            ws.cell(row=row, column=i).fill = STATUS_FILL[status]


def write_note(ws, row: int, text: str, font=NOTE_FONT) -> int:
    ws.cell(row=row, column=1, value=_text(text)).font = font
    return row + 1


def write_block(ws, start: int, title: str, rows: list[tuple]) -> dict:
    """三列块：(标签, 值或公式, 说明, 状态)。返回各关键行号。"""
    ws.cell(row=start, column=1, value=title).font = SEC_FONT
    for i, h in enumerate(["项目", "金额 / 值", "说明"], start=1):
        c = ws.cell(row=start + 1, column=i, value=_text(h))
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    first = start + 2
    r = first
    marks: dict[str, int] = {}
    for label, value, note, status in rows:
        c1 = ws.cell(row=r, column=1, value=_text(label))
        c1.font = BOLD_FONT if str(label).startswith(("Level", "S0", "S1", "S2")) else BODY_FONT
        c1.border = BORDER
        c1.alignment = Alignment(vertical="top", wrap_text=True)
        c2 = ws.cell(row=r, column=2, value=value)
        c2.font = BODY_FONT
        c2.border = BORDER
        if status and STATUS_FILL.get(status):
            c2.fill = STATUS_FILL[status]
        c3 = ws.cell(row=r, column=3, value=_text(note))
        c3.font = NOTE_FONT
        c3.border = BORDER
        c3.alignment = Alignment(vertical="top", wrap_text=True)
        marks[label] = r
        r += 1
    return {"first": first, "last": r - 1, "next": r + 1, "marks": marks}


def box(ws, row: int, lines: list[str], font=WARN_FONT) -> int:
    for t in lines:
        row = write_note(ws, row, t, font)
    return row


# ================================================================ 机会/修正推导
# ★ 2026-11-04：5-Opportunities 原为整块写死的 DC-DC 内容。通用工具必须由实际
#   BOM 推导候选项，否则换一个产品就会在表里讨论**根本不存在的零件**。
def _fmt_qty(q) -> str:
    return f"{q:g}" if isinstance(q, (int, float)) else str(q)


def build_opps(parts: list[dict], *, ev_n: int, unknown_n: int,
               has_attr_split: bool, std_part_n: int, std_qty: float,
               mass_rows: list[dict]) -> list[tuple]:
    """按实际 BOM 推导降本机会。所有条目一律 Hypothetical（无基线价、无报价）。"""
    out: list[tuple] = []
    oid = 1

    def nid() -> str:
        nonlocal oid
        oid += 1
        return f"O{oid - 1:02d}"

    bu = [p for p in parts if p["bottomup"]]
    unpriced = [p for p in parts if not p["priced"] and not p["bottomup"]]

    # ---- 1. 最大金额件（按 bottom-up 售价，或若全无价则按件数最集中处）
    if bu:
        top = max(bu, key=lambda p: p["qty"])
        if top["qty"] >= 2:
            out.append((
                nid(), "设计", f"多用量件规格归并：{top['name'][:24]}（Qty {_fmt_qty(top['qty'])}）",
                f"{top['name']} 单件 Qty = {_fmt_qty(top['qty'])}，"
                f"料号 {top['part']}，功能域 {top['func']}/{top['cat']}",
                "同功能位置若有多个相似料号，复核能否共用一种规格或一套模具",
                "同族件多料号 = 多套模具 + 多套库存 + 多套尺寸验证；"
                "归并的收益主要在工装摊销与换型损失，材料金额变化很小",
                "材料金额影响小（同族材料单价接近）",
                "减少换型与检验批次",
                "可省一套模具（金额 Unknown）",
                "★ 单件收益通常 < 5 元/台；价值主要在料号精简与管理成本",
                "无",
                "归并后承载/定位能力必须复核，不能只看外形相近",
                "按实际工况做承载/振动/装配干涉验证",
                "MAT", "无", "Tooling", "物料归并", f"G{oid}", "HG1",
                "低：需确认承载与定位能力", "4–6 周",
                "Hypothetical Opportunity", "中"))

    # ---- 2. 密封/紧固件包（件数多但金额小 —— 这是本类产品最有价值的结构性结论）
    std = [p for p in parts if p["standard"]]
    if std:
        out.append((
            nid(), "设计",
            f"标准紧固/密封件的规格精简（{std_part_n} 料号 / {std_qty:.0f} 件实物）",
            f"{std_part_n} 个料号、{std_qty:.0f} 件实物，在 BOM 中占 {len(std)} 行；"
            "单价极低但采购与库存管理成本不随单价缩放",
            "同规格不同长度/表面处理的螺钉归并为 2–3 种；O 圈按内径与线径系列化",
            "标准件的成本主体不在材料，而在采购批次、库存与错料风险；"
            "件数多不等于钱多，但件数多等于管理动作多",
            "材料影响可忽略（标准件单价极低）",
            "减少采购批次与库存 SKU",
            "无",
            "★ 采购价影响 < 1 元/台；此项属『管理型』而非『设计型』降本",
            "无",
            "归并后必须复核夹紧力与密封压缩率，不可只按长度近似",
            "夹紧力核算 + 压缩永久变形（ISO 815）+ 气密",
            "MAT", "无", "无", "物料归并", f"G{oid}", "HG1",
            "低", "4–8 周", "Hypothetical Opportunity", "中"))

    # ---- 3. 无价件（最大的实质缺口 —— 先取证再谈降本）
    if unknown_n:
        out.append((
            nid(), "取证",
            f"先补 {unknown_n} 项无价件的价格证据（本表最优先动作）",
            f"{unknown_n} 行金额为空（不是 0）；Level 1 合计只反映已建模件，"
            "是结构性低估",
            "按件号向采购系统/供应商询价；无报价的按参数化建模（面积/重量/工艺类比）",
            "★ 在无价的件上讨论降本，等于在没有基线的数字上算百分比；"
            "先把 Unknown 变成 Comparable，降本结论才有意义",
            "不改变材料成本，改变的是**可见的**成本边界",
            "无",
            "无",
            "★ 这不是降本措施：它会把 Level 1 金额**抬高**到真实水平，"
            "必须与降本严格分开",
            "无",
            "询价结果可能显著高于当前估算",
            "逐件取得合同价或书面报价；无价件单列清单跟踪",
            "无", "无", "无", "取证补齐", f"G{oid}", "HG2",
            "无（必须先做）", "6–10 周", "Evidence Action", "高"))

    # ---- 4. 材料分层缺失（只有真的缺时才出这一条）
    if not has_attr_split:
        out.append((
            nid(), "取证", "补齐材料分层（当前 Material Input = 0）",
            "无 attrs 或 Object 未识别 → 材料分层为空，"
            "Supplier Mfg Cost 仅含加工费，材料输入列为 0",
            "补齐 purchased-attributes 表：每件的九类材料净重与工序路由",
            "材料与加工不可分时，任何『材料替代』类降本都无法量化——"
            "只能定性说方向，不能给金额",
            "补齐后材料成本会**上升**（当前是 0，不是真的 0）",
            "补齐后可拆到具体工序，暴露出真正的加工杠杆",
            "无",
            "★ 这不是降本：是让材料成本从『隐身为 0』变成真实值",
            "无",
            "分层重量需要拆解称重或供应商物料清单，二者都可能拿不到",
            "逐件称重 + 材料牌号确认（丝印/光谱）",
            "无", "无", "无", "取证补齐", f"G{oid}", "HG2",
            "高：数据可得性不确定", "8–12 周", "Evidence Action", "高"))

    # ---- 5. 界面点数（装配动作的规模指标）
    if len(mass_rows) >= 20:
        out.append((
            nid(), "设计", "界面点数复核（螺钉数 ≠ 连接功能数）",
            f"BOM 共 {len(mass_rows)} 个计量叶件，界面点数见 1-Assumptions IFACE_CNT",
            "按『连接功能』而非『螺钉数』复核界面；可合并的界面合并，"
            "可改卡扣/压筋的取消紧固件",
            "每个界面 = 一次定位 + 一次紧固 + 一次扭矩验证；"
            "界面数是装配成本与直通率的一阶驱动，且与件数无关",
            "材料影响小",
            "供应商侧减少拧紧节拍与扭矩验证点",
            "需新增卡扣模（金额 Unknown）",
            "★ 采购价影响取决于装配工时占比；本模型工时来自模板节拍，"
            "未实测，因此只能定性排序",
            "卡扣模具一次性投入",
            "取消紧固件后必须复核振动下的松脱与密封",
            "振动台 + 扭矩衰减测量 + 气密",
            "无", "VA", "Tooling", "工序简化", f"G{oid}", "HG0",
            "中：需振动与松动验证", "8–12 周",
            "Hypothetical Opportunity", "中"))

    # ---- 6. 铜/铝用量（只有真的有机加工净重时才出）
    metal = [p for p in parts if p["cat"] in ("HV 连接", "壳体", "冷却与热") and p["bottomup"]]
    if metal:
        out.append((
            nid(), "设计",
            f"金属件牌号与用量复核（{len(metal)} 个金属 bottom-up 件）",
            "、".join(f"{p['name'][:20]}" for p in metal[:4]) + " 等",
            "复核牌号是否过度选材（如易切削铜合金用于纯冲压件）；"
            "复核净重与加工余量是否匹配",
            "金属件成本随净重与利用率变；牌号只影响单价，"
            "过度选材不会在重量上暴露，只会在单价上暴露",
            "材料价差视牌号而定（同族 ±2 元/kg）",
            "减少一套 CNC 参数与一次换料",
            "无",
            "采购价影响小（< 3 元/台）；价值在料号精简与供应弹性",
            "无",
            "导电率/导热率/强度必须复核，牌号变更可能牵动载流与热",
            "载流温升 + 折弯成形性 + 端子压接拉拔力",
            "MAT", "无", "无", "材料替代", f"G{oid}", "HG0",
            "低：需确认电导率与成形性", "6–8 周",
            "Hypothetical Opportunity", "中"))

    return out


def ac_topic(title: str) -> str:
    """把 AC 条目标题归到一个『修正主题』键，用于产品 CSV 与通用补充区去重。

    ★ 为什么必须归一化而不是比字符串：产品 CSV 里的标题是人工写的
    （「两铝盖由『整件按铝价』改为九类材料分层计价」），通用区是模板句
    （「混合材料件不再『整件按单一材料价』计价」）——同一个修正，两种说法。
    直接比标题会漏判，表格就会出现两条语义重复的 AC。
    """
    t = (title or "").lower()
    rules = [
        ("MAT_SPLIT", ("分层", "单一材料", "混合材料", "material split")),
        ("PCBA_PROXY", ("裸板", "pcba", "pcb 面积", "bare pcb")),
        ("RATE_DOUBLE", ("叠加人工", "全包费率", "operator", "重复计价")),
        ("MAKE_BUY", ("自制", "外购", "make/buy", "purchased")),
        ("YIELD_CHAIN", ("良率", "累积串联", "yield")),
        ("STD_PART", ("紧固件", "密封标准件", "标准件", "standard part",
                      "漏项", "0 元计价", "螺母", "o 圈")),
        ("NO_PRICE", ("金额为空", "无价", "coverage", "缺数据")),
    ]
    for key, kws in rules:
        if any(kw in t for kw in kws):
            return key
    return ""


def build_ac_rows(parts: list[dict], *, leaves: list[dict], mass_rows: list[dict],
                  std_part_n: int, std_qty: float, has_attr_split: bool,
                  ev_n: int, unknown_n: int) -> list[tuple]:
    """会计修正条目。★ 一律不计入降本；金额影响只描述『从错到对』的位移。

    ★ 2026-11-04：原 AC01/AC02 是 DC-DC 专属算式（Top Cover 26.22→31.28、
    PCBA 51.80/304.05）。改为只在本产品真的存在该结构时才生成该条。
    """
    rows: list[tuple] = []
    aid = 1

    def nid() -> str:
        nonlocal aid
        aid += 1
        return f"AC{aid - 1:02d}"

    # AC-A：材料分层口径（只有当存在分层时才可能发生过该错误）
    if has_attr_split:
        rows.append((
            nid(), "混合材料件不再『整件按单一材料价』计价",
            "整件净重 × 主体材料单价",
            "按 AL / CU / STEEL / PLASTIC / ELASTOMER / FERRITE / FOAM / "
            "ADHESIVE / OTHER 九类净重各自单价；废料回收只对金属",
            "混合材料件按单一材料计价，会把辅料（胶、泡棉、弹性体，单价远高于铝）"
            "藏进主体材料价里，也会把钢质件按铝价计",
            "★ 修正后金额**上升**；具体位移见 2-BOM 材料列与 4-Summary 透视",
            "★ 这是把算错的成本算对。修正后数字上升，不存在『节省』；"
            "把口径修正计入降本收益，等于用错误基线虚构收益"))
    else:
        rows.append((
            nid(), "材料分层未建立 —— 当前材料输入为 0（待补，不是已修正）",
            "无 attrs：Material Input 列全为 0，Supplier Mfg Cost 仅含加工费",
            "应补：每件九类材料净重 + 工序路由（purchased-attributes 表）",
            "材料与加工不可分时，Level 1 合计只含加工与加价，"
            "材料部分被隐藏成 0——这不是低估一点，是整块缺失",
            "★ 补齐后材料成本会从 0 跳到真实值，Level 1 合计将**显著上升**",
            "★ 这不是降本也不是既有错误修正，而是**尚未取得的数据**；"
            "不得把它当作任何形式的收益，也不得用行业均价填位"))

    # AC-B：裸板价 vs PCBA 交付价
    pcba = [p for p in parts if p["cat"] in ("PCB / PCBA",)]
    if pcba:
        rows.append((
            nid(), "Bare PCB 面积价不得冒充 PCBA 交付价",
            "裸板面积价直接作为 PCBA 交付价",
            "裸板面积价单列为 Bare PCB Proxy；交付价 = 裸板 × PCBA 系数"
            "（已并入 Price Low/High）",
            "裸板只占 PCBA 交付价 20–40%，直接把裸板价当整板价会低估 2.5–5 倍",
            "PCBA 交付价将高于裸板价（倍数见 Price Basis 列）",
            "★ 同上，是口径修正而非节省；但它把此前被隐藏的最大成本项暴露出来"))

    # AC-C：费率叠加人工
    rows.append((
        nid(), "全包费率行不再叠加人工",
        "压铸/冲压/注塑/检漏行：机时费（含人工）+ 人数 × 人工费率",
        "全包费率行 operators = 0；裸机费率行（CNC）才计人工；纯人工行机时费 = 0",
        "同一份人工被计入两次",
        "供应商加工成本下降（按工序结构不同，约 −8…−12%）",
        "★ 是重复计价纠正；下降后的数字才是可信基线"))

    # AC-D：自制/外购分类（v6：口径本身按用户要求改为二分，此条改为声明差异）
    make_n = sum(1 for p in parts if p["bottomup"])
    rows.append((
        nid(), "v6 口径：恢复 Make/Buy 二分（与 v5 全外购口径的差异声明）",
        "v5：全部叶件按 Purchased，结构件也取供应商交付价（含 SG&A/利润/物流）",
        f"v6：{make_n} 个结构件按 Make 自制成本口径 = (材料+工序)×(1+OH)；"
        "其余按 Purchased（交付价或 Class E 行业均价）",
        "用户要求按『有自制件』的口径建模：自制件没有供应商加成链，"
        "把 SG&A/利润/物流加进自制件价格会高估",
        "★ 两版金额不可直接对比：v6 的 Make 件金额系统性低于 v5 同件金额"
        "（差 = SG&A + 利润 + 物流段）；差异量级见 4-Summary Level 1 合计对照",
        "★ 这是口径选择，不是修正也不是降本；『自制』是情景假设，"
        "拆解无法判断零件由谁制造"))

    # AC-E：良率串联
    rows.append((
        nid(), "良率由『单步除法』改为累积串联",
        "Step Cost = 该步成本 ÷ 该步良率；各步相加",
        "C_i = (直接成本_i + C_{i-1}) ÷ 良率_i；取最后一道工序累积值",
        "单步除法不体现前道损失，且全列求和会重复计入中间累积值",
        "多工序件成本上升（累积良率低于单步良率）",
        "★ 是计算逻辑纠正"))

    # AC-F：标准件漏项补齐（只有存在标准件时才出）
    if std_part_n:
        rows.append((
            nid(), f"紧固件 / 密封标准件不再以 0 元计价（{std_part_n} 料号补齐）",
            f"旧口径：{std_part_n} 个标准件挂靠位置（{std_qty:.0f} 件实物）未进 BOM，"
            f"Level 1 只含 {len(mass_rows)} 个计量叶件",
            f"现口径：{len(parts)} 行 = {len(mass_rows)} 个计量叶件 + "
            f"{std_part_n} 个标准件行（{std_qty:.0f} 件，按料号归并）",
            "叶件过滤条件是 `total_weight` 非空，而紧固件的 Total Weight 全部为空"
            "（其重量已并入父件的 Fastener Weight 列）——"
            "于是它们被当成『无重量』整行丢弃，零成本进入 BOM",
            f"Level 1 上升 {std_qty:.0f} 件对应的包金额（见 4-Summary 紧固件包块）；"
            "金额量级远小于漏项估算，因为标准件单价极低",
            "★ 这是**漏项补齐**，不是降本；它使 Level 1 更完整但也更高，"
            "同时把 Coverage 的分母从计量叶件数改为 BOM 全行数"))

    # AC-G：无价件
    if unknown_n:
        rows.append((
            nid(), f"{unknown_n} 行金额为空（不是 0）——不得按 0 参与占比",
            "旧口径把『无取证价』当作『有估算』，Coverage 报 100%",
            f"现口径：取证价 {ev_n} 项 / 无价 {unknown_n} 项单列；"
            "Coverage 只按真正有价的件计",
            "Excel 空单元格不参与求和，于是 Level 1 合计只反映已建模件，"
            "但旧口径把它当成完整合计",
            "★ Level 1 合计本身不变，但它被明确标注为**结构性低估**",
            "★ 这不是降本也不是修正已算错的数，而是把『缺数据』如实标出来；"
            "把缺失当成 0 参与占比，会让所有 Share % 失真"))

    return rows


# ================================================================ 主构建
def build(leaves: list[dict], attrs: dict, prices: dict, out: Path, product: str,
          func_prefix: str | None = None,
          src_names: dict[str, str] | None = None,
          opps_in: list[dict] | None = None,
          ac_in: list[dict] | None = None,
          catalog: dict | None = None,
          bench_fill: bool = True) -> None:
    """★ func_prefix：功能域前缀（INV/DCDC/OBC/GEN）。缺省由产品名推导，
    不再硬编码成 DCDC（旧版把某个产品的默认值写死在通用工具里）。

    ★ src_names：证据来源文件名（leaves/attrs/prices），用于 7-Evidence 表。
    缺省写成「未提供」，不再把 dcdc-*.csv 硬编码进通用工具。

    ★ opps_in / ac_in：产品专属降本机会与会计修正。**产品分析是数据，不是代码**——
    通用工具不内置任何产品的降本次目；缺省时只用 build_opps() 按 BOM 推导
    通用候选，绝不写出别的产品的零件。

    ★ catalog / bench_fill：基准价目录（Step 5.8）。当既无取证价、又未建 bottom-up
    模型时，按目录三级补位（3a 严格类别 / 3b 包络类别 / 3c 质量轴兜底）。
    `bench_fill=False` 恢复旧行为（留 Unknown、不补位），供回归对照使用。
    """
    func_prefix = func_prefix_for(product, func_prefix)
    catalog = catalog if catalog is not None else load_catalog()
    sn = src_names or {}
    src_leaves = sn.get("leaves") or "（未指定叶件清单名）"
    src_attrs = sn.get("attrs") or "（未提供 attrs）"
    src_prices = sn.get("prices") or "（未提供取证价表）"
    n_attr = sum(1 for lf in leaves if lf["part"] in attrs and attrs[lf["part"]].get("object"))
    evidence_level = ("1 级（attrs 精确路由）" if n_attr else
                      "2 级（件名+材料推断）")
    # 是否有任何材料分层（决定 7-Evidence 里 materials 列是否为 0 的说明口径）
    has_attr_split = any(any(v > 0 for v in (a.get("splits") or {}).values())
                         for a in attrs.values()) if attrs else False
    wb = Workbook()
    wb.remove(wb.active)

    # 3-ProcessCost 的列定义必须在使用它的公式之前就存在（2-BOM 的 Supplier Process 引用它）。
    p3_labels, p3_widths, P3 = make_cols([
        ("step", "Step", 6), ("part", "Part No", 15), ("proc", "Process", 22),
        ("cat", "Category", 8), ("bearer", "Cost Bearer", 10), ("site", "Site", 11),
        ("cycle", "Cycle (s)", 10), ("cycle_basis", "Cycle Basis", 12),
        ("rate", "Machine Rate", 11), ("rate_basis", "Rate Basis", 18),
        ("ops", "Operators", 9), ("labor_rate", "Labor Rate", 10),
        ("labor", "Labor Cost", 11), ("cons", "Consumable", 11),
        ("yld", "Step Yield", 9), ("direct", "Step Direct (CNY)", 13),
        ("cumul", "Cumulative Cost (CNY)", 14), ("is_last", "Is Last Step (Y/N)", 11),
    ])

    # ★ 质量口径：只有 total_weight > 0 的行进入基线。标准件的重量已在父件里，
    #   再计一次就是重复计重（父件 Total = Part Weight + Fastener Weight 已核验）。
    total_kg = sum(r["total_weight"] for r in leaves)
    mass_rows = [r for r in leaves if r["total_weight"] > 0]
    std_rows = [r for r in leaves if r["total_weight"] <= 0]
    std_qty = sum(r["qty"] for r in std_rows)
    std_kg = sum(r["fastener_weight"] for r in std_rows)
    iface_cnt = sum(r["itf"] * r["qty"] for r in leaves)
    # ★ v6 修正：attrs 九类分层是**单件口径**（= Part Weight/Qty，与 mat_formula 的
    #   单件计价一致），汇总成整机公斤数必须 ×Qty——漏乘会把多件数铜排的铜量按 1 件计。
    al_kg = sum(attrs.get(r["part"], {}).get("splits", {}).get("AL (kg)", 0.0) * r["qty"]
                for r in leaves)
    cu_kg = sum(attrs.get(r["part"], {}).get("splits", {}).get("CU (kg)", 0.0) * r["qty"]
                for r in leaves)
    # v6：铝按路由拆分——压铸件用 K_CAST、冲压/磨削件用 K_STAMP，
    # 6-Sensitivity 的铝价行必须分项计价，不能一律乘 K_CAST。
    al_kg_cast = al_kg_stamp = 0.0
    for _r in leaves:
        _at = attrs.get(_r["part"])
        if not _at:
            continue
        _a = _at["splits"].get("AL (kg)", 0.0) * _r["qty"]
        if ROUTE_K_METAL.get(_at["object"]) == "K_CAST":
            al_kg_cast += _a
        else:
            al_kg_stamp += _a

    # ---------------------------------------------------------- 1-Assumptions
    params = [list(p) for p in PARAMS]
    for pid, label, val, unit, note in (
        ("AL_KG", "铝件净重合计（九类分层 ×Qty）", round(al_kg, 3), "kg", f"由 {src_attrs} 分层重量求和"),
        ("AL_KG_CAST", "其中：压铸路由铝净重", round(al_kg_cast, 3), "kg", "×K_CAST"),
        ("AL_KG_STAMP", "其中：冲压/磨削路由铝净重", round(al_kg_stamp, 3), "kg", "×K_STAMP"),
        ("CU_KG", "铜件净重合计（九类分层 ×Qty）", round(cu_kg, 3), "kg", "同上"),
        ("IFACE_CNT", "整机界面点数（拆解求和）", iface_cnt, "点", "Number of Interface × Qty 求和"),
    ):
        params.append([pid, label, val, unit, "Calculated", note])

    ws1 = wb.create_sheet("1-Assumptions")
    ws1["A1"] = f"{product} — 假设与费率（自制 + 外购混合口径）"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = ("成本边界：结构件按 Make（自制）计价 = 材料+工序+制造间接费，不含 SG&A/利润/物流；"
                 "其余叶件按 Purchased 交付价（无取证价时 Class E 行业均价补位）；"
                 "Level 2 = 整机集成 VA，全部 Scenario；Level 3 = S0/S1/S2 三种互斥口径。")
    ws1["A2"].font = WARN_FONT
    ws1["A3"] = (f"★ v3 口径变更：BOM 由 {len(mass_rows)} 行（仅计量叶件）扩为 {len(leaves)} 行"
                 f"（+{len(std_rows)} 行紧固/密封标准件、{std_qty:.0f} 件实物）。"
                 "标准件重量已含在父件 Total Weight 内，故**不再进入质量基线**（基线仍为"
                 f"{total_kg:.3f} kg），但其成本此前为 0，v3 起按件独立计价。")
    ws1["A3"].font = WARN_FONT
    write_header(ws1, 4, ["参数ID", "参数", "值", "单位", "状态", "来源/说明"],
                 [18, 38, 14, 10, 12, 78])

    ref: dict[str, str] = {}
    r = 5
    for pid, label, value, unit, status, note in params:
        write_body(ws1, r, [pid, label, value, unit, status, note], status)
        ws1.cell(row=r, column=3).number_format = RATIO if isinstance(value, float) and abs(value) < 2 else '#,##0.####'
        ref[pid] = f"'1-Assumptions'!$C${r}"
        r += 1

    # 派生系数：金属"毛坯投入 − 废料回收"的等效倍率
    #   k = 1/Y − (1/Y − 1) × 回收率
    r += 1
    ws1.cell(row=r, column=1, value="派生系数（材料等效倍率，供 2-BOM 引用）").font = SEC_FONT
    write_header(ws1, r + 1, ["参数ID", "参数", "值", "单位", "状态", "来源/说明"],
                 [18, 38, 14, 10, 12, 78])
    rr = r + 2
    for pid, label, ysrc, with_scrap, note in (
        ("K_CAST", "压铸等效倍率（金属，含废料回收）", "Y_CAST", True,
         "压铸件材料成本 = 净重 × 单价 × K_CAST"),
        ("K_STAMP", "冲压等效倍率（金属，含废料回收）", "Y_STAMP", True,
         "冲压件材料成本 = 净重 × 单价 × K_STAMP"),
        ("K_MOLD", "注塑/塑料倍率（非金属，不减废料）", "Y_INJ", False,
         "塑料件材料成本 = 净重 ÷ 利用率 × 单价；塑料无废料回收价"),
        ("K_SINT", "烧结倍率（铁氧体，非金属）", "Y_SINT", False,
         "铁氧体材料成本 = 净重 ÷ 利用率 × 单价；磨削余量无回收价"),
        ("K_DIECUT", "模切倍率（泡棉，非金属）", "Y_DIECUT", False,
         "泡棉材料成本 = 净重 ÷ 利用率 × 单价"),
    ):
        formula = (f'=1/{ref[ysrc]}-(1/{ref[ysrc]}-1)*{ref["SCRAP_REC"]}'
                   if with_scrap else f'=1/{ref[ysrc]}')
        write_body(ws1, rr, [pid, label, formula, "-", "Calculated", note], "Calculated")
        ws1.cell(row=rr, column=3).number_format = RATIO
        ref[pid] = f"'1-Assumptions'!$C${rr}"
        rr += 1
    rr = box(ws1, rr + 1, [
        "★ 废料回收只对金属生效（铝/铜/钢）：K = 1/Y − (1/Y − 1) × 回收率。"
        "塑料、铁氧体、泡棉、胶无回收依据，K = 1/Y，不减废料。"
        "胶（ADHESIVE）直接用净重 × 单价（一次性消耗，无利用率损失）。",
        "★ 费率口径：RATE_CAST/STAMP/INJ/TEST 含人工 → 工序表 operators 强制为 0；"
        "RATE_CNC 不含人工 → 必须另加；RATE_ASSY 纯人工 → 机时费必须为 0。"
        "同一份人工只能出现一次。",
    ])
    ws1.freeze_panes = "A5"

    # ---------------------------------------------------------- 2-BOM
    ws2 = wb.create_sheet("2-BOM")
    ws2["A1"] = ("Make/Buy BOM —— 结构件 = Make（自制成本口径）；"
                 "电子件/密封件/标准件 = Purchased（交付价或 Class E 行业均价）")
    ws2["A1"].font = TITLE_FONT
    ws2["A2"] = (f"★ v3：{len(leaves)} 行 = {len(mass_rows)} 个计量叶件（{total_kg:.3f} kg）+ "
                 f"{len(std_rows)} 个标准件行（{std_qty:.0f} 件，Net Weight 列为空："
                 "其重量已计入父件，不重复计量）。标准件只贡献成本，不贡献质量。")
    ws2["A2"].font = WARN_FONT

    bom_labels, bom_widths, B = make_cols([
        ("level", "Level", 6), ("part", "Part No", 15), ("name", "Name", 30),
        ("func", "Function", 7), ("cat", "Category", 13), ("qty", "Qty", 5),
        ("group", "Material Group", 14), ("species", "Material Species", 26),
        ("netwt", "Net Weight (kg)", 11),
        ("al", "AL (kg)", 8), ("cu", "CU (kg)", 8), ("steel", "STEEL (kg)", 8),
        ("plastic", "PLASTIC (kg)", 8), ("elastomer", "ELASTOMER (kg)", 8),
        ("ferrite", "FERRITE (kg)", 8), ("foam", "FOAM (kg)", 8),
        ("adhesive", "ADHESIVE (kg)", 8), ("other", "OTHER (kg)", 8),
        ("mat_cost", "Material Input (CNY)", 12),
        ("obj", "Procurement Object", 22), ("route", "Route", 34),
        ("route_status", "Route Status", 11), ("makebuy", "Make/Buy", 10),
        ("bottomup", "Bottom-up Model (Y/N)", 13),
        ("proc_cost", "Supplier Process (CNY)", 13),
        ("mfg_cost", "Supplier Mfg Cost (CNY)", 13),
        ("sell_price", "Supplier Selling Price (CNY)", 14),
        ("route_basis", "Route Cost Basis", 26),
        ("price_src", "Price Source", 22),
        ("p_low", "Price Low (CNY)", 11), ("p_high", "Price High (CNY)", 11),
        ("p_base", "Price Base (CNY)", 12),
        ("p_class", "Price Class", 9), ("p_type", "Price Type", 20),
        ("p_status", "Price Status", 16), ("boundary", "Cost Boundary", 30),
        ("p_basis", "Price Basis", 60),
        ("bare", "Bare PCB Proxy (CNY)", 12), ("bare_basis", "Bare PCB Basis", 46),
        ("tooling", "Tooling CapEx (kCNY)", 12),
        ("amt_low", "Amt Low", 11), ("amt_high", "Amt High", 11), ("amt_base", "Amt Base", 11),
    ])
    write_header(ws2, 4, bom_labels, bom_widths)

    parts: list[dict] = []
    r = 5
    for lf in leaves:
        at = attrs.get(lf["part"], {})
        ev = prices.get(lf["part"], {})
        obj = at.get("object", "")
        route_key = OBJECT_ROUTE.get(obj)
        splits = at.get("splits") or {c: 0.0 for c in SPLIT_COLS}
        has_split = any(v > 0 for v in splits.values())
        func, cat = function_category(lf["name"], lf["species"], func_prefix)

        # ---- 工艺路由三级降级（★ 2026-11-04 修复：此前无 attrs 即完全不建模型）
        # 1) attrs 给出可识别 Procurement Object → 精确路由（证据最全）
        # 2) 无取证价（即没有替代成本来源）→ 按件名+材料推断路由，标 Inferred
        # 3) 有取证价 → 不建自下而上模型，按取证价采购（保留 DC-DC 已交付口径）
        # 4) 都没有 → 显式 Unknown，留空成本，绝不当 0
        #
        # ★ 外购电子件（PCB/MOSFET/电容/连接器/传感器/保险丝…）**永远不建工艺模型**：
        #   它们的成本是市场价，不是"某条工艺路线的材料+工时"。给它们硬套冲压/注塑模板
        #   会产出一个看起来精确、实际毫无依据的数（v1 用 is_buy() 正是这么挡的）。
        buy_candidate = is_buy(lf["name"])
        # 标准件（螺钉/螺母/垫圈/卡扣…）永远不建工艺模型：它们是目录件按件计价，
        # 套"装配"模板会得出一个没有依据的数（且标准件重量已在父件内，更不该建模）。
        standard_part = bool(lf.get("standard"))
        steps: list[dict] = []
        inferred_note = ""
        if route_key:
            route_text = ROUTES[route_key][0]
            steps = ROUTES[route_key][1]
            bottomup = bool(at.get("bottomup"))
            route_status = at.get("route_status") or "Assumed"
        elif standard_part or buy_candidate:
            # 外购件/标准件优先：不建工艺。有取证价按取证价；无取证价为 Unknown（留空）。
            # ★ 2026-11-04：route_text 不得在无价时谎称「按取证价采购」——
            #   逆变器无 price-evidence CSV 时，61 行曾全部显示「按取证价采购」而金额为 0。
            route_text = ("不建工艺（按取证价采购）" if ev
                          else "不建工艺（无价格证据：待补取证价）")
            bottomup = False
            route_status = "N/A"
        elif at:
            # 有属性表但 Object 未识别、且不在外购/标准件名单 → 不允许静默跳过，
            # 显式标 Unknown 交由人工补 Procurement Object（旧版在此静默返回"不建工艺"）。
            route_text = "未建工艺（Unknown：Procurement Object 未识别）"
            bottomup = False
            route_status = "Unknown"
        elif ev:
            # 无属性表、但有取证价 → 按取证价采购，不重复建供应商成本
            route_text = "不建工艺（按取证价采购）"
            bottomup = False
            route_status = "N/A"
        else:
            inf_text, inf_steps = route_for(lf["name"], lf["species"])
            strength = route_strength(inf_text)
            if strength == "Weak":
                # 只判到装配级 = 弱推断，仍然建模型但必须标出来
                route_text = f"推断：{inf_text}（仅装配级）"
                route_status = "Inferred-Weak"
            else:
                route_text = f"推断：{inf_text}"
                route_status = "Inferred"
            steps = to_step_dicts(inf_steps, basis="模板（按件名推断）")
            bottomup = True
            inferred_note = ("按件名+材料推断，无供应商工艺证据；"
                             "材料分层缺失，Supplier Mfg Cost 仅含加工费")

        if bottomup and not steps:
            # 防御：给了 bottomup 却拿不到工序，明确降级而不是产出空成本
            bottomup = False
            route_status = "Unknown"
            route_text = "未建工艺（Unknown：无可用路由证据）"

        if bottomup and has_split:
            pl = plastic_param(lf["species"])
            # v6：金属等效倍率按路由选择——压铸 K_CAST，冲压/磨削 K_STAMP
            k_metal = ROUTE_K_METAL.get(route_key, "K_STAMP") if route_key else "K_STAMP"
            mat_formula = (
                f'=IF(SUM(${B["al"]}{r}:${B["other"]}{r})=0,"",'
                f'${B["al"]}{r}*{ref["AL_PRICE"]}*{ref[k_metal]}'
                f'+${B["cu"]}{r}*{ref["CU_PRICE"]}*{ref["K_STAMP"]}'
                f'+${B["steel"]}{r}*{ref["ST_PRICE"]}*{ref["K_STAMP"]}'
                f'+${B["plastic"]}{r}*{ref[pl]}*{ref["K_MOLD"]}'
                f'+${B["elastomer"]}{r}*{ref["ELASTOMER_PRICE"]}'
                f'+${B["ferrite"]}{r}*{ref["FERRITE_PRICE"]}*{ref["K_SINT"]}'
                f'+${B["foam"]}{r}*{ref["FOAM_PRICE"]}*{ref["K_DIECUT"]}'
                f'+${B["adhesive"]}{r}*{ref["ADHESIVE_PRICE"]}'
                f'+${B["other"]}{r}*{ref["PLASTIC_PRICE"]}*{ref["K_MOLD"]})'
            )
        else:
            mat_formula = None

        # ★ 基准价补位结果（只在 bottomup=False 且无取证价时才会被赋值）
        bench = None
        if bottomup:
            price_source = "Bottom-up（自制成本模型）"
            p_type, p_status, p_class = "Bottom-up Estimate", "Bottom-up Estimate", "—"
            boundary = ("Self-made cost（材料+工序+制造间接费；不含 SG&A/利润/物流）")
            p_basis = (f"材料分层 + 工序路由（{route_status}）；"
                       f"见 {B['mfg_cost']} 列 + 3-ProcessCost + 1-Assumptions OH_RATE")
        elif ev:
            price_source = f"取证价（Class {ev['cls']}）" if ev.get("cls") else "取证价"
            p_type = ev.get("ptype") or "Unknown"
            p_status = PTYPE2STATUS.get(p_type, ev.get("pstatus") or "Unknown")
            p_class = ev.get("cls") or ""
            boundary = ev.get("boundary") or "Estimated delivered at assembly plant"
            p_basis = ev.get("basis") or ""
        else:
            # ---- Step 5.8 基准价补位（Class E）：每行都要有价，取消留空
            #   三级降精度，绝不跳类别：
            #     3a 严格类别（有实证锚点）      → Price Status = Comparable
            #     3b 包络类别（类别对、锚点缺）  → Price Status = Assumption
            #     3c 质量轴兜底（按净重×元/kg）  → Price Status = Assumption
            #   ★ 门控：本分支只在「无取证价」时进入；bottomup 为 True 时走上面的 if。
            bench = None
            if bench_fill and catalog:
                bench = benchmark_price(
                    lf["name"], lf["species"],
                    net_weight=lf["weight"],
                    length=lf.get("length") or 0.0,
                    width=lf.get("width") or 0.0,
                    catalog=catalog,
                    material_group=lf["group"],
                )
            if bench:
                price_source = f"基准价（Class E·{bench['tier']}）"
                p_type = "Benchmark Average"
                p_status = bench["status"]
                p_class = "E"
                boundary = "Benchmark range, not a quotation"
                p_basis = bench["basis_text"]
            else:
                price_source = "Unknown"
                p_type = p_status = "Unknown"
                p_class, boundary, p_basis = "", "", "无任何价格证据"

        row_status = STATUS2EVIDENCE.get(p_status, "Unknown")
        values = [
            3, lf["part"], lf["name"], func, cat, lf["qty"],
            lf["group"], lf["species"], round(lf["weight"], 4),
            *[(round(splits[c], 4) if splits.get(c) else None) for c in SPLIT_COLS],
            mat_formula,
            obj or "Unknown", route_text, route_status,
            "Make" if bottomup else "Purchased",
            "Y" if bottomup else "N",
            # Supplier Process：只取 3-ProcessCost 中标记为最后一道工序的累积成本
            f'=IF(COUNTIFS(\'3-ProcessCost\'!${P3["part"]}:${P3["part"]},$B{r})=0,"",'
            f'SUMIFS(\'3-ProcessCost\'!${P3["cumul"]}:${P3["cumul"]},'
            f'\'3-ProcessCost\'!${P3["part"]}:${P3["part"]},$B{r},'
            f'\'3-ProcessCost\'!${P3["is_last"]}:${P3["is_last"]},"Y"))',
            # Supplier Mfg Cost = 材料 + 工序
            f'=IF(AND(${B["mat_cost"]}{r}="",${B["proc_cost"]}{r}=""),"",'
            f'N(${B["mat_cost"]}{r})+N(${B["proc_cost"]}{r}))',
            # Supplier Selling Price（v6：仅 bottom-up=Make 件；自制口径 = 制造成本 × (1+OH)，
            # 不含 SG&A/物流/利润——自制件没有供应商加成链）
            f'=IF(${B["bottomup"]}{r}<>"Y","",IF(${B["mfg_cost"]}{r}="","",'
            f'${B["mfg_cost"]}{r}*(1+{ref["OH_RATE"]})))',
            at.get("split_basis") or ("不适用" if not bottomup else inferred_note or "不适用"),
            price_source,
            # ★ 基准价补位行：把 bench 的 Low/High 写进原有的价格列，
            #   让既有的 Price Base 公式自己决定 AVERAGE / GEOMEAN——绝不在这里写均值。
            (ev.get("low") if ev else (bench["low"] if bench else None)),
            (ev.get("high") if ev else (bench["high"] if bench else None)),
            # Price Base：Make 件用自制成本价；取证价若只有单点则用该点，否则按区间取值
            #   High/Low ≤ RANGE_MAX_RATIO → 算术均值（保留区间）；
            #   > 该值 → 几何均值（宽区间不再取算术中心，避免被高档拉动）
            f'=IF(${B["bottomup"]}{r}="Y",IF(${B["sell_price"]}{r}="","",${B["sell_price"]}{r}),'
            f'IF(${B["p_low"]}{r}="","",'
            f'IF(${B["p_high"]}{r}="",${B["p_low"]}{r},'
            f'IF(${B["p_high"]}{r}/${B["p_low"]}{r}<={ref["RANGE_MAX_RATIO"]},'
            f'AVERAGE(${B["p_low"]}{r},${B["p_high"]}{r}),'
            f'GEOMEAN(${B["p_low"]}{r},${B["p_high"]}{r})))))',
            p_class, p_type, p_status, boundary, p_basis,
            ev.get("bare"), ev.get("bare_basis") or "",
            None,  # Tooling CapEx：供应商模具费 Unknown，留空（≠ 免费模具）
            # 金额列
            f'=IF(ISNUMBER(${B["p_low"]}{r}),${B["p_low"]}{r}*${B["qty"]}{r},'
            f'IF(ISNUMBER(${B["p_base"]}{r}),${B["p_base"]}{r}*${B["qty"]}{r},""))',
            f'=IF(ISNUMBER(${B["p_high"]}{r}),${B["p_high"]}{r}*${B["qty"]}{r},'
            f'IF(ISNUMBER(${B["p_base"]}{r}),${B["p_base"]}{r}*${B["qty"]}{r},""))',
            f'=IF(ISNUMBER(${B["p_base"]}{r}),${B["p_base"]}{r}*${B["qty"]}{r},"")',
        ]
        write_body(ws2, r, values, row_status)
        ws2.cell(row=r, column=1).number_format = '0'
        for key in ("netwt", "al", "cu", "steel", "plastic", "elastomer",
                    "ferrite", "foam", "adhesive", "other"):
            ws2[f"{B[key]}{r}"].number_format = KG
        for key in ("mat_cost", "proc_cost", "mfg_cost", "sell_price", "p_low", "p_high",
                    "p_base", "bare", "amt_low", "amt_high", "amt_base"):
            ws2[f"{B[key]}{r}"].number_format = MONEY

        parts.append({
            "row": r, "part": lf["part"], "name": lf["name"], "qty": lf["qty"],
            "func": func, "cat": cat, "weight": lf["weight"], "obj": obj or "Unknown",
            "bottomup": bottomup, "route_key": route_key,
            "steps": steps, "route_status": route_status,
            "price_source": price_source, "pstatus": p_status, "pcls": p_class,
            "status": row_status, "standard": lf.get("standard", False),
            "mass": lf["total_weight"], "inferred_note": inferred_note,
            "make": bottomup,  # v6：Make = bottom-up 自制件
            # ★ 基准价三级的层级要带出去，供 4-Summary 分行披露与 RFQ 优先级排序
            "bench_tier": (bench["tier"] if bench else ""),
            "bench_key": (bench["row"]["key"] if bench else ""),
            # ★ priced 必须是「真的有价」：取证价 Low/High 至少一个有效，
            #   或 bottom-up 售价链能算出数，或基准目录给出了区间。
            #   仅 lookup 命中不算。
            "priced": (bool(ev.get("low") or ev.get("high")) or bottomup
                       or bool(bench)),
        })
        r += 1
    bom_last = r - 1
    ws2.freeze_panes = "D5"

    fn = bom_last + 2
    fn = box(ws2, fn, [
        "口径 1（v6）：Make/Buy 二分。Make = 结构件按自制成本计价："
        "价格 = (材料 + 工序) × (1 + 制造间接费率)，**不含 SG&A / 利润 / 物流**"
        "（自制件没有供应商加成链）。Purchased = 电子件/密封件/标准件按交付价；"
        "无取证价时按基准价目录 Class E 补位（行业均价，不是报价）。",
        "★ 『自制』是情景假设（按用户要求以有自制件的口径建模），不是拆解结论——"
        "拆解无法判断零件由谁制造。与 v5（全外购口径）的金额不可直接对比："
        "v5 的 Make 件金额含 SG&A/利润/物流，v6 不含。",
        "口径 2：材料为九类净重分列（单件口径 = Part Weight/Qty，不含紧固件重量——"
        "紧固件已由标准件行单独计价）。分层依据见 Route Cost Basis 列"
        "（Scenario 假设，需拆件称重核实）。整件按单一材料计价会系统性误判混合材料件。",
        "口径 3：Bare PCB Proxy 只是裸板面积价，**不得作为 PCBA 交付价**。"
        "PCBA 交付价 = 裸板 × PCBA 系数（系数已并入 Price Low/High，见取证价 CSV 的 Price Basis）。",
        "口径 4：Amt Low / Base / High 只对已有值求和。Make 件按点值处理"
        "（不确定性在 1-Assumptions 的参数里，不在价格区间里），"
        "故 Low/High 应读作『取证价部分的区间 + Make 点值』。",
        "口径 5：良率按工序累积串联 C_i=(直接成本_i + C_{i-1})/良率_i；"
        "Supplier Process 列只取该件最后一道工序的累积值（3-ProcessCost!R = Y），避免全列求和重复计。",
        "口径 6（v6）：金属材料等效倍率按路由选择——压铸件 K_CAST（利用率 0.85），"
        "冲压/磨削件 K_STAMP（利用率 0.72）。v2 对铝一律用 K_CAST，"
        "冲压铝件的材料成本被低估约 7%。",
    ])

    # ---------------------------------------------------------- 3-ProcessCost
    ws3 = wb.create_sheet("3-ProcessCost")
    ws3["A1"] = "Manufacturing Cost Breakdown —— Make 件的工序成本（自制口径）"
    ws3["A1"].font = TITLE_FONT
    ws3["A2"] = ("★ v6：本表是 Make（自制）件的内部工序成本，其合计通过 2-BOM 的"
                 "自制价格（×(1+OH)）进入 Level 1，**不得再计入 Level 2 整机集成**。"
                 "Purchased 件不进本表。")
    ws3["A2"].font = WARN_FONT
    write_header(ws3, 4, p3_labels, p3_widths)

    r = 5
    n_bottomup = n_inferred = 0
    for p in parts:
        if not p["bottomup"]:
            continue
        n_bottomup += 1
        if str(p.get("route_status", "")).startswith("Inferred"):
            n_inferred += 1
        # 工序来自三级降级的最终结果（attrs 精确 或 件名推断），不再硬依赖 route_key
        steps = p["steps"]
        w = round(p["weight"], 4)
        prev_cum: str | None = None
        for i, st in enumerate(steps, start=1):
            kind = RATE_KIND.get(st["rate"]) if st["rate"] else "outsrc"
            cycle = st["cycle"]
            if isinstance(cycle, tuple):
                cycle_v = (f'={ref["CAST_BASE_S"]}+{ref["CAST_S_PER_KG"]}*{w}'
                           if cycle[0] == "cast"
                           else f'={ref["INJ_BASE_S"]}+{ref["INJ_S_PER_KG"]}*{w}')
                cycle_basis = st["basis"] or "重量式"
            elif cycle is None:
                cycle_v, cycle_basis = None, "外协（无设备节拍）"
            else:
                cycle_v, cycle_basis = cycle, st["basis"] or "模板"

            if kind == "allin":
                machine_f, ops_v = f'={ref[st["rate"]]}', 0.0
                labor_f = 0.0
            elif kind == "machine":
                machine_f, ops_v = f'={ref[st["rate"]]}', st["ops"]
                labor_f = f'=${P3["ops"]}{r}*{ref["RATE_ASSY"]}*${P3["cycle"]}{r}/3600'
            elif kind == "labor":
                machine_f, ops_v = 0.0, st["ops"]
                labor_f = f'=${P3["ops"]}{r}*{ref["RATE_ASSY"]}*${P3["cycle"]}{r}/3600'
            else:
                machine_f, ops_v, labor_f = 0.0, 0.0, 0.0

            cons = st["cons"]
            cons_v = f'={w}*{ref[cons[1]]}' if isinstance(cons, tuple) else cons

            direct_f = (f'=${P3["cons"]}{r}' if kind == "outsrc"
                        else (f'=${P3["cycle"]}{r}/3600*${P3["rate"]}{r}'
                              f'+${P3["labor"]}{r}+${P3["cons"]}{r}'))
            cumul_f = (f'=(IF(${P3["direct"]}{r}="",0,${P3["direct"]}{r})'
                       f'+IF({prev_cum or 0}=0,0,{prev_cum or 0}))/${P3["yld"]}{r}'
                       if prev_cum else
                       f'=${P3["direct"]}{r}/${P3["yld"]}{r}')
            prev_cum = f'${P3["cumul"]}{r}'

            values = [i, p["part"], st["proc"], st["cat"], "Supplier",
                      "in-house" if kind != "outsrc" else "outsource",
                      cycle_v, cycle_basis, machine_f, RATE_KIND_LABEL[kind],
                      ops_v, f'={ref["RATE_ASSY"]}', labor_f, cons_v, st["yld"],
                      direct_f, cumul_f, "Y" if i == len(steps) else ""]
            write_body(ws3, r, values, "Estimated")
            for key in ("rate", "labor_rate", "labor", "cons", "direct", "cumul"):
                ws3[f"{P3[key]}{r}"].number_format = MONEY
            ws3[f"{P3['cycle']}{r}"].number_format = '0.0'
            ws3[f"{P3['yld']}{r}"].number_format = RATIO
            r += 1
    ws3.freeze_panes = "D5"
    p3_rows = r - 5
    fn3 = r + 1
    fn3 = box(ws3, fn3, [
        "★ 费率去重：Rate Basis =『全包（含人工）』的行 Operators 强制为 0，人工已含在机时费内；"
        "只有『裸机（不含人工）』行才计 Operators × 人工费率；『纯人工』行的机时费为 0。"
        "v1 曾在全包费率行上再次叠加人工——同一份人工被计了两次。",
    ])
    fn3 = box(ws3, fn3, [
        "★ 节拍口径：Cycle Basis =『模板』的节拍与零件重量无关（同一套工序模板套用到所有同类件），"
        "对小件偏高、对大件偏低；只有压铸/注塑按重量式估算。这是本表已知的系统性偏差，"
        "需供应商 OEE 与瓶颈节拍实测才能消除。",
        "★ 供应商工装（模具/夹具/检具）在本模型中假定已包含在供应商售价里"
        "（行业常规：模具费摊入单价或单独一次性收取）。因此 2-BOM 的 Tooling CapEx 留空，"
        "且不另行计提摊销。若供应商实际单独收取模具费，必须另开一次性成本行，"
        "**且不得同时摊入单价**——否则是同一次支出计两次。",
    ])

    # ---------------------------------------------------------- 4-Summary
    ws4 = wb.create_sheet("4-Summary")
    ws4["A1"] = "Should-Cost Summary —— Level 1 组件（Make+Purchased）/ Level 2 集成情景 / Level 3 整机口径"
    ws4["A1"].font = TITLE_FONT
    s_labels, s_widths, S = make_cols([
        ("part", "Part No", 15), ("name", "Name", 30), ("func", "Function", 11),
        ("cat", "Category", 15), ("qty", "Qty", 12), ("obj", "Procurement Object", 22),
        ("mat", "Material Input", 12), ("proc", "Supplier Process", 12),
        ("mfg", "Supplier Mfg Cost", 13), ("sell", "Supplier Selling Price", 14),
        ("src", "Price Source", 22), ("low", "Delivered Low", 11),
        ("base", "Delivered Base", 12), ("high", "Delivered High", 12),
        ("pstatus", "Price Status", 16), ("pclass", "Price Class", 9),
        ("amt_low", "Amt Low", 11), ("amt_base", "Amt Base", 12),
        ("amt_high", "Amt High", 11), ("share", "Share %", 8), ("status", "Status", 11),
        # ★ 材料/加工的**金额**列：功能域透视必须用金额口径求和。
        #   早前透视直接用 SUMPRODUCT(…*N(range)…) —— N() 不按数组展开，
        #   整列塌缩成首行值，导致四列相加 ≠ 交付价（实测闭合差 36.05）。
        ("mat_amt", "Mat Amt", 11), ("proc_amt", "Proc Amt", 11),
    ])
    write_header(ws4, 4, s_labels, s_widths)

    r = 5
    first = r
    for p in parts:
        br = p["row"]
        vals = [
            p["part"], p["name"], p["func"], p["cat"], p["qty"], p["obj"],
            f"='2-BOM'!${B['mat_cost']}{br}",
            f"='2-BOM'!${B['proc_cost']}{br}",
            f"='2-BOM'!${B['mfg_cost']}{br}",
            f"='2-BOM'!${B['sell_price']}{br}",
            p["price_source"],
            f"='2-BOM'!${B['p_low']}{br}",
            f"='2-BOM'!${B['p_base']}{br}",
            f"='2-BOM'!${B['p_high']}{br}",
            p["pstatus"], p["pcls"] or "",
            f'=IF(ISNUMBER(${S["low"]}{r}),${S["low"]}{r}*${S["qty"]}{r},'
            f'IF(ISNUMBER(${S["base"]}{r}),${S["base"]}{r}*${S["qty"]}{r},""))',
            f'=IF(ISNUMBER(${S["base"]}{r}),${S["base"]}{r}*${S["qty"]}{r},"")',
            f'=IF(ISNUMBER(${S["high"]}{r}),${S["high"]}{r}*${S["qty"]}{r},'
            f'IF(ISNUMBER(${S["base"]}{r}),${S["base"]}{r}*${S["qty"]}{r},""))',
            None,  # Share % 在合计行确定后再补
            p["status"],
            # 材料金额 = 单价 × Qty（空则空，绝不当 0）
            f'=IF(ISNUMBER(${S["mat"]}{r}),${S["mat"]}{r}*${S["qty"]}{r},"")',
            f'=IF(ISNUMBER(${S["proc"]}{r}),${S["proc"]}{r}*${S["qty"]}{r},"")',
        ]
        write_body(ws4, r, vals, p["status"])
        for key in ("mat", "proc", "mfg", "sell", "low", "base", "high",
                    "amt_low", "amt_base", "amt_high", "mat_amt", "proc_amt"):
            ws4[f"{S[key]}{r}"].number_format = MONEY
        r += 1
    last = r - 1
    total_row = r

    # 合计行
    tot: dict[str, object] = {S["part"]: "合计（Level 1 组件成本：Make + Purchased）"}
    tot[S["qty"]] = f'=SUM({S["qty"]}{first}:{S["qty"]}{last})'
    for key in ("mat", "proc", "mfg", "sell", "amt_low", "amt_base", "amt_high"):
        col = S[key]
        tot[col] = f'=IF(COUNT({col}{first}:{col}{last})=0,"",SUM({col}{first}:{col}{last}))'
    tot[S["share"]] = 1
    tot[S["status"]] = "Calculated"
    write_body(ws4, total_row, [tot.get(get_column_letter(i)) for i in range(1, len(s_labels) + 1)],
               "Calculated")
    for key in ("mat", "proc", "mfg", "sell", "amt_low", "amt_base", "amt_high", "part"):
        ws4[f"{S[key]}{total_row}"].font = BOLD_FONT
        ws4[f"{S[key]}{total_row}"].number_format = MONEY
    ws4[f"{S['share']}{total_row}"].number_format = PCT
    ws4[f"{S['share']}{total_row}"].font = BOLD_FONT

    # 补 Share %
    for rr in range(first, last + 1):
        c = ws4[f"{S['share']}{rr}"]
        c.value = f'=IF({S["amt_base"]}{rr}="","",{S["amt_base"]}{rr}/${S["amt_base"]}${total_row})'
        c.number_format = PCT
    ws4.freeze_panes = "C5"

    l1_base = f"${S['amt_base']}${total_row}"
    l1_low = f"${S['amt_low']}${total_row}"
    l1_high = f"${S['amt_high']}${total_row}"
    bu_n = sum(1 for p in parts if p["bottomup"])
    # ★ 2026-11-04：覆盖说明里的工艺族名必须按本产品实际 bottom-up 件统计。
    #   旧版硬编码「压铸/冲压/注塑/包胶端子」——那是 DC-DC 的实际结构，
    #   换到逆变器（六件冲压件套同一模板）就变成一句与数据无关的断言。
    _bu_procs = sorted({p["steps"][0]["proc"]
                        for p in parts if p["bottomup"] and p.get("steps")})
    bu_family_text = "、".join(_bu_procs) if _bu_procs else "无可用首道工序"
    # ★ 2026-11-04 修复：ev_n 原为 len(parts) - bu_n，隐含假定「非 bottom-up 必有取证价」。
    #   逆变器无 price-evidence CSV 时不成立 → 61 行零价却报 Model Coverage=100%。
    #   现在按实际是否取到价（priced）统计，并把无价件单列为 unknown_n。
    ev_n = sum(1 for p in parts if (not p["bottomup"]) and p.get("priced")
               and p.get("pcls") != "E")
    # ★ 基准价补位件数（Class E）。必须单列：它是「有价」，但绝不算
    #   Verified / Comparable——4-Summary 与 7-Evidence 都把这一层单独披露。
    bench_n = sum(1 for p in parts if p.get("pcls") == "E")
    bench_tier_n: dict[str, int] = {}
    for _p in parts:
        if _p.get("pcls") == "E":
            _t = _p.get("bench_tier") or "3?"
            bench_tier_n[_t] = bench_tier_n.get(_t, 0) + 1
    unknown_n = len(parts) - bu_n - ev_n - bench_n
    std_n = sum(1 for p in parts if p["standard"])
    std_part_n = len({p["part"] for p in parts if p["standard"]})
    rng = f"${S['amt_base']}${first}:${S['amt_base']}${last}"
    src_rng = f"${S['src']}${first}:${S['src']}${last}"
    cls_rng = f"${S['pclass']}${first}:${S['pclass']}${last}"
    bench_amt_f = f'SUMIFS({rng},{cls_rng},"E")'
    # 紧固件 / 密封件包金额：**按行号显式求和**，不靠 Procurement Object 文本。
    # ★ 2026-11-05 修复：旧版用 SUMIFS(Procurement Object="Standard fastener")，
    #   而无 attrs 时该列全为 Unknown（见上面 obj or "Unknown"）→ 包内金额恒为 0，
    #   却在同一块里声称「11 个料号 / 134 件」，自相矛盾。
    #   补位前这些行价格为空白，0 是「真无价」；现在有价了，必须真算出来。
    std_rows = [p["row"] for p in parts if p["standard"]]
    std_bundle_f = ("(" + "+".join(f"${S['amt_base']}{_rr}" for _rr in std_rows) + ")"
                    if std_rows else "0")

    # ---- Level 1
    b1 = write_block(ws4, total_row + 2,
                     "Level 1 — Component Cost（Make 自制 + Purchased 外购，交付/自制口径）", [
        ("BOM 行数（Make + Purchased）", len(parts),
         f"由 {len(mass_rows)} 个计量叶件 + {std_part_n} 个标准件行（紧固/密封，"
         f"{std_qty:.0f} 件，按料号归并）构成；父件不参与求和", "Calculated"),
        ("① Make 自制件（件数 / 金额）",
         f'=COUNTIFS({src_rng},"Bottom-up*")&" 项 / "&TEXT(SUMIFS({rng},{src_rng},"Bottom-up*"),"#,##0.00")&" 元"',
         "结构件按自制成本口径：材料+工序+制造间接费；不含 SG&A/利润/物流", "Calculated"),
        ("② 取证价（件数 / 金额）",
         f'=COUNTIFS({src_rng},"取证价*")&" 项 / "&TEXT(SUMIFS({rng},{src_rng},"取证价*"),"#,##0.00")&" 元"',
         # ★ 2026-11-04：原说明列举「透气阀/磁芯/导热膏」等 DC-DC 专属件类，
         #   通用工具里不得写死某产品的清单。改为按实际统计 + 通用定义。
         (f"取证价 = 有外部价格证据的件（非 Make 破拆）。"
          f"本产品 {ev_n} 项取证价" + ("；其余非 Make 件无价" if unknown_n else "，无价 0 项")
          if unknown_n else f"取证价 = 有外部价格证据的件（非 Make 破拆）。本产品 {ev_n} 项全部有价"),
         "Calculated"),
        ("③ Actual Procurement 已发生采购价", 0, "★ 本项目无任何合同/订单价，故为 0", "Unknown"),
        ("Level 1 合计（Base）", f"={l1_base}",
         "Make 件 = 自制成本（含 OH，不含 SG&A/利润/物流）；Purchased 件 = 交付价或 Class E 行业均价",
         "Calculated"),
        ("Level 1 合计（Low / High）",
         # ★ 2026-11-04：无取证价时 Low/High 全空，旧式 TEXT 输出「0.00 / 0.00」，
         #   读起来像一个真实区间。改为条件文本：无区间就明说「不适用」。
         (f'=IF({l1_low}=0,IF({l1_high}=0,'
          f'"不适用（无取证价 → 区间无从产生，以上 Base 为点值）",'
          f'"— / "&TEXT({l1_high},"#,##0.00")),'
          f'TEXT({l1_low},"#,##0.00")&" / "&TEXT({l1_high},"#,##0.00"))'),
         "区间只来自取证价部分；bottom-up 件为点值。全部为 bottom-up/无价时区间不成立",
         "Calculated"),
    ])

    # ---- 紧固件 / 密封件包（v3 新增：单列一块，避免被大件金额掩盖）
    std_amt = f"={std_bundle_f}"
    b1b = write_block(ws4, b1["next"],
                      f"紧固件 / 密封件包（v3 补齐；{std_part_n} 个料号 / {std_qty:.0f} 件实物）", [
        ("包内金额（Base，元）", std_amt,
         "标准紧固件 + O 圈两类之和；已含在 Level 1 合计内，不得再与之相加", "Calculated"),
        ("占 Level 1 比例", f'=IF({l1_base}=0,"",{std_amt.replace("=", "")}/{l1_base})',
         "★ 件数占比高 ≠ 成本占比高：这是本包最重要的结论", "Calculated"),
        ("包内实物件数", std_qty, "螺钉 + 螺母 + O 圈，按宽表 Quantity 求和", "Verified"),
        ("件数占 BOM 实物件数比例",
         f'=IF(SUM({S["qty"]}{first}:{S["qty"]}{last})=0,"",'
         f'{std_qty}/SUM({S["qty"]}{first}:{S["qty"]}{last}))',
         "分母 = 计量叶件件数 + 标准件件数", "Calculated"),
        ("平均单件价（Base，元/件）", f'=IF({std_qty}=0,"",'
         f'{std_amt.replace("=", "")}/{std_qty})',
         "用于判断『按件计价』是否合理；标准件单件均在 0.05–0.6 元区间，"
         "单件均值过百即说明分母或范围写错", "Calculated"),
        ("包内材料重量（kg）", round(std_kg, 4),
         "★ 该重量已包含在父件 Total Weight 内，**不再计入质量基线**；"
         f"此处仅列出以说明 {total_kg:.3f} kg 已覆盖这 {std_qty:.0f} 件", "Verified"),
        ("计重口径", "不重复计量",
         f"宽表父件的 Total Weight = Part Weight + Fastener Weight；"
         f"若标准件行再填重量，整机基线会重复计入 {std_kg:.3f} kg", "Verified"),
    ])

    # ---- Level 2
    b2 = write_block(ws4, b1b["next"],
                     "Level 2 — Final Integration VA（**Scenario，不是已确认生产成本**）", [
        ("整机集成方 / 地点", "Unknown", "★ 未取证前不得把整机总装标为『我方自制』", "Unknown"),
        ("总装人工", f"={ref['INTEG_CT_MIN']}/60*{ref['RATE_ASSY']}",
         f"{ref['INTEG_CT_MIN']} min × 45 元/人·h；节拍为 Scenario，非实测工时", "Scenario"),
        ("EOL 测试（Hi-Pot / 功能 / 气密）", f"={ref['INTEG_EOL_CT_MIN']}/60*{ref['RATE_TEST']}",
         "逐台 EOL，非抽检；无供应商工艺记录", "Scenario"),
        ("密封胶 FIPG", f"={ref['INTEG_FIPG_G']}/1000*{ref['INTEG_FIPG_PRICE']}",
         "当前 0 = 假设密封靠成型垫圈；若拆解确认有胶道须改 >0 并重算", "Scenario"),
        ("其他辅料（清洗/擦拭/扎带/标签）", f"={ref['INTEG_CONS']}", "元/台包干", "Scenario"),
        ("设备折旧（总装线）", f"={ref['INTEG_INV']}/({ref['INTEG_LIFE']}*{ref['VOL_A']})",
         "投资额 Unknown → 本项只能标 Scenario；年产量翻倍则折旧减半", "Scenario"),
    ])
    # 小计行紧接 Level 2 数据
    l2_sub_row = b2["last"] + 1
    l2_data_first = b2["first"] + 1  # 跳过『整机集成方』（文本行）
    l2_data_last = b2["last"]
    c = ws4.cell(row=l2_sub_row, column=1, value="Level 2 小计（整机集成 VA）")
    c.font = BOLD_FONT
    c.border = BORDER
    c = ws4.cell(row=l2_sub_row, column=2, value=f"=SUM(B{l2_data_first}:B{l2_data_last})")
    c.font = BOLD_FONT
    c.border = BORDER
    c.number_format = MONEY
    c.fill = STATUS_FILL["Scenario"]
    c = ws4.cell(row=l2_sub_row, column=3, value="★ Scenario 值，不是已确认生产成本")
    c.font = WARN_FONT
    c.border = BORDER
    l2_sub = f"$B${l2_sub_row}"

    # ---- Level 3
    b3 = write_block(ws4, l2_sub_row + 2,
                     "Level 3 — Complete Assembly Cost（三种**互斥**口径，按实际边界只用其一）", [
        ("S0 组件成本口径（不含整机总装）", f"={l1_base}",
         "即 Level 1。适用于『零件清单已完整、整机总装边界不在本项目内』", "Calculated"),
        ("S1 我方总装情景（Level 1 + Level 2）", f"={l1_base}+{l2_sub}",
         "Scenario：假设由我方做总装/EOL。若集成方不是我们，此口径不适用", "Scenario"),
        ("S2 整机外购（Turnkey）", None,
         "★ 必须由整机供应商报价确定。整机报价已含其内部零件采购与总装，"
         "不得再把 Level 1 与 Level 2 叠加进去——那就是三重计价", "Unknown"),
        ("S2 的件级构成", None,
         "件级 Level 1 只能作为 S2 报价的『下限参照』，不是 S2 的成本", "Unknown"),
    ])

    # ---- 成本完整度
    bu_amt = f'SUMIFS({rng},{src_rng},"Bottom-up*")'
    b4 = write_block(ws4, b3["next"], "成本完整度（Model Coverage ≠ Verified Price Coverage）", [
        ("BOM 行总数（Make + Purchased）", len(parts),
         f"由 {len(mass_rows)} 个计量叶件 + {std_part_n} 个标准件行构成", "Calculated"),
        ("Actual Procurement Price 件数", 0, "★ 无合同/订单价", "Unknown"),
        ("Supplier Quotation 件数", 0, "★ 无供应商报价", "Unknown"),
        ("Bottom-up 建模件数", bu_n, "Make 自制件（材料+工序+OH）", "Calculated"),
        ("Comparable / Parametric 建模件数", ev_n,
         (f"取证价 {ev_n} 项（类比外推 Class D，含标准件 {std_part_n} 项）"
          f"——本产品无 PCBA/导热膏类参数化件") if ev_n else
         "★ 本产品无价格取证表 → 0 项；所有非 Make 件落入 Unknown 价格件数",
         "Calculated"),
        ("Unknown 价格件数", unknown_n,
         ("无缺价件；但『有价』≠『有真实采购价』" if unknown_n == 0 else
          f"★ {unknown_n} 行无任何价格证据（金额为空，不等于 0）。"
          "需补齐 --prices 取证价表或对其建模，否则 Level 1 合计是**低估**"),
         "Calculated" if unknown_n == 0 else "Unknown"),
        ("基准价补位件数（Class E）", bench_n,
         (f"★ 其中 3a 严格类别 {bench_tier_n.get('3a', 0)} / 3b 包络类别 "
          f"{bench_tier_n.get('3b', 0)} / 3c 质量轴兜底 {bench_tier_n.get('3c', 0)}。"
          "**行业均价，不是报价**，只能判结构不能谈价（用户口径：无法算自制价的件用行业均值）")
         if bench_n else "★ 本次未触发补位（全部行已有取证价或 Make 模型）",
         "Estimated" if bench_n else "N/A"),
        ("Model Coverage（有估算 ÷ 全部）", f"=({bu_n}+{ev_n}+{bench_n})/{len(parts)}",
         "★ 有估算金额 ≠ 取得真实采购价格；分母为 BOM 全行数（v3 起含标准件行）", "Calculated"),
        ("Verified Price Coverage（Actual Procurement ÷ 全部）", f"=0/{len(parts)}",
         "★ 这才是可用于定点谈判的覆盖率。基准价补位**不计入**本行", "Unknown"),
        ("Class C+D+E 金额占比（分母 = Level 1 合计）",
         f'=IF({l1_base}=0,"",(SUMIFS({rng},{cls_rng},"C")+SUMIFS({rng},{cls_rng},"D")'
         f'+SUMIFS({rng},{cls_rng},"E"))/{l1_base})',
         "★ > 30% 时结论必须声明『外购价以估算为主』", "Calculated"),
        ("其中 Class E 基准价金额占比", f'=IF({l1_base}=0,"",{bench_amt_f}/{l1_base})',
         "★ 这部分金额的置信度最低：件数少但金额可占比很高（大件走质量轴兜底）",
         "Estimated" if bench_n else "N/A"),
        ("紧固件包金额占比（分母 = Level 1 合计）",
         f'=IF({l1_base}=0,"",{std_bundle_f}/{l1_base})',
         f"★ {std_part_n} 个料号 / {std_qty:.0f} 件只占极小比例——"
         "件数多不等于钱多", "Calculated"),
    ])
    ws4[f"B{b4['marks']['Model Coverage（有估算 ÷ 全部）']}"].number_format = PCT
    ws4[f"B{b4['marks']['Verified Price Coverage（Actual Procurement ÷ 全部）']}"].number_format = PCT
    if "Class C+D+E 金额占比（分母 = Level 1 合计）" in b4["marks"]:
        ws4[f"B{b4['marks']['Class C+D+E 金额占比（分母 = Level 1 合计）']}"].number_format = PCT
    if "其中 Class E 基准价金额占比" in b4["marks"]:
        ws4[f"B{b4['marks']['其中 Class E 基准价金额占比']}"].number_format = PCT

    # ---- 基准价补位（Class E）三级披露
    #   ★ 三行必须分开：3a 是「类别对且锚点实证」，3b 是「类别对但只有包络」，
    #     3c 是「连类别都没有，按质量轴兜底」。合并成一个总数会让读者
    #     把「有实证的 PCBA 面积价」和「按金属 18–60 元/kg 兜底的支架」
    #     当成同一个置信度。
    if bench_n:
        b5 = write_block(ws4, b4["next"], "基准价补位（Class E）—— 行业均价，不是报价", [
            ("3a 严格类别（类别对 + 实证锚点）件数", bench_tier_n.get("3a", 0),
             "类别名逐字匹配，锚点有来源与日期；置信度等同 Class D 类比外推", "Estimated"),
            ("3a 金额（Base，元）",
             f'=SUMIFS({rng},{cls_rng},"E",{src_rng},"基准价（Class E·3a）*")',
             "★ 只统计 3a 行，不含 3b/3c", "Estimated"),
            ("3a 占 Level 1 比例", None, "与 3b/3c 同一分母（Level 1 合计）", "Estimated"),
        ])
        # 上面第 3 行需要引用本块内两个金额单元格，写回真实引用
        _r3a_amt = b5["marks"]["3a 金额（Base，元）"]
        _r3a_share = b5["marks"]["3a 占 Level 1 比例"]
        ws4[f"B{_r3a_share}"] = (f'=IF({l1_base}=0,"",'
                                 f'IF(B{_r3a_amt}="","",B{_r3a_amt}/{l1_base}))')
        ws4[f"B{_r3a_amt}"].number_format = MONEY
        ws4[f"B{_r3a_share}"].number_format = PCT

        b6 = write_block(ws4, b5["next"], "基准价补位（Class E）—— 3b 包络类别", [
            ("3b 包络类别（类别对、锚点缺）件数", bench_tier_n.get("3b", 0),
             "★ 区间为行业最宽可信包络（如 MOSFET 5–400 元/件 = 80 倍宽）。"
             "类别没错，精度很低——**这一档最该优先 RFQ**", "Assumption"),
            ("3b 金额（Base，元）",
             f'=SUMIFS({rng},{cls_rng},"E",{src_rng},"基准价（Class E·3b）*")',
             "★ 只统计 3b 行", "Assumption"),
            ("3b 占 Level 1 比例", None, "与 3a 同一分母（Level 1 合计）", "Assumption"),
        ])
        _r3b_amt = b6["marks"]["3b 金额（Base，元）"]
        _r3b_share = b6["marks"]["3b 占 Level 1 比例"]
        ws4[f"B{_r3b_share}"] = (f'=IF({l1_base}=0,"",'
                                 f'IF(B{_r3b_amt}="","",B{_r3b_amt}/{l1_base}))')
        ws4[f"B{_r3b_amt}"].number_format = MONEY
        ws4[f"B{_r3b_share}"].number_format = PCT

        b7 = write_block(ws4, b6["next"], "基准价补位（Class E）—— 3c 质量轴兜底", [
            ("3c 质量轴兜底（无类别匹配，按净重 × 元/kg）件数", bench_tier_n.get("3c", 0),
             "★ 仅用于结构件（金属/塑料/弹性体/泡棉），电子件一律不许走本档。"
             "价格由该件自身净重推出，不是同类别均价", "Assumption"),
            ("3c 金额（Base，元）",
             f'=SUMIFS({rng},{cls_rng},"E",{src_rng},"基准价（Class E·3c）*")',
             "★ 只统计 3c 行", "Assumption"),
            ("3c 占 Level 1 比例", None, "与 3a 同一分母（Level 1 合计）", "Assumption"),
        ])
        _r3c_amt = b7["marks"]["3c 金额（Base，元）"]
        _r3c_share = b7["marks"]["3c 占 Level 1 比例"]
        ws4[f"B{_r3c_share}"] = (f'=IF({l1_base}=0,"",'
                                 f'IF(B{_r3c_amt}="","",B{_r3c_amt}/{l1_base}))')
        ws4[f"B{_r3c_amt}"].number_format = MONEY
        ws4[f"B{_r3c_share}"].number_format = PCT

        b8 = write_block(ws4, b7["next"], "基准价补位（Class E）—— 口径声明", [
            ("补位前 Unknown 价格件数", unknown_n + bench_n,
             "即未补位时需要留空的行数", "Calculated"),
            ("补位后 Unknown 价格件数", unknown_n,
             f"★ {unknown_n} = 0 只表示『每行都有价』，**不表示成本已验证**。"
             "Verified Price Coverage 仍为 0", "Calculated"),
            ("三级合计金额（Base，元）",
             f'=SUMIFS({rng},{cls_rng},"E")',
             "★ 此数已在 Level 1 合计内，**不得再加一次**（否则重复计价）", "Estimated"),
            ("三级合计占 Level 1 比例", f'=IF({l1_base}=0,"",'
             + f'SUMIFS({rng},{cls_rng},"E")/{l1_base})',
             "★ 该比例越低，Level 1 越可信；> 40% 时 Level 1 只能当结构参照", "Estimated"),
        ])
        ws4[f"B{b8['marks']['三级合计金额（Base，元）']}"].number_format = MONEY
        ws4[f"B{b8['marks']['三级合计占 Level 1 比例']}"].number_format = PCT
        bench_blk_next = b8["next"]
    else:
        bench_blk_next = b4["next"]
    ws4[f"B{b4['marks']['紧固件包金额占比（分母 = Level 1 合计）']}"].number_format = PCT
    for _k in ("占 Level 1 比例", "件数占 BOM 实物件数比例"):
        if _k in b1b["marks"]:
            ws4[f"B{b1b['marks'][_k]}"].number_format = PCT
    for _k in ("包内金额（Base，元）", "平均单件价（Base，元/件）"):
        if _k in b1b["marks"]:
            ws4[f"B{b1b['marks'][_k]}"].number_format = MONEY
    if "占 Level 1 比例" in b1b["marks"]:
        _v = ws4[f"B{b1b['marks']['占 Level 1 比例']}"].value
        if isinstance(_v, str) and _v.startswith("="):
            ws4[f"B{b1b['marks']['占 Level 1 比例']}"].value = "=" + _v[1:]

    # ---- 功能域透视
    piv0 = bench_blk_next
    ws4.cell(row=piv0, column=1, value="功能域透视（材料输入 / 加工 / 加价 / 交付价）").font = SEC_FONT
    ws4.cell(row=piv0, column=2, value=(
        "★ 全部按**金额**口径（单价 × Qty），四列可相加；材料+加工+加价 = 交付价。"
        "v6：Make 件的『加价』= 制造间接费 OH（无 SG&A/利润/物流）；"
        "Purchased 取证价件不可拆，其材料/加工留空，全部落在『加价』列"
        "（= 交付价 − 供应商制造成本）。"
    )).font = NOTE_FONT
    piv_hdr = ["Function", "Category", "材料输入", "供应商加工", "供应商加价",
               "交付价（Base）", "Share %"]
    for i, h in enumerate(piv_hdr, start=1):
        c = ws4.cell(row=piv0 + 1, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")

    combos: list[tuple[str, str]] = []
    for p in parts:
        k = (p["func"], p["cat"])
        if k not in combos:
            combos.append(k)
    pr = piv0 + 2
    piv_first = pr
    for func, cat in combos:
        cond = (f'${S["func"]}${first}:${S["func"]}${last},"{func}",'
                f'${S["cat"]}${first}:${S["cat"]}${last},"{cat}"')
        # 金额 = 单价 × Qty（已在上面的 mat_amt / proc_amt 列算好）；
        # 用 SUMIFS 取金额列，不用 SUMPRODUCT——N() 不按数组展开会让整列塌缩成首行值。
        vals = [
            func, cat,
            f'=SUMIFS(${S["mat_amt"]}${first}:${S["mat_amt"]}${last},{cond})',
            f'=SUMIFS(${S["proc_amt"]}${first}:${S["proc_amt"]}${last},{cond})',
            None,  # 供应商加价 = 交付价 − 材料 − 加工（下行统一补）
            f'=SUMIFS(${S["amt_base"]}${first}:${S["amt_base"]}${last},{cond})',
            None,
        ]
        write_body(ws4, pr, vals, "Calculated")
        for i in (3, 4, 5, 6):
            ws4.cell(row=pr, column=i).number_format = MONEY
        pr += 1
    piv_last = pr - 1
    piv_tot = pr
    # 供应商加价 = 交付价 − 材料 − 加工（逐行，保证四列相加恒等于交付价）
    for rr in range(piv_first, piv_last + 1):
        c = ws4.cell(row=rr, column=5,
                     value=f'=IF(F{rr}=0,"",F{rr}-N(C{rr})-N(D{rr}))')
        c.number_format = MONEY
        c.border = BORDER
    ws4.cell(row=piv_tot, column=1, value="合计").font = BOLD_FONT
    ws4.cell(row=piv_tot, column=1).border = BORDER
    for i in (3, 4, 5, 6):
        c = ws4.cell(row=piv_tot, column=i,
                     value=f"=SUM({get_column_letter(i)}{piv_first}:{get_column_letter(i)}{piv_last})")
        c.font = BOLD_FONT
        c.border = BORDER
        c.number_format = MONEY
        c.fill = STATUS_FILL["Calculated"]
    for rr in range(piv_first, piv_last + 1):
        c = ws4[f"G{rr}"]
        c.value = f'=IF(F{rr}=0,"",F{rr}/F{piv_tot})'
        c.number_format = PCT
    c = ws4.cell(row=piv_tot, column=7, value=1)
    c.font = BOLD_FONT
    c.border = BORDER
    c.number_format = PCT
    c.fill = STATUS_FILL["Calculated"]
    ws4.cell(row=piv_tot + 1, column=1,
             value=("★ 校验 1：末行 Share % 合计必须 = 100%。"
                    "★ 校验 2：材料 + 加工 + 加价 必须 = 交付价（四列同为金额口径）——"
                    "若不等，说明单价与金额混算，不是口径差异而是错误。")
             ).font = NOTE_FONT
    ws4.cell(row=piv_tot + 2, column=1,
             value="闭合差（材料+加工+加价−交付价，应恒为 0）").font = BOLD_FONT
    c = ws4.cell(row=piv_tot + 2, column=6,
                 value=(f"=ROUND(SUM(C{piv_first}:C{piv_last})"
                        f"+SUM(D{piv_first}:D{piv_last})"
                        f"+SUM(E{piv_first}:E{piv_last})"
                        f"-SUM(F{piv_first}:F{piv_last}),6)"))
    c.font = BOLD_FONT
    c.number_format = MONEY
    c.fill = STATUS_FILL["Calculated"]

    # ---- Quotation 对比
    q0 = piv_tot + 3
    ws4.cell(row=q0, column=1, value="Quotation 对比（拿到零件报价或整机报价后填入；空 = Unknown）").font = SEC_FONT
    for i, h in enumerate(["项目", "金额", "说明"], start=1):
        c = ws4.cell(row=q0 + 1, column=i, value=_text(h))
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    q_rows = [
        (f"整机供应商报价（输入）", None, "Turnkey 口径；填入后 S2 才有值", "Unknown"),
        ("S1 模型值（Level 1 + 我方总装情景）", f"={l1_base}+{l2_sub}", "Scenario 口径", "Scenario"),
        ("Variance（报价 − 模型）",
         f'=IF(B{q0 + 2}="","",B{q0 + 2}-B{q0 + 3})', "", "Calculated"),
        ("Variance %（对报价）",
         f'=IF(B{q0 + 2}="","",(B{q0 + 2}-B{q0 + 3})/B{q0 + 2})', "", "Calculated"),
    ]
    rr = q0 + 2
    for label, value, note_txt, status in q_rows:
        c1 = ws4.cell(row=rr, column=1, value=_text(label))
        c1.font = BODY_FONT
        c1.border = BORDER
        c2 = ws4.cell(row=rr, column=2, value=value)
        c2.font = BODY_FONT
        c2.border = BORDER
        c2.number_format = PCT if "%" in label else MONEY
        if STATUS_FILL.get(status):
            c2.fill = STATUS_FILL[status]
        c3 = ws4.cell(row=rr, column=3, value=note_txt)
        c3.font = NOTE_FONT
        c3.border = BORDER
        rr += 1

    # ---- 目标价反推
    t0 = rr + 1
    ws4.cell(row=t0, column=1, value="目标价反推（Top-down）：整机目标价 → Level 1 允许合计").font = SEC_FONT
    ws4.cell(row=t0 + 1, column=1, value=(
        "★ 反推得到的是『全部外购件合计允许上限』，不是任何单件的采购价。"
        "把合计上限按比例拆到各件，等于编造价格。"
    )).font = WARN_FONT
    hdr_row = t0 + 3
    for i, h in enumerate(["项目", "Low", "Base", "High"], start=1):
        c = ws4.cell(row=hdr_row, column=i, value=_text(h))
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    price_row = hdr_row + 1
    l2_plain = l2_sub
    # ★ 目标价未输入时必须显示空白，不能算出 −67.82（v3 首跑就是负数摆在那里）。
    defs = [
        ("整机目标单价（输入）", None, False),
        ("扣除 Level 2 集成 VA（Scenario，不随目标价变化）",
         '=IF({c}{r}="","",{c}{r}-$B$' + str(l2_sub_row) + ')', False),
        ("Level 1 允许合计（整机目标价 − Level 2）",
         '=IF({c}{r}="","",{c}{r1})', False),
        ("当前 Level 1（模型 Base）", f"={l1_base}", False),
        ("余量（允许 − 当前）", '=IF({c}{r}="","",{c}{r2}-{c}{r3})', False),
        ("余量占当前 Level 1", '=IF({c}{r}="","",{c}{r4}/{c}{r3})', True),
    ]
    for i, (label, tpl, is_pct) in enumerate(defs):
        row = price_row + i
        c1 = ws4.cell(row=row, column=1, value=_text(label))
        c1.font = BOLD_FONT if i == 0 else BODY_FONT
        c1.border = BORDER
        for ci, cl in zip((2, 3, 4), ("B", "C", "D")):
            if tpl is None:
                c = ws4.cell(row=row, column=ci)
                c.fill = STATUS_FILL["Unknown"]
            else:
                c = ws4.cell(row=row, column=ci, value=tpl.format(
                    c=cl, r=price_row, r1=price_row + 1, r2=price_row + 2,
                    r3=price_row + 3, r4=price_row + 4))
                c.fill = STATUS_FILL["Calculated"]
                c.number_format = PCT if is_pct else MONEY
            c.font = BODY_FONT
            c.border = BORDER
    judge = price_row + len(defs)
    ws4.cell(row=judge, column=1, value="判定").font = BOLD_FONT
    for ci, cl in zip((2, 3, 4), ("B", "C", "D")):
        c = ws4.cell(row=judge, column=ci, value=(
            f'=IF({cl}{price_row}="","待输入",'
            f'IF(ISNUMBER({cl}{price_row + 4})=FALSE,"待输入",'
            f'IF({cl}{price_row + 4}<0,"不可行：目标价低于当前 Level 1",'
            f'"可行：Level 1 需 ≤ "&TEXT({cl}{price_row + 2},"#,##0"))))'))
        c.font = BODY_FONT
        c.border = BORDER
        c.fill = STATUS_FILL["Calculated"]

    # ---------------------------------------------------------- 5-Opportunities
    ws5 = wb.create_sheet("5-Opportunities")
    ws5["A1"] = "Cost Reduction Opportunities —— 会计修正与真实降本严格分开"
    ws5["A1"].font = TITLE_FONT
    o_labels, o_widths, _ = make_cols([
        ("id", "ID", 6), ("kind", "类别", 10), ("title", "题目", 26),
        ("now", "现有方案", 32), ("alt", "替代方案", 30), ("mech", "工程机理", 36),
        ("s_mat", "供应商材料成本影响", 16), ("s_mfg", "供应商制造成本影响", 16),
        ("s_tool", "供应商工装影响", 14), ("price", "预计采购价影响", 18),
        ("new", "新增成本", 12), ("trade", "工程权衡", 30), ("valid", "验证要求", 32),
        ("mat_imp", "MAT Impact", 11), ("va_imp", "VA Impact", 11),
        ("tool_imp", "Tooling Impact", 11), ("mech_tag", "降本机制", 14),
        ("dedup", "防重复计数组", 12), ("hg", "HG", 6), ("risk", "风险", 22),
        ("weeks", "验证周期", 10), ("status", "状态", 22), ("prio", "优先级", 8),
    ])
    write_header(ws5, 4, o_labels, o_widths)

    # ★ 2026-11-04：原 opps 是整块写死的 DC-DC 分析（Top Cover 1.097 kg、
    #   TPV/TPS 密封圈、Se-Cu58 铜排、PCB 2 290×233…）。通用工具里挂某产品的
    #   具体零件，会在别的产品工作簿里讲**不存在的件**——逆变器 BOM 无 Top Cover、
    #   无磁芯、无 TPV 密封圈。
    #   现在：给了 --opps 就用该产品自己的分析；没给则按 BOM 推导通用候选。
    opps = build_opps(parts, ev_n=ev_n, unknown_n=unknown_n,
                      has_attr_split=has_attr_split, std_part_n=std_part_n,
                      std_qty=std_qty, mass_rows=mass_rows)
    if opps_in:
        # ★ 产品条目是 dict，必须按列序转成 tuple 再写表——
        #   直接 list(dict) 会得到键名（id/kind/title），整表变成列名。
        ORDER = ("id", "kind", "title", "now", "alt", "mech", "s_mat", "s_mfg",
                 "s_tool", "price", "new", "trade", "valid", "mat_imp", "va_imp",
                 "tool_imp", "mech_tag", "dedup", "hg", "risk", "weeks",
                 "status", "prio")
        prod = [tuple(d.get(k, "") for k in ORDER) for d in opps_in]
        # ★ 2026-11-04：给了产品机会表（--opps）时，**不再追加通用候选**。
        #   通用候选是按 BOM 自动推导的（紧固件精简 / 界面点数复核 / 金属件牌号复核），
        #   它们与人工分析常常主题重叠（如 O09『金属件牌号与用量复核』≈ O04
        #   『Se-Cu58 与 C10300 统一牌号』），且分析师从未审阅过——
        #   追加进去会让工作簿的 O 条目数多于报告，报告与工作簿对不上。
        #   产品机会表视为该产品 O 区的**完备集合**；通用候选只用于没给表的场合。
        opps = prod

    r = 5
    for row_v in opps:
        write_body(ws5, r, list(row_v), "Estimated")
        r += 1

    # ---- Accounting Correction 区
    r += 1
    ws5.cell(row=r, column=1,
             value="Accounting Correction（成本核算修正，**不计入 Cost Saving**）").font = SEC_FONT
    r += 1
    ac_hdr = ["ID", "题目", "原口径（错）", "修正后口径", "修正原因", "金额影响", "为什么不算降本"]
    for i, h in enumerate(ac_hdr, start=1):
        c = ws5.cell(row=r, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    # ★ 2026-11-04：AC01/AC02 原是 DC-DC 专属算式（Top Cover 26.22→31.28、
    #   PCBA 51.80/304.05）。通用工具不得携带某产品的具体数字。
    #   改为「按本产品实际统计」的条件条目：只有真的存在该结构时才出这一行。
    ac_rows = build_ac_rows(parts, leaves=leaves, mass_rows=mass_rows,
                            std_part_n=std_part_n, std_qty=std_qty,
                            has_attr_split=has_attr_split,
                            ev_n=ev_n, unknown_n=unknown_n)
    if ac_in:
        # ★ 产品 AC 在前、通用补充在后，且必须重新编号——否则两区都从 AC01 起，
        #   同一张表出现两个 AC01/AC02…（旧版实测就是这样）。
        prod_ac = [(a["id"], a["title"], a["before"], a["after"], a["reason"],
                    a["amt"], a["why"]) for a in ac_in]
        # ★ 2026-11-04：产品 CSV 常常已经写了与通用补充同主题的条目
        #   （如「分层计价」「裸板价 ≠ PCBA 价」「费率叠加人工」「良率串联」
        #   「自制/外购二分」「标准件漏项」）。此前只改编号，结果同一张表
        #   出现两组语义相同的行（AC01–AC06 与 AC07–AC12 讲同一件事），
        #   读者会以为有 12 项独立修正。
        #   现按『修正主题』去重：产品 CSV 已声明的主题，通用区不再重复。
        prod_topics = {ac_topic(r[1]) for r in prod_ac} - {""}
        extra_ac = []
        mx = max((int(i[2:]) for i in {r[0] for r in prod_ac}
                  if i[:2] == "AC" and i[2:].isdigit()), default=0)
        for r0 in ac_rows:
            if ac_topic(r0[1]) in prod_topics:
                continue
            mx += 1
            rl = list(r0)
            rl[0] = f"AC{mx:02d}"
            extra_ac.append(tuple(rl))
        ac_rows = prod_ac + extra_ac

    for row_v in ac_rows:
        write_body(ws5, r, list(row_v), "Estimated")
        r += 1

    r += 1
    r = box(ws5, r, [
        "防重复计数规则：①『防重复计数组』相同的措施不得在合计中重复计节省；"
        "② 同一降本机制不得同时计入 MAT 与 VA（主归属二选一）；"
        "③ 供应商制造成本下降 ≠ OEM 采购价等额下降——本表『预计采购价影响』列已按"
        "1/(1−利润率) 与 OH/SG&A 链放大（★ v6：该链只适用于 Purchased 件；"
        "Make 自制件无加成链，措施若落在 Make 件上，成本影响即价格影响）；"
        "但能否传导取决于议价与合同条款；"
        "④ 会计修正（AC 区）一律不计入降本合计。",
        "HG 定义（Hardness Grade）：HG0 描述想法 / HG1 毛节省潜力 / HG2 净节省潜力 / "
        "HG3 时间表已排 / HG4 已实施 / HG5 节省已兑现。"
        "HG 只记实施成熟度，与『状态』列是两个独立坐标。",
        "★ 本表全部条目为 Hypothetical Opportunity：无 Baseline 采购价、无供应商报价、"
        "无实施动作。不得把任何一条写成 HG4/HG5 或 Verified Saving。",
    ])

    # ---------------------------------------------------------- 6-Sensitivity
    ws6 = wb.create_sheet("6-Sensitivity")
    ws6["A1"] = "Sensitivity —— 按真实成本驱动排序，不做整体百分比"
    ws6["A1"].font = TITLE_FONT
    write_header(ws6, 4, ["驱动因素", "Low", "Base", "High",
                          "成本影响（元，边界差绝对值）", "占 Level 1 比例", "说明"],
                 [30, 11, 11, 11, 24, 13, 60])
    # ★ 这两个 SUMIFS 在 6-Sensitivity 里被使用，范围列都在 4-Summary，必须带表名。
    P4 = "'4-Summary'!"
    # v6：本产品无 PCBA 对象 → PCBA 行恒为 0，改为 Class E 基准价杠杆（用户口径：
    # 无法算自制价的件用行业均值，其不确定性必须披露）。
    # ★ SUMIFS 参数顺序：sum_range 在前，criteria_range 在后。
    #   旧版把字符串 "E" 写在 criteria_range1 位置 → Excel 判定文件非法、拒绝打开
    #   （最小复现：=SUMIFS(rng,"E") 必拒；正确顺序 + ">0" 条件合法）。
    bench_amt = f'SUMIFS({P4}${S["amt_base"]}${first}:${S["amt_base"]}${last},' \
                f'{P4}${S["pclass"]}${first}:${S["pclass"]}${last},"E")'
    proc_dc = f'SUMIFS({P4}${S["proc"]}${first}:${S["proc"]}${last},' \
              f'{P4}${S["obj"]}${first}:${S["obj"]}${last},"Die-cast*")'
    yield_rng = f'{P4}${S["proc"]}${first}:${S["proc"]}${last}'
    obj_rng = f'{P4}${S["obj"]}${first}:${S["obj"]}${last}'
    src_rng4 = f'{P4}${S["src"]}${first}:${S["src"]}${last}'
    amt_rng4 = f'{P4}${S["amt_base"]}${first}:${S["amt_base"]}${last}'
    # v6：Make 件售价 = 成本 × (1+OH)，无利润率 → 利润率行失效，改为 OH_RATE 行。
    bu_sell = f'SUMIFS({amt_rng4},{src_rng4},"Bottom-up*")'
    # ★ 跨表引用必须带表名：只写 $R$26 会被 Excel 当成 6-Sensitivity 自己的 R26（空），
    #   更要紧的是 SUMIFS 只给一个参数，Excel 会直接判定文件非法、拒绝打开。
    l1_on6 = f"'4-Summary'!{l1_base}"
    sens = [
        # 每行 E 列 = 边界两端成本差的**绝对值**（方向写在说明里），
        # 不用有符号差，是因为驱动方向不统一（利润率↑→价↑，良率↑→价↓），
        # 混符号会使『按影响大小排序』失效。
        # v6：铝价按路由拆两项——压铸 ×K_CAST、冲压/磨削 ×K_STAMP（口径 6）。
        ("铝价 ±15%（压铸路由）", 0.85, 1.0, 1.15,
         f"={ref['AL_KG_CAST']}*{ref['AL_PRICE']}*{ref['K_CAST']}*(1.15-0.85)",
         "只影响 Make 件的压铸铝材料输入（Housing/Bottom Cover）。铝价↑ → 成本↑"),
        ("铝价 ±15%（冲压/磨削路由）", 0.85, 1.0, 1.15,
         f"={ref['AL_KG_STAMP']}*{ref['AL_PRICE']}*{ref['K_STAMP']}*(1.15-0.85)",
         "只影响 Make 件的冲压铝材料输入（Access/Top Cover 等）。铝价↑ → 成本↑"),
        ("铜价 ±30%", 0.70, 1.0, 1.30,
         f"={ref['CU_KG']}*{ref['CU_PRICE']}*{ref['K_STAMP']}*(1.30-0.70)",
         "铜排与端子；铜净重见 1-Assumptions CU_KG（×Qty 口径）。铜价↑ → 成本↑"),
        ("Class E 基准价 ±30%", 0.70, 1.0, 1.30,
         f"={bench_amt}*(1.30-0.70)",
         "★ 无取证价且无法自制建模的件按行业均价补位（用户口径）；"
         "该行是这批补位件的不确定性披露。倍数↑ → 成本↑"),
        ("压铸件工序成本 ±20%", 0.80, 1.0, 1.20,
         f"={proc_dc}*(1.20-0.80)",
         "仅压铸件工序成本（不含材料）；注塑件量小可忽略。工序成本↑ → 成本↑"),
        ("制造间接费率 OH 8% → 18%", 0.08, 0.12, 0.18,
         f"={bu_sell}/(1+{ref['OH_RATE']})*(0.18-0.08)",
         "v6：Make 件售价 = 成本 × (1+OH)，无利润率 → 利润率行失效，"
         "改为 OH 行：Δ价 = 成本 × ΔOH。OH↑ → 成本↑"),
        ("累积良率 90% → 100%", 0.90, 0.95, 1.00,
         f'=(SUMIFS({yield_rng},{obj_rng},"Die-cast*")'
         f'+SUMIFS({yield_rng},{obj_rng},"Stamped*")'
         f'+SUMIFS({yield_rng},{obj_rng},"Machined*")'
         f'+SUMIFS({yield_rng},{obj_rng},"Injection molding")'
         f'+SUMIFS({yield_rng},{obj_rng},"Overmolded*"))'
         f"*0.95*(1/0.90-1/1.00)",
         "良率按工序乘性串联，非简单加成；只作用于 Make 件。"
         "★ 良率↑ → 成本↓（与上面几行方向相反）"),
        ("年产量 0.5× / 2×", 0.5, 1.0, 2.0,
         f"=IF({ref['VOL_A']}=0,\"\",{ref['INTEG_INV']}/"
         f"({ref['INTEG_LIFE']}*{ref['VOL_A']})*(2-0.5))",
         "只影响集成设备折旧/件；供应商工装摊销未单列（假定含在售价内）。"
         "产量↑ → 单件折旧↓"),
    ]
    rr = 5
    for label, lo, base_v, hi, formula, note in sens:
        write_body(ws6, rr, [label, lo, base_v, hi, formula,
                             f'=IF(E{rr}="","",E{rr}/{l1_on6})', note],
                   "Calculated", literal={1, 7})
        ws6.cell(row=rr, column=5).number_format = MONEY
        ws6.cell(row=rr, column=6).number_format = PCT
        rr += 1
    rr = box(ws6, rr + 1, [
        "读法 1：影响额 = 该驱动在 Low→High 两端造成的成本差**绝对值**，只用于排序"
        "『先查哪个数据』，不作报价依据。方向见说明列：铝/铜价、Class E 倍数、"
        "工序成本、OH 率、产量为『因子↑ → 成本↑』；良率为『良率↑ → 成本↓』。",
        "读法 2：第一顺位按 E 列金额从大到小读（本表生成后以 E 列实际值为准，"
        "不预设哪一行最大）。本表 E 列一律为**绝对值**，方向见说明列；"
        "不要用 E 列的符号判断增减。",
    ])

    # ---- 杠杆阶梯（以当前模型为锚，k 为 Class E 基准价金额倍数）
    l0 = rr + 1
    ws6.cell(row=l0, column=1, value="Level 1 杠杆：Class E 基准价金额倍数 → 整机口径").font = SEC_FONT
    ws6.cell(row=l0 + 1, column=1, value=(
        "k = 1.0 即当前模型（Class E 补位件 = Amt Base 之和）。"
        "该表是自洽的：k=1 行必须与 4-Summary 的 Level 1 合计相等。"
    )).font = NOTE_FONT
    for i, h in enumerate(["k（Class E 金额倍数）", "Class E 补位件（元）",
                           "Level 1（元）", "S1 情景（元）"], start=1):
        c = ws6.cell(row=l0 + 3, column=i, value=_text(h))
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    l2_on4 = f"'4-Summary'!{l2_sub}"
    for j, k in enumerate([0.5, 0.75, 1.0, 1.25, 1.5, 2.0]):
        row = l0 + 4 + j
        vals = [
            k,
            f"={bench_amt}*A{row}",
            f"={l1_on6}-{bench_amt}+{bench_amt}*A{row}",
            f"=C{row}+{l2_on4}",
        ]
        for i, v in enumerate(vals, start=1):
            c = ws6.cell(row=row, column=i, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            if i == 1:
                c.number_format = '0.00'
            else:
                c.number_format = MONEY
                c.fill = STATUS_FILL["Calculated"]
    ws6.cell(row=l0 + 11, column=1, value=(
        "★ k 只缩放 Class E 补位件（行业均价），不动 Make 自制件与取证价件。"
        "k=1 行必须等于 4-Summary 的 Level 1 合计——不等即口径混算。"
    )).font = WARN_FONT

    # ---------------------------------------------------------- 7-Evidence
    ws7 = wb.create_sheet("7-Evidence")
    ws7["A1"] = "Evidence / Cost Boundary / Missing Data Audit"
    ws7["A1"].font = TITLE_FONT
    write_header(ws7, 4, ["项", "内容", "来源/位置", "状态", "置信度",
                          "影响成本项", "待确认问题（缺什么证据能回答）"],
                 [22, 60, 30, 11, 9, 20, 54])
    supplier_list = "、".join(sorted({s.split(",")[0].strip()
                                      for s in (r["supplier"] for r in leaves) if s})) or "无"
    ev_rows = [
        ("质量基线", f"{len(mass_rows)} 个计量叶件合计 {total_kg:.3f} kg（父件不参与求和）",
         src_leaves, "Verified", "高", "全部",
         "无；紧固件的重量已含在父件 Fastener Weight 内（见下一条）"),
        ("紧固件包计数口径",
         f"{std_part_n} 个料号 / {std_qty:.0f} 件实物 / {std_kg:.3f} kg；"
         f"以 {std_part_n} 行进 BOM，但 Total Weight 全部留空",
         f"{src_leaves} 的 Mass Accounting 列", "Verified", "高", "Level 1",
         f"无；★ 重量不重复计量，成本单独计价。若以后有人把标准件行的重量填回去，"
         f"整机质量会凭空多出 {std_kg:.3f} kg"),
        ("成本边界（Level 1）",
         f"Make {bu_n} 行按自制成本口径（材料+工序+OH，不含 SG&A/利润/物流）；"
         f"Purchased {len(parts) - bu_n} 行按交付价（Delivered at assembly plant）",
         "2-BOM Make/Buy 列与 Cost Boundary 列", "Scenario", "中", "Level 1",
         "供应商交付条款（FCA/FOB/DAP）未取得；自制/外购真实边界需采购系统与制造分工证据"),
        ("成本边界（Level 2）", "整机集成 VA 全部标 Scenario；集成方身份 Unknown",
         "1-Assumptions INTEG_IDENTITY", "Unknown", "无", "Level 2",
         "★ 整机总装在谁做？OEM 内部？Tier1？还是总成级供应商？"
         "这一条决定用 S1 还是 S2 口径"),
        ("供应商结构", supplier_list, "Teardown Data Manufacturer 字段", "Verified", "中",
         "Level 1",
         "★ 无供应商名 ≠ 自制：需通过件号向采购系统反查，"
         "不能因为没有厂商名就推断为 OEM 自制"),        ("Bottom-up 覆盖",
         f"{bu_n} 项（Make 自制）：{bu_family_text}；"
         f"{'九类材料分层 + 工序路由' if has_attr_split else '仅工序路由，材料分层缺失'}",
         f"{src_attrs} + 2-BOM 材料列 + 3-ProcessCost", "Calculated", "中", "Level 1 材料与加工",
         ("分层重量为 Scenario 假设，需拆件称重核实" if has_attr_split else
          "★ 无 attrs 或分层缺失 → 材料输入列为 0，Mfg Cost 仅含加工费；"
          "Level 1 因此**低估**，补齐分层后须重算")),
        ("自制口径声明",
         "★ Make/Buy 二分是情景假设（按用户要求以有自制件的口径建模），不是拆解结论——"
         "拆解无法判断零件由谁制造。Make 价格 = (材料+工序)×(1+OH)，不含 SG&A/利润/物流",
         "2-BOM Make/Buy 列 + 1-Assumptions OH_RATE", "Scenario", "中", "Level 1 全部 Make 件",
         "★ 若实际为外购，应改用 v5（全外购口径）的供应商交付价——两者金额不可直接对比；"
         "自制/外购的真实边界需采购系统与制造分工证据"),
        ("取证价覆盖",
         (f"{ev_n} 项：PCBA 为 Class C 参数化 + 类比外推 Class D（含标准件 {std_part_n}）"
          if ev_n else
          f"0 项 —— ★ 本产品无价格取证表，{unknown_n} 行金额为空"),
         src_prices, "Estimated" if ev_n else "Unknown", "低" if ev_n else "无",
         "Level 1 交付价",
         ("无 Actual Procurement、无 Supplier Quotation；Class C+D 金额占比见 4-Summary。"
          if ev_n else
          f"★ {unknown_n} 行无任何价格证据。Excel 空单元格不参与求和，"
          "Level 1 合计只反映已建模件，是**结构性低估**，不得当作完整采购价")),
        ("基准价补位（Class E）",
         (f"{bench_n} 行：3a 严格类别 {bench_tier_n.get('3a', 0)} / "
          f"3b 包络类别 {bench_tier_n.get('3b', 0)} / 3c 质量轴兜底 "
          f"{bench_tier_n.get('3c', 0)}；合计金额占比见 4-Summary 补位块"
          if bench_n else "0 行——本次未触发补位"),
         f"{sn.get('catalog') or '（未指定基准价目录）'} + 4-Summary 补位块",
         "Estimated" if bench_n else "N/A", "低" if bench_n else "无", "Level 1 交付价",
         ("★ 这些是**行业均价，不是供应商报价**，只能用于结构判断与 RFQ 优先级排序，"
          "**不得用于定点谈判**。3b/3c 两档的区间是包络或按净重推算，"
          "宽达数倍属正常，不要把它当精度。"
          "★ 若某行同时存在取证价或 bottom-up 模型，补位一律不覆盖——"
          "补位只填原本会留空的行")
         if bench_n else
         "无；如需对缺价件补位，确认 --catalog 指向的目录存在且未被 --no-benchmark-fill 关闭"),
        ("标准件计价口径",
         f"{std_part_n} 个料号按件计价（{std_qty:.0f} 件 / {std_kg:.3f} kg），"
         "不进 3-ProcessCost、不做材料破拆",
         f"{src_prices} + 2-BOM Price Class 列", "Estimated", "中",
         "Level 1 紧固件包",
         "紧固件是标准外购件，其价值主要在热处理与表面处理，不靠材料重量驱动；"
         "但尺寸口径、法兰/平头、是否自锁未逐件实测，"
         "也未区分不锈钢与碳钢镀层，因此留 D 类"),
        ("PCBA 采购对象",
         "PCB 类件按『已贴装 PCBA』处理；裸板面积价单列为代理值，不作交付价",
         "2-BOM Bare PCB Proxy 列 + 取证价表", "Estimated", "低", "Level 1 最大单项",
         "★ 板层数、叠构、铜厚、表面处理、器件清单、贴装点数、拼板数、直通率——全缺失。"
         "没有器件清单时，任何『PCBA 应成本』都只能叫参数化估算"),
        ("PCBA 系数", "2.5–5.0（裸板约占 PCBA 20–40%）",
         "取证价表 Price Basis 列", "Estimated", "低", "Level 1",
         "★ 无任何取证。这是本模型最大的单一不确定性"),
        ("费率口径", "CAST/STAMP/INJ/TEST 全包含人工；CNC 裸机；ASSY 纯人工",
         "3-ProcessCost Rate Basis 列", "Estimated", "中", "供应商加工成本",
         "供应商实际费率口径（是否含折旧、能耗、维护、间接费）未知；"
         "口径比数值更影响结果"),
        ("节拍", "压铸/注塑按重量式；CNC/冲压/装配用工序模板",
         "3-ProcessCost Cycle Basis 列", "Estimated", "低", "供应商加工成本",
         "需供应商 OEE 与瓶颈节拍实测；模板节拍对小件偏高、对大件偏低"),
        ("Tooling", "供应商工装假定含在售价内、未单列；集成设备按 1-Assumptions 值为 Unknown",
         "2-BOM Tooling CapEx 空 + 1-Assumptions INTEG_INV", "Unknown", "无",
         "工装摊销 / Level 2 折旧",
         "★ Tooling Unknown 不等于 0。需模具清单与报价；若供应商单独收取模具费，"
         "必须另开一次性成本行，且不得同时摊入单价"),
        ("年产量 / 产量规划", "按 1-Assumptions VOL_A / VOL_L 假设",
         "1-Assumptions VOL_A / VOL_L", "Unknown", "无", "折旧 / 工装摊销",
         "需 SOR 或产量规划"),
        ("FIPG", "INTEG_FIPG_G 按 1-Assumptions 值（缺省 0 = 密封靠成型垫圈）",
         "1-Assumptions INTEG_FIPG_G", "Scenario", "中", "Level 2",
         "拆解确认是否有胶道。有则须改 >0 并重算 Level 2"),
        ("功率级位置",
         "★ 需按本产品实际清单核对：功率器件/模块/冷板是否为独立件号，"
         "还是集成在 PCBA 或壳体上（导热膏件的存在是功率级存在的旁证）",
         "Teardown Data 清单 + 过程取证", "Unknown", "无", "PCBA 交付价",
         "★ 这一条是清单检查项，不是结论：若清单确无功率级件号，"
         "需正反面俯拍 + 器件丝印，或导热膏压痕位置反推热源"),
    ]
    r = 5
    for row_v in ev_rows:
        write_body(ws7, r, list(row_v), row_v[3])
        r += 1

    # ---- Missing Data Audit
    r += 1
    ws7.cell(row=r, column=1, value="Missing Data Audit（空缺与错误逐项归类）").font = SEC_FONT
    r += 1
    for i, h in enumerate(["Sheet", "Part No / Cell", "Field", "Blank Type",
                           "Root Cause", "Proposed Fix", "Result"], start=1):
        c = ws7.cell(row=r, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    audit = [
        ("2-BOM", f"全部 {len(parts)} 行 / Make/Buy", "Make/Buy", "D → 分类错误",
         "v1 用件名正则 is_buy() 二分；Cover/Busbar/Terminal/Plug 因无供应商名被视为自制",
         "v2 全部强制 Purchased；v6 按用户口径恢复二分：结构件 Make（自制成本口径）、"
         "电子件/密封件/标准件 Purchased；『自制』为情景假设并已在表内声明",
         f"已改：Make {bu_n} / Purchased {len(parts) - bu_n}"),
        ("4-Summary", "5–24 行", "Purchased 金额", "D → 映射错误",
         "v1 把外购价留在 2-BOM 价格列，而汇总表 Purchased 列引用的是另一列（永远为空），"
         "导致 11 个外购件在明细行全空",
         "v2 逐行同时给出 Low/Base/High + Price Source + Price Status；不依赖空列",
         f"已修：{len(parts)} 行全部有价"),
        ("4-Summary", "合计行", "外购件合计", "D → 映射错误",
         "同上；∑ 恒为空，故 v1 的整机金额里外购件只出现在三层块，明细行看不到",
         "Level 1 合计行 = SUM(Amt Base)", "已修"),
        ("2-BOM", "Top/Bottom Cover 材料列", "材料分层", "C → 口径错误",
         "Mixed Material 整件按 species 首个命中材料（铝）计价；"
         "冷却管/堵盖的 Fastener Weight 亦被并入塑料/铝",
         "九类净重分列 + 各自单价 + 废料回收仅对金属；紧固件剥离到 STEEL 列",
         "已修（当 attrs 提供分层时生效）"),
        ("3-ProcessCost", "全包费率行 Operators", "Operators", "D → 重复计价",
         "RATE_CAST/STAMP/INJ/TEST 已含人工，工序表仍填人数并另计人工",
         "RATE_KIND 判定：allin → operators 强制 0；machine → 另加；labor → 机时费 0",
         "已修"),
        ("3-ProcessCost", "Cumulative Cost", "良率", "D → 计算逻辑错误",
         "v1 各步独立除良率后相加，未体现前道损失；且若对整列求和会把中间累积值重复计入",
         "C_i=(直接+i−1 累积)/良率_i，并加 Is Last Step 列，2-BOM 只取最后一行的累积值",
         "已修"),
        ("2-BOM", "PCB 类行", "PCBA 交付价", "C → 口径错误",
         "取证表若只有裸板面积价与 PCBA 系数，生成公式必须显式使用系数；"
         "否则会把裸板价直接当成 PCBA 交付价",
         "系数并入 Low/High；裸板单列 Bare PCB Proxy 并明标不得作交付价",
         "已修（当 prices 提供时生效）"),
        ("price-evidence CSV", "裸板代理价行", "Bare PCB Proxy", "D → 列错位",
         "裸板代理价曾误填为 PCBA 交付价 Base（应填裸板几何均值）。"
         "该错误发生在某个项目的取证表里，与本生成器无关——生成器只读不推断",
         "取证表内改正，并在 Basis 里写明修正；生成器不做自动纠错", "已修（数据侧）"),
        ("build_leaf_list.py", "标准件行", "叶件纳入条件", "D → 静默丢行",
         "旧过滤条件为 `r[\"path\"] not in prefix and r[\"total_weight\"]`——"
         "紧固件的 Total Weight 为空（重量已归入父件），整行被当成『无重量』丢弃",
         "按 Part Type == 'Standard' 单独识别并强制纳入，同时在 Total Weight 置空、"
         "另加 Parent Part No / Mass Accounting 两列声明计重口径",
         f"已修：BOM = {len(parts)} 行，其中标准件 {std_part_n} 行 / {std_qty:.0f} 件"),
        ("4-Summary", "—", "自制件三层块", "A → N/A",
         "『① 自制件供应商报价』在本项目不存在（无自制件）",
         "改为 Level 1 / Level 2 / Level 3 三种互斥口径", "已改：S0/S1/S2"),        ("4-Summary", "—", "Level 2 Final Integration", "C → 估算但已标 Scenario",
         "集成方与地点 Unknown；节拍由『界面点 × 每点秒数』推出，界面点数 ≠ 装配动作数",
         "整块标 Scenario 并写明『不是已确认生产成本』；节拍参数状态改 Scenario", "已修"),
        ("1-Assumptions", "INTEG_INV", "设备投资", "B → Unknown",
         "设备投资为未经确认的假设值，若据此算出折旧/件会被当实际成本",
         "状态改 Unknown；折旧行在 Level 2 块内标 Scenario", "已修"),
        ("2-BOM", "全部行 Tooling CapEx", "Tooling CapEx", "B → Unknown（有意留空）",
         "供应商模具费未取得",
         "留空并在 3-ProcessCost 脚注声明『假定含在售价内』；"
         "不得把 Unknown 当 0 后宣称模型不含工装", "已声明"),
        ("5-Opportunities", "原『混合材料分层计价』条", "降本归类", "D → 归类错误",
         "会计口径修正被列为 Cost Saving，用户会误以为有净收益",
         "移入 Accounting Correction 区，明确不计入 Cost Saving", "已修"),
        ("7-Evidence", "—", "覆盖率声明", "D → 表述错误",
         "旧版称『全部外购件均有价，Coverage = 100%』——把『有估算』当成『有采购价』",
         "区分 Model Coverage（有估算）与 Verified Price Coverage（有真实采购价），"
         "并声明 Class C+D 金额占比", "已修"),
        ("2-BOM / 4-Summary", "Make 行", "自制口径（v6 新增）", "C → 口径选择",
         "v5 全外购口径把结构件也按供应商交付价（含 SG&A/利润/物流）计价；"
         "用户要求按『有自制件』的口径另建一版",
         "v6 恢复 Make/Buy 二分：结构件 Make = (材料+工序)×(1+OH)，"
         "不含 SG&A/利润/物流；无法自制建模的件按 Class E 行业均值补位；"
         "『自制』为情景假设，已在 2-BOM 口径 1 与 7-Evidence 声明",
         f"已改：Make {bu_n} 件 / Class E 补位 {bench_n} 件；与 v5 金额不可直接对比"),
    ]
    for row_v in audit:
        write_body(ws7, r, list(row_v), "Estimated")
        r += 1

    # ---- 页码/打印
    for ws in wb:
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws7.freeze_panes = "A5"

    wb.save(out)
    try:
        n_inf = sum(1 for p in parts
                    if str(p.get("route_status", "")).startswith("Inferred"))
        n_weak = sum(1 for p in parts if p.get("route_status") == "Inferred-Weak")
        n_unk = sum(1 for p in parts if p.get("route_status") == "Unknown")
        print(f"已写入 {out}")
        print(f"  产品 {product}；功能域前缀 {func_prefix}；"
              f"工艺证据级别 {evidence_level}")
        print(f"  BOM {len(parts)} 行（Make {bu_n} + Purchased {len(parts) - bu_n}）= "
              f"{len(mass_rows)} 计量叶件 + "
              f"{std_part_n} 标准件行 / {std_qty:.0f} 件；"
              f"取证价 {ev_n} 项 / Class E 补位 {bench_n} 项 / 无价 {unknown_n} 项；工序 {p3_rows} 条")
        print(f"  Route Status：Assumed {bu_n - n_inf - n_weak - n_unk} / "
              f"Inferred {n_inf} / Inferred-Weak {n_weak} / Unknown {n_unk} / N/A {ev_n}")
        print(f"  质量 {total_kg:.3f} kg（标准件 {std_kg:.3f} kg 已含在父件内，不重复计）；"
              f"界面点 {iface_cnt}；铝 {al_kg:.3f} kg / 铜 {cu_kg:.3f} kg")
        if evidence_level != "1 级（attrs 精确路由）":
            print("  ⚠ 降级提示：无 attrs 或 Object 未识别，材料分层为空、"
                  "Supplier Mfg Cost 仅含加工费；补齐 --attrs 可升级为 1 级")
        if unknown_n:
            print(f"  ⛔ 缺价提示：{unknown_n} 行无任何价格证据（金额为空，不是 0）——"
                  f"Level 1 合计是**结构性低估**；补齐 --prices 取证价表或对其建模")
    except UnicodeEncodeError:
        pass


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 Make/Buy 混合口径 Should-Cost 工作簿（v6）")
    ap.add_argument("--bom", type=Path, required=True)
    ap.add_argument("--attrs", type=Path, default=None)
    ap.add_argument("--prices", type=Path, default=None)
    ap.add_argument("--opps", type=Path, default=None,
                    help="产品专属降本机会表 CSV（可选）。★ 不提供时只输出按 BOM 推导的"
                         "通用候选；绝不把某个产品的零件写进别的产品工作簿")
    ap.add_argument("--opps-ac", type=Path, default=None,
                    help="产品专属会计修正表 CSV（可选）")
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--product", default="产品")
    ap.add_argument("--function-prefix", default=None,
                    help="功能域前缀（INV/DCDC/OBC/PDU/BMS/GEN）。缺省由 --product 推导；"
                         "产品名不含关键词时用它显式指定，避免落到通用前缀 GEN")
    ap.add_argument("--catalog", type=Path, default=None,
                    help="基准价目录 CSV（Step 5.8 Class E 补位）。缺省用 "
                         "data/benchmark/benchmark-price-catalog.csv")
    ap.add_argument("--no-benchmark-fill", action="store_true",
                    help="关闭基准价补位：无取证价且无 bottom-up 的行恢复留空（Unknown）。"
                         "仅供回归对照，正式交付不要用")
    a = ap.parse_args()
    if not a.bom.exists():
        raise SystemExit(f"找不到 BOM：{a.bom}")
    leaves = load_leaves(a.bom)
    if not leaves:
        raise SystemExit("BOM 为空")
    a.out.parent.mkdir(parents=True, exist_ok=True)
    cat_path = a.catalog or CATALOG_DEFAULT
    if a.catalog and not cat_path.exists():
        raise SystemExit(f"找不到基准价目录：{cat_path}")
    build(leaves, load_attrs(a.attrs), load_prices(a.prices), a.out, a.product,
          a.function_prefix,
          src_names={"leaves": a.bom.name,
                     "attrs": a.attrs.name if a.attrs else "",
                     "prices": a.prices.name if a.prices else "",
                     "opps": a.opps.name if a.opps else "",
                     "opps_ac": a.opps_ac.name if a.opps_ac else "",
                     "catalog": cat_path.name},
          opps_in=load_opps(a.opps), ac_in=load_opps_ac(a.opps_ac),
          catalog=load_catalog(cat_path) if not a.no_benchmark_fill else None,
          bench_fill=not a.no_benchmark_fill)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())