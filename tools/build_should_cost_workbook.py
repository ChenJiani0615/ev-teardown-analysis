#!/usr/bin/env python3
"""生成 7 表 Should-Cost 工作簿（全部为活公式，参数集中在 1-Assumptions）。

用法:
    python tools/build_should_cost_workbook.py \
        --bom data/ex90-rear-inverter/derived/inverter-assy-leaves.csv \
        --out reports/EX90_RearInverter_ShouldCost_Analysis.xlsx \
        --product "Volvo EX90 Ultra 2025 后逆变器"

设计约束（与 should-cost-engineering/SKILL.md 一致）:
    * 所有费率/单价/良率只写在 1-Assumptions，公式只引用单元格；
    * 外购件不建工艺、不计材料，价格缺失时留空（Unknown），不补零；
    * 只用叶件求和，父件不参与，避免重复计价；
    * 降本机会的节省金额一律用公式，输入不全则返回空。

成本完整度（2026-11-03 Phase 2A）:
    * Unknown 价格不得被 N()/SUM 静默当 0：合计行旁给 Known/Modeled 小计、
      Unknown 计数、覆盖率、Complete PPC 判定四件套；
    * PPC 只有在 Unknown 计数 = 0 时才可标"完整"，否则标"部分（仅自制件）"；
    * 外购件价格状态七类：Actual Procurement / Supplier Quotation / Comparable /
      Bottom-up Estimate / Parametric / Assumption / Unknown；
      Supplier Quotation 是报价不是已发生成本，不得标 Verified。
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ 参数表

PARAMS = [
    ("CUR", "币种", "CNY", "-", "Unknown", "未取得报价，币种为假设"),
    ("REG", "生产地区", "中国华东", "-", "Estimated", "供应商产地未取得"),
    ("VOL_A", "年产量", 20000, "件/年", "Unknown", "待项目输入"),
    ("VOL_L", "生命周期产量（合格）", 100000, "件", "Unknown", "待项目输入"),
    ("AL_PRICE", "铝压铸合金单价", 23.9, "CNY/kg", "Verified", "长江现货 2026-10-28；与 v4 逆变器工作簿同锚"),
    ("CU_PRICE", "铜单价", 110.5, "CNY/kg", "Verified", "长江现货 2026-10-28；与 v4 逆变器工作簿同锚"),
    ("ST_PRICE", "钢单价", 8, "CNY/kg", "Estimated", "冷轧/镀锌板"),
    ("PBT_PRICE", "PBT 单价", 25, "CNY/kg", "Estimated", ""),
    ("PPS_PRICE", "PPS 单价", 60, "CNY/kg", "Estimated", ""),
    ("PPA_PRICE", "PPA 单价", 45, "CNY/kg", "Estimated", ""),
    ("PA_PRICE", "PA66 单价", 30, "CNY/kg", "Estimated", ""),
    ("PLASTIC_PRICE", "通用工程塑料单价", 25, "CNY/kg", "Estimated", "PC-ABS/ABS/POM 等未单独列价的牌号"),
    ("FERRITE_PRICE", "铁氧体单价", 30, "CNY/kg", "Estimated", ""),
    ("SCRAP_REC", "废料回收率", 0.60, "-", "Estimated", "废料按材料价 × 该比例回收"),
    ("Y_CAST", "压铸材料利用率", 0.85, "-", "Estimated", "含浇口/溢流"),
    ("Y_STAMP", "冲压材料利用率", 0.72, "-", "Estimated", "含排样废料，须实测"),
    ("Y_INJ", "注塑材料利用率", 0.92, "-", "Estimated", "含水口料回用"),
    ("RATE_CAST", "压铸机费率（含人工）", 180, "CNY/h", "Estimated", "全包口径，不再加人工"),
    ("RATE_CNC", "CNC/清洗费率（不含人工）", 90, "CNY/h", "Estimated", "裸机口径，需另加人工"),
    ("RATE_STAMP", "冲压机费率（含人工）", 120, "CNY/h", "Estimated", "全包口径"),
    ("RATE_INJ", "注塑机费率（含人工）", 110, "CNY/h", "Estimated", "全包口径"),
    ("RATE_ASSY", "装配人工费率", 45, "CNY/人·h", "Estimated", ""),
    ("RATE_TEST", "测试台费率（含人工）", 80, "CNY/h", "Estimated", ""),
    ("PLATE_PRICE", "电镀外协单价", 20, "CNY/kg", "Estimated", "外协工序按价格×基数计价，不走设备费率"),
    ("TOOL_LIFE_SHOTS", "工装寿命（模次）", 500000, "模次", "Estimated", "摊销基数 = MIN(寿命产量, 生命周期产量)"),
    ("CALIB_YEAR", "单价口径年份", 2026, "年", "Estimated", "全部 Estimated 单价的时间锚点，跨年对比须重估"),
    ("OH_RATE", "制造间接费率", 0.12, "-", "Estimated", "对直接成本"),
    ("SGA_RATE", "SG&A 率", 0.05, "-", "Estimated", ""),
    ("LOG_RATE", "物流包装率", 0.01, "-", "Estimated", "对工厂成本"),
    ("PROFIT_RATE", "利润率（按销售额）", 0.08, "-", "Estimated", "不是成本加成率"),
    ("CU_KG", "铜件总重（数据推导）", 0.0, "kg", "Calculated", "由 BOM Material Species 含 Copper 的行汇总"),
    ("AL_KG", "铝件总重（数据推导）", 0.0, "kg", "Calculated", "由 BOM Material Species 含 Aluminium 的行汇总"),
    # ---- v4 增补（2026-11-03）：三层口径 + 区间取值 + 接口点法 + 工艺材料三件套
    ("RANGE_MAX_RATIO", "区间可用上限（High/Low）", 1.5, "倍", "Estimated",
     "≤ 该值视为窄区间，直接保留区间；> 该值改用均值，不展示区间"),
    ("CENTER_METHOD", "宽区间取值方法", "GEOMEAN", "-", "Estimated",
     "GEOMEAN = 几何均值 √(Low×High)；改成 AVERAGE 即为算术均值，公式会自动跟随"),
    ("PASS_RATE", "外购件处置费率（资金占用/来料检验）", 0, "-", "Estimated",
     "默认 0 = 外购件只付供应商价 + 物流，不叠加间接费与利润；若需计入内部处置成本可填 0.02–0.03"),
    ("IFACE_CNT", "整机界面点数（紧固/插接/定位）", 0, "点", "Unknown",
     "拆解数据 Number of Interface 列逐件求和；生成后按产品核对"),
    ("ASSY_CT_MIN", "整机总装节拍", 0, "min/台", "Estimated",
     "界面点 × 8–10 s + 点胶 + 合盖 + 搬运；生成后按产品重估"),
    ("EOL_CT_MIN", "EOL 测试节拍（Hi-Pot/功能/气密）", 5, "min/台", "Estimated",
     "逐台 EOL，非抽检；含老化另计"),
    ("FIPG_G", "每台密封胶用量（FIPG）", 0, "g/台", "Estimated",
     "按密封界面胶道总长 × 25–30 g/m 估；无 FIPG 界面的产品填 0"),
    ("FIPG_PRICE", "密封胶单价（脱醇型硅酮 FIPG）", 90, "CNY/kg", "Estimated",
     "车规 FIPG 品类量级 60–150 元/kg"),
    ("CONS_OTHER", "其他辅料（清洗剂/擦拭/扎带/标签/防错包装）", 2, "CNY/台", "Estimated",
     "小件耗材，按每台 2 元估"),
    ("ASSY_INV", "总装单元设备投资（点胶+拧紧+压装+气密+EOL）", 5000000, "CNY", "Unknown",
     "半自动总装单元；若为全自动线需上调至 1.2–2 亿"),
    ("ASSY_LIFE", "设备折旧年限", 5, "年", "Estimated", "直线折旧"),
]

# 工序模板：(工序, 类别, 节拍s, 费率参数, 人数, 耗材, 步骤良率)
# 外协行：费率参数为 None，耗材为 ("outsrc", 单价参数, 基数)——按外协价×基数计价，不走设备费率
CAST = [
    ("压铸", "加工", 60, "RATE_CAST", 0.5, 0.0, 0.95),
    ("修边", "加工", 20, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("CNC 加工", "加工", 120, "RATE_CNC", 1.0, 3.0, 0.98),
    ("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99),
    ("检漏", "测试", 30, "RATE_TEST", 1.0, 2.0, 0.99),
]
STAMP = [
    ("冲压", "加工", 8, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("去毛刺", "加工", 15, "RATE_ASSY", 1.0, 0.0, 0.99),
    ("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99),
]
PLATE = [
    ("冲压", "加工", 10, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("平面磨", "加工", 60, "RATE_CNC", 1.0, 2.0, 0.97),
    ("清洗", "加工", 15, "RATE_CNC", 0.5, 1.0, 0.99),
]
DIE_CUT = [("模切", "加工", 10, "RATE_STAMP", 1.0, 0.5, 0.98)]
INJ = [
    ("注塑", "加工", 45, "RATE_INJ", 1.0, 0.0, 0.97),
    ("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99),
]
INJ_INSERT = INJ + [("嵌件装配", "装配", 40, "RATE_ASSY", 1.0, 0.5, 0.98)]
OVERMOLD = [
    ("注塑包胶", "加工", 50, "RATE_INJ", 1.0, 0.0, 0.96),
    ("去浇口", "加工", 10, "RATE_ASSY", 1.0, 0.0, 0.99),
]
BUSBAR = [
    ("冲压", "加工", 12, "RATE_STAMP", 1.0, 0.0, 0.98),
    ("电镀（外协）", "加工", 0, None, 0.0, ("outsrc", "PLATE_PRICE", "weight"), 0.98),
    ("清洗", "加工", 10, "RATE_CNC", 0.5, 1.0, 0.99),
]
ASSY = [("装配", "装配", 60, "RATE_ASSY", 1.0, 0.5, 0.98)]

# 外购判定必须锚定名称开头，否则 "Cover, Rear Inverter PCB" 会被 "PCB" 误判为外购
BUY_RE = re.compile(
    r"^(PCB\b|SH Film Capacitor|Capacitor|MOSFET|Ferrite Core|Connector|Fuse|Sensor|"
    r"Jumper Harness|Thermal Paste|Breather|Inductor|Transformer|Gasket|O Ring|NVH Foam|Spring)",
    re.IGNORECASE,
)


def is_buy(name: str) -> bool:
    return bool(BUY_RE.match(name.strip()))


# ------------------------------------------------------------------ 价格状态
# 七类价格状态（Phase 2A 任务 A）。语义约束：
#   * Actual Procurement：已发生的采购成本（合同/订单价），唯一可标 Verified 的价格类
#   * Supplier Quotation：供应商报价，是"要价"不是"已发生成本"，不得标 Verified
#   * Comparable：可比件市场价（分销/零售价 ≠ 量产价，必须注明）
#   * Bottom-up Estimate：自下而上模型值（本工作簿 3/4 表可复算）
#   * Parametric Estimate：参数化/情景推算值
#   * Assumption：人为假设值（敏感性/情景专用，不得进 PPC 基线）
#   * Unknown：缺价，留空，不补零
PRICE_STATUS = {
    "Actual Procurement": "Verified",
    "Supplier Quotation": "Estimated",
    "Comparable": "Estimated",
    "Bottom-up Estimate": "Calculated",
    "Parametric Estimate": "Estimated",
    "Assumption": "Estimated",
    "Unknown": "Unknown",
}


def price_status_of(row: dict | None) -> str:
    """外购件价格状态：默认 Unknown；有输入时由调用方显式指定，不自动猜。"""
    if row and row.get("price_status"):
        ps = str(row["price_status"]).strip()
        if ps in PRICE_STATUS:
            return ps
    return "Unknown"


# ------------------------------------------------------------------ 功能域切割
# 按 PDF 对标逻辑两级切割：Function（电气功能域，Box 为兜底）→ Category（域内部件类别）。
# 边界公式（脚注声明，防口径漂移）：
#   INV function = PM + DC-Link 电容 + 电流传感 + PCB + E-Parts
#   Box function  = 壳体 + 冷却与热 + HV 连接 + 信号连接 + EMC + 密封紧固其他
# E-Parts 独立于 PCB：电子件是降本主战场，混入 PCB 会被大数淹没。
# 冷却与 TIM 归 Box 不归功率级：热管理是结构件成本，不是器件成本。

def function_category(name: str, species: str, func_prefix: str = "INV") -> tuple[str, str]:
    n = name.strip().lower()
    s = species.lower()
    # --- 电气功能域（前缀随产品：INV / DCDC / OBC）---
    if "mosfet" in n or "power module" in n:
        return func_prefix, "PM 功率级"
    if "film capacitor" in n or ("capacitor" in n and "sh" in n):
        return func_prefix, "DC-Link 电容"
    if "current sensor" in n or ("sensor" in n and "current" in n):
        return func_prefix, "电流传感"
    if "ferrite core" in n and "current sensor" in n:
        return func_prefix, "电流传感"
    if n.startswith("pcb"):
        return func_prefix, "PCB"
    if "inductor" in n or "transformer" in n:
        return func_prefix, "功率磁件"
    if "fuse" in n or "breather" in n or "terminal" in n:
        return func_prefix, "E-Parts"
    # --- Box function ---
    if "housing" in n or "bottom cover" in n or "top cover" in n or "access cover" in n:
        return "Box", "壳体"
    if "cooling plate" in n or "coolant" in n or "cooling pipe" in n or "thermal paste" in n:
        return "Box", "冷却与热"
    if "busbar" in n:
        return "Box", "HV 连接"
    if "connector" in n or "jumper harness" in n:
        return "Box", "信号连接"
    if "ferrite core" in n:
        return "Box", "EMC"
    if n == "cap":
        return "Box", "密封紧固其他"
    if "gasket" in n or "o ring" in n or "nvh foam" in n or "spacer" in n or "locator" in n:
        return "Box", "密封紧固其他"
    if "bracket" in n or "holder" in n or "cover" in n:
        return "Box", "密封紧固其他"
    # 兜底：按材料属性归 Box
    if "copper" in s or "aluminium" in s or "steel" in s:
        return "Box", "密封紧固其他"
    return func_prefix, "E-Parts"


def route_for(name: str, species: str) -> tuple[str, list]:
    n = name.lower()
    plastic = "plastic" in species.lower()
    if "housing" in n or "casing" in n:
        return "压铸 → 修边 → CNC → 清洗 → 检漏", CAST
    if "bottom cover" in n:
        return "压铸 → 修边 → CNC → 清洗 → 检漏", CAST
    if "top cover" in n:
        # DC-DC 上盖为 1.097 kg 压铸铝（96 mm 深腔，AC44300）→ CAST；
        # 钣金上盖（浅、轻、无压铸牌号）→ STAMP。按牌号分界。
        s = species.lower()
        die_cast = "aluminium" in s and any(k in s for k in ("ac44300", "ac43500", "enac", "adc", "alsi"))
        return ("压铸 → 修边 → CNC → 清洗 → 检漏", CAST) if die_cast else ("冲压 → 去毛刺 → 清洗", STAMP)
    if "plate" in n and "cooling" in n:
        return "冲压 → 平面磨 → 清洗", PLATE
    if "cooling pipe" in n:
        return "注塑 → 去浇口", INJ
    if n == "cap" and plastic:
        return "注塑 → 去浇口", INJ
    if "gasket" in n:
        return "模切", DIE_CUT
    if "nvh foam" in n:
        return "模切", DIE_CUT
    if "busbar with ferrite" in n:
        return "冲压 → 注塑包胶 → 去浇口", OVERMOLD
    if "busbar" in n:
        return "冲压 → 电镀 → 清洗", BUSBAR
    if "terminal" in n:
        return "注塑 → 去浇口 → 嵌件装配", INJ_INSERT
    if "adapter" in n:
        return "注塑 → 去浇口 → 嵌件装配", INJ_INSERT
    if "bracket" in n:
        return ("注塑 → 去浇口", INJ) if plastic else ("冲压 → 去毛刺", STAMP)
    if "cover" in n:
        return ("注塑 → 去浇口", INJ) if plastic else ("冲压 → 去毛刺", STAMP)
    if "spacer" in n:
        return "注塑 → 去浇口", INJ
    if "plug" in n and "plastic" not in species.lower():
        return "冲压 → 去毛刺 → 清洗", STAMP
    if "locator" in n or "holder" in n:
        return "冲压 → 去毛刺", STAMP
    return "装配", ASSY


def material_param(species: str) -> tuple[str, str] | None:
    s = species.lower()
    if "aluminium" in s:
        return "AL_PRICE", "Y_CAST"
    if "copper" in s:
        return "CU_PRICE", "Y_STAMP"
    if "steel" in s:
        return "ST_PRICE", "Y_STAMP"
    if "pps" in s:
        return "PPS_PRICE", "Y_INJ"
    if "ppa" in s:
        return "PPA_PRICE", "Y_INJ"
    if "pbt" in s:
        return "PBT_PRICE", "Y_INJ"
    if "pa66" in s or "pa6" in s or "para" in s or "hdpe" in s:
        return "PA_PRICE", "Y_INJ"
    if any(k in s for k in ("pc-abs", "pc/abs", "abs", "pom", "pc ")):
        return "PLASTIC_PRICE", "Y_INJ"
    if "ferrite" in s:
        return "FERRITE_PRICE", "Y_STAMP"
    return None


def load_bom(path: Path) -> list[dict]:
    rows = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            qty = float(r.get("Quantity") or 1) or 1.0
            total = float(r.get("Total Weight (kg)") or 0)
            pprice_raw = (r.get("Purchased Price") or "").strip()
            itf_raw = (r.get("Number of Interface") or "").strip()
            rows.append({
                "part": (r.get("CS Part") or "").strip(),
                "qty": qty,
                "weight": total / qty if qty else total,
                "group": (r.get("Material Group") or "").strip(),
                "species": (r.get("Material Species") or "").strip(),
                "supplier": (r.get("Manufacturer") or "").strip(),
                "name": (r.get("Name") or "").strip() or (r.get("Part Type") or ""),
                # 价格状态与采购价（Phase 2A 任务 A）：显式输入，不自动猜
                "price_status": (r.get("Price Status") or "").strip(),
                "purchased_price": float(pprice_raw) if pprice_raw else None,
                # v4：件级界面点数（接口点法装配工时的数据基础）
                "itf": int(float(itf_raw)) if itf_raw else 0,
            })
    return rows


def load_prices(path: Path | None) -> dict[str, dict]:
    """读价格取证 CSV（Step 5.7）：CS Part → Low/High/Class/Basis/Status。

    证据数据放独立文件而不是生成器硬编码：价格取证是可复核交付物
    （7-Evidence 要求来源+日期+换算可追溯），且生成器保持产品无关。
    """
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
            out[part] = {
                "low": float(low) if low else None,
                "high": float(high) if high else None,
                "cls": (r.get("Price Class") or "").strip(),
                "basis": (r.get("Price Basis") or "").strip(),
                "status": (r.get("Price Status") or "").strip(),
            }
    return out


# ------------------------------------------------------------------ 样式

HDR_FILL = PatternFill("solid", fgColor="305496")
HDR_FONT = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=9)
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STATUS_FILL = {
    "Verified": None,
    "Calculated": PatternFill("solid", fgColor="DDEBF7"),
    "Estimated": PatternFill("solid", fgColor="FFF2CC"),
    "Unknown": PatternFill("solid", fgColor="FCE4E4"),
}
MONEY = '#,##0.00'
KG = '0.000'


def header(ws, row: int, labels: list[str], widths: list[int]) -> None:
    for i, (label, width) in enumerate(zip(labels, widths), start=1):
        cell = ws.cell(row=row, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = 28


def body(ws, row: int, values: list, status: str | None = None) -> None:
    for i, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=value)
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    if status and STATUS_FILL.get(status):
        for i in range(1, len(values) + 1):
            ws.cell(row=row, column=i).fill = STATUS_FILL[status]


# ------------------------------------------------------------------ 构建

def _emit(text: str) -> None:
    """stdout 被重定向到管道时系统代码页可能是 cp1252，中文会抛 UnicodeEncodeError。"""
    try:
        print(text)
    except UnicodeEncodeError:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            return
        print(text)


def build(bom_rows: list[dict], out: Path, product: str, prices: dict[str, dict] | None = None) -> None:
    prices = prices or {}
    wb = Workbook()
    wb.remove(wb.active)

    # 功能域前缀随产品：逆变器 INV / DC-DC 转换器 DCDC / OBC 车载充电机 OBC
    pl = product.lower()
    if "dc-dc" in pl or "dcdc" in pl:
        func_prefix = "DCDC"
    elif "obc" in pl:
        func_prefix = "OBC"
    else:
        func_prefix = "INV"

    # 材料重量由 BOM 推导，且只统计自制件（外购件的材料已在供应商价里，不能重复计入敏感性）
    make_rows = [r for r in bom_rows if not is_buy(r["name"] or r["part"])]
    cu_kg = sum(r["weight"] * r["qty"] for r in make_rows
                if "copper" in (r["species"] or "").lower())
    al_kg = sum(r["weight"] * r["qty"] for r in make_rows
                if "aluminium" in (r["species"] or "").lower())
    # v4：整机界面点数 = 件级 Number of Interface 求和（接口点法的数据基础）
    iface_cnt = sum(r["itf"] * r["qty"] for r in bom_rows)
    for i, (pid, label, value, unit, status, note) in enumerate(PARAMS):
        if pid == "CU_KG":
            PARAMS[i] = (pid, label, round(cu_kg, 3), unit, status, note)
        elif pid == "AL_KG":
            PARAMS[i] = (pid, label, round(al_kg, 3), unit, status, note)
        elif pid == "IFACE_CNT":
            PARAMS[i] = (pid, label, iface_cnt, unit,
                         "Verified" if iface_cnt else "Unknown",
                         "拆解数据 Number of Interface 列逐件求和"
                         + (f"（{len(bom_rows)} 个叶件）" if iface_cnt else "；CSV 无该列数据"))
        elif pid == "ASSY_CT_MIN":
            # 界面点 × 9 s + 合盖 60 s + 搬运 60 s，折成分钟
            est_min = round((iface_cnt * 9 + 120) / 60, 1) if iface_cnt else 0
            PARAMS[i] = (pid, label, est_min, unit, "Estimated" if iface_cnt else "Unknown",
                         f"{iface_cnt} 点 × 9 s + 合盖 60 s + 搬运 60 s ≈ {est_min} min"
                         if iface_cnt else "无界面点数据，待人工重估")

    # ---------- 1 Assumptions
    ws1 = wb.create_sheet("1-Assumptions")
    ws1["A1"] = f"{product} — Should-Cost 假设与费率"
    ws1["A1"].font = TITLE_FONT
    header(ws1, 3, ["参数ID", "参数", "值", "单位", "状态", "来源/说明", "备注"], [16, 30, 12, 10, 12, 46, 24])
    ref: dict[str, str] = {}
    r = 4
    for pid, label, value, unit, status, note in PARAMS:
        body(ws1, r, [pid, label, value, unit, status, note, ""], status)
        ws1.cell(row=r, column=3).number_format = KG if isinstance(value, float) else '#,##0.####'
        ref[pid] = f"'1-Assumptions'!$C${r}"
        r += 1
    ws1.freeze_panes = "A4"
    note_row = r + 1
    ws1.cell(row=note_row, column=1, value=(
        "说明：本表是全簿唯一的数值来源，其他工作表只引用这里的单元格。"
        "Estimated 项必须用供应商费率或当期行情替换后才可用于定点。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="C00000")

    # ---------- 2 BOM
    ws2 = wb.create_sheet("2-BOM")
    ws2["A1"] = "Engineering BOM（仅叶件，父件不参与求和，避免重复计价）"
    ws2["A1"].font = TITLE_FONT
    cols = ["Level", "Part No", "Name", "Function", "Category", "Qty", "Material", "Net Weight (kg)",
            "Material Yield", "Material Price", "Gross Input (kg)", "Scrap Value", "Material Cost",
            "Route", "Make/Buy", "Status", "Source", "Purchased Price", "Price Type", "Price Status",
            "Cost Boundary", "Quantity Basis", "Tooling CapEx (kCNY)", "Tooling",
            "Price Low (CNY)", "Price High (CNY)", "Price Base (CNY)", "Price Range Basis",
            "Price Class", "Amt Low", "Amt High", "Amt Base"]
    header(ws2, 3, cols, [7, 15, 34, 8, 13, 6, 26, 13, 12, 13, 13, 11, 13, 30, 11, 12, 30, 14,
                          16, 14, 22, 14, 13, 11, 13, 13, 13, 40, 10, 11, 11, 11])
    r = 4
    parts: list[dict] = []
    for row in bom_rows:
        name = row["name"] or row["part"]
        buy = is_buy(name)
        route_text, steps = route_for(name, row["species"])
        mat = material_param(row["species"]) if not buy else None
        make_buy = "Buy" if buy else "Make"
        func, cat = function_category(name, row["species"], func_prefix)

        # 价格状态（任务 A）：外购件默认 Unknown；自制件不适用（留空）
        ptype = ""
        pstatus = ""
        cboundary = ""
        qbasis = ""
        pprice = None
        ev: dict = {}
        if buy:
            ptype = "Unit Price"
            pstatus = price_status_of(row)
            # 有价但状态未标 → 按七类语义回落 Bottom-up Estimate 之外最诚实的档：
            # 价格来自输入文件，无来源说明时不允许冒充 Actual Procurement
            if row.get("purchased_price") is not None and pstatus == "Unknown":
                pstatus = "Comparable" if row.get("price_status") else "Unknown"
            pprice = row.get("purchased_price")
            # v4 价格取证（Step 5.7）：--prices CSV 的 Low/High/Class/Basis 优先于
            # BOM 内联价格；有取证价时升级价格状态与口径列（七类语义见 PRICE_STATUS）
            ev = prices.get(row["part"], {})
            ev_status = ev.get("status") or ""
            if ev.get("low") is not None and ev_status in PRICE_STATUS:
                pstatus = ev_status
            if pstatus != "Unknown":
                cboundary = "Delivered at assembly plant（到厂价，含包装运输）"
                qbasis = f"per piece × {row['qty']}"

        row_status = "Estimated" if not buy else PRICE_STATUS.get(pstatus, "Unknown")
        values = [
            3, row["part"], name, func, cat, row["qty"], row["species"], round(row["weight"], 4),
            None, None, None, None, None, route_text, make_buy,
            row_status,
            f"Teardown Data: {row['part']}" + (f"｜供应商 {row['supplier']}" if row["supplier"] else ""),
            pprice, ptype, pstatus, cboundary, qbasis, None, None,
        ]
        body(ws2, r, values, row_status)
        ws2.cell(row=r, column=8).number_format = KG

        if buy:
            # v4 价格取证五列：Low/High 为输入，Base 由区间规则公式取值，
            # Range Basis 说明来源，Class 标 A/B/C/D；X/Y/Z 为金额辅助列。
            # 取证数据来自 --prices CSV（Step 5.7），未取证件保持空 = Unknown。
            ev_low = ev.get("low")
            ev_high = ev.get("high")
            ev_cls = ev.get("cls") or ""
            ev_basis = ev.get("basis") or ""
            ws2.cell(row=r, column=25, value=ev_low)   # Price Low（取证输入）
            ws2.cell(row=r, column=26, value=ev_high)  # Price High（取证输入）
            ws2.cell(row=r, column=27, value=(
                f'=IF($Y{r}="","",IF($Z{r}/$Y{r}<={ref["RANGE_MAX_RATIO"]},AVERAGE($Y{r}:$Z{r}),'
                f'IF(UPPER({ref["CENTER_METHOD"]})="AVERAGE",AVERAGE($Y{r}:$Z{r}),'
                f'GEOMEAN($Y{r}:$Z{r}))))'))
            ws2.cell(row=r, column=28, value=ev_basis)  # Price Range Basis（来源+日期+换算）
            ws2.cell(row=r, column=29, value=ev_cls)    # Price Class（A/B/C/D）
            ws2.cell(row=r, column=30, value=f'=IF(ISNUMBER($Y{r}),$Y{r}*$F{r},"")')
            ws2.cell(row=r, column=31, value=f'=IF(ISNUMBER($Z{r}),$Z{r}*$F{r},"")')
            ws2.cell(row=r, column=32, value=f'=IF(ISNUMBER($AA{r}),$AA{r}*$F{r},"")')
            for c in (25, 26, 27, 30, 31, 32):
                ws2.cell(row=r, column=c).number_format = MONEY
            parts.append({"row": r, "part": row["part"], "name": name, "qty": row["qty"],
                          "steps": [], "weight": row["weight"], "buy": True,
                          "status": row_status,
                          "func": func, "cat": cat, "price_status": pstatus,
                          "purchased_price": pprice, "price_class": ev_cls})
        else:
            status = "Estimated"
            if mat and row["weight"] > 0:
                price_pid, yield_pid = mat
                ws2.cell(row=r, column=9, value=f"={ref[yield_pid]}")
                ws2.cell(row=r, column=10, value=f"={ref[price_pid]}")
                ws2.cell(row=r, column=11, value=f"=H{r}/I{r}")
                ws2.cell(row=r, column=12, value=f"=(K{r}-H{r})*J{r}*{ref['SCRAP_REC']}")
                ws2.cell(row=r, column=13, value=f"=K{r}*J{r}-L{r}")
                ws2.cell(row=r, column=11).number_format = KG
                for c in (12, 13):
                    ws2.cell(row=r, column=c).number_format = MONEY
            else:
                # 自制但材料单价未映射（如 PC-ABS）：工艺照算，材料留空，状态标 Unknown
                status = "Unknown"
                ws2.cell(row=r, column=16, value="Unknown")
                for c in range(1, len(values) + 1):
                    ws2.cell(row=r, column=c).fill = STATUS_FILL["Unknown"]
            # 工艺步骤照样生成，避免“Make 但零成本”
            parts.append({"row": r, "part": row["part"], "name": name, "qty": row["qty"],
                          "steps": steps, "weight": row["weight"], "buy": False,
                          "status": status, "func": func, "cat": cat, "price_status": ""})
        r += 1
    ws2.freeze_panes = "D4"
    # 功能域边界公式（脚注声明，口径随产品重新声明，不全局冻结）
    fn = r + 1
    ws2.cell(row=fn, column=1, value="功能域边界公式：").font = Font(name="微软雅黑", size=9, bold=True)
    ws2.cell(row=fn + 1, column=1, value=(
        "INV function = PM 功率级 + DC-Link 电容 + 电流传感 + PCB + E-Parts"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    ws2.cell(row=fn + 2, column=1, value=(
        "Box function = 壳体 + 冷却与热 + HV 连接 + 信号连接 + EMC + 密封紧固其他"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    ws2.cell(row=fn + 3, column=1, value=(
        "E-Parts 独立于 PCB；冷却与 TIM 归 Box 不归功率级。多合一产品扩展时 Function 加 DCDC/OBC，"
        "Category 加 PFC/LLC/Power Stage。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")

    # ---------- 3 ProcessCost
    ws3 = wb.create_sheet("3-ProcessCost")
    ws3["A1"] = "Process Cost Breakdown（逐工序，外购件不在此表）"
    ws3["A1"].font = TITLE_FONT
    header(ws3, 3, ["Step", "Part No", "Process", "Category", "Site", "Cycle (s)", "Machine Rate",
                    "Operators", "Labor Rate", "Consumable", "Step Yield", "Step Cost"],
           [7, 15, 20, 9, 10, 10, 12, 9, 12, 12, 9, 12])
    r = 4
    for p in parts:
        if p["buy"]:
            continue
        for idx, (proc, category, cycle, rate_pid, ops, cons, yld) in enumerate(p["steps"], start=1):
            if isinstance(cons, tuple) and cons[0] == "outsrc":
                # 外协行：按外协单价 × 基数计价，不走设备费率
                site = "outsource"
                cons_cell = f"={round(p['weight'], 4)}*{ref[cons[1]]}"
                cost_formula = f'=IF(J{r}="","",J{r}/K{r})'
                rate_cell = None
            elif isinstance(cons, tuple):
                site = "in-house"
                cons_cell = f"={round(p['weight'], 4)}*{ref[cons[1]]}"
                cost_formula = None
                rate_cell = f"={ref[rate_pid]}"
            else:
                site = "in-house"
                cons_cell = cons
                cost_formula = None
                rate_cell = f"={ref[rate_pid]}"
            values = [idx, p["part"], proc, category, site, cycle, rate_cell, ops,
                      f"={ref['RATE_ASSY']}", cons_cell, yld,
                      cost_formula or
                      f'=IF(OR(F{r}="",G{r}=""),"",(F{r}/3600*G{r}+H{r}*I{r}*F{r}/3600+J{r})/K{r})']
            body(ws3, r, values, "Estimated")
            ws3.cell(row=r, column=10).number_format = MONEY
            ws3.cell(row=r, column=12).number_format = MONEY
            r += 1
    ws3.freeze_panes = "C4"

    # ---------- 4 Summary
    ws4 = wb.create_sheet("4-Summary")
    ws4["A1"] = "Should-Cost Summary（叶件汇总 + 应成本推算）"
    ws4["A1"].font = TITLE_FONT
    header(ws4, 3, ["Part No", "Name", "Function", "Category", "Qty", "Material", "Purchased", "Process",
                    "Assembly", "Test", "Tooling Amort", "Direct Cost", "Overhead", "Factory Cost",
                    "Should-Cost Price", "Share %", "Status"],
           [15, 32, 8, 13, 6, 12, 12, 11, 11, 10, 13, 12, 11, 12, 15, 9, 12])
    r = 4
    first_data = r
    for p in parts:
        def idx(sheet_col: str) -> str:
            """INDEX/MATCH 取同件号单元格；对空单元格 Excel 返回 0，故配合 &"" 判空。"""
            return f"INDEX('2-BOM'!${sheet_col}:${sheet_col},MATCH($A{r},'2-BOM'!$B:$B,0))"

        def look(sheet_col: str) -> str:
            raw = idx(sheet_col)
            return f'=IFERROR(IF(TRIM({raw}&"")="","",{raw}),"")'

        tooling_capex = idx("W")
        values = [
            p["part"], p["name"], p["func"], p["cat"], p["qty"],
            # F 材料成本 ← 2-BOM M 列（Material Cost）；G 外购件 ← R 列（Purchased Price 输入列）。
            # v4 忠实做法：R 留空、价格在 Y/Z/AA，外购价不进 L 合计 → 三层块①不重复计价。
            look("M"),
            look("R"),
            f'=IF(COUNTIFS(\'3-ProcessCost\'!$B:$B,$A{r},\'3-ProcessCost\'!$D:$D,"加工")=0,"",'
            f'SUMIFS(\'3-ProcessCost\'!$L:$L,\'3-ProcessCost\'!$B:$B,$A{r},\'3-ProcessCost\'!$D:$D,"加工"))',
            f'=IF(COUNTIFS(\'3-ProcessCost\'!$B:$B,$A{r},\'3-ProcessCost\'!$D:$D,"装配")=0,"",'
            f'SUMIFS(\'3-ProcessCost\'!$L:$L,\'3-ProcessCost\'!$B:$B,$A{r},\'3-ProcessCost\'!$D:$D,"装配"))',
            f'=IF(COUNTIFS(\'3-ProcessCost\'!$B:$B,$A{r},\'3-ProcessCost\'!$D:$D,"测试")=0,"",'
            f'SUMIFS(\'3-ProcessCost\'!$L:$L,\'3-ProcessCost\'!$B:$B,$A{r},\'3-ProcessCost\'!$D:$D,"测试"))',
            f'=IF(TRIM({tooling_capex}&"")="","",{tooling_capex}*1000/MIN({ref["TOOL_LIFE_SHOTS"]},{ref["VOL_L"]}))',
            f'=IF(COUNT(F{r}:K{r})=0,"",SUM(F{r}:K{r}))',
            # v4 三层口径：外购件只加物流（PASS_RATE 默认 0，不加 OH/SG&A/利润）；
            # 自制件加 OH（供应商侧的辅料与设备费率已在工序成本里）
            f'=IF(L{r}="","",IF({idx("O")}="Buy",L{r}*{ref["PASS_RATE"]},L{r}*{ref["OH_RATE"]}))',
            f'=IF(L{r}="","",L{r}+M{r})',
            # 外购件 = 进厂成本（+物流）；自制件 = 供应商报价（+SG&A+物流，÷(1−利润)）
            f'=IF(N{r}="","",IF({idx("O")}="Buy",N{r}*(1+{ref["LOG_RATE"]}),'
            f'N{r}*(1+{ref["SGA_RATE"]})*(1+{ref["LOG_RATE"]})/(1-{ref["PROFIT_RATE"]})))',
            f'=IF(O{r}="","",O{r}/$O${len(parts) + 4})',
            p["status"],
        ]
        body(ws4, r, values, p["status"])
        for c in (6, 7, 8, 9, 10, 11, 12, 13, 14, 15):
            ws4.cell(row=r, column=c).number_format = MONEY
        ws4.cell(row=r, column=16).number_format = '0.0%'
        r += 1
    last_data = r - 1
    total_row = r
    # 17 列：A 合计 / B-D 空 / E Qty / F-O 数值列 / P Share% / Q Status
    totals = ["合计", "", "", "",
              f'=IF(COUNT(E{first_data}:E{last_data})=0,"",SUM(E{first_data}:E{last_data}))']
    for col in "FGHIJKLMNO":
        totals.append(
            f'=IF(COUNT({col}{first_data}:{col}{last_data})=0,"",SUM({col}{first_data}:{col}{last_data}))'
        )
    totals.append(f'=IF(O{total_row}="","",1)')
    totals.append("Calculated")
    body(ws4, total_row, totals, "Calculated")
    for c in range(6, 16):
        ws4.cell(row=total_row, column=c).number_format = MONEY
        ws4.cell(row=total_row, column=c).font = Font(name="微软雅黑", size=9, bold=True)
    ws4.cell(row=total_row, column=1).font = Font(name="微软雅黑", size=9, bold=True)

    # ---- PPC 口径块：MAT / VA 分离（材料性成本 vs 增值成本）
    ppc0 = total_row + 2
    ws4.cell(row=ppc0, column=1, value="成本结构（PPC = MAT + VA）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    ppc_rows = [
        ("材料性成本 MAT = Σ材料（自制件；外购件在②层，G 列留空防重复计价）",
         f"=N(F{total_row})+N(G{total_row})"),
        ("增值成本 VA = Σ加工 + Σ装配 + Σ测试", f"=N(H{total_row})+N(I{total_row})+N(J{total_row})"),
        ("自制件直接成本 = MAT_自制 + VA", f"=B{ppc0 + 1}+B{ppc0 + 2}"),
        ("外购件采购价合计（见下方 ② 层，不计入上行）",
         f"=SUM('2-BOM'!AF4:AF{3 + len(parts)})"),
        ("工装 CapEx 合计（一次性，kCNY）", f"=SUM('2-BOM'!W4:W{3 + len(parts)})"),
        ("工装摊销/件（单件口径）", f"=N(K{total_row})"),
        ("应成本价 = 工厂成本×(1+物流)/(1−利润)（仅自制件口径）", f"=N(O{total_row})"),
    ]
    for i, (label, formula) in enumerate(ppc_rows, start=1):
        row = ppc0 + i
        cell = ws4.cell(row=row, column=1, value=label)
        cell.font = BODY_FONT
        cell.border = BORDER
        c = ws4.cell(row=row, column=2, value=formula)
        c.font = BODY_FONT
        c.border = BORDER
        c.fill = STATUS_FILL["Calculated"]
        c.number_format = MONEY

    # ---- 三层口径块（v4 Step 4.5）：① 自制件供应商报价 ② 外购件进厂 ③ 我方总装增值
    # 边界：①② 付给供应商（含各自 OH/SG&A/物流/利润），③ 是我们自己的；
    # 外购件 PASS_RATE 默认 0 = 不叠加我方间接费与利润，防止 20.92→76.52 式重复计价。
    tri0 = ppc0 + len(ppc_rows) + 2
    buy_rows_n = sum(1 for p in parts if p["buy"])
    ws4.cell(row=tri0, column=1, value="整机成本（供应商件价 + 我方总装工艺成本）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    tri_hdr = ["层级", "金额（CNY）", "说明"]
    for i, label in enumerate(tri_hdr, start=1):
        cell = ws4.cell(row=tri0 + 1, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    make_n = sum(1 for p in parts if not p["buy"])
    assy_ct_val = PARAMS[[p[0] for p in PARAMS].index("ASSY_CT_MIN")][2]
    tri_rows = [
        ("① 自制(非标)件 — 供应商报价", None,
         f"{make_n} 项叶件；含供应商侧的辅料与设备费率"),
        ("    直接成本（材料+加工+装配+测试）", f"=$L${total_row}",
         "4-Summary 合计行 L 列（Direct）"),
        ("    + 制造间接费", f"=$B${tri0 + 2}*{ref['OH_RATE']}", "折旧/厂房/能源/质量"),
        ("    + SG&A", f"=($B${tri0 + 2}+$B${tri0 + 3})*{ref['SGA_RATE']}", "管理与销售费用"),
        ("    + 物流包装", f"=($B${tri0 + 2}+$B${tri0 + 3}+$B${tri0 + 4})*{ref['LOG_RATE']}", "包装与运输"),
        ("    小计：供应商成本", f"=SUM($B${tri0 + 2}:$B${tri0 + 5})", ""),
        ("    ÷ (1 − 利润) = 供应商报价", f"=$B${tri0 + 6}/(1-{ref['PROFIT_RATE']})", "利润率按销售额"),
        ("② 外购件 — 进厂成本", None, ""),
        (f"    采购价（{buy_rows_n} 项）",
         f"=SUMIFS('2-BOM'!$AF$4:$AF${3 + len(parts)},'2-BOM'!$O$4:$O${3 + len(parts)},\"Buy\")",
         "已是供应商价；含导热膏等外购辅料"),
        ("    + 处置与物流",
         f"=$B${tri0 + 9}*((1+{ref['PASS_RATE']})*(1+{ref['LOG_RATE']})-1)",
         "PASS_RATE 默认 0"),
        ("    小计：外购件进厂成本", f"=$B${tri0 + 9}+$B${tri0 + 10}", ""),
        ("③ 我方增值 — 总装工艺成本", None, ""),
        (f"    人工：总装（{iface_cnt} 界面点，{assy_ct_val} min）",
         f"={ref['ASSY_CT_MIN']}/60*{ref['RATE_ASSY']}", "45 元/人·h"),
        ("    人工：EOL 测试", f"={ref['EOL_CT_MIN']}/60*{ref['RATE_TEST']}", "80 元/h（含测试台）"),
        ("    辅料：密封胶 FIPG", f"={ref['FIPG_G']}/1000*{ref['FIPG_PRICE']}", "g × 元/kg"),
        ("    辅料：清洗剂/擦拭/扎带/标签", f"={ref['CONS_OTHER']}", "元/台"),
        ("    设备折旧（总装线摊销）",
         f"={ref['ASSY_INV']}/({ref['ASSY_LIFE']}*{ref['VOL_A']})",
         "投资额 ÷（折旧年限 × 年产量）"),
        ("    小计：我方增值", f"=SUM($B${tri0 + 13}:$B${tri0 + 17})", "不含我方间接费与利润"),
        ("＝ 整机成本（我方口径）", f"=$B${tri0 + 7}+$B${tri0 + 11}+$B${tri0 + 18}",
         "供应商件价 + 我方总装工艺成本"),
    ]
    for i, (label, formula, note) in enumerate(tri_rows, start=1):
        row = tri0 + i
        cell = ws4.cell(row=row, column=1, value=label)
        cell.font = BODY_FONT
        cell.border = BORDER
        if formula is None:
            cell.font = Font(name="微软雅黑", size=9, bold=True)
        c = ws4.cell(row=row, column=2, value=formula)
        c.font = BODY_FONT
        c.border = BORDER
        if formula is not None:
            c.fill = STATUS_FILL["Calculated"]
            c.number_format = MONEY
        n = ws4.cell(row=row, column=3, value=note)
        n.font = Font(name="微软雅黑", size=9, color="595959")
    tri_end = tri0 + len(tri_rows)
    # 边界行（仅判断用）：外购件全取 Low / 全取 High
    bnd0 = tri_end + 2
    ws4.cell(row=bnd0, column=1, value="边界（仅判断用，不作结论）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    for i, label in enumerate(["情形", "外购件进厂成本", "整机"], start=1):
        cell = ws4.cell(row=bnd0 + 1, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    bnd_rows = [
        ("外购件全取 Low",
         f"=SUMIFS('2-BOM'!$AD$4:$AD${3 + len(parts)},'2-BOM'!$O$4:$O${3 + len(parts)},\"Buy\")"
         f"*((1+{ref['PASS_RATE']})*(1+{ref['LOG_RATE']}))",
         f"=$B${tri0 + 7}+B{bnd0 + 2}+$B${tri0 + 18}"),
        ("外购件全取 High",
         f"=SUMIFS('2-BOM'!$AE$4:$AE${3 + len(parts)},'2-BOM'!$O$4:$O${3 + len(parts)},\"Buy\")"
         f"*((1+{ref['PASS_RATE']})*(1+{ref['LOG_RATE']}))",
         f"=$B${tri0 + 7}+B{bnd0 + 3}+$B${tri0 + 18}"),
    ]
    for i, (label, f1, f2) in enumerate(bnd_rows, start=1):
        row = bnd0 + 1 + i
        for col, v in [(1, label), (2, f1), (3, f2)]:
            c = ws4.cell(row=row, column=col, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            if col > 1:
                c.fill = STATUS_FILL["Calculated"]
                c.number_format = MONEY
    tri_note = bnd0 + len(bnd_rows) + 2
    ws4.cell(row=tri_note, column=1, value=(
        "注 1：三段 = ①供应商自制件报价 + ②外购件进厂成本 + ③我方总装工艺成本。①②付给供应商，③是我们自己的。"
    )).font = Font(name="微软雅黑", size=9, color="595959")
    ws4.cell(row=tri_note + 1, column=1, value=(
        "注 2：③ 不是只有人工。工艺成本 = 人工（总装 + EOL）+ 辅料（密封胶、清洗剂、扎带、标签）+ 设备折旧。"
    )).font = Font(name="微软雅黑", size=9, color="595959")
    ws4.cell(row=tri_note + 2, column=1, value=(
        "注 3：设备折旧 = 投资额 ÷（折旧年限 × 年产量）。年产量翻倍则折旧减半；若折旧高于人工，说明低于盈亏平衡产量，应外协总装。"
    )).font = Font(name="微软雅黑", size=9, color="595959")
    ws4.cell(row=tri_note + 3, column=1, value=(
        "注 4：外购件按进厂成本计（供应商价 + 物流），不叠加我方 OH/SG&A/利润（PASS_RATE=0）；"
        "若把外购件也套自制件的加成链，会把 20.92 元的成本算成 76.52 元。"
    )).font = Font(name="微软雅黑", size=9, color="595959")

    # ---- 成本完整度块（任务 B）：Unknown 不得被 N()/SUM 静默当 0
    # 四件套：Known/Modeled 小计 / Unknown 计数 / 覆盖率 / Complete PPC 判定。
    # 判定规则：外购件价格存在 Unknown（或外购件数 > 0 且价格全空）时，
    # PPC 只能标"部分（仅自制件）"，不得标"完整"。
    cov0 = tri_note + 4
    ws4.cell(row=cov0, column=1, value="成本完整度（Unknown ≠ 0，缺价不得静默计入）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    # 2-BOM Make/Buy 列（O 列）、价格状态列（T 列）、采购价列（R 列）、取值列（AA 列）
    bom_first, bom_last = 4, 3 + len(parts)
    cov_rows = [
        ("外购件总数", f"=COUNTIFS('2-BOM'!$O:$O,\"Buy\")"),
        ("有价外购件数（Known/Modeled）",
         f"=COUNTIFS('2-BOM'!$O:$O,\"Buy\",'2-BOM'!$AA:$AA,\">0\")"),
        ("Unknown 价格外购件数",
         f"=COUNTIFS('2-BOM'!$O:$O,\"Buy\",'2-BOM'!$T:$T,\"Unknown\")"),
        ("Known/Modeled 外购成本小计（元）",
         f"=SUMIFS('2-BOM'!$AF:$AF,'2-BOM'!$O:$O,\"Buy\",'2-BOM'!$AA:$AA,\">0\")"),
        ("成本数据覆盖率（有价外购件数 ÷ 外购件总数）",
         f"=IF(B{cov0 + 1}=0,\"\",B{cov0 + 2}/B{cov0 + 1})"),
        ("Complete PPC 是否可计算",
         f"=IF(B{cov0 + 3}=0,\"完整：全部外购件有价\",\"部分（仅自制件）：\"&B{cov0 + 3}&\" 项外购件价格 Unknown\")"),
    ]
    for i, (label, formula) in enumerate(cov_rows, start=1):
        row = cov0 + i
        cell = ws4.cell(row=row, column=1, value=label)
        cell.font = BODY_FONT
        cell.border = BORDER
        c = ws4.cell(row=row, column=2, value=formula)
        c.font = BODY_FONT
        c.border = BORDER
        c.fill = STATUS_FILL["Calculated"]
        if "覆盖率" in label:
            c.number_format = '0.0%'
        elif "小计" in label:
            c.number_format = MONEY
    ws4.cell(row=cov0 + len(cov_rows) + 1, column=1, value=(
        "读法：Unknown 计数 > 0 时，上方 PPC/应成本价只是「仅自制件下限」，不是完整产品成本；"
        "合计行里外购件金额为空（不是 0），SUM 不会把它当 0 计入——但读者必须知道缺了哪一块。"
        "Supplier Quotation 是报价不是已发生成本，即使填入也不得标 Verified。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="C00000")

    # ---- 功能域透视块：钱花在哪个功能上
    piv0 = cov0 + len(cov_rows) + 3
    ws4.cell(row=piv0, column=1, value="功能域透视（自制件 MAT/VA + 外购件进厂价 = 全成本口径）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    ws4.cell(row=piv0, column=2, value=(
        "注：4-Summary 逐件行的 G 列按 v4 留空（价格在 2-BOM!Y:AA，防①/②重复计价），"
        "故本透视的外购件金额单独从 2-BOM!AF 列聚合，Direct Cost 才含外购件。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    piv_hdr = ["Function", "Category", "MAT", "VA", "Purchased (②层)", "Direct Cost", "Share %"]
    for i, label in enumerate(piv_hdr, start=1):
        cell = ws4.cell(row=piv0 + 1, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    # 收集 Function×Category 组合（保持 BOM 出现顺序）
    combos: list[tuple[str, str]] = []
    for p in parts:
        key = (p["func"], p["cat"])
        if key not in combos:
            combos.append(key)
    bom_last = 3 + len(parts)
    pr = piv0 + 2
    for func, cat in combos:
        mat_f = (f'=SUMIFS($F${first_data}:$F${last_data},$C${first_data}:$C${last_data},"{func}",'
                 f'$D${first_data}:$D${last_data},"{cat}")')
        # SUMIFS 求和区与条件区必须同尺寸：H/I/J 三列分别 SUMIFS 后相加
        va_f = "+".join(
            f'SUMIFS(${col}${first_data}:${col}${last_data},'
            f'$C${first_data}:$C${last_data},"{func}",'
            f'$D${first_data}:$D${last_data},"{cat}")'
            for col in "HIJ"
        )
        va_f = f"={va_f}"
        # 外购件：从 2-BOM 的 Function(D)/Category(E) × Amt Base(AF) 聚合
        buy_f = (f'=SUMIFS(\'2-BOM\'!$AF$4:$AF${bom_last},'
                 f'\'2-BOM\'!$D$4:$D${bom_last},"{func}",'
                 f'\'2-BOM\'!$E$4:$E${bom_last},"{cat}")')
        dc_f = f'=IF(COUNT(C{pr}:E{pr})=0,"",SUM(C{pr}:E{pr}))'
        # 合计行位置可预先算出（piv0+2 起 len(combos) 行，合计紧随其后）
        sh_f = f'=IF(F{pr}="","",F{pr}/$F${piv0 + 2 + len(combos)})'
        for i, v in enumerate([func, cat, mat_f, va_f, buy_f, dc_f, sh_f], start=1):
            c = ws4.cell(row=pr, column=i, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            if i >= 3:
                c.fill = STATUS_FILL["Calculated"]
                c.number_format = '0.0%' if i == 7 else MONEY
        pr += 1
    piv_total = pr
    ws4.cell(row=piv_total, column=1, value="合计").font = Font(name="微软雅黑", size=9, bold=True)
    for col_idx, col_letter in [(3, "C"), (4, "D"), (5, "E"), (6, "F")]:
        c = ws4.cell(row=piv_total, column=col_idx,
                     value=f"=SUM({col_letter}{piv0 + 2}:{col_letter}{piv_total - 1})")
        c.font = Font(name="微软雅黑", size=9, bold=True)
        c.border = BORDER
        c.number_format = MONEY
        c.fill = STATUS_FILL["Calculated"]
    c = ws4.cell(row=piv_total, column=7, value=1)
    c.font = Font(name="微软雅黑", size=9, bold=True)
    c.border = BORDER
    c.number_format = '0.0%'
    c.fill = STATUS_FILL["Calculated"]

    # ---- Quotation 对比区（spec 既有要求，本次补实现）
    q0 = piv_total + 2
    ws4.cell(row=q0, column=1, value="Quotation 对比（拿到报价后填入，空 = Unknown）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    q_rows = [
        ("供应商报价（输入）", None, "Unknown"),
        ("模型应成本价", f"=O{total_row}", "Calculated"),
        ("Variance（报价 − 模型）", f'=IF(B{q0 + 1}="","",B{q0 + 1}-B{q0 + 2})', "Calculated"),
        ("Variance %（对报价）", f'=IF(B{q0 + 1}="","",(B{q0 + 1}-B{q0 + 2})/B{q0 + 1})', "Calculated"),
    ]
    for i, (label, formula, status) in enumerate(q_rows, start=1):
        row = q0 + i
        cell = ws4.cell(row=row, column=1, value=label)
        cell.font = BODY_FONT
        cell.border = BORDER
        c = ws4.cell(row=row, column=2, value=formula)
        c.font = BODY_FONT
        c.border = BORDER
        c.fill = STATUS_FILL[status]
        c.number_format = '0.0%' if "%" in label else MONEY

    info = q0 + len(q_rows) + 2
    ws4.cell(row=info, column=1, value="说明").font = Font(name="微软雅黑", size=9, bold=True)
    ws4.cell(row=info + 1, column=1, value=(
        "1) 只用叶件求和，父件（子总成行）不参与，避免重复计价。"
    )).font = Font(name="微软雅黑", size=9)
    ws4.cell(row=info + 2, column=1, value=(
        "2) 外购件价格在 Y/Z/AA 列（取证区间 Low/High + 宽区间几何均值），R 列留空防重复计价；"
        "三层块②按 AA 列合计，不叠加 OH/SG&A/利润。"
    )).font = Font(name="微软雅黑", size=9)
    ws4.cell(row=info + 3, column=1, value=(
        "3) Should-Cost Price = 工厂成本 ×(1+物流率) /(1−利润率)，利润率按销售额计，不是成本加成。"
    )).font = Font(name="微软雅黑", size=9)

    # ---- Top-down 目标价反推（无报价时的主要用法）
    t0 = info + 5
    ws4.cell(row=t0, column=1, value="目标价反推（Top-down）：填入目标/竞品单价，反推外购件允许合计").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    ws4.cell(row=t0 + 1, column=1, value=(
        "★ 只在 B/C/D 三格填实际掌握的目标价或竞品价；没有就留空，不要填估算值。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="C00000")

    price_row = t0 + 3
    hdr = ["项目", "Low", "Base", "High"]
    for i, label in enumerate(hdr, start=1):
        cell = ws4.cell(row=price_row - 1, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    rows_def = [
        ("目标单价（输入）", None),
        ("允许工厂成本", "=IF({c}{p}=\"\",\"\",{c}{p}*(1-{profit})/(1+{log}))"),
        ("允许直接成本", "=IF({c}{r1}=\"\",\"\",{c}{r1}/(1+{oh}))"),
        ("自制件直接成本（模型）", f"=$L${total_row}"),
        ("外购件允许合计", "=IF({c}{r2}=\"\",\"\",{c}{r2}-{c}{r3})"),
        ("外购件占允许直接成本", "=IF({c}{r2}=\"\",\"\",{c}{r4}/{c}{r2})"),
    ]
    for i, (label, tpl) in enumerate(rows_def):
        row = price_row + i
        cell = ws4.cell(row=row, column=1, value=label)
        cell.font = Font(name="微软雅黑", size=9, bold=(i == 0))
        cell.border = BORDER
        for col_idx, col_letter in enumerate(["B", "C", "D"], start=2):
            if tpl is None:
                c = ws4.cell(row=row, column=col_idx, value=None)
                c.fill = STATUS_FILL["Unknown"]
            else:
                c = ws4.cell(row=row, column=col_idx, value=tpl.format(
                    c=col_letter, p=price_row,
                    r1=price_row + 1, r2=price_row + 2, r3=price_row + 3, r4=price_row + 4,
                    profit=ref["PROFIT_RATE"], log=ref["LOG_RATE"], oh=ref["OH_RATE"]))
                c.fill = STATUS_FILL["Calculated"]
                c.number_format = '0.0%' if i == 5 else MONEY
            c.font = BODY_FONT
            c.border = BORDER

    judge_row = price_row + len(rows_def)
    ws4.cell(row=judge_row, column=1, value="判定").font = Font(name="微软雅黑", size=9, bold=True)
    for col_idx, col_letter in enumerate(["B", "C", "D"], start=2):
        c = ws4.cell(row=judge_row, column=col_idx, value=(
            f'=IF({col_letter}{price_row}="","待输入",'
            f'IF({col_letter}{price_row + 4}<0,"不可行：目标价低于自制件成本",'
            f'"可行：外购件需 ≤ "&TEXT({col_letter}{price_row + 4},"#,##0")))'))
        c.font = BODY_FONT
        c.border = BORDER
        c.fill = STATUS_FILL["Calculated"]

    ws4.cell(row=judge_row + 2, column=1, value=(
        "用法：目标价是「采购单价」口径，必须与模型同边界（同质量口径、同含件范围）。"
        f"反推结果 = PCB+电感+铁氧体+导热膏+线束+密封件等 {buy_rows_n} 项外购件的合计上限"
        "（外购件总数见成本完整度块）；"
        "若该上限明显低于可见器件的合理量级，说明目标价本身不具备可达性。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")

    # ---------- 5 Opportunities
    ws5 = wb.create_sheet("5-Opportunities")
    ws5["A1"] = "Cost Reduction Opportunities（单件节省一律用公式）"
    ws5["A1"].font = TITLE_FONT
    header(ws5, 3, ["ID", "题目", "现有方案", "替代方案", "工程依据", "Baseline 单价",
                    "Proposed 单价", "单件节省", "净单件节省", "年产量", "年度节省", "一次性投入",
                    "净生命周期收益", "MAT Impact", "VA Impact", "Tooling Impact",
                    "降本机制", "防重复计数组", "HG", "风险", "验证周期", "状态", "优先级"],
           [6, 26, 30, 30, 34, 12, 12, 11, 12, 10, 12, 12, 14, 12, 12, 12,
            18, 14, 6, 22, 10, 20, 8])
    is_obcm = "obcm" in product.lower()
    is_dcdc = "dc-dc" in product.lower() or "dcdc" in product.lower()
    if is_dcdc:
        opps = [
            ("O01", "上盖集成泡沫的成本显性化",
             "Top Cover 1.097 kg：铝壳 + PE 泡沫 + 钢 + 胶（4 种材料）",
             "泡沫/钢/胶改供应商预制或取消（NVH 泡沫与上盖合一）",
             "混合材料件把辅料藏在铝价里；拆开计价才能看出泡沫与胶的占比",
             "", "", "MAT", "无", "无", "材料显性化", "G1", "HG1",
             "低：需确认泡沫是否承担 NVH 功能", "4–6 周", "中"),
            ("O02", "TPV/TPS 两种密封圈合并",
             "Gasket TPV 0.015 kg + Gasket 2 TPS 0.004 kg（两种弹性体，两处密封界面）",
             "统一为一种 TPV（或一种 TPS）",
             "两种弹性体 = 两套模切模与两种库存；两者合计仅 19 g，功能高度重叠",
             "", "", "MAT", "无", "Tooling", "物料归并", "G2", "HG1",
             "低：需压缩永久变形与耐液验证", "4–6 周", "高"),
            ("O03", "铁氧体磁芯规格归并",
             "Ferrite Core 0.050 kg ×2 + Ferrite Core 2 0.096 kg（两种磁芯，三种料号含 2 个磁芯件）",
             "归并为单一磁芯规格或共用模压件",
             "两种磁芯 = 两套模具与两种烧结参数；磁芯单价主要随体积而非形状变",
             "", "", "MAT", "无", "Tooling", "物料归并", "G3", "HG0",
             "中：需重核电感量与 EMI", "6–10 周", "中"),
            ("O04", "Se-Cu58 与 C10300 铜排统一牌号",
             "Busbar（Se-Cu58）0.108 kg + Busbar 2（C10300）0.037 kg，两种铜合金",
             "统一为 C10300（或 Se-Cu58）",
             "两种铜牌号 = 两套采购规格；载流相同时统一牌号可减库存与换料",
             "", "", "MAT", "无", "无", "物料归并", "G4", "HG0",
             "低：需确认电导率与成形性", "6–8 周", "中"),
            ("O05", "两铝盖由全铝计价改分层计价",
             "Top Cover 1.097 kg（铝+PE 泡沫+钢+胶）+ Bottom Cover 1.909 kg（铝+铁氧体+密封剂），"
             "合并 3.006 kg 现按全铝 23.9 元/kg 计价",
             "按实际材料分层核算（铝件按铝价，泡沫/钢/胶/铁氧体按各自单价）",
             "这是计价口径修正而非设计变更：全铝计价高估铝、低估辅料，应先把盖子的成本算对",
             "", "", "MAT", "无", "无", "计价修正", "G5", "HG2",
             "低：仅需拆件分层称重", "2–3 周", "高"),
        ]
    elif is_obcm:
        opps = [
            ("O01", "Casing 壁厚减薄 15%",
             "压铸铝壳 4.789 kg，占总质量 44.6%",
             "拓扑优化 + 局部减薄，加强筋补刚度",
             "壳体是最大质量与成本项；铝件成本 ≈ 质量 × 铝价，减重收益可直接算",
             f"={ref['AL_KG']}*{ref['AL_PRICE']}",
             f"={ref['AL_KG']}*0.85*{ref['AL_PRICE']}",
             "MAT", "无", "无", "材料减重", "G1", "HG1",
             "中：需模态、密封面平面度与气密验证", "10–14 周", "高"),
            ("O02", "两个电感合并为集成磁件",
             "Inductor 1.250 kg + Inductor 2 1.314 kg（均为 PULSE）",
             "单一集成磁件（PFC 与 LLC 共磁芯）",
             "磁件占 23.9% 质量；双磁件=双份磁芯/绕线/灌封与两个料号",
             "", "", "无", "VA", "Tooling", "零件合并", "G2", "HG0",
             "高：磁集成需重做 EMI 与热设计", "20–28 周", "中"),
            ("O03", "TIM 由 14 处减为 4 处",
             "Thermal Paste ×14，0.196 kg",
             "集中式导热垫/相变膜，减少点胶点",
             "14 个点胶点 = 14 次参数控制与错料机会；工序成本被低估",
             "", "", "无", "VA", "无", "工序简化", "G3", "HG0",
             "中：需界面温升实测", "6–10 周", "中"),
            ("O04", "PCB 3 合并进主板",
             "PCB 1.877 kg + PCB 3 0.136 kg（9 个界面）",
             "单主板 + 局部屏蔽",
             "每增加一块板：治具、料号、板间连接器与测试工时",
             "", "", "无", "VA", "Tooling", "零件合并", "G4", "HG0",
             "中：布局与 EMC 需重做", "12–16 周", "中"),
            ("O05", "Cover 2 铝件改塑胶或减薄",
             "Aluminium ENAC46000，0.537 kg",
             "塑胶件（PC-ABS）或铝件减薄",
             "该件非承力、非散热主路径时，铝件属过度设计",
             "", "", "MAT", "无", "Tooling", "材料替代", "G5", "HG0",
             "低：需屏蔽与耐温评估", "8–10 周", "低"),
        ]
    else:
        opps = [
            ("O01", "铜冷却板改铝基板 + 铜嵌件",
             "4 件全铜冷却板，共 0.766 kg",
             "铝基板 + 局部铜嵌件（铜用量按 40% 保留）",
             "铜密度 3.3× 铝；若热流仅在器件下方集中，全铜属过量设计",
             f"={ref['CU_KG']}*{ref['CU_PRICE']}",
             f"={ref['CU_KG']}*0.4*{ref['CU_PRICE']}+{ref['CU_KG']}*0.6*1.5*{ref['AL_PRICE']}",
             "MAT", "无", "无", "材料替代", "G1", "HG2",
             "中：铝板按 1.5× 厚度补偿导热率差，需热仿真确认 q'' 与结温余量", "8–12 周", "高"),
            ("O02", "三种 TIM 合并为一种",
             "Thermal Paste / 2 / 3 三种牌号",
             "统一为一种导热脂（或相变膜）",
             "三种材料 = 三套胶量/固化/有效期参数，错料与调试成本被低估",
             "", "", "MAT", "无", "无", "物料归并", "G2", "HG0",
             "低：需重做界面温升验证", "4–6 周", "中"),
            ("O03", "12 颗分立 MOSFET 改模块化",
             "12 颗分立器件 + 4 个支架 + 16 定位销",
             "1 个功率模块 + 压紧结构",
             "分立方案装配点与筛选配对数成倍；但器件单价低、供应链开放",
             "", "", "无", "VA", "Tooling", "集成度提升", "G3", "HG0",
             "高：需重新验证并联均流与热", "16–24 周", "中"),
            ("O04", "三块 PCB 合并为两块",
             "PCB / PCB 2 / PCB 3 + 2 个钢屏蔽盖",
             "两块板 + 1 个屏蔽盖",
             "每增加一块板的固定成本（治具/料号/测试工时）通常被低估",
             "", "", "无", "VA", "Tooling", "零件合并", "G4", "HG0",
             "中：EMC 与布局需重做", "12–16 周", "中"),
            ("O05", "三件铝壳改一体腔",
             "Housing + Bottom Cover + Top Cover",
             "一体压铸腔 + 单盖",
             "少一道对合面与一套密封/检具；但模具吨位与报废牵连上升",
             "", "", "无", "VA", "Tooling", "零件合并", "G5", "HG0",
             "中：取决产量与良率", "12–20 周", "低"),
        ]
    r = 4
    for oid, title, now, alt, basis, base, prop, mat_imp, va_imp, tool_imp, mech, dedup, hg, risk, weeks, prio in opps:
        status = "Modeled Saving" if base else "Hypothetical Opportunity"
        values = [oid, title, now, alt, basis, base, prop,
                  f'=IF(OR(F{r}="",G{r}=""),"",F{r}-G{r})',
                  f'=IF(H{r}="","",IF(L{r}="",H{r},H{r}-L{r}/{ref["VOL_L"]}))',
                  f"={ref['VOL_A']}",
                  f'=IF(H{r}="","",H{r}*J{r})',
                  None,
                  f'=IF(OR(H{r}="",L{r}=""),"",H{r}*{ref["VOL_L"]}-L{r})',
                  mat_imp, va_imp, tool_imp, mech, dedup, hg, risk, weeks, status, prio]
        body(ws5, r, values, "Calculated" if status == "Modeled Saving" else "Estimated")
        for c in (6, 7, 8, 9, 11, 12, 13):
            ws5.cell(row=r, column=c).number_format = MONEY
        r += 1
    # 防重复计数说明（任务 C）
    ws5.cell(row=r + 1, column=1, value=(
        "防重复计数规则：①「防重复计数组」相同的措施不得在合计中重复计节省"
        "（同组多措施只按组内最大可叠加项计）；②同一降本机制不得同时计入 MAT 与 VA"
        "（MAT Impact / VA Impact 二选一为主归属，另一列写「无」或写次级影响的定性说明）；"
        "③外购总成价与内部展开成本不得双重计价（见 2-BOM Make/Buy 与 7-Evidence）；"
        "④跨措施叠加时先查组号，同组措施需声明是否可叠加及依据。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="C00000")
    ws5.cell(row=r + 2, column=1, value=(
        "HG 定义（Hardness Grade，与 Bosch CTM 报告口径一致）：HG0 描述想法 / HG1 毛节省潜力 / "
        "HG2 净节省潜力 / HG3 时间表已排 / HG4 已实施 / HG5 节省已兑现。"
        "技术证据可信度、经济测算成熟度、实施状态三者分别记录：HG 只记实施成熟度，"
        "不替代「状态」列（Verified/Modeled/Hypothetical）的证据分级。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")

    # ---------- 6 Sensitivity
    ws6 = wb.create_sheet("6-Sensitivity")
    ws6["A1"] = "Sensitivity（只动关键驱动，不是整体百分比）"
    ws6["A1"].font = TITLE_FONT
    header(ws6, 3, ["驱动因素", "Low", "Base", "High", "对单件成本影响(元, High−Low)",
                    "影响占比", "说明"], [30, 12, 12, 12, 20, 10, 44])
    rows = [
        ("铜价 ±30%", 0.7, 1.0, 1.3,
         f"={ref['CU_KG']}*{ref['CU_PRICE']}*0.6", "铜件为价格弹性最高项（质量 × 单价双高）"),
        ("铝价 ±15%", 0.85, 1.0, 1.15,
         f"={ref['AL_KG']}*{ref['AL_PRICE']}*0.3", "铝件用量大，单价波动直接进材料成本"),
        ("装配工时 ±20%", 0.8, 1.0, 1.2,
         f"=IF('4-Summary'!$I${total_row}=\"\",\"\",'4-Summary'!$I${total_row}*0.4)",
         "仅自制件装配工时部分"),
        ("良率 95% → 90%", 0.90, 0.95, 1.0,
         f"=IF('4-Summary'!$N${total_row}=\"\",\"\",'4-Summary'!$N${total_row}*(0.95/0.90-1))",
         "良率链乘性影响，非简单加成"),
        ("年产量 0.5× / 2×", 0.5, 1.0, 2.0,
         f"=IF('4-Summary'!$K${total_row}=\"\",\"\",'4-Summary'!$K${total_row}*(2-0.5))",
         "仅影响工装摊销；当前工装费未填，故为空"),
    ]
    r = 4
    for label, lo, base, hi, formula, note in rows:
        values = [label, lo, base, hi, formula,
                  f'=IF(E{r}="","",E{r}/IF(\'4-Summary\'!$N${total_row}="",1,\'4-Summary\'!$N${total_row}))',
                  note]
        body(ws6, r, values, "Calculated")
        ws6.cell(row=r, column=5).number_format = MONEY
        ws6.cell(row=r, column=6).number_format = '0.0%'
        r += 1
    ws6.cell(row=r + 1, column=1, value=(
        "说明：影响额 = 单位驱动变化量 × 对应成本桶，属线性近似；"
        "用于排序「先查哪个数据」，不作为报价依据。"
    )).font = Font(name="微软雅黑", size=9, italic=True)

    # ---- 工装摊销阶梯：同一工装投入，产量不同单件完全不同
    t0 = r + 3
    ws6.cell(row=t0, column=1, value="工装摊销阶梯（CapEx 一次性 vs 摊销/件 两本账）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    ws6.cell(row=t0 + 1, column=1, value=(
        "工装 CapEx 未填时全表为空；填入后可见：同一措施在 50k 与 200k 生命周期下净收益完全不同。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    tool_capex_sum = f"SUM('2-BOM'!W4:W{3 + len(parts)})"
    for i, label in enumerate(["生命周期产量", "摊销/件 (CNY)", "对 PPC 影响 (CNY)"], start=1):
        cell = ws6.cell(row=t0 + 3, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for j, vol in enumerate([50000, 100000, 200000, 500000], start=0):
        row = t0 + 4 + j
        vals = [
            vol,
            f"=IF({tool_capex_sum}=0,\"\",{tool_capex_sum}*1000/MIN({ref['TOOL_LIFE_SHOTS']},A{row}))",
            f"=IF(B{row}=\"\",\"\",B{row}*(1+{ref['OH_RATE']})*(1+{ref['LOG_RATE']})/(1-{ref['PROFIT_RATE']}))",
        ]
        for i, v in enumerate(vals, start=1):
            c = ws6.cell(row=row, column=i, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            if i > 1:
                c.fill = STATUS_FILL["Calculated"]
                c.number_format = MONEY
            else:
                c.number_format = '#,##0'

    # ---- 外购件合计杠杆（逐项价格未定或需交叉校验时，用合计作为单一变量）
    l0 = t0 + 9
    ws6.cell(row=l0, column=1, value="外购件合计杠杆（逐项价格未定或需交叉校验时，用合计作为单一变量）").font = Font(
        name="微软雅黑", size=11, bold=True, color="1F3864")
    ws6.cell(row=l0 + 1, column=1, value=(
        f"{buy_rows_n} 项外购件已按 2-BOM!Y:AA 取证填价（几何均值 {ref['CENTER_METHOD']} 取点值）。"
        "本阶梯表用于交叉校验：填入一个假设合计，即可得到整机应成本，"
        "与 4-Summary 三层块的 ② 对比，即可检验取证件是否偏离量级。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")

    selfmade = f"'4-Summary'!$L${total_row}"
    coef = f"((1+{ref['OH_RATE']})*(1+{ref['LOG_RATE']})/(1-{ref['PROFIT_RATE']}))"

    inp = l0 + 3
    ws6.cell(row=inp, column=1, value="假设外购件合计（CNY/件，输入）").font = Font(name="微软雅黑", size=9, bold=True)
    c = ws6.cell(row=inp, column=2, value=None)
    c.fill = STATUS_FILL["Unknown"]
    c.border = BORDER
    c.number_format = MONEY

    calc = [
        ("直接成本", f"=IF(B{inp}=\"\",\"\",{selfmade}+B{inp})"),
        ("工厂成本", f"=IF(B{inp}=\"\",\"\",B{inp + 1}*(1+{ref['OH_RATE']}))"),
        ("应成本价", f"=IF(B{inp}=\"\",\"\",B{inp + 2}*(1+{ref['LOG_RATE']})/(1-{ref['PROFIT_RATE']}))"),
        ("外购件占直接成本", f"=IF(B{inp}=\"\",\"\",B{inp}/B{inp + 1})"),
        ("边际系数：外购每 +1 元 → 应成本 +X 元", f"={coef}"),
        ("外购每 +10 元 → 应成本变化", f"=10*B{inp + 5}"),
    ]
    for i, (label, formula) in enumerate(calc, start=1):
        row = inp + i
        cell = ws6.cell(row=row, column=1, value=label)
        cell.font = BODY_FONT
        cell.border = BORDER
        c = ws6.cell(row=row, column=2, value=formula)
        c.font = BODY_FONT
        c.border = BORDER
        c.fill = STATUS_FILL["Calculated"]
        c.number_format = '0.0%' if label.startswith("外购件占") else MONEY

    # 阶梯表
    ladder = inp + 8
    ws6.cell(row=ladder, column=1, value="阶梯表：外购件合计 → 整机应成本价").font = Font(
        name="微软雅黑", size=9, bold=True)
    for i, label in enumerate(["外购件合计", "直接成本", "工厂成本", "应成本价"], start=1):
        cell = ws6.cell(row=ladder + 1, column=i, value=label)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for j, p_val in enumerate([0, 200, 400, 600, 800, 1000], start=0):
        row = ladder + 2 + j
        vals = [
            p_val,
            f"={selfmade}+A{row}",
            f"=B{row}*(1+{ref['OH_RATE']})",
            f"=C{row}*(1+{ref['LOG_RATE']})/(1-{ref['PROFIT_RATE']})",
        ]
        for i, v in enumerate(vals, start=1):
            c = ws6.cell(row=row, column=i, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            if i > 1:
                c.fill = STATUS_FILL["Calculated"]
                c.number_format = MONEY
        ws6.column_dimensions[get_column_letter(1)].width = max(
            ws6.column_dimensions[get_column_letter(1)].width or 0, 34)
        for i in range(2, 5):
            ws6.column_dimensions[get_column_letter(i)].width = max(
                ws6.column_dimensions[get_column_letter(i)].width or 0, 16)

    ws6.cell(row=ladder + 9, column=1, value=(
        "阶梯表不含任何外购件单价假设，只是把「合计」当变量；"
        "0 元那一行 = 当前模型（仅自制件）的下限，数值随 1-Assumptions 参数联动。"
    )).font = Font(name="微软雅黑", size=9, italic=True, color="595959")

    # ---------- 7 Evidence
    ws7 = wb.create_sheet("7-Evidence")
    ws7["A1"] = "Evidence & Open Questions"
    ws7["A1"].font = TITLE_FONT
    header(ws7, 3, ["项", "内容", "来源/位置", "状态", "置信度", "影响成本项", "待确认问题", "负责人"],
           [24, 48, 34, 12, 10, 20, 40, 10])
    total_kg = sum(r["weight"] * r["qty"] for r in bom_rows)
    make_n = sum(1 for p in parts if not p["buy"])
    buy_n = len(parts) - make_n
    buy_kg = total_kg - sum(p["weight"] * p["qty"] for p in parts if not p["buy"])
    supplier_list = "、".join(sorted({r["supplier"].split(",")[0].strip()
                                      for r in bom_rows if r["supplier"]})) or "无"
    # 价格取证汇总（Step 5.7）：Class 分布 + D 类占比判定（>30% 必须声明并出 RFQ 清单）
    # 硬性规则 12 的口径是**金额占比**（D 类金额 ÷ 已取证总金额），不是项数占比
    priced = [p for p in parts if p["buy"] and p.get("price_class")]
    cls_n = {c: sum(1 for p in priced if p["price_class"] == c) for c in "ABCD"}
    d_n = cls_n["D"]

    def _ev_amount(p: dict) -> float:
        ev = prices.get(p["part"], {})
        lo, hi = ev.get("low"), ev.get("high")
        if not lo or not hi:
            return 0.0
        return math.sqrt(lo * hi) * p["qty"]

    total_amt = sum(_ev_amount(p) for p in priced)
    d_amt = sum(_ev_amount(p) for p in priced if p["price_class"] == "D")
    d_ratio = (d_amt / total_amt) if total_amt else 0
    # RFQ 清单：D 类按金额降序（宽区间取几何均值近似排序）
    rfq = sorted(
        (p for p in priced if p["price_class"] == "D"),
        key=lambda p: -_ev_amount(p))
    rfq_text = "；".join(f"{p['name']}×{p['qty']}" for p in rfq) or "无"
    evidence = [
        ("质量", f"叶件 {len(parts)} 个，合计 {total_kg:.3f} kg（父件不参与求和）",
         "Teardown Data 宽表 + 上游 Word 记录", "Verified", "高", "全部",
         "是否有未称重的散装紧固件未计入", ""),
        ("BOM 结构", f"自制 {make_n} 项 / 外购 {buy_n} 项；外购占质量 {buy_kg:.3f} kg"
         f"（{buy_kg / total_kg * 100 if total_kg else 0:.1f}%）",
         "本工作簿 2-BOM", "Verified", "高", "MAT 外购件",
         "外购件是否已含全部装配与测试，需供应商 BOM 确认", ""),
        ("供应商", supplier_list, "上游 Word 记录 / Teardown Data Manufacturer 字段",
         "Verified", "中", "外购件价格",
         "两铝盖（Top/Bottom Cover）无供应商标注，压铸是否外包需确认", ""),
        ("混合材料计价偏差", "Top Cover（铝+PE 泡沫+钢+胶）与 Bottom Cover（铝+铁氧体+密封剂）"
         "均为 Mixed Material，材料成本按全铝计价",
         "本工作簿 2-BOM!G 列 + Teardown Data Material Species", "Estimated", "中", "MAT 材料",
         "泡沫/钢/胶/铁氧体与铝价差大：全铝计价高估铝、低估辅料；需分层称重或拆件记录修正", ""),
        ("功率级位置", "清单无功率模块/冷板/连接器件号，但 Thermal Paste 0.107 kg（3 处）"
         "是功率级存在的硬证据；PCB 2（290×233 mm，31 界面点）大概率承载功率级",
         "Teardown Data + 2-BOM", "Estimated", "中", "MAT 外购件",
         "功率器件是板载（集成在 PCB 2）还是独立模块，需 PCB 正反面照片或器件丝印确认", ""),
        ("工艺路线", "各件 Routing 为本模型假设", "本工作簿 2-BOM!N 列", "Estimated", "中", "VA 全部",
         "需供应商工艺流程图确认（压铸后是否需 T6、焊接还是螺钉）", ""),
        ("节拍", "压铸 60 s / 冲压 8 s / 注塑 45 s / CNC 120 s", "3-ProcessCost", "Estimated", "低", "VA 加工",
         "需供应商 OEE 与瓶颈节拍实测", ""),
        ("费率", "压铸 180、CNC 90、注塑 110、装配 45 CNY/h", "1-Assumptions", "Estimated", "低", "VA 全部",
         "需确认费率是否含人工、能耗、维护与间接费（口径比数值更重要）", ""),
        ("材料单价", "铝 23.9 / 铜 110.5 / 钢 8 CNY/kg（铝铜为长江现货 2026-10-28，Verified）",
         "1-Assumptions", "Verified", "中", "MAT 材料",
         "钢/塑料/铁氧体仍为 Estimated，需当期采购合同价与牌号", ""),
        ("外购件价格", f"{buy_n} 项外购件已全部取证填价（Low/High 区间 + A/B/C/D 证据类）；"
         f"Class 分布 A:{cls_n['A']} / B:{cls_n['B']} / C:{cls_n['C']} / D:{cls_n['D']}",
         "2-BOM!Y:AC 列 + 价格取证 CSV", "Estimated", "中", "MAT 外购件",
         "全部为宽区间（High/Low > 1.5×），按几何均值取点值；"
         "无 Actual Procurement / Supplier Quotation，量产定点前必须询价", ""),
        ("Class D 占比声明", f"D 类（先前估计）占已取证金额 {d_ratio * 100:.0f}%"
         + ("（> 30%，外购价以旧估计为主，结论降级）" if d_ratio > 0.3 else ""),
         "2-BOM!AC 列统计", "Estimated", "低" if d_ratio > 0.3 else "中", "MAT 外购件",
         f"RFQ 优先级（按金额降序）：{rfq_text}", ""),
        ("价格状态口径", "七类：Actual Procurement / Supplier Quotation / Comparable / "
         "Bottom-up Estimate / Parametric Estimate / Assumption / Unknown（2-BOM!T 列）",
         "2-BOM!T 列", "Verified", "高", "MAT 外购件",
         "Supplier Quotation 是报价不是已发生成本，不得标 Verified；"
         "Assumption 值只用于敏感性，不得进 PPC 基线", ""),
        ("工装", "未取得模具/夹具/检具费用", "2-BOM!W 列为空", "Unknown", "无", "工装摊销",
         "需模具报价与寿命（模次）", ""),
        ("年产量", "假设 20000 件/年，生命周期 100000 件", "1-Assumptions", "Unknown", "无", "工装摊销/费率",
         "需项目 SOR 或产量规划", ""),
        ("良率", "工序良率 95–99%、装配 98% 为本模型假设", "3-ProcessCost", "Estimated", "低", "VA 全部",
         "需供应商直通率与返修率", ""),
        ("FIPG 假设", "FIPG_G = 0：密封靠成型圈（TPV/TPS 垫圈）+ O 圈，无点胶界面假设",
         "1-Assumptions FIPG_G", "Estimated", "中", "③ 我方增值",
         "Bottom Cover species 含 Sealants，若拆解确认有胶道需改 FIPG_G > 0 并重算③", ""),
        ("报价", "无供应商报价，未做 Modeled vs Quotation 差异",
         "—", "Unknown", "无", "Quotation 对比",
         "拿到报价后填入 4-Summary Quotation 区", ""),
        ("型号/供应商反查", "Inductor 件号 E30500102 在 findchips 公开检索无结果（2026-11-03）；"
         "OEM 内部工程号不在公开网络",
         "线上检索 2026-11-03（findchips.com）", "Unknown", "无", "外购件价格",
         "外购件型号需靠微距照片或 Cares4 报告 PDF；PCB 供应商 LEAR/KINGSHINE 可尝试商务渠道询价", ""),
    ]
    r = 4
    for item in evidence:
        body(ws7, r, list(item), item[3])
        r += 1
    ws7.freeze_panes = "A4"

    for ws in wb:
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(out)
    _emit(f"已写入 {out}")
    _emit(f"  叶件 {len(parts)} 个；工序 {sum(len(p['steps']) for p in parts)} 条")


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 Should-Cost 七表工作簿")
    ap.add_argument("--bom", type=Path, required=True, help="叶件清单 CSV（含 Total Weight (kg) 列）")
    ap.add_argument("--prices", type=Path, default=None,
                    help="价格取证 CSV（CS Part + Price Low/High/Class/Basis/Status，Step 5.7）")
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--product", default="产品")
    args = ap.parse_args()

    if not args.bom.exists():
        raise SystemExit(f"找不到 BOM：{args.bom}")
    rows = load_bom(args.bom)
    if not rows:
        raise SystemExit("BOM 为空")
    if args.prices and not args.prices.exists():
        raise SystemExit(f"找不到价格取证文件：{args.prices}")
    prices = load_prices(args.prices) if args.prices else {}
    args.out.parent.mkdir(parents=True, exist_ok=True)
    build(rows, args.out, args.product, prices)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
