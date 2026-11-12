#!/usr/bin/env python3
"""把 Markdown 报告导出为 Excel（默认）；需要时再额外导出 Word。

只支持本 skill 产出用到的子集：标题、段落、引用块、有序/无序列表、管道表格、
行内 **粗体** 与 `代码`。不追求通用 Markdown 兼容。

用法:
    python tools/export_report.py reports/xxx.md            # 只出 xlsx
    python tools/export_report.py reports/xxx.md --docx     # 额外出 docx
    python tools/export_report.py reports/xxx.md --docx out.docx

依赖: openpyxl（必需）、python-docx（仅 --docx 时）
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 解析

LIST_RE = re.compile(r"^([-*]|\d+\.)\s+")
SEP_CELL_RE = re.compile(r"^:?-{2,}:?$")


def parse_blocks(text: str) -> list[tuple]:
    lines = text.splitlines()
    blocks: list[tuple] = []
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue

        if stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            blocks.append(("h", level, stripped[level:].strip()))
            i += 1
            continue

        if stripped.startswith("|"):
            rows: list[list[str]] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(SEP_CELL_RE.match(c) or not c for c in cells):
                    rows.append(cells)
                i += 1
            if rows:
                width = max(len(r) for r in rows)
                rows = [r + [""] * (width - len(r)) for r in rows]
                blocks.append(("table", rows))
            continue

        if stripped.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append(("quote", " ".join(b for b in buf if b)))
            continue

        if LIST_RE.match(stripped):
            items, ordered = [], bool(re.match(r"^\d+\.", stripped))
            while i < len(lines) and LIST_RE.match(lines[i].strip()):
                items.append(LIST_RE.sub("", lines[i].strip()))
                i += 1
            blocks.append(("list", items, ordered))
            continue

        buf = []
        while i < len(lines):
            cur = lines[i].strip()
            if not cur or cur.startswith(("#", "|", ">")) or LIST_RE.match(cur):
                break
            buf.append(cur)
            i += 1
        blocks.append(("p", " ".join(buf)))
    return blocks


def plain(text: str) -> str:
    text = text.replace("<br>", " / ")
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    return text.replace("`", "")


def add_runs(paragraph, text: str, size: float | None = None, bold: bool = False,
             color: tuple[int, int, int] | None = None) -> None:
    for part in re.split(r"(\*\*.+?\*\*)", text):
        if not part:
            continue
        is_bold = bold
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part, is_bold = part[2:-2], True
        run = paragraph.add_run(part.replace("`", ""))
        run.bold = is_bold
        if size:
            run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor(*color)
        run.font.name = "微软雅黑"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "微软雅黑")


# ---------------------------------------------------------------- Excel

COL_WIDTHS = {
    2: [12, 60],
    3: [26, 34, 44],
    4: [24, 34, 34, 44],
    5: [28, 46, 40, 40, 40],
    7: [7, 30, 26, 30, 26, 30, 22],
}
HEADER_FILL = PatternFill("solid", fgColor="305496")
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
SUB_FONT = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
BODY_FONT = Font(name="微软雅黑", size=9)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def disp_len(value: object) -> int:
    return sum(2 if ord(ch) > 127 else 1 for ch in str(value))


def est_lines(text: object, width: float) -> int:
    if text is None or str(text) == "":
        return 1
    return sum(max(1, math.ceil(disp_len(seg) / max(4.0, width * 1.0)))
               for seg in str(text).split("\n"))


def safe_title(name: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip() or "Sheet"
    return cleaned[:31]


def write_xlsx(blocks: list[tuple], path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    # 按二级标题切分；文档开头的标题并入第一节，避免留一个只有标题的空表
    sections: list[tuple[str, list[tuple]]] = []
    pending: list[tuple] = []
    for block in blocks:
        if block[0] == "h" and block[1] == 2:
            sections.append((safe_title(block[2]), [block]))
        elif sections:
            sections[-1][1].append(block)
        else:
            pending.append(block)
    if not sections:
        sections = [("总览", pending)]
    else:
        sections[0][1][:0] = pending

    used: set[str] = set()
    for name, sheet_blocks in sections:
        title = name
        suffix = 2
        while title in used:
            title = safe_title(f"{name}-{suffix}")
            suffix += 1
        used.add(title)

        ws = wb.create_sheet(title)
        ws.sheet_view.showGridLines = False
        col_count = 5
        for block in sheet_blocks:
            if block[0] == "table":
                col_count = max(col_count, len(block[1][0]))
        widths = COL_WIDTHS.get(col_count, [28] * col_count)
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        row = 1
        for block in sheet_blocks:
            kind = block[0]
            if kind == "h":
                level, text = block[1], plain(block[2])
                cell = ws.cell(row=row, column=1, value=text)
                cell.font = TITLE_FONT if level <= 2 else SUB_FONT
                cell.alignment = Alignment(vertical="center")
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_count)
                ws.row_dimensions[row].height = 26 if level <= 2 else 20
                row += 2
            elif kind in ("p", "quote"):
                text = plain(block[1])
                cell = ws.cell(row=row, column=1, value=text)
                cell.font = Font(name="微软雅黑", size=9, italic=(kind == "quote"),
                                 color="595959" if kind == "quote" else "000000")
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_count)
                ws.row_dimensions[row].height = min(409, max(15, est_lines(text, sum(widths)) * 13 + 4))
                row += 1
            elif kind == "list":
                items, ordered = block[1], block[2]
                for n, item in enumerate(items, start=1):
                    prefix = f"{n}. " if ordered else "• "
                    cell = ws.cell(row=row, column=1, value=prefix + plain(item))
                    cell.font = BODY_FONT
                    cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_count)
                    ws.row_dimensions[row].height = min(
                        409, max(15, est_lines(item, sum(widths)) * 13 + 4)
                    )
                    row += 1
                row += 1
            elif kind == "table":
                rows = block[1]
                for r_idx, cells in enumerate(rows):
                    is_header = r_idx == 0
                    for c_idx, value in enumerate(cells):
                        cell = ws.cell(row=row, column=c_idx + 1, value=plain(value))
                        cell.border = BORDER
                        cell.alignment = Alignment(wrap_text=True, vertical="top" if not is_header else "center")
                        if is_header:
                            cell.font = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
                            cell.fill = HEADER_FILL
                        else:
                            cell.font = BODY_FONT
                    ws.row_dimensions[row].height = min(
                        409,
                        max(18, max(est_lines(c, w) for c, w in zip(cells, widths)) * 13 + 6),
                    )
                    row += 1
                row += 1

        ws.freeze_panes = "A2"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(path)


# ---------------------------------------------------------------- Word

DOC_COL_WIDTHS = {
    2: [6.0, 20.0],
    3: [5.0, 9.0, 12.5],
    4: [4.2, 7.2, 7.0, 8.3],
    5: [4.8, 6.4, 5.1, 5.1, 5.2],
    7: [1.3, 4.6, 4.0, 4.4, 3.8, 4.4, 3.0],
}


def no_split(row) -> None:
    """禁止行跨页断开：否则一行会从页底开始、被切成两页。"""
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def set_cell_bg(cell, hex_color: str) -> None:
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def mark_header_row(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def write_docx(blocks: list[tuple], path: Path, footer_text: str) -> None:
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Cm(29.7), Cm(21)
    for attr in ("top_margin", "bottom_margin"):
        setattr(section, attr, Cm(1.8))
    section.left_margin = section.right_margin = Cm(1.5)

    normal = doc.styles["Normal"]
    normal.font.name = "微软雅黑"
    normal.font.size = Pt(10)
    normal.element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "微软雅黑")

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_runs(footer, footer_text, size=8, color=(0x80, 0x80, 0x80))
    add_runs(footer, "　　p.", size=8, color=(0x80, 0x80, 0x80))
    run = footer.add_run()
    for tag, attr in (("w:fldChar", "begin"), ("w:instrText", None), ("w:fldChar", "end")):
        el = OxmlElement(tag)
        if attr:
            el.set(qn("w:fldCharType"), attr)
        else:
            el.set(qn("xml:space"), "preserve")
            el.text = " PAGE "
        run._r.append(el)
    run.font.size = Pt(8)

    table_index = 0
    for block in blocks:
        kind = block[0]
        if kind == "h":
            level, text = block[1], plain(block[2])
            if level == 1:
                para = doc.add_paragraph()
                add_runs(para, text, size=18, bold=True, color=(0x1F, 0x38, 0x64))
            else:
                para = doc.add_paragraph()
                add_runs(para, text, size=13 if level == 2 else 11, bold=True,
                         color=(0x1F, 0x38, 0x64))
                para.paragraph_format.space_before = Pt(10 if level == 2 else 6)
        elif kind == "p":
            para = doc.add_paragraph()
            add_runs(para, block[1], size=10)
            para.paragraph_format.space_after = Pt(4)
        elif kind == "quote":
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Cm(0.6)
            add_runs(para, block[1], size=9, color=(0x59, 0x59, 0x59))
        elif kind == "list":
            items, ordered = block[1], block[2]
            for n, item in enumerate(items, start=1):
                para = doc.add_paragraph()
                para.paragraph_format.left_indent = Cm(0.6)
                para.paragraph_format.space_after = Pt(2)
                add_runs(para, (f"{n}. " if ordered else "• ") + item, size=10)
        elif kind == "table":
            rows = block[1]
            cols = len(rows[0])
            table_index += 1
            table = doc.add_table(rows=1, cols=cols)
            table.style = "Table Grid"
            table.autofit = False
            layout = OxmlElement("w:tblLayout")
            layout.set(qn("w:type"), "fixed")
            table._tbl.tblPr.append(layout)
            widths = DOC_COL_WIDTHS.get(cols, [26.5 / cols] * cols)

            header = table.rows[0]
            mark_header_row(header)
            no_split(header)
            for idx, value in enumerate(rows[0]):
                cell = header.cells[idx]
                cell.width = Cm(widths[idx])
                set_cell_bg(cell, "305496")
                para = cell.paragraphs[0]
                para.paragraph_format.space_after = Pt(0)
                add_runs(para, value, size=8.5, bold=True, color=(0xFF, 0xFF, 0xFF))

            for cells in rows[1:]:
                row = table.add_row()
                no_split(row)
                for idx, value in enumerate(cells):
                    cell = row.cells[idx]
                    cell.width = Cm(widths[idx])
                    para = cell.paragraphs[0]
                    para.paragraph_format.space_after = Pt(0)
                    add_runs(para, value, size=8.5)

            doc.add_paragraph().paragraph_format.space_after = Pt(2)

    doc.save(path)


# ---------------------------------------------------------------- CLI

def main() -> int:
    parser = argparse.ArgumentParser(description="Markdown 报告 → Excel（默认）/ Word（可选）")
    parser.add_argument("src", type=Path, help="Markdown 报告路径")
    parser.add_argument("--xlsx", type=Path, default=None, help="Excel 输出路径，缺省同名")
    parser.add_argument(
        "--docx", type=Path, nargs="?", const=Path("__auto__"), default=None,
        help="额外导出 Word；不带值时用同名 .docx",
    )
    parser.add_argument("--footer", default="ME 团队内部资料 · 拆解对标分析", help="Word 页脚文字")
    args = parser.parse_args()

    if not args.src.exists():
        raise SystemExit(f"找不到文件：{args.src}")

    blocks = parse_blocks(args.src.read_text(encoding="utf-8"))
    if not blocks:
        raise SystemExit(f"{args.src} 没有解析出任何内容")

    # 默认只出 Excel；Word 需显式 --docx
    xlsx_path = args.xlsx or args.src.with_suffix(".xlsx")
    write_xlsx(blocks, xlsx_path)
    print(f"已写入 {xlsx_path}")

    if args.docx is not None:
        docx_path = args.src.with_suffix(".docx") if args.docx == Path("__auto__") else args.docx
        write_docx(blocks, docx_path, args.footer)
        print(f"已写入 {docx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
