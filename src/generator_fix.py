"""
generator_fix.py — 填入維修單模板，輸出 xlsx
與 generator.py 相同邏輯，標題與檔名改為「維修單」
"""

import re
import shutil
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TEMPLATE_PATH = Path(__file__).parent.parent / "template" / "template_fix.xlsx"
OUTPUT_DIR    = Path(__file__).parent.parent / "output"

ITEM_ROW     = 8
FOOTER_START = 9


def _format_date(s):
    try:
        dt = datetime.strptime(s, "%Y/%m/%d")
        return f"{dt.year} 年  {dt.month:02d} 月  {dt.day:02d}  日"
    except Exception:
        return s


def _safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def _copy_row_style(ws, src_row, dst_row):
    for col in range(1, 11):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        if isinstance(dst, MergedCell):
            continue
        if src.has_style:
            dst.font          = copy(src.font)
            dst.border        = copy(src.border)
            dst.fill          = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment     = copy(src.alignment)
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def _shift_footer_merges(ws, n_extra):
    to_shift = [
        (r.min_col, r.min_row, r.max_col, r.max_row)
        for r in ws.merged_cells.ranges
        if r.min_row >= FOOTER_START
    ]
    for (c1, r1, c2, r2) in to_shift:
        ws.unmerge_cells(start_row=r1, start_column=c1,
                         end_row=r2,   end_column=c2)
    for (c1, r1, c2, r2) in to_shift:
        ws.merge_cells(start_row=r1 + n_extra, start_column=c1,
                       end_row=r2   + n_extra, end_column=c2)


def _draw_invoice_section(ws, start_row, invoice_choice):
    thin    = Side(style="thin")
    box     = Border(top=thin, bottom=thin, left=thin, right=thin)
    center  = Alignment(horizontal="center", vertical="center")
    left_al = Alignment(horizontal="left",   vertical="center")
    black_fill = PatternFill(fill_type="solid", fgColor="FF000000")
    white_fill = PatternFill(fill_type="solid", fgColor="FFFFFFFF")
    text_font  = Font(size=11, color="000000")

    r1, r2 = start_row, start_row + 1
    ws.row_dimensions[r1].height = 20
    ws.row_dimensions[r2].height = 20

    def _merge(col1, col2):
        try:
            ws.merge_cells(start_row=r1, start_column=col1,
                           end_row=r2,   end_column=col2)
        except Exception:
            pass

    _merge(1, 5)
    lc = ws.cell(row=r1, column=1)
    if not isinstance(lc, MergedCell):
        lc.value, lc.alignment, lc.font = "發票開立方式：", left_al, text_font

    _merge(6, 6)
    cb1 = ws.cell(row=r1, column=6)
    if not isinstance(cb1, MergedCell):
        cb1.value     = ""
        cb1.alignment = center
        cb1.border    = box
        cb1.fill      = black_fill if invoice_choice == "隨貨" else white_fill

    _merge(7, 8)
    tc1 = ws.cell(row=r1, column=7)
    if not isinstance(tc1, MergedCell):
        tc1.value, tc1.alignment, tc1.font = "  發票隨貨", left_al, text_font

    _merge(9, 9)
    cb2 = ws.cell(row=r1, column=9)
    if not isinstance(cb2, MergedCell):
        cb2.value     = ""
        cb2.alignment = center
        cb2.border    = box
        cb2.fill      = black_fill if invoice_choice == "直寄" else white_fill

    _merge(10, 10)
    tc2 = ws.cell(row=r1, column=10)
    if not isinstance(tc2, MergedCell):
        tc2.value, tc2.alignment, tc2.font = "  發票直寄", left_al, text_font


def _clear_invoice_section(ws, start_row):
    no_fill, no_border = PatternFill(fill_type=None), Border()
    no_font, no_align  = Font(), Alignment()

    for rng in [f"F{start_row}:H{start_row+1}",
                f"I{start_row}:J{start_row+1}",
                f"A{start_row}:E{start_row+1}"]:
        try:
            ws.unmerge_cells(rng)
        except Exception:
            pass

    for r in (start_row, start_row + 1):
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell, MergedCell):
                continue
            cell.value, cell.fill   = None, no_fill
            cell.border, cell.font  = no_border, no_font
            cell.alignment          = no_align


def _norm_text(s):
    return re.sub(r'[\s\u3000]+', '', str(s or ''))


def _insert_invoice_flag_below(ws, target_text, invoice_choice='尚未確認'):
    if invoice_choice == '隨貨':
        sym1, sym2 = '■', '□'
    elif invoice_choice == '直寄':
        sym1, sym2 = '□', '■'
    else:
        sym1, sym2 = '□', '□'

    required_parts = ['第一聯', '白聯', '立善留存',
                      '第二聯', '紅聯', '貨運公司',
                      '第三聯', '黃聯', '客戶收執聯']

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            norm_value = _norm_text(cell.value)
            if (_norm_text(target_text) in norm_value
                    or all(p in norm_value for p in required_parts)):
                dst_row, dst_col = cell.row + 1, cell.column + 8
                dst = ws.cell(row=dst_row, column=dst_col)
                while isinstance(dst, MergedCell):
                    dst_row += 1
                    dst = ws.cell(row=dst_row, column=dst_col)
                dst.value     = f"{sym1}發票隨貨"
                dst.alignment = Alignment(horizontal='left', vertical='center')
                if cell.has_style:
                    dst.font = copy(cell.font)

                dst2 = ws.cell(row=dst_row, column=dst_col + 1)
                if isinstance(dst2, MergedCell):
                    dst2 = ws.cell(row=dst_row + 1, column=dst_col + 1)
                dst2.value     = f"{sym2}發票直寄"
                dst2.alignment = Alignment(horizontal='left', vertical='center')
                if cell.has_style:
                    dst2.font = copy(cell.font)
                return


def generate_fix(data, extra, out_filename=""):
    OUTPUT_DIR.mkdir(exist_ok=True)
    tmp = OUTPUT_DIR / "_working_fix.xlsx"
    shutil.copy(TEMPLATE_PATH, tmp)

    wb = openpyxl.load_workbook(tmp)
    ws = wb.active

    h       = data["header"]
    items   = data["items"]
    n       = len(items)
    n_extra = n - 1

    # ── 1. 表頭 ───────────────────────────────────────────────
    ws["A3"] = "維修日期：" + _format_date(extra.get("ship_date", ""))
    _safe_write(ws, 4, 3, h.get("customer", ""))
    _safe_write(ws, 5, 3, h.get("phone", ""))
    _safe_write(ws, 6, 3, h.get("address", ""))
    ws["F5"] = "聯絡人：" + h.get("contact", "")

    sale_no = extra.get("sale_no", "").strip()
    if sale_no:
        ws["I4"] = "銷貨單號：" + sale_no

    tax_id = h.get("tax_id", "").strip()
    if tax_id:
        ws["F4"] = "統一編號：" + tax_id

    # ── 2. 插入多餘列 ─────────────────────────────────────────
    if n_extra > 0:
        _shift_footer_merges(ws, n_extra)
        ws.insert_rows(FOOTER_START, amount=n_extra)

    # ── 3. 品項 ───────────────────────────────────────────────
    for i, item in enumerate(items):
        r = ITEM_ROW + i
        if i > 0:
            ws.merge_cells(f"A{r}:B{r}")
            ws.merge_cells(f"C{r}:E{r}")
            _copy_row_style(ws, ITEM_ROW, r)
        _safe_write(ws, r, 1, i + 1)
        _safe_write(ws, r, 3, item["name"].replace("\n", " "))
        _safe_write(ws, r, 6, item["qty"])
        _safe_write(ws, r, 7, item["unit"])
        _safe_write(ws, r, 8, item["unit_price"])
        _safe_write(ws, r, 9, item["subtotal"])

    # ── 4. Footer ─────────────────────────────────────────────
    note_row     = ITEM_ROW + n
    operator_row = note_row + 1

    note_text = extra.get("note", "").strip()
    if note_text:
        _safe_write(ws, note_row, 2, note_text)
    _safe_write(ws, operator_row, 3, extra.get("operator", ""))

    # ── 5. 發票區 ─────────────────────────────────────────────
    invoice_row = 12 + n_extra
    _clear_invoice_section(ws, invoice_row)
    _insert_invoice_flag_below(
        ws,
        "※第一聯(白聯)為立善留存；第二聯(紅聯)為貨運公司留存；第三聯(黃聯)為客戶收執聯",
        extra.get("invoice_choice", "尚未確認")
    )

    # ── 6. 存檔 ───────────────────────────────────────────────
    if not out_filename:
        customer = h.get("customer", "客戶")
        date_tag = extra.get("ship_date", "").replace("/", "") or datetime.today().strftime("%Y%m%d")
        out_filename = f"維修單-{customer}-{date_tag}.xlsx"

    out_path = OUTPUT_DIR / out_filename
    wb.save(out_path)
    tmp.unlink(missing_ok=True)
    return out_path


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from parser import parse

    if len(sys.argv) < 2:
        print("用法：python generator_fix.py 報價單.xlsx")
        sys.exit(1)

    data  = parse(sys.argv[1])
    extra = {
        "ship_date":      datetime.today().strftime("%Y/%m/%d"),
        "sale_no":        "",
        "operator":       "小皋",
        "note":           "",
        "invoice_choice": "隨貨",
    }
    out = generate_fix(data, extra)
    print(f"維修單已輸出：{out}")