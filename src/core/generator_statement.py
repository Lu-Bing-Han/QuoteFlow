"""
generator_statement.py — 依 Trello「本周下單」卡片填入對帳單模板，輸出 xlsx
"""
import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.cell import MergedCell

from _paths import TEMPLATE_DIR, OUTPUT_DIR
TEMPLATE_PATH = TEMPLATE_DIR / "template_statement.xlsx"

TITLE_ROW  = 6
HEADER_ROW = 11
ITEM_ROW   = 12   # 模板保留 12~14 三列供品項使用（總金額為 SUM 公式）


def _safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def _append_to_label(ws, row, col, value):
    """保留模板原有的標籤文字（含其間距），於後方補上資料。"""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return
    label = cell.value or ""
    cell.value = f"{label}{value}"


def parse_amount(raw) -> float:
    """從「應收總金額」之類的原始字串解析數字金額，無法解析則回傳 0。"""
    s = re.sub(r'[^\d.]', '', str(raw or ''))
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def parse_paid_amount(payment_raw: str) -> float:
    """從付款方式原始文字（如「匯款，已付訂金5000」）抓出已收金額，無法判斷則回傳 0。"""
    m = re.search(r'(?:已付|已收|訂金|預收)[^\d]*([\d,]+(?:\.\d+)?)', payment_raw or '')
    return parse_amount(m.group(1)) if m else 0.0


def generate_statement(card: dict, location_lookup: dict | None = None,
                        output_dir=None) -> Path:
    """依單張 Trello 卡片資料產出對帳單，回傳輸出檔案路徑。"""
    out_dir = Path(output_dir) if output_dir else OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    title_company = card.get("company", "")
    company       = card.get("company_desc", "") or title_company
    address = (
        card.get("address", "")
        or (location_lookup or {}).get(company, "")
        or (location_lookup or {}).get(title_company, "")
        or card.get("location", "")
    )

    total_amount = parse_amount(card.get("amount", ""))
    paid_amount  = parse_paid_amount(card.get("payment_raw", ""))
    due_amount   = total_amount - paid_amount

    quantity   = card.get("quantity", "") or ""
    unit_price = total_amount / float(quantity) if str(quantity).replace('.', '', 1).isdigit() and float(quantity) else ""

    today = datetime.today()
    title_cell = ws.cell(row=TITLE_ROW, column=1)
    if not isinstance(title_cell, MergedCell) and title_cell.value:
        title_cell.value = f"{title_cell.value.rstrip()}  {today.year}.{today.month:02d}"

    _append_to_label(ws, HEADER_ROW - 3, 1, company)
    _append_to_label(ws, HEADER_ROW - 2, 1, address)
    _append_to_label(ws, HEADER_ROW - 1, 1, card.get("payment_raw", ""))

    _append_to_label(ws, HEADER_ROW - 3, 6, card.get("contact", ""))
    _append_to_label(ws, HEADER_ROW - 2, 6, card.get("phone", ""))
    _append_to_label(ws, HEADER_ROW - 1, 6, card.get("fax", ""))

    _safe_write(ws, ITEM_ROW, 2, card.get("product", ""))
    _safe_write(ws, ITEM_ROW, 3, quantity)
    _safe_write(ws, ITEM_ROW, 4, unit_price)
    _safe_write(ws, ITEM_ROW, 5, total_amount)
    _safe_write(ws, ITEM_ROW, 6, paid_amount)
    _safe_write(ws, ITEM_ROW, 7, due_amount)
    _safe_write(ws, ITEM_ROW, 8, card.get("card_url", ""))

    date_tag = datetime.today().strftime("%Y%m%d")
    safe_company = re.sub(r'[\\/:*?"<>|]', '_', company) or "客戶"
    out_path = out_dir / f"對帳單-{safe_company}-{date_tag}.xlsx"
    wb.save(out_path)
    return out_path
