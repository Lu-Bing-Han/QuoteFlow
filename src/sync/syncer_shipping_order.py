"""
syncer_shipping_order.py — 從出貨一覽表 Google Sheet 讀取資料，推送至出貨指示單
"""
from pathlib import Path

from sync.syncer_sheets import get_service, _SPREADSHEET_ID as _SRC_ID, _SHEET_GID as _SRC_GID

_DST_ID       = "15H37eDyC2MtqreSggyj8QR32vpN3n7R7Bdb8NTOXJdU"
_DST_SHEET    = "2026"

# 出貨一覽表欄位索引（0-based）
_COL_PREFIX     = 0   # A: %
_COL_ORDER_DATE = 3   # D: 下單日期
_COL_DELIVERY   = 4   # E: 預計出貨
_COL_COMPANY    = 5   # F: 客戶名稱
_COL_PRODUCT    = 8   # I: 品號（HYPERLINK 顯示文字）
_COL_PAYMENT    = 12  # M: 付款方式

 
def _get_sheet_name(service, spreadsheet_id: str, gid: int) -> str:
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["sheetId"] == gid:
            return sheet["properties"]["title"]
    raise ValueError(f"找不到 gid={gid} 的工作表")


def fetch_from_overview(credentials_path: Path, token_path: Path) -> list[dict]:
    """從出貨一覽表 Google Sheet 讀取所有資料列，回傳 list[dict]。"""
    service    = get_service(credentials_path, token_path)
    sheet_name = _get_sheet_name(service, _SRC_ID, _SRC_GID)

    result = service.spreadsheets().values().get(
        spreadsheetId=_SRC_ID,
        range=f"'{sheet_name}'!A:N",
        valueRenderOption="FORMATTED_VALUE",   # HYPERLINK 回傳顯示文字
    ).execute()

    rows = result.get("values", [])
    records = []
    for row in rows:
        # 跳過空列與標題列（標題列 A 欄通常為 "%" 或空）
        if not row:
            continue
        prefix = row[_COL_PREFIX] if len(row) > _COL_PREFIX else ""
        if not prefix or prefix in ("%", "0"):
            continue

        records.append({
            "prefix":       prefix,
            "order_date":   row[_COL_ORDER_DATE] if len(row) > _COL_ORDER_DATE else "",
            "delivery":     row[_COL_DELIVERY]   if len(row) > _COL_DELIVERY   else "",
            "company":      row[_COL_COMPANY]    if len(row) > _COL_COMPANY    else "",
            "location":     "",   # 出貨一覽表無此欄
            "product":      row[_COL_PRODUCT]    if len(row) > _COL_PRODUCT    else "",
            "payment_type": row[_COL_PAYMENT]    if len(row) > _COL_PAYMENT    else "",
        })
    return records



def _row_to_shipping(record: dict) -> list:
    """對應出貨指示單欄位：A業務|B出貨日期|C客戶名稱|D地區|E機台|F備貨|G發票|H付款|I出貨單|J序號|K換輪"""
    return [
        record.get("prefix",       ""),   # A 業務
        record.get("delivery",     ""),   # B 出貨日期
        record.get("company",      ""),   # C 客戶名稱
        record.get("location",     ""),   # D 地區（空白）
        record.get("product",      ""),   # E 機台
        "",                               # F 備貨狀態
        "",                               # G 發票
        record.get("payment_type", ""),   # H 付款
        "",                               # I 出貨單
        "",                               # J 序號貼紙
        "",                               # K 換輪
    ]


def push_shipping_orders(records: list[dict],
                         credentials_path: Path,
                         token_path: Path) -> int:
    """將選取的資料列推送至出貨指示單 Google Sheet（工作表 2026）。"""
    if not records:
        return 0
    service = get_service(credentials_path, token_path)
    rows = [_row_to_shipping(r) for r in records]
    service.spreadsheets().values().append(
        spreadsheetId=_DST_ID,
        range=f"'{_DST_SHEET}'!A:K",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()
    return len(rows)
