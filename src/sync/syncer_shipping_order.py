"""
syncer_shipping_order.py — 從出貨一覽表 Google Sheet 讀取資料，推送至出貨指示單
"""
from pathlib import Path
from collections import defaultdict

from sync.syncer_sheets import get_service, _SPREADSHEET_ID as _SRC_ID, _SHEET_GID as _SRC_GID

_DST_ID    = "15H37eDyC2MtqreSggyj8QR32vpN3n7R7Bdb8NTOXJdU"
_DST_SHEET = "2026"

# 出貨一覽表欄位索引（0-based）
_COL_PREFIX     = 0   # A: %
_COL_ORDER_DATE = 3   # D: 下單日期
_COL_DELIVERY   = 4   # E: 預計出貨
_COL_COMPANY    = 5   # F: 客戶名稱
_COL_PRODUCT    = 8   # I: 品號（HYPERLINK 顯示文字）
_COL_QUANTITY   = 9   # J: 數量
_COL_PAYMENT    = 12  # M: 付款方式
_COL_AMOUNT     = 13  # N: 貨款金額(含稅)

# 區域關鍵字對照
_REGION_MAP = {
    "北區": ["台北", "新北", "基隆", "桃園", "新竹"],
    "中區": ["苗栗", "台中", "彰化", "雲林", "南投"],
    "南區": ["嘉義", "台南", "高雄", "屏東"],
}
_REGION_ORDER = ["北區", "中區", "南區"]


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
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    rows = result.get("values", [])
    records = []
    for row in rows:
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
            "product":      row[_COL_PRODUCT]    if len(row) > _COL_PRODUCT    else "",
            "quantity":     row[_COL_QUANTITY]   if len(row) > _COL_QUANTITY   else "",
            "payment_type": row[_COL_PAYMENT]    if len(row) > _COL_PAYMENT    else "",
            "amount":       row[_COL_AMOUNT]     if len(row) > _COL_AMOUNT     else "",
        })
    return records


def _classify_region(location: str) -> str:
    """根據地區字串判斷屬於北區/中區/南區，無法判斷回傳空字串。"""
    for region, keywords in _REGION_MAP.items():
        if any(kw in location for kw in keywords):
            return region
    return ""


def _build_location_map(service) -> dict[str, str]:
    """建立 客戶名稱 → 地區 對照表：先載入 Trello 快取，再以出貨指示單現有資料覆蓋。"""
    import json
    from _paths import _LOCATION_CACHE_PATH

    mapping: dict[str, str] = {}

    # 1. Trello 快取（fallback）
    if _LOCATION_CACHE_PATH.exists():
        try:
            mapping.update(json.loads(_LOCATION_CACHE_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass

    # 2. 出貨指示單現有資料（優先，覆蓋快取）
    result = service.spreadsheets().values().get(
        spreadsheetId=_DST_ID,
        range=f"'{_DST_SHEET}'!C:D",
    ).execute()
    for row in result.get("values", []):
        if len(row) >= 2:
            company  = str(row[0]).strip()
            location = str(row[1]).strip()
            if company and location:
                mapping[company] = location

    return mapping


def _row_to_shipping(record: dict, loc_map: dict) -> list:
    """對應出貨指示單欄位：A業務|B出貨日期|C客戶名稱|D地區|E機台|F備貨|G發票|H付款|I出貨單|J序號|K換輪"""
    company = record.get("company", "")
    product = record.get("product", "")
    qty     = str(record.get("quantity", "")).strip()
    machine = f"{product}*{qty}台" if qty else product

    return [
        record.get("prefix",   ""),      # A 業務
        record.get("delivery", ""),      # B 出貨日期
        company,                         # C 客戶名稱
        loc_map.get(company, ""),        # D 地區
        machine,                         # E 機台
        "",                              # F 備貨狀態
        "",                              # G 發票
        "",                              # H 付款
        "",                              # I 出貨單
        "",                              # J 序號貼紙
        "",                              # K 換輪
    ]


def _insert_rows_at(service, sheet_id: int, insert_index: int, count: int):
    """在 insert_index（0-based）前插入 count 列，繼承上方格式。"""
    service.spreadsheets().batchUpdate(
        spreadsheetId=_DST_ID,
        body={"requests": [{
            "insertDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": insert_index,
                    "endIndex":   insert_index + count,
                },
                "inheritFromBefore": True,
            }
        }]},
    ).execute()


def _center_column_b(service, sheet_id: int, start_index: int, count: int):
    """將指定列範圍的 B 欄設定水平置中。"""
    service.spreadsheets().batchUpdate(
        spreadsheetId=_DST_ID,
        body={"requests": [{
            "repeatCell": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    start_index,
                    "endRowIndex":      start_index + count,
                    "startColumnIndex": 1,
                    "endColumnIndex":   2,
                },
                "cell":   {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment",
            }
        }]},
    ).execute()


def push_shipping_orders(records: list[dict],
                         credentials_path: Path,
                         token_path: Path) -> int:
    """將選取的資料列推送至出貨指示單，依地區插入對應區域末尾。"""
    if not records:
        return 0
    service = get_service(credentials_path, token_path)

    # 取得 sheetId
    meta = service.spreadsheets().get(spreadsheetId=_DST_ID).execute()
    sheet_id = next(
        (s["properties"]["sheetId"] for s in meta["sheets"]
         if s["properties"]["title"] == _DST_SHEET),
        None,
    )

    # 建立地區對照表
    loc_map = _build_location_map(service)

    # 讀取現有 A 欄，找出各區域標籤的列位置（0-based）
    all_a = service.spreadsheets().values().get(
        spreadsheetId=_DST_ID,
        range=f"'{_DST_SHEET}'!A:A",
    ).execute().get("values", [])

    section_label_row: dict[str, int] = {}
    for i, row in enumerate(all_a):
        cell = str(row[0]).strip() if row else ""
        if cell in _REGION_ORDER:
            section_label_row[cell] = i

    # 找每個區域最後一個有資料的列（0-based），作為插入點
    def _last_data_row(section: str) -> int:
        """回傳該區域最後一筆資料的 0-based row index（即新列插入在此之後）。"""
        start = section_label_row[section]
        # 下一個區域的位置，或 sheet 末尾
        next_starts = [
            section_label_row[s] for s in _REGION_ORDER
            if s in section_label_row and section_label_row[s] > start
        ]
        end = min(next_starts) if next_starts else len(all_a)

        last = start  # 最少是標籤列本身
        for k in range(start + 1, end):
            row = all_a[k] if k < len(all_a) else []
            if row and str(row[0]).strip():
                last = k
        return last

    # 按區域分組
    grouped: dict[str, list[dict]] = defaultdict(list)
    unmatched: list[dict] = []
    for r in records:
        location = loc_map.get(r.get("company", ""), "")
        region   = _classify_region(location)
        if region and region in section_label_row:
            grouped[region].append(r)
        else:
            unmatched.append(r)

    # 由下往上插入，避免列位移影響上方區域
    total = 0
    for region in reversed(_REGION_ORDER):
        batch = grouped.get(region, [])
        if not batch:
            continue

        insert_at = _last_data_row(region) + 1  # 0-based 插入位置
        n = len(batch)

        _insert_rows_at(service, sheet_id, insert_at, n)

        rows_data = [_row_to_shipping(r, loc_map) for r in batch]
        service.spreadsheets().values().update(
            spreadsheetId=_DST_ID,
            range=f"'{_DST_SHEET}'!A{insert_at + 1}",
            valueInputOption="USER_ENTERED",
            body={"values": rows_data},
        ).execute()

        if sheet_id is not None:
            _center_column_b(service, sheet_id, insert_at, n)

        total += n

    # 無法判斷區域的，附加至 sheet 末尾
    if unmatched:
        last_row  = len(all_a)
        start_row = last_row + 1
        rows_data = [_row_to_shipping(r, loc_map) for r in unmatched]
        service.spreadsheets().values().update(
            spreadsheetId=_DST_ID,
            range=f"'{_DST_SHEET}'!A{start_row}",
            valueInputOption="USER_ENTERED",
            body={"values": rows_data},
        ).execute()
        if sheet_id is not None:
            _center_column_b(service, sheet_id, last_row, len(unmatched))
        total += len(unmatched)

    return total
