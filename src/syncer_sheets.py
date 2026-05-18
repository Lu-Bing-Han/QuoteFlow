"""
syncer_sheets.py — 透過 OAuth 將 Trello 卡片同步到 Google Sheets
"""
import json
from pathlib import Path

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

_SCOPES          = ["https://www.googleapis.com/auth/spreadsheets"]
_SPREADSHEET_ID  = "1lmQR4CG7dqOqiXIIRF_ztAu2mjw1XyGPorZM_7La6Dg"
_SHEET_GID       = 584074203


def get_service(credentials_path: Path, token_path: Path):
    """回傳已授權的 Sheets API service；首次執行會開瀏覽器授權。"""
    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), _SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow  = InstalledAppFlow.from_client_secrets_file(str(credentials_path), _SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("sheets", "v4", credentials=creds)


def _get_sheet_name(service, spreadsheet_id: str, gid: int) -> str:
    """從 gid 查出工作表名稱。"""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["sheetId"] == gid:
            return sheet["properties"]["title"]
    raise ValueError(f"找不到 gid={gid} 的工作表")


def _read_synced_ids(synced_path: Path) -> set:
    if synced_path.exists():
        return set(json.loads(synced_path.read_text(encoding="utf-8")))
    return set()


def _save_synced_ids(synced_path: Path, ids: set):
    synced_path.write_text(json.dumps(sorted(ids), ensure_ascii=False, indent=2),
                           encoding="utf-8")


def _card_to_row(card: dict) -> list:
    """對應 Sheet 欄位順序：%, Y, 付託Y/N, 下單日期, 預計出貨, 客戶名稱, 性質, OB, 品號, 數量, 改造Y/N, 改造項目, 付款方式, 貨款金額"""
    product_safe = card["product"].replace('"', "")
    return [
        card["prefix"],                                                              # %
        "Y",                                                                         # Y
        card.get("payment_raw", ""),                                                 # 付託Y/N
        card["created_date"],                                                        # 下單日期
        card.get("delivery", ""),                                                    # 預計出貨
        card["company"],                                                             # 客戶名稱
        "BUYER",                                                                     # 性質
        "N",                                                                         # OB
        f'=HYPERLINK("{card.get("card_url","")}", "{product_safe}")',                # 品號
        card["quantity"],                                                            # 數量
        card.get("has_remodel", ""),                                                 # 改造Y/N
        "",                                                                          # 改造項目
        card.get("payment_type", ""),                                                # 付款方式
        card.get("amount", ""),                                                      # 貨款金額(含稅)
    ]


def sync_cards(cards: list[dict], credentials_path: Path,
               token_path: Path, synced_path: Path) -> int:
    """
    將 cards 中尚未同步的卡片寫入 Google Sheet。
    回傳新增筆數。
    """
    service     = get_service(credentials_path, token_path)
    sheet_name  = _get_sheet_name(service, _SPREADSHEET_ID, _SHEET_GID)
    synced_ids  = _read_synced_ids(synced_path)

    new_cards = [c for c in cards if c["card_id"] not in synced_ids]
    if not new_cards:
        return 0

    rows = [_card_to_row(c) for c in new_cards]
    service.spreadsheets().values().append(
        spreadsheetId=_SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:N",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()

    synced_ids.update(c["card_id"] for c in new_cards)
    _save_synced_ids(synced_path, synced_ids)
    return len(new_cards)
