"""
creator_trello.py — 從 Excel 讀取資料並在 Trello 建立卡片
"""
from pathlib import Path

import openpyxl
import requests

_API_BASE   = "https://api.trello.com/1"
_BOARD_NAME = "物流事業部1"
_LIST_NAME  = "0.待評估"


def _auth(api_key: str, token: str) -> dict:
    return {"key": api_key, "token": token}


def _get_target_list_id(api_key: str, token: str) -> str:
    resp = requests.get(
        f"{_API_BASE}/members/me/boards",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    board_id = None
    for b in resp.json():
        if b["name"] == _BOARD_NAME:
            board_id = b["id"]
            break
    if not board_id:
        raise ValueError(f"找不到看板「{_BOARD_NAME}」")

    resp = requests.get(
        f"{_API_BASE}/boards/{board_id}/lists",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    for lst in resp.json():
        if _LIST_NAME in lst["name"]:
            return lst["id"]
    raise ValueError(f"找不到清單「{_LIST_NAME}」")


def read_excel_cards(excel_path: Path) -> list[dict]:
    """讀 Excel：C欄第一行為標題，其餘為描述；D欄需求附加在描述下方。

    回傳 list of {"row": int, "title": str, "desc": str, "needs": str}
    """
    wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    ws = wb.active

    cards = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        c_val = row[2].value  # C欄
        if not c_val:
            continue

        lines = str(c_val).strip().splitlines()
        title = lines[0].strip() if lines else ""
        rest  = "\n".join(ln for ln in lines[1:] if ln.strip())

        d_val  = row[3].value  # D欄（需求）
        needs  = str(d_val).strip() if d_val else ""

        if not title:
            continue

        desc_parts = []
        if rest:
            desc_parts.append(rest)
        if needs:
            desc_parts.append(f"需求：\n{needs}")

        cards.append({
            "row":   row_idx,
            "title": title,
            "desc":  "\n\n".join(desc_parts),
            "needs": needs,
        })

    wb.close()
    return cards


def create_cards(cards: list[dict], api_key: str, token: str) -> int:
    """建立 Trello 卡片，回傳成功建立筆數。"""
    list_id = _get_target_list_id(api_key, token)
    created = 0
    for card in cards:
        resp = requests.post(
            f"{_API_BASE}/cards",
            params={
                **_auth(api_key, token),
                "idList": list_id,
                "name":   card["title"],
                "desc":   card["desc"],
            },
            timeout=15,
        )
        resp.raise_for_status()
        created += 1
    return created
