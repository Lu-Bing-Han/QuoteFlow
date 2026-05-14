"""
syncer_trello.py — 從 Trello「本周下單(PO)」清單抓取卡片並解析欄位
"""
import re
from datetime import datetime, timezone

import requests

_API_BASE   = "https://api.trello.com/1"
_BOARD_NAME = "物流事業部1"
_LIST_NAME  = "本周下單"       # 包含此字串即符合


def _auth(api_key: str, token: str) -> dict:
    return {"key": api_key, "token": token}


def _get_board_id(api_key: str, token: str) -> str:
    resp = requests.get(
        f"{_API_BASE}/members/me/boards",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    for b in resp.json():
        if b["name"] == _BOARD_NAME:
            return b["id"]
    raise ValueError(f"找不到看板「{_BOARD_NAME}」")


def _get_list_id(api_key: str, token: str, board_id: str) -> str:
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


def _parse_title(title: str) -> dict:
    """從卡片標題萃取各欄位。"""
    # 移除 【...】 標記
    t = re.sub(r'【[^】]*】', '', title).strip()

    # 前綴（WI / CH / JE …）
    prefix = ""
    m = re.match(r'^([A-Za-z]{1,6})\s*[-\s]', t)
    if m:
        prefix = m.group(1).upper()
        t = t[m.end():].strip()

    # 客戶名稱（括號前的文字）
    company = ""
    m = re.match(r'^(.+?)[\(（]', t)
    if m:
        company = m.group(1).strip()
        t = re.sub(r'^.+?[\)）]\s*', '', t).strip()   # 移除地點括號
    else:
        parts = t.split()
        company = parts[0] if parts else ""
        t = ' '.join(parts[1:])

    # 聯絡人（先生/小姐/女士 結尾）
    m = re.match(r'^.+?(?:先生|小姐|女士)\s*[-\s]*', t)
    if m:
        t = t[m.end():].strip()

    t = re.sub(r'^[-\s]+', '', t).strip()

    # 數量（*N 格式）
    quantity = "1"
    m = re.search(r'\*(\d+)\s*$', t)
    if m:
        quantity = m.group(1)
        t = t[:m.start()].strip()

    product = re.sub(r'【[^】]*】', '', t).strip()

    return {"prefix": prefix, "company": company, "product": product, "quantity": quantity}


def fetch_po_cards(api_key: str, token: str) -> list[dict]:
    """回傳本周下單清單的所有卡片（已解析欄位）。"""
    board_id = _get_board_id(api_key, token)
    list_id  = _get_list_id(api_key, token, board_id)

    resp = requests.get(
        f"{_API_BASE}/lists/{list_id}/cards",
        params={**_auth(api_key, token), "fields": "id,name,due"},
        timeout=15,
    )
    resp.raise_for_status()

    result = []
    for card in resp.json():
        fields = _parse_title(card["name"])
        # 建立日期從 card ID 前 8 位推算
        ts  = int(card["id"][:8], 16)
        dt  = datetime.fromtimestamp(ts, tz=timezone.utc)
        created_date = f"{dt.month}/{dt.day}"

        result.append({
            "card_id":      card["id"],
            "title":        card["name"],
            "prefix":       fields["prefix"],
            "company":      fields["company"],
            "product":      fields["product"],
            "quantity":     fields["quantity"],
            "created_date": created_date,
            "due_date":     card.get("due") or "",
        })
    return result
