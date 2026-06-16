"""
syncer_trello.py — 從 Trello「本周下單(PO)」清單抓取卡片並解析欄位
"""
import json
import re
from datetime import datetime, timezone, date as date_type
from pathlib import Path

import requests

_API_BASE   = "https://api.trello.com/1"
_BOARD_NAME = "物流事業部1"
_LIST_NAME  = "本周下單"       # 包含此字串即符合


def _auth(api_key: str, token: str) -> dict:
    return {"key": api_key, "token": token}


def _get_board_id(api_key: str, token: str, board_name: str) -> str:
    resp = requests.get(
        f"{_API_BASE}/members/me/boards",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    for b in resp.json():
        if b["name"] == board_name:
            return b["id"]
    raise ValueError(f"找不到看板「{board_name}」")


def _get_list_id(api_key: str, token: str, board_id: str, list_name: str) -> str:
    resp = requests.get(
        f"{_API_BASE}/boards/{board_id}/lists",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    for lst in resp.json():
        if list_name in lst["name"]:
            return lst["id"]
    raise ValueError(f"找不到清單「{list_name}」")


def _parse_desc(desc: str) -> dict:
    """從卡片描述萃取付款方式、交期、應收總金額、下單日期。"""
    def _find(pattern):
        m = re.search(pattern, desc or "", re.MULTILINE)
        return m.group(1).strip() if m else ""

    payment_raw    = _find(r'付款方式[：:]\s*(.+)')
    delivery       = _find(r'交期[：:]\s*(.+)')
    amount         = _find(r'應收總金額[：:]\s*(.+)')
    order_date_raw = _find(r'下單日期[\s　]*[：:﹕][\s　]*(.+)')

    if "現金" in payment_raw:
        payment_type = "現金"
    elif "匯款" in payment_raw:
        payment_type = "匯款"
    elif "支票" in payment_raw:
        payment_type = "支票"
    else:
        payment_type = ""

    order_date_str = ""
    order_date_dt  = None
    if order_date_raw:
        raw = order_date_raw.strip()
        # 三段式：114/6/8、2025/6/8、2025-06-08
        m3 = re.match(r'(\d{2,4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})', raw)
        # 兩段式：6/8、06/08
        m2 = re.match(r'(\d{1,2})[/\-\.](\d{1,2})', raw)
        if m3:
            yr, mo, dy = int(m3.group(1)), int(m3.group(2)), int(m3.group(3))
            if yr < 200:
                yr += 1911
            order_date_str = f"{mo}/{dy}"
            try:
                order_date_dt = date_type(yr, mo, dy)
            except ValueError:
                pass
        elif m2:
            mo, dy = int(m2.group(1)), int(m2.group(2))
            yr = date_type.today().year
            order_date_str = f"{mo}/{dy}"
            try:
                order_date_dt = date_type(yr, mo, dy)
            except ValueError:
                pass

    return {
        "payment_raw":    payment_raw,
        "payment_type":   payment_type,
        "delivery":       delivery,
        "amount":         amount,
        "order_date_str": order_date_str,
        "order_date_dt":  order_date_dt,
    }


def _parse_bracket_date(title: str) -> tuple[str, date_type | None]:
    """從標題 【YYY.MM.DD ...】 解析民國日期，回傳 ('M/D', date) 或 ('', None)。"""
    m = re.search(r'【(\d{2,3})\.(\d{1,2})\.(\d{1,2})', title)
    if m:
        roc, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{month}/{day}", date_type(roc + 1911, month, day)
    return "", None


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

    # 客戶名稱（括號前的文字）與地區（括號內）
    company = ""
    location = ""
    m = re.match(r'^(.+?)[\(（](.+?)[\)）]', t)
    if m:
        company  = m.group(1).strip()
        location = m.group(2).strip()
        t = t[m.end():].strip()
    else:
        m2 = re.match(r'^(.+?)[\(（]', t)
        if m2:
            company = m2.group(1).strip()
            t = re.sub(r'^.+?[\)）]\s*', '', t).strip()
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

    return {"prefix": prefix, "company": company, "location": location, "product": product, "quantity": quantity}


def fetch_po_cards(api_key: str, token: str,
                   board_name: str = _BOARD_NAME,
                   list_name:  str = _LIST_NAME) -> list[dict]:
    """回傳指定看板/清單的所有卡片（已解析欄位）。"""
    board_id = _get_board_id(api_key, token, board_name)
    list_id  = _get_list_id(api_key, token, board_id, list_name)

    resp = requests.get(
        f"{_API_BASE}/lists/{list_id}/cards",
        params={**_auth(api_key, token), "fields": "id,name,due,shortUrl,desc,labels"},
        timeout=15,
    )
    resp.raise_for_status()

    result = []
    for card in resp.json():
        fields = _parse_title(card["name"])
        desc_data = _parse_desc(card.get("desc", ""))

        # 優先使用描述裡的下單日期，否則 fallback 到標題括號日期，最後用卡片建立時間
        if desc_data["order_date_str"]:
            order_date = desc_data["order_date_str"]
            order_dt   = desc_data["order_date_dt"]
        else:
            order_date, order_dt = _parse_bracket_date(card["name"])
            if not order_date:
                ts = int(card["id"][:8], 16)
                dt = datetime.fromtimestamp(ts, tz=timezone.utc)
                order_date = f"{dt.month}/{dt.day}"
                order_dt   = dt.date()

        label_names = [lb.get("name", "") for lb in (card.get("labels") or [])]
        has_remodel = "Y" if any("改造" in n for n in label_names) else "N"

        result.append({
            "card_id":      card["id"],
            "title":        card["name"],
            "prefix":       fields["prefix"],
            "company":      fields["company"],
            "location":     fields["location"],
            "product":      fields["product"],
            "quantity":     fields["quantity"],
            "created_date": order_date,
            "created_dt":   order_dt,
            "due_date":     card.get("due") or "",
            "card_url":     card.get("shortUrl", ""),
            "payment_raw":  desc_data["payment_raw"],
            "payment_type": desc_data["payment_type"],
            "delivery":     desc_data["delivery"],
            "amount":       desc_data["amount"],
            "has_remodel":  has_remodel,
        })
    return result


def update_location_cache(cards: list[dict], cache_path: Path) -> None:
    """從 Trello 卡片的 location 欄位更新客戶地區快取（只補新資料，不覆蓋已有的）。"""
    existing: dict[str, str] = {}
    if cache_path.exists():
        try:
            existing = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    for card in cards:
        company  = card.get("company",  "").strip()
        location = card.get("location", "").strip()
        if company and location and company not in existing:
            existing[company] = location
    cache_path.write_text(
        json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8"
    )
