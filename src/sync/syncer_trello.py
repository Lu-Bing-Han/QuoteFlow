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


def _parse_order_date(raw: str) -> tuple[str, date_type | None]:
    """解析下單日期字串，回傳 ('M/D', date) 或 ('', None)。"""
    raw = re.sub(r'\\(?=[/\-.])', '', (raw or "").strip())
    # 三段式：114/6/8、2025/6/8、2025-06-08、115.06.18
    m3 = re.match(r'(\d{2,4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})', raw)
    # 兩段式：6/8、06/08
    m2 = re.match(r'(\d{1,2})[/\-\.](\d{1,2})', raw)
    if m3:
        yr, mo, dy = int(m3.group(1)), int(m3.group(2)), int(m3.group(3))
        if yr < 200:
            yr += 1911
        try:
            return f"{mo}/{dy}", date_type(yr, mo, dy)
        except ValueError:
            return "", None
    if m2:
        mo, dy = int(m2.group(1)), int(m2.group(2))
        yr = date_type.today().year
        try:
            return f"{mo}/{dy}", date_type(yr, mo, dy)
        except ValueError:
            return "", None
    return "", None


def _parse_comment_order_date(text: str) -> tuple[str, date_type | None]:
    """從留言中的「115.06.18已下單」格式解析下單日期。"""
    if "已下單" not in (text or ""):
        return "", None
    m = re.search(r'(\d{2,4}\\?[/\-\.]\d{1,2}\\?[/\-\.]\d{1,2})\s*已下單', text)
    if not m:
        return "", None
    return _parse_order_date(m.group(1))


def _parse_compact_roc_date(text: str) -> tuple[str, date_type | None]:
    """從附件檔名中的 1150618 格式解析民國日期。"""
    for m in re.finditer(r'(?<!\d)(\d{3})(\d{2})(\d{2})(?!\d)', text or ""):
        roc, mo, dy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return f"{mo}/{dy}", date_type(roc + 1911, mo, dy)
        except ValueError:
            continue
    return "", None


def _parse_attachment_order_date(attachments: list[dict]) -> tuple[str, date_type | None]:
    """從附件名稱解析訂購單檔名中的日期，例如 ...-1150618.jpg。"""
    texts = [
        f"{att.get('name') or ''} {att.get('url') or ''}"
        for att in (attachments or [])
    ]
    texts.sort(key=lambda text: 0 if "訂購單" in text else 1)
    for text in texts:
        m = re.search(r'(\d{2,4}\\?[/\-\.]\d{1,2}\\?[/\-\.]\d{1,2})', text)
        if m:
            order_date, order_dt = _parse_order_date(m.group(1))
            if order_date:
                return order_date, order_dt
        order_date, order_dt = _parse_compact_roc_date(text)
        if order_date:
            return order_date, order_dt
    return "", None


def _parse_actions_order_date(actions: list[dict]) -> tuple[str, date_type | None]:
    """從已載入的 Trello actions 解析留言中的已下單日期。"""
    for action in actions or []:
        data = action.get("data") or {}
        text = (data.get("text") or "").strip()
        order_date, order_dt = _parse_comment_order_date(text)
        if order_date:
            return order_date, order_dt
    return "", None


def _parse_action_date(action: dict) -> tuple[str, date_type | None]:
    """解析 Trello action date，回傳 ('M/D', date) 或 ('', None)。"""
    raw = (action.get("date") or "").strip()
    if not raw:
        return "", None
    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError:
        return "", None
    return f"{dt.month}/{dt.day}", dt.date()


def _parse_move_to_list_date(
    actions: list[dict],
    list_id: str,
) -> tuple[str, date_type | None]:
    """從 Trello actions 解析卡片移動到指定清單的日期。"""
    for action in actions or []:
        data = action.get("data") or {}
        list_after = data.get("listAfter") or {}
        if action.get("type") == "updateCard" and list_after.get("id") == list_id:
            return _parse_action_date(action)
    return "", None


def _get_comment_order_date(
    card_id: str,
    api_key: str,
    token: str,
) -> tuple[str, date_type | None]:
    """讀取 Trello 留言，抓出第一個「已下單」日期；失敗時回傳空值走 fallback。"""
    try:
        resp = requests.get(
            f"{_API_BASE}/cards/{card_id}/actions",
            params={
                **_auth(api_key, token),
                "filter": "commentCard",
                "fields": "data,date",
                "limit": 50,
            },
            timeout=15,
        )
        resp.raise_for_status()
    except requests.RequestException:
        return "", None

    return _parse_actions_order_date(resp.json())


def _get_move_to_list_date(
    card_id: str,
    list_id: str,
    api_key: str,
    token: str,
) -> tuple[str, date_type | None]:
    """讀取 Trello 移動紀錄，抓出卡片移入目前清單的日期。"""
    try:
        resp = requests.get(
            f"{_API_BASE}/cards/{card_id}/actions",
            params={
                **_auth(api_key, token),
                "filter": "updateCard:idList",
                "fields": "data,date,type",
                "limit": 100,
            },
            timeout=15,
        )
        resp.raise_for_status()
    except requests.RequestException:
        return "", None

    return _parse_move_to_list_date(resp.json(), list_id)


def _resolve_order_date(
    card: dict,
    desc_data: dict,
    list_id: str,
    api_key: str,
    token: str,
) -> tuple[str, date_type | None]:
    """依序從描述、留言、附件、標題與移入清單日期決定下單日期。"""
    if desc_data["order_date_str"]:
        return desc_data["order_date_str"], desc_data["order_date_dt"]

    order_date, order_dt = _parse_actions_order_date(card.get("actions") or [])
    if order_date:
        return order_date, order_dt

    order_date, order_dt = _get_comment_order_date(card["id"], api_key, token)
    if order_date:
        return order_date, order_dt

    order_date, order_dt = _parse_attachment_order_date(card.get("attachments") or [])
    if order_date:
        return order_date, order_dt

    order_date, order_dt = _parse_bracket_date(card["name"])
    if order_date:
        return order_date, order_dt

    order_date, order_dt = _parse_move_to_list_date(card.get("actions") or [], list_id)
    if order_date:
        return order_date, order_dt

    order_date, order_dt = _get_move_to_list_date(card["id"], list_id, api_key, token)
    if order_date:
        return order_date, order_dt

    ts = int(card["id"][:8], 16)
    dt = datetime.fromtimestamp(ts, tz=timezone.utc)
    return f"{dt.month}/{dt.day}", dt.date()


def _parse_desc(desc: str) -> dict:
    """從卡片描述萃取聯絡人/電話/傳真/地址、付款方式、交期、應收總金額、下單日期。"""
    def _find(pattern):
        m = re.search(pattern, desc or "", re.MULTILINE)
        return m.group(1).strip() if m else ""

    company_desc   = _find(r'公司名稱[ \t]*[：:][ \t]*(.+)')
    contact        = _find(r'聯絡人[ \t]*[：:][ \t]*(.+)')
    phone          = _find(r'電話[ \t]*[：:][ \t]*(.+)')
    fax            = _find(r'傳真[ \t]*[：:][ \t]*(.+)')
    address        = _find(r'地址[ \t]*[：:][ \t]*(.+)')
    payment_raw    = _find(r'付款方式[ \t]*[：:][ \t]*(.+)')
    delivery       = _find(r'交期[ \t]*[：:][ \t]*(.+)')
    amount         = _find(r'應收總金額[ \t]*[：:][ \t]*(.+)')
    order_date_raw = _find(r'下單日期[\s　]*[：:﹕][\s　]*(.+)')

    if "現金" in payment_raw:
        payment_type = "現金"
    elif "匯款" in payment_raw:
        payment_type = "匯款"
    elif "支票" in payment_raw:
        payment_type = "支票"
    else:
        payment_type = ""

    order_date_str, order_date_dt = _parse_order_date(order_date_raw)

    return {
        "company_desc":   company_desc,
        "contact":        contact,
        "phone":          phone,
        "fax":            fax,
        "address":        address,
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
        params={
            **_auth(api_key, token),
            "fields": "id,name,due,shortUrl,desc,labels,attachments",
            "attachments": "true",
            "actions": "commentCard,updateCard:idList",
            "action_fields": "data,date,type",
            "action_limit": 100,
        },
        timeout=15,
    )
    resp.raise_for_status()

    result = []
    for card in resp.json():
        fields = _parse_title(card["name"])
        desc_data = _parse_desc(card.get("desc", ""))

        order_date, order_dt = _resolve_order_date(card, desc_data, list_id, api_key, token)

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
            "company_desc": desc_data["company_desc"],
            "contact":      desc_data["contact"],
            "phone":        desc_data["phone"],
            "fax":          desc_data["fax"],
            "address":      desc_data["address"],
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
