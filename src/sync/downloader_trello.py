"""
downloader_trello.py — 從 Trello 下載卡片資料（Excel + 附件）到本機資料夾
"""
import re
from pathlib import Path

import openpyxl
import requests

_API_BASE   = "https://api.trello.com/1"
_BOARD_NAME = "物流事業部1"


def _auth(api_key: str, token: str) -> dict:
    return {"key": api_key, "token": token}


def get_board_lists(api_key: str, token: str, board_name: str = _BOARD_NAME) -> list[dict]:
    """回傳指定看板的所有清單 [{"id": ..., "name": ...}, ...]。"""
    resp = requests.get(
        f"{_API_BASE}/members/me/boards",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    board_id = None
    for b in resp.json():
        if b["name"] == board_name:
            board_id = b["id"]
            break
    if not board_id:
        raise ValueError(f"找不到看板「{board_name}」")

    resp = requests.get(
        f"{_API_BASE}/boards/{board_id}/lists",
        params={**_auth(api_key, token), "fields": "name,id"},
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()


def get_list_cards(list_id: str, api_key: str, token: str) -> list[dict]:
    """抓取清單內所有卡片（含標籤、附件清單），供 GUI 預覽使用。"""
    resp = requests.get(
        f"{_API_BASE}/lists/{list_id}/cards",
        params={
            **_auth(api_key, token),
            "attachments": "true",
            "fields": "name,desc,labels,attachments,shortUrl",
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()


def download_cards(
    cards: list[dict],
    output_dir: Path,
    api_key: str,
    token: str,
    progress_cb=None,
) -> tuple[int, list[str]]:
    """下載指定卡片到 output_dir，每張卡片建一個資料夾。

    cards 為 get_list_cards() 回傳的子集（使用者選取的卡片）。
    資料夾內包含：
    - 卡片資料.xlsx（標題、標籤、描述）
    - 所有附件原檔

    progress_cb(current, total, card_name) 可選，用於更新進度。
    回傳 (卡片數, 失敗清單)。
    """
    auth = _auth(api_key, token)
    output_dir.mkdir(parents=True, exist_ok=True)
    failed: list[str] = []

    for idx, card in enumerate(cards, 1):
        card_name = card.get("name", f"card_{idx}")
        if progress_cb:
            progress_cb(idx, len(cards), card_name)

        # ── 建立卡片資料夾 ─────────────────────────────
        safe_name = re.sub(r'[\\/:*?"<>|]', '_', card_name).strip() or f"card_{idx}"
        card_dir = output_dir / safe_name
        card_dir.mkdir(parents=True, exist_ok=True)

        # ── 寫入 Excel ────────────────────────────────
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "卡片資料"
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 60

        labels = "、".join(
            lbl.get("name") or lbl.get("color", "")
            for lbl in card.get("labels", [])
        )

        for field, value in [
            ("標題", card_name),
            ("標籤", labels),
            ("描述", card.get("desc", "")),
        ]:
            ws.append([field, value])

        wb.save(str(card_dir / "卡片資料.xlsx"))
        wb.close()

        # ── 下載附件 ──────────────────────────────────
        for att in card.get("attachments") or []:
            url = att.get("url", "")
            if not url:
                continue
            fname = att.get("name") or url.split("/")[-1]
            safe_fname = re.sub(r'[\\/:*?"<>|]', '_', fname) or "attachment"
            try:
                # Trello 新版 Token（ATTA…）需用 Authorization header 才能下載附件
                headers = {
                    "Authorization": (
                        f'OAuth oauth_consumer_key="{api_key}",'
                        f' oauth_token="{token}"'
                    )
                }
                r = requests.get(
                    url,
                    params=auth,
                    headers=headers,
                    timeout=60,
                    stream=True,
                )
                r.raise_for_status()
                with open(card_dir / safe_fname, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            except Exception as e:
                failed.append(f"{safe_name}/{safe_fname}：{e}")

    return len(cards), failed
