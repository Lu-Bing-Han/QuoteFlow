"""
syncer_gemini.py — 透過既有 Railway LINE server 呼叫 Gemini，輔助判斷對帳單已收金額
"""
import requests


def extract_paid_amount(comments_text: str, total_amount: float,
                         server_url: str, secret: str, timeout: int = 30) -> dict:
    """請伺服器用 Gemini 讀留言判斷已收金額。
    回傳 {"found": bool, "paid_amount": float, "reason": str}；
    伺服器未設定、沒有留言、或呼叫失敗時 found=False。
    """
    empty = {"found": False, "paid_amount": 0.0, "reason": ""}
    url = (server_url or "").rstrip("/")
    if not url or not (comments_text or "").strip():
        return empty

    resp = requests.post(
        f"{url}/api/extract_paid_amount",
        json={"comments": comments_text, "total_amount": total_amount},
        headers={"X-API-Secret": secret, "Content-Type": "application/json"},
        timeout=timeout,
    )
    resp.raise_for_status()
    data = resp.json()
    return {
        "found":       bool(data.get("found")),
        "paid_amount": float(data.get("paid_amount") or 0),
        "reason":      str(data.get("reason", "")),
    }
