"""
tests/test_syncer_gemini.py — syncer_gemini.py 離線情境測試（不觸發真實網路請求）
"""
from sync.syncer_gemini import extract_paid_amount


def test_extract_paid_amount_no_server_url():
    result = extract_paid_amount("留言：已付訂金5000", 10000, "", "secret")
    assert result == {"found": False, "paid_amount": 0.0, "reason": ""}

def test_extract_paid_amount_no_comments():
    result = extract_paid_amount("", 10000, "https://example.com", "secret")
    assert result == {"found": False, "paid_amount": 0.0, "reason": ""}
