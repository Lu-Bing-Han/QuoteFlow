from datetime import date

import requests

from sync.syncer_trello import (
    _extract_comments,
    _fmt_trello_error,
    _parse_attachment_order_date,
    _parse_comment_order_date,
    _parse_desc,
    _parse_move_to_list_date,
    _resolve_order_date,
)


def test_resolve_order_date_skips_api_calls_when_disabled():
    # 標題有【...】日期標記，resolve_via_api=False 時應直接從標題判斷，
    # 完全不打 API（用假的 api_key/token 也不會出錯，因為根本不會用到）。
    card = {"id": "abc123", "name": "CH測試公司【114.8.14訂】產品", "actions": [], "attachments": []}
    desc_data = {"order_date_str": "", "order_date_dt": None}
    order_date, order_dt = _resolve_order_date(
        card, desc_data, list_id="dummy",
        api_key="not-a-real-key", token="not-a-real-token",
        resolve_via_api=False)
    assert order_date == "8/14"
    assert order_dt == date(2025, 8, 14)


def test_fmt_trello_error_timeout():
    assert _fmt_trello_error(requests.exceptions.Timeout("timed out")) == \
        "連線逾時，請檢查網路連線後再試一次"

def test_fmt_trello_error_connection_error_redacts_url():
    e = requests.exceptions.ConnectionError(
        "Max retries exceeded with url: /1/lists/x/cards?key=abc&token=ATTAxyz")
    assert _fmt_trello_error(e) == "無法連線到 Trello，請檢查網路連線後再試一次"

def test_fmt_trello_error_401_redacts_token():
    e = Exception("401 Client Error: Unauthorized for url: "
                   "https://api.trello.com/1/members/me/boards?key=abc&token=ATTAxyz")
    assert _fmt_trello_error(e) == "Trello 憑證無效或已過期，請至「出貨一覽表」頁籤重新填入並儲存 API Key/Token"

def test_fmt_trello_error_unrecognized_redacts_and_truncates():
    long_url = "https://api.trello.com/1/x?key=abc&token=ATTAxyz" + "a" * 300
    e = Exception(f"weird error for url: {long_url}")
    result = _fmt_trello_error(e)
    assert "abc" not in result and "ATTAxyz" not in result
    assert len(result) <= 201


def test_extract_comments_sorts_chronologically_and_filters_type():
    actions = [
        {"type": "commentCard", "date": "2026-06-20T00:00:00.000Z", "data": {"text": "後來的留言"}},
        {"type": "updateCard",  "date": "2026-06-15T00:00:00.000Z", "data": {"text": "不應該出現"}},
        {"type": "commentCard", "date": "2023-06-21T00:00:00.000Z", "data": {"text": "收到50%訂金"}},
    ]
    result = _extract_comments(actions)
    assert result == "[2023-06-21] 收到50%訂金\n[2026-06-20] 後來的留言"

def test_extract_comments_empty():
    assert _extract_comments([]) == ""


def test_parse_desc_tax_id_with_colon():
    parsed = _parse_desc("統一編號:55971514\n聯絡人: 王小姐")
    assert parsed["tax_id"] == "55971514"

def test_parse_desc_tax_id_short_label_strips_trailing_note():
    parsed = _parse_desc("統編 : 54159940(GOOGLE)\n地址 : 台北市")
    assert parsed["tax_id"] == "54159940"

def test_parse_desc_tax_id_missing():
    assert _parse_desc("聯絡人: 王小姐")["tax_id"] == ""


def test_parse_comment_order_date_roc_dots():
    assert _parse_comment_order_date("115.06.18已下單") == ("6/18", date(2026, 6, 18))


def test_parse_comment_order_date_roc_slashes_with_space():
    assert _parse_comment_order_date("115/6/5 已下單") == ("6/5", date(2026, 6, 5))


def test_parse_comment_order_date_trello_escaped_dot():
    assert _parse_comment_order_date(r"115\.06.18已下單") == ("6/18", date(2026, 6, 18))


def test_parse_comment_order_date_requires_ordered_marker():
    assert _parse_comment_order_date("115.06.18 客戶確認") == ("", None)


def test_desc_order_date_still_parses_first():
    parsed = _parse_desc("下單日期：115.06.16\n付款方式：匯款")

    assert parsed["order_date_str"] == "6/16"
    assert parsed["order_date_dt"] == date(2026, 6, 16)


def test_parse_attachment_order_date_compact_roc_filename():
    attachments = [
        {
            "name": (
                "訂購單-台灣大食品股份有限公司(台中清水)"
                "陳尚鴻先生RBT25LL-1150618.jpg"
            )
        }
    ]

    assert _parse_attachment_order_date(attachments) == ("6/18", date(2026, 6, 18))


def test_parse_attachment_order_date_prefers_purchase_order_dot_date():
    attachments = [
        {"name": "報價單-客戶-1150518.jpg"},
        {"name": "訂購單-客戶-115.6.18.jpg"},
    ]

    assert _parse_attachment_order_date(attachments) == ("6/18", date(2026, 6, 18))


def test_parse_move_to_list_date_uses_matching_list_after():
    actions = [
        {
            "type": "updateCard",
            "date": "2026-06-12T03:04:05.000Z",
            "data": {"listAfter": {"id": "other"}},
        },
        {
            "type": "updateCard",
            "date": "2026-06-20T08:30:00.000Z",
            "data": {"listAfter": {"id": "target-list"}},
        },
    ]

    assert _parse_move_to_list_date(actions, "target-list") == ("6/20", date(2026, 6, 20))
