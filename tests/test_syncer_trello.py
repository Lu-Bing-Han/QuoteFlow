from datetime import date

from sync.syncer_trello import (
    _parse_attachment_order_date,
    _parse_comment_order_date,
    _parse_desc,
    _parse_move_to_list_date,
)


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
