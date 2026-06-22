"""
tests/test_generator_statement.py — generator_statement.py 工具函式測試
"""
from core.generator_statement import parse_amount, parse_paid_amount


# ── parse_amount ─────────────────────────────────────────────────────────────

def test_parse_amount_plain_number():
    assert parse_amount("12000") == 12000.0

def test_parse_amount_with_currency_and_commas():
    assert parse_amount("NT$12,000") == 12000.0

def test_parse_amount_empty():
    assert parse_amount("") == 0.0

def test_parse_amount_non_numeric():
    assert parse_amount("未確認") == 0.0


# ── parse_paid_amount ─────────────────────────────────────────────────────────

def test_parse_paid_amount_with_deposit():
    assert parse_paid_amount("匯款，已付訂金5000") == 5000.0

def test_parse_paid_amount_with_prefix_label():
    assert parse_paid_amount("現金，已收3000元") == 3000.0

def test_parse_paid_amount_no_hint():
    assert parse_paid_amount("匯款") == 0.0

def test_parse_paid_amount_empty():
    assert parse_paid_amount("") == 0.0
