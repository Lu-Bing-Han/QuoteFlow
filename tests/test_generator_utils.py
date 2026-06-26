"""
tests/test_generator_utils.py — generator.py 工具函式測試
"""
import pytest
from core.generator import _format_date, trello_card_to_data


# ── _format_date ─────────────────────────────────────────────────────────────

def test_format_date_valid():
    result = _format_date("2026/06/01")
    assert "2026" in result
    assert "06" in result
    assert "01" in result

def test_format_date_preserves_year():
    result = _format_date("2026/01/15")
    assert result.startswith("2026")

def test_format_date_invalid_passthrough():
    assert _format_date("不是日期") == "不是日期"

def test_format_date_empty():
    assert _format_date("") == ""


# ── trello_card_to_data ────────────────────────────────────────────────────

def test_trello_card_to_data_basic_fields():
    card = {
        "company": "ABC公司", "phone": "04-1234567", "contact": "王先生",
        "address": "台中市", "tax_id": "12345678",
        "product": "電池*5", "quantity": "1", "amount": "NT$10,000",
    }
    data = trello_card_to_data(card)
    assert data["header"] == {
        "customer": "ABC公司", "phone": "04-1234567", "contact": "王先生",
        "address": "台中市", "tax_id": "12345678", "quote_no": "",
    }
    item = data["items"][0]
    assert item["name"] == "電池"
    assert item["qty"] == "5"
    assert item["unit_price"] == 2000.0
    assert item["subtotal"] == 10000.0

def test_trello_card_to_data_prefers_company_desc_over_title():
    card = {"company": "title公司-分廠", "company_desc": "正式公司名稱",
            "product": "機台", "quantity": "1", "amount": "1000"}
    assert trello_card_to_data(card)["header"]["customer"] == "正式公司名稱"

def test_trello_card_to_data_address_falls_back_to_location_lookup():
    card = {"company": "ABC公司", "product": "機台", "quantity": "1", "amount": "1000"}
    data = trello_card_to_data(card, location_lookup={"ABC公司": "台南市"})
    assert data["header"]["address"] == "台南市"

def test_trello_card_to_data_no_quantity_marker_keeps_fallback():
    card = {"company": "ABC公司", "product": "機台", "quantity": "3", "amount": "3000"}
    item = trello_card_to_data(card)["items"][0]
    assert item["qty"] == "3"
    assert item["unit_price"] == 1000.0
