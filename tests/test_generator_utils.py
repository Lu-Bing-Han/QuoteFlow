"""
tests/test_generator_utils.py — generator.py 工具函式測試
"""
import pytest
from core.generator import _format_date


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
