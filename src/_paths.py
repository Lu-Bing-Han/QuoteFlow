"""
_paths.py — 統一管理路徑，相容開發環境與 PyInstaller 打包後環境
"""
import sys
from pathlib import Path


def _exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def _template_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "template"  # type: ignore[attr-defined]
    return Path(__file__).parent.parent / "template"


EXE_DIR      = _exe_dir()
TEMPLATE_DIR = _template_dir()
OUTPUT_DIR   = EXE_DIR / "output"
CONFIG_PATH  = EXE_DIR / "config.json"
ICON_PATH    = TEMPLATE_DIR / "icon.png"

# Derived paths used by sync tabs
_GSHEETS_TOKEN_PATH     = EXE_DIR  / "gsheets_token.json"
_SYNCED_CARDS_PATH      = EXE_DIR  / "synced_cards.json"
_GSHEETS_CREDS_PATH     = TEMPLATE_DIR / "credentials.json"
_PRODUCTION_SYNCED_PATH = EXE_DIR  / "production_synced_cards.json"
