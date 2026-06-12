"""pytest conftest — 設定 sys.path 讓 tests 能 import src/"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
