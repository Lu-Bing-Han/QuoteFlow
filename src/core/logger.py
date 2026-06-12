"""
logger.py — 統一 logging 設定

使用方式：
    from core.logger import get_logger
    log = get_logger(__name__)
    log.info("...")
    log.warning("...")
    log.error("...", exc_info=True)

log 檔位置：執行檔同目錄 / quoteflow.log（最多 5 MB × 3 個備份）
"""
import logging
import logging.handlers
from pathlib import Path


def _log_path() -> Path:
    try:
        from _paths import EXE_DIR
        return EXE_DIR / "quoteflow.log"
    except Exception:
        return Path(__file__).parent.parent.parent / "quoteflow.log"


def _setup():
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)-8s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    root = logging.getLogger()
    if root.handlers:
        return  # 已初始化，避免重複
    root.setLevel(logging.DEBUG)

    fh = logging.handlers.RotatingFileHandler(
        _log_path(), maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(fmt)
    root.addHandler(ch)


_setup()


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)
