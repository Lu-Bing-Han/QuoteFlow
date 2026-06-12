"""
db.py — 資料庫連線與初始化
使用 SQLite（Python 內建，不需安裝額外套件）

資料庫路徑優先順序：
  1. config.json 裡的 "db_path"（使用者在 ⚙ 設定中指定，可指向 NAS 共用路徑）
  2. 執行檔同目錄下的 quoteflow.db（預設）
"""
import sqlite3
from pathlib import Path

from _paths import EXE_DIR

_DEFAULT_DB_PATH = EXE_DIR / "quoteflow.db"
_config_db_path: Path | None = None


def set_db_path(path: str | Path | None):
    """程式啟動時由 app_core 依 config 呼叫，設定資料庫路徑。"""
    global _config_db_path
    _config_db_path = Path(path) if path else None


def get_db_path() -> Path:
    return _config_db_path if _config_db_path else _DEFAULT_DB_PATH


def get_connection() -> sqlite3.Connection:
    """取得資料庫連線（呼叫端負責 .close()）。
    timeout=10：多人同時寫入時最多等 10 秒，避免衝突報錯。
    """
    conn = sqlite3.connect(str(get_db_path()), timeout=10)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """建立資料表（若已存在則不動）。程式啟動時呼叫一次即可。"""
    conn = get_connection()
    try:
        conn.executescript("""
            -- 報價單主檔
            CREATE TABLE IF NOT EXISTS quotes (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_no    TEXT    NOT NULL,        -- 報價單號，例如 QF-2026-001
                date        TEXT    NOT NULL,        -- 日期，格式 YYYY-MM-DD
                customer    TEXT    NOT NULL,        -- 客戶名稱
                contact     TEXT,                   -- 聯絡人
                phone       TEXT,                   -- 電話
                total       REAL    DEFAULT 0,      -- 總金額
                created_at  TEXT    DEFAULT (datetime('now','localtime'))
            );

            -- 報價單品項明細
            CREATE TABLE IF NOT EXISTS quote_items (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_id    INTEGER NOT NULL REFERENCES quotes(id) ON DELETE CASCADE,
                seq         INTEGER,                -- 序號
                code        TEXT,                   -- 型號
                name        TEXT,                   -- 品名
                spec        TEXT,                   -- 規格
                qty         INTEGER DEFAULT 1,      -- 數量
                unit        TEXT,                   -- 單位
                unit_price  REAL    DEFAULT 0,      -- 單價
                subtotal    REAL    DEFAULT 0       -- 小計
            );

            -- LINE 顧客詢問
            CREATE TABLE IF NOT EXISTS line_inquiries (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                line_user_id    TEXT    NOT NULL,
                display_name    TEXT    DEFAULT '未知顧客',
                message         TEXT    NOT NULL,
                inquiry_type    TEXT    DEFAULT '未分類',
                status          TEXT    DEFAULT '待處理',
                trello_card_id  TEXT,
                company_name    TEXT    DEFAULT '',
                tax_id          TEXT    DEFAULT '',
                contact_name    TEXT    DEFAULT '',
                mobile          TEXT    DEFAULT '',
                phone           TEXT    DEFAULT '',
                fax             TEXT    DEFAULT '',
                address         TEXT    DEFAULT '',
                email           TEXT    DEFAULT '',
                inquiry_product TEXT    DEFAULT '',
                area            TEXT    DEFAULT '',
                created_at      TEXT    DEFAULT (datetime('now','localtime')),
                updated_at      TEXT    DEFAULT (datetime('now','localtime'))
            );
        """)
        conn.commit()
        # 相容舊版 DB：若新欄位不存在則補上
        new_cols = [
            ("company_name",    "TEXT DEFAULT ''"),
            ("tax_id",          "TEXT DEFAULT ''"),
            ("contact_name",    "TEXT DEFAULT ''"),
            ("mobile",          "TEXT DEFAULT ''"),
            ("phone",           "TEXT DEFAULT ''"),
            ("fax",             "TEXT DEFAULT ''"),
            ("address",         "TEXT DEFAULT ''"),
            ("email",           "TEXT DEFAULT ''"),
            ("inquiry_product", "TEXT DEFAULT ''"),
            ("area",            "TEXT DEFAULT ''"),
            ("sender",          "TEXT DEFAULT 'customer'"),
        ]
        for col, typedef in new_cols:
            try:
                conn.execute(
                    f"ALTER TABLE line_inquiries ADD COLUMN {col} {typedef}"
                )
            except Exception:
                pass  # 欄位已存在屬正常，不需記錄
        conn.commit()
    finally:
        conn.close()
