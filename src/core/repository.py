"""
repository.py — 報價單 CRUD 操作
UI 層只需 import 這裡的函式，不需要直接碰 SQL。
"""
from __future__ import annotations
from typing import Any
from .db import get_connection


# ════════════════════════════════════════════════════════
#  新增
# ════════════════════════════════════════════════════════

def save_quote(
    quote_no: str,
    date: str,
    customer: str,
    contact: str,
    phone: str,
    total: float,
    items: list[dict],
) -> int:
    """
    儲存一筆報價單（主檔 + 品項明細）。
    回傳新建的 quotes.id。

    items 格式（list of dict）：
        [{"seq": 1, "code": "APB-15", "name": "油壓拖板車", "spec": "...",
          "qty": 1, "unit": "台", "unit_price": 12000, "subtotal": 12000}, ...]
    """
    conn = get_connection()
    try:
        cur = conn.execute(
            """INSERT INTO quotes (quote_no, date, customer, contact, phone, total)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (quote_no, date, customer, contact, phone, total),
        )
        quote_id = cur.lastrowid

        conn.executemany(
            """INSERT INTO quote_items
               (quote_id, seq, code, name, spec, qty, unit, unit_price, subtotal)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            [
                (quote_id,
                 it.get("seq"), it.get("code"), it.get("name"), it.get("spec"),
                 it.get("qty", 1), it.get("unit"), it.get("unit_price", 0),
                 it.get("subtotal", 0))
                for it in items
            ],
        )
        conn.commit()
        return quote_id
    finally:
        conn.close()


# ════════════════════════════════════════════════════════
#  查詢
# ════════════════════════════════════════════════════════

def list_quotes(keyword: str = "") -> list[dict]:
    """
    列出所有報價單（可用關鍵字搜尋客戶名稱或報價單號）。
    回傳 list of dict，每筆包含主檔欄位（不含品項明細）。
    """
    conn = get_connection()
    try:
        like = f"%{keyword}%"
        rows = conn.execute(
            """SELECT id, quote_no, date, customer, contact, phone, total, created_at
               FROM quotes
               WHERE customer LIKE ? OR quote_no LIKE ?
               ORDER BY id DESC""",
            (like, like),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_quote(quote_id: int) -> dict[str, Any] | None:
    """
    取得單筆報價單（含品項明細）。
    回傳 {"quote": {...}, "items": [...]}，找不到回傳 None。
    """
    conn = get_connection()
    try:
        row = conn.execute(
            "SELECT * FROM quotes WHERE id = ?", (quote_id,)
        ).fetchone()
        if not row:
            return None
        items = conn.execute(
            "SELECT * FROM quote_items WHERE quote_id = ? ORDER BY seq",
            (quote_id,),
        ).fetchall()
        return {"quote": dict(row), "items": [dict(i) for i in items]}
    finally:
        conn.close()


def search_quotes_by_codes(codes: list[str], customer: str = "") -> list[dict]:
    """
    用品號組合搜尋報價單——回傳同時包含所有指定品號的報價單。
    codes 支援模糊比對，例如 "AP-30" 可以找到 "AP-30B"。
    customer 可額外過濾客戶名稱（模糊）。
    """
    if not codes:
        return list_quotes(keyword=customer)

    conn = get_connection()
    try:
        # 基礎 SQL：每個品號產生一個 EXISTS 子查詢，全部都要符合
        conditions = []
        params: list[str] = []
        for code in codes:
            conditions.append(
                "EXISTS (SELECT 1 FROM quote_items qi "
                "WHERE qi.quote_id = q.id AND qi.code LIKE ?)"
            )
            params.append(f"%{code.strip()}%")

        if customer:
            conditions.append("q.customer LIKE ?")
            params.append(f"%{customer.strip()}%")

        where = " AND ".join(conditions)
        sql = f"""
            SELECT q.id, q.quote_no, q.date, q.customer, q.contact,
                   q.phone, q.total, q.created_at
            FROM quotes q
            WHERE {where}
            ORDER BY q.id DESC
        """
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def find_same_combination(codes: list[str]) -> list[dict]:
    """
    找出品號組合完全相同的歷史報價單。
    「完全相同」= 品號集合一致（不多也不少），數量不限。
    回傳 list of dict，含 quote_no, date, customer, total，最多 10 筆，最新優先。
    """
    if not codes:
        return []

    unique_codes = list(dict.fromkeys(c.strip() for c in codes if c.strip()))
    n = len(unique_codes)
    if n == 0:
        return []

    conn = get_connection()
    try:
        # 條件 1：包含所有指定品號
        has_all = " AND ".join(
            f"EXISTS (SELECT 1 FROM quote_items qi "
            f"WHERE qi.quote_id = q.id AND qi.code = ?)"
            for _ in unique_codes
        )
        # 條件 2：品號種類數量完全一致（不多也不少）
        sql = f"""
            SELECT q.id, q.quote_no, q.date, q.customer, q.total
            FROM quotes q
            WHERE {has_all}
              AND (SELECT COUNT(DISTINCT qi2.code)
                   FROM quote_items qi2 WHERE qi2.quote_id = q.id) = ?
            ORDER BY q.date DESC
            LIMIT 10
        """
        rows = conn.execute(sql, unique_codes + [n]).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def find_similar_combinations(codes: list[str]) -> list[dict]:
    """
    找出「相似」組合的歷史報價單：
    - 多一個品號：包含所有指定品號，且剛好多 1 個其他品號
    - 少一個品號：包含所有指定品號中的任意 N-1 個，且品號種類剛好是 N-1

    回傳 list of dict，每筆多一個 "similarity" 欄位說明差異。
    """
    unique_codes = list(dict.fromkeys(c.strip() for c in codes if c.strip()))
    n = len(unique_codes)
    if n == 0:
        return []

    conn = get_connection()
    results = []
    try:
        # ── 多一個品號（superset +1）────────────────────────
        has_all = " AND ".join(
            f"EXISTS (SELECT 1 FROM quote_items qi WHERE qi.quote_id = q.id AND qi.code = ?)"
            for _ in unique_codes
        )
        sql_plus = f"""
            SELECT q.id, q.quote_no, q.date, q.customer, q.total
            FROM quotes q
            WHERE {has_all}
              AND (SELECT COUNT(DISTINCT qi2.code)
                   FROM quote_items qi2 WHERE qi2.quote_id = q.id) = ?
            ORDER BY q.date DESC LIMIT 5
        """
        for row in conn.execute(sql_plus, unique_codes + [n + 1]).fetchall():
            # 找出多出來的那個品號
            extra = conn.execute(
                """SELECT code FROM quote_items
                   WHERE quote_id = ? AND code NOT IN ({})
                   GROUP BY code""".format(",".join("?" * n)),
                [row["id"]] + unique_codes,
            ).fetchall()
            extra_str = "、".join(r["code"] for r in extra)
            d = dict(row)
            d["similarity"] = f"多：{extra_str}"
            results.append(d)

        # ── 少一個品號（subset -1）──────────────────────────
        if n >= 2:
            for i, missing_code in enumerate(unique_codes):
                subset = [c for j, c in enumerate(unique_codes) if j != i]
                has_subset = " AND ".join(
                    f"EXISTS (SELECT 1 FROM quote_items qi WHERE qi.quote_id = q.id AND qi.code = ?)"
                    for _ in subset
                )
                sql_minus = f"""
                    SELECT q.id, q.quote_no, q.date, q.customer, q.total
                    FROM quotes q
                    WHERE {has_subset}
                      AND (SELECT COUNT(DISTINCT qi2.code)
                           FROM quote_items qi2 WHERE qi2.quote_id = q.id) = ?
                    ORDER BY q.date DESC LIMIT 3
                """
                for row in conn.execute(sql_minus, subset + [n - 1]).fetchall():
                    d = dict(row)
                    d["similarity"] = f"少：{missing_code}"
                    results.append(d)

        # 去重（同一筆 id 可能被多條路徑命中）
        seen, unique_results = set(), []
        for r in results:
            if r["id"] not in seen:
                seen.add(r["id"])
                unique_results.append(r)
        return unique_results
    finally:
        conn.close()


# ════════════════════════════════════════════════════════
#  刪除
# ════════════════════════════════════════════════════════

def delete_quote(quote_id: int):
    """刪除一筆報價單（品項明細會連帶刪除）。"""
    conn = get_connection()
    try:
        conn.execute("DELETE FROM quotes WHERE id = ?", (quote_id,))
        conn.commit()
    finally:
        conn.close()
