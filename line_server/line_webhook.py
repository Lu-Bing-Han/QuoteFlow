"""
line_webhook.py — LINE Messaging API Webhook Server (PostgreSQL)

接收顧客訊息 → Gemini 辨識結構化資料 → 存入 PostgreSQL → 提供 REST API 讓 QuoteFlow 同步。

環境變數（在 Railway 的 Variables 頁面設定）：
  LINE_CHANNEL_SECRET       — LINE Messaging API Channel Secret
  LINE_CHANNEL_ACCESS_TOKEN — LINE Channel Access Token
  API_SECRET                — 自訂字串，QuoteFlow 連線時用來驗證身份
  GEMINI_API_KEY            — Google Gemini API 金鑰
  DATABASE_URL              — Railway PostgreSQL 自動注入
  PORT                      — Railway 自動注入，不需手動設定
"""
import base64, hashlib, hmac, json, os
from datetime import datetime, timezone, timedelta

from flask import Flask, abort, jsonify, request

app = Flask(__name__)

CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
CHANNEL_TOKEN  = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
API_SECRET     = os.environ.get("API_SECRET", "change-me-before-deploy")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

_raw_db_url   = os.environ.get("DATABASE_URL", "")
DATABASE_URL  = _raw_db_url.replace("postgres://", "postgresql://", 1) \
                if _raw_db_url.startswith("postgres://") else _raw_db_url

_EMPTY_INFO = {
    "company_name":    "",
    "tax_id":          "",
    "contact_name":    "",
    "mobile":          "",
    "phone":           "",
    "fax":             "",
    "address":         "",
    "email":           "",
    "inquiry_product": "",
    "area":            "",
}

_GEMINI_MODEL = "gemini-2.5-flash-lite"

_FIELD_PROMPT = """欄位說明：
- company_name：公司名稱
- tax_id：統一編號（8位數字）
- contact_name：聯絡人姓名
- mobile：手機號碼（09開頭）
- phone：市話號碼
- fax：傳真號碼
- address：地址
- email：電子郵件
- inquiry_product：詢價或維修的商品名稱/型號
- area：公司所在市區（如「台北市」「新北市」「台中市」）

只回傳 JSON 物件，不要包含 markdown 或其他文字。"""


# ── 時間 ─────────────────────────────────────────────────

def _now_tw() -> str:
    """台灣時間（UTC+8）字串 YYYY-MM-DD HH:MM:SS"""
    tz = timezone(timedelta(hours=8))
    return datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')


# ── 資料庫 ────────────────────────────────────────────────

def _conn():
    import psycopg2
    import psycopg2.extras
    return psycopg2.connect(DATABASE_URL,
                            cursor_factory=psycopg2.extras.RealDictCursor)


def _init_db():
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS line_inquiries (
            id              SERIAL PRIMARY KEY,
            line_user_id    TEXT    NOT NULL,
            display_name    TEXT    DEFAULT '未知顧客',
            message         TEXT    NOT NULL DEFAULT '',
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
            sender          TEXT    DEFAULT 'customer',
            created_at      TEXT,
            updated_at      TEXT
        )
    """)
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
            cur.execute(
                f"ALTER TABLE line_inquiries ADD COLUMN IF NOT EXISTS {col} {typedef}"
            )
        except Exception:
            pass
    conn.commit()
    cur.close()
    conn.close()


def _merge_duplicate_names():
    """將相同 display_name（忽略空白）但不同 line_user_id 的記錄合併為最舊那個 ID。"""
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("""
        SELECT TRIM(display_name) as name
        FROM line_inquiries
        GROUP BY TRIM(display_name)
        HAVING COUNT(DISTINCT line_user_id) > 1
    """)
    names = [r["name"] for r in cur.fetchall()]
    print(f"[merge] 找到 {len(names)} 個需要合併的名稱: {names}", flush=True)
    for name in names:
        cur.execute("""
            SELECT line_user_id FROM line_inquiries
            WHERE TRIM(display_name) = %s
            GROUP BY line_user_id
            ORDER BY MIN(created_at) ASC
        """, (name,))
        uids = [r["line_user_id"] for r in cur.fetchall()]
        if len(uids) > 1:
            canonical    = uids[0]
            placeholders = ",".join(["%s"] * (len(uids) - 1))
            cur.execute(
                f"UPDATE line_inquiries SET line_user_id=%s "
                f"WHERE line_user_id IN ({placeholders})",
                [canonical] + uids[1:],
            )
            print(f"[merge] '{name}': {len(uids)} IDs → {canonical}", flush=True)
    conn.commit()
    cur.close()
    conn.close()


_init_db()
_merge_duplicate_names()


# ── Gemini 辨識 ───────────────────────────────────────────

def _parse_json(text: str) -> dict:
    text = text.strip()
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.split("```")[0]
    result = json.loads(text.strip())
    return {k: str(result.get(k, "")).strip() for k in _EMPTY_INFO}


def _gemini_client():
    from google import genai
    return genai.Client(api_key=GEMINI_API_KEY)


def _gemini_generate(client, contents, retries: int = 3, delay: float = 3.0):
    import time
    last_err = None
    for attempt in range(retries):
        try:
            return client.models.generate_content(model=_GEMINI_MODEL, contents=contents)
        except Exception as e:
            last_err = e
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                raise
            elif "503" in err_str or "UNAVAILABLE" in err_str:
                wait = delay * (attempt + 1)
                print(f"[Gemini RETRY] 第 {attempt+1} 次，等 {wait:.0f}s 後重試…", flush=True)
                time.sleep(wait)
            else:
                raise
    raise last_err


def _extract_info(message: str) -> dict:
    if not GEMINI_API_KEY:
        print("[Gemini text SKIP] GEMINI_API_KEY 未設定", flush=True)
        return dict(_EMPTY_INFO)
    try:
        client = _gemini_client()
        prompt = (
            "你是一個資料提取助手。從以下顧客 LINE 訊息中提取資訊，以 JSON 格式回傳。\n"
            "找不到的欄位填空字串 \"\"，不要猜測或捏造資料。\n\n"
            + _FIELD_PROMPT
            + f"\n\n顧客訊息：\n{message}"
        )
        resp = _gemini_generate(client, prompt)
        print(f"[Gemini text RAW] {resp.text!r}", flush=True)
        return _parse_json(resp.text)
    except Exception as e:
        print(f"[Gemini text ERROR] {type(e).__name__}: {e}", flush=True)
        return dict(_EMPTY_INFO)


def _download_image(message_id: str) -> bytes | None:
    import urllib.request
    req = urllib.request.Request(
        f"https://api-data.line.me/v2/bot/message/{message_id}/content",
        headers={"Authorization": f"Bearer {CHANNEL_TOKEN}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return r.read()
    except Exception as e:
        print(f"[Download image ERROR] {type(e).__name__}: {e}", flush=True)
        return None


def _extract_info_from_image(image_bytes: bytes) -> dict:
    if not GEMINI_API_KEY:
        print("[Gemini image SKIP] GEMINI_API_KEY 未設定", flush=True)
        return dict(_EMPTY_INFO)
    if not image_bytes:
        return dict(_EMPTY_INFO)
    try:
        import io
        import PIL.Image
        from google.genai import types

        img = PIL.Image.open(io.BytesIO(image_bytes)).convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="JPEG")

        client = _gemini_client()
        prompt = (
            "從這張圖片中提取聯絡資訊（名片、文件等），以 JSON 格式回傳。\n"
            "找不到的欄位填空字串 \"\"，不要猜測或捏造資料。\n\n"
            + _FIELD_PROMPT
        )
        resp = _gemini_generate(client, [
            prompt,
            types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"),
        ])
        print(f"[Gemini image RAW] {resp.text!r}", flush=True)
        return _parse_json(resp.text)
    except Exception as e:
        print(f"[Gemini image ERROR] {type(e).__name__}: {e}", flush=True)
        return dict(_EMPTY_INFO)


# ── LINE 工具函式 ─────────────────────────────────────────

def _verify_signature(body: bytes, signature: str) -> bool:
    digest = hmac.new(CHANNEL_SECRET.encode(), body, hashlib.sha256).digest()
    return hmac.compare_digest(base64.b64encode(digest).decode(), signature)


def _get_display_name(user_id: str) -> str:
    import urllib.request
    req = urllib.request.Request(
        f"https://api.line.me/v2/bot/profile/{user_id}",
        headers={"Authorization": f"Bearer {CHANNEL_TOKEN}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=5) as r:
            return json.loads(r.read()).get("displayName", "未知顧客")
    except Exception:
        return "未知顧客"


# ── Webhook ───────────────────────────────────────────────

@app.route("/webhook", methods=["POST"])
def webhook():
    body = request.get_data()
    sig  = request.headers.get("X-Line-Signature", "")
    if not _verify_signature(body, sig):
        abort(400)

    for event in request.json.get("events", []):
        if event.get("type") != "message":
            continue
        msg          = event.get("message", {})
        msg_type     = msg.get("type")
        user_id      = event["source"]["userId"]
        display_name = _get_display_name(user_id)

        if msg_type == "text":
            text = msg["text"]
            info = _extract_info(text)
        elif msg_type == "image":
            image_bytes = _download_image(msg["id"])
            text        = "（顧客傳送了圖片）"
            info        = _extract_info_from_image(image_bytes) if image_bytes else dict(_EMPTY_INFO)
        else:
            continue

        now  = _now_tw()
        conn = _conn()
        cur  = conn.cursor()
        cur.execute("""
            INSERT INTO line_inquiries
                (line_user_id, display_name, message,
                 company_name, tax_id, contact_name, mobile, phone,
                 fax, address, email, inquiry_product, area,
                 created_at, updated_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (user_id, display_name, text,
              info["company_name"], info["tax_id"], info["contact_name"],
              info["mobile"], info["phone"], info["fax"],
              info["address"], info["email"], info["inquiry_product"], info["area"],
              now, now))
        conn.commit()
        cur.close()
        conn.close()

    return "OK", 200


# ── REST API（供 QuoteFlow 桌面 app 呼叫）────────────────

def _auth():
    if request.headers.get("X-API-Secret") != API_SECRET:
        abort(401)


@app.route("/api/inquiries", methods=["GET"])
def list_inquiries():
    _auth()
    status = request.args.get("status", "待處理")
    conn = _conn()
    cur  = conn.cursor()
    if status == "全部":
        cur.execute("SELECT * FROM line_inquiries ORDER BY created_at DESC")
    else:
        cur.execute(
            "SELECT * FROM line_inquiries WHERE status=%s OR sender='staff' "
            "ORDER BY created_at DESC",
            (status,),
        )
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/inquiries/<int:inquiry_id>", methods=["PATCH"])
def update_inquiry(inquiry_id: int):
    _auth()
    data   = request.json or {}
    fields = ["status", "inquiry_type", "trello_card_id",
              "company_name", "tax_id", "contact_name", "mobile",
              "phone", "fax", "address", "email", "inquiry_product", "area"]
    sets   = ", ".join(f"{f} = COALESCE(%s, {f})" for f in fields)
    values = [data.get(f) for f in fields] + [_now_tw(), inquiry_id]
    conn = _conn()
    cur  = conn.cursor()
    cur.execute(
        f"UPDATE line_inquiries SET {sets}, updated_at=%s WHERE id=%s",
        values,
    )
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({"ok": True})


@app.route("/api/push_message", methods=["POST"])
def push_message():
    _auth()
    data    = request.json or {}
    to      = data.get("to", "").strip()
    message = data.get("message", "").strip()
    if not to or not message:
        return jsonify({"ok": False, "error": "missing to or message"}), 400

    import urllib.request as _ur
    body = json.dumps({
        "to": to,
        "messages": [{"type": "text", "text": message}],
    }).encode()
    req = _ur.Request(
        "https://api.line.me/v2/bot/message/push",
        data=body,
        headers={
            "Authorization": f"Bearer {CHANNEL_TOKEN}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with _ur.urlopen(req, timeout=10):
            pass
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

    now  = _now_tw()
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("""
        INSERT INTO line_inquiries
            (line_user_id, display_name, message, sender, status,
             inquiry_type, created_at, updated_at)
        VALUES (%s, '我方', %s, 'staff', '已忽略', '', %s, %s)
        RETURNING id
    """, (to, message, now, now))
    row_id = cur.fetchone()["id"]
    conn.commit()
    cur.execute("SELECT * FROM line_inquiries WHERE id=%s", (row_id,))
    row = dict(cur.fetchone())
    cur.close()
    conn.close()
    return jsonify({"ok": True, "record": row})


@app.route("/api/merge_by_name", methods=["POST"])
def merge_by_name():
    _auth()
    data         = request.json or {}
    display_name = data.get("display_name", "").strip()
    if not display_name:
        return jsonify({"ok": False, "error": "missing display_name"}), 400
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("""
        SELECT line_user_id
        FROM line_inquiries
        WHERE display_name = %s
        GROUP BY line_user_id
        ORDER BY MIN(created_at) ASC
    """, (display_name,))
    uids = [r["line_user_id"] for r in cur.fetchall()]
    if len(uids) <= 1:
        cur.close()
        conn.close()
        return jsonify({"ok": True, "merged": 0, "canonical": uids[0] if uids else None})
    canonical     = uids[0]
    placeholders  = ",".join(["%s"] * (len(uids) - 1))
    cur.execute(
        f"UPDATE line_inquiries SET line_user_id=%s WHERE line_user_id IN ({placeholders})",
        [canonical] + uids[1:],
    )
    merged = cur.rowcount
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({"ok": True, "merged": merged, "canonical": canonical})


@app.route("/api/extract_text", methods=["POST"])
def extract_text():
    _auth()
    data    = request.json or {}
    message = data.get("message", "").strip()
    if not message:
        return jsonify(_EMPTY_INFO)
    info = _extract_info(message)
    return jsonify(info)


@app.route("/health")
def health():
    return "ok", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
