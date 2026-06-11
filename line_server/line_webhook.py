"""
line_webhook.py — LINE Messaging API Webhook Server

接收顧客訊息 → Gemini 辨識結構化資料 → 存入 SQLite → 提供 REST API 讓 QuoteFlow 同步。

環境變數（在 Railway 的 Variables 頁面設定）：
  LINE_CHANNEL_SECRET       — LINE Messaging API Channel Secret
  LINE_CHANNEL_ACCESS_TOKEN — LINE Channel Access Token
  API_SECRET                — 自訂字串，QuoteFlow 連線時用來驗證身份
  GEMINI_API_KEY            — Google Gemini API 金鑰
  PORT                      — Railway 自動注入，不需手動設定
"""
import base64, hashlib, hmac, json, os, sqlite3
from pathlib import Path

from flask import Flask, abort, jsonify, request

app = Flask(__name__)

CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
CHANNEL_TOKEN  = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
API_SECRET     = os.environ.get("API_SECRET", "change-me-before-deploy")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

DB_PATH = Path("line_inquiries.db")

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


# ── 資料庫 ────────────────────────────────────────────────────

def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(str(DB_PATH), timeout=10)
    c.row_factory = sqlite3.Row
    return c


def _init_db():
    c = _conn()
    c.executescript("""
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
            created_at      TEXT    DEFAULT (datetime('now','+8 hours')),
            updated_at      TEXT    DEFAULT (datetime('now','+8 hours'))
        );
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
            c.execute(f"ALTER TABLE line_inquiries ADD COLUMN {col} {typedef}")
        except Exception:
            pass
    c.commit()
    c.close()


_init_db()


# ── Gemini 辨識 ───────────────────────────────────────────────

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
    """呼叫 Gemini，遇到 503 UNAVAILABLE 時自動重試最多 retries 次。"""
    import time
    last_err = None
    for attempt in range(retries):
        try:
            return client.models.generate_content(model=_GEMINI_MODEL, contents=contents)
        except Exception as e:
            last_err = e
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                # 每日額度耗盡，重試無效
                raise
            elif "503" in err_str or "UNAVAILABLE" in err_str:
                wait = delay * (attempt + 1)
                print(f"[Gemini RETRY] 第 {attempt+1} 次，等 {wait:.0f}s 後重試…", flush=True)
                time.sleep(wait)
            else:
                raise
    raise last_err


def _extract_info(message: str) -> dict:
    """從文字訊息提取結構化資訊。"""
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
    """從 LINE Content API 下載圖片。"""
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
    """從圖片（名片等）提取結構化資訊。"""
    if not GEMINI_API_KEY:
        print("[Gemini image SKIP] GEMINI_API_KEY 未設定", flush=True)
        return dict(_EMPTY_INFO)
    if not image_bytes:
        print("[Gemini image SKIP] 圖片下載失敗，image_bytes 為空", flush=True)
        return dict(_EMPTY_INFO)
    try:
        import io
        import PIL.Image
        from google.genai import types

        # 統一轉成 JPEG，相容 JFIF / PNG / WebP 等格式
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


# ── LINE 工具函式 ─────────────────────────────────────────────

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


# ── Webhook ───────────────────────────────────────────────────

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

        c = _conn()
        c.execute("""
            INSERT INTO line_inquiries
                (line_user_id, display_name, message,
                 company_name, tax_id, contact_name, mobile, phone,
                 fax, address, email, inquiry_product, area)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (user_id, display_name, text,
              info["company_name"], info["tax_id"], info["contact_name"],
              info["mobile"], info["phone"], info["fax"],
              info["address"], info["email"], info["inquiry_product"], info["area"]))
        c.commit()
        c.close()

    return "OK", 200


# ── REST API（供 QuoteFlow 桌面 app 呼叫）────────────────────

def _auth():
    if request.headers.get("X-API-Secret") != API_SECRET:
        abort(401)


@app.route("/api/inquiries", methods=["GET"])
def list_inquiries():
    _auth()
    status = request.args.get("status", "待處理")
    c = _conn()
    if status == "全部":
        rows = c.execute(
            "SELECT * FROM line_inquiries ORDER BY created_at DESC"
        ).fetchall()
    else:
        # 無論篩選條件，always 包含 sender='staff' 的我方回覆
        rows = c.execute(
            "SELECT * FROM line_inquiries WHERE status=? OR sender='staff' "
            "ORDER BY created_at DESC",
            (status,),
        ).fetchall()
    c.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/inquiries/<int:inquiry_id>", methods=["PATCH"])
def update_inquiry(inquiry_id: int):
    _auth()
    data = request.json or {}
    fields = ["status", "inquiry_type", "trello_card_id",
              "company_name", "tax_id", "contact_name", "mobile",
              "phone", "fax", "address", "email", "inquiry_product", "area"]
    sets   = ", ".join(f"{f} = COALESCE(?, {f})" for f in fields)
    values = [data.get(f) for f in fields] + [inquiry_id]
    c = _conn()
    c.execute(
        f"UPDATE line_inquiries SET {sets}, updated_at=datetime('now','+8 hours') WHERE id=?",
        values,
    )
    c.commit()
    c.close()
    return jsonify({"ok": True})


@app.route("/api/push_message", methods=["POST"])
def push_message():
    """發送 LINE Push Message 給顧客，並將訊息存入 DB。"""
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

    c = _conn()
    cur = c.execute("""
        INSERT INTO line_inquiries
            (line_user_id, display_name, message, sender, status,
             inquiry_type, created_at, updated_at)
        VALUES (?, '我方', ?, 'staff', '已忽略', '',
                datetime('now','+8 hours'), datetime('now','+8 hours'))
    """, (to, message))
    c.commit()
    row_id = cur.lastrowid
    row = dict(c.execute(
        "SELECT * FROM line_inquiries WHERE id=?", (row_id,)
    ).fetchone())
    c.close()
    return jsonify({"ok": True, "record": row})


@app.route("/api/extract_text", methods=["POST"])
def extract_text():
    """接收合併訊息文字，呼叫 Gemini 辨識並回傳結構化欄位。"""
    _auth()
    data = request.json or {}
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
