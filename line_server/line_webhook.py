"""
line_webhook.py — LINE Messaging API Webhook Server

接收顧客訊息 → 存入 SQLite → 提供 REST API 讓 QuoteFlow 桌面 app 同步。

環境變數（在 Railway 的 Variables 頁面設定）：
  LINE_CHANNEL_SECRET       — LINE Messaging API Channel Secret
  LINE_CHANNEL_ACCESS_TOKEN — LINE Channel Access Token
  API_SECRET                — 自訂字串，QuoteFlow 連線時用來驗證身份
  PORT                      — Railway 自動注入，不需手動設定
"""
import base64, hashlib, hmac, json, os, sqlite3
from pathlib import Path

from flask import Flask, abort, jsonify, request

app = Flask(__name__)

CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
CHANNEL_TOKEN  = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
API_SECRET     = os.environ.get("API_SECRET", "change-me-before-deploy")

DB_PATH = Path("line_inquiries.db")


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
            created_at      TEXT    DEFAULT (datetime('now','localtime')),
            updated_at      TEXT    DEFAULT (datetime('now','localtime'))
        );
    """)
    c.commit()
    c.close()


_init_db()


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


def _reply(reply_token: str, text: str):
    import urllib.request
    data = json.dumps({
        "replyToken": reply_token,
        "messages": [{"type": "text", "text": text}],
    }).encode()
    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/reply",
        data=data,
        headers={
            "Authorization": f"Bearer {CHANNEL_TOKEN}",
            "Content-Type": "application/json",
        },
    )
    try:
        urllib.request.urlopen(req, timeout=5)
    except Exception:
        pass


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
        msg = event.get("message", {})
        if msg.get("type") != "text":
            continue

        user_id      = event["source"]["userId"]
        text         = msg["text"]
        reply_token  = event["replyToken"]
        display_name = _get_display_name(user_id)

        c = _conn()
        c.execute(
            "INSERT INTO line_inquiries (line_user_id, display_name, message) VALUES (?,?,?)",
            (user_id, display_name, text),
        )
        c.commit()
        c.close()

        _reply(reply_token,
               f"您好 {display_name}，感謝您的詢問！\n"
               "我們的客服人員將盡快確認並與您聯繫，請稍候。")

    return "OK", 200


# ── REST API（供 QuoteFlow 桌面 app 呼叫）────────────────────

def _auth():
    if request.headers.get("X-API-Secret") != API_SECRET:
        abort(401)


@app.route("/api/inquiries", methods=["GET"])
def list_inquiries():
    """取得詢問列表。?status=待處理|已建卡|已忽略|全部"""
    _auth()
    status = request.args.get("status", "待處理")
    c = _conn()
    if status == "全部":
        rows = c.execute(
            "SELECT * FROM line_inquiries ORDER BY created_at DESC"
        ).fetchall()
    else:
        rows = c.execute(
            "SELECT * FROM line_inquiries WHERE status=? ORDER BY created_at DESC",
            (status,),
        ).fetchall()
    c.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/inquiries/<int:inquiry_id>", methods=["PATCH"])
def update_inquiry(inquiry_id: int):
    """更新詢問狀態（建卡或忽略後由桌面 app 呼叫）"""
    _auth()
    data = request.json or {}
    c = _conn()
    c.execute(
        """UPDATE line_inquiries
           SET status          = COALESCE(?, status),
               inquiry_type    = COALESCE(?, inquiry_type),
               trello_card_id  = COALESCE(?, trello_card_id),
               updated_at      = datetime('now','localtime')
           WHERE id = ?""",
        (data.get("status"), data.get("inquiry_type"),
         data.get("trello_card_id"), inquiry_id),
    )
    c.commit()
    c.close()
    return jsonify({"ok": True})


# ── 健康檢查（Railway 部署用）────────────────────────────────

@app.route("/health")
def health():
    return "ok", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
