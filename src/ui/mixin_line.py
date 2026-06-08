"""
mixin_line.py — LINE 顧客詢問頁籤

顯示透過 LINE 官方帳號傳入的顧客詢問，讓客服確認後手動決定是否在 Trello 建立卡片。
"""
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from ui.app_core import _mk_lf, _ctk_btn


# ── 顏色常數 ────────────────────────────────────────────────
_STATUS_COLOR = {
    "待處理": "#e67e22",
    "已建卡": "#1e8449",
    "已忽略": "#7f8c8d",
}


class _LineTab:

    def _build_tab_line(self, parent, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        # ══════════════════════════════════════════════════════
        #  伺服器設定
        # ══════════════════════════════════════════════════════
        srv_outer, srv_lf = _mk_lf(parent, "LINE 伺服器設定", BG, FONTB)
        srv_outer.pack(fill="x", padx=12, pady=(12, 4))
        srv_lf.columnconfigure(1, weight=1)

        line_cfg = self._config.get("line_server", {})

        ctk.CTkLabel(srv_lf, text="伺服器網址：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        srv_url_var = tk.StringVar(value=line_cfg.get("url", ""))
        ctk.CTkEntry(srv_lf, textvariable=srv_url_var, font=FONT_S,
                      placeholder_text="https://your-app.railway.app",
                      corner_radius=4, border_width=1
                      ).grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=4)

        ctk.CTkLabel(srv_lf, text="API Secret：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        srv_secret_var = tk.StringVar(value=line_cfg.get("secret", ""))
        ctk.CTkEntry(srv_lf, textvariable=srv_secret_var, font=FONT_S,
                      show="*", corner_radius=4, border_width=1
                      ).grid(row=1, column=1, sticky="ew", padx=(0, 4), pady=4)

        def _save_srv_cfg():
            self._config.setdefault("line_server", {})
            self._config["line_server"]["url"]    = srv_url_var.get().strip().rstrip("/")
            self._config["line_server"]["secret"] = srv_secret_var.get().strip()
            self._save_config(self._config)
            status_bar.configure(text="✔  伺服器設定已儲存", text_color="#1e8449")

        _ctk_btn(srv_lf, text="儲存設定", command=_save_srv_cfg,
                 width=90, height=28
                 ).grid(row=0, column=2, rowspan=2, padx=(0, 8), pady=4)

        # ══════════════════════════════════════════════════════
        #  篩選列
        # ══════════════════════════════════════════════════════
        filter_outer, filter_lf = _mk_lf(parent, "篩選", BG, FONTB)
        filter_outer.pack(fill="x", padx=12, pady=(12, 4))

        row = ctk.CTkFrame(filter_lf, fg_color="transparent", corner_radius=0)
        row.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(row, text="狀態：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")

        status_var = tk.StringVar(value="待處理")
        status_menu = ctk.CTkOptionMenu(
            row, variable=status_var,
            values=["全部", "待處理", "已建卡", "已忽略"],
            font=FONT_S, width=100, height=28, corner_radius=4,
            command=lambda _: _refresh(),
        )
        status_menu.pack(side="left", padx=(0, 12))

        _ctk_btn(row, text="🔄  重新整理", command=lambda: _refresh(),
                 width=110, height=28).pack(side="left")

        count_lbl = ctk.CTkLabel(row, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY)
        count_lbl.pack(side="left", padx=(16, 0))

        # ══════════════════════════════════════════════════════
        #  主體：左側列表 + 右側詳情
        # ══════════════════════════════════════════════════════
        body = tk.Frame(parent, bg=BG)
        body.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # ── 左：詢問列表 ──────────────────────────────────────
        list_outer, list_lf = _mk_lf(body, "詢問列表", BG, FONTB)
        list_outer.pack(side="left", fill="both", expand=True, padx=(0, 6))

        cols = ("created_at", "display_name", "inquiry_type", "status", "message")
        tree = ttk.Treeview(list_lf, columns=cols, show="headings",
                             selectmode="browse")
        tree.heading("created_at",   text="時間")
        tree.heading("display_name", text="顧客名稱")
        tree.heading("inquiry_type", text="類型")
        tree.heading("status",       text="狀態")
        tree.heading("message",      text="訊息摘要")
        tree.column("created_at",   width=130, anchor="center")
        tree.column("display_name", width=100, anchor="w")
        tree.column("inquiry_type", width=70,  anchor="center")
        tree.column("status",       width=60,  anchor="center")
        tree.column("message",      width=200, anchor="w")

        sb = ctk.CTkScrollbar(list_lf, orientation="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=4, pady=4)

        # ── 右：詳情 + 操作 ───────────────────────────────────
        detail_outer, detail_lf = _mk_lf(body, "詳細內容", BG, FONTB)
        detail_outer.pack(side="left", fill="y", ipadx=4)
        detail_outer.configure(width=280)
        detail_outer.pack_propagate(False)

        _detail_labels: dict[str, ctk.CTkLabel] = {}

        def _detail_row(key: str, label: str, row_idx: int):
            ctk.CTkLabel(detail_lf, text=label + "：", fg_color="transparent",
                          font=FONT_S, text_color=GRAY,
                          anchor="w").grid(row=row_idx, column=0, sticky="nw",
                                           padx=(8, 2), pady=4)
            lbl = ctk.CTkLabel(detail_lf, text="—", fg_color="transparent",
                                font=FONT_S, anchor="nw", wraplength=160,
                                justify="left")
            lbl.grid(row=row_idx, column=1, sticky="nw", padx=(0, 8), pady=4)
            _detail_labels[key] = lbl

        detail_lf.columnconfigure(1, weight=1)
        _detail_row("display_name", "顧客",   0)
        _detail_row("created_at",   "時間",   1)
        _detail_row("inquiry_type", "類型",   2)
        _detail_row("status",       "狀態",   3)

        # 訊息完整內容
        ctk.CTkLabel(detail_lf, text="訊息內容：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=4, column=0, columnspan=2,
                                       sticky="w", padx=8, pady=(8, 2))
        msg_box = ctk.CTkTextbox(detail_lf, font=FONT_S, height=120,
                                  corner_radius=4, border_width=1,
                                  state="disabled", wrap="word")
        msg_box.grid(row=5, column=0, columnspan=2,
                     sticky="ew", padx=8, pady=(0, 8))

        # 類型選擇（客服可手動修改）
        ctk.CTkLabel(detail_lf, text="調整類型：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=6, column=0, columnspan=2,
                                       sticky="w", padx=8, pady=(4, 0))
        type_var = tk.StringVar(value="未分類")
        type_menu = ctk.CTkOptionMenu(
            detail_lf, variable=type_var,
            values=["未分類", "新產品詢問", "維修需求", "其他"],
            font=FONT_S, width=140, height=28, corner_radius=4,
        )
        type_menu.grid(row=7, column=0, columnspan=2,
                       sticky="w", padx=8, pady=(0, 12))

        # 操作按鈕
        btn_frame = ctk.CTkFrame(detail_lf, fg_color="transparent", corner_radius=0)
        btn_frame.grid(row=8, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))

        btn_create = _ctk_btn(
            btn_frame, text="✔  建立 Trello 卡片",
            fg_color="#117a65", hover_color="#0e6655",
            width=180, height=34,
            command=lambda: _on_create_card(),
        )
        btn_create.pack(fill="x", pady=(0, 6))

        btn_ignore = _ctk_btn(
            btn_frame, text="✕  忽略此詢問",
            fg_color="#922b21", hover_color="#7b241c",
            width=180, height=34,
            command=lambda: _on_ignore(),
        )
        btn_ignore.pack(fill="x")

        # ══════════════════════════════════════════════════════
        #  狀態列（底部）
        # ══════════════════════════════════════════════════════
        status_bar = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY, anchor="w")
        status_bar.pack(anchor="w", padx=14, pady=(0, 4))

        # ══════════════════════════════════════════════════════
        #  內部狀態
        # ══════════════════════════════════════════════════════
        _selected_id: list[int] = [0]   # 用 list 讓 closure 可寫入

        def _get_trello_creds() -> tuple[str, str]:
            tr_cfg = self._config.get("trello", {})
            return tr_cfg.get("api_key", ""), tr_cfg.get("token", "")

        # ══════════════════════════════════════════════════════
        #  資料庫操作
        # ══════════════════════════════════════════════════════
        def _srv_headers() -> dict:
            secret = self._config.get("line_server", {}).get("secret", "")
            return {"X-API-Secret": secret, "Content-Type": "application/json"}

        def _srv_url() -> str:
            return self._config.get("line_server", {}).get("url", "")

        def _sync_from_server(status_filter: str):
            """從雲端拉取最新詢問並 upsert 進本機 DB。"""
            import requests as _req
            url = _srv_url()
            if not url:
                return
            resp = _req.get(f"{url}/api/inquiries",
                            params={"status": status_filter},
                            headers=_srv_headers(), timeout=10)
            resp.raise_for_status()
            rows = resp.json()
            from core.db import get_connection
            conn = get_connection()
            try:
                for r in rows:
                    conn.execute("""
                        INSERT INTO line_inquiries
                            (id, line_user_id, display_name, message,
                             inquiry_type, status, trello_card_id, created_at, updated_at)
                        VALUES (?,?,?,?,?,?,?,?,?)
                        ON CONFLICT(id) DO UPDATE SET
                            display_name   = excluded.display_name,
                            message        = excluded.message,
                            inquiry_type   = excluded.inquiry_type,
                            status         = excluded.status,
                            trello_card_id = excluded.trello_card_id,
                            updated_at     = excluded.updated_at
                    """, (r["id"], r["line_user_id"], r["display_name"], r["message"],
                          r["inquiry_type"], r["status"], r.get("trello_card_id"),
                          r["created_at"], r["updated_at"]))
                conn.commit()
            finally:
                conn.close()

        def _push_status_to_server(inquiry_id: int, status: str,
                                   inquiry_type: str = "", trello_card_id: str = ""):
            """把狀態變更推回雲端（失敗不中斷，靜默忽略）。"""
            import requests as _req
            url = _srv_url()
            if not url:
                return
            try:
                _req.patch(f"{url}/api/inquiries/{inquiry_id}",
                           json={"status": status,
                                 "inquiry_type": inquiry_type or None,
                                 "trello_card_id": trello_card_id or None},
                           headers=_srv_headers(), timeout=8)
            except Exception:
                pass

        def _fetch_inquiries(status_filter: str) -> list[dict]:
            from core.db import get_connection
            conn = get_connection()
            try:
                if status_filter == "全部":
                    rows = conn.execute(
                        "SELECT * FROM line_inquiries ORDER BY created_at DESC"
                    ).fetchall()
                else:
                    rows = conn.execute(
                        "SELECT * FROM line_inquiries WHERE status=? ORDER BY created_at DESC",
                        (status_filter,)
                    ).fetchall()
                return [dict(r) for r in rows]
            finally:
                conn.close()

        def _update_status(inquiry_id: int, new_status: str,
                           trello_card_id: str = ""):
            from core.db import get_connection
            conn = get_connection()
            try:
                conn.execute(
                    """UPDATE line_inquiries
                       SET status=?, trello_card_id=?,
                           updated_at=datetime('now','localtime')
                       WHERE id=?""",
                    (new_status, trello_card_id, inquiry_id),
                )
                conn.commit()
            finally:
                conn.close()

        def _update_type(inquiry_id: int, new_type: str):
            from core.db import get_connection
            conn = get_connection()
            try:
                conn.execute(
                    """UPDATE line_inquiries
                       SET inquiry_type=?,
                           updated_at=datetime('now','localtime')
                       WHERE id=?""",
                    (new_type, inquiry_id),
                )
                conn.commit()
            finally:
                conn.close()

        # ══════════════════════════════════════════════════════
        #  UI 更新
        # ══════════════════════════════════════════════════════
        def _refresh():
            status_bar.configure(text="⏳  正在從伺服器同步…", text_color="#e67e22")
            parent.update_idletasks()
            try:
                _sync_from_server(status_var.get())
            except Exception as e:
                status_bar.configure(text=f"⚠  同步失敗（{e}），顯示本機資料",
                                     text_color="#e67e22")
            else:
                status_bar.configure(text="", text_color=GRAY)
            rows = _fetch_inquiries(status_var.get())
            tree.delete(*tree.get_children())
            for r in rows:
                preview = r["message"][:40].replace("\n", " ")
                if len(r["message"]) > 40:
                    preview += "…"
                tree.insert("", "end",
                            values=(r["created_at"][:16],
                                    r["display_name"],
                                    r["inquiry_type"],
                                    r["status"],
                                    preview),
                            tags=(r["id"],))
            n = len(rows)
            count_lbl.configure(text=f"共 {n} 筆" if n else "無資料")
            _clear_detail()

        def _clear_detail():
            for lbl in _detail_labels.values():
                lbl.configure(text="—")
            msg_box.configure(state="normal")
            msg_box.delete("1.0", "end")
            msg_box.configure(state="disabled")
            _selected_id[0] = 0
            btn_create.configure(state="disabled")
            btn_ignore.configure(state="disabled")

        def _show_detail(row_data: dict):
            _selected_id[0] = row_data["id"]
            _detail_labels["display_name"].configure(text=row_data["display_name"])
            _detail_labels["created_at"].configure(text=row_data["created_at"][:16])
            _detail_labels["inquiry_type"].configure(text=row_data["inquiry_type"])

            status = row_data["status"]
            color  = _STATUS_COLOR.get(status, GRAY)
            _detail_labels["status"].configure(text=status, text_color=color)

            type_var.set(row_data["inquiry_type"])

            msg_box.configure(state="normal")
            msg_box.delete("1.0", "end")
            msg_box.insert("1.0", row_data["message"])
            msg_box.configure(state="disabled")

            is_pending = (status == "待處理")
            btn_create.configure(state="normal" if is_pending else "disabled")
            btn_ignore.configure(state="normal" if is_pending else "disabled")

        def _on_select(event=None):
            sel = tree.selection()
            if not sel:
                return
            inquiry_id = int(tree.item(sel[0])["tags"][0])
            from core.db import get_connection
            conn = get_connection()
            try:
                row = conn.execute(
                    "SELECT * FROM line_inquiries WHERE id=?", (inquiry_id,)
                ).fetchone()
            finally:
                conn.close()
            if row:
                _show_detail(dict(row))

        tree.bind("<<TreeviewSelect>>", _on_select)

        # ══════════════════════════════════════════════════════
        #  建立 Trello 卡片
        # ══════════════════════════════════════════════════════
        def _on_create_card():
            inquiry_id = _selected_id[0]
            if not inquiry_id:
                return

            # 先儲存客服可能調整的類型
            _update_type(inquiry_id, type_var.get())

            api_key, token = _get_trello_creds()
            if not api_key or not token:
                messagebox.showwarning(
                    "尚未設定 Trello 憑證",
                    "請先至「出貨一覽表」頁籤填入 Trello API Key 與 Token 並儲存。",
                    parent=self,
                )
                return

            # 讀取最新資料
            from core.db import get_connection
            conn = get_connection()
            try:
                row = dict(conn.execute(
                    "SELECT * FROM line_inquiries WHERE id=?", (inquiry_id,)
                ).fetchone())
            finally:
                conn.close()

            card_title = f"[LINE] {row['display_name']} — {row['inquiry_type']}"
            card_desc  = (
                f"顧客：{row['display_name']}\n"
                f"類型：{row['inquiry_type']}\n"
                f"時間：{row['created_at']}\n\n"
                f"訊息內容：\n{row['message']}"
            )

            status_bar.configure(text="⏳  正在建立 Trello 卡片…", text_color="#e67e22")
            btn_create.configure(state="disabled")

            def _do_create():
                try:
                    from sync.creator_trello import create_cards
                    cards = [{"title": card_title, "desc": card_desc, "notes": ""}]
                    create_cards(cards, api_key, token)
                    _update_status(inquiry_id, "已建卡")
                    _push_status_to_server(inquiry_id, "已建卡",
                                           inquiry_type=type_var.get())
                    self.after(0, lambda: _on_create_success())
                except Exception as e:
                    self.after(0, lambda err=e: _on_create_error(err))

            def _on_create_success():
                status_bar.configure(text="✔  Trello 卡片建立成功！", text_color="#1e8449")
                _refresh()

            def _on_create_error(err):
                status_bar.configure(text=f"✕  建立失敗：{err}", text_color="#c0392b")
                btn_create.configure(state="normal")
                messagebox.showerror("建立失敗", str(err), parent=self)

            threading.Thread(target=_do_create, daemon=True).start()

        # ══════════════════════════════════════════════════════
        #  忽略詢問
        # ══════════════════════════════════════════════════════
        def _on_ignore():
            inquiry_id = _selected_id[0]
            if not inquiry_id:
                return
            if not messagebox.askyesno("確認忽略",
                                        "確定要忽略這筆詢問？",
                                        parent=self):
                return
            _update_status(inquiry_id, "已忽略")
            _push_status_to_server(inquiry_id, "已忽略")
            status_bar.configure(text="已忽略此筆詢問", text_color=GRAY)
            _refresh()

        # 初始狀態：按鈕停用，等待選取
        btn_create.configure(state="disabled")
        btn_ignore.configure(state="disabled")

        # 首次載入
        _refresh()
