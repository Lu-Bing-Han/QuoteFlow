"""
mixin_line.py — LINE 顧客詢問頁籤

三欄設計：
  左  — 顧客列表（每人一張卡片）
  中  — 對話串（點選顧客後顯示所有訊息氣泡）
  右  — 選取訊息的資料表單 + 建卡操作
"""
import threading
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from ui.app_core import _mk_lf, _ctk_btn


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
                      font=FONT_S, text_color=GRAY, anchor="w"
                      ).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        srv_url_var = tk.StringVar(value=line_cfg.get("url", ""))
        ctk.CTkEntry(srv_lf, textvariable=srv_url_var, font=FONT_S,
                      placeholder_text="https://your-app.railway.app",
                      corner_radius=4, border_width=1
                      ).grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=4)

        ctk.CTkLabel(srv_lf, text="API Secret：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY, anchor="w"
                      ).grid(row=1, column=0, sticky="w", padx=8, pady=4)
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
        filter_outer.pack(fill="x", padx=12, pady=(4, 4))

        frow = ctk.CTkFrame(filter_lf, fg_color="transparent", corner_radius=0)
        frow.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frow, text="狀態：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        status_var = tk.StringVar(value="待處理")
        ctk.CTkOptionMenu(
            frow, variable=status_var,
            values=["全部", "待處理", "已建卡", "已忽略"],
            font=FONT_S, width=100, height=28, corner_radius=4,
            command=lambda _: _refresh(),
        ).pack(side="left", padx=(0, 12))

        _ctk_btn(frow, text="🔄  重新整理", command=lambda: _refresh(),
                 width=110, height=28).pack(side="left")
        _ctk_btn(frow, text="🗑  清空列表",
                 fg_color="#7f8c8d", hover_color="#626d72",
                 width=100, height=28,
                 command=lambda: _on_clear()).pack(side="left", padx=(8, 0))

        count_lbl = ctk.CTkLabel(frow, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY)
        count_lbl.pack(side="left", padx=(16, 0))

        # ══════════════════════════════════════════════════════
        #  三欄主體
        # ══════════════════════════════════════════════════════
        body = tk.Frame(parent, bg=BG)
        body.pack(fill="both", expand=True, padx=12, pady=(0, 4))

        # ── 欄一：顧客列表（固定寬度）────────────────────────
        col_list = tk.Frame(body, bg=BG, width=200)
        col_list.pack(side="left", fill="y")
        col_list.pack_propagate(False)

        list_outer, list_lf = _mk_lf(col_list, "詢問列表", BG, FONTB)
        list_outer.pack(fill="both", expand=True)

        cards_scroll = ctk.CTkScrollableFrame(list_lf, fg_color="#f7f7f7",
                                               corner_radius=0)
        cards_scroll.pack(fill="both", expand=True)

        # ── 欄二：對話串（彈性寬度）──────────────────────────
        col_thread = tk.Frame(body, bg=BG)
        col_thread.pack(side="left", fill="both", expand=True, padx=6)

        thread_outer, thread_lf = _mk_lf(col_thread, "對話紀錄", BG, FONTB)
        thread_outer.pack(fill="both", expand=True)

        # 頂部顧客標題列
        thread_header = tk.Frame(thread_lf, bg="#f0f0f0")
        thread_header.pack(fill="x", padx=0, pady=0)
        thread_name_lbl = tk.Label(
            thread_header, text="← 點選左側顧客", bg="#f0f0f0",
            font=("Microsoft JhengHei UI", 10, "bold"), fg="#636e72",
            anchor="w", padx=12, pady=6,
        )
        thread_name_lbl.pack(side="left", fill="x", expand=True)
        tk.Frame(thread_lf, height=1, bg="#dddddd").pack(fill="x")

        # 對話氣泡捲動區
        thread_scroll = ctk.CTkScrollableFrame(thread_lf, fg_color="#f5f5f5",
                                                corner_radius=0)
        thread_scroll.pack(fill="both", expand=True)

        # ── 欄三：資料表單（固定寬度）────────────────────────
        col_detail = tk.Frame(body, bg=BG, width=280)
        col_detail.pack(side="left", fill="y")
        col_detail.pack_propagate(False)

        detail_outer, detail_lf = _mk_lf(col_detail, "詳細內容", BG, FONTB)
        detail_outer.pack(fill="both", expand=True)

        # ── 右側：唯讀訊息資訊 ────────────────────────────────
        _detail_labels: dict[str, ctk.CTkLabel] = {}
        info_frame = ctk.CTkFrame(detail_lf, fg_color="transparent", corner_radius=0)
        info_frame.pack(fill="x", padx=8, pady=(4, 0))
        info_frame.columnconfigure(1, weight=1)

        def _info_row(key: str, label: str, r: int):
            ctk.CTkLabel(info_frame, text=label + "：", fg_color="transparent",
                          font=FONT_S, text_color=GRAY, anchor="w"
                          ).grid(row=r, column=0, sticky="w", pady=1)
            lbl = ctk.CTkLabel(info_frame, text="—", fg_color="transparent",
                                font=FONT_S, anchor="w")
            lbl.grid(row=r, column=1, sticky="w", padx=(4, 0), pady=1)
            _detail_labels[key] = lbl

        _info_row("display_name", "LINE 名稱", 0)
        _info_row("created_at",   "時間",       1)
        _info_row("status",       "狀態",       2)

        ctk.CTkLabel(detail_lf, text="Gemini 辨識結果（可手動補填）：",
                      fg_color="transparent", font=FONT_S, text_color=GRAY,
                      anchor="w").pack(anchor="w", padx=8, pady=(6, 1))

        # 操作按鈕（固定底部）
        btn_frame = ctk.CTkFrame(detail_lf, fg_color="transparent", corner_radius=0)
        btn_frame.pack(side="bottom", fill="x", padx=8, pady=(2, 8))

        # 類型 + 人員（固定底部）
        bottom_frame = ctk.CTkFrame(detail_lf, fg_color="transparent", corner_radius=0)
        bottom_frame.pack(side="bottom", fill="x", padx=8, pady=(4, 2))
        bottom_frame.columnconfigure(1, weight=1)
        bottom_frame.columnconfigure(3, weight=1)

        ctk.CTkLabel(bottom_frame, text="類型：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).grid(row=0, column=0, sticky="w", pady=3)
        type_var = tk.StringVar(value="未分類")
        ctk.CTkOptionMenu(bottom_frame, variable=type_var,
                           values=["未分類", "新產品詢問", "維修需求", "其他"],
                           font=FONT_S, width=110, height=26, corner_radius=4
                           ).grid(row=0, column=1, sticky="w", padx=(2, 8), pady=3)

        ctk.CTkLabel(bottom_frame, text="人員：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).grid(row=0, column=2, sticky="w", pady=3)
        operators = self._config.get("operators", [""])
        op_codes  = self._config.get("operator_codes", {})
        op_var = tk.StringVar(value=operators[0] if operators else "")
        ctk.CTkOptionMenu(bottom_frame, variable=op_var,
                           values=operators if operators else [""],
                           font=FONT_S, width=80, height=26, corner_radius=4
                           ).grid(row=0, column=3, sticky="w", padx=(2, 0), pady=3)

        # 表單欄位
        scroll_frame = ctk.CTkScrollableFrame(detail_lf, fg_color="transparent",
                                               corner_radius=0)
        scroll_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        scroll_frame.columnconfigure(1, weight=1)

        _field_vars: dict[str, tk.StringVar] = {}

        def _form_row(key: str, label: str, r: int):
            ctk.CTkLabel(scroll_frame, text=label + "：", fg_color="transparent",
                          font=FONT_S, text_color=GRAY, anchor="w"
                          ).grid(row=r, column=0, sticky="w", padx=(4, 2), pady=2)
            var = tk.StringVar()
            ctk.CTkEntry(scroll_frame, textvariable=var, font=FONT_S,
                          height=26, corner_radius=4, border_width=1
                          ).grid(row=r, column=1, sticky="ew", padx=(0, 4), pady=2)
            _field_vars[key] = var

        _form_row("company_name",    "公司名稱", 0)
        _form_row("area",            "市區區域", 1)
        _form_row("contact_name",    "聯絡人",   2)
        _form_row("tax_id",          "統一編號", 3)
        _form_row("mobile",          "手機",     4)
        _form_row("phone",           "電話",     5)
        _form_row("fax",             "FAX",      6)
        _form_row("address",         "地址",     7)
        _form_row("email",           "Mail",     8)
        _form_row("inquiry_product", "詢價商品", 9)

        btn_create = _ctk_btn(
            btn_frame, text="✔  建立卡片",
            fg_color="#117a65", hover_color="#0e6655",
            height=34, command=lambda: _on_create_card(),
        )
        btn_create.pack(side="left", fill="x", expand=True, padx=(0, 4))

        btn_ignore = _ctk_btn(
            btn_frame, text="✕  忽略",
            fg_color="#922b21", hover_color="#7b241c",
            height=34, command=lambda: _on_ignore(),
        )
        btn_ignore.pack(side="left", fill="x", expand=True)

        # ══════════════════════════════════════════════════════
        #  狀態列
        # ══════════════════════════════════════════════════════
        status_bar = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY, anchor="w")
        status_bar.pack(anchor="w", padx=14, pady=(0, 4))

        # ══════════════════════════════════════════════════════
        #  內部狀態
        # ══════════════════════════════════════════════════════
        _card_frames: dict[str, tk.Frame] = {}   # user_id → card widget
        _selected_user: list[str] = [""]
        _selected_id:   list[int] = [0]
        _selected_bubble: list    = [None]
        _thread_bubbles:  list    = []

        _AVATAR_COLORS = ["#2e86c1","#117a65","#8e44ad","#d35400",
                          "#c0392b","#1a5276","#0e6655","#784212"]

        def _avatar_color(name: str) -> str:
            return _AVATAR_COLORS[sum(ord(c) for c in (name or "?")) % len(_AVATAR_COLORS)]

        def _set_bg_all(widget, color: str):
            try: widget.configure(bg=color)
            except Exception: pass
            for child in widget.winfo_children():
                _set_bg_all(child, color)

        # ── 左欄：顧客卡片 ────────────────────────────────────
        def _build_card(user_id: str, msgs: list[dict]) -> tk.Frame:
            latest  = msgs[0]
            name    = latest.get("display_name") or "未知顧客"
            t       = (latest.get("created_at") or "")[:16]
            msg     = (latest.get("message") or "").replace("\n", " ")
            preview = (msg[:28] + "…") if len(msg) > 28 else msg
            pending = sum(1 for m in msgs if m.get("status") == "待處理")
            av_color = _avatar_color(name)
            CBG = "#ffffff"

            card = tk.Frame(cards_scroll, bg=CBG, cursor="hand2")
            card.pack(fill="x")

            SZ = 40
            av = tk.Canvas(card, width=SZ, height=SZ, bg=CBG, highlightthickness=0)
            av.pack(side="left", padx=(10, 6), pady=8)
            av.create_oval(2, 2, SZ-2, SZ-2, fill=av_color, outline="")
            av.create_text(SZ//2, SZ//2,
                           text=(name[0].upper() if name else "?"),
                           fill="white",
                           font=("Microsoft JhengHei UI", 13, "bold"))

            txt = tk.Frame(card, bg=CBG)
            txt.pack(side="left", fill="both", expand=True, pady=8)

            r1 = tk.Frame(txt, bg=CBG)
            r1.pack(fill="x")
            tk.Label(r1, text=name, bg=CBG,
                     font=("Microsoft JhengHei UI", 9, "bold"),
                     fg="#111111", anchor="w").pack(side="left")
            tk.Label(r1, text=t[5:] if len(t) > 5 else t, bg=CBG,
                     font=("Microsoft JhengHei UI", 7),
                     fg="#b2bec3").pack(side="right", padx=(0, 8))

            r2 = tk.Frame(txt, bg=CBG)
            r2.pack(fill="x", pady=(2, 0))
            tk.Label(r2, text=preview, bg=CBG,
                     font=("Microsoft JhengHei UI", 8),
                     fg="#636e72", anchor="w").pack(side="left", fill="x", expand=True)
            if pending > 0:
                tk.Label(r2, text=str(pending), bg="#e67e22", fg="white",
                         font=("Microsoft JhengHei UI", 7, "bold"),
                         padx=4, pady=1).pack(side="right", padx=(0, 8))

            tk.Frame(card, height=1, bg="#eeeeee").pack(fill="x", side="bottom")
            return card

        def _bind_card_events(card: tk.Frame, user_id: str):
            def on_click(e, uid=user_id): _on_card_click(uid)
            def on_enter(e, c=card, uid=user_id):
                if uid != _selected_user[0]: _set_bg_all(c, "#f0faf5")
            def on_leave(e, c=card, uid=user_id):
                if uid != _selected_user[0]: _set_bg_all(c, "#ffffff")
            def _bind_w(w):
                w.bind("<Button-1>", on_click)
                w.bind("<Enter>",    on_enter)
                w.bind("<Leave>",    on_leave)
                for child in w.winfo_children(): _bind_w(child)
            _bind_w(card)

        def _on_card_click(user_id: str):
            for uid, frame in _card_frames.items():
                _set_bg_all(frame, "#ffffff")
            if user_id in _card_frames:
                _set_bg_all(_card_frames[user_id], "#e8f8f0")
            _selected_user[0] = user_id

            from core.db import get_connection
            conn = get_connection()
            try:
                sf = status_var.get()
                if sf == "全部":
                    rows = conn.execute(
                        "SELECT * FROM line_inquiries WHERE line_user_id=? "
                        "ORDER BY created_at ASC", (user_id,)
                    ).fetchall()
                else:
                    rows = conn.execute(
                        "SELECT * FROM line_inquiries WHERE line_user_id=? AND status=? "
                        "ORDER BY created_at ASC", (user_id, sf)
                    ).fetchall()
            finally:
                conn.close()

            msgs = [dict(r) for r in rows]
            name = msgs[0]["display_name"] if msgs else "未知顧客"
            thread_name_lbl.configure(text=name, fg="#111111")
            _show_thread(msgs)

        # ── 中欄：對話串氣泡 ──────────────────────────────────
        def _build_bubble(msg_data: dict):
            mid    = msg_data["id"]
            status = msg_data.get("status", "待處理")
            t      = (msg_data.get("created_at") or "")[:16]
            text   = msg_data.get("message", "")
            stat_color = _STATUS_COLOR.get(status, GRAY)
            BBGN = "#f5f5f5"

            outer = tk.Frame(thread_scroll, bg=BBGN, cursor="hand2")
            outer.pack(fill="x", padx=8, pady=4)

            # 時間 + 狀態
            hdr = tk.Frame(outer, bg=BBGN)
            hdr.pack(fill="x", padx=2)
            tk.Label(hdr, text=t, bg=BBGN,
                     font=("Microsoft JhengHei UI", 7), fg="#b2bec3").pack(side="left")
            tk.Label(hdr, text=status, bg=BBGN,
                     font=("Microsoft JhengHei UI", 7), fg=stat_color).pack(side="right")

            # 訊息氣泡（左對齊，模擬顧客訊息）
            bub_wrap = tk.Frame(outer, bg=BBGN)
            bub_wrap.pack(fill="x", pady=(2, 2))
            bub = tk.Frame(bub_wrap, bg="#ffffff", padx=10, pady=7,
                           relief="flat", bd=0)
            bub.pack(side="left", padx=(4, 40))

            # 使用 tk.Text 支援長訊息自動換行
            txt_widget = tk.Text(bub, font=FONT_S, fg="#2d3436", bg="#ffffff",
                                 wrap="word", relief="flat", borderwidth=0,
                                 highlightthickness=0, cursor="arrow",
                                 state="normal", width=32)
            txt_widget.insert("1.0", text)
            txt_widget.configure(state="disabled")
            txt_widget.pack(fill="x")
            # 動態調整高度
            lines = int(txt_widget.index("end-1c").split(".")[0])
            txt_widget.configure(height=max(1, lines))

            # 底線分隔
            tk.Frame(outer, height=1, bg="#ebebeb").pack(fill="x", pady=(2, 0))

            _thread_bubbles.append(outer)

            def on_click(e, m=mid, b=outer): _on_bubble_click(m, b)
            def on_enter(e, b=outer):
                if _selected_bubble[0] is not b: _set_bg_all(b, "#eaf4fb")
            def on_leave(e, b=outer):
                if _selected_bubble[0] is not b: _set_bg_all(b, BBGN)

            def _bind_w(w):
                w.bind("<Button-1>", on_click)
                w.bind("<Enter>",    on_enter)
                w.bind("<Leave>",    on_leave)
                for child in w.winfo_children(): _bind_w(child)
            _bind_w(outer)

        def _show_thread(msgs: list[dict]):
            for w in list(thread_scroll.winfo_children()): w.destroy()
            _thread_bubbles.clear()
            _selected_bubble[0] = None
            _clear_form()
            for m in msgs:
                _build_bubble(m)
            thread_scroll.after(80, lambda: thread_scroll._parent_canvas.yview_moveto(1.0))

        def _on_bubble_click(msg_id: int, bubble_frame: tk.Frame):
            if _selected_bubble[0] is not None:
                _set_bg_all(_selected_bubble[0], "#f5f5f5")
            _selected_bubble[0] = bubble_frame
            _set_bg_all(bubble_frame, "#d5eaf8")

            from core.db import get_connection
            conn = get_connection()
            try:
                row = conn.execute(
                    "SELECT * FROM line_inquiries WHERE id=?", (msg_id,)
                ).fetchone()
            finally:
                conn.close()
            if row:
                _show_detail(dict(row))

        # ══════════════════════════════════════════════════════
        #  伺服器 / DB 工具
        # ══════════════════════════════════════════════════════
        def _srv_headers() -> dict:
            secret = self._config.get("line_server", {}).get("secret", "")
            return {"X-API-Secret": secret, "Content-Type": "application/json"}

        def _srv_url() -> str:
            return self._config.get("line_server", {}).get("url", "")

        def _sync_from_server(status_filter: str):
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
                struct_fields = [
                    "company_name", "tax_id", "contact_name", "mobile", "phone",
                    "fax", "address", "email", "inquiry_product", "area",
                ]
                for r in rows:
                    conn.execute(f"""
                        INSERT INTO line_inquiries
                            (id, line_user_id, display_name, message,
                             inquiry_type, status, trello_card_id, created_at, updated_at,
                             {", ".join(struct_fields)})
                        VALUES (?,?,?,?,?,?,?,?,?,{",".join("?" * len(struct_fields))})
                        ON CONFLICT(id) DO UPDATE SET
                            display_name   = excluded.display_name,
                            message        = excluded.message,
                            inquiry_type   = excluded.inquiry_type,
                            status         = excluded.status,
                            trello_card_id = excluded.trello_card_id,
                            updated_at     = excluded.updated_at,
                            {", ".join(f"{f} = excluded.{f}" for f in struct_fields)}
                    """, (r["id"], r["line_user_id"], r["display_name"], r["message"],
                          r["inquiry_type"], r["status"], r.get("trello_card_id"),
                          r["created_at"], r["updated_at"],
                          *[r.get(f, "") or "" for f in struct_fields]))
                conn.commit()
            finally:
                conn.close()

        def _push_status_to_server(inquiry_id: int, status: str,
                                   inquiry_type: str = "", trello_card_id: str = "",
                                   extra: dict | None = None):
            import requests as _req
            url = _srv_url()
            if not url:
                return
            payload = {"status": status,
                       "inquiry_type": inquiry_type or None,
                       "trello_card_id": trello_card_id or None}
            if extra:
                payload.update({k: v or None for k, v in extra.items()})
            try:
                _req.patch(f"{url}/api/inquiries/{inquiry_id}",
                           json=payload, headers=_srv_headers(), timeout=8)
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

        def _update_status(inquiry_id: int, new_status: str, trello_card_id: str = ""):
            from core.db import get_connection
            conn = get_connection()
            try:
                conn.execute(
                    "UPDATE line_inquiries SET status=?, trello_card_id=?, "
                    "updated_at=datetime('now','localtime') WHERE id=?",
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
                    "UPDATE line_inquiries SET inquiry_type=?, "
                    "updated_at=datetime('now','localtime') WHERE id=?",
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

            from collections import OrderedDict
            grouped: OrderedDict[str, list[dict]] = OrderedDict()
            for r in rows:
                uid = r.get("line_user_id") or "__unknown__"
                grouped.setdefault(uid, []).append(r)

            for w in list(cards_scroll.winfo_children()): w.destroy()
            _card_frames.clear()

            for uid, msgs in grouped.items():
                card = _build_card(uid, msgs)
                _card_frames[uid] = card
                _bind_card_events(card, uid)

            n_users = len(grouped)
            n_msgs  = len(rows)
            count_lbl.configure(
                text=f"{n_users} 位顧客，{n_msgs} 則" if n_users else "無資料"
            )
            _clear_detail()

        def _clear_form():
            for lbl in _detail_labels.values():
                lbl.configure(text="—")
            for var in _field_vars.values():
                var.set("")
            _selected_id[0] = 0
            btn_create.configure(state="disabled")
            btn_ignore.configure(state="disabled")

        def _clear_detail():
            for frame in _card_frames.values():
                _set_bg_all(frame, "#ffffff")
            _selected_user[0] = ""
            for w in list(thread_scroll.winfo_children()): w.destroy()
            _thread_bubbles.clear()
            _selected_bubble[0] = None
            thread_name_lbl.configure(text="← 點選左側顧客", fg="#636e72")
            _clear_form()

        def _show_detail(row_data: dict):
            _selected_id[0] = row_data["id"]
            _detail_labels["display_name"].configure(text=row_data["display_name"])
            _detail_labels["created_at"].configure(text=row_data["created_at"][:16])
            status = row_data["status"]
            _detail_labels["status"].configure(
                text=status, text_color=_STATUS_COLOR.get(status, GRAY)
            )
            type_var.set(row_data.get("inquiry_type", "未分類"))
            for key, var in _field_vars.items():
                var.set(row_data.get(key, "") or "")
            is_pending = (status == "待處理")
            btn_create.configure(state="normal" if is_pending else "disabled")
            btn_ignore.configure(state="normal" if is_pending else "disabled")

        # ══════════════════════════════════════════════════════
        #  建立 Trello 卡片
        # ══════════════════════════════════════════════════════
        def _on_create_card():
            inquiry_id = _selected_id[0]
            if not inquiry_id:
                return
            _update_type(inquiry_id, type_var.get())

            tr_cfg = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "")
            token   = tr_cfg.get("token", "")
            if not api_key or not token:
                messagebox.showwarning(
                    "尚未設定 Trello 憑證",
                    "請先至「出貨一覽表」頁籤填入 Trello API Key 與 Token 並儲存。",
                    parent=self,
                )
                return

            from core.db import get_connection
            conn = get_connection()
            try:
                row = dict(conn.execute(
                    "SELECT * FROM line_inquiries WHERE id=?", (inquiry_id,)
                ).fetchone())
            finally:
                conn.close()

            op_name = op_var.get()
            op_code = op_codes.get(op_name, op_name[:1].upper() if op_name else "")
            company  = _field_vars["company_name"].get().strip()
            area     = _field_vars["area"].get().strip()
            contact  = _field_vars["contact_name"].get().strip()
            product  = _field_vars["inquiry_product"].get().strip()

            card_title = (
                f"【{op_code} {company}"
                f"{'(' + area + ')' if area else ''}"
                f"{'-' + contact if contact else ''}"
                f"{' -' + product if product else ''}】"
            )
            card_desc = (
                f"公司名稱: {_field_vars['company_name'].get()}\n"
                f"統一編號: {_field_vars['tax_id'].get()}\n"
                f"聯絡人: {_field_vars['contact_name'].get()}\n"
                f"手機: {_field_vars['mobile'].get()}\n"
                f"電話: {_field_vars['phone'].get()}\n"
                f"FAX: {_field_vars['fax'].get()}\n"
                f"地址: {_field_vars['address'].get()}\n"
                f"Mail: {_field_vars['email'].get()}\n"
                f"客戶詢價來源: LINE官方帳號\n\n"
                f"原始訊息：\n{row['message']}"
            )

            status_bar.configure(text="⏳  正在建立 Trello 卡片…", text_color="#e67e22")
            btn_create.configure(state="disabled")
            _inq_type = type_var.get()

            def _do_create():
                try:
                    from sync.creator_trello import (
                        create_cards, _REPAIR_BOARD, _REPAIR_LIST,
                        _BOARD_NAME, _LIST_NAME,
                    )
                    cards = [{"title": card_title, "desc": card_desc, "notes": ""}]
                    if _inq_type == "維修需求":
                        create_cards(cards, api_key, token,
                                     board_name=_REPAIR_BOARD, list_name=_REPAIR_LIST)
                    else:
                        create_cards(cards, api_key, token,
                                     board_name=_BOARD_NAME, list_name=_LIST_NAME)
                    _update_status(inquiry_id, "已建卡")
                    _push_status_to_server(
                        inquiry_id, "已建卡", inquiry_type=type_var.get(),
                        extra={k: v.get() for k, v in _field_vars.items()},
                    )
                    self.after(0, _on_create_success)
                except Exception as e:
                    self.after(0, lambda err=e: _on_create_error(err))

            def _on_create_success():
                status_bar.configure(text="✔  Trello 卡片建立成功！", text_color="#1e8449")
                cur_user = _selected_user[0]
                _refresh()
                if cur_user and cur_user in _card_frames:
                    _on_card_click(cur_user)

            def _on_create_error(err):
                status_bar.configure(text=f"✕  建立失敗：{err}", text_color="#c0392b")
                btn_create.configure(state="normal")
                messagebox.showerror("建立失敗", str(err), parent=self)

            threading.Thread(target=_do_create, daemon=True).start()

        # ══════════════════════════════════════════════════════
        #  忽略
        # ══════════════════════════════════════════════════════
        def _on_ignore():
            inquiry_id = _selected_id[0]
            if not inquiry_id:
                return
            if not messagebox.askyesno("確認忽略", "確定要忽略這筆詢問？", parent=self):
                return
            _update_status(inquiry_id, "已忽略")
            _push_status_to_server(inquiry_id, "已忽略")
            status_bar.configure(text="已忽略此筆詢問", text_color=GRAY)
            cur_user = _selected_user[0]
            _refresh()
            if cur_user and cur_user in _card_frames:
                _on_card_click(cur_user)

        # ══════════════════════════════════════════════════════
        #  清空
        # ══════════════════════════════════════════════════════
        def _on_clear():
            if not _card_frames:
                return
            from core.db import get_connection
            sf = status_var.get()
            conn = get_connection()
            try:
                if sf == "全部":
                    ids = [r[0] for r in conn.execute(
                        "SELECT id FROM line_inquiries").fetchall()]
                else:
                    ids = [r[0] for r in conn.execute(
                        "SELECT id FROM line_inquiries WHERE status=?", (sf,)).fetchall()]
            finally:
                conn.close()
            if not ids:
                return
            if not messagebox.askyesno(
                "確認清空",
                f"確定要刪除目前顯示的 {len(ids)} 筆詢問紀錄？\n此操作無法復原。",
                parent=self,
            ):
                return
            conn = get_connection()
            try:
                placeholders = ",".join("?" * len(ids))
                conn.execute(
                    f"DELETE FROM line_inquiries WHERE id IN ({placeholders})", ids)
                conn.commit()
            finally:
                conn.close()
            status_bar.configure(text=f"已刪除 {len(ids)} 筆紀錄", text_color=GRAY)
            _refresh()

        btn_create.configure(state="disabled")
        btn_ignore.configure(state="disabled")
        _refresh()
