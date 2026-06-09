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

        _ctk_btn(row, text="🗑  清空列表",
                 fg_color="#7f8c8d", hover_color="#626d72",
                 width=100, height=28,
                 command=lambda: _on_clear()).pack(side="left", padx=(8, 0))

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

        # ── 上方：唯讀基本資訊 ───────────────────────────────
        info_frame = ctk.CTkFrame(detail_lf, fg_color="transparent", corner_radius=0)
        info_frame.pack(fill="x", padx=8, pady=(4, 0))
        info_frame.columnconfigure(1, weight=1)

        def _info_row(key: str, label: str, row: int):
            ctk.CTkLabel(info_frame, text=label + "：", fg_color="transparent",
                          font=FONT_S, text_color=GRAY, anchor="w"
                          ).grid(row=row, column=0, sticky="w", pady=1)
            lbl = ctk.CTkLabel(info_frame, text="—", fg_color="transparent",
                                font=FONT_S, anchor="w")
            lbl.grid(row=row, column=1, sticky="w", padx=(4, 0), pady=1)
            _detail_labels[key] = lbl

        _info_row("display_name", "LINE 名稱", 0)
        _info_row("created_at",   "時間",       1)
        _info_row("status",       "狀態",       2)

        # ── 訊息內容 ─────────────────────────────────────────
        ctk.CTkLabel(detail_lf, text="訊息內容：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY, anchor="w"
                      ).pack(anchor="w", padx=8, pady=(4, 1))
        msg_box = ctk.CTkTextbox(detail_lf, font=FONT_S, height=52,
                                  corner_radius=4, border_width=1,
                                  state="disabled", wrap="word")
        msg_box.pack(fill="x", padx=8, pady=(0, 3))

        btn_history = _ctk_btn(
            detail_lf, text="📋  查看同顧客歷史詢問",
            fg_color=GRAY, hover_color="#4d5d6e",
            width=180, height=26,
            command=lambda: _on_show_history(),
        )
        btn_history.pack(anchor="w", padx=8, pady=(0, 3))

        # ── 可捲動的編輯表單 ─────────────────────────────────
        ctk.CTkLabel(detail_lf, text="Gemini 辨識結果（可手動補填）：",
                      fg_color="transparent", font=FONT_S, text_color=GRAY,
                      anchor="w").pack(anchor="w", padx=8, pady=(1, 1))

        # ── 操作按鈕（固定在底部）────────────────────────────
        btn_frame = ctk.CTkFrame(detail_lf, fg_color="transparent", corner_radius=0)
        btn_frame.pack(side="bottom", fill="x", padx=8, pady=(2, 8))

        # ── 類型 + 人員代號（固定在底部）────────────────────
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
        operators  = self._config.get("operators", [""])
        op_codes   = self._config.get("operator_codes", {})
        op_var = tk.StringVar(value=operators[0] if operators else "")
        ctk.CTkOptionMenu(bottom_frame, variable=op_var,
                           values=operators if operators else [""],
                           font=FONT_S, width=80, height=26, corner_radius=4
                           ).grid(row=0, column=3, sticky="w", padx=(2, 0), pady=3)

        # ── 可捲動的編輯表單（填滿剩餘空間）────────────────
        scroll_frame = ctk.CTkScrollableFrame(detail_lf, fg_color="transparent",
                                               corner_radius=0)
        scroll_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        scroll_frame.columnconfigure(1, weight=1)

        _field_vars: dict[str, tk.StringVar] = {}

        def _form_row(key: str, label: str, row: int):
            ctk.CTkLabel(scroll_frame, text=label + "：", fg_color="transparent",
                          font=FONT_S, text_color=GRAY, anchor="w"
                          ).grid(row=row, column=0, sticky="w", padx=(4, 2), pady=2)
            var = tk.StringVar()
            ctk.CTkEntry(scroll_frame, textvariable=var, font=FONT_S,
                          height=26, corner_radius=4, border_width=1
                          ).grid(row=row, column=1, sticky="ew", padx=(0, 4), pady=2)
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
        #  狀態列（底部）
        # ══════════════════════════════════════════════════════
        status_bar = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY, anchor="w")
        status_bar.pack(anchor="w", padx=14, pady=(0, 4))

        # ══════════════════════════════════════════════════════
        #  內部狀態
        # ══════════════════════════════════════════════════════
        _selected_id: list[int] = [0]   # 用 list 讓 closure 可寫入
        _selected_user: list[str] = [""]

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
            """把狀態變更推回雲端（失敗不中斷，靜默忽略）。"""
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
                           json=payload,
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
            _selected_user[0] = ""
            btn_create.configure(state="disabled")
            btn_ignore.configure(state="disabled")
            btn_history.configure(state="disabled")

        def _show_detail(row_data: dict):
            _selected_id[0] = row_data["id"]
            _selected_user[0] = row_data.get("line_user_id", "")
            btn_history.configure(state="normal")
            _detail_labels["display_name"].configure(text=row_data["display_name"])
            _detail_labels["created_at"].configure(text=row_data["created_at"][:16])

            status = row_data["status"]
            color  = _STATUS_COLOR.get(status, GRAY)
            _detail_labels["status"].configure(text=status, text_color=color)

            type_var.set(row_data.get("inquiry_type", "未分類"))

            msg_box.configure(state="normal")
            msg_box.delete("1.0", "end")
            msg_box.insert("1.0", row_data["message"])
            msg_box.configure(state="disabled")

            # 填入 Gemini 辨識結果
            for key, var in _field_vars.items():
                var.set(row_data.get(key, "") or "")

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

            # 取人員代號
            op_name = op_var.get()
            op_code = op_codes.get(op_name, op_name[:1].upper() if op_name else "")

            # 從表單讀取最新填入值（客服可能手動補填）
            company  = _field_vars["company_name"].get().strip()
            area     = _field_vars["area"].get().strip()
            contact  = _field_vars["contact_name"].get().strip()
            product  = _field_vars["inquiry_product"].get().strip()

            # 標題：【代號 公司名(市區區域)-客戶名 -詢價商品】
            area_part    = f"({area})" if area else ""
            contact_part = f"-{contact}" if contact else ""
            product_part = f" -{product}" if product else ""
            card_title = (
                f"【{op_code} {company}{area_part}{contact_part}{product_part}】"
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
                        inquiry_id, "已建卡",
                        inquiry_type=type_var.get(),
                        extra={k: v.get() for k, v in _field_vars.items()},
                    )
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

        # ══════════════════════════════════════════════════════
        #  清空目前顯示的紀錄
        # ══════════════════════════════════════════════════════
        def _on_clear():
            items = tree.get_children()
            if not items:
                return
            ids = [int(tree.item(i)["tags"][0]) for i in items]
            if not messagebox.askyesno(
                "確認清空",
                f"確定要刪除目前顯示的 {len(ids)} 筆詢問紀錄？\n此操作無法復原。",
                parent=self,
            ):
                return
            from core.db import get_connection
            conn = get_connection()
            try:
                placeholders = ",".join("?" * len(ids))
                conn.execute(
                    f"DELETE FROM line_inquiries WHERE id IN ({placeholders})", ids
                )
                conn.commit()
            finally:
                conn.close()
            status_bar.configure(text=f"已刪除 {len(ids)} 筆紀錄", text_color=GRAY)
            _refresh()

        # ══════════════════════════════════════════════════════
        #  同顧客歷史詢問（資訊分散在多則訊息時，方便人工拼湊）
        # ══════════════════════════════════════════════════════
        _STRUCT_LABELS = [
            ("company_name",    "公司"),
            ("tax_id",          "統編"),
            ("contact_name",    "聯絡人"),
            ("mobile",          "手機"),
            ("phone",           "電話"),
            ("fax",             "FAX"),
            ("address",         "地址"),
            ("email",           "Mail"),
            ("inquiry_product", "詢價商品"),
            ("area",            "區域"),
        ]

        def _show_customer_history(user_id: str, name: str, current_id: int):
            from core.db import get_connection
            conn = get_connection()
            try:
                rows = [dict(r) for r in conn.execute(
                    "SELECT * FROM line_inquiries WHERE line_user_id=? "
                    "ORDER BY created_at ASC",
                    (user_id,)
                ).fetchall()]
            finally:
                conn.close()

            dlg = ctk.CTkToplevel(self)
            dlg.title(f"同顧客歷史詢問 — {name}")
            dlg.configure(fg_color=BG)
            dlg.after(100, dlg.grab_set)
            dlg.geometry("640x540")

            ctk.CTkLabel(
                dlg,
                text=f"共 {len(rows)} 筆歷史詢問（依時間排序，可參考各則內容人工拼湊完整資料）",
                fg_color="transparent", font=FONT_S, text_color=GRAY,
            ).pack(anchor="w", padx=16, pady=(12, 4))

            scroll = ctk.CTkScrollableFrame(dlg, fg_color="transparent", corner_radius=0)
            scroll.pack(fill="both", expand=True, padx=12, pady=(0, 8))

            for r in rows:
                title = f"{r['created_at'][:16]}　[{r['status']}]"
                if r["id"] == current_id:
                    title += "　← 目前檢視這筆"
                card_outer, card_lf = _mk_lf(scroll, title, BG, FONTB)
                card_outer.pack(fill="x", padx=4, pady=4)

                msg_preview = r["message"][:200]
                ctk.CTkLabel(card_lf, text=msg_preview, fg_color="transparent",
                              font=FONT_S, anchor="w", justify="left",
                              wraplength=560
                              ).pack(anchor="w", padx=8, pady=(4, 2))

                bits = [f"{lbl}: {r.get(key)}" for key, lbl in _STRUCT_LABELS if r.get(key)]
                if bits:
                    ctk.CTkLabel(card_lf, text="　|　".join(bits),
                                  fg_color="transparent", font=FONT_S,
                                  text_color="#1e8449", anchor="w", justify="left",
                                  wraplength=560
                                  ).pack(anchor="w", padx=8, pady=(0, 6))
                else:
                    ctk.CTkLabel(card_lf, text="（未辨識出結構化資訊）",
                                  fg_color="transparent", font=FONT_S,
                                  text_color=GRAY, anchor="w"
                                  ).pack(anchor="w", padx=8, pady=(0, 6))

            ctk.CTkButton(dlg, text="關閉", command=dlg.destroy,
                           fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                           font=FONT, width=100, height=34, corner_radius=6
                           ).pack(pady=(0, 10))

        def _on_show_history():
            user_id = _selected_user[0]
            if not user_id:
                return
            name = _detail_labels["display_name"].cget("text")
            _show_customer_history(user_id, name, _selected_id[0])

        # 初始狀態：按鈕停用，等待選取
        btn_create.configure(state="disabled")
        btn_ignore.configure(state="disabled")
        btn_history.configure(state="disabled")

        # 首次載入
        _refresh()
