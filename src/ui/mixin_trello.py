"""
mixin_trello.py — Trello 相關頁籤 mixin（出貨一覽表、生產群組紀錄、建立卡片、下載卡片）
"""
import json
import os
import threading
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from pathlib import Path
from ui.app_core import _mk_lf


def _bind_drag_select(tree: "ttk.Treeview"):
    """為 Treeview 加上滑鼠拖曳範圍選取與 Ctrl+點擊多選。"""
    _anchor = [None]

    def _on_press(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        _anchor[0] = item
        tree.selection_set(item)

    def _on_ctrl_press(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        _anchor[0] = None
        if item in tree.selection():
            tree.selection_remove(item)
        else:
            tree.selection_add(item)
        return "break"

    def _on_drag(event):
        if not _anchor[0]:
            return
        item = tree.identify_row(event.y)
        if not item:
            return
        all_items = tree.get_children()
        try:
            a = all_items.index(_anchor[0])
            b = all_items.index(item)
        except ValueError:
            return
        lo, hi = min(a, b), max(a, b)
        tree.selection_set(all_items[lo:hi + 1])

    def _on_release(event):
        _anchor[0] = None

    tree.bind("<Button-1>",         _on_press)
    tree.bind("<Control-Button-1>", _on_ctrl_press)
    tree.bind("<B1-Motion>",        _on_drag)
    tree.bind("<ButtonRelease-1>",  _on_release)


class _TrelloTab:
    """Mixin providing Trello-related tab builders and callbacks."""

    # ════════════════════════════════════════════════════════
    #  Tab 7：出貨一覽表
    # ════════════════════════════════════════════════════════
    def _build_tab_overview(self, parent, FONT, FONTB, BG):
        from _paths import _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH  # noqa: F401
        GRAY  = "#5d6d7e"
        GREEN = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _all_cards:       list[dict] = []
        _displayed_cards: list[dict] = []

        # ── Trello 憑證 ───────────────────────────────────
        cred_outer, cred_frame = _mk_lf(parent, "Trello 憑證", BG, FONTB)
        cred_outer.pack(fill="x", padx=12, pady=(12, 4))
        cred_frame.columnconfigure(1, weight=1)

        tr_cfg  = self._config.get("trello", {})
        key_var = tk.StringVar(value=tr_cfg.get("api_key", ""))
        tok_var = tk.StringVar(value=tr_cfg.get("token",   ""))

        ctk.CTkLabel(cred_frame, text="API Key：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        key_entry = ctk.CTkEntry(cred_frame, textvariable=key_var, font=FONT_S,
                                  show="*", corner_radius=4, border_width=1)
        key_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=3)

        ctk.CTkLabel(cred_frame, text="Token：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=3)
        tok_entry = ctk.CTkEntry(cred_frame, textvariable=tok_var, font=FONT_S,
                                  show="*", corner_radius=4, border_width=1)
        tok_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=3)

        def _show_hide(entry, btn):
            if entry.cget("show") == "*":
                entry.configure(show=""); btn.configure(text="隱藏")
            else:
                entry.configure(show="*"); btn.configure(text="顯示")

        btn_key = ctk.CTkButton(cred_frame, text="顯示", font=FONT_S,
                                 fg_color=GRAY, hover_color="#4d5d6e",
                                 text_color="white", width=50, height=26, corner_radius=4)
        btn_key.configure(command=lambda: _show_hide(key_entry, btn_key))
        btn_key.grid(row=0, column=2, padx=(0, 8), pady=3)

        btn_tok = ctk.CTkButton(cred_frame, text="顯示", font=FONT_S,
                                 fg_color=GRAY, hover_color="#4d5d6e",
                                 text_color="white", width=50, height=26, corner_radius=4)
        btn_tok.configure(command=lambda: _show_hide(tok_entry, btn_tok))
        btn_tok.grid(row=1, column=2, padx=(0, 8), pady=3)

        def _save_trello_creds():
            self._config.setdefault("trello", {})
            self._config["trello"]["api_key"] = key_var.get().strip()
            self._config["trello"]["token"]   = tok_var.get().strip()
            self._save_config(self._config)
            messagebox.showinfo("已儲存", "Trello 憑證已儲存", parent=parent)

        ctk.CTkButton(cred_frame, text="儲存憑證", command=_save_trello_creds,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT, width=90, height=28, corner_radius=4
                       ).grid(row=2, column=1, sticky="w", padx=8, pady=(4, 6))

        # ── Google Sheets 狀態 ────────────────────────────
        gs_outer, gs_frame = _mk_lf(parent, "Google Sheets", BG, FONTB)
        gs_outer.pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(gs_frame,
                      text="✔  credentials.json 已就緒" if _GSHEETS_CREDS_PATH.exists()
                           else "✘  找不到 credentials.json（請放到 template 資料夾）",
                      fg_color="transparent", font=FONT_S, anchor="w",
                      text_color=GREEN if _GSHEETS_CREDS_PATH.exists() else "#c0392b",
                      ).pack(fill="x", padx=8, pady=6)

        token_status = ctk.CTkLabel(gs_frame,
                                     text="✔  已授權（gsheets_token.json 存在）" if _GSHEETS_TOKEN_PATH.exists()
                                          else "尚未授權，點「同步」時會自動開啟瀏覽器",
                                     fg_color="transparent", font=FONT_S, anchor="w",
                                     text_color=GREEN if _GSHEETS_TOKEN_PATH.exists() else GRAY)
        token_status.pack(fill="x", padx=8, pady=(0, 6))

        # ── 卡片抓取列 ────────────────────────────────────
        _BOARD_SOURCES = {
            "物流事業部1 — 本周下單":  ("物流事業部1",  "本周下單"),
            "維修保養部門 — 已下單":   ("維修保養部門", "已下單"),
        }

        fetch_row = tk.Frame(parent, bg=BG)
        fetch_row.pack(fill="x", padx=12, pady=(6, 2))

        ctk.CTkLabel(fetch_row, text="來源：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(0, 2))
        source_var = tk.StringVar(value=list(_BOARD_SOURCES.keys())[0])
        ctk.CTkOptionMenu(
            fetch_row, variable=source_var,
            values=list(_BOARD_SOURCES.keys()),
            font=FONT_S, width=200, height=28, corner_radius=4,
        ).pack(side="left", padx=(0, 12))

        fetch_status = ctk.CTkLabel(fetch_row, text="", fg_color="transparent",
                                     font=FONT_S, text_color=GRAY)
        fetch_status.pack(side="left", padx=(8, 0))

        ctk.CTkLabel(fetch_row, text="下單日期 從：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(8, 2))
        from tkcalendar import DateEntry
        from datetime import date as _dt_cls
        date_entry = DateEntry(fetch_row, width=10, date_pattern='y/m/d',
                                font=FONT_S, maxdate=_dt_cls.today())
        date_entry.pack(side="left", padx=(0, 2))
        ctk.CTkButton(fetch_row, text="全部",
                       command=lambda: _show_all(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(2, 0))

        def _render_tree(cards: list[dict]):
            _displayed_cards.clear()
            _displayed_cards.extend(cards)
            tree.delete(*tree.get_children())
            for c in cards:
                tree.insert("", "end", values=(
                    c["created_date"],
                    c["company"],
                    c["product"],
                    c["quantity"],
                    c.get("payment_type", ""),
                ))
            tree.selection_set(tree.get_children())
            prev_title_lbl.configure(
                text=f"  本周下單卡片（共 {len(cards)} 張，可多選）  ")

        def _apply_days_filter():
            from_date = date_entry.get_date()
            filtered = [c for c in _all_cards
                        if c.get("created_dt") and c["created_dt"] >= from_date]
            _render_tree(filtered)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 張，篩選後 {len(filtered)} 張",
                text_color=GREEN)

        def _show_all():
            _render_tree(_all_cards)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 張（全部顯示）",
                text_color=GREEN)

        def _fetch():
            from sync.syncer_trello import fetch_po_cards, update_location_cache
            from _paths import _LOCATION_CACHE_PATH
            api_key = key_var.get().strip()
            token   = tok_var.get().strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未填", "請先填入並儲存 Trello 憑證", parent=parent); return

            board_name, list_name = _BOARD_SOURCES[source_var.get()]

            def _worker(bn=board_name, ln=list_name):
                cards = fetch_po_cards(api_key, token, board_name=bn, list_name=ln)
                update_location_cache(cards, _LOCATION_CACHE_PATH)
                return cards

            def _on_done(cards):
                _all_cards.clear(); _all_cards.extend(cards); _apply_days_filter()

            def _on_error(e):
                from sync.syncer_sheets import _fmt_api_error
                fetch_status.configure(text=f"✘  {_fmt_api_error(e)}", text_color="#c0392b")

            self._run_task(_worker,
                            buttons=[fetch_btn],
                            status_label=fetch_status,
                            loading_text="抓取中…",
                            success_text=lambda cards: f"✔  共 {len(cards)} 張",
                            on_success=_on_done,
                            on_error=_on_error)

        fetch_btn = ctk.CTkButton(fetch_row, text="🔄 抓取卡片", command=_fetch,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=90, height=28, corner_radius=4)
        fetch_btn.pack(side="left")
        ctk.CTkButton(fetch_row, text="篩選", command=_apply_days_filter,
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_outer = tk.Frame(parent, bg="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_inner = tk.Frame(prev_outer, bg=BG)
        prev_inner.pack(fill="both", expand=True, padx=1, pady=1)

        prev_title_lbl = ctk.CTkLabel(prev_inner,
                                       text="  本周下單卡片（共 0 張，可多選）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(4, 0))

        tree_frame = tk.Frame(prev_inner, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        cols = ("date", "company", "product", "qty", "payment")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                             selectmode="extended", height=6)
        tree.heading("date",    text="下單日期")
        tree.heading("company", text="客戶名稱")
        tree.heading("product", text="品號 / 產品")
        tree.heading("qty",     text="數量")
        tree.heading("payment", text="付款方式")
        tree.column("date",    width=70,  anchor="center", stretch=False)
        tree.column("company", width=140, anchor="w",      stretch=False)
        tree.column("product", width=260, anchor="w",      stretch=True)
        tree.column("qty",     width=50,  anchor="center", stretch=False)
        tree.column("payment", width=80,  anchor="center", stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # ── 拖曳框選 ──────────────────────────────────────
        _bind_drag_select(tree)

        # ── 全選 / 取消全選 / 刪除 ───────────────────────────
        sel_row = tk.Frame(parent, bg=BG)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))

        def _delete_selected_overview():
            sel_ids = tree.selection()
            if not sel_ids:
                return
            indices = sorted([tree.index(i) for i in sel_ids], reverse=True)
            to_remove = [_displayed_cards[i] for i in indices]
            for card in to_remove:
                if card in _all_cards:
                    _all_cards.remove(card)
                if card in _displayed_cards:
                    _displayed_cards.remove(card)
            for iid in sel_ids:
                tree.delete(iid)
            prev_title_lbl.configure(
                text=f"  本周下單卡片（共 {len(_displayed_cards)} 張，可多選）  ")

        ctk.CTkButton(sel_row, text="🗑 刪除選取",
                       command=_delete_selected_overview,
                       fg_color="#922b21", hover_color="#7b241c", text_color="white",
                       font=FONT_S, width=90, height=26, corner_radius=4
                       ).pack(side="left")

        # ── 狀態列 & 推送按鈕 ─────────────────────────────
        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _push():
            from sync.syncer_sheets import push_cards
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先選取要同步的卡片", parent=parent); return
            if not _GSHEETS_CREDS_PATH.exists():
                messagebox.showerror("缺少憑證", f"找不到 {_GSHEETS_CREDS_PATH}", parent=parent); return

            indices = [tree.index(i) for i in sel_ids]
            selected = [_displayed_cards[i] for i in indices]

            if not messagebox.askyesno("確認同步",
                    f"即將推送 {len(selected)} 筆卡片至 Google Sheets，確定繼續？",
                    parent=parent):
                return

            def _worker():
                return push_cards(selected, _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH)

            def _on_done(added):
                token_status.configure(text="✔  已授權（gsheets_token.json 存在）", text_color=GREEN)
                if messagebox.askyesno("完成", f"新增 {added} 筆資料\n\n是否立即開啟出貨一覽表？", parent=parent):
                    import webbrowser
                    from sync.syncer_sheets import _SPREADSHEET_ID, _SHEET_GID
                    webbrowser.open(f"https://docs.google.com/spreadsheets/d/{_SPREADSHEET_ID}/edit#gid={_SHEET_GID}")

            def _on_error(e):
                from sync.syncer_sheets import _fmt_api_error
                msg = _fmt_api_error(e)
                out_label.configure(text=f"✘  {msg}", text_color="#c0392b")
                messagebox.showerror("同步失敗", msg, parent=parent)

            self._run_task(_worker,
                            buttons=[push_btn],
                            status_label=out_label,
                            loading_text="同步中…",
                            success_text=lambda added: f"✔  同步完成，新增 {added} 筆資料",
                            on_success=_on_done,
                            on_error=_on_error)

        bb = tk.Frame(parent, bg=BG)
        bb.pack(fill="x", padx=12, pady=8)
        push_btn = ctk.CTkButton(bb, text="☁  推送選取的卡片 → Google Sheets", command=_push,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8)
        push_btn.pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab：對帳單
    # ════════════════════════════════════════════════════════
    def _build_tab_statement(self, parent, FONT, FONTB, BG):
        GRAY  = "#5d6d7e"
        GREEN = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _all_cards:       list[dict] = []
        _displayed_cards: list[dict] = []

        # ── 抓取列 ────────────────────────────────────────
        fetch_row = tk.Frame(parent, bg=BG)
        fetch_row.pack(fill="x", padx=12, pady=(12, 2))

        ctk.CTkLabel(fetch_row, text="來源：物流事業部1 — 本周下單",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left", padx=(0, 12))

        fetch_status = ctk.CTkLabel(fetch_row, text="", fg_color="transparent",
                                     font=FONT_S, text_color=GRAY)
        fetch_status.pack(side="left", padx=(8, 0))

        ctk.CTkLabel(fetch_row, text="下單日期 從：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(8, 2))
        from tkcalendar import DateEntry
        from datetime import date as _dt_cls
        date_entry = DateEntry(fetch_row, width=10, date_pattern='y/m/d',
                                font=FONT_S, maxdate=_dt_cls.today())
        date_entry.pack(side="left", padx=(0, 2))
        ctk.CTkButton(fetch_row, text="全部",
                       command=lambda: _show_all(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(2, 0))

        def _render_tree(cards: list[dict]):
            _displayed_cards.clear()
            _displayed_cards.extend(cards)
            tree.delete(*tree.get_children())
            for c in cards:
                tree.insert("", "end", values=(
                    c["created_date"],
                    c["company"],
                    c["product"],
                    c["quantity"],
                    c.get("amount", ""),
                    c.get("payment_raw", ""),
                ))
            tree.selection_set(tree.get_children())
            prev_title_lbl.configure(
                text=f"  本周下單卡片（共 {len(cards)} 張，可多選）  ")

        def _apply_days_filter():
            from_date = date_entry.get_date()
            filtered = [c for c in _all_cards
                        if c.get("created_dt") and c["created_dt"] >= from_date]
            _render_tree(filtered)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 張，篩選後 {len(filtered)} 張",
                text_color=GREEN)

        def _show_all():
            _render_tree(_all_cards)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 張（全部顯示）",
                text_color=GREEN)

        def _fetch():
            from sync.syncer_trello import fetch_po_cards, update_location_cache
            from _paths import _LOCATION_CACHE_PATH
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未填",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent); return

            def _worker():
                cards = fetch_po_cards(api_key, token)
                update_location_cache(cards, _LOCATION_CACHE_PATH)
                return cards

            def _on_done(cards):
                _all_cards.clear(); _all_cards.extend(cards); _apply_days_filter()

            def _on_error(e):
                from sync.syncer_sheets import _fmt_api_error
                fetch_status.configure(text=f"✘  {_fmt_api_error(e)}", text_color="#c0392b")

            self._run_task(_worker,
                            buttons=[fetch_btn],
                            status_label=fetch_status,
                            loading_text="抓取中…",
                            success_text=lambda cards: f"✔  共 {len(cards)} 張",
                            on_success=_on_done,
                            on_error=_on_error)

        fetch_btn = ctk.CTkButton(fetch_row, text="🔄 抓取卡片", command=_fetch,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=90, height=28, corner_radius=4)
        fetch_btn.pack(side="left")
        ctk.CTkButton(fetch_row, text="篩選", command=_apply_days_filter,
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_outer = tk.Frame(parent, bg="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_inner = tk.Frame(prev_outer, bg=BG)
        prev_inner.pack(fill="both", expand=True, padx=1, pady=1)

        prev_title_lbl = ctk.CTkLabel(prev_inner,
                                       text="  本周下單卡片（共 0 張，可多選）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(4, 0))

        tree_frame = tk.Frame(prev_inner, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        cols = ("date", "company", "product", "qty", "amount", "payment")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                             selectmode="extended", height=10)
        tree.heading("date",    text="下單日期")
        tree.heading("company", text="客戶名稱")
        tree.heading("product", text="品號 / 產品")
        tree.heading("qty",     text="數量")
        tree.heading("amount",  text="應收總金額")
        tree.heading("payment", text="付款方式")
        tree.column("date",    width=70,  anchor="center", stretch=False)
        tree.column("company", width=140, anchor="w",      stretch=False)
        tree.column("product", width=240, anchor="w",      stretch=True)
        tree.column("qty",     width=50,  anchor="center", stretch=False)
        tree.column("amount",  width=90,  anchor="e",      stretch=False)
        tree.column("payment", width=120, anchor="w",      stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # ── 拖曳框選 ──────────────────────────────────────
        _bind_drag_select(tree)

        # ── 全選 / 取消全選 ───────────────────────────────
        sel_row = tk.Frame(parent, bg=BG)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=26, corner_radius=4
                       ).pack(side="left")

        # ── 狀態列 & 生成按鈕 ─────────────────────────────
        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _generate():
            from core.generator_statement import generate_statement
            from _paths import _LOCATION_CACHE_PATH
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先選取要產出對帳單的卡片", parent=parent); return

            indices  = [tree.index(i) for i in sel_ids]
            selected = [_displayed_cards[i] for i in indices]

            location_lookup = {}
            if _LOCATION_CACHE_PATH.exists():
                try:
                    location_lookup = json.loads(_LOCATION_CACHE_PATH.read_text(encoding="utf-8"))
                except Exception:
                    pass

            def _worker():
                out_dir = self._get_path("output_statement")
                return [generate_statement(c, location_lookup, output_dir=out_dir) for c in selected]

            def _on_done(paths):
                msg = "\n".join(str(p) for p in paths)
                if messagebox.askyesno("生成成功",
                        f"已生成 {len(paths)} 份對帳單：\n{msg}\n\n是否立即開啟？", parent=parent):
                    for p in paths:
                        os.startfile(p)

            self._run_task(_worker,
                            buttons=[gen_btn],
                            status_label=out_label,
                            loading_text="生成中…",
                            success_text=lambda paths: f"✔  已生成 {len(paths)} 份對帳單",
                            on_success=_on_done)

        bb = tk.Frame(parent, bg=BG)
        bb.pack(fill="x", padx=12, pady=8)
        gen_btn = ctk.CTkButton(bb, text="📄  生成對帳單（選取的卡片）", command=_generate,
                       fg_color="#1e8449", hover_color="#196f3d", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8)
        gen_btn.pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 8：生產群組紀錄
    # ════════════════════════════════════════════════════════
    def _build_tab_production(self, parent, FONT, FONTB, BG):
        from _paths import _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH  # noqa: F401
        GRAY  = "#5d6d7e"
        GREEN = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _all_cards:       list[dict] = []
        _displayed_cards: list[dict] = []

        # ── 抓取列 ────────────────────────────────────────
        fetch_row = tk.Frame(parent, bg=BG)
        fetch_row.pack(fill="x", padx=12, pady=(12, 2))

        ctk.CTkLabel(fetch_row, text="來源：出貨一覽表 Google Sheet",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(fetch_row, text="下單日期 從：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(8, 2))
        from tkcalendar import DateEntry
        from datetime import date as _dt_cls
        date_entry = DateEntry(fetch_row, width=10, date_pattern='y/m/d',
                                font=FONT_S, maxdate=_dt_cls.today())
        date_entry.pack(side="left", padx=(0, 2))
        ctk.CTkButton(fetch_row, text="全部",
                       command=lambda: _show_all(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(2, 8))

        fetch_status = ctk.CTkLabel(fetch_row, text="", fg_color="transparent",
                                     font=FONT_S, text_color=GRAY)
        fetch_status.pack(side="left", padx=(0, 0))

        def _parse_date_key(s: str):
            import re
            s = str(s).replace("前", "").strip()
            m = re.search(r'(\d{1,2})[/\-](\d{1,2})', s)
            if m:
                return (int(m.group(1)), int(m.group(2)))
            return (99, 99)

        def _apply_days_filter(base: list | None = None):
            import re
            from datetime import date
            source = base if base is not None else _all_cards
            from_date = date_entry.get_date()
            def _parse_d(s):
                m = re.match(r'(\d{1,2})/(\d{1,2})', str(s).strip())
                if m:
                    try: return date(date.today().year, int(m.group(1)), int(m.group(2)))
                    except ValueError: pass
                return None
            filtered = [r for r in source if (_parse_d(r.get("order_date","")) or date.min) >= from_date]
            if base is not None:
                _all_cards.clear(); _all_cards.extend(source)
            _render_tree(filtered)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 筆，篩選後 {len(filtered)} 筆",
                text_color=GREEN)

        def _show_all(base: list | None = None):
            source = base if base is not None else _all_cards
            if base is not None:
                _all_cards.clear(); _all_cards.extend(source)
            _render_tree(_all_cards)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 筆（全部顯示）",
                text_color=GREEN)

        def _render_tree(records: list):
            _displayed_cards.clear(); _displayed_cards.extend(records)
            tree.delete(*tree.get_children())
            for r in records:
                tree.insert("", "end", values=(
                    r.get("prefix",       ""),
                    r.get("order_date",   ""),
                    r.get("company",      ""),
                    r.get("product",      ""),
                    r.get("delivery",     ""),
                    r.get("payment_type", ""),
                ))
            tree.selection_set(tree.get_children())
            prev_title_lbl.configure(
                text=f"  出貨一覽表資料（共 {len(records)} 筆，可多選）  ")

        def _fetch():
            from sync.syncer_shipping_order import fetch_from_overview
            if not _GSHEETS_CREDS_PATH.exists():
                messagebox.showerror("缺少憑證", f"找不到 {_GSHEETS_CREDS_PATH}", parent=parent); return
            fetch_status.configure(text="讀取中…", text_color=GRAY)

            def _run():
                try:
                    records = fetch_from_overview(_GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH)
                    records.sort(key=lambda r: _parse_date_key(r.get("order_date", "")), reverse=True)
                    parent.after(0, lambda r=records: _apply_days_filter(r))
                except Exception as e:
                    from sync.syncer_sheets import _fmt_api_error
                    parent.after(0, lambda e=e: fetch_status.configure(
                        text=f"✘  {_fmt_api_error(e)}", text_color="#c0392b"))
            threading.Thread(target=_run, daemon=True).start()

        ctk.CTkButton(fetch_row, text="🔄 讀取", command=_fetch,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=70, height=28, corner_radius=4
                       ).pack(side="left")
        ctk.CTkButton(fetch_row, text="篩選", command=lambda: _apply_days_filter(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_outer = tk.Frame(parent, bg="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_inner = tk.Frame(prev_outer, bg=BG)
        prev_inner.pack(fill="both", expand=True, padx=1, pady=1)

        prev_title_lbl = ctk.CTkLabel(prev_inner,
                                       text="  出貨一覽表資料（共 0 筆，可多選）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(4, 0))

        tree_frame = tk.Frame(prev_inner, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        cols = ("prefix", "order_date", "company", "product", "delivery", "payment")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                             selectmode="extended", height=6)
        tree.heading("prefix",     text="業務")
        tree.heading("order_date", text="下單日期")
        tree.heading("company",    text="客戶名稱")
        tree.heading("product",    text="品號 / 產品")
        tree.heading("delivery",   text="交期")
        tree.heading("payment",    text="付款條件")
        tree.column("prefix",     width=55,  anchor="center", stretch=False)
        tree.column("order_date", width=70,  anchor="center", stretch=False)
        tree.column("company",    width=135, anchor="w",      stretch=False)
        tree.column("product",    width=240, anchor="w",      stretch=True)
        tree.column("delivery",   width=75,  anchor="center", stretch=False)
        tree.column("payment",    width=90,  anchor="w",      stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # ── 拖曳框選 ──────────────────────────────────────
        _bind_drag_select(tree)

        # ── 全選 / 取消全選 / 刪除 ───────────────────────────
        sel_row = tk.Frame(parent, bg=BG)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))

        def _delete_selected_production():
            sel_ids = tree.selection()
            if not sel_ids:
                return
            to_remove = [_displayed_cards[tree.index(i)] for i in sel_ids]
            for card in to_remove:
                if card in _all_cards:
                    _all_cards.remove(card)
                if card in _displayed_cards:
                    _displayed_cards.remove(card)
            for iid in sel_ids:
                tree.delete(iid)
            prev_title_lbl.configure(
                text=f"  出貨一覽表資料（共 {len(_displayed_cards)} 筆，可多選）  ")

        ctk.CTkButton(sel_row, text="🗑 刪除選取",
                       command=_delete_selected_production,
                       fg_color="#922b21", hover_color="#7b241c", text_color="white",
                       font=FONT_S, width=90, height=26, corner_radius=4
                       ).pack(side="left")

        # ── 狀態列 & 寫入按鈕 ─────────────────────────────
        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _write():
            from sync.syncer_production import write_sheet_records
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先選取要寫入的資料", parent=parent); return

            indices  = [tree.index(i) for i in sel_ids]
            selected = [_displayed_cards[i] for i in indices]

            prod_file = self._get_path("production_file")
            if not messagebox.askyesno("確認寫入",
                    f"即將寫入 {len(selected)} 筆資料至「總表」，確定繼續？",
                    parent=parent):
                return

            out_label.configure(text="寫入中…", text_color=GRAY)

            def _run():
                try:
                    added = write_sheet_records(selected, production_file=prod_file)
                    def _done(n=added):
                        out_label.configure(text=f"✔  寫入完成，新增 {n} 筆資料", text_color=GREEN)
                        if messagebox.askyesno("完成", f"新增 {n} 筆資料\n\n是否立即開啟生產群組紀錄？", parent=parent):
                            os.startfile(str(prod_file))
                    parent.after(0, _done)
                except Exception as e:
                    parent.after(0, lambda e=e: (
                        out_label.configure(text=f"✘  {e}", text_color="#c0392b"),
                        messagebox.showerror("寫入失敗", str(e), parent=parent)
                    ))
            threading.Thread(target=_run, daemon=True).start()

        bb = tk.Frame(parent, bg=BG)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="📝  寫入選取的卡片 → 生產群組紀錄.xlsx", command=_write,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 9：出貨指示單
    # ════════════════════════════════════════════════════════
    def _build_tab_shipping_order(self, parent, FONT, FONTB, BG):
        from _paths import _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH  # noqa: F401
        GRAY  = "#5d6d7e"
        GREEN = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _all_cards:       list[dict] = []
        _displayed_cards: list[dict] = []

        # ── 抓取列 ────────────────────────────────────────
        fetch_row = tk.Frame(parent, bg=BG)
        fetch_row.pack(fill="x", padx=12, pady=(12, 2))

        ctk.CTkLabel(fetch_row, text="來源：出貨一覽表 Google Sheet",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(fetch_row, text="下單日期 從：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(0, 2))
        from tkcalendar import DateEntry
        from datetime import date as _dt_cls
        date_entry = DateEntry(fetch_row, width=10, date_pattern='y/m/d',
                                font=FONT_S, maxdate=_dt_cls.today())
        date_entry.pack(side="left", padx=(0, 2))
        ctk.CTkButton(fetch_row, text="全部",
                       command=lambda: _show_all(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 8))

        fetch_status = ctk.CTkLabel(fetch_row, text="", fg_color="transparent",
                                     font=FONT_S, text_color=GRAY)
        fetch_status.pack(side="left", padx=(0, 0))

        def _parse_date_key(s: str):
            """將 '6/1前'、'5/29' 等格式轉為可排序的 (month, day)，無法解析放最後。"""
            import re
            s = str(s).replace("前", "").strip()
            m = re.search(r'(\d{1,2})[/\-](\d{1,2})', s)
            if m:
                return (int(m.group(1)), int(m.group(2)))
            return (99, 99)

        def _get_region(company: str) -> str:
            from sync.syncer_shipping_order import _classify_region
            from _paths import _LOCATION_CACHE_PATH
            cache = {}
            if _LOCATION_CACHE_PATH.exists():
                try: cache = json.loads(_LOCATION_CACHE_PATH.read_text(encoding="utf-8"))
                except Exception: pass
            return _classify_region(cache.get(company, "")) or "？"

        def _populate_tree(records: list):
            _displayed_cards.clear()
            _displayed_cards.extend(records)
            tree.delete(*tree.get_children())
            for r in records:
                tree.insert("", "end", values=(
                    r.get("prefix",       ""),
                    r.get("order_date",   ""),
                    r.get("delivery",     ""),
                    r.get("company",      ""),
                    _get_region(r.get("company", "")),
                    r.get("product",      ""),
                ))
            tree.selection_set(tree.get_children())
            prev_title_lbl.configure(
                text=f"  出貨一覽表資料（共 {len(records)} 筆，可多選）  ")
            fetch_status.configure(text=f"✔  找到 {len(records)} 筆", text_color=GREEN)

        def _apply_days_filter(base: list | None = None):
            import re
            from datetime import date
            source = base if base is not None else _all_cards
            from_date = date_entry.get_date()
            def _parse_d(s):
                m = re.match(r'(\d{1,2})/(\d{1,2})', str(s).replace("前", "").strip())
                if m:
                    try: return date(date.today().year, int(m.group(1)), int(m.group(2)))
                    except ValueError: pass
                return None
            filtered = [r for r in source if (_parse_d(r.get("order_date", "")) or date.min) >= from_date]
            if base is not None:
                _all_cards.clear(); _all_cards.extend(source)
            _populate_tree(filtered)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 筆，篩選後 {len(filtered)} 筆",
                text_color=GREEN)

        def _show_all():
            _populate_tree(_all_cards)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 筆（全部顯示）",
                text_color=GREEN)

        def _fetch():
            from sync.syncer_shipping_order import fetch_from_overview
            if not _GSHEETS_CREDS_PATH.exists():
                messagebox.showerror("缺少憑證", f"找不到 {_GSHEETS_CREDS_PATH}", parent=parent); return
            fetch_status.configure(text="讀取中…", text_color=GRAY)

            def _run():
                try:
                    records = fetch_from_overview(_GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH)
                    records.sort(key=lambda r: _parse_date_key(r.get("order_date", "")), reverse=True)
                    parent.after(0, lambda r=records: _apply_days_filter(r))
                except Exception as e:
                    from sync.syncer_sheets import _fmt_api_error
                    parent.after(0, lambda e=e: fetch_status.configure(
                        text=f"✘  {_fmt_api_error(e)}", text_color="#c0392b"))
            threading.Thread(target=_run, daemon=True).start()

        ctk.CTkButton(fetch_row, text="🔄 讀取出貨一覽表", command=_fetch,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=120, height=28, corner_radius=4
                       ).pack(side="left")
        ctk.CTkButton(fetch_row, text="篩選", command=lambda: _apply_days_filter(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))
        ctk.CTkButton(fetch_row, text="地區管理", command=lambda: self._open_location_editor(parent),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=70, height=28, corner_radius=4
                       ).pack(side="left", padx=(6, 0))

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_outer = tk.Frame(parent, bg="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_inner = tk.Frame(prev_outer, bg=BG)
        prev_inner.pack(fill="both", expand=True, padx=1, pady=1)

        prev_title_lbl = ctk.CTkLabel(prev_inner,
                                       text="  出貨一覽表資料（共 0 筆，可多選）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(4, 0))

        tree_frame = tk.Frame(prev_inner, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        cols = ("prefix", "order_date", "delivery", "company", "region", "product")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                             selectmode="extended", height=6)
        tree.heading("prefix",     text="業務")
        tree.heading("order_date", text="下單日期")
        tree.heading("delivery",   text="出貨日期")
        tree.heading("company",    text="客戶名稱")
        tree.heading("region",     text="區域")
        tree.heading("product",    text="機台")
        tree.column("prefix",     width=50,  anchor="center", stretch=False)
        tree.column("order_date", width=68,  anchor="center", stretch=False)
        tree.column("delivery",   width=68,  anchor="center", stretch=False)
        tree.column("company",    width=130, anchor="w",      stretch=False)
        tree.column("region",     width=50,  anchor="center", stretch=False)
        tree.column("product",    width=260, anchor="w",      stretch=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # ── 拖曳框選 ──────────────────────────────────────
        _bind_drag_select(tree)

        # ── 全選 / 取消全選 / 刪除 ───────────────────────────
        sel_row = tk.Frame(parent, bg=BG)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))

        def _delete_selected_shipping():
            sel_ids = tree.selection()
            if not sel_ids:
                return
            to_remove = [_displayed_cards[tree.index(i)] for i in sel_ids]
            for card in to_remove:
                if card in _displayed_cards:
                    _displayed_cards.remove(card)
            for iid in sel_ids:
                tree.delete(iid)
            prev_title_lbl.configure(
                text=f"  出貨一覽表資料（共 {len(_displayed_cards)} 筆，可多選）  ")

        ctk.CTkButton(sel_row, text="🗑 刪除選取",
                       command=_delete_selected_shipping,
                       fg_color="#922b21", hover_color="#7b241c", text_color="white",
                       font=FONT_S, width=90, height=26, corner_radius=4
                       ).pack(side="left")

        # ── 狀態列 & 推送按鈕 ─────────────────────────────
        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _push():
            from sync.syncer_shipping_order import push_shipping_orders
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先選取要推送的卡片", parent=parent); return
            if not _GSHEETS_CREDS_PATH.exists():
                messagebox.showerror("缺少憑證", f"找不到 {_GSHEETS_CREDS_PATH}", parent=parent); return

            indices  = [tree.index(i) for i in sel_ids]
            selected = [_displayed_cards[i] for i in indices]

            if not messagebox.askyesno("確認推送",
                    f"即將推送 {len(selected)} 筆至出貨指示單（工作表 2026），確定繼續？",
                    parent=parent):
                return

            out_label.configure(text="推送中…", text_color=GRAY)
            _DST_URL = "https://docs.google.com/spreadsheets/d/15H37eDyC2MtqreSggyj8QR32vpN3n7R7Bdb8NTOXJdU/edit"

            def _run():
                try:
                    added = push_shipping_orders(selected, _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH)
                    def _done(n=added):
                        out_label.configure(text=f"✔  推送完成，新增 {n} 筆資料", text_color=GREEN)
                        if messagebox.askyesno("完成", f"新增 {n} 筆資料\n\n是否立即開啟出貨指示單？", parent=parent):
                            webbrowser.open(_DST_URL)
                    parent.after(0, _done)
                except Exception as e:
                    from sync.syncer_sheets import _fmt_api_error
                    msg = _fmt_api_error(e)
                    parent.after(0, lambda m=msg: (
                        out_label.configure(text=f"✘  {m}", text_color="#c0392b"),
                        messagebox.showerror("推送失敗", m, parent=parent)
                    ))
            threading.Thread(target=_run, daemon=True).start()

        bb = tk.Frame(parent, bg=BG)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="☁  推送選取的卡片 → 出貨指示單", command=_push,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 9：建立卡片
    # ════════════════════════════════════════════════════════
    def _open_location_editor(self, parent):
        """開啟地區對照表編輯器（customer_locations.json）。"""
        from _paths import _LOCATION_CACHE_PATH
        GRAY  = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        win = tk.Toplevel(parent)
        win.title("地區對照表編輯")
        win.geometry("480x420")
        win.grab_set()

        # 讀取現有資料
        data: dict[str, str] = {}
        if _LOCATION_CACHE_PATH.exists():
            try: data = json.loads(_LOCATION_CACHE_PATH.read_text(encoding="utf-8"))
            except Exception: pass

        # Treeview
        frame = tk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        cols = ("company", "location")
        tv = ttk.Treeview(frame, columns=cols, show="headings", height=14)
        tv.heading("company",  text="客戶名稱")
        tv.heading("location", text="地區")
        tv.column("company",  width=220, anchor="w")
        tv.column("location", width=180, anchor="w")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def _reload():
            tv.delete(*tv.get_children())
            for c, l in sorted(data.items()):
                tv.insert("", "end", values=(c, l))

        _reload()

        # 編輯列
        edit_row = tk.Frame(win)
        edit_row.pack(fill="x", padx=10, pady=2)
        tk.Label(edit_row, text="客戶名稱", font=FONT_S).pack(side="left")
        company_var = tk.StringVar()
        tk.Entry(edit_row, textvariable=company_var, width=18,
                 font=FONT_S).pack(side="left", padx=(2, 6))
        tk.Label(edit_row, text="地區", font=FONT_S).pack(side="left")
        location_var = tk.StringVar()
        tk.Entry(edit_row, textvariable=location_var, width=14,
                 font=FONT_S).pack(side="left", padx=(2, 6))

        def _on_select(event):
            sel = tv.selection()
            if sel:
                v = tv.item(sel[0], "values")
                company_var.set(v[0]); location_var.set(v[1])
        tv.bind("<<TreeviewSelect>>", _on_select)

        def _upsert():
            c = company_var.get().strip(); l = location_var.get().strip()
            if not c or not l: return
            data[c] = l; _reload()

        def _delete():
            sel = tv.selection()
            if not sel: return
            c = tv.item(sel[0], "values")[0]
            data.pop(c, None); _reload()
            company_var.set(""); location_var.set("")

        btn_row = tk.Frame(win)
        btn_row.pack(fill="x", padx=10, pady=(2, 4))
        ctk.CTkButton(btn_row, text="新增/更新", command=_upsert,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row, text="刪除選取", command=_delete,
                       fg_color="#c0392b", hover_color="#922b21", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 4))

        def _save():
            _LOCATION_CACHE_PATH.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            win.destroy()

        ctk.CTkButton(btn_row, text="儲存並關閉", command=_save,
                       fg_color="#1e8449", hover_color="#196f3d", text_color="white",
                       font=FONT_S, width=90, height=28, corner_radius=4
                       ).pack(side="right")

    # ════════════════════════════════════════════════════════
    #  Tab：會計對帳
    # ════════════════════════════════════════════════════════
    def _build_tab_accounting(self, parent, FONT, FONTB, BG):
        GRAY  = "#5d6d7e"
        GREEN = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _all_records:       list[dict] = []
        _displayed_records: list[dict] = []

        # ── 來源檔案列 ────────────────────────────────────
        src_row = tk.Frame(parent, bg=BG)
        src_row.pack(fill="x", padx=12, pady=(12, 4))

        ctk.CTkLabel(src_row, text="生產群組紀錄：",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left")

        src_var = tk.StringVar()
        ctk.CTkEntry(src_row, textvariable=src_var, font=FONT_S,
                      width=320, height=28, corner_radius=4, border_width=1
                      ).pack(side="left", padx=(4, 4))

        def _pick_file():
            p = filedialog.askopenfilename(
                title="選擇生產群組紀錄 Excel",
                filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
            if p:
                src_var.set(p)

        ctk.CTkButton(src_row, text="選擇", command=_pick_file,
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=56, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(src_row, text="工作表：",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left")
        sheet_var = tk.StringVar(value="收款紀錄")
        sheet_cb = ttk.Combobox(src_row, textvariable=sheet_var, font=FONT_S,
                                 state="readonly", width=10,
                                 values=["收款紀錄", "支票收款"])
        sheet_cb.pack(side="left", padx=(4, 10))

        fetch_status = ctk.CTkLabel(src_row, text="", fg_color="transparent",
                                     font=FONT_S, text_color=GRAY)
        fetch_status.pack(side="left")

        # 自動填入路徑
        try:
            src_var.set(str(self._get_path("production_file")))
        except Exception:
            pass

        # ── 篩選列：搜尋 + 未付款勾選 ────────────────────
        filter_row = tk.Frame(parent, bg=BG)
        filter_row.pack(fill="x", padx=12, pady=(2, 2))

        ctk.CTkLabel(filter_row, text="搜尋客戶：",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left")
        search_var = tk.StringVar()
        search_entry = ctk.CTkEntry(filter_row, textvariable=search_var,
                                     font=FONT_S, width=160, height=26,
                                     corner_radius=4, border_width=1)
        search_entry.pack(side="left", padx=(4, 4))

        unpaid_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(filter_row, text="只顯示尚未付款（原色背景）",
                         variable=unpaid_var,
                         font=FONT_S, text_color=GRAY,
                         fg_color="#2e86c1", hover_color="#1a5276",
                         command=lambda: _apply_filter()
                         ).pack(side="left", padx=(10, 8))

        ctk.CTkButton(filter_row, text="篩選",
                       command=lambda: _apply_filter(),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(filter_row, text="清除",
                       command=lambda: (search_var.set(""),
                                        unpaid_var.set(False),
                                        _apply_filter()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=26, corner_radius=4
                       ).pack(side="left")
        search_entry.bind("<Return>", lambda _e: _apply_filter())

        def _apply_filter():
            q           = search_var.get().strip().lower()
            unpaid_only = unpaid_var.get()
            filtered = []
            for r in _all_records:
                if q and q not in r.get("company", "").lower():
                    continue
                if unpaid_only and not r.get("is_original_color", True):
                    continue
                filtered.append(r)
            _render_tree(filtered)

        def _render_tree(records: list):
            _displayed_records.clear()
            _displayed_records.extend(records)
            tree.delete(*tree.get_children())
            for r in records:
                raw_short = r["raw"][:55] + "…" if len(r["raw"]) > 55 else r["raw"]
                tag = "unpaid" if r.get("is_original_color", True) else "paid"
                tree.insert("", "end", tags=(tag,), values=(
                    r.get("date_str",   ""),
                    r.get("company",    ""),
                    r.get("amount_str", ""),
                    raw_short,
                ))
            tree.selection_set(tree.get_children())
            prev_title_lbl.configure(
                text=f"  收款紀錄（共 {len(records)} 筆，可多選）  ")

        def _fetch():
            p = src_var.get().strip()
            if not p:
                messagebox.showwarning("未選擇檔案", "請先選擇生產群組紀錄 Excel 檔案",
                                       parent=parent)
                return
            fetch_status.configure(text="讀取中…", text_color=GRAY)

            def _run():
                try:
                    from sync.syncer_accounting import read_payment_records
                    records = read_payment_records(Path(p), sheet_name=sheet_var.get())
                    # 日期由大到小排序
                    records.sort(key=lambda r: r["date_sort_key"], reverse=True)
                    def _done(r=records):
                        _all_records.clear()
                        _all_records.extend(r)
                        _apply_filter()
                        fetch_status.configure(
                            text=f"✔  共 {len(r)} 筆", text_color=GREEN)
                    parent.after(0, _done)
                except Exception as e:
                    parent.after(0, lambda e=e: fetch_status.configure(
                        text=f"✘  {e}", text_color="#c0392b"))
            threading.Thread(target=_run, daemon=True).start()

        ctk.CTkButton(src_row, text="🔄 讀取收款紀錄", command=_fetch,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=110, height=28, corner_radius=4
                       ).pack(side="left")

        # ── Treeview ──────────────────────────────────────
        prev_outer = tk.Frame(parent, bg="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_inner = tk.Frame(prev_outer, bg=BG)
        prev_inner.pack(fill="both", expand=True, padx=1, pady=1)

        prev_title_lbl = ctk.CTkLabel(prev_inner,
                                       text="  收款紀錄（共 0 筆，可多選）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(4, 0))

        tree_frame = tk.Frame(prev_inner, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        cols = ("date", "company", "amount", "raw")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                             selectmode="extended", height=10)
        tree.heading("date",    text="日期")
        tree.heading("company", text="解析客戶名稱")
        tree.heading("amount",  text="金額")
        tree.heading("raw",     text="收款明細")
        tree.column("date",    width=80,  anchor="center", stretch=False)
        tree.column("company", width=110, anchor="w",      stretch=False)
        tree.column("amount",  width=90,  anchor="e",      stretch=False)
        tree.column("raw",     width=380, anchor="w",      stretch=True)

        # 原色（未付款）顯示正常；有背景色（已付款）的列顯示為灰色
        tree.tag_configure("unpaid", foreground="#1b2631")
        tree.tag_configure("paid",   foreground="#aab7b8")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # 拖曳框選
        _bind_drag_select(tree)

        # ── 全選 / 取消全選 ───────────────────────────────
        sel_row = tk.Frame(parent, bg=BG)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=26, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=26, corner_radius=4
                       ).pack(side="left")

        # ── 狀態列 & 主動作按鈕 ───────────────────────────
        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _match_and_move():
            from sync.syncer_accounting import preview_matches, execute_moves
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先選取要對帳的付款紀錄", parent=parent)
                return

            indices  = [tree.index(i) for i in sel_ids]
            selected = [_displayed_records[i] for i in indices]

            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent)
                return

            out_label.configure(text="比對中…", text_color=GRAY)

            def _run():
                try:
                    result = preview_matches(selected, api_key, token)
                    parent.after(0, lambda r=result: _show_confirm(r))
                except Exception as e:
                    parent.after(0, lambda e=e: (
                        out_label.configure(text=f"✘  {e}", text_color="#c0392b"),
                        messagebox.showerror("比對失敗", str(e), parent=parent)
                    ))

            def _show_confirm(result):
                matched   = result["matched"]
                unmatched = result["unmatched"]
                dst_id    = result["dst_id"]

                lines = []
                if matched:
                    lines.append(f"✔  找到 {len(matched)} 筆符合：")
                    for pay, card in matched:
                        lines.append(
                            f"  • {pay['company']}  ${pay['amount_str']}"
                            f"  →  {card['title'][:40]}")
                if unmatched:
                    lines.append(f"\n✘  未找到 {len(unmatched)} 筆：")
                    for pay in unmatched:
                        lines.append(f"  • {pay['company']}  ${pay['amount_str']}")

                if not matched:
                    messagebox.showinfo("比對結果",
                        "\n".join(lines) or "沒有任何符合的卡片", parent=parent)
                    out_label.configure(text="比對完成，無符合卡片", text_color=GRAY)
                    return

                lines.append(f"\n確定將 {len(matched)} 張卡片移至「會計對帳完成」？")
                if not messagebox.askyesno("比對結果", "\n".join(lines), parent=parent):
                    out_label.configure(text="已取消", text_color=GRAY)
                    return

                out_label.configure(text="移動中…", text_color=GRAY)

                def _do_move():
                    try:
                        n = execute_moves(matched, dst_id, api_key, token)
                        parent.after(0, lambda n=n: out_label.configure(
                            text=f"✔  完成！已移動 {n} 張卡片至「會計對帳完成」",
                            text_color=GREEN))
                    except Exception as e:
                        parent.after(0, lambda e=e: (
                            out_label.configure(text=f"✘  {e}", text_color="#c0392b"),
                            messagebox.showerror("移動失敗", str(e), parent=parent)
                        ))
                threading.Thread(target=_do_move, daemon=True).start()

            threading.Thread(target=_run, daemon=True).start()

        bb = tk.Frame(parent, bg=BG)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🔍  比對並移動選取的紀錄 → 會計對帳完成",
                       command=_match_and_move,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    def _build_tab_create_cards(self, parent, FONT, FONTB, BG):
        from tksheet import Sheet
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _HEADERS = ["序號", "標題", "描述", "備註/需求"]
        _EMPTY   = ["", "", "", ""]

        top_row = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        top_row.pack(fill="x", padx=12, pady=(12, 4))

        src_outer, src_frame = _mk_lf(top_row, "來源 Excel", BG, FONTB)
        src_outer.pack(side="left", fill="y", padx=(0, 8))
        src_frame.columnconfigure(1, weight=1)

        hint_outer, hint_frame = _mk_lf(top_row, "使用說明", BG, FONTB)
        hint_outer.pack(side="left", fill="both", expand=True)
        hint_text = "導入 Excel 後即可編輯資料\n可以選取多筆資料搬移刪除\n不可跳號選取"
        ctk.CTkLabel(hint_frame, text=hint_text, fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      justify="left", anchor="nw").pack(padx=10, pady=8, anchor="nw")

        ctk.CTkLabel(src_frame, text="檔案路徑：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        path_var = tk.StringVar()
        ctk.CTkEntry(src_frame, textvariable=path_var, font=FONT_S,
                      corner_radius=4, border_width=1
                      ).grid(row=0, column=1, sticky="ew", padx=4, pady=6)

        ctk.CTkLabel(src_frame, text="工作表：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        sheet_var = tk.StringVar()
        sheet_cb  = ttk.Combobox(src_frame, textvariable=sheet_var, font=FONT_S,
                                  state="readonly", width=20)
        sheet_cb.grid(row=1, column=1, sticky="w", padx=4, pady=4)

        def _pick_file():
            from sync.creator_trello import get_sheet_names
            p = filedialog.askopenfilename(
                title="選擇 Excel 檔案",
                filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
            if not p: return
            path_var.set(p)
            try:
                names = get_sheet_names(Path(p))
                all_options = ["全部"] + names
                sheet_cb["values"] = all_options
                sheet_var.set(all_options[0])
            except Exception:
                sheet_cb["values"] = []
                sheet_var.set("")

        ctk.CTkButton(src_frame, text="選擇", command=_pick_file,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=56, height=26, corner_radius=4
                       ).grid(row=0, column=2, padx=(0, 4), pady=6)

        status_lbl = ctk.CTkLabel(src_frame, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY)
        status_lbl.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        # ── 序號篩選列 ────────────────────────────────────
        filter_frame = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        filter_frame.pack(fill="x", padx=12, pady=(0, 2))
        ctk.CTkLabel(filter_frame, text="序號篩選：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        filter_var = tk.StringVar()
        ctk.CTkEntry(filter_frame, textvariable=filter_var, font=FONT_S,
                      width=240, height=28, corner_radius=4
                      ).pack(side="left", padx=(4, 6))
        ctk.CTkLabel(filter_frame, text="（逗號分隔，例如 1,3,5；空白=全部顯示）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY).pack(side="left")

        # ── 可編輯 Sheet 預覽 ─────────────────────────────
        prev_outer = ctk.CTkFrame(parent, fg_color=BG, corner_radius=8,
                                   border_width=1, border_color="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_title_lbl = ctk.CTkLabel(prev_outer,
                                       text="  卡片預覽（雙擊儲存格可編輯，0 筆）  ",
                                       fg_color=BG, text_color=GRAY,
                                       font=FONT)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(6, 0))
        prev_frame = ctk.CTkFrame(prev_outer, fg_color=BG, corner_radius=0)
        prev_frame.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        sheet = Sheet(prev_frame,
                      headers=_HEADERS,
                      data=[_EMPTY[:] for _ in range(10)],
                      column_width=200,
                      row_height=28)
        sheet.set_column_widths([70, 300, 420, 200])
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

        _all_data: list[dict] = []

        def _apply_filter():
            raw = filter_var.get().strip()
            if raw:
                keys = {s.strip() for s in raw.split(",") if s.strip()}
                filtered = [c for c in _all_data if str(c["seq"]) in keys]
            else:
                filtered = _all_data
            rows = [[c["seq"], c["title"], c["desc"], c.get("notes", "")] for c in filtered]
            while len(rows) < len(filtered) + 5:
                rows.append(_EMPTY[:])
            sheet.data = rows
            prev_title_lbl.configure(text=f"  卡片預覽（雙擊儲存格可編輯，{len(filtered)} 筆）  ")

        ctk.CTkButton(filter_frame, text="篩選", command=_apply_filter,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=56, height=28, corner_radius=4
                       ).pack(side="left")
        ctk.CTkButton(filter_frame, text="清除",
                       command=lambda: (filter_var.set(""), _apply_filter()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=56, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))

        def _load_preview():
            from sync.creator_trello import read_excel_cards
            p = path_var.get().strip()
            if not p:
                messagebox.showwarning("未選擇檔案", "請先選擇 Excel 檔案", parent=parent); return
            status_lbl.configure(text="讀取中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                selected = sheet_var.get()
                _all_data.clear()
                if selected == "全部":
                    all_sheets = [v for v in sheet_cb["values"] if v != "全部"]
                    for sname in all_sheets:
                        _all_data.extend(read_excel_cards(Path(p), sheet_name=sname))
                else:
                    _all_data.extend(read_excel_cards(Path(p), sheet_name=selected or None))
                filter_var.set("")
                _apply_filter()
                status_lbl.configure(text=f"✔  讀取完成，共 {len(_all_data)} 筆", text_color="#1e8449")
            except Exception as e:
                status_lbl.configure(text=f"✘  {e}", text_color="#c0392b")

        ctk.CTkButton(src_frame, text="讀取預覽", command=_load_preview,
                       fg_color="#117a65", hover_color="#0e6655", text_color="white",
                       font=FONT_S, width=70, height=26, corner_radius=4
                       ).grid(row=0, column=3, padx=(0, 8), pady=6)

        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _create():
            from sync.creator_trello import create_cards as trello_create_cards
            rows = sheet.data
            cards = []
            for row in rows:
                title = str(row[1]).strip() if len(row) > 1 else ""
                desc  = str(row[2]).strip() if len(row) > 2 else ""
                notes = str(row[3]).strip() if len(row) > 3 else ""
                if not title: continue
                if notes:
                    desc = desc + f"\n需求：\n{notes}"
                cards.append({"title": title, "desc": desc})
            if not cards:
                messagebox.showwarning("無資料", "表格中沒有填入標題的列", parent=parent); return
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent); return
            if not messagebox.askyesno("確認建立",
                    f"即將在「0.待評估」清單建立 {len(cards)} 張卡片，確定繼續？", parent=parent): return

            out_label.configure(text="建立中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                created = trello_create_cards(cards, api_key, token)
                out_label.configure(text=f"✔  成功建立 {created} 張卡片", text_color="#1e8449")
            except Exception as e:
                out_label.configure(text=f"✘  {e}", text_color="#c0392b")
                messagebox.showerror("建立失敗", str(e), parent=parent)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🃏  建立全部卡片 → Trello 0.待評估", command=_create,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 10：下載卡片
    # ════════════════════════════════════════════════════════
    def _build_tab_download_cards(self, parent, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _list_map: dict[str, str] = {}
        _all_cards: list[dict]    = []

        top_outer, top = _mk_lf(parent, "Trello 清單", BG, FONTB)
        top_outer.pack(fill="x", padx=12, pady=(12, 4))
        top.columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="選擇清單：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)

        list_var = tk.StringVar()
        list_cb  = ttk.Combobox(top, textvariable=list_var, font=FONT_S,
                                 state="readonly", width=28)
        list_cb.grid(row=0, column=1, sticky="w", padx=4, pady=6)

        status_lbl = ctk.CTkLabel(top, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY)

        def _get_creds():
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent)
                return None, None
            return api_key, token

        def _fetch_lists():
            from sync.downloader_trello import get_board_lists
            api_key, token = _get_creds()
            if not api_key: return

            def _worker():
                return get_board_lists(api_key, token)

            def _on_done(lists):
                _list_map.clear()
                _list_map.update({lst["name"]: lst["id"] for lst in lists})
                names = list(_list_map.keys())
                list_cb["values"] = names
                if names:
                    list_var.set(names[0])

            self._run_task(_worker,
                            buttons=[fetch_lists_btn, preview_btn],
                            status_label=status_lbl,
                            loading_text="抓取中…",
                            success_text=lambda lists: f"找到 {len(lists)} 個清單",
                            on_success=_on_done)

        def _preview_cards():
            from sync.downloader_trello import get_list_cards
            selected = list_var.get().strip()
            if not selected or selected not in _list_map:
                messagebox.showwarning("未選擇清單", "請先抓取清單並選擇", parent=parent); return
            api_key, token = _get_creds()
            if not api_key: return

            def _worker():
                return get_list_cards(_list_map[selected], api_key, token)

            def _on_done(cards):
                _all_cards.clear()
                _all_cards.extend(cards)
                tree.delete(*tree.get_children())
                for card in cards:
                    labels = "、".join(
                        lbl.get("name") or lbl.get("color", "")
                        for lbl in card.get("labels", []))
                    att_count = len(card.get("attachments") or [])
                    tree.insert("", "end", values=(card["name"], labels, att_count))
                tree.selection_set(tree.get_children())
                prev_title_lbl.configure(text=f"  卡片預覽（共 {len(cards)} 張，可多選）  ")

            self._run_task(_worker,
                            buttons=[fetch_lists_btn, preview_btn],
                            status_label=status_lbl,
                            loading_text="載入卡片中…",
                            success_text=lambda cards: f"找到 {len(cards)} 張卡片",
                            on_success=_on_done)

        fetch_lists_btn = ctk.CTkButton(top, text="抓取清單", command=_fetch_lists,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=70, height=26, corner_radius=4)
        fetch_lists_btn.grid(row=0, column=2, padx=(0, 4), pady=6)
        preview_btn = ctk.CTkButton(top, text="預覽卡片", command=_preview_cards,
                       fg_color="#117a65", hover_color="#0e6655", text_color="white",
                       font=FONT_S, width=70, height=26, corner_radius=4)
        preview_btn.grid(row=0, column=3, padx=(0, 4), pady=6)
        status_lbl.grid(row=0, column=4, sticky="w", padx=4, pady=6)

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_outer = ctk.CTkFrame(parent, fg_color=BG, corner_radius=8,
                                   border_width=1, border_color="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_title_lbl = ctk.CTkLabel(prev_outer,
                                       text="  卡片預覽（共 0 張，可多選）  ",
                                       fg_color=BG, text_color=GRAY, font=FONT)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(6, 0))
        prev_inner = ctk.CTkFrame(prev_outer, fg_color=BG, corner_radius=0)
        prev_inner.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        cols = ("title", "labels", "att")
        tree = ttk.Treeview(prev_inner, columns=cols, show="headings",
                             selectmode="extended", height=8)
        tree.heading("title",  text="標題")
        tree.heading("labels", text="標籤")
        tree.heading("att",    text="附件")
        tree.column("title",  width=340, anchor="w",      stretch=True)
        tree.column("labels", width=130, anchor="w",      stretch=False)
        tree.column("att",    width=50,  anchor="center", stretch=False)

        vsb = ttk.Scrollbar(prev_inner, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        sel_row = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="left")

        pf = ctk.CTkFrame(parent, fg_color="#e8ecf0", corner_radius=6)
        pf.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(pf, text="輸出資料夾：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY, anchor="w", width=96).pack(side="left", padx=8, pady=5)
        ctk.CTkLabel(pf, text="（依⚙路徑設定）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY).pack(side="left", pady=5)

        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _download():
            from sync.downloader_trello import download_cards as trello_download_cards
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先預覽並選取要下載的卡片", parent=parent); return
            api_key, token = _get_creds()
            if not api_key: return

            indices  = [tree.index(i) for i in sel_ids]
            selected_cards = [_all_cards[i] for i in indices]
            list_name  = list_var.get().strip()
            output_dir = self._get_path("download_cards_dir") / list_name

            report = self._make_progress_reporter(out_label)

            def _worker():
                def _progress_cb(cur, total, name):
                    report(f"下載中… ({cur}/{total}) {name}")
                return trello_download_cards(
                    selected_cards, output_dir, api_key, token, progress_cb=_progress_cb)

            def _on_done(result):
                count, failed = result
                if failed:
                    out_label.configure(text_color="#e67e22")
                    messagebox.showwarning("部分附件失敗",
                        "以下附件下載失敗：\n" + "\n".join(failed), parent=parent)
                if messagebox.askyesno("下載完成",
                        f"共 {count} 張卡片\n\n是否立即開啟資料夾？", parent=parent):
                    os.startfile(str(output_dir))

            self._run_task(_worker,
                            buttons=[download_btn],
                            status_label=out_label,
                            loading_text="下載中…",
                            success_text=lambda result: (
                                f"✔  完成 {result[0]} 張，{len(result[1])} 個附件失敗" if result[1]
                                else f"✔  成功下載 {result[0]} 張卡片至：{output_dir}"),
                            on_success=_on_done)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=6)
        download_btn = ctk.CTkButton(bb, text="⬇  下載選取的卡片", command=_download,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8)
        download_btn.pack(fill="x")
