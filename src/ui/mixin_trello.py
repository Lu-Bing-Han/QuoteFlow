"""
mixin_trello.py — Trello 相關頁籤 mixin（出貨一覽表、生產群組紀錄、建立卡片、下載卡片）
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path


class _TrelloTab:
    """Mixin providing Trello-related tab builders and callbacks."""

    # ════════════════════════════════════════════════════════
    #  Tab 7：出貨一覽表
    # ════════════════════════════════════════════════════════
    def _build_tab_overview(self, parent, FONT, FONTB, BG):
        from _paths import _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH, _SYNCED_CARDS_PATH  # noqa: F401
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        # ── Trello 憑證 ───────────────────────────────────
        cred_frame = tk.LabelFrame(parent, text="Trello 憑證", bg=BG, font=FONTB)
        cred_frame.pack(fill="x", padx=12, pady=(12, 4))
        cred_frame.columnconfigure(1, weight=1)

        tr_cfg   = self._config.get("trello", {})
        key_var  = tk.StringVar(value=tr_cfg.get("api_key", ""))
        tok_var  = tk.StringVar(value=tr_cfg.get("token",   ""))

        tk.Label(cred_frame, text="API Key：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        key_entry = tk.Entry(cred_frame, textvariable=key_var, font=FONT_S, show="*")
        key_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=3)

        tk.Label(cred_frame, text="Token：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=3)
        tok_entry = tk.Entry(cred_frame, textvariable=tok_var, font=FONT_S, show="*")
        tok_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=3)

        def _show_hide(entry, btn):
            if entry.cget("show") == "*":
                entry.config(show=""); btn.config(text="隱藏")
            else:
                entry.config(show="*"); btn.config(text="顯示")

        btn_key = tk.Button(cred_frame, text="顯示", font=FONT_S,
                            command=lambda: _show_hide(key_entry, btn_key))
        btn_key.grid(row=0, column=2, padx=(0, 8), pady=3)
        btn_tok = tk.Button(cred_frame, text="顯示", font=FONT_S,
                            command=lambda: _show_hide(tok_entry, btn_tok))
        btn_tok.grid(row=1, column=2, padx=(0, 8), pady=3)

        def _save_trello_creds():
            self._config.setdefault("trello", {})
            self._config["trello"]["api_key"] = key_var.get().strip()
            self._config["trello"]["token"]   = tok_var.get().strip()
            self._save_config(self._config)
            messagebox.showinfo("已儲存", "Trello 憑證已儲存", parent=parent)

        tk.Button(cred_frame, text="儲存憑證", command=_save_trello_creds,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT, padx=8).grid(row=2, column=1, sticky="w", padx=8, pady=(4, 6))

        # ── Google Sheets 設定 ────────────────────────────
        gs_frame = tk.LabelFrame(parent, text="Google Sheets", bg=BG, font=FONTB)
        gs_frame.pack(fill="x", padx=12, pady=4)

        creds_status = tk.Label(
            gs_frame,
            text="✔  credentials.json 已就緒" if _GSHEETS_CREDS_PATH.exists()
                 else "✘  找不到 credentials.json（請放到 template 資料夾）",
            bg=BG, font=FONT_S,
            fg="#1e8449" if _GSHEETS_CREDS_PATH.exists() else "#c0392b",
            anchor="w",
        )
        creds_status.pack(fill="x", padx=8, pady=6)

        token_status = tk.Label(
            gs_frame,
            text="✔  已授權（gsheets_token.json 存在）" if _GSHEETS_TOKEN_PATH.exists()
                 else "尚未授權，點「同步」時會自動開啟瀏覽器",
            bg=BG, font=FONT_S,
            fg="#1e8449" if _GSHEETS_TOKEN_PATH.exists() else GRAY,
            anchor="w",
        )
        token_status.pack(fill="x", padx=8, pady=(0, 6))

        # ── 同步按鈕與狀態 ────────────────────────────────
        out_label = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY,
                             anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _sync():
            from sync.syncer_trello import fetch_po_cards
            from sync.syncer_sheets import sync_cards
            api_key = key_var.get().strip()
            token   = tok_var.get().strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未填", "請先填入 Trello API Key 與 Token", parent=parent)
                return
            if not _GSHEETS_CREDS_PATH.exists():
                messagebox.showerror("缺少憑證", f"找不到 {_GSHEETS_CREDS_PATH}", parent=parent)
                return

            out_label.config(text="抓取 Trello 卡片中…", fg=GRAY)
            parent.update_idletasks()
            try:
                cards = fetch_po_cards(api_key, token)
                out_label.config(text=f"找到 {len(cards)} 張卡片，同步至 Google Sheets…", fg=GRAY)
                parent.update_idletasks()

                added = sync_cards(cards, _GSHEETS_CREDS_PATH,
                                   _GSHEETS_TOKEN_PATH, _SYNCED_CARDS_PATH)

                token_status.config(text="✔  已授權（gsheets_token.json 存在）", fg="#1e8449")
                if added:
                    out_label.config(text=f"✔  同步完成，新增 {added} 筆資料", fg="#1e8449")
                else:
                    out_label.config(text="✔  同步完成，無新卡片", fg="#1e8449")
            except Exception as e:
                out_label.config(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("同步失敗", str(e), parent=parent)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(fill="x", padx=12)
        tk.Button(bb, text="🔄  同步 Trello → Google Sheets", command=_sync,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 8：生產群組紀錄
    # ════════════════════════════════════════════════════════
    def _build_tab_production(self, parent, FONT, FONTB, BG):
        from _paths import _PRODUCTION_SYNCED_PATH  # noqa: F401
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        # ── 說明 ──────────────────────────────────────────
        info = tk.LabelFrame(parent, text="說明", bg=BG, font=FONTB)
        info.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(info,
                 text="從 Trello「本周下單」抓取 2026/5/15 之後的新卡片，附加到生產群組紀錄 Excel。\n"
                      "Trello 憑證與「出貨一覽表」頁籤共用，請先在該頁籤儲存憑證。",
                 bg=BG, font=FONT_S, fg=GRAY, justify="left",
                 ).pack(padx=12, pady=8, anchor="w")

        # ── 檔案路徑 ──────────────────────────────────────
        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(fill="x", padx=12, pady=4)
        tk.Label(pf, text="寫入檔案：", bg="#e8ecf0", font=FONT, fg=GRAY,
                 anchor="w", width=10).pack(side="left", padx=8, pady=6)
        tk.Label(pf, text="（依⚙路徑設定）",
                 bg="#e8ecf0", font=FONT_S, fg=GRAY).pack(side="left", pady=6)

        # ── 狀態與同步按鈕 ────────────────────────────────
        out_label = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY,
                             anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(8, 0))

        def _sync():
            from sync.syncer_trello import fetch_po_cards
            from sync.syncer_production import sync_production
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning(
                    "憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證",
                    parent=parent)
                return

            out_label.config(text="抓取 Trello 卡片中…", fg=GRAY)
            parent.update_idletasks()
            try:
                cards = fetch_po_cards(api_key, token)
                out_label.config(text=f"找到 {len(cards)} 張卡片，寫入 Excel 中…", fg=GRAY)
                parent.update_idletasks()

                added = sync_production(cards, _PRODUCTION_SYNCED_PATH,
                                        production_file=self._get_path("production_file"))
                if added:
                    out_label.config(text=f"✔  同步完成，新增 {added} 筆資料", fg="#1e8449")
                else:
                    out_label.config(text="✔  同步完成，無新卡片（2026/5/15 之後）", fg="#1e8449")
                if messagebox.askyesno("同步完成",
                        f"{'新增 ' + str(added) + ' 筆資料' if added else '無新卡片'}\n\n是否立即開啟生產群組紀錄？",
                        parent=parent):
                    os.startfile(str(self._get_path("production_file")))
            except Exception as e:
                out_label.config(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("同步失敗", str(e), parent=parent)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(fill="x", padx=12)
        tk.Button(bb, text="🔄  同步 Trello → 生產群組紀錄.xlsx", command=_sync,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 9：建立卡片
    # ════════════════════════════════════════════════════════
    def _build_tab_create_cards(self, parent, FONT, FONTB, BG):
        from tksheet import Sheet
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _HEADERS = ["序號", "標題", "描述"]
        _EMPTY   = ["", "", ""]

        # ── 頂部：來源 Excel（左）+ 使用說明（右）──────────
        top_row = tk.Frame(parent, bg=BG)
        top_row.pack(fill="x", padx=12, pady=(12, 4))

        src_frame = tk.LabelFrame(top_row, text="來源 Excel", bg=BG, font=FONTB)
        src_frame.pack(side="left", fill="y", padx=(0, 8))
        src_frame.columnconfigure(1, weight=1)

        hint_frame = tk.LabelFrame(top_row, text="使用說明", bg=BG, font=FONTB)
        hint_frame.pack(side="left", fill="both", expand=True)
        hint_text = (
            "導入 Excel 後即可編輯資料\n"
            "可以選取多筆資料搬移刪除\n"
            "不可跳號選取"
        )
        tk.Label(hint_frame, text=hint_text, bg=BG, font=FONT_S, fg=GRAY,
                 justify="left", anchor="nw").pack(padx=10, pady=8, anchor="nw")

        tk.Label(src_frame, text="檔案路徑：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        path_var = tk.StringVar()
        tk.Entry(src_frame, textvariable=path_var, font=FONT_S
                 ).grid(row=0, column=1, sticky="ew", padx=4, pady=6)

        tk.Label(src_frame, text="工作表：", bg=BG, font=FONT_S, fg=GRAY,
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
            if not p:
                return
            path_var.set(p)
            try:
                names = get_sheet_names(Path(p))
                # 「全部」放在最前面，可一次讀入所有工作表
                all_options = ["全部"] + names
                sheet_cb["values"] = all_options
                sheet_var.set(all_options[0])
            except Exception:
                sheet_cb["values"] = []
                sheet_var.set("")

        tk.Button(src_frame, text="選擇", command=_pick_file,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=6).grid(row=0, column=2, padx=(0, 4), pady=6)

        status_lbl = tk.Label(src_frame, text="", bg=BG, font=FONT_S, fg=GRAY)
        status_lbl.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        # ── 序號篩選列 ────────────────────────────────────
        filter_frame = tk.Frame(parent, bg=BG)
        filter_frame.pack(fill="x", padx=12, pady=(0, 2))
        tk.Label(filter_frame, text="序號篩選：", bg=BG, font=FONT_S, fg=GRAY
                 ).pack(side="left")
        filter_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=filter_var, font=FONT_S, width=30
                 ).pack(side="left", padx=(4, 6))
        tk.Label(filter_frame, text="（逗號分隔，例如 1,3,5；空白=全部顯示）",
                 bg=BG, font=FONT_S, fg=GRAY).pack(side="left")

        # ── 可編輯 Sheet 預覽 ─────────────────────────────
        prev_frame = tk.LabelFrame(parent, text="卡片預覽（雙擊儲存格可編輯，0 筆）",
                                   bg=BG, font=FONT)
        prev_frame.pack(fill="both", expand=True, padx=12, pady=4)

        sheet = Sheet(prev_frame,
                      headers=_HEADERS,
                      data=[_EMPTY[:] for _ in range(10)],
                      column_width=200,
                      row_height=28)
        sheet.set_column_widths([70, 300, 500])
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
            rows = [[c["seq"], c["title"], c["desc"]] for c in filtered]
            while len(rows) < len(filtered) + 5:
                rows.append(_EMPTY[:])
            sheet.data = rows
            prev_frame.config(text=f"卡片預覽（雙擊儲存格可編輯，{len(filtered)} 筆）")

        tk.Button(filter_frame, text="篩選", command=_apply_filter,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=6).pack(side="left")
        tk.Button(filter_frame, text="清除", command=lambda: (filter_var.set(""), _apply_filter()),
                  bg="#5d6d7e", fg="white", relief="flat",
                  font=FONT_S, padx=6).pack(side="left", padx=(4, 0))

        def _load_preview():
            from sync.creator_trello import read_excel_cards
            p = path_var.get().strip()
            if not p:
                messagebox.showwarning("未選擇檔案", "請先選擇 Excel 檔案", parent=parent)
                return
            status_lbl.config(text="讀取中…", fg=GRAY)
            parent.update_idletasks()
            try:
                selected = sheet_var.get()
                _all_data.clear()
                if selected == "全部":
                    # 取所有工作表名稱（排除「全部」偽選項）
                    all_sheets = [v for v in sheet_cb["values"] if v != "全部"]
                    for sname in all_sheets:
                        _all_data.extend(read_excel_cards(Path(p), sheet_name=sname))
                else:
                    _all_data.extend(read_excel_cards(Path(p), sheet_name=selected or None))
                filter_var.set("")
                _apply_filter()
                status_lbl.config(text=f"✔  讀取完成，共 {len(_all_data)} 筆", fg="#1e8449")
            except Exception as e:
                status_lbl.config(text=f"✘  {e}", fg="#c0392b")

        tk.Button(src_frame, text="讀取預覽", command=_load_preview,
                  bg="#117a65", fg="white", relief="flat",
                  font=FONT_S, padx=6).grid(row=0, column=3, padx=(0, 8), pady=6)

        # ── 建立按鈕 ──────────────────────────────────────
        out_label = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY,
                             anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _create():
            from sync.creator_trello import create_cards as trello_create_cards
            rows = sheet.data
            cards = []
            for row in rows:
                title = str(row[1]).strip() if len(row) > 1 else ""
                desc  = str(row[2]).strip() if len(row) > 2 else ""
                if not title:
                    continue
                cards.append({"title": title, "desc": desc})

            if not cards:
                messagebox.showwarning("無資料", "表格中沒有填入標題的列", parent=parent)
                return
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning(
                    "憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證",
                    parent=parent)
                return
            if not messagebox.askyesno(
                    "確認建立",
                    f"即將在「0.待評估」清單建立 {len(cards)} 張卡片，確定繼續？",
                    parent=parent):
                return

            out_label.config(text="建立中…", fg=GRAY)
            parent.update_idletasks()
            try:
                created = trello_create_cards(cards, api_key, token)
                out_label.config(text=f"✔  成功建立 {created} 張卡片", fg="#1e8449")
            except Exception as e:
                out_label.config(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("建立失敗", str(e), parent=parent)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(fill="x", padx=12)
        tk.Button(bb, text="🃏  建立全部卡片 → Trello 0.待評估", command=_create,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 10：下載卡片
    # ════════════════════════════════════════════════════════
    def _build_tab_download_cards(self, parent, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _list_map: dict[str, str] = {}   # list name → list id
        _all_cards: list[dict]    = []   # 目前預覽的卡片資料

        # ── 清單選擇 ──────────────────────────────────────
        top = tk.LabelFrame(parent, text="Trello 清單", bg=BG, font=FONTB)
        top.pack(fill="x", padx=12, pady=(12, 4))
        top.columnconfigure(1, weight=1)

        tk.Label(top, text="選擇清單：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)

        list_var = tk.StringVar()
        list_cb  = ttk.Combobox(top, textvariable=list_var, font=FONT_S,
                                 state="readonly", width=28)
        list_cb.grid(row=0, column=1, sticky="w", padx=4, pady=6)

        status_lbl = tk.Label(top, text="", bg=BG, font=FONT_S, fg=GRAY)
        status_lbl.grid(row=0, column=3, sticky="w", padx=8, pady=6)

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
            if not api_key:
                return
            status_lbl.config(text="抓取中…", fg=GRAY)
            parent.update_idletasks()
            try:
                lists = get_board_lists(api_key, token)
                _list_map.clear()
                _list_map.update({lst["name"]: lst["id"] for lst in lists})
                names = list(_list_map.keys())
                list_cb["values"] = names
                if names:
                    list_var.set(names[0])
                status_lbl.config(text=f"找到 {len(lists)} 個清單", fg="#1e8449")
            except Exception as e:
                status_lbl.config(text=f"✘ {e}", fg="#c0392b")

        def _preview_cards():
            from sync.downloader_trello import get_list_cards
            selected = list_var.get().strip()
            if not selected or selected not in _list_map:
                messagebox.showwarning("未選擇清單", "請先抓取清單並選擇", parent=parent)
                return
            api_key, token = _get_creds()
            if not api_key:
                return
            status_lbl.config(text="載入卡片中…", fg=GRAY)
            parent.update_idletasks()
            try:
                cards = get_list_cards(_list_map[selected], api_key, token)
                _all_cards.clear()
                _all_cards.extend(cards)
                tree.delete(*tree.get_children())
                for card in cards:
                    labels = "、".join(
                        lbl.get("name") or lbl.get("color", "")
                        for lbl in card.get("labels", []))
                    att_count = len(card.get("attachments") or [])
                    tree.insert("", "end", values=(
                        card["name"], labels, att_count))
                tree.selection_set(tree.get_children())   # 預設全選
                status_lbl.config(text=f"找到 {len(cards)} 張卡片", fg="#1e8449")
                prev_frame.config(text=f"卡片預覽（共 {len(cards)} 張，可多選）")
            except Exception as e:
                status_lbl.config(text=f"✘ {e}", fg="#c0392b")

        tk.Button(top, text="抓取清單", command=_fetch_lists,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=6).grid(row=0, column=2, padx=(0, 4), pady=6)
        tk.Button(top, text="預覽卡片", command=_preview_cards,
                  bg="#117a65", fg="white", relief="flat",
                  font=FONT_S, padx=6).grid(row=0, column=3, padx=(0, 8), pady=6)
        # 覆蓋 status_lbl 位置（改放 column=4）
        status_lbl.grid(row=0, column=4, sticky="w", padx=4, pady=6)

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_frame = tk.LabelFrame(parent, text="卡片預覽（共 0 張，可多選）",
                                   bg=BG, font=FONT)
        prev_frame.pack(fill="both", expand=True, padx=12, pady=4)

        cols = ("title", "labels", "att")
        tree = ttk.Treeview(prev_frame, columns=cols, show="headings",
                            selectmode="extended", height=8)
        tree.heading("title",  text="標題")
        tree.heading("labels", text="標籤")
        tree.heading("att",    text="附件")
        tree.column("title",  width=340, anchor="w",      stretch=True)
        tree.column("labels", width=130, anchor="w",      stretch=False)
        tree.column("att",    width=50,  anchor="center", stretch=False)

        vsb = ttk.Scrollbar(prev_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # 全選 / 取消全選
        sel_row = tk.Frame(parent, bg=BG)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        tk.Button(sel_row, text="全選", command=lambda: tree.selection_set(tree.get_children()),
                  bg="#5d6d7e", fg="white", relief="flat", font=FONT_S, padx=8
                  ).pack(side="left", padx=(0, 4))
        tk.Button(sel_row, text="取消全選", command=lambda: tree.selection_remove(tree.get_children()),
                  bg="#5d6d7e", fg="white", relief="flat", font=FONT_S, padx=8
                  ).pack(side="left")

        # ── 輸出路徑 ──────────────────────────────────────
        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(fill="x", padx=12, pady=4)
        tk.Label(pf, text="輸出資料夾：", bg="#e8ecf0", font=FONT_S, fg=GRAY,
                 anchor="w", width=12).pack(side="left", padx=8, pady=5)
        tk.Label(pf, text="（依⚙路徑設定）",
                 bg="#e8ecf0", font=FONT_S, fg=GRAY).pack(side="left", pady=5)

        # ── 狀態與下載按鈕 ────────────────────────────────
        out_label = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY,
                             anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _download():
            from sync.downloader_trello import download_cards as trello_download_cards
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先預覽並選取要下載的卡片", parent=parent)
                return
            api_key, token = _get_creds()
            if not api_key:
                return

            indices  = [tree.index(i) for i in sel_ids]
            selected_cards = [_all_cards[i] for i in indices]
            list_name  = list_var.get().strip()
            output_dir = self._get_path("download_cards_dir") / list_name

            out_label.config(text="下載中…", fg=GRAY)
            parent.update_idletasks()

            def _progress(cur, total, name):
                out_label.config(text=f"下載中… ({cur}/{total}) {name}", fg=GRAY)
                parent.update_idletasks()

            try:
                count, failed = trello_download_cards(
                    selected_cards, output_dir, api_key, token, progress_cb=_progress)
                if failed:
                    out_label.config(
                        text=f"✔  完成 {count} 張，{len(failed)} 個附件失敗", fg="#e67e22")
                    messagebox.showwarning("部分附件失敗",
                        "以下附件下載失敗：\n" + "\n".join(failed), parent=parent)
                else:
                    out_label.config(
                        text=f"✔  成功下載 {count} 張卡片至：{output_dir}", fg="#1e8449")
                if messagebox.askyesno("下載完成",
                        f"共 {count} 張卡片\n\n是否立即開啟資料夾？", parent=parent):
                    os.startfile(str(output_dir))
            except Exception as e:
                out_label.config(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("下載失敗", str(e), parent=parent)

        bb = tk.Frame(parent, bg=BG, pady=6)
        bb.pack(fill="x", padx=12)
        tk.Button(bb, text="⬇  下載選取的卡片", command=_download,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=8).pack(fill="x")
