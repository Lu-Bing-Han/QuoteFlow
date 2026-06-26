"""
mixin_documents.py — 出貨單、驗機單、維修單、維修掛件 頁籤 mixin
"""
import json, os, sys, subprocess
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from datetime import datetime
from pathlib import Path
from ui.app_core import _mk_lf
from ui.mixin_trello import _bind_drag_select


class _DocumentsTab:
    """Mixin providing shipping, inspection, fix, and tag tab builders + callbacks."""

    # ── Tab 1：出貨單 ─────────────────────────────────────────
    def _build_tab_shipping(self, parent, PAD, FONT, FONTB, BG):
        GRAY = "#5d6d7e"

        # 內容包進可捲動容器，避免視窗不夠高時下方內容（例如生成按鈕）被裁掉
        page = ctk.CTkScrollableFrame(parent, fg_color=BG, corner_radius=0)
        page.pack(fill="both", expand=True)
        parent = page

        # ── 資料來源切換 ──────────────────────────────────
        mode_row = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        mode_row.pack(fill="x", padx=12, pady=(10, 0))
        ctk.CTkLabel(mode_row, text="資料來源：", fg_color="transparent",
                      font=FONT, text_color="#2c3e50").pack(side="left", padx=(0, 8))
        mode_var = tk.StringVar(value="quote")

        def _switch_mode():
            if mode_var.get() == "quote":
                trello_frame.pack_forget()
                quote_frame.pack(fill="both", expand=True)
            else:
                quote_frame.pack_forget()
                trello_frame.pack(fill="both", expand=True)

        ctk.CTkRadioButton(mode_row, text="📄 從報價單讀入", variable=mode_var, value="quote",
                            command=_switch_mode, font=FONT,
                            fg_color="#1a5276", hover_color="#2e6da4"
                            ).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(mode_row, text="🃏 從 Trello 抓取", variable=mode_var, value="trello",
                            command=_switch_mode, font=FONT,
                            fg_color="#1a5276", hover_color="#2e6da4"
                            ).pack(side="left")

        # ── 補填欄位（兩種資料來源共用）─────────────────────
        rf_outer, rf = _mk_lf(parent, "補填欄位", BG, FONTB)
        rf_outer.pack(fill="x", padx=12, pady=(6, 4))
        rf.columnconfigure(1, weight=1)
        rf.columnconfigure(3, weight=1)
        SMALL_PAD = {"padx": 6, "pady": 2}

        self._fill_vars = {}
        for label, key, default, row, col in [
            ("出貨日期", "ship_date", datetime.today().strftime("%Y/%m/%d"), 0, 0),
            ("銷貨單號", "sale_no",   "",                                    0, 2),
        ]:
            ctk.CTkLabel(rf, text=label + "：", fg_color="transparent",
                          anchor="w", font=FONT, text_color="#2c3e50"
                          ).grid(row=row, column=col, sticky="w", **SMALL_PAD)
            var = tk.StringVar(value=default)
            ctk.CTkEntry(rf, textvariable=var, font=FONT,
                          corner_radius=4, border_width=1
                          ).grid(row=row, column=col + 1, sticky="ew", **SMALL_PAD)
            self._fill_vars[key] = var

        ctk.CTkLabel(rf, text="附註：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=1, column=0, sticky="w", **SMALL_PAD)
        note_var = tk.StringVar(value="")
        ctk.CTkEntry(rf, textvariable=note_var, font=FONT,
                      corner_radius=4, border_width=1
                      ).grid(row=1, column=1, columnspan=3, sticky="ew", **SMALL_PAD)
        self._fill_vars["note"] = note_var

        ctk.CTkLabel(rf, text="製表人員：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=2, column=0, sticky="w", **SMALL_PAD)
        op_f = ctk.CTkFrame(rf, fg_color="transparent", corner_radius=0)
        op_f.grid(row=2, column=1, sticky="ew", **SMALL_PAD)
        self._operator_var = tk.StringVar()
        self._operator_cb  = ttk.Combobox(op_f, textvariable=self._operator_var,
                                           values=self._config["operators"],
                                           width=10, font=FONT, state="readonly")
        if self._config["operators"]:
            self._operator_var.set(self._config["operators"][0])
        self._operator_cb.pack(side="left")
        ctk.CTkButton(op_f, text="＋", command=self._add_operator,
                       fg_color="#27ae60", hover_color="#1e8449", text_color="white",
                       font=FONT, width=24, height=24, corner_radius=4
                       ).pack(side="left", padx=(4, 0))
        ctk.CTkButton(op_f, text="－", command=self._del_operator,
                       fg_color="#c0392b", hover_color="#a93226", text_color="white",
                       font=FONT, width=24, height=24, corner_radius=4
                       ).pack(side="left", padx=(2, 0))

        ctk.CTkLabel(rf, text="發票方式：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=2, column=2, sticky="w", **SMALL_PAD)
        inv_f = ctk.CTkFrame(rf, fg_color="transparent", corner_radius=0)
        inv_f.grid(row=2, column=3, sticky="w", **SMALL_PAD)
        self._invoice_var = tk.StringVar(value="尚未確認")
        for lbl, val in [("尚未確認", "尚未確認"), ("隨貨", "隨貨"), ("直寄", "直寄")]:
            ctk.CTkRadioButton(inv_f, text=lbl, variable=self._invoice_var,
                                value=val, font=FONT,
                                fg_color="#1a5276", hover_color="#2e6da4"
                                ).pack(side="left", padx=(0, 6))

        # ── 兩種資料來源各自的內容區 ──────────────────────
        content = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        content.pack(fill="both", expand=True)

        trello_frame = ctk.CTkFrame(content, fg_color=BG, corner_radius=0)
        quote_frame  = ctk.CTkFrame(content, fg_color=BG, corner_radius=0)
        self._build_shipping_trello_block(trello_frame, FONTB, BG)
        quote_frame.pack(fill="both", expand=True)   # 預設顯示「從報價單讀入」

        bb = ctk.CTkFrame(quote_frame, fg_color=BG, corner_radius=0)
        bb.pack(side="bottom", fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="＋ 新增", command=self._add_row,
                       fg_color="#27ae60", hover_color="#1e8449", text_color="white",
                       font=FONT, width=90, height=34, corner_radius=6
                       ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(bb, text="－ 刪除", command=self._del_row,
                       fg_color="#c0392b", hover_color="#a93226", text_color="white",
                       font=FONT, width=90, height=34, corner_radius=6
                       ).pack(side="left")
        ctk.CTkButton(bb, text="✏ 編輯所選列", command=self._edit_selected_row,
                       fg_color="#d68910", hover_color="#b7770d", text_color="white",
                       font=FONT, width=110, height=34, corner_radius=6
                       ).pack(side="left", padx=(6, 0))
        ctk.CTkButton(bb, text="⬇  生成出貨單", command=self._generate,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 11, "bold"),
                       width=160, height=34, corner_radius=6
                       ).pack(side="right")

        lf_outer, lf = _mk_lf(quote_frame, "從報價單讀入", BG, FONTB)
        lf_outer.pack(fill="x", padx=12, pady=6)
        lf.columnconfigure(1, weight=1)
        lf.columnconfigure(3, weight=1)

        self._read_vars = {}
        for i, (label, key) in enumerate([
            ("客戶名稱", "customer"), ("聯絡電話", "phone"),
            ("聯絡人",   "contact"),  ("地址",     "address"),
            ("報價單號", "quote_no"), ("報價日期", "quote_date"),
        ]):
            r, c = divmod(i, 2)
            ctk.CTkLabel(lf, text=label + "：", fg_color="transparent",
                          anchor="w", font=FONT, text_color="#2c3e50"
                          ).grid(row=r, column=c * 2, sticky="w", **PAD)
            var = tk.StringVar(value="—")
            ctk.CTkEntry(lf, textvariable=var, font=FONT,
                          corner_radius=4, border_width=1
                          ).grid(row=r, column=c * 2 + 1, sticky="ew", **PAD)
            self._read_vars[key] = var

        # 品項列表
        tf_outer, tf = _mk_lf(quote_frame, "品項列表", BG, FONTB)
        tf_outer.pack(fill="both", expand=True, padx=12, pady=4)

        ctk.CTkLabel(tf, text="💡  雙擊儲存格、右鍵點選，或選取列後按下方「✏ 編輯所選列」即可編輯",
                      fg_color="transparent", font=("Microsoft JhengHei UI", 9),
                      text_color="#d68910", anchor="w"
                      ).pack(fill="x", padx=4, pady=(2, 4))

        tree_area = ctk.CTkFrame(tf, fg_color="transparent", corner_radius=0)
        tree_area.pack(fill="both", expand=True)

        cols     = ("seq", "name", "qty", "unit", "unit_price", "subtotal", "part_no")
        col_lbls = ("序號", "品名 / 規格", "數量", "單位", "單價", "小計", "品號")
        col_ws   = (45, 330, 65, 65, 85, 85, 0)

        self._tree = ttk.Treeview(tree_area, columns=cols, show="headings",
                                   selectmode="browse", height=8)
        self._tree["displaycolumns"] = cols[:-1]   # part_no 僅供內部使用，不顯示
        for col, lbl, w in zip(cols, col_lbls, col_ws):
            self._tree.heading(col, text=lbl)
            self._tree.column(col, width=w, minwidth=w, anchor="center")
        self._tree.column("name", anchor="w")

        vsb = ttk.Scrollbar(tree_area, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._tree.bind("<Double-1>", self._on_cell_dclick)
        self._tree.bind("<Button-3>", self._on_cell_rclick)

    # ── Tab 1 附加區塊：或從 Trello 抓取 ─────────────────────
    def _build_shipping_trello_block(self, parent, FONTB, BG):
        GRAY   = "#5d6d7e"
        GREEN  = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)

        # board_name, list_name, resolve_order_date_via_api
        # 「已報價」清單卡片數量極多、且大多還沒下單日期可言，關掉逐卡 API 日期判斷避免抓取過慢
        _BOARD_SOURCES = {
            "物流事業部1 — 2.已報價(Quoted)": ("物流事業部1", "已報價",   False),
            "物流事業部1 — 本周下單(PO)":     ("物流事業部1", "本周下單", True),
        }
        _all_cards:       list[dict] = []
        _displayed_cards: list[dict] = []

        tr_outer, tr_f = _mk_lf(parent, "或從 Trello 抓取（每張卡片各產出一份出貨單）", BG, FONTB)
        tr_outer.pack(fill="both", expand=True, padx=12, pady=(8, 4))

        fetch_row = tk.Frame(tr_f, bg=BG)
        fetch_row.pack(fill="x", padx=4, pady=(2, 2))

        ctk.CTkLabel(fetch_row, text="來源：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(0, 2))
        source_var = tk.StringVar(value=list(_BOARD_SOURCES.keys())[0])
        ctk.CTkOptionMenu(
            fetch_row, variable=source_var,
            values=list(_BOARD_SOURCES.keys()),
            font=FONT_S, width=210, height=28, corner_radius=4,
        ).pack(side="left", padx=(0, 8))

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
                    c["created_date"], c["company"], c["product"],
                    c["quantity"], c.get("amount", "")))
            tree.selection_set(tree.get_children())
            prev_title_lbl.configure(text=f"  卡片（共 {len(cards)} 張，可多選）  ")

        def _apply_days_filter():
            from_date = date_entry.get_date()
            filtered = [c for c in _all_cards
                        if c.get("created_dt") and c["created_dt"] >= from_date]
            _render_tree(filtered)
            fetch_status.configure(
                text=f"✔  共 {len(_all_cards)} 張，篩選後 {len(filtered)} 張", text_color=GREEN)

        def _show_all():
            _render_tree(_all_cards)
            fetch_status.configure(text=f"✔  共 {len(_all_cards)} 張（全部顯示）", text_color=GREEN)

        _FETCH_TIMEOUT_MS = 45_000   # 背景執行緒卡住（極端網路狀況）時，逾時後仍讓使用者能重試

        def _fetch():
            import threading
            from sync.syncer_trello import fetch_po_cards, update_location_cache
            from _paths import _LOCATION_CACHE_PATH
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未填",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent); return

            board_name, list_name, resolve_via_api = _BOARD_SOURCES[source_var.get()]

            fetch_btn.configure(state="disabled")
            fetch_status.configure(text="抓取中…", text_color=GRAY)
            result: dict = {}

            def _worker(bn=board_name, ln=list_name, rv=resolve_via_api):
                try:
                    cards = fetch_po_cards(api_key, token, board_name=bn, list_name=ln,
                                            resolve_order_date_via_api=rv)
                    update_location_cache(cards, _LOCATION_CACHE_PATH)
                    result["cards"] = cards
                except Exception as e:
                    result["error"] = e
                finally:
                    result["done"] = True

            threading.Thread(target=_worker, daemon=True).start()

            def _poll(elapsed=0):
                if result.get("done"):
                    fetch_btn.configure(state="normal")
                    if "error" in result:
                        from sync.syncer_trello import _fmt_trello_error
                        fetch_status.configure(text=f"✘  {_fmt_trello_error(result['error'])}",
                                                text_color="#c0392b")
                    else:
                        cards = result["cards"]
                        _all_cards.clear(); _all_cards.extend(cards); _apply_days_filter()
                    return
                if elapsed >= _FETCH_TIMEOUT_MS:
                    # 逾時：放棄等待並重新啟用按鈕，讓使用者可以直接重試（不需要重開程式）；
                    # 背景執行緒仍會自己跑完，只是結果不再被採用。
                    fetch_btn.configure(state="normal")
                    fetch_status.configure(text="✘  抓取逾時，請檢查網路連線後點「抓取卡片」重試",
                                            text_color="#c0392b")
                    return
                parent.after(300, lambda: _poll(elapsed + 300))

            _poll()

        fetch_btn = ctk.CTkButton(fetch_row, text="🔄 抓取卡片", command=_fetch,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=90, height=28, corner_radius=4)
        fetch_btn.pack(side="left")

        # 獨立成一整列並設定 wraplength，避免錯誤訊息過長時把上面的按鈕擠出畫面外
        fetch_status = ctk.CTkLabel(tr_f, text="", fg_color="transparent",
                                     font=FONT_S, text_color=GRAY,
                                     anchor="w", justify="left", wraplength=600)
        fetch_status.pack(fill="x", padx=4, pady=(2, 2))

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_title_lbl = ctk.CTkLabel(tr_f, text="  卡片（共 0 張，可多選）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        prev_title_lbl.pack(anchor="w", padx=4, pady=(4, 0))

        tree_frame = tk.Frame(tr_f, bg=BG)
        tree_frame.pack(fill="x", padx=4, pady=(0, 4))

        cols = ("date", "company", "product", "qty", "amount")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                             selectmode="extended", height=5)
        tree.heading("date",    text="下單日期")
        tree.heading("company", text="客戶名稱")
        tree.heading("product", text="品號 / 產品")
        tree.heading("qty",     text="數量")
        tree.heading("amount",  text="應收總金額")
        tree.column("date",    width=70,  anchor="center", stretch=False)
        tree.column("company", width=160, anchor="w",      stretch=False)
        tree.column("product", width=300, anchor="w",      stretch=True)
        tree.column("qty",     width=50,  anchor="center", stretch=False)
        tree.column("amount",  width=90,  anchor="e",      stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        _bind_drag_select(tree)

        sel_row = tk.Frame(tr_f, bg=BG)
        sel_row.pack(fill="x", padx=4, pady=(0, 2))
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

        def _load_location_lookup() -> dict:
            from _paths import _LOCATION_CACHE_PATH
            if _LOCATION_CACHE_PATH.exists():
                try:
                    return json.loads(_LOCATION_CACHE_PATH.read_text(encoding="utf-8"))
                except Exception:
                    pass
            return {}

        # ── 預覽（依目前選取的卡片，即時更新；可直接編輯後再產生）─────
        _overrides: dict[str, dict] = {}   # card_id -> 編輯過的 {"header":..., "items":...}
        _current_card_id = [None]
        _suppress_trace  = [False]

        preview_outer, preview_f = _mk_lf(
            tr_f, "預覽（點選卡片即時更新，可直接編輯，編輯結果會套用到「產生出貨單」）", BG, FONTB)
        preview_outer.pack(fill="both", expand=True, padx=4, pady=(4, 4))

        preview_hdr = ctk.CTkFrame(preview_f, fg_color=BG, corner_radius=0)
        preview_hdr.pack(fill="x", padx=4, pady=(2, 2))
        preview_hdr.columnconfigure(1, weight=1)
        preview_hdr.columnconfigure(3, weight=1)

        def _save_override(*_args):
            if _suppress_trace[0]:
                return
            card_id = _current_card_id[0]
            if not card_id:
                return
            header = {key: ("" if var.get() == "—" else var.get())
                      for key, var in _preview_vars.items()}
            items = []
            for iid in preview_tree.get_children():
                v = preview_tree.item(iid, "values")
                items.append({"seq": self._to_num(v[0]), "name": v[1],
                              "qty": self._to_num(v[2], default=v[2]), "unit": v[3],
                              "unit_price": self._to_num(v[4]), "subtotal": self._to_num(v[5]),
                              "part_no": ""})
            _overrides[card_id] = {"header": header, "items": items}

        _preview_vars = {}
        _PREVIEW_FIELDS = [
            ("客戶名稱", "customer"), ("聯絡電話", "phone"),
            ("聯絡人",   "contact"),  ("地址",     "address"),
            ("統一編號", "tax_id"),
        ]
        for i, (label, key) in enumerate(_PREVIEW_FIELDS):
            r, c = divmod(i, 2)
            ctk.CTkLabel(preview_hdr, text=label + "：", fg_color="transparent",
                          anchor="w", font=FONT_S, text_color="#2c3e50"
                          ).grid(row=r, column=c * 2, sticky="w", padx=(0, 4), pady=2)
            var = tk.StringVar(value="—")
            var.trace_add("write", _save_override)
            ctk.CTkEntry(preview_hdr, textvariable=var, font=FONT_S,
                          corner_radius=4, border_width=1
                          ).grid(row=r, column=c * 2 + 1, sticky="ew", padx=(0, 12), pady=2)
            _preview_vars[key] = var

        preview_tree_frame = tk.Frame(preview_f, bg=BG)
        preview_tree_frame.pack(fill="both", expand=True, padx=4, pady=(2, 4))

        p_cols   = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
        p_lbls   = ("序號", "品名 / 規格", "數量", "單位", "單價", "小計")
        p_widths = (45, 280, 60, 60, 80, 90)
        preview_tree = ttk.Treeview(preview_tree_frame, columns=p_cols, show="headings", height=4)
        for col, lbl, w in zip(p_cols, p_lbls, p_widths):
            preview_tree.heading(col, text=lbl)
            preview_tree.column(col, width=w, minwidth=w, anchor="center")
        preview_tree.column("name", anchor="w")

        preview_vsb = ttk.Scrollbar(preview_tree_frame, orient="vertical", command=preview_tree.yview)
        preview_tree.configure(yscrollcommand=preview_vsb.set)
        preview_tree.pack(side="left", fill="both", expand=True)
        preview_vsb.pack(side="right", fill="y")

        ctk.CTkLabel(preview_f, text="💡  雙擊儲存格、右鍵點選，或選取列後按下方「✏ 編輯所選列」即可編輯",
                      fg_color="transparent", font=("Microsoft JhengHei UI", 9),
                      text_color="#d68910", anchor="w"
                      ).pack(fill="x", padx=4, pady=(0, 2))

        preview_sel_row = tk.Frame(preview_f, bg=BG)
        preview_sel_row.pack(fill="x", padx=4, pady=(0, 4))
        ctk.CTkButton(preview_sel_row, text="✏ 編輯所選列", command=lambda: _edit_preview_selected(),
                       fg_color="#d68910", hover_color="#b7770d", text_color="white",
                       font=FONT_S, width=90, height=26, corner_radius=4
                       ).pack(side="left")

        def _open_preview_cell_editor(item_id, col_id):
            col_idx  = int(col_id.replace("#", "")) - 1
            col_keys = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
            col_disp = ("序號", "品名", "數量", "單位", "單價", "小計")
            old_val  = preview_tree.item(item_id, "values")[col_idx]

            bbox = preview_tree.bbox(item_id, col_id)
            if not bbox:
                return
            x, y, _, h = bbox

            pop = ctk.CTkToplevel(parent)
            pop.title(f"編輯「{col_disp[col_idx]}」")
            pop.geometry(f"300x80+{parent.winfo_rootx()+x}+{parent.winfo_rooty()+y+h}")
            pop.after(100, pop.grab_set)

            var   = tk.StringVar(value=old_val)
            entry = ctk.CTkEntry(pop, textvariable=var,
                                  font=("Microsoft JhengHei UI", 11),
                                  corner_radius=4, border_width=1)
            entry.pack(fill="x", padx=10, pady=8)
            entry.focus()

            def save(_=None):
                vals = list(preview_tree.item(item_id, "values"))
                new  = var.get()
                col_name = col_keys[col_idx]
                if col_name in ("seq", "qty", "unit_price", "subtotal"):
                    if new.strip():
                        try:
                            new = float(new) if "." in new else int(new)
                        except ValueError:
                            messagebox.showwarning("格式錯誤", "此欄位請輸入數字", parent=pop)
                            return
                vals[col_idx] = new
                if col_name in ("qty", "unit_price"):
                    try:
                        vals[5] = round(float(vals[2]) * float(vals[4]), 2)
                    except Exception:
                        pass
                preview_tree.item(item_id, values=vals)
                _save_override()
                pop.destroy()

            entry.bind("<Return>", save)
            ctk.CTkButton(pop, text="確認", command=save,
                           fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                           font=("Microsoft JhengHei UI", 10),
                           width=80, height=28, corner_radius=4).pack(pady=2)

        def _on_preview_dclick(event):
            item_id = preview_tree.identify_row(event.y)
            col_id  = preview_tree.identify_column(event.x)
            if not item_id or not col_id:
                return
            _open_preview_cell_editor(item_id, col_id)

        def _on_preview_rclick(event):
            item_id = preview_tree.identify_row(event.y)
            col_id  = preview_tree.identify_column(event.x)
            if not item_id or not col_id:
                return
            preview_tree.selection_set(item_id)
            col_idx  = int(col_id.replace("#", "")) - 1
            col_disp = ("序號", "品名", "數量", "單位", "單價", "小計")
            menu = tk.Menu(parent, tearoff=0)
            menu.add_command(label=f"✏  編輯「{col_disp[col_idx]}」",
                              command=lambda: _open_preview_cell_editor(item_id, col_id))
            menu.tk_popup(event.x_root, event.y_root)

        def _edit_preview_selected():
            sel = preview_tree.selection()
            if not sel:
                messagebox.showinfo("尚未選取", "請先在表格中點選一列，再按「✏ 編輯所選列」", parent=parent)
                return
            _open_preview_cell_editor(sel[0], "#2")   # 預設開啟「品名 / 規格」欄

        preview_tree.bind("<Double-1>", _on_preview_dclick)
        preview_tree.bind("<Button-3>", _on_preview_rclick)

        def _update_preview(_event=None):
            sel = tree.selection()
            if not sel:
                _current_card_id[0] = None
                for var in _preview_vars.values():
                    var.set("—")
                preview_tree.delete(*preview_tree.get_children())
                return
            focus = tree.focus()
            idx = tree.index(focus) if focus in sel else tree.index(sel[-1])
            card = _displayed_cards[idx]
            card_id = card.get("card_id")
            _current_card_id[0] = card_id

            from core.generator import trello_card_to_data
            data = _overrides.get(card_id) or trello_card_to_data(card, _load_location_lookup())

            _suppress_trace[0] = True
            try:
                h = data["header"]
                for key, var in _preview_vars.items():
                    var.set(h.get(key, "") or "—")

                preview_tree.delete(*preview_tree.get_children())
                for item in data["items"]:
                    preview_tree.insert("", "end", values=(
                        item["seq"], item["name"], item["qty"], item["unit"],
                        item["unit_price"], item["subtotal"]))
            finally:
                _suppress_trace[0] = False

        tree.bind("<<TreeviewSelect>>", _update_preview)

        def _generate_from_trello():
            from core.generator import generate, trello_card_to_data
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先選取要產出出貨單的卡片", parent=parent); return

            indices  = [tree.index(i) for i in sel_ids]
            selected = [_displayed_cards[i] for i in indices]
            location_lookup = _load_location_lookup()

            extra = {
                "ship_date":      self._fill_vars["ship_date"].get(),
                "sale_no":        self._fill_vars["sale_no"].get(),
                "note":           self._fill_vars["note"].get(),
                "operator":       self._operator_var.get(),
                "invoice_choice": self._invoice_var.get(),
            }

            def _worker():
                paths = []
                for c in selected:
                    data   = _overrides.get(c.get("card_id")) or trello_card_to_data(c, location_lookup)
                    result = generate(data, extra, output_dir=self._OUT_SHIPPING)
                    paths.extend(result if isinstance(result, list) else [result])
                return paths

            def _on_done(paths):
                msg = "\n".join(str(p) for p in paths)
                if messagebox.askyesno("生成成功",
                        f"已生成 {len(paths)} 份出貨單：\n{msg}\n\n是否立即開啟？", parent=parent):
                    for p in paths:
                        os.startfile(p) if sys.platform == "win32" else subprocess.run(["open", str(p)])

            self._run_task(_worker,
                            buttons=[gen_btn],
                            status_label=fetch_status,
                            loading_text="生成中…",
                            success_text=lambda paths: f"✔  已生成 {len(paths)} 份出貨單",
                            on_success=_on_done)

        gen_btn = ctk.CTkButton(tr_f, text="📤  產生出貨單（選取的卡片，每張各一份）",
                       command=_generate_from_trello,
                       fg_color="#1e8449", hover_color="#196f3d", text_color="white",
                       font=("Microsoft JhengHei UI", 11, "bold"),
                       height=36, corner_radius=6)
        gen_btn.pack(side="bottom", fill="x", padx=4, pady=(2, 6))

    # ── Tab 2：驗機單 ─────────────────────────────────────────
    def _build_tab_inspection(self, parent, PAD, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        info_outer, info = _mk_lf(parent, "說明", BG, FONTB)
        info_outer.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(info,
                      text="載入報價單後，點擊下方按鈕自動生成驗機單 Excel 及 Word。",
                      fg_color="transparent", font=FONT, text_color=GRAY
                      ).pack(padx=12, pady=8, anchor="w")

        # ── 附加選項 ────────────────────────────────────────
        opts_outer, opts = _mk_lf(parent, "附加選項（Word）", BG, FONTB)
        opts_outer.pack(fill="x", padx=12, pady=4)

        self._insp_vars = {}
        preview_var = tk.StringVar(value="（未勾選）")

        def _update_preview(*_):
            lines = []
            acc = []
            if self._insp_vars["電線"].get():   acc.append("□電線")
            if self._insp_vars["充電器"].get(): acc.append("□充電器")
            if acc:
                lines.append("附配件 " + "/".join(acc))
            if self._insp_vars["把手拆折"].get():
                lines.append("□把手拆折")
            if self._insp_vars["腳踏拆"].get():
                lines.append("□腳踏拆")
            preview_var.set("\n".join(lines) if lines else "（未勾選）")

        row0 = ctk.CTkFrame(opts, fg_color="transparent")
        row0.pack(anchor="w", padx=8, pady=(4, 2))
        for key, label in [("把手拆折", "□把手拆折"), ("腳踏拆", "□腳踏拆")]:
            var = tk.BooleanVar()
            self._insp_vars[key] = var
            ctk.CTkCheckBox(row0, text=label, variable=var, command=_update_preview,
                             font=FONT, text_color="#2c3e50",
                             checkbox_width=18, checkbox_height=18
                             ).pack(side="left", padx=12)

        row1 = ctk.CTkFrame(opts, fg_color="transparent")
        row1.pack(anchor="w", padx=8, pady=(2, 4))
        ctk.CTkLabel(row1, text="附配件：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(12, 4))
        for key, label in [("電線", "電線"), ("充電器", "充電器")]:
            var = tk.BooleanVar()
            self._insp_vars[key] = var
            ctk.CTkCheckBox(row1, text=label, variable=var, command=_update_preview,
                             font=FONT, text_color="#2c3e50",
                             checkbox_width=18, checkbox_height=18
                             ).pack(side="left", padx=8)

        # ── Word 預覽 ────────────────────────────────────────
        prev_outer, prev = _mk_lf(parent, "Word 預覽", BG, FONTB)
        prev_outer.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(prev, textvariable=preview_var, fg_color="transparent",
                      font=FONT_S, text_color="#6c3483", anchor="w", justify="left"
                      ).pack(padx=12, pady=6, anchor="w")

        pf = ctk.CTkFrame(parent, fg_color="#e8ecf0", corner_radius=6)
        pf.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(pf, text="輸出路徑：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY, anchor="w", width=96
                      ).pack(side="left", padx=8, pady=6)
        ctk.CTkLabel(pf, text="（依⚙路徑設定）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left", pady=6)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(side="bottom", fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🔍  生成驗機單", command=self._generate_inspection,
                       fg_color="#6c3483", hover_color="#5b2c6f", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ── Tab 3：維修單 ─────────────────────────────────────────
    def _build_tab_fix(self, parent, PAD, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        info_outer, info = _mk_lf(parent, "說明", BG, FONTB)
        info_outer.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(info,
                      text="載入報價單後，點擊下方按鈕生成維修單。",
                      fg_color="transparent", font=FONT, text_color=GRAY
                      ).pack(padx=12, pady=8, anchor="w")

        pf = ctk.CTkFrame(parent, fg_color="#e8ecf0", corner_radius=6)
        pf.pack(fill="x", padx=12, pady=4)
        row = ctk.CTkFrame(pf, fg_color="transparent", corner_radius=0)
        row.pack(fill="x")
        ctk.CTkLabel(row, text="輸出路徑：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY, anchor="w", width=80
                      ).pack(side="left", padx=8)
        ctk.CTkLabel(row, text="（依⚙路徑設定）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY
                      ).pack(side="left")

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(side="bottom", fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🔧  生成維修單", command=self._generate_fix,
                       fg_color="#d68910", hover_color="#b7770d", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ── Tab 4：維修掛件 ───────────────────────────────────────
    def _build_tab_tag(self, parent, PAD, FONT, FONTB, BG):
        from tkcalendar import DateEntry
        GRAY = "#5d6d7e"

        tgf_outer, tgf = _mk_lf(parent, "維修掛件資料", BG, FONTB)
        tgf_outer.pack(fill="x", padx=12, pady=(12, 4))
        tgf.columnconfigure(1, weight=1)
        tgf.columnconfigure(3, weight=1)

        self._tag_vars = {}

        cust_var = tk.StringVar()
        self._tag_vars["customer"] = cust_var
        ctk.CTkLabel(tgf, text="客戶名稱：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=0, column=0, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(tgf, textvariable=cust_var, font=FONT,
                      corner_radius=4, border_width=1
                      ).grid(row=0, column=1, sticky="ew", padx=8, pady=2)

        def _load_customer():
            if self._parsed_data:
                cust_var.set(self._parsed_data["header"].get("customer", ""))
                part_nos = [item.get("part_no", "") or item.get("name", "")
                            for item in self._parsed_data.get("items", [])]
                self._tag_partno_cb["values"] = part_nos
                if part_nos:
                    self._tag_partno_var.set(part_nos[0])
            else:
                messagebox.showwarning("尚未載入", "請先選擇並載入報價單", parent=parent)

        ctk.CTkButton(tgf, text="從報價單帶入", command=_load_customer,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT, width=110, height=28, corner_radius=4
                       ).grid(row=0, column=2, padx=8, pady=2)

        no_var = tk.StringVar(value="1")
        self._tag_vars["no"] = no_var
        ctk.CTkLabel(tgf, text="No.：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=1, column=0, sticky="w", padx=8, pady=2)
        ttk.Combobox(tgf, textvariable=no_var,
                     values=[str(i) for i in range(1, 21)],
                     width=8, font=FONT).grid(row=1, column=1, sticky="w", padx=8, pady=2)

        self._tag_partno_var = tk.StringVar()
        self._tag_vars["part_no"] = self._tag_partno_var
        ctk.CTkLabel(tgf, text="品號：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=1, column=2, sticky="w", padx=8, pady=2)
        self._tag_partno_cb = ttk.Combobox(tgf, textvariable=self._tag_partno_var,
                                            font=FONT, width=20)
        self._tag_partno_cb.grid(row=1, column=3, sticky="ew", padx=8, pady=2)

        seq_var = tk.StringVar()
        self._tag_vars["seq_no"] = seq_var
        ctk.CTkLabel(tgf, text="序號：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=2, column=0, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(tgf, textvariable=seq_var, font=FONT,
                      corner_radius=4, border_width=1
                      ).grid(row=2, column=1, sticky="ew", padx=8, pady=2)

        ctk.CTkLabel(tgf, text="拉回：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=2, column=2, sticky="w", padx=8, pady=2)
        self._tag_date_entry = DateEntry(
            tgf, font=FONT, date_pattern="yyyy/mm/dd",
            background="#2e86c1", foreground="white", width=14)
        self._tag_date_entry.grid(row=2, column=3, sticky="w", padx=8, pady=2)

        prob_var = tk.StringVar()
        self._tag_vars["problem"] = prob_var
        ctk.CTkLabel(tgf, text="問題：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=3, column=0, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(tgf, textvariable=prob_var, font=FONT,
                      corner_radius=4, border_width=1
                      ).grid(row=3, column=1, sticky="ew", padx=8, pady=2)

        status_var = tk.StringVar()
        self._tag_vars["repair_status"] = status_var
        ctk.CTkLabel(tgf, text="維修狀況：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=3, column=2, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(tgf, textvariable=status_var, font=FONT,
                      corner_radius=4, border_width=1
                      ).grid(row=3, column=3, sticky="ew", padx=8, pady=2)

        pf = ctk.CTkFrame(parent, fg_color="#e8ecf0", corner_radius=6)
        pf.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(pf, text="輸出路徑：", fg_color="transparent",
                      font=("Microsoft JhengHei UI", 9), text_color=GRAY,
                      anchor="w", width=80).pack(side="left", padx=8, pady=6)
        ctk.CTkLabel(pf, text=r"Z:\待維修機台資料",
                      fg_color="transparent",
                      font=("Microsoft JhengHei UI", 9), text_color=GRAY
                      ).pack(side="left", pady=6)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(side="bottom", fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="📋  生成維修掛件", command=self._generate_tag_doc,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ── Callbacks ─────────────────────────────────────────────

    def _on_cell_dclick(self, event):
        item_id = self._tree.identify_row(event.y)
        col_id  = self._tree.identify_column(event.x)
        if not item_id or not col_id:
            return
        self._open_cell_editor(item_id, col_id)

    def _on_cell_rclick(self, event):
        item_id = self._tree.identify_row(event.y)
        col_id  = self._tree.identify_column(event.x)
        if not item_id or not col_id:
            return
        self._tree.selection_set(item_id)
        col_idx  = int(col_id.replace("#", "")) - 1
        col_disp = ("序號", "品名", "數量", "單位", "單價", "小計")
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=f"✏  編輯「{col_disp[col_idx]}」",
                          command=lambda: self._open_cell_editor(item_id, col_id))
        menu.tk_popup(event.x_root, event.y_root)

    def _open_cell_editor(self, item_id, col_id):
        col_idx  = int(col_id.replace("#", "")) - 1
        col_keys = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
        col_disp = ("序號", "品名", "數量", "單位", "單價", "小計")
        old_val  = self._tree.item(item_id, "values")[col_idx]

        bbox = self._tree.bbox(item_id, col_id)
        if not bbox:
            return
        x, y, _, h = bbox

        pop = ctk.CTkToplevel(self)
        pop.title(f"編輯「{col_disp[col_idx]}」")
        pop.geometry(f"300x80+{self.winfo_rootx()+x}+{self.winfo_rooty()+y+h}")
        pop.after(100, pop.grab_set)

        var   = tk.StringVar(value=old_val)
        entry = ctk.CTkEntry(pop, textvariable=var,
                              font=("Microsoft JhengHei UI", 11),
                              corner_radius=4, border_width=1)
        entry.pack(fill="x", padx=10, pady=8)
        entry.focus()

        def save(_=None):
            vals = list(self._tree.item(item_id, "values"))
            new  = var.get()
            col_name = col_keys[col_idx]
            if col_name in ("seq", "qty", "unit_price", "subtotal"):
                if new.strip():
                    try:
                        new = float(new) if "." in new else int(new)
                    except ValueError:
                        messagebox.showwarning("格式錯誤", "此欄位請輸入數字", parent=pop)
                        return
            vals[col_idx] = new
            if col_name in ("qty", "unit_price"):
                try:
                    vals[5] = round(float(vals[2]) * float(vals[4]), 2)
                except Exception:
                    pass
            self._tree.item(item_id, values=vals)
            pop.destroy()

        entry.bind("<Return>", save)
        ctk.CTkButton(pop, text="確認", command=save,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=("Microsoft JhengHei UI", 10),
                       width=80, height=28, corner_radius=4).pack(pady=2)

    def _edit_selected_row(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("尚未選取", "請先在表格中點選一列，再按「✏ 編輯所選列」")
            return
        self._open_cell_editor(sel[0], "#2")   # 預設開啟「品名 / 規格」欄

    def _add_row(self):
        n = len(self._tree.get_children()) + 1
        self._tree.insert("", "end", values=(n, "新品項", 1, "組", 0, 0, ""))

    def _del_row(self):
        sel = self._tree.selection()
        if not sel:
            return
        self._tree.delete(sel[0])
        for i, rid in enumerate(self._tree.get_children()):
            v = list(self._tree.item(rid, "values"))
            v[0] = i + 1
            self._tree.item(rid, values=v)

    def _add_operator(self):
        pop = ctk.CTkToplevel(self)
        pop.title("新增製表人員")
        pop.geometry("260x80")
        pop.after(100, pop.grab_set)
        var   = tk.StringVar()
        entry = ctk.CTkEntry(pop, textvariable=var,
                              font=("Microsoft JhengHei UI", 11),
                              corner_radius=4, border_width=1)
        entry.pack(fill="x", padx=10, pady=8)
        entry.focus()

        def save(_=None):
            name = var.get().strip()
            if not name:
                return
            if name not in self._config["operators"]:
                self._config["operators"].append(name)
                self._save_config(self._config)
                self._operator_cb["values"] = self._config["operators"]
            self._operator_var.set(name)
            pop.destroy()

        entry.bind("<Return>", save)
        ctk.CTkButton(pop, text="新增", command=save,
                       fg_color="#27ae60", hover_color="#1e8449", text_color="white",
                       font=("Microsoft JhengHei UI", 10),
                       width=80, height=28, corner_radius=4).pack(pady=2)

    def _del_operator(self):
        cur = self._operator_var.get()
        if not cur:
            return
        if len(self._config["operators"]) <= 1:
            messagebox.showwarning("無法刪除", "至少要保留一位製表人員")
            return
        if messagebox.askyesno("確認刪除", f"刪除「{cur}」？"):
            self._config["operators"].remove(cur)
            self._save_config(self._config)
            self._operator_cb["values"] = self._config["operators"]
            self._operator_var.set(self._config["operators"][0])

    def _sync_header(self):
        for key, var in self._read_vars.items():
            val = var.get()
            self._parsed_data["header"][key] = "" if val == "—" else val

    @staticmethod
    def _to_num(s, default=0):
        try:
            return float(s) if "." in str(s) else int(s)
        except (ValueError, TypeError):
            return default

    def _generate(self):
        from core.generator import generate
        if not self._parsed_data:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return
        items = []
        for i, rid in enumerate(self._tree.get_children()):
            v = self._tree.item(rid, "values")
            items.append({"seq": i+1, "name": v[1],
                          "qty": self._to_num(v[2]), "unit": v[3],
                          "unit_price": self._to_num(v[4]),
                          "subtotal":   self._to_num(v[5]),
                          "part_no":    v[6] if len(v) > 6 else ""})
        self._parsed_data["items"] = items
        self._sync_header()
        extra = {
            "ship_date":      self._fill_vars["ship_date"].get(),
            "sale_no":        self._fill_vars["sale_no"].get(),
            "note":           self._fill_vars["note"].get(),
            "operator":       self._operator_var.get(),
            "invoice_choice": self._invoice_var.get(),
        }
        try:
            result = generate(self._parsed_data, extra, output_dir=self._OUT_SHIPPING)
            paths  = result if isinstance(result, list) else [result]
            msg    = "\n".join(str(p) for p in paths)
            self._set_status(f"出貨單已生成（{len(paths)} 份）", ok=True)
            if messagebox.askyesno("生成成功",
                    f"已生成 {len(paths)} 份出貨單：\n{msg}\n\n是否立即開啟？"):
                for p in paths:
                    os.startfile(p) if sys.platform == "win32" else subprocess.run(["open", str(p)])
        except Exception as e:
            self._set_status(f"出貨單生成失敗：{e}", ok=False)
            messagebox.showerror("生成失敗", str(e))

    def _generate_inspection(self):
        from core.generator_inspection import generate_inspection
        if not self._parsed_data or not self._src_path:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return
        accessories = {k: v.get() for k, v in getattr(self, "_insp_vars", {}).items()}
        try:
            excel_path, word_paths = generate_inspection(
                self._src_path, self._parsed_data,
                output_dir=self._get_path("output_inspection"),
                accessories=accessories)
            msg = f"驗機單 Excel 已儲存至：\n{excel_path}"
            if word_paths:
                msg += f"\n\n驗機單 Word（共 {len(word_paths)} 份）："
                for wp in word_paths:
                    msg += f"\n  {wp.name}"
            self._set_status("驗機單已生成", ok=True)
            if messagebox.askyesno("生成成功", msg + "\n\n是否立即開啟？"):
                os.startfile(excel_path)
                for wp in word_paths:
                    os.startfile(wp)
        except Exception as e:
            self._set_status(f"驗機單生成失敗：{e}", ok=False)
            messagebox.showerror("生成失敗", str(e))

    def _generate_fix(self):
        from core.generator_fix import generate_fix
        if not self._parsed_data:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return
        items = []
        for i, rid in enumerate(self._tree.get_children()):
            v = self._tree.item(rid, "values")
            items.append({"seq": i+1, "name": v[1],
                          "qty": self._to_num(v[2]), "unit": v[3],
                          "unit_price": self._to_num(v[4]),
                          "subtotal":   self._to_num(v[5])})
        self._parsed_data["items"] = items
        self._sync_header()
        extra = {
            "ship_date":      self._fill_vars["ship_date"].get(),
            "sale_no":        self._fill_vars["sale_no"].get(),
            "note":           self._fill_vars["note"].get(),
            "operator":       self._operator_var.get(),
            "invoice_choice": self._invoice_var.get(),
        }
        try:
            result = generate_fix(self._parsed_data, extra, output_dir=self._OUT_SHIPPING)
            paths  = result if isinstance(result, list) else [result]
            msg = "\n".join(str(p) for p in paths)
            self._set_status(f"維修單已生成（{len(paths)} 份）", ok=True)
            if messagebox.askyesno("生成成功",
                    f"已生成 {len(paths)} 份檔案：\n{msg}\n\n是否立即開啟？"):
                for p in paths:
                    os.startfile(p) if sys.platform == "win32" else subprocess.run(["open", str(p)])
        except Exception as e:
            self._set_status(f"維修單生成失敗：{e}", ok=False)
            messagebox.showerror("生成失敗", str(e))

    def _generate_tag_doc(self):
        from core.generator_tag import generate_tag
        customer = self._tag_vars["customer"].get().strip()
        if not customer:
            messagebox.showwarning("客戶名稱未填", "請填入客戶名稱或從報價單帶入")
            return
        tag_data = {
            "no":            self._tag_vars["no"].get(),
            "part_no":       self._tag_vars["part_no"].get(),
            "seq_no":        self._tag_vars["seq_no"].get(),
            "problem":       self._tag_vars["problem"].get(),
            "pullback_date": self._tag_date_entry.get_date().strftime("%Y/%m/%d"),
            "repair_status": self._tag_vars["repair_status"].get(),
        }
        data = {"header": {"customer": customer}}
        try:
            path = generate_tag(data, tag_data, output_dir=self._OUT_TAG)
            self._set_status("維修掛件已生成", ok=True)
            if messagebox.askyesno("生成成功",
                    f"維修掛件已生成：\n{path}\n\n是否立即開啟？"):
                os.startfile(str(path)) if sys.platform == "win32" else subprocess.run(["open", str(path)])
        except Exception as e:
            self._set_status(f"維修掛件生成失敗：{e}", ok=False)
            messagebox.showerror("生成失敗", str(e))
