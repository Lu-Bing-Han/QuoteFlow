"""
mixin_quote.py — 報價單生成頁籤 mixin
"""
import tkinter as tk
from tkinter import messagebox, ttk
import os
from pathlib import Path


class _QuoteTab:
    """Mixin providing _build_tab_quote and all quote-related callbacks."""

    def _build_tab_quote(self, parent, FONT, FONTB, BG):
        import json as _json
        from core.generator_quote import load_product_catalog, generate_quote_from_cart
        from _paths import TEMPLATE_DIR

        GRAY   = "#5d6d7e"
        BLUE   = "#1a4a7a"
        GREEN  = "#1e8449"
        FONT_S = ("Microsoft JhengHei UI", 9)
        MONO   = ("Consolas", 9)
        _TARGET_LIST = "1.待報價(evaluate)"

        # ── 狀態 ────────────────────────────────────────────
        _all_cards:      list[dict]      = []
        _filtered_cards: list[dict]      = []
        _card:           list[dict|None] = [None]   # 選中的 Trello card
        _customer:       dict            = {}        # 解析後客戶資料
        _cart:           dict            = {}        # code → {product, qty, price}
        _cur_cat:        list[str]       = ["all"]
        _products: list[dict] = load_product_catalog(TEMPLATE_DIR)
        CATS = ["全部"] + sorted({p["category"] for p in _products})

        # ════════════════════════════════════════════════════
        # 右側購物車 Panel（固定 272px，所有步驟皆顯示）
        # ════════════════════════════════════════════════════
        body = tk.Frame(parent, bg=BG)
        body.pack(fill="both", expand=True)

        left_area = tk.Frame(body, bg=BG)
        left_area.pack(side="left", fill="both", expand=True)

        right_border = tk.Frame(body, bg="#c8cfd6", width=1)
        right_border.pack(side="right", fill="y")
        right_border.pack_propagate(False)

        cart_panel = tk.Frame(body, bg=BG, width=272)
        cart_panel.pack(side="right", fill="y")
        cart_panel.pack_propagate(False)

        # 購物車 Header
        ch = tk.Frame(cart_panel, bg="#e8ecf0")
        ch.pack(fill="x")
        tk.Label(ch, text="🛒  報價清單", bg="#e8ecf0", fg=BLUE,
                 font=(FONTB[0], 10, "bold"), pady=7).pack(side="left", padx=10)
        cart_cnt_lbl = tk.Label(ch, text="", bg="#e8ecf0", fg=GRAY, font=FONT_S)
        cart_cnt_lbl.pack(side="right", padx=10)

        # 購物車 可捲動 items
        cart_mid = tk.Frame(cart_panel, bg=BG)
        cart_mid.pack(fill="both", expand=True)
        cart_cvs = tk.Canvas(cart_mid, bg=BG, highlightthickness=0)
        cart_sb  = ttk.Scrollbar(cart_mid, orient="vertical", command=cart_cvs.yview)
        cart_cvs.configure(yscrollcommand=cart_sb.set)
        cart_sb.pack(side="right", fill="y")
        cart_cvs.pack(side="left", fill="both", expand=True)
        cart_inner = tk.Frame(cart_cvs, bg=BG)
        _cart_win  = cart_cvs.create_window((0, 0), window=cart_inner, anchor="nw")

        def _on_cart_resize(e=None):
            cart_cvs.configure(scrollregion=cart_cvs.bbox("all"))
            if cart_cvs.winfo_width() > 1:
                cart_cvs.itemconfig(_cart_win, width=cart_cvs.winfo_width())
        cart_inner.bind("<Configure>", _on_cart_resize)
        cart_cvs.bind("<Configure>",   _on_cart_resize)

        def _cart_scroll(e):
            cart_cvs.yview_scroll(int(-1 * (e.delta / 120)), "units")

        def _rebind_cart_scroll(widget=None):
            """購物車重建後遞迴綁定所有子元件。"""
            w = widget or cart_inner
            w.bind("<MouseWheel>", _cart_scroll)
            for child in w.winfo_children():
                _rebind_cart_scroll(child)

        cart_cvs.bind("<MouseWheel>", _cart_scroll)

        # 購物車 Footer
        cf = tk.Frame(cart_panel, bg="#f0f4f8", bd=1, relief="ridge")
        cf.pack(fill="x")
        cart_sub_lbl = tk.Label(cf, text="小計：—", bg="#f0f4f8", fg=GRAY,
                                 font=FONT_S, anchor="e")
        cart_sub_lbl.pack(fill="x", padx=10, pady=(7, 1))
        cart_tax_lbl = tk.Label(cf, text="營業稅 5%：—", bg="#f0f4f8", fg=GRAY,
                                 font=FONT_S, anchor="e")
        cart_tax_lbl.pack(fill="x", padx=10, pady=1)
        cart_tot_lbl = tk.Label(cf, text="合計：—", bg="#f0f4f8",
                                 font=(FONTB[0], 10, "bold"), fg=BLUE, anchor="e")
        cart_tot_lbl.pack(fill="x", padx=10, pady=(1, 4))

        # ── 運費設定 ────────────────────────────────────────
        tk.Frame(cf, bg="#c8cfd6", height=1).pack(fill="x", padx=6)
        ship_outer = tk.Frame(cf, bg="#f0f4f8")
        ship_outer.pack(fill="x", padx=8, pady=(3, 0))

        ship_var       = tk.BooleanVar(value=False)
        ship_price_var = tk.StringVar(value="1500")
        ship_promo_var = tk.BooleanVar(value=False)

        ship_row1 = tk.Frame(ship_outer, bg="#f0f4f8")
        ship_row1.pack(fill="x")
        ship_cb = tk.Checkbutton(ship_row1, text="加入運費", variable=ship_var,
                                  bg="#f0f4f8", font=FONT_S, fg=GRAY,
                                  activebackground="#f0f4f8")
        ship_cb.pack(side="left")
        tk.Label(ship_row1, text="$", bg="#f0f4f8", font=FONT_S, fg=GRAY).pack(side="left")
        ship_price_entry = tk.Entry(ship_row1, textvariable=ship_price_var,
                                     width=7, font=FONT_S, state="disabled")
        ship_price_entry.pack(side="left", padx=(0, 2))

        ship_row2 = tk.Frame(ship_outer, bg="#f0f4f8")
        ship_row2.pack(fill="x", pady=(1, 4))
        ship_promo_cb = tk.Checkbutton(ship_row2, text="顯示免運優惠文字",
                                        variable=ship_promo_var,
                                        bg="#f0f4f8", font=FONT_S, fg="#c0392b",
                                        activebackground="#f0f4f8", state="disabled")
        ship_promo_cb.pack(side="left")

        def _on_ship_toggle(*_):
            state = "normal" if ship_var.get() else "disabled"
            ship_price_entry.config(state=state)
            ship_promo_cb.config(state=state)
            if not ship_var.get():
                ship_promo_var.set(False)
            _update_cart_totals()

        ship_var.trace_add("write", _on_ship_toggle)
        ship_price_var.trace_add("write", lambda *_: _update_cart_totals())
        # ────────────────────────────────────────────────────

        checkout_btn = tk.Button(cf, text="確認報價內容 →", state="disabled",
                                  bg=BLUE, fg="white", relief="flat",
                                  font=FONT_S, pady=5)
        checkout_btn.pack(fill="x", padx=8, pady=(0, 8))

        def _fmt(n): return f"${int(round(n)):,}"

        def _render_cart():
            for w in cart_inner.winfo_children():
                w.destroy()
            items = [(c, v) for c, v in _cart.items() if v["qty"] > 0]
            if not items:
                tk.Label(cart_inner, text="\n📦\n點擊左側產品加入",
                         bg=BG, fg=GRAY, font=FONT_S, justify="center"
                         ).pack(pady=20)
                cart_cnt_lbl.config(text="")
                cart_sub_lbl.config(text="小計：—")
                cart_tax_lbl.config(text="營業稅 5%：—")
                cart_tot_lbl.config(text="合計：—")
                checkout_btn.config(state="disabled")
                return
            total_qty = sum(v["qty"] for _, v in items)
            cart_cnt_lbl.config(text=f"{total_qty} 項")
            for code, item in items:
                if_f = tk.Frame(cart_inner, bg=BG)
                if_f.pack(fill="x", padx=6, pady=3)
                tk.Frame(cart_inner, bg="#dee2e6", height=1).pack(fill="x", padx=4)
                tk.Label(if_f, text=code, bg=BG, fg=GRAY, font=MONO, anchor="w"
                         ).pack(fill="x")
                name_t = item["product"]["name"]
                if len(name_t) > 22: name_t = name_t[:20] + "…"
                tk.Label(if_f, text=name_t, bg=BG, fg="#2c3e50",
                         font=(FONT_S[0], 9, "bold"), anchor="w").pack(fill="x")
                ctrl = tk.Frame(if_f, bg=BG)
                ctrl.pack(fill="x", pady=(2, 0))

                def _dec(c=code):
                    _cart[c]["qty"] -= 1
                    if _cart[c]["qty"] <= 0:
                        del _cart[c]
                    _render_cart()
                    _render_products()

                def _inc(c=code):
                    _cart[c]["qty"] += 1
                    _render_cart()
                    _render_products()

                tk.Button(ctrl, text="−", command=_dec, width=2, font=FONT_S,
                          relief="flat", bg="#dee2e6", fg="#333").pack(side="left")
                tk.Label(ctrl, text=str(item["qty"]), bg=BG, fg="#2c3e50",
                         font=FONT_S, width=3).pack(side="left")
                tk.Button(ctrl, text="＋", command=_inc, width=2, font=FONT_S,
                          relief="flat", bg="#dee2e6", fg="#333").pack(side="left")
                tk.Label(ctrl, text="  ￥", bg=BG, fg=GRAY, font=FONT_S
                         ).pack(side="left")
                pv = tk.StringVar(value=str(int(item["price"])))

                def _price_changed(pvar=pv, c=code, *_):
                    try:
                        _cart[c]["price"] = float(pvar.get().replace(",", ""))
                    except ValueError:
                        pass
                    _update_cart_totals()

                pv.trace_add("write", _price_changed)
                tk.Entry(ctrl, textvariable=pv, width=7, font=FONT_S
                         ).pack(side="left")
                sub = item["qty"] * item["price"]
                tk.Label(if_f, text=f"= {_fmt(sub)}", bg=BG, fg=GRAY,
                         font=FONT_S, anchor="e").pack(fill="x")
            # 購物車重建後綁滾輪
            _rebind_cart_scroll()
            _update_cart_totals()

        def _update_cart_totals():
            items = [v for v in _cart.values() if v["qty"] > 0]
            sub   = sum(v["qty"] * v["price"] for v in items)
            freight = 0
            if ship_var.get():
                try:
                    freight = float(ship_price_var.get().replace(",", ""))
                except ValueError:
                    pass
            tax   = round((sub + freight) * 0.05)
            tot   = sub + freight + tax
            cart_sub_lbl.config(text=f"小計：{_fmt(sub)}" + (f" + 運費 {_fmt(freight)}" if freight else ""))
            cart_tax_lbl.config(text=f"營業稅 5%：{_fmt(tax)}")
            cart_tot_lbl.config(text=f"合計：{_fmt(tot)}")
            checkout_btn.config(state="normal" if items else "disabled")

        # ════════════════════════════════════════════════════
        # Step Bar（步驟進度列）
        # ════════════════════════════════════════════════════
        step_bar = tk.Frame(left_area, bg="#e8ecf0")
        step_bar.pack(fill="x")
        _step_btns: list[tk.Label] = []
        for i, txt in enumerate(["① 顧客資料", "② 選取品項", "③ 報價確認"], 1):
            lbl = tk.Label(step_bar, text=txt, bg="#e8ecf0", fg=GRAY,
                           font=FONT_S, pady=7, cursor="hand2")
            lbl.pack(side="left", expand=True, fill="x")
            lbl.bind("<Button-1>", lambda e, n=i: _go_step(n))
            _step_btns.append(lbl)
            if i < 3:
                tk.Frame(step_bar, bg="#c8cfd6", width=1).pack(side="left", fill="y")

        step_content = tk.Frame(left_area, bg=BG)
        step_content.pack(fill="both", expand=True)

        # ════════════════════════════════════════════════════
        # Step 1 — 顧客資料
        # ════════════════════════════════════════════════════
        p1 = tk.Frame(step_content, bg=BG)

        # Trello 載入列
        p1_top = tk.Frame(p1, bg=BG)
        p1_top.pack(fill="x", padx=10, pady=(8, 2))
        tk.Label(p1_top, text=f"來源：{_TARGET_LIST}",
                 bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        p1_status = tk.Label(p1_top, text="", bg=BG, font=FONT_S, fg=GRAY)
        p1_status.pack(side="left", padx=8)

        def _get_tr_creds():
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent)
                return None, None
            return api_key, token

        def _refresh_card_tree(cards):
            from core.generator_quote import parse_card_desc
            _filtered_cards.clear()
            _filtered_cards.extend(cards)
            card_tree.delete(*card_tree.get_children())
            for card in cards:
                desc    = parse_card_desc(card.get("desc", ""))
                company = desc.get("公司名") or desc.get("公司名稱", "")
                card_tree.insert("", "end", values=(
                    card["name"], company, desc.get("聯絡人", "")))
            if card_tree.get_children():
                card_tree.selection_set(card_tree.get_children()[0])
                _on_card_select()
            shown = len(cards)
            total = len(_all_cards)
            card_lf.config(
                text=f"待報價卡片（共 {total} 張{'，顯示 '+str(shown)+' 張' if shown<total else ''}）")

        def _apply_card_search(*_):
            kw = p1_search_var.get().strip().lower()
            matched = [c for c in _all_cards if kw in c["name"].lower()] if kw else list(_all_cards)
            _refresh_card_tree(matched)

        def _load_cards():
            from sync.downloader_trello import get_board_lists, get_list_cards
            api_key, token = _get_tr_creds()
            if not api_key: return
            p1_status.config(text="載入中…", fg=GRAY)
            parent.update_idletasks()
            try:
                lists  = get_board_lists(api_key, token)
                target = next((l for l in lists
                               if "待報價" in l["name"] or "evaluate" in l["name"].lower()), None)
                if not target:
                    p1_status.config(text=f'✘ 找不到清單', fg="#c0392b"); return
                cards = get_list_cards(target["id"], api_key, token)
                _all_cards.clear(); _all_cards.extend(cards)
                p1_search_var.set("")
                _refresh_card_tree(_all_cards)
                p1_status.config(text=f"✔ 找到 {len(cards)} 張", fg=GREEN)
            except Exception as e:
                p1_status.config(text=f"✘ {e}", fg="#c0392b")

        tk.Button(p1_top, text="🔄 載入", command=_load_cards,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=8).pack(side="right")

        # 搜索列
        p1_search_row = tk.Frame(p1, bg=BG)
        p1_search_row.pack(fill="x", padx=10, pady=(0, 2))
        tk.Label(p1_search_row, text="🔍", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        p1_search_var = tk.StringVar()
        tk.Entry(p1_search_row, textvariable=p1_search_var,
                 font=FONT_S, width=28).pack(side="left", padx=(2, 4))
        tk.Button(p1_search_row, text="清除", command=lambda: p1_search_var.set(""),
                  bg="#5d6d7e", fg="white", relief="flat", font=FONT_S, padx=4
                  ).pack(side="left")
        p1_search_var.trace_add("write", _apply_card_search)

        # 卡片 Treeview
        card_lf = tk.LabelFrame(p1, text="待報價卡片（共 0 張）", bg=BG, font=FONT_S)
        card_lf.pack(fill="both", expand=True, padx=10, pady=2)
        _ccols = ("name", "company", "contact")
        card_tree = ttk.Treeview(card_lf, columns=_ccols, show="headings",
                                  selectmode="browse", height=5)
        card_tree.heading("name",    text="卡片標題")
        card_tree.heading("company", text="公司名")
        card_tree.heading("contact", text="聯絡人")
        card_tree.column("name",    width=230, anchor="w", stretch=True)
        card_tree.column("company", width=130, anchor="w", stretch=False)
        card_tree.column("contact", width=80,  anchor="w", stretch=False)
        ct_vsb = ttk.Scrollbar(card_lf, orient="vertical", command=card_tree.yview)
        card_tree.configure(yscrollcommand=ct_vsb.set)
        card_tree.pack(side="left", fill="both", expand=True)
        ct_vsb.pack(side="right", fill="y")

        # Trello 卡片連結
        _card_url: list[str] = [""]
        trello_link_row = tk.Frame(p1, bg=BG)
        trello_link_row.pack(fill="x", padx=10, pady=(0, 2))
        trello_link_lbl = tk.Label(
            trello_link_row, text="", bg=BG, fg="#1a5276",
            font=(FONT_S[0], 9, "underline"), cursor="hand2", anchor="w")
        trello_link_lbl.pack(side="left")

        def _open_trello_card(e=None):
            import webbrowser as _wb
            if _card_url[0]:
                _wb.open(_card_url[0])

        trello_link_lbl.bind("<Button-1>", _open_trello_card)

        # 報價設定列（放在客戶資料上方，日曆才不會被遮住）
        cfg_f = tk.LabelFrame(p1, text="報價設定", bg=BG, font=FONT_S)
        cfg_f.pack(fill="x", padx=10, pady=(4, 2))

        cfg_row0 = tk.Frame(cfg_f, bg=BG)
        cfg_row0.pack(fill="x", padx=6, pady=(4, 2))
        tk.Label(cfg_row0, text="製表人員：", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        operators = self._config.get("operators", ["小皋"])
        op_var  = tk.StringVar(value=operators[0] if operators else "")
        op_cb   = ttk.Combobox(cfg_row0, textvariable=op_var, values=operators,
                                font=FONT_S, state="readonly", width=10)
        op_cb.pack(side="left", padx=(0, 16))

        from tkcalendar import DateEntry
        tk.Label(cfg_row0, text="報價日期：", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        date_entry = DateEntry(cfg_row0, font=FONT_S, date_pattern="yyyy/mm/dd",
                               width=12, background="#2e86c1", foreground="white",
                               borderwidth=1)
        date_entry.pack(side="left", padx=(0, 6))

        cfg_row1 = tk.Frame(cfg_f, bg=BG)
        cfg_row1.pack(fill="x", padx=6, pady=(2, 2))
        valid_lbl = tk.Label(cfg_row1, text="", bg=BG, font=FONT_S, fg=GRAY)
        valid_lbl.pack(side="left", padx=(0, 16))
        tk.Label(cfg_row1, text="報價單號：", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        quote_no_var = tk.StringVar()
        tk.Entry(cfg_row1, textvariable=quote_no_var, font=FONT_S, width=16
                 ).pack(side="left", padx=(0, 6))

        cfg_row2 = tk.Frame(cfg_f, bg=BG)
        cfg_row2.pack(fill="x", padx=6, pady=(2, 6))
        tk.Label(cfg_row2, text="報價單類型：", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        _QUOTE_TYPE_OPTIONS = ["一般報價單", "補助報價單", "對比報價單"]
        quote_type_var = tk.StringVar(value=_QUOTE_TYPE_OPTIONS[0])
        ttk.Combobox(cfg_row2, textvariable=quote_type_var,
                     values=_QUOTE_TYPE_OPTIONS, state="readonly",
                     width=14, font=FONT_S).pack(side="left", padx=(0, 6))

        def _update_valid_label(*_):
            try:
                from datetime import timedelta as _td
                d  = date_entry.get_date()
                vd = d + _td(days=15)
                valid_lbl.config(text=f"有效日期：{vd.strftime('%Y/%m/%d')}")
                _refresh_quote_no()
            except Exception:
                pass

        def _refresh_quote_no(*_):
            from core.generator_quote import next_quote_no
            try:
                d     = date_entry.get_date()
                op    = op_var.get().strip()
                codes = self._config.get("operator_codes", {})
                code  = codes.get(op, op[:1].upper() if op else "X")
                no    = next_quote_no(self._get_path("output_quote"), code, d)
                quote_no_var.set(no)
            except Exception:
                pass

        date_entry.bind("<<DateEntrySelected>>", _update_valid_label)
        op_var.trace_add("write", _refresh_quote_no)
        _update_valid_label()

        # 客戶資訊預覽區
        cust_lf = tk.LabelFrame(p1, text="客戶資料", bg=BG, font=FONT_S)
        cust_lf.pack(fill="x", padx=10, pady=(2, 4))
        cust_lf.columnconfigure(1, weight=1)
        cust_lf.columnconfigure(3, weight=1)
        _cust_vars: dict[str, tk.StringVar] = {}
        for r, (label, key) in enumerate([
            ("公司名稱", "company"), ("聯絡人", "contact"),
            ("電話",     "phone"),   ("傳真",   "fax"),
            ("地址",     "address"), ("統一編號","tax_id"),
            ("E-MAIL",   "email"),
        ]):
            col = (r % 2) * 2
            row_i = r // 2
            tk.Label(cust_lf, text=label + "：", bg=BG, fg=GRAY,
                     font=FONT_S, anchor="e").grid(row=row_i, column=col, sticky="e",
                                                    padx=(6, 2), pady=2)
            sv = tk.StringVar()
            _cust_vars[key] = sv
            tk.Entry(cust_lf, textvariable=sv, font=FONT_S, state="readonly",
                     width=18).grid(row=row_i, column=col+1, sticky="ew", padx=(0, 8), pady=2)

        def _strip_md_link(text: str) -> str:
            """去除 Trello Markdown 連結格式，只保留顯示文字。
            [text](url) → text
            """
            import re as _re
            return _re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', text).strip()

        def _on_card_select(*_):
            from core.generator_quote import parse_card_desc
            sel = card_tree.selection()
            if not sel: return
            idx  = card_tree.index(sel[0])
            card = _filtered_cards[idx]
            _card[0] = card
            desc = parse_card_desc(card.get("desc", ""))

            # 更新 Trello 連結
            url = card.get("shortUrl") or card.get("url", "")
            _card_url[0] = url
            if url:
                name_short = card["name"][:40] + ("…" if len(card["name"]) > 40 else "")
                trello_link_lbl.config(text=f"🔗 {name_short}")
            else:
                trello_link_lbl.config(text="")

            def _g(*keys):
                for k in keys:
                    v = desc.get(k)
                    if v: return _strip_md_link(v)
                return ""

            _customer.clear()
            _customer.update({
                "company": _g("公司名", "公司名稱"),
                "contact": _g("聯絡人"),
                "phone":   _g("手機", "電話"),
                "fax":     _g("傳真"),
                "address": _g("地址", "聯絡地址"),
                "tax_id":  _g("統一編號", "统一編號", "统編", "統編"),
                "email":   _g("電子信箱", "E-MAIL", "EMAIL"),
            })
            for key, sv in _cust_vars.items():
                sv.set(_customer.get(key, ""))

        card_tree.bind("<<TreeviewSelect>>", _on_card_select)

        p1_next_btn = tk.Button(p1, text="下一步：選取品項 →",
                                 command=lambda: _go_step(2),
                                 bg=BLUE, fg="white", relief="flat",
                                 font=(FONT_S[0], 10, "bold"), pady=7)
        p1_next_btn.pack(fill="x", padx=10, pady=(2, 8))

        # ════════════════════════════════════════════════════
        # Step 2 — 選取品項
        # ════════════════════════════════════════════════════
        p2 = tk.Frame(step_content, bg=BG)

        # 分類 Tabs（可橫向捲動，支援 37 個系列）
        cat_bar_outer = tk.Frame(p2, bg=BG, height=48)
        cat_bar_outer.pack(fill="x", padx=10, pady=(8, 0))
        cat_bar_outer.pack_propagate(False)

        cat_cvs = tk.Canvas(cat_bar_outer, bg=BG, highlightthickness=0, height=48)
        cat_hsb = ttk.Scrollbar(cat_bar_outer, orient="horizontal", command=cat_cvs.xview)
        cat_cvs.configure(xscrollcommand=cat_hsb.set)
        cat_hsb.pack(side="bottom", fill="x")
        cat_cvs.pack(side="top", fill="both", expand=True)

        cat_bar  = tk.Frame(cat_cvs, bg=BG)
        _cat_win = cat_cvs.create_window((0, 0), window=cat_bar, anchor="nw")
        cat_bar.bind("<Configure>",
                     lambda e: cat_cvs.configure(scrollregion=cat_cvs.bbox("all")))

        # 滑鼠滾輪捲動（Windows 需綁到所有子元件）
        def _cat_scroll(e):
            cat_cvs.xview_scroll(int(-1 * (e.delta / 120)), "units")

        def _prod_scroll(e):
            prod_cvs.yview_scroll(int(-1 * (e.delta / 120)), "units")

        def _bind_scroll(widget, fn):
            widget.bind("<MouseWheel>", fn)
            for child in widget.winfo_children():
                _bind_scroll(child, fn)

        cat_cvs.bind("<MouseWheel>", _cat_scroll)
        cat_bar.bind("<MouseWheel>",  _cat_scroll)

        _cat_btns: dict[str, tk.Button] = {}

        def _set_cat(cat: str):
            _cur_cat[0] = cat
            for c, b in _cat_btns.items():
                b.config(bg=BLUE if c == cat else "#dee2e6",
                         fg="white" if c == cat else "#333")
            _render_products()

        for cat in CATS:
            b = tk.Button(cat_bar, text=cat, font=FONT_S, relief="flat",
                          bg="#dee2e6", fg="#333", padx=8, pady=3,
                          command=lambda c=cat: _set_cat(c))
            b.pack(side="left", padx=(0, 4))
            b.bind("<MouseWheel>", _cat_scroll)   # 按鈕本身也綁
            _cat_btns[cat] = b

        # 產品格狀列表（Canvas + 2欄）
        prod_outer = tk.Frame(p2, bg=BG)
        prod_outer.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        prod_cvs = tk.Canvas(prod_outer, bg=BG, highlightthickness=0)
        prod_sb  = ttk.Scrollbar(prod_outer, orient="vertical", command=prod_cvs.yview)
        prod_cvs.configure(yscrollcommand=prod_sb.set)
        prod_sb.pack(side="right", fill="y")
        prod_cvs.pack(side="left", fill="both", expand=True)

        prod_inner = tk.Frame(prod_cvs, bg=BG)
        _prod_win  = prod_cvs.create_window((0, 0), window=prod_inner, anchor="nw")

        def _on_prod_resize(e=None):
            prod_cvs.configure(scrollregion=prod_cvs.bbox("all"))
            if prod_cvs.winfo_width() > 1:
                prod_cvs.itemconfig(_prod_win, width=prod_cvs.winfo_width())
        prod_inner.bind("<Configure>", _on_prod_resize)
        prod_cvs.bind("<Configure>",   _on_prod_resize)

        def _render_products():
            for w in prod_inner.winfo_children():
                w.destroy()
            cat = _cur_cat[0]
            filtered = [p for p in _products
                        if cat == "全部" or p.get("category") == cat]
            for idx, prod in enumerate(filtered):
                col   = idx % 2
                row_i = idx // 2
                in_cart = prod["code"] in _cart and _cart[prod["code"]]["qty"] > 0
                card_bg = "#d8eaf8" if in_cart else "white"
                bdr     = BLUE if in_cart else "#d0d7de"

                cf2 = tk.Frame(prod_inner, bg=card_bg, bd=1, relief="solid",
                               highlightbackground=bdr, highlightthickness=1)
                cf2.grid(row=row_i, column=col, padx=4, pady=4, sticky="nsew")
                prod_inner.columnconfigure(col, weight=1)

                tk.Label(cf2, text=prod["code"], bg=card_bg, fg=GRAY,
                         font=MONO, anchor="w").pack(fill="x", padx=8, pady=(6, 0))
                # 品名（去掉 code 前綴）
                disp_name = prod["name"].replace(prod["code"] + " ", "").replace(prod["code"], "")
                tk.Label(cf2, text=disp_name.strip(), bg=card_bg, fg="#1a1a1a",
                         font=(FONT_S[0], 9, "bold"), anchor="w",
                         wraplength=160, justify="left").pack(fill="x", padx=8, pady=(0, 2))
                tk.Label(cf2, text=prod.get("spec", ""), bg=card_bg, fg=GRAY,
                         font=(MONO[0], 8), anchor="w",
                         wraplength=160, justify="left").pack(fill="x", padx=8, pady=(0, 4))

                bot_f = tk.Frame(cf2, bg=card_bg)
                bot_f.pack(fill="x", padx=8, pady=(0, 6))
                tk.Label(bot_f, text=f"${prod['price']:,}/{prod['unit']}",
                         bg=card_bg, fg=BLUE,
                         font=(FONTB[0], 10, "bold")).pack(side="left")

                def _add(p=prod):
                    code = p["code"]
                    if code not in _cart:
                        _cart[code] = {"product": p, "qty": 0, "price": p["price"]}
                    _cart[code]["qty"] += 1
                    _render_cart()
                    _render_products()

                add_lbl = "✔" if in_cart else "+"
                add_bg  = GREEN if in_cart else BLUE
                tk.Button(bot_f, text=add_lbl, command=_add,
                          bg=add_bg, fg="white", relief="flat",
                          font=(FONTB[0], 12), width=2, pady=0
                          ).pack(side="right")

            # 產品卡每次重建後，把所有子元件都綁上垂直滾輪
            _bind_scroll(prod_inner, _prod_scroll)

        prod_cvs.bind("<MouseWheel>", _prod_scroll)

        # 初始化：設定第一個分類並渲染
        if CATS:
            _cur_cat[0] = CATS[0]
            for c, b in _cat_btns.items():
                b.config(bg=BLUE if c == CATS[0] else "#dee2e6",
                         fg="white" if c == CATS[0] else "#333")
        _render_products()

        p2_back_btn = tk.Button(p2, text="← 返回顧客資料",
                                 command=lambda: _go_step(1),
                                 bg="#5d6d7e", fg="white", relief="flat",
                                 font=FONT_S, padx=8, pady=4)
        p2_back_btn.pack(anchor="w", padx=10, pady=(0, 6))

        # ════════════════════════════════════════════════════
        # Step 3 — 報價確認
        # ════════════════════════════════════════════════════
        p3 = tk.Frame(step_content, bg=BG)

        # 客戶摘要
        p3_cust_f = tk.LabelFrame(p3, text="客戶資料", bg=BG, font=FONT_S)
        p3_cust_f.pack(fill="x", padx=10, pady=(8, 4))
        p3_cust_f.columnconfigure(1, weight=1)
        p3_cust_f.columnconfigure(3, weight=1)
        _p3_cust_lbls: dict[str, tk.Label] = {}
        for r, (label, key) in enumerate([
            ("公司名稱", "company"), ("聯絡人",  "contact"),
            ("報價單號", "_quote_no"), ("報價日期", "_quote_date"),
        ]):
            col = (r % 2) * 2
            row_i = r // 2
            tk.Label(p3_cust_f, text=label + "：", bg=BG, fg=GRAY,
                     font=FONT_S).grid(row=row_i, column=col, sticky="e", padx=(6, 2), pady=2)
            lbl = tk.Label(p3_cust_f, text="—", bg=BG, fg="#1a1a1a",
                           font=(FONT_S[0], 9, "bold"), anchor="w")
            lbl.grid(row=row_i, column=col+1, sticky="ew", padx=(0, 8), pady=2)
            _p3_cust_lbls[key] = lbl

        # 品項確認表
        conf_lf = tk.LabelFrame(p3, text="報價品項", bg=BG, font=FONT_S)
        conf_lf.pack(fill="both", expand=True, padx=10, pady=4)
        _ccols3 = ("code", "name", "qty", "price", "subtotal")
        conf_tree = ttk.Treeview(conf_lf, columns=_ccols3, show="headings",
                                  selectmode="none", height=6)
        conf_tree.heading("code",     text="品號")
        conf_tree.heading("name",     text="品名")
        conf_tree.heading("qty",      text="數量")
        conf_tree.heading("price",    text="單價")
        conf_tree.heading("subtotal", text="小計")
        conf_tree.column("code",     width=100, anchor="w")
        conf_tree.column("name",     width=200, anchor="w", stretch=True)
        conf_tree.column("qty",      width=55,  anchor="center", stretch=False)
        conf_tree.column("price",    width=80,  anchor="e", stretch=False)
        conf_tree.column("subtotal", width=90,  anchor="e", stretch=False)
        conf_tree.pack(fill="both", expand=True)

        # 摘要
        sum_f = tk.Frame(p3, bg="#f0f4f8", bd=1, relief="ridge")
        sum_f.pack(fill="x", padx=10, pady=(0, 4))
        p3_sub_lbl = tk.Label(sum_f, text="小計：—", bg="#f0f4f8", fg=GRAY,
                               font=FONT_S, anchor="e")
        p3_sub_lbl.pack(fill="x", padx=12, pady=(6, 1))
        p3_tax_lbl = tk.Label(sum_f, text="營業稅 5%：—", bg="#f0f4f8", fg=GRAY,
                               font=FONT_S, anchor="e")
        p3_tax_lbl.pack(fill="x", padx=12, pady=1)
        p3_tot_lbl = tk.Label(sum_f, text="應收總金額：—", bg="#f0f4f8",
                               font=(FONTB[0], 11, "bold"), fg=BLUE, anchor="e")
        p3_tot_lbl.pack(fill="x", padx=12, pady=(1, 6))

        p3_out_lbl = tk.Label(p3, text="", bg=BG, font=FONT_S, fg=GRAY,
                               anchor="w", wraplength=620)
        p3_out_lbl.pack(fill="x", padx=10)

        # ── 生成按鈕（放在歷史面板前，確保一定看得到）──────────
        hist_lf = tk.LabelFrame(p3, text="📋  歷史相同組合報價", bg=BG, font=FONT_S)
        hist_lf.pack(fill="x", padx=10, pady=(6, 2))

        hist_cols = ("date", "customer", "total")
        hist_tree = ttk.Treeview(hist_lf, columns=hist_cols,
                                 show="headings", height=4)
        hist_tree.heading("date",     text="日期")
        hist_tree.heading("customer", text="客戶")
        hist_tree.heading("total",    text="應收總金額（含稅）")
        hist_tree.column("date",     width=100, anchor="center")
        hist_tree.column("customer", width=200, anchor="w")
        hist_tree.column("total",    width=140, anchor="e")
        hist_tree.pack(fill="x", padx=4, pady=(4, 2))

        hist_hint = tk.Label(hist_lf, text="查詢中…", bg=BG,
                             font=("Microsoft JhengHei UI", 8), fg=GRAY, anchor="w")
        hist_hint.pack(anchor="w", padx=6, pady=(0, 4))

        def _show_hist_detail(event=None):
            sel = hist_tree.selection()
            if not sel:
                return
            quote_id = hist_tree.item(sel[0])["tags"][0]
            from core.repository import get_quote
            data = get_quote(int(quote_id))
            if not data:
                return
            q, items_d = data["quote"], data["items"]

            dlg = tk.Toplevel(self)
            dlg.title(f"報價單明細 — {q['quote_no']}")
            dlg.configure(bg=BG)
            dlg.grab_set()
            dlg.geometry("680x480")

            info_lf = tk.LabelFrame(dlg, text="基本資訊", bg=BG, font=FONTB)
            info_lf.pack(fill="x", padx=12, pady=(12, 4))
            info_lf.columnconfigure(1, weight=1)
            info_lf.columnconfigure(3, weight=1)

            FONT_S2 = ("Microsoft JhengHei UI", 9)
            for r, (l1, v1, l2, v2) in enumerate([
                ("報價單號", q["quote_no"],       "日期",   q["date"]),
                ("客戶",     q["customer"],        "聯絡人", q["contact"] or "—"),
                ("電話",     q["phone"] or "—",   "總金額", f"{q['total']:,.0f}"),
            ]):
                tk.Label(info_lf, text=l1+"：", bg=BG, font=FONT_S2,
                         fg=GRAY, anchor="e").grid(row=r, column=0, sticky="e", padx=(8,2), pady=3)
                tk.Label(info_lf, text=v1, bg=BG, font=FONT_S2,
                         anchor="w").grid(row=r, column=1, sticky="w", pady=3)
                tk.Label(info_lf, text=l2+"：", bg=BG, font=FONT_S2,
                         fg=GRAY, anchor="e").grid(row=r, column=2, sticky="e", padx=(16,2), pady=3)
                tk.Label(info_lf, text=v2, bg=BG, font=FONT_S2,
                         anchor="w").grid(row=r, column=3, sticky="w", pady=3)

            item_lf = tk.LabelFrame(dlg, text="品項明細", bg=BG, font=FONTB)
            item_lf.pack(fill="both", expand=True, padx=12, pady=(0, 8))

            icols = ("seq", "code", "name", "qty", "unit", "unit_price", "subtotal")
            itree = ttk.Treeview(item_lf, columns=icols, show="headings", height=10)
            itree.heading("seq",        text="#")
            itree.heading("code",       text="型號")
            itree.heading("name",       text="品名")
            itree.heading("qty",        text="數量")
            itree.heading("unit",       text="單位")
            itree.heading("unit_price", text="單價")
            itree.heading("subtotal",   text="小計")
            itree.column("seq",        width=35,  anchor="center")
            itree.column("code",       width=110, anchor="w")
            itree.column("name",       width=150, anchor="w")
            itree.column("qty",        width=50,  anchor="center")
            itree.column("unit",       width=50,  anchor="center")
            itree.column("unit_price", width=90,  anchor="e")
            itree.column("subtotal",   width=90,  anchor="e")
            isb = ttk.Scrollbar(item_lf, orient="vertical", command=itree.yview)
            itree.configure(yscrollcommand=isb.set)
            isb.pack(side="right", fill="y")
            itree.pack(fill="both", expand=True, padx=4, pady=4)

            for it in items_d:
                itree.insert("", "end", values=(
                    it["seq"], it["code"], it["name"],
                    it["qty"], it["unit"],
                    f"{it['unit_price']:,.0f}",
                    f"{it['subtotal']:,.0f}",
                ))

            tk.Button(dlg, text="關閉", command=dlg.destroy,
                      bg=GRAY, fg="white", relief="flat",
                      font=FONT, padx=16, pady=4).pack(pady=(0, 10))

        hist_tree.bind("<Double-1>", _show_hist_detail)

        # ── 相似組合（多/少一個品號）────────────────────────
        sim_lf = tk.LabelFrame(p3, text="🔍  相似組合報價（多或少一個品號）",
                               bg=BG, font=FONT_S)
        sim_lf.pack(fill="x", padx=10, pady=(4, 2))

        sim_cols = ("similarity", "date", "customer", "total")
        sim_tree = ttk.Treeview(sim_lf, columns=sim_cols,
                                show="headings", height=4)
        sim_tree.heading("similarity", text="差異")
        sim_tree.heading("date",       text="日期")
        sim_tree.heading("customer",   text="客戶")
        sim_tree.heading("total",      text="應收總金額（含稅）")
        sim_tree.column("similarity", width=120, anchor="center")
        sim_tree.column("date",       width=100, anchor="center")
        sim_tree.column("customer",   width=180, anchor="w")
        sim_tree.column("total",      width=140, anchor="e")
        sim_tree.pack(fill="x", padx=4, pady=(4, 2))

        sim_hint = tk.Label(sim_lf, text="", bg=BG,
                            font=("Microsoft JhengHei UI", 8), fg=GRAY, anchor="w")
        sim_hint.pack(anchor="w", padx=6, pady=(0, 4))

        def _show_sim_detail(event=None):
            sel = sim_tree.selection()
            if not sel:
                return
            quote_id = sim_tree.item(sel[0])["tags"][0]
            from core.repository import get_quote
            data = get_quote(int(quote_id))
            if not data:
                return
            q, items_d = data["quote"], data["items"]
            dlg = tk.Toplevel(self)
            dlg.title(f"報價單明細 — {q['quote_no']}")
            dlg.configure(bg=BG)
            dlg.grab_set()
            dlg.geometry("680x480")
            FONT_S2 = ("Microsoft JhengHei UI", 9)
            info_lf = tk.LabelFrame(dlg, text="基本資訊", bg=BG, font=FONTB)
            info_lf.pack(fill="x", padx=12, pady=(12, 4))
            info_lf.columnconfigure(1, weight=1)
            info_lf.columnconfigure(3, weight=1)
            for r, (l1, v1, l2, v2) in enumerate([
                ("報價單號", q["quote_no"],     "日期",   q["date"]),
                ("客戶",     q["customer"],      "聯絡人", q["contact"] or "—"),
                ("電話",     q["phone"] or "—", "總金額", f"{q['total']:,.0f}"),
            ]):
                tk.Label(info_lf, text=l1+"：", bg=BG, font=FONT_S2, fg=GRAY,
                         anchor="e").grid(row=r, column=0, sticky="e", padx=(8,2), pady=3)
                tk.Label(info_lf, text=v1, bg=BG, font=FONT_S2,
                         anchor="w").grid(row=r, column=1, sticky="w", pady=3)
                tk.Label(info_lf, text=l2+"：", bg=BG, font=FONT_S2, fg=GRAY,
                         anchor="e").grid(row=r, column=2, sticky="e", padx=(16,2), pady=3)
                tk.Label(info_lf, text=v2, bg=BG, font=FONT_S2,
                         anchor="w").grid(row=r, column=3, sticky="w", pady=3)
            item_lf = tk.LabelFrame(dlg, text="品項明細", bg=BG, font=FONTB)
            item_lf.pack(fill="both", expand=True, padx=12, pady=(0, 8))
            icols = ("seq", "code", "name", "qty", "unit", "unit_price", "subtotal")
            itree = ttk.Treeview(item_lf, columns=icols, show="headings", height=10)
            for col, hdr, w, anc in [
                ("seq","#",35,"center"),("code","型號",110,"w"),("name","品名",150,"w"),
                ("qty","數量",50,"center"),("unit","單位",50,"center"),
                ("unit_price","單價",90,"e"),("subtotal","小計",90,"e"),
            ]:
                itree.heading(col, text=hdr)
                itree.column(col, width=w, anchor=anc)
            isb = ttk.Scrollbar(item_lf, orient="vertical", command=itree.yview)
            itree.configure(yscrollcommand=isb.set)
            isb.pack(side="right", fill="y")
            itree.pack(fill="both", expand=True, padx=4, pady=4)
            for it in items_d:
                itree.insert("", "end", values=(
                    it["seq"], it["code"], it["name"], it["qty"], it["unit"],
                    f"{it['unit_price']:,.0f}", f"{it['subtotal']:,.0f}",
                ))
            tk.Button(dlg, text="關閉", command=dlg.destroy,
                      bg=GRAY, fg="white", relief="flat",
                      font=FONT, padx=16, pady=4).pack(pady=(0, 10))

        sim_tree.bind("<Double-1>", _show_sim_detail)

        def _refresh_history_panel(cart_items):
            from core.repository import find_same_combination
            codes = [it["product"]["code"] for it in cart_items
                     if it.get("product", {}).get("code")]
            hist_tree.delete(*hist_tree.get_children())
            if not codes:
                hist_hint.config(text="購物車無品項")
                return
            try:
                results = find_same_combination(codes)
            except Exception:
                hist_hint.config(text="（資料庫未連線或尚無記錄）")
                return
            if not results:
                hist_hint.config(text="無相同品號組合的歷史記錄")
            else:
                totals = []
                for r in results:
                    hist_tree.insert("", "end", values=(
                        r["date"], r["customer"], f"${r['total']:,.0f}"
                    ), tags=(r["id"],))
                    totals.append(r["total"])
                avg = sum(totals) / len(totals)
                hist_hint.config(
                    text=f"共 {len(results)} 筆  ｜  最近：${totals[0]:,.0f}  ｜  平均：${avg:,.0f}"
                )

            # ── 相似組合 ──────────────────────────────────
            sim_tree.delete(*sim_tree.get_children())
            try:
                from core.repository import find_similar_combinations
                sim_results = find_similar_combinations(codes)
            except Exception:
                sim_hint.config(text="（資料庫未連線或尚無記錄）")
                return
            if not sim_results:
                sim_hint.config(text="無相似組合的歷史記錄")
                return
            for r in sim_results:
                sim_tree.insert("", "end", values=(
                    r["similarity"], r["date"], r["customer"], f"${r['total']:,.0f}"
                ), tags=(r["id"],))
            sim_hint.config(text=f"共 {len(sim_results)} 筆")

        def _build_confirm():
            conf_tree.delete(*conf_tree.get_children())
            items = [v for v in _cart.values() if v["qty"] > 0]
            sub   = 0
            for v in items:
                s = v["qty"] * v["price"]
                sub += s
                conf_tree.insert("", "end", values=(
                    v["product"]["code"],
                    v["product"]["name"][:30],
                    f"{v['qty']} {v['product']['unit']}",
                    f"${int(v['price']):,}",
                    f"${int(s):,}",
                ))
            freight = 0
            if ship_var.get():
                try:
                    freight = float(ship_price_var.get().replace(",", ""))
                except ValueError:
                    freight = 0
                if freight:
                    promo_txt = "（含免運優惠）" if ship_promo_var.get() else ""
                    conf_tree.insert("", "end", values=(
                        "N/A", f"一次性運費{promo_txt}", "1 趟",
                        f"${int(freight):,}", f"${int(freight):,}",
                    ))
            tax = round((sub + freight) * 0.05)
            p3_sub_lbl.config(text=f"小計：{_fmt(sub)}" + (f" + 運費 {_fmt(freight)}" if freight else ""))
            p3_tax_lbl.config(text=f"營業稅 5%：{_fmt(tax)}")
            p3_tot_lbl.config(text=f"應收總金額：{_fmt(sub + freight + tax)}")
            _p3_cust_lbls["company"].config(    text=_customer.get("company", "—"))
            _p3_cust_lbls["contact"].config(    text=_customer.get("contact", "—"))
            _p3_cust_lbls["_quote_no"].config(  text=quote_no_var.get())
            _p3_cust_lbls["_quote_date"].config( text=date_entry.get_date().strftime("%Y/%m/%d"))

            # ── 歷史相同組合對比 ──────────────────────────────
            _refresh_history_panel(items)

        def _generate():
            from core.generator_quote import generate_quote_from_cart
            if not _cart:
                messagebox.showwarning("購物車空白", "請先選取品項", parent=parent); return
            if not _customer.get("company"):
                messagebox.showwarning("無客戶資料", "請先在步驟①選取 Trello 卡片", parent=parent); return
            # 依報價單類型選擇範本
            from core.generator_quote import (
                QUOTE_TYPE_REGULAR, QUOTE_TYPE_ALLOWANCE, QUOTE_TYPE_COMPARE)
            _type_map = {
                "一般報價單": (QUOTE_TYPE_REGULAR,   "template_quote.xlsx"),
                "補助報價單": (QUOTE_TYPE_ALLOWANCE, "template_quote_allowance.xlsx"),
                "對比報價單": (QUOTE_TYPE_COMPARE,   "template_quote_compare.xlsx"),
            }
            _qt, _tpl_name = _type_map.get(
                quote_type_var.get(), (QUOTE_TYPE_REGULAR, "template_quote.xlsx"))
            tpl_path = TEMPLATE_DIR / _tpl_name
            if not tpl_path.exists():
                messagebox.showerror("找不到範本", f"找不到 {tpl_path}", parent=parent); return

            cart_items = [
                {**{"code":  v["product"]["code"],
                    "name":  v["product"]["name"],
                    "spec":  v["product"].get("spec", ""),
                    "unit":  v["product"]["unit"],
                    "qty":   v["qty"],
                    "price": v["price"]}}
                for v in _cart.values() if v["qty"] > 0
            ]
            q_date   = date_entry.get_date()
            quote_no = quote_no_var.get().strip()
            out_dir  = self._get_path("output_quote")

            p3_out_lbl.config(text="生成中…", fg=GRAY)
            parent.update_idletasks()
            shipping_info = None
            if ship_var.get():
                try:
                    sp = float(ship_price_var.get().replace(",", ""))
                except ValueError:
                    sp = 1500
                shipping_info = {
                    "enabled": True,
                    "price":   sp,
                    "promo":   ship_promo_var.get(),
                }
            try:
                out_path = generate_quote_from_cart(
                    _customer, cart_items, tpl_path, out_dir, quote_no, q_date,
                    operator=op_var.get().strip(),
                    shipping=shipping_info,
                    quote_type=_qt)
                p3_out_lbl.config(text=f"✔  已生成：{out_path}", fg=GREEN)

                # 詢問是否存入資料庫
                if messagebox.askyesno("儲存記錄",
                        "是否將此報價單儲存到資料庫記錄？", parent=parent):
                    try:
                        from core.repository import save_quote
                        freight = 0.0
                        if shipping_info and shipping_info.get("enabled"):
                            freight = float(shipping_info.get("price", 0))
                        subtotal = sum(
                            it.get("qty", 1) * it.get("price", 0)
                            for it in cart_items
                        )
                        save_quote(
                            quote_no = quote_no,
                            date     = q_date.strftime("%Y-%m-%d"),
                            customer = _customer.get("company", ""),
                            contact  = _customer.get("contact", ""),
                            phone    = _customer.get("phone", ""),
                            total    = round(subtotal + freight, 2),
                            items    = [
                                {
                                    "seq":        i + 1,
                                    "code":       it.get("code", ""),
                                    "name":       it.get("name", ""),
                                    "spec":       it.get("spec", ""),
                                    "qty":        it.get("qty", 1),
                                    "unit":       it.get("unit", ""),
                                    "unit_price": it.get("price", 0),
                                    "subtotal":   it.get("qty", 1) * it.get("price", 0),
                                }
                                for i, it in enumerate(cart_items)
                            ],
                        )
                    except Exception as db_err:
                        messagebox.showwarning(
                            "資料庫寫入失敗",
                            f"報價單已生成，但儲存記錄時發生錯誤：\n{db_err}",
                            parent=parent,
                        )

                if messagebox.askyesno("完成",
                        f"報價單已生成\n{out_path}\n\n是否立即開啟？", parent=parent):
                    os.startfile(str(out_path))
            except Exception as e:
                p3_out_lbl.config(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("生成失敗", str(e), parent=parent)

        p3_back_btn = tk.Button(p3, text="← 返回選取品項",
                                 command=lambda: _go_step(2),
                                 bg="#5d6d7e", fg="white", relief="flat",
                                 font=FONT_S, padx=8, pady=3)
        p3_back_btn.pack(anchor="w", padx=10, pady=(0, 4))

        # ════════════════════════════════════════════════════
        # Step 切換
        # ════════════════════════════════════════════════════
        _panels = {1: p1, 2: p2, 3: p3}

        def _go_step(n: int):
            for i, lbl in enumerate(_step_btns, 1):
                if i == n:
                    lbl.config(bg=BLUE, fg="white")
                elif i < n:
                    lbl.config(bg="#d5e8f5", fg=BLUE)
                else:
                    lbl.config(bg="#e8ecf0", fg=GRAY)
            for i, panel in _panels.items():
                if i == n:
                    panel.pack(fill="both", expand=True)
                else:
                    panel.pack_forget()
            # 第 1 步不顯示購物車
            if n == 1:
                right_border.pack_forget()
                cart_panel.pack_forget()
            else:
                right_border.pack(side="right", fill="y")
                cart_panel.pack(side="right", fill="y")
            if n == 3:
                _build_confirm()
                checkout_btn.config(
                    text="📄  生成報價單 .xlsx",
                    bg=GREEN,
                    command=_generate,
                )
            else:
                checkout_btn.config(
                    text="確認報價內容 →",
                    bg=BLUE,
                    command=lambda: _go_step(3),
                )

        checkout_btn.config(command=lambda: _go_step(3))
        _render_cart()
        right_border.pack_forget()
        cart_panel.pack_forget()
        _go_step(1)
