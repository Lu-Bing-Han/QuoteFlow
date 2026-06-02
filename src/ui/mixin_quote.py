"""
mixin_quote.py — 報價單生成頁籤 mixin
"""
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
import os
from pathlib import Path
from ui.app_core import _mk_lf


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

        _PIC_DIR   = TEMPLATE_DIR / "Picture"
        _img_cache: dict = {}

        def _load_prod_photo(code: str, category: str):
            if code in _img_cache:
                return _img_cache[code]
            try:
                from PIL import Image, ImageTk

                def _open(f):
                    _im = Image.open(f).convert("RGBA")
                    _im.thumbnail((140, 105), Image.LANCZOS)
                    return ImageTk.PhotoImage(_im)

                for _ext in ("png", "jpg", "jpeg", "bmp"):
                    _f = _PIC_DIR / category / f"{code}.{_ext}" if category else None
                    if not (_f and _f.exists()):
                        _m = list(_PIC_DIR.glob(f"*/{code}.{_ext}"))
                        _f = _m[0] if _m else None
                    if _f and _f.exists():
                        _img_cache[code] = _open(_f)
                        return _img_cache[code]

                _search_dirs = ([_PIC_DIR / category] if category else []) + \
                               [d for d in _PIC_DIR.iterdir() if d.is_dir()]
                _best: tuple = ("", None)
                for _d in _search_dirs:
                    for _f in _d.iterdir():
                        if _f.suffix.lower() not in (".png", ".jpg", ".jpeg", ".bmp"):
                            continue
                        _stem = _f.stem
                        if code.startswith(_stem) and len(_stem) > len(_best[0]):
                            _best = (_stem, _f)
                if _best[1]:
                    _img_cache[code] = _open(_best[1])
                    return _img_cache[code]

            except Exception:
                pass
            _img_cache[code] = None
            return None

        # ── 狀態 ────────────────────────────────────────────
        _all_cards:      list[dict]      = []
        _filtered_cards: list[dict]      = []
        _card:           list[dict|None] = [None]
        _customer:       dict            = {}
        _cart:           dict            = {}
        _cur_cat:        list[str]       = ["all"]
        _products: list[dict] = load_product_catalog(TEMPLATE_DIR)
        CATS = ["全部"] + sorted({p["category"] for p in _products})

        # ════════════════════════════════════════════════════
        # 右側購物車 Panel
        # ════════════════════════════════════════════════════
        body = tk.Frame(parent, bg=BG)
        body.pack(fill="both", expand=True)

        left_area = tk.Frame(body, bg=BG)
        left_area.pack(side="left", fill="both", expand=True)

        right_border = ctk.CTkFrame(body, fg_color="#c8cfd6", corner_radius=0, width=1)
        right_border.pack(side="right", fill="y")
        right_border.pack_propagate(False)

        cart_panel = ctk.CTkFrame(body, fg_color=BG, corner_radius=0, width=272)
        cart_panel.pack(side="right", fill="y")
        cart_panel.pack_propagate(False)

        # 購物車 Header
        ch = ctk.CTkFrame(cart_panel, fg_color="#e8ecf0", corner_radius=0)
        ch.pack(fill="x")
        ctk.CTkLabel(ch, text="🛒  報價清單", fg_color="transparent",
                      text_color=BLUE,
                      font=(FONTB[0], 10, "bold")).pack(side="left", padx=10, pady=7)
        cart_cnt_lbl = ctk.CTkLabel(ch, text="", fg_color="transparent",
                                     text_color=GRAY, font=FONT_S)
        cart_cnt_lbl.pack(side="right", padx=10)

        # 購物車可捲動 items
        cart_mid = ctk.CTkFrame(cart_panel, fg_color=BG, corner_radius=0)
        cart_mid.pack(fill="both", expand=True)
        cart_cvs = tk.Canvas(cart_mid, bg=BG, highlightthickness=0)
        cart_sb  = ctk.CTkScrollbar(cart_mid, orientation="vertical", command=cart_cvs.yview)
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
            w = widget or cart_inner
            w.bind("<MouseWheel>", _cart_scroll)
            for child in w.winfo_children():
                _rebind_cart_scroll(child)

        cart_cvs.bind("<MouseWheel>", _cart_scroll)

        # 購物車 Footer
        cf = ctk.CTkFrame(cart_panel, fg_color="#f0f4f8", corner_radius=0,
                           border_width=1, border_color="#c8cfd6")
        cf.pack(fill="x")
        cart_sub_lbl = ctk.CTkLabel(cf, text="小計：—", fg_color="transparent",
                                     text_color=GRAY, font=FONT_S, anchor="e")
        cart_sub_lbl.pack(fill="x", padx=10, pady=(7, 1))
        cart_tax_lbl = ctk.CTkLabel(cf, text="營業稅 5%：—", fg_color="transparent",
                                     text_color=GRAY, font=FONT_S, anchor="e")
        cart_tax_lbl.pack(fill="x", padx=10, pady=1)
        cart_tot_lbl = ctk.CTkLabel(cf, text="合計：—", fg_color="transparent",
                                     font=(FONTB[0], 10, "bold"), text_color=BLUE, anchor="e")
        cart_tot_lbl.pack(fill="x", padx=10, pady=(1, 4))

        # ── 運費設定 ────────────────────────────────────────
        ctk.CTkFrame(cf, fg_color="#c8cfd6", height=1, corner_radius=0).pack(fill="x", padx=6)
        ship_outer = ctk.CTkFrame(cf, fg_color="#f0f4f8", corner_radius=0)
        ship_outer.pack(fill="x", padx=8, pady=(3, 0))

        ship_var       = tk.BooleanVar(value=False)
        ship_price_var = tk.StringVar(value="1500")
        ship_promo_var = tk.BooleanVar(value=False)

        ship_row1 = ctk.CTkFrame(ship_outer, fg_color="transparent", corner_radius=0)
        ship_row1.pack(fill="x")
        ship_cb = ctk.CTkCheckBox(ship_row1, text="加入運費", variable=ship_var,
                                   font=FONT_S, text_color=GRAY,
                                   fg_color=BLUE, hover_color="#2e6da4",
                                   checkmark_color="white")
        ship_cb.pack(side="left")
        ctk.CTkLabel(ship_row1, text="$", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        ship_price_entry = ctk.CTkEntry(ship_row1, textvariable=ship_price_var,
                                         width=56, height=24, font=FONT_S,
                                         corner_radius=4, state="disabled")
        ship_price_entry.pack(side="left", padx=(0, 2))

        ship_row2 = ctk.CTkFrame(ship_outer, fg_color="transparent", corner_radius=0)
        ship_row2.pack(fill="x", pady=(1, 4))
        ship_promo_cb = ctk.CTkCheckBox(ship_row2, text="顯示免運優惠文字",
                                         variable=ship_promo_var,
                                         font=FONT_S, text_color="#c0392b",
                                         fg_color="#c0392b", hover_color="#a93226",
                                         checkmark_color="white", state="disabled")
        ship_promo_cb.pack(side="left")

        def _on_ship_toggle(*_):
            state = "normal" if ship_var.get() else "disabled"
            ship_price_entry.configure(state=state)
            ship_promo_cb.configure(state=state)
            if not ship_var.get():
                ship_promo_var.set(False)
            _update_cart_totals()

        ship_var.trace_add("write", _on_ship_toggle)
        ship_price_var.trace_add("write", lambda *_: _update_cart_totals())

        checkout_btn = ctk.CTkButton(cf, text="確認報價內容 →", state="disabled",
                                      fg_color=BLUE, hover_color="#112d4e",
                                      text_color="white", font=FONT_S,
                                      height=34, corner_radius=0)
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
                cart_cnt_lbl.configure(text="")
                cart_sub_lbl.configure(text="小計：—")
                cart_tax_lbl.configure(text="營業稅 5%：—")
                cart_tot_lbl.configure(text="合計：—")
                checkout_btn.configure(state="disabled")
                return
            total_qty = sum(v["qty"] for _, v in items)
            cart_cnt_lbl.configure(text=f"{total_qty} 項")
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
                    _update_card_visual(c)

                def _inc(c=code):
                    _cart[c]["qty"] += 1
                    _render_cart()
                    _update_card_visual(c)

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
            cart_sub_lbl.configure(text=f"小計：{_fmt(sub)}" + (f" + 運費 {_fmt(freight)}" if freight else ""))
            cart_tax_lbl.configure(text=f"營業稅 5%：{_fmt(tax)}")
            cart_tot_lbl.configure(text=f"合計：{_fmt(tot)}")
            checkout_btn.configure(state="normal" if items else "disabled")

        # ════════════════════════════════════════════════════
        # Step Bar
        # ════════════════════════════════════════════════════
        step_bar = tk.Frame(left_area, bg="#e8ecf0")
        step_bar.pack(fill="x")
        _step_btns: list[ctk.CTkButton] = []
        for i, txt in enumerate(["① 顧客資料", "② 選取品項", "③ 報價確認"], 1):
            btn = ctk.CTkButton(step_bar, text=txt,
                                 fg_color="#e8ecf0", hover_color="#d5e8f5",
                                 text_color=GRAY, font=FONT_S,
                                 height=36, corner_radius=0,
                                 command=lambda n=i: _go_step(n))
            btn.pack(side="left", expand=True, fill="x")
            _step_btns.append(btn)
            if i < 3:
                tk.Frame(step_bar, bg="#c8cfd6", width=1).pack(side="left", fill="y")

        # Step 1 footer（固定在 left_area 底部，Step 2/3 時隱藏）
        p1_footer = tk.Frame(left_area, bg=BG)
        p1_footer.pack(side="bottom", fill="x")

        step_content = tk.Frame(left_area, bg=BG)
        step_content.pack(fill="both", expand=True)

        # ════════════════════════════════════════════════════
        # Step 1 — 顧客資料
        # ════════════════════════════════════════════════════
        p1 = tk.Frame(step_content, bg=BG)

        p1_top = tk.Frame(p1, bg=BG)
        p1_top.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(p1_top, text=f"來源：{_TARGET_LIST}",
                      fg_color="transparent", font=FONT_S,
                      text_color=GRAY).pack(side="left")
        p1_status = ctk.CTkLabel(p1_top, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY)
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
            card_title_lbl.configure(
                text=f"  待報價卡片（共 {total} 張{'，顯示 '+str(shown)+' 張' if shown<total else ''}）  ")

        def _apply_card_search(*_):
            kw = p1_search_var.get().strip().lower()
            matched = [c for c in _all_cards if kw in c["name"].lower()] if kw else list(_all_cards)
            _refresh_card_tree(matched)

        def _load_cards():
            from sync.downloader_trello import get_board_lists, get_list_cards
            api_key, token = _get_tr_creds()
            if not api_key: return
            p1_status.configure(text="載入中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                lists  = get_board_lists(api_key, token)
                target = next((l for l in lists
                               if "待報價" in l["name"] or "evaluate" in l["name"].lower()), None)
                if not target:
                    p1_status.configure(text="✘ 找不到清單", text_color="#c0392b"); return
                cards = get_list_cards(target["id"], api_key, token)
                _all_cards.clear(); _all_cards.extend(cards)
                p1_search_var.set("")
                _refresh_card_tree(_all_cards)
                p1_status.configure(text=f"✔ 找到 {len(cards)} 張", text_color=GREEN)
            except Exception as e:
                p1_status.configure(text=f"✘ {e}", text_color="#c0392b")

        ctk.CTkButton(p1_top, text="🔄 載入", command=_load_cards,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="right")

        # 搜索列
        p1_search_row = tk.Frame(p1, bg=BG)
        p1_search_row.pack(fill="x", padx=10, pady=(0, 2))
        ctk.CTkLabel(p1_search_row, text="🔍", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        p1_search_var = tk.StringVar()
        ctk.CTkEntry(p1_search_row, textvariable=p1_search_var,
                      font=FONT_S, width=220, height=28, corner_radius=4
                      ).pack(side="left", padx=(2, 4))
        ctk.CTkButton(p1_search_row, text="清除",
                       command=lambda: p1_search_var.set(""),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=50, height=28, corner_radius=4
                       ).pack(side="left")
        p1_search_var.trace_add("write", _apply_card_search)

        # 卡片 Treeview（手動建立，需動態更新標題）
        card_outer = tk.Frame(p1, bg="#d0d7de")
        card_outer.pack(fill="both", expand=True, padx=10, pady=2)
        card_inner = tk.Frame(card_outer, bg=BG)
        card_inner.pack(fill="both", expand=True, padx=1, pady=1)
        card_title_lbl = ctk.CTkLabel(card_inner, text="  待報價卡片（共 0 張）  ",
                                       fg_color="transparent", text_color=GRAY, font=FONT_S)
        card_title_lbl.pack(anchor="w", padx=10, pady=(4, 0))
        card_lf = tk.Frame(card_inner, bg=BG)
        card_lf.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        _ccols = ("name", "company", "contact")
        card_tree = ttk.Treeview(card_lf, columns=_ccols, show="headings",
                                  selectmode="browse", height=3)
        card_tree.heading("name",    text="卡片標題")
        card_tree.heading("company", text="公司名")
        card_tree.heading("contact", text="聯絡人")
        card_tree.column("name",    width=230, anchor="w", stretch=True)
        card_tree.column("company", width=130, anchor="w", stretch=False)
        card_tree.column("contact", width=80,  anchor="w", stretch=False)
        ct_vsb = ttk.Scrollbar(card_lf, orient="vertical", command=card_tree.yview)
        card_tree.configure(yscrollcommand=ct_vsb.set)
        ct_vsb.pack(side="right", fill="y")
        card_tree.pack(side="left", fill="both", expand=True)

        _card_url: list[str] = [""]
        trello_link_row = tk.Frame(p1, bg=BG)
        trello_link_row.pack(fill="x", padx=10, pady=(0, 2))
        trello_link_lbl = ctk.CTkLabel(
            trello_link_row, text="", fg_color="transparent",
            text_color="#1a5276",
            font=(FONT_S[0], 9, "underline"), anchor="w", cursor="hand2")
        trello_link_lbl.pack(side="left")

        def _open_trello_card(e=None):
            import webbrowser as _wb
            if _card_url[0]:
                _wb.open(_card_url[0])

        trello_link_lbl.bind("<Button-1>", _open_trello_card)

        # 報價設定
        cfg_outer, cfg_f = _mk_lf(p1, "報價設定", BG, FONT_S)
        cfg_outer.pack(fill="x", padx=10, pady=(4, 2))

        cfg_row0 = ctk.CTkFrame(cfg_f, fg_color="transparent", corner_radius=0)
        cfg_row0.pack(fill="x", padx=6, pady=(4, 2))
        ctk.CTkLabel(cfg_row0, text="製表人員：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        operators = self._config.get("operators", ["小皋"])
        op_var  = tk.StringVar(value=operators[0] if operators else "")
        op_cb   = ctk.CTkComboBox(cfg_row0, variable=op_var, values=operators,
                                   font=FONT_S, width=100, height=28)
        op_cb.pack(side="left", padx=(0, 16))

        from tkcalendar import DateEntry
        ctk.CTkLabel(cfg_row0, text="報價日期：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        date_entry = DateEntry(cfg_row0, font=FONT_S, date_pattern="yyyy/mm/dd",
                               width=12, background="#2e86c1", foreground="white",
                               borderwidth=1)
        date_entry.pack(side="left", padx=(0, 6))

        cfg_row1 = ctk.CTkFrame(cfg_f, fg_color="transparent", corner_radius=0)
        cfg_row1.pack(fill="x", padx=6, pady=(2, 2))
        valid_lbl = ctk.CTkLabel(cfg_row1, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY)
        valid_lbl.pack(side="left", padx=(0, 16))
        ctk.CTkLabel(cfg_row1, text="報價單號：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        quote_no_var = tk.StringVar()
        ctk.CTkEntry(cfg_row1, textvariable=quote_no_var, font=FONT_S,
                      width=128, height=28, corner_radius=4
                      ).pack(side="left", padx=(0, 6))

        cfg_row2 = ctk.CTkFrame(cfg_f, fg_color="transparent", corner_radius=0)
        cfg_row2.pack(fill="x", padx=6, pady=(2, 6))
        ctk.CTkLabel(cfg_row2, text="報價單類型：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        _QUOTE_TYPE_OPTIONS = ["一般報價單", "補助報價單", "對比報價單"]
        quote_type_var = tk.StringVar(value=_QUOTE_TYPE_OPTIONS[0])
        ctk.CTkComboBox(cfg_row2, variable=quote_type_var,
                         values=_QUOTE_TYPE_OPTIONS,
                         font=FONT_S, width=120, height=28).pack(side="left", padx=(0, 6))

        def _update_valid_label(*_):
            try:
                from datetime import timedelta as _td
                d  = date_entry.get_date()
                vd = d + _td(days=15)
                valid_lbl.configure(text=f"有效日期：{vd.strftime('%Y/%m/%d')}")
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
        cust_outer, cust_lf = _mk_lf(p1, "客戶資料", BG, FONT_S)
        cust_outer.pack(fill="x", padx=10, pady=(2, 4))
        cust_lf.columnconfigure(1, weight=1)
        cust_lf.columnconfigure(3, weight=1)
        _cust_vars: dict[str, tk.StringVar] = {}
        for r, (label, key) in enumerate([
            ("公司名稱", "company"), ("聯絡人", "contact"),
            ("電話",     "phone"),   ("傳真",   "fax"),
            ("地址",     "address"), ("統一編號","tax_id"),
            ("E-MAIL",   "email"),
        ]):
            col   = (r % 2) * 2
            row_i = r // 2
            ctk.CTkLabel(cust_lf, text=label + "：", fg_color="transparent",
                          text_color=GRAY, font=FONT_S,
                          anchor="e").grid(row=row_i, column=col, sticky="e",
                                           padx=(6, 2), pady=2)
            sv = tk.StringVar()
            _cust_vars[key] = sv
            ctk.CTkEntry(cust_lf, textvariable=sv, font=FONT_S,
                          state="readonly", width=144, height=26, corner_radius=4
                          ).grid(row=row_i, column=col+1, sticky="ew", padx=(0, 8), pady=2)

        def _strip_md_link(text: str) -> str:
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

            url = card.get("shortUrl") or card.get("url", "")
            _card_url[0] = url
            if url:
                name_short = card["name"][:40] + ("…" if len(card["name"]) > 40 else "")
                trello_link_lbl.configure(text=f"🔗 {name_short}")
            else:
                trello_link_lbl.configure(text="")

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
                "email":   _g("電子信箱", "E-MAIL", "EMAIL", "E-Mail", "Mail"),
            })
            for key, sv in _cust_vars.items():
                sv.set(_customer.get(key, ""))

        card_tree.bind("<<TreeviewSelect>>", _on_card_select)

        p1_next_btn = ctk.CTkButton(p1_footer, text="下一步：選取品項 →",
                                     command=lambda: _go_step(2),
                                     fg_color=BLUE, hover_color="#112d4e",
                                     text_color="white",
                                     font=(FONT_S[0], 10, "bold"),
                                     height=38, corner_radius=0)
        p1_next_btn.pack(fill="x")

        # ════════════════════════════════════════════════════
        # Step 2 — 選取品項
        # ════════════════════════════════════════════════════
        p2 = tk.Frame(step_content, bg=BG)

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
            prod_cvs.yview_moveto(0)

        for cat in CATS:
            b = tk.Button(cat_bar, text=cat, font=FONT_S, relief="flat",
                          bg="#dee2e6", fg="#333", padx=8, pady=3,
                          command=lambda c=cat: _set_cat(c))
            b.pack(side="left", padx=(0, 4))
            b.bind("<MouseWheel>", _cat_scroll)
            _cat_btns[cat] = b

        # 產品格狀列表
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

        # card frame registry: code -> (outer_border, inner_bg, add_btn)
        _card_widgets: dict[str, tuple] = {}

        def _card_colors(code: str):
            in_cart = code in _cart and _cart[code]["qty"] > 0
            return ("#d8eaf8" if in_cart else "white",
                    BLUE if in_cart else "#d0d7de",
                    "✔" if in_cart else "+",
                    GREEN if in_cart else BLUE,
                    "#0e6655" if in_cart else "#112d4e")

        def _update_card_visual(code: str):
            if code not in _card_widgets:
                return
            border_f, inner_f, children, add_btn = _card_widgets[code]
            card_bg, bdr, lbl, btn_fg, btn_hv = _card_colors(code)
            border_f.configure(bg=bdr)
            inner_f.configure(bg=card_bg)
            for w in children:
                try:
                    w.configure(bg=card_bg)
                except Exception:
                    pass
            add_btn.configure(text=lbl, fg_color=btn_fg, hover_color=btn_hv)

        def _render_products():
            for w in prod_inner.winfo_children():
                w.destroy()
            _card_widgets.clear()
            cat = _cur_cat[0]
            filtered = [p for p in _products
                        if cat == "全部" or p.get("category") == cat]
            for idx, prod in enumerate(filtered):
                col   = idx % 2
                row_i = idx // 2
                code  = prod["code"]
                card_bg, bdr, add_lbl, add_col, add_hv = _card_colors(code)

                border_f = tk.Frame(prod_inner, bg=bdr)
                border_f.grid(row=row_i, column=col, padx=4, pady=4, sticky="nsew")
                prod_inner.columnconfigure(col, weight=1)
                inner_f = tk.Frame(border_f, bg=card_bg)
                inner_f.pack(fill="both", expand=True, padx=1, pady=1)

                child_widgets = []

                _ph = _load_prod_photo(code, prod.get("category", ""))
                if _ph:
                    img_lbl = tk.Label(inner_f, image=_ph, bg=card_bg)
                    img_lbl.pack(pady=(6, 2))
                    child_widgets.append(img_lbl)

                for txt, fg, fnt, pkw in [
                    (code,                    GRAY,     MONO,                    dict(fill="x", padx=8, pady=(6, 0))),
                    (prod["name"].replace(code + " ", "").replace(code, "").strip(),
                                              "#1a1a1a",(FONT_S[0], 9, "bold"), dict(fill="x", padx=8, pady=(0, 2))),
                    (prod.get("spec", ""),    GRAY,     (MONO[0], 8),            dict(fill="x", padx=8, pady=(0, 4))),
                ]:
                    w = tk.Label(inner_f, text=txt, bg=card_bg, fg=fg,
                                 font=fnt, anchor="w", wraplength=160, justify="left")
                    w.pack(**pkw)
                    child_widgets.append(w)

                bot_f = tk.Frame(inner_f, bg=card_bg)
                bot_f.pack(fill="x", padx=8, pady=(0, 6))
                child_widgets.append(bot_f)

                tk.Label(bot_f, text=f"${prod['price']:,}/{prod['unit']}",
                         bg=card_bg, fg=BLUE,
                         font=(FONTB[0], 10, "bold")).pack(side="left")

                def _add(p=prod):
                    c = p["code"]
                    if c not in _cart:
                        _cart[c] = {"product": p, "qty": 0,
                                    "price": p["price"], "category": p.get("category", "")}
                    _cart[c]["qty"] += 1
                    _render_cart()
                    _update_card_visual(c)

                add_btn = ctk.CTkButton(bot_f, text=add_lbl, command=_add,
                                        fg_color=add_col, hover_color=add_hv,
                                        text_color="white",
                                        font=(FONTB[0], 12), width=32, height=28,
                                        corner_radius=4)
                add_btn.pack(side="right")

                _card_widgets[code] = (border_f, inner_f, child_widgets, add_btn)

            _bind_scroll(prod_inner, _prod_scroll)

        prod_cvs.bind("<MouseWheel>", _prod_scroll)

        if CATS:
            _cur_cat[0] = CATS[0]
            for c, b in _cat_btns.items():
                b.config(bg=BLUE if c == CATS[0] else "#dee2e6",
                         fg="white" if c == CATS[0] else "#333")
        _render_products()

        p2_back_btn = ctk.CTkButton(p2, text="← 返回顧客資料",
                                     command=lambda: _go_step(1),
                                     fg_color=GRAY, hover_color="#4d5d6e",
                                     text_color="white", font=FONT_S,
                                     width=120, height=30, corner_radius=6)
        p2_back_btn.pack(anchor="w", padx=10, pady=(0, 6))

        # ════════════════════════════════════════════════════
        # Step 3 — 報價確認
        # ════════════════════════════════════════════════════
        p3 = tk.Frame(step_content, bg=BG)

        p3_cust_border = tk.Frame(p3, bg="#d0d7de")
        p3_cust_border.pack(fill="x", padx=10, pady=(8, 4))
        p3_cust_wrap = tk.Frame(p3_cust_border, bg=BG)
        p3_cust_wrap.pack(fill="x", padx=1, pady=1)
        tk.Label(p3_cust_wrap, text="  客戶資料  ", bg=BG, fg=GRAY,
                 font=FONT_S, anchor="w").pack(anchor="w", padx=6, pady=(4, 0))
        p3_cust_f = tk.Frame(p3_cust_wrap, bg=BG)
        p3_cust_f.pack(fill="x", padx=4, pady=(0, 4))
        p3_cust_f.columnconfigure(1, weight=1)
        p3_cust_f.columnconfigure(3, weight=1)
        _p3_cust_lbls: dict[str, tk.Label] = {}
        for r, (label, key) in enumerate([
            ("公司名稱", "company"), ("聯絡人",  "contact"),
            ("報價單號", "_quote_no"), ("報價日期", "_quote_date"),
        ]):
            col   = (r % 2) * 2
            row_i = r // 2
            tk.Label(p3_cust_f, text=label + "：", bg=BG, fg=GRAY,
                     font=FONT_S, anchor="e"
                     ).grid(row=row_i, column=col, sticky="e", padx=(6, 2), pady=2)
            lbl = tk.Label(p3_cust_f, text="—", bg=BG, fg="#1a1a1a",
                           font=(FONT_S[0], 9, "bold"), anchor="w")
            lbl.grid(row=row_i, column=col+1, sticky="ew", padx=(0, 8), pady=2)
            _p3_cust_lbls[key] = lbl

        conf_outer = tk.Frame(p3, bg="#d0d7de")
        conf_outer.pack(fill="x", padx=10, pady=4)
        conf_wrap = tk.Frame(conf_outer, bg=BG)
        conf_wrap.pack(fill="x", padx=1, pady=1)
        tk.Label(conf_wrap, text="  報價品項  ", bg=BG, fg=GRAY,
                 font=FONT_S, anchor="w").pack(anchor="w", padx=6, pady=(4, 0))
        conf_inner = tk.Frame(conf_wrap, bg=BG)
        conf_inner.pack(fill="x", padx=4, pady=(0, 4))
        _ccols3 = ("code", "name", "qty", "price", "subtotal")
        conf_tree = ttk.Treeview(conf_inner, columns=_ccols3, show="headings",
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
        conf_tree.pack(fill="x")

        sum_f = tk.Frame(p3, bg="#f0f4f8", relief="solid", bd=1)
        sum_f.pack(fill="x", padx=10, pady=(0, 4))
        p3_sub_lbl = tk.Label(sum_f, text="小計：—", bg="#f0f4f8",
                               fg=GRAY, font=FONT_S, anchor="e")
        p3_sub_lbl.pack(fill="x", padx=12, pady=(6, 1))
        p3_tax_lbl = tk.Label(sum_f, text="營業稅 5%：—", bg="#f0f4f8",
                               fg=GRAY, font=FONT_S, anchor="e")
        p3_tax_lbl.pack(fill="x", padx=12, pady=1)
        p3_tot_lbl = tk.Label(sum_f, text="應收總金額：—", bg="#f0f4f8",
                               fg=BLUE, font=(FONTB[0], 11, "bold"), anchor="e")
        p3_tot_lbl.pack(fill="x", padx=12, pady=(1, 6))

        p3_out_lbl = tk.Label(p3, text="", bg=BG, fg=GRAY,
                              font=FONT_S, anchor="w", wraplength=620,
                              justify="left")
        p3_out_lbl.pack(fill="x", padx=10)

        # 歷史區段容器（expand 填滿剩餘空間）
        hist_container = tk.Frame(p3, bg=BG)
        hist_container.pack(fill="both", expand=True, padx=0, pady=0)

        # 歷史相同組合
        hist_outer, hist_lf = _mk_lf(hist_container, "📋  歷史相同組合報價", BG, FONT_S)
        hist_outer.pack(fill="x", padx=10, pady=(6, 2))

        hist_cols = ("date", "customer", "total")
        hist_tree = ttk.Treeview(hist_lf, columns=hist_cols, show="headings", height=4)
        hist_tree.heading("date",     text="日期")
        hist_tree.heading("customer", text="客戶")
        hist_tree.heading("total",    text="應收總金額（含稅）")
        hist_tree.column("date",     width=100, anchor="center")
        hist_tree.column("customer", width=200, anchor="w")
        hist_tree.column("total",    width=140, anchor="e")
        hist_tree.pack(fill="x", padx=4, pady=(4, 2))

        hist_hint = ctk.CTkLabel(hist_lf, text="查詢中…", fg_color="transparent",
                                  font=("Microsoft JhengHei UI", 8),
                                  text_color=GRAY, anchor="w")
        hist_hint.pack(anchor="w", padx=6, pady=(0, 4))

        def _show_detail_dlg(q, items_d):
            dlg = ctk.CTkToplevel(self)
            dlg.title(f"報價單明細 — {q['quote_no']}")
            dlg.configure(fg_color=BG)
            dlg.after(100, dlg.grab_set)
            dlg.geometry("680x480")
            info_outer, info_lf = _mk_lf(dlg, "基本資訊", BG, FONTB)
            info_outer.pack(fill="x", padx=12, pady=(12, 4))
            info_lf.columnconfigure(1, weight=1)
            info_lf.columnconfigure(3, weight=1)
            FONT_S2 = ("Microsoft JhengHei UI", 9)
            for r, (l1, v1, l2, v2) in enumerate([
                ("報價單號", q["quote_no"],       "日期",   q["date"]),
                ("客戶",     q["customer"],        "聯絡人", q["contact"] or "—"),
                ("電話",     q["phone"] or "—",   "總金額", f"{q['total']:,.0f}"),
            ]):
                ctk.CTkLabel(info_lf, text=l1+"：", fg_color="transparent",
                              font=FONT_S2, text_color=GRAY, anchor="e"
                              ).grid(row=r, column=0, sticky="e", padx=(8,2), pady=3)
                ctk.CTkLabel(info_lf, text=v1, fg_color="transparent",
                              font=FONT_S2, anchor="w"
                              ).grid(row=r, column=1, sticky="w", pady=3)
                ctk.CTkLabel(info_lf, text=l2+"：", fg_color="transparent",
                              font=FONT_S2, text_color=GRAY, anchor="e"
                              ).grid(row=r, column=2, sticky="e", padx=(16,2), pady=3)
                ctk.CTkLabel(info_lf, text=v2, fg_color="transparent",
                              font=FONT_S2, anchor="w"
                              ).grid(row=r, column=3, sticky="w", pady=3)
            item_outer2, item_lf = _mk_lf(dlg, "品項明細", BG, FONTB)
            item_outer2.pack(fill="both", expand=True, padx=12, pady=(0, 8))
            icols = ("seq", "code", "name", "qty", "unit", "unit_price", "subtotal")
            itree = ttk.Treeview(item_lf, columns=icols, show="headings", height=10)
            for col, hdr, w, anc in [
                ("seq","#",35,"center"),("code","型號",110,"w"),("name","品名",150,"w"),
                ("qty","數量",50,"center"),("unit","單位",50,"center"),
                ("unit_price","單價",90,"e"),("subtotal","小計",90,"e"),
            ]:
                itree.heading(col, text=hdr)
                itree.column(col, width=w, anchor=anc)
            isb = ctk.CTkScrollbar(item_lf, orientation="vertical", command=itree.yview)
            itree.configure(yscrollcommand=isb.set)
            isb.pack(side="right", fill="y")
            itree.pack(fill="both", expand=True, padx=4, pady=4)
            for it in items_d:
                itree.insert("", "end", values=(
                    it["seq"], it["code"], it["name"], it["qty"], it["unit"],
                    f"{it['unit_price']:,.0f}", f"{it['subtotal']:,.0f}",
                ))
            ctk.CTkButton(dlg, text="關閉", command=dlg.destroy,
                           fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                           font=FONT, width=100, height=34, corner_radius=6
                           ).pack(pady=(0, 10))

        def _show_hist_detail(event=None):
            sel = hist_tree.selection()
            if not sel: return
            quote_id = hist_tree.item(sel[0])["tags"][0]
            from core.repository import get_quote
            data = get_quote(int(quote_id))
            if not data: return
            _show_detail_dlg(data["quote"], data["items"])

        hist_tree.bind("<Double-1>", _show_hist_detail)

        sim_outer, sim_lf = _mk_lf(hist_container, "🔍  相似組合報價（多或少一個品號）", BG, FONT_S)
        sim_outer.pack(fill="x", padx=10, pady=(4, 2))

        sim_cols = ("similarity", "date", "customer", "total")
        sim_tree = ttk.Treeview(sim_lf, columns=sim_cols, show="headings", height=4)
        sim_tree.heading("similarity", text="差異")
        sim_tree.heading("date",       text="日期")
        sim_tree.heading("customer",   text="客戶")
        sim_tree.heading("total",      text="應收總金額（含稅）")
        sim_tree.column("similarity", width=120, anchor="center")
        sim_tree.column("date",       width=100, anchor="center")
        sim_tree.column("customer",   width=180, anchor="w")
        sim_tree.column("total",      width=140, anchor="e")
        sim_tree.pack(fill="x", padx=4, pady=(4, 2))

        sim_hint = ctk.CTkLabel(sim_lf, text="", fg_color="transparent",
                                 font=("Microsoft JhengHei UI", 8),
                                 text_color=GRAY, anchor="w")
        sim_hint.pack(anchor="w", padx=6, pady=(0, 4))

        def _show_sim_detail(event=None):
            sel = sim_tree.selection()
            if not sel: return
            quote_id = sim_tree.item(sel[0])["tags"][0]
            from core.repository import get_quote
            data = get_quote(int(quote_id))
            if not data: return
            _show_detail_dlg(data["quote"], data["items"])

        sim_tree.bind("<Double-1>", _show_sim_detail)

        def _refresh_history_panel(cart_items):
            from core.repository import find_same_combination
            codes = [it["product"]["code"] for it in cart_items
                     if it.get("product", {}).get("code")]
            hist_tree.delete(*hist_tree.get_children())
            if not codes:
                hist_hint.configure(text="購物車無品項")
                return
            try:
                results = find_same_combination(codes)
            except Exception:
                hist_hint.configure(text="（資料庫未連線或尚無記錄）")
                return
            if not results:
                hist_hint.configure(text="無相同品號組合的歷史記錄")
            else:
                totals = []
                for r in results:
                    hist_tree.insert("", "end", values=(
                        r["date"], r["customer"], f"${r['total']:,.0f}"
                    ), tags=(r["id"],))
                    totals.append(r["total"])
                avg = sum(totals) / len(totals)
                hist_hint.configure(
                    text=f"共 {len(results)} 筆  ｜  最近：${totals[0]:,.0f}  ｜  平均：${avg:,.0f}")

            sim_tree.delete(*sim_tree.get_children())
            try:
                from core.repository import find_similar_combinations
                sim_results = find_similar_combinations(codes)
            except Exception:
                sim_hint.configure(text="（資料庫未連線或尚無記錄）")
                return
            if not sim_results:
                sim_hint.configure(text="無相似組合的歷史記錄")
                return
            for r in sim_results:
                sim_tree.insert("", "end", values=(
                    r["similarity"], r["date"], r["customer"], f"${r['total']:,.0f}"
                ), tags=(r["id"],))
            sim_hint.configure(text=f"共 {len(sim_results)} 筆")

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
            p3_sub_lbl.configure(text=f"小計：{_fmt(sub)}" + (f" + 運費 {_fmt(freight)}" if freight else ""))
            p3_tax_lbl.configure(text=f"營業稅 5%：{_fmt(tax)}")
            p3_tot_lbl.configure(text=f"應收總金額：{_fmt(sub + freight + tax)}")
            _p3_cust_lbls["company"].configure(text=_customer.get("company", "—"))
            _p3_cust_lbls["contact"].configure(text=_customer.get("contact", "—"))
            _p3_cust_lbls["_quote_no"].configure(text=quote_no_var.get())
            _p3_cust_lbls["_quote_date"].configure(text=date_entry.get_date().strftime("%Y/%m/%d"))
            _refresh_history_panel(items)

        def _generate():
            from core.generator_quote import generate_quote_from_cart
            if not _cart:
                messagebox.showwarning("購物車空白", "請先選取品項", parent=parent); return
            if not _customer.get("company"):
                messagebox.showwarning("無客戶資料", "請先在步驟①選取 Trello 卡片", parent=parent); return
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
                {**{"code":     v["product"]["code"],
                    "name":     v["product"]["name"],
                    "spec":     v["product"].get("spec", ""),
                    "unit":     v["product"]["unit"],
                    "qty":      v["qty"],
                    "price":    v["price"],
                    "category": v["product"].get("category", "")}}
                for v in _cart.values() if v["qty"] > 0
            ]
            q_date   = date_entry.get_date()
            quote_no = quote_no_var.get().strip()
            out_dir  = self._get_path("output_quote")

            p3_out_lbl.configure(text="生成中…", fg=GRAY)
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
                    quote_type=_qt,
                    card_title=_card[0].get("name", "") if _card[0] else "")
                p3_out_lbl.configure(text=f"✔  已生成：{out_path}", fg=GREEN)

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
                p3_out_lbl.configure(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("生成失敗", str(e), parent=parent)

        p3_back_btn = ctk.CTkButton(p3, text="← 返回選取品項",
                                     command=lambda: _go_step(2),
                                     fg_color=GRAY, hover_color="#4d5d6e",
                                     text_color="white", font=FONT_S,
                                     width=120, height=30, corner_radius=6)
        p3_back_btn.pack(anchor="w", padx=10, pady=(0, 4))

        # ════════════════════════════════════════════════════
        # Step 切換
        # ════════════════════════════════════════════════════
        # Step 1 不需要 expand（內容固定高度），Step 2/3 需要 expand 讓產品格展開
        _panels = {1: p1, 2: p2, 3: p3}

        def _go_step(n: int):
            for i, btn in enumerate(_step_btns, 1):
                if i == n:
                    btn.configure(fg_color=BLUE, text_color="white")
                elif i < n:
                    btn.configure(fg_color="#d5e8f5", text_color=BLUE)
                else:
                    btn.configure(fg_color="#e8ecf0", text_color=GRAY)
            for i, panel in _panels.items():
                if i == n:
                    panel.pack(fill="both", expand=True)
                else:
                    panel.pack_forget()
            if n == 1:
                right_border.pack_forget()
                cart_panel.pack_forget()
                p1_footer.pack(side="bottom", fill="x")
            else:
                p1_footer.pack_forget()
                right_border.pack(side="right", fill="y")
                cart_panel.pack(side="right", fill="y")
            if n == 3:
                _build_confirm()
                checkout_btn.configure(
                    text="📄  生成報價單 .xlsx",
                    fg_color=GREEN,
                    command=_generate,
                )
            else:
                checkout_btn.configure(
                    text="確認報價內容 →",
                    fg_color=BLUE,
                    command=lambda: _go_step(3),
                )

        checkout_btn.configure(command=lambda: _go_step(3))
        _render_cart()
        right_border.pack_forget()
        cart_panel.pack_forget()
        _go_step(1)
