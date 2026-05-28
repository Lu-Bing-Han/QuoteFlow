"""
mixin_history.py — 報價記錄查詢頁籤
支援用品號組合（模糊比對）+ 客戶名稱搜尋歷史報價單。
"""
import tkinter as tk
from tkinter import messagebox, ttk


class _HistoryTab:

    def _build_tab_history(self, parent, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)
        GREEN  = "#1e8449"

        # ════════════════════════════════════════════════
        #  搜尋區
        # ════════════════════════════════════════════════
        search_lf = tk.LabelFrame(parent, text="搜尋條件", bg=BG, font=FONTB)
        search_lf.pack(fill="x", padx=12, pady=(12, 6))

        # ── 品號輸入 ──────────────────────────────────
        code_row = tk.Frame(search_lf, bg=BG)
        code_row.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(code_row, text="品號組合：", bg=BG, font=FONT_S, fg=GRAY,
                 width=8, anchor="w").pack(side="left")

        # 用來存放已加入的品號 tag
        _code_tags: list[str] = []
        _tag_frames: list[tk.Frame] = []

        tag_area = tk.Frame(search_lf, bg=BG)
        tag_area.pack(fill="x", padx=8, pady=(0, 4))

        code_input_var = tk.StringVar()
        code_entry = tk.Entry(code_row, textvariable=code_input_var,
                              font=FONT_S, width=18)
        code_entry.pack(side="left", padx=(0, 6))

        def _add_code_tag(event=None):
            raw = code_input_var.get().strip()
            if not raw or raw in _code_tags:
                code_input_var.set("")
                return
            _code_tags.append(raw)
            # 建立 tag chip
            chip = tk.Frame(tag_area, bg="#d5e8f7", bd=1, relief="solid")
            chip.pack(side="left", padx=(0, 4), pady=2)
            tk.Label(chip, text=raw, bg="#d5e8f7", font=FONT_S,
                     padx=4).pack(side="left")
            def _remove(c=raw, f=chip):
                _code_tags.remove(c)
                f.destroy()
            tk.Button(chip, text="✕", bg="#d5e8f7", fg="#c0392b",
                      relief="flat", font=("Arial", 7), padx=2,
                      command=_remove).pack(side="left")
            _tag_frames.append(chip)
            code_input_var.set("")

        code_entry.bind("<Return>", _add_code_tag)
        tk.Button(code_row, text="＋ 加入", command=_add_code_tag,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=8).pack(side="left")

        tk.Label(code_row, text="（輸入品號後按 Enter 或＋加入，可加多個）",
                 bg=BG, fg=GRAY, font=("Microsoft JhengHei UI", 8)
                 ).pack(side="left", padx=(8, 0))

        # ── 客戶名稱 ──────────────────────────────────
        cust_row = tk.Frame(search_lf, bg=BG)
        cust_row.pack(fill="x", padx=8, pady=(4, 8))
        tk.Label(cust_row, text="客戶名稱：", bg=BG, font=FONT_S, fg=GRAY,
                 width=8, anchor="w").pack(side="left")
        cust_var = tk.StringVar()
        tk.Entry(cust_row, textvariable=cust_var, font=FONT_S, width=24
                 ).pack(side="left", padx=(0, 12))

        tk.Button(cust_row, text="🔍  搜尋", font=FONTB,
                  bg="#1a5276", fg="white", relief="flat",
                  padx=14, pady=4,
                  command=lambda: _do_search()
                  ).pack(side="left")
        tk.Button(cust_row, text="清除", font=FONT_S,
                  bg=GRAY, fg="white", relief="flat", padx=8,
                  command=lambda: _clear()
                  ).pack(side="left", padx=(6, 0))

        # ════════════════════════════════════════════════
        #  結果列表
        # ════════════════════════════════════════════════
        result_lf = tk.LabelFrame(parent, text="搜尋結果", bg=BG, font=FONTB)
        result_lf.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        cols = ("quote_no", "date", "customer", "contact", "total")
        tree = ttk.Treeview(result_lf, columns=cols, show="headings",
                            selectmode="browse")
        tree.heading("quote_no", text="報價單號")
        tree.heading("date",     text="日期")
        tree.heading("customer", text="客戶")
        tree.heading("contact",  text="聯絡人")
        tree.heading("total",    text="總金額")
        tree.column("quote_no", width=130, anchor="center")
        tree.column("date",     width=100, anchor="center")
        tree.column("customer", width=200, anchor="w")
        tree.column("contact",  width=90,  anchor="center")
        tree.column("total",    width=100, anchor="e")

        sb = ttk.Scrollbar(result_lf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=4, pady=4)

        count_lbl = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY)
        count_lbl.pack(anchor="w", padx=14)

        # ── 點兩下查看明細 ─────────────────────────────
        def _on_dclick(event):
            sel = tree.selection()
            if not sel:
                return
            quote_id = tree.item(sel[0])["tags"][0]
            _show_detail(int(quote_id))

        tree.bind("<Double-1>", _on_dclick)

        # ════════════════════════════════════════════════
        #  搜尋邏輯
        # ════════════════════════════════════════════════
        def _do_search():
            from core.repository import search_quotes_by_codes
            results = search_quotes_by_codes(
                codes    = list(_code_tags),
                customer = cust_var.get().strip(),
            )
            tree.delete(*tree.get_children())
            for r in results:
                total_str = f"{r['total']:,.0f}"
                tree.insert("", "end",
                            values=(r["quote_no"], r["date"], r["customer"],
                                    r["contact"] or "—", total_str),
                            tags=(r["id"],))
            n = len(results)
            count_lbl.config(text=f"共找到 {n} 筆" if n else "查無資料")

        def _clear():
            for chip in list(_tag_frames):
                chip.destroy()
            _tag_frames.clear()
            _code_tags.clear()
            cust_var.set("")
            tree.delete(*tree.get_children())
            count_lbl.config(text="")

        # ════════════════════════════════════════════════
        #  明細彈窗
        # ════════════════════════════════════════════════
        def _show_detail(quote_id: int):
            from core.repository import get_quote
            data = get_quote(quote_id)
            if not data:
                return
            q, items = data["quote"], data["items"]

            dlg = tk.Toplevel(self)
            dlg.title(f"報價單明細 — {q['quote_no']}")
            dlg.configure(bg=BG)
            dlg.grab_set()
            dlg.geometry("680x480")

            # 基本資訊
            info_lf = tk.LabelFrame(dlg, text="基本資訊", bg=BG, font=FONTB)
            info_lf.pack(fill="x", padx=12, pady=(12, 4))
            info_lf.columnconfigure(1, weight=1)
            info_lf.columnconfigure(3, weight=1)

            fields = [
                ("報價單號", q["quote_no"], "日期",    q["date"]),
                ("客戶",     q["customer"], "聯絡人",  q["contact"] or "—"),
                ("電話",     q["phone"] or "—", "總金額", f"{q['total']:,.0f}"),
            ]
            for r, (l1, v1, l2, v2) in enumerate(fields):
                tk.Label(info_lf, text=l1 + "：", bg=BG, font=FONT_S,
                         fg=GRAY, anchor="e").grid(row=r, column=0, sticky="e",
                                                   padx=(8, 2), pady=3)
                tk.Label(info_lf, text=v1, bg=BG, font=FONT_S,
                         anchor="w").grid(row=r, column=1, sticky="w", pady=3)
                tk.Label(info_lf, text=l2 + "：", bg=BG, font=FONT_S,
                         fg=GRAY, anchor="e").grid(row=r, column=2, sticky="e",
                                                   padx=(16, 2), pady=3)
                tk.Label(info_lf, text=v2, bg=BG, font=FONT_S,
                         anchor="w").grid(row=r, column=3, sticky="w", pady=3)

            # 品項明細
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

            for it in items:
                itree.insert("", "end", values=(
                    it["seq"], it["code"], it["name"],
                    it["qty"], it["unit"],
                    f"{it['unit_price']:,.0f}",
                    f"{it['subtotal']:,.0f}",
                ))

            tk.Button(dlg, text="關閉", command=dlg.destroy,
                      bg=GRAY, fg="white", relief="flat",
                      font=FONT, padx=16, pady=4
                      ).pack(pady=(0, 10))
