"""
mixin_history.py — 報價記錄查詢頁籤
"""
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from ui.app_core import _mk_lf


class _HistoryTab:

    def _build_tab_history(self, parent, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        # ════════════════════════════════════════════════
        #  搜尋區
        # ════════════════════════════════════════════════
        search_outer, search_lf = _mk_lf(parent, "搜尋條件", BG, FONTB)
        search_outer.pack(fill="x", padx=12, pady=(12, 6))

        # ── 品號輸入 ──────────────────────────────────
        code_row = ctk.CTkFrame(search_lf, fg_color="transparent", corner_radius=0)
        code_row.pack(fill="x", padx=8, pady=(4, 4))

        ctk.CTkLabel(code_row, text="品號組合：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w", width=72).pack(side="left")

        _code_tags:  list[str]          = []
        _tag_frames: list[ctk.CTkFrame] = []

        tag_area = tk.Frame(search_lf, bg=BG)
        tag_area.pack(fill="x", padx=8, pady=(0, 4))

        code_input_var = tk.StringVar()
        code_entry = ctk.CTkEntry(code_row, textvariable=code_input_var,
                                   font=FONT_S, width=150, height=28, corner_radius=4)
        code_entry.pack(side="left", padx=(0, 6))

        def _add_code_tag(event=None):
            raw = code_input_var.get().strip()
            if not raw or raw in _code_tags:
                code_input_var.set("")
                return
            _code_tags.append(raw)
            chip = ctk.CTkFrame(tag_area, fg_color="#d5e8f7", corner_radius=6,
                                 border_width=1, border_color="#aed6f1")
            chip.pack(side="left", padx=(0, 4), pady=2)
            ctk.CTkLabel(chip, text=raw, fg_color="transparent",
                          font=FONT_S, text_color="#1a5276").pack(side="left", padx=(6, 2))
            def _remove(c=raw, f=chip):
                _code_tags.remove(c)
                f.destroy()
            ctk.CTkButton(chip, text="✕", command=_remove,
                           fg_color="transparent", hover_color="#aed6f1",
                           text_color="#c0392b",
                           font=("Arial", 8), width=20, height=20,
                           corner_radius=4).pack(side="left", padx=(0, 2))
            _tag_frames.append(chip)
            code_input_var.set("")

        code_entry.bind("<Return>", _add_code_tag)
        ctk.CTkButton(code_row, text="＋ 加入", command=_add_code_tag,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="left")
        ctk.CTkLabel(code_row, text="（輸入品號後按 Enter 或＋加入，可加多個）",
                      fg_color="transparent",
                      text_color=GRAY, font=("Microsoft JhengHei UI", 8)
                      ).pack(side="left", padx=(8, 0))

        # ── 客戶名稱 ──────────────────────────────────
        cust_row = ctk.CTkFrame(search_lf, fg_color="transparent", corner_radius=0)
        cust_row.pack(fill="x", padx=8, pady=(4, 8))
        ctk.CTkLabel(cust_row, text="客戶名稱：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w", width=72).pack(side="left")
        cust_var = tk.StringVar()
        ctk.CTkEntry(cust_row, textvariable=cust_var, font=FONT_S,
                      width=200, height=28, corner_radius=4
                      ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(cust_row, text="🔍  搜尋", font=FONTB,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       width=100, height=32, corner_radius=6,
                       command=lambda: _do_search()
                       ).pack(side="left")
        ctk.CTkButton(cust_row, text="清除", font=FONT_S,
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       width=60, height=32, corner_radius=6,
                       command=lambda: _clear()
                       ).pack(side="left", padx=(6, 0))

        # ════════════════════════════════════════════════
        #  結果列表
        # ════════════════════════════════════════════════
        result_outer, result_lf = _mk_lf(parent, "搜尋結果", BG, FONTB)
        result_outer.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        ctk.CTkLabel(result_lf, text="💡  雙擊列、右鍵點選，或選取列後按「🔍 查看明細」可查看完整報價單內容",
                      fg_color="transparent", font=FONT_S,
                      text_color="#d68910", anchor="w"
                      ).pack(fill="x", padx=4, pady=(2, 4))

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

        sb = ctk.CTkScrollbar(result_lf, orientation="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=4, pady=4)

        bottom_row = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        bottom_row.pack(fill="x", padx=14, pady=(0, 6))

        count_lbl = ctk.CTkLabel(bottom_row, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w")
        count_lbl.pack(side="left")

        def _selected_quote_id():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("尚未選取", "請先在表格中點選一列，再按「🔍 查看明細」")
                return None
            return int(tree.item(sel[0])["tags"][0])

        def _view_selected():
            quote_id = _selected_quote_id()
            if quote_id is not None:
                _show_detail(quote_id)

        ctk.CTkButton(bottom_row, text="🔍  查看明細", command=_view_selected,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=FONT_S, width=110, height=28, corner_radius=4
                       ).pack(side="right")

        def _on_dclick(event):
            sel = tree.selection()
            if not sel:
                return
            quote_id = tree.item(sel[0])["tags"][0]
            _show_detail(int(quote_id))

        def _on_rclick(event):
            row_id = tree.identify_row(event.y)
            if not row_id:
                return
            tree.selection_set(row_id)
            quote_id = tree.item(row_id)["tags"][0]
            menu = tk.Menu(parent, tearoff=0)
            menu.add_command(label="🔍  查看明細",
                              command=lambda: _show_detail(int(quote_id)))
            menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Double-1>", _on_dclick)
        tree.bind("<Button-3>", _on_rclick)

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
            count_lbl.configure(text=f"共找到 {n} 筆" if n else "查無資料")

        def _clear():
            for chip in list(_tag_frames):
                chip.destroy()
            _tag_frames.clear()
            _code_tags.clear()
            cust_var.set("")
            tree.delete(*tree.get_children())
            count_lbl.configure(text="")

        # ════════════════════════════════════════════════
        #  明細彈窗
        # ════════════════════════════════════════════════
        def _show_detail(quote_id: int):
            from core.repository import get_quote
            data = get_quote(quote_id)
            if not data:
                return
            q, items = data["quote"], data["items"]

            dlg = ctk.CTkToplevel(self)
            dlg.title(f"報價單明細 — {q['quote_no']}")
            dlg.configure(fg_color=BG)
            dlg.after(100, dlg.grab_set)
            dlg.geometry("680x480")

            info_outer, info_lf = _mk_lf(dlg, "基本資訊", BG, FONTB)
            info_outer.pack(fill="x", padx=12, pady=(12, 4))
            info_lf.columnconfigure(1, weight=1)
            info_lf.columnconfigure(3, weight=1)

            fields = [
                ("報價單號", q["quote_no"], "日期",    q["date"]),
                ("客戶",     q["customer"], "聯絡人",  q["contact"] or "—"),
                ("電話",     q["phone"] or "—", "總金額", f"{q['total']:,.0f}"),
            ]
            for r, (l1, v1, l2, v2) in enumerate(fields):
                ctk.CTkLabel(info_lf, text=l1 + "：", fg_color="transparent",
                              font=FONT_S, text_color=GRAY,
                              anchor="e").grid(row=r, column=0, sticky="e",
                                               padx=(8, 2), pady=3)
                ctk.CTkLabel(info_lf, text=v1, fg_color="transparent",
                              font=FONT_S, anchor="w"
                              ).grid(row=r, column=1, sticky="w", pady=3)
                ctk.CTkLabel(info_lf, text=l2 + "：", fg_color="transparent",
                              font=FONT_S, text_color=GRAY,
                              anchor="e").grid(row=r, column=2, sticky="e",
                                               padx=(16, 2), pady=3)
                ctk.CTkLabel(info_lf, text=v2, fg_color="transparent",
                              font=FONT_S, anchor="w"
                              ).grid(row=r, column=3, sticky="w", pady=3)

            item_outer, item_lf = _mk_lf(dlg, "品項明細", BG, FONTB)
            item_outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

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

            isb = ctk.CTkScrollbar(item_lf, orientation="vertical", command=itree.yview)
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

            ctk.CTkButton(dlg, text="關閉", command=dlg.destroy,
                           fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                           font=FONT, width=100, height=34, corner_radius=6
                           ).pack(pady=(0, 10))
