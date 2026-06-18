"""
mixin_documents.py — 出貨單、驗機單、維修單、維修掛件 頁籤 mixin
"""
import os, sys, subprocess
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from datetime import datetime
from pathlib import Path
from ui.app_core import _mk_lf


class _DocumentsTab:
    """Mixin providing shipping, inspection, fix, and tag tab builders + callbacks."""

    # ── Tab 1：出貨單 ─────────────────────────────────────────
    def _build_tab_shipping(self, parent, PAD, FONT, FONTB, BG):
        GRAY = "#5d6d7e"

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
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

        mid = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        mid.pack(fill="x", padx=12, pady=6)
        mid.columnconfigure(0, weight=1)
        mid.columnconfigure(1, weight=1)

        lf_outer, lf = _mk_lf(mid, "從報價單讀入", BG, FONTB)
        lf_outer.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        lf.columnconfigure(1, weight=1)

        self._read_vars = {}
        for i, (label, key) in enumerate([
            ("客戶名稱", "customer"), ("聯絡電話", "phone"),
            ("聯絡人",   "contact"),  ("地址",     "address"),
            ("報價單號", "quote_no"), ("報價日期", "quote_date"),
        ]):
            ctk.CTkLabel(lf, text=label + "：", fg_color="transparent",
                          anchor="w", font=FONT, text_color="#2c3e50"
                          ).grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value="—")
            ctk.CTkEntry(lf, textvariable=var, font=FONT,
                          corner_radius=4, border_width=1
                          ).grid(row=i, column=1, sticky="ew", **PAD)
            self._read_vars[key] = var

        rf_outer, rf = _mk_lf(mid, "補填欄位", BG, FONTB)
        rf_outer.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        rf.columnconfigure(1, weight=1)

        self._fill_vars = {}
        for i, (label, key, default) in enumerate([
            ("出貨日期", "ship_date", datetime.today().strftime("%Y/%m/%d")),
            ("銷貨單號", "sale_no",   ""),
            ("附註",     "note",      ""),
        ]):
            ctk.CTkLabel(rf, text=label + "：", fg_color="transparent",
                          anchor="w", font=FONT, text_color="#2c3e50"
                          ).grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value=default)
            ctk.CTkEntry(rf, textvariable=var, font=FONT,
                          corner_radius=4, border_width=1
                          ).grid(row=i, column=1, sticky="ew", **PAD)
            self._fill_vars[key] = var

        ctk.CTkLabel(rf, text="製表人員：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=3, column=0, sticky="w", **PAD)
        op_f = ctk.CTkFrame(rf, fg_color="transparent", corner_radius=0)
        op_f.grid(row=3, column=1, sticky="ew", **PAD)
        self._operator_var = tk.StringVar()
        self._operator_cb  = ttk.Combobox(op_f, textvariable=self._operator_var,
                                           values=self._config["operators"],
                                           width=12, font=FONT, state="readonly")
        if self._config["operators"]:
            self._operator_var.set(self._config["operators"][0])
        self._operator_cb.pack(side="left")
        ctk.CTkButton(op_f, text="＋", command=self._add_operator,
                       fg_color="#27ae60", hover_color="#1e8449", text_color="white",
                       font=FONT, width=28, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))
        ctk.CTkButton(op_f, text="－", command=self._del_operator,
                       fg_color="#c0392b", hover_color="#a93226", text_color="white",
                       font=FONT, width=28, height=28, corner_radius=4
                       ).pack(side="left", padx=(2, 0))

        ctk.CTkLabel(rf, text="發票方式：", fg_color="transparent",
                      anchor="w", font=FONT, text_color="#2c3e50"
                      ).grid(row=4, column=0, sticky="w", **PAD)
        inv_f = ctk.CTkFrame(rf, fg_color="transparent", corner_radius=0)
        inv_f.grid(row=4, column=1, sticky="w", **PAD)
        self._invoice_var = tk.StringVar(value="尚未確認")
        for lbl, val in [("尚未確認", "尚未確認"), ("隨貨", "隨貨"), ("直寄", "直寄")]:
            ctk.CTkRadioButton(inv_f, text=lbl, variable=self._invoice_var,
                                value=val, font=FONT,
                                fg_color="#1a5276", hover_color="#2e6da4"
                                ).pack(side="left", padx=(0, 8))

        # 品項列表
        tf_outer, tf = _mk_lf(parent, "品項列表", BG, FONTB)
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
