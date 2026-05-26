"""
mixin_documents.py — 出貨單、驗機單、維修單、維修掛件 頁籤 mixin
"""
import os, sys, subprocess
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
from pathlib import Path


class _DocumentsTab:
    """Mixin providing shipping, inspection, fix, and tag tab builders + callbacks."""

    # ── Tab 1：出貨單 ─────────────────────────────────────────
    def _build_tab_shipping(self, parent, PAD, FONT, FONTB, BG):
        # 生成按鈕（先 pack bottom 確保不被擠掉）
        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(side="bottom", fill="x", padx=12)
        tk.Button(bb, text="＋ 新增", command=self._add_row,
                  bg="#27ae60", fg="white", relief="flat",
                  font=FONT, padx=10, pady=3).pack(side="left", padx=(0, 6))
        tk.Button(bb, text="－ 刪除", command=self._del_row,
                  bg="#c0392b", fg="white", relief="flat",
                  font=FONT, padx=10, pady=3).pack(side="left")
        tk.Button(bb, text="⬇  生成出貨單", command=self._generate,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 11, "bold"),
                  padx=16, pady=6).pack(side="right")

        # 欄位區
        mid = tk.Frame(parent, bg=BG)
        mid.pack(fill="x", padx=12, pady=6)
        mid.columnconfigure(0, weight=1)
        mid.columnconfigure(1, weight=1)

        lf = tk.LabelFrame(mid, text="從報價單讀入", bg=BG, font=FONTB)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        lf.columnconfigure(1, weight=1)

        self._read_vars = {}
        for i, (label, key) in enumerate([
            ("客戶名稱", "customer"), ("聯絡電話", "phone"),
            ("聯絡人",   "contact"),  ("地址",     "address"),
            ("報價單號", "quote_no"), ("報價日期", "quote_date"),
        ]):
            tk.Label(lf, text=label + "：", bg=BG, anchor="w", font=FONT
                     ).grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value="—")
            tk.Entry(lf, textvariable=var, font=FONT
                     ).grid(row=i, column=1, sticky="ew", **PAD)
            self._read_vars[key] = var

        rf = tk.LabelFrame(mid, text="補填欄位", bg=BG, font=FONTB)
        rf.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        rf.columnconfigure(1, weight=1)

        self._fill_vars = {}
        for i, (label, key, default) in enumerate([
            ("出貨日期", "ship_date", datetime.today().strftime("%Y/%m/%d")),
            ("銷貨單號", "sale_no",   ""),
            ("附註",     "note",      ""),
        ]):
            tk.Label(rf, text=label + "：", bg=BG, anchor="w", font=FONT
                     ).grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value=default)
            tk.Entry(rf, textvariable=var, font=FONT
                     ).grid(row=i, column=1, sticky="ew", **PAD)
            self._fill_vars[key] = var

        tk.Label(rf, text="製表人員：", bg=BG, anchor="w", font=FONT
                 ).grid(row=3, column=0, sticky="w", **PAD)
        op_f = tk.Frame(rf, bg=BG)
        op_f.grid(row=3, column=1, sticky="ew", **PAD)
        self._operator_var = tk.StringVar()
        self._operator_cb  = ttk.Combobox(op_f, textvariable=self._operator_var,
                                           values=self._config["operators"],
                                           width=12, font=FONT, state="readonly")
        if self._config["operators"]:
            self._operator_var.set(self._config["operators"][0])
        self._operator_cb.pack(side="left")
        tk.Button(op_f, text="＋", command=self._add_operator,
                  bg="#27ae60", fg="white", relief="flat",
                  font=FONT, width=3).pack(side="left", padx=(4, 0))
        tk.Button(op_f, text="－", command=self._del_operator,
                  bg="#c0392b", fg="white", relief="flat",
                  font=FONT, width=3).pack(side="left", padx=(2, 0))

        tk.Label(rf, text="發票方式：", bg=BG, anchor="w", font=FONT
                 ).grid(row=4, column=0, sticky="w", **PAD)
        inv_f = tk.Frame(rf, bg=BG)
        inv_f.grid(row=4, column=1, sticky="w", **PAD)
        self._invoice_var = tk.StringVar(value="尚未確認")
        for lbl, val in [("尚未確認", "尚未確認"), ("隨貨", "隨貨"), ("直寄", "直寄")]:
            tk.Radiobutton(inv_f, text=lbl, variable=self._invoice_var,
                           value=val, bg=BG, font=FONT,
                           activebackground=BG).pack(side="left", padx=(0, 8))

        # 品項列表
        tf = tk.LabelFrame(parent, text="品項列表（雙擊儲存格可編輯）",
                           bg=BG, font=FONTB)
        tf.pack(fill="both", expand=True, padx=12, pady=4)

        cols     = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
        col_lbls = ("序號", "品名 / 規格", "數量", "單位", "單價", "小計")
        col_ws   = (45, 330, 65, 65, 85, 85)

        self._tree = ttk.Treeview(tf, columns=cols, show="headings",
                                   selectmode="browse", height=8)
        for col, lbl, w in zip(cols, col_lbls, col_ws):
            self._tree.heading(col, text=lbl)
            self._tree.column(col, width=w, minwidth=w, anchor="center")
        self._tree.column("name", anchor="w")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._tree.bind("<Double-1>", self._on_cell_dclick)

    # ── Tab 2：驗機單 ─────────────────────────────────────────
    def _build_tab_inspection(self, parent, PAD, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        info = tk.LabelFrame(parent, text="說明", bg=BG, font=FONTB)
        info.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(info, text="載入報價單後，點擊下方按鈕自動生成驗機單 Excel 及 Word。",
                 bg=BG, font=FONT, fg=GRAY).pack(padx=12, pady=8, anchor="w")

        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(fill="x", padx=12, pady=4)
        tk.Label(pf, text="輸出路徑：", bg="#e8ecf0", font=FONT_S, fg=GRAY,
                 anchor="w", width=12).pack(side="left", padx=8, pady=6)
        tk.Label(pf, text="（依⚙路徑設定）",
                 bg="#e8ecf0", font=FONT_S, fg=GRAY).pack(side="left", pady=6)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(side="bottom", fill="x", padx=12)
        tk.Button(bb, text="🔍  生成驗機單", command=self._generate_inspection,
                  bg="#6c3483", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=10).pack(fill="x")

    # ── Tab 3：維修單 ─────────────────────────────────────────
    def _build_tab_fix(self, parent, PAD, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        info = tk.LabelFrame(parent, text="說明", bg=BG, font=FONTB)
        info.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(info, text="載入報價單後，點擊下方按鈕生成維修單。",
                 bg=BG, font=FONT, fg=GRAY).pack(padx=12, pady=8, anchor="w")

        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(fill="x", padx=12, pady=4)
        row = tk.Frame(pf, bg="#e8ecf0")
        row.pack(fill="x")
        tk.Label(row, text="輸出路徑：", bg="#e8ecf0", font=FONT_S,
                 fg=GRAY, anchor="w", width=10).pack(side="left", padx=8)
        tk.Label(row, text="（依⚙路徑設定）",
                 bg="#e8ecf0", font=FONT_S, fg=GRAY).pack(side="left")

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(side="bottom", fill="x", padx=12)
        tk.Button(bb, text="🔧  生成維修單", command=self._generate_fix,
                  bg="#d68910", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=10).pack(fill="x")

    # ── Tab 4：維修掛件 ───────────────────────────────────────
    def _build_tab_tag(self, parent, PAD, FONT, FONTB, BG):
        from tkcalendar import DateEntry
        GRAY = "#5d6d7e"

        tgf = tk.LabelFrame(parent, text="維修掛件資料", bg=BG, font=FONTB)
        tgf.pack(fill="x", padx=12, pady=(12, 4))
        tgf.columnconfigure(1, weight=1)
        tgf.columnconfigure(3, weight=1)

        self._tag_vars = {}

        # 客戶名稱
        cust_var = tk.StringVar()
        self._tag_vars["customer"] = cust_var
        tk.Label(tgf, text="客戶名稱：", bg=BG, anchor="w", font=FONT
                 ).grid(row=0, column=0, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=cust_var, font=FONT
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

        tk.Button(tgf, text="從報價單帶入", command=_load_customer,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT, padx=6).grid(row=0, column=2, padx=8, pady=2)

        # No.
        no_var = tk.StringVar(value="1")
        self._tag_vars["no"] = no_var
        tk.Label(tgf, text="No.：", bg=BG, anchor="w", font=FONT
                 ).grid(row=1, column=0, sticky="w", padx=8, pady=2)
        ttk.Combobox(tgf, textvariable=no_var,
                     values=[str(i) for i in range(1, 21)],
                     width=8, font=FONT).grid(row=1, column=1, sticky="w", padx=8, pady=2)

        # 品號
        self._tag_partno_var = tk.StringVar()
        self._tag_vars["part_no"] = self._tag_partno_var
        tk.Label(tgf, text="品號：", bg=BG, anchor="w", font=FONT
                 ).grid(row=1, column=2, sticky="w", padx=8, pady=2)
        self._tag_partno_cb = ttk.Combobox(tgf, textvariable=self._tag_partno_var,
                                            font=FONT, width=20)
        self._tag_partno_cb.grid(row=1, column=3, sticky="ew", padx=8, pady=2)

        # 序號 / 拉回日期
        seq_var = tk.StringVar()
        self._tag_vars["seq_no"] = seq_var
        tk.Label(tgf, text="序號：", bg=BG, anchor="w", font=FONT
                 ).grid(row=2, column=0, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=seq_var, font=FONT
                 ).grid(row=2, column=1, sticky="ew", padx=8, pady=2)

        tk.Label(tgf, text="拉回：", bg=BG, anchor="w", font=FONT
                 ).grid(row=2, column=2, sticky="w", padx=8, pady=2)
        self._tag_date_entry = DateEntry(
            tgf, font=FONT, date_pattern="yyyy/mm/dd",
            background="#2e86c1", foreground="white", width=14)
        self._tag_date_entry.grid(row=2, column=3, sticky="w", padx=8, pady=2)

        # 問題 / 維修狀況
        prob_var = tk.StringVar()
        self._tag_vars["problem"] = prob_var
        tk.Label(tgf, text="問題：", bg=BG, anchor="w", font=FONT
                 ).grid(row=3, column=0, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=prob_var, font=FONT
                 ).grid(row=3, column=1, sticky="ew", padx=8, pady=2)

        status_var = tk.StringVar()
        self._tag_vars["repair_status"] = status_var
        tk.Label(tgf, text="維修狀況：", bg=BG, anchor="w", font=FONT
                 ).grid(row=3, column=2, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=status_var, font=FONT
                 ).grid(row=3, column=3, sticky="ew", padx=8, pady=2)

        # 輸出路徑
        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(fill="x", padx=12, pady=4)
        tk.Label(pf, text="輸出路徑：", bg="#e8ecf0",
                 font=("Microsoft JhengHei UI", 9), fg=GRAY,
                 anchor="w", width=10).pack(side="left", padx=8, pady=6)
        tk.Label(pf, text=r"Z:\待維修機台資料",
                 bg="#e8ecf0", font=("Microsoft JhengHei UI", 9), fg=GRAY
                 ).pack(side="left", pady=6)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(side="bottom", fill="x", padx=12)
        tk.Button(bb, text="📋  生成維修掛件", command=self._generate_tag_doc,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=10).pack(fill="x")

    # ── Callbacks ─────────────────────────────────────────────

    def _on_cell_dclick(self, event):
        item_id = self._tree.identify_row(event.y)
        col_id  = self._tree.identify_column(event.x)
        if not item_id or not col_id:
            return
        col_idx  = int(col_id.replace("#", "")) - 1
        col_keys = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
        col_disp = ("序號", "品名", "數量", "單位", "單價", "小計")
        old_val  = self._tree.item(item_id, "values")[col_idx]

        bbox = self._tree.bbox(item_id, col_id)
        if not bbox:
            return
        x, y, _, h = bbox

        pop = tk.Toplevel(self)
        pop.title(f"編輯「{col_disp[col_idx]}」")
        pop.geometry(f"300x80+{self.winfo_rootx()+x}+{self.winfo_rooty()+y+h}")
        pop.grab_set()

        var   = tk.StringVar(value=old_val)
        entry = tk.Entry(pop, textvariable=var, font=("Microsoft JhengHei UI", 11))
        entry.pack(fill="x", padx=10, pady=8)
        entry.select_range(0, "end")
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
        tk.Button(pop, text="確認", command=save,
                  bg="#2e86c1", fg="white", relief="flat").pack(pady=2)

    def _add_row(self):
        n = len(self._tree.get_children()) + 1
        self._tree.insert("", "end", values=(n, "新品項", 1, "組", 0, 0))

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
        pop = tk.Toplevel(self)
        pop.title("新增製表人員")
        pop.geometry("260x80")
        pop.grab_set()
        var   = tk.StringVar()
        entry = tk.Entry(pop, textvariable=var, font=("Microsoft JhengHei UI", 11))
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
        tk.Button(pop, text="新增", command=save,
                  bg="#27ae60", fg="white", relief="flat").pack(pady=2)

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
            result = generate(self._parsed_data, extra, output_dir=self._OUT_SHIPPING)
            paths  = result if isinstance(result, list) else [result]
            msg    = "\n".join(str(p) for p in paths)
            if messagebox.askyesno("生成成功",
                    f"已生成 {len(paths)} 份出貨單：\n{msg}\n\n是否立即開啟？"):
                for p in paths:
                    os.startfile(p) if sys.platform == "win32" else subprocess.run(["open", str(p)])
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))

    def _generate_inspection(self):
        from core.generator_inspection import generate_inspection
        if not self._parsed_data or not self._src_path:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return
        try:
            excel_path, word_paths = generate_inspection(
                self._src_path, self._parsed_data,
                output_dir=self._get_path("output_inspection"))
            msg = f"驗機單 Excel 已儲存至：\n{excel_path}"
            if word_paths:
                msg += f"\n\n驗機單 Word（共 {len(word_paths)} 份）："
                for wp in word_paths:
                    msg += f"\n  {wp.name}"
            if messagebox.askyesno("生成成功", msg + "\n\n是否立即開啟？"):
                os.startfile(excel_path)
                for wp in word_paths:
                    os.startfile(wp)
        except Exception as e:
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
            if messagebox.askyesno("生成成功",
                    f"已生成 {len(paths)} 份檔案：\n{msg}\n\n是否立即開啟？"):
                for p in paths:
                    os.startfile(p) if sys.platform == "win32" else subprocess.run(["open", str(p)])
        except Exception as e:
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
            if messagebox.askyesno("生成成功",
                    f"維修掛件已生成：\n{path}\n\n是否立即開啟？"):
                os.startfile(str(path)) if sys.platform == "win32" else subprocess.run(["open", str(path)])
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))
