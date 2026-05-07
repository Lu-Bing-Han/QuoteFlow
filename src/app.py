"""
app.py  —  報價單 → 出貨單 / 驗機單 / 維修單 轉換工具 (Tkinter GUI)
"""

import json, os, subprocess, sys, tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from parser import parse
from generator import generate
from generator_inspection import generate_inspection
from generator_fix import generate_fix

CONFIG_PATH = Path(__file__).parent.parent / "config.json"

def _load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"operators": ["小皋"]}

def _save_config(cfg):
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("報價單 → 出貨單 轉換工具｜立善科技")
        self.geometry("960x760")
        self.resizable(True, True)
        self.configure(bg="#f4f6f8")
        self._parsed_data = None
        self._src_path = None
        self._config = _load_config()
        self._build_ui()

    def _build_ui(self):
        PAD   = {"padx": 12, "pady": 4}
        FONT  = ("Microsoft JhengHei", 10)
        FONTB = ("Microsoft JhengHei", 10, "bold")

        top = tk.Frame(self, bg="#1a5276", pady=10)
        top.pack(fill="x")
        tk.Label(top, text="立善科技｜報價單轉出貨單",
                 bg="#1a5276", fg="white",
                 font=("Microsoft JhengHei", 14, "bold")).pack(side="left", padx=16)
        tk.Button(top, text="📂  選擇報價單 .xlsx", command=self._open_file,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=("Microsoft JhengHei", 11), padx=12, pady=4).pack(side="right", padx=16)

        self._file_label = tk.Label(self, text="尚未選擇報價單",
                                    bg="#f4f6f8", fg="#888", font=FONT)
        self._file_label.pack(anchor="w", padx=16, pady=(4, 0))

        mid = tk.Frame(self, bg="#f4f6f8")
        mid.pack(fill="x", padx=12, pady=6)
        mid.columnconfigure(0, weight=1)
        mid.columnconfigure(1, weight=1)

        lf = tk.LabelFrame(mid, text="從報價單讀入", bg="#f4f6f8", font=FONTB)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        lf.columnconfigure(1, weight=1)

        self._read_vars = {}
        READ_FIELDS = [
            ("客戶名稱", "customer"), ("聯絡電話", "phone"),
            ("聯絡人",   "contact"),  ("地址",     "address"),
            ("報價單號", "quote_no"), ("報價日期", "quote_date"),
        ]
        for i, (label, key) in enumerate(READ_FIELDS):
            tk.Label(lf, text=label + "：", bg="#f4f6f8",
                     anchor="w", font=FONT).grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value="—")
            tk.Label(lf, textvariable=var, bg="#eaf4fb", anchor="w",
                     relief="groove", font=FONT
                     ).grid(row=i, column=1, sticky="ew", **PAD)
            self._read_vars[key] = var

        rf = tk.LabelFrame(mid, text="補填欄位", bg="#f4f6f8", font=FONTB)
        rf.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        rf.columnconfigure(1, weight=1)

        self._fill_vars = {}
        FILL_FIELDS = [
            ("出貨日期", "ship_date", datetime.today().strftime("%Y/%m/%d")),
            ("銷貨單號", "sale_no",   ""),
            ("附註",     "note",      ""),
        ]
        for i, (label, key, default) in enumerate(FILL_FIELDS):
            tk.Label(rf, text=label + "：", bg="#f4f6f8",
                     anchor="w", font=FONT).grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value=default)
            tk.Entry(rf, textvariable=var, font=FONT
                     ).grid(row=i, column=1, sticky="ew", **PAD)
            self._fill_vars[key] = var

        tk.Label(rf, text="製表人員：", bg="#f4f6f8",
                 anchor="w", font=FONT).grid(row=3, column=0, sticky="w", **PAD)
        op_f = tk.Frame(rf, bg="#f4f6f8")
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
                  font=FONT, width=3).pack(side="left", padx=(4,0))
        tk.Button(op_f, text="－", command=self._del_operator,
                  bg="#c0392b", fg="white", relief="flat",
                  font=FONT, width=3).pack(side="left", padx=(2,0))

        tk.Label(rf, text="發票方式：", bg="#f4f6f8",
                 anchor="w", font=FONT).grid(row=4, column=0, sticky="w", **PAD)
        inv_f = tk.Frame(rf, bg="#f4f6f8")
        inv_f.grid(row=4, column=1, sticky="w", **PAD)
        self._invoice_var = tk.StringVar(value="尚未確認")
        for lbl, val in [("發票尚未確認", "尚未確認"),
                         ("發票隨貨", "隨貨"),
                         ("發票直寄", "直寄")]:
            tk.Radiobutton(inv_f, text=lbl, variable=self._invoice_var,
                           value=val, bg="#f4f6f8", font=FONT,
                           activebackground="#f4f6f8").pack(side="left", padx=(0,8))

        tf = tk.LabelFrame(self, text="品項列表（雙擊儲存格可編輯）",
                           bg="#f4f6f8", font=FONTB)
        tf.pack(fill="both", expand=True, padx=12, pady=4)

        cols      = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
        col_lbls  = ("序號", "品名 / 規格", "數量", "單位", "單價", "小計")
        col_ws    = (45, 330, 65, 65, 85, 85)

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

        bb = tk.Frame(self, bg="#f4f6f8")
        bb.pack(fill="x", padx=12)
        tk.Button(bb, text="＋ 新增品項", command=self._add_row,
                  bg="#27ae60", fg="white", relief="flat",
                  font=FONT, padx=10, pady=3).pack(side="left", padx=(0,6))
        tk.Button(bb, text="－ 刪除選取", command=self._del_row,
                  bg="#c0392b", fg="white", relief="flat",
                  font=FONT, padx=10, pady=3).pack(side="left")

        bot = tk.Frame(self, bg="#f4f6f8", pady=10)
        bot.pack(fill="x")
        for text, cmd, color in [
            ("⬇  生成出貨單", self._generate,             "#1a5276"),
            ("🔍  生成驗機單", self._generate_inspection,  "#6c3483"),
            ("🔧  生成維修單", self._generate_fix,         "#d68910"),
        ]:
            tk.Button(bot, text=text, command=cmd, bg=color, fg="white",
                      font=("Microsoft JhengHei", 13, "bold"),
                      relief="flat", padx=16, pady=8).pack(side="left", expand=True, fill="x", padx=6)

    # ── 開檔 ──────────────────────────────────────────────────
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="選擇報價單",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
        if not path:
            return
        self._src_path = path
        try:
            data = parse(path)
            self._parsed_data = data
            self._file_label.config(text=f"已載入：{path}", fg="#1a5276")
            h = data["header"]
            for key, var in self._read_vars.items():
                var.set(h.get(key, "") or "—")
            for row_id in self._tree.get_children():
                self._tree.delete(row_id)
            for item in data["items"]:
                self._tree.insert("", "end", values=(
                    item["seq"], item["name"].replace("\n", " "),
                    item["qty"], item["unit"],
                    item["unit_price"], item["subtotal"]))
        except Exception as e:
            messagebox.showerror("讀取失敗", f"無法解析報價單：\n{e}")

    # ── 雙擊編輯 ──────────────────────────────────────────────
    def _on_cell_dclick(self, event):
        item_id = self._tree.identify_row(event.y)
        col_id  = self._tree.identify_column(event.x)
        if not item_id or not col_id:
            return
        col_idx  = int(col_id.replace("#", "")) - 1
        col_keys = ("seq", "name", "qty", "unit", "unit_price", "subtotal")
        col_disp = ("序號", "品名", "數量", "單位", "單價", "小計")
        col_name = col_keys[col_idx]
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
        entry = tk.Entry(pop, textvariable=var, font=("Microsoft JhengHei", 11))
        entry.pack(fill="x", padx=10, pady=8)
        entry.select_range(0, "end")
        entry.focus()

        def save(_=None):
            vals = list(self._tree.item(item_id, "values"))
            new  = var.get()
            if col_name in ("seq", "qty", "unit_price", "subtotal"):
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

    # ── 品項操作 ──────────────────────────────────────────────
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

    # ── 製表人員 ──────────────────────────────────────────────
    def _add_operator(self):
        pop = tk.Toplevel(self)
        pop.title("新增製表人員")
        pop.geometry("260x80")
        pop.grab_set()
        var   = tk.StringVar()
        entry = tk.Entry(pop, textvariable=var, font=("Microsoft JhengHei", 11))
        entry.pack(fill="x", padx=10, pady=8)
        entry.focus()

        def save(_=None):
            name = var.get().strip()
            if not name:
                return
            if name not in self._config["operators"]:
                self._config["operators"].append(name)
                _save_config(self._config)
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
            _save_config(self._config)
            self._operator_cb["values"] = self._config["operators"]
            self._operator_var.set(self._config["operators"][0])

    # ── 生成出貨單 ────────────────────────────────────────────
    def _generate(self):
        if not self._parsed_data:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return

        items = []
        for i, rid in enumerate(self._tree.get_children()):
            v = self._tree.item(rid, "values")
            items.append({"seq": i+1, "name": v[1],
                          "qty": v[2], "unit": v[3],
                          "unit_price": v[4], "subtotal": v[5]})
        self._parsed_data["items"] = items

        extra = {
            "ship_date":      self._fill_vars["ship_date"].get(),
            "sale_no":        self._fill_vars["sale_no"].get(),
            "note":           self._fill_vars["note"].get(),
            "operator":       self._operator_var.get(),
            "invoice_choice": self._invoice_var.get(),
        }

        try:
            out_path = generate(self._parsed_data, extra)
            ans = messagebox.askyesno("生成成功",
                f"出貨單已儲存至：\n{out_path}\n\n是否立即開啟？")
            if ans:
                if sys.platform == "win32":
                    os.startfile(out_path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(out_path)])
                else:
                    subprocess.run(["xdg-open", str(out_path)])
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))

    # ── 生成驗機單 ────────────────────────────────────────────
    def _generate_inspection(self):
        if not self._parsed_data or not self._src_path:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return
        try:
            out_path = generate_inspection(self._src_path, self._parsed_data)
            ans = messagebox.askyesno("生成成功",
                f"驗機單已儲存至：\n{out_path}\n\n是否立即開啟？")
            if ans:
                if sys.platform == "win32":
                    os.startfile(out_path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(out_path)])
                else:
                    subprocess.run(["xdg-open", str(out_path)])
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))

    # ── 生成維修單 ────────────────────────────────────────────
    def _generate_fix(self):
        if not self._parsed_data:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return

        items = []
        for i, rid in enumerate(self._tree.get_children()):
            v = self._tree.item(rid, "values")
            items.append({"seq": i+1, "name": v[1],
                          "qty": v[2], "unit": v[3],
                          "unit_price": v[4], "subtotal": v[5]})
        self._parsed_data["items"] = items

        extra = {
            "ship_date":      self._fill_vars["ship_date"].get(),
            "sale_no":        self._fill_vars["sale_no"].get(),
            "note":           self._fill_vars["note"].get(),
            "operator":       self._operator_var.get(),
            "invoice_choice": self._invoice_var.get(),
        }

        try:
            out_path = generate_fix(self._parsed_data, extra)
            ans = messagebox.askyesno("生成成功",
                f"維修單已儲存至：\n{out_path}\n\n是否立即開啟？")
            if ans:
                if sys.platform == "win32":
                    os.startfile(out_path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(out_path)])
                else:
                    subprocess.run(["xdg-open", str(out_path)])
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))


if __name__ == "__main__":
    App().mainloop()