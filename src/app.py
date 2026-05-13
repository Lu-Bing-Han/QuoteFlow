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
from generator_tag import generate_tag
from generator_label import generate_labels

from _paths import CONFIG_PATH, ICON_PATH

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
        self.title("報價單轉單工具｜立善科技")
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        win_h = min(sh - 60, 920)
        self.geometry(f"960x{win_h}+{(sw-960)//2}+0")
        self.resizable(True, True)
        self.configure(bg="#f4f6f8")
        if ICON_PATH.exists():
            self._icon = tk.PhotoImage(file=str(ICON_PATH))
            self.iconphoto(True, self._icon)
            _ico = ICON_PATH.with_suffix(".ico")
            if not _ico.exists():
                try:
                    from PIL import Image
                    Image.open(ICON_PATH).save(_ico, format="ICO")
                except Exception:
                    pass
            if _ico.exists():
                try:
                    self.iconbitmap(str(_ico))
                except Exception:
                    pass
        self._parsed_data = None
        self._src_path = None
        self._config = _load_config()
        self._build_ui()

    # ════════════════════════════════════════════════════════
    #  UI 建構
    # ════════════════════════════════════════════════════════
    def _build_ui(self):
        PAD   = {"padx": 12, "pady": 4}
        FONT  = ("Microsoft JhengHei", 10)
        FONTB = ("Microsoft JhengHei", 10, "bold")
        BG    = "#f4f6f8"

        # ── Top bar ──────────────────────────────────────────
        top = tk.Frame(self, bg="#1a5276", pady=8)
        top.pack(fill="x")
        tk.Label(top, text="立善科技｜報價單轉單工具",
                 bg="#1a5276", fg="white",
                 font=("Microsoft JhengHei", 14, "bold")).pack(side="left", padx=16)
        tk.Button(top, text="選擇報價單 .xlsx ▶", command=self._open_file,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=("Microsoft JhengHei", 10), padx=10, pady=3).pack(side="right", padx=16)

        self._file_label = tk.Label(self, text="⚠  尚未選擇報價單",
                                    bg=BG, fg="#c0392b", font=FONT)
        self._file_label.pack(anchor="w", padx=16, pady=(3, 0))

        # ── Notebook ─────────────────────────────────────────
        style = ttk.Style()
        style.configure("TNotebook", background="#dde1e6", tabmargins=[0, 4, 0, 0])
        style.configure("TNotebook.Tab", font=FONT, padding=[14, 6])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        tab_ship  = tk.Frame(nb, bg=BG)
        tab_insp  = tk.Frame(nb, bg=BG)
        tab_fix   = tk.Frame(nb, bg=BG)
        tab_label = tk.Frame(nb, bg=BG)

        nb.add(tab_ship,  text="  出貨單  ")
        nb.add(tab_insp,  text="  驗機單  ")
        nb.add(tab_fix,   text="  維修單  ")
        nb.add(tab_label, text="  標籤生成  ")

        self._build_tab_shipping(tab_ship,  PAD, FONT, FONTB, BG)
        self._build_tab_inspection(tab_insp, PAD, FONT, FONTB, BG)
        self._build_tab_fix(tab_fix,        PAD, FONT, FONTB, BG)
        self._build_tab_label(tab_label,    FONT, FONTB, BG)

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
                  font=("Microsoft JhengHei", 11, "bold"),
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
        FONT_S = ("Microsoft JhengHei", 9)

        info = tk.LabelFrame(parent, text="說明", bg=BG, font=FONTB)
        info.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(info, text="載入報價單後，點擊下方按鈕自動生成驗機單 Excel 及 Word。",
                 bg=BG, font=FONT, fg=GRAY).pack(padx=12, pady=8, anchor="w")

        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(fill="x", padx=12, pady=4)
        tk.Label(pf, text="輸出路徑：", bg="#e8ecf0", font=FONT_S, fg=GRAY,
                 anchor="w", width=12).pack(side="left", padx=8, pady=6)
        tk.Label(pf, text=r"Z:\Mika\驗收單及改造記錄單\Quoteflow_output",
                 bg="#e8ecf0", font=FONT_S, fg=GRAY).pack(side="left", pady=6)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(side="bottom", fill="x", padx=12)
        tk.Button(bb, text="🔍  生成驗機單", command=self._generate_inspection,
                  bg="#6c3483", fg="white", relief="flat",
                  font=("Microsoft JhengHei", 12, "bold"), pady=10).pack(fill="x")

    # ── Tab 3：維修單 ─────────────────────────────────────────
    def _build_tab_fix(self, parent, PAD, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei", 9)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(side="bottom", fill="x", padx=12)
        tk.Button(bb, text="🔧  生成維修單", command=self._generate_fix,
                  bg="#d68910", fg="white", relief="flat",
                  font=("Microsoft JhengHei", 12, "bold"), pady=10).pack(fill="x")

        pf = tk.Frame(parent, bg="#e8ecf0")
        pf.pack(side="bottom", fill="x", padx=12, pady=(0, 2))
        for label, path in [
            ("出貨單",   r"Z:\出貨單\Quoteflow_output"),
            ("維修掛件", r"Z:\待維修機台資料"),
        ]:
            row = tk.Frame(pf, bg="#e8ecf0")
            row.pack(fill="x")
            tk.Label(row, text=label + "：", bg="#e8ecf0", font=FONT_S,
                     fg=GRAY, anchor="w", width=10).pack(side="left", padx=8)
            tk.Label(row, text=path, bg="#e8ecf0", font=FONT_S,
                     fg=GRAY).pack(side="left")

        # 維修掛件
        tgf = tk.LabelFrame(parent, text="維修掛件", bg=BG, font=FONTB)
        tgf.pack(fill="x", padx=12, pady=(12, 4))
        tgf.columnconfigure(1, weight=1)
        tgf.columnconfigure(3, weight=1)

        self._tag_vars = {}

        no_var = tk.StringVar(value="1")
        self._tag_vars["no"] = no_var
        tk.Label(tgf, text="No.：", bg=BG, anchor="w", font=FONT
                 ).grid(row=0, column=0, sticky="w", padx=8, pady=2)
        ttk.Combobox(tgf, textvariable=no_var,
                     values=[str(i) for i in range(1, 21)],
                     width=8, font=FONT).grid(row=0, column=1, sticky="w", padx=8, pady=2)

        self._tag_partno_var = tk.StringVar()
        self._tag_vars["part_no"] = self._tag_partno_var
        tk.Label(tgf, text="品號：", bg=BG, anchor="w", font=FONT
                 ).grid(row=0, column=2, sticky="w", padx=8, pady=2)
        self._tag_partno_cb = ttk.Combobox(tgf, textvariable=self._tag_partno_var,
                                            font=FONT, width=20)
        self._tag_partno_cb.grid(row=0, column=3, sticky="ew", padx=8, pady=2)

        seq_var = tk.StringVar()
        self._tag_vars["seq_no"] = seq_var
        tk.Label(tgf, text="序號：", bg=BG, anchor="w", font=FONT
                 ).grid(row=1, column=0, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=seq_var, font=FONT
                 ).grid(row=1, column=1, sticky="ew", padx=8, pady=2)

        from tkcalendar import DateEntry
        tk.Label(tgf, text="拉回：", bg=BG, anchor="w", font=FONT
                 ).grid(row=1, column=2, sticky="w", padx=8, pady=2)
        self._tag_date_entry = DateEntry(
            tgf, font=FONT, date_pattern="yyyy/mm/dd",
            background="#2e86c1", foreground="white", width=14)
        self._tag_date_entry.grid(row=1, column=3, sticky="w", padx=8, pady=2)

        prob_var = tk.StringVar()
        self._tag_vars["problem"] = prob_var
        tk.Label(tgf, text="問題：", bg=BG, anchor="w", font=FONT
                 ).grid(row=2, column=0, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=prob_var, font=FONT
                 ).grid(row=2, column=1, sticky="ew", padx=8, pady=2)

        status_var = tk.StringVar()
        self._tag_vars["repair_status"] = status_var
        tk.Label(tgf, text="維修狀況：", bg=BG, anchor="w", font=FONT
                 ).grid(row=2, column=2, sticky="w", padx=8, pady=2)
        tk.Entry(tgf, textvariable=status_var, font=FONT
                 ).grid(row=2, column=3, sticky="ew", padx=8, pady=2)

    # ── Tab 4：標籤生成 ───────────────────────────────────────
    def _build_tab_label(self, parent, FONT, FONTB, BG):
        from tksheet import Sheet
        import re as _re

        # 模板選擇
        tpl_frame = tk.Frame(parent, bg=BG)
        tpl_frame.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(tpl_frame, text="標籤樣式：", bg=BG, font=FONTB).pack(side="left")
        _tpl_var = tk.StringVar(value="銀標")
        _tpl_cb  = ttk.Combobox(tpl_frame, textvariable=_tpl_var,
                                 values=["銀標", "APT標", "無公司標", "上銀標"],
                                 state="readonly", font=FONT, width=10)
        _tpl_cb.pack(side="left", padx=(8, 0))

        # 欄位索引：0=型號 1=荷重 2=序號 3=機台尺寸 4=機台重量 5=出廠年份
        #           6=供應商代碼 7=機台序號 8=訂單編號 9=收貨人
        _HEADERS = ["型號", "荷重", "序號", "機台尺寸", "機台重量", "出廠年份",
                    "供應商代碼", "機台序號", "訂單編號", "收貨人"]
        _EMPTY   = [""] * len(_HEADERS)
        _COLS    = {
            "銀標":    [0, 1, 2],
            "APT標":   [0, 1, 2, 3, 4, 5],
            "無公司標": [0, 1, 2],
            "上銀標":   [6, 7, 8, 9],
        }
        _ALL_COLS = list(range(len(_HEADERS)))

        tf = tk.LabelFrame(parent, text="標籤資料", bg=BG, font=FONTB)
        tf.pack(fill="both", expand=True, padx=12, pady=(10, 4))

        sheet = Sheet(tf,
                      headers=_HEADERS,
                      data=[_EMPTY[:] for _ in range(50)],
                      column_width=130,
                      row_height=28)
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

        def _on_tpl_change(*_):
            visible = _COLS.get(_tpl_var.get(), [0, 1, 2])
            hidden  = [c for c in _ALL_COLS if c not in visible]
            sheet.show_columns(_ALL_COLS)
            if hidden:
                sheet.hide_columns(hidden)

        _tpl_var.trace_add("write", _on_tpl_change)
        _on_tpl_change()

        def _re_find(text, *patterns):
            for pat in patterns:
                m = _re.search(pat, text)
                if m:
                    return m.group(1).strip()
            return ""

        def _load_from_quote():
            today  = datetime.today()
            serial = f"{today.year % 100 + 12:02d}{today.month + 12:02d}"
            year   = str(today.year)
            rows   = []
            if self._parsed_data:
                for item in self._parsed_data.get("items", []):
                    name = item.get("name", "")
                    raw_load  = _re_find(name, r'載重[：:]\s*(\S+)')
                    load_num  = _re.sub(r'[kK][gG][sS]?$', '', raw_load).strip()
                    load      = (load_num + "kgs") if load_num else ""
                    length    = _re_find(name, r'牙叉長度\s*[：: ]+(\d+(?:\.\d+)?)')
                    width     = _re_find(name, r'牙叉外寬\s*[：: ]+(\d+(?:\.\d+)?)')
                    size      = (f"{length}mm*{width}mm" if length and width
                                 else (f"{length}mm" if length else f"{width}mm" if width else ""))
                    weight_raw = _re_find(name, r'自重\s*[：:]\s*(\d+(?:\.\d+)?)')
                    weight     = (weight_raw + "kgs") if weight_raw else ""
                    rows.append([item.get("part_no", ""), load, serial,
                                 size, weight, year])
            while len(rows) < 50:
                rows.append(_EMPTY[:])
            sheet.data = rows

        _load_from_quote()

        def _autofill_serial():
            data = sheet.data
            start_row, start_val = None, ""
            for i, row in enumerate(data):
                v = str(row[2]).strip() if len(row) > 2 else ""
                if v:
                    start_row, start_val = i, v
                    break
            if start_row is None:
                messagebox.showwarning("無起始序號", "請先在製造序號欄填入第一個序號", parent=parent)
                return
            m = _re.match(r'^(.*?)(\d+)$', start_val)
            if not m:
                messagebox.showwarning("格式不符", "序號結尾需為數字，例如 SN001 或 2026001", parent=parent)
                return
            prefix, num_str = m.group(1), m.group(2)
            w = len(num_str)
            counter = int(num_str)
            for i in range(start_row, len(data)):
                has_content = str(data[i][0]).strip() or str(data[i][1]).strip()
                if i == start_row or has_content:
                    data[i][2] = prefix + str(counter).zfill(w)
                    counter += 1
            sheet.data = data

        # 操作按鈕
        bb = tk.Frame(parent, bg=BG)
        bb.pack(fill="x", padx=12, pady=(0, 4))
        tk.Button(bb, text="從報價單讀入", command=_load_from_quote,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT, padx=8).pack(side="left", padx=(0, 6))
        tk.Button(bb, text="＋ 新增列",
                  command=lambda: sheet.insert_rows(number=1),
                  bg="#27ae60", fg="white", relief="flat",
                  font=FONT, padx=8).pack(side="left", padx=(0, 6))
        tk.Button(bb, text="－ 刪除列",
                  command=lambda: [sheet.delete_rows(row=r)
                                   for r in sorted(sheet.get_selected_rows(), reverse=True)],
                  bg="#c0392b", fg="white", relief="flat",
                  font=FONT, padx=8).pack(side="left", padx=(0, 6))
        tk.Button(bb, text="流水號↓", command=_autofill_serial,
                  bg="#7d3c98", fg="white", relief="flat",
                  font=FONT, padx=8).pack(side="left")

        # 生成按鈕
        def _generate():
            rows = sheet.data
            def _kgs(v):
                v = str(v).strip()
                return (v + "kgs") if v and not v.lower().endswith("kgs") else v
            def _s(v): return str(v).strip()
            def _g(r, i): return _s(r[i]) if len(r) > i else ""
            tpl = _tpl_var.get()
            is_silver_top = (tpl == "上銀標")
            data_list = [
                {"型號": _s(r[0]), "荷重": _kgs(r[1]),
                 "序號": _g(r, 7) if is_silver_top else _g(r, 2),
                 "製造序號": _g(r, 2),
                 "機台尺寸": _g(r, 3), "機台重量": _kgs(_g(r, 4)),
                 "出廠年份": _g(r, 5),
                 "供應商代碼": _g(r, 6), "機台序號": _g(r, 7),
                 "訂單編號": _g(r, 8), "收貨人": _g(r, 9)}
                for r in rows if any(str(r[i]).strip() for i in _COLS.get(tpl, [0, 1, 2]) if i < len(r))
            ]
            if not data_list:
                messagebox.showwarning("無資料", "請先填入標籤資料", parent=parent)
                return
            date_tag = datetime.today().strftime("%Y%m%d%H%M%S")
            out_path = Path(r"Z:\出貨單\Quoteflow_output") / f"標籤-{date_tag}.pdf"
            try:
                result = generate_labels(data_list, out_path, template_key=tpl)
                if messagebox.askyesno("生成成功",
                        f"已生成 {len(data_list)} 張標籤：\n{result}\n\n是否立即開啟？",
                        parent=parent):
                    os.startfile(str(result))
            except Exception as e:
                messagebox.showerror("生成失敗", str(e), parent=parent)

        gf = tk.Frame(parent, bg=BG, pady=8)
        gf.pack(fill="x", padx=12)
        tk.Button(gf, text="🖨  生成標籤 PDF", command=_generate,
                  bg="#1e8449", fg="white", relief="flat",
                  font=("Microsoft JhengHei", 12, "bold"), pady=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  開檔
    # ════════════════════════════════════════════════════════
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
            self._file_label.config(text=f"✔  已載入：{path}", fg="#1e8449")
            h = data["header"]
            for key, var in self._read_vars.items():
                var.set(h.get(key, "") or "—")
            for row_id in self._tree.get_children():
                self._tree.delete(row_id)
            for item in data["items"]:
                self._tree.insert("", "end", values=(
                    item["seq"], item.get("part_no", ""),
                    item["qty"], item["unit"], "", ""))
            self._fill_vars["sale_no"].set(h.get("quote_no", ""))
            part_nos = [item.get("part_no", "") or item.get("name", "")
                        for item in data["items"]]
            self._tag_partno_cb["values"] = part_nos
            if part_nos:
                self._tag_partno_var.set(part_nos[0])
        except Exception as e:
            messagebox.showerror("讀取失敗", f"無法解析報價單：\n{e}")

    # ════════════════════════════════════════════════════════
    #  品項表操作
    # ════════════════════════════════════════════════════════
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
        entry = tk.Entry(pop, textvariable=var, font=("Microsoft JhengHei", 11))
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

    # ════════════════════════════════════════════════════════
    #  製表人員
    # ════════════════════════════════════════════════════════
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

    # ════════════════════════════════════════════════════════
    #  生成
    # ════════════════════════════════════════════════════════
    _OUT_SHIPPING = Path(r"Z:\出貨單\Quoteflow_output")
    _OUT_TAG      = Path(r"Z:\待維修機台資料")

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
        if not self._parsed_data or not self._src_path:
            messagebox.showwarning("尚未載入", "請先選擇並載入報價單")
            return
        try:
            excel_path, word_paths = generate_inspection(self._src_path, self._parsed_data)
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
        tag_fields = [
            self._tag_vars["seq_no"].get().strip(),
            self._tag_vars["problem"].get().strip(),
            self._tag_vars["repair_status"].get().strip(),
        ]
        gen_tag = True
        if not any(tag_fields):
            gen_tag = messagebox.askyesno(
                "維修掛件未填寫",
                "維修掛件欄位尚未填寫，是否仍要繼續生成維修單（不含掛件）？")
            if not gen_tag:
                return
        try:
            result = generate_fix(self._parsed_data, extra, output_dir=self._OUT_SHIPPING)
            paths  = result if isinstance(result, list) else [result]
            if gen_tag and any(tag_fields):
                tag_data = {
                    "no":            self._tag_vars["no"].get(),
                    "part_no":       self._tag_vars["part_no"].get(),
                    "seq_no":        self._tag_vars["seq_no"].get(),
                    "problem":       self._tag_vars["problem"].get(),
                    "pullback_date": self._tag_date_entry.get_date().strftime("%Y/%m/%d"),
                    "repair_status": self._tag_vars["repair_status"].get(),
                }
                paths.append(generate_tag(self._parsed_data, tag_data,
                                          output_dir=self._OUT_TAG))
            msg = "\n".join(str(p) for p in paths)
            if messagebox.askyesno("生成成功",
                    f"已生成 {len(paths)} 份檔案：\n{msg}\n\n是否立即開啟？"):
                for p in paths:
                    os.startfile(p) if sys.platform == "win32" else subprocess.run(["open", str(p)])
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))


if __name__ == "__main__":
    App().mainloop()
