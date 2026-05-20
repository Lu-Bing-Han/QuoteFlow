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
from generator_schedule import generate_schedule, fetch_events, events_to_rows, calculate_travel_times, sort_rows_by_location
from syncer_trello import fetch_po_cards
from syncer_sheets import sync_cards
from syncer_production import sync_production, PRODUCTION_FILE as _PRODUCTION_EXCEL
from creator_trello import read_excel_cards, create_cards as trello_create_cards, get_sheet_names

from _paths import CONFIG_PATH, ICON_PATH, TEMPLATE_DIR, EXE_DIR

_GSHEETS_TOKEN_PATH      = EXE_DIR  / "gsheets_token.json"
_SYNCED_CARDS_PATH       = EXE_DIR  / "synced_cards.json"
_GSHEETS_CREDS_PATH      = TEMPLATE_DIR / "credentials.json"
_PRODUCTION_SYNCED_PATH  = EXE_DIR  / "production_synced_cards.json"

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
        FONT  = ("Microsoft JhengHei UI", 10)
        FONTB = ("Microsoft JhengHei UI", 10, "bold")
        BG    = "#f4f6f8"

        # ── Top bar ──────────────────────────────────────────
        top = tk.Frame(self, bg="#1a5276", pady=8)
        top.pack(fill="x")
        tk.Label(top, text="立善科技｜報價單轉單工具",
                 bg="#1a5276", fg="white",
                 font=("Microsoft JhengHei UI", 14, "bold")).pack(side="left", padx=16)
        tk.Button(top, text="選擇報價單 .xlsx ▶", command=self._open_file,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 10), padx=10, pady=3).pack(side="right", padx=16)

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
        tab_tag   = tk.Frame(nb, bg=BG)
        tab_label = tk.Frame(nb, bg=BG)
        tab_sched    = tk.Frame(nb, bg=BG)
        tab_overview = tk.Frame(nb, bg=BG)
        tab_prod     = tk.Frame(nb, bg=BG)
        tab_create   = tk.Frame(nb, bg=BG)

        nb.add(tab_ship,     text="  出貨單  ")
        nb.add(tab_insp,     text="  驗機單  ")
        nb.add(tab_fix,      text="  維修單  ")
        nb.add(tab_tag,      text="  維修掛件  ")
        nb.add(tab_label,    text="  標籤生成  ")
        nb.add(tab_sched,    text="  出貨排程  ")
        nb.add(tab_overview, text="  出貨一覽表  ")
        nb.add(tab_prod,     text="  生產群組紀錄  ")
        nb.add(tab_create,   text="  建立卡片  ")

        self._build_tab_shipping(tab_ship,    PAD, FONT, FONTB, BG)
        self._build_tab_inspection(tab_insp,  PAD, FONT, FONTB, BG)
        self._build_tab_fix(tab_fix,          PAD, FONT, FONTB, BG)
        self._build_tab_tag(tab_tag,          PAD, FONT, FONTB, BG)
        self._build_tab_label(tab_label,      FONT, FONTB, BG)
        self._build_tab_schedule(tab_sched,   FONT, FONTB, BG)
        self._build_tab_overview(tab_overview, FONT, FONTB, BG)
        self._build_tab_production(tab_prod,  FONT, FONTB, BG)
        self._build_tab_create_cards(tab_create, FONT, FONTB, BG)

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
        tk.Label(pf, text=r"Z:\Mika\驗收單及改造記錄單\Quoteflow_output",
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
        tk.Label(row, text=r"Z:\出貨單\Quoteflow_output",
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

    # ── Tab 5：標籤生成 ───────────────────────────────────────
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
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=8).pack(fill="x")

    # ── Tab 5：出貨排程 ───────────────────────────────────────
    def _build_tab_schedule(self, parent, FONT, FONTB, BG):
        from tkcalendar import DateEntry
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _rows = []  # list of {"ev": dict, "location": str, "note_suffix": str, "travel_time": str}

        # 從 template_schedule.xlsx「地址工作區」C2:C29 載入地址選項
        _addr_options: list[str] = []
        try:
            import openpyxl as _oxl
            _tmpl_path = TEMPLATE_DIR / "template_schedule.xlsx"
            if _tmpl_path.exists():
                _wb = _oxl.load_workbook(str(_tmpl_path), read_only=True, data_only=True)
                if "地址" in _wb.sheetnames:
                    _ws = _wb["地址"]
                    _addr_options = [
                        str(_ws.cell(row=r, column=3).value).strip()
                        for r in range(2, 30)
                        if _ws.cell(row=r, column=3).value
                    ]
                _wb.close()
        except Exception:
            pass

        # ── Credential section ────────────────────────────
        cred_frame = tk.LabelFrame(parent, text="Timetree 登入憑證", bg=BG, font=FONTB)
        cred_frame.pack(fill="x", padx=12, pady=(12, 4))
        cred_frame.columnconfigure(1, weight=1)

        tt_cfg = self._config.get("timetree", {})
        sid_var  = tk.StringVar(value=tt_cfg.get("session_id", ""))
        csrf_var = tk.StringVar(value=tt_cfg.get("csrf_token", ""))

        tk.Label(cred_frame, text="Session ID：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        sid_entry = tk.Entry(cred_frame, textvariable=sid_var, font=FONT_S, show="*")
        sid_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=3)

        tk.Label(cred_frame, text="CSRF Token：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=3)
        csrf_entry = tk.Entry(cred_frame, textvariable=csrf_var, font=FONT_S, show="*")
        csrf_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=3)

        def _show_hide(entry, btn):
            if entry.cget("show") == "*":
                entry.config(show="")
                btn.config(text="隱藏")
            else:
                entry.config(show="*")
                btn.config(text="顯示")

        btn_show_sid  = tk.Button(cred_frame, text="顯示", font=FONT_S,
                                  command=lambda: _show_hide(sid_entry,  btn_show_sid))
        btn_show_sid.grid(row=0, column=2, padx=(0, 8), pady=3)
        btn_show_csrf = tk.Button(cred_frame, text="顯示", font=FONT_S,
                                  command=lambda: _show_hide(csrf_entry, btn_show_csrf))
        btn_show_csrf.grid(row=1, column=2, padx=(0, 8), pady=3)

        def _save_creds():
            self._config.setdefault("timetree", {})
            self._config["timetree"]["session_id"] = sid_var.get().strip()
            self._config["timetree"]["csrf_token"]  = csrf_var.get().strip()
            _save_config(self._config)
            messagebox.showinfo("已儲存", "Timetree 憑證已儲存", parent=parent)

        tk.Button(cred_frame, text="儲存憑證", command=_save_creds,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT, padx=8).grid(row=2, column=1, sticky="w", padx=8, pady=(4, 6))

        # ── Google Maps 設定 ──────────────────────────────
        maps_frame = tk.LabelFrame(parent, text="Google Maps（行車時間）", bg=BG, font=FONTB)
        maps_frame.pack(fill="x", padx=12, pady=(0, 4))
        maps_frame.columnconfigure(1, weight=1)

        gm_cfg     = self._config.get("google_maps", {})
        gm_key_var = tk.StringVar(value=gm_cfg.get("api_key", ""))
        gm_org_var = tk.StringVar(value=gm_cfg.get("origin",  "406臺中市北屯區水景里景南巷1-1號"))

        tk.Label(maps_frame, text="API Key：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        gm_key_entry = tk.Entry(maps_frame, textvariable=gm_key_var, font=FONT_S, show="*")
        gm_key_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=3)
        btn_gm_key = tk.Button(maps_frame, text="顯示", font=FONT_S,
                               command=lambda: _show_hide(gm_key_entry, btn_gm_key))
        btn_gm_key.grid(row=0, column=2, padx=(0, 8), pady=3)

        tk.Label(maps_frame, text="出發地：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=3)
        tk.Entry(maps_frame, textvariable=gm_org_var, font=FONT_S
                 ).grid(row=1, column=1, columnspan=2, sticky="ew", padx=8, pady=3)

        def _save_maps_cfg():
            self._config.setdefault("google_maps", {})
            self._config["google_maps"]["api_key"] = gm_key_var.get().strip()
            self._config["google_maps"]["origin"]  = gm_org_var.get().strip()
            _save_config(self._config)
            messagebox.showinfo("已儲存", "Google Maps 設定已儲存", parent=parent)

        tk.Button(maps_frame, text="儲存設定", command=_save_maps_cfg,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT, padx=8).grid(row=2, column=1, sticky="w", padx=8, pady=(4, 6))

        # ── Preview section ───────────────────────────────
        prev_frame = tk.LabelFrame(parent, text="排程預覽", bg=BG, font=FONTB)
        prev_frame.pack(fill="both", expand=True, padx=12, pady=4)

        # Date + fetch row
        date_row = tk.Frame(prev_frame, bg=BG)
        date_row.pack(fill="x", padx=8, pady=(6, 4))
        tk.Label(date_row, text="日期：", bg=BG, font=FONT).pack(side="left")
        date_entry = DateEntry(date_row, font=FONT, date_pattern="yyyy/mm/dd",
                               background="#2e86c1", foreground="white", width=14)
        date_entry.pack(side="left", padx=(0, 6))

        fetch_status = tk.Label(date_row, text="", bg=BG, font=FONT_S, fg=GRAY)
        fetch_status.pack(side="left", padx=8)

        # Treeview
        tree_frame = tk.Frame(prev_frame, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 2))

        cols = ("seq", "location", "note", "travel_time")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                            height=7, selectmode="browse")
        tree.heading("seq",         text="順序")
        tree.heading("location",    text="地點")
        tree.heading("note",        text="備註")
        tree.heading("travel_time", text="行車時間")
        tree.column("seq",         width=45,  anchor="center", stretch=False)
        tree.column("location",    width=160, anchor="w",      stretch=False)
        tree.column("note",        width=200, anchor="w",      stretch=True)
        tree.column("travel_time", width=75,  anchor="center", stretch=False)
        tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)

        def _display_loc(location: str) -> str:
            idx = location.find("(")
            return location[:idx].strip() if idx != -1 else location

        def _refresh_tree():
            tree.delete(*tree.get_children())
            for i, row in enumerate(_rows, 1):
                tree.insert("", "end", values=(
                    i, _display_loc(row["location"]), row["note_suffix"], row.get("travel_time", "")))

        # Row action buttons
        btn_row = tk.Frame(prev_frame, bg=BG)
        btn_row.pack(fill="x", padx=8, pady=(2, 8))

        def _move(delta):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            new_idx = idx + delta
            if 0 <= new_idx < len(_rows):
                _rows[idx], _rows[new_idx] = _rows[new_idx], _rows[idx]
                _refresh_tree()
                tree.selection_set(tree.get_children()[new_idx])

        def _delete_row():
            sel = tree.selection()
            if not sel:
                return
            _rows.pop(tree.index(sel[0]))
            _refresh_tree()

        # display name → full address 對應表
        _addr_map = {}
        for _full in _addr_options:
            _idx = _full.find("(")
            _disp = _full[:_idx].strip() if _idx != -1 else _full
            _addr_map[_disp] = _full
        _addr_display = list(_addr_map.keys())

        def _open_row_dialog(title, location="", note_suffix="", on_confirm=None):
            dlg = tk.Toplevel(parent)
            dlg.title(title)
            dlg.resizable(False, False)
            dlg.grab_set()

            tk.Label(dlg, text="地點：", font=FONT).grid(row=0, column=0, padx=10, pady=6, sticky="w")
            # 顯示用：只顯示公司名稱
            loc_var = tk.StringVar(value=_display_loc(location))
            cb = ttk.Combobox(dlg, textvariable=loc_var, font=FONT, width=26, values=_addr_display)
            cb.grid(row=0, column=1, padx=10, pady=6)

            tk.Label(dlg, text="備註：", font=FONT).grid(row=1, column=0, padx=10, pady=6, sticky="w")
            note_var = tk.StringVar(value=note_suffix)
            tk.Entry(dlg, textvariable=note_var, font=FONT, width=28
                     ).grid(row=1, column=1, padx=10, pady=6)

            def _confirm():
                typed = loc_var.get().strip()
                # 從下拉選的 → 換回完整地址；手動輸入的 → 直接使用
                full_loc = _addr_map.get(typed, typed)
                if on_confirm:
                    on_confirm(full_loc, note_var.get().strip())
                dlg.destroy()

            tk.Button(dlg, text="確認", command=_confirm,
                      bg="#1a5276", fg="white", relief="flat",
                      font=FONT, padx=10).grid(row=2, column=1, sticky="e", padx=10, pady=8)

        def _edit_row(_event=None):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            row = _rows[idx]

            def _apply(loc, note):
                _rows[idx]["location"]    = loc
                _rows[idx]["note_suffix"] = note
                _refresh_tree()

            _open_row_dialog("編輯事件", row["location"], row["note_suffix"], _apply)

        def _add_row():
            def _apply(loc, note):
                if not loc:
                    return
                _rows.append({"ev": {}, "location": loc, "note_suffix": note, "travel_time": ""})
                _refresh_tree()
                tree.selection_set(tree.get_children()[-1])

            _open_row_dialog("新增事件", on_confirm=_apply)

        tree.bind("<Double-1>", _edit_row)

        def _calc_travel():
            if not _rows:
                messagebox.showwarning("無資料", "請先抓取事件清單", parent=parent)
                return
            api_key = gm_key_var.get().strip()
            origin  = gm_org_var.get().strip()
            if not api_key:
                messagebox.showwarning("未設定", "請先填入 Google Maps API Key", parent=parent)
                return
            fetch_status.config(text="計算行車時間中…", fg=GRAY)
            parent.update_idletasks()
            try:
                _, failed = calculate_travel_times(_rows, api_key, origin)
                _refresh_tree()
                if not failed:
                    fetch_status.config(text="行車時間計算完成", fg="#1e8449")
                else:
                    detail = "\n".join(
                        f"  第{seq}站「{loc}」— {status}" for seq, loc, status in failed
                    )
                    fetch_status.config(
                        text=f"完成，{len(failed)} 筆失敗（地址找不到）", fg="#e67e22")
                    messagebox.showwarning(
                        "部分地址無法計算",
                        f"以下站點無法取得行車時間：\n{detail}\n\n"
                        "請在 Timetree 的「地點」欄位填入完整地址，或手動在備註中修改。",
                        parent=parent)
            except Exception as e:
                fetch_status.config(text=f"✘ {e}", fg="#c0392b")

        def _sort_location(south_to_north: bool):
            if not _rows:
                messagebox.showwarning("無資料", "請先抓取事件清單", parent=parent)
                return
            api_key = gm_key_var.get().strip()
            if not api_key:
                messagebox.showwarning("未設定", "請先填入 Google Maps API Key", parent=parent)
                return
            fetch_status.config(text="Geocoding 中…", fg=GRAY)
            parent.update_idletasks()
            try:
                sorted_rows, failed = sort_rows_by_location(_rows, api_key, south_to_north)
                _rows.clear()
                _rows.extend(sorted_rows)
                _refresh_tree()
                if failed:
                    names = ", ".join(r["location"].split("(")[0] for r in failed)
                    fetch_status.config(text=f"排序完成，{len(failed)} 筆無法定位：{names}", fg="#e67e22")
                else:
                    fetch_status.config(text="排序完成", fg="#1e8449")
            except Exception as e:
                fetch_status.config(text=f"✘ {e}", fg="#c0392b")

        for text, cmd, color in [
            ("↑ 上移",       lambda: _move(-1),               "#5d6d7e"),
            ("↓ 下移",       lambda: _move(1),                "#5d6d7e"),
            ("➕ 新增",      _add_row,                         "#117a65"),
            ("✏ 編輯",       _edit_row,                        "#1a5276"),
            ("🗑 刪除",      _delete_row,                      "#922b21"),
            ("🧭 南→北",    lambda: _sort_location(True),     "#1a5276"),
            ("🧭 北→南",    lambda: _sort_location(False),    "#1a5276"),
            ("📍 計算時間",  _calc_travel,                     "#6c3483"),
        ]:
            tk.Button(btn_row, text=text, command=cmd,
                      bg=color, fg="white", relief="flat",
                      font=FONT_S, padx=8, pady=3).pack(side="left", padx=(0, 6))

        def _fetch_preview():
            sid  = sid_var.get().strip()
            csrf = csrf_var.get().strip()
            if not sid or not csrf:
                messagebox.showwarning("憑證未填", "請先填入 Session ID 與 CSRF Token", parent=parent)
                return
            target = date_entry.get_date()
            fetch_status.config(text="抓取中…", fg=GRAY)
            parent.update_idletasks()
            try:
                evs = fetch_events(target, sid, csrf)
                _rows.clear()
                _rows.extend(events_to_rows(evs))
                _refresh_tree()
                fetch_status.config(text=f"找到 {len(evs)} 筆事件", fg="#1e8449")
            except Exception as e:
                fetch_status.config(text=f"✘ {e}", fg="#c0392b")

        tk.Button(date_row, text="🔍 抓取", command=_fetch_preview,
                  bg="#117a65", fg="white", relief="flat",
                  font=FONT, padx=8).pack(side="left")

        # ── Write button ──────────────────────────────────
        out_label = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY,
                             anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(2, 0))

        def _write_schedule():
            if not _rows:
                messagebox.showwarning("無資料", "請先抓取並確認事件清單", parent=parent)
                return
            sid  = sid_var.get().strip()
            csrf = csrf_var.get().strip()
            target = date_entry.get_date()
            try:
                out = generate_schedule(target, sid, csrf, rows=list(_rows))
                out_label.config(text=f"✔  已寫入：{out}", fg="#1e8449")
                if messagebox.askyesno("寫入成功",
                        f"排程已寫入：\n{out}\n\n是否立即開啟？", parent=parent):
                    os.startfile(str(out))
            except Exception as e:
                out_label.config(text=f"✘  {e}", fg="#c0392b")
                messagebox.showerror("寫入失敗", str(e), parent=parent)

        bb = tk.Frame(parent, bg=BG, pady=8)
        bb.pack(fill="x", padx=12)
        tk.Button(bb, text="✅  確認寫入出貨行程表.xlsx", command=_write_schedule,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 12, "bold"), pady=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 7：出貨一覽表
    # ════════════════════════════════════════════════════════
    def _build_tab_overview(self, parent, FONT, FONTB, BG):
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
            _save_config(self._config)
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
        tk.Label(pf,
                 text=r"Z:\會計\●使用表格\公司帳務\1.帳務資料\▲生產群組紀錄(新版)\生產群組紀錄2026(115年).xlsx",
                 bg="#e8ecf0", font=FONT_S, fg=GRAY).pack(side="left", pady=6)

        # ── 狀態與同步按鈕 ────────────────────────────────
        out_label = tk.Label(parent, text="", bg=BG, font=FONT_S, fg=GRAY,
                             anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(8, 0))

        def _sync():
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

                added = sync_production(cards, _PRODUCTION_SYNCED_PATH)
                if added:
                    out_label.config(text=f"✔  同步完成，新增 {added} 筆資料", fg="#1e8449")
                else:
                    out_label.config(text="✔  同步完成，無新卡片（2026/5/15 之後）", fg="#1e8449")
                if messagebox.askyesno("同步完成",
                        f"{'新增 ' + str(added) + ' 筆資料' if added else '無新卡片'}\n\n是否立即開啟生產群組紀錄？",
                        parent=parent):
                    os.startfile(str(_PRODUCTION_EXCEL))
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

        _HEADERS = ["標題", "描述"]
        _EMPTY   = ["", ""]

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
            p = filedialog.askopenfilename(
                title="選擇 Excel 檔案",
                filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
            if not p:
                return
            path_var.set(p)
            try:
                names = get_sheet_names(Path(p))
                sheet_cb["values"] = names
                if names:
                    sheet_var.set(names[0])
            except Exception:
                sheet_cb["values"] = []
                sheet_var.set("")

        tk.Button(src_frame, text="選擇", command=_pick_file,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=6).grid(row=0, column=2, padx=(0, 4), pady=6)

        status_lbl = tk.Label(src_frame, text="", bg=BG, font=FONT_S, fg=GRAY)
        status_lbl.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        # ── 可編輯 Sheet 預覽 ─────────────────────────────
        prev_frame = tk.LabelFrame(parent, text="卡片預覽（雙擊儲存格可編輯，0 筆）",
                                   bg=BG, font=FONT)
        prev_frame.pack(fill="both", expand=True, padx=12, pady=4)

        sheet = Sheet(prev_frame,
                      headers=_HEADERS,
                      data=[_EMPTY[:] for _ in range(10)],
                      column_width=400,
                      row_height=28)
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

        def _load_preview():
            p = path_var.get().strip()
            if not p:
                messagebox.showwarning("未選擇檔案", "請先選擇 Excel 檔案", parent=parent)
                return
            status_lbl.config(text="讀取中…", fg=GRAY)
            parent.update_idletasks()
            try:
                selected_sheet = sheet_var.get() or None
                data = read_excel_cards(Path(p), sheet_name=selected_sheet)
                rows = [[c["title"], c["desc"]] for c in data]
                while len(rows) < len(data) + 5:
                    rows.append(_EMPTY[:])
                sheet.data = rows
                prev_frame.config(
                    text=f"卡片預覽（雙擊儲存格可編輯，{len(data)} 筆）")
                status_lbl.config(text=f"✔  讀取完成，共 {len(data)} 筆", fg="#1e8449")
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
            rows = sheet.data
            cards = []
            for row in rows:
                title = str(row[0]).strip() if len(row) > 0 else ""
                desc  = str(row[1]).strip() if len(row) > 1 else ""
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

    # ════════════════════════════════════════════════════════
    #  製表人員
    # ════════════════════════════════════════════════════════
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


if __name__ == "__main__":
    App().mainloop()
