"""
app_core.py — 主應用程式類別 AppCore / App
整合所有 mixin，包含 __init__、_build_ui、共用方法與導覽邏輯。
"""

import json, os, sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# ── 路徑 ──────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent.parent))   # src/
from _paths import CONFIG_PATH, ICON_PATH, TEMPLATE_DIR, EXE_DIR

# ── Mixin imports ─────────────────────────────────────────
from ui.mixin_documents import _DocumentsTab
from ui.mixin_quote     import _QuoteTab
from ui.mixin_label     import _LabelTab
from ui.mixin_schedule  import _ScheduleTab
from ui.mixin_trello    import _TrelloTab
from ui.mixin_history   import _HistoryTab

# ── Core imports ──────────────────────────────────────────
from core.parser import parse
from core.db import init_db, set_db_path


# ════════════════════════════════════════════════════════
#  Config helpers
# ════════════════════════════════════════════════════════
def _load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"operators": ["小皋"]}


def _save_config(cfg):
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ════════════════════════════════════════════════════════
#  App class
# ════════════════════════════════════════════════════════
class App(
    _DocumentsTab,
    _QuoteTab,
    _LabelTab,
    _ScheduleTab,
    _TrelloTab,
    _HistoryTab,
    tk.Tk,
):
    """Main application window — inherits all mixin tab builders."""

    _PATH_DEFAULTS = {
        "output_shipping":   r"Z:\出貨單\Quoteflow_output",
        "output_inspection": r"Z:\Mika\驗收單及改造記錄單\Quoteflow_output",
        "output_tag":        r"Z:\待維修機台資料",
        "output_label":      r"Z:\出貨單\Quoteflow_output",
        "output_quote":      r"Z:\出貨單\Quoteflow_output\報價單",
        "schedule_file":     r"Z:\會計\5.出貨相關\出貨行程表.xlsx",
        "production_file":   r"Z:\會計\●使用表格\公司帳務\1.帳務資料\▲生產群組紀錄(新版)\生產群組紀錄2026(115年).xlsx",
        "download_cards_dir": r"Z:\出貨單\Quoteflow_output\下載卡片",
    }

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
        set_db_path(self._config.get("db_path"))   # 若有設定共用路徑則套用
        init_db()       # 建立資料庫 / 資料表（已存在則略過）
        self._build_ui()

    # ════════════════════════════════════════════════════════
    #  Config helpers exposed as instance methods so mixins can call self._save_config(...)
    # ════════════════════════════════════════════════════════
    def _load_config(self):
        return _load_config()

    def _save_config(self, cfg):
        _save_config(cfg)

    # ════════════════════════════════════════════════════════
    #  Path helpers
    # ════════════════════════════════════════════════════════
    def _get_path(self, key: str) -> Path:
        return Path(self._config.get("paths", {}).get(key) or self._PATH_DEFAULTS[key])

    @property
    def _OUT_SHIPPING(self):   return self._get_path("output_shipping")
    @property
    def _OUT_TAG(self):        return self._get_path("output_tag")

    # ════════════════════════════════════════════════════════
    #  UI 建構
    # ════════════════════════════════════════════════════════
    def _build_ui(self):
        PAD   = {"padx": 12, "pady": 4}
        FONT  = ("Microsoft JhengHei UI", 10)
        FONTB = ("Microsoft JhengHei UI", 10, "bold")
        BG    = "#f4f6f8"
        NAV_BG      = "#1b2631"
        NAV_ACTIVE  = "#1a5276"
        NAV_HOVER   = "#2e4057"
        NAV_FG      = "#d5d8dc"
        NAV_GRP_FG  = "#7f8c8d"

        # ── Top bar ──────────────────────────────────────────
        top = tk.Frame(self, bg="#1a5276", pady=8)
        top.pack(fill="x")
        tk.Label(top, text="立善科技｜QuoteFlow",
                 bg="#1a5276", fg="white",
                 font=("Microsoft JhengHei UI", 14, "bold")).pack(side="left", padx=16)
        tk.Button(top, text="選擇報價單 .xlsx ▶", command=self._open_file,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 10), padx=10, pady=3).pack(side="right", padx=(4, 16))
        tk.Button(top, text="⚙", command=self._open_paths_dialog,
                  bg="#1a5276", fg="white", relief="flat",
                  font=("Microsoft JhengHei UI", 13), padx=6, pady=1).pack(side="right")

        self._file_label = tk.Label(self, text="⚠  尚未選擇報價單",
                                    bg=BG, fg="#c0392b", font=FONT)
        self._file_label.pack(anchor="w", padx=16, pady=(3, 0))

        # ── 主體：側邊列 + 內容區 ────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)

        # ── 左側導覽列 ────────────────────────────────────────
        nav = tk.Frame(body, bg=NAV_BG, width=130)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)

        content_area = tk.Frame(body, bg=BG)
        content_area.pack(side="left", fill="both", expand=True)

        # ── 頁面 Frame（每個功能一個，重疊在 content_area）──────
        self._pages: dict[str, tk.Frame] = {}
        self._nav_btns: dict[str, tk.Label] = {}
        self._active_page: str = ""

        def _show(key: str):
            if self._active_page and self._active_page in self._nav_btns:
                self._nav_btns[self._active_page].config(bg=NAV_BG)
            for f in self._pages.values():
                f.pack_forget()
            self._pages[key].pack(fill="both", expand=True)
            self._nav_btns[key].config(bg=NAV_ACTIVE)
            self._active_page = key

        self._show_page = _show

        def _make_page(key: str) -> tk.Frame:
            f = tk.Frame(content_area, bg=BG)
            self._pages[key] = f
            return f

        # ── 導覽列項目生成 ────────────────────────────────────
        def _nav_group(text: str):
            tk.Label(nav, text=text, bg=NAV_BG, fg=NAV_GRP_FG,
                     font=("Microsoft JhengHei UI", 8),
                     anchor="w", padx=10, pady=0).pack(fill="x", pady=(10, 1))
            tk.Frame(nav, bg=NAV_HOVER, height=1).pack(fill="x", padx=8)

        def _nav_item(text: str, key: str):
            lbl = tk.Label(nav, text=f"  {text}", bg=NAV_BG, fg=NAV_FG,
                           font=("Microsoft JhengHei UI", 9),
                           anchor="w", padx=6, pady=7, cursor="hand2")
            lbl.pack(fill="x")
            lbl.bind("<Button-1>", lambda e, k=key: _show(k))
            lbl.bind("<Enter>",    lambda e, b=lbl, k=key:
                         b.config(bg=NAV_HOVER) if self._active_page != k else None)
            lbl.bind("<Leave>",    lambda e, b=lbl, k=key:
                         b.config(bg=NAV_ACTIVE if self._active_page == k else NAV_BG))
            self._nav_btns[key] = lbl

        # ── 導覽結構 ──────────────────────────────────────────
        _nav_group("📄  單據生成")
        _nav_item("出貨單",   "shipping")
        _nav_item("報價單",   "quote")
        _nav_item("驗機單",   "inspection")
        _nav_item("維修單",   "fix")
        _nav_item("維修掛件", "tag")
        _nav_item("標籤生成", "label")

        _nav_group("📅  排程")
        _nav_item("出貨排程", "schedule")

        _nav_group("🃏  Trello")
        _nav_item("出貨一覽表",   "overview")
        _nav_item("生產群組紀錄", "production")
        _nav_item("建立卡片",     "create")
        _nav_item("下載卡片",     "download")

        _nav_group("🗂  記錄查詢")
        _nav_item("報價記錄",     "history")

        # ── 建立各頁面內容 ────────────────────────────────────
        self._build_tab_shipping(   _make_page("shipping"),   PAD, FONT, FONTB, BG)
        self._build_tab_quote(      _make_page("quote"),      FONT, FONTB, BG)
        self._build_tab_inspection( _make_page("inspection"), PAD, FONT, FONTB, BG)
        self._build_tab_fix(        _make_page("fix"),        PAD, FONT, FONTB, BG)
        self._build_tab_tag(        _make_page("tag"),        PAD, FONT, FONTB, BG)
        self._build_tab_label(      _make_page("label"),      FONT, FONTB, BG)
        self._build_tab_schedule(   _make_page("schedule"),   FONT, FONTB, BG)
        self._build_tab_overview(   _make_page("overview"),   FONT, FONTB, BG)
        self._build_tab_production( _make_page("production"), FONT, FONTB, BG)
        self._build_tab_create_cards( _make_page("create"),   FONT, FONTB, BG)
        self._build_tab_download_cards(_make_page("download"),FONT, FONTB, BG)
        self._build_tab_history(    _make_page("history"),    FONT, FONTB, BG)

        # 預設顯示出貨單
        _show("shipping")

    # ════════════════════════════════════════════════════════
    #  路徑設定 Dialog
    # ════════════════════════════════════════════════════════
    def _open_paths_dialog(self):
        BG     = "#f4f6f8"
        GRAY   = "#5d6d7e"
        FONT   = ("Microsoft JhengHei UI", 10)
        FONTB  = ("Microsoft JhengHei UI", 10, "bold")
        FONT_S = ("Microsoft JhengHei UI", 9)
        PAD    = {"padx": 8, "pady": 5}

        dlg = tk.Toplevel(self)
        dlg.title("⚙  路徑與人員設定")
        dlg.configure(bg=BG)
        dlg.resizable(True, False)
        dlg.grab_set()
        dlg.update_idletasks()
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"720x620+{(sw-720)//2}+{(sh-620)//2}")

        # ── 資料庫路徑設定 ────────────────────────────────────
        db_lf = tk.LabelFrame(dlg, text="資料庫設定", bg=BG, font=FONTB)
        db_lf.pack(fill="x", padx=16, pady=(16, 6))
        db_lf.columnconfigure(1, weight=1)

        tk.Label(db_lf, text="資料庫路徑：", bg=BG, font=FONT_S, fg=GRAY,
                 anchor="w").grid(row=0, column=0, sticky="w", padx=(8, 2), pady=5)
        db_path_var = tk.StringVar(value=self._config.get("db_path", ""))
        tk.Entry(db_lf, textvariable=db_path_var, font=FONT_S
                 ).grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=5)

        def _pick_db():
            p = filedialog.askopenfilename(
                title="選擇資料庫檔案",
                filetypes=[("SQLite 資料庫", "*.db"), ("所有檔案", "*.*")],
                parent=dlg)
            if p:
                db_path_var.set(p)
                # 選完立即寫入 config，不需要另外按「儲存並關閉」
                self._config["db_path"] = p
                set_db_path(p)
                self._save_config(self._config)

        tk.Button(db_lf, text="選擇", command=_pick_db,
                  bg="#2e86c1", fg="white", relief="flat",
                  font=FONT_S, padx=6).grid(row=0, column=2, padx=(0, 8), pady=5)
        tk.Label(db_lf,
                 text="留空 = 使用本機預設路徑；填入 NAS 路徑可多台電腦共用同一份資料庫",
                 bg=BG, font=("Microsoft JhengHei UI", 8), fg=GRAY
                 ).grid(row=1, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        items = [
            ("output_shipping",   "出貨單 / 維修單 輸出資料夾", False),
            ("output_inspection", "驗機單 輸出資料夾",          False),
            ("output_tag",        "維修掛件 輸出資料夾",        False),
            ("output_label",      "標籤 輸出資料夾",            False),
            ("output_quote",      "報價單 輸出資料夾",          False),
            ("download_cards_dir","下載卡片 輸出資料夾",        False),
            ("schedule_file",     "出貨行程表 .xlsx",           True),
            ("production_file",   "生產群組紀錄 .xlsx",         True),
        ]

        lf = tk.LabelFrame(dlg, text="輸出路徑設定", bg=BG, font=FONTB)
        lf.pack(fill="x", padx=16, pady=(16, 6))
        lf.columnconfigure(1, weight=1)

        path_vars: dict[str, tk.StringVar] = {}
        paths_cfg = self._config.get("paths", {})

        for i, (key, label, is_file) in enumerate(items):
            tk.Label(lf, text=label + "：", bg=BG, font=FONT_S, fg=GRAY,
                     anchor="w").grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value=paths_cfg.get(key) or self._PATH_DEFAULTS[key])
            path_vars[key] = var
            tk.Entry(lf, textvariable=var, font=FONT_S
                     ).grid(row=i, column=1, sticky="ew", padx=(0, 4), pady=5)

            def _pick(v=var, f=is_file):
                if f:
                    p = filedialog.askopenfilename(
                        filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
                else:
                    p = filedialog.askdirectory()
                if p:
                    v.set(p)

            tk.Button(lf, text="選擇", command=_pick,
                      bg="#2e86c1", fg="white", relief="flat",
                      font=FONT_S, padx=6).grid(row=i, column=2, padx=(0, 8), pady=5)

        # ── 人員設定 ─────────────────────────────────────────
        op_lf = tk.LabelFrame(dlg, text="人員設定（報價單製表人員與單號代號）",
                              bg=BG, font=FONTB)
        op_lf.pack(fill="x", padx=16, pady=(0, 6))

        # Treeview 顯示現有人員
        op_cols = ("name", "code")
        op_tree = ttk.Treeview(op_lf, columns=op_cols, show="headings",
                               selectmode="browse", height=4)
        op_tree.heading("name", text="名稱（放入 Excel 製表人欄）")
        op_tree.heading("code", text="代號（報價單號前綴）")
        op_tree.column("name", width=200, anchor="w")
        op_tree.column("code", width=120, anchor="center")
        op_tree.pack(fill="x", padx=8, pady=(6, 2))

        # 填入現有資料
        operators = self._config.get("operators", ["小皋"])
        op_codes  = self._config.get("operator_codes", {})
        for name in operators:
            op_tree.insert("", "end", values=(name, op_codes.get(name, "")))

        # 新增 / 刪除 列
        edit_row = tk.Frame(op_lf, bg=BG)
        edit_row.pack(fill="x", padx=8, pady=(0, 6))

        tk.Label(edit_row, text="名稱：", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        name_var = tk.StringVar()
        tk.Entry(edit_row, textvariable=name_var, font=FONT_S, width=12,
                 relief="solid", borderwidth=1).pack(side="left", padx=(0, 8))

        tk.Label(edit_row, text="代號：", bg=BG, font=FONT_S, fg=GRAY).pack(side="left")
        code_var = tk.StringVar()
        tk.Entry(edit_row, textvariable=code_var, font=FONT_S, width=8,
                 relief="solid", borderwidth=1).pack(side="left", padx=(0, 8))

        tk.Label(edit_row, text="（英文字串，如 K）", bg=BG, font=FONT_S, fg=GRAY
                 ).pack(side="left", padx=(0, 12))

        def _op_add():
            n = name_var.get().strip()
            c = code_var.get().strip().upper()
            if not n or not c:
                messagebox.showwarning("欄位不完整", "請填入名稱和代號", parent=dlg)
                return
            # 更新已存在的同名項
            for item in op_tree.get_children():
                if op_tree.item(item)["values"][0] == n:
                    op_tree.item(item, values=(n, c))
                    name_var.set(""); code_var.set("")
                    return
            op_tree.insert("", "end", values=(n, c))
            name_var.set(""); code_var.set("")

        def _op_del():
            sel = op_tree.selection()
            if sel:
                op_tree.delete(sel[0])

        def _op_select(*_):
            sel = op_tree.selection()
            if sel:
                vals = op_tree.item(sel[0])["values"]
                name_var.set(vals[0])
                code_var.set(vals[1])

        op_tree.bind("<<TreeviewSelect>>", _op_select)

        tk.Button(edit_row, text="新增 / 更新", command=_op_add,
                  bg="#117a65", fg="white", relief="flat",
                  font=FONT_S, padx=8).pack(side="left", padx=(0, 4))
        tk.Button(edit_row, text="刪除選取", command=_op_del,
                  bg="#c0392b", fg="white", relief="flat",
                  font=FONT_S, padx=8).pack(side="left")

        def _reset_defaults():
            for key, var in path_vars.items():
                var.set(self._PATH_DEFAULTS[key])

        def _save():
            self._config.setdefault("paths", {})
            for key, var in path_vars.items():
                self._config["paths"][key] = var.get().strip()

            # 收集人員資料
            new_operators: list[str] = []
            new_codes:     dict[str, str] = {}
            for item in op_tree.get_children():
                vals = op_tree.item(item)["values"]
                name = str(vals[0]).strip()
                code = str(vals[1]).strip().upper()
                if name:
                    new_operators.append(name)
                    new_codes[name] = code or name[:1].upper()
            if new_operators:
                self._config["operators"]      = new_operators
                self._config["operator_codes"] = new_codes

            db_path = db_path_var.get().strip()
            if db_path:
                self._config["db_path"] = db_path
            else:
                self._config.pop("db_path", None)
            set_db_path(db_path or None)

            self._save_config(self._config)
            messagebox.showinfo("已儲存", "設定已儲存", parent=dlg)
            dlg.destroy()

        bb = tk.Frame(dlg, bg=BG)
        bb.pack(fill="x", padx=16, pady=8)
        tk.Button(bb, text="還原路徑預設值", command=_reset_defaults,
                  bg="#5d6d7e", fg="white", relief="flat",
                  font=FONT, padx=10).pack(side="left", padx=(0, 8))
        tk.Button(bb, text="儲存並關閉", command=_save,
                  bg="#1a5276", fg="white", relief="flat",
                  font=FONTB, padx=16, pady=6).pack(side="left")

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
