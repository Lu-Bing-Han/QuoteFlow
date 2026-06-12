"""
app_core.py — 主應用程式類別 AppCore / App
整合所有 mixin，包含 __init__、_build_ui、共用方法與導覽邏輯。
"""

import json, os, sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from pathlib import Path

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# ── CTk 共用輔助 ──────────────────────────────────────────
def _mk_lf(parent, text: str, fg_color: str = "#f4f6f8",
           font=("Microsoft JhengHei UI", 10, "bold"),
           text_color: str = "#5d6d7e"):
    """LabelFrame 替代：回傳 (outer, inner)。
    outer.pack(...)；將子 widget 加到 inner。
    """
    outer = ctk.CTkFrame(parent, fg_color=fg_color, corner_radius=8,
                          border_width=1, border_color="#d0d7de")
    ctk.CTkLabel(outer, text=f"  {text}  ", fg_color=fg_color,
                  text_color=text_color, font=font).pack(anchor="w", padx=10, pady=(6, 0))
    inner = ctk.CTkFrame(outer, fg_color=fg_color, corner_radius=0)
    inner.pack(fill="both", expand=True, padx=4, pady=(0, 6))
    return outer, inner


def _ctk_btn(parent, text, command=None, fg_color="#1a5276",
             hover_color="#154360", text_color="white",
             font=("Microsoft JhengHei UI", 9),
             width=0, height=32, padx=0, pady=0,
             corner_radius=6, **kw) -> ctk.CTkButton:
    w = ctk.CTkButton(parent, text=text, command=command,
                       fg_color=fg_color, hover_color=hover_color,
                       text_color=text_color, font=font,
                       corner_radius=corner_radius,
                       height=height, **kw)
    if width:
        w.configure(width=width)
    return w


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
from ui.mixin_line      import _LineTab

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
            from core.logger import get_logger
            get_logger(__name__).warning("config.json 格式錯誤，使用預設值", exc_info=True)
    return {"operators": ["小皋"]}


def _save_config(cfg):
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def _validate_config(cfg: dict):
    """啟動時驗證 config，有問題用 messagebox 警告（不阻止啟動）。"""
    from core.logger import get_logger
    log = get_logger(__name__)
    warnings: list[str] = []

    db_path = cfg.get("db_path", "")
    if db_path and not Path(db_path).exists():
        msg = f"設定的資料庫路徑不存在：\n{db_path}\n\n將使用本機預設路徑。"
        warnings.append(msg)
        log.warning("db_path 不存在: %s", db_path)

    line_cfg = cfg.get("line_server", {})
    if line_cfg.get("url") and not line_cfg.get("secret"):
        msg = "LINE 伺服器已設定網址，但 API Secret 為空，同步將會失敗。"
        warnings.append(msg)
        log.warning("LINE server url 已設定但 secret 為空")

    if not cfg.get("operators"):
        msg = "尚未設定任何人員，請至 ⚙ 設定新增人員與代號。"
        warnings.append(msg)
        log.warning("operators 未設定")

    if warnings:
        messagebox.showwarning("設定警告", "\n\n".join(warnings))


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
    _LineTab,
    ctk.CTk,
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
        set_db_path(self._config.get("db_path"))
        init_db()
        self.after(200, lambda: _validate_config(self._config))
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
        top = ctk.CTkFrame(self, fg_color="#1a5276", corner_radius=0, height=48)
        top.pack(fill="x")
        top.pack_propagate(False)
        ctk.CTkLabel(top, text="立善科技｜QuoteFlow",
                     text_color="white",
                     font=ctk.CTkFont("Microsoft JhengHei UI", 16, "bold")
                     ).pack(side="left", padx=16)
        ctk.CTkButton(top, text="選擇報價單 .xlsx ▶", command=self._open_file,
                      fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                      font=ctk.CTkFont("Microsoft JhengHei UI", 10),
                      width=160, height=32, corner_radius=6
                      ).pack(side="right", padx=(4, 16), pady=8)
        ctk.CTkButton(top, text="⚙", command=self._open_paths_dialog,
                      fg_color="transparent", hover_color="#2e4057", text_color="white",
                      font=ctk.CTkFont("Microsoft JhengHei UI", 16),
                      width=36, height=32, corner_radius=6
                      ).pack(side="right", padx=(0, 2), pady=8)
        ctk.CTkButton(top, text="開啟成本表", command=self._open_cost_file,
                      fg_color="transparent", hover_color="#2e4057", text_color="#d0d8e0",
                      font=ctk.CTkFont("Microsoft JhengHei UI", 9),
                      width=80, height=32, corner_radius=6
                      ).pack(side="right", padx=(0, 2), pady=8)

        self._file_label = ctk.CTkLabel(self, text="⚠  尚未選擇報價單",
                                        fg_color="transparent",
                                        text_color="#c0392b", font=FONT, anchor="w")
        self._file_label.pack(anchor="w", padx=16, pady=(3, 0))

        # ── 主體：側邊列 + 內容區 ────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)

        # ── 左側導覽列 ────────────────────────────────────────
        nav = ctk.CTkFrame(body, fg_color=NAV_BG, corner_radius=0, width=134)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)

        content_area = tk.Frame(body, bg=BG)
        content_area.pack(side="left", fill="both", expand=True)

        # ── 頁面 Frame（每個功能一個，重疊在 content_area）──────
        self._pages:         dict[str, tk.Frame]    = {}
        self._page_builders: dict[str, object]      = {}
        self._page_built:    set[str]               = set()
        self._nav_btns:      dict[str, ctk.CTkButton] = {}
        self._active_page: str = ""

        def _show(key: str):
            if self._active_page and self._active_page in self._nav_btns:
                self._nav_btns[self._active_page].configure(fg_color="transparent")
            for f in self._pages.values():
                f.pack_forget()
            if key not in self._page_built:
                self._page_builders[key]()
                self._page_built.add(key)
            self._pages[key].pack(fill="both", expand=True)
            self._nav_btns[key].configure(fg_color=NAV_ACTIVE)
            self._active_page = key

        self._show_page = _show

        def _make_page(key: str) -> tk.Frame:
            f = tk.Frame(content_area, bg=BG)
            self._pages[key] = f
            return f

        def _register(key: str, builder, *args):
            page = _make_page(key)
            self._page_builders[key] = lambda p=page: builder(p, *args)

        # ── 導覽列項目生成 ────────────────────────────────────
        def _nav_group(text: str):
            ctk.CTkLabel(nav, text=text, fg_color="transparent",
                          text_color=NAV_GRP_FG,
                          font=("Microsoft JhengHei UI", 8),
                          anchor="w").pack(fill="x", padx=10, pady=(10, 1))
            ctk.CTkFrame(nav, fg_color=NAV_HOVER, height=1,
                          corner_radius=0).pack(fill="x", padx=8)

        def _nav_item(text: str, key: str):
            btn = ctk.CTkButton(nav, text=f"  {text}",
                                 fg_color="transparent",
                                 hover_color=NAV_HOVER,
                                 text_color=NAV_FG,
                                 anchor="w",
                                 font=("Microsoft JhengHei UI", 9),
                                 height=34, corner_radius=4,
                                 command=lambda k=key: _show(k))
            btn.pack(fill="x", padx=4, pady=1)
            self._nav_btns[key] = btn

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
        _nav_item("出貨指示單",   "shipping_order")
        _nav_item("建立卡片",     "create")
        _nav_item("下載卡片",     "download")
        _nav_item("會計對帳",     "accounting")

        _nav_group("🗂  記錄查詢")
        _nav_item("報價記錄",     "history")
        _nav_item("LINE 詢問",    "line")

        # ── 登記各頁面（lazy：第一次切換時才建立 widgets）────────
        _register("shipping",      self._build_tab_shipping,      PAD, FONT, FONTB, BG)
        _register("quote",         self._build_tab_quote,         FONT, FONTB, BG)
        _register("inspection",    self._build_tab_inspection,    PAD, FONT, FONTB, BG)
        _register("fix",           self._build_tab_fix,           PAD, FONT, FONTB, BG)
        _register("tag",           self._build_tab_tag,           PAD, FONT, FONTB, BG)
        _register("label",         self._build_tab_label,         FONT, FONTB, BG)
        _register("schedule",      self._build_tab_schedule,      FONT, FONTB, BG)
        _register("overview",      self._build_tab_overview,      FONT, FONTB, BG)
        _register("production",    self._build_tab_production,    FONT, FONTB, BG)
        _register("shipping_order",self._build_tab_shipping_order,FONT, FONTB, BG)
        _register("create",        self._build_tab_create_cards,  FONT, FONTB, BG)
        _register("download",      self._build_tab_download_cards,FONT, FONTB, BG)
        _register("accounting",    self._build_tab_accounting,    FONT, FONTB, BG)
        _register("history",       self._build_tab_history,       FONT, FONTB, BG)
        _register("line",          self._build_tab_line,          FONT, FONTB, BG)

        # 預設顯示出貨單（此時才真正建立 shipping 頁面）
        _show("shipping")

    # ════════════════════════════════════════════════════════
    #  路徑設定 Dialog
    def _open_cost_file(self):
        import os
        from _paths import TEMPLATE_DIR
        cost_path = TEMPLATE_DIR / "template_cost.xlsx"
        if not cost_path.exists():
            from tkinter import messagebox
            messagebox.showwarning("找不到檔案", f"找不到成本表：\n{cost_path}")
            return
        os.startfile(str(cost_path))

    # ════════════════════════════════════════════════════════
    def _open_paths_dialog(self):
        BG     = "#f4f6f8"
        GRAY   = "#5d6d7e"
        FONT   = ("Microsoft JhengHei UI", 10)
        FONTB  = ("Microsoft JhengHei UI", 10, "bold")
        FONT_S = ("Microsoft JhengHei UI", 9)
        PAD    = {"padx": 8, "pady": 5}

        dlg = ctk.CTkToplevel(self)
        dlg.title("⚙  路徑與人員設定")
        dlg.configure(fg_color=BG)
        dlg.resizable(True, True)
        dlg.after(100, dlg.grab_set)
        dlg.update_idletasks()
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        h = min(720, sh - 80)
        dlg.geometry(f"720x{h}+{(sw-720)//2}+{(sh-h)//2}")

        # ── 資料庫路徑設定 ────────────────────────────────────
        db_outer, db_lf = _mk_lf(dlg, "資料庫設定", BG, FONTB)
        db_outer.pack(fill="x", padx=16, pady=(16, 6))
        db_lf.columnconfigure(1, weight=1)

        ctk.CTkLabel(db_lf, text="資料庫路徑：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=(8, 2), pady=5)
        db_path_var = tk.StringVar(value=self._config.get("db_path", ""))
        ctk.CTkEntry(db_lf, textvariable=db_path_var, font=FONT_S,
                      corner_radius=4, border_width=1
                      ).grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=5)

        def _pick_db():
            p = filedialog.askopenfilename(
                title="選擇資料庫檔案",
                filetypes=[("SQLite 資料庫", "*.db"), ("所有檔案", "*.*")],
                parent=dlg)
            if p:
                db_path_var.set(p)
                self._config["db_path"] = p
                set_db_path(p)
                self._save_config(self._config)

        ctk.CTkButton(db_lf, text="選擇", command=_pick_db,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=60, height=28, corner_radius=4
                       ).grid(row=0, column=2, padx=(0, 8), pady=5)
        ctk.CTkLabel(db_lf,
                      text="留空 = 使用本機預設路徑；填入 NAS 路徑可多台電腦共用同一份資料庫",
                      fg_color="transparent",
                      font=("Microsoft JhengHei UI", 8), text_color=GRAY
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

        path_outer, path_lf = _mk_lf(dlg, "輸出路徑設定", BG, FONTB)
        path_outer.pack(fill="x", padx=16, pady=(0, 6))
        path_lf.columnconfigure(1, weight=1)

        path_vars: dict[str, tk.StringVar] = {}
        paths_cfg = self._config.get("paths", {})

        for i, (key, label, is_file) in enumerate(items):
            ctk.CTkLabel(path_lf, text=label + "：", fg_color="transparent",
                          font=FONT_S, text_color=GRAY,
                          anchor="w").grid(row=i, column=0, sticky="w", **PAD)
            var = tk.StringVar(value=paths_cfg.get(key) or self._PATH_DEFAULTS[key])
            path_vars[key] = var
            ctk.CTkEntry(path_lf, textvariable=var, font=FONT_S,
                          corner_radius=4, border_width=1
                          ).grid(row=i, column=1, sticky="ew", padx=(0, 4), pady=5)

            def _pick(v=var, f=is_file):
                if f:
                    p = filedialog.askopenfilename(
                        filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
                else:
                    p = filedialog.askdirectory()
                if p:
                    v.set(p)

            ctk.CTkButton(path_lf, text="選擇", command=_pick,
                           fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                           font=FONT_S, width=60, height=28, corner_radius=4
                           ).grid(row=i, column=2, padx=(0, 8), pady=5)

        # ── 人員設定 ─────────────────────────────────────────
        op_outer, op_inner = _mk_lf(dlg, "人員設定（報價單製表人員與單號代號）", BG, FONTB)
        op_outer.pack(fill="x", padx=16, pady=(0, 6))

        op_cols = ("name", "code")
        op_tree = ttk.Treeview(op_inner, columns=op_cols, show="headings",
                                selectmode="browse", height=4)
        op_tree.heading("name", text="名稱（放入 Excel 製表人欄）")
        op_tree.heading("code", text="代號（報價單號前綴）")
        op_tree.column("name", width=200, anchor="w")
        op_tree.column("code", width=120, anchor="center")
        op_tree.pack(fill="x", padx=8, pady=(4, 2))

        operators = self._config.get("operators", ["小皋"])
        op_codes  = self._config.get("operator_codes", {})
        for name in operators:
            op_tree.insert("", "end", values=(name, op_codes.get(name, "")))

        edit_row = ctk.CTkFrame(op_inner, fg_color="transparent", corner_radius=0)
        edit_row.pack(fill="x", padx=8, pady=(0, 6))

        ctk.CTkLabel(edit_row, text="名稱：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        name_var = tk.StringVar()
        ctk.CTkEntry(edit_row, textvariable=name_var, font=FONT_S,
                      width=100, height=28, corner_radius=4
                      ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(edit_row, text="代號：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        code_var = tk.StringVar()
        ctk.CTkEntry(edit_row, textvariable=code_var, font=FONT_S,
                      width=70, height=28, corner_radius=4
                      ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(edit_row, text="（英文字串，如 K）", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left", padx=(0, 12))

        def _op_add():
            n = name_var.get().strip()
            c = code_var.get().strip().upper()
            if not n or not c:
                messagebox.showwarning("欄位不完整", "請填入名稱和代號", parent=dlg)
                return
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

        ctk.CTkButton(edit_row, text="新增 / 更新", command=_op_add,
                       fg_color="#117a65", hover_color="#0e6655", text_color="white",
                       font=FONT_S, width=90, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(edit_row, text="刪除選取", command=_op_del,
                       fg_color="#c0392b", hover_color="#a93226", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="left")

        def _reset_defaults():
            for key, var in path_vars.items():
                var.set(self._PATH_DEFAULTS[key])

        def _save():
            self._config.setdefault("paths", {})
            for key, var in path_vars.items():
                self._config["paths"][key] = var.get().strip()

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

        bb = ctk.CTkFrame(dlg, fg_color="transparent", corner_radius=0)
        bb.pack(fill="x", padx=16, pady=8)
        ctk.CTkButton(bb, text="還原路徑預設值", command=_reset_defaults,
                       fg_color="#5d6d7e", hover_color="#4d5d6e", text_color="white",
                       font=FONT, height=34, corner_radius=6
                       ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(bb, text="儲存並關閉", command=_save,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=FONTB, height=34, corner_radius=6
                       ).pack(side="left")

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
            self._file_label.configure(text=f"✔  已載入：{path}", text_color="#1e8449")
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
