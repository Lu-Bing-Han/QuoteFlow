"""
mixin_trello.py — Trello 相關頁籤 mixin（出貨一覽表、生產群組紀錄、建立卡片、下載卡片）
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from pathlib import Path
from ui.app_core import _mk_lf


class _TrelloTab:
    """Mixin providing Trello-related tab builders and callbacks."""

    # ════════════════════════════════════════════════════════
    #  Tab 7：出貨一覽表
    # ════════════════════════════════════════════════════════
    def _build_tab_overview(self, parent, FONT, FONTB, BG):
        from _paths import _GSHEETS_CREDS_PATH, _GSHEETS_TOKEN_PATH, _SYNCED_CARDS_PATH  # noqa: F401
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        # ── Trello 憑證 ───────────────────────────────────
        cred_outer, cred_frame = _mk_lf(parent, "Trello 憑證", BG, FONTB)
        cred_outer.pack(fill="x", padx=12, pady=(12, 4))
        cred_frame.columnconfigure(1, weight=1)

        tr_cfg   = self._config.get("trello", {})
        key_var  = tk.StringVar(value=tr_cfg.get("api_key", ""))
        tok_var  = tk.StringVar(value=tr_cfg.get("token",   ""))

        ctk.CTkLabel(cred_frame, text="API Key：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        key_entry = ctk.CTkEntry(cred_frame, textvariable=key_var, font=FONT_S,
                                  show="*", corner_radius=4, border_width=1)
        key_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=3)

        ctk.CTkLabel(cred_frame, text="Token：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=3)
        tok_entry = ctk.CTkEntry(cred_frame, textvariable=tok_var, font=FONT_S,
                                  show="*", corner_radius=4, border_width=1)
        tok_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=3)

        def _show_hide(entry, btn):
            if entry.cget("show") == "*":
                entry.configure(show=""); btn.configure(text="隱藏")
            else:
                entry.configure(show="*"); btn.configure(text="顯示")

        btn_key = ctk.CTkButton(cred_frame, text="顯示", font=FONT_S,
                                 fg_color=GRAY, hover_color="#4d5d6e",
                                 text_color="white", width=50, height=26, corner_radius=4)
        btn_key.configure(command=lambda: _show_hide(key_entry, btn_key))
        btn_key.grid(row=0, column=2, padx=(0, 8), pady=3)

        btn_tok = ctk.CTkButton(cred_frame, text="顯示", font=FONT_S,
                                 fg_color=GRAY, hover_color="#4d5d6e",
                                 text_color="white", width=50, height=26, corner_radius=4)
        btn_tok.configure(command=lambda: _show_hide(tok_entry, btn_tok))
        btn_tok.grid(row=1, column=2, padx=(0, 8), pady=3)

        def _save_trello_creds():
            self._config.setdefault("trello", {})
            self._config["trello"]["api_key"] = key_var.get().strip()
            self._config["trello"]["token"]   = tok_var.get().strip()
            self._save_config(self._config)
            messagebox.showinfo("已儲存", "Trello 憑證已儲存", parent=parent)

        ctk.CTkButton(cred_frame, text="儲存憑證", command=_save_trello_creds,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT, width=90, height=28, corner_radius=4
                       ).grid(row=2, column=1, sticky="w", padx=8, pady=(4, 6))

        # ── Google Sheets 設定 ────────────────────────────
        gs_outer, gs_frame = _mk_lf(parent, "Google Sheets", BG, FONTB)
        gs_outer.pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(gs_frame,
                      text="✔  credentials.json 已就緒" if _GSHEETS_CREDS_PATH.exists()
                           else "✘  找不到 credentials.json（請放到 template 資料夾）",
                      fg_color="transparent", font=FONT_S, anchor="w",
                      text_color="#1e8449" if _GSHEETS_CREDS_PATH.exists() else "#c0392b",
                      ).pack(fill="x", padx=8, pady=6)

        token_status = ctk.CTkLabel(gs_frame,
                                     text="✔  已授權（gsheets_token.json 存在）" if _GSHEETS_TOKEN_PATH.exists()
                                          else "尚未授權，點「同步」時會自動開啟瀏覽器",
                                     fg_color="transparent", font=FONT_S, anchor="w",
                                     text_color="#1e8449" if _GSHEETS_TOKEN_PATH.exists() else GRAY)
        token_status.pack(fill="x", padx=8, pady=(0, 6))

        # ── 同步按鈕與狀態 ────────────────────────────────
        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _sync():
            from sync.syncer_trello import fetch_po_cards
            from sync.syncer_sheets import sync_cards
            api_key = key_var.get().strip()
            token   = tok_var.get().strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未填", "請先填入 Trello API Key 與 Token", parent=parent); return
            if not _GSHEETS_CREDS_PATH.exists():
                messagebox.showerror("缺少憑證", f"找不到 {_GSHEETS_CREDS_PATH}", parent=parent); return

            out_label.configure(text="抓取 Trello 卡片中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                cards = fetch_po_cards(api_key, token)
                out_label.configure(text=f"找到 {len(cards)} 張卡片，同步至 Google Sheets…", text_color=GRAY)
                parent.update_idletasks()
                added = sync_cards(cards, _GSHEETS_CREDS_PATH,
                                   _GSHEETS_TOKEN_PATH, _SYNCED_CARDS_PATH)
                token_status.configure(text="✔  已授權（gsheets_token.json 存在）", text_color="#1e8449")
                if added:
                    out_label.configure(text=f"✔  同步完成，新增 {added} 筆資料", text_color="#1e8449")
                else:
                    out_label.configure(text="✔  同步完成，無新卡片", text_color="#1e8449")
            except Exception as e:
                out_label.configure(text=f"✘  {e}", text_color="#c0392b")
                messagebox.showerror("同步失敗", str(e), parent=parent)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🔄  同步 Trello → Google Sheets", command=_sync,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 8：生產群組紀錄
    # ════════════════════════════════════════════════════════
    def _build_tab_production(self, parent, FONT, FONTB, BG):
        from _paths import _PRODUCTION_SYNCED_PATH  # noqa: F401
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        info_outer, info = _mk_lf(parent, "說明", BG, FONTB)
        info_outer.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(info,
                      text="從 Trello「本周下單」抓取 2026/5/15 之後的新卡片，附加到生產群組紀錄 Excel。\n"
                           "Trello 憑證與「出貨一覽表」頁籤共用，請先在該頁籤儲存憑證。",
                      fg_color="transparent", font=FONT_S, text_color=GRAY,
                      justify="left", anchor="w",
                      ).pack(padx=12, pady=8, anchor="w")

        pf = ctk.CTkFrame(parent, fg_color="#e8ecf0", corner_radius=6)
        pf.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(pf, text="寫入檔案：", fg_color="transparent",
                      font=FONT, text_color=GRAY, anchor="w", width=80).pack(side="left", padx=8, pady=6)
        ctk.CTkLabel(pf, text="（依⚙路徑設定）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY).pack(side="left", pady=6)

        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(8, 0))

        def _sync():
            from sync.syncer_trello import fetch_po_cards
            from sync.syncer_production import sync_production
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent); return

            out_label.configure(text="抓取 Trello 卡片中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                cards = fetch_po_cards(api_key, token)
                out_label.configure(text=f"找到 {len(cards)} 張卡片，寫入 Excel 中…", text_color=GRAY)
                parent.update_idletasks()
                added = sync_production(cards, _PRODUCTION_SYNCED_PATH,
                                        production_file=self._get_path("production_file"))
                if added:
                    out_label.configure(text=f"✔  同步完成，新增 {added} 筆資料", text_color="#1e8449")
                else:
                    out_label.configure(text="✔  同步完成，無新卡片（2026/5/15 之後）", text_color="#1e8449")
                if messagebox.askyesno("同步完成",
                        f"{'新增 ' + str(added) + ' 筆資料' if added else '無新卡片'}\n\n是否立即開啟生產群組紀錄？",
                        parent=parent):
                    os.startfile(str(self._get_path("production_file")))
            except Exception as e:
                out_label.configure(text=f"✘  {e}", text_color="#c0392b")
                messagebox.showerror("同步失敗", str(e), parent=parent)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🔄  同步 Trello → 生產群組紀錄.xlsx", command=_sync,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 9：建立卡片
    # ════════════════════════════════════════════════════════
    def _build_tab_create_cards(self, parent, FONT, FONTB, BG):
        from tksheet import Sheet
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _HEADERS = ["序號", "標題", "描述", "備註/需求"]
        _EMPTY   = ["", "", "", ""]

        top_row = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        top_row.pack(fill="x", padx=12, pady=(12, 4))

        src_outer, src_frame = _mk_lf(top_row, "來源 Excel", BG, FONTB)
        src_outer.pack(side="left", fill="y", padx=(0, 8))
        src_frame.columnconfigure(1, weight=1)

        hint_outer, hint_frame = _mk_lf(top_row, "使用說明", BG, FONTB)
        hint_outer.pack(side="left", fill="both", expand=True)
        hint_text = "導入 Excel 後即可編輯資料\n可以選取多筆資料搬移刪除\n不可跳號選取"
        ctk.CTkLabel(hint_frame, text=hint_text, fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      justify="left", anchor="nw").pack(padx=10, pady=8, anchor="nw")

        ctk.CTkLabel(src_frame, text="檔案路徑：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        path_var = tk.StringVar()
        ctk.CTkEntry(src_frame, textvariable=path_var, font=FONT_S,
                      corner_radius=4, border_width=1
                      ).grid(row=0, column=1, sticky="ew", padx=4, pady=6)

        ctk.CTkLabel(src_frame, text="工作表：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        sheet_var = tk.StringVar()
        sheet_cb  = ttk.Combobox(src_frame, textvariable=sheet_var, font=FONT_S,
                                  state="readonly", width=20)
        sheet_cb.grid(row=1, column=1, sticky="w", padx=4, pady=4)

        def _pick_file():
            from sync.creator_trello import get_sheet_names
            p = filedialog.askopenfilename(
                title="選擇 Excel 檔案",
                filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")])
            if not p: return
            path_var.set(p)
            try:
                names = get_sheet_names(Path(p))
                all_options = ["全部"] + names
                sheet_cb["values"] = all_options
                sheet_var.set(all_options[0])
            except Exception:
                sheet_cb["values"] = []
                sheet_var.set("")

        ctk.CTkButton(src_frame, text="選擇", command=_pick_file,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=56, height=26, corner_radius=4
                       ).grid(row=0, column=2, padx=(0, 4), pady=6)

        status_lbl = ctk.CTkLabel(src_frame, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY)
        status_lbl.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        # ── 序號篩選列 ────────────────────────────────────
        filter_frame = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        filter_frame.pack(fill="x", padx=12, pady=(0, 2))
        ctk.CTkLabel(filter_frame, text="序號篩選：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY).pack(side="left")
        filter_var = tk.StringVar()
        ctk.CTkEntry(filter_frame, textvariable=filter_var, font=FONT_S,
                      width=240, height=28, corner_radius=4
                      ).pack(side="left", padx=(4, 6))
        ctk.CTkLabel(filter_frame, text="（逗號分隔，例如 1,3,5；空白=全部顯示）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY).pack(side="left")

        # ── 可編輯 Sheet 預覽 ─────────────────────────────
        prev_outer = ctk.CTkFrame(parent, fg_color=BG, corner_radius=8,
                                   border_width=1, border_color="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_title_lbl = ctk.CTkLabel(prev_outer,
                                       text="  卡片預覽（雙擊儲存格可編輯，0 筆）  ",
                                       fg_color=BG, text_color=GRAY,
                                       font=FONT)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(6, 0))
        prev_frame = ctk.CTkFrame(prev_outer, fg_color=BG, corner_radius=0)
        prev_frame.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        sheet = Sheet(prev_frame,
                      headers=_HEADERS,
                      data=[_EMPTY[:] for _ in range(10)],
                      column_width=200,
                      row_height=28)
        sheet.set_column_widths([70, 300, 420, 200])
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

        _all_data: list[dict] = []

        def _apply_filter():
            raw = filter_var.get().strip()
            if raw:
                keys = {s.strip() for s in raw.split(",") if s.strip()}
                filtered = [c for c in _all_data if str(c["seq"]) in keys]
            else:
                filtered = _all_data
            rows = [[c["seq"], c["title"], c["desc"], c.get("notes", "")] for c in filtered]
            while len(rows) < len(filtered) + 5:
                rows.append(_EMPTY[:])
            sheet.data = rows
            prev_title_lbl.configure(text=f"  卡片預覽（雙擊儲存格可編輯，{len(filtered)} 筆）  ")

        ctk.CTkButton(filter_frame, text="篩選", command=_apply_filter,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=56, height=28, corner_radius=4
                       ).pack(side="left")
        ctk.CTkButton(filter_frame, text="清除",
                       command=lambda: (filter_var.set(""), _apply_filter()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=56, height=28, corner_radius=4
                       ).pack(side="left", padx=(4, 0))

        def _load_preview():
            from sync.creator_trello import read_excel_cards
            p = path_var.get().strip()
            if not p:
                messagebox.showwarning("未選擇檔案", "請先選擇 Excel 檔案", parent=parent); return
            status_lbl.configure(text="讀取中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                selected = sheet_var.get()
                _all_data.clear()
                if selected == "全部":
                    all_sheets = [v for v in sheet_cb["values"] if v != "全部"]
                    for sname in all_sheets:
                        _all_data.extend(read_excel_cards(Path(p), sheet_name=sname))
                else:
                    _all_data.extend(read_excel_cards(Path(p), sheet_name=selected or None))
                filter_var.set("")
                _apply_filter()
                status_lbl.configure(text=f"✔  讀取完成，共 {len(_all_data)} 筆", text_color="#1e8449")
            except Exception as e:
                status_lbl.configure(text=f"✘  {e}", text_color="#c0392b")

        ctk.CTkButton(src_frame, text="讀取預覽", command=_load_preview,
                       fg_color="#117a65", hover_color="#0e6655", text_color="white",
                       font=FONT_S, width=70, height=26, corner_radius=4
                       ).grid(row=0, column=3, padx=(0, 8), pady=6)

        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _create():
            from sync.creator_trello import create_cards as trello_create_cards
            rows = sheet.data
            cards = []
            for row in rows:
                title = str(row[1]).strip() if len(row) > 1 else ""
                desc  = str(row[2]).strip() if len(row) > 2 else ""
                notes = str(row[3]).strip() if len(row) > 3 else ""
                if not title: continue
                if notes:
                    desc = desc + f"\n需求：\n{notes}"
                cards.append({"title": title, "desc": desc})
            if not cards:
                messagebox.showwarning("無資料", "表格中沒有填入標題的列", parent=parent); return
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent); return
            if not messagebox.askyesno("確認建立",
                    f"即將在「0.待評估」清單建立 {len(cards)} 張卡片，確定繼續？", parent=parent): return

            out_label.configure(text="建立中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                created = trello_create_cards(cards, api_key, token)
                out_label.configure(text=f"✔  成功建立 {created} 張卡片", text_color="#1e8449")
            except Exception as e:
                out_label.configure(text=f"✘  {e}", text_color="#c0392b")
                messagebox.showerror("建立失敗", str(e), parent=parent)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(bb, text="🃏  建立全部卡片 → Trello 0.待評估", command=_create,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")

    # ════════════════════════════════════════════════════════
    #  Tab 10：下載卡片
    # ════════════════════════════════════════════════════════
    def _build_tab_download_cards(self, parent, FONT, FONTB, BG):
        GRAY   = "#5d6d7e"
        FONT_S = ("Microsoft JhengHei UI", 9)

        _list_map: dict[str, str] = {}
        _all_cards: list[dict]    = []

        top_outer, top = _mk_lf(parent, "Trello 清單", BG, FONTB)
        top_outer.pack(fill="x", padx=12, pady=(12, 4))
        top.columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="選擇清單：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY,
                      anchor="w").grid(row=0, column=0, sticky="w", padx=8, pady=6)

        list_var = tk.StringVar()
        list_cb  = ttk.Combobox(top, textvariable=list_var, font=FONT_S,
                                 state="readonly", width=28)
        list_cb.grid(row=0, column=1, sticky="w", padx=4, pady=6)

        status_lbl = ctk.CTkLabel(top, text="", fg_color="transparent",
                                   font=FONT_S, text_color=GRAY)

        def _get_creds():
            tr_cfg  = self._config.get("trello", {})
            api_key = tr_cfg.get("api_key", "").strip()
            token   = tr_cfg.get("token",   "").strip()
            if not api_key or not token:
                messagebox.showwarning("憑證未設定",
                    "請先至「出貨一覽表」頁籤填入並儲存 Trello 憑證", parent=parent)
                return None, None
            return api_key, token

        def _fetch_lists():
            from sync.downloader_trello import get_board_lists
            api_key, token = _get_creds()
            if not api_key: return
            status_lbl.configure(text="抓取中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                lists = get_board_lists(api_key, token)
                _list_map.clear()
                _list_map.update({lst["name"]: lst["id"] for lst in lists})
                names = list(_list_map.keys())
                list_cb["values"] = names
                if names:
                    list_var.set(names[0])
                status_lbl.configure(text=f"找到 {len(lists)} 個清單", text_color="#1e8449")
            except Exception as e:
                status_lbl.configure(text=f"✘ {e}", text_color="#c0392b")

        def _preview_cards():
            from sync.downloader_trello import get_list_cards
            selected = list_var.get().strip()
            if not selected or selected not in _list_map:
                messagebox.showwarning("未選擇清單", "請先抓取清單並選擇", parent=parent); return
            api_key, token = _get_creds()
            if not api_key: return
            status_lbl.configure(text="載入卡片中…", text_color=GRAY)
            parent.update_idletasks()
            try:
                cards = get_list_cards(_list_map[selected], api_key, token)
                _all_cards.clear()
                _all_cards.extend(cards)
                tree.delete(*tree.get_children())
                for card in cards:
                    labels = "、".join(
                        lbl.get("name") or lbl.get("color", "")
                        for lbl in card.get("labels", []))
                    att_count = len(card.get("attachments") or [])
                    tree.insert("", "end", values=(card["name"], labels, att_count))
                tree.selection_set(tree.get_children())
                status_lbl.configure(text=f"找到 {len(cards)} 張卡片", text_color="#1e8449")
                prev_title_lbl.configure(text=f"  卡片預覽（共 {len(cards)} 張，可多選）  ")
            except Exception as e:
                status_lbl.configure(text=f"✘ {e}", text_color="#c0392b")

        ctk.CTkButton(top, text="抓取清單", command=_fetch_lists,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT_S, width=70, height=26, corner_radius=4
                       ).grid(row=0, column=2, padx=(0, 4), pady=6)
        ctk.CTkButton(top, text="預覽卡片", command=_preview_cards,
                       fg_color="#117a65", hover_color="#0e6655", text_color="white",
                       font=FONT_S, width=70, height=26, corner_radius=4
                       ).grid(row=0, column=3, padx=(0, 4), pady=6)
        status_lbl.grid(row=0, column=4, sticky="w", padx=4, pady=6)

        # ── 卡片預覽 Treeview ─────────────────────────────
        prev_outer = ctk.CTkFrame(parent, fg_color=BG, corner_radius=8,
                                   border_width=1, border_color="#d0d7de")
        prev_outer.pack(fill="both", expand=True, padx=12, pady=4)
        prev_title_lbl = ctk.CTkLabel(prev_outer,
                                       text="  卡片預覽（共 0 張，可多選）  ",
                                       fg_color=BG, text_color=GRAY, font=FONT)
        prev_title_lbl.pack(anchor="w", padx=10, pady=(6, 0))
        prev_inner = ctk.CTkFrame(prev_outer, fg_color=BG, corner_radius=0)
        prev_inner.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        cols = ("title", "labels", "att")
        tree = ttk.Treeview(prev_inner, columns=cols, show="headings",
                             selectmode="extended", height=8)
        tree.heading("title",  text="標題")
        tree.heading("labels", text="標籤")
        tree.heading("att",    text="附件")
        tree.column("title",  width=340, anchor="w",      stretch=True)
        tree.column("labels", width=130, anchor="w",      stretch=False)
        tree.column("att",    width=50,  anchor="center", stretch=False)

        vsb = ttk.Scrollbar(prev_inner, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        sel_row = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        sel_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkButton(sel_row, text="全選",
                       command=lambda: tree.selection_set(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=60, height=28, corner_radius=4
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(sel_row, text="取消全選",
                       command=lambda: tree.selection_remove(tree.get_children()),
                       fg_color=GRAY, hover_color="#4d5d6e", text_color="white",
                       font=FONT_S, width=80, height=28, corner_radius=4
                       ).pack(side="left")

        pf = ctk.CTkFrame(parent, fg_color="#e8ecf0", corner_radius=6)
        pf.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(pf, text="輸出資料夾：", fg_color="transparent",
                      font=FONT_S, text_color=GRAY, anchor="w", width=96).pack(side="left", padx=8, pady=5)
        ctk.CTkLabel(pf, text="（依⚙路徑設定）",
                      fg_color="transparent", font=FONT_S, text_color=GRAY).pack(side="left", pady=5)

        out_label = ctk.CTkLabel(parent, text="", fg_color="transparent",
                                  font=FONT_S, text_color=GRAY, anchor="w", wraplength=700)
        out_label.pack(fill="x", padx=16, pady=(4, 0))

        def _download():
            from sync.downloader_trello import download_cards as trello_download_cards
            sel_ids = tree.selection()
            if not sel_ids:
                messagebox.showwarning("未選擇", "請先預覽並選取要下載的卡片", parent=parent); return
            api_key, token = _get_creds()
            if not api_key: return

            indices  = [tree.index(i) for i in sel_ids]
            selected_cards = [_all_cards[i] for i in indices]
            list_name  = list_var.get().strip()
            output_dir = self._get_path("download_cards_dir") / list_name

            out_label.configure(text="下載中…", text_color=GRAY)
            parent.update_idletasks()

            def _progress(cur, total, name):
                out_label.configure(text=f"下載中… ({cur}/{total}) {name}", text_color=GRAY)
                parent.update_idletasks()

            try:
                count, failed = trello_download_cards(
                    selected_cards, output_dir, api_key, token, progress_cb=_progress)
                if failed:
                    out_label.configure(
                        text=f"✔  完成 {count} 張，{len(failed)} 個附件失敗", text_color="#e67e22")
                    messagebox.showwarning("部分附件失敗",
                        "以下附件下載失敗：\n" + "\n".join(failed), parent=parent)
                else:
                    out_label.configure(
                        text=f"✔  成功下載 {count} 張卡片至：{output_dir}", text_color="#1e8449")
                if messagebox.askyesno("下載完成",
                        f"共 {count} 張卡片\n\n是否立即開啟資料夾？", parent=parent):
                    os.startfile(str(output_dir))
            except Exception as e:
                out_label.configure(text=f"✘  {e}", text_color="#c0392b")
                messagebox.showerror("下載失敗", str(e), parent=parent)

        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=6)
        ctk.CTkButton(bb, text="⬇  下載選取的卡片", command=_download,
                       fg_color="#1a5276", hover_color="#154360", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")
