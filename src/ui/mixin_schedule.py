"""
mixin_schedule.py — 出貨排程頁籤 mixin
"""
import os
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path


class _ScheduleTab:
    """Mixin providing _build_tab_schedule and its callbacks."""

    def _build_tab_schedule(self, parent, FONT, FONTB, BG):
        from tkcalendar import DateEntry
        from _paths import TEMPLATE_DIR
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
            self._save_config(self._config)
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
            self._save_config(self._config)
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
            from core.generator_schedule import calculate_travel_times
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
            from core.generator_schedule import sort_rows_by_location
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
            from core.generator_schedule import fetch_events, events_to_rows
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
            from core.generator_schedule import generate_schedule
            if not _rows:
                messagebox.showwarning("無資料", "請先抓取並確認事件清單", parent=parent)
                return
            sid  = sid_var.get().strip()
            csrf = csrf_var.get().strip()
            target = date_entry.get_date()
            try:
                out = generate_schedule(target, sid, csrf, rows=list(_rows),
                                        schedule_file=self._get_path("schedule_file"))
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
