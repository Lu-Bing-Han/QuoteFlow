"""
mixin_label.py — 標籤生成頁籤 mixin
"""
import os
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from datetime import datetime
from pathlib import Path
from ui.app_core import _mk_lf


class _LabelTab:
    """Mixin providing _build_tab_label and its callbacks."""

    def _build_tab_label(self, parent, FONT, FONTB, BG):
        from tksheet import Sheet
        import re as _re

        GRAY = "#5d6d7e"

        # 模板選擇
        tpl_frame = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        tpl_frame.pack(fill="x", padx=12, pady=(10, 0))
        ctk.CTkLabel(tpl_frame, text="標籤樣式：", fg_color="transparent",
                      font=FONTB, text_color="#2c3e50").pack(side="left")
        _tpl_var = tk.StringVar(value="銀標")
        _tpl_cb  = ctk.CTkComboBox(tpl_frame, variable=_tpl_var,
                                    values=["銀標", "APT標", "無公司標", "上銀標"],
                                    font=FONT, width=120, height=28)
        _tpl_cb.pack(side="left", padx=(8, 0))

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

        tf_outer, tf = _mk_lf(parent, "標籤資料", BG, FONTB)
        tf_outer.pack(fill="both", expand=True, padx=12, pady=(10, 4))

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
        bb = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        bb.pack(fill="x", padx=12, pady=(0, 4))
        ctk.CTkButton(bb, text="從報價單讀入", command=_load_from_quote,
                       fg_color="#2e86c1", hover_color="#1a5276", text_color="white",
                       font=FONT, width=110, height=32, corner_radius=6
                       ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(bb, text="＋ 新增列",
                       command=lambda: sheet.insert_rows(number=1),
                       fg_color="#27ae60", hover_color="#1e8449", text_color="white",
                       font=FONT, width=90, height=32, corner_radius=6
                       ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(bb, text="－ 刪除列",
                       command=lambda: [sheet.delete_rows(row=r)
                                        for r in sorted(sheet.get_selected_rows(), reverse=True)],
                       fg_color="#c0392b", hover_color="#a93226", text_color="white",
                       font=FONT, width=90, height=32, corner_radius=6
                       ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(bb, text="流水號↓", command=_autofill_serial,
                       fg_color="#7d3c98", hover_color="#6c3483", text_color="white",
                       font=FONT, width=90, height=32, corner_radius=6
                       ).pack(side="left")

        # 生成按鈕
        def _generate():
            from core.generator_label import generate_labels
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
            out_path = self._get_path("output_label") / f"標籤-{date_tag}.pdf"
            try:
                result = generate_labels(data_list, out_path, template_key=tpl)
                if messagebox.askyesno("生成成功",
                        f"已生成 {len(data_list)} 張標籤：\n{result}\n\n是否立即開啟？",
                        parent=parent):
                    os.startfile(str(result))
            except Exception as e:
                messagebox.showerror("生成失敗", str(e), parent=parent)

        gf = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        gf.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(gf, text="🖨  生成標籤 PDF", command=_generate,
                       fg_color="#1e8449", hover_color="#176a3a", text_color="white",
                       font=("Microsoft JhengHei UI", 12, "bold"),
                       height=44, corner_radius=8).pack(fill="x")
