"""
generator_label.py — 批量生成產品標籤 PDF
將 {{型號}} / {{荷重}} / {{製造序號}} 替換後輸出多頁 PDF
"""
from pathlib import Path
from datetime import datetime

import fitz  # pymupdf

from _paths import TEMPLATE_DIR

TEMPLATES = {
    "銀標":    TEMPLATE_DIR / "template_logo.pdf",
    "無公司標": TEMPLATE_DIR / "template.pdf",
}
_FONT  = r"C:\Windows\Fonts\msjhbd.ttc"
_WHITE = (1, 1, 1)
_BLACK = (0, 0, 0)


def generate_labels(data_list: list, output_path: Path,
                    template_key: str = "銀標") -> Path:
    """
    data_list    : [{"型號": "...", "荷重": "...", "製造序號": "..."}, ...]
    output_path  : 輸出 PDF 完整路徑
    template_key : "銀標" 或 "無公司標"
    """
    tpl_path = TEMPLATES.get(template_key, TEMPLATES["銀標"])
    template = fitz.open(str(tpl_path))
    out      = fitz.open()

    for data in data_list:
        out.insert_pdf(template, from_page=0, to_page=0)
        page = out[-1]

        replacements = {
            "{{型號}}":    data.get("型號", ""),
            "{{荷重}}":    data.get("荷重", ""),
            "{{製造序號}}": data.get("製造序號", ""),
        }

        for placeholder, value in replacements.items():
            if not value:
                continue
            for rect in page.search_for(placeholder):
                page.draw_rect(rect, color=None, fill=_WHITE)
                page.insert_text(
                    (rect.x0 + 2, rect.y1 - 2),
                    value,
                    fontsize=rect.height * 0.75,
                    fontfile=_FONT,
                    color=_BLACK,
                )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    out.save(str(output_path))
    template.close()
    out.close()
    return output_path
