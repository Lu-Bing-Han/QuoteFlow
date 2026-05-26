"""
generator_label.py — 批量生成產品標籤 PDF
將 {{型號}} / {{荷重}} / {{製造序號}} 替換後輸出多頁 PDF
"""
from pathlib import Path

import fitz  # pymupdf

from _paths import TEMPLATE_DIR

TEMPLATES = {
    "銀標":    TEMPLATE_DIR / "template_logo.pdf",
    "APT標":   TEMPLATE_DIR / "template_APT.pdf",
    "無公司標": TEMPLATE_DIR / "template.pdf",
    "上銀標":   TEMPLATE_DIR / "template_silver.pdf",
}
_FONT  = r"C:\Windows\Fonts\msjhbd.ttc"
_WHITE = (1, 1, 1)
_BLACK = (0, 0, 0)

_font_obj = None

def _get_font():
    global _font_obj
    if _font_obj is None:
        _font_obj = fitz.Font(fontfile=_FONT)
    return _font_obj


def _find_placeholder_rects(page, placeholder: str):
    """
    Return a list of fitz.Rect bounding the full placeholder text.

    Works for both normal PDFs and Illustrator-exported PDFs where each
    character is stored as a separate span.  Falls back to search_for first
    (fast path), then reconstructs from individual chars (slow path).
    """
    # Fast path — works when text is stored contiguously
    rects = page.search_for(placeholder)
    if rects:
        return rects

    # Slow path — reconstruct from per-character spans
    blocks = page.get_text("rawdict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
    results = []

    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            # Collect all chars in the line as (char, bbox) pairs
            chars = []
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    chars.append((ch["c"], fitz.Rect(ch["bbox"])))

            # Slide a window of len(placeholder) over chars
            n = len(placeholder)
            for i in range(len(chars) - n + 1):
                window = chars[i:i + n]
                text = "".join(c for c, _ in window)
                if text == placeholder:
                    # Union all bboxes in the window
                    r = window[0][1]
                    for _, bbox in window[1:]:
                        r |= bbox
                    results.append(r)

    return results


def generate_labels(data_list: list, output_path: Path,
                    template_key: str = "銀標") -> Path:
    """
    data_list    : [{"型號": "...", "荷重": "...", "製造序號": "..."}, ...]
    output_path  : 輸出 PDF 完整路徑
    template_key : "銀標" / "無公司標" / "APT標"
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
            "{{序號}}":    data.get("序號", "") or data.get("機台序號", ""),
            "{{尺寸}}":    data.get("機台尺寸", ""),
            "{{重量}}":    data.get("機台重量", ""),
            "{{機台尺寸}}": data.get("機台尺寸", ""),
            "{{機台重量}}": data.get("機台重量", ""),
            "{{出廠年份}}": data.get("出廠年份", ""),
            "{{年分}}":    data.get("出廠年份", ""),
            "{{代碼}}":   data.get("供應商代碼", ""),
            "{{機台序號}}": data.get("機台序號", ""),
            "{{編號}}":   data.get("訂單編號", ""),
            "{{收貨人}}":  data.get("收貨人", ""),
        }

        tw = fitz.TextWriter(page.rect)
        for placeholder, value in replacements.items():
            if not value:
                continue
            for rect in _find_placeholder_rects(page, placeholder):
                page.draw_rect(rect, color=None, fill=_WHITE)
                fs = max(rect.height * 0.75, 6)
                tw.append(
                    (rect.x0 + 2, rect.y1 - 2),
                    value,
                    font=_get_font(),
                    fontsize=fs,
                )
        tw.write_text(page, color=_BLACK)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    out.save(str(output_path))
    template.close()
    out.close()
    return output_path
