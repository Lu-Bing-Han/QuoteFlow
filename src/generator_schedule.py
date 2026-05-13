"""
generator_schedule.py — 從 Timetree 行事曆抓取出貨排程，生成 Excel 工作表
"""
import re
from datetime import datetime, timezone, timedelta, date as date_type
from pathlib import Path

import requests
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from _paths import TEMPLATE_DIR, OUTPUT_DIR

_RED    = "FFFF0000"
_THIN   = Side(style="thin")
_THICK  = Side(style="thick")
_ALL_BORDER   = Border(left=_THIN,  right=_THIN,  top=_THIN,  bottom=_THIN)
_THICK_BORDER = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)
_CENTER = Alignment(horizontal="center", vertical="center")

CALENDAR_ID = 25024642
_API_URL    = f"https://timetreeapp.com/api/v1/calendar/{CALENDAR_ID}/events"
_TEMPLATE   = TEMPLATE_DIR / "template_schedule.xlsx"
_TW_TZ      = timezone(timedelta(hours=8))


def _extract_company(title: str) -> str:
    """Return company name, stripping 【...】, (N), (拉回) leading prefixes."""
    t = re.sub(r'【[^】]*】', '', title).strip()
    # strip leading (N) and (拉回) tokens, e.g. "(3)(拉回)"
    t = re.sub(r'^(\(\d+\)\s*)*(\(拉回\)\s*)?', '', t).strip()
    m = re.match(r'^(.+?)(?=\()', t)
    if m:
        return m.group(1).strip(' -')
    return t.split()[0] if t else title


def _location_cell(ev: dict) -> str:
    company = _extract_company(ev.get("title", ""))
    loc = (ev.get("location") or "").strip()
    if loc:
        return f"{company}({loc})"
    # use area from title parentheses if present
    m = re.search(r'\(([^)]+)\)', ev.get("title", ""))
    area = m.group(1) if m else ""
    return f"{company}({area})" if area else company


def _note_cell(seq: int, ev: dict) -> str:
    company = _extract_company(ev.get("title", ""))
    t = re.sub(r'【[^】]*】', '', ev.get("title", "")).strip()
    # strip leading company name and any (area) that follows
    t = re.sub(r'^' + re.escape(company) + r'\s*(\([^)]*\))?\s*', '', t).strip()
    t = re.sub(r'^[-\s]+', '', t).strip()
    return f"{seq}.{company}-{t}" if t else f"{seq}.{company}"


def fetch_events(target: date_type, session_id: str, csrf_token: str) -> list:
    """Return active Timetree events on target date (Taiwan time)."""
    start_tw = datetime(target.year, target.month, target.day, tzinfo=_TW_TZ)
    # Use 60-day lookback so events created weeks ago are still returned
    since_ms = int((start_tw - timedelta(days=60)).timestamp() * 1000)

    headers = {
        "accept": "application/json, text/plain, */*",
        "content-type": "application/json",
        "referer": "https://timetreeapp.com/",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "x-csrf-token": csrf_token,
        "x-timetreea": "web/2.1.0/zh-TW",
        "cookie": f"_session_id={session_id}",
    }

    resp = requests.get(_API_URL, params={"since": since_ms}, headers=headers, timeout=15)
    resp.raise_for_status()

    result = []
    for ev in resp.json().get("events", []):
        if ev.get("deactivated_at"):
            continue
        dt = datetime.fromtimestamp(ev.get("start_at", 0) / 1000, tz=_TW_TZ)
        if dt.date() == target:
            result.append({**ev, "_dt": dt})

    # all-day events first, then by start time
    result.sort(key=lambda e: (not e.get("all_day"), e["_dt"]))
    return result


def generate_schedule(target: date_type, output_dir: Path,
                      session_id: str, csrf_token: str) -> Path:
    events = fetch_events(target, session_id, csrf_token)

    wb = openpyxl.load_workbook(str(_TEMPLATE))

    roc_year = target.year - 1911
    sheet_name = f"{roc_year}年{target.month}月{target.day}日"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    def _merge_ab_cf(r, center_c=False):
        ws.merge_cells(f"A{r}:B{r}")
        ws.cell(row=r, column=1).alignment = _CENTER
        ws.merge_cells(f"C{r}:F{r}")
        if center_c:
            ws.cell(row=r, column=3).alignment = _CENTER

    def _merge_gh_border(r):
        ws.merge_cells(f"G{r}:H{r}")
        for col in (7, 8):
            ws.cell(row=r, column=col).border = _THICK_BORDER

    def _table_border(start_row, end_row):
        """Apply thin all-border to the event table A:H."""
        for r in range(start_row, end_row + 1):
            for col in range(1, 9):
                ws.cell(row=r, column=col).border = _ALL_BORDER

    # Column widths
    ws.column_dimensions["F"].width = 40
    ws.column_dimensions["G"].width = 11.5
    ws.column_dimensions["H"].width = 40

    # ── 上方事件表格 ─────────────────────────────────────────
    table_end = 1 + len(events)   # row 1 = header, rows 2..N+1 = events

    # Header row
    ws["A1"] = f"{target.month:02d}/{target.day:02d}行程順序"
    ws["C1"] = "地點"
    ws["G1"] = "Google時間"
    ws.cell(row=1, column=8, value="備註").alignment = _CENTER
    _merge_ab_cf(1, center_c=True)

    # Event rows
    for i, ev in enumerate(events, 1):
        r = i + 1
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=3, value=_location_cell(ev))
        _merge_ab_cf(r)  # C left-aligned

    # All-border on event table (rows 1..table_end, cols A:H)
    _table_border(1, table_end)

    # ── 下方備註列（每筆之間空一列）────────────────────────────
    note_start = table_end + 2   # one blank row gap after table
    for i, ev in enumerate(events, 1):
        r = note_start + (i - 1) * 2   # skip one row between each entry
        ws.cell(row=r, column=1, value=_note_cell(i, ev))
        if "拉回" in ev.get("title", ""):
            # 拉回事件：純文字，不合併、不加框線、不加物流士確認
            continue
        ws.merge_cells(f"A{r}:F{r}")
        for col in range(1, 7):
            ws.cell(row=r, column=col).border = _THICK_BORDER
        ws.cell(row=r, column=7, value="物流士確認          □")
        _merge_gh_border(r)

    # ── 頁尾 ─────────────────────────────────────────────────
    footer_r = note_start + len(events) * 2

    # 要給司機零用金+手機 — red
    ws.cell(row=footer_r, column=7, value="要給司機零用金+手機").font = Font(color=_RED)

    # 物流士確認 — merged G:H, thick border
    ws.cell(row=footer_r + 1, column=7, value="物流士確認")
    _merge_gh_border(footer_r + 1)

    # 物流士確認 主管覆核(蓋章) — red text, black □, merged G:H, thick border
    ws.cell(row=footer_r + 2, column=7).value = CellRichText(
        TextBlock(InlineFont(color=_RED), "物流士確認          主管覆核(蓋章) "),
        TextBlock(InlineFont(color="FF000000"), "□"),
    )
    _merge_gh_border(footer_r + 2)

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"{target.strftime('%Y-%m-%d')}.xlsx"
    wb.save(str(out_path))
    wb.close()
    return out_path
