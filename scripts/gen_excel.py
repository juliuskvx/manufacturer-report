"""
gen_excel.py — EuroAir Intel Manufacturer Report
Maintains a persistent historical archive: data/historical_data.xlsx
Each run appends one row of data. The file grows month by month.

Sheets:
  1. Historical Data  — one row per month, all key metrics
  2. Monthly Deliveries — time series, auto-extending
  3. Model Breakdown   — latest month detail
  4. Backlog Analysis  — latest month detail
  5. Stock Data        — latest month + running stock history
"""

import json
import os
from datetime import datetime

try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, LineChart, Reference
except ImportError:
    os.system("pip install openpyxl --break-system-packages -q")
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, LineChart, Reference

with open("output/manufacturer_data.json") as f:
    d = json.load(f)

B  = d["boeing"]
A  = d["airbus"]
Bs = B.get("stock", {})
As = A.get("stock", {})

ARCHIVE_PATH = "data/historical_data.xlsx"
os.makedirs("data", exist_ok=True)

# ── Style helpers ──────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)

def font(bold=False, size=11, color="1E293B", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")

def border_thin():
    s = Side(style="thin", color="E2E8F0")
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def style_cell(c, bold=False, size=10, color="1E293B", bg=None, h_align="left", italic=False):
    c.font = font(bold=bold, size=size, color=color, italic=italic)
    if bg: c.fill = fill(bg)
    c.alignment = align(h_align)
    c.border = border_thin()

NAVY   = "0D1F3C"
BOEING = "0033A0"
AIRBUS = "005B8E"
GOLD   = "C9A84C"
OFFWHITE = "F8FAFC"
LGRAY  = "E2E8F0"
SLATE  = "64748B"
WHITE  = "FFFFFF"
GREEN  = "16A34A"
RED    = "DC2626"
BLIGHT = "D6E4FF"
ALIGHT = "D0EEFF"

# ── Historical headers ─────────────────────────────────────────────────────────
HIST_HEADERS = [
    "Report Month", "Report Date",
    "Boeing Deliveries YTD", "Boeing Orders YTD", "Boeing Backlog", "Boeing Backlog Years",
    "Boeing 737MAX YTD", "Boeing 787 YTD", "Boeing 777 YTD", "Boeing 767 YTD",
    "Airbus Deliveries YTD", "Airbus Orders YTD", "Airbus Backlog", "Airbus Backlog Years",
    "Airbus A320neo YTD", "Airbus A350 YTD", "Airbus A220 YTD", "Airbus A330 YTD",
    "BA Price (USD)", "BA Change %", "BA 52w High", "BA 52w Low",
    "AIR.PA Price (EUR)", "AIR.PA Change %", "AIR.PA 52w High", "AIR.PA 52w Low",
    "Data Source", "Notes",
]

def new_data_row():
    return [
        d["reportMonth"],
        d["reportDate"],
        B["deliveries_ytd"],
        B["orders_ytd"],
        B["backlog"],
        B["backlog_years"],
        B["models"].get("737 MAX", 0),
        B["models"].get("787", 0),
        B["models"].get("777", 0),
        B["models"].get("767", 0),
        A["deliveries_ytd"],
        A["orders_ytd"],
        A["backlog"],
        A["backlog_years"],
        A["models"].get("A320neo", 0),
        A["models"].get("A350", 0),
        A["models"].get("A220", 0),
        A["models"].get("A330", 0),
        Bs.get("price", 0),
        Bs.get("change_pct", 0),
        Bs.get("week52_high", 0),
        Bs.get("week52_low", 0),
        As.get("price", 0),
        As.get("change_pct", 0),
        As.get("week52_high", 0),
        As.get("week52_low", 0),
        d.get("dataSource", "fallback_verified"),
        "",
    ]

# ── Load or create workbook ────────────────────────────────────────────────────
if os.path.exists(ARCHIVE_PATH):
    print(f"  [excel] loading existing archive: {ARCHIVE_PATH}")
    wb = load_workbook(ARCHIVE_PATH)
else:
    print(f"  [excel] creating new archive: {ARCHIVE_PATH}")
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Historical Data (append-only)
# ═══════════════════════════════════════════════════════════════════════════════
if "Historical Data" in wb.sheetnames:
    ws1 = wb["Historical Data"]
    # Check if this month already exists — don't duplicate
    existing_months = [ws1.cell(row=r, column=1).value for r in range(2, ws1.max_row + 1)]
    if d["reportMonth"] in existing_months:
        print(f"  [excel] {d['reportMonth']} already in archive — updating row")
        for r in range(2, ws1.max_row + 1):
            if ws1.cell(row=r, column=1).value == d["reportMonth"]:
                new_row = new_data_row()
                for ci, val in enumerate(new_row, 1):
                    c = ws1.cell(row=r, column=ci, value=val)
                    style_cell(c, h_align="center", bg=OFFWHITE if r % 2 == 0 else WHITE)
                break
    else:
        print(f"  [excel] appending new row for {d['reportMonth']}")
        next_row = ws1.max_row + 1
        new_row = new_data_row()
        for ci, val in enumerate(new_row, 1):
            c = ws1.cell(row=next_row, column=ci, value=val)
            bg = OFFWHITE if next_row % 2 == 0 else WHITE
            # Color Boeing cols blue, Airbus cols teal, stock cols slate
            col_color = "1E293B"
            if 3 <= ci <= 10:   col_color = BOEING
            elif 11 <= ci <= 18: col_color = AIRBUS
            elif 19 <= ci <= 22: col_color = "1A5276"
            elif 23 <= ci <= 26: col_color = "1A5276"
            style_cell(c, h_align="center", bg=bg, color=col_color)
else:
    ws1 = wb.create_sheet("Historical Data", 0)
    ws1.sheet_view.showGridLines = False

    # Title
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HIST_HEADERS))
    tc = ws1.cell(row=1, column=1, value="EuroAir Intel — Manufacturer Historical Archive")
    tc.font = font(bold=True, size=16, color=WHITE)
    tc.fill = fill(NAVY)
    tc.alignment = align("center")
    ws1.row_dimensions[1].height = 36

    # Header row
    ws1.row_dimensions[2].height = 28
    col_groups = {
        (1,2): NAVY,
        (3,10): BOEING,
        (11,18): AIRBUS,
        (19,22): "1A3A5C",
        (23,26): "0A3D5C",
        (27,28): SLATE,
    }
    for ci, h in enumerate(HIST_HEADERS, 1):
        c = ws1.cell(row=2, column=ci, value=h)
        bg = NAVY
        for (lo, hi), col in col_groups.items():
            if lo <= ci <= hi:
                bg = col
                break
        c.font = font(bold=True, size=9, color=WHITE)
        c.fill = fill(bg)
        c.alignment = align("center", wrap=True)
        c.border = border_thin()

    # First data row
    new_row_data = new_data_row()
    ws1.row_dimensions[3].height = 20
    for ci, val in enumerate(new_row_data, 1):
        c = ws1.cell(row=3, column=ci, value=val)
        col_color = "1E293B"
        if 3 <= ci <= 10:    col_color = BOEING
        elif 11 <= ci <= 18: col_color = AIRBUS
        style_cell(c, h_align="center", bg=OFFWHITE, color=col_color)

    # Column widths
    col_widths = [14, 12] + [14]*16 + [12]*8 + [16, 18]
    for ci, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 2: Monthly Deliveries (this month's detail + running YTD chart)
# ═══════════════════════════════════════════════════════════════════════════════
sheet_name2 = f"Deliveries {d['reportMonth']}"
if sheet_name2 in wb.sheetnames:
    del wb[sheet_name2]
ws2 = wb.create_sheet(sheet_name2)
ws2.sheet_view.showGridLines = False

for col, w in enumerate([14, 16, 16, 16, 16], 1):
    ws2.column_dimensions[get_column_letter(col)].width = w

# Title
ws2.merge_cells("A1:E2")
t = ws2.cell(row=1, column=1, value=f"Monthly Deliveries 2026 YTD — {d['reportMonth']}")
t.font = font(bold=True, size=15, color=WHITE)
t.fill = fill(NAVY)
t.alignment = align("center")
ws2.row_dimensions[1].height = 32

# Headers
headers2 = ["Month", "Boeing", "Airbus", "Combined", "Boeing Lead"]
hcols = [NAVY, BOEING, AIRBUS, SLATE, NAVY]
for ci, (h, hc) in enumerate(zip(headers2, hcols), 1):
    c = ws2.cell(row=3, column=ci, value=h)
    c.font = font(bold=True, size=10, color=WHITE)
    c.fill = fill(hc)
    c.alignment = align("center")
    c.border = border_thin()

months    = d.get("chartLabels", [])
b_series  = d.get("boeingMonthlySeries", [])
a_series  = d.get("airbusMonthlySeries", [])

for i, (m, bv, av) in enumerate(zip(months, b_series, a_series), 4):
    ws2.row_dimensions[i].height = 20
    bg = OFFWHITE if i % 2 == 0 else WHITE
    lead = bv - av
    for ci, (val, col) in enumerate(zip(
        [m, bv, av, bv+av, lead],
        ["1E293B", BOEING, AIRBUS, SLATE, GREEN if lead >= 0 else RED]
    ), 1):
        c = ws2.cell(row=i, column=ci, value=val)
        style_cell(c, h_align="center", bg=bg, color=col,
                   bold=(ci in [2,3] and val > 0))

# Totals
tr = len(months) + 4
ws2.row_dimensions[tr].height = 24
totals = [sum(b_series), sum(a_series)]
for ci, (val, col) in enumerate(zip(
    ["TOTAL YTD", totals[0], totals[1], sum(totals), totals[0]-totals[1]],
    [WHITE, "90CAFF", "90E0FF", WHITE, "90FFB0" if totals[0]>=totals[1] else "FFB0B0"]
), 1):
    c = ws2.cell(row=tr, column=ci, value=val)
    style_cell(c, bold=True, h_align="center", bg=NAVY, color=col)

# Chart
chart2 = BarChart()
chart2.type = "col"
chart2.grouping = "clustered"
chart2.title = f"Monthly Deliveries — {d['reportMonth']}"
chart2.y_axis.title = "Aircraft Delivered"
chart2.style = 10
chart2.width = 22
chart2.height = 14
data_r  = Reference(ws2, min_col=2, max_col=3, min_row=3, max_row=3+len(months))
cats_r  = Reference(ws2, min_col=1, min_row=4, max_row=3+len(months))
chart2.add_data(data_r, titles_from_data=True)
chart2.set_categories(cats_r)
chart2.series[0].graphicalProperties.solidFill = BOEING
chart2.series[1].graphicalProperties.solidFill = "00AEEF"
ws2.add_chart(chart2, "G3")

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 3: Latest Month Summary
# ═══════════════════════════════════════════════════════════════════════════════
sheet_name3 = f"Summary {d['reportMonth']}"
if sheet_name3 in wb.sheetnames:
    del wb[sheet_name3]
ws3 = wb.create_sheet(sheet_name3)
ws3.sheet_view.showGridLines = False

for col, w in enumerate([24, 18, 18], 1):
    ws3.column_dimensions[get_column_letter(col)].width = w

ws3.merge_cells("A1:C2")
t3 = ws3.cell(row=1, column=1, value=f"EuroAir Intel — {d['reportMonth']} Summary")
t3.font = font(bold=True, size=15, color=WHITE)
t3.fill = fill(NAVY)
t3.alignment = align("center")
ws3.row_dimensions[1].height = 32

# Headers
for ci, (h, hc) in enumerate(zip(["Metric", "Boeing", "Airbus"], [NAVY, BOEING, AIRBUS]), 1):
    c = ws3.cell(row=3, column=ci, value=h)
    c.font = font(bold=True, size=10, color=WHITE)
    c.fill = fill(hc)
    c.alignment = align("center")
    c.border = border_thin()

summary_rows = [
    ("Deliveries YTD",       B["deliveries_ytd"],               A["deliveries_ytd"]),
    ("Gross Orders YTD",     B["orders_ytd"],                   A["orders_ytd"]),
    ("Total Backlog (ac)",   B["backlog"],                      A["backlog"]),
    ("Backlog Runway (yrs)", B["backlog_years"],                A["backlog_years"]),
    ("Narrowbody Deliv.",    B["models"].get("737 MAX", 0),     A["models"].get("A320neo", 0)),
    ("Widebody Deliv.",      B["models"].get("787",0)+B["models"].get("777",0), A["models"].get("A350",0)+A["models"].get("A330",0)),
    ("Stock Price",          f"${Bs.get('price','N/A')} USD",   f"€{As.get('price','N/A')} EUR"),
    ("Stock Change Today",   f"{float(Bs.get('change_pct',0)):+.2f}%", f"{float(As.get('change_pct',0)):+.2f}%"),
    ("52-Week High",         f"${Bs.get('week52_high','N/A')}", f"€{As.get('week52_high','N/A')}"),
    ("52-Week Low",          f"${Bs.get('week52_low','N/A')}",  f"€{As.get('week52_low','N/A')}"),
]

for i, (metric, bv, av) in enumerate(summary_rows, 4):
    ws3.row_dimensions[i].height = 22
    bg = OFFWHITE if i % 2 == 0 else WHITE
    for ci, (val, col) in enumerate(zip([metric, bv, av], ["1E293B", BOEING, AIRBUS]), 1):
        c = ws3.cell(row=i, column=ci, value=val)
        style_cell(c, h_align="center" if ci > 1 else "left", bg=bg, color=col, bold=ci > 1)

# Source note
src_row = len(summary_rows) + 5
ws3.merge_cells(f"A{src_row}:C{src_row}")
src = ws3.cell(row=src_row, column=1, value=f"Source: Boeing.com · Airbus.com · Yahoo Finance · Generated {d['reportDate']}")
src.font = font(size=8, color=SLATE, italic=True)
src.fill = fill(OFFWHITE)
src.alignment = align("left")

# ── Save ───────────────────────────────────────────────────────────────────────
wb.save(ARCHIVE_PATH)
print(f"✓ Saved archive: {ARCHIVE_PATH}")

# Also save a dated copy to output/ for email attachment
dated_copy = f"output/Manufacturer_Data_{d['reportDate']}.xlsx"
import shutil
shutil.copy(ARCHIVE_PATH, dated_copy)
print(f"✓ Copied to: {dated_copy}")
