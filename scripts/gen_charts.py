"""
gen_charts.py — EuroAir Intel Manufacturer Report
Generates 4 PNG charts:
  1. Monthly deliveries comparison (grouped bar, Boeing vs Airbus)
  2. YTD deliveries by model (stacked bar, side by side)
  3. Backlog comparison (horizontal bar)
  4. Stock price 30-day trend (line chart, dual axis)
"""

import json
import os
import sys

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import numpy as np
except ImportError:
    os.system("pip install matplotlib numpy --break-system-packages -q")
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import numpy as np

try:
    import yfinance as yf
except ImportError:
    os.system("pip install yfinance --break-system-packages -q")
    import yfinance as yf

# ── Palette (matches EuroAir Intel PPTX + Boeing/Airbus brand colors) ─────────
BOEING_BLUE   = "#0033A0"   # Boeing brand blue
AIRBUS_TEAL   = "#00205B"   # Airbus navy (their brand)
AIRBUS_LIGHT  = "#00AEEF"   # Airbus accent cyan
BOEING_LIGHT  = "#4B9CD3"   # Boeing lighter blue
GOLD          = "#C9A84C"   # EuroAir accent
BG            = "#F8FAFC"   # slide background
GRID_CLR      = "#E2E8F0"
TEXT_CLR      = "#1E293B"
SLATE         = "#64748B"

FONT_TITLE  = {"fontsize": 13, "fontweight": "bold", "color": TEXT_CLR, "fontfamily": "DejaVu Sans"}
FONT_AXIS   = {"fontsize": 9,  "color": SLATE,        "fontfamily": "DejaVu Sans"}
FONT_LABEL  = {"fontsize": 8,  "color": TEXT_CLR,     "fontfamily": "DejaVu Sans"}

def apply_style(ax, title, xlabel="", ylabel=""):
    ax.set_facecolor(BG)
    ax.figure.patch.set_facecolor(BG)
    ax.set_title(title, pad=10, **FONT_TITLE)
    if xlabel: ax.set_xlabel(xlabel, **FONT_AXIS)
    if ylabel: ax.set_ylabel(ylabel, **FONT_AXIS)
    ax.tick_params(colors=SLATE, labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_CLR)
    ax.spines["bottom"].set_color(GRID_CLR)
    ax.yaxis.grid(True, color=GRID_CLR, linewidth=0.7, zorder=0)
    ax.set_axisbelow(True)

def save(fig, path):
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print(f"  ✓ {path}")

def chart_monthly_deliveries(d, outdir):
    """Grouped bar: monthly deliveries Boeing vs Airbus YTD."""
    labels  = d["chartLabels"]
    boeing  = d["boeingMonthlySeries"]
    airbus  = d["airbusMonthlySeries"]
    n = len(labels)
    x = np.arange(n)
    w = 0.38

    fig, ax = plt.subplots(figsize=(10, 4.2))
    bars_b = ax.bar(x - w/2, boeing, w, label="Boeing",  color=BOEING_BLUE, zorder=3, linewidth=0)
    bars_a = ax.bar(x + w/2, airbus, w, label="Airbus",  color=AIRBUS_LIGHT, zorder=3, linewidth=0)

    # Value labels on bars
    for bar in bars_b:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + 1.5, str(int(h)),
                    ha="center", va="bottom", fontsize=7.5, color=BOEING_BLUE, fontweight="bold")
    for bar in bars_a:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + 1.5, str(int(h)),
                    ha="center", va="bottom", fontsize=7.5, color="#005B8E", fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(labels, **FONT_LABEL)
    apply_style(ax, "Monthly Deliveries 2026 YTD — Boeing vs Airbus", ylabel="Aircraft Delivered")
    ax.legend(loc="upper left", fontsize=9, framealpha=0.8,
              handles=[mpatches.Patch(color=BOEING_BLUE, label="Boeing"),
                       mpatches.Patch(color=AIRBUS_LIGHT, label="Airbus")])
    fig.tight_layout()
    save(fig, os.path.join(outdir, "chart_monthly.png"))

def chart_model_breakdown(d, outdir):
    """Stacked bar: deliveries by model for Boeing and Airbus side by side."""
    b = d["boeing"]["models"]
    a = d["airbus"]["models"]

    # Boeing models
    b_labels = list(b.keys())
    b_vals   = [b[k] for k in b_labels]
    b_colors = ["#0033A0", "#4B9CD3", "#7BB3D9", "#B0CCEC"]

    # Airbus models
    a_labels = list(a.keys())
    a_vals   = [a[k] for k in a_labels]
    a_colors = ["#00AEEF", "#00205B", "#005B8E", "#7BCAE1"]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.2))

    # Boeing donut-style bar
    wedges_b = ax1.pie(b_vals, labels=None, colors=b_colors, startangle=90,
                       wedgeprops={"linewidth": 2, "edgecolor": BG},
                       pctdistance=0.75)
    ax1.set_facecolor(BG)
    ax1.figure.patch.set_facecolor(BG)
    # Add centre circle for donut effect
    centre = plt.Circle((0, 0), 0.52, fc=BG)
    ax1.add_patch(centre)
    total_b = sum(b_vals)
    ax1.text(0, 0.06, str(total_b), ha="center", va="center",
             fontsize=22, fontweight="bold", color=BOEING_BLUE)
    ax1.text(0, -0.18, "delivered", ha="center", va="center",
             fontsize=8, color=SLATE)
    ax1.set_title("Boeing — Deliveries by Model\n(2026 YTD)", pad=8, **FONT_TITLE)
    # Legend
    legend_b = [mpatches.Patch(color=b_colors[i], label=f"{b_labels[i]}: {b_vals[i]}") for i in range(len(b_labels))]
    ax1.legend(handles=legend_b, loc="lower center", bbox_to_anchor=(0.5, -0.12),
               ncol=2, fontsize=8, framealpha=0.8)

    # Airbus donut
    wedges_a = ax2.pie(a_vals, labels=None, colors=a_colors, startangle=90,
                       wedgeprops={"linewidth": 2, "edgecolor": BG},
                       pctdistance=0.75)
    ax2.set_facecolor(BG)
    centre2 = plt.Circle((0, 0), 0.52, fc=BG)
    ax2.add_patch(centre2)
    total_a = sum(a_vals)
    ax2.text(0, 0.06, str(total_a), ha="center", va="center",
             fontsize=22, fontweight="bold", color="#005B8E")
    ax2.text(0, -0.18, "delivered", ha="center", va="center",
             fontsize=8, color=SLATE)
    ax2.set_title("Airbus — Deliveries by Model\n(2026 YTD)", pad=8, **FONT_TITLE)
    legend_a = [mpatches.Patch(color=a_colors[i], label=f"{a_labels[i]}: {a_vals[i]}") for i in range(len(a_labels))]
    ax2.legend(handles=legend_a, loc="lower center", bbox_to_anchor=(0.5, -0.12),
               ncol=2, fontsize=8, framealpha=0.8)

    fig.patch.set_facecolor(BG)
    fig.tight_layout(pad=1.5)
    save(fig, os.path.join(outdir, "chart_models.png"))

def chart_backlog(d, outdir):
    """Horizontal grouped bar: backlog by model."""
    b_backlog = d["boeing"]["backlog_by_model"]
    a_backlog = d["airbus"]["backlog_by_model"]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.2))

    # Boeing
    b_labels = list(b_backlog.keys())
    b_vals   = [b_backlog[k] for k in b_labels]
    b_cols   = ["#0033A0", "#4B9CD3", "#7BB3D9", "#B0CCEC"]
    bars = ax1.barh(b_labels, b_vals, color=b_cols, zorder=3, linewidth=0, height=0.5)
    for bar, val in zip(bars, b_vals):
        if val > 0:
            ax1.text(val + 20, bar.get_y() + bar.get_height()/2,
                     f"{val:,}", va="center", fontsize=8, color=BOEING_BLUE, fontweight="bold")
    apply_style(ax1, f"Boeing Backlog by Model\n(Total: {d['boeing']['backlog']:,} aircraft)", ylabel="")
    ax1.set_xlabel("Aircraft on Order", **FONT_AXIS)
    ax1.invert_yaxis()

    # Airbus
    a_labels = list(a_backlog.keys())
    a_vals   = [a_backlog[k] for k in a_labels]
    a_cols   = ["#00AEEF", "#00205B", "#005B8E", "#7BCAE1"]
    bars2 = ax2.barh(a_labels, a_vals, color=a_cols, zorder=3, linewidth=0, height=0.5)
    for bar, val in zip(bars2, a_vals):
        if val > 0:
            ax2.text(val + 20, bar.get_y() + bar.get_height()/2,
                     f"{val:,}", va="center", fontsize=8, color="#005B8E", fontweight="bold")
    apply_style(ax2, f"Airbus Backlog by Model\n(Total: {d['airbus']['backlog']:,} aircraft)", ylabel="")
    ax2.set_xlabel("Aircraft on Order", **FONT_AXIS)
    ax2.invert_yaxis()

    fig.patch.set_facecolor(BG)
    fig.tight_layout(pad=1.5)
    save(fig, os.path.join(outdir, "chart_backlog.png"))

def chart_stock(d, outdir):
    """Dual-axis line chart: BA (USD) and AIR.PA (EUR) 30-day price trend."""
    print("  [chart_stock] fetching 30-day price history...")

    fig, ax1 = plt.subplots(figsize=(10, 4.2))
    ax2 = ax1.twinx()

    try:
        ba  = yf.Ticker("BA").history(period="1mo")
        air = yf.Ticker("AIR.PA").history(period="1mo")

        if not ba.empty:
            dates_ba = [dt.strftime("%d %b") for dt in ba.index]
            ax1.plot(range(len(ba)), ba["Close"].values,
                     color=BOEING_BLUE, linewidth=2, label="BA (USD)", zorder=3)
            ax1.fill_between(range(len(ba)), ba["Close"].values,
                             alpha=0.08, color=BOEING_BLUE)

        if not air.empty:
            dates_air = [dt.strftime("%d %b") for dt in air.index]
            ax2.plot(range(len(air)), air["Close"].values,
                     color=AIRBUS_LIGHT, linewidth=2, linestyle="--",
                     label="AIR.PA (EUR)", zorder=3)
            ax2.fill_between(range(len(air)), air["Close"].values,
                             alpha=0.06, color=AIRBUS_LIGHT)

        # X-axis: show ~6 date labels evenly spaced
        n = max(len(ba), len(air)) if not ba.empty or not air.empty else 20
        tick_positions = list(range(0, n, max(1, n // 6)))
        all_dates = dates_ba if not ba.empty else dates_air
        ax1.set_xticks(tick_positions)
        ax1.set_xticklabels([all_dates[i] for i in tick_positions if i < len(all_dates)],
                            rotation=0, **FONT_LABEL)

    except Exception as e:
        print(f"  [chart_stock] live data failed: {e} — using placeholder")
        # Draw placeholder bars with known prices
        ba_price  = d["boeing"]["stock"].get("price", 229)
        air_price = d["airbus"]["stock"].get("price", 148)
        ax1.bar([0], [ba_price],  color=BOEING_BLUE,  label="BA (USD)")
        ax2.bar([1], [air_price], color=AIRBUS_LIGHT, label="AIR.PA (EUR)")

    apply_style(ax1, "Stock Price — Boeing (BA) vs Airbus (AIR.PA) · 30 Days")
    ax1.set_ylabel("BA Price (USD)", color=BOEING_BLUE, fontsize=9)
    ax2.set_ylabel("AIR.PA Price (EUR)", color=AIRBUS_LIGHT, fontsize=9)
    ax1.tick_params(axis="y", labelcolor=BOEING_BLUE, labelsize=8)
    ax2.tick_params(axis="y", labelcolor=AIRBUS_LIGHT, labelsize=8)
    ax2.spines["right"].set_visible(True)
    ax2.spines["right"].set_color(GRID_CLR)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left",
               fontsize=9, framealpha=0.8)

    fig.patch.set_facecolor(BG)
    fig.tight_layout()
    save(fig, os.path.join(outdir, "chart_stock.png"))

def main():
    with open("output/manufacturer_data.json") as f:
        d = json.load(f)

    outdir = "output"
    os.makedirs(outdir, exist_ok=True)

    print("\n[1/4] Monthly deliveries chart...")
    chart_monthly_deliveries(d, outdir)

    print("[2/4] Model breakdown chart...")
    chart_model_breakdown(d, outdir)

    print("[3/4] Backlog chart...")
    chart_backlog(d, outdir)

    print("[4/4] Stock price chart...")
    chart_stock(d, outdir)

    print("\n✓ All charts generated")

if __name__ == "__main__":
    main()
