"""
Task 3: Customer Spending Report
RaushByte Technologies – Data Analytics Internship
Objective: Identify high-value customers based on spending behavior.
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
import os

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE    = os.path.dirname(os.path.abspath(__file__))
DATA    = os.path.join(BASE, "data",   "customer_analytics_dataset.xlsx")
CHARTS  = os.path.join(BASE, "charts")
OUTPUT  = os.path.join(BASE, "output", "Task3_Customer_Spending_Report.xlsx")
os.makedirs(CHARTS, exist_ok=True)
os.makedirs(os.path.join(BASE, "output"), exist_ok=True)

# ── Load & Validate ───────────────────────────────────────────────────────────
print("=" * 60)
print("  TASK 3 – CUSTOMER SPENDING REPORT")
print("=" * 60)

df = pd.read_excel(DATA)
print(f"\n✅ Dataset loaded: {df.shape[0]:,} rows × {df.shape[1]} columns")
print(f"   Columns: {list(df.columns)}")
print(f"   Missing values: {df.isnull().sum().sum()}")
print(f"   Duplicates    : {df.duplicated().sum()}")

# ── Colours ───────────────────────────────────────────────────────────────────
C1, C2, C3, C4, C5 = "#2E86AB", "#A23B72", "#F18F01", "#C73E1D", "#3B1F2B"
PALETTE = [C1, C2, C3, C4, C5,
           "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7", "#DDA0DD"]

# ══════════════════════════════════════════════════════════════════════════════
# 1. TOP 10 SPENDING CUSTOMERS
# ══════════════════════════════════════════════════════════════════════════════
top10 = (df.nlargest(10, "Total_Spending")
           [["Customer_ID", "Customer_Name", "City", "Customer_Type",
             "Total_Spending", "Purchase_Count", "Satisfaction_Score"]]
           .reset_index(drop=True))
top10.index = top10.index + 1

print("\n📊 TOP 10 SPENDING CUSTOMERS")
print(top10.to_string())

# Chart 1 – Horizontal bar: Top 10 customers
fig, ax = plt.subplots(figsize=(12, 7))
bars = ax.barh(top10["Customer_Name"][::-1],
               top10["Total_Spending"][::-1] / 1_000,
               color=PALETTE[:10][::-1], edgecolor="white", linewidth=0.5)
for bar, val in zip(bars, top10["Total_Spending"][::-1] / 1_000):
    ax.text(bar.get_width() + 2, bar.get_y() + bar.get_height() / 2,
            f"₹{val:.0f}K", va="center", fontsize=9, fontweight="bold",
            color="#333")
ax.set_xlabel("Total Spending (₹ Thousands)", fontsize=11)
ax.set_title("Top 10 Highest Spending Customers", fontsize=14,
             fontweight="bold", pad=15)
ax.set_facecolor("#f8f9fa")
fig.patch.set_facecolor("white")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.grid(axis="x", alpha=0.3, linestyle="--")
plt.tight_layout()
plt.savefig(os.path.join(CHARTS, "1_top10_spenders.png"), dpi=150,
            bbox_inches="tight")
plt.close()
print("\n✅ Chart 1 saved: 1_top10_spenders.png")

# ══════════════════════════════════════════════════════════════════════════════
# 2. AVERAGE CUSTOMER SPENDING
# ══════════════════════════════════════════════════════════════════════════════
avg_overall   = df["Total_Spending"].mean()
median_spend  = df["Total_Spending"].median()
max_spend     = df["Total_Spending"].max()
min_spend     = df["Total_Spending"].min()
total_revenue = df["Total_Spending"].sum()

avg_by_type = df.groupby("Customer_Type")["Total_Spending"].mean().sort_values(ascending=False)

print(f"\n💰 AVERAGE SPENDING ANALYSIS")
print(f"   Overall Average : ₹{avg_overall:,.0f}")
print(f"   Median          : ₹{median_spend:,.0f}")
print(f"   Max             : ₹{max_spend:,.0f}")
print(f"   Min             : ₹{min_spend:,.0f}")
print(f"   Total Revenue   : ₹{total_revenue:,.0f}")
print(f"\n   By Customer Type:")
for t, v in avg_by_type.items():
    print(f"   {t:12s}: ₹{v:,.0f}")

# Chart 2 – Bar: Avg spending by customer type
fig, ax = plt.subplots(figsize=(9, 6))
colors_type = [C1, C2, C3]
bars = ax.bar(avg_by_type.index, avg_by_type.values / 1_000,
              color=colors_type, edgecolor="white", linewidth=0.8, width=0.5)
for bar, val in zip(bars, avg_by_type.values / 1_000):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
            f"₹{val:.0f}K", ha="center", fontsize=11, fontweight="bold",
            color="#333")
ax.set_ylabel("Average Spending (₹ Thousands)", fontsize=11)
ax.set_title("Average Spending by Customer Type", fontsize=14,
             fontweight="bold", pad=15)
ax.set_facecolor("#f8f9fa")
fig.patch.set_facecolor("white")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.grid(axis="y", alpha=0.3, linestyle="--")
ax.set_ylim(0, avg_by_type.max() / 1_000 * 1.2)
plt.tight_layout()
plt.savefig(os.path.join(CHARTS, "2_avg_spending_by_type.png"), dpi=150,
            bbox_inches="tight")
plt.close()
print("\n✅ Chart 2 saved: 2_avg_spending_by_type.png")

# ══════════════════════════════════════════════════════════════════════════════
# 3. AGE-GROUP SPENDING ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
bins   = [0, 24, 34, 44, 54, 64, 120]
labels = ["18–24", "25–34", "35–44", "45–54", "55–64", "65+"]
df["Age_Group"] = pd.cut(df["Age"], bins=bins, labels=labels, right=True)

age_stats = df.groupby("Age_Group", observed=True).agg(
    Count        = ("Customer_ID",   "count"),
    Avg_Spending = ("Total_Spending", "mean"),
    Total_Spend  = ("Total_Spending", "sum"),
    Avg_Purchases= ("Purchase_Count", "mean"),
    Avg_Satisfaction=("Satisfaction_Score","mean")
).reset_index()

print(f"\n👥 AGE-GROUP SPENDING ANALYSIS")
print(age_stats.to_string(index=False))

# Chart 3 – Grouped bar: Count + Avg Spending by age group
fig, ax1 = plt.subplots(figsize=(12, 6))
x     = np.arange(len(age_stats))
width = 0.4

bars1 = ax1.bar(x - width/2, age_stats["Count"],
                width=width, color=C1, alpha=0.85,
                label="Customer Count", edgecolor="white")
ax1.set_xlabel("Age Group", fontsize=11)
ax1.set_ylabel("Number of Customers", fontsize=11, color=C1)
ax1.tick_params(axis="y", labelcolor=C1)

ax2 = ax1.twinx()
bars2 = ax2.bar(x + width/2, age_stats["Avg_Spending"] / 1_000,
                width=width, color=C2, alpha=0.85,
                label="Avg Spending (₹K)", edgecolor="white")
ax2.set_ylabel("Avg Spending (₹ Thousands)", fontsize=11, color=C2)
ax2.tick_params(axis="y", labelcolor=C2)

for bar in bars1:
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10,
             f"{int(bar.get_height())}", ha="center", fontsize=8,
             color=C1, fontweight="bold")
for bar in bars2:
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
             f"₹{bar.get_height():.0f}K", ha="center", fontsize=8,
             color=C2, fontweight="bold")

ax1.set_xticks(x)
ax1.set_xticklabels(age_stats["Age_Group"])
ax1.set_title("Age-Group Spending Analysis", fontsize=14,
              fontweight="bold", pad=15)
ax1.set_facecolor("#f8f9fa")
fig.patch.set_facecolor("white")
ax1.spines["top"].set_visible(False)
ax2.spines["top"].set_visible(False)

lines = [mpatches.Patch(color=C1, label="Customer Count"),
         mpatches.Patch(color=C2, label="Avg Spending (₹K)")]
ax1.legend(handles=lines, loc="upper left", fontsize=9)
plt.tight_layout()
plt.savefig(os.path.join(CHARTS, "3_age_group_analysis.png"), dpi=150,
            bbox_inches="tight")
plt.close()
print("\n✅ Chart 3 saved: 3_age_group_analysis.png")

# ══════════════════════════════════════════════════════════════════════════════
# 4. SPENDING TIER SEGMENTATION (Dashboard insight)
# ══════════════════════════════════════════════════════════════════════════════
spend_bins  = [0, 100000, 250000, 400000, 500001]
spend_labels= ["Budget (<₹1L)", "Mid-Range (₹1L–2.5L)",
               "High (₹2.5L–4L)", "Top Tier (>₹4L)"]
df["Spend_Tier"] = pd.cut(df["Total_Spending"], bins=spend_bins,
                           labels=spend_labels, right=False)
tier_counts = df["Spend_Tier"].value_counts().reindex(spend_labels)

print(f"\n🏆 SPENDING TIER SEGMENTATION")
for tier, cnt in tier_counts.items():
    pct = cnt / len(df) * 100
    print(f"   {tier:30s}: {cnt:,} ({pct:.1f}%)")

# Chart 4 – Pie: Spending tiers
fig, ax = plt.subplots(figsize=(9, 7))
wedge_colors = [C1, C2, C3, C4]
wedges, texts, autotexts = ax.pie(
    tier_counts.values,
    labels=tier_counts.index,
    autopct="%1.1f%%",
    colors=wedge_colors,
    startangle=140,
    pctdistance=0.75,
    wedgeprops={"edgecolor": "white", "linewidth": 2}
)
for t in texts:
    t.set_fontsize(9)
for at in autotexts:
    at.set_fontsize(9)
    at.set_fontweight("bold")
    at.set_color("white")
ax.set_title("Customer Spending Tier Distribution", fontsize=14,
             fontweight="bold", pad=20)
fig.patch.set_facecolor("white")
plt.tight_layout()
plt.savefig(os.path.join(CHARTS, "4_spending_tiers.png"), dpi=150,
            bbox_inches="tight")
plt.close()
print("\n✅ Chart 4 saved: 4_spending_tiers.png")

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL WORKBOOK – Dashboard + Data + Summary
# ══════════════════════════════════════════════════════════════════════════════
print("\n📁 Building Excel workbook …")
wb = Workbook()

# ── Helper styles ─────────────────────────────────────────────────────────────
def hdr_font(bold=True, size=11, color="FFFFFF"):
    return Font(bold=bold, size=size, color=color, name="Calibri")

def cell_font(bold=False, size=10, color="333333"):
    return Font(bold=bold, size=size, color=color, name="Calibri")

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def write_cell(ws, row, col, value, font=None, fill_=None, align=None, border=None, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:   c.font      = font
    if fill_:  c.fill      = fill_
    if align:  c.alignment = align
    if border: c.border    = border
    if num_fmt: c.number_format = num_fmt
    return c

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 – DATA
# ══════════════════════════════════════════════════════════════════════════════
ws_data = wb.active
ws_data.title = "Data"
ws_data.sheet_view.showGridLines = False

headers = list(df.columns)
# Remove Age_Group and Spend_Tier (derived cols) – keep originals only
orig_cols = ["Customer_ID","Customer_Name","City","Age","Customer_Type",
             "Purchase_Count","Total_Spending","Feedback",
             "Satisfaction_Score","Membership_Years"]

hdr_fills = ["1B4F72","2E86AB","A23B72","F18F01","C73E1D",
             "1B4F72","2E86AB","A23B72","F18F01","C73E1D"]

for ci, (col, fc) in enumerate(zip(orig_cols, hdr_fills), start=1):
    c = ws_data.cell(row=1, column=ci, value=col.replace("_"," "))
    c.font      = hdr_font(size=10)
    c.fill      = fill(fc)
    c.alignment = center()
    c.border    = thin_border()

col_widths = [14, 22, 16, 8, 16, 16, 18, 12, 18, 18]
for ci, w in enumerate(col_widths, start=1):
    ws_data.column_dimensions[get_column_letter(ci)].width = w
ws_data.row_dimensions[1].height = 30

light_row  = fill("EBF5FB")
white_row  = fill("FFFFFF")

for ri, row in enumerate(df[orig_cols].itertuples(index=False), start=2):
    rf = light_row if ri % 2 == 0 else white_row
    for ci, val in enumerate(row, start=1):
        c = ws_data.cell(row=ri, column=ci, value=val)
        c.fill      = rf
        c.alignment = center() if ci != 2 else left()
        c.border    = thin_border()
        c.font      = cell_font(size=9)
        if ci == 7:   # Total_Spending
            c.number_format = '₹#,##0'

ws_data.freeze_panes = "A2"
print("   ✅ Data sheet written (8,000 rows)")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 – SUMMARY (KPI Dashboard)
# ══════════════════════════════════════════════════════════════════════════════
ws_sum = wb.create_sheet("Summary")
ws_sum.sheet_view.showGridLines = False
ws_sum.sheet_view.zoomScale     = 100

# Title banner
ws_sum.merge_cells("A1:J1")
c = ws_sum["A1"]
c.value     = "CUSTOMER SPENDING REPORT – RAUSHBYTE TECHNOLOGIES"
c.font      = Font(bold=True, size=16, color="FFFFFF", name="Calibri")
c.fill      = fill("1B4F72")
c.alignment = center()
ws_sum.row_dimensions[1].height = 40

ws_sum.merge_cells("A2:J2")
c = ws_sum["A2"]
c.value     = "Data Analytics Internship  |  Task 3  |  Dataset: customer_analytics_dataset.xlsx  |  Records: 8,000"
c.font      = Font(italic=True, size=10, color="FFFFFF", name="Calibri")
c.fill      = fill("2E86AB")
c.alignment = center()
ws_sum.row_dimensions[2].height = 22

# ── KPI Cards row ─────────────────────────────────────────────────────────────
kpi_configs = [
    ("Total Customers",   "=COUNTA(Data!A2:A8001)",    "1B4F72", '0'),
    ("Total Revenue",     "=SUM(Data!G2:G8001)",        "A23B72", '₹#,##0'),
    ("Avg Spending",      "=AVERAGE(Data!G2:G8001)",    "C73E1D", '₹#,##0'),
    ("Max Spending",      "=MAX(Data!G2:G8001)",        "F18F01", '₹#,##0'),
    ("Avg Purchases",     "=AVERAGE(Data!F2:F8001)",    "1B4F72", '0.0'),
]

ws_sum.row_dimensions[4].height = 18
ws_sum.row_dimensions[5].height = 40
ws_sum.row_dimensions[6].height = 30
ws_sum.row_dimensions[7].height = 18

for i, (label, formula, color, nfmt) in enumerate(kpi_configs):
    col_s = 1 + i * 2       # A,C,E,G,I
    col_e = col_s + 1
    col_l = get_column_letter(col_s)
    col_r = get_column_letter(col_e)

    ws_sum.merge_cells(f"{col_l}4:{col_r}4")
    lc = ws_sum[f"{col_l}4"]
    lc.value     = label
    lc.font      = Font(bold=True, size=9, color="FFFFFF", name="Calibri")
    lc.fill      = fill(color)
    lc.alignment = center()

    ws_sum.merge_cells(f"{col_l}5:{col_r}5")
    vc = ws_sum[f"{col_l}5"]
    vc.value          = formula
    vc.number_format  = nfmt
    vc.font           = Font(bold=True, size=16, color=color, name="Calibri")
    vc.fill           = fill("F8F9FA")
    vc.alignment      = center()
    vc.border         = thin_border()

    ws_sum.merge_cells(f"{col_l}6:{col_r}6")
    ws_sum[f"{col_l}6"].fill = fill("EBF5FB")

# col widths
for ci in range(1, 11):
    ws_sum.column_dimensions[get_column_letter(ci)].width = 16

# ── Section: Top 10 Spenders ──────────────────────────────────────────────────
ROW = 9
ws_sum.merge_cells(f"A{ROW}:J{ROW}")
c = ws_sum[f"A{ROW}"]
c.value     = "🏆  TOP 10 HIGHEST SPENDING CUSTOMERS"
c.font      = Font(bold=True, size=12, color="FFFFFF", name="Calibri")
c.fill      = fill("1B4F72")
c.alignment = left()
ws_sum.row_dimensions[ROW].height = 26
ROW += 1

t10_headers = ["Rank","Customer ID","Customer Name","City","Type",
               "Purchase Count","Total Spending","Satisfaction Score"]
t10_cols    = [1,2,3,4,5,6,7,9]   # columns A–H for this section
for ci, hdr in zip(range(1,9), t10_headers):
    c = ws_sum.cell(row=ROW, column=ci, value=hdr)
    c.font      = Font(bold=True, size=9, color="FFFFFF", name="Calibri")
    c.fill      = fill("2E86AB")
    c.alignment = center()
    c.border    = thin_border()
ws_sum.row_dimensions[ROW].height = 22
ROW += 1

for rank_i, row in enumerate(top10.itertuples(), start=1):
    rf = fill("EBF5FB") if rank_i % 2 == 0 else fill("FFFFFF")
    vals = [rank_i, row.Customer_ID, row.Customer_Name, row.City,
            row.Customer_Type, row.Purchase_Count,
            row.Total_Spending, row.Satisfaction_Score]
    for ci, val in enumerate(vals, start=1):
        c = ws_sum.cell(row=ROW, column=ci, value=val)
        c.font      = cell_font(size=9, bold=(ci==1))
        c.fill      = rf
        c.alignment = center()
        c.border    = thin_border()
        if ci == 7:
            c.number_format = '₹#,##0'
    ws_sum.row_dimensions[ROW].height = 18
    ROW += 1

ROW += 1  # spacer

# ── Section: Avg Spending by Type ────────────────────────────────────────────
ws_sum.merge_cells(f"A{ROW}:E{ROW}")
c = ws_sum[f"A{ROW}"]
c.value     = "💰  AVERAGE SPENDING BY CUSTOMER TYPE"
c.font      = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
c.fill      = fill("A23B72")
c.alignment = left()
ws_sum.row_dimensions[ROW].height = 26
ROW += 1

for ci, hdr in enumerate(["Customer Type","Avg Spending","Avg Purchases","Avg Satisfaction","Count"], start=1):
    c = ws_sum.cell(row=ROW, column=ci, value=hdr)
    c.font      = Font(bold=True, size=9, color="FFFFFF", name="Calibri")
    c.fill      = fill("C45B8A")
    c.alignment = center()
    c.border    = thin_border()
ws_sum.row_dimensions[ROW].height = 20
ROW += 1

type_formulas = [
    ("New",       '=AVERAGEIF(Data!E2:E8001,"New",Data!G2:G8001)',
                  '=AVERAGEIF(Data!E2:E8001,"New",Data!F2:F8001)',
                  '=AVERAGEIF(Data!E2:E8001,"New",Data!I2:I8001)',
                  '=COUNTIF(Data!E2:E8001,"New")'),
    ("Premium",   '=AVERAGEIF(Data!E2:E8001,"Premium",Data!G2:G8001)',
                  '=AVERAGEIF(Data!E2:E8001,"Premium",Data!F2:F8001)',
                  '=AVERAGEIF(Data!E2:E8001,"Premium",Data!I2:I8001)',
                  '=COUNTIF(Data!E2:E8001,"Premium")'),
    ("Returning", '=AVERAGEIF(Data!E2:E8001,"Returning",Data!G2:G8001)',
                  '=AVERAGEIF(Data!E2:E8001,"Returning",Data!F2:F8001)',
                  '=AVERAGEIF(Data!E2:E8001,"Returning",Data!I2:I8001)',
                  '=COUNTIF(Data!E2:E8001,"Returning")'),
]
type_colors = ["EBF5FB","FFFFFF","EBF5FB"]
for ri, (typ, avg_s, avg_p, avg_sat, cnt) in enumerate(type_formulas):
    rf  = fill(type_colors[ri])
    row_vals = [typ, avg_s, avg_p, avg_sat, cnt]
    for ci, val in enumerate(row_vals, start=1):
        c = ws_sum.cell(row=ROW, column=ci, value=val)
        c.font      = cell_font(size=9, bold=(ci==1))
        c.fill      = rf
        c.alignment = center()
        c.border    = thin_border()
        if ci == 2:  c.number_format = '₹#,##0'
        if ci == 3:  c.number_format = '0.0'
        if ci == 4:  c.number_format = '0.00'
    ws_sum.row_dimensions[ROW].height = 18
    ROW += 1

ROW += 1  # spacer

# ── Section: Age-Group Analysis ───────────────────────────────────────────────
ws_sum.merge_cells(f"A{ROW}:G{ROW}")
c = ws_sum[f"A{ROW}"]
c.value     = "👥  AGE-GROUP SPENDING ANALYSIS"
c.font      = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
c.fill      = fill("C73E1D")
c.alignment = left()
ws_sum.row_dimensions[ROW].height = 26
ROW += 1

age_hdrs = ["Age Group","Customers","Avg Spending","Total Spending",
            "Avg Purchases","Avg Satisfaction","% of Total"]
for ci, hdr in enumerate(age_hdrs, start=1):
    c = ws_sum.cell(row=ROW, column=ci, value=hdr)
    c.font      = Font(bold=True, size=9, color="FFFFFF", name="Calibri")
    c.fill      = fill("E05A3A")
    c.alignment = center()
    c.border    = thin_border()
ws_sum.row_dimensions[ROW].height = 20
ROW += 1

age_row_colors = ["EBF5FB","FFFFFF"] * 3
for ri, row_data in enumerate(age_stats.itertuples(index=False)):
    rf = fill(age_row_colors[ri])
    vals = [str(row_data.Age_Group),
            row_data.Count,
            row_data.Avg_Spending,
            row_data.Total_Spend,
            row_data.Avg_Purchases,
            row_data.Avg_Satisfaction,
            row_data.Count / 8000]
    for ci, val in enumerate(vals, start=1):
        c = ws_sum.cell(row=ROW, column=ci, value=val)
        c.font      = cell_font(size=9, bold=(ci==1))
        c.fill      = rf
        c.alignment = center()
        c.border    = thin_border()
        if ci == 3:  c.number_format = '₹#,##0'
        if ci == 4:  c.number_format = '₹#,##0'
        if ci == 5:  c.number_format = '0.0'
        if ci == 6:  c.number_format = '0.00'
        if ci == 7:  c.number_format = '0.0%'
    ws_sum.row_dimensions[ROW].height = 18
    ROW += 1

ROW += 1

# ── Section: Spending Tier ────────────────────────────────────────────────────
ws_sum.merge_cells(f"A{ROW}:E{ROW}")
c = ws_sum[f"A{ROW}"]
c.value     = "🏅  SPENDING TIER SEGMENTATION"
c.font      = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
c.fill      = fill("1B4F72")
c.alignment = left()
ws_sum.row_dimensions[ROW].height = 26
ROW += 1

for ci, hdr in enumerate(["Tier","Customers","% of Total","Cumulative %"], start=1):
    c = ws_sum.cell(row=ROW, column=ci, value=hdr)
    c.font      = Font(bold=True, size=9, color="FFFFFF", name="Calibri")
    c.fill      = fill("2E86AB")
    c.alignment = center()
    c.border    = thin_border()
ws_sum.row_dimensions[ROW].height = 20
ROW += 1

cumulative = 0
tier_row_colors = ["EBF5FB","FFFFFF"] * 3
for ti, (tier, cnt) in enumerate(tier_counts.items()):
    pct = cnt / 8000
    cumulative += pct
    rf = fill(tier_row_colors[ti % 2])
    for ci, val in enumerate([tier, cnt, pct, cumulative], start=1):
        c = ws_sum.cell(row=ROW, column=ci, value=val)
        c.font      = cell_font(size=9, bold=(ci==1))
        c.fill      = rf
        c.alignment = center()
        c.border    = thin_border()
        if ci in (3,4): c.number_format = '0.0%'
    ws_sum.row_dimensions[ROW].height = 18
    ROW += 1

# ── Embed Charts ──────────────────────────────────────────────────────────────
from openpyxl.drawing.image import Image as XLImage

chart_files = [
    "1_top10_spenders.png",
    "2_avg_spending_by_type.png",
    "3_age_group_analysis.png",
    "4_spending_tiers.png"
]
# Place charts in a third sheet
ws_charts = wb.create_sheet("Charts")
ws_charts.sheet_view.showGridLines = False
ws_charts.merge_cells("A1:N1")
c = ws_charts["A1"]
c.value     = "CUSTOMER SPENDING REPORT – CHARTS DASHBOARD"
c.font      = Font(bold=True, size=14, color="FFFFFF", name="Calibri")
c.fill      = fill("1B4F72")
c.alignment = center()
ws_charts.row_dimensions[1].height = 35

chart_positions = ["A2", "H2", "A28", "H28"]
for chart_file, pos in zip(chart_files, chart_positions):
    fpath = os.path.join(CHARTS, chart_file)
    if os.path.exists(fpath):
        img = XLImage(fpath)
        img.width  = 480
        img.height = 280
        ws_charts.add_image(img, pos)

# ── Embed 2 Excel-native charts in Summary sheet ──────────────────────────────
# Native Bar chart: Avg spending by type (using type_formulas data written above)
# We'll reference the written data in Summary sheet rows
# Find the row where type data starts (after "AVG SPENDING BY TYPE" header+col header)
# It's easier to add a small hidden data table and reference it.

ws_cd = wb.create_sheet("ChartData")   # helper sheet, hidden later
ws_cd["A1"] = "Type"
ws_cd["B1"] = "AvgSpending"
ws_cd["C1"] = "Count"
type_raw = df.groupby("Customer_Type")["Total_Spending"].agg(["mean","count"]).reset_index()
for ri2, row2 in enumerate(type_raw.itertuples(index=False), start=2):
    ws_cd.cell(row=ri2, column=1, value=row2.Customer_Type)
    ws_cd.cell(row=ri2, column=2, value=round(row2.mean, 0))
    ws_cd.cell(row=ri2, column=3, value=int(row2.count))

ws_cd["E1"] = "AgeGroup"
ws_cd["F1"] = "AvgSpending"
for ri2, row2 in enumerate(age_stats.itertuples(index=False), start=2):
    ws_cd.cell(row=ri2, column=5, value=str(row2.Age_Group))
    ws_cd.cell(row=ri2, column=6, value=round(row2.Avg_Spending, 0))

# Bar chart – Type avg spending
bc = BarChart()
bc.type    = "col"
bc.title   = "Avg Spending by Customer Type"
bc.y_axis.title = "Avg Spending (₹)"
bc.x_axis.title = "Customer Type"
bc.style   = 10
bc.width   = 14
bc.height  = 10
data_ref   = Reference(ws_cd, min_col=2, max_col=2, min_row=1, max_row=4)
cats_ref   = Reference(ws_cd, min_col=1, min_row=2, max_row=4)
bc.add_data(data_ref, titles_from_data=True)
bc.set_categories(cats_ref)
ws_sum.add_chart(bc, f"F{ROW+2}")

# Bar chart – Age group avg spending
bc2 = BarChart()
bc2.type    = "col"
bc2.title   = "Avg Spending by Age Group"
bc2.y_axis.title = "Avg Spending (₹)"
bc2.x_axis.title = "Age Group"
bc2.style   = 10
bc2.width   = 14
bc2.height  = 10
data_ref2  = Reference(ws_cd, min_col=6, max_col=6, min_row=1, max_row=7)
cats_ref2  = Reference(ws_cd, min_col=5, min_row=2, max_row=7)
bc2.add_data(data_ref2, titles_from_data=True)
bc2.set_categories(cats_ref2)
ws_sum.add_chart(bc2, f"F{ROW+18}")

ws_cd.sheet_state = "hidden"

wb.save(OUTPUT)
print(f"\n✅ Excel workbook saved: {OUTPUT}")

# ══════════════════════════════════════════════════════════════════════════════
# FINAL SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("  TASK 3 COMPLETE – RESULTS SUMMARY")
print("=" * 60)
print(f"  Total Customers     : 8,000")
print(f"  Total Revenue       : ₹{total_revenue:,.0f}")
print(f"  Avg Spending        : ₹{avg_overall:,.0f}")
print(f"  Median Spending     : ₹{median_spend:,.0f}")
print(f"  Highest Spender     : {top10.iloc[0]['Customer_Name']} (₹{top10.iloc[0]['Total_Spending']:,.0f})")
print(f"  Highest Age Group   : {age_stats.loc[age_stats['Avg_Spending'].idxmax(),'Age_Group']} (₹{age_stats['Avg_Spending'].max():,.0f} avg)")
print(f"  Top Tier (>₹4L)     : {tier_counts.get('Top Tier (>₹4L)',0):,} customers")
print("\n  Charts:  charts/1_top10_spenders.png")
print("           charts/2_avg_spending_by_type.png")
print("           charts/3_age_group_analysis.png")
print("           charts/4_spending_tiers.png")
print(f"\n  Excel:   output/Task3_Customer_Spending_Report.xlsx")
print("=" * 60)
