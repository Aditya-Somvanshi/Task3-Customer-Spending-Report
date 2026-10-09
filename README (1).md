# Task 3: Customer Spending Report
**RaushByte Technologies – Data Analytics Internship**

---

## 📌 Objective
Identify high-value customers based on spending behavior from 8,000 customer records — covering top spenders, average spending metrics, age-group analysis, and spending tier segmentation.

---

## ✅ Requirements Completed

| # | Requirement | Status |
|---|-------------|--------|
| 1 | Find top spending customers | ✅ Done |
| 2 | Calculate average customer spending | ✅ Done |
| 3 | Perform age-group spending analysis | ✅ Done |
| 4 | Create a simple dashboard | ✅ Done |

---

## 📊 Dataset

- **File:** `customer_analytics_dataset.xlsx`
- **Rows:** 8,000 customers
- **Columns:** 10
  - Customer_ID, Customer_Name, City, Age
  - Customer_Type (Premium / Returning / New)
  - Purchase_Count, Total_Spending
  - Feedback, Satisfaction_Score, Membership_Years

---

## 🔍 Data Quality

| Check | Result |
|-------|--------|
| Missing Values | 0 |
| Duplicates | 0 |
| Total Records | 8,000 |
| Data Integrity | 100% ✅ |

---

## 📈 Key Findings

### 1. Top 10 Spending Customers
| Rank | Customer Name | Type | Total Spending |
|------|---------------|------|---------------|
| 1 | Natalie Krueger | Returning | ₹4,99,950 |
| 2 | Clayton Clark | New | ₹4,99,932 |
| 3 | Keith Pierce | New | ₹4,99,912 |
| 4 | Rachel Parker | Premium | ₹4,99,905 |
| 5 | Louis Hurst | Premium | ₹4,99,893 |
| 6 | Dawn Hodges | Returning | ₹4,99,851 |
| 7 | Jerry Miller | Premium | ₹4,99,825 |
| 8 | Elizabeth Bridges | Returning | ₹4,99,792 |
| 9 | Deborah Howe | New | ₹4,99,688 |
| 10 | Monica Liu | Returning | ₹4,99,643 |

### 2. Average Spending Analysis
| Metric | Value |
|--------|-------|
| Overall Average | ₹2,51,068 |
| Median | ₹2,50,250 |
| Maximum | ₹4,99,950 |
| Minimum | ₹5,012 |
| **Total Revenue** | **₹200.85 Crore** |

**By Customer Type:**
| Type | Avg Spending |
|------|-------------|
| Premium | ₹2,53,878 |
| New | ₹2,51,725 |
| Returning | ₹2,47,561 |

### 3. Age-Group Spending Analysis
| Age Group | Customers | Avg Spending | % of Total |
|-----------|-----------|-------------|------------|
| 18–24 | 1,245 | ₹2,55,099 | 15.6% |
| 25–34 | 1,839 | ₹2,49,142 | 23.0% |
| 35–44 | 1,906 | ₹2,54,526 | 23.8% |
| 45–54 | 1,884 | ₹2,47,529 | 23.6% |
| 55–64 | 1,126 | ₹2,49,823 | 14.1% |

🏆 **Highest Avg Spending:** 18–24 age group (₹2,55,099)

### 4. Spending Tier Segmentation
| Tier | Customers | % |
|------|-----------|---|
| Budget (<₹1L) | 1,529 | 19.1% |
| Mid-Range (₹1L–2.5L) | 2,468 | 30.9% |
| High (₹2.5L–4L) | 2,414 | 30.2% |
| Top Tier (>₹4L) | 1,589 | 19.9% |

---

## 📁 Project Structure

```
Task3-Customer-Spending-Report/
│
├── data/
│   └── customer_analytics_dataset.xlsx    ← Raw dataset
│
├── charts/
│   ├── 1_top10_spenders.png               ← Bar chart: Top 10 customers
│   ├── 2_avg_spending_by_type.png         ← Bar chart: Avg spending by type
│   ├── 3_age_group_analysis.png           ← Grouped bar: Age-group analysis
│   └── 4_spending_tiers.png               ← Pie chart: Spending tiers
│
├── output/
│   └── Task3_Customer_Spending_Report.xlsx ← Excel dashboard workbook
│
├── customer_spending_report.py             ← Python analysis script
├── requirements.txt                        ← Python dependencies
└── README.md                               ← This file
```

---

## 🛠️ Tools Used

- **Python 3** — Data analysis and chart generation
- **pandas** — Data loading, groupby, aggregation
- **matplotlib** — Visualizations (bar charts, grouped bars, pie chart)
- **openpyxl** — Excel workbook creation with formulas and embedded charts
- **Excel** — Final dashboard with KPIs, AVERAGEIF/COUNTIF formulas, charts

---

## 📋 Excel Workbook Structure

**Sheet 1 – Data:** All 8,000 customer records with formatted table and alternating row colors

**Sheet 2 – Summary (Dashboard):**
- 5 KPI cards: Total Customers, Total Revenue, Avg Spending, Max Spending, Avg Purchases
- Top 10 Spending Customers table
- Avg Spending by Customer Type (`AVERAGEIF` formulas)
- Age-Group Spending Analysis table
- Spending Tier Segmentation table
- 2 embedded Excel-native charts (Bar: Type spending, Bar: Age-group spending)

**Sheet 3 – Charts:**
- All 4 PNG charts embedded as a visual dashboard

---

## 🚀 How to Run

1. Install dependencies:
   ```bash
   pip install pandas openpyxl matplotlib
   ```

2. Run the analysis script:
   ```bash
   python customer_spending_report.py
   ```

3. Open the Excel workbook:
   - Go to `output/Task3_Customer_Spending_Report.xlsx`
   - Open in Excel → click **Enable Editing**
   - Review **Summary** sheet for the full dashboard

---

*Task 3 of 3 | RaushByte Technologies Data Analytics Internship*
