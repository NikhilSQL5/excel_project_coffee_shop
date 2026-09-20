# BrewTrack — Coffee Shop Sales Performance Analysis

An Excel Data Analyst portfolio project: cleaning, analyzing, and visualizing 3 months of multi-outlet coffee shop sales data using core Excel formulas — no VBA, no macros, no Power Pivot.

## Project Overview

BrewTrack is a fictional 4-outlet coffee shop chain in Bengaluru. This project takes a messy, real-world-style POS export, cleans it with Excel formulas, calculates business KPIs, builds PivotTable-style summary tables, and presents the findings on a one-page interactive dashboard.

## Business Problem

BrewTrack's owner exports daily sales from the POS billing system into a spreadsheet but has never analyzed it. This project answers: what's selling, which outlet performs best, what a typical order looks like, and where to focus attention next quarter.

## Objectives

- Clean 3 months of raw POS export data (200 rows, with realistic data-quality issues)
- Calculate core KPIs: Total Sales, Total Profit, Profit Margin, Average Order Value, Total Orders
- Identify best-selling products and the most profitable category
- Compare outlet performance
- Spot the monthly sales trend across the quarter
- Build summary tables (PivotTable-style) answering every business question
- Build one clean, interactive dashboard
- Turn the analysis into insights and recommendations the owner could act on

## Dataset

- **Period:** June 1 – August 31, 2026
- **Outlets:** Koramangala, Indiranagar, Whitefield, HSR Layout
- **Raw export:** 200 order-line rows, with intentional data-quality issues (see Data Cleaning below)
- **Cleaned dataset:** 197 unique orders (3 duplicate Order_IDs removed)

| File | Description |
|---|---|
| `data/sales_data.csv` | Final cleaned dataset (197 rows) — matches the Cleaned_Data tab in the workbook exactly |
| `excel/Excel_Data_Analyst_Portfolio_Project.xlsx` | Full workbook: raw data, cleaning formulas, KPIs, summary tables, dashboard |

## Data Dictionary

| Column | Type | Description |
|---|---|---|
| Order_ID | Text | Unique order identifier |
| Order_Date | Date | Date the order was placed |
| Outlet | Text | Koramangala / Indiranagar / Whitefield / HSR Layout |
| Product | Text | Item ordered |
| Category | Text | Hot Beverages / Cold Beverages / Bakery / Snacks |
| Quantity | Number | Units ordered |
| Unit_Price | Number | Price per unit (₹) |
| Sales | Number (calc.) | Quantity × Unit_Price |
| Cost | Number | Cost of goods for the line (₹) |
| Profit | Number (calc.) | Sales − Cost |
| Profit_Margin | % (calc.) | Profit ÷ Sales |
| Month | Text (calc.) | e.g. "Jun-2026" — for trend grouping |
| Year | Number (calc.) | Order year |
| Payment_Method | Text | UPI / Cash / Card |
| Order_Status | Text | Completed / Cancelled |
| Order_Status_Category | Text (calc.) | Groups status into Completed / Not Completed |

## Tools & Technologies

- Microsoft Excel (formulas written for broad version compatibility — see Formulas Used)
- GitHub (portfolio hosting)

## Data Cleaning

The raw 200-row export had 5 realistic issues, all fixed with formulas (nothing retyped by hand):

| Issue | Rows | Fix |
|---|---|---|
| Duplicate Order_IDs | 3 | Cleaned_Data references only the first occurrence of each ID |
| Category casing/spacing/naming (e.g. "hot beverages", "Hot Beverage") | 19 | TRIM + PROPER, plus a 2-row INDEX+MATCH exception table for the singular/plural mismatch |
| Outlet casing/spacing/abbreviation (e.g. "koramangala", "HSR") | 19 | TRIM + PROPER, plus a 2-row INDEX+MATCH exception table for the "HSR" abbreviation |
| Order_Date stored as DD/MM/YYYY text | 11 | IF(ISNUMBER(...)) test + LEFT/MID/RIGHT + DATE() reconstruction |
| Blank Payment_Method | 7 | Labeled "Not Specified" instead of left blank |
| Missing Quantity | 3 | Defaulted to 1 — a documented assumption, not a silent guess |

## Formulas Used

| Category | Formulas |
|---|---|
| Basic Aggregation | SUM, AVERAGE, COUNT, COUNTA |
| Conditional Analysis | COUNTIF, COUNTIFS, SUMIFS |
| Logical | IF, IFERROR |
| Lookup | INDEX + MATCH (with IFERROR fallback) |
| Text Cleaning | TRIM, PROPER, LEFT, MID, RIGHT |
| Date Analysis | YEAR, DATE, TEXT, ISNUMBER |
| Value Conversion | VALUE |

**Note on XLOOKUP:** every lookup in this project uses INDEX+MATCH instead of XLOOKUP, since INDEX+MATCH works in every Excel version back to 2010, while XLOOKUP needs Excel 365/2021+. On a modern Excel install, any INDEX+MATCH formula here could be rewritten as a single XLOOKUP.

Full explanations (purpose, syntax, business use, common mistakes) for every formula are in `documentation/Excel_Data_Analyst_Project_Documentation.docx`.

## Analysis Performed

Five SUMIFS/COUNTIFS-based summary tables (PivotTable-style):

1. **Sales by Month** — Rows: Month · Values: Sum of Sales (Completed only)
2. **Sales & Profit by Category** — Rows: Category · Values: Sales, Profit, Margin %, % of Total
3. **Sales by Outlet** — Rows: Outlet · Values: Sales, Order Count, % of Total
4. **Top 5 Products** — Rows: Product · Values: Sum of Sales
5. **Order Status Breakdown** — Rows: Order Status · Values: Count, % of Total

## Dashboard

One page, built entirely from formulas — nothing is a typed-in number:

- **5 KPI cards:** Total Sales, Total Profit, Total Orders, Average Order Value, Profit Margin
- **5 charts:** Monthly Sales Trend (line), Sales by Category (bar), Sales by Outlet (bar), Top 5 Products (horizontal bar), Order Status Distribution (pie)
- **Filtering:** AutoFilter dropdowns on the Cleaned_Data tab (Outlet, Category, Month, Order Status) — used instead of native Slicers. See *Project Limitations* below for why.

## Key KPIs

| KPI | Value |
|---|---|
| Total Orders | 197 |
| Completed Orders | 170 (86.3%) |
| Cancelled Orders | 27 (13.7%) |
| Total Sales (Completed only) | ₹47,560 |
| Total Profit (Completed only) | ₹26,599 |
| Profit Margin | 55.9% |
| Average Order Value | ₹280 |
| MoM Growth (Aug vs. Jul) | +57.2% |

## Key Insights

1. **Cold Beverages drives the most revenue (₹14,960, 31.5%) but Hot Beverages has the best margin (59.6% vs. 57.6%)** — revenue leadership and profitability leadership are different categories here.
2. **Snacks is the weakest category on both dimensions** — lowest revenue share (20.8%) and lowest margin (51.3%).
3. **HSR Layout underperforms the other three outlets by ~18–19%** (₹10,230 vs. ₹12,210–₹12,580 for the others), while the other three are within 3 points of each other.
4. **Frappe is a runaway best-seller** — ₹6,960 in sales, 83% higher than the next two products (Sandwich and Latte, tied at ₹3,800).
5. **Sales dipped sharply in July (−22.6%) before rebounding in August (+57.2%)** — worth investigating against any known operational disruption.
6. **13.7% of orders are cancelled** — nearly 1 in 7, and worth a root-cause check by outlet, payment method, or time period.

## Business Recommendations

1. Review Cold Beverages' ingredient/supplier costs or pricing to close some of the margin gap with Hot Beverages, without slowing its revenue growth.
2. Audit Snacks pricing and supplier costs — the only category trailing on both revenue and margin.
3. Investigate HSR Layout's operations directly (staffing, footfall, competition) given its consistent ~18% gap versus the other three outlets.
4. Protect Frappe's supply chain and consider a Frappe-anchored promotion, given its outsized share of revenue.
5. Cross-check the July dip and the 13.7% cancellation rate against outlet- and date-level detail — they may share one fixable root cause.

## Project Workflow

1. **Raw_Data** — original POS export, untouched
2. **Cleaned_Data** — formulas reference Raw_Data; casing/spacing/naming standardized, dates reconstructed, duplicates removed, calculated columns added
3. **Calculations** — core KPI formulas
4. **PivotTables** — 5 SUMIFS/COUNTIFS summary tables
5. **Dashboard** — KPI cards + 5 charts, built from PivotTables

## Folder Structure

```
BrewTrack-Excel-Portfolio/
│
├── README.md
│
├── data/
│   └── sales_data.csv
│
├── excel/
│   └── Excel_Data_Analyst_Portfolio_Project.xlsx
│
├── documentation/
    └── Excel_Data_Analyst_Project_Documentation.docx
```

## How to Use This Project

1. Open `excel/Excel_Data_Analyst_Portfolio_Project.xlsx` — start on the **README** tab for an in-workbook guide.
2. Review **Raw_Data** to see the original, unedited export.
3. Review **Cleaned_Data** to see every cleaning formula in action (click any cell to see the formula).
4. Check **Calculations** and **PivotTables** for the KPI and summary-table logic.
5. View **Dashboard** for the one-page summary.
6. To filter, use the AutoFilter dropdowns on the Cleaned_Data tab's headers.

## Skills Demonstrated

- Data cleaning (TRIM, PROPER, IF-based blank handling, text-to-date reconstruction)
- Conditional aggregation (SUMIFS, COUNTIFS)
- Lookup formulas (INDEX + MATCH with IFERROR fallback)
- Calculated columns and KPI formula design
- Pivot-style summary table design
- Dashboard layout and chart selection
- Data-driven insight and recommendation writing

## Future Improvements

- Convert the PivotTables tab to native Excel PivotTables with real Slicers
- Add Power Query for a fully repeatable cleaning pipeline on new monthly exports
- Extend to 12 months of data to confirm whether the July dip is seasonal
- Add a customer loyalty ID to enable repeat-customer analysis
- Rebuild in Power BI to compare static-workbook vs. interactive-BI-tool delivery

## Author

Author: Nikhil
