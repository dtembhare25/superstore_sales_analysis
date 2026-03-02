# 🛒 Superstore Sales Analysis — Canada

> 🌱 *This is the first project I completed after deciding to transition from Civil Engineering into Business Analytics. It marks the beginning of my journey into data-driven decision making.*

**Tools Used:** Microsoft Excel (Pivot Tables, Conditional Formatting, Data Sorting & Filtering)  
**Dataset:** Superstore Canada Sales Data | 8,400 rows across 7 regions  
**Course:** Business Analyst Program — UpGrad  
**Focus Segment:** Corporate Customers

---

## 📌 Project Overview

This project analyses retail sales data from one of Canada's largest superstore chains. As a simulated Sales Manager, the goal was to identify which product sub-categories are driving profits and which are causing losses — across 7 Canadian regions — to support strategic business decisions like product discontinuation or regional pricing changes.

---

## 🎯 Objectives

- Filter and segment 8,400+ orders by customer type (Corporate, Consumer, Home Office, Small Business)
- Clean, format, and organise the Corporate segment data for reporting
- Identify the **top 3 most profitable** product sub-categories per region using Pivot Tables
- Identify the **most loss-making** sub-categories and the regions most affected
- Apply conditional formatting to make profit/loss patterns visually intuitive

---

## 📁 File Structure

```
superstore-sales-analysis.xlsx
│
├── RAW DATA S.S.P.          ← Original dataset (8,400 rows, 18 columns)
├── CORPORATE SEGMENT        ← Filtered & formatted report (3,076 rows)
├── CONSUMER SEGMENT         ← Filtered data
├── HOME OFFICE SEGMENT      ← Filtered data
├── SMALL BUSINESS SEGMENT   ← Filtered data
├── HIGH PROFIT CAT. PIVOT TABLE   ← Activity 2: Top performers by region
└── LOSS MAKING CAT. PIVOT TABLE   ← Activity 2: Loss makers by region
```

---

## 🔍 Key Insights

### ✅ Top Profitable Sub-Categories

| Sub-Category | Best Region | Insight |
|---|---|---|
| **Office Machines** | Prarie ($91,913) | Highest profit generator across all regions |
| **Telephones & Communication** | Ontario & West (~$25–26K) | Consistently profitable everywhere |
| **Binders & Accessories** | West ($43,849) | Strong performer across all 7 regions |

### ❌ Consistent Loss-Making Sub-Categories

| Sub-Category | Problem Regions | Insight |
|---|---|---|
| **Tables** | Ontario (−$16,169), Quebec (−$6,563) | Loss in every region — candidate for discontinuation |
| **Bookcases** | Quebec (−$11,367), NW Territories (−$3,270) | Loss in 5 out of 7 regions |
| **Storage & Organization** | Atlantic (−$7,288), Quebec (−$2,369) | Mixed performance — needs regional review |
| **Scissors, Rulers & Trimmers** | Nunavut (−$1,760) | Extreme loss concentrated in one region |

### 💡 Business Recommendation

> **Tables** are the biggest concern — they generate losses across *all* 7 regions with no profitable exception. This suggests a pricing or cost structure issue rather than a regional demand problem. Discontinuation or a significant price revision should be evaluated.
>
> **Copiers & Fax** show a massive loss in Atlantic (−$17,507) but strong profits elsewhere — indicating a regional distribution or competition issue, not a product-level problem.

---

## 🛠️ Excel Skills Applied

- **Data Filtering & Segmentation** — Split 8,400 rows into 4 customer segment sheets
- **Formatting** — Header styling, date formatting (DD-MMM-YYYY), currency formatting ($USD), column width adjustment
- **Freeze Panes** — Locked header row for large dataset navigation
- **Multi-level Sorting** — Region → Province → Sales (descending)
- **Conditional Formatting** — Top 10% orders highlighted; profit/loss colour scale (green = profit, red = loss)
- **Pivot Tables** — Profit aggregated by Sub-Category × Region for both high-profit and loss-making analysis
- **Border Demarcation** — Double bottom borders applied to separate each Region visually

---

## 📊 Dataset Columns (Corporate Segment)

`Order ID` | `Customer Name` | `Order Date` | `Order Priority` | `Ship Date` | `Ship Mode` | `Region` | `Province` | `Order Quantity` | `Unit Price` | `Shipping Cost` | `Discount` | `Product Base Margin` | `Profit` | `Sales` | `Product Category` | `Product Sub-Category`

---

*This project was completed as part of the UpGrad Business Analyst certification programme.*
