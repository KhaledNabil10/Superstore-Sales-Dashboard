# Superstore Sales Dashboard (Excel)

An interactive multi-page Excel dashboard analyzing sales, profit, 
shipping, and product performance for a retail superstore dataset.

## 📊 Data Source
Dataset based on the classic "Superstore" retail sales dataset 
(9,994 order records), publicly available on Kaggle.

## 🔍 Key Insight
While reviewing the data model, I found the original COGS calculation 
was logically inconsistent (an incorrect sign in the Discount 
adjustment), which understated true product cost by ~19%. I traced 
this back through Power Query, corrected the formula using 
Unit Cost × Quantity, and rebuilt the dependent margin calculations. 
This revealed that Furniture, while profitable, operates on a much 
thinner margin (~2.5%) compared to Technology and Office Supplies (~17%).

## 📊 Dashboard Pages
- **Home** — Overall KPIs, profit trend, top customers by sales
- **Country** — Regional & state-level performance, profit/loss by state
- **Shipping** — Delivery mode efficiency, shipping cost vs. profit
- **Sales** — Customer behavior, discount analysis, Sales vs COGS vs Profit
- **Product** — Sub-category and product-level performance

## 🛠️ Tools & Techniques
**Power Query**
- Data cleaning & transformation
- Custom calculated columns (e.g. COGS)
- Splitting & merging columns
- Conditional column logic (e.g. date/year extraction)
- Merging a separate Shipping Cost source into the main data model

**Power Pivot / DAX**
- Relationships across Orders, Returns, People, and Shipping Cost tables
- Calculated Measures (e.g. Profit Margin %, Profit Per Order)

**Excel Dashboard Design**
- PivotTables & PivotCharts
- Slicers connected across all report pages (Report Connections)
- Conditional formatting (Invert if Negative for loss-making states/products)
- VBA Macro for one-click "Clear All Filters"

## 💾 Data
All data is embedded within the workbook (loaded via Power Query) — 
no external files needed. Simply download and open the .xlsm file.

## ⚠️ Note
This workbook contains a VBA macro (`Clear_All_Slicers`) used to reset 
all dashboard filters with one click. When opening, Excel may show a 
security warning — this is expected and safe; enable macros to use 
the "Clear Filter" button.

## 📁 Files
- `Superstore-Sales-Dashboard.xlsm` — the full workbook
- `screenshots/` — preview images of each dashboard page

---
**Author:** Khaled Nabil  
🔗 [LinkedIn](https://www.linkedin.com/in/khaled-nabil-270a95253) | 📧 khalednabil369@gmail.com
