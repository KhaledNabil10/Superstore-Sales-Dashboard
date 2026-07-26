# 📊 Superstore Sales Dashboard (Excel)

An interactive **multi-page Excel dashboard** built to analyze sales, profit, shipping, customer, and product performance using the Superstore retail dataset.

The project demonstrates the complete BI workflow—from **data cleaning and transformation** to **data modeling, DAX calculations, and interactive dashboard design** using Microsoft Excel.

---

# 📷 Dashboard Preview

## 🏠 Home Dashboard

![Home](screenshots/Home.png)

**Overview of business performance including:**
- Total Sales
- Total Profit
- Total Orders
- Total Customers
- Profit Margin
- Sales Distribution
- Monthly Sales Trend
- Top Customers

---

## 🌍 Country Dashboard

![Country](screenshots/Country.png)

Analyze geographic performance through:

- Regional Sales
- State Performance
- Top States
- Sales & Profit by Region
- Monthly Sales Trend

---

## 🚚 Shipping Dashboard

![Shipping](screenshots/Shipping.png)

Shipping analysis including:

- Shipping Cost
- Delivery Duration
- Average Profit
- Ship Mode Distribution
- Shipping Cost vs Profit
- Orders by Ship Mode

---

## 💰 Sales Dashboard

![Sales](screenshots/Sales.png)

Sales analysis including:

- Customer Purchasing Behavior
- Sales vs Profit
- Sales vs COGS
- Average Discount
- Top Customers
- Sales by Segment

---

## 📦 Product Dashboard

![Product](screenshots/Product.png)

Product performance analysis including:

- Top Selling Products
- Sales vs Profit by Sub-Category
- Quantity Sold
- Profit by Category
- Product KPIs

---

# 📊 Dataset

The project uses the classic **Superstore Sales Dataset**.

- Nearly **10,000 sales records**
- Orders
- Customers
- Products
- Returns
- Regions
- Shipping information

Dataset Source:

https://www.kaggle.com/datasets

---

# 💡 Business Insights

During development, I reviewed the data model and rebuilt several calculations to improve reporting quality.

Key findings include:

- Technology generated the highest overall profit.
- Furniture produced high sales but relatively low profitability.
- Some sub-categories generated high revenue while producing negative profit.
- Office Supplies sold the largest quantity of products.
- Shipping cost and delivery performance varied significantly across shipping modes.

---

# 🛠️ Tools & Technologies

## Microsoft Excel

- PivotTables
- PivotCharts
- Dashboard Design
- Slicers
- Timeline Filters
- Conditional Formatting

---

## Power Query

- Data Cleaning
- ETL Process
- Merge Queries
- Append Queries
- Custom Columns
- Conditional Columns
- Data Transformation

---

## Power Pivot

- Data Modeling
- Relationships
- Measures
- Calculated Columns

---

## DAX

Custom Measures including:

- Total Sales
- Total Profit
- Profit Margin
- Profit per Order
- Average Shipping Cost
- Average Delivery Duration
- Quantity Sold
- Total Products

---

## VBA

Created a macro to reset every slicer across all dashboard pages with one click.

```vb
Sub Clear_All_Slicers()

    Dim sc As SlicerCache

    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc

End Sub
```

---

# ✨ Dashboard Features

- Interactive multi-page dashboard
- Fully connected slicers across all reports
- Dynamic KPI Cards
- Power Query ETL
- Power Pivot Data Model
- DAX Measures
- VBA Automation
- Responsive PivotCharts
- Business-focused visualizations

---

# 📈 Dashboard Pages

| Page | Description |
|-------|-------------|
| Home | Executive overview of business performance |
| Country | Regional and state analysis |
| Shipping | Shipping performance analysis |
| Sales | Customer and sales analysis |
| Product | Product and sub-category analysis |

---

# 📁 Repository Structure

```
Superstore-Sales-Dashboard
│
├── screenshots
│   ├── Home.png
│   ├── Country.png
│   ├── Shipping.png
│   ├── Sales.png
│   └── Product.png
│
├── Superstore-Sales-Dashboard.xlsm
└── README.md
```

---

# ⚠️ Note

The workbook contains a VBA macro:

```
Clear_All_Slicers
```

Excel may display a security warning when opening the workbook.

Simply click:

**Enable Content**

to activate the **Clear Filter** button.

---

# 💾 Download

Download the workbook:

```
Superstore-Sales-Dashboard.xlsm
```

Open the workbook using **Microsoft Excel** with macros enabled.

---

# 🧠 Skills Demonstrated

- Data Cleaning
- ETL
- Data Modeling
- Data Visualization
- Dashboard Design
- Business Intelligence
- Business Analysis
- DAX
- Power Query
- Power Pivot
- Excel Automation
- VBA

---

# 👨‍💻 Author

**Khaled Nabil**

📧 khalednabil369@gmail.com

💼 LinkedIn

https://www.linkedin.com/in/khaled-nabil-270a95253

---

## ⭐ If you found this project useful, consider giving it a Star.
