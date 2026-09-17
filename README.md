# SuperStore-Retail-Dashboard-
SuperStore Retail data analysis using Excel 


# 📊 Superstore Sales & Profit Analysis Dashboard — Excel

An interactive **Sales & Profit Analysis Dashboard** built in Microsoft Excel using **Power Query, Power Pivot, DAX, Pivot Tables, Pivot Charts, Charts, Slicers, and Filters**.

The project analyzes Superstore sales data to identify business performance across **categories, regions, customers, products, sales, profit, discounts, and time periods**.

---

## 📌 Project Overview

This project demonstrates an end-to-end **Excel data analytics workflow**, starting from raw data preparation and transformation to data modeling, DAX calculations, visualization, and interactive dashboard development.

The dashboard is designed to help business users quickly understand:

* Overall sales performance
* Profitability and profit margins
* Category and sub-category performance
* Regional performance
* Customer and product trends
* Impact of discounts
* Sales and profit trends over time
* Returned orders

---

## 🎯 Project Objectives

The main objectives of this project are:

1. Clean and transform raw Superstore data using **Power Query**.
2. Build a structured analytical data model using **Power Pivot**.
3. Create relationships between fact and dimension tables.
4. Develop reusable **DAX measures** for business KPIs.
5. Analyze sales, profit, quantity, discount, and returns.
6. Create interactive Pivot Charts and Excel charts.
7. Add **Slicers and Filters** for dynamic analysis.
8. Build a management-friendly dashboard for data-driven decision-making.

---

## 💼 Business Questions

The dashboard answers the following business questions:

### Sales Analysis

* What is the total sales revenue?
* How does sales performance change over time?
* Which product categories generate the highest sales?
* Which regions contribute the most sales?
* Which sub-categories generate the most revenue?

### Profitability Analysis

* What is the total profit?
* What is the overall profit margin?
* Which categories are most and least profitable?
* Which sub-categories generate losses?
* How does profit vary by region?

### Customer & Product Analysis

* How many unique customers are there?
* Which products contribute significantly to sales?
* Which product categories have stronger profitability?
* Which customer segments contribute to revenue?

### Operational Analysis

* Which shipping modes are used most frequently?
* How many orders are returned?
* How does discounting relate to profitability?

---

## 🔄 Project Workflow

```text
Raw Excel Data
      ↓
Power Query
      ↓
Data Cleaning & Transformation
      ↓
Fact & Dimension Tables
      ↓
Power Pivot Data Model
      ↓
Relationships
      ↓
DAX Measures
      ↓
Pivot Tables
      ↓
Pivot Charts & Charts
      ↓
Slicers & Filters
      ↓
Interactive Excel Dashboard
```

---

## 🗂️ Dataset

The project uses the **Superstore dataset**.

### Raw Data Sources

The workbook contains the following raw and supporting tables:

| Table        | Description                     |
| ------------ | ------------------------------- |
| `RawOrders`  | Original order-level sales data |
| `RawReturns` | Returned order information      |
| `RawPeople`  | Regional manager information    |

### Main Fact Table

`FactOrders` contains the analytical order-level data with fields such as:

* Order ID
* Order Date
* Ship Date
* Ship Mode
* Customer ID
* Customer Name
* Segment
* Country
* City
* State
* Postal Code
* Region
* Product ID
* Category
* Sub-Category
* Product Name
* Sales
* Quantity
* Discount
* Profit
* Regional Manager
* Returned

---

## 🧩 Data Model

The project uses a **Star Schema-style data model** in Power Pivot.

### Fact Table

**FactOrders**

Contains transactional/order-level measures:

* Sales
* Quantity
* Discount
* Profit
* Returned

### Dimension Tables

**DimProducts**

* Product ID
* Product Name
* Category
* Sub-Category

**DimCustomers**

* Customer ID
* Customer Name
* Segment

**DimRegion**

* Region
* Country
* Regional Manager

**DimShipMode**

* Ship Mode

**DimDate**

* Date
* Year
* Quarter
* Month Number
* Month
* Weekday

**DimShipDate**

* Date
* Year
* Quarter
* Month Number
* Month
* Weekday

### Simplified Model

```text
                    DimDate
                       │
                       │
DimCustomers ───── FactOrders ───── DimProducts
                       │
                       │
                  DimRegion
                       │
                       │
                 DimShipMode

                  RawReturns
                      │
                      ↓
                  Returned
```

---

## 🧮 DAX Measures

The dashboard uses DAX measures to calculate important business KPIs.

### Total Sales

```DAX
Total Sales :=
SUM(FactOrders[Sales])
```

### Total Profit

```DAX
Total Profit :=
SUM(FactOrders[Profit])
```

### Total Quantity

```DAX
Total Quantity :=
SUM(FactOrders[Quantity])
```

### Total Orders

```DAX
Total Orders :=
DISTINCTCOUNT(FactOrders[Order ID])
```

### Total Customers

```DAX
Total Customers :=
DISTINCTCOUNT(FactOrders[Customer ID])
```

### Profit Margin

```DAX
Profit Margin :=
DIVIDE([Total Profit], [Total Sales], 0)
```

### Total Discount

```DAX
Total Discount :=
SUM(FactOrders[Discount])
```

> **Note:** The exact DAX implementation can be adjusted depending on the final Power Pivot model and whether certain calculations are implemented as measures, calculated columns, or Pivot calculations.

---

## 📈 Dashboard Features

The Excel dashboard includes interactive analytical features such as:

### KPI Cards

* 💰 **Total Sales**
* 📈 **Total Profit**
* 📊 **Profit Margin**

The current workbook contains approximately:

| KPI           |    Value |
| ------------- | -------: |
| Total Sales   |   $2.30M |
| Total Profit  | $286.40K |
| Profit Margin |   12.47% |

---

### 📊 Visualizations

The dashboard uses:

* Pivot Charts
* Standard Excel Charts
* Category analysis
* Regional analysis
* Sales & profit comparisons
* Profit margin analysis
* Time-based analysis

---

### 🎛️ Interactive Controls

Users can dynamically analyze the dashboard using:

* **Slicers**
* **Filters**
* Pivot Table filters
* Category selection
* Region selection
* Segment selection
* Time/date filtering

These controls allow users to change the analysis without modifying the underlying data.

---

## 🔍 Key Findings

Based on the current Superstore dataset:

### Overall Performance

* Total sales are approximately **$2.30 million**.
* Total profit is approximately **$286.40K**.
* Overall profit margin is approximately **12.47%**.
* The dataset contains approximately **5,009 unique orders**.
* There are approximately **793 unique customers**.
* The dataset covers orders from **2014 through 2017**.

### Category Performance

| Category        |    Sales |   Profit |
| --------------- | -------: | -------: |
| Technology      | $836.15K | $145.45K |
| Furniture       | $742.00K |  $18.45K |
| Office Supplies | $719.05K | $122.49K |

Technology generates the largest sales volume and also has the highest total profit among the three categories.

Furniture has substantial sales but a considerably lower profit contribution.

### Regional Performance

| Region  |    Sales |   Profit |
| ------- | -------: | -------: |
| West    | $725.46K | $108.42K |
| East    | $678.78K |  $91.52K |
| Central | $501.24K |  $39.71K |
| South   | $391.72K |  $46.75K |

The West region has the highest sales and profit contribution in the dataset.

### Sub-Category Observation

The analysis also highlights sub-categories with negative profitability.

For example:

* **Tables** have approximately **-$17.73K profit**.
* **Bookcases** have approximately **-$3.47K profit**.
* **Supplies** have approximately **-$1.19K profit**.

This demonstrates why analyzing **profit alongside sales** is important: high sales do not necessarily mean high profitability.

---

## 🖼️ Dashboard Screenshot

Add your exported dashboard screenshot to the repository, for example:

```text
screenshots/
└── superstore-dashboard.png
```

Then display it in this README:

```markdown
![Superstore Excel Dashboard](screenshots/superstore-dashboard.png)
```

### Dashboard Preview

> Replace the image path above with the actual screenshot filename in your GitHub repository.

---

## 📁 Repository Structure

Recommended GitHub repository structure:

```text
Superstore-Excel-Dashboard/
│
├── README.md
│
├── data/
│   └── Superstore_PowerQuery_DataModel.xlsx
│
├── screenshots/
│   └── superstore-dashboard.png
│
├── documentation/
│   └── data-model.png
│
└── outputs/
    └── dashboard-preview.png
```

---

## 🛠️ Tools & Technologies

| Tool / Technology   | Purpose                         |
| ------------------- | ------------------------------- |
| **Microsoft Excel** | Dashboard development           |
| **Power Query**     | Data cleaning & transformation  |
| **Power Pivot**     | Data modeling                   |
| **DAX**             | KPI and analytical calculations |
| **Pivot Tables**    | Data summarization              |
| **Pivot Charts**    | Interactive visualization       |
| **Excel Charts**    | Data visualization              |
| **Slicers**         | Interactive filtering           |
| **Filters**         | Data exploration                |

---

## 🔧 Data Preparation

Power Query was used to prepare the raw data before analysis.

Typical transformation steps include:

1. Import raw datasets.
2. Remove unnecessary columns.
3. Check data types.
4. Handle missing values.
5. Standardize column names.
6. Merge required datasets.
7. Append or transform tables where required.
8. Create structured tables.
9. Load transformed data into the Power Pivot model.

---

## 🔗 Data Relationships

The Power Pivot model connects transactional data with descriptive dimensions.

Example relationships:

```text
FactOrders[Product ID]
        ↓
DimProducts[Product ID]

FactOrders[Customer ID]
        ↓
DimCustomers[Customer ID]

FactOrders[Region]
        ↓
DimRegion[Region]

FactOrders[Ship Mode]
        ↓
DimShipMode[Ship Mode]

FactOrders[Order Date]
        ↓
DimDate[Date]
```

This structure makes it easier to perform dynamic calculations and slice the data across multiple dimensions.

---

## 📊 Analytical Approach

The project follows a complete **ETL → Data Model → DAX → Visualization** approach:

### 1. Extract

Raw Superstore datasets are imported into Excel.

### 2. Transform

Power Query cleans and prepares the data.

### 3. Model

Power Pivot is used to create fact and dimension tables.

### 4. Calculate

DAX measures calculate business KPIs.

### 5. Analyze

Pivot Tables and Pivot Charts summarize the data.

### 6. Visualize

Charts and KPI cards communicate the results.

### 7. Interact

Slicers and filters allow users to perform dynamic analysis.

---

## 🚀 How to Use the Dashboard

1. Download the Excel workbook from this repository.
2. Open it using **Microsoft Excel**.
3. Enable content/connections if Excel asks for permission.
4. Navigate to the **Superstore Dashboard** sheet.
5. Use the available slicers and filters.
6. Select different categories, regions, segments, or dates.
7. Review the KPI cards and charts.
8. Use Pivot Tables for detailed analysis.

---

## 💡 Business Value

This dashboard converts raw transactional data into an interactive reporting solution.

It can help users:

* Monitor sales performance
* Track profitability
* Identify loss-making product areas
* Compare regional performance
* Analyze customer segments
* Understand product performance
* Investigate discounts and profitability
* Support data-driven business analysis

---

## 📌 Skills Demonstrated

This project demonstrates practical skills in:

* Excel Data Analysis
* Power Query
* Power Pivot
* DAX
* Data Modeling
* Star Schema
* Data Cleaning
* ETL
* KPI Development
* Pivot Tables
* Pivot Charts
* Dashboard Design
* Business Analysis
* Data Visualization
* Interactive Reporting

---

## 👨‍💻 Project Author

**Prasad Kulkarni**

### Project: Superstore Sales & Profit Analysis Dashboard

**Tools:** Microsoft Excel | Power Query | Power Pivot | DAX | Pivot Charts | Slicers

---

## ⭐ Project Highlights

> **Raw Data → Power Query → Power Pivot → DAX → Pivot Tables → Charts → Interactive Dashboard**

This project demonstrates how Excel can be used as an end-to-end **Business Intelligence and Data Analytics tool** to transform raw business data into actionable insights.
