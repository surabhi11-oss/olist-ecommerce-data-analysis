# Olist E-Commerce Business Analytics & Executive Dashboard

An end-to-end Data Analytics project leveraging **MySQL 8.4**, **Python (Pandas & OpenPyXL)**, and multi-platform spreadsheet tools (**Excel, Google Sheets, Apple Numbers**) to analyze **99,441 orders** from the Olist Brazilian E-Commerce public dataset.

---

## 📌 Project Overview

This project provides a comprehensive quantitative analysis of Olist's Brazilian e-commerce platform operations between 2016 and 2018. It covers the full analytics lifecycle:
1. **Relational Database Design & ETL**: Schema creation with explicit data types, primary/foreign key constraints, indexes, and `LOAD DATA LOCAL INFILE` data ingestion.
2. **Advanced Business Analysis (SQL)**: 8 distinct analytical modules using Common Table Expressions (CTEs), window functions (`NTILE`, `ROW_NUMBER`, `LAG`, `SUM OVER`), multi-table joins, and temporal aggregations.
3. **Automated Excel Workbook Exporter (Python)**: A robust Python pipeline (`03_export_analysis.py`) that executes queries, computes MoM growth metrics, and exports a professionally styled 11-sheet Excel workbook (`olist_business_analysis.xlsx`).
4. **Cross-Platform Executive Dashboard**: A clean, Numbers-/Excel-/Sheets-compatible **Dashboard** tab featuring 5 KPI cards and 6 chart-ready summary data tables formatted for native spreadsheet charting.

---

## 🎯 Business Problems & Analytical Questions

The analysis answers critical operational and commercial questions for platform leadership:

* **Sales & Growth**: What is the historical trajectory of delivered revenue and order volume? What is the true month-over-month growth when accounting for calendar gaps?
* **Customer Geography & Retention**: Which Brazilian states and cities generate the highest revenue? What proportion of customers return for repeat purchases?
* **Category & Seller Concentration**: Which product categories drive merchandise revenue? Does seller revenue follow the Pareto Principle (80/20 rule)?
* **Logistics & Delivery Performance**: What is Olist's on-time delivery rate? How many days ahead or behind estimated delivery dates do orders arrive?
* **Customer Satisfaction**: How are review scores distributed? How drastically do shipping delays impact 1-star review rates?
* **Payment Preferences**: What payment methods do Brazilian consumers prefer? How do credit card installment choices impact average order payment values?

---

## 📊 Dataset Overview

The Olist dataset consists of 8 interconnected tables storing transaction records from 2016 to 2018:

| Table Name | Description | Record Count |
| :--- | :--- | :--- |
| `customers` | Customer locations and unique identifiers | 99,441 |
| `orders` | Order timestamps, status, and delivery dates | 99,441 |
| `order_items` | Products, prices, freight values, and sellers per order item | 112,650 |
| `payments` | Payment methods, installments, and payment values | 103,886 |
| `reviews` | Review scores (1–5 stars) and feedback timestamps | 99,223 |
| `products` | Product dimensions, weights, and category names | 32,951 |
| `sellers` | Seller locations and unique seller identifiers | 3,095 |
| `category_translation` | Translation of Portuguese category names to English | 71 |

---

## 🛠️ Tech Stack & Skills Demonstrated

* **Database & Querying**: MySQL 8.4, Relational Schema Design, Foreign Key Constraints, Indexes (`idx_orders_customer`, `idx_order_items_product`, etc.), Complex CTEs, Window Functions (`ROW_NUMBER`, `NTILE`, `LAG`, `SUM() OVER`), Conditional Aggregations.
* **Programming & Automation**: Python 3.14.0, Pandas (Data Manipulation, Date Periods, MoM Shift Logic), MySQL Connector Python.
* **Spreadsheet & Visualization**: OpenPyXL, Executive Styling (Hex Fills, Custom Typography, Borders), Cross-Platform Compatibility (Microsoft Excel, Google Sheets, Apple Numbers).
* **Data Analyst Intern Core Competencies**: Business Metrics Definition (AOV, MoM Growth, Pareto Distribution, On-Time Rates), Data Validation & Hygiene, Technical Documentation, Executive Reporting.

---

## 🧹 Data Hygiene & Validation Approach

To ensure statistical integrity and avoid common data traps, explicit cleaning rules were enforced across SQL queries and Python code:

1. **Delivered Orders Filtering**: Financial, logistics, and customer metrics filter strictly for `order_status = 'delivered'` with non-null customer delivery timestamps (`order_delivered_customer_date IS NOT NULL`), isolating completed commercial transactions from canceled or in-transit orders.
2. **Review Deduplication per Order-Category**: Joining `order_items` directly to `reviews` inflates review counts when an order contains multiple items. Query 7.3 utilizes a `SELECT DISTINCT order_id, category_name_english` CTE to ensure each customer review is counted exactly once per category.
3. **Order-Level Payment Aggregation**: Payments in Olist can be split across multiple vouchers or installments. Query 8.2 aggregates total payments to the `order_id` level before evaluating installment counts, preventing payment record granularity from distorting Average Order Payment Value.
4. **Calendar Gap Awareness**: The dataset lacks transaction data for certain calendar periods (e.g., 2016-11). MoM growth in Python converts timestamps to `pd.PeriodIndex(freq='M')` and verifies `prev_period == curr_period - 1` before calculating growth percentages, setting non-consecutive months to `NaN`/blank.

---

## 📈 Executive Dashboard & KPI Overview

The exported Excel workbook (`olist_business_analysis.xlsx`) places an executive **Dashboard** as its first tab (`index=0`), styled in an Executive Navy (`#1F497D`) and Soft Blue theme:

### Key Performance Indicators (KPI Cards)
* **Total Gross Revenue**: **$15,419,773.75** (Merchandise + Freight across delivered orders)
* **Total Delivered Orders**: **96,478**
* **Total Items Sold**: **110,197**
* **Average Order Value (AOV)**: **$159.83**
* **Average Review Score**: **4.09 / 5.00**

### Native Chart-Ready Data Tables
To ensure 100% cross-platform compatibility without rendering crashes in Apple Numbers or Google Sheets, the Dashboard avoids openpyxl chart drawing objects and provides clean, dedicated, chart-ready data tables:
1. `[ LINE CHART DATA ]` **Monthly Revenue Trend**: Chronological monthly revenue progression ($143.46 in Sep 2016 to peak of $1,128,774.52 in May 2018).
2. `[ COLUMN CHART DATA ]` **Monthly Orders Volume**: Monthly order counts (scaling from 1 order to 6,749 orders/month).
3. `[ DONUT CHART DATA ]` **Delivery Performance**: On-Time / Early delivery share (93.23%) vs Late delivery share (6.77%).
4. `[ BAR CHART DATA ]` **Top 10 Product Categories by Revenue**: Top revenue-generating product categories led by `health_beauty` ($1.23M) and `watches_gifts` ($1.17M).
5. `[ PIE / DONUT DATA ]` **Payment Method Mix**: Distribution of payment methods led by Credit Card (78.34%) and Boleto (17.92%).
6. `[ COLUMN CHART DATA ]` **Customer Review Score Distribution**: 5-star (57.78%), 4-star (19.29%), 3-star (8.24%), 2-star (3.18%), 1-star (11.51%).

---

## 🔍 Key Quantified Business Insights

Based strictly on empirical findings from the SQL queries and exported data:

1. **Revenue Growth & Scale**: Delivered gross revenue expanded significantly from **$46,490.66** in October 2016 to a peak of **$1,128,774.52** in May 2018 across **96,478 delivered orders**, achieving an overall platform Average Order Value (AOV) of **$159.83**.
2. **Seller Revenue Concentration (Pareto Principle)**: Merchandise revenue is heavily concentrated among top sellers. The top **20% of active sellers (594 sellers)** generate **82.29% ($10,879,659.64)** of total merchandise revenue, whereas the bottom **80% (2,376 sellers)** contribute only **17.71% ($2,341,838.47)**.
3. **Consumer Payment Preferences**: Credit card is the dominant payment channel in Brazil, capturing **78.34% ($12,542,084.19)** of total payment value across 76,505 transactions. Bank slips (*Boleto*) represent the second largest channel at **17.92% ($2,869,361.27)**, while Vouchers (2.37%) and Debit Cards (1.36%) account for minor shares.
4. **Logistics & Delivery Lead Times**: Olist maintains a strong **93.23% on-time delivery rate (89,936 orders)**, with packages arriving an average of **12.5 days ahead** of the estimated delivery date. However, **6.77% of orders (6,534 orders)** experienced delays, arriving an average of **10.6 days late**.
5. **Impact of Delivery Delays on Customer Satisfaction**: Shipping delays severely penalize customer reviews. On-time orders maintain an average review score of **4.15 stars** with only **8.63% 1-star ratings**. In contrast, delayed orders suffer a steep drop to an average of **2.25 stars**, with **53.64% of delayed orders receiving a 1-star review**.
6. **Category Performance Concentration**: The top 5 product categories account for **$5,266,320.62** (~34%) of total merchandise revenue, led by `health_beauty` (**$1,233,131.72** across 8,647 orders) and `watches_gifts` (**$1,166,176.98** across 5,495 orders).
7. **Customer Purchase Frequency & Cohort Retention**: Approximately **97% of unique customers** in the dataset made only 1 purchase during the observed 2016–2018 window, indicating that Olist's historical revenue volume was predominantly driven by new customer acquisition rather than repeat buyers.

---

## 📌 Key Data-Quality & Analytical Assumptions

To ensure complete transparency, the analytical choices and data assumptions are explicitly documented below:

* **Delivered Status Baseline**: Orders with status `canceled`, `unavailable`, `invoiced`, or `processing` are excluded from financial and delivery metrics unless explicitly analyzing status distributions (e.g., Query 1.2).
* **First Purchase Definition**: Customer cohort analysis (Query 3.4) evaluates a customer's first *delivered* purchase date (`MIN(order_purchase_timestamp)` where `order_status = 'delivered'`), rather than absolute platform registration.
* **Category Review Threshold**: When ranking highest and lowest rated product categories (Query 7.3), a minimum sample size threshold of **100 reviews** is required to filter out low-volume statistical noise.
* **Payment Installment Values**: Installment analysis aggregates credit card payments to the `order_id` level first to measure actual order payment amounts rather than split payment line items.

---

## 📁 Project Structure

```
olist-project/
├── 01_database_setup.sql        # Database initialization, table DDL, constraints, indexing, LOAD DATA INFILE
├── 02_business_analysis.sql     # 8 modules of advanced MySQL business queries (CTEs, Window Functions)
├── 03_export_analysis.py        # Python script: MySQL execution, Pandas MoM calculations, OpenPyXL Excel export
├── olist_business_analysis.xlsx # Formatted 11-sheet Excel workbook (Dashboard + 10 Analysis Tabs)
├── README.md                    # Comprehensive project documentation
└── dataset/                     # Olist CSV data files (8 tables)
    ├── olist_customers_dataset.csv
    ├── olist_geolocation_dataset.csv
    ├── olist_order_items_dataset.csv
    ├── olist_order_payments_dataset.csv
    ├── olist_order_reviews_dataset.csv
    ├── olist_orders_dataset.csv
    ├── olist_products_dataset.csv
    ├── olist_sellers_dataset.csv
    └── product_category_name_translation.csv
```

---

## 🚀 How to Reproduce the Analysis

### Prerequisites
* **MySQL Server 8.4+**
* **Python 3.14.0**
* Required Python packages:
  ```bash
  pip install mysql-connector-python pandas openpyxl
  ```

### Step 1: Database Setup & Data Ingestion
1. Start your local MySQL service.
2. Open MySQL CLI or workbench and execute `sql/01_database_setup.sql`:
   ```bash
   mysql -u root -p < sql/01_database_setup.sql
   ```
   *Note: Ensure the CSV file paths in `LOAD DATA LOCAL INFILE` match your local directory.*

### Step 2: Execute SQL Business Analysis
To run the SQL analysis directly in MySQL:
```bash
mysql -u root -p olist_analysis < sql/02_business_analysis.sql
```

### Step 3: Export Excel Workbook & Dashboard
Run the Python exporter script:
```bash
python3 03_export_analysis.py
```
When prompted, enter your MySQL root password. The script will output `olist_business_analysis.xlsx` containing the Executive **Dashboard** tab and 10 detailed analysis worksheets.

---

## 💼 Skills & Competencies Demonstrated (Data Analyst Intern)

* **SQL Mastery**: Writing clean, performant, readable SQL queries using CTEs, window functions (`ROW_NUMBER`, `NTILE`, `LAG`), aggregations, and multi-table joins.
* **Data Pipelines**: Building robust Python ETL scripts connecting relational databases to reporting outputs.
* **Business Acumen**: Translating raw transactional data into actionable commercial insights across revenue growth, logistics, customer satisfaction, and seller dynamics.
* **Executive Presentation**: Designing structured, cross-platform compatible dashboards with intuitive KPI cards and clear chart-ready data layouts.
* **Data Rigor & Validation**: Identifying data traps (duplicate review counts, missing calendar months, split payments) and implementing strict analytical controls.
