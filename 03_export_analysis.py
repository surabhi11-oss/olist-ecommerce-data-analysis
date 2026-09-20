#!/usr/bin/env python3
"""
Olist E-Commerce Business Analysis Exporter
-------------------------------------------
Connects to the `olist_analysis` MySQL database, executes analytical queries
from `sql/02_business_analysis.sql`, and exports the formatted results into
an Excel workbook named `olist_business_analysis.xlsx`.

Requirements:
    - mysql-connector-python
    - pandas
    - openpyxl

Usage:
    python3 03_export_analysis.py
"""

import sys
import os
import getpass
import pandas as pd
import mysql.connector
from mysql.connector import Error
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, DoughnutChart, Reference


def get_db_connection():
    """Prompt user for MySQL password and establish a database connection."""
    print("=================================================================")
    print(" Olist E-Commerce Database Exporter - MySQL to Excel")
    print("=================================================================")
    host = "localhost"
    port = 3306
    user = "root"
    database = "olist_analysis"
    
    password = getpass.getpass(prompt=f"Enter MySQL password for user '{user}': ")
    
    try:
        conn = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database
        )
        if conn.is_connected():
            print(f"[SUCCESS] Connected to MySQL database '{database}'.\n")
            return conn
    except Error as e:
        print(f"\n[ERROR] Failed to connect to MySQL database: {e}")
        sys.exit(1)


def execute_query_to_df(conn, query, query_name, column_labels=None):
    """Execute a SELECT query via MySQL cursor directly and return a formatted DataFrame."""
    try:
        cursor = conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        col_names = [desc[0] for desc in cursor.description]
        cursor.close()
        
        df = pd.DataFrame(rows, columns=col_names)
        
        if df.empty:
            print(f" [INFO] '{query_name}' returned 0 rows.")
        else:
            print(f" [SUCCESS] '{query_name}' fetched {len(df)} rows.")
            
        if column_labels and not df.empty:
            if len(column_labels) == len(df.columns):
                df.columns = column_labels
            else:
                print(f" [WARNING] Column mismatch for '{query_name}': Expected {len(column_labels)}, got {len(df.columns)}.")
                
        return df
    except Error as e:
        print(f" [ERROR] Query '{query_name}' MySQL Error: {e}")
        return pd.DataFrame()
    except Exception as e:
        print(f" [ERROR] Query '{query_name}' Unexpected Error: {e}")
        return pd.DataFrame()


def format_excel_worksheet(ws, freeze_row=2):
    """Apply professional formatting, frozen headers, and auto-adjusted column widths."""
    ws.freeze_panes = f"A{freeze_row}"
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and not cell.border.left.style:
                cell.border = thin_border

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


def append_df_to_worksheet(ws, df, start_row, section_title=None):
    """Append a DataFrame with section header styling to an openpyxl worksheet."""
    current_row = start_row
    
    # Write Section Header if provided
    if section_title:
        ws.cell(row=current_row, column=1, value=section_title)
        cell = ws.cell(row=current_row, column=1)
        cell.font = Font(name="Calibri", size=12, bold=True, color="1F497D")
        cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        current_row += 1

    if df.empty:
        ws.cell(row=current_row, column=1, value="[No data returned for this query or query failed]")
        return current_row + 2

    # Write Table Column Headers
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_num, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=current_row, column=col_num, value=str(col_name))
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    current_row += 1

    # Write Data Rows
    data_font = Font(name="Calibri", size=11)
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )
    
    for _, row in df.iterrows():
        for col_num, val in enumerate(row, 1):
            cell = ws.cell(row=current_row, column=col_num, value=val)
            cell.font = data_font
            cell.border = thin_border
            # Format numbers
            if isinstance(val, (int, float)):
                if isinstance(val, float):
                    cell.number_format = '#,##0.00'
                else:
                    cell.number_format = '#,##0'
        current_row += 1

    return current_row + 2  # Leave 2 blank rows between tables


def append_dashboard_table(ws, df, start_row, start_col, title, chart_type_label=None, highlight_top=False):
    """Write a visually polished, chart-ready data table onto the Dashboard worksheet for native charting."""
    if df is None or df.empty:
        return start_row
    
    # Subheader / Title for the summary data section with visual chart type tag
    if chart_type_label:
        badge_cell = ws.cell(row=start_row, column=start_col, value=f"[ {chart_type_label} ]  {title}")
        badge_cell.font = Font(name="Calibri", size=11, bold=True, color="1F497D")
    else:
        ws.cell(row=start_row, column=start_col, value=title).font = Font(name="Calibri", size=11, bold=True, color="1F497D")
    start_row += 1
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Table Headers
    for col_idx, col_name in enumerate(df.columns):
        cell = ws.cell(row=start_row, column=start_col + col_idx, value=str(col_name))
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    start_row += 1
    
    # Alternate Row Fill & Top Row Highlight Fill
    even_row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    top_row_fill = PatternFill(start_color="EBF2FA", end_color="EBF2FA", fill_type="solid")
    data_font = Font(name="Calibri", size=10)
    bold_data_font = Font(name="Calibri", size=10, bold=True, color="1F497D")
    
    # Find max numeric value for top highlighting if requested
    max_val_row_idx = -1
    if highlight_top and len(df) > 0:
        for c_i, col_c in enumerate(df.columns):
            if pd.api.types.is_numeric_dtype(df[col_c]):
                max_val_row_idx = df[col_c].idxmax()
                break

    # Table Data Rows
    for r_idx, row in df.iterrows():
        is_top = (r_idx == max_val_row_idx)
        for col_idx, val in enumerate(row):
            cell = ws.cell(row=start_row, column=start_col + col_idx, value=val)
            cell.font = bold_data_font if is_top else data_font
            cell.border = thin_border
            
            if is_top:
                cell.fill = top_row_fill
            elif r_idx % 2 == 1:
                cell.fill = even_row_fill
            
            # Format based on column name and value type
            col_name_str = str(df.columns[col_idx]).lower()
            if isinstance(val, (int, float)):
                if any(kw in col_name_str for kw in ["revenue", "spend", "value", "price"]):
                    cell.number_format = '$#,##0.00'
                elif any(kw in col_name_str for kw in ["percentage", "share", "pct", "%"]):
                    cell.number_format = '0.00"%"'
                elif isinstance(val, float):
                    cell.number_format = '#,##0.00'
                else:
                    cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        start_row += 1
        
    return start_row


def build_dashboard_sheet(wb, kpi_data, dashboard_dfs=None):
    """Create a polished, executive-level Dashboard as the first worksheet (Excel & Google Sheets compatible)."""
    ws = wb.create_sheet(title="Dashboard", index=0)
    ws.views.sheetView[0].showGridLines = True
    
    # -------------------------------------------------------------------------
    # 1. Header Banner
    # -------------------------------------------------------------------------
    ws.merge_cells("A1:O1")
    title_cell = ws["A1"]
    title_cell.value = "Olist E-Commerce Business Analytics Dashboard"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:O2")
    sub_cell = ws["A2"]
    sub_cell.value = "Brazilian E-Commerce | 2016–2018 Performance Summary"
    sub_cell.font = Font(name="Calibri", size=11, italic=True, color="DCE6F1")
    sub_cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # -------------------------------------------------------------------------
    # 2. KPI Cards Section
    # -------------------------------------------------------------------------
    ws.cell(row=4, column=1, value="KEY PERFORMANCE INDICATORS (KPIs)").font = Font(name="Calibri", size=12, bold=True, color="1F497D")
    
    kpi_cards = [
        {"title": "Total Revenue", "val": kpi_data["total_revenue"], "fmt": '$#,##0.00,,"M"', "col": "B"},
        {"title": "Total Orders", "val": kpi_data["total_orders"], "fmt": "#,##0", "col": "E"},
        {"title": "Total Items Sold", "val": kpi_data["total_items"], "fmt": "#,##0", "col": "H"},
        {"title": "Average Order Value", "val": kpi_data["aov"], "fmt": "$#,##0.00", "col": "K"},
        {"title": "Average Review Score", "val": kpi_data["avg_review"], "fmt": "0.00", "col": "N"},
    ]

    card_title_font = Font(name="Calibri", size=10, bold=True, color="595959")
    card_val_font = Font(name="Calibri", size=15, bold=True, color="1F497D")
    card_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    card_border = Border(
        left=Side(style='medium', color='1F497D'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='medium', color='1F497D'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for card in kpi_cards:
        c = card["col"]
        # Header cell
        ws[f"{c}5"] = card["title"]
        ws[f"{c}5"].font = card_title_font
        ws[f"{c}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c}5"].fill = card_fill
        ws[f"{c}5"].border = card_border

        # Value cell
        ws[f"{c}6"] = card["val"]
        ws[f"{c}6"].font = card_val_font
        ws[f"{c}6"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c}6"].fill = card_fill
        ws[f"{c}6"].number_format = card["fmt"]
        ws[f"{c}6"].border = card_border

    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 28

    # -------------------------------------------------------------------------
    # 3. OpenPyXL Embedded Visual Charts Section (2x2 Grid)
    # -------------------------------------------------------------------------
    ws.cell(row=8, column=1, value="EXECUTIVE VISUAL ANALYTICS").font = Font(name="Calibri", size=12, bold=True, color="1F497D")

    ws_sales_monthly = wb["Sales_Monthly"]
    ws_categories = wb["Categories"]
    ws_payments = wb["Payments"]
    ws_reviews = wb["Reviews"]

    # Chart 1: Monthly Revenue Trend (LineChart)
    chart1 = LineChart()
    chart1.title = "Monthly Revenue Trend ($)"
    chart1.style = 13
    chart1.legend = None  # Single series - remove unnecessary legend
    chart1.y_axis.title = "Gross Revenue ($)"
    chart1.x_axis.title = "Month"
    chart1.width = 15
    chart1.height = 9.5
    
    num_rows_monthly = ws_sales_monthly.max_row
    if num_rows_monthly > 1:
        data1 = Reference(ws_sales_monthly, min_col=6, min_row=1, max_row=num_rows_monthly)
        cats1 = Reference(ws_sales_monthly, min_col=1, min_row=2, max_row=num_rows_monthly)
        chart1.add_data(data1, titles_from_data=True)
        chart1.set_categories(cats1)
        ws.add_chart(chart1, "B10")

    # Chart 2: Top 10 Product Categories by Revenue (BarChart - Horizontal)
    chart2 = BarChart()
    chart2.type = "bar"
    chart2.style = 11
    chart2.legend = None  # Single series - remove unnecessary legend
    chart2.title = "Top 10 Product Categories by Revenue"
    chart2.x_axis.title = "Total Revenue ($)"
    chart2.y_axis.title = "Category"
    chart2.width = 15
    chart2.height = 9.5
    
    num_rows_cats = min(ws_categories.max_row, 11)
    if num_rows_cats > 1:
        data2 = Reference(ws_categories, min_col=2, min_row=1, max_row=num_rows_cats)
        cats2 = Reference(ws_categories, min_col=1, min_row=2, max_row=num_rows_cats)
        chart2.add_data(data2, titles_from_data=True)
        chart2.set_categories(cats2)
        ws.add_chart(chart2, "I10")

    # Chart 3: Payment Method Breakdown (DoughnutChart)
    chart3 = DoughnutChart()
    chart3.title = "Payment Method Breakdown"
    chart3.width = 15
    chart3.height = 9.5
    
    data3 = Reference(ws_payments, min_col=2, min_row=2, max_row=7)
    cats3 = Reference(ws_payments, min_col=1, min_row=3, max_row=7)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    ws.add_chart(chart3, "B26")

    # Chart 4: Customer Review Score Distribution (BarChart - Column)
    chart4 = BarChart()
    chart4.type = "col"
    chart4.style = 12
    chart4.legend = None  # Single series - remove unnecessary legend
    chart4.title = "Customer Review Score Distribution"
    chart4.y_axis.title = "Review Count"
    chart4.x_axis.title = "Review Score (Stars)"
    chart4.width = 15
    chart4.height = 9.5
    
    data4 = Reference(ws_reviews, min_col=2, min_row=2, max_row=7)
    cats4 = Reference(ws_reviews, min_col=1, min_row=3, max_row=7)
    chart4.add_data(data4, titles_from_data=True)
    chart4.set_categories(cats4)
    ws.add_chart(chart4, "I26")

    # -------------------------------------------------------------------------
    # 4. Key Business Insights Section (Rows 42+)
    # -------------------------------------------------------------------------
    ws.cell(row=42, column=1, value="KEY BUSINESS INSIGHTS").font = Font(name="Calibri", size=12, bold=True, color="1F497D")

    insights = [
        ("Payment Preferences", "Credit cards account for 78.34% of payment transactions, followed by boleto at 17.92%."),
        ("Top Category", "Health & Beauty is the highest-revenue product category at approximately $1.23M across 8,647 orders."),
        ("Logistics Lead Time", "93.23% of delivered orders arrived on time or early, delivering an average of 12.5 days ahead of estimate."),
        ("Customer Feedback", "57.78% of customer reviews are 5-star, achieving an overall platform average review score of 4.09 stars out of 5.00."),
        ("Shipping Impact", "Delayed orders experience substantially lower review scores (2.25 avg) compared to on-time orders (4.15 avg)."),
    ]

    badge_font = Font(name="Calibri", size=10, bold=True, color="1F497D")
    text_font = Font(name="Calibri", size=10, color="333333")
    insight_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    insight_border = Border(
        left=Side(style='medium', color='1F497D'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    current_row = 44
    for topic, desc in insights:
        # Topic badge cell (Col B)
        cell_b = ws.cell(row=current_row, column=2, value=topic)
        cell_b.font = badge_font
        cell_b.fill = insight_fill
        cell_b.alignment = Alignment(horizontal="left", vertical="center")
        cell_b.border = insight_border

        # Description merged cell (Cols C to N)
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=14)
        cell_c = ws.cell(row=current_row, column=3, value=desc)
        cell_c.font = text_font
        cell_c.fill = insight_fill
        cell_c.alignment = Alignment(horizontal="left", vertical="center")
        
        # Apply borders across merged cells
        for col_i in range(3, 15):
            ws.cell(row=current_row, column=col_i).border = insight_border
            ws.cell(row=current_row, column=col_i).fill = insight_fill

        ws.row_dimensions[current_row].height = 22
        current_row += 1

    # Set column widths for Dashboard layout
    ws.column_dimensions['A'].width = 3
    for c in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O']:
        ws.column_dimensions[c].width = 14


def main():
    conn = get_db_connection()
    output_filename = "olist_business_analysis.xlsx"
    output_filepath = os.path.abspath(output_filename)

    # Initialize Excel Workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    print("[INFO] Processing and exporting business analysis queries...\n")

    # =========================================================================
    # SHEET 1: Overview
    # =========================================================================
    print("Processing Sheet 1: Overview...")
    ws_overview = wb.create_sheet(title="Overview")
    
    q1_1 = """
    SELECT 'customers' AS table_name, COUNT(*) AS row_count FROM customers
    UNION ALL SELECT 'orders', COUNT(*) FROM orders
    UNION ALL SELECT 'order_items', COUNT(*) FROM order_items
    UNION ALL SELECT 'payments', COUNT(*) FROM payments
    UNION ALL SELECT 'reviews', COUNT(*) FROM reviews
    UNION ALL SELECT 'products', COUNT(*) FROM products
    UNION ALL SELECT 'sellers', COUNT(*) FROM sellers
    UNION ALL SELECT 'category_translation', COUNT(*) FROM category_translation;
    """
    df1_1 = execute_query_to_df(conn, q1_1, "Query 1.1 Table Record Counts", 
                                ["Table Name", "Total Record Count"])

    q1_2 = """
    SELECT 
        order_status,
        COUNT(*) AS total_orders,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS status_percentage
    FROM orders
    GROUP BY order_status
    ORDER BY total_orders DESC;
    """
    df1_2 = execute_query_to_df(conn, q1_2, "Query 1.2 Order Status Breakdown", 
                                ["Order Status", "Total Orders", "Status Percentage (%)"])

    q1_3 = """
    SELECT 
        MIN(order_purchase_timestamp) AS earliest_order_date,
        MAX(order_purchase_timestamp) AS latest_order_date,
        TIMESTAMPDIFF(DAY, MIN(order_purchase_timestamp), MAX(order_purchase_timestamp)) AS dataset_span_days
    FROM orders;
    """
    df1_3 = execute_query_to_df(conn, q1_3, "Query 1.3 Date Coverage Range", 
                                ["Earliest Order Date", "Latest Order Date", "Dataset Time Span (Days)"])

    r = append_df_to_worksheet(ws_overview, df1_1, start_row=1, section_title="1. Core Table Record Counts")
    r = append_df_to_worksheet(ws_overview, df1_2, start_row=r, section_title="2. Order Status Breakdown")
    r = append_df_to_worksheet(ws_overview, df1_3, start_row=r, section_title="3. Dataset Date Horizon")
    format_excel_worksheet(ws_overview, freeze_row=2)

    # =========================================================================
    # SHEET 2: Sales_Monthly
    # =========================================================================
    print("Processing Sheet 2: Sales_Monthly...")
    ws_sales_monthly = wb.create_sheet(title="Sales_Monthly")
    
    q2_1 = """
    SELECT
        YEAR(o.order_purchase_timestamp) AS order_year,
        MONTH(o.order_purchase_timestamp) AS order_month,
        COUNT(DISTINCT o.order_id) AS delivered_orders,
        COUNT(oi.order_item_id) AS total_items,
        ROUND(SUM(oi.price), 2) AS merchandise_revenue,
        ROUND(SUM(oi.freight_value), 2) AS freight_revenue,
        ROUND(SUM(oi.price + oi.freight_value), 2) AS gross_revenue,
        ROUND(
            SUM(oi.price + oi.freight_value) /
            COUNT(DISTINCT o.order_id),
            2
        ) AS average_order_value
    FROM orders o
    INNER JOIN order_items oi
        ON o.order_id = oi.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY
        YEAR(o.order_purchase_timestamp),
        MONTH(o.order_purchase_timestamp)
    ORDER BY
        YEAR(o.order_purchase_timestamp),
        MONTH(o.order_purchase_timestamp);
    """
    df2_1_raw = execute_query_to_df(conn, q2_1, "Query 2.1 Monthly Revenue")
    
    if not df2_1_raw.empty:
        # Create year_month column in Python
        df2_1_raw["year_month"] = (
            df2_1_raw["order_year"].astype(str)
            + "-"
            + df2_1_raw["order_month"].astype(str).str.zfill(2)
        )
        
        # Prepare df2_1 for Sales_Monthly sheet export
        df2_1_export = df2_1_raw[
            [
                "year_month",
                "delivered_orders",
                "total_items",
                "merchandise_revenue",
                "freight_revenue",
                "gross_revenue",
                "average_order_value",
            ]
        ].copy()
        df2_1_export.columns = [
            "Month",
            "Delivered Orders",
            "Total Items",
            "Merchandise Revenue ($)",
            "Freight Revenue ($)",
            "Gross Revenue ($)",
            "Average Order Value ($)",
        ]
    else:
        df2_1_export = pd.DataFrame()
    
    append_df_to_worksheet(ws_sales_monthly, df2_1_export, start_row=1)
    format_excel_worksheet(ws_sales_monthly, freeze_row=2)

    # =========================================================================
    # SHEET 3: Sales_Growth (Calculated in Python from Monthly Revenue)
    # =========================================================================
    print("Processing Sheet 3: Sales_Growth...")
    ws_sales_growth = wb.create_sheet(title="Sales_Growth")
    
    if not df2_1_raw.empty:
        df_growth = df2_1_raw[["year_month", "gross_revenue"]].copy()
        df_growth = df_growth.rename(columns={"gross_revenue": "monthly_revenue"})
        
        # Convert year_month to Monthly Period to verify consecutive calendar months
        df_growth["period"] = pd.PeriodIndex(df_growth["year_month"], freq="M")
        df_growth = df_growth.sort_values("period").reset_index(drop=True)
        
        # Check if the previous row is the immediately preceding calendar month (period - 1)
        prev_period = df_growth["period"].shift(1)
        expected_prev_period = df_growth["period"] - 1
        is_consecutive = (prev_period == expected_prev_period)
        
        # Shift revenue only where consecutive, else NaN
        raw_prev_revenue = df_growth["monthly_revenue"].shift(1)
        df_growth["previous_month_revenue"] = raw_prev_revenue.where(is_consecutive, None)
        
        # Calculate MoM growth percentage only where consecutive
        growth_calc = (
            (df_growth["monthly_revenue"] - df_growth["previous_month_revenue"])
            / df_growth["previous_month_revenue"]
            * 100
        ).round(2)
        
        df_growth["month_over_month_growth_pct"] = growth_calc.where(is_consecutive, None)
        
        df2_2_export = df_growth[
            ["year_month", "monthly_revenue", "previous_month_revenue", "month_over_month_growth_pct"]
        ].copy()
        df2_2_export.columns = [
            "Month",
            "Monthly Revenue ($)",
            "Previous Month Revenue ($)",
            "MoM Growth (%)",
        ]
        print(f" [SUCCESS] Calculated Sales_Growth in Python for {len(df2_2_export)} monthly rows.")
    else:
        df2_2_export = pd.DataFrame()
        print(" [WARNING] Cannot calculate Sales_Growth because Monthly Revenue returned 0 rows.")

    append_df_to_worksheet(ws_sales_growth, df2_2_export, start_row=1)
    format_excel_worksheet(ws_sales_growth, freeze_row=2)

    # =========================================================================
    # SHEET 4: Customers
    # =========================================================================
    print("Processing Sheet 4: Customers...")
    ws_customers = wb.create_sheet(title="Customers")
    
    q3_1a = """
    SELECT 
        c.customer_state AS state,
        COUNT(DISTINCT c.customer_unique_id) AS unique_customers,
        COUNT(DISTINCT o.order_id) AS total_orders,
        ROUND(SUM(oi.price + oi.freight_value), 2) AS total_spend,
        ROUND(SUM(oi.price + oi.freight_value) / COUNT(DISTINCT c.customer_unique_id), 2) AS spend_per_customer
    FROM customers c
    JOIN orders o ON c.customer_id = o.customer_id
    JOIN order_items oi ON o.order_id = oi.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY c.customer_state
    ORDER BY total_spend DESC
    LIMIT 10;
    """
    df3_1a = execute_query_to_df(conn, q3_1a, "Query 3.1a Top 10 Customer States", 
                                 ["Customer State", "Unique Customers", "Total Orders", "Total Spend ($)", "Spend Per Customer ($)"])

    q3_1b = """
    SELECT 
        c.customer_city AS city,
        c.customer_state AS state,
        COUNT(DISTINCT c.customer_unique_id) AS unique_customers,
        COUNT(DISTINCT o.order_id) AS total_orders,
        ROUND(SUM(oi.price + oi.freight_value), 2) AS total_spend,
        ROUND(SUM(oi.price + oi.freight_value) / COUNT(DISTINCT c.customer_unique_id), 2) AS spend_per_customer
    FROM customers c
    JOIN orders o ON c.customer_id = o.customer_id
    JOIN order_items oi ON o.order_id = oi.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY c.customer_city, c.customer_state
    ORDER BY total_spend DESC
    LIMIT 10;
    """
    df3_1b = execute_query_to_df(conn, q3_1b, "Query 3.1b Top 10 Customer Cities", 
                                 ["Customer City", "State", "Unique Customers", "Total Orders", "Total Spend ($)", "Spend Per Customer ($)"])

    r = append_df_to_worksheet(ws_customers, df3_1a, start_row=1, section_title="1. Top 10 Customer States by Revenue")
    r = append_df_to_worksheet(ws_customers, df3_1b, start_row=r, section_title="2. Top 10 Customer Cities by Revenue")
    format_excel_worksheet(ws_customers, freeze_row=2)

    # =========================================================================
    # SHEET 5: Customer_Segments
    # =========================================================================
    print("Processing Sheet 5: Customer_Segments...")
    ws_cust_seg = wb.create_sheet(title="Customer_Segments")
    
    q3_2 = """
    WITH customer_order_counts AS (
        SELECT 
            c.customer_unique_id,
            COUNT(DISTINCT o.order_id) AS order_count
        FROM customers c
        JOIN orders o ON c.customer_id = o.customer_id
        WHERE o.order_status = 'delivered'
        GROUP BY c.customer_unique_id
    )
    SELECT 
        CASE 
            WHEN order_count = 1 THEN '1 Order (One-time)'
            WHEN order_count = 2 THEN '2 Orders'
            WHEN order_count = 3 THEN '3 Orders'
            ELSE '4+ Orders (Power Buyers)'
        END AS purchase_frequency_tier,
        COUNT(*) AS customer_count,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage_of_customers
    FROM customer_order_counts
    GROUP BY 
        CASE 
            WHEN order_count = 1 THEN '1 Order (One-time)'
            WHEN order_count = 2 THEN '2 Orders'
            WHEN order_count = 3 THEN '3 Orders'
            ELSE '4+ Orders (Power Buyers)'
        END
    ORDER BY customer_count DESC;
    """
    df3_2 = execute_query_to_df(conn, q3_2, "Query 3.2 Customer Purchase Tiers", 
                                ["Purchase Frequency Tier", "Customer Count", "Percentage of Customers (%)"])

    q3_3 = """
    WITH max_dataset_date AS (
        SELECT MAX(order_purchase_timestamp) AS ref_date FROM orders
    ),
    customer_rfm_raw AS (
        SELECT 
            c.customer_unique_id,
            TIMESTAMPDIFF(DAY, MAX(o.order_purchase_timestamp), m.ref_date) AS recency_days,
            COUNT(DISTINCT o.order_id) AS frequency,
            SUM(oi.price + oi.freight_value) AS monetary_val
        FROM customers c
        JOIN orders o ON c.customer_id = o.customer_id
        JOIN order_items oi ON o.order_id = oi.order_id
        CROSS JOIN max_dataset_date m
        WHERE o.order_status = 'delivered'
        GROUP BY c.customer_unique_id, m.ref_date
    )
    SELECT 
        customer_unique_id,
        recency_days,
        frequency,
        ROUND(monetary_val, 2) AS monetary_value,
        NTILE(4) OVER (ORDER BY recency_days ASC) AS r_quartile,
        NTILE(4) OVER (ORDER BY monetary_val DESC) AS m_quartile
    FROM customer_rfm_raw
    ORDER BY monetary_val DESC
    LIMIT 20;
    """
    df3_3 = execute_query_to_df(conn, q3_3, "Query 3.3 Top 20 RFM Quartiles", 
                                ["Customer Unique ID", "Recency (Days)", "Frequency (Orders)", "Monetary Spend ($)", "Recency Quartile", "Monetary Quartile"])

    r = append_df_to_worksheet(ws_cust_seg, df3_2, start_row=1, section_title="1. One-Time vs Repeat Purchase Tiers")
    r = append_df_to_worksheet(ws_cust_seg, df3_3, start_row=r, section_title="2. Top 20 Customer RFM Segmentation Quartiles")
    format_excel_worksheet(ws_cust_seg, freeze_row=2)

    # =========================================================================
    # SHEET 6: Categories
    # =========================================================================
    print("Processing Sheet 6: Categories...")
    ws_categories = wb.create_sheet(title="Categories")
    
    q4_1 = """
    SELECT 
        COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized') AS category,
        ROUND(SUM(oi.price), 2) AS revenue,
        COUNT(oi.order_item_id) AS items_sold,
        COUNT(DISTINCT oi.order_id) AS order_count,
        ROUND(AVG(oi.price), 2) AS average_item_price,
        ROUND(AVG(oi.freight_value), 2) AS average_freight_cost
    FROM order_items oi
    JOIN products p ON oi.product_id = p.product_id
    LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
    JOIN orders o ON oi.order_id = o.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized')
    ORDER BY revenue DESC
    LIMIT 15;
    """
    df4_1 = execute_query_to_df(conn, q4_1, "Query 4.1 Category Performance", 
                                ["Category Name (English)", "Total Revenue ($)", "Items Sold", "Order Count", "Average Item Price ($)", "Average Freight Cost ($)"])
    append_df_to_worksheet(ws_categories, df4_1, start_row=1)
    format_excel_worksheet(ws_categories, freeze_row=2)

    # =========================================================================
    # SHEET 7: Sellers
    # =========================================================================
    print("Processing Sheet 7: Sellers...")
    ws_sellers = wb.create_sheet(title="Sellers")
    
    q5_1 = """
    SELECT 
        s.seller_id,
        s.seller_state AS state,
        s.seller_city AS city,
        COUNT(DISTINCT oi.order_id) AS fulfilled_orders,
        COUNT(oi.order_item_id) AS items_sold,
        ROUND(SUM(oi.price), 2) AS total_merchandise_revenue,
        ROUND(AVG(oi.price), 2) AS avg_item_price
    FROM sellers s
    JOIN order_items oi ON s.seller_id = oi.seller_id
    JOIN orders o ON oi.order_id = o.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY s.seller_id, s.seller_state, s.seller_city
    ORDER BY total_merchandise_revenue DESC
    LIMIT 10;
    """
    df5_1 = execute_query_to_df(conn, q5_1, "Query 5.1 Top 10 Sellers", 
                                ["Seller ID", "State", "City", "Fulfilled Orders", "Items Sold", "Merchandise Revenue ($)", "Avg Item Price ($)"])

    q5_2 = """
    WITH seller_revenue AS (
        SELECT 
            s.seller_id,
            SUM(oi.price) AS total_seller_revenue
        FROM sellers s
        JOIN order_items oi ON s.seller_id = oi.seller_id
        JOIN orders o ON oi.order_id = o.order_id
        WHERE o.order_status = 'delivered'
        GROUP BY s.seller_id
    ),
    seller_ranked AS (
        SELECT 
            seller_id,
            total_seller_revenue,
            ROW_NUMBER() OVER (ORDER BY total_seller_revenue DESC) AS seller_rank,
            COUNT(*) OVER () AS total_seller_count,
            SUM(total_seller_revenue) OVER () AS grand_total_revenue
        FROM seller_revenue
    ),
    seller_pareto AS (
        SELECT 
            seller_id,
            total_seller_revenue,
            grand_total_revenue,
            seller_rank,
            total_seller_count,
            CASE 
                WHEN seller_rank <= CEIL(0.20 * total_seller_count) THEN 'Top 20% Sellers'
                ELSE 'Remaining 80% Sellers'
            END AS seller_tier
        FROM seller_ranked
    )
    SELECT 
        seller_tier,
        COUNT(seller_id) AS seller_count,
        ROUND(COUNT(seller_id) * 100.0 / MAX(total_seller_count), 2) AS percentage_of_total_sellers,
        ROUND(SUM(total_seller_revenue), 2) AS total_tier_revenue,
        ROUND(SUM(total_seller_revenue) * 100.0 / MAX(grand_total_revenue), 2) AS percentage_of_total_revenue
    FROM seller_pareto
    GROUP BY seller_tier
    ORDER BY total_tier_revenue DESC;
    """
    df5_2 = execute_query_to_df(conn, q5_2, "Query 5.2 Pareto Seller Contribution", 
                                ["Seller Tier", "Seller Count", "Percentage of Sellers (%)", "Tier Total Revenue ($)", "Percentage of Total Revenue (%)"])

    r = append_df_to_worksheet(ws_sellers, df5_1, start_row=1, section_title="1. Top 10 Seller Performance")
    r = append_df_to_worksheet(ws_sellers, df5_2, start_row=r, section_title="2. Pareto Principle: Top 20% Seller Revenue Share")
    format_excel_worksheet(ws_sellers, freeze_row=2)

    # =========================================================================
    # SHEET 8: Delivery
    # =========================================================================
    print("Processing Sheet 8: Delivery...")
    ws_delivery = wb.create_sheet(title="Delivery")
    
    q6_1 = """
    SELECT 
        c.customer_state AS state,
        COUNT(o.order_id) AS delivered_orders,
        ROUND(AVG(TIMESTAMPDIFF(DAY, o.order_purchase_timestamp, o.order_delivered_customer_date)), 1) AS avg_total_delivery_days,
        ROUND(AVG(TIMESTAMPDIFF(DAY, o.order_purchase_timestamp, o.order_delivered_carrier_date)), 1) AS avg_fulfillment_days,
        ROUND(AVG(TIMESTAMPDIFF(DAY, o.order_delivered_carrier_date, o.order_delivered_customer_date)), 1) AS avg_carrier_transit_days
    FROM orders o
    JOIN customers c ON o.customer_id = c.customer_id
    WHERE o.order_status = 'delivered'
      AND o.order_delivered_customer_date IS NOT NULL
      AND o.order_delivered_carrier_date IS NOT NULL
    GROUP BY c.customer_state
    ORDER BY avg_total_delivery_days DESC;
    """
    df6_1 = execute_query_to_df(conn, q6_1, "Query 6.1 State Delivery Lead Times", 
                                ["Customer State", "Delivered Orders", "Avg Total Delivery (Days)", "Avg Seller Dispatch (Days)", "Avg Carrier Transit (Days)"])

    q6_2 = """
    SELECT 
        CASE 
            WHEN DATE(order_delivered_customer_date) <= DATE(order_estimated_delivery_date) THEN 'On-Time / Early'
            ELSE 'Late Delivery'
        END AS delivery_performance_status,
        COUNT(*) AS order_count,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage,
        ROUND(AVG(TIMESTAMPDIFF(DAY, order_estimated_delivery_date, order_delivered_customer_date)), 1) AS avg_days_diff_from_estimate
    FROM orders
    WHERE order_status = 'delivered'
      AND order_delivered_customer_date IS NOT NULL
    GROUP BY 
        CASE 
            WHEN DATE(order_delivered_customer_date) <= DATE(order_estimated_delivery_date) THEN 'On-Time / Early'
            ELSE 'Late Delivery'
        END;
    """
    df6_2 = execute_query_to_df(conn, q6_2, "Query 6.2 On-Time vs Delayed Rates", 
                                ["Delivery Performance Status", "Order Count", "Percentage (%)", "Avg Days Diff vs Estimate"])

    r = append_df_to_worksheet(ws_delivery, df6_1, start_row=1, section_title="1. State-Level Delivery Lead Times")
    r = append_df_to_worksheet(ws_delivery, df6_2, start_row=r, section_title="2. On-Time vs Delayed Delivery Performance")
    format_excel_worksheet(ws_delivery, freeze_row=2)

    # =========================================================================
    # SHEET 9: Reviews
    # =========================================================================
    print("Processing Sheet 9: Reviews...")
    ws_reviews = wb.create_sheet(title="Reviews")
    
    q7_1 = """
    SELECT 
        review_score,
        COUNT(*) AS review_count,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage_share
    FROM reviews
    GROUP BY review_score
    ORDER BY review_score DESC;
    """
    df7_1 = execute_query_to_df(conn, q7_1, "Query 7.1 Review Score Distribution", 
                                ["Review Score (Stars)", "Review Count", "Percentage Share (%)"])

    q7_2 = """
    SELECT 
        CASE 
            WHEN DATE(o.order_delivered_customer_date) <= DATE(o.order_estimated_delivery_date) THEN 'On-Time / Early'
            ELSE 'Late Delivery'
        END AS delivery_status,
        COUNT(r.review_id) AS total_reviews,
        ROUND(AVG(r.review_score), 2) AS average_review_score,
        SUM(CASE WHEN r.review_score = 5 THEN 1 ELSE 0 END) AS five_star_reviews,
        SUM(CASE WHEN r.review_score = 1 THEN 1 ELSE 0 END) AS one_star_reviews,
        ROUND(SUM(CASE WHEN r.review_score = 1 THEN 1 ELSE 0 END) * 100.0 / COUNT(r.review_id), 2) AS percentage_1_star
    FROM orders o
    JOIN reviews r ON o.order_id = r.order_id
    WHERE o.order_status = 'delivered'
      AND o.order_delivered_customer_date IS NOT NULL
    GROUP BY 
        CASE 
            WHEN DATE(o.order_delivered_customer_date) <= DATE(o.order_estimated_delivery_date) THEN 'On-Time / Early'
            ELSE 'Late Delivery'
        END;
    """
    df7_2 = execute_query_to_df(conn, q7_2, "Query 7.2 Delay Impact on Reviews", 
                                ["Delivery Status", "Total Reviews", "Average Review Score", "5-Star Reviews", "1-Star Reviews", "1-Star Percentage (%)"])

    q7_3 = """
    WITH order_product_categories AS (
        SELECT DISTINCT 
            oi.order_id,
            COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized') AS category_name_english
        FROM order_items oi
        JOIN products p ON oi.product_id = p.product_id
        LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
    ),
    category_review_scores AS (
        SELECT 
            opc.category_name_english,
            COUNT(r.review_id) AS review_count,
            ROUND(AVG(r.review_score), 2) AS avg_review_score
        FROM order_product_categories opc
        JOIN reviews r ON opc.order_id = r.order_id
        GROUP BY opc.category_name_english
        HAVING COUNT(r.review_id) >= 100
    ),
    ranked_categories AS (
        SELECT 
            category_name_english,
            review_count,
            avg_review_score,
            ROW_NUMBER() OVER (ORDER BY avg_review_score DESC, review_count DESC) AS rank_highest,
            ROW_NUMBER() OVER (ORDER BY avg_review_score ASC, review_count DESC) AS rank_lowest
        FROM category_review_scores
    )
    SELECT 
        'Highest Rated' AS performance_tier,
        category_name_english,
        review_count,
        avg_review_score
    FROM ranked_categories
    WHERE rank_highest <= 5

    UNION ALL

    SELECT 
        'Lowest Rated' AS performance_tier,
        category_name_english,
        review_count,
        avg_review_score
    FROM ranked_categories
    WHERE rank_lowest <= 5
    ORDER BY performance_tier DESC, avg_review_score DESC;
    """
    df7_3 = execute_query_to_df(conn, q7_3, "Query 7.3 Highest & Lowest Category Ratings", 
                                ["Performance Tier", "Category Name (English)", "Review Count", "Average Review Score"])

    r = append_df_to_worksheet(ws_reviews, df7_1, start_row=1, section_title="1. Overall Review Score Distribution")
    r = append_df_to_worksheet(ws_reviews, df7_2, start_row=r, section_title="2. Delivery Delay Impact on Review Scores")
    r = append_df_to_worksheet(ws_reviews, df7_3, start_row=r, section_title="3. Highest & Lowest Rated Product Categories")
    format_excel_worksheet(ws_reviews, freeze_row=2)

    # =========================================================================
    # SHEET 10: Payments
    # =========================================================================
    print("Processing Sheet 10: Payments...")
    ws_payments = wb.create_sheet(title="Payments")
    
    q8_1 = """
    SELECT 
        payment_type,
        COUNT(DISTINCT order_id) AS transaction_count,
        ROUND(SUM(payment_value), 2) AS total_payment_value,
        ROUND(SUM(payment_value) * 100.0 / SUM(SUM(payment_value)) OVER (), 2) AS percentage_of_total_payment,
        ROUND(AVG(payment_value), 2) AS avg_payment_amount
    FROM payments
    GROUP BY payment_type
    ORDER BY total_payment_value DESC;
    """
    df8_1 = execute_query_to_df(conn, q8_1, "Query 8.1 Payment Method Breakdown", 
                                ["Payment Type", "Transaction Count", "Total Payment Value ($)", "Percentage Share (%)", "Avg Payment Amount ($)"])

    q8_2 = """
    WITH credit_card_orders AS (
        SELECT 
            order_id,
            MAX(payment_installments) AS payment_installments,
            SUM(payment_value) AS total_order_payment_value
        FROM payments
        WHERE payment_type = 'credit_card'
        GROUP BY order_id
    )
    SELECT 
        payment_installments,
        COUNT(order_id) AS total_orders,
        ROUND(SUM(total_order_payment_value), 2) AS total_payment_value,
        ROUND(AVG(total_order_payment_value), 2) AS avg_order_payment_value
    FROM credit_card_orders
    GROUP BY payment_installments
    ORDER BY payment_installments ASC;
    """
    df8_2 = execute_query_to_df(conn, q8_2, "Query 8.2 Installment Analysis", 
                                ["Credit Card Installments", "Total Orders", "Total Payment Value ($)", "Avg Order Payment Value ($)"])

    q8_3 = """
    WITH order_payment_summary AS (
        SELECT 
            order_id,
            COUNT(payment_sequential) AS payment_splits_count,
            COUNT(DISTINCT payment_type) AS distinct_payment_types,
            SUM(payment_value) AS total_order_payment
        FROM payments
        GROUP BY order_id
    )
    SELECT 
        CASE 
            WHEN payment_splits_count = 1 THEN 'Single Payment'
            WHEN payment_splits_count > 1 AND distinct_payment_types = 1 THEN 'Split Payment (Same Method/Vouchers)'
            ELSE 'Multi-Method Payment (e.g., Credit + Voucher)'
        END AS payment_behavior,
        COUNT(*) AS order_count,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage_of_orders,
        ROUND(AVG(total_order_payment), 2) AS avg_order_payment_value
    FROM order_payment_summary
    GROUP BY 
        CASE 
            WHEN payment_splits_count = 1 THEN 'Single Payment'
            WHEN payment_splits_count > 1 AND distinct_payment_types = 1 THEN 'Split Payment (Same Method/Vouchers)'
            ELSE 'Multi-Method Payment (e.g., Credit + Voucher)'
        END
    ORDER BY order_count DESC;
    """
    df8_3 = execute_query_to_df(conn, q8_3, "Query 8.3 Multi-Payment Behavior", 
                                ["Payment Behavior Tier", "Order Count", "Percentage of Orders (%)", "Avg Order Payment Value ($)"])

    r = append_df_to_worksheet(ws_payments, df8_1, start_row=1, section_title="1. Payment Method Breakdown")
    r = append_df_to_worksheet(ws_payments, df8_2, start_row=r, section_title="2. Credit Card Installments & Order Value Analysis")
    r = append_df_to_worksheet(ws_payments, df8_3, start_row=r, section_title="3. Multi-Payment & Voucher Usage Behavior")
    format_excel_worksheet(ws_payments, freeze_row=2)

    # =========================================================================
    # CREATE DASHBOARD WORKSHEET AS FIRST SHEET (index=0)
    # =========================================================================
    print("Processing Dashboard Sheet (index=0)...")
    total_rev = float(df2_1_raw["gross_revenue"].sum()) if not df2_1_raw.empty else 0.0
    total_ord = int(df2_1_raw["delivered_orders"].sum()) if not df2_1_raw.empty else 0
    total_itm = int(df2_1_raw["total_items"].sum()) if not df2_1_raw.empty else 0
    avg_order_val = total_rev / total_ord if total_ord > 0 else 0.0
    
    if not df7_1.empty:
        avg_review_val = float((df7_1["Review Score (Stars)"] * df7_1["Review Count"]).sum() / df7_1["Review Count"].sum())
    else:
        avg_review_val = 0.0

    kpi_dict = {
        "total_revenue": total_rev,
        "total_orders": total_ord,
        "total_items": total_itm,
        "aov": avg_order_val,
        "avg_review": avg_review_val
    }

    dashboard_dfs = {
        "revenue_trend": df2_1_export[["Month", "Gross Revenue ($)"]] if not df2_1_export.empty else None,
        "orders_trend": df2_1_export[["Month", "Delivered Orders"]] if not df2_1_export.empty else None,
        "top_categories": df4_1[["Category Name (English)", "Total Revenue ($)"]].head(10) if not df4_1.empty else None,
        "payment_mix": df8_1[["Payment Type", "Transaction Count", "Percentage Share (%)"]] if not df8_1.empty else None,
        "review_dist": df7_1[["Review Score (Stars)", "Review Count", "Percentage Share (%)"]] if not df7_1.empty else None,
        "delivery_perf": df6_2[["Delivery Performance Status", "Order Count", "Percentage (%)"]] if not df6_2.empty else None,
    }

    build_dashboard_sheet(wb, kpi_dict, dashboard_dfs)

    # Close connection
    conn.close()

    # Save Workbook
    wb.save(output_filepath)
    print("=================================================================")
    print(f"[SUCCESS] Business Analysis Excel Workbook generated successfully!")
    print(f"File Path: {output_filepath}")
    print("=================================================================")


if __name__ == "__main__":
    main()
